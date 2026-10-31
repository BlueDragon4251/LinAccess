#include "dao_database.hpp"
#include "access_query.hpp"
#include "sqlite_execution.hpp"
#include <QSqlError>
#include <QSqlQuery>
#include <QSqlRecord>
#include <QSet>
#include <QVariant>
#include <algorithm>

namespace linaccess {
namespace {
QString text(std::string_view value) { return QString::fromUtf8(value.data(), value.size()); }
QString quote(QString name) { return '"' + name.replace('"', "\"\"") + '"'; }
std::string lower(std::string_view value) { return text(value).toCaseFolded().toStdString(); }
void require(bool valid, const char* message) { if (!valid) throw std::runtime_error(message); }
void exec(QSqlQuery& query, const QString& sql) {
    if (!query.exec(sql)) throw std::runtime_error("DAO: SQL operation failed");
}
[[noreturn]] void queryFailure(const QSqlQuery& query) {
    const auto message = query.lastError().text().toCaseFolded();
    if (message.contains("no such function")) throw VbaRuntimeError(3085,"DAO: SQL function unsupported");
    if (message.contains("no such table")) throw VbaRuntimeError(3078,"DAO: SQL dependency unavailable");
    if (message.contains("syntax error") || message.contains("circularly defined")) throw VbaRuntimeError(3075,"DAO: SQL grammar unsupported");
    if (message.contains("interrupted")) throw VbaRuntimeError(3035,"DAO: query execution budget exceeded");
    throw VbaRuntimeError(3146,"DAO: SQL operation failed");
}
VbaValue fromSql(const QVariant& value) {
    if (value.isNull()) return {};
    switch (value.metaType().id()) {
    case QMetaType::Int: case QMetaType::UInt: case QMetaType::LongLong:
    case QMetaType::ULongLong: case QMetaType::Double:
        return VbaValue::numeric(value.toDouble());
    case QMetaType::QByteArray: throw std::runtime_error("DAO: BLOB Variant not implemented");
    default: return VbaValue::string(value.toString().toUtf8().toStdString());
    }
}
QVariant toSql(const VbaValue& value) {
    switch (value.kind) {
    case VbaValue::Kind::Null: return {};
    case VbaValue::Kind::Empty: return QString{};
    case VbaValue::Kind::Number: return value.number;
    case VbaValue::Kind::Boolean: return value.boolean ? -1 : 0;
    case VbaValue::Kind::Text: return text(value.text);
    default: throw std::runtime_error("DAO: Object cannot be stored in a field");
    }
}
class Recordset;
class Field final : public VbaObject {
public:
    Field(std::shared_ptr<Recordset> owner, QString name) : owner_(std::move(owner)), name_(std::move(name)) {}
    VbaValue member(std::string_view name, const std::vector<VbaValue>& args) override;
    void assign(std::string_view name, const VbaValue& value) override;
private:
    std::shared_ptr<Recordset> owner_;
    QString name_;
};
class Recordset final : public VbaObject, public std::enable_shared_from_this<Recordset> {
public:
    Recordset(QSqlDatabase db, QString source, bool snapshot, QMap<QString,QVariant> bindings = {}, int topLimit = 0, bool hiddenRank = false)
        : db_(std::move(db)), source_(std::move(source)), readOnly_(snapshot), bindings_(std::move(bindings)), topLimit_(topLimit), hiddenRank_(hiddenRank) {
        table_ = db_.tables().contains(source_, Qt::CaseInsensitive);
        readOnly_ |= !table_;
        reload();
    }
    VbaValue member(std::string_view name, const std::vector<VbaValue>& args) override {
        require(!closed_, "DAO: Recordset is closed");
        const auto key = lower(name);
        if (key == "fields") {
            require(args.size() == 1, "DAO: Fields requires one field name/index");
            QString field;
            if (args[0].kind == VbaValue::Kind::Number) {
                const auto index = static_cast<int>(args[0].asNumber());
                require(index >= 0 && index < columns_.size(), "DAO: Field index out of range");
                field = columns_[index];
            } else field = text(args[0].asText());
            (void)column(field);
            return VbaValue::reference(std::make_shared<Field>(shared_from_this(), field));
        }
        require(args.empty(), "DAO: Unexpected Recordset arguments");
        if (key == "eof") return VbaValue::logical(rows_.empty() || position_ >= rows_.size());
        if (key == "bof") return VbaValue::logical(rows_.empty() || position_ < 0);
        if (key == "recordcount") return VbaValue::numeric(rows_.size());
        if (key == "movefirst" || key == "movelast" || key == "movenext" || key == "moveprevious") {
            require(mode_ == Mode::Browse, "DAO: Update or CancelUpdate before navigation");
            if (key == "movefirst") position_ = 0;
            if (key == "movelast") position_ = rows_.empty() ? 0 : rows_.size() - 1;
            if (key == "movenext") { require(position_ < rows_.size(), "DAO: Already at EOF"); ++position_; }
            if (key == "moveprevious") { require(position_ >= 0 && !rows_.empty(), "DAO: Already at BOF"); --position_; }
            return {};
        }
        if (key == "edit" || key == "addnew") {
            require(!readOnly_, "DAO: Recordset is read-only");
            require(mode_ == Mode::Browse, "DAO: Edit already in progress");
            if (key == "edit") { current(); pending_ = rows_[position_]; mode_ = Mode::Edit; }
            else { pending_ = QVariantList(columns_.size()); mode_ = Mode::Add; }
            changed_.clear();
            return {};
        }
        if (key == "cancelupdate") { pending_.clear(); changed_.clear(); mode_ = Mode::Browse; return {}; }
        if (key == "update") { update(); return {}; }
        if (key == "delete") {
            require(!readOnly_ && mode_ == Mode::Browse, "DAO: Cannot delete this Recordset");
            current();
            QSqlQuery query(db_);
            require(query.prepare("DELETE FROM " + quote(source_) + " WHERE rowid=?"), "DAO: Delete prepare failed");
            query.addBindValue(ids_[position_]);
            require(query.exec() && query.numRowsAffected() == 1, "DAO: Delete failed or record changed");
            reload();
            return {};
        }
        if (key == "requery") {
            require(mode_ == Mode::Browse, "DAO: Pending edit");
            reload(); return {};
        }
        if (key == "close") { closed_ = true; rows_.clear(); ids_.clear(); return {}; }
        return value(text(name)); // implicit default Fields(name).Value
    }
    void assign(std::string_view name, const VbaValue& value) override { set(text(name), value); }
    VbaValue value(const QString& name) const {
        require(!closed_, "DAO: Recordset is closed");
        const int index = column(name);
        if (mode_ != Mode::Browse) return fromSql(pending_[index]);
        current(); return fromSql(rows_[position_][index]);
    }
    void set(const QString& name, const VbaValue& value) {
        require(!closed_ && mode_ != Mode::Browse, "DAO: Edit or AddNew required");
        const auto index = column(name);
        pending_[index] = toSql(value); changed_.insert(index);
    }
private:
    enum class Mode { Browse, Edit, Add };
    int column(const QString& name) const {
        for (int i = 0; i < columns_.size(); ++i)
            if (columns_[i].compare(name, Qt::CaseInsensitive) == 0) return i;
        throw std::runtime_error("DAO: Unknown field");
    }
    void current() const { require(position_ >= 0 && position_ < rows_.size(), "DAO: No current record"); }
    void reload() {
        SqliteExecutionBudget budget(db_);
        QSqlQuery query(db_);
        if (table_) exec(query,"SELECT rowid,* FROM " + quote(source_));
        else {
            if (!query.prepare(source_)) queryFailure(query);
            for (auto it = bindings_.cbegin(); it != bindings_.cend(); ++it) query.bindValue(it.key(),it.value());
            if (!query.exec()) queryFailure(query);
        }
        const int offset = table_ ? 1 : 0;
        columns_.clear(); rows_.clear(); ids_.clear(); position_ = 0;
        const int end = query.record().count() - (hiddenRank_ ? 1 : 0);
        for (int i = offset; i < end; ++i) columns_ << query.record().fieldName(i);
        while (query.next()) {
            if (hiddenRank_ && query.value(query.record().count()-1).toLongLong() > topLimit_) break;
            require(rows_.size() < 100000, "DAO: Recordset row limit exceeded");
            QVariantList row;
            if (table_) ids_.append(query.value(0));
            for (int i = offset; i < end; ++i) row << query.value(i);
            rows_.append(row);
        }
        if (query.lastError().isValid()) queryFailure(query);
    }
    void update() {
        require(!closed_ && !readOnly_ && mode_ != Mode::Browse, "DAO: No editable record buffer");
        const bool adding = mode_ == Mode::Add;
        QVariant id = adding ? QVariant{} : ids_[position_];
        QStringList columns, values, assignments;
        QList<int> indices = changed_.values();
        std::sort(indices.begin(), indices.end());
        for (const int i : indices) {
            columns << quote(columns_[i]); values << "?"; assignments << quote(columns_[i]) + "=?";
        }
        QSqlQuery query(db_);
        if (adding && columns.empty()) exec(query, "INSERT INTO " + quote(source_) + " DEFAULT VALUES");
        else if (!columns.empty()) {
            require(query.prepare(adding ? "INSERT INTO " + quote(source_) + " (" + columns.join(',')
                        + ") VALUES (" + values.join(',') + ")"
                    : "UPDATE " + quote(source_) + " SET " + assignments.join(',') + " WHERE rowid=?"),
                    "DAO: Update prepare failed");
            for (const int i : indices) query.addBindValue(pending_[i]);
            if (!adding) query.addBindValue(id);
            require(query.exec() && query.numRowsAffected() == 1, "DAO: Update failed or record changed");
        }
        if (adding) id = query.lastInsertId();
        mode_ = Mode::Browse; pending_.clear(); changed_.clear();
        reload(); position_ = ids_.indexOf(id);
    }
    QSqlDatabase db_;
    QString source_;
    QMap<QString,QVariant> bindings_;
    bool table_{}, readOnly_{}, closed_{};
    QStringList columns_;
    QVector<QVariantList> rows_;
    QVector<QVariant> ids_;
    QVariantList pending_;
    QSet<int> changed_;
    qsizetype position_{};
    Mode mode_{Mode::Browse};
    int topLimit_{};
    bool hiddenRank_{};
};
VbaValue Field::member(std::string_view name, const std::vector<VbaValue>& args) {
    require(args.empty(), "DAO: Unexpected field arguments");
    if (lower(name) == "value") return owner_->value(name_);
    if (lower(name) == "name") return VbaValue::string(name_.toUtf8().toStdString());
    throw std::runtime_error("DAO: Unknown Field member");
}
void Field::assign(std::string_view name, const VbaValue& value) {
    require(lower(name) == "value", "DAO: Field property is read-only");
    owner_->set(name_, value);
}
class QueryDefinition;
class QueryParameter final : public VbaObject {
public:
    QueryParameter(std::shared_ptr<QueryDefinition> query, QString name) : query_(std::move(query)), name_(std::move(name)) {}
    VbaValue member(std::string_view name, const std::vector<VbaValue>& args) override;
    void assign(std::string_view name, const VbaValue& value) override;
private:
    std::shared_ptr<QueryDefinition> query_;
    QString name_;
};
class QueryDefinition final : public VbaObject, public std::enable_shared_from_this<QueryDefinition> {
public:
    QueryDefinition(QSqlDatabase db, QString source) : db_(std::move(db)), source_(std::move(source)), plan_(compileAccessQuery(source_)) {}
    VbaValue member(std::string_view name, const std::vector<VbaValue>& args) override {
        require(!closed_,"DAO: QueryDef is closed");
        const auto key = lower(name);
        if (key == "parameters") {
            require(args.size() == 1,"DAO: Parameter name/index required");
            QString parameter;
            if (args[0].kind == VbaValue::Kind::Number) {
                const auto index = static_cast<int>(args[0].asNumber());
                require(index >= 0 && index < plan_.parameters.size(),"DAO: Parameter index out of range");
                parameter = plan_.parameters[index].name;
            } else parameter = text(args[0].asText());
            checkParameter(parameter);
            return VbaValue::reference(std::make_shared<QueryParameter>(shared_from_this(),parameter));
        }
        if (key == "openrecordset") {
            require(args.empty() || (args.size() == 1 && (args[0].asNumber() == 2 || args[0].asNumber() == 4)),"DAO: Recordset type unsupported");
            require(plan_.sql.trimmed().startsWith("SELECT ",Qt::CaseInsensitive),"DAO: SELECT QueryDef required");
            return VbaValue::reference(std::make_shared<Recordset>(db_,plan_.sql,true,bindAccessQueryParameters(plan_,values_),plan_.topLimit,plan_.hiddenRank));
        }
        if (key == "execute") {
            require(args.empty() || (args.size() == 1 && args[0].asNumber() == 128),"DAO: Execute options unsupported");
            SqliteExecutionBudget budget(db_);
            QSqlQuery query(db_);
            if (!query.prepare(plan_.sql)) queryFailure(query);
            const auto bindings = bindAccessQueryParameters(plan_,values_);
            for (auto it = bindings.cbegin(); it != bindings.cend(); ++it) query.bindValue(it.key(),it.value());
            if (!query.exec()) queryFailure(query);
            if (query.isSelect()) throw VbaRuntimeError(3065,"DAO: Execute cannot run a SELECT query");
            affected_ = query.numRowsAffected(); return {};
        }
        require(args.empty(),"DAO: Unexpected QueryDef arguments");
        if (key == "sql") return VbaValue::string(source_.toUtf8().toStdString());
        if (key == "recordsaffected") return VbaValue::numeric(affected_);
        if (key == "close") { closed_ = true; return {}; }
        throw VbaRuntimeError(438,"DAO: QueryDef member unsupported");
    }
    void assign(std::string_view name, const VbaValue& value) override {
        require(!closed_ && lower(name) == "sql","DAO: QueryDef property is not writable");
        const auto source = text(value.asText());
        const auto compiled = compileAccessQuery(source);
        source_ = source; plan_ = compiled; values_.clear();
    }
    VbaValue parameter(const QString& name) const {
        checkParameter(name); return fromSql(values_.value(name.toCaseFolded()));
    }
    void setParameter(const QString& name, const VbaValue& value) {
        checkParameter(name); values_.insert(name.toCaseFolded(),toSql(value));
    }
private:
    void checkParameter(const QString& name) const {
        require(!closed_,"DAO: QueryDef is closed");
        for (const auto& parameter : plan_.parameters)
            if (parameter.name.compare(name,Qt::CaseInsensitive) == 0) return;
        throw VbaRuntimeError(3265,"DAO: Parameter not found");
    }
    QSqlDatabase db_;
    QString source_;
    AccessQueryPlan plan_;
    QMap<QString,QVariant> values_;
    qlonglong affected_{};
    bool closed_{};
};
VbaValue QueryParameter::member(std::string_view name, const std::vector<VbaValue>& args) {
    require(args.empty(),"DAO: Unexpected Parameter arguments");
    if (lower(name) == "value") return query_->parameter(name_);
    if (lower(name) == "name") return VbaValue::string(name_.toUtf8().toStdString());
    throw VbaRuntimeError(438,"DAO: Parameter member unsupported");
}
void QueryParameter::assign(std::string_view name, const VbaValue& value) {
    require(lower(name) == "value","DAO: Parameter property is not writable");
    query_->setParameter(name_,value);
}
class Database final : public VbaObject {
public:
    explicit Database(QSqlDatabase db) : db_(std::move(db)) {}
    VbaValue member(std::string_view name, const std::vector<VbaValue>& args) override {
        require(db_.isOpen(), "DAO: Database is closed");
        const auto key = lower(name);
        if (key == "querydefs") {
            require(args.size() == 1,"DAO: QueryDef name required");
            return VbaValue::reference(std::make_shared<QueryDefinition>(db_,accessQuerySource(db_,text(args[0].asText()))));
        }
        if (key == "createquerydef") {
            require(args.size() == 2 && args[0].asText().empty(),"DAO: Only temporary CreateQueryDef is currently supported");
            return VbaValue::reference(std::make_shared<QueryDefinition>(db_,text(args[1].asText())));
        }
        if (key == "openrecordset") {
            require(args.size() == 1 || args.size() == 2, "DAO: OpenRecordset arguments unsupported");
            const int mode = args.size() == 2 ? static_cast<int>(args[1].asNumber()) : 2;
            require(mode == 2 || mode == 4, "DAO: Recordset type unsupported");
            const auto source = text(args[0].asText());
            if (db_.tables().contains(source,Qt::CaseInsensitive))
                return VbaValue::reference(std::make_shared<Recordset>(db_,source,mode == 4));
            const auto sql = source.trimmed().startsWith("SELECT ",Qt::CaseInsensitive) ? source : accessQuerySource(db_,source);
            const auto plan = compileAccessQuery(sql);
            require(plan.sql.trimmed().startsWith("SELECT ",Qt::CaseInsensitive),"DAO: SELECT required");
            return VbaValue::reference(std::make_shared<Recordset>(db_,plan.sql,true,bindAccessQueryParameters(plan,{}),plan.topLimit,plan.hiddenRank));
        }
        if (key == "execute") {
            require(args.size() == 1 || (args.size() == 2 && args[1].asNumber() == 128), "DAO: Execute arguments unsupported");
            SqliteExecutionBudget budget(db_);
            const auto plan = compileAccessQuery(text(args[0].asText()));
            const auto bindings = bindAccessQueryParameters(plan,{});
            QSqlQuery query(db_);
            if (!query.prepare(plan.sql)) queryFailure(query);
            for (auto it = bindings.cbegin(); it != bindings.cend(); ++it) query.bindValue(it.key(),it.value());
            if (!query.exec()) queryFailure(query);
            if (query.isSelect()) throw VbaRuntimeError(3065,"DAO: Execute cannot run a SELECT query");
            return {};
        }
        throw std::runtime_error("DAO: Database member unsupported");
    }
private:
    QSqlDatabase db_;
};
}
std::shared_ptr<VbaObject> makeDaoDatabase(QSqlDatabase database) {
    return std::make_shared<Database>(std::move(database));
}
}
