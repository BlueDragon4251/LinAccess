#include "dao_database.hpp"
#include "vba_qt_value.hpp"
#include "access_query.hpp"
#include "jet_sql_functions.hpp"
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
    if (const auto code=jetSqlFunctionErrorNumber(message.toUtf8().toStdString()))
        throw VbaRuntimeError(*code,"Jet SQL function argument invalid");
    if (message.contains("already exists")) throw VbaRuntimeError(3010,"DAO: destination table already exists");
    if (message.contains("no such function")) throw VbaRuntimeError(3085,"DAO: SQL function unsupported");
    if (message.contains("no such table")) throw VbaRuntimeError(3078,"DAO: SQL dependency unavailable");
    if ((message.contains("syntax error") || message.contains("incomplete input")) || message.contains("circularly defined")) throw VbaRuntimeError(3075,"DAO: SQL grammar unsupported");
    if (message.contains("interrupted")) throw VbaRuntimeError(3035,"DAO: query execution budget exceeded");
    throw VbaRuntimeError(3146,"DAO: SQL operation failed");
}
qlonglong executeAction(QSqlDatabase db,const AccessQueryPlan& plan,const QMap<QString,QVariant>& bindings) {
    if (plan.sql.trimmed().startsWith("SELECT",Qt::CaseInsensitive)) throw VbaRuntimeError(3065,"DAO: Execute cannot run a SELECT query");
    const bool makeTable=!plan.makeTable.isEmpty();
    QSqlQuery guard(db);
    if (makeTable && !guard.exec("SAVEPOINT la_make_table")) queryFailure(guard);
    try {
        qlonglong affected{};
        {
            SqliteExecutionBudget budget(db);
            QSqlQuery query(db);
            if (!query.prepare(plan.sql)) queryFailure(query);
            for (auto it=bindings.cbegin();it!=bindings.cend();++it) query.bindValue(it.key(),it.value());
            if (!query.exec()) queryFailure(query);
            if (query.isSelect()) throw VbaRuntimeError(3065,"DAO: Execute cannot run a SELECT query");
            affected=std::max<qlonglong>(0,query.numRowsAffected()); query.finish();
            if (makeTable) {
                if (!query.exec("SELECT COUNT(*) FROM "+quote(plan.makeTable)) || !query.next()) queryFailure(query);
                affected=query.value(0).toLongLong(); query.finish();
            }
        }
        if (makeTable && !guard.exec("RELEASE la_make_table")) queryFailure(guard);
        return affected;
    } catch (...) {
        if (makeTable) { guard.exec("ROLLBACK TO la_make_table"); guard.exec("RELEASE la_make_table"); }
        throw;
    }
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
VbaValue fieldCollection(std::shared_ptr<Recordset> owner);
class Field final : public VbaObject {
public:
    Field(std::shared_ptr<Recordset> owner, QString name) : owner_(std::move(owner)), name_(std::move(name)) {}
    VbaValue scalarValue() override { return member("Value",{}); }
    VbaValue member(std::string_view name, const std::vector<VbaValue>& args) override;
    void assign(std::string_view name, const VbaValue& value) override;
private:
    std::shared_ptr<Recordset> owner_;
    QString name_;
};
class Recordset final : public VbaObject, public std::enable_shared_from_this<Recordset> {
public:
    Recordset(QSqlDatabase db, QString source, bool snapshot, QMap<QString,QVariant> bindings = {},AccessQueryPlan plan={})
        : db_(std::move(db)), source_(std::move(source)), readOnly_(snapshot), bindings_(std::move(bindings)),queryPlan_(std::move(plan)) {
        table_ = db_.tables().contains(source_, Qt::CaseInsensitive);
        readOnly_ |= !table_;
        reload();
    }
    VbaValue defaultMember(const std::vector<VbaValue>& args,bool reference) override {
        if (args.size()!=1) throw VbaRuntimeError(450,"DAO: Recordset default field requires one index");
        const auto field=member("Fields",args); return reference ? field : field.object->member("Value",{});
    }
    VbaValue item(std::string_view name) override { return value(text(name)); }
    VbaValue objectItem(std::string_view name) override { return member("Fields",{VbaValue::string(std::string(name))}); }
    void assignItem(std::string_view name,const VbaValue& value) override { set(text(name),value); }
    QStringList fieldNames() const { require(!closed_,"DAO: Recordset is closed"); return columns_; }
    VbaValue member(std::string_view name, const std::vector<VbaValue>& args) override {
        require(!closed_, "DAO: Recordset is closed");
        const auto key = lower(name);
        if (key == "fields") {
            if (args.empty()) return fieldCollection(shared_from_this());
            require(args.size() == 1, "DAO: Fields requires one field name/index");
            QString field;
            if (args[0].kind == VbaValue::Kind::Number) {
                const auto index = sqlFieldValue(args[0],"INTEGER").toLongLong();
                if (index < 0 || index >= columns_.size()) throw VbaRuntimeError(3265,"DAO: Field index out of range");
                field = columns_[static_cast<int>(index)];
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
        if (mode_ != Mode::Browse) return qtFieldValue(pending_[index],types_.value(name.toCaseFolded()));
        current(); return qtFieldValue(rows_[position_][index],types_.value(name.toCaseFolded()));
    }
    void set(const QString& name, const VbaValue& value) {
        require(!closed_ && mode_ != Mode::Browse, "DAO: Edit or AddNew required");
        const auto index = column(name);
        pending_[index] = sqlFieldValue(value,types_.value(name.toCaseFolded())); changed_.insert(index);
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
        types_ = table_ ? sqliteColumnTypes(db_,source_) : QMap<QString,QString>{};
        if (!table_) {
            auto plan=queryPlan_;
            if (plan.sql.isEmpty()) plan.sql=source_;
            const auto result = executeAccessQueryPlan(db_,plan,bindings_);
            columns_ = result.columns; rows_ = result.rows; ids_.clear(); position_ = 0;
            return;
        }
        SqliteExecutionBudget budget(db_);
        QSqlQuery query(db_);
        exec(query,"SELECT rowid,* FROM " + quote(source_));
        const int offset = table_ ? 1 : 0;
        columns_.clear(); rows_.clear(); ids_.clear(); position_ = 0;
        const int end = query.record().count();
        for (int i = offset; i < end; ++i) columns_ << query.record().fieldName(i);
        while (query.next()) {
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
    QMap<QString,QString> types_;
    QVector<QVariantList> rows_;
    QVector<QVariant> ids_;
    QVariantList pending_;
    QSet<int> changed_;
    qsizetype position_{};
    Mode mode_{Mode::Browse};
    AccessQueryPlan queryPlan_;
};
class Fields final : public VbaObject {
public:
    explicit Fields(std::shared_ptr<Recordset> owner) : owner_(std::move(owner)) {}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        const auto key = lower(name);
        if (key == "count" && args.empty()) return VbaValue::numeric(owner_->fieldNames().size());
        if (key == "item" && args.size() == 1) return owner_->member("Fields",args);
        throw VbaRuntimeError(438,"DAO: Fields member unsupported");
    }
    VbaValue item(std::string_view name) override { return owner_->member("Fields",{VbaValue::string(std::string(name))}); }
    std::vector<VbaValue> enumerate() override {
        std::vector<VbaValue> values;
        for (const auto& name : owner_->fieldNames())
            values.push_back(VbaValue::reference(std::make_shared<Field>(owner_,name)));
        return values;
    }
private:
    std::shared_ptr<Recordset> owner_;
};
VbaValue fieldCollection(std::shared_ptr<Recordset> owner) {
    return VbaValue::reference(std::make_shared<Fields>(std::move(owner)));
}

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
    QueryParameter(std::shared_ptr<QueryDefinition> query, QString name,quint64 generation) : query_(std::move(query)), name_(std::move(name)),generation_(generation) {}
    VbaValue scalarValue() override { return member("Value",{}); }
    VbaValue member(std::string_view name, const std::vector<VbaValue>& args) override;
    void assign(std::string_view name, const VbaValue& value) override;
private:
    std::shared_ptr<QueryDefinition> query_;
    QString name_;
    quint64 generation_{};
};
class QueryParameters;
class QueryDefinition final : public VbaObject, public std::enable_shared_from_this<QueryDefinition> {
public:
    QueryDefinition(QSqlDatabase db, QString source) : db_(std::move(db)), source_(std::move(source)), plan_(compileAccessQuery(source_)) {}
    VbaValue member(std::string_view name, const std::vector<VbaValue>& args) override {
        if (closed_) throw VbaRuntimeError(3420,"DAO: QueryDef is closed");
        const auto key = lower(name);
        if (key == "parameters") {
            auto collection=parameters();
            return args.empty() ? VbaValue::reference(collection) : collection->member("Item",args);
        }
        if (key == "openrecordset") {
            require(args.empty() || (args.size() == 1 && (args[0].asNumber() == 2 || args[0].asNumber() == 4)),"DAO: Recordset type unsupported");
            require(plan_.sql.trimmed().startsWith("SELECT ",Qt::CaseInsensitive),"DAO: SELECT QueryDef required");
            return VbaValue::reference(std::make_shared<Recordset>(db_,plan_.sql,true,bindAccessQueryParameters(plan_,values_),plan_));
        }
        if (key == "execute") {
            require(args.empty() || (args.size() == 1 && args[0].asNumber() == 128),"DAO: Execute options unsupported");
            affected_=executeAction(db_,plan_,bindAccessQueryParameters(plan_,values_)); return {};
        }
        require(args.empty(),"DAO: Unexpected QueryDef arguments");
        if (key == "sql") return VbaValue::string(source_.toUtf8().toStdString());
        if (key == "recordsaffected") return VbaValue::numeric(affected_);
        if (key == "close") { closed_ = true; return {}; }
        throw VbaRuntimeError(438,"DAO: QueryDef member unsupported");
    }
    void assign(std::string_view name, const VbaValue& value) override {
        if (closed_) throw VbaRuntimeError(3420,"DAO: QueryDef is closed");
        require(lower(name) == "sql","DAO: QueryDef property is not writable");
        const auto source = text(value.asText());
        const auto compiled = compileAccessQuery(source);
        source_ = source; plan_ = compiled; values_.clear(); parameterObjects_.clear(); ++generation_;
    }
    VbaValue parameter(const QString& name) const {
        checkParameter(name); return fromSql(values_.value(name.toCaseFolded()));
    }
    void setParameter(const QString& name, const VbaValue& value) {
        checkParameter(name); values_.insert(name.toCaseFolded(),toSql(value));
    }
private:
    friend class QueryParameter;
    friend class QueryParameters;
    std::shared_ptr<VbaObject> parameters();
    void parameterGeneration(quint64 generation) const {
        if (closed_ || generation!=generation_) throw VbaRuntimeError(3420,"DAO: Parameter is no longer valid");
    }
    void checkParameter(const QString& name) const {
        if (closed_) throw VbaRuntimeError(3420,"DAO: QueryDef is closed");
        for (const auto& parameter : plan_.parameters)
            if (parameter.name.compare(name,Qt::CaseInsensitive) == 0) return;
        throw VbaRuntimeError(3265,"DAO: Parameter not found");
    }
    QSqlDatabase db_;
    QString source_;
    AccessQueryPlan plan_;
    QMap<QString,QVariant> values_;
    std::weak_ptr<VbaObject> parameters_;
    QMap<QString,std::weak_ptr<VbaObject>> parameterObjects_;
    quint64 generation_{};
    qlonglong affected_{};
    bool closed_{};
};
class QueryParameters final : public VbaObject {
public:
    explicit QueryParameters(std::shared_ptr<QueryDefinition> query) : query_(std::move(query)) {}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        if (query_->closed_) throw VbaRuntimeError(3420,"DAO: QueryDef is closed");
        const auto key=lower(name);
        if (key=="count") {
            if (!args.empty()) throw VbaRuntimeError(450,"DAO: Count takes no arguments");
            return VbaValue::numeric(query_->plan_.parameters.size());
        }
        if (key!="item") throw VbaRuntimeError(438,"DAO: Parameters collection member unsupported");
        if (args.size()!=1) throw VbaRuntimeError(450,"DAO: Parameter name/index required");
        const auto& value=args.front();
        if (value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Invalid Null Parameter index");
        QString target;
        if (value.kind==VbaValue::Kind::Number || value.kind==VbaValue::Kind::Currency || value.kind==VbaValue::Kind::Boolean || value.kind==VbaValue::Kind::Empty) {
            const auto index=sqlFieldValue(value,"INTEGER").toLongLong();
            if (index<0 || index>=query_->plan_.parameters.size()) throw VbaRuntimeError(3265,"DAO: Parameter index not found");
            target=query_->plan_.parameters[index].name;
        } else {
            const auto requested=text(value.asText());
            for (const auto& parameter : query_->plan_.parameters)
                if (parameter.name.compare(requested,Qt::CaseInsensitive)==0) { target=parameter.name; break; }
            if (target.isEmpty()) throw VbaRuntimeError(3265,"DAO: Parameter not found");
        }
        auto& cached=query_->parameterObjects_[target.toCaseFolded()];
        auto parameter=cached.lock();
        if (!parameter) { parameter=std::make_shared<QueryParameter>(query_,target,query_->generation_); cached=parameter; }
        return VbaValue::reference(parameter);
    }
    VbaValue item(std::string_view name) override { return member("Item",{VbaValue::string(std::string(name))}); }
    std::vector<VbaValue> enumerate() override {
        if (query_->closed_) throw VbaRuntimeError(3420,"DAO: QueryDef is closed");
        std::vector<VbaValue> result;
        for (int i=0;i<query_->plan_.parameters.size();++i) result.push_back(member("Item",{VbaValue::numeric(i)}));
        return result;
    }
private:
    std::shared_ptr<QueryDefinition> query_;
};
std::shared_ptr<VbaObject> QueryDefinition::parameters() {
    auto collection=parameters_.lock();
    if (!collection) { collection=std::make_shared<QueryParameters>(shared_from_this()); parameters_=collection; }
    return collection;
}
VbaValue QueryParameter::member(std::string_view name, const std::vector<VbaValue>& args) {
    query_->parameterGeneration(generation_);
    if (!args.empty()) throw VbaRuntimeError(450,"DAO: Unexpected Parameter arguments");
    if (lower(name) == "value") return query_->parameter(name_);
    if (lower(name) == "name") return VbaValue::string(name_.toUtf8().toStdString());
    throw VbaRuntimeError(438,"DAO: Parameter member unsupported");
}
void QueryParameter::assign(std::string_view name, const VbaValue& value) {
    query_->parameterGeneration(generation_);
    require(lower(name) == "value","DAO: Parameter property is not writable");
    query_->setParameter(name_,value);
}
QStringList nativeTableNames(QSqlDatabase db) {
    if (!db.isOpen()) throw VbaRuntimeError(3420,"DAO: Database is closed");
    QStringList result;
    QSqlQuery query(db);
    if (!query.exec("SELECT name FROM sqlite_schema WHERE type='table' ORDER BY rowid")) queryFailure(query);
    while (query.next()) {
        const auto name=query.value(0).toString();
        if (!name.startsWith("sqlite_",Qt::CaseInsensitive) && !name.startsWith("_la_",Qt::CaseInsensitive)) result.append(name);
    }
    if (query.lastError().isValid()) queryFailure(query);
    return result;
}
class TableDefinition final : public VbaObject {
public:
    TableDefinition(QSqlDatabase db,QString name) : db_(std::move(db)),name_(std::move(name)) {}
    bool valid() {
        if (invalidated_ || !db_.isOpen() || !nativeTableNames(db_).contains(name_,Qt::CaseInsensitive)) invalidated_=true;
        return !invalidated_;
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        if (!valid()) throw VbaRuntimeError(3420,"DAO: TableDef is unavailable");
        const auto key=lower(name);
        if (key=="name") {
            if (!args.empty()) throw VbaRuntimeError(450,"DAO: Name takes no arguments");
            return VbaValue::string(name_.toStdString());
        }
        if (key=="openrecordset") {
            if (args.size()>1) throw VbaRuntimeError(450,"DAO: TableDef OpenRecordset arguments unsupported");
            if (!args.empty() && args[0].kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Invalid Null recordset mode");
            const auto mode=args.empty() ? 2 : sqlFieldValue(args[0],"INTEGER").toLongLong();
            if (mode!=2 && mode!=4) throw VbaRuntimeError(3251,"DAO: TableDef recordset mode unsupported");
            return VbaValue::reference(std::make_shared<Recordset>(db_,name_,mode==4));
        }
        throw VbaRuntimeError(438,"DAO: TableDef member unsupported");
    }
private:
    QSqlDatabase db_;
    QString name_;
    bool invalidated_{};
};
class TableDefinitions final : public VbaObject {
public:
    explicit TableDefinitions(QSqlDatabase db) : db_(std::move(db)) {}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        const auto names=nativeTableNames(db_); const auto key=lower(name);
        if (key=="count") {
            if (!args.empty()) throw VbaRuntimeError(450,"DAO: Count takes no arguments");
            return VbaValue::numeric(names.size());
        }
        if (key!="item") throw VbaRuntimeError(438,"DAO: TableDefs member unsupported");
        if (args.size()!=1) throw VbaRuntimeError(450,"DAO: TableDef name/index required");
        const auto& value=args[0]; QString target;
        if (value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Invalid Null TableDef index");
        if (value.kind==VbaValue::Kind::Number || value.kind==VbaValue::Kind::Currency || value.kind==VbaValue::Kind::Boolean || value.kind==VbaValue::Kind::Empty) {
            const auto index=sqlFieldValue(value,"INTEGER").toLongLong();
            if (index<0 || index>=names.size()) throw VbaRuntimeError(3265,"DAO: TableDef index not found");
            target=names[index];
        } else {
            const auto requested=text(value.asText());
            for (const auto& candidate:names) if (candidate.compare(requested,Qt::CaseInsensitive)==0) { target=candidate; break; }
            if (target.isEmpty()) throw VbaRuntimeError(3265,"DAO: TableDef not found");
        }
        auto& cached=definitions_[target.toCaseFolded()]; auto definition=cached.lock();
        if (!definition || !definition->valid()) { definition=std::make_shared<TableDefinition>(db_,target); cached=definition; }
        return VbaValue::reference(definition);
    }
    VbaValue item(std::string_view name) override { return member("Item",{VbaValue::string(std::string(name))}); }
    std::vector<VbaValue> enumerate() override {
        const auto names=nativeTableNames(db_); std::vector<VbaValue> result;
        for (const auto& name:names) result.push_back(member("Item",{VbaValue::string(name.toStdString())}));
        return result;
    }
private:
    QSqlDatabase db_;
    QMap<QString,std::weak_ptr<TableDefinition>> definitions_;
};
class Database final : public VbaObject {
public:
    explicit Database(QSqlDatabase db) : db_(std::move(db)) {}
    VbaValue member(std::string_view name, const std::vector<VbaValue>& args) override {
        require(db_.isOpen(), "DAO: Database is closed");
        const auto key = lower(name);
        if (key=="tabledefs") {
            auto definitions=tableDefinitions_.lock();
            if (!definitions) { definitions=std::make_shared<TableDefinitions>(db_); tableDefinitions_=definitions; }
            if (args.empty()) return VbaValue::reference(definitions);
            return definitions->member("Item",args);
        }
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
            return VbaValue::reference(std::make_shared<Recordset>(db_,plan.sql,true,bindAccessQueryParameters(plan,{}),plan));
        }
        if (key == "execute") {
            require(args.size() == 1 || (args.size() == 2 && args[1].asNumber() == 128), "DAO: Execute arguments unsupported");
            const auto plan = compileAccessQuery(text(args[0].asText()));
            affected_=executeAction(db_,plan,bindAccessQueryParameters(plan,{}));
            return {};
        }
        if (key=="recordsaffected" && args.empty()) return VbaValue::numeric(affected_);
        throw std::runtime_error("DAO: Database member unsupported");
    }
private:
    QSqlDatabase db_;
    std::weak_ptr<TableDefinitions> tableDefinitions_;
    qlonglong affected_{};
};
}
std::shared_ptr<VbaObject> makeDaoDatabase(QSqlDatabase database) {
    return std::make_shared<Database>(std::move(database));
}
}
