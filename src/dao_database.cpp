#include "dao_database.hpp"
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
    Recordset(QSqlDatabase db, QString source, bool snapshot)
        : db_(std::move(db)), source_(std::move(source)), readOnly_(snapshot) {
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
        QSqlQuery query(db_);
        exec(query, table_ ? "SELECT rowid,* FROM " + quote(source_) : source_);
        const int offset = table_ ? 1 : 0;
        columns_.clear(); rows_.clear(); ids_.clear(); position_ = 0;
        for (int i = offset; i < query.record().count(); ++i) columns_ << query.record().fieldName(i);
        while (query.next()) {
            require(rows_.size() < 100000, "DAO: Recordset row limit exceeded");
            QVariantList row;
            if (table_) ids_.append(query.value(0));
            for (int i = offset; i < query.record().count(); ++i) row << query.value(i);
            rows_.append(row);
        }
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
    bool table_{}, readOnly_{}, closed_{};
    QStringList columns_;
    QVector<QVariantList> rows_;
    QVector<QVariant> ids_;
    QVariantList pending_;
    QSet<int> changed_;
    qsizetype position_{};
    Mode mode_{Mode::Browse};
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
class Database final : public VbaObject {
public:
    explicit Database(QSqlDatabase db) : db_(std::move(db)) {}
    VbaValue member(std::string_view name, const std::vector<VbaValue>& args) override {
        require(db_.isOpen(), "DAO: Database is closed");
        const auto key = lower(name);
        if (key == "openrecordset") {
            require(args.size() == 1 || args.size() == 2, "DAO: OpenRecordset arguments unsupported");
            const int mode = args.size() == 2 ? static_cast<int>(args[1].asNumber()) : 2;
            require(mode == 2 || mode == 4, "DAO: Recordset type unsupported");
            const auto source = text(args[0].asText());
            require(db_.tables().contains(source, Qt::CaseInsensitive)
                || source.trimmed().startsWith("SELECT ", Qt::CaseInsensitive), "DAO: Table or SELECT required");
            return VbaValue::reference(std::make_shared<Recordset>(db_, source, mode == 4));
        }
        if (key == "execute") {
            require(args.size() == 1 || (args.size() == 2 && args[1].asNumber() == 128), "DAO: Execute arguments unsupported");
            QSqlQuery query(db_); exec(query, text(args[0].asText())); return {};
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
