#include "dao_database.hpp"
#include "access_autonumber.hpp"
#include "vba_qt_value.hpp"
#include "access_query.hpp"
#include "jet_sql_functions.hpp"
#include "sqlite_execution.hpp"
#include <QSqlError>
#include <QSqlQuery>
#include <QSqlRecord>
#include <QSqlDriver>
#include <sqlite3.h>
#include <QSet>
#include <QVariant>
#include <QUuid>
#include <QElapsedTimer>
#include <QScopeGuard>
#include <QPointer>
#include <QThread>
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
    if (message.contains("unique constraint") || message.contains("primary key")) throw VbaRuntimeError(3022,"DAO: Duplicate key");
    if (message.contains("not null")) throw VbaRuntimeError(3314,"DAO: Required value missing");
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
            if (!query.prepare(accessActionSql(db,plan))) queryFailure(query);
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
    case VbaValue::Kind::Currency: return text(value.asText());
    case VbaValue::Kind::Boolean: return value.boolean ? -1 : 0;
    case VbaValue::Kind::Text: return text(value.text);
    default: throw std::runtime_error("DAO: Object cannot be stored in a field");
    }
}
struct WorkspaceState {
    QPointer<QSqlDriver> driver;
    sqlite3* connection{};
    bool alive{true};
    quint64 rollbackRevision{};
    int clones{};
    QStringList savepoints;
    QString prefix=QString("la_dao_")+QUuid::createUuid().toString(QUuid::Id128);
    std::weak_ptr<VbaObject> engine,workspace,database;
    void check() const {
        if (!alive || !driver || !driver->isOpen() || !connection) throw VbaRuntimeError(3420,"DAO: Workspace connection unavailable");
        if (driver->thread()!=QThread::currentThread()) throw VbaRuntimeError(3251,"DAO: Cross-thread workspace unsupported");
        const auto handle=driver->handle();
        if (!handle.isValid() || QByteArray(handle.typeName())!="sqlite3*" || *static_cast<sqlite3* const*>(handle.constData())!=connection)
            throw VbaRuntimeError(3420,"DAO: Workspace connection was replaced");
        if (!savepoints.isEmpty() && sqlite3_get_autocommit(connection)) throw VbaRuntimeError(3420,"DAO: Workspace transaction was ended externally");
    }
    void sql(const QString& command) {
        check(); const auto bytes=command.toUtf8();
        const int result=sqlite3_exec(connection,bytes.constData(),nullptr,nullptr,nullptr);
        if (result!=SQLITE_OK) throw VbaRuntimeError(result==SQLITE_BUSY || result==SQLITE_LOCKED ? 3260 : 3146,"DAO: Transaction operation failed");
    }
    void begin() {
        check();
        if (clones) throw VbaRuntimeError(3251,"DAO: Transactions with recordset clones unsupported");
        if (savepoints.size()>=5) throw VbaRuntimeError(3003,"DAO: Too many nested transactions");
        const auto name=prefix+"_"+QString::number(savepoints.size());
        sql("SAVEPOINT "+name); savepoints.append(name);
    }
    void commit() {
        check(); if (savepoints.isEmpty()) throw VbaRuntimeError(3034,"DAO: BeginTrans required");
        sql("RELEASE "+savepoints.last()); savepoints.removeLast();
    }
    void rollback() {
        check(); if (savepoints.isEmpty()) throw VbaRuntimeError(3034,"DAO: BeginTrans required");
        if (savepoints.size()!=1) throw VbaRuntimeError(3251,"DAO: Nested rollback unsupported");
        sql("ROLLBACK TO "+savepoints.first()); ++rollbackRevision;
        sql("RELEASE "+savepoints.first()); savepoints.clear();
    }
};
std::shared_ptr<WorkspaceState> workspaceState(QSqlDatabase db) {
    if (!db.isOpen() || !db.driver()) throw VbaRuntimeError(3420,"DAO: Database is closed");
    const auto handle=db.driver()->handle();
    if (!handle.isValid() || QByteArray(handle.typeName())!="sqlite3*") throw VbaRuntimeError(3251,"DAO: SQLite workspace required");
    auto* native=*static_cast<sqlite3* const*>(handle.constData());
    auto state=db.driver()->property("la_dao_workspace_state").value<std::shared_ptr<WorkspaceState>>();
    if (state && state->alive && state->connection==native) { state->check(); return state; }
    state=std::make_shared<WorkspaceState>(); state->driver=db.driver(); state->connection=native;
    // The SQLite destructor invalidates wrappers even if the Qt driver is reopened.
    auto* token=new std::shared_ptr<WorkspaceState>(state);
    if (sqlite3_create_function_v2(native,"linaccess_dao_connection_generation",0,SQLITE_UTF8|SQLITE_DIRECTONLY,token,
        [](sqlite3_context* context,int,sqlite3_value**) { sqlite3_result_int(context,1); },nullptr,nullptr,
        [](void* data) { auto* holder=static_cast<std::shared_ptr<WorkspaceState>*>(data); (*holder)->alive=false; (*holder)->connection=nullptr; delete holder; })!=SQLITE_OK)
        throw VbaRuntimeError(3146,"DAO: Workspace lifetime tracking unavailable");
    db.driver()->setProperty("la_dao_workspace_state",QVariant::fromValue(state)); return state;
}
class Recordset;
struct RecordsetData {
    std::shared_ptr<WorkspaceState> workspace;
    quint64 rollbackRevision{};
    QStringList columns;
    QMap<QString,QString> types;
    QVector<QVariantList> rows;
    QVector<QVariant> ids;
    QVector<quint64> bookmarks;
    quint64 nextBookmark{1};
    quint64 revision{};
    QByteArray family{QUuid::createUuid().toRfc4122()};
    QSet<Recordset*> cursors;
};
VbaValue fieldCollection(std::shared_ptr<Recordset> owner);
class Field final : public VbaObject {
public:
    Field(std::shared_ptr<Recordset> owner, QString name) : owner_(std::move(owner)), name_(std::move(name)) {}
    bool hasScalarDefault() const override { return true; }
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
        data_->workspace=workspaceState(db_); data_->rollbackRevision=data_->workspace->rollbackRevision;
        table_ = db_.tables().contains(source_, Qt::CaseInsensitive);
        readOnly_ |= !table_;
        reload();
        data_->cursors.insert(this);
    }
    Recordset(QSqlDatabase db,QStringList columns,QMap<QString,QString> types,QVector<QVariantList> rows,QVector<QVariant> ids)
        : db_(std::move(db)),readOnly_(true),frozen_(true) {
        data_->workspace=workspaceState(db_); data_->rollbackRevision=data_->workspace->rollbackRevision;
        data_->columns=std::move(columns); data_->types=std::move(types); data_->rows=std::move(rows); data_->ids=std::move(ids);
        for (qsizetype i=0;i<data_->rows.size();++i) data_->bookmarks.append(data_->nextBookmark++);
        data_->cursors.insert(this);
    }
    ~Recordset() override { if (cloneCounted_) --data_->workspace->clones; data_->cursors.remove(this); }
    void checkTransaction() const {
        data_->workspace->check();
        if (data_->rollbackRevision!=data_->workspace->rollbackRevision) throw VbaRuntimeError(3420,"DAO: Recordset requires Requery after Rollback");
    }
    void refreshSnapshot(const QStringList& columns,const QMap<QString,QString>& types,
                         const QVector<QVariantList>& rows,const QVector<QVariant>& ids) {
        data_->workspace->check();
        if (!frozen_ || columns!=data_->columns || ids.size()!=rows.size() || data_->ids.size()!=data_->rows.size())
            throw VbaRuntimeError(3251,"DAO: Incompatible form snapshot refresh");
        if (data_->rollbackRevision!=data_->workspace->rollbackRevision) { data_->family=QUuid::createUuid().toRfc4122(); data_->rollbackRevision=data_->workspace->rollbackRevision; }
        QMap<qlonglong,quint64> tokens;
        QMap<Recordset*,QVariant> positions;
        for (qsizetype i=0;i<data_->ids.size();++i) tokens.insert(data_->ids[i].toLongLong(),data_->bookmarks[i]);
        for (auto* cursor:data_->cursors)
            if (cursor->position_>=0 && cursor->position_<data_->ids.size()) positions.insert(cursor,data_->ids[cursor->position_]);
        QVector<quint64> bookmarks;
        for (const auto& id:ids) {
            const auto previous=tokens.value(id.toLongLong());
            bookmarks.append(previous ? previous : data_->nextBookmark++);
        }
        data_->types=types; data_->rows=rows; data_->ids=ids; data_->bookmarks=std::move(bookmarks);
        ++data_->revision;
        for (auto* cursor:data_->cursors) {
            if (positions.contains(cursor)) cursor->position_=data_->ids.indexOf(positions.value(cursor));
            else if (cursor->position_>=0) cursor->position_=data_->rows.size();
        }
    }
    VbaValue bookmarkAt(qsizetype position) const {
        checkTransaction();
        if (closed_ || position<0 || position>=data_->rows.size()) throw VbaRuntimeError(3021,"DAO: No current record");
        std::vector<unsigned char> bytes(data_->family.begin(),data_->family.end());
        const auto id=data_->bookmarks[position];
        for (int shift=56;shift>=0;shift-=8) bytes.push_back(static_cast<unsigned char>(id>>shift));
        return VbaValue::byteArray(std::move(bytes));
    }
    qsizetype bookmarkPosition(const VbaValue& value) const {
        checkTransaction();
        if (value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Invalid Null bookmark");
        if (value.kind!=VbaValue::Kind::ByteArray) throw VbaRuntimeError(13,"DAO: Bookmark requires a byte array");
        if (value.bytes.size()!=24 || !std::equal(data_->family.begin(),data_->family.end(),value.bytes.begin(),
                [](char a,unsigned char b) { return static_cast<unsigned char>(a)==b; }))
            throw VbaRuntimeError(3159,"DAO: Invalid bookmark");
        quint64 id=0; for (std::size_t i=16;i<24;++i) id=(id<<8)|value.bytes[i];
        const auto position=data_->bookmarks.indexOf(id);
        if (position<0) throw VbaRuntimeError(3159,"DAO: Invalid bookmark");
        return position;
    }
    VbaValue defaultMember(const std::vector<VbaValue>& args,bool reference) override {
        if (args.size()!=1) throw VbaRuntimeError(450,"DAO: Recordset default field requires one index");
        const auto field=member("Fields",args); return reference ? field : field.object->member("Value",{});
    }
    VbaValue item(std::string_view name) override { return value(text(name)); }
    VbaValue objectItem(std::string_view name) override { return member("Fields",{VbaValue::string(std::string(name))}); }
    void assignItem(std::string_view name,const VbaValue& value) override { set(text(name),value); }
    QStringList fieldNames() const { checkTransaction(); if (closed_) throw VbaRuntimeError(3420,"DAO: Recordset is closed"); return data_->columns; }
    VbaValue member(std::string_view name, const std::vector<VbaValue>& args) override {
        if (closed_) throw VbaRuntimeError(3420,"DAO: Recordset is closed");
        const auto key = lower(name);
        if (key!="close" && key!="requery") checkTransaction();
        if (key=="transactions" && args.empty()) return VbaValue::logical(true);
        if (key=="findfirst" || key=="findlast" || key=="findnext" || key=="findprevious") {
            if (args.size()!=1) throw VbaRuntimeError(450,"DAO: Find requires one criterion");
            if (args[0].kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Invalid Null criterion");
            find(key,text(args[0].asText())); return {};
        }
        if (key=="clone") {
            if (!data_->workspace->savepoints.isEmpty()) throw VbaRuntimeError(3251,"DAO: Clone inside transaction unsupported");
            if (!args.empty()) throw VbaRuntimeError(450,"DAO: Clone takes no arguments");
            if (mode_!=Mode::Browse) throw VbaRuntimeError(3219,"DAO: Pending edit");
            return VbaValue::reference(std::shared_ptr<Recordset>(new Recordset(*this,CloneTag{})));
        }
        if (key == "fields") {
            if (args.empty()) return fieldCollection(shared_from_this());
            require(args.size() == 1, "DAO: Fields requires one field name/index");
            QString field;
            if (args[0].kind == VbaValue::Kind::Number) {
                const auto index = sqlFieldValue(args[0],"INTEGER").toLongLong();
                if (index < 0 || index >= data_->columns.size()) throw VbaRuntimeError(3265,"DAO: Field index out of range");
                field = data_->columns[static_cast<int>(index)];
            } else field = text(args[0].asText());
            (void)column(field);
            return VbaValue::reference(std::make_shared<Field>(shared_from_this(), field));
        }
        require(args.empty(), "DAO: Unexpected Recordset arguments");
        if (key=="bookmark") return bookmarkAt(position_);
        if (key=="bookmarkable") return VbaValue::logical(true);
        if (key=="nomatch") return VbaValue::logical(noMatch_);
        if (key == "eof") return VbaValue::logical(data_->rows.empty() || position_ >= data_->rows.size());
        if (key == "bof") return VbaValue::logical(data_->rows.empty() || position_ < 0);
        if (key == "recordcount") return VbaValue::numeric(data_->rows.size());
        if (key == "movefirst" || key == "movelast" || key == "movenext" || key == "moveprevious") {
            require(mode_ == Mode::Browse, "DAO: Update or CancelUpdate before navigation");
            if (key == "movefirst") position_ = 0;
            if (key == "movelast") position_ = data_->rows.empty() ? 0 : data_->rows.size() - 1;
            if (key == "movenext") { require(position_ < data_->rows.size(), "DAO: Already at EOF"); ++position_; }
            if (key == "moveprevious") { require(position_ >= 0 && !data_->rows.empty(), "DAO: Already at BOF"); --position_; }
            noMatch_=false;
            return {};
        }
        if (key == "edit" || key == "addnew") {
            require(!readOnly_, "DAO: Recordset is read-only");
            require(mode_ == Mode::Browse, "DAO: Edit already in progress");
            if (key == "edit") { current(); pending_ = data_->rows[position_]; mode_ = Mode::Edit; }
            else { pending_ = QVariantList(data_->columns.size()); mode_ = Mode::Add; }
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
            query.addBindValue(data_->ids[position_]);
            require(query.exec() && query.numRowsAffected() == 1, "DAO: Delete failed or record changed");
            reload(false);
            return {};
        }
        if (key == "requery") {
            data_->workspace->check();
            if (data_->rollbackRevision!=data_->workspace->rollbackRevision) { mode_=Mode::Browse; pending_.clear(); changed_.clear(); }
            require(mode_ == Mode::Browse, "DAO: Pending edit");
            reload(); return {};
        }
        if (key == "close") { if (cloneCounted_) { --data_->workspace->clones; cloneCounted_=false; } closed_ = true; data_->cursors.remove(this); pending_.clear(); changed_.clear(); return {}; }
        return value(text(name)); // implicit default Fields(name).Value
    }
    void assign(std::string_view name, const VbaValue& value) override {
        if (lower(name)=="bookmark") {
            if (closed_) throw VbaRuntimeError(3420,"DAO: Recordset is closed");
            if (mode_!=Mode::Browse) throw VbaRuntimeError(3219,"DAO: Pending edit");
            position_=bookmarkPosition(value); noMatch_=false; return;
        }
        set(text(name), value);
    }
    QString fieldType(const QString& name) const {
        checkTransaction();
        if(closed_) throw VbaRuntimeError(3420,"DAO: Recordset is closed");
        (void)column(name); return data_->types.value(name.toCaseFolded());
    }
    VbaValue value(const QString& name) const {
        checkTransaction();
        if (closed_) throw VbaRuntimeError(3420,"DAO: Recordset is closed");
        const int index = column(name);
        if (mode_ != Mode::Browse) return qtFieldValue(pending_[index],data_->types.value(name.toCaseFolded()));
        current(); return qtFieldValue(data_->rows[position_][index],data_->types.value(name.toCaseFolded()));
    }
    void set(const QString& name, const VbaValue& value) {
        checkTransaction();
        if (closed_) throw VbaRuntimeError(3420,"DAO: Recordset is closed");
        require(mode_ != Mode::Browse, "DAO: Edit or AddNew required");
        const auto index = column(name);
        pending_[index] = sqlFieldValue(value,data_->types.value(name.toCaseFolded())); changed_.insert(index);
    }
private:
    struct CloneTag {};
    Recordset(const Recordset& source,CloneTag)
        : data_(source.data_),db_(source.db_),source_(source.source_),bindings_(source.bindings_),
          table_(source.table_),readOnly_(source.readOnly_),frozen_(source.frozen_),position_(-1),queryPlan_(source.queryPlan_) {
        cloneCounted_=true; ++data_->workspace->clones;
        data_->cursors.insert(this);
    }
    enum class Mode { Browse, Edit, Add };
    int column(const QString& name) const {
        for (int i = 0; i < data_->columns.size(); ++i)
            if (data_->columns[i].compare(name, Qt::CaseInsensitive) == 0) return i;
        throw std::runtime_error("DAO: Unknown field");
    }
    void current() const { if (position_<0 || position_>=data_->rows.size()) throw VbaRuntimeError(3021,"DAO: No current record"); }
    void find(const std::string& kind,const QString& criteria) {
        if (mode_!=Mode::Browse) throw VbaRuntimeError(3219,"DAO: Pending edit");
        if (criteria.trimmed().isEmpty()) throw VbaRuntimeError(3075,"DAO: Empty find criterion");
        auto plan=compileAccessQuery("SELECT IIf(("+criteria+"),1,0) FROM [__la_find_scope]");
        if (!plan.parameters.empty()) throw VbaRuntimeError(3061,"DAO: Unbound find parameter");
        QStringList projections;
        for (const auto& column:data_->columns) projections.append("? AS "+quote(column));
        const auto from=plan.sql.lastIndexOf(" FROM ",Qt::CaseInsensitive);
        if (from<0) throw VbaRuntimeError(3075,"DAO: Find criterion grammar unsupported");
        SqliteExecutionBudget budget(db_); QElapsedTimer clock; clock.start();
        const auto handle=db_.driver()->handle();
        auto* connection=*static_cast<sqlite3* const*>(handle.constData());
        int previousDqs=0;
        if (sqlite3_db_config(connection,SQLITE_DBCONFIG_DQS_DML,-1,&previousDqs)!=SQLITE_OK
            || sqlite3_db_config(connection,SQLITE_DBCONFIG_DQS_DML,0,nullptr)!=SQLITE_OK)
            throw VbaRuntimeError(3251,"DAO: Strict find preparation unavailable");
        const auto restoreDqs=qScopeGuard([&] { sqlite3_db_config(connection,SQLITE_DBCONFIG_DQS_DML,previousDqs,nullptr); });
        QSqlQuery query(db_);
        const auto fail=[&] {
            if (query.lastError().text().contains("no such column",Qt::CaseInsensitive)) throw VbaRuntimeError(3070,"DAO: Find field unavailable");
            queryFailure(query);
        };
        if (!query.prepare(plan.sql.left(from)+" FROM (SELECT "+projections.join(',')+")")) fail();
        const bool backwards=kind=="findlast" || kind=="findprevious";
        const auto revision=data_->revision;
        const qsizetype step=backwards ? -1 : 1;
        qsizetype candidate=kind=="findfirst" ? 0 : kind=="findlast" ? data_->rows.size()-1 : position_+step;
        for (;candidate>=0 && candidate<data_->rows.size();candidate+=step) {
            if (clock.hasExpired(5000)) throw VbaRuntimeError(3035,"DAO: Find execution budget exceeded");
            const auto& row=data_->rows[candidate];
            for (qsizetype column=0;column<row.size();++column) query.bindValue(static_cast<int>(column),row[column]);
            const bool fetched=query.exec() && query.next();
            if (data_->revision!=revision) throw VbaRuntimeError(3219,"DAO: Recordset changed during find");
            if (!fetched) fail();
            const bool match=query.value(0).toInt()!=0; query.finish();
            if (match) { position_=candidate; noMatch_=false; return; }
        }
        position_=backwards ? -1 : data_->rows.size(); noMatch_=true;
    }
    void reload(bool invalidate=true) {
        if (frozen_) throw VbaRuntimeError(3251,"DAO: Form snapshot requires form Requery");
        const auto original=*data_;
        std::vector<std::shared_ptr<Recordset>> leases;
        for (auto* cursor:data_->cursors)
            if (auto lease=cursor->weak_from_this().lock()) leases.push_back(std::move(lease));
        QMap<Recordset*,QPair<qsizetype,bool>> originalPositions;
        for (auto* cursor:data_->cursors) originalPositions.insert(cursor,{cursor->position_,cursor->noMatch_});
        bool completed=false;
        const auto restore=qScopeGuard([&] {
            if (completed) return;
            const auto live=data_->cursors; *data_=original; data_->cursors=live;
            for (auto it=originalPositions.cbegin();it!=originalPositions.cend();++it) {
                if (live.contains(it.key())) { it.key()->position_=it.value().first; it.key()->noMatch_=it.value().second; }
            }
        });
        QMap<qlonglong,quint64> previousTokens;
        QMap<Recordset*,QVariant> previousRows;
        if (!invalidate && table_) {
            for (qsizetype i=0;i<data_->ids.size();++i) previousTokens.insert(data_->ids[i].toLongLong(),data_->bookmarks[i]);
            for (auto* cursor:data_->cursors)
                if (cursor->position_>=0 && cursor->position_<data_->ids.size()) previousRows.insert(cursor,data_->ids[cursor->position_]);
        }
        const auto finalize=[&] {
            if (data_->rollbackRevision!=data_->workspace->rollbackRevision) data_->family=QUuid::createUuid().toRfc4122();
            data_->rollbackRevision=data_->workspace->rollbackRevision;
            data_->bookmarks.clear();
            for (qsizetype i=0;i<data_->rows.size();++i) {
                const auto old=table_ ? previousTokens.value(data_->ids[i].toLongLong()) : 0;
                data_->bookmarks.append(old ? old : data_->nextBookmark++);
            }
            for (auto* cursor:data_->cursors) {
                if (invalidate) cursor->position_=-1;
                else if (previousRows.contains(cursor)) cursor->position_=data_->ids.indexOf(previousRows.value(cursor));
                else if (cursor->position_>=0) cursor->position_=data_->rows.size();
                cursor->noMatch_=false;
            }
            position_=0; noMatch_=false;
            ++data_->revision;
            completed=true;
        };
        data_->types = table_ ? sqliteColumnTypes(db_,source_) : QMap<QString,QString>{};
        if (!table_) {
            auto plan=queryPlan_;
            if (plan.sql.isEmpty()) plan.sql=source_;
            const auto result = executeAccessQueryPlan(db_,plan,bindings_);
            data_->columns = result.columns; data_->types=result.types; data_->rows = result.rows; data_->ids.clear(); finalize();
            return;
        }
        SqliteExecutionBudget budget(db_);
        QSqlQuery query(db_);
        exec(query,"SELECT rowid,* FROM " + quote(source_));
        const int offset = table_ ? 1 : 0;
        data_->columns.clear(); data_->rows.clear(); data_->ids.clear(); position_ = 0;
        const int end = query.record().count();
        for (int i = offset; i < end; ++i) data_->columns << query.record().fieldName(i);
        while (query.next()) {
            require(data_->rows.size() < 100000, "DAO: Recordset row limit exceeded");
            QVariantList row;
            if (table_) data_->ids.append(query.value(0));
            for (int i = offset; i < end; ++i) row << query.value(i);
            data_->rows.append(row);
        }
        if (query.lastError().isValid()) queryFailure(query);
        finalize();
    }
    void update() {
        require(!closed_ && !readOnly_ && mode_ != Mode::Browse, "DAO: No editable record buffer");
        const bool adding = mode_ == Mode::Add;
        SqliteExecutionBudget budget(db_);
        const auto generated=adding ? generatedAccessColumns(db_,source_) : QSet<QString>{};
        QVariant id = adding ? QVariant{} : data_->ids[position_];
        QStringList columns, values, assignments;
        QList<int> indices = changed_.values();
        if (adding) indices.removeIf([&](int index) { return pending_[index].isNull() && generated.contains(data_->columns[index].toCaseFolded()); });
        std::sort(indices.begin(), indices.end());
        for (const int i : indices) {
            columns << quote(data_->columns[i]); values << "?"; assignments << quote(data_->columns[i]) + "=?";
        }
        QSqlQuery query(db_);
        if (adding && columns.empty()) { if (!query.exec("INSERT INTO " + quote(source_) + " DEFAULT VALUES")) queryFailure(query); }
        else if (!columns.empty()) {
            require(query.prepare(adding ? "INSERT INTO " + quote(source_) + " (" + columns.join(',')
                        + ") VALUES (" + values.join(',') + ")"
                    : "UPDATE " + quote(source_) + " SET " + assignments.join(',') + " WHERE rowid=?"),
                    "DAO: Update prepare failed");
            for (const int i : indices) query.addBindValue(pending_[i]);
            if (!adding) query.addBindValue(id);
            if (!query.exec()) queryFailure(query);
            require(query.numRowsAffected() == 1, "DAO: Update failed or record changed");
        }
        if (adding) id = query.lastInsertId();
        mode_ = Mode::Browse; pending_.clear(); changed_.clear();
        reload(false); position_ = data_->ids.indexOf(id);
    }
    std::shared_ptr<RecordsetData> data_{std::make_shared<RecordsetData>()};
    QSqlDatabase db_;
    QString source_;
    QMap<QString,QVariant> bindings_;
    bool table_{}, readOnly_{}, closed_{}, frozen_{},noMatch_{},cloneCounted_{};
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
    if(lower(name)=="type") {
        if(!args.empty()) throw VbaRuntimeError(450,"Field.Type requires no arguments");
        const auto type=owner_->fieldType(name_).toUpper();
        const int code=type=="INTEGER1" ? 2 : type=="INTEGER2" ? 3 : type=="INTEGER4" ? 4 : type=="CURRENCY_REAL" ? 5 : 0;
        if(!code) throw VbaRuntimeError(3251,"Field.Type is not available for this source type");
        return VbaValue::integer(static_cast<std::int16_t>(code));
    }
    require(args.empty(), "DAO: Unexpected field arguments");
    if (lower(name) == "value") return owner_->value(name_);
    if (lower(name) == "name") return VbaValue::string(name_.toUtf8().toStdString());
    throw std::runtime_error("DAO: Unknown Field member");
}
void Field::assign(std::string_view name, const VbaValue& value) {
    if(lower(name)=="type") { (void)owner_->fieldType(name_); throw VbaRuntimeError(383,"Field.Type is read-only"); }
    require(lower(name) == "value", "DAO: Field property is read-only");
    owner_->set(name_, value);
}
class QueryDefinition;
class QueryParameter final : public VbaObject {
public:
    QueryParameter(std::shared_ptr<QueryDefinition> query, QString name,quint64 generation) : query_(std::move(query)), name_(std::move(name)),generation_(generation) {}
    bool hasScalarDefault() const override { return true; }
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
        checkParameter(name);
        const auto stored=values_.value(name.toCaseFolded());
        for(const auto& parameter:plan_.parameters)
            if(parameter.type=="currency" && parameter.name.compare(name,Qt::CaseInsensitive)==0)
                return qtFieldValue(stored,"CURRENCY_REAL");
        return fromSql(stored);
    }
    void setParameter(const QString& name, const VbaValue& value) {
        checkParameter(name);
        for(const auto& parameter:plan_.parameters)
            if(parameter.type=="currency" && parameter.name.compare(name,Qt::CaseInsensitive)==0) {
                const auto converted=sqlFieldValue(value,"CURRENCY_REAL");
                values_.insert(name.toCaseFolded(),converted); return;
            }
        values_.insert(name.toCaseFolded(),toSql(value));
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
    if (lower(name)=="type") {
        static const QMap<QString,int> types{{"boolean",1},{"yesno",1},{"byte",2},
            {"short",3},{"integer",4},{"long",4},{"currency",5},{"single",6},{"double",7},
            {"date",8},{"datetime",8},{"text",10},{"varchar",10},{"char",10}};
        for (const auto& parameter:query_->plan_.parameters) if (parameter.name.compare(name_,Qt::CaseInsensitive)==0) {
            const auto type=types.constFind(parameter.type);
            if (type==types.cend()) throw VbaRuntimeError(3251,"DAO: inferred parameter type is not validated");
            return VbaValue::integer(static_cast<std::int16_t>(type.value()));
        }
        throw VbaRuntimeError(3265,"DAO: Parameter not found");
    }
    throw VbaRuntimeError(438,"DAO: Parameter member unsupported");
}
void QueryParameter::assign(std::string_view name, const VbaValue& value) {
    query_->parameterGeneration(generation_);
    if (lower(name)=="type") throw VbaRuntimeError(383,"DAO: Parameter Type is read-only in an Access workspace");
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
    explicit Database(QSqlDatabase db) : db_(std::move(db)),workspace_(workspaceState(db_)) {}
    VbaValue member(std::string_view name, const std::vector<VbaValue>& args) override {
        workspace_->check(); require(db_.isOpen(), "DAO: Database is closed");
        const auto key = lower(name);
        if (key=="transactions" && args.empty()) return VbaValue::logical(true);
        if (key=="name" && args.empty()) return VbaValue::string(db_.databaseName().toUtf8().toStdString());
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
    std::shared_ptr<WorkspaceState> workspace_;
    qlonglong affected_{};
};
class SingletonDaoCollection final : public VbaObject {
public:
    SingletonDaoCollection(std::shared_ptr<VbaObject> object,QString name):object_(std::move(object)),name_(std::move(name)) {}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        const auto key=lower(name);
        if (key=="count" && args.empty()) { (void)object_->member("Name",{}); return VbaValue::numeric(1); }
        if (key=="item") return defaultMember(args,true);
        throw VbaRuntimeError(3251,"DAO: Collection member unsupported");
    }
    VbaValue defaultMember(const std::vector<VbaValue>& args,bool) override {
        if (args.size()!=1) throw VbaRuntimeError(450,"DAO: Collection requires one index");
        const auto value=scalarVbaValue(args[0]);
        if (value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Invalid Null index");
        if (value.kind==VbaValue::Kind::Text) {
            if (text(value.asText()).compare(name_,Qt::CaseInsensitive)!=0) throw VbaRuntimeError(3265,"DAO: Collection item unavailable");
        } else if (sqlFieldValue(value,"INTEGER").toLongLong()!=0) throw VbaRuntimeError(3265,"DAO: Collection item unavailable");
        // Validate the underlying connection before returning an item.
        (void)object_->member("Name",{}); return VbaValue::reference(object_);
    }
    VbaValue item(std::string_view name) override { return defaultMember({VbaValue::string(std::string(name))},true); }
    std::vector<VbaValue> enumerate() override { (void)object_->member("Name",{}); return {VbaValue::reference(object_)}; }
private:
    std::shared_ptr<VbaObject> object_;
    QString name_;
};
class Workspace final : public VbaObject {
public:
    explicit Workspace(QSqlDatabase db):db_(std::move(db)),state_(workspaceState(db_)) {}
    VbaValue defaultMember(const std::vector<VbaValue>& args,bool) override { return databases()->defaultMember(args,true); }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        state_->check(); const auto key=lower(name);
        if (key=="name" && args.empty()) return VbaValue::string("#Default Workspace#");
        if (key=="databases") { auto collection=databases(); return args.empty() ? VbaValue::reference(collection) : collection->defaultMember(args,true); }
        if (key=="begintrans" || key=="rollback") {
            if (!args.empty()) throw VbaRuntimeError(450,"DAO: Transaction method takes no arguments");
            if (key=="begintrans") state_->begin(); else state_->rollback(); return {};
        }
        if (key=="committrans") {
            if (args.size()>1) throw VbaRuntimeError(450,"DAO: CommitTrans arguments unsupported");
            if (!args.empty()) {
                const auto flags=sqlFieldValue(scalarVbaValue(args[0]),"INTEGER");
                if (args[0].kind==VbaValue::Kind::Null || flags.isNull()) throw VbaRuntimeError(94,"DAO: Invalid Null commit flags");
                if (flags.toLongLong()==1) throw VbaRuntimeError(3251,"DAO: Forced OS flush unsupported");
                if (flags.toLongLong()!=0) throw VbaRuntimeError(5,"DAO: Invalid commit flags");
            }
            state_->commit(); return {};
        }
        throw VbaRuntimeError(3251,"DAO: Workspace member unsupported");
    }
private:
    std::shared_ptr<VbaObject> databases() {
        state_->check(); auto value=databases_.lock();
        if (!value) {
            auto database=state_->database.lock();
            if (!database) { database=makeDaoDatabase(db_); state_->database=database; }
            value=std::make_shared<SingletonDaoCollection>(database,db_.databaseName()); databases_=value;
        }
        return value;
    }
    QSqlDatabase db_;
    std::shared_ptr<WorkspaceState> state_;
    std::weak_ptr<VbaObject> databases_;
};
class Engine final : public VbaObject {
public:
    explicit Engine(QSqlDatabase db):db_(std::move(db)),state_(workspaceState(db_)) {}
    VbaValue defaultMember(const std::vector<VbaValue>& args,bool) override { return workspaces()->defaultMember(args,true); }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        state_->check(); const auto key=lower(name);
        if (key=="workspaces") { auto value=workspaces(); return args.empty() ? VbaValue::reference(value) : value->defaultMember(args,true); }
        if (key=="begintrans" || key=="committrans" || key=="rollback") return workspace()->member(name,args);
        throw VbaRuntimeError(3251,"DAO: DBEngine member unsupported");
    }
private:
    std::shared_ptr<VbaObject> workspace() {
        auto value=state_->workspace.lock();
        if (!value) { value=std::make_shared<Workspace>(db_); state_->workspace=value; }
        return value;
    }
    std::shared_ptr<VbaObject> workspaces() {
        state_->check(); auto value=workspaces_.lock();
        if (!value) { value=std::make_shared<SingletonDaoCollection>(workspace(),"#Default Workspace#"); workspaces_=value; }
        return value;
    }
    QSqlDatabase db_;
    std::shared_ptr<WorkspaceState> state_;
    std::weak_ptr<VbaObject> workspaces_;
};
}
std::shared_ptr<VbaObject> makeDaoEngine(QSqlDatabase database) {
    const auto state=workspaceState(database); auto value=state->engine.lock();
    if (!value) { value=std::make_shared<Engine>(std::move(database)); state->engine=value; }
    return value;
}
std::shared_ptr<VbaObject> makeDaoDatabase(QSqlDatabase database) {
    return std::make_shared<Database>(std::move(database));
}
std::shared_ptr<VbaObject> makeDaoSnapshot(QSqlDatabase database,const QStringList& columns,
                                        const QMap<QString,QString>& types,const QVector<QVariantList>& rows,const QVector<QVariant>& ids) {
    return std::make_shared<Recordset>(std::move(database),columns,types,rows,ids);
}
void refreshDaoSnapshot(const std::shared_ptr<VbaObject>& recordset,const QStringList& columns,
                        const QMap<QString,QString>& types,const QVector<QVariantList>& rows,const QVector<QVariant>& ids) {
    const auto native=std::dynamic_pointer_cast<Recordset>(recordset);
    if (!native) throw VbaRuntimeError(13,"DAO: Native recordset required");
    native->refreshSnapshot(columns,types,rows,ids);
}
VbaValue daoBookmarkAt(const std::shared_ptr<VbaObject>& recordset,qsizetype position) {
    const auto native=std::dynamic_pointer_cast<Recordset>(recordset);
    if (!native) throw VbaRuntimeError(13,"DAO: Native recordset required");
    return native->bookmarkAt(position);
}
qsizetype daoBookmarkPosition(const std::shared_ptr<VbaObject>& recordset,const VbaValue& bookmark) {
    const auto native=std::dynamic_pointer_cast<Recordset>(recordset);
    if (!native) throw VbaRuntimeError(13,"DAO: Native recordset required");
    return native->bookmarkPosition(bookmark);
}
}
