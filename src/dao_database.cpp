#include "dao_database.hpp"
#include "vba_array.hpp"
#include "access_autonumber.hpp"
#include "access_table_import.hpp"
#include "vba_qt_value.hpp"
#include "access_query.hpp"
#include "jet_sql_functions.hpp"
#include "sqlite_execution.hpp"
#include <QSqlError>
#include <QJsonDocument>
#include <QJsonArray>
#include <QJsonObject>
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
#include <limits>

namespace linaccess {
namespace {
QString text(std::string_view value) { return QString::fromUtf8(value.data(), value.size()); }
QString quote(QString name) { return '"' + name.replace('"', "\"\"") + '"'; }
std::string lower(std::string_view value) { return text(value).toCaseFolded().toStdString(); }
VbaValue longCount(qlonglong value) {
    if(value<std::numeric_limits<std::int32_t>::min() || value>std::numeric_limits<std::int32_t>::max()) throw VbaRuntimeError(6,"DAO count exceeds Long range");
    return VbaValue::longInteger(static_cast<std::int32_t>(value));
}
VbaValue integerCount(qlonglong value) {
    if(value<std::numeric_limits<std::int16_t>::min() || value>std::numeric_limits<std::int16_t>::max()) throw VbaRuntimeError(6,"DAO count exceeds Integer range");
    return VbaValue::integer(static_cast<std::int16_t>(value));
}
int provenDaoFieldType(const QString& declaration) {
    const auto type=declaration.toUpper();
    return type=="INTEGER1" ? 2 : type=="INTEGER2" ? 3 : type=="INTEGER4" ? 4 : type=="CURRENCY_REAL" ? 5 : (type=="DATE" || type=="DATETIME") ? 8 : type=="BLOB" ? 11 : 0;
}
VbaValue fixedDaoFieldSize(int type) {
    const int size=(type==1 || type==2) ? 1 : type==3 ? 2 : (type==4 || type==6) ? 4
        : (type==5 || type==7 || type==8) ? 8 : (type==11 || type==12) ? 0 : type==15 ? 16 : -1;
    if(size<0) throw VbaRuntimeError(3251,"DAO field size is not proven for this type");
    return VbaValue::longInteger(size);
}
void require(bool valid, const char* message) { if (!valid) throw std::runtime_error(message); }
void exec(QSqlQuery& query, const QString& sql) {
    if (!query.exec(sql)) throw std::runtime_error("DAO: SQL operation failed");
}
[[noreturn]] void queryFailure(const QSqlQuery& query) {
    const auto message = query.lastError().text().toCaseFolded();
    if (const auto code=jetSqlFunctionErrorNumber(message.toUtf8().toStdString()))
        throw VbaRuntimeError(*code,jetSqlFunctionErrorExplanation(message.toUtf8().toStdString()));
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
    const auto statement=plan.sql.trimmed();
    const bool schemaAction=statement.startsWith("CREATE ",Qt::CaseInsensitive) || statement.startsWith("ALTER ",Qt::CaseInsensitive) || statement.startsWith("DROP ",Qt::CaseInsensitive);
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
            affected=schemaAction ? 0 : std::max<qlonglong>(0,query.numRowsAffected()); query.finish();
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
    case VbaValue::Kind::Date: return sqlFieldValue(value,"DATETIME");
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
    QMap<QString,quint64> queryGenerations;
    QVariantMap queryViewSnapshot;
    static constexpr const char* queryViewProperties[]{"la_query_database","la_query_views","la_query_internal_views","la_query_rank_pending","la_query_revision"};
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
        sql("SAVEPOINT "+name);
        if(savepoints.isEmpty()) for(const auto property:queryViewProperties)
            queryViewSnapshot.insert(QString::fromLatin1(property),driver->property(property));
        savepoints.append(name);
    }
    void commit() {
        check(); if (savepoints.isEmpty()) throw VbaRuntimeError(3034,"DAO: BeginTrans required");
        sql("RELEASE "+savepoints.last()); savepoints.removeLast();
        if(savepoints.isEmpty()) queryViewSnapshot.clear();
    }
    void rollback() {
        check(); if (savepoints.isEmpty()) throw VbaRuntimeError(3034,"DAO: BeginTrans required");
        if (savepoints.size()!=1) throw VbaRuntimeError(3251,"DAO: Nested rollback unsupported");
        sql("ROLLBACK TO "+savepoints.first()); ++rollbackRevision;
        sql("RELEASE "+savepoints.first()); savepoints.clear();
        for(const auto property:queryViewProperties) driver->setProperty(property,queryViewSnapshot.value(QString::fromLatin1(property)));
        queryViewSnapshot.clear();
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
QString sqliteLiteralDefaultSql(const QVariant& declaration) {
    if(declaration.isNull()) return {};
    const auto token=declaration.toString().trimmed(); const auto literal=accessLiteralDefaultSql(token);
    // Raw SQLite TRUE is1; imported Access TRUE is already normalized to-1.
    return token.compare("TRUE",Qt::CaseInsensitive)==0 ? QString("1") : literal.value_or(QString{});
}
struct TableSchemaColumn { QString name,type; int sourceType{},sourceSize{}; QString literalDefault; };
QVector<TableSchemaColumn> readTableSchemaColumns(QSqlDatabase db,const QString& tableName) {
    QSqlQuery shadow(db);
    shadow.prepare("SELECT 1 FROM temp.sqlite_schema WHERE type IN ('table','view') AND (name=? COLLATE NOCASE OR name='_la_table_schema' COLLATE NOCASE)");
    shadow.addBindValue(tableName);
    if(!shadow.exec()) queryFailure(shadow);
    if(shadow.next()) throw VbaRuntimeError(3251,"DAO schema metadata has a TEMP shadow");
    shadow.finish();
    const auto types=sqliteColumnTypes(db,tableName);
    QSqlQuery query(db);
    if(!query.exec("PRAGMA main.table_info("+quote(tableName)+")")) queryFailure(query);
    QVector<TableSchemaColumn> columns;
    while(query.next()) {
        const auto name=query.value(1).toString();
        columns.push_back({name,types.value(name.toCaseFolded()),0,0,sqliteLiteralDefaultSql(query.value(4))});
    }
    if(query.lastError().isValid()) queryFailure(query);
    if(columns.isEmpty()) throw VbaRuntimeError(3420,"DAO main-schema table unavailable");
    query.finish();
    if(!query.exec("SELECT 1 FROM main.sqlite_schema WHERE type='table' AND name='_la_table_schema' COLLATE NOCASE")) queryFailure(query);
    const bool metadata=query.next(); query.finish();
    if(metadata) {
        query.prepare("SELECT definition FROM main._la_table_schema WHERE table_name=? COLLATE NOCASE"); query.addBindValue(tableName);
        if(!query.exec()) queryFailure(query);
        if(query.next()) {
            const auto document=QJsonDocument::fromJson(query.value(0).toByteArray());
            if(!document.isArray()) throw VbaRuntimeError(3251,"DAO source schema metadata invalid");
            if(query.next()) throw VbaRuntimeError(5,"DAO source schema metadata ambiguous");
            QSet<QString> seen;
            for(const auto& entry:document.array()) {
                const auto object=entry.toObject();
                if(!object.value("name").isString() || (object.contains("source_type") && !object.value("source_type").isString())) throw VbaRuntimeError(3251,"DAO source field metadata invalid");
                const auto name=object.value("name").toString().toCaseFolded();
                if(seen.contains(name)) throw VbaRuntimeError(5,"DAO source field metadata ambiguous"); seen.insert(name);
                const auto source=object.value("source_type").toString().toUpper();
                const auto origin=object.value("source_origin").toString();
                const auto size=object.value("source_size").toDouble();
                if(object.contains("source_size") && (!object.value("source_size").isDouble() || size<0 || size>255 || std::floor(size)!=size))
                    throw VbaRuntimeError(3251,"DAO source field capacity invalid");
                const int code=!object.contains("source_type") ? 0 : (source=="BOOLEAN" || source=="BIT") ? 1 : source=="TINYINT" ? 2 : source=="SMALLINT" ? 3 : source=="INT" ? 4
                    : source=="CURRENCY" ? 5 : source=="SINGLE" ? 6 : source=="DOUBLE" ? 7 : (source=="DATE" || source=="DATETIME") ? 8
                    : source=="VARCHAR" ? 10 : source=="BLOB" ? 11 : (source=="MEMO" || (source=="TEXT" && origin=="MDBTOOLS_MYSQL")) ? 12 : source=="GUID" ? 15 : -1;
                bool found=false;
                if(size>0 && (code!=10 || origin!="MDBTOOLS_MYSQL")) throw VbaRuntimeError(3251,"DAO source field capacity provenance unavailable");
                for(auto& column:columns) if(column.name.toCaseFolded()==name) { column.sourceType=code; column.sourceSize=static_cast<int>(size); found=true; break; }
                if(!found) throw VbaRuntimeError(3251,"DAO source schema field unavailable");
            }
        }
        if(query.lastError().isValid()) queryFailure(query);
    }
    return columns;
}
QMap<QString,QVariant> literalDefaultValues(QSqlDatabase db,const QVector<TableSchemaColumn>& columns) {
    QStringList keys,expressions;
    for(const auto& field:columns) if(!field.literalDefault.isEmpty()) { keys.append(field.name.toCaseFolded()); expressions.append(field.literalDefault); }
    QMap<QString,QVariant> result;
    if(expressions.isEmpty()) return result;
    QSqlQuery values(db);
    if(!values.exec("SELECT "+expressions.join(',')) || !values.next()) queryFailure(values);
    for(int i=0;i<keys.size();++i) result.insert(keys[i],values.value(i));
    return result;
}
class Recordset;
struct RecordsetData {
    std::shared_ptr<WorkspaceState> workspace;
    quint64 rollbackRevision{};
    QStringList columns;
    QMap<QString,QString> types;
    QMap<QString,int> daoTypes,definitionSizes;
    QMap<QString,QVariant> literalDefaults;
    QVector<QVariantList> rows;
    QVector<QVariant> ids;
    QVector<quint64> bookmarks;
    quint64 nextBookmark{1};
    quint64 revision{};
    QByteArray family{QUuid::createUuid().toRfc4122()};
    QSet<Recordset*> cursors;
};
VbaValue fieldCollection(std::shared_ptr<Recordset> owner);
bool binaryFieldType(const QString& type) {
    const auto normalized=type.toUpper();
    return normalized=="BLOB" || normalized=="BINARY" || normalized=="VARBINARY" || normalized=="LONGBINARY";
}
class Field final : public VbaObject {
public:
    std::string_view className() const override { return "Field"; }
    bool isNativeType(std::string_view type) const override { return type.starts_with("dao.") && type.substr(4)==lower(className()); }
    Field(std::shared_ptr<Recordset> owner, QString name) : owner_(std::move(owner)), name_(std::move(name)) {}
    bool hasScalarDefault() const override { return true; }
    VbaValue scalarValue() override { return member("Value",{}); }
    void assignDefaultMember(const std::vector<VbaValue>& args,const VbaValue& value) override {
        if(!args.empty()) throw VbaRuntimeError(450,"DAO scalar default assignment takes no indices");
        assign("Value",value);
    }
    VbaValue member(std::string_view name, const std::vector<VbaValue>& args) override;
    void assign(std::string_view name, const VbaValue& value) override;
private:
    std::shared_ptr<Recordset> owner_;
    QString name_;
};
class Recordset final : public VbaObject, public std::enable_shared_from_this<Recordset> {
public:
    std::string_view className() const override { return "Recordset"; }
    bool isNativeType(std::string_view type) const override { return type.starts_with("dao.") && type.substr(4)==lower(className()); }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        const auto key=lower(name);
        if (key=="findfirst" || key=="findlast" || key=="findnext" || key=="findprevious")
            return VbaArgumentSignature{{"Criteria"},1};
        if (key=="move") return VbaArgumentSignature{{"Rows","Start"},1};
        if (key=="getrows") return VbaArgumentSignature{{"NumRows"},0};
        return {};
    }
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
                         const QVector<QVariantList>& rows,const QVector<QVariant>& ids,
                         const QMap<qlonglong,qlonglong>& identityChanges) {
        data_->workspace->check();
        if (!frozen_ || columns!=data_->columns || ids.size()!=rows.size() || data_->ids.size()!=data_->rows.size())
            throw VbaRuntimeError(3251,"DAO: Incompatible form snapshot refresh");
        QSet<qlonglong> changedTargets;
        for(auto it=identityChanges.cbegin();it!=identityChanges.cend();++it) {
            if(!data_->ids.contains(QVariant::fromValue(it.key())) || changedTargets.contains(it.value())
                || (it.key()!=it.value() && data_->ids.contains(QVariant::fromValue(it.value())) && !identityChanges.contains(it.value())))
                throw VbaRuntimeError(3251,"DAO: Ambiguous form snapshot identity change");
            changedTargets.insert(it.value());
        }
        const auto refreshedId=[&](const QVariant& id) {
            const auto found=identityChanges.constFind(id.toLongLong());
            return found==identityChanges.cend() ? id : QVariant::fromValue(found.value());
        };
        if (data_->rollbackRevision!=data_->workspace->rollbackRevision) { data_->family=QUuid::createUuid().toRfc4122(); data_->rollbackRevision=data_->workspace->rollbackRevision; }
        QMap<qlonglong,quint64> tokens;
        QMap<Recordset*,QVariant> positions;
        for (qsizetype i=0;i<data_->ids.size();++i) tokens.insert(refreshedId(data_->ids[i]).toLongLong(),data_->bookmarks[i]);
        for (auto* cursor:data_->cursors)
            if (cursor->position_>=0 && cursor->position_<data_->ids.size()) positions.insert(cursor,refreshedId(data_->ids[cursor->position_]));
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
    void assignDefaultMember(const std::vector<VbaValue>& args,const VbaValue& value) override { assignIndexedMember("Fields",args,value); }
    void assignIndexedMember(std::string_view name,const std::vector<VbaValue>& args,const VbaValue& value) override {
        if(lower(name)!="fields") throw VbaRuntimeError(438,"DAO: Indexed Recordset property unsupported");
        if(args.size()!=1) throw VbaRuntimeError(450,"DAO: Recordset field assignment requires one index");
        member("Fields",args).object->assign("Value",value);
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
        if(key=="getrows") {
            if(args.size()>1) throw VbaRuntimeError(450,"DAO: GetRows accepts one row count");
            if(args.empty() || args[0].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(3251,"DAO: GetRows currently requires explicit positive NumRows");
            const auto requested=static_cast<qsizetype>(coerceVbaModuleValue(scalarVbaValue(args[0]),"long").number);
            if(requested<=0) throw VbaRuntimeError(3251,"DAO: GetRows currently requires positive NumRows");
            if(mode_!=Mode::Browse) throw VbaRuntimeError(3219,"DAO: Pending edit");
            current();
            const auto count=std::min(requested,data_->rows.size()-position_); const auto fields=data_->columns.size();
            if(fields<=0 || count>1000000/fields) throw VbaRuntimeError(3251,"DAO: GetRows exceeds the supported cell limit");
            std::vector<VbaValue> cells; cells.reserve(static_cast<std::size_t>(count*fields));
            for(qsizetype row=position_;row<position_+count;++row) for(qsizetype field=0;field<fields;++field) {
                const auto raw=data_->rows[row][field];
                if(!raw.isNull() && raw.metaType().id()==QMetaType::QByteArray) {
                    const auto bytes=raw.toByteArray(); cells.push_back(VbaValue::byteArray(std::vector<unsigned char>(bytes.begin(),bytes.end())));
                } else cells.push_back(qtFieldValue(raw,data_->types.value(data_->columns[field].toCaseFolded())));
            }
            auto result=makeVbaArrayWithBounds(std::move(cells),{{0,static_cast<int>(fields-1)},{0,static_cast<int>(count-1)}});
            position_+=count; noMatch_=false; return result;
        }
        if (key=="findfirst" || key=="findlast" || key=="findnext" || key=="findprevious") {
            if (args.size()!=1) throw VbaRuntimeError(450,"DAO: Find requires one criterion");
            if (args[0].kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Invalid Null criterion");
            find(key,text(args[0].asText())); return {};
        }
        if (key=="move") {
            if (args.empty() || args.size()>2) throw VbaRuntimeError(450,"DAO: Move requires Rows and optional Start");
            if (mode_!=Mode::Browse) throw VbaRuntimeError(3219,"DAO: Pending edit");
            if (args[0].kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Invalid Null movement");
            const auto distance=sqlFieldValue(args[0],"INTEGER4").toLongLong();
            const auto origin=args.size()==2 ? bookmarkPosition(args[1]) : position_;
            if (data_->rows.empty() || (origin<0 && distance<=0) ||
                (origin>=data_->rows.size() && distance>=0))
                throw VbaRuntimeError(3021,"DAO: No current record");
            const auto destination=static_cast<qlonglong>(origin)+distance;
            position_=static_cast<qsizetype>(std::clamp<qlonglong>(destination,-1,data_->rows.size()));
            noMatch_=false;
            return {};
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
        if (key=="lastmodified") {
            if(readOnly_) throw VbaRuntimeError(3251,"DAO: LastModified requires an editable recordset");
            if(!lastModified_) throw VbaRuntimeError(3021,"DAO: No record modified by this cursor");
            (void)bookmarkPosition(*lastModified_); return *lastModified_;
        }
        if (key=="bookmarkable") return VbaValue::logical(true);
        if (key=="updatable") return VbaValue::logical(!readOnly_);
        if (key=="editmode") return VbaValue::integer(mode_==Mode::Browse ? 0 : mode_==Mode::Edit ? 1 : 2);
        if (key=="nomatch") return VbaValue::logical(noMatch_);
        if (key == "eof") return VbaValue::logical(data_->rows.empty() || position_ >= data_->rows.size());
        if (key == "bof") return VbaValue::logical(data_->rows.empty() || position_ < 0);
        if (key == "recordcount") return longCount(data_->rows.size());
        if (key=="absoluteposition") return longCount(position_<0 || position_>=data_->rows.size() ? -1 : position_);
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
            else {
                pending_=QVariantList(data_->columns.size());
                for(int i=0;i<data_->columns.size();++i) {
                    const auto key=data_->columns[i].toCaseFolded();
                    if(data_->literalDefaults.contains(key)) pending_[i]=data_->literalDefaults.value(key);
                }
                mode_=Mode::Add;
            }
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
        if(lower(name)=="absoluteposition") {
            checkTransaction();
            if(closed_) throw VbaRuntimeError(3420,"DAO: Recordset is closed");
            if(mode_!=Mode::Browse) throw VbaRuntimeError(3219,"DAO: Pending edit");
            const auto scalar=scalarVbaValue(value);
            if(scalar.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Invalid Null position");
            const auto position=sqlFieldValue(scalar,"INTEGER4").toLongLong();
            if(position<0 || position>=data_->rows.size()) throw VbaRuntimeError(3021,"DAO: Position outside populated recordset");
            position_=static_cast<qsizetype>(position); noMatch_=false; return;
        }
        if (lower(name)=="editmode" || lower(name)=="updatable" || lower(name)=="lastmodified")
            throw VbaRuntimeError(438,"DAO: Recordset metadata is read-only");
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
    VbaValue declaredFieldType(const QString& name) const {
        const auto declaration=fieldType(name),key=name.toCaseFolded();
        const int code=data_->daoTypes.contains(key) ? data_->daoTypes.value(key) : provenDaoFieldType(declaration);
        if(code<=0) throw VbaRuntimeError(3251,"Field.Type is not proven for this source");
        return VbaValue::integer(static_cast<std::int16_t>(code));
    }
    VbaValue definitionFieldSize(const QString& name) const {
        const auto declaration=fieldType(name),key=name.toCaseFolded();
        if(data_->definitionSizes.contains(key)) {
            const auto size=data_->definitionSizes.value(key);
            if(size<0) throw VbaRuntimeError(3251,"Field.Size is not proven for this source");
            return VbaValue::longInteger(size);
        }
        return fixedDaoFieldSize(provenDaoFieldType(declaration));
    }
    VbaValue fieldOrdinal(const QString& name) const {
        (void)fieldType(name); return integerCount(column(name));
    }
    VbaValue fieldSize(const QString& name) const {
        const auto type=fieldType(name);
        if(!binaryFieldType(type) || mode_!=Mode::Browse)
            throw VbaRuntimeError(3251,"FieldSize supports stored binary observations only");
        current(); const auto& raw=data_->rows[position_][column(name)];
        if(raw.isNull()) throw VbaRuntimeError(3251,"Null FieldSize requires reference validation");
        if(raw.metaType().id()!=QMetaType::QByteArray) throw VbaRuntimeError(13,"Binary field requires BLOB storage");
        const auto size=raw.toByteArray().size();
        if(size>std::numeric_limits<std::int32_t>::max()) throw VbaRuntimeError(6,"FieldSize Long overflow");
        return VbaValue::longInteger(static_cast<std::int32_t>(size));
    }
    VbaValue value(const QString& name) const {
        checkTransaction();
        if (closed_) throw VbaRuntimeError(3420,"DAO: Recordset is closed");
        const int index = column(name);
        if(mode_==Mode::Browse) current();
        const auto& raw=mode_!=Mode::Browse ? pending_[index] : data_->rows[position_][index];
        const auto type=data_->types.value(name.toCaseFolded());
        if(!raw.isNull() && raw.metaType().id()==QMetaType::QByteArray) {
            if(!type.isEmpty() && !binaryFieldType(type)) throw VbaRuntimeError(13,"Binary storage does not match declared field type");
            const auto bytes=raw.toByteArray();
            return VbaValue::byteArray(std::vector<unsigned char>(bytes.begin(),bytes.end()));
        }
        if(!raw.isNull() && binaryFieldType(type)) throw VbaRuntimeError(13,"Binary field requires BLOB storage");
        return qtFieldValue(raw,type);
    }
    void set(const QString& name, const VbaValue& value) {
        checkTransaction();
        if (closed_) throw VbaRuntimeError(3420,"DAO: Recordset is closed");
        require(mode_ != Mode::Browse, "DAO: Edit or AddNew required");
        const auto index = column(name);
        const auto type=data_->types.value(name.toCaseFolded());
        if(value.kind==VbaValue::Kind::ByteArray) {
            if(!binaryFieldType(type)) throw VbaRuntimeError(13,"Byte array requires a binary field");
            if(value.bytes.size()>static_cast<std::size_t>(std::numeric_limits<qsizetype>::max()))
                throw VbaRuntimeError(6,"Binary field length overflow");
            pending_[index]=value.bytes.empty() ? QByteArray("",0)
                : QByteArray(reinterpret_cast<const char*>(value.bytes.data()),static_cast<qsizetype>(value.bytes.size()));
        } else {
            if(binaryFieldType(type) && value.kind!=VbaValue::Kind::Null)
                throw VbaRuntimeError(3251,"Binary field assignment requires Byte() or Null");
            pending_[index]=sqlFieldValue(value,type);
        }
        changed_.insert(index);
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
        data_->types.clear(); data_->daoTypes.clear(); data_->definitionSizes.clear(); data_->literalDefaults.clear();
        if(table_) {
            QSqlQuery origin(db_); origin.prepare("SELECT name FROM main.sqlite_schema WHERE type='table' AND name=? COLLATE NOCASE"); origin.addBindValue(source_);
            if(!origin.exec()) queryFailure(origin);
            const auto mainName=origin.next() ? origin.value(0).toString() : QString{}; origin.finish();
            if(mainName.isEmpty()) data_->types=sqliteColumnTypes(db_,source_);
            else {
                const auto schema=readTableSchemaColumns(db_,mainName);
                data_->literalDefaults=literalDefaultValues(db_,schema);
                for(const auto& field:schema) {
                const auto key=field.name.toCaseFolded(); data_->types.insert(key,field.type);
                const auto code=field.sourceType ? field.sourceType : provenDaoFieldType(field.type);
                data_->daoTypes.insert(key,code);
                const auto size=code<=0 ? -1 : code==10 ? (field.sourceSize>0 ? field.sourceSize : -1)
                    : static_cast<int>(fixedDaoFieldSize(code).number);
                data_->definitionSizes.insert(key,size);
                }
            }
        }
        if (!table_) {
            if(installAccessQueryViews(db_)<0) throw VbaRuntimeError(3251,"DAO: Query views are busy");
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
                    : "UPDATE " + quote(source_) + " SET " + assignments.join(',') + " WHERE rowid=? RETURNING rowid"),
                    "DAO: Update prepare failed");
            for (const int i : indices) query.addBindValue(pending_[i]);
            if (!adding) query.addBindValue(id);
            if (!query.exec()) queryFailure(query);
            if(adding) require(query.numRowsAffected() == 1, "DAO: Insert failed or record changed");
            else {
                require(query.next(),"DAO: Update failed or record changed");
                id=query.value(0);
                require(!query.next(),"DAO: Update returned multiple identities");
                if(query.lastError().isValid()) queryFailure(query);
            }
        }
        if (adding) id = query.lastInsertId();
        if(!adding && position_>=0 && position_<data_->ids.size()) data_->ids[position_]=id;
        mode_ = Mode::Browse; pending_.clear(); changed_.clear();
        reload(false); position_ = data_->ids.indexOf(id);
        if(position_>=0) lastModified_=bookmarkAt(position_); else lastModified_.reset();
    }
    std::shared_ptr<RecordsetData> data_{std::make_shared<RecordsetData>()};
    QSqlDatabase db_;
    QString source_;
    QMap<QString,QVariant> bindings_;
    bool table_{}, readOnly_{}, closed_{}, frozen_{},noMatch_{},cloneCounted_{};
    std::optional<VbaValue> lastModified_;
    QVariantList pending_;
    QSet<int> changed_;
    qsizetype position_{};
    Mode mode_{Mode::Browse};
    AccessQueryPlan queryPlan_;
};
class Fields final : public VbaObject {
public:
    std::string_view className() const override { return "Fields"; }
    bool isNativeType(std::string_view type) const override { return type.starts_with("dao.") && type.substr(4)==lower(className()); }
    explicit Fields(std::shared_ptr<Recordset> owner) : owner_(std::move(owner)) {}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        const auto key = lower(name);
        if (key == "count" && args.empty()) return integerCount(owner_->fieldNames().size());
        if (key == "item" && args.size() == 1) return owner_->member("Fields",args);
        throw VbaRuntimeError(438,"DAO: Fields member unsupported");
    }
    VbaValue item(std::string_view name) override { return owner_->member("Fields",{VbaValue::string(std::string(name))}); }
    void assignIndexedMember(std::string_view name,const std::vector<VbaValue>& args,const VbaValue& value) override {
        if(lower(name)!="item") throw VbaRuntimeError(438,"DAO: Indexed Fields property unsupported");
        if(args.size()!=1) throw VbaRuntimeError(450,"DAO: Fields assignment requires one index");
        owner_->member("Fields",args).object->assign("Value",value);
    }
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
    if(lower(name)=="fieldsize") {
        if(!args.empty()) throw VbaRuntimeError(450,"Field.FieldSize requires no arguments");
        return owner_->fieldSize(name_);
    }
    if(lower(name)=="ordinalposition") {
        if(!args.empty()) throw VbaRuntimeError(450,"Field.OrdinalPosition requires no arguments");
        return owner_->fieldOrdinal(name_);
    }
    if(lower(name)=="size") {
        if(!args.empty()) throw VbaRuntimeError(450,"Field.Size requires no arguments");
        return owner_->definitionFieldSize(name_);
    }
    if(lower(name)=="type") {
        if(!args.empty()) throw VbaRuntimeError(450,"Field.Type requires no arguments");
        return owner_->declaredFieldType(name_);
    }
    require(args.empty(), "DAO: Unexpected field arguments");
    if (lower(name) == "value") return owner_->value(name_);
    if (lower(name) == "name") return VbaValue::string(name_.toUtf8().toStdString());
    throw std::runtime_error("DAO: Unknown Field member");
}
void Field::assign(std::string_view name, const VbaValue& value) {
    if(lower(name)=="size") { (void)owner_->fieldType(name_); throw VbaRuntimeError(383,"Field.Size is read-only"); }
    if(lower(name)=="fieldsize") { (void)owner_->fieldSize(name_); throw VbaRuntimeError(383,"FieldSize is read-only"); }
    if(lower(name)=="ordinalposition") { (void)owner_->fieldOrdinal(name_); throw VbaRuntimeError(383,"Recordset Field.OrdinalPosition is read-only"); }
    if(lower(name)=="type") { (void)owner_->fieldType(name_); throw VbaRuntimeError(383,"Field.Type is read-only"); }
    require(lower(name) == "value", "DAO: Field property is read-only");
    owner_->set(name_, value);
}
class QueryDefinition;
class QueryParameter final : public VbaObject {
public:
    std::string_view className() const override { return "Parameter"; }
    bool isNativeType(std::string_view type) const override { return type.starts_with("dao.") && type.substr(4)==lower(className()); }
    QueryParameter(std::shared_ptr<QueryDefinition> query, QString name,quint64 generation) : query_(std::move(query)), name_(std::move(name)),generation_(generation) {}
    bool hasScalarDefault() const override { return true; }
    VbaValue scalarValue() override { return member("Value",{}); }
    void assignDefaultMember(const std::vector<VbaValue>& args,const VbaValue& value) override {
        if(!args.empty()) throw VbaRuntimeError(450,"DAO scalar default assignment takes no indices");
        assign("Value",value);
    }
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
    std::string_view className() const override { return "QueryDef"; }
    bool isNativeType(std::string_view type) const override { return type.starts_with("dao.") && type.substr(4)==lower(className()); }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        const auto key=lower(name);
        if (key=="openrecordset") return VbaArgumentSignature{{"Type"},0};
        if (key=="execute") return VbaArgumentSignature{{"Options"},0};
        return {};
    }
    QueryDefinition(QSqlDatabase db, QString source,QString name={},bool owned=false)
        : db_(std::move(db)), source_(std::move(source)), plan_(source_.trimmed().isEmpty() ? AccessQueryPlan{} : compileAccessQuery(source_)),name_(std::move(name)),owned_(owned),workspace_(workspaceState(db_)) {
        if(owned_) namedGeneration_=workspace_->queryGenerations.value(name_.toCaseFolded());
    }
    VbaValue member(std::string_view name, const std::vector<VbaValue>& args) override {
        refresh();
        const auto key = lower(name);
        if (key == "parameters") {
            auto collection=parameters();
            return args.empty() ? VbaValue::reference(collection) : collection->member("Item",args);
        }
        if (key == "openrecordset") {
            if(plan_.sql.isEmpty()) throw VbaRuntimeError(3251,"DAO: QueryDef SQL is not set");
            require(args.empty() || (args.size() == 1 && (args[0].asNumber() == 2 || args[0].asNumber() == 4)),"DAO: Recordset type unsupported");
            require(plan_.sql.trimmed().startsWith("SELECT ",Qt::CaseInsensitive),"DAO: SELECT QueryDef required");
            return VbaValue::reference(std::make_shared<Recordset>(db_,plan_.sql,true,bindAccessQueryParameters(plan_,values_),plan_));
        }
        if (key == "execute") {
            if(plan_.sql.isEmpty()) throw VbaRuntimeError(3251,"DAO: QueryDef SQL is not set");
            if(owned_ && installAccessQueryViews(db_)<0) throw VbaRuntimeError(3251,"DAO: Query views are busy");
            require(args.empty() || (args.size() == 1 && args[0].asNumber() == 128),"DAO: Execute options unsupported");
            affected_=executeAction(db_,plan_,bindAccessQueryParameters(plan_,values_)); return {};
        }
        require(args.empty(),"DAO: Unexpected QueryDef arguments");
        if (key == "sql") return VbaValue::string(source_.toUtf8().toStdString());
        if (key == "name") return VbaValue::string(name_.toStdString());
        if (key == "recordsaffected") return longCount(affected_);
        if (key == "close") { closed_ = true; return {}; }
        throw VbaRuntimeError(438,"DAO: QueryDef member unsupported");
    }
    void assign(std::string_view name, const VbaValue& value) override {
        refresh();
        if(lower(name)=="name") throw VbaRuntimeError(3251,"DAO: QueryDef renaming is unsupported");
        require(lower(name) == "sql","DAO: QueryDef property is not writable");
        const auto scalar=scalarVbaValue(value);
        if(scalar.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Invalid Null SQL");
        const auto source=text(scalar.asText());
        const auto compiled=source.trimmed().isEmpty() ? AccessQueryPlan{} : compileAccessQuery(source);
        if(owned_) {
            QSqlQuery query(db_); query.prepare("SELECT definition FROM main._la_objects WHERE kind='query' AND name=?"); query.addBindValue(name_);
            if(!query.exec() || !query.next()) throw VbaRuntimeError(3420,"DAO: Owned query was removed");
            const auto previous=query.value(0); const auto document=QJsonDocument::fromJson(previous.toByteArray());
            if(!document.isObject() || document.object().value("sql").toString()!=source_)
                throw VbaRuntimeError(3197,"DAO: Owned query changed concurrently");
            auto definition=document.object(); definition.insert("sql",source); query.finish();
            query.prepare("UPDATE main._la_objects SET definition=? WHERE kind='query' AND name=? AND definition=?");
            query.addBindValue(QString::fromUtf8(QJsonDocument(definition).toJson(QJsonDocument::Compact)));
            query.addBindValue(name_); query.addBindValue(previous);
            if(!query.exec()) queryFailure(query);
            if(query.numRowsAffected()!=1) throw VbaRuntimeError(3197,"DAO: Owned query changed concurrently");
        }
        installSource(source,compiled);
    }
    VbaValue parameter(const QString& name) const {
        checkParameter(name);
        const auto stored=values_.value(name.toCaseFolded());
        for(const auto& parameter:plan_.parameters) if(parameter.name.compare(name,Qt::CaseInsensitive)==0) {
            if(parameter.type=="currency") return qtFieldValue(stored,"CURRENCY_REAL");
            if(parameter.type=="date" || parameter.type=="datetime") return qtFieldValue(stored,"DATETIME");
        }
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
    friend class QueryDefinitions;
    friend class Database;
    std::shared_ptr<VbaObject> parameters();
    void installSource(const QString& source,const AccessQueryPlan& plan) {
        source_=source; plan_=plan; values_.clear(); parameterObjects_.clear(); ++generation_;
    }
    void alive() const {
        workspace_->check(); if(closed_) throw VbaRuntimeError(3420,"DAO: QueryDef is closed");
        if(owned_ && namedGeneration_!=workspace_->queryGenerations.value(name_.toCaseFolded()))
            throw VbaRuntimeError(3420,"DAO: QueryDef namespace was replaced");
    }
    QString ownedSource() const {
        alive(); bool owned=false;
        try {
            const auto source=accessQuerySource(db_,name_,nullptr,&owned,true);
            if(!owned) throw VbaRuntimeError(3420,"DAO: Owned query identity unavailable");
            return source;
        } catch(const VbaRuntimeError& failure) {
            if(failure.number==3265) throw VbaRuntimeError(3420,"DAO: Owned query was removed"); throw;
        }
    }
    void refresh() {
        alive();
        if(owned_) { const auto source=ownedSource(); if(source!=source_) { const auto plan=source.trimmed().isEmpty() ? AccessQueryPlan{} : compileAccessQuery(source); installSource(source,plan); } }
    }
    void parameterGeneration(quint64 generation) const {
        alive();
        if(owned_ && ownedSource()!=source_) throw VbaRuntimeError(3420,"DAO: Parameter source changed");
        if (closed_ || generation!=generation_) throw VbaRuntimeError(3420,"DAO: Parameter is no longer valid");
    }
    void checkParameter(const QString& name) const {
        alive();
        if(owned_ && ownedSource()!=source_) throw VbaRuntimeError(3420,"DAO: Parameter source changed");
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
    QString name_; bool owned_{};
    quint64 namedGeneration_{};
    std::shared_ptr<WorkspaceState> workspace_;
};
class QueryParameters final : public VbaObject {
public:
    std::string_view className() const override { return "Parameters"; }
    bool isNativeType(std::string_view type) const override { return type.starts_with("dao.") && type.substr(4)==lower(className()); }
    explicit QueryParameters(std::shared_ptr<QueryDefinition> query) : query_(std::move(query)) {}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        query_->refresh();
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
        query_->refresh();
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
class QueryDefinitions final : public VbaObject {
public:
    explicit QueryDefinitions(QSqlDatabase db) : db_(std::move(db)),workspace_(workspaceState(db_)) {}
    std::string_view className() const override { return "QueryDefs"; }
    bool isNativeType(std::string_view type) const override { return type=="dao.querydefs"; }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        if(lower(name)=="item") return VbaArgumentSignature{{"Index"},1};
        if(lower(name)=="delete") return VbaArgumentSignature{{"Name"},1}; return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        workspace_->check(); const auto key=lower(name);
        if(key=="delete") {
            if(args.size()!=1 || args[0].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(450,"DAO: QueryDefs.Delete requires a name");
            const auto value=scalarVbaValue(args[0]);
            if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Invalid Null QueryDef name");
            QSqlQuery query(db_); const auto savepoint="la_delete_query_"+QUuid::createUuid().toString(QUuid::Id128);
            if(!query.exec("SAVEPOINT "+savepoint)) queryFailure(query);
            bool completed=false;
            const auto rollback=qScopeGuard([&] { if(!completed) { query.exec("ROLLBACK TO "+savepoint); query.exec("RELEASE "+savepoint); } });
            QString canonical; bool owned=false; (void)accessQuerySource(db_,text(value.asText()),&canonical,&owned,true);
            if(!owned) throw VbaRuntimeError(3251,"DAO: Imported query deletion requires an archive-preserving overlay");
            auto& generation=workspace_->queryGenerations[canonical.toCaseFolded()];
            query.prepare("DELETE FROM main._la_objects WHERE kind='query' AND name=?"); query.addBindValue(canonical);
            if(!query.exec()) queryFailure(query);
            if(query.numRowsAffected()!=1) throw VbaRuntimeError(3265,"DAO: QueryDef no longer exists"); query.finish();
            if(!query.exec("RELEASE "+savepoint)) queryFailure(query);
            ++generation; objects_.remove(canonical.toCaseFolded()); completed=true; return {};
        }
        if(key=="count" && args.empty()) return longCount(names().size());
        if(key=="refresh" && args.empty()) { (void)names(); return {}; }
        if(key!="item") throw VbaRuntimeError(438,"DAO: QueryDefs member unsupported");
        if(args.size()!=1 || args[0].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(450,"DAO: QueryDef name/index required");
        const auto value=scalarVbaValue(args[0]); QString requested;
        if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Invalid Null QueryDef index");
        if(value.kind==VbaValue::Kind::Number || value.kind==VbaValue::Kind::Currency || value.kind==VbaValue::Kind::Boolean || value.kind==VbaValue::Kind::Empty) {
            const auto index=sqlFieldValue(value,"INTEGER4").toLongLong(); const auto entries=names();
            if(index<0 || index>=entries.size()) throw VbaRuntimeError(3265,"DAO: QueryDef index not found");
            requested=entries[index];
        } else requested=text(value.asText());
        QString canonical; bool owned=false; const auto source=accessQuerySource(db_,requested,&canonical,&owned,true);
        auto& weak=objects_[canonical.toCaseFolded()]; auto object=weak.lock();
        if(!object || object->closed_ || (object->owned_ && object->namedGeneration_!=workspace_->queryGenerations.value(canonical.toCaseFolded()))) { object=std::make_shared<QueryDefinition>(db_,source,canonical,owned); weak=object; }
        return VbaValue::reference(object);
    }
    VbaValue item(std::string_view name) override { return member("Item",{VbaValue::string(std::string(name))}); }
    std::vector<VbaValue> enumerate() override {
        std::vector<VbaValue> values;
        for(const auto& name:names()) values.push_back(member("Item",{VbaValue::string(name.toStdString())}));
        return values;
    }
private:
    QStringList names() const {
        workspace_->check(); QSqlQuery query(db_); QMap<QString,QString> entries;
        if(query.exec("SELECT name FROM main._la_access_artifacts WHERE kind='query' ORDER BY id"))
            while(query.next()) { const auto name=query.value(0).toString(); entries.insert(name.toCaseFolded(),name); }
        query.finish();
        if(query.exec("SELECT name FROM main._la_objects WHERE kind='query'")) {
            int rows=0;
            while(query.next()) {
                if(++rows>1000000) throw VbaRuntimeError(3251,"DAO: QueryDefs catalog limit exceeded");
                const auto name=query.value(0).toString(),folded=name.toCaseFolded();
                if(entries.contains(folded)) throw VbaRuntimeError(5,"DAO: QueryDef name ambiguous");
                entries.insert(folded,name);
            }
            if(query.lastError().isValid()) throw VbaRuntimeError(3146,"DAO: QueryDefs catalog read failed");
        }
        return entries.values();
    }
    QSqlDatabase db_; std::shared_ptr<WorkspaceState> workspace_;
    QMap<QString,std::weak_ptr<QueryDefinition>> objects_;
};

using TableSchemaReader=std::function<QVector<TableSchemaColumn>()>;
class TableSchemaField final : public VbaObject {
public:
    TableSchemaField(TableSchemaReader reader,QString name) : reader_(std::move(reader)),name_(std::move(name)) {}
    std::string_view className() const override { return "Field"; }
    bool isNativeType(std::string_view type) const override { return type=="dao.field"; }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        if(!args.empty()) throw VbaRuntimeError(450,"DAO schema field properties take no arguments");
        const auto columns=reader_();
        int position=-1;
        for(int i=0;i<columns.size();++i) if(columns[i].name.compare(name_,Qt::CaseInsensitive)==0) {
            if(position>=0) throw VbaRuntimeError(5,"DAO schema field name ambiguous"); position=i;
        }
        if(position<0) throw VbaRuntimeError(3420,"DAO schema field unavailable");
        const auto key=lower(name);
        if(key=="name") return VbaValue::string(columns[position].name.toStdString());
        if(key=="ordinalposition") return integerCount(position);
        if(key=="type" || key=="size") {
            const auto& column=columns[position];
            const int code=column.sourceType ? column.sourceType : provenDaoFieldType(column.type);
            if(code<=0) throw VbaRuntimeError(3251,"DAO schema field type is not proven");
            if(key=="size" && code==10 && column.sourceSize>0) return VbaValue::longInteger(column.sourceSize);
            return key=="size" ? fixedDaoFieldSize(code) : VbaValue::integer(static_cast<std::int16_t>(code));
        }
        throw VbaRuntimeError(3251,"DAO schema field property is unsupported");
    }
    void assign(std::string_view,const VbaValue&) override { (void)reader_(); throw VbaRuntimeError(383,"DAO schema fields are read-only"); }
private: TableSchemaReader reader_; QString name_;
};
class TableSchemaFields final : public VbaObject {
public:
    explicit TableSchemaFields(TableSchemaReader reader) : reader_(std::move(reader)) {}
    std::string_view className() const override { return "Fields"; }
    bool isNativeType(std::string_view type) const override { return type=="dao.fields"; }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        if(lower(name)=="item") return VbaArgumentSignature{{"Index"},1}; return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        const auto columns=reader_(); const auto key=lower(name);
        if(key=="count" && args.empty()) return integerCount(columns.size());
        if(key=="refresh" && args.empty()) return {};
        if(key!="item") throw VbaRuntimeError(438,"DAO schema Fields member unsupported");
        if(args.size()!=1) throw VbaRuntimeError(450,"DAO schema Fields requires one index");
        const auto value=scalarVbaValue(args[0]);
        if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null schema field index");
        int position=-1;
        if(value.kind==VbaValue::Kind::Text) {
            const auto requested=text(value.asText());
            for(int i=0;i<columns.size();++i) if(columns[i].name.compare(requested,Qt::CaseInsensitive)==0) {
                if(position>=0) throw VbaRuntimeError(5,"DAO schema field name ambiguous"); position=i;
            }
        } else {
            const auto index=coerceVbaModuleValue(value,"long").number;
            if(index>=0 && index<columns.size()) position=static_cast<int>(index);
        }
        if(position<0) throw VbaRuntimeError(3265,"DAO schema field index not found");
        const auto column=columns[position].name;
        auto& cached=fields_[column.toCaseFolded()]; auto field=cached.lock();
        if(!field) { field=std::make_shared<TableSchemaField>(reader_,column); cached=field; }
        return VbaValue::reference(field);
    }
    VbaValue item(std::string_view name) override { return member("Item",{VbaValue::string(std::string(name))}); }
    std::vector<VbaValue> enumerate() override {
        const auto columns=reader_(); std::vector<VbaValue> result;
        for(const auto& column:columns) result.push_back(member("Item",{VbaValue::string(column.name.toStdString())}));
        return result;
    }
private: TableSchemaReader reader_; QMap<QString,std::weak_ptr<TableSchemaField>> fields_;
};
class TableDefinition final : public VbaObject, public std::enable_shared_from_this<TableDefinition> {
public:
    std::string_view className() const override { return "TableDef"; }
    bool isNativeType(std::string_view type) const override { return type.starts_with("dao.") && type.substr(4)==lower(className()); }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        const auto key=lower(name);
        if (key=="fields") return VbaArgumentSignature{{"Index"},0};
        if (key=="openrecordset") return VbaArgumentSignature{{"Type"},0};
        return {};
    }
    TableDefinition(QSqlDatabase db,QString name) : db_(std::move(db)),name_(std::move(name)),workspace_(workspaceState(db_)) {}
    bool valid() {
        workspace_->check();
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
        if(key=="fields") {
            if(args.size()>1) throw VbaRuntimeError(450,"DAO TableDef.Fields takes at most one index");
            const auto version=schemaVersion(),rollback=workspace_->rollbackRevision;
            auto fields=fields_.lock();
            if(!fields || version!=fieldsVersion_ || rollback!=fieldsRollback_) {
                const auto self=shared_from_this();
                fields=std::make_shared<TableSchemaFields>([self,version,rollback] {
                    if(!self->valid() || self->schemaVersion()!=version || self->workspace_->rollbackRevision!=rollback)
                        throw VbaRuntimeError(3420,"DAO schema metadata invalidated; reacquire TableDef.Fields");
                    return readTableSchemaColumns(self->db_,self->name_);
                });
                fields_=fields; fieldsVersion_=version; fieldsRollback_=rollback;
            }
            return args.empty() || args[0].kind==VbaValue::Kind::Missing ? VbaValue::reference(fields) : fields->member("Item",args);
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
    quint64 schemaVersion() const {
        workspace_->check(); QSqlQuery query(db_);
        if(!query.exec("PRAGMA main.schema_version") || !query.next()) queryFailure(query);
        return query.value(0).toULongLong();
    }
    QSqlDatabase db_;
    QString name_;
    std::shared_ptr<WorkspaceState> workspace_;
    std::weak_ptr<TableSchemaFields> fields_;
    quint64 fieldsVersion_{},fieldsRollback_{};
    bool invalidated_{};
};
class TableDefinitions final : public VbaObject {
public:
    std::string_view className() const override { return "TableDefs"; }
    bool isNativeType(std::string_view type) const override { return type.starts_with("dao.") && type.substr(4)==lower(className()); }
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
    std::string_view className() const override { return "Database"; }
    bool isNativeType(std::string_view type) const override { return type.starts_with("dao.") && type.substr(4)==lower(className()); }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        const auto key=lower(name);
        if (key=="openrecordset") return VbaArgumentSignature{{"Name","Type"},1};
        if (key=="execute") return VbaArgumentSignature{{"Query","Options"},1};
        if (key=="createquerydef") return VbaArgumentSignature{{"Name","SQLText"},1};
        if (key=="querydefs") return VbaArgumentSignature{{"Index"},0};
        return {};
    }
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
            if(!queryDefinitions_) queryDefinitions_=std::make_shared<QueryDefinitions>(db_);
            return args.empty() || (args.size()==1 && args[0].kind==VbaValue::Kind::Missing)
                ? VbaValue::reference(queryDefinitions_) : queryDefinitions_->member("Item",args);
        }
        if (key == "createquerydef") {
            if(args.empty() || args.size()>2) throw VbaRuntimeError(450,"DAO: CreateQueryDef requires Name and optional SQLText");
            const auto nameValue=scalarVbaValue(args[0]);
            const auto sqlValue=args.size()==2 && args[1].kind!=VbaValue::Kind::Missing ? scalarVbaValue(args[1]) : VbaValue::string("");
            if(nameValue.kind==VbaValue::Kind::Null || sqlValue.kind==VbaValue::Kind::Null)
                throw VbaRuntimeError(94,"DAO: Invalid Null CreateQueryDef argument");
            const auto name=text(nameValue.asText()),source=text(sqlValue.asText());
            if(name.isEmpty()) return VbaValue::reference(std::make_shared<QueryDefinition>(db_,source));
            if(name.trimmed().isEmpty() || name.size()>64 || name.contains(QChar(0)) || name.contains('\n') || name.contains('\r'))
                throw VbaRuntimeError(5,"DAO: Unsupported QueryDef name");
            auto& generation=workspace_->queryGenerations[name.toCaseFolded()];
            auto object=std::make_shared<QueryDefinition>(db_,source,name,true);
            QSqlQuery query(db_); const auto savepoint="la_create_query_"+QUuid::createUuid().toString(QUuid::Id128);
            if(!query.exec("SAVEPOINT "+savepoint)) queryFailure(query);
            bool completed=false;
            const auto rollback=qScopeGuard([&] { if(!completed) { query.exec("ROLLBACK TO "+savepoint); query.exec("RELEASE "+savepoint); } });
            if(!query.exec("CREATE TABLE IF NOT EXISTS main._la_objects(kind TEXT NOT NULL,name TEXT NOT NULL,definition TEXT NOT NULL,PRIMARY KEY(kind,name))")) queryFailure(query);
            bool exists=false;
            try { (void)accessQuerySource(db_,name,nullptr,nullptr,true); exists=true; }
            catch(const VbaRuntimeError& failure) { if(failure.number!=3265) throw; }
            bool mainView=false;
            { QSqlQuery schema(db_);
                if(!schema.exec("SELECT name FROM main.sqlite_schema WHERE type='view'")) queryFailure(schema);
                while(schema.next()) if(schema.value(0).toString().compare(name,Qt::CaseInsensitive)==0) mainView=true;
                if(schema.lastError().isValid()) queryFailure(schema);
            }
            const bool managedView=!mainView && !db_.tables(QSql::Tables).contains(name,Qt::CaseInsensitive)
                && db_.tables(QSql::Views).contains(name,Qt::CaseInsensitive)
                && db_.driver()->property("la_query_database").toString()==db_.databaseName()
                && db_.driver()->property("la_query_views").toStringList().contains(name,Qt::CaseInsensitive);
            if(exists || (db_.tables(QSql::AllTables).contains(name,Qt::CaseInsensitive) && !managedView))
                throw VbaRuntimeError(3012,"DAO: QueryDef name already exists");
            query.prepare("INSERT INTO main._la_objects(kind,name,definition) VALUES('query',?,?)");
            query.addBindValue(name); query.addBindValue(QString::fromUtf8(QJsonDocument(QJsonObject{{"sql",source}}).toJson(QJsonDocument::Compact)));
            if(!query.exec()) queryFailure(query); query.finish();
            if(!query.exec("RELEASE "+savepoint)) queryFailure(query);
            ++generation; object->namedGeneration_=generation;
            completed=true; return VbaValue::reference(std::move(object));
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
        if (key=="recordsaffected" && args.empty()) return longCount(affected_);
        throw std::runtime_error("DAO: Database member unsupported");
    }
private:
    QSqlDatabase db_;
    std::weak_ptr<TableDefinitions> tableDefinitions_;
    std::shared_ptr<VbaObject> queryDefinitions_;
    std::shared_ptr<WorkspaceState> workspace_;
    qlonglong affected_{};
};
class SingletonDaoCollection final : public VbaObject {
public:
    SingletonDaoCollection(std::shared_ptr<VbaObject> object,QString name,std::string_view type):object_(std::move(object)),name_(std::move(name)),type_(type) {}
    std::string_view className() const override { return type_; }
    bool isNativeType(std::string_view type) const override { return type.starts_with("dao.") && type.substr(4)==lower(className()); }
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
    std::string_view type_;
};
class Workspace final : public VbaObject {
public:
    std::string_view className() const override { return "Workspace"; }
    bool isNativeType(std::string_view type) const override { return type.starts_with("dao.") && type.substr(4)==lower(className()); }
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
            value=std::make_shared<SingletonDaoCollection>(database,db_.databaseName(),"Databases"); databases_=value;
        }
        return value;
    }
    QSqlDatabase db_;
    std::shared_ptr<WorkspaceState> state_;
    std::weak_ptr<VbaObject> databases_;
};
class Engine final : public VbaObject {
public:
    std::string_view className() const override { return "DBEngine"; }
    bool isNativeType(std::string_view type) const override { return type.starts_with("dao.") && type.substr(4)==lower(className()); }
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
        if (!value) { value=std::make_shared<SingletonDaoCollection>(workspace(),"#Default Workspace#","Workspaces"); workspaces_=value; }
        return value;
    }
    QSqlDatabase db_;
    std::shared_ptr<WorkspaceState> state_;
    std::weak_ptr<VbaObject> workspaces_;
};
}
QMap<QString,QVariant> nativeTableLiteralDefaults(QSqlDatabase database,const QString& table) {
    workspaceState(database)->check();
    QSqlQuery origin(database); origin.prepare("SELECT name FROM main.sqlite_schema WHERE type='table' AND name=? COLLATE NOCASE"); origin.addBindValue(table);
    if(!origin.exec()) queryFailure(origin);
    if(!origin.next()) return {};
    const auto canonical=origin.value(0).toString(); origin.finish();
    QSqlQuery shadow(database);
    shadow.prepare("SELECT 1 FROM temp.sqlite_schema WHERE type IN ('table','view') AND (name=? COLLATE NOCASE OR name='_la_table_schema' COLLATE NOCASE)"); shadow.addBindValue(canonical);
    if(!shadow.exec()) queryFailure(shadow);
    if(shadow.next()) throw VbaRuntimeError(3251,"Native defaults have a TEMP shadow"); shadow.finish();
    QSqlQuery columns(database);
    if(!columns.exec("PRAGMA main.table_info("+quote(canonical)+")")) queryFailure(columns);
    QVector<TableSchemaColumn> schema;
    while(columns.next()) schema.push_back({columns.value(1).toString(),QString{},0,0,sqliteLiteralDefaultSql(columns.value(4))});
    if(columns.lastError().isValid()) queryFailure(columns); columns.finish();
    return literalDefaultValues(database,schema);
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
                        const QMap<QString,QString>& types,const QVector<QVariantList>& rows,const QVector<QVariant>& ids,
                        const QMap<qlonglong,qlonglong>& identityChanges) {
    const auto native=std::dynamic_pointer_cast<Recordset>(recordset);
    if (!native) throw VbaRuntimeError(13,"DAO: Native recordset required");
    native->refreshSnapshot(columns,types,rows,ids,identityChanges);
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
