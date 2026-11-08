#include "dao_database.hpp"
#include "vba_array.hpp"
#include "access_autonumber.hpp"
#include "access_index_metadata.hpp"
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
int nativeRecordsetMode(const std::vector<VbaValue>& args,std::size_t index) {
    if(index>=args.size() || args[index].kind==VbaValue::Kind::Missing) return 2;
    const auto mode=static_cast<int>(coerceVbaModuleValue(scalarVbaValue(args[index]),"long").number);
    if(mode!=2 && mode!=4) throw VbaRuntimeError(3251,"DAO: Native Recordset mode is not supported");
    return mode;
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
QStringList nativeTableNames(QSqlDatabase db);
void observeNativeTableCatalog(QSqlDatabase db,const QStringList& before);
void checkReservedCounterSchemaAction(QSqlDatabase db);
void removeDroppedTableMetadata(QSqlDatabase db,const QStringList& before) {
    const auto after=nativeTableNames(db); const auto catalog=db.tables();
    QStringList removed,added;
    for(const auto& name:before) if(!after.contains(name,Qt::CaseInsensitive)) removed.append(name);
    for(const auto& name:after) if(!before.contains(name,Qt::CaseInsensitive)) added.append(name);
    // A rename needs rebinding of defaults, index identities and source metadata.
    // Refuse it atomically for managed/imported tables until that is implemented.
    for(const auto& name:removed) {
        for(const auto& metadata:{QString("_la_counter_sequences"),QString("_la_counter32_sequences"),QString("_la_table_schema"),QString("_la_access_indexes")}) {
            if(!catalog.contains(metadata,Qt::CaseInsensitive)) continue;
            QSqlQuery query(db);
            if(!added.isEmpty()) {
                query.prepare("SELECT 1 FROM main."+quote(metadata)+" WHERE table_name=?");query.addBindValue(name);
                if(!query.exec()) queryFailure(query);
                if(query.next()) throw VbaRuntimeError(3251,"DAO: Rename requires counter/source metadata rebinding");
                continue;
            }
            query.prepare("DELETE FROM main."+quote(metadata)+" WHERE table_name=?");query.addBindValue(name);
            if(!query.exec()) queryFailure(query);
        }
    }
}
qlonglong executeAction(QSqlDatabase db,const AccessQueryPlan& plan,const QMap<QString,QVariant>& bindings) {
    if (plan.sql.trimmed().startsWith("SELECT",Qt::CaseInsensitive)) throw VbaRuntimeError(3065,"DAO: Execute cannot run a SELECT query");
    const bool makeTable=!plan.makeTable.isEmpty();
    const bool schemaAction=plan.daoQueryType==96; // Checked compiler dbQDDL classification, including whitespace/comments.
    if(schemaAction || makeTable) checkReservedCounterSchemaAction(db);
    const auto before=(schemaAction || makeTable) ? nativeTableNames(db) : QStringList{};
    QSqlQuery guard(db);
    const bool schemaScope=makeTable || schemaAction;
    const auto schemaSavepoint="la_schema_"+QUuid::createUuid().toString(QUuid::Id128);
    if(schemaScope && !guard.exec("SAVEPOINT "+schemaSavepoint)) queryFailure(guard);
    try {
        qlonglong affected{};
        {
            SqliteExecutionBudget budget(db);
            QSqlQuery query(db);
            QString action;
            if(makeTable) {
                const auto select=accessMakeTableSource(db,plan);
                QVector<QString> types;
                const auto columns=accessQueryColumns(db,select,&types);
                if(types.contains("CURRENCY_REAL") || types.contains("DATETIME") || types.contains("DATE")) {
                    if(select.hiddenRank || select.topLimit || select.daoQueryType!=0 || !select.rightJoinSql.isEmpty())
                        throw VbaRuntimeError(3251,"Complex typed make-table source unsupported");
                    // Build actual declared columns before executing any source rows.
                    // Unknown output types remain untyped; never infer Jet types from values.
                    const QSet<QString> allowed={"CURRENCY_REAL","DATETIME","DATE","TEXT","BOOLEAN","INTEGER1","INTEGER2","INTEGER4","INTEGER","INT","REAL","DOUBLE","FLOAT"};
                    QSet<QString> names; QStringList definitions;
                    for(qsizetype i=0;i<columns.size();++i) {
                        const auto name=columns[i];
                        if(name.isEmpty() || names.contains(name.toCaseFolded()))
                            throw VbaRuntimeError(3251,"Make-table output names must be unique");
                        names.insert(name.toCaseFolded());
                        const auto type=types.value(i).toUpper();
                        definitions.append(quote(name)+(allowed.contains(type) ? " "+type : QString{}));
                    }
                    if(definitions.isEmpty()) throw VbaRuntimeError(3251,"Make-table columns unavailable");
                    if(!query.exec("CREATE TABLE "+quote(plan.makeTable)+" ("+definitions.join(',')+")")) queryFailure(query);
                    query.finish();
                    action="INSERT INTO "+quote(plan.makeTable)+" "+select.sql;
                }
            }
            if(action.isEmpty()) action=accessActionSql(db,plan);
            if (!query.prepare(action)) queryFailure(query);
            for (auto it=bindings.cbegin();it!=bindings.cend();++it) query.bindValue(it.key(),it.value());
            if (!query.exec()) queryFailure(query);
            if (query.isSelect()) throw VbaRuntimeError(3065,"DAO: Execute cannot run a SELECT query");
            affected=schemaAction ? 0 : std::max<qlonglong>(0,query.numRowsAffected()); query.finish();
            if (makeTable) {
                if (!query.exec("SELECT COUNT(*) FROM "+quote(plan.makeTable)) || !query.next()) queryFailure(query);
                affected=query.value(0).toLongLong(); query.finish();
            }
        }
        if(schemaScope) {
            removeDroppedTableMetadata(db,before);
            const auto after=nativeTableNames(db);
            const bool removed=std::any_of(before.cbegin(),before.cend(),[&](const QString& name){return !after.contains(name,Qt::CaseInsensitive);});
            if(schemaAction && !removed) for(const auto& table:after) if(!before.contains(table,Qt::CaseInsensitive)) configureNewAccessCounters(db,table);
            observeNativeTableCatalog(db,before);
        }
        if(schemaScope && !guard.exec("RELEASE "+schemaSavepoint)) queryFailure(guard);
        return affected;
    } catch (...) {
        if(schemaScope) { guard.exec("ROLLBACK TO "+schemaSavepoint); guard.exec("RELEASE "+schemaSavepoint); }
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
    QMap<QString,quint64> queryGenerations,tableGenerations;
    QVariantMap queryViewSnapshot;
    static constexpr const char* queryViewProperties[]{"la_query_database","la_query_views","la_query_internal_views","la_query_rank_pending","la_query_revision"};
    int clones{};
    QStringList savepoints;
    QString connectionName;
    QMap<QString,AccessCounterReservation> counterReservations;
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
        if(savepoints.isEmpty()) { queryViewSnapshot.clear(); counterReservations.clear(); }
    }
    void rollback() {
        check(); if (savepoints.isEmpty()) throw VbaRuntimeError(3034,"DAO: BeginTrans required");
        if (savepoints.size()!=1) throw VbaRuntimeError(3251,"DAO: Nested rollback unsupported");
        const auto db=QSqlDatabase::database(connectionName,false);
        for(auto& reservation:counterReservations) reservation=currentAccessCounterReservation(db,reservation);
        sql("ROLLBACK TO "+savepoints.first()); ++rollbackRevision;
        for(const auto& reservation:counterReservations) retainAccessCounterReservation(db,reservation);
        sql("RELEASE "+savepoints.first()); savepoints.clear();counterReservations.clear();
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
    state=std::make_shared<WorkspaceState>(); state->driver=db.driver(); state->connection=native; state->connectionName=db.connectionName();
    // The SQLite destructor invalidates wrappers even if the Qt driver is reopened.
    auto* token=new std::shared_ptr<WorkspaceState>(state);
    if (sqlite3_create_function_v2(native,"linaccess_dao_connection_generation",0,SQLITE_UTF8|SQLITE_DIRECTONLY,token,
        [](sqlite3_context* context,int,sqlite3_value**) { sqlite3_result_int(context,1); },nullptr,nullptr,
        [](void* data) { auto* holder=static_cast<std::shared_ptr<WorkspaceState>*>(data); (*holder)->alive=false; (*holder)->connection=nullptr; delete holder; })!=SQLITE_OK)
        throw VbaRuntimeError(3146,"DAO: Workspace lifetime tracking unavailable");
    db.driver()->setProperty("la_dao_workspace_state",QVariant::fromValue(state)); return state;
}
void checkReservedCounterSchemaAction(QSqlDatabase db) {
    const auto state=workspaceState(db);
    if(!state->savepoints.isEmpty() && !state->counterReservations.isEmpty())
        throw VbaRuntimeError(3251,"DAO: Schema changes after counter allocation require transaction completion");
}
void observeNativeTableCatalog(QSqlDatabase db,const QStringList& before) {
    const auto after=nativeTableNames(db); const auto state=workspaceState(db);
    QSet<QString> previous,current;
    for(const auto& name:before) previous.insert(name.toCaseFolded());
    for(const auto& name:after) current.insert(name.toCaseFolded());
    auto names=previous; names.unite(current);
    for(const auto& name:names) if(previous.contains(name)!=current.contains(name)) {
        auto& generation=state->tableGenerations[name];
        if(generation==std::numeric_limits<quint64>::max()) throw VbaRuntimeError(6,"DAO: Table generation overflow");
        ++generation;
    }
}
QString sqliteLiteralDefaultSql(const QVariant& declaration) {
    if(declaration.isNull()) return {};
    const auto token=declaration.toString().trimmed(); const auto literal=accessLiteralDefaultSql(token);
    // Raw SQLite TRUE is1; imported Access TRUE is already normalized to-1.
    return token.compare("TRUE",Qt::CaseInsensitive)==0 ? QString("1") : literal.value_or(QString{});
}
struct TableSchemaColumn { QString name,type; int sourceType{},sourceSize{}; QString literalDefault; bool hasDefault{}; QString foreignName; int indexAttributes{-1}; };
std::optional<QString> nativeFieldDefaultExpression(const TableSchemaColumn& field) {
    if(provenDaoFieldType(field.type)==11 || field.sourceType==11) return {};
    if(!field.hasDefault) return QString{};
    auto expression=field.literalDefault;
    if(expression.isEmpty() || expression.size()>512) return {};
    if(expression=="NULL") expression="Null";
    else if(expression.startsWith(QChar(39))) {
        auto text=expression.mid(1,expression.size()-2); text.replace("''","'");
        if(text.contains(QChar(0)) || text.contains('\r') || text.contains('\n')) return {};
        text.replace(QChar(34),QString("\"\"")); expression=QChar(34)+text+QChar(34);
    } else if(expression.startsWith('.')) expression.prepend('0');
    else if(expression.startsWith("-.") || expression.startsWith("+.")) expression.insert(1,'0');
    if(expression.size()>255) return {};
    return expression;
}
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
        columns.push_back({name,types.value(name.toCaseFolded()),0,0,sqliteLiteralDefaultSql(query.value(4)),!query.value(4).isNull()});
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
    const auto counter32=jetLongAccessColumns(db,tableName);
    for(auto& column:columns) if(counter32.contains(column.name.toCaseFolded())) {column.type="INTEGER4";column.sourceType=4;}
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
    QMap<QString,QString> defaultExpressions;
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
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        if(lower(name)=="getchunk") return VbaArgumentSignature{{"Offset","Bytes"},2};
        if(lower(name)=="appendchunk") return VbaArgumentSignature{{"Val"},1}; return {};
    }
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
        if (key=="fields") return VbaArgumentSignature{{"Index"},0};
        return {};
    }
    Recordset(QSqlDatabase db, QString source, bool snapshot, QMap<QString,QVariant> bindings = {},AccessQueryPlan plan={},int cursorKind=0)
        : db_(std::move(db)), source_(std::move(source)), bindings_(std::move(bindings)), readOnly_(snapshot), recordsetType_(cursorKind ? cursorKind : snapshot ? 4 : 2),queryPlan_(std::move(plan)) {
        data_->workspace=workspaceState(db_); data_->rollbackRevision=data_->workspace->rollbackRevision;
        table_ = db_.tables().contains(source_, Qt::CaseInsensitive);
        if(table_ && nativeTableNames(db_).contains(source_,Qt::CaseInsensitive)) tableGeneration_=data_->workspace->tableGenerations.value(source_.toCaseFolded());
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
    void checkTableGeneration() const {
        data_->workspace->check();
        if(tableGeneration_ && *tableGeneration_!=data_->workspace->tableGenerations.value(source_.toCaseFolded()))
            throw VbaRuntimeError(3420,"DAO: Physical table identity was replaced");
    }
    void checkTransaction() const {
        checkTableGeneration();
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
            if (args.empty() || (args.size()==1 && args[0].kind==VbaValue::Kind::Missing)) {
                auto collection=fields_.lock();
                if(!collection) { collection=fieldCollection(shared_from_this()).object; fields_=collection; }
                return VbaValue::reference(collection);
            }
            if(args.size()!=1) throw VbaRuntimeError(450,"DAO: Fields requires one field name/index");
            const auto value=scalarVbaValue(args[0]);
            if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Invalid Null field index");
            int position=-1;
            if(value.kind==VbaValue::Kind::Text) {
                const auto requested=text(value.asText());
                for(int i=0;i<data_->columns.size();++i) if(data_->columns[i].compare(requested,Qt::CaseInsensitive)==0) {
                    if(position>=0) throw VbaRuntimeError(3251,"DAO: Ambiguous field name"); position=i;
                }
            } else {
                const auto index=coerceVbaModuleValue(value,"long").number;
                if(index>=0 && index<data_->columns.size()) position=static_cast<int>(index);
            }
            if(position<0) throw VbaRuntimeError(3265,"DAO: Field index/name not found");
            const auto canonical=data_->columns[position];
            int matches=0; for(const auto& name:data_->columns) if(name.compare(canonical,Qt::CaseInsensitive)==0) ++matches;
            if(matches!=1) throw VbaRuntimeError(3251,"DAO: Duplicate result fields require unambiguous labels");
            auto& cached=fieldObjects_[canonical.toCaseFolded()];
            auto field=cached.lock();
            if(!field) { field=std::make_shared<Field>(shared_from_this(),canonical); cached=field; }
            return VbaValue::reference(field);
        }
        if(key=="type") {
            if(!args.empty()) throw VbaRuntimeError(450,"DAO: Recordset.Type takes no arguments");
            if(!recordsetType_) throw VbaRuntimeError(3251,"DAO: Frozen form cursor has no proven Recordset.Type");
            return VbaValue::integer(static_cast<std::int16_t>(*recordsetType_));
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
            changed_.clear();chunkFields_.clear();
            if (key == "edit") { current(); pending_ = data_->rows[position_]; mode_ = Mode::Edit; }
            else {
                pending_=QVariantList(data_->columns.size());
                for(int i=0;i<data_->columns.size();++i) {
                    const auto key=data_->columns[i].toCaseFolded();
                    if(data_->literalDefaults.contains(key)) pending_[i]=data_->literalDefaults.value(key);
                }
                const auto generators=generatedAccessColumns(db_,source_);
                for(int i=0;i<data_->columns.size();++i) if(generators.contains(data_->columns[i].toCaseFolded())) {
                    const auto reservation=reserveAccessCounter(db_,source_,data_->columns[i]);
                    pending_[i]=reservation.value;changed_.insert(i);
                    if(!data_->workspace->savepoints.isEmpty()) {
                        const auto key=reservation.table+QChar(0)+reservation.column;
                        auto& retained=data_->workspace->counterReservations[key];
                        if(reservation.jetLong || retained.table.isEmpty() || reservation.value>retained.value) retained=reservation;
                    }
                }
                mode_=Mode::Add;
            }
            return {};
        }
        if (key == "cancelupdate") { pending_.clear(); changed_.clear(); chunkFields_.clear(); mode_ = Mode::Browse; return {}; }
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
            checkTableGeneration();
            if (data_->rollbackRevision!=data_->workspace->rollbackRevision) { mode_=Mode::Browse; pending_.clear(); changed_.clear(); chunkFields_.clear(); }
            require(mode_ == Mode::Browse, "DAO: Pending edit");
            reload(); return {};
        }
        if (key == "close") { if (cloneCounted_) { --data_->workspace->clones; cloneCounted_=false; } closed_ = true; data_->cursors.remove(this); pending_.clear(); changed_.clear(); chunkFields_.clear(); return {}; }
        return value(text(name)); // implicit default Fields(name).Value
    }
    void assign(std::string_view name, const VbaValue& value) override {
        if(lower(name)=="type") {
            if(closed_) throw VbaRuntimeError(3420,"DAO: Recordset is closed");
            checkTransaction();
            throw VbaRuntimeError(383,"DAO: Recordset.Type is read-only");
        }
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
    VbaValue generatedFieldProperty(const QString& name,bool attributes) const {
        (void)fieldType(name);
        if(!table_ || readOnly_ || declaredFieldType(name).asNumber()!=4
            || !generatedAccessColumns(db_,source_).contains(name.toCaseFolded()))
            throw VbaRuntimeError(3251,"DAO: Field generator property is not proven in this context");
        // Access2003/DAO3.6 writable Long counter observation: Fixed|AutoIncr|Updatable.
        return attributes ? VbaValue::longInteger(49) : VbaValue::logical(true);
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
    VbaValue defaultExpression(const QString& name) const {
        (void)fieldType(name); const auto key=name.toCaseFolded();
        if(!data_->defaultExpressions.contains(key)) throw VbaRuntimeError(3251,"Field.DefaultValue is not proven for this source");
        return VbaValue::string(data_->defaultExpressions.value(key).toStdString());
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
    VbaValue binaryChunk(const QString& name,const VbaValue& offsetValue,const VbaValue& bytesValue) const {
        const auto type=fieldType(name);
        if(!binaryFieldType(type) || mode_!=Mode::Browse)
            throw VbaRuntimeError(3251,"GetChunk supports stored binary observations only");
        current();
        const auto offset=static_cast<qint64>(coerceVbaModuleValue(scalarVbaValue(offsetValue),"long").number);
        const auto requested=static_cast<qint64>(coerceVbaModuleValue(scalarVbaValue(bytesValue),"long").number);
        const auto& raw=data_->rows[position_][column(name)];
        if(raw.isNull()) throw VbaRuntimeError(3251,"Null GetChunk requires reference validation");
        if(raw.metaType().id()!=QMetaType::QByteArray) throw VbaRuntimeError(13,"Binary field requires BLOB storage");
        const auto bytes=raw.toByteArray();
        if(offset<0 || offset>=bytes.size() || requested<=0)
            throw VbaRuntimeError(3251,"GetChunk boundary case requires reference validation");
        const auto count=std::min(requested,static_cast<qint64>(bytes.size())-offset);
        if(count>16*1024*1024) throw VbaRuntimeError(3251,"GetChunk exceeds native output bound");
        const auto* begin=reinterpret_cast<const unsigned char*>(bytes.constData())+offset;
        return VbaValue::byteArray(std::vector<unsigned char>(begin,begin+count));
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
        const auto value=qtFieldValue(raw,type);
        return data_->daoTypes.value(name.toCaseFolded())==6 && value.kind!=VbaValue::Kind::Null
            ? coerceVbaModuleValue(value,"single") : value;
    }
    void appendBinaryChunk(const QString& name,const VbaValue& argument) {
        const auto type=fieldType(name);
        if(!binaryFieldType(type) || mode_==Mode::Browse)
            throw VbaRuntimeError(3251,"AppendChunk requires a native binary Edit/AddNew buffer");
        const auto value=scalarVbaValue(argument);
        if(value.kind!=VbaValue::Kind::ByteArray) throw VbaRuntimeError(3251,"AppendChunk currently requires Byte()");
        const int index=column(name);
        if(changed_.contains(index) && !chunkFields_.contains(index))
            throw VbaRuntimeError(3251,"Mixed binary Value assignment and AppendChunk is unsupported");
        auto bytes=chunkFields_.contains(index) ? pending_[index].toByteArray() : QByteArray("",0);
        constexpr std::size_t limit=16*1024*1024;
        if(value.bytes.size()>limit || static_cast<std::size_t>(bytes.size())>limit-value.bytes.size())
            throw VbaRuntimeError(3251,"AppendChunk exceeds native binary buffer bound");
        if(!value.bytes.empty()) bytes.append(reinterpret_cast<const char*>(value.bytes.data()),static_cast<qsizetype>(value.bytes.size()));
        pending_[index]=bytes; changed_.insert(index); chunkFields_.insert(index);
    }
    void set(const QString& name, const VbaValue& value) {
        checkTransaction();
        if (closed_) throw VbaRuntimeError(3420,"DAO: Recordset is closed");
        require(mode_ != Mode::Browse, "DAO: Edit or AddNew required");
        const auto index = column(name);
        if(mode_==Mode::Add && value.kind==VbaValue::Kind::Null && !pending_[index].isNull()
            && generatedAccessColumns(db_,source_).contains(name.toCaseFolded())) return;
        if(mode_==Mode::Edit && table_ && generatedAccessColumns(db_,source_).contains(name.toCaseFolded()))
            throw VbaRuntimeError(3164,"DAO: AutoNumber field cannot be updated");
        const auto type=data_->types.value(name.toCaseFolded());
        if(chunkFields_.contains(index)) throw VbaRuntimeError(3251,"Mixed AppendChunk and binary Value assignment is unsupported");
        if(value.kind==VbaValue::Kind::ByteArray) {
            if(!binaryFieldType(type)) throw VbaRuntimeError(13,"Byte array requires a binary field");
            if(value.bytes.size()>static_cast<std::size_t>(std::numeric_limits<qsizetype>::max()))
                throw VbaRuntimeError(6,"Binary field length overflow");
            pending_[index]=value.bytes.empty() ? QByteArray("",0)
                : QByteArray(reinterpret_cast<const char*>(value.bytes.data()),static_cast<qsizetype>(value.bytes.size()));
        } else {
            if(binaryFieldType(type) && value.kind!=VbaValue::Kind::Null)
                throw VbaRuntimeError(3251,"Binary field assignment requires Byte() or Null");
            const auto scalar=data_->daoTypes.value(name.toCaseFolded())==6 ? scalarVbaValue(value) : value;
            const auto converted=data_->daoTypes.value(name.toCaseFolded())==6 && scalar.kind!=VbaValue::Kind::Null
                ? coerceVbaModuleValue(scalar,"single") : scalar;
            pending_[index]=sqlFieldValue(converted,type);
        }
        changed_.insert(index);
    }
private:
    struct CloneTag {};
    Recordset(const Recordset& source,CloneTag)
        : data_(source.data_),db_(source.db_),source_(source.source_),bindings_(source.bindings_),
          table_(source.table_),readOnly_(source.readOnly_),frozen_(source.frozen_),recordsetType_(source.recordsetType_),tableGeneration_(source.tableGeneration_),position_(-1),queryPlan_(source.queryPlan_) {
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
        data_->types.clear(); data_->daoTypes.clear(); data_->definitionSizes.clear(); data_->literalDefaults.clear(); data_->defaultExpressions.clear();
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
                if(const auto expression=nativeFieldDefaultExpression(field)) data_->defaultExpressions.insert(key,*expression);
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
        mode_ = Mode::Browse; pending_.clear(); changed_.clear(); chunkFields_.clear();
        reload(false); position_ = data_->ids.indexOf(id);
        if(position_>=0) lastModified_=bookmarkAt(position_); else lastModified_.reset();
    }
    std::shared_ptr<RecordsetData> data_{std::make_shared<RecordsetData>()};
    QSqlDatabase db_;
    QString source_;
    QMap<QString,QVariant> bindings_;
    bool table_{}, readOnly_{}, closed_{}, frozen_{},noMatch_{},cloneCounted_{};
    std::optional<int> recordsetType_; // Opening kind, independent of write capability; unknown for host-only frozen adapters.
    std::optional<quint64> tableGeneration_;
    std::optional<VbaValue> lastModified_;
    QVariantList pending_;
    QSet<int> changed_,chunkFields_;
    qsizetype position_{};
    Mode mode_{Mode::Browse};
    AccessQueryPlan queryPlan_;
    std::weak_ptr<VbaObject> fields_;
    QMap<QString,std::weak_ptr<Field>> fieldObjects_;
};
class Fields final : public VbaObject {
public:
    std::string_view className() const override { return "Fields"; }
    bool isNativeType(std::string_view type) const override { return type.starts_with("dao.") && type.substr(4)==lower(className()); }
    explicit Fields(std::shared_ptr<Recordset> owner) : owner_(std::move(owner)) {}
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        if(lower(name)=="item") return VbaArgumentSignature{{"Index"},1}; return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        const auto key = lower(name);
        if(key=="count") {
            if(!args.empty()) throw VbaRuntimeError(450,"DAO: Fields.Count takes no arguments");
            return integerCount(owner_->fieldNames().size());
        }
        if(key=="item") {
            if(args.size()!=1 || args[0].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(450,"DAO: Fields.Item requires one index");
            return owner_->member("Fields",args);
        }
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
            values.push_back(owner_->member("Fields",{VbaValue::string(name.toStdString())}));
        return values;
    }
private:
    std::shared_ptr<Recordset> owner_;
};
VbaValue fieldCollection(std::shared_ptr<Recordset> owner) {
    return VbaValue::reference(std::make_shared<Fields>(std::move(owner)));
}

VbaValue Field::member(std::string_view name, const std::vector<VbaValue>& args) {
    if(lower(name)=="appendchunk") {
        if(args.size()!=1) throw VbaRuntimeError(450,"Field.AppendChunk requires Val");
        owner_->appendBinaryChunk(name_,args[0]); return {};
    }
    if(lower(name)=="getchunk") {
        if(args.size()!=2) throw VbaRuntimeError(450,"Field.GetChunk requires Offset and Bytes");
        return owner_->binaryChunk(name_,args[0],args[1]);
    }
    if(lower(name)=="attributes" || lower(name)=="dataupdatable") {
        if(!args.empty()) throw VbaRuntimeError(450,"Field generator property requires no arguments");
        return owner_->generatedFieldProperty(name_,lower(name)=="attributes");
    }
    if(lower(name)=="fieldsize") {
        if(!args.empty()) throw VbaRuntimeError(450,"Field.FieldSize requires no arguments");
        return owner_->fieldSize(name_);
    }
    if(lower(name)=="ordinalposition") {
        if(!args.empty()) throw VbaRuntimeError(450,"Field.OrdinalPosition requires no arguments");
        return owner_->fieldOrdinal(name_);
    }
    if(lower(name)=="defaultvalue") {
        if(!args.empty()) throw VbaRuntimeError(450,"Field.DefaultValue requires no arguments");
        return owner_->defaultExpression(name_);
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
    if (lower(name) == "name") {
        (void)owner_->fieldType(name_);
        return VbaValue::string(name_.toUtf8().toStdString());
    }
    throw std::runtime_error("DAO: Unknown Field member");
}
void Field::assign(std::string_view name, const VbaValue& value) {
    if(lower(name)=="defaultvalue") { (void)owner_->fieldType(name_); throw VbaRuntimeError(383,"Recordset Field.DefaultValue is read-only"); }
    if(lower(name)=="size") { (void)owner_->fieldType(name_); throw VbaRuntimeError(383,"Field.Size is read-only"); }
    if(lower(name)=="attributes" || lower(name)=="dataupdatable") {
        (void)owner_->generatedFieldProperty(name_,lower(name)=="attributes");
        throw VbaRuntimeError(383,"Field generator property is read-only");
    }
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
        if (key=="fields" || key=="parameters") return VbaArgumentSignature{{"Index"},0};
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
        if(key=="fields") {
            auto collection=fields();
            return args.empty() || (args.size()==1 && args[0].kind==VbaValue::Kind::Missing)
                ? VbaValue::reference(collection) : collection->member("Item",args);
        }
        if (key == "parameters") {
            auto collection=parameters();
            return args.empty() || (args.size()==1 && args[0].kind==VbaValue::Kind::Missing)
                ? VbaValue::reference(collection) : collection->member("Item",args);
        }
        if (key == "openrecordset") {
            if(plan_.sql.isEmpty()) throw VbaRuntimeError(3251,"DAO: QueryDef SQL is not set");
            if(args.size()>1) throw VbaRuntimeError(450,"DAO: QueryDef OpenRecordset accepts one mode");
            const int mode=nativeRecordsetMode(args,0);
            require(plan_.sql.trimmed().startsWith("SELECT ",Qt::CaseInsensitive),"DAO: SELECT QueryDef required");
            return VbaValue::reference(std::make_shared<Recordset>(db_,plan_.sql,true,bindAccessQueryParameters(plan_,values_),plan_,mode));
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
        if(key=="type") {
            if(!name_.isEmpty() && !owned_) throw VbaRuntimeError(3251,"DAO: Imported QueryDef type metadata is unavailable");
            if(plan_.daoQueryType<0) throw VbaRuntimeError(3251,"DAO: QueryDef type is not proven for this source");
            return VbaValue::integer(static_cast<std::int16_t>(plan_.daoQueryType));
        }
        if (key == "recordsaffected") return longCount(affected_);
        if (key == "close") { closed_ = true; return {}; }
        throw VbaRuntimeError(438,"DAO: QueryDef member unsupported");
    }
    void assign(std::string_view name, const VbaValue& value) override {
        refresh();
        if(lower(name)=="name") throw VbaRuntimeError(3251,"DAO: QueryDef renaming is unsupported");
        if(lower(name)=="type") throw VbaRuntimeError(383,"DAO: QueryDef.Type is read-only");
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
    std::shared_ptr<VbaObject> fields();
    quint64 fieldSchemaVersion() const {
        workspace_->check(); QSqlQuery query(db_);
        if(!query.exec("PRAGMA main.schema_version") || !query.next()) queryFailure(query);
        return query.value(0).toULongLong();
    }
    QVector<TableSchemaColumn> queryFieldColumns() const {
        const auto fresh=source_.trimmed().isEmpty() ? AccessQueryPlan{} : compileAccessQuery(source_);
        if(!fresh.makeTable.isEmpty() || !fresh.sql.trimmed().startsWith("SELECT",Qt::CaseInsensitive))
            throw VbaRuntimeError(3251,"DAO: QueryDef.Fields requires a supported SELECT definition");
        QVector<QString> types; const auto names=accessQueryColumns(db_,fresh,&types);
        QVector<TableSchemaColumn> columns; QSet<QString> seen;
        for(int i=0;i<names.size();++i) {
            if(names[i].isEmpty() || seen.contains(names[i].toCaseFolded()))
                throw VbaRuntimeError(3251,"DAO: QueryDef field names are empty or ambiguous");
            seen.insert(names[i].toCaseFolded());
            TableSchemaColumn column; column.name=names[i]; column.type=types.value(i); columns.append(std::move(column));
        }
        return columns;
    }
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
    std::weak_ptr<VbaObject> fields_;
    QVector<TableSchemaColumn> fieldColumns_;
    quint64 fieldGeneration_{},fieldSchema_{},fieldRollback_{};

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
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        if(lower(name)=="item") return VbaArgumentSignature{{"Index"},1}; return {};
    }
    explicit QueryParameters(std::shared_ptr<QueryDefinition> query) : query_(std::move(query)) {}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        query_->refresh();
        const auto key=lower(name);
        if (key=="count") {
            if (!args.empty()) throw VbaRuntimeError(450,"DAO: Count takes no arguments");
            return integerCount(query_->plan_.parameters.size());
        }
        if(key=="refresh") {
            if(!args.empty()) throw VbaRuntimeError(450,"DAO: Refresh takes no arguments");
            return {};
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
    if (!query.exec("SELECT name FROM main.sqlite_schema WHERE type='table' ORDER BY rowid")) queryFailure(query);
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
        if(key=="count" && args.empty()) return integerCount(names().size());
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
    TableSchemaField(TableSchemaReader reader,QString name,bool defaults=true) : reader_(std::move(reader)),name_(std::move(name)),defaults_(defaults) {}
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
        if(key=="attributes" && columns[position].indexAttributes>=0) return VbaValue::longInteger(columns[position].indexAttributes);
        if(key=="foreignname") {
            if(columns[position].foreignName.isNull()) throw VbaRuntimeError(3251,"ForeignName is proven only for Relation fields");
            return VbaValue::string(columns[position].foreignName.toStdString());
        }
        if(key=="ordinalposition") return integerCount(position);
        if(key=="defaultvalue") {
            if(!defaults_) throw VbaRuntimeError(3251,"QueryDef field default is not proven");
            const auto expression=nativeFieldDefaultExpression(columns[position]);
            if(!expression) throw VbaRuntimeError(3251,"Schema Field.DefaultValue is not proven");
            return VbaValue::string(expression->toStdString());
        }
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
private: TableSchemaReader reader_; QString name_; bool defaults_{};
};
class TableSchemaFields final : public VbaObject {
public:
    explicit TableSchemaFields(TableSchemaReader reader,bool defaults=true) : reader_(std::move(reader)),defaults_(defaults) {}
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
        if(!field) { field=std::make_shared<TableSchemaField>(reader_,column,defaults_); cached=field; }
        return VbaValue::reference(field);
    }
    VbaValue item(std::string_view name) override { return member("Item",{VbaValue::string(std::string(name))}); }
    std::vector<VbaValue> enumerate() override {
        const auto columns=reader_(); std::vector<VbaValue> result;
        for(const auto& column:columns) result.push_back(member("Item",{VbaValue::string(column.name.toStdString())}));
        return result;
    }
private: TableSchemaReader reader_; QMap<QString,std::weak_ptr<TableSchemaField>> fields_; bool defaults_{};
};
bool sameQueryFieldColumns(const QVector<TableSchemaColumn>& left,const QVector<TableSchemaColumn>& right) {
    if(left.size()!=right.size()) return false;
    for(int i=0;i<left.size();++i) if(left[i].name!=right[i].name || left[i].type!=right[i].type) return false;
    return true;
}
std::shared_ptr<VbaObject> QueryDefinition::fields() {
    const auto version=fieldSchemaVersion(),rollback=workspace_->rollbackRevision;
    const auto columns=queryFieldColumns(); auto collection=fields_.lock();
    if(!collection || fieldGeneration_!=generation_ || fieldSchema_!=version || fieldRollback_!=rollback
        || !sameQueryFieldColumns(fieldColumns_,columns)) {
        const auto self=shared_from_this(); const auto generation=generation_;
        collection=std::make_shared<TableSchemaFields>([self,generation,version,rollback,columns] {
            self->alive();
            if(self->generation_!=generation || self->workspace_->rollbackRevision!=rollback || self->fieldSchemaVersion()!=version
                || (self->owned_ && self->ownedSource()!=self->source_))
                throw VbaRuntimeError(3420,"DAO: QueryDef.Fields metadata was invalidated");
            const auto fresh=self->queryFieldColumns();
            if(!sameQueryFieldColumns(columns,fresh)) throw VbaRuntimeError(3420,"DAO: QueryDef field metadata changed");
            return fresh;
        },false);
        fields_=collection; fieldColumns_=columns; fieldGeneration_=generation_; fieldSchema_=version; fieldRollback_=rollback;
    }
    return collection;
}
struct NativeIndexData {
    AccessSourceIndex source;
    QByteArray definition;
    bool operator==(const NativeIndexData&) const = default;
};
QVector<NativeIndexData> readNativeIndexCatalog(QSqlDatabase db,const QString& table) {
    workspaceState(db)->check(); QSqlQuery query(db);
    if(!query.exec("SELECT type FROM main.sqlite_schema WHERE name='_la_access_indexes'")) queryFailure(query);
    if(!query.next() || query.value(0).toString()!="table" || query.next()) throw VbaRuntimeError(3251,"DAO: Original index identities are unavailable"); query.finish();
    if(!query.exec("PRAGMA main.table_info('_la_access_indexes')")) queryFailure(query); int shape=0;
    const QStringList expected{"table_name","definition"};
    while(query.next()) if(shape>=expected.size() || query.value(1).toString()!=expected[shape++]) throw VbaRuntimeError(3251,"DAO: Index catalog shape invalid");
    if(shape!=2) throw VbaRuntimeError(3251,"DAO: Index catalog shape invalid"); query.finish();
    query.prepare("SELECT definition FROM main._la_access_indexes WHERE table_name=? COLLATE NOCASE"); query.addBindValue(table);
    if(!query.exec()) queryFailure(query);
    if(!query.next() || query.value(0).isNull()) throw VbaRuntimeError(3251,"DAO: Original index identities are unavailable");
    const auto definition=query.value(0).toString().toUtf8();
    if(query.next()) throw VbaRuntimeError(3251,"DAO: Original index definitions are ambiguous"); query.finish();
    QStringList columns,primary; QMap<QString,bool> required; QMap<QString,QString> types; QMap<int,QString> keyOrder;
    if(!query.exec("PRAGMA main.table_info("+quote(table)+")")) queryFailure(query);
    while(query.next()) {
        const auto name=query.value(1).toString(); columns.append(name); types.insert(name,query.value(2).toString()); required.insert(name,query.value(3).toBool());
        if(query.value(5).toInt()>0) keyOrder.insert(query.value(5).toInt(),name);
    }
    for(auto it=keyOrder.cbegin();it!=keyOrder.cend();++it) { if(it.key()!=primary.size()+1) throw VbaRuntimeError(3251,"DAO: Primary key order invalid"); primary.append(it.value()); }
    QVector<AccessSourceIndex> source;
    try { source=checkedAccessSourceIndexes(definition,columns); }
    catch(const std::exception&) { throw VbaRuntimeError(3251,"DAO: Original index definitions cannot be verified"); }
    struct Actual { QString name,origin; bool unique{}; QStringList columns; QVector<bool> descending; };
    QVector<Actual> actual;
    if(!query.exec("PRAGMA main.index_list("+quote(table)+")")) queryFailure(query);
    while(query.next()) {
        if(actual.size()>=4096) throw VbaRuntimeError(3251,"DAO: Index catalog bound exceeded");
        if(query.value(4).toBool()) throw VbaRuntimeError(3251,"DAO: Partial indexes are unsupported");
        Actual index; index.name=query.value(1).toString(); index.unique=query.value(2).toBool(); index.origin=query.value(3).toString();
        QSqlQuery fields(db); if(!fields.exec("PRAGMA main.index_xinfo("+quote(index.name)+")")) queryFailure(fields);
        while(fields.next()) if(fields.value(5).toBool()) {
            if(fields.value(1).toInt()<0 || fields.value(4).toString()!="BINARY") throw VbaRuntimeError(3251,"DAO: Expression/collated index metadata unsupported");
            index.columns.append(fields.value(2).toString()); index.descending.append(fields.value(3).toBool());
        }
        actual.append(index);
    }
    QVector<NativeIndexData> result; int primaryCount=0;
    for(const auto& index:source) {
        if(index.kind==2 || (index.flags&~(128|8|2|1)) || (index.flags&2)) throw VbaRuntimeError(3251,"DAO: Logical/unknown/IgnoreNulls index semantics unsupported");
        if(index.flags&8) for(const auto& field:index.columns) if(!required.value(field)) throw VbaRuntimeError(3251,"DAO: Required index constraint is unproven");
        int found=-1;
        if(index.kind==1) { ++primaryCount; if(index.columns!=primary) throw VbaRuntimeError(3251,"DAO: Original primary key does not match table"); }
        for(int i=0;i<actual.size();++i) {
            const auto& native=actual[i];
            const bool identity=index.kind==1 ? native.origin=="pk" : native.origin=="c" && native.name==table+"_"+index.name;
            if(identity && native.unique==((index.flags&1)!=0) && native.columns==index.columns && native.descending==index.descending) {
                if(found>=0) throw VbaRuntimeError(3251,"DAO: Native index identity ambiguous"); found=i;
            }
        }
        if(found>=0) actual.removeAt(found);
        else if(!(index.kind==1 && index.columns.size()==1 && !index.descending[0] && types.value(index.columns[0]).compare("INTEGER",Qt::CaseInsensitive)==0))
            throw VbaRuntimeError(3251,"DAO: Original index does not match native constraint");
        result.append({index,definition});
    }
    if(!actual.isEmpty() || primaryCount!=(primary.isEmpty() ? 0 : 1)) throw VbaRuntimeError(3251,"DAO: Native index catalog is incomplete");
    return result;
}
using NativeIndexReader=std::function<QVector<NativeIndexData>()>;
class NativeIndex final : public VbaObject,public std::enable_shared_from_this<NativeIndex> {
public:
    NativeIndex(NativeIndexReader reader,NativeIndexData data) : reader_(std::move(reader)),data_(std::move(data)) {}
    std::string_view className() const override { return "Index"; }
    bool isNativeType(std::string_view type) const override { return type=="dao.index"; }
    bool matches(const NativeIndexData& data) const { return data_==data; }
    void check() const {
        for(const auto& row:reader_()) if(row.source.name.compare(data_.source.name,Qt::CaseInsensitive)==0 && row==data_) return;
        throw VbaRuntimeError(3420,"DAO: Index definition changed");
    }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        if(lower(name)=="fields") return VbaArgumentSignature{{"Index"},0}; return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        check(); const auto key=lower(name);
        if(key=="fields") {
            if(args.size()>1) throw VbaRuntimeError(450,"DAO: Index.Fields takes at most one index");
            auto fields=fields_.lock();
            if(!fields) {
                const auto self=shared_from_this(); fields=std::make_shared<TableSchemaFields>([self] {
                    self->check(); QVector<TableSchemaColumn> result;
                    for(int i=0;i<self->data_.source.columns.size();++i) {
                        TableSchemaColumn field; field.name=self->data_.source.columns[i]; field.indexAttributes=self->data_.source.descending[i] ? 1 : 0; result.append(field);
                    }
                    return result;
                },false); fields_=fields;
            }
            return args.empty() || args[0].kind==VbaValue::Kind::Missing ? VbaValue::reference(fields) : fields->member("Item",args);
        }
        if(!args.empty()) throw VbaRuntimeError(450,"DAO: Index property takes no arguments");
        if(key=="name") return VbaValue::string(data_.source.name.toStdString());
        if(key=="primary") return VbaValue::logical(data_.source.kind==1);
        if(key=="unique") return VbaValue::logical(data_.source.flags&1);
        if(key=="required") return VbaValue::logical(data_.source.flags&8);
        if(key=="ignorenulls") return VbaValue::logical(data_.source.flags&2);
        throw VbaRuntimeError(3251,"DAO: Index property unsupported");
    }
    void assign(std::string_view,const VbaValue&) override { check(); throw VbaRuntimeError(383,"DAO: Imported index metadata is read-only"); }
private:
    NativeIndexReader reader_; NativeIndexData data_; std::weak_ptr<VbaObject> fields_;
};
class NativeIndexes final : public VbaObject {
public:
    explicit NativeIndexes(NativeIndexReader reader) : reader_(std::move(reader)) {}
    std::string_view className() const override { return "Indexes"; }
    bool isNativeType(std::string_view type) const override { return type=="dao.indexes"; }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        if(lower(name)=="item") return VbaArgumentSignature{{"Index"},1}; return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        const auto rows=reader_(); const auto key=lower(name);
        if(key=="count" || key=="refresh") {
            if(!args.empty()) throw VbaRuntimeError(450,"DAO: Index collection property takes no arguments");
            return key=="count" ? integerCount(rows.size()) : VbaValue{};
        }
        if(key!="item") throw VbaRuntimeError(3251,"DAO: Index catalog mutation unsupported");
        if(args.size()!=1 || args[0].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(450,"DAO: Index required");
        const auto value=scalarVbaValue(args[0]); if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Invalid Null index");
        int position=-1;
        if(value.kind==VbaValue::Kind::Text) {
            const auto requested=text(value.asText()); for(int i=0;i<rows.size();++i) if(rows[i].source.name.compare(requested,Qt::CaseInsensitive)==0) position=i;
        } else { const auto index=coerceVbaModuleValue(value,"long").number; if(index>=0 && index<rows.size()) position=static_cast<int>(index); }
        if(position<0) throw VbaRuntimeError(3265,"DAO: Index not found");
        const auto& row=rows[position]; auto& cached=objects_[row.source.name.toCaseFolded()]; auto object=cached.lock();
        if(!object || !object->matches(row)) { object=std::make_shared<NativeIndex>(reader_,row); cached=object; }
        return VbaValue::reference(object);
    }
    VbaValue item(std::string_view name) override { return member("Item",{VbaValue::string(std::string(name))}); }
    std::vector<VbaValue> enumerate() override {
        const auto rows=reader_(); std::vector<VbaValue> result;
        for(const auto& row:rows) result.push_back(member("Item",{VbaValue::string(row.source.name.toStdString())})); return result;
    }
private:
    NativeIndexReader reader_; QMap<QString,std::weak_ptr<NativeIndex>> objects_;
};

class TableDefinition final : public VbaObject, public std::enable_shared_from_this<TableDefinition> {
public:
    std::string_view className() const override { return "TableDef"; }
    bool isNativeType(std::string_view type) const override { return type.starts_with("dao.") && type.substr(4)==lower(className()); }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        const auto key=lower(name);
        if (key=="fields" || key=="indexes") return VbaArgumentSignature{{"Index"},0};
        if (key=="openrecordset") return VbaArgumentSignature{{"Type"},0};
        return {};
    }
    TableDefinition(QSqlDatabase db,QString name) : db_(std::move(db)),name_(std::move(name)),workspace_(workspaceState(db_)),generation_(workspace_->tableGenerations.value(name_.toCaseFolded())) {}
    bool valid() {
        workspace_->check();
        if (invalidated_ || generation_!=workspace_->tableGenerations.value(name_.toCaseFolded()) || !db_.isOpen() || !nativeTableNames(db_).contains(name_,Qt::CaseInsensitive)) invalidated_=true;
        return !invalidated_;
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        if (!valid()) throw VbaRuntimeError(3420,"DAO: TableDef is unavailable");
        const auto key=lower(name);
        if (key=="name") {
            if (!args.empty()) throw VbaRuntimeError(450,"DAO: Name takes no arguments");
            return VbaValue::string(name_.toStdString());
        }
        if(key=="recordcount") {
            if(!args.empty()) throw VbaRuntimeError(450,"DAO: RecordCount takes no arguments");
            SqliteExecutionBudget budget(db_); QSqlQuery count(db_);
            if(!count.exec("SELECT COUNT(*) FROM main."+quote(name_)) || !count.next()) queryFailure(count);
            return longCount(count.value(0).toLongLong());
        }
        if(key=="indexes") {
            if(args.size()>1) throw VbaRuntimeError(450,"DAO: TableDef.Indexes takes at most one index");
            const auto version=schemaVersion(),rollback=workspace_->rollbackRevision; auto indexes=indexes_.lock();
            if(!indexes || version!=indexesVersion_ || rollback!=indexesRollback_) {
                const auto self=shared_from_this();
                indexes=std::make_shared<NativeIndexes>([self,version,rollback] {
                    if(!self->valid() || self->schemaVersion()!=version || self->workspace_->rollbackRevision!=rollback)
                        throw VbaRuntimeError(3420,"DAO: Index metadata invalidated; reacquire TableDef.Indexes");
                    return readNativeIndexCatalog(self->db_,self->name_);
                }); indexes_=indexes; indexesVersion_=version; indexesRollback_=rollback;
            }
            return args.empty() || args[0].kind==VbaValue::Kind::Missing ? VbaValue::reference(indexes) : indexes->member("Item",args);
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
            const int mode=nativeRecordsetMode(args,0);
            return VbaValue::reference(std::make_shared<Recordset>(db_,name_,mode==4));
        }
        throw VbaRuntimeError(438,"DAO: TableDef member unsupported");
    }
    void assign(std::string_view name,const VbaValue& value) override {
        if(!valid()) throw VbaRuntimeError(3420,"DAO: TableDef is unavailable");
        if(lower(name)=="recordcount") throw VbaRuntimeError(383,"DAO: TableDef.RecordCount is read-only");
        VbaObject::assign(name,value);
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
    std::weak_ptr<VbaObject> indexes_;
    quint64 indexesVersion_{},indexesRollback_{};
    bool invalidated_{};
    quint64 generation_{};
};
class TableDefinitions final : public VbaObject {
public:
    std::string_view className() const override { return "TableDefs"; }
    bool isNativeType(std::string_view type) const override { return type.starts_with("dao.") && type.substr(4)==lower(className()); }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        if(lower(name)=="item") return VbaArgumentSignature{{"Index"},1}; return {};
    }
    explicit TableDefinitions(QSqlDatabase db) : db_(std::move(db)),workspace_(workspaceState(db_)) {}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        workspace_->check(); const auto names=nativeTableNames(db_); const auto key=lower(name);
        if (key=="count") {
            if (!args.empty()) throw VbaRuntimeError(450,"DAO: Count takes no arguments");
            return integerCount(names.size());
        }
        if (key=="refresh") {
            if(!args.empty()) throw VbaRuntimeError(450,"DAO: Refresh takes no arguments");
            for(auto it=definitions_.begin();it!=definitions_.end();) {
                const auto object=it.value().lock();
                if(!object || !object->valid()) it=definitions_.erase(it); else ++it;
            }
            return {};
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
        workspace_->check(); const auto names=nativeTableNames(db_); std::vector<VbaValue> result;
        for (const auto& name:names) result.push_back(member("Item",{VbaValue::string(name.toStdString())}));
        return result;
    }
private:
    QSqlDatabase db_; std::shared_ptr<WorkspaceState> workspace_;
    QMap<QString,std::weak_ptr<TableDefinition>> definitions_;
};
struct NativeRelationData {
    QString name,source,parent,child;
    quint32 attributes{};
    QVector<QPair<QString,QString>> columns; // primary Name, foreign ForeignName
    bool operator==(const NativeRelationData&) const = default;
};
struct NativeRelationCatalog {
    QSqlDatabase db; std::shared_ptr<WorkspaceState> workspace;
    explicit NativeRelationCatalog(QSqlDatabase database) : db(std::move(database)),workspace(workspaceState(db)) {}
    quint64 version() const {
        workspace->check(); QSqlQuery query(db);
        if(!query.exec("PRAGMA main.schema_version") || !query.next()) queryFailure(query);
        return query.value(0).toULongLong();
    }
    QVector<NativeRelationData> read() const {
        workspace->check();
        QMap<QString,QString> tables;
        for(const auto& name:nativeTableNames(db)) {
            const auto key=name.toCaseFolded(); if(tables.contains(key)) throw VbaRuntimeError(3251,"DAO: Relation table names ambiguous"); tables.insert(key,name);
        }
        const auto table=[&](const QString& name) {
            if(name.isEmpty() || name.contains(QChar(0)) || !name.isValidUtf16() || !tables.contains(name.toCaseFolded())) throw VbaRuntimeError(3251,"DAO: Relation table is unavailable");
            return tables.value(name.toCaseFolded());
        };
        const auto field=[&](const QString& owner,const QString& name) {
            if(name.isEmpty() || name.contains(QChar(0)) || !name.isValidUtf16()) throw VbaRuntimeError(3251,"DAO: Relation field invalid");
            QSqlQuery query(db); if(!query.exec("PRAGMA main.table_info("+quote(owner)+")")) queryFailure(query);
            QString result; int matches=0;
            while(query.next()) if(query.value(1).toString().compare(name,Qt::CaseInsensitive)==0) { result=query.value(1).toString(); ++matches; }
            if(matches!=1) throw VbaRuntimeError(3251,"DAO: Relation field is unavailable or ambiguous");
            return result;
        };
        struct Key { QString parent,child,update,erase; QMap<int,QPair<QString,QString>> columns; };
        QVector<Key> actual;
        for(const auto& child:tables) {
            QSqlQuery query(db); if(!query.exec("PRAGMA main.foreign_key_list("+quote(child)+")")) queryFailure(query);
            QMap<int,Key> groups; int count=0;
            while(query.next()) {
                if(++count>1000000) throw VbaRuntimeError(3251,"DAO: Relation catalog bound exceeded");
                const int id=query.value(0).toInt(),ordinal=query.value(1).toInt(); auto& key=groups[id];
                const auto parent=table(query.value(2).toString()),update=query.value(5).toString(),erase=query.value(6).toString();
                if(!key.parent.isEmpty() && (key.parent!=parent || key.update!=update || key.erase!=erase)) throw VbaRuntimeError(3251,"DAO: Foreign key metadata inconsistent");
                if(ordinal<0 || key.columns.contains(ordinal)) throw VbaRuntimeError(3251,"DAO: Foreign key ordinal invalid");
                key.parent=parent; key.child=child; key.update=update; key.erase=erase;
                key.columns.insert(ordinal,{field(parent,query.value(4).toString()),field(child,query.value(3).toString())});
            }
            for(const auto& key:groups) actual.append(key);
        }
        QMap<QString,NativeRelationData> definitions;
        QSqlQuery query(db);
        if(!query.exec("SELECT type FROM main.sqlite_schema WHERE name='_la_access_relations'")) queryFailure(query);
        const bool stored=query.next();
        if(stored && (query.value(0).toString()!="table" || query.next())) throw VbaRuntimeError(3251,"DAO: Relation catalog shape invalid"); query.finish();
        if(stored) {
            const QStringList expected{"source_file","name","flags","ordinal","child_table","child_column","parent_table","parent_column"};
            if(!query.exec("PRAGMA main.table_info('_la_access_relations')")) queryFailure(query); int position=0;
            while(query.next()) { if(position>=expected.size() || query.value(1).toString()!=expected[position++]) throw VbaRuntimeError(3251,"DAO: Relation catalog shape invalid"); }
            if(position!=expected.size()) throw VbaRuntimeError(3251,"DAO: Relation catalog shape invalid"); query.finish();
            if(!query.exec("SELECT source_file,name,flags,ordinal,child_table,child_column,parent_table,parent_column FROM main._la_access_relations ORDER BY name,ordinal")) queryFailure(query);
            int rows=0; constexpr quint32 allowed=1u|2u|4u|256u|4096u|16777216u|33554432u;
            while(query.next()) {
                if(++rows>1000000) throw VbaRuntimeError(3251,"DAO: Relation catalog bound exceeded");
                const auto name=query.value(1).toString();
                if(name.isEmpty() || name.contains(QChar(0)) || !name.isValidUtf16()) throw VbaRuntimeError(3251,"DAO: Relation name invalid");
                bool flagOk=false,ordinalOk=false; const auto flags=query.value(2).toLongLong(&flagOk),ordinal=query.value(3).toLongLong(&ordinalOk);
                if(!flagOk || !ordinalOk || flags<0 || flags>std::numeric_limits<std::uint32_t>::max() || (static_cast<quint32>(flags)&~allowed)
                    || query.value(2).toDouble()!=static_cast<double>(flags) || query.value(3).toDouble()!=static_cast<double>(ordinal) || ordinal<0 || ordinal>=32767)
                    throw VbaRuntimeError(3251,"DAO: Relation flags/ordinal unsupported");
                const auto parent=table(query.value(6).toString()),child=table(query.value(4).toString()),source=query.value(0).toString();
                auto& item=definitions[name.toCaseFolded()];
                if(item.name.isEmpty()) { item.name=name; item.parent=parent; item.child=child; item.source=source; item.attributes=static_cast<quint32>(flags); }
                if(item.name!=name || item.parent!=parent || item.child!=child || item.source!=source || item.attributes!=static_cast<quint32>(flags) || ordinal!=item.columns.size())
                    throw VbaRuntimeError(3251,"DAO: Relation definitions are ambiguous or incomplete");
                const auto pair=qMakePair(field(parent,query.value(7).toString()),field(child,query.value(5).toString()));
                for(const auto& previous:item.columns) if(previous.first.compare(pair.first,Qt::CaseInsensitive)==0 || previous.second.compare(pair.second,Qt::CaseInsensitive)==0)
                    throw VbaRuntimeError(3251,"DAO: Relation fields duplicated");
                item.columns.append(pair);
            }
            if(query.lastError().isValid()) queryFailure(query);
        }
        QVector<NativeRelationData> result;
        for(const auto& item:definitions) {
            if(!(item.attributes&2)) {
                int found=-1;
                for(int i=0;i<actual.size();++i) {
                    const auto& key=actual[i]; QVector<QPair<QString,QString>> columns; int ordinal=0;
                    for(auto it=key.columns.cbegin();it!=key.columns.cend();++it) { if(it.key()!=ordinal++) throw VbaRuntimeError(3251,"DAO: Foreign key ordinal gap"); columns.append(it.value()); }
                    if(key.parent==item.parent && key.child==item.child && columns==item.columns
                        && key.update==((item.attributes&256) ? "CASCADE" : "NO ACTION") && key.erase==((item.attributes&4096) ? "CASCADE" : "NO ACTION")) { found=i; break; }
                }
                if(found<0) throw VbaRuntimeError(3251,"DAO: Declared relation does not match an enforced native key"); actual.removeAt(found);
            }
            result.append(item);
        }
        if(!actual.isEmpty()) throw VbaRuntimeError(3251,"DAO: Native foreign keys without proven relation names are unsupported");
        return result;
    }
};
class NativeRelation final : public VbaObject, public std::enable_shared_from_this<NativeRelation> {
public:
    NativeRelation(std::shared_ptr<NativeRelationCatalog> catalog,NativeRelationData data,quint64 version,quint64 rollback)
        : catalog_(std::move(catalog)),data_(std::move(data)),version_(version),rollback_(rollback) {}
    std::string_view className() const override { return "Relation"; }
    bool isNativeType(std::string_view type) const override { return type=="dao.relation"; }
    bool matches(const NativeRelationData& data,quint64 version,quint64 rollback) const { return data_==data && version_==version && rollback_==rollback; }
    void check() const {
        if(catalog_->version()!=version_ || catalog_->workspace->rollbackRevision!=rollback_) throw VbaRuntimeError(3420,"DAO: Relation metadata invalidated");
        for(const auto& current:catalog_->read()) if(current.name.compare(data_.name,Qt::CaseInsensitive)==0) {
            if(current==data_) return; break;
        }
        throw VbaRuntimeError(3420,"DAO: Relation definition changed");
    }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        if(lower(name)=="fields") return VbaArgumentSignature{{"Index"},0}; return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        check(); const auto key=lower(name);
        if(key=="fields") {
            if(args.size()>1) throw VbaRuntimeError(450,"DAO: Relation.Fields takes one index");
            auto fields=fields_.lock();
            if(!fields) {
                const auto self=shared_from_this();
                fields=std::make_shared<TableSchemaFields>([self] {
                    self->check(); QVector<TableSchemaColumn> columns;
                    for(const auto& pair:self->data_.columns) { TableSchemaColumn field; field.name=pair.first; field.foreignName=pair.second; columns.append(field); }
                    return columns;
                },false); fields_=fields;
            }
            return args.empty() || args[0].kind==VbaValue::Kind::Missing ? VbaValue::reference(fields) : fields->member("Item",args);
        }
        if(!args.empty()) throw VbaRuntimeError(450,"DAO: Relation property takes no arguments");
        if(key=="name") return VbaValue::string(data_.name.toStdString());
        if(key=="table") return VbaValue::string(data_.parent.toStdString());
        if(key=="foreigntable") return VbaValue::string(data_.child.toStdString());
        if(key=="attributes") return VbaValue::longInteger(static_cast<std::int32_t>(data_.attributes));
        throw VbaRuntimeError(3251,"DAO: Relation property unsupported");
    }
    void assign(std::string_view,const VbaValue&) override { check(); throw VbaRuntimeError(383,"DAO: Imported relation metadata is read-only"); }
private:
    std::shared_ptr<NativeRelationCatalog> catalog_; NativeRelationData data_; quint64 version_{},rollback_{};
    std::weak_ptr<VbaObject> fields_;
};
class NativeRelations final : public VbaObject {
public:
    explicit NativeRelations(QSqlDatabase db) : catalog_(std::make_shared<NativeRelationCatalog>(std::move(db))) {}
    std::string_view className() const override { return "Relations"; }
    bool isNativeType(std::string_view type) const override { return type=="dao.relations"; }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        if(lower(name)=="item") return VbaArgumentSignature{{"Index"},1}; return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        const auto rows=catalog_->read(); const auto key=lower(name);
        if(key=="count" || key=="refresh") {
            if(!args.empty()) throw VbaRuntimeError(450,"DAO: Relation collection property takes no arguments");
            return key=="count" ? integerCount(rows.size()) : VbaValue{};
        }
        if(key!="item") throw VbaRuntimeError(3251,"DAO: Relation catalog mutation unsupported");
        if(args.size()!=1 || args[0].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(450,"DAO: Relation index required");
        const auto value=scalarVbaValue(args[0]); if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Invalid Null relation index");
        int position=-1;
        if(value.kind==VbaValue::Kind::Text) {
            const auto requested=text(value.asText()); for(int i=0;i<rows.size();++i) if(rows[i].name.compare(requested,Qt::CaseInsensitive)==0) position=i;
        } else { const auto index=coerceVbaModuleValue(value,"long").number; if(index>=0 && index<rows.size()) position=static_cast<int>(index); }
        if(position<0) throw VbaRuntimeError(3265,"DAO: Relation not found");
        const auto& row=rows[position]; const auto version=catalog_->version(),rollback=catalog_->workspace->rollbackRevision;
        auto& cached=objects_[row.name.toCaseFolded()]; auto object=cached.lock();
        if(!object || !object->matches(row,version,rollback)) { object=std::make_shared<NativeRelation>(catalog_,row,version,rollback); cached=object; }
        return VbaValue::reference(object);
    }
    VbaValue item(std::string_view name) override { return member("Item",{VbaValue::string(std::string(name))}); }
    std::vector<VbaValue> enumerate() override {
        const auto rows=catalog_->read(); std::vector<VbaValue> result;
        for(const auto& row:rows) result.push_back(member("Item",{VbaValue::string(row.name.toStdString())})); return result;
    }
private:
    std::shared_ptr<NativeRelationCatalog> catalog_; QMap<QString,std::weak_ptr<NativeRelation>> objects_;
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
        if (key=="tabledefs" || key=="relations") return VbaArgumentSignature{{"Index"},0};
        if (key=="querydefs") return VbaArgumentSignature{{"Index"},0};
        return {};
    }
    explicit Database(QSqlDatabase db) : db_(std::move(db)),workspace_(workspaceState(db_)) {}
    VbaValue member(std::string_view name, const std::vector<VbaValue>& args) override {
        workspace_->check(); require(db_.isOpen(), "DAO: Database is closed");
        const auto key = lower(name);
        if (key=="transactions" && args.empty()) return VbaValue::logical(true);
        if (key=="name" && args.empty()) return VbaValue::string(db_.databaseName().toUtf8().toStdString());
        if(key=="relations") {
            auto collection=relations_.lock(); if(!collection) { collection=std::make_shared<NativeRelations>(db_); relations_=collection; }
            return args.empty() || (args.size()==1 && args[0].kind==VbaValue::Kind::Missing) ? VbaValue::reference(collection) : collection->member("Item",args);
        }
        if (key=="tabledefs") {
            auto definitions=tableDefinitions_.lock();
            if (!definitions) { definitions=std::make_shared<TableDefinitions>(db_); tableDefinitions_=definitions; }
            if (args.empty() || (args.size()==1 && args[0].kind==VbaValue::Kind::Missing)) return VbaValue::reference(definitions);
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
            if(args.empty() || args.size()>2) throw VbaRuntimeError(450,"DAO: Database OpenRecordset accepts Name and optional Type");
            const int mode=nativeRecordsetMode(args,1);
            const auto source = text(args[0].asText());
            if (db_.tables().contains(source,Qt::CaseInsensitive))
                return VbaValue::reference(std::make_shared<Recordset>(db_,source,mode == 4));
            const auto sql = source.trimmed().startsWith("SELECT ",Qt::CaseInsensitive) ? source : accessQuerySource(db_,source);
            const auto plan = compileAccessQuery(sql);
            require(plan.sql.trimmed().startsWith("SELECT ",Qt::CaseInsensitive),"DAO: SELECT required");
            return VbaValue::reference(std::make_shared<Recordset>(db_,plan.sql,true,bindAccessQueryParameters(plan,{}),plan,mode));
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
    std::weak_ptr<VbaObject> relations_;
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
        if (key=="count" && args.empty()) { (void)object_->member("Name",{}); return VbaValue::integer(1); }
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
