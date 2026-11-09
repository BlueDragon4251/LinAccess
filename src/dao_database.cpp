#include "dao_database.hpp"
#include "vba_currency.hpp"
#include "dao_schema_change.hpp"
#include "native_schema_editor.hpp"
#include "native_field_defaults.hpp"
#include "native_field_validation.hpp"
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
#include <QCryptographicHash>
#include <bit>
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
    return type=="SINGLE_REAL" ? 6 : type=="INTEGER1" ? 2 : type=="INTEGER2" ? 3 : type=="INTEGER4" ? 4 : (type=="CURRENCY_REAL" || type=="CURRENCY_TICKS") ? 5 : (type=="DATE" || type=="DATETIME") ? 8 : type=="BLOB" ? 11 : 0;
}
VbaValue fixedDaoFieldSize(int type) {
    const int size=(type==1 || type==2) ? 1 : type==3 ? 2 : (type==4 || type==6) ? 4
        : (type==5 || type==7 || type==8) ? 8 : (type==11 || type==12) ? 0 : type==15 ? 16 : -1;
    if(size<0) throw VbaRuntimeError(3251,"DAO field size is not proven for this type");
    return VbaValue::longInteger(size);
}
int nativeRecordsetMode(const std::vector<VbaValue>& args,std::size_t index,int fallback=2) {
    if(index>=args.size() || args[index].kind==VbaValue::Kind::Missing) return fallback;
    const auto mode=static_cast<int>(coerceVbaModuleValue(scalarVbaValue(args[index]),"long").number);
    if(mode!=1 && mode!=2 && mode!=4) throw VbaRuntimeError(3251,"DAO: Native Recordset mode is not supported");
    return mode;
}
void require(bool valid, const char* message) { if (!valid) throw std::runtime_error(message); }
void exec(QSqlQuery& query, const QString& sql) {
    if (!query.exec(sql)) throw std::runtime_error("DAO: SQL operation failed");
}
[[noreturn]] void queryFailure(const QSqlQuery& query) {
    const auto message = query.lastError().text().toCaseFolded();
    if (const auto code=jetSqlFunctionErrorNumber(message.toUtf8().toStdString()))
        throw VbaRuntimeError(*code,jetSqlFunctionErrorExplanation(query.lastError().databaseText().toUtf8().toStdString()));
    if (message.contains("unique constraint") || message.contains("primary key")) throw VbaRuntimeError(3022,"DAO: Duplicate key");
    bool nativeOk=false;const int nativeCode=query.lastError().nativeErrorCode().toInt(&nativeOk);if((nativeOk && (nativeCode&255)==SQLITE_READONLY) || message.contains("attempt to write a readonly database")) throw VbaRuntimeError(3027,"DAO: Database or object is read-only");
    if (message.contains("not null")) throw VbaRuntimeError(3314,"DAO: Required value missing");
    if (message.contains("already exists")) throw VbaRuntimeError(3010,"DAO: destination table already exists");
    if (message.contains("no such function")) throw VbaRuntimeError(3085,"DAO: SQL function unsupported");
    if (message.contains("no such table")) throw VbaRuntimeError(3078,"DAO: SQL dependency unavailable");
    if ((message.contains("syntax error") || message.contains("incomplete input")) || message.contains("circularly defined")) throw VbaRuntimeError(3075,"DAO: SQL grammar unsupported");
    if(message.contains("rowid alias")) throw VbaRuntimeError(3251,"DAO: Altering rowid-named fields requires supported table reconstruction");
    if (message.contains("interrupted")) throw VbaRuntimeError(3035,"DAO: query execution budget exceeded");
    throw VbaRuntimeError(3146,"DAO: SQL operation failed");
}
bool connectionUpdatable(QSqlDatabase db);
QString physicalRowIdentity(QSqlDatabase db,const QString& table) {
    QSqlQuery query(db);if(!query.exec("PRAGMA table_xinfo("+quote(table)+")")) queryFailure(query);QSet<QString> names;while(query.next()) names.insert(query.value(1).toString().toCaseFolded());if(query.lastError().isValid()) queryFailure(query);query.finish();
    for(const auto& alias:{QString("rowid"),QString("_rowid_"),QString("oid")}) if(!names.contains(alias)) {
        // Keep the fixed alias unquoted: SQLite must resolve an actual hidden
        // identity, never a double-quoted fallback string on WITHOUT ROWID.
        if(!query.exec("SELECT "+alias+" FROM "+quote(table)+" LIMIT 0")) throw VbaRuntimeError(3251,"DAO: Physical row identity unavailable");return alias;
    }
    throw VbaRuntimeError(3251,"DAO: All physical row identity aliases are shadowed");
}
QStringList nativeTableNames(QSqlDatabase db);
void observeNativeTableCatalog(QSqlDatabase db,const QStringList& before);
void checkReservedCounterSchemaAction(QSqlDatabase db);
bool checkedDatasheetPageCatalog(QSqlDatabase db);
void removeDatasheetPageCatalog(QSqlDatabase db,const QString& kind,const QString& name);
void removeDroppedTableMetadata(QSqlDatabase db,const QStringList& before) {
    const auto after=nativeTableNames(db); const auto catalog=db.tables();
    QStringList removed,added;
    for(const auto& name:before) if(!after.contains(name,Qt::CaseInsensitive)) removed.append(name);
    for(const auto& name:after) if(!before.contains(name,Qt::CaseInsensitive)) added.append(name);
    // A rename needs rebinding of defaults, index identities and source metadata.
    // Refuse it atomically for managed/imported tables until that is implemented.
    for(const auto& name:removed) {
        if(added.isEmpty()) removeDatasheetPageCatalog(db,"table",name);
        else if(checkedDatasheetPageCatalog(db)) {QSqlQuery pages(db);pages.prepare("SELECT 1 FROM main._la_datasheet_page_setup WHERE kind='table' AND name=?");pages.addBindValue(name.toCaseFolded());if(!pages.exec()) queryFailure(pages);if(pages.next()) throw VbaRuntimeError(3251,"DAO: Rename requires datasheet page rebinding");}
        if(catalog.contains("_la_access_relations",Qt::CaseInsensitive)) {
            QSqlQuery relations(db);relations.prepare("SELECT 1 FROM main._la_access_relations WHERE child_table=? COLLATE NOCASE OR parent_table=? COLLATE NOCASE");relations.addBindValue(name);relations.addBindValue(name);
            if(!relations.exec()) queryFailure(relations);if(relations.next()) throw VbaRuntimeError(3251,"DAO: Remove dependent relations before dropping/renaming a table");
        }
        for(const auto& metadata:{QString("_la_counter_sequences"),QString("_la_counter32_sequences"),QString("_la_counter32_steps"),QString("_la_table_schema"),QString("_la_access_indexes")}) {
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
    if(!connectionUpdatable(db)) throw VbaRuntimeError(3027,"DAO: Database or object is read-only");
    const bool makeTable=!plan.makeTable.isEmpty();
    const bool schemaAction=plan.daoQueryType==96; // Checked compiler dbQDDL classification, including whitespace/comments.
    if(schemaAction || makeTable) checkReservedCounterSchemaAction(db);
    const auto before=(schemaAction || makeTable) ? nativeTableNames(db) : QStringList{};
    for(const auto& counter:plan.counters) if(counter.reseed) {
        if(!before.contains(counter.table,Qt::CaseInsensitive)) throw VbaRuntimeError(3078,"Counter table unavailable");
        QString canonical;for(const auto& name:before) if(name.compare(counter.table,Qt::CaseInsensitive)==0) {canonical=name;break;}
        if(!jetLongAccessColumns(db,canonical).contains(counter.column.toCaseFolded())) throw VbaRuntimeError(3251,"ALTER COUNTER requires an existing managed generator");
    }
    AccessCounterSqlScope counterScope(db);
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
            for(const auto& counter:plan.counters) {
                if(!after.contains(counter.table,Qt::CaseInsensitive) || (before.contains(counter.table,Qt::CaseInsensitive)!=counter.reseed)) throw VbaRuntimeError(3251,"Counter DDL table state is unsupported");
                QString canonical;
                for(const auto& name:after) if(name.compare(counter.table,Qt::CaseInsensitive)==0) {canonical=name;break;}
                QString column;
                QSqlQuery fields(db);if(!fields.exec("PRAGMA main.table_info("+quote(canonical)+")")) queryFailure(fields);
                while(fields.next()) if(fields.value(1).toString().compare(counter.column,Qt::CaseInsensitive)==0) {column=fields.value(1).toString();break;}fields.finish();
                if(column.isEmpty()) throw VbaRuntimeError(3265,"Counter DDL field unavailable");
                configureNewAccessCounterStep(db,canonical,column,counter.seed,counter.increment,counter.reseed);
            }
            observeNativeTableCatalog(db,before);
        }
        if(schemaScope && !guard.exec("RELEASE "+schemaSavepoint)) queryFailure(guard);
        counterScope.complete(true);
        return affected;
    } catch (...) {
        if(schemaScope) { guard.exec("ROLLBACK TO "+schemaSavepoint); guard.exec("RELEASE "+schemaSavepoint); }
        counterScope.complete(false);
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
    QMap<QString,quint64> queryGenerations,tableGenerations,tableFieldRevisions,relationGenerations;
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
        if(driver->property("la_schema_restore_failed").toBool()) throw VbaRuntimeError(3251,"DAO: Schema connection settings failed; reopen the project");
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
        [](void* data) { auto* holder=static_cast<std::shared_ptr<WorkspaceState>*>(data); if((*holder)->driver) {(*holder)->driver->setProperty("la_counter_sql_sync_failed",QVariant{});(*holder)->driver->setProperty("la_schema_restore_failed",QVariant{});} (*holder)->alive=false; (*holder)->connection=nullptr; delete holder; })!=SQLITE_OK)
        throw VbaRuntimeError(3146,"DAO: Workspace lifetime tracking unavailable");
    db.driver()->setProperty("la_dao_workspace_state",QVariant::fromValue(state)); return state;
}
bool connectionUpdatable(QSqlDatabase db) {
    const auto workspace=workspaceState(db);workspace->check();const int readonly=sqlite3_db_readonly(workspace->connection,"main");if(readonly<0) throw VbaRuntimeError(3251,"DAO: Database write capability unavailable");
    QSqlQuery query(db);if(!query.exec("PRAGMA query_only") || !query.next()) queryFailure(query);const auto queryOnly=query.value(0).toInt();if(queryOnly!=0 && queryOnly!=1) throw VbaRuntimeError(3251,"DAO: Connection write mode unknown");return readonly==0 && queryOnly==0;
}
void checkReservedCounterSchemaAction(QSqlDatabase db) {
    if(!connectionUpdatable(db)) throw VbaRuntimeError(3027,"DAO: Database or object is read-only");
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
    auto token=declaration.toString().trimmed();
    // The SQL compiler separates a unary sign from its numeric token.
    if(token.startsWith('+') || token.startsWith('-')) token=token.left(1)+token.mid(1).trimmed();
    const auto literal=accessLiteralDefaultSql(token);
    // Raw SQLite TRUE is1; imported Access TRUE is already normalized to-1.
    return token.compare("TRUE",Qt::CaseInsensitive)==0 ? QString("1") : literal.value_or(QString{});
}
struct TableSchemaColumn { QString name,type; int sourceType{},sourceSize{}; QString literalDefault; bool hasDefault{}; QString foreignName; int indexAttributes{-1},generatorAttributes{-1}; std::optional<bool> required;std::optional<int> nativeAttributes;std::optional<bool> allowZeroLength;QString storedDefault;std::optional<bool> canonicalDate;QString validationRule,validationText;std::optional<QString> caption,description,format;QString captionIdentity,descriptionIdentity,formatIdentity;bool sourceDisplayProven{},sourceFormatProven{}; };
using TableSchemaReader=std::function<QVector<TableSchemaColumn>()>;
using NativeFieldSetter=std::function<TableSchemaReader(const QString&,std::string_view,const VbaValue&)>;
std::shared_ptr<VbaObject> makeNativeFieldProperties(TableSchemaReader reader,QString field,NativeFieldSetter setter={},std::function<VbaValue(std::string_view)> intrinsic={});
std::optional<QString> nativeFieldDefaultExpression(const TableSchemaColumn& field) {
    if(provenDaoFieldType(field.type)==11 || field.sourceType==11) return {};
    if(!field.hasDefault) return QString{};
    if(field.sourceType==8 && field.nativeAttributes) if(const auto clock=nativeDateDefaultExpression(field.storedDefault)) return *clock;
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
    QVector<TableSchemaColumn> columns;QMap<QString,bool> notNull;
    while(query.next()) {
        const auto name=query.value(1).toString();
        columns.push_back({name,types.value(name.toCaseFolded()),0,0,sqliteLiteralDefaultSql(query.value(4)),!query.value(4).isNull()});
        columns.last().storedDefault=query.value(4).toString();
        notNull.insert(name.toCaseFolded(),query.value(3).toBool());
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
                if(size>0 && (code!=10 || (origin!="MDBTOOLS_MYSQL" && origin!="DAO_NATIVE"))) throw VbaRuntimeError(3251,"DAO source field capacity provenance unavailable");
                for(auto& column:columns) if(column.name.toCaseFolded()==name) {
                    column.sourceType=code;column.sourceSize=static_cast<int>(size);
                    if(origin=="DAO_NATIVE" && object.contains("native_required")) {
                        const auto required=object.value("native_required");
                        if(!required.isBool() || (required.toBool() && !notNull.value(name))) throw VbaRuntimeError(3251,"DAO native Required definition is not enforced");
                        column.required=required.toBool();
                    }
                    if(origin=="DAO_NATIVE" && object.contains("native_attributes")) {
                        const auto attribute=object.value("native_attributes");const auto expected=(code==10 || code==11 || code==12) ? 2 : 1;
                        if(!attribute.isDouble() || (attribute.toDouble()!=expected && !(code==4 && attribute.toDouble()==17))) throw VbaRuntimeError(3251,"DAO native field attributes invalid");
                        column.nativeAttributes=static_cast<int>(attribute.toDouble());
                    }
                    if(object.contains("source_display_properties_proven")) {const auto proven=object.value("source_display_properties_proven");if(!proven.isBool() || (origin!="MDBTOOLS_MYSQL" && origin!="MDBTOOLS_ACCESS")) throw VbaRuntimeError(3251,"DAO: Source field property provenance invalid");column.sourceDisplayProven=proven.toBool();if(column.sourceDisplayProven) for(const auto& property:{QString("caption"),QString("description")}) if(object.contains("source_"+property)) {const auto value=object.value("source_"+property),identity=object.value("source_"+property+"_id");const auto text=value.toString();if(!value.isString() || text.size()>4096 || text.contains(QChar(0)) || !text.isValidUtf16() || !identity.isString() || QUuid::fromString(identity.toString()).isNull()) throw VbaRuntimeError(3251,"DAO: Source field display property invalid");if(property=="caption") {column.caption=text;column.captionIdentity=identity.toString();}else {column.description=text;column.descriptionIdentity=identity.toString();}}}
                    if(column.sourceDisplayProven) {
                        if(object.contains("source_required")) {const auto value=object.value("source_required");if(!value.isBool() || (value.toBool() && !notNull.value(name))) throw VbaRuntimeError(3251,"DAO: Source Required definition is not enforced");column.required=value.toBool();}
                        if(object.contains("source_allow_zero_length")) {const auto value=object.value("source_allow_zero_length");if(!value.isBool() || (code!=10 && code!=12) || column.type!="TEXT") throw VbaRuntimeError(3251,"DAO: Source AllowZeroLength definition invalid");column.allowZeroLength=value.toBool();}
                    }
                    if(object.contains("source_format_properties_proven")) {
                        const auto proof=object.value("source_format_properties_proven");
                        if(!proof.isBool() || !column.sourceDisplayProven || (origin!="MDBTOOLS_MYSQL" && origin!="MDBTOOLS_ACCESS")) throw VbaRuntimeError(3251,"DAO: Original field Format provenance invalid");
                        column.sourceFormatProven=proof.toBool();
                    }
                    if(object.contains("source_format") || object.contains("source_format_id")) {
                        const auto value=object.value("source_format"),identity=object.value("source_format_id");const auto text=value.toString();
                        if(!column.sourceFormatProven || !value.isString() || text.size()>4096 || text.contains(QChar(0)) || !text.isValidUtf16() || !identity.isString() || QUuid::fromString(identity.toString()).isNull()) throw VbaRuntimeError(3251,"DAO: Original field Format property invalid");
                        column.format=text;column.formatIdentity=identity.toString();
                    }
                    if(origin=="DAO_NATIVE") for(const auto& property:{QString("caption"),QString("description"),QString("format")}) if(object.contains("native_"+property)) {const auto value=object.value("native_"+property);const auto text=value.toString();if(!value.isString() || text.size()>255 || text.contains(QChar(0)) || !text.isValidUtf16() || (property=="caption" && (text.contains('\n') || text.contains('\r')))) throw VbaRuntimeError(3251,"DAO: Native field display property invalid");const auto identity=object.value("native_"+property+"_id");if(!identity.isString() || QUuid::fromString(identity.toString()).isNull()) throw VbaRuntimeError(3251,"DAO: Native field property identity invalid");if(property=="caption") {column.caption=text;column.captionIdentity=identity.toString();}else if(property=="description") {column.description=text;column.descriptionIdentity=identity.toString();}else {column.format=text;column.formatIdentity=identity.toString();}}
                    if(origin=="DAO_NATIVE" && object.contains("native_validation_text")) {const auto message=object.value("native_validation_text");if(!message.isString() || message.toString().size()>255 || message.toString().contains(QChar(0)) || !message.toString().isValidUtf16() || (code!=1 && code!=2 && code!=3 && code!=4 && code!=6 && code!=7 && !(code==8 && object.value("native_date_canonical").toBool()))) throw VbaRuntimeError(3251,"DAO: Native ValidationText declaration invalid");column.validationText=message.toString();}
                    if(origin=="DAO_NATIVE" && object.contains("native_validation_rule")) {const auto rule=object.value("native_validation_rule");if(!rule.isString() || (code!=1 && code!=2 && code!=3 && code!=4 && code!=6 && code!=7 && !(code==8 && object.value("native_date_canonical").toBool()))) throw VbaRuntimeError(3251,"DAO: Native ValidationRule declaration invalid");column.validationRule=rule.toString();(void)nativeFieldValidationPredicate(code,column.validationRule,quote(column.name));}
                    if(origin=="DAO_NATIVE" && object.contains("native_date_canonical")) {const auto canonical=object.value("native_date_canonical");if(!canonical.isBool() || code!=8) throw VbaRuntimeError(3251,"DAO: Native Date storage declaration invalid");column.canonicalDate=canonical.toBool();}
                    if(origin=="DAO_NATIVE" && object.contains("native_allow_zero_length")) {
                        const auto allowed=object.value("native_allow_zero_length");if(!allowed.isBool() || (code!=10 && code!=12)) throw VbaRuntimeError(3251,"DAO: Native AllowZeroLength declaration invalid");column.allowZeroLength=allowed.toBool();
                    }
                    found=true;break;
                }
                if(!found) throw VbaRuntimeError(3251,"DAO source schema field unavailable");
            }
        }
        if(query.lastError().isValid()) queryFailure(query);
    }
    const auto counter32=jetLongAccessColumns(db,tableName);
    for(auto& column:columns) {
        if(column.nativeAttributes && (*column.nativeAttributes&16) && !counter32.contains(column.name.toCaseFolded())) throw VbaRuntimeError(3251,"DAO native AutoNumber attribute lacks a generator");
        if(counter32.contains(column.name.toCaseFolded())) {column.type="INTEGER4";column.sourceType=4;column.generatorAttributes=17;}
    }
    return columns;
}
bool canonicalNativeDateColumn(QSqlDatabase db,const QString& table,const TableSchemaColumn& column) {
    if(column.sourceType!=8 || !column.nativeAttributes || *column.nativeAttributes!=1 || !column.required || !column.canonicalDate.value_or(false)) return false;
    QSqlQuery declared(db);if(!declared.exec("PRAGMA main.table_info("+quote(table)+")")) queryFailure(declared);
    while(declared.next()) if(declared.value(1).toString()==column.name) return declared.value(2).toString().trimmed().compare("DATETIME",Qt::CaseInsensitive)==0;
    return false;
}
QMap<QString,QVariant> literalDefaultValues(QSqlDatabase db,const QVector<TableSchemaColumn>& columns,bool clocks=false) {
    QStringList keys,expressions;
    for(const auto& field:columns) {auto expression=field.literalDefault;if(expression.isEmpty() && clocks && field.sourceType==8 && field.nativeAttributes) if(const auto clock=nativeDateDefaultExpression(field.storedDefault)) expression=*nativeDateDefaultSql(*clock);if(!expression.isEmpty()) {keys.append(field.name.toCaseFolded());expressions.append(expression);}}
    QMap<QString,QVariant> result;
    if(expressions.isEmpty()) return result;
    if(clocks) installJetSqlFunctions(db);
    QSqlQuery values(db);
    if(!values.exec("SELECT "+expressions.join(',')) || !values.next()) queryFailure(values);
    for(int i=0;i<keys.size();++i) {auto value=values.value(i);for(const auto& field:columns) if(field.name.toCaseFolded()==keys[i] && (field.type=="SINGLE_REAL" || field.type.contains("DATE",Qt::CaseInsensitive))) {value=sqlFieldValue(fromSql(value),field.type);break;}result.insert(keys[i],value);}
    return result;
}
struct NativeRecordsetIndex {QString name;QStringList fields,types;QVector<bool> descending;bool operator==(const NativeRecordsetIndex&) const=default;};
NativeRecordsetIndex checkedRecordsetIndex(QSqlDatabase db,const QString& table,const QString& requested);
class Recordset;
struct RecordsetData {
    std::shared_ptr<WorkspaceState> workspace;
    quint64 rollbackRevision{},fieldRevision{};
    QStringList columns;
    QString rowIdentity{"rowid"};
    std::optional<NativeRecordsetIndex> selectedIndex;
    QMap<QString,QString> types;
    QMap<QString,int> daoTypes,definitionSizes;
    QMap<QString,QVariant> literalDefaults;
    QMap<QString,QString> defaultExpressions,validationRules,validationTexts;
    QMap<QString,bool> requiredFields,allowZeroLengthFields;
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
        if(lower(name)=="properties") return VbaArgumentSignature{{"Index"},0};
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
    void assignIndexedMember(std::string_view name,const std::vector<VbaValue>& args,const VbaValue& value) override {if(lower(name)!="properties" || args.size()!=1) throw VbaRuntimeError(438,"DAO: Indexed field property assignment unsupported");member("Properties",args).object->assign("Value",value);}

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
        if (key=="seek") return VbaArgumentSignature{{"Comparison","Key1","Key2","Key3","Key4","Key5","Key6","Key7","Key8","Key9","Key10"},2};
        if (key=="getrows") return VbaArgumentSignature{{"NumRows"},0};
        if (key=="fields") return VbaArgumentSignature{{"Index"},0};
        return {};
    }
    Recordset(QSqlDatabase db, QString source, bool snapshot, QMap<QString,QVariant> bindings = {},AccessQueryPlan plan={},int cursorKind=0)
        : db_(std::move(db)), source_(std::move(source)), bindings_(std::move(bindings)), readOnly_(snapshot), recordsetType_(cursorKind ? cursorKind : snapshot ? 4 : 2),queryPlan_(std::move(plan)) {
        data_->workspace=workspaceState(db_); data_->rollbackRevision=data_->workspace->rollbackRevision;
        table_ = db_.tables().contains(source_, Qt::CaseInsensitive);
        if(table_ && nativeTableNames(db_).contains(source_,Qt::CaseInsensitive)) tableGeneration_=data_->workspace->tableGenerations.value(source_.toCaseFolded());
        readOnly_ |= !table_ || !connectionUpdatable(db_);
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
        if(table_ && data_->fieldRevision!=data_->workspace->tableFieldRevisions.value(source_.toCaseFolded())) throw VbaRuntimeError(3420,"DAO: Recordset requires Requery after field schema changes");
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
        if((key=="edit" || key=="addnew" || key=="update" || key=="delete") && !connectionUpdatable(db_)) throw VbaRuntimeError(3027,"DAO: Database or object is read-only");
        if (key=="transactions" && args.empty()) return VbaValue::logical(true);
        if(key=="index") {if(!args.empty()) throw VbaRuntimeError(450,"DAO: Index property takes no arguments");if(!table_ || recordsetType_!=1) throw VbaRuntimeError(3251,"DAO: Index requires a table-type Recordset");if(!indexName_.isEmpty() && (!data_->selectedIndex || checkedRecordsetIndex(db_,source_,indexName_)!=*data_->selectedIndex)) throw VbaRuntimeError(3251,"DAO: Index changed; Requery is required");return VbaValue::string(indexName_.toStdString());}
        if(key=="seek") {seek(args);return {};}
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
        if (key=="updatable") return VbaValue::logical(!readOnly_ && connectionUpdatable(db_));
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
                QStringList expressions,names;for(auto it=data_->defaultExpressions.cbegin();it!=data_->defaultExpressions.cend();++it) if(data_->daoTypes.value(it.key())==8) if(const auto expression=nativeDateDefaultSql(it.value())) {names.append(it.key());expressions.append(*expression);}
                if(!expressions.isEmpty()) {installJetSqlFunctions(db_);QSqlQuery defaults(db_);if(!defaults.exec("SELECT "+expressions.join(',')) || !defaults.next()) queryFailure(defaults);for(int i=0;i<names.size();++i) {const auto index=column(names[i]);pending_[index]=defaults.value(i);changed_.insert(index);}}
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
            require(query.prepare("DELETE FROM " + quote(source_) + " WHERE "+data_->rowIdentity+"=?"), "DAO: Delete prepare failed");
            query.addBindValue(data_->ids[position_]);
            require(query.exec() && query.numRowsAffected() == 1, "DAO: Delete failed or record changed");
            reload(false);
            return {};
        }
        if (key == "requery") {
            checkTableGeneration();
            if (data_->rollbackRevision!=data_->workspace->rollbackRevision || (table_ && data_->fieldRevision!=data_->workspace->tableFieldRevisions.value(source_.toCaseFolded()))) { mode_=Mode::Browse; pending_.clear(); changed_.clear(); chunkFields_.clear(); }
            require(mode_ == Mode::Browse, "DAO: Pending edit");
            reload(); return {};
        }
        if (key == "close") { if (cloneCounted_) { --data_->workspace->clones; cloneCounted_=false; } closed_ = true; data_->cursors.remove(this); pending_.clear(); changed_.clear(); chunkFields_.clear(); return {}; }
        return value(text(name)); // implicit default Fields(name).Value
    }
    void assign(std::string_view name, const VbaValue& value) override {
        if(lower(name)=="index") {
            if(closed_) throw VbaRuntimeError(3420,"DAO: Recordset is closed");checkTransaction();if(!table_ || recordsetType_!=1) throw VbaRuntimeError(3251,"DAO: Index requires a table-type Recordset");if(mode_!=Mode::Browse) throw VbaRuntimeError(3219,"DAO: Pending edit");const auto revision=data_->revision;const auto scalar=scalarVbaValue(value);checkTransaction();if(closed_) throw VbaRuntimeError(3420,"DAO: Recordset closed during index evaluation");if(mode_!=Mode::Browse || revision!=data_->revision) throw VbaRuntimeError(3219,"DAO: Recordset changed during index evaluation");if(scalar.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null index name");
            const auto requested=text(scalar.asText());std::optional<NativeRecordsetIndex> next;if(!requested.isEmpty()) next=checkedRecordsetIndex(db_,source_,requested);const auto name=next ? next->name : QString{};if(name==indexName_ && next==data_->selectedIndex) return;if(data_->cursors.size()!=1) throw VbaRuntimeError(3251,"DAO: Change Index before creating clones or after closing them");const auto old=indexName_;const auto id=position_>=0 && position_<data_->ids.size() ? data_->ids[position_] : QVariant{};indexName_=name;try {reload(false);if(!id.isNull()) position_=data_->ids.indexOf(id);}catch(...) {indexName_=old;throw;}return;
        }
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
        if(!table_ || declaredFieldType(name).asNumber()!=4
            || !generatedAccessColumns(db_,source_).contains(name.toCaseFolded()))
            throw VbaRuntimeError(3251,"DAO: Field generator property is not proven in this context");
        // Access2003/DAO3.6 writable Long counter observation: Fixed|AutoIncr|Updatable.
        const bool writable=!readOnly_ && connectionUpdatable(db_);return attributes ? VbaValue::longInteger(writable ? 49 : 17) : VbaValue::logical(writable);
    }
    VbaValue schemaFieldBoolean(const QString& name,bool allowZeroLength) const {
        (void)fieldType(name);const auto& properties=allowZeroLength ? data_->allowZeroLengthFields : data_->requiredFields;const auto key=name.toCaseFolded();
        if(!properties.contains(key)) throw VbaRuntimeError(3251,"DAO: Recordset field property has no proven native declaration");return VbaValue::logical(properties.value(key));
    }
    QVector<TableSchemaColumn> propertySchema() const {checkTransaction();if(closed_) throw VbaRuntimeError(3420,"DAO: Recordset is closed");if(!table_) throw VbaRuntimeError(3251,"DAO: Query field properties are not proven");return readTableSchemaColumns(db_,source_);}
    VbaValue schemaFieldValidation(const QString& name,bool message) const {
        (void)fieldType(name);const auto& properties=message ? data_->validationTexts : data_->validationRules;const auto key=name.toCaseFolded();if(!properties.contains(key)) throw VbaRuntimeError(3251,"DAO: Recordset validation metadata has no proven native declaration");return VbaValue::string(properties.value(key).toStdString());
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
        if(!connectionUpdatable(db_)) throw VbaRuntimeError(3027,"DAO: Database or object is read-only");
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
        if(!connectionUpdatable(db_)) throw VbaRuntimeError(3027,"DAO: Database or object is read-only");
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
          table_(source.table_),readOnly_(source.readOnly_),frozen_(source.frozen_),recordsetType_(source.recordsetType_),tableGeneration_(source.tableGeneration_),indexName_(source.indexName_),position_(-1),queryPlan_(source.queryPlan_) {
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
    void seek(const std::vector<VbaValue>& arguments) {
        if(!table_ || recordsetType_!=1) throw VbaRuntimeError(3251,"DAO: Seek requires a table-type Recordset");if(mode_!=Mode::Browse) throw VbaRuntimeError(3219,"DAO: Pending edit");if(arguments.size()<2 || arguments.size()>11) throw VbaRuntimeError(450,"DAO: Seek requires comparison and complete keys");if(indexName_.isEmpty()) throw VbaRuntimeError(3251,"DAO: Set Index before Seek");
        const auto index=checkedRecordsetIndex(db_,source_,indexName_);if(!data_->selectedIndex || index!=*data_->selectedIndex) throw VbaRuntimeError(3251,"DAO: Index changed; Requery is required");if(arguments.size()>static_cast<std::size_t>(index.fields.size()+1)) throw VbaRuntimeError(3251,"DAO: Too many Seek keys for selected index");const auto revision=data_->revision;const auto comparison=text(scalarVbaValue(arguments[0]).asText());if(comparison!="=" && comparison!="<" && comparison!="<=" && comparison!=">" && comparison!=">=") throw VbaRuntimeError(3251,"DAO: Seek comparison unsupported");if(comparison!="=" && index.descending.contains(true)) throw VbaRuntimeError(3251,"DAO: Descending range Seek requires Access reference validation");
        QVector<QVariant> keys;QVector<int> positions;bool suffix=false;for(int field=0;field<index.fields.size();++field) {const auto value=field+1<arguments.size() ? scalarVbaValue(arguments[field+1]) : VbaValue::missing();if(value.kind==VbaValue::Kind::Missing) {if(field==0) throw VbaRuntimeError(3251,"DAO: Seek requires the first index key");suffix=true;keys.append(QVariant{});positions.append(column(index.fields[field]));continue;}if(suffix || value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(3251,"DAO: Only omitted trailing Seek keys are supported, not explicit Null or holes");const bool date=index.types[field].contains("DATE",Qt::CaseInsensitive);if(date ? value.kind!=VbaValue::Kind::Date : (value.kind!=VbaValue::Kind::Number && value.kind!=VbaValue::Kind::Boolean)) throw VbaRuntimeError(13,"DAO: Seek requires a matching numeric or Date key");keys.append(sqlFieldValue(value,index.types[field]));positions.append(column(index.fields[field]));}
        checkTransaction();if(closed_) throw VbaRuntimeError(3420,"DAO: Recordset closed during Seek evaluation");if(mode_!=Mode::Browse || revision!=data_->revision) throw VbaRuntimeError(3219,"DAO: Recordset changed during Seek evaluation");const bool backwards=comparison=="<" || comparison=="<=";QElapsedTimer clock;clock.start();
        for(qsizetype row=backwards ? data_->rows.size()-1 : 0;row>=0 && row<data_->rows.size();row+=backwards ? -1 : 1) {if(clock.hasExpired(5000)) throw VbaRuntimeError(3035,"DAO: Seek execution budget exceeded");int order=0;for(int field=0;field<positions.size();++field) {const auto& value=data_->rows[row][positions[field]];if(value.isNull()) throw VbaRuntimeError(3251,"DAO: Nullable Seek index is not supported");if(keys[field].isNull()) {order=1;break;}if(index.types[field].contains("DATE",Qt::CaseInsensitive)) order=QString::compare(value.toString(),keys[field].toString(),Qt::CaseSensitive);else {const auto number=value.toDouble(),key=keys[field].toDouble();order=number<key ? -1 : number>key ? 1 : 0;}if(order) break;}
            const bool matches=comparison=="=" ? order==0 : comparison=="<" ? order<0 : comparison=="<=" ? order<=0 : comparison==">" ? order>0 : order>=0;if(matches) {position_=row;noMatch_=false;return;}}
        position_=data_->rows.size();noMatch_=true;
    }
    void find(const std::string& kind,const QString& criteria) {
        if(data_->types.values().contains("CURRENCY_TICKS")) throw VbaRuntimeError(3251,"DAO: Textual Find criteria on exact Currency projections are not implemented");
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
        const auto fieldRevision=table_ ? data_->workspace->tableFieldRevisions.value(source_.toCaseFolded()) : 0;
        const bool schemaChanged=table_ && data_->fieldRevision!=fieldRevision;
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
            if (schemaChanged || data_->rollbackRevision!=data_->workspace->rollbackRevision) data_->family=QUuid::createUuid().toRfc4122();
            data_->rollbackRevision=data_->workspace->rollbackRevision;data_->fieldRevision=fieldRevision;
            data_->bookmarks.clear();
            for (qsizetype i=0;i<data_->rows.size();++i) {
                const auto old=table_ ? previousTokens.value(data_->ids[i].toLongLong()) : 0;
                data_->bookmarks.append(old ? old : data_->nextBookmark++);
            }
            for (auto* cursor:data_->cursors) {
                if(schemaChanged) {cursor->mode_=Mode::Browse;cursor->pending_.clear();cursor->changed_.clear();cursor->chunkFields_.clear();cursor->lastModified_.reset();}
                if (invalidate) cursor->position_=-1;
                else if (previousRows.contains(cursor)) cursor->position_=data_->ids.indexOf(previousRows.value(cursor));
                else if (cursor->position_>=0) cursor->position_=data_->rows.size();
                cursor->noMatch_=false;
            }
            position_=0; noMatch_=false;
            ++data_->revision;
            completed=true;
        };
        data_->types.clear(); data_->daoTypes.clear(); data_->definitionSizes.clear(); data_->literalDefaults.clear(); data_->defaultExpressions.clear();data_->requiredFields.clear();data_->allowZeroLengthFields.clear();data_->validationRules.clear();data_->validationTexts.clear();
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
                if(field.nativeAttributes) {data_->validationRules.insert(key,field.validationRule);data_->validationTexts.insert(key,field.validationText);}
                if(field.required) data_->requiredFields.insert(key,*field.required);if(field.allowZeroLength) data_->allowZeroLengthFields.insert(key,*field.allowZeroLength);
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
        data_->rowIdentity=physicalRowIdentity(db_,source_);std::optional<NativeRecordsetIndex> selected;QString order;if(!indexName_.isEmpty()) {selected=checkedRecordsetIndex(db_,source_,indexName_);QStringList terms;for(int field=0;field<selected->fields.size();++field) terms.append(quote(selected->fields[field])+(selected->descending[field] ? " DESC" : " ASC"));terms.append(data_->rowIdentity+" ASC");order=" ORDER BY "+terms.join(',');}
        exec(query,"SELECT "+data_->rowIdentity+",* FROM " + quote(source_)+order);data_->selectedIndex=std::move(selected);
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
                    : "UPDATE " + quote(source_) + " SET " + assignments.join(',') + " WHERE "+data_->rowIdentity+"=? RETURNING "+data_->rowIdentity),
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
    QString source_,indexName_;
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
    if(lower(name)=="properties") {const auto field=std::make_shared<Field>(owner_,name_);auto properties=makeNativeFieldProperties([owner=owner_] {return owner->propertySchema();},name_,{},[field](std::string_view key) {return field->member(key,{});});if(args.empty() || (args.size()==1 && args[0].kind==VbaValue::Kind::Missing)) return VbaValue::reference(properties);return properties->member("Item",args);}
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
    if(lower(name)=="required" || lower(name)=="allowzerolength") {
        if(!args.empty()) throw VbaRuntimeError(450,"Field schema property requires no arguments");return owner_->schemaFieldBoolean(name_,lower(name)=="allowzerolength");
    }
    if(lower(name)=="validationrule" || lower(name)=="validationtext") {if(!args.empty()) throw VbaRuntimeError(450,"Field validation property requires no arguments");return owner_->schemaFieldValidation(name_,lower(name)=="validationtext");}
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
    if(lower(name)=="validationrule" || lower(name)=="validationtext") {(void)owner_->schemaFieldValidation(name_,lower(name)=="validationtext");throw VbaRuntimeError(383,"Recordset field validation properties are read-only");}
    if(lower(name)=="required" || lower(name)=="allowzerolength") {(void)owner_->schemaFieldBoolean(name_,lower(name)=="allowzerolength");throw VbaRuntimeError(383,"Recordset schema field property is read-only");}
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
// Callers own a surrounding mutation savepoint; no QtGui dependency.
bool checkedDatasheetPageCatalog(QSqlDatabase db) {
    QSqlQuery query(db);if(!query.exec("SELECT type FROM main.sqlite_schema WHERE name='_la_datasheet_page_setup'")) queryFailure(query);if(!query.next()) return false;if(query.value(0).toString()!="table" || query.next()) throw VbaRuntimeError(3251,"DAO: Datasheet page catalog invalid");query.finish();
    if(!query.exec("PRAGMA main.table_info('_la_datasheet_page_setup')")) queryFailure(query);const QStringList shape{"kind","name","layout"};int count=0;while(query.next()) {if(count>=3 || query.value(1).toString()!=shape[count] || query.value(2).toString().toUpper()!="TEXT" || query.value(3).toInt()!=1 || query.value(5).toInt()!=(count<2 ? count+1 : 0)) throw VbaRuntimeError(3251,"DAO: Datasheet page catalog shape invalid");++count;}if(count!=3) throw VbaRuntimeError(3251,"DAO: Datasheet page catalog shape invalid");return true;
}
void removeDatasheetPageCatalog(QSqlDatabase db,const QString& kind,const QString& name) {
    if(!checkedDatasheetPageCatalog(db)) return;QSqlQuery query(db);query.prepare("DELETE FROM main._la_datasheet_page_setup WHERE kind=? AND name=?");query.addBindValue(kind);query.addBindValue(name.toCaseFolded());if(!query.exec()) queryFailure(query);
}
void renameDatasheetPageCatalog(QSqlDatabase db,const QString& kind,const QString& oldName,const QString& newName) {
    if(!checkedDatasheetPageCatalog(db)) return;QSqlQuery query(db);
    const auto oldKey=oldName.toCaseFolded(),newKey=newName.toCaseFolded();if(oldKey==newKey) return;
    query.prepare("SELECT 1 FROM main._la_datasheet_page_setup WHERE kind=? AND name=?");query.addBindValue(kind);query.addBindValue(newKey);if(!query.exec()) queryFailure(query);if(query.next()) throw VbaRuntimeError(3251,"DAO: Target datasheet page settings already exist");query.finish();
    query.prepare("UPDATE main._la_datasheet_page_setup SET name=? WHERE kind=? AND name=?");query.addBindValue(newKey);query.addBindValue(kind);query.addBindValue(oldKey);if(!query.exec()) queryFailure(query);
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
            if(mode==1) throw VbaRuntimeError(3251,"DAO: QueryDef cannot open a table-type Recordset");
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
        if(lower(name)=="name") {renameOwned(value);return;}
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
    void renameOwned(const VbaValue& value) {
        const auto scalar=scalarVbaValue(value);refresh();
        if(!owned_ || name_.isEmpty()) throw VbaRuntimeError(3251,"DAO: QueryDef rename requires an owned persistent query");
        if(scalar.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Invalid Null QueryDef name");
        const auto next=text(scalar.asText());
        if(next.trimmed().isEmpty() || next.size()>64 || !next.isValidUtf16() || next.contains(QChar(0)) || next.contains('\n') || next.contains('\r')) throw VbaRuntimeError(5,"DAO: Unsupported QueryDef name");
        if(next==name_) return;
        if(!connectionUpdatable(db_)) throw VbaRuntimeError(3027,"DAO: Database or object is read-only");
        if(!sqlite3_get_autocommit(workspace_->connection)) throw VbaRuntimeError(3251,"DAO: Complete the current transaction before renaming a QueryDef");
        const auto oldKey=name_.toCaseFolded(),newKey=next.toCaseFolded();
        QSqlQuery query(db_);const auto savepoint="la_rename_query_"+QUuid::createUuid().toString(QUuid::Id128);
        QVariantMap registry;for(const auto* key:WorkspaceState::queryViewProperties) registry.insert(key,db_.driver()->property(key));
        if(!query.exec("SAVEPOINT "+savepoint)) queryFailure(query);bool completed=false;
        const auto rollback=qScopeGuard([&] {if(!completed) {query.exec("ROLLBACK TO "+savepoint);query.exec("RELEASE "+savepoint);for(auto it=registry.cbegin();it!=registry.cend();++it) db_.driver()->setProperty(it.key().toUtf8().constData(),it.value());}});
        if(newKey!=oldKey) {
            bool exists=false;try {(void)accessQuerySource(db_,next,nullptr,nullptr,true);exists=true;}catch(const VbaRuntimeError& failure) {if(failure.number!=3265) throw;}
            if(exists || db_.tables(QSql::AllTables).contains(next,Qt::CaseInsensitive)) throw VbaRuntimeError(3012,"DAO: QueryDef name already exists");
        }
        query.prepare("SELECT definition FROM main._la_objects WHERE kind='query' AND name=?");query.addBindValue(name_);
        if(!query.exec()) queryFailure(query);if(!query.next()) throw VbaRuntimeError(3420,"DAO: Owned query was removed");const auto previous=query.value(0);const auto document=QJsonDocument::fromJson(previous.toByteArray());
        if(!document.isObject() || document.object().value("sql").toString()!=source_) throw VbaRuntimeError(3197,"DAO: Owned query changed concurrently");query.finish();
        query.prepare("UPDATE main._la_objects SET name=? WHERE kind='query' AND name=? AND definition=?");query.addBindValue(next);query.addBindValue(name_);query.addBindValue(previous);
        if(!query.exec()) queryFailure(query);if(query.numRowsAffected()!=1) throw VbaRuntimeError(3197,"DAO: Owned query changed concurrently");query.finish();
        renameDatasheetPageCatalog(db_,"query",name_,next);
        if(installAccessQueryViews(db_)<0) throw VbaRuntimeError(3251,"DAO: Query views are busy");
        if(!query.exec("RELEASE "+savepoint)) queryFailure(query);completed=true;
        ++workspace_->queryGenerations[oldKey];if(newKey!=oldKey) ++workspace_->queryGenerations[newKey];name_=next;namedGeneration_=workspace_->queryGenerations.value(newKey);
    }
    VbaValue parameter(const QString& name) const {
        checkParameter(name);
        const auto stored=values_.value(name.toCaseFolded());
        for(const auto& parameter:plan_.parameters) if(parameter.name.compare(name,Qt::CaseInsensitive)==0) {
            if(parameter.type=="currency") return stored.isNull() ? VbaValue{} : convertCurrency(fromSql(stored));
            if(parameter.type=="date" || parameter.type=="datetime") return qtFieldValue(stored,"DATETIME");
        }
        return fromSql(stored);
    }
    void setParameter(const QString& name, const VbaValue& value) {
        checkParameter(name);
        for(const auto& parameter:plan_.parameters)
            if(parameter.type=="currency" && parameter.name.compare(name,Qt::CaseInsensitive)==0) {
                const auto converted=value.kind==VbaValue::Kind::Null ? QVariant{} : QVariant(QString::fromStdString(convertCurrency(value).asText()));
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
    const auto scalar=scalarVbaValue(value);query_->parameterGeneration(generation_);
    query_->setParameter(name_,scalar);
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
            if(!connectionUpdatable(db_)) throw VbaRuntimeError(3027,"DAO: Database or object is read-only");
            QVariantMap registry;for(const auto* key:WorkspaceState::queryViewProperties) registry.insert(key,db_.driver()->property(key));
            QSqlQuery query(db_); const auto savepoint="la_delete_query_"+QUuid::createUuid().toString(QUuid::Id128);
            if(!query.exec("SAVEPOINT "+savepoint)) queryFailure(query);
            bool completed=false;
            const auto rollback=qScopeGuard([&] { if(!completed) { query.exec("ROLLBACK TO "+savepoint); query.exec("RELEASE "+savepoint);for(auto it=registry.cbegin();it!=registry.cend();++it) db_.driver()->setProperty(it.key().toUtf8().constData(),it.value()); } });
            QString canonical; bool owned=false; (void)accessQuerySource(db_,text(value.asText()),&canonical,&owned,true);
            if(!owned) throw VbaRuntimeError(3251,"DAO: Imported query deletion requires an archive-preserving overlay");
            auto& generation=workspace_->queryGenerations[canonical.toCaseFolded()];
            query.prepare("DELETE FROM main._la_objects WHERE kind='query' AND name=?"); query.addBindValue(canonical);
            if(!query.exec()) queryFailure(query);
            if(query.numRowsAffected()!=1) throw VbaRuntimeError(3265,"DAO: QueryDef no longer exists"); query.finish();
            removeDatasheetPageCatalog(db_,"query",canonical);
            if(installAccessQueryViews(db_)<0) throw VbaRuntimeError(3251,"DAO: Query views are busy");
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
        if(!object) for(const auto& candidate:objects_) if(auto held=candidate.lock();held && !held->closed_ && held->name_==canonical && held->owned_==owned && (!owned || held->namedGeneration_==workspace_->queryGenerations.value(canonical.toCaseFolded()))) {object=held;weak=held;break;}
        if(!object || object->name_!=canonical || object->closed_ || (object->owned_ && object->namedGeneration_!=workspace_->queryGenerations.value(canonical.toCaseFolded()))) { object=std::make_shared<QueryDefinition>(db_,source,canonical,owned); weak=object; }
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


TableSchemaReader updateNativeFieldPropertiesImpl(QSqlDatabase db,const QString& table,const QString& field,const QMap<QString,VbaValue>& properties);
QString nativeTextPropertyKey(const QString& name) {const auto key=name.toCaseFolded();if(key!="caption" && key!="description" && key!="format") throw VbaRuntimeError(3251,"DAO: Only native Caption, Description and Format properties are implemented");return key;}
QString nativeTextPropertyValue(const QString& key,const VbaValue& input) {const auto value=scalarVbaValue(input);if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null text property value");if(value.kind!=VbaValue::Kind::Text) throw VbaRuntimeError(13,"DAO: Text property requires a String");const auto text=QString::fromStdString(value.text);if(text.size()>255 || text.contains(QChar(0)) || !text.isValidUtf16() || (key=="caption" && (text.contains('\n') || text.contains('\r')))) throw VbaRuntimeError(3251,"DAO: Text property exceeds supported text limits");return text;}
struct NativeFieldPropertyState {
    TableSchemaReader reader;QString field;NativeFieldSetter setter;std::function<VbaValue(std::string_view)> intrinsic;
    TableSchemaColumn column() const {std::optional<TableSchemaColumn> selected;for(const auto& column:reader()) if(column.name.compare(field,Qt::CaseInsensitive)==0) {if(selected) throw VbaRuntimeError(3251,"DAO: Ambiguous property field");selected=column;}if(!selected) throw VbaRuntimeError(3420,"DAO: Property field unavailable");if(!selected->nativeAttributes && !selected->sourceDisplayProven) throw VbaRuntimeError(3251,"DAO: Original/unproven field properties are not inferred");return *selected;}
    std::optional<QString> value(const QString& key) const {const auto field=column();if(key=="format" && !field.nativeAttributes && !field.sourceFormatProven) throw VbaRuntimeError(3251,"DAO: Original field Format provenance is unavailable");return key=="caption" ? field.caption : key=="description" ? field.description : field.format;}
    QString identity(const QString& key) const {const auto field=column();return key=="caption" ? field.captionIdentity : key=="description" ? field.descriptionIdentity : field.formatIdentity;}
    void store(const QString& key,const VbaValue& value) {const auto scalar=scalarVbaValue(value);const auto field=column();if(!setter) throw VbaRuntimeError(383,"DAO: Recordset field properties are readonly");auto next=setter(field.name,key.toStdString(),scalar);reader=std::move(next);if(key=="name") this->field=QString::fromStdString(scalar.asText());}
};
class NativeFieldProperty final : public VbaObject {
public:
    NativeFieldProperty(std::shared_ptr<NativeFieldPropertyState> state,QString key):state_(std::move(state)),key_(std::move(key)) {}
    std::string_view className() const override {return "Property";}bool isNativeType(std::string_view type) const override {return type=="dao.property";}
    void bindIdentity() {identity_=state_->identity(key_);if(identity_.isEmpty()) throw VbaRuntimeError(3420,"DAO: Stored property identity unavailable");}
    bool hasScalarDefault() const override {return true;}VbaValue scalarValue() override {return member("Value",{});}
    void assignDefaultMember(const std::vector<VbaValue>& args,const VbaValue& value) override {if(!args.empty()) throw VbaRuntimeError(450,"DAO: Scalar property takes no indices");assign("Value",value);}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {if(!args.empty()) throw VbaRuntimeError(450,"DAO: Property takes no arguments");if(identity_.isEmpty()) bindIdentity();if(identity_!=state_->identity(key_)) throw VbaRuntimeError(3420,"DAO: Property identity replaced");const auto value=state_->value(key_);if(!value) throw VbaRuntimeError(3270,"DAO: Property not found");const auto key=lower(name);if(key=="name") return VbaValue::string(key_=="caption" ? "Caption" : key_=="description" ? "Description" : "Format");if(key=="type") return VbaValue::integer(10);if(key=="value") return VbaValue::string(value->toStdString());if(key=="inherited") {if(!state_->column().nativeAttributes) throw VbaRuntimeError(3251,"DAO: Source property inheritance flag is not proven");return VbaValue::logical(false);}throw VbaRuntimeError(3251,"DAO: Property member unsupported");}
    void assign(std::string_view name,const VbaValue& value) override {if(identity_.isEmpty()) bindIdentity();if(identity_!=state_->identity(key_)) throw VbaRuntimeError(3420,"DAO: Property identity replaced");if(!state_->value(key_)) throw VbaRuntimeError(3270,"DAO: Property not found");if(lower(name)!="value") throw VbaRuntimeError(383,"DAO: Property metadata is readonly");const auto text=nativeTextPropertyValue(key_,value);state_->store(key_,VbaValue::string(text.toStdString()));}
private:std::shared_ptr<NativeFieldPropertyState> state_;QString key_,identity_;
};
const QMap<QString,QPair<QString,int>>& nativeIntrinsicFieldProperties() {static const QMap<QString,QPair<QString,int>> properties{{"name",{"Name",10}},{"type",{"Type",3}},{"size",{"Size",4}},{"attributes",{"Attributes",4}},{"dataupdatable",{"DataUpdatable",1}},{"ordinalposition",{"OrdinalPosition",3}},{"required",{"Required",1}},{"allowzerolength",{"AllowZeroLength",1}},{"defaultvalue",{"DefaultValue",10}},{"validationrule",{"ValidationRule",10}},{"validationtext",{"ValidationText",10}}};return properties;}
class NativeIntrinsicFieldProperty final : public VbaObject {
public:
    NativeIntrinsicFieldProperty(std::shared_ptr<NativeFieldPropertyState> state,QString key):state_(std::move(state)),key_(std::move(key)) {}
    std::string_view className() const override {return "Property";}bool isNativeType(std::string_view type) const override {return type=="dao.property";}
    bool hasScalarDefault() const override {return true;}VbaValue scalarValue() override {return member("Value",{});}
    void assignDefaultMember(const std::vector<VbaValue>& args,const VbaValue& value) override {if(!args.empty()) throw VbaRuntimeError(450,"DAO: Scalar property takes no indices");assign("Value",value);}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {if(!args.empty()) throw VbaRuntimeError(450,"DAO: Property takes no arguments");(void)state_->column();if(!state_->intrinsic) throw VbaRuntimeError(3251,"DAO: Intrinsic field property context unavailable");const auto value=state_->intrinsic(key_.toStdString());const auto key=lower(name);if(key=="name") return VbaValue::string(nativeIntrinsicFieldProperties().value(key_).first.toStdString());if(key=="type") return VbaValue::integer(nativeIntrinsicFieldProperties().value(key_).second);if(key=="value") return value;if(key=="inherited") {if(!state_->column().nativeAttributes) throw VbaRuntimeError(3251,"DAO: Source property inheritance flag is not proven");return VbaValue::logical(false);}throw VbaRuntimeError(3251,"DAO: Property member unsupported");}
    void assign(std::string_view name,const VbaValue& value) override {if(lower(name)!="value") throw VbaRuntimeError(383,"DAO: Property metadata is readonly");(void)state_->column();if(!state_->intrinsic) throw VbaRuntimeError(3251,"DAO: Intrinsic field property context unavailable");(void)state_->intrinsic(key_.toStdString());state_->store(key_,value);}
private:std::shared_ptr<NativeFieldPropertyState> state_;QString key_;
};
class DetachedTextProperty final : public VbaObject {
public:
    DetachedTextProperty(QString key,QString value):key_(std::move(key)),value_(std::move(value)) {}
    std::string_view className() const override {return "Property";}bool isNativeType(std::string_view type) const override {return type=="dao.property";}
    bool hasScalarDefault() const override {return true;}VbaValue scalarValue() override {return member("Value",{});}
    void assignDefaultMember(const std::vector<VbaValue>& args,const VbaValue& value) override {if(!args.empty()) throw VbaRuntimeError(450,"DAO: Scalar property takes no indices");assign("Value",value);}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {if(attached_) return attached_->member(name,args);if(!args.empty()) throw VbaRuntimeError(450,"DAO: Property takes no arguments");const auto key=lower(name);if(key=="name") return VbaValue::string(key_=="caption" ? "Caption" : key_=="description" ? "Description" : "Format");if(key=="type") return VbaValue::integer(10);if(key=="value") return VbaValue::string(value_.toStdString());if(key=="inherited") return VbaValue::logical(false);throw VbaRuntimeError(3251,"DAO: Property member unsupported");}
    void assign(std::string_view name,const VbaValue& value) override {if(attached_) {attached_->assign(name,value);return;}if(lower(name)!="value") throw VbaRuntimeError(383,"DAO: Property metadata is readonly");value_=nativeTextPropertyValue(key_,value);}
    QString key_,value_;std::shared_ptr<VbaObject> attached_;
};
class NativeFieldProperties final : public VbaObject {
public:
    explicit NativeFieldProperties(std::shared_ptr<NativeFieldPropertyState> state):state_(std::move(state)) {}
    std::string_view className() const override {return "Properties";}bool isNativeType(std::string_view type) const override {return type=="dao.properties";}
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {const auto key=lower(name);if(key=="item") return VbaArgumentSignature{{"Index"},1};if(key=="append") return VbaArgumentSignature{{"Object"},1};if(key=="delete") return VbaArgumentSignature{{"Name"},1};return {};}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        (void)state_->column();const auto key=lower(name);if(key=="refresh" && args.empty()) return {};if(key=="count") throw VbaRuntimeError(3251,"DAO: Complete Field.Properties catalog is not implemented");
        if(key=="item" || key=="delete") {if(args.size()!=1) throw VbaRuntimeError(450,"DAO: Properties lookup/delete takes one name");const auto scalar=scalarVbaValue(args[0]);if(scalar.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null property name");if(scalar.kind!=VbaValue::Kind::Text) throw VbaRuntimeError(3251,"DAO: Property ordinal catalog is not implemented");const auto requested=QString::fromStdString(scalar.text).toCaseFolded();if(nativeIntrinsicFieldProperties().contains(requested)) {if(!state_->intrinsic) throw VbaRuntimeError(3251,"DAO: Intrinsic field property context unavailable");(void)state_->intrinsic(requested.toStdString());if(key=="delete") throw VbaRuntimeError(383,"DAO: Built-in field properties cannot be deleted");return VbaValue::reference(std::make_shared<NativeIntrinsicFieldProperty>(state_,requested));}const auto property=nativeTextPropertyKey(requested);if(!state_->value(property)) throw VbaRuntimeError(3270,"DAO: Property not found");if(key=="delete") {state_->store(property,VbaValue{});return {};}auto value=std::make_shared<NativeFieldProperty>(state_,property);value->bindIdentity();return VbaValue::reference(value);}
        if(key=="append") {if(args.size()!=1 || args[0].kind!=VbaValue::Kind::Object) throw VbaRuntimeError(450,"DAO: Properties.Append requires a created Property");const auto property=std::dynamic_pointer_cast<DetachedTextProperty>(args[0].object);if(!property || property->attached_) throw VbaRuntimeError(3219,"DAO: Created unattached text Property required");if(state_->value(property->key_)) throw VbaRuntimeError(3367,"DAO: Property already exists");auto attached=std::make_shared<NativeFieldProperty>(state_,property->key_);state_->store(property->key_,VbaValue::string(property->value_.toStdString()));property->attached_=attached;try {attached->bindIdentity();}catch(...) {throw DaoStoredSchemaChange("DAO: Property stored; metadata reload failed, reopen the project");}return {};}
        throw VbaRuntimeError(3251,"DAO: Field.Properties operation is not implemented");
    }
    std::vector<VbaValue> enumerate() override {throw VbaRuntimeError(3251,"DAO: Complete Field.Properties enumeration is not implemented");}
private:std::shared_ptr<NativeFieldPropertyState> state_;
};
std::shared_ptr<VbaObject> makeNativeFieldProperties(TableSchemaReader reader,QString field,NativeFieldSetter setter,std::function<VbaValue(std::string_view)> intrinsic) {auto state=std::make_shared<NativeFieldPropertyState>(NativeFieldPropertyState{std::move(reader),std::move(field),std::move(setter),std::move(intrinsic)});(void)state->column();return std::make_shared<NativeFieldProperties>(std::move(state));}

class TableSchemaField final : public VbaObject,public std::enable_shared_from_this<TableSchemaField> {
public:
    TableSchemaField(TableSchemaReader reader,QString name,bool defaults=true,NativeFieldSetter setter={}) : reader_(std::move(reader)),name_(std::move(name)),defaults_(defaults),setter_(std::move(setter)) {}
    TableSchemaReader displayPropertyReader() const {return reader_;}
    TableSchemaReader storeDisplayProperty(const QString& key,const VbaValue& value) {if(!setter_) throw VbaRuntimeError(383,"DAO: Field properties are readonly");auto next=setter_(name_,key.toStdString(),value);reader_=std::move(next);return reader_;}
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {if(lower(name)=="properties") return VbaArgumentSignature{{"Index"},0};return {};}
    std::string_view className() const override { return "Field"; }
    bool isNativeType(std::string_view type) const override { return type=="dao.field"; }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        if(lower(name)=="properties") {if(!defaults_) throw VbaRuntimeError(3251,"DAO: Query field properties are not proven");const auto self=shared_from_this();bool native=false;for(const auto& column:reader_()) if(column.name.compare(name_,Qt::CaseInsensitive)==0) native=column.nativeAttributes.has_value();NativeFieldSetter setter;if(setter_ && native) setter=[self](const QString&,std::string_view key,const VbaValue& value) {if(key=="caption" || key=="description" || key=="format") return self->storeDisplayProperty(QString::fromUtf8(key.data(),key.size()),value);self->assign(key,value);return self->displayPropertyReader();};auto properties=makeNativeFieldProperties(reader_,name_,setter,[self](std::string_view key) {return self->member(key,{});});if(args.empty() || (args.size()==1 && args[0].kind==VbaValue::Kind::Missing)) return VbaValue::reference(properties);return properties->member("Item",args);}
        if(!args.empty()) throw VbaRuntimeError(450,"DAO schema field properties take no arguments");
        const auto columns=reader_();
        int position=-1;
        for(int i=0;i<columns.size();++i) if(columns[i].name.compare(name_,Qt::CaseInsensitive)==0) {
            if(position>=0) throw VbaRuntimeError(5,"DAO schema field name ambiguous"); position=i;
        }
        if(position<0) throw VbaRuntimeError(3420,"DAO schema field unavailable");
        const auto key=lower(name);
        if(key=="name") return VbaValue::string(columns[position].name.toStdString());
        if(defaults_ && key=="validationtext" && columns[position].nativeAttributes) return VbaValue::string(columns[position].validationText.toStdString());
        if(defaults_ && key=="validationrule" && columns[position].nativeAttributes) return VbaValue::string(columns[position].validationRule.toStdString());
        if(defaults_ && key=="required" && columns[position].required) return VbaValue::logical(*columns[position].required);
        if(defaults_ && key=="allowzerolength" && columns[position].allowZeroLength) return VbaValue::logical(*columns[position].allowZeroLength);
        if(defaults_ && columns[position].nativeAttributes) {
            if(key=="attributes") return VbaValue::longInteger(*columns[position].nativeAttributes);
            if(key=="dataupdatable") return VbaValue::logical(false);
        }
        if(key=="attributes" && columns[position].indexAttributes>=0) return VbaValue::longInteger(columns[position].indexAttributes);
        if(defaults_ && columns[position].generatorAttributes>=0) {
            if(key=="attributes") return VbaValue::longInteger(columns[position].generatorAttributes);
            if(key=="dataupdatable") return VbaValue::logical(false); // TableDef field has no data buffer.
        }
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
    void assignIndexedMember(std::string_view name,const std::vector<VbaValue>& args,const VbaValue& value) override {if(lower(name)!="properties" || args.size()!=1) throw VbaRuntimeError(438,"DAO: Indexed field property assignment unsupported");member("Properties",args).object->assign("Value",value);}
    void assign(std::string_view name,const VbaValue& value) override {
        (void)reader_();const auto key=lower(name);
        if(!setter_ || (key!="required" && key!="defaultvalue" && key!="size" && key!="name" && key!="allowzerolength" && key!="validationrule" && key!="validationtext" && key!="ordinalposition")) throw VbaRuntimeError(383,"DAO schema field property is read-only");
        bool native=false;for(const auto& column:reader_()) if(column.name.compare(name_,Qt::CaseInsensitive)==0) native=column.nativeAttributes.has_value() && column.required.has_value();
        if(!native) throw VbaRuntimeError(383,"DAO: Original/unproven schema fields remain read-only");
        const auto scalar=scalarVbaValue(value);(void)reader_();auto next=setter_(name_,key,scalar);reader_=std::move(next);if(key=="name") name_=text(scalar.asText());
    }
private: TableSchemaReader reader_; QString name_; bool defaults_{};NativeFieldSetter setter_;
};
std::shared_ptr<VbaObject> createExistingTableField(QSqlDatabase db,const QString& table,const std::vector<VbaValue>& args);
void mutateExistingTableField(QSqlDatabase db,const QString& table,const VbaValue& value,bool erase);
class TableSchemaFields final : public VbaObject {
public:
    explicit TableSchemaFields(TableSchemaReader reader,bool defaults=true,std::function<void(const VbaValue&,bool)> mutation={},std::function<TableSchemaReader()> refresher={},NativeFieldSetter setter={}) : reader_(std::move(reader)),defaults_(defaults),mutation_(std::move(mutation)),refresher_(std::move(refresher)),setter_(std::move(setter)) {}
    std::string_view className() const override { return "Fields"; }
    bool isNativeType(std::string_view type) const override { return type=="dao.fields"; }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        const auto key=lower(name);if(key=="item") return VbaArgumentSignature{{"Index"},1};if(mutation_ && key=="append") return VbaArgumentSignature{{"Object"},1};if(mutation_ && key=="delete") return VbaArgumentSignature{{"Name"},1};return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        const auto key=lower(name);
        if(key=="refresh" && refresher_) {
            if(!args.empty()) throw VbaRuntimeError(450,"DAO: Fields.Refresh takes no arguments");
            if(*epoch_==std::numeric_limits<quint64>::max()) throw VbaRuntimeError(6,"DAO: Field collection generation overflow");
            auto candidate=refresher_();(void)candidate();reader_=std::move(candidate);++*epoch_;fields_.clear();return {};
        }
        const auto columns=reader_();
        if(mutation_ && (key=="append" || key=="delete")) {if(args.size()!=1) throw VbaRuntimeError(450,"DAO: Field/name required");mutation_(args[0],key=="delete");return {};}
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
        if(!field) {const auto reader=reader_;const auto epoch=epoch_;const auto generation=*epoch;
            field=std::make_shared<TableSchemaField>([reader,epoch,generation] {if(*epoch!=generation) throw VbaRuntimeError(3420,"DAO: Field object predates collection refresh");return reader();},column,defaults_,setter_);cached=field;}
        return VbaValue::reference(field);
    }
    VbaValue item(std::string_view name) override { return member("Item",{VbaValue::string(std::string(name))}); }
    std::vector<VbaValue> enumerate() override {
        const auto columns=reader_(); std::vector<VbaValue> result;
        for(const auto& column:columns) result.push_back(member("Item",{VbaValue::string(column.name.toStdString())}));
        return result;
    }
private: TableSchemaReader reader_; QMap<QString,std::weak_ptr<TableSchemaField>> fields_; bool defaults_{};std::function<void(const VbaValue&,bool)> mutation_;std::function<TableSchemaReader()> refresher_;std::shared_ptr<quint64> epoch_{std::make_shared<quint64>(0)};NativeFieldSetter setter_;
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
NativeRecordsetIndex checkedRecordsetIndex(QSqlDatabase db,const QString& source,const QString& requested) {
    QString table;for(const auto& name:nativeTableNames(db)) if(name.compare(source,Qt::CaseInsensitive)==0) {table=name;break;}if(table.isEmpty()) throw VbaRuntimeError(3251,"DAO: Index requires a physical main table");std::optional<AccessSourceIndex> selected;for(const auto& index:readNativeIndexCatalog(db,table)) if(index.source.name.compare(requested,Qt::CaseInsensitive)==0) selected=index.source;if(!selected) throw VbaRuntimeError(3265,"DAO: Index not found");const auto schema=readTableSchemaColumns(db,table);QSet<QString> required;QMap<QString,QString> declarations;QSqlQuery columns(db);if(!columns.exec("PRAGMA main.table_info("+quote(table)+")")) queryFailure(columns);while(columns.next()) {const auto name=columns.value(1).toString();if(columns.value(3).toBool()) required.insert(name.toCaseFolded());declarations.insert(name.toCaseFolded(),columns.value(2).toString().trimmed().toUpper());}
    NativeRecordsetIndex result{selected->name,selected->columns,{},selected->descending};for(const auto& name:selected->columns) {std::optional<TableSchemaColumn> field;for(const auto& column:schema) if(column.name==name) field=column;if(!field || !field->nativeAttributes || !required.contains(name.toCaseFolded()) || (field->sourceType!=1 && field->sourceType!=2 && field->sourceType!=3 && field->sourceType!=4 && field->sourceType!=6 && field->sourceType!=7 && !(field->sourceType==8 && field->canonicalDate.value_or(false) && declarations.value(name.toCaseFolded())=="DATETIME")) || (field->sourceType==6 && declarations.value(name.toCaseFolded())!="SINGLE_REAL")) throw VbaRuntimeError(3251,"DAO: Index navigation requires non-null native numeric or canonical Date keys");result.types.append(field->type);}return result;
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
std::shared_ptr<VbaObject> createExistingTableIndex(QSqlDatabase db,const QString& table,const QString& name);
void mutateExistingTableIndex(QSqlDatabase db,const QString& table,const VbaValue& value,bool erase);
class NativeIndexes final : public VbaObject {
public:
    NativeIndexes(NativeIndexReader reader,QSqlDatabase db,QString table,std::function<NativeIndexReader()> refresher={}) : reader_(std::move(reader)),db_(std::move(db)),table_(std::move(table)),refresher_(std::move(refresher)) {}
    std::string_view className() const override { return "Indexes"; }
    bool isNativeType(std::string_view type) const override { return type=="dao.indexes"; }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        const auto key=lower(name);if(key=="item") return VbaArgumentSignature{{"Index"},1};if(key=="append") return VbaArgumentSignature{{"Object"},1};if(key=="delete") return VbaArgumentSignature{{"Name"},1};return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        const auto key=lower(name);
        if(key=="refresh" && refresher_) {if(!args.empty()) throw VbaRuntimeError(450,"DAO: Indexes.Refresh takes no arguments");auto candidate=refresher_();(void)candidate();reader_=std::move(candidate);objects_.clear();return {};}
        const auto rows=reader_();
        if(key=="count" || key=="refresh") {
            if(!args.empty()) throw VbaRuntimeError(450,"DAO: Index collection property takes no arguments");
            return key=="count" ? integerCount(rows.size()) : VbaValue{};
        }
        if(key=="append" || key=="delete") {if(args.size()!=1) throw VbaRuntimeError(450,"DAO: Index/name required");mutateExistingTableIndex(db_,table_,args[0],key=="delete");return {};}
        if(key!="item") throw VbaRuntimeError(3251,"DAO: Index catalog member unsupported");
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
    NativeIndexReader reader_; QSqlDatabase db_;QString table_;std::function<NativeIndexReader()> refresher_;QMap<QString,std::weak_ptr<NativeIndex>> objects_;
};

quint64 renameExistingNativeTable(QSqlDatabase db,const QString& table,const QString& renamed);
class TableDefinition final : public VbaObject, public std::enable_shared_from_this<TableDefinition> {
public:
    std::string_view className() const override { return "TableDef"; }
    bool isNativeType(std::string_view type) const override { return type.starts_with("dao.") && type.substr(4)==lower(className()); }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        const auto key=lower(name);
        if(key=="createfield") return VbaArgumentSignature{{"Name","Type","Size"},0};
        if(key=="createindex") return VbaArgumentSignature{{"Name"},0};
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
        if(key=="updatable") {if(!args.empty()) throw VbaRuntimeError(450,"DAO: Updatable takes no arguments");return VbaValue::logical(connectionUpdatable(db_));}
        if(key=="attributes" || key=="connect" || key=="sourcetablename") {
            if(!args.empty()) throw VbaRuntimeError(450,"DAO: Local TableDef property takes no arguments");for(const auto& column:readTableSchemaColumns(db_,name_)) if(!column.nativeAttributes || !column.required) throw VbaRuntimeError(3251,"DAO: Local TableDef declaration not proven");return key=="attributes" ? VbaValue::longInteger(0) : VbaValue::string("");
        }
        if(key=="recordcount") {
            if(!args.empty()) throw VbaRuntimeError(450,"DAO: RecordCount takes no arguments");
            SqliteExecutionBudget budget(db_); QSqlQuery count(db_);
            if(!count.exec("SELECT COUNT(*) FROM main."+quote(name_)) || !count.next()) queryFailure(count);
            return longCount(count.value(0).toLongLong());
        }
        if(key=="createfield") return VbaValue::reference(createExistingTableField(db_,name_,args));
        if(key=="createindex") {
            if(args.size()>1) throw VbaRuntimeError(450,"DAO: CreateIndex takes Name only");
            if(!args.empty() && args[0].kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null index name");
            return VbaValue::reference(createExistingTableIndex(db_,name_,args.empty() ? QString{} : text(args[0].asText())));
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
                },db_,name_,[self]() -> NativeIndexReader {
                    if(!self->valid()) throw VbaRuntimeError(3420,"DAO: Index table unavailable");
                    const auto version=self->schemaVersion(),rollback=self->workspace_->rollbackRevision;self->indexes_.reset();
                    return [self,version,rollback] {if(!self->valid() || self->schemaVersion()!=version || self->workspace_->rollbackRevision!=rollback) throw VbaRuntimeError(3420,"DAO: Refreshed index metadata invalidated");return readNativeIndexCatalog(self->db_,self->name_);};
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
                },true,[self](const VbaValue& value,bool erase) {
                    if(!self->valid()) throw VbaRuntimeError(3420,"DAO: Field table unavailable");
                    mutateExistingTableField(self->db_,self->name_,value,erase);
                },[self]() -> TableSchemaReader {
                    if(!self->valid()) throw VbaRuntimeError(3420,"DAO: Field table unavailable");
                    const auto version=self->schemaVersion(),rollback=self->workspace_->rollbackRevision;self->fields_.reset();
                    return [self,version,rollback] {if(!self->valid() || self->schemaVersion()!=version || self->workspace_->rollbackRevision!=rollback) throw VbaRuntimeError(3420,"DAO: Refreshed field metadata invalidated");return readTableSchemaColumns(self->db_,self->name_);};
                },[self](const QString& field,std::string_view property,const VbaValue& value) -> TableSchemaReader {
                    if(!self->valid()) throw VbaRuntimeError(3420,"DAO: Field table unavailable");
                    return updateNativeFieldPropertiesImpl(self->db_,self->name_,field,{{text(property),value}});
                });
                fields_=fields; fieldsVersion_=version; fieldsRollback_=rollback;
            }
            return args.empty() || args[0].kind==VbaValue::Kind::Missing ? VbaValue::reference(fields) : fields->member("Item",args);
        }
        if (key=="openrecordset") {
            if (args.size()>1) throw VbaRuntimeError(450,"DAO: TableDef OpenRecordset arguments unsupported");
            const int mode=nativeRecordsetMode(args,0,1);
            return VbaValue::reference(std::make_shared<Recordset>(db_,name_,mode==4,QMap<QString,QVariant>{},AccessQueryPlan{},mode));
        }
        throw VbaRuntimeError(438,"DAO: TableDef member unsupported");
    }
    void assign(std::string_view name,const VbaValue& value) override {
        if(!valid()) throw VbaRuntimeError(3420,"DAO: TableDef is unavailable");
        if(lower(name)=="updatable") throw VbaRuntimeError(383,"DAO: TableDef.Updatable is read-only");
        if(lower(name)=="attributes" || lower(name)=="connect" || lower(name)=="sourcetablename") {
            (void)member(name,{});const auto scalar=scalarVbaValue(value);if(!valid()) throw VbaRuntimeError(3420,"DAO: Table identity changed during property evaluation");(void)member(name,{});if(scalar.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null local TableDef property");
            if(lower(name)=="attributes" ? coerceVbaModuleValue(scalar,"long").number!=0 : !scalar.asText().empty()) throw VbaRuntimeError(3251,"DAO: Linked/system TableDef changes are not implemented");return;
        }
        if(lower(name)=="name") {
            const auto scalar=scalarVbaValue(value);if(!valid()) throw VbaRuntimeError(3420,"DAO: Table identity changed while evaluating its new name");if(scalar.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null table name");auto renamed=text(scalar.asText());const auto generation=renameExistingNativeTable(db_,name_,renamed);name_=std::move(renamed);generation_=generation;fields_.reset();indexes_.reset();return;
        }
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
// Detached DAO schema objects; mutations become physical only at TableDefs.Append.
QString schemaJetName(QString name) {return '['+name.replace("]","]]")+']';}
struct DetachedTableState {
    QSqlDatabase db;std::shared_ptr<WorkspaceState> workspace;QString name;
    std::shared_ptr<TableDefinition> attached,existing;
};
class DetachedField final : public VbaObject,public std::enable_shared_from_this<DetachedField> {
public:
    DetachedField(std::shared_ptr<DetachedTableState> owner,QString name,int type,int size):owner_(std::move(owner)),name_(std::move(name)),type_(type),size_(size) {
        attributes_=(type==10 || type==11 || type==12) ? 2 : 1;
        if(type!=10) {switch(type) {case 1:case 2:size_=1;break;case 3:size_=2;break;case 4:case 6:size_=4;break;case 5:case 7:case 8:size_=8;break;case 15:size_=16;break;case 11:case 12:size_=0;break;default:break;}}
    }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {if(lower(name)=="properties") return VbaArgumentSignature{{"Index"},0};return {};}
    TableSchemaReader displayPropertyReader() {const auto self=shared_from_this();return [self] {self->owner_->workspace->check();if(self->attachedField_) {const auto field=std::dynamic_pointer_cast<TableSchemaField>(self->attachedField_);if(!field) throw VbaRuntimeError(3251,"DAO: Attached field property metadata unavailable");return field->displayPropertyReader()();}if(self->owner_->attached || (self->owner_->existing && !self->owner_->existing->valid())) throw VbaRuntimeError(3420,"DAO: Detached property field unavailable");TableSchemaColumn field;field.name=self->name_;field.sourceType=self->type_;field.nativeAttributes=self->attributes_;field.required=self->required_;field.caption=self->caption_;field.description=self->description_;field.format=self->format_;field.captionIdentity=self->captionIdentity_;field.descriptionIdentity=self->descriptionIdentity_;field.formatIdentity=self->formatIdentity_;return QVector<TableSchemaColumn>{field};};}
    std::string_view className() const override {return "Field";}
    bool isNativeType(std::string_view type) const override {return type=="dao.field";}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        owner_->workspace->check();if(owner_->existing && !owner_->existing->valid()) throw VbaRuntimeError(3420,"DAO: Field builder table invalidated");if(!args.empty() && lower(name)!="properties") throw VbaRuntimeError(450,"DAO: Detached field properties take no arguments");
        const auto key=lower(name);
        if(attachedField_) return attachedField_->member(name,args);
        if(owner_->attached) throw VbaRuntimeError(3420,"DAO: Detached field was not attached");
        if(key=="properties") {const auto self=shared_from_this();auto properties=makeNativeFieldProperties(displayPropertyReader(),name_,[self](const QString&,std::string_view name,const VbaValue& value) {const auto key=QString::fromUtf8(name.data(),name.size());(void)self->displayPropertyReader()();if(key!="caption" && key!="description" && key!="format") {self->assign(name,value);return self->displayPropertyReader();}if(self->attachedField_) {const auto field=std::dynamic_pointer_cast<TableSchemaField>(self->attachedField_);return field->storeDisplayProperty(key,value);}auto& slot=key=="caption" ? self->caption_ : key=="description" ? self->description_ : self->format_;auto& identity=key=="caption" ? self->captionIdentity_ : key=="description" ? self->descriptionIdentity_ : self->formatIdentity_;if(value.kind==VbaValue::Kind::Null) {slot.reset();identity.clear();}else {const auto text=nativeTextPropertyValue(key,value);if(!slot) identity=QUuid::createUuid().toString(QUuid::WithoutBraces);slot=text;}return self->displayPropertyReader();},[self](std::string_view key) {return self->member(key,{});});if(args.empty() || (args.size()==1 && args[0].kind==VbaValue::Kind::Missing)) return VbaValue::reference(properties);return properties->member("Item",args);}
        if(key=="name") return VbaValue::string(name_.toStdString());
        if(key=="type") return VbaValue::integer(static_cast<std::int16_t>(type_));
        if(key=="size") return VbaValue::longInteger(size_);
        if(key=="attributes") return VbaValue::longInteger(attributes_);
        if(key=="required") return VbaValue::logical(required_);
        if(key=="allowzerolength" && (type_==10 || type_==12)) return VbaValue::logical(allowZeroLength_);
        if(key=="defaultvalue") return VbaValue::string(default_.toStdString());
        if(key=="validationtext") return VbaValue::string(validationText_.toStdString());
        if(key=="validationrule") return VbaValue::string(validationRule_.toStdString());
        throw VbaRuntimeError(3251,"DAO: Detached field property unsupported");
    }
    void assignIndexedMember(std::string_view name,const std::vector<VbaValue>& args,const VbaValue& value) override {if(lower(name)!="properties" || args.size()!=1) throw VbaRuntimeError(438,"DAO: Indexed field property assignment unsupported");member("Properties",args).object->assign("Value",value);}
    void assign(std::string_view name,const VbaValue& raw) override {
        owner_->workspace->check();if(owner_->existing && !owner_->existing->valid()) throw VbaRuntimeError(3420,"DAO: Field builder table invalidated");if(attachedField_) {attachedField_->assign(name,raw);return;}if(owner_->attached) throw VbaRuntimeError(383,"DAO: Attached schema field is read-only");
        const auto value=scalarVbaValue(raw);if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null schema field property");
        const auto key=lower(name);
        if(key=="name") {name_=text(value.asText());return;}
        if(key=="type") {type_=static_cast<int>(coerceVbaModuleValue(value,"integer").number);
            if(type_!=10) {const DetachedField intrinsic(owner_,name_,type_,0);size_=intrinsic.size_;}
            attributes_=(type_==10 || type_==11 || type_==12) ? 2 : 1;return;}
        if(key=="size") {const auto size=static_cast<int>(coerceVbaModuleValue(value,"long").number);if(type_==10) size_=size;return;}
        if(key=="attributes") {const auto attributes=coerceVbaModuleValue(value,"long").number;if(attributes!=0 && attributes!=1 && attributes!=2 && attributes!=16 && attributes!=17) throw VbaRuntimeError(3251,"DAO: Unsupported field attributes");const int intrinsic=(type_==10 || type_==11 || type_==12) ? 2 : 1;
            if(((static_cast<int>(attributes)&3) && (static_cast<int>(attributes)&3)!=intrinsic) || ((static_cast<int>(attributes)&16) && type_!=4)) throw VbaRuntimeError(3251,"DAO: Attributes conflict with the field type");
            attributes_=(static_cast<int>(attributes)&16)|intrinsic;return;}
        if(key=="required") {required_=coerceVbaModuleValue(value,"boolean").truthy();return;}
        if(key=="allowzerolength") {if(type_!=10 && type_!=12) throw VbaRuntimeError(3251,"DAO: AllowZeroLength requires Text or Memo");allowZeroLength_=coerceVbaModuleValue(value,"boolean").truthy();return;}
        if(key=="validationtext") {const auto message=text(value.asText());if(message.size()>255 || message.contains(QChar(0)) || !message.isValidUtf16() || (!message.isEmpty() && ((attributes_&16) || (type_!=1 && type_!=2 && type_!=3 && type_!=4 && type_!=6 && type_!=7 && !(type_==8 && canonicalDate_))))) throw VbaRuntimeError(3251,"DAO: ValidationText requires a native numeric or canonical Date non-counter field and at most255 valid characters");validationText_=message;return;}
        if(key=="validationrule") {const auto rule=text(value.asText()).trimmed();if(!rule.isEmpty() && ((attributes_&16) || (type_!=1 && type_!=2 && type_!=3 && type_!=4 && type_!=6 && type_!=7 && !(type_==8 && canonicalDate_)))) throw VbaRuntimeError(3251,"DAO: ValidationRule requires native numeric or canonical Date non-counter fields");(void)nativeFieldValidationPredicate(type_,rule,schemaJetName(name_));validationRule_=rule;return;}
        if(key=="defaultvalue") {const auto token=text(value.asText());if(!token.isEmpty() && !accessLiteralDefaultSql(token) && !(type_==8 && (nativeDateDefaultSql(token) || nativeDateLiteralSql(token)))) throw VbaRuntimeError(3251,"DAO: Field default requires a supported literal or Date clock expression");default_=token;return;}
        throw VbaRuntimeError(3251,"DAO: Detached field assignment unsupported");
    }
    QJsonObject metadata() const {
        const auto source=type_==1 ? "BOOLEAN" : type_==2 ? "TINYINT" : type_==3 ? "SMALLINT" : type_==4 ? "INT" : type_==6 ? "SINGLE" : type_==7 ? "DOUBLE" : type_==8 ? "DATETIME" : type_==10 ? "VARCHAR" : type_==11 ? "BLOB" : "MEMO";
        auto object=QJsonObject{{"name",name_},{"source_type",source},{"source_origin","DAO_NATIVE"},{"source_size",type_==10 ? size_ : 0},{"native_attributes",attributes_},{"native_required",required_}};if(type_==10 || type_==12) object.insert("native_allow_zero_length",allowZeroLength_);if(type_==8) object.insert("native_date_canonical",canonicalDate_);if(!validationRule_.isEmpty()) object.insert("native_validation_rule",validationRule_);if(!validationText_.isEmpty()) object.insert("native_validation_text",validationText_);if(caption_) {object.insert("native_caption",*caption_);object.insert("native_caption_id",captionIdentity_);}if(description_) {object.insert("native_description",*description_);object.insert("native_description_id",descriptionIdentity_);}if(format_) {object.insert("native_format",*format_);object.insert("native_format_id",formatIdentity_);}return object;
    }
    QString definition(bool storedDateLiteral=false) const {
        if(name_.isEmpty() || name_.trimmed()!=name_ || name_.size()>64 || name_.contains(QChar(0)) || !name_.isValidUtf16()) throw VbaRuntimeError(5,"DAO: Invalid field name");
        if(attributes_&16) {
            if(type_!=4 || !default_.isEmpty() || !validationRule_.isEmpty() || !validationText_.isEmpty()) throw VbaRuntimeError(3251,"DAO: AutoNumber requires a Long field without a scalar default or numeric ValidationRule");
            return schemaJetName(name_)+" COUNTER";
        }
        QString type;
        switch(type_) {
        case 1:type="BOOLEAN";break;case 2:type="INTEGER1";break;case 3:type="INTEGER2";break;case 4:type="INTEGER4";break;
        case 6:type="SINGLE_REAL";break;
        case 7:type="DOUBLE";break;case 8:type="DATETIME";break;case 10:
            if(size_<1 || size_>255) throw VbaRuntimeError(3251,"DAO: Text field requires an explicit supported size");
            type="VARCHAR("+QString::number(size_)+")";break;
        case 11:type="BLOB";break;case 12:type="TEXT";break;
        default:throw VbaRuntimeError(3251,"DAO: Field type is not supported for native creation");
        }
        auto result=schemaJetName(name_)+" "+type+(required_ ? " NOT NULL" : "");
        if(type_==8 && nativeDateDefaultSql(default_)) result+=" DEFAULT ("+*nativeDateDefaultSql(default_)+")";
        else if(type_==8 && nativeDateLiteralSql(default_)) result+=" DEFAULT "+*nativeDateLiteralSql(default_);
        else if(!default_.isEmpty()) {auto literal=*accessLiteralDefaultSql(default_);if(type_==8 && !storedDateLiteral) {QSqlQuery query(owner_->db);if(!query.exec("SELECT "+literal) || !query.next()) queryFailure(query);const auto stored=sqlFieldValue(fromSql(query.value(0)),"DATETIME");if(stored.isNull()) literal="NULL";else {auto value=stored.toString();literal="'"+value.replace("'","''")+"'";}}if(type_==6) {QSqlQuery query(owner_->db);if(!query.exec("SELECT "+literal) || !query.next()) queryFailure(query);const auto stored=sqlFieldValue(fromSql(query.value(0)),"SINGLE_REAL");literal=stored.isNull() ? "NULL" : QString::number(stored.toDouble(),'g',17);}result+=" DEFAULT "+literal;}
        if(type_==6) result+=" CHECK ("+schemaJetName(name_)+" IS NULL OR (typeof("+schemaJetName(name_)+") IN ('integer','real') AND "+schemaJetName(name_)+"=linaccess_jet_field_store('SINGLE_REAL',"+schemaJetName(name_)+")))";
        if(type_==10) result+=" CHECK (linaccess_jet_text_capacity("+schemaJetName(name_)+","+QString::number(size_)+"))";
        const auto column=schemaJetName(name_);
        if((type_==10 || type_==12) && !allowZeroLength_) result+=" CHECK (linaccess_jet_nonempty_text("+column+"))";
        if(type_==2 || type_==3 || type_==4 || type_==1) result+=" CHECK ("+column+" IS NULL OR (typeof("+column+")='integer' AND "+column+"=linaccess_jet_field_store('"+type+"',"+column+")))";
        if(type_==7) result+=" CHECK ("+column+" IS NULL OR (typeof("+column+") IN ('integer','real') AND abs("+column+")<=1.7976931348623157e308))";
        if(type_==8) result+=" CHECK ("+column+" IS NULL OR (typeof("+column+")='text' AND "+(canonicalDate_ ? column+"=linaccess_jet_field_store('DATETIME',"+column+")" : "linaccess_jet_field_store('DATETIME',"+column+") IS NOT NULL")+"))";
        if(type_==11 || type_==12) result+=" CHECK ("+column+" IS NULL OR typeof("+column+")='"+QString(type_==11 ? "blob" : "text")+"')";
        if(!validationRule_.isEmpty()) result+=" CHECK (linaccess_jet_validation("+nativeFieldValidationPredicate(type_,validationRule_,column)+(validationText_.isEmpty() ? QString{} : ",'"+QString(validationText_).replace("'","''")+"'")+"))";
        return result;
    }
    std::shared_ptr<VbaObject> attachedField_;std::shared_ptr<DetachedTableState> owner_;QString name_;int type_{},size_{},attributes_{};bool required_{},appended_{},allowZeroLength_{true},canonicalDate_{true};QString default_,validationRule_,validationText_;std::optional<QString> caption_,description_,format_;QString captionIdentity_,descriptionIdentity_,formatIdentity_;
};
class DetachedFields final : public VbaObject {
public:
    explicit DetachedFields(std::shared_ptr<DetachedTableState> owner):owner_(std::move(owner)) {}
    std::string_view className() const override {return "Fields";}
    bool isNativeType(std::string_view type) const override {return type=="dao.fields";}
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        const auto key=lower(name);if(key=="item") return VbaArgumentSignature{{"Index"},1};if(key=="append") return VbaArgumentSignature{{"Object"},1};if(key=="delete") return VbaArgumentSignature{{"Name"},1};return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        owner_->workspace->check();const auto key=lower(name);
        if(owner_->attached) return owner_->attached->member("Fields",{}).object->member(name,args);
        if(key=="count" && args.empty()) return integerCount(fields_.size());
        if(key=="append") {
            if(args.size()!=1 || args[0].kind!=VbaValue::Kind::Object) throw VbaRuntimeError(450,"DAO: Fields.Append requires a field");
            auto field=std::dynamic_pointer_cast<DetachedField>(args[0].object);
            if(!field || field->owner_!=owner_ || field->appended_) throw VbaRuntimeError(3219,"DAO: Field belongs to a different or attached schema");
            if(fields_.size()>=255) throw VbaRuntimeError(3190,"DAO: Too many fields defined");
            (void)field->definition();for(const auto& existing:fields_) if(existing->name_.compare(field->name_,Qt::CaseInsensitive)==0) throw VbaRuntimeError(3191,"DAO: Duplicate field name");
            field->appended_=true;fields_.append(field);return {};
        }
        if(key=="refresh" && args.empty()) return {};
        if((key!="item" && key!="delete") || args.size()!=1) throw VbaRuntimeError(450,"DAO: Field name/index required");
        const auto value=scalarVbaValue(args[0]);if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null field index");
        int position=-1;
        if(value.kind==VbaValue::Kind::Text) {for(int i=0;i<fields_.size();++i) if(fields_[i]->name_.compare(text(value.asText()),Qt::CaseInsensitive)==0) position=i;}
        else {const auto index=coerceVbaModuleValue(value,"long").number;if(index>=0 && index<fields_.size()) position=static_cast<int>(index);}
        if(position<0) throw VbaRuntimeError(3265,"DAO: Field not found");
        if(key=="delete") {fields_.removeAt(position);return {};}
        return VbaValue::reference(fields_[position]);
    }
    VbaValue item(std::string_view name) override {return member("Item",{VbaValue::string(std::string(name))});}
    std::vector<VbaValue> enumerate() override {owner_->workspace->check();if(owner_->attached) return owner_->attached->member("Fields",{}).object->enumerate();std::vector<VbaValue> result;for(const auto& field:fields_) result.push_back(VbaValue::reference(field));return result;}
    std::shared_ptr<DetachedTableState> owner_;QVector<std::shared_ptr<DetachedField>> fields_;
};
struct DetachedIndexState {
    std::shared_ptr<DetachedTableState> table;AccessSourceIndex data;bool appended{};std::shared_ptr<VbaObject> attachedIndex;
    void check() const {table->workspace->check();if(table->attached && !attachedIndex) throw VbaRuntimeError(3420,"DAO: Detached index was not attached");if(attachedIndex) (void)attachedIndex->member("Name",{});if(table->existing && !table->existing->valid()) throw VbaRuntimeError(3420,"DAO: Index builder table invalidated");if(table->attached && !table->attached->valid()) throw VbaRuntimeError(3420,"DAO: Index table unavailable");}
    void mutableCheck() const {check();if(table->attached) throw VbaRuntimeError(383,"DAO: Attached index is read-only");}
};
class DetachedIndexField final : public VbaObject {
public:
    DetachedIndexField(std::shared_ptr<DetachedIndexState> owner,QString name):owner_(std::move(owner)),name_(std::move(name)) {}
    std::string_view className() const override {return "Field";}
    bool isNativeType(std::string_view type) const override {return type=="dao.field";}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        owner_->check();if(!args.empty()) throw VbaRuntimeError(450,"DAO: Index field property takes no arguments");
        if(owner_->table->attached) return owner_->attachedIndex->member("Fields",{VbaValue::string(name_.toStdString())}).object->member(name,args);
        const auto key=lower(name);if(key=="name") return VbaValue::string(name_.toStdString());if(key=="attributes") return VbaValue::longInteger(descending_ ? 1 : 0);
        throw VbaRuntimeError(3251,"DAO: Index field property unsupported");
    }
    void assign(std::string_view name,const VbaValue& raw) override {
        owner_->mutableCheck();const auto value=scalarVbaValue(raw);if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null index field property");
        if(lower(name)=="name") {name_=text(value.asText());return;}
        if(lower(name)=="attributes") {const auto flag=coerceVbaModuleValue(value,"long").number;if(flag!=0 && flag!=1) throw VbaRuntimeError(3251,"DAO: Only dbDescending index field attribute supported");descending_=flag==1;return;}
        throw VbaRuntimeError(3251,"DAO: Index field assignment unsupported");
    }
    std::shared_ptr<DetachedIndexState> owner_;QString name_;bool descending_{},appended_{};
};
class DetachedIndexFields final : public VbaObject {
public:
    explicit DetachedIndexFields(std::shared_ptr<DetachedIndexState> owner):owner_(std::move(owner)) {}
    std::string_view className() const override {return "Fields";}
    bool isNativeType(std::string_view type) const override {return type=="dao.fields";}
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {const auto key=lower(name);if(key=="append") return VbaArgumentSignature{{"Object"},1};if(key=="item") return VbaArgumentSignature{{"Index"},1};return {};}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        owner_->check();const auto key=lower(name);
        if(owner_->table->attached) return owner_->attachedIndex->member("Fields",{}).object->member(name,args);
        if(key=="count" && args.empty()) return integerCount(fields_.size());
        if(key=="refresh" && args.empty()) return {};
        if(key=="append") {
            owner_->mutableCheck();if(args.size()!=1 || args[0].kind!=VbaValue::Kind::Object) throw VbaRuntimeError(450,"DAO: Index field required");
            auto field=std::dynamic_pointer_cast<DetachedIndexField>(args[0].object);
            if(!field || field->owner_!=owner_ || field->appended_) throw VbaRuntimeError(3219,"DAO: Foreign index field");
            if(fields_.size()>=10) throw VbaRuntimeError(3251,"DAO: Index field bound exceeded");
            for(const auto& existing:fields_) if(existing->name_.compare(field->name_,Qt::CaseInsensitive)==0) throw VbaRuntimeError(3191,"DAO: Duplicate index field");
            field->appended_=true;fields_.append(field);return {};
        }
        if(key!="item" || args.size()!=1) throw VbaRuntimeError(450,"DAO: Index field name/index required");
        const auto value=scalarVbaValue(args[0]);if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null index field index");
        if(value.kind==VbaValue::Kind::Text) {for(const auto& field:fields_) if(field->name_.compare(text(value.asText()),Qt::CaseInsensitive)==0) return VbaValue::reference(field);}
        else {const auto index=coerceVbaModuleValue(value,"long").number;if(index>=0 && index<fields_.size()) return VbaValue::reference(fields_[static_cast<int>(index)]);}
        throw VbaRuntimeError(3265,"DAO: Index field not found");
    }
    VbaValue item(std::string_view name) override {return member("Item",{VbaValue::string(std::string(name))});}
    std::vector<VbaValue> enumerate() override {owner_->check();if(owner_->table->attached) return owner_->attachedIndex->member("Fields",{}).object->enumerate();std::vector<VbaValue> result;for(const auto& field:fields_) result.push_back(VbaValue::reference(field));return result;}
    std::shared_ptr<DetachedIndexState> owner_;QVector<std::shared_ptr<DetachedIndexField>> fields_;
};
class DetachedIndex final : public VbaObject {
public:
    DetachedIndex(std::shared_ptr<DetachedTableState> owner,QString name):state_(std::make_shared<DetachedIndexState>()) {state_->table=std::move(owner);state_->data.name=std::move(name);fields_=std::make_shared<DetachedIndexFields>(state_);}
    std::string_view className() const override {return "Index";}
    bool isNativeType(std::string_view type) const override {return type=="dao.index";}
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {const auto key=lower(name);if(key=="createfield") return VbaArgumentSignature{{"Name"},0};if(key=="fields") return VbaArgumentSignature{{"Index"},0};return {};}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        state_->check();const auto key=lower(name);
        if(state_->table->attached) return state_->attachedIndex->member(name,args);
        if(key=="fields") return args.empty() ? VbaValue::reference(fields_) : fields_->member("Item",args);
        if(key=="createfield") {if(args.size()>1) throw VbaRuntimeError(450,"DAO: Index.CreateField takes Name only");return VbaValue::reference(std::make_shared<DetachedIndexField>(state_,args.empty() ? QString{} : text(args[0].asText())));}
        if(!args.empty()) throw VbaRuntimeError(450,"DAO: Index property takes no arguments");
        if(key=="name") return VbaValue::string(state_->data.name.toStdString());
        if(key=="primary") return VbaValue::logical(state_->data.kind==1);
        if(key=="unique") return VbaValue::logical(state_->data.flags&1);
        if(key=="required") return VbaValue::logical(state_->data.flags&8);
        if(key=="ignorenulls") return VbaValue::logical(false);
        throw VbaRuntimeError(3251,"DAO: Index member unsupported");
    }
    void assign(std::string_view name,const VbaValue& raw) override {
        state_->mutableCheck();const auto key=lower(name);const auto value=scalarVbaValue(raw);if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null index property");
        if(key=="name") {state_->data.name=text(value.asText());return;}
        const auto enabled=coerceVbaModuleValue(value,"boolean").truthy();
        if(key=="primary") {state_->data.kind=enabled ? 1 : 0;if(enabled) state_->data.flags|=1;return;}
        if(key=="unique") {if(!enabled && state_->data.kind==1) throw VbaRuntimeError(3251,"DAO: Primary index must be unique");state_->data.flags=enabled ? state_->data.flags|1 : state_->data.flags&~1;return;}
        if(key=="required") {state_->data.flags=enabled ? state_->data.flags|8 : state_->data.flags&~8;return;}
        if(key=="ignorenulls" && !enabled) return;
        throw VbaRuntimeError(3251,"DAO: Index option unsupported");
    }
    std::shared_ptr<DetachedIndexState> state_;std::shared_ptr<DetachedIndexFields> fields_;
};
class DetachedIndexes final : public VbaObject {
public:
    explicit DetachedIndexes(std::shared_ptr<DetachedTableState> owner):owner_(std::move(owner)) {}
    std::string_view className() const override {return "Indexes";}
    bool isNativeType(std::string_view type) const override {return type=="dao.indexes";}
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {const auto key=lower(name);if(key=="append") return VbaArgumentSignature{{"Object"},1};if(key=="item") return VbaArgumentSignature{{"Index"},1};return {};}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        owner_->workspace->check();const auto key=lower(name);
        if(owner_->attached) return owner_->attached->member("Indexes",{}).object->member(name,args);
        if(key=="count" && args.empty()) return integerCount(indexes_.size());
        if(key=="refresh" && args.empty()) return {};
        if(key=="append") {
            if(args.size()!=1 || args[0].kind!=VbaValue::Kind::Object) throw VbaRuntimeError(450,"DAO: Index required");
            auto index=std::dynamic_pointer_cast<DetachedIndex>(args[0].object);
            if(!index || index->state_->table!=owner_ || index->state_->appended) throw VbaRuntimeError(3219,"DAO: Foreign/attached index");
            for(const auto& existing:indexes_) if(existing->state_->data.name.compare(index->state_->data.name,Qt::CaseInsensitive)==0) throw VbaRuntimeError(3191,"DAO: Duplicate index name");
            index->state_->appended=true;indexes_.append(index);return {};
        }
        if(key!="item" || args.size()!=1) throw VbaRuntimeError(450,"DAO: Index name/index required");
        const auto value=scalarVbaValue(args[0]);if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null index");
        if(value.kind==VbaValue::Kind::Text) {for(const auto& index:indexes_) if(index->state_->data.name.compare(text(value.asText()),Qt::CaseInsensitive)==0) return VbaValue::reference(index);}
        else {const auto index=coerceVbaModuleValue(value,"long").number;if(index>=0 && index<indexes_.size()) return VbaValue::reference(indexes_[static_cast<int>(index)]);}
        throw VbaRuntimeError(3265,"DAO: Index not found");
    }
    VbaValue item(std::string_view name) override {return member("Item",{VbaValue::string(std::string(name))});}
    std::vector<VbaValue> enumerate() override {owner_->workspace->check();if(owner_->attached) return owner_->attached->member("Indexes",{}).object->enumerate();std::vector<VbaValue> result;for(const auto& index:indexes_) result.push_back(VbaValue::reference(index));return result;}
    std::shared_ptr<DetachedTableState> owner_;QVector<std::shared_ptr<DetachedIndex>> indexes_;
};
class DetachedTableDefinition final : public VbaObject {
public:
    DetachedTableDefinition(QSqlDatabase db,QString name):state_(std::make_shared<DetachedTableState>()) {state_->db=db;state_->workspace=workspaceState(db);state_->name=std::move(name);fields_=std::make_shared<DetachedFields>(state_);indexes_=std::make_shared<DetachedIndexes>(state_);}
    std::string_view className() const override {return "TableDef";}
    bool isNativeType(std::string_view type) const override {return type=="dao.tabledef";}
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {const auto key=lower(name);if(key=="createfield") return VbaArgumentSignature{{"Name","Type","Size"},0};if(key=="createindex") return VbaArgumentSignature{{"Name"},0};if(key=="fields" || key=="indexes") return VbaArgumentSignature{{"Index"},0};return {};}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        state_->workspace->check();const auto key=lower(name);
        if(state_->attached) return state_->attached->member(name,args);
        if(args.empty() && (key=="attributes" || key=="connect" || key=="sourcetablename")) return key=="attributes" ? VbaValue::longInteger(0) : VbaValue::string("");
        if(key=="name" && args.empty()) return VbaValue::string(state_->name.toStdString());
        if(key=="fields") return args.empty() ? VbaValue::reference(fields_) : fields_->member("Item",args);
        if(key=="indexes") return args.empty() ? VbaValue::reference(indexes_) : indexes_->member("Item",args);
        if(key=="createindex") {
            if(args.size()>1) throw VbaRuntimeError(450,"DAO: CreateIndex takes Name only");
            if(!args.empty() && args[0].kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null index name");
            return VbaValue::reference(std::make_shared<DetachedIndex>(state_,args.empty() ? QString{} : text(args[0].asText())));
        }
        if(key=="createfield") {
            if(args.size()>3) throw VbaRuntimeError(450,"DAO: CreateField accepts Name, Type and Size");
            QString fieldName;int type=0,size=0;
            if(!args.empty() && args[0].kind!=VbaValue::Kind::Missing) {if(args[0].kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null field name");fieldName=text(args[0].asText());}
            if(args.size()>1 && args[1].kind!=VbaValue::Kind::Missing) type=static_cast<int>(coerceVbaModuleValue(scalarVbaValue(args[1]),"integer").number);
            if(args.size()>2 && args[2].kind!=VbaValue::Kind::Missing) size=static_cast<int>(coerceVbaModuleValue(scalarVbaValue(args[2]),"long").number);
            return VbaValue::reference(std::make_shared<DetachedField>(state_,fieldName,type,size));
        }
        throw VbaRuntimeError(3251,"DAO: Detached TableDef member unsupported");
    }
    void assign(std::string_view name,const VbaValue& raw) override {
        state_->workspace->check();if(state_->attached) {state_->attached->assign(name,raw);return;}const auto key=lower(name);if(key!="name" && key!="attributes" && key!="connect" && key!="sourcetablename") throw VbaRuntimeError(383,"DAO: TableDef property is read-only");
        const auto value=scalarVbaValue(raw);state_->workspace->check();if(state_->attached) {state_->attached->assign(name,value);return;}if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null TableDef property");if(key=="name") {state_->name=text(value.asText());return;}
        if(key=="attributes" ? coerceVbaModuleValue(value,"long").number!=0 : !value.asText().empty()) throw VbaRuntimeError(3251,"DAO: Linked/system TableDef creation is not implemented");
    }
    void append(QSqlDatabase db) {
        state_->workspace->check();if(state_->attached || workspaceState(db)!=state_->workspace) throw VbaRuntimeError(3219,"DAO: TableDef belongs to a different or attached database");
        const auto name=state_->name;
        if(name.isEmpty() || name.trimmed()!=name || name.size()>64 || name.contains(QChar(0)) || !name.isValidUtf16() || name.startsWith("_la_",Qt::CaseInsensitive) || name.startsWith("sqlite_",Qt::CaseInsensitive)) throw VbaRuntimeError(5,"DAO: Invalid/reserved table name");
        if(fields_->fields_.isEmpty()) throw VbaRuntimeError(3264,"DAO: Cannot append a table without fields");
        QVector<AccessSourceIndex> indexes;QSet<QString> indexNames,requiredColumns;int primaryCount=0;
        for(const auto& index:indexes_->indexes_) {
            auto data=index->state_->data;
            if(data.name.isEmpty() || data.name.trimmed()!=data.name || data.name.size()>64 || data.name.contains(QChar(0)) || !data.name.isValidUtf16()) throw VbaRuntimeError(5,"DAO: Invalid index name");
            if(indexNames.contains(data.name.toCaseFolded())) throw VbaRuntimeError(3191,"DAO: Duplicate index name");indexNames.insert(data.name.toCaseFolded());
            if(index->fields_->fields_.isEmpty()) throw VbaRuntimeError(3264,"DAO: Index requires fields");
            if(data.kind==1 && ++primaryCount>1) throw VbaRuntimeError(3283,"DAO: Multiple primary indexes");
            QSet<QString> seen;
            for(const auto& indexField:index->fields_->fields_) {
                QString canonical;int type=0;
                for(const auto& field:fields_->fields_) if(field->name_.compare(indexField->name_,Qt::CaseInsensitive)==0) {canonical=field->name_;type=field->type_;break;}
                if(canonical.isEmpty()) throw VbaRuntimeError(3265,"DAO: Index field absent from table");
                if(type==11 || type==12) throw VbaRuntimeError(3251,"DAO: Binary/Memo indexes unsupported");
                if(seen.contains(canonical.toCaseFolded())) throw VbaRuntimeError(3191,"DAO: Duplicate index field");seen.insert(canonical.toCaseFolded());
                data.columns.append(canonical);data.descending.append(indexField->descending_);
                if(data.kind==1 || (data.flags&8)) requiredColumns.insert(canonical.toCaseFolded());
            }
            indexes.append(data);
        }
        QStringList definitions;QSet<QString> names;
        for(const auto& field:fields_->fields_) {
            if(names.contains(field->name_.toCaseFolded())) throw VbaRuntimeError(3191,"DAO: Duplicate field name");
            names.insert(field->name_.toCaseFolded());auto definition=field->definition();
            if(requiredColumns.contains(field->name_.toCaseFolded()) && !field->required_ && !(field->attributes_&16)) definition+=" NOT NULL";
            definitions.append(definition);
        }
        for(const auto& index:indexes) if(index.kind==1) {
            QStringList columns;for(int i=0;i<index.columns.size();++i) columns.append(schemaJetName(index.columns[i])+(index.descending[i] ? " DESC" : " ASC"));
            definitions.append("PRIMARY KEY ("+columns.join(',')+")");
        }
        QSqlQuery query(db);const auto savepoint="la_schema_builder_"+QUuid::createUuid().toString(QUuid::Id128);
        if(!query.exec("SAVEPOINT "+savepoint)) queryFailure(query);
        bool completed=false;const auto rollback=qScopeGuard([&] {if(!completed) {query.exec("ROLLBACK TO "+savepoint);query.exec("RELEASE "+savepoint);}});
        executeAction(db,compileAccessQuery("CREATE TABLE "+schemaJetName(name)+" ("+definitions.join(',')+")"),{});
        QJsonArray indexMetadata,indexColumns;for(const auto& field:fields_->fields_) indexColumns.append(field->name_);
        for(int position=0;position<indexes.size();++position) {
            const auto& index=indexes[position];QStringList columns;QJsonArray entries;
            for(int i=0;i<index.columns.size();++i) {columns.append(quote(index.columns[i])+(index.descending[i] ? " DESC" : " ASC"));entries.append(QJsonObject{{"name",index.columns[i]},{"descending",index.descending[i]}});}
            if(index.kind!=1 && !query.exec(QString("CREATE ")+((index.flags&1) ? "UNIQUE " : "")+"INDEX "+quote(name+"_"+index.name)+" ON "+quote(name)+" ("+columns.join(',')+")")) queryFailure(query);
            indexMetadata.append(QJsonObject{{"name",index.name},{"kind",index.kind},{"flags",index.flags},{"fields",entries}});
        }
        const auto indexDefinition=QJsonDocument(QJsonObject{{"version",1},{"source_table",name},{"columns",indexColumns},{"indexes",indexMetadata}}).toJson(QJsonDocument::Compact);
        if(!query.exec("CREATE TABLE IF NOT EXISTS main._la_access_indexes(table_name TEXT PRIMARY KEY,definition TEXT NOT NULL)")) queryFailure(query);
        query.prepare("INSERT INTO main._la_access_indexes(table_name,definition) VALUES(?,?)");query.addBindValue(name);query.addBindValue(QString::fromUtf8(indexDefinition));if(!query.exec()) queryFailure(query);
        QJsonArray metadata;
        for(const auto& field:fields_->fields_) metadata.append(field->metadata());
        if(!query.exec("CREATE TABLE IF NOT EXISTS main._la_table_schema(table_name TEXT PRIMARY KEY,definition TEXT NOT NULL)")) queryFailure(query);
        query.prepare("INSERT INTO main._la_table_schema(table_name,definition) VALUES(?,?)");query.addBindValue(name);query.addBindValue(QJsonDocument(metadata).toJson(QJsonDocument::Compact));
        if(!query.exec()) queryFailure(query);
        (void)readNativeIndexCatalog(db,name);
        const auto attached=std::make_shared<TableDefinition>(db,name);
        QVector<std::shared_ptr<VbaObject>> fieldBindings,indexBindings;
        for(const auto& field:fields_->fields_) fieldBindings.append(attached->member("Fields",{VbaValue::string(field->name_.toStdString())}).object);
        for(const auto& index:indexes) indexBindings.append(attached->member("Indexes",{VbaValue::string(index.name.toStdString())}).object);
        if(!query.exec("RELEASE "+savepoint)) queryFailure(query);completed=true;
        state_->attached=attached;
        for(int i=0;i<fields_->fields_.size();++i) fields_->fields_[i]->attachedField_=fieldBindings[i];
        for(int i=0;i<indexes.size();++i) {indexes_->indexes_[i]->state_->data=indexes[i];indexes_->indexes_[i]->state_->attachedIndex=indexBindings[i];}
    }
private:std::shared_ptr<DetachedTableState> state_;std::shared_ptr<DetachedFields> fields_;std::shared_ptr<DetachedIndexes> indexes_;
};
std::shared_ptr<VbaObject> createExistingTableField(QSqlDatabase db,const QString& table,const std::vector<VbaValue>& args) {
    if(args.size()>3) throw VbaRuntimeError(450,"DAO: CreateField accepts Name, Type and Size");
    auto owner=std::make_shared<DetachedTableState>();owner->db=db;owner->workspace=workspaceState(db);owner->name=table;owner->existing=std::make_shared<TableDefinition>(db,table);
    QString name;int type=0,size=0;
    if(!args.empty() && args[0].kind!=VbaValue::Kind::Missing) {const auto value=scalarVbaValue(args[0]);if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null field name");name=text(value.asText());}
    if(args.size()>1 && args[1].kind!=VbaValue::Kind::Missing) type=static_cast<int>(coerceVbaModuleValue(scalarVbaValue(args[1]),"integer").number);
    if(args.size()>2 && args[2].kind!=VbaValue::Kind::Missing) size=static_cast<int>(coerceVbaModuleValue(scalarVbaValue(args[2]),"long").number);
    return std::make_shared<DetachedField>(owner,name,type,size);
}
void mutateExistingTableField(QSqlDatabase db,const QString& table,const VbaValue& value,bool erase) {
    checkReservedCounterSchemaAction(db);
    auto& fieldRevision=workspaceState(db)->tableFieldRevisions[table.toCaseFolded()];
    if(fieldRevision==std::numeric_limits<quint64>::max()) throw VbaRuntimeError(6,"DAO: Field schema revision overflow");
    QSqlQuery query(db);const auto savepoint="la_field_mutation_"+QUuid::createUuid().toString(QUuid::Id128);
    if(!query.exec("SAVEPOINT "+savepoint)) queryFailure(query);
    bool completed=false;const auto rollback=qScopeGuard([&] {if(!completed) {query.exec("ROLLBACK TO "+savepoint);query.exec("RELEASE "+savepoint);}});
    const auto columns=readTableSchemaColumns(db,table);
    QStringList names;for(const auto& column:columns) names.append(column.name);
    QJsonArray metadata;
    if(db.tables().contains("_la_table_schema")) {
        query.prepare("SELECT definition FROM main._la_table_schema WHERE table_name=? COLLATE NOCASE");query.addBindValue(table);
        if(!query.exec()) queryFailure(query);
        if(query.next()) {const auto document=QJsonDocument::fromJson(query.value(0).toByteArray());if(!document.isArray() || query.next()) throw VbaRuntimeError(3251,"DAO: Field metadata unavailable or ambiguous");metadata=document.array();}query.finish();
    }
    std::optional<QJsonObject> indexMetadata;
    if(db.tables().contains("_la_access_indexes")) {
        query.prepare("SELECT definition FROM main._la_access_indexes WHERE table_name=? COLLATE NOCASE");query.addBindValue(table);
        if(!query.exec()) queryFailure(query);
        if(query.next()) {const auto encoded=query.value(0).toString().toUtf8();if(query.next()) throw VbaRuntimeError(3251,"DAO: Index metadata ambiguous");indexMetadata=QJsonDocument::fromJson(encoded).object();}query.finish();
        if(indexMetadata) (void)readNativeIndexCatalog(db,table);
    }
    QString fieldName;std::shared_ptr<DetachedField> builder;
    if(erase) {
        const auto scalar=scalarVbaValue(value);if(scalar.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null field name");
        const auto requested=text(scalar.asText());for(const auto& name:names) if(name.compare(requested,Qt::CaseInsensitive)==0) {fieldName=name;break;}
        if(fieldName.isEmpty()) throw VbaRuntimeError(3265,"DAO: Field not found");
        if(names.size()<=1) throw VbaRuntimeError(3251,"DAO: Cannot remove the last field");
        if(generatedAccessColumns(db,table).contains(fieldName.toCaseFolded())) throw VbaRuntimeError(3251,"DAO: Counter deletion requires dedicated schema migration");
        if(!query.exec("PRAGMA main.table_info("+quote(table)+")")) queryFailure(query);
        while(query.next()) if(query.value(1).toString().compare(fieldName,Qt::CaseInsensitive)==0 && query.value(5).toInt()>0) throw VbaRuntimeError(3251,"DAO: Primary field deletion requires table reconstruction");query.finish();
        if(!query.exec("PRAGMA main.index_list("+quote(table)+")")) queryFailure(query);
        while(query.next()) {
            if(query.value(4).toBool()) throw VbaRuntimeError(3251,"DAO: Partial index dependencies are not proven");
            QSqlQuery fields(db);if(!fields.exec("PRAGMA main.index_xinfo("+quote(query.value(1).toString())+")")) queryFailure(fields);
            while(fields.next()) if(fields.value(5).toBool() && (fields.value(1).toInt()<0 || fields.value(2).toString().compare(fieldName,Qt::CaseInsensitive)==0)) throw VbaRuntimeError(3251,"DAO: Remove dependent indexes before deleting this field");
        }query.finish();
        for(const auto& child:nativeTableNames(db)) {
            QSqlQuery keys(db);if(!keys.exec("PRAGMA main.foreign_key_list("+quote(child)+")")) queryFailure(keys);
            while(keys.next()) if((child.compare(table,Qt::CaseInsensitive)==0 && keys.value(3).toString().compare(fieldName,Qt::CaseInsensitive)==0)
                || (keys.value(2).toString().compare(table,Qt::CaseInsensitive)==0 && keys.value(4).toString().compare(fieldName,Qt::CaseInsensitive)==0)) throw VbaRuntimeError(3251,"DAO: Field participates in a relation");
        }
        if(db.tables().contains("_la_access_relations")) {
            QSqlQuery relations(db);relations.prepare("SELECT 1 FROM main._la_access_relations WHERE (child_table=? COLLATE NOCASE AND child_column=? COLLATE NOCASE) OR (parent_table=? COLLATE NOCASE AND parent_column=? COLLATE NOCASE)");relations.addBindValue(table);relations.addBindValue(fieldName);relations.addBindValue(table);relations.addBindValue(fieldName);
            if(!relations.exec()) queryFailure(relations);if(relations.next()) throw VbaRuntimeError(3251,"DAO: Field participates in a declared relation");
        }
        executeAction(db,compileAccessQuery("ALTER TABLE "+schemaJetName(table)+" DROP COLUMN "+schemaJetName(fieldName)),{});
        names.removeAll(fieldName);QJsonArray retained;for(const auto& entry:metadata) if(entry.toObject().value("name").toString().compare(fieldName,Qt::CaseInsensitive)!=0) retained.append(entry);metadata=retained;
    } else {
        if(value.kind!=VbaValue::Kind::Object) throw VbaRuntimeError(450,"DAO: Detached field required");builder=std::dynamic_pointer_cast<DetachedField>(value.object);
        if(!builder || builder->owner_->workspace!=workspaceState(db) || !builder->owner_->existing || builder->owner_->name.compare(table,Qt::CaseInsensitive)!=0 || builder->appended_ || builder->owner_->attached) throw VbaRuntimeError(3219,"DAO: Field belongs to a different or attached table");
        if(!builder->owner_->existing->valid()) throw VbaRuntimeError(3420,"DAO: Field builder table invalidated");
        if(names.size()>=255) throw VbaRuntimeError(3190,"DAO: Too many fields defined");
        if(builder->attributes_&16) throw VbaRuntimeError(3251,"DAO: Adding AutoNumber requires dedicated schema migration");
        fieldName=builder->name_;if(names.contains(fieldName,Qt::CaseInsensitive)) throw VbaRuntimeError(3191,"DAO: Duplicate field name");
        executeAction(db,compileAccessQuery("ALTER TABLE "+schemaJetName(table)+" ADD COLUMN "+builder->definition()),{});
        names.append(fieldName);metadata.append(builder->metadata());
    }
    if(!query.exec("CREATE TABLE IF NOT EXISTS main._la_table_schema(table_name TEXT PRIMARY KEY,definition TEXT NOT NULL)")) queryFailure(query);
    query.prepare("INSERT INTO main._la_table_schema(table_name,definition) VALUES(?,?) ON CONFLICT(table_name) DO UPDATE SET definition=excluded.definition");query.addBindValue(table);query.addBindValue(QJsonDocument(metadata).toJson(QJsonDocument::Compact));if(!query.exec()) queryFailure(query);
    if(indexMetadata) {
        QJsonArray declared;for(const auto& name:names) declared.append(name);indexMetadata->insert("columns",declared);
        query.prepare("UPDATE main._la_access_indexes SET definition=? WHERE table_name=? COLLATE NOCASE");query.addBindValue(QString::fromUtf8(QJsonDocument(*indexMetadata).toJson(QJsonDocument::Compact)));query.addBindValue(table);
        if(!query.exec() || query.numRowsAffected()!=1) throw VbaRuntimeError(3251,"DAO: Index metadata changed during field mutation");
        (void)readNativeIndexCatalog(db,table);
    }
    (void)readTableSchemaColumns(db,table);
    std::shared_ptr<TableDefinition> attached;std::shared_ptr<VbaObject> fieldBinding;
    if(builder) {attached=std::make_shared<TableDefinition>(db,table);fieldBinding=attached->member("Fields",{VbaValue::string(fieldName.toStdString())}).object;}
    if(!query.exec("RELEASE "+savepoint)) queryFailure(query);completed=true;++fieldRevision;
    if(builder) {builder->appended_=true;builder->owner_->attached=std::move(attached);builder->attachedField_=std::move(fieldBinding);}
}

std::shared_ptr<VbaObject> createExistingTableIndex(QSqlDatabase db,const QString& table,const QString& name) {
    (void)readNativeIndexCatalog(db,table);
    auto owner=std::make_shared<DetachedTableState>();owner->db=db;owner->workspace=workspaceState(db);owner->name=table;owner->existing=std::make_shared<TableDefinition>(db,table);
    return std::make_shared<DetachedIndex>(owner,name);
}
bool mutateNativeSchemaIndex(QSqlDatabase db,const QString& table,const VbaValue& value,bool erase);
void mutateExistingTableIndex(QSqlDatabase db,const QString& table,const VbaValue& raw,bool erase) {
    const auto value=erase ? scalarVbaValue(raw) : raw;
    if(mutateNativeSchemaIndex(db,table,value,erase)) return;
    checkReservedCounterSchemaAction(db);
    QSqlQuery query(db);const auto savepoint="la_index_mutation_"+QUuid::createUuid().toString(QUuid::Id128);
    if(!query.exec("SAVEPOINT "+savepoint)) queryFailure(query);
    bool completed=false;const auto rollback=qScopeGuard([&] {if(!completed) {query.exec("ROLLBACK TO "+savepoint);query.exec("RELEASE "+savepoint);}});
    const auto current=readNativeIndexCatalog(db,table);
    QVector<AccessSourceIndex> definitions;for(const auto& row:current) definitions.append(row.source);
    const auto columns=readTableSchemaColumns(db,table);QStringList columnNames;for(const auto& column:columns) columnNames.append(column.name);
    std::shared_ptr<DetachedIndex> builder;
    if(erase) {
        const auto scalar=scalarVbaValue(value);if(scalar.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null index name");
        const auto requested=text(scalar.asText());int position=-1;
        for(int i=0;i<definitions.size();++i) if(definitions[i].name.compare(requested,Qt::CaseInsensitive)==0) {position=i;break;}
        if(position<0) throw VbaRuntimeError(3265,"DAO: Index not found");
        if(definitions[position].kind==1) throw VbaRuntimeError(3251,"DAO: Primary index deletion requires table reconstruction");
        if((definitions[position].flags&1) && db.tables().contains("_la_access_relations")) {
            QSqlQuery relations(db);relations.prepare("SELECT 1 FROM main._la_access_relations WHERE (flags & 1)!=0 AND (child_table=? COLLATE NOCASE OR parent_table=? COLLATE NOCASE)");relations.addBindValue(table);relations.addBindValue(table);
            if(!relations.exec()) queryFailure(relations);if(relations.next()) throw VbaRuntimeError(3251,"DAO: Remove one-to-one relations before deleting unique indexes");
        }
        if(definitions[position].flags&1) for(const auto& child:nativeTableNames(db)) {
            QSqlQuery references(db);if(!references.exec("PRAGMA main.foreign_key_list("+quote(child)+")")) queryFailure(references);
            while(references.next()) if(references.value(2).toString().compare(table,Qt::CaseInsensitive)==0)
                throw VbaRuntimeError(3251,"DAO: Unique index deletion on a referenced table requires dependency analysis");
        }
        if(!query.exec("DROP INDEX main."+quote(table+"_"+definitions[position].name))) queryFailure(query);
        definitions.removeAt(position);
    } else {
        if(value.kind!=VbaValue::Kind::Object) throw VbaRuntimeError(450,"DAO: Index builder required");
        builder=std::dynamic_pointer_cast<DetachedIndex>(value.object);
        if(!builder || builder->state_->table->workspace!=workspaceState(db) || !builder->state_->table->existing
            || builder->state_->table->name.compare(table,Qt::CaseInsensitive)!=0 || builder->state_->appended)
            throw VbaRuntimeError(3219,"DAO: Index builder belongs to a different or attached table");
        builder->state_->mutableCheck();auto definition=builder->state_->data;
        if(definition.name.isEmpty() || definition.name.trimmed()!=definition.name || definition.name.size()>64 || definition.name.contains(QChar(0)) || !definition.name.isValidUtf16()) throw VbaRuntimeError(5,"DAO: Invalid index name");
        for(const auto& existing:definitions) if(existing.name.compare(definition.name,Qt::CaseInsensitive)==0) throw VbaRuntimeError(3191,"DAO: Duplicate index name");
        if(definition.kind==1) throw VbaRuntimeError(3251,"DAO: Adding primary index requires table reconstruction");
        if(builder->fields_->fields_.isEmpty()) throw VbaRuntimeError(3264,"DAO: Index requires fields");
        QMap<QString,bool> required;
        if(!query.exec("PRAGMA main.table_info("+quote(table)+")")) queryFailure(query);
        while(query.next()) required.insert(query.value(1).toString().toCaseFolded(),query.value(3).toBool());query.finish();
        QSet<QString> seen;QStringList expressions;
        for(const auto& field:builder->fields_->fields_) {
            QString canonical;int type=0;
            for(const auto& column:columns) if(column.name.compare(field->name_,Qt::CaseInsensitive)==0) {canonical=column.name;type=column.sourceType ? column.sourceType : provenDaoFieldType(column.type);break;}
            if(canonical.isEmpty()) throw VbaRuntimeError(3265,"DAO: Index field absent from table");
            if(type<=0 || type==11 || type==12 || type==15) throw VbaRuntimeError(3251,"DAO: Index field type unsupported");
            if(seen.contains(canonical.toCaseFolded())) throw VbaRuntimeError(3191,"DAO: Duplicate index field");seen.insert(canonical.toCaseFolded());
            if((definition.flags&8) && !required.value(canonical.toCaseFolded())) throw VbaRuntimeError(3251,"DAO: Required index needs existing NOT NULL columns");
            definition.columns.append(canonical);definition.descending.append(field->descending_);
            expressions.append(quote(canonical)+(field->descending_ ? " DESC" : " ASC"));
        }
        if(!query.exec(QString("CREATE ")+((definition.flags&1) ? "UNIQUE " : "")+"INDEX main."+quote(table+"_"+definition.name)+" ON "+quote(table)+" ("+expressions.join(',')+")")) queryFailure(query);
        definitions.append(definition);
    }
    QJsonArray entries,names;for(const auto& column:columnNames) names.append(column);
    for(const auto& definition:definitions) {
        QJsonArray fields;for(int i=0;i<definition.columns.size();++i) fields.append(QJsonObject{{"name",definition.columns[i]},{"descending",definition.descending[i]}});
        entries.append(QJsonObject{{"name",definition.name},{"kind",definition.kind},{"flags",definition.flags},{"fields",fields}});
    }
    const auto encoded=QJsonDocument(QJsonObject{{"version",1},{"source_table",table},{"columns",names},{"indexes",entries}}).toJson(QJsonDocument::Compact);
    (void)checkedAccessSourceIndexes(encoded,columnNames);
    query.prepare("UPDATE main._la_access_indexes SET definition=? WHERE table_name=? COLLATE NOCASE");query.addBindValue(QString::fromUtf8(encoded));query.addBindValue(table);
    if(!query.exec() || query.numRowsAffected()!=1) throw VbaRuntimeError(3251,"DAO: Index catalog changed during mutation");
    (void)readNativeIndexCatalog(db,table);
    std::shared_ptr<TableDefinition> attached;std::shared_ptr<VbaObject> indexBinding;
    if(builder) {attached=std::make_shared<TableDefinition>(db,table);indexBinding=attached->member("Indexes",{VbaValue::string(definitions.last().name.toStdString())}).object;}
    if(!query.exec("RELEASE "+savepoint)) queryFailure(query);completed=true;
    if(builder) {builder->state_->data=definitions.last();builder->state_->appended=true;builder->state_->table->attached=std::move(attached);builder->state_->attachedIndex=std::move(indexBinding);}
}

class TableDefinitions final : public VbaObject {
public:
    std::string_view className() const override { return "TableDefs"; }
    bool isNativeType(std::string_view type) const override { return type.starts_with("dao.") && type.substr(4)==lower(className()); }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        const auto key=lower(name);if(key=="item") return VbaArgumentSignature{{"Index"},1};if(key=="append") return VbaArgumentSignature{{"Object"},1};if(key=="delete") return VbaArgumentSignature{{"Name"},1};return {};
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
        if(key=="append") {
            if(args.size()!=1 || args[0].kind!=VbaValue::Kind::Object) throw VbaRuntimeError(450,"DAO: TableDefs.Append requires a detached TableDef");
            const auto table=std::dynamic_pointer_cast<DetachedTableDefinition>(args[0].object);if(!table) throw VbaRuntimeError(3219,"DAO: Detached TableDef required");table->append(db_);return {};
        }
        if(key=="delete") {
            if(args.size()!=1 || args[0].kind==VbaValue::Kind::Null) throw VbaRuntimeError(450,"DAO: TableDefs.Delete requires a table name");
            const auto requested=text(args[0].asText());QString canonical;for(const auto& candidate:names) if(candidate.compare(requested,Qt::CaseInsensitive)==0) {canonical=candidate;break;}
            if(canonical.isEmpty()) throw VbaRuntimeError(3265,"DAO: TableDef not found");executeAction(db_,compileAccessQuery("DROP TABLE "+schemaJetName(canonical)),{});return {};
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
bool checkedRelationKey(QSqlDatabase db,const QString& table,const QStringList& fields) {
    for(const auto& index:readNativeIndexCatalog(db,table)) if((index.source.flags&1) && index.source.columns==fields) return true;
    return false;
}
struct NativeRelationCatalog {
    QSqlDatabase db; std::shared_ptr<WorkspaceState> workspace;
    explicit NativeRelationCatalog(QSqlDatabase database) : db(std::move(database)),workspace(workspaceState(db)) {}
    quint64 version() const {
        workspace->check(); QSqlQuery query(db);
        if(!query.exec("PRAGMA main.schema_version") || !query.next()) queryFailure(query);
        return query.value(0).toULongLong();
    }
    QVector<NativeRelationData> read(const QString& independentlyEditedTable={}) const {
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
            if(item.source=="DAO_NATIVE" && (item.attributes&1)) {
                QStringList parentFields,childFields;for(const auto& pair:item.columns) {parentFields.append(pair.first);childFields.append(pair.second);}
                if(!checkedRelationKey(db,item.parent,parentFields) || !checkedRelationKey(db,item.child,childFields)) throw VbaRuntimeError(3251,"DAO: Native one-to-one relation lacks checked unique keys");
            }
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
        for(const auto& key:actual) if(independentlyEditedTable.isEmpty() || key.child.compare(independentlyEditedTable,Qt::CaseInsensitive)==0 || key.parent.compare(independentlyEditedTable,Qt::CaseInsensitive)==0) throw VbaRuntimeError(3251,"DAO: Native foreign keys without proven relation names are unsupported");
        return result;
    }
};
class NativeRelation final : public VbaObject, public std::enable_shared_from_this<NativeRelation> {
public:
    NativeRelation(std::shared_ptr<NativeRelationCatalog> catalog,NativeRelationData data,quint64 version,quint64 rollback)
        : catalog_(std::move(catalog)),data_(std::move(data)),version_(version),rollback_(rollback),generation_(catalog_->workspace->relationGenerations.value(data_.name.toCaseFolded())) {}
    std::string_view className() const override { return "Relation"; }
    bool isNativeType(std::string_view type) const override { return type=="dao.relation"; }
    bool matches(const NativeRelationData& data,quint64 version,quint64 rollback) const { return data_==data && version_==version && rollback_==rollback && generation_==catalog_->workspace->relationGenerations.value(data_.name.toCaseFolded()); }
    void check() const {
        if(catalog_->version()!=version_ || catalog_->workspace->rollbackRevision!=rollback_ || generation_!=catalog_->workspace->relationGenerations.value(data_.name.toCaseFolded())) throw VbaRuntimeError(3420,"DAO: Relation metadata invalidated");
        for(const auto& current:catalog_->read()) if(current.name.compare(data_.name,Qt::CaseInsensitive)==0) {
            if(current==data_) return; break;
        }
        throw VbaRuntimeError(3420,"DAO: Relation definition changed");
    }
    void publishGeneration(quint64 value) {generation_=value;}
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
    std::shared_ptr<NativeRelationCatalog> catalog_; NativeRelationData data_; quint64 version_{},rollback_{},generation_{};
    std::weak_ptr<VbaObject> fields_;
};
struct DetachedRelationState {
    QSqlDatabase db;std::shared_ptr<WorkspaceState> workspace;QString name,parent,child;quint32 attributes{};std::shared_ptr<NativeRelation> attached;
    void check() const {workspace->check();if(attached) (void)attached->member("Name",{});}
    void mutableCheck() const {check();if(attached) throw VbaRuntimeError(383,"DAO: Attached relation is read-only");}
};
class DetachedRelationField final : public VbaObject {
public:
    DetachedRelationField(std::shared_ptr<DetachedRelationState> owner,QString name):owner_(std::move(owner)),name_(std::move(name)) {}
    std::string_view className() const override {return "Field";}
    bool isNativeType(std::string_view type) const override {return type=="dao.field";}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        owner_->check();if(owner_->attached) return owner_->attached->member("Fields",{VbaValue::string(name_.toStdString())}).object->member(name,args);
        if(!args.empty()) throw VbaRuntimeError(450,"DAO: Relation field property takes no arguments");
        const auto key=lower(name);if(key=="name") return VbaValue::string(name_.toStdString());if(key=="foreignname") return VbaValue::string(foreign_.toStdString());
        throw VbaRuntimeError(3251,"DAO: Relation field property unsupported");
    }
    void assign(std::string_view name,const VbaValue& raw) override {
        owner_->mutableCheck();const auto value=scalarVbaValue(raw);if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null relation field property");
        if(lower(name)=="name") {name_=text(value.asText());return;}if(lower(name)=="foreignname") {foreign_=text(value.asText());return;}
        throw VbaRuntimeError(3251,"DAO: Relation field assignment unsupported");
    }
    std::shared_ptr<DetachedRelationState> owner_;QString name_,foreign_;bool appended_{};
};
class DetachedRelationFields final : public VbaObject {
public:
    explicit DetachedRelationFields(std::shared_ptr<DetachedRelationState> owner):owner_(std::move(owner)) {}
    std::string_view className() const override {return "Fields";}
    bool isNativeType(std::string_view type) const override {return type=="dao.fields";}
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {const auto key=lower(name);if(key=="append") return VbaArgumentSignature{{"Object"},1};if(key=="item") return VbaArgumentSignature{{"Index"},1};return {};}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        owner_->check();if(owner_->attached) return owner_->attached->member("Fields",{}).object->member(name,args);
        const auto key=lower(name);if(key=="count" && args.empty()) return integerCount(fields_.size());if(key=="refresh" && args.empty()) return {};
        if(key=="append") {
            if(args.size()!=1 || args[0].kind!=VbaValue::Kind::Object) throw VbaRuntimeError(450,"DAO: Relation field required");
            auto field=std::dynamic_pointer_cast<DetachedRelationField>(args[0].object);
            if(!field || field->owner_!=owner_ || field->appended_) throw VbaRuntimeError(3219,"DAO: Foreign/attached relation field");
            if(fields_.size()>=10) throw VbaRuntimeError(3251,"DAO: Relation field bound exceeded");
            for(const auto& existing:fields_) if(existing->name_.compare(field->name_,Qt::CaseInsensitive)==0 || existing->foreign_.compare(field->foreign_,Qt::CaseInsensitive)==0) throw VbaRuntimeError(3191,"DAO: Duplicate relation field");
            field->appended_=true;fields_.append(field);return {};
        }
        if(key!="item" || args.size()!=1) throw VbaRuntimeError(450,"DAO: Relation field name/index required");
        const auto value=scalarVbaValue(args[0]);if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null relation field index");
        if(value.kind==VbaValue::Kind::Text) {for(const auto& field:fields_) if(field->name_.compare(text(value.asText()),Qt::CaseInsensitive)==0) return VbaValue::reference(field);}
        else {const auto index=coerceVbaModuleValue(value,"long").number;if(index>=0 && index<fields_.size()) return VbaValue::reference(fields_[static_cast<int>(index)]);}
        throw VbaRuntimeError(3265,"DAO: Relation field not found");
    }
    VbaValue item(std::string_view name) override {return member("Item",{VbaValue::string(std::string(name))});}
    std::vector<VbaValue> enumerate() override {owner_->check();if(owner_->attached) return owner_->attached->member("Fields",{}).object->enumerate();std::vector<VbaValue> result;for(const auto& field:fields_) result.push_back(VbaValue::reference(field));return result;}
    std::shared_ptr<DetachedRelationState> owner_;QVector<std::shared_ptr<DetachedRelationField>> fields_;
};
class DetachedRelation final : public VbaObject {
public:
    DetachedRelation(QSqlDatabase db,QString name,QString parent,QString child,quint32 attributes):state_(std::make_shared<DetachedRelationState>()) {state_->db=db;state_->workspace=workspaceState(db);state_->name=std::move(name);state_->parent=std::move(parent);state_->child=std::move(child);state_->attributes=attributes;fields_=std::make_shared<DetachedRelationFields>(state_);}
    std::string_view className() const override {return "Relation";}
    bool isNativeType(std::string_view type) const override {return type=="dao.relation";}
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {const auto key=lower(name);if(key=="createfield") return VbaArgumentSignature{{"Name"},0};if(key=="fields") return VbaArgumentSignature{{"Index"},0};return {};}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        state_->check();if(state_->attached) return state_->attached->member(name,args);const auto key=lower(name);
        if(key=="fields") return args.empty() ? VbaValue::reference(fields_) : fields_->member("Item",args);
        if(key=="createfield") {if(args.size()>1) throw VbaRuntimeError(450,"DAO: Relation.CreateField takes Name only");if(!args.empty() && args[0].kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null relation field name");return VbaValue::reference(std::make_shared<DetachedRelationField>(state_,args.empty() ? QString{} : text(args[0].asText())));}
        if(!args.empty()) throw VbaRuntimeError(450,"DAO: Relation property takes no arguments");
        if(key=="name") return VbaValue::string(state_->name.toStdString());if(key=="table") return VbaValue::string(state_->parent.toStdString());if(key=="foreigntable") return VbaValue::string(state_->child.toStdString());if(key=="attributes") return VbaValue::longInteger(static_cast<std::int32_t>(state_->attributes));
        throw VbaRuntimeError(3251,"DAO: Relation member unsupported");
    }
    void assign(std::string_view name,const VbaValue& raw) override {
        state_->mutableCheck();const auto value=scalarVbaValue(raw);if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null relation property");
        const auto key=lower(name);if(key=="name") {state_->name=text(value.asText());return;}if(key=="table") {state_->parent=text(value.asText());return;}if(key=="foreigntable") {state_->child=text(value.asText());return;}if(key=="attributes") {state_->attributes=static_cast<quint32>(static_cast<qint32>(coerceVbaModuleValue(value,"long").number));return;}
        throw VbaRuntimeError(3251,"DAO: Relation property assignment unsupported");
    }
    std::shared_ptr<DetachedRelationState> state_;std::shared_ptr<DetachedRelationFields> fields_;
};
class NativeSchemaSettings {
public:
    explicit NativeSchemaSettings(QSqlDatabase db):db_(std::move(db)),query_(db_) {}
    ~NativeSchemaSettings() {if(changed_) (void)restore();}
    void begin(bool requireForeignKeys) {
        const auto workspace=workspaceState(db_);workspace->check();
        if(!sqlite3_get_autocommit(workspace->connection)) throw VbaRuntimeError(3251,"DAO: Native schema reconstruction requires no active transaction");
        foreignKeys_=pragma("foreign_keys");legacy_=pragma("legacy_alter_table");
        if(requireForeignKeys && foreignKeys_!=1) throw VbaRuntimeError(3251,"DAO: Enforced relation migration requires foreign_keys enabled");
        changed_=true;
        if(!query_.exec("PRAGMA foreign_keys=OFF") || !query_.exec("PRAGMA legacy_alter_table=ON") || pragma("foreign_keys")!=0 || pragma("legacy_alter_table")!=1) throw VbaRuntimeError(3251,"DAO: Native schema migration settings unavailable");
    }
    bool restore() noexcept {
        if(!changed_) return true;
        query_.finish();const bool first=query_.exec("PRAGMA legacy_alter_table="+QString::number(legacy_));const bool second=query_.exec("PRAGMA foreign_keys="+QString::number(foreignKeys_));
        bool checked=false;if(first && second && query_.exec("PRAGMA foreign_keys") && query_.next()) checked=query_.value(0).toInt()==foreignKeys_;query_.finish();
        if(checked && query_.exec("PRAGMA legacy_alter_table") && query_.next()) checked=query_.value(0).toInt()==legacy_;else checked=false;query_.finish();
        changed_=false;if(!checked && db_.driver()) db_.driver()->setProperty("la_schema_restore_failed",true);return checked;
    }
private:
    int pragma(const QString& name) {if(!query_.exec("PRAGMA "+name) || !query_.next()) queryFailure(query_);const int result=query_.value(0).toInt();query_.finish();return result;}
    QSqlDatabase db_;QSqlQuery query_;int foreignKeys_{},legacy_{};bool changed_{};
};
struct NativeFieldChange {std::optional<bool> required,allowZeroLength;std::optional<int> size,type;std::optional<QString> defaultValue,validationRule,validationText;};

QPair<qint64,QByteArray> nativeTableValueDigest(QSqlDatabase db,const QString& table,const QStringList& columns,const QSet<QString>& numericColumns={}) {
    QCryptographicHash hash(QCryptographicHash::Sha256);QSqlQuery query(db);QStringList projection{"rowid"};for(const auto& column:columns) projection.append(quote(column));
    if(!query.exec("SELECT "+projection.join(',')+" FROM main."+quote(table)+" ORDER BY rowid")) queryFailure(query);
    qint64 count=0;
    const auto add=[&](char kind,const QByteArray& bytes) {hash.addData(QByteArray(1,kind));hash.addData(QByteArray::number(bytes.size()));hash.addData(":");hash.addData(bytes);};
    while(query.next()) {
        if(++count>1000000) throw VbaRuntimeError(3251,"DAO: Relation migration row bound exceeded");
        for(int column=0;column<projection.size();++column) {
            const auto value=query.value(column);const auto type=value.metaType().id();
            if(value.isNull()) add('N',{});
            else if(column>0 && numericColumns.contains(columns[column-1].toCaseFolded())) {if(type!=QMetaType::Double && type!=QMetaType::Float && type!=QMetaType::Int && type!=QMetaType::LongLong && type!=QMetaType::Bool) throw VbaRuntimeError(3251,"DAO: Non-numeric storage in numeric type migration");add('Q',QByteArray::number(std::bit_cast<quint64>(value.toDouble()),16));}
            else if(type==QMetaType::Double || type==QMetaType::Float) add('R',QByteArray::number(std::bit_cast<quint64>(value.toDouble()),16));
            else if(type==QMetaType::QByteArray) add('B',value.toByteArray());
            else if(type==QMetaType::QString) add('T',value.toString().toUtf8());
            else if(type==QMetaType::Int || type==QMetaType::LongLong || type==QMetaType::Bool) add('I',QByteArray::number(value.toLongLong()));
            else throw VbaRuntimeError(3251,"DAO: Relation migration value encoding unproven");
        }
    }
    if(query.lastError().isValid()) queryFailure(query);return {count,hash.result()};
}
void checkNativeForeignKeys(QSqlDatabase db) {
    QSqlQuery query(db);if(!query.exec("PRAGMA main.foreign_key_check")) queryFailure(query);
    if(query.next()) throw VbaRuntimeError(3251,"DAO: Existing rows violate a foreign-key constraint");if(query.lastError().isValid()) queryFailure(query);
}
void rebuildNativeRelationTable(QSqlDatabase db,const QString& table,const QVector<NativeRelationData>& before,const QVector<NativeRelationData>& after,const QMap<QString,NativeFieldChange>& changes={},const std::optional<QVector<AccessSourceIndex>>& replacementIndexes={},const QString& counterDestination={},const QStringList& reorderedColumns={}) {
    SqliteExecutionBudget budget(db,100000000,30000);
    checkNativeForeignKeys(db);
    const auto columns=readTableSchemaColumns(db,table);const auto indexes=readNativeIndexCatalog(db,table);
    QSqlQuery query(db);query.prepare("SELECT sql FROM main.sqlite_schema WHERE type='table' AND name=?");query.addBindValue(table);
    if(!query.exec() || !query.next()) throw VbaRuntimeError(3251,"DAO: Relation table definition unavailable");const auto originalSql=query.value(0).toString();query.finish();
    if(originalSql.size()>1024*1024) throw VbaRuntimeError(3251,"DAO: Relation table definition bound exceeded");
    QJsonArray metadata;query.prepare("SELECT definition FROM main._la_table_schema WHERE table_name=? COLLATE NOCASE");query.addBindValue(table);
    if(!query.exec() || !query.next()) throw VbaRuntimeError(3251,"DAO: Relation migration requires native declared field metadata");
    const auto document=QJsonDocument::fromJson(query.value(0).toByteArray());if(!document.isArray() || query.next()) throw VbaRuntimeError(3251,"DAO: Relation field metadata ambiguous");metadata=document.array();query.finish();
    if(metadata.size()!=columns.size()) throw VbaRuntimeError(3251,"DAO: Relation migration requires complete native metadata");
    QMap<QString,QJsonObject> declarations;for(const auto& value:metadata) {const auto object=value.toObject();if(object.value("source_origin").toString()!="DAO_NATIVE") throw VbaRuntimeError(3251,"DAO: Original/unknown field migration is not yet supported");declarations.insert(object.value("name").toString().toCaseFolded(),object);}
    QVector<AccessSourceIndex> following;for(const auto& index:indexes) following.append(index.source);if(replacementIndexes) following=*replacementIndexes;
    QSet<QString> required,nextRequired;
    for(const auto& index:indexes) if(index.source.kind==1 || (index.source.flags&8)) for(const auto& field:index.source.columns) required.insert(field.toCaseFolded());
    for(const auto& index:following) if(index.kind==1 || (index.flags&8)) for(const auto& field:index.columns) nextRequired.insert(field.toCaseFolded());
    QStringList definitions,modifiedDefinitions,names;QSet<QString> numericColumns;for(auto it=changes.cbegin();it!=changes.cend();++it) if(it.value().type) numericColumns.insert(it.key());
    auto owner=std::make_shared<DetachedTableState>();owner->db=db;owner->workspace=workspaceState(db);owner->name=table;
    for(const auto& column:columns) {
        if(column.name.compare("rowid",Qt::CaseInsensitive)==0 || column.name.compare("_rowid_",Qt::CaseInsensitive)==0 || column.name.compare("oid",Qt::CaseInsensitive)==0) throw VbaRuntimeError(3251,"DAO: Shadowed row identity migration unsupported");
        if(!declarations.contains(column.name.toCaseFolded()) || !column.nativeAttributes || !column.required) throw VbaRuntimeError(3251,"DAO: Native field properties incomplete for migration");
        DetachedField field(owner,column.name,column.sourceType,column.sourceSize);field.attributes_=*column.nativeAttributes;field.required_=*column.required;field.allowZeroLength_=column.allowZeroLength.value_or(true);field.canonicalDate_=column.canonicalDate.value_or(false);field.validationRule_=column.validationRule;field.validationText_=column.validationText;
        if(!(field.attributes_&16) && column.hasDefault) {if(column.sourceType==8 && nativeDateDefaultExpression(column.storedDefault)) field.default_=*nativeDateDefaultExpression(column.storedDefault);else {if(column.literalDefault.isEmpty()) throw VbaRuntimeError(3251,"DAO: Nonliteral default migration unsupported");field.default_=column.literalDefault;}}
        const auto definition=[&](const DetachedField& value,const QSet<QString>& mandatory) {auto result=value.attributes_&16 ? schemaJetName(column.name)+" INT NOT NULL DEFAULT ("+accessCounterDefault(table,column.name)+")" : value.definition(true);if(mandatory.contains(column.name.toCaseFolded()) && !value.required_ && !(value.attributes_&16)) result+=" NOT NULL";return result;};
        definitions.append(definition(field,required));
        if(changes.contains(column.name.toCaseFolded())) {const auto change=changes.value(column.name.toCaseFolded());if(change.type) field.type_=*change.type;if(change.required) field.required_=*change.required;if(change.allowZeroLength) field.allowZeroLength_=*change.allowZeroLength;if(change.size) field.size_=*change.size;if(change.defaultValue) field.default_=*change.defaultValue;if(change.validationRule) field.validationRule_=*change.validationRule;if(change.validationText) field.validationText_=*change.validationText;}
        auto modified=definition(field,nextRequired);if((field.attributes_&16) && !counterDestination.isEmpty()) modified=schemaJetName(column.name)+" INT NOT NULL DEFAULT ("+accessCounterDefault(counterDestination,column.name)+")";modifiedDefinitions.append(modified);names.append(column.name);
    }
    if(!reorderedColumns.isEmpty()) {if(reorderedColumns.size()!=names.size()) throw VbaRuntimeError(3251,"DAO: Incomplete field reorder");QStringList ordered;QSet<QString> seen;for(const auto& name:reorderedColumns) {const int index=names.indexOf(name);if(index<0 || seen.contains(name.toCaseFolded())) throw VbaRuntimeError(3251,"DAO: Invalid field reorder");seen.insert(name.toCaseFolded());ordered.append(modifiedDefinitions[index]);}modifiedDefinitions=ordered;}
    const auto primaryClause=[](const AccessSourceIndex& index) {QStringList fields;for(int i=0;i<index.columns.size();++i) fields.append(schemaJetName(index.columns[i])+(index.descending[i] ? " DESC" : " ASC"));return "PRIMARY KEY ("+fields.join(',')+")";};
    for(const auto& index:indexes) if(index.source.kind==1) definitions.append(primaryClause(index.source));
    for(const auto& index:following) if(index.kind==1) modifiedDefinitions.append(primaryClause(index));
    const auto makeSql=[&](const QString& target,const QVector<NativeRelationData>& relations,const QStringList& columns) {
        auto parts=columns;
        for(const auto& relation:relations) if(relation.child.compare(table,Qt::CaseInsensitive)==0 && !(relation.attributes&2)) {
            if(relation.attributes&~(1u|256u|4096u)) throw VbaRuntimeError(3251,"DAO: Relation attributes require unsupported migration");
            QStringList parentFields,childFields;for(const auto& pair:relation.columns) {parentFields.append(schemaJetName(pair.first));childFields.append(schemaJetName(pair.second));}
            parts.append("FOREIGN KEY ("+childFields.join(',')+") REFERENCES "+schemaJetName(relation.parent)+" ("+parentFields.join(',')+") ON UPDATE "+((relation.attributes&256) ? "CASCADE" : "NO ACTION")+" ON DELETE "+((relation.attributes&4096) ? "CASCADE" : "NO ACTION"));
        }
        return compileAccessQuery("CREATE TABLE "+schemaJetName(target)+" ("+parts.join(',')+")").sql;
    };
    if(!equivalentNativeSchemaSql(originalSql,makeSql(table,before,definitions))) throw VbaRuntimeError(3251,"DAO: Unknown/custom table constraints prevent safe reconstruction");
    QSet<QString> retainedIndexes;for(const auto& index:following) if(index.kind!=1) retainedIndexes.insert(table+"_"+index.name);
    QVector<QString> objects;query.prepare("SELECT name,type,sql FROM main.sqlite_schema WHERE tbl_name=? AND type IN ('index','trigger') AND sql IS NOT NULL ORDER BY type,name");query.addBindValue(table);
    if(!query.exec()) queryFailure(query);qint64 objectBytes=0;
    while(query.next()) {if(replacementIndexes && query.value(1).toString()=="index" && !retainedIndexes.contains(query.value(0).toString())) continue;const auto sql=query.value(2).toString();objectBytes+=sql.size();if(objects.size()>=4096 || objectBytes>4*1024*1024) throw VbaRuntimeError(3251,"DAO: Relation schema object bound exceeded");objects.append(sql);}query.finish();
    const auto original=nativeTableValueDigest(db,table,names,numericColumns);
    const auto temporary="_la_relation_rebuild_"+QUuid::createUuid().toString(QUuid::Id128);
    if(!query.exec(makeSql(temporary,after,modifiedDefinitions))) queryFailure(query);
    QStringList quoted{"rowid"};for(const auto& name:names) quoted.append(quote(name));
    if(!query.exec("INSERT INTO main."+quote(temporary)+" ("+quoted.join(',')+") SELECT "+quoted.join(',')+" FROM main."+quote(table))) queryFailure(query);
    if(nativeTableValueDigest(db,temporary,names,numericColumns)!=original) throw VbaRuntimeError(3251,"DAO: Relation migration changed row identities or values");
    if(!query.exec("DROP TABLE main."+quote(table)) || !query.exec("ALTER TABLE main."+quote(temporary)+" RENAME TO "+quote(table))) queryFailure(query);
    for(const auto& sql:objects) if(!query.exec(sql)) queryFailure(query);
    if(nativeTableValueDigest(db,table,names,numericColumns)!=original) throw VbaRuntimeError(3251,"DAO: Relation migration final value verification failed");
    if(counterDestination.isEmpty() && reorderedColumns.isEmpty()) {if(!replacementIndexes) (void)readNativeIndexCatalog(db,table);if(changes.isEmpty()) (void)readTableSchemaColumns(db,table);}checkNativeForeignKeys(db);
}

quint64 nativeSchemaVersion(QSqlDatabase db) {
    QSqlQuery query(db);if(!query.exec("PRAGMA main.schema_version") || !query.next()) queryFailure(query);return query.value(0).toULongLong();
}
TableSchemaReader currentNativeFieldReader(QSqlDatabase db,const QString& table) {
    const auto workspace=workspaceState(db);const auto version=nativeSchemaVersion(db),rollback=workspace->rollbackRevision,generation=workspace->tableGenerations.value(table.toCaseFolded());
    return [db,table,workspace,version,rollback,generation] {workspace->check();if(workspace->rollbackRevision!=rollback || workspace->tableGenerations.value(table.toCaseFolded())!=generation || nativeSchemaVersion(db)!=version) throw VbaRuntimeError(3420,"DAO: Edited field metadata invalidated");return readTableSchemaColumns(db,table);};
}
bool mutateNativeSchemaIndex(QSqlDatabase db,const QString& table,const VbaValue& value,bool erase) {
    const auto current=readNativeIndexCatalog(db,table);std::shared_ptr<DetachedIndex> builder;int removed=-1;
    if(erase) {
        if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null index name");const auto name=text(value.asText());
        for(int i=0;i<current.size();++i) if(current[i].source.name.compare(name,Qt::CaseInsensitive)==0 && (current[i].source.kind==1 || (current[i].source.flags&8))) removed=i;
        if(removed<0) return false;
    } else {
        if(value.kind!=VbaValue::Kind::Object || !(builder=std::dynamic_pointer_cast<DetachedIndex>(value.object)) || (builder->state_->data.kind!=1 && !(builder->state_->data.flags&8))) return false;
        if(builder->state_->table->workspace!=workspaceState(db) || !builder->state_->table->existing || builder->state_->table->name.compare(table,Qt::CaseInsensitive)!=0 || builder->state_->appended) throw VbaRuntimeError(3219,"DAO: Schema index builder belongs to a different or attached table");builder->state_->mutableCheck();
    }
    const auto workspace=workspaceState(db);checkReservedCounterSchemaAction(db);NativeSchemaSettings settings(db);settings.begin(false);
    QSqlQuery query(db);const auto savepoint="la_primary_mutation_"+QUuid::createUuid().toString(QUuid::Id128);
    if(!query.exec("SAVEPOINT "+savepoint)) queryFailure(query);bool completed=false;const auto rollback=qScopeGuard([&] {if(!completed) {query.exec("ROLLBACK TO "+savepoint);query.exec("RELEASE "+savepoint);}});
    const auto checked=readNativeIndexCatalog(db,table);if(checked!=current) throw VbaRuntimeError(3420,"DAO: Index catalog changed before schema index mutation");
    auto catalog=std::make_shared<NativeRelationCatalog>(db);const auto relations=catalog->read(table);
    if(erase) for(const auto& relation:relations) if(relation.parent==table || relation.child==table) throw VbaRuntimeError(3251,"DAO: Remove dependent relations before removing a constraint-bearing index");
    const auto columns=readTableSchemaColumns(db,table);QStringList names;for(const auto& column:columns) names.append(column.name);
    QVector<AccessSourceIndex> definitions;for(int i=0;i<current.size();++i) if(i!=removed) definitions.append(current[i].source);
    AccessSourceIndex added;
    if(builder) {
        added=builder->state_->data;
        if(added.name.isEmpty() || added.name.trimmed()!=added.name || added.name.size()>64 || added.name.contains(QChar(0)) || !added.name.isValidUtf16()) throw VbaRuntimeError(5,"DAO: Invalid schema index name");
        for(const auto& existing:definitions) {if(existing.name.compare(added.name,Qt::CaseInsensitive)==0) throw VbaRuntimeError(3191,"DAO: Duplicate index name");if(added.kind==1 && existing.kind==1) throw VbaRuntimeError(3283,"DAO: Table already has a primary key");}
        if(builder->fields_->fields_.isEmpty()) throw VbaRuntimeError(3264,"DAO: Schema index requires fields");
        if(builder->fields_->fields_.size()>10) throw VbaRuntimeError(3251,"DAO: Schema index field bound exceeded");QSet<QString> seen;added.columns.clear();added.descending.clear();
        for(const auto& field:builder->fields_->fields_) {
            QString canonical;int type=0;for(const auto& column:columns) if(column.name.compare(field->name_,Qt::CaseInsensitive)==0) {canonical=column.name;type=column.sourceType;break;}
            if(canonical.isEmpty()) throw VbaRuntimeError(3265,"DAO: Schema index field missing");if(type<=0 || type==11 || type==12 || type==15) throw VbaRuntimeError(3251,"DAO: Schema index field type unsupported");
            if(seen.contains(canonical.toCaseFolded())) throw VbaRuntimeError(3191,"DAO: Duplicate schema index field");seen.insert(canonical.toCaseFolded());added.columns.append(canonical);added.descending.append(field->descending_);
        }
        if(added.kind==1) added.flags|=1;definitions.append(added);
    }
    query.prepare("SELECT definition FROM main._la_access_indexes WHERE table_name=? COLLATE NOCASE");query.addBindValue(table);if(!query.exec() || !query.next()) throw VbaRuntimeError(3251,"DAO: Schema index metadata unavailable");const auto previous=query.value(0);if(query.next()) throw VbaRuntimeError(3251,"DAO: Schema index metadata ambiguous");query.finish();
    QJsonArray entries,columnNames;for(const auto& name:names) columnNames.append(name);
    for(const auto& index:definitions) {QJsonArray fields;for(int i=0;i<index.columns.size();++i) fields.append(QJsonObject{{"name",index.columns[i]},{"descending",index.descending[i]}});entries.append(QJsonObject{{"name",index.name},{"kind",index.kind},{"flags",index.flags},{"fields",fields}});}
    const auto encoded=QJsonDocument(QJsonObject{{"version",1},{"source_table",table},{"columns",columnNames},{"indexes",entries}}).toJson(QJsonDocument::Compact);(void)checkedAccessSourceIndexes(encoded,names);
    if(workspace->tableFieldRevisions.value(table.toCaseFolded())==std::numeric_limits<quint64>::max()) throw VbaRuntimeError(6,"DAO: Field schema revision overflow");
    rebuildNativeRelationTable(db,table,relations,relations,{},definitions);
    if(builder && added.kind!=1) {QStringList fields;for(int i=0;i<added.columns.size();++i) fields.append(quote(added.columns[i])+(added.descending[i] ? " DESC" : " ASC"));if(!query.exec(QString("CREATE ")+((added.flags&1) ? "UNIQUE " : "")+"INDEX main."+quote(table+"_"+added.name)+" ON "+quote(table)+" ("+fields.join(',')+")")) queryFailure(query);}
    query.prepare("UPDATE main._la_access_indexes SET definition=? WHERE table_name=? COLLATE NOCASE AND definition=?");query.addBindValue(QString::fromUtf8(encoded));query.addBindValue(table);query.addBindValue(previous);if(!query.exec() || query.numRowsAffected()!=1) throw VbaRuntimeError(3251,"DAO: Schema index metadata changed during write");
    (void)readNativeIndexCatalog(db,table);(void)readTableSchemaColumns(db,table);(void)catalog->read(table);
    std::shared_ptr<TableDefinition> attached;std::shared_ptr<VbaObject> binding;if(builder) {attached=std::make_shared<TableDefinition>(db,table);binding=attached->member("Indexes",{VbaValue::string(added.name.toStdString())}).object;}
    if(!query.exec("RELEASE "+savepoint)) queryFailure(query);completed=true;++workspace->tableFieldRevisions[table.toCaseFolded()];
    if(builder) {builder->state_->data=added;builder->state_->appended=true;builder->state_->table->attached=std::move(attached);builder->state_->attachedIndex=std::move(binding);}
    if(!settings.restore()) throw DaoStoredSchemaChange("DAO: Schema index change stored; connection settings failed to restore, reopen the project");return true;
}

quint64 renameExistingNativeTable(QSqlDatabase db,const QString& table,const QString& renamed) {
    const auto workspace=workspaceState(db);workspace->check();
    if(renamed.isEmpty() || renamed.trimmed()!=renamed || renamed.size()>64 || renamed.contains(QChar(0)) || !renamed.isValidUtf16() || renamed.startsWith("_la_",Qt::CaseInsensitive) || renamed.startsWith("sqlite_",Qt::CaseInsensitive)) throw VbaRuntimeError(5,"DAO: Invalid table name");
    if(renamed==table) return workspace->tableGenerations.value(table.toCaseFolded());if(renamed.compare(table,Qt::CaseInsensitive)==0) throw VbaRuntimeError(3251,"DAO: Case-only table renaming requires separate reconstruction");
    checkReservedCounterSchemaAction(db);NativeSchemaSettings settings(db);settings.begin(false);
    QSqlQuery query(db);const auto savepoint="la_table_rename_"+QUuid::createUuid().toString(QUuid::Id128);if(!query.exec("SAVEPOINT "+savepoint)) queryFailure(query);bool completed=false;
    const auto rollback=qScopeGuard([&] {if(!completed) {query.exec("ROLLBACK TO "+savepoint);query.exec("RELEASE "+savepoint);}});
    query.prepare("SELECT 1 FROM main.sqlite_schema WHERE name=? COLLATE NOCASE");query.addBindValue(renamed);if(!query.exec()) queryFailure(query);if(query.next()) throw VbaRuntimeError(3010,"DAO: Table name already exists");query.finish();
    const auto columns=readTableSchemaColumns(db,table);QStringList names;for(const auto& column:columns) {if(!column.nativeAttributes || !column.required) throw VbaRuntimeError(3251,"DAO: Table rename requires complete native declarations");names.append(column.name);}
    const auto indexes=readNativeIndexCatalog(db,table);auto catalog=std::make_shared<NativeRelationCatalog>(db);const auto relations=catalog->read(table);
    for(const auto& relation:relations) if((relation.parent==table || relation.child==table) && relation.source!="DAO_NATIVE") throw VbaRuntimeError(3251,"DAO: Original relationship table names remain protected");
    if(!query.exec("SELECT sql FROM main.sqlite_schema WHERE type IN ('view','trigger') AND sql IS NOT NULL")) queryFailure(query);
    while(query.next()) if(nativeSchemaUsesIndexHints(query.value(0).toString())) throw VbaRuntimeError(3251,"DAO: Named-index SQL dependencies require explicit reconstruction");query.finish();
    const auto oldKey=table.toCaseFolded(),newKey=renamed.toCaseFolded();auto& oldGeneration=workspace->tableGenerations[oldKey];auto& newGeneration=workspace->tableGenerations[newKey];if(oldGeneration==std::numeric_limits<quint64>::max() || newGeneration==std::numeric_limits<quint64>::max()) throw VbaRuntimeError(6,"DAO: Table generation overflow");
    const auto digest=nativeTableValueDigest(db,table,names);
    query.prepare("SELECT definition FROM main._la_access_indexes WHERE table_name=? COLLATE NOCASE");query.addBindValue(table);if(!query.exec() || !query.next()) throw VbaRuntimeError(3251,"DAO: Rename index catalog unavailable");const auto previousIndexes=query.value(0);if(query.next()) throw VbaRuntimeError(3251,"DAO: Rename index catalog ambiguous");query.finish();
    auto indexMetadata=QJsonDocument::fromJson(previousIndexes.toByteArray()).object();indexMetadata.insert("source_table",renamed);const auto nextIndexes=QJsonDocument(indexMetadata).toJson(QJsonDocument::Compact);(void)checkedAccessSourceIndexes(nextIndexes,names);
    // Explicit column copying consumes no counters; only new defaults use the
    // destination name. Metadata are rebound before any generator is exposed.
    rebuildNativeRelationTable(db,table,relations,relations,{},{},renamed);
    if(!query.exec("PRAGMA legacy_alter_table=OFF") || !query.exec("PRAGMA legacy_alter_table") || !query.next() || query.value(0).toInt()!=0) throw VbaRuntimeError(3251,"DAO: Table reference propagation unavailable");query.finish();
    if(!query.exec("ALTER TABLE main."+quote(table)+" RENAME TO "+quote(renamed))) queryFailure(query);
    for(const auto& item:indexes) if(item.source.kind!=1) {
        const auto& index=item.source;QStringList fields;for(int i=0;i<index.columns.size();++i) fields.append(quote(index.columns[i])+(index.descending[i] ? " DESC" : " ASC"));
        if(!query.exec("DROP INDEX main."+quote(table+"_"+index.name)) || !query.exec(QString("CREATE ")+((index.flags&1) ? "UNIQUE " : "")+"INDEX main."+quote(renamed+"_"+index.name)+" ON "+quote(renamed)+" ("+fields.join(',')+")")) queryFailure(query);
    }
    for(const auto& storage:{QString("_la_table_schema"),QString("_la_counter32_sequences"),QString("_la_counter32_steps"),QString("_la_counter_sequences")}) if(db.tables().contains(storage)) {
        query.prepare("UPDATE main."+quote(storage)+" SET table_name=? WHERE table_name=? COLLATE NOCASE");query.addBindValue(renamed);query.addBindValue(table);if(!query.exec()) queryFailure(query);if(storage=="_la_table_schema" && query.numRowsAffected()!=1) throw VbaRuntimeError(3251,"DAO: Native table declaration changed during rename");
    }
    query.prepare("UPDATE main._la_access_indexes SET table_name=?,definition=? WHERE table_name=? COLLATE NOCASE AND definition=?");query.addBindValue(renamed);query.addBindValue(QString::fromUtf8(nextIndexes));query.addBindValue(table);query.addBindValue(previousIndexes);if(!query.exec() || query.numRowsAffected()!=1) throw VbaRuntimeError(3251,"DAO: Index metadata changed during rename");
    if(db.tables().contains("_la_access_relations")) for(const auto& side:{QString("parent"),QString("child")}) {query.prepare("UPDATE main._la_access_relations SET "+side+"_table=? WHERE "+side+"_table=? COLLATE NOCASE");query.addBindValue(renamed);query.addBindValue(table);if(!query.exec()) queryFailure(query);}
    renameDatasheetPageCatalog(db,"table",table,renamed);
    if(nativeTableValueDigest(db,renamed,names)!=digest) throw VbaRuntimeError(3251,"DAO: Rename changed stored values/row identities");(void)readTableSchemaColumns(db,renamed);(void)readNativeIndexCatalog(db,renamed);(void)catalog->read(renamed);checkNativeForeignKeys(db);
    if(!query.exec("RELEASE "+savepoint)) queryFailure(query);completed=true;++oldGeneration;++newGeneration;
    if(!settings.restore()) throw DaoStoredSchemaChange("DAO: Table rename stored; connection settings failed to restore, reopen the project");return newGeneration;
}

TableSchemaReader renameNativeField(QSqlDatabase db,const QString& table,const TableSchemaColumn& column,const QString& renamed) {
    const auto workspace=workspaceState(db);const auto columns=readTableSchemaColumns(db,table);
    if(renamed.isEmpty() || renamed.trimmed()!=renamed || renamed.size()>64 || renamed.contains(QChar(0)) || !renamed.isValidUtf16()) throw VbaRuntimeError(5,"DAO: Invalid field name");
    for(const auto& item:columns) {
        if(!item.nativeAttributes || !item.required) throw VbaRuntimeError(3251,"DAO: Field rename requires complete native declarations");
        if(item.name!=column.name && item.name.compare(renamed,Qt::CaseInsensitive)==0) throw VbaRuntimeError(3191,"DAO: Field name already exists");
        if(item.name.compare("rowid",Qt::CaseInsensitive)==0 || item.name.compare("_rowid_",Qt::CaseInsensitive)==0 || item.name.compare("oid",Qt::CaseInsensitive)==0) throw VbaRuntimeError(3251,"DAO: Shadowed row identity rename unsupported");
    }
    if(renamed.compare("rowid",Qt::CaseInsensitive)==0 || renamed.compare("_rowid_",Qt::CaseInsensitive)==0 || renamed.compare("oid",Qt::CaseInsensitive)==0) throw VbaRuntimeError(3251,"DAO: Rename cannot shadow row identity");
    if(renamed==column.name) return currentNativeFieldReader(db,table);
    checkReservedCounterSchemaAction(db);NativeSchemaSettings settings(db);settings.begin(false);
    QSqlQuery query(db);const auto savepoint="la_field_rename_"+QUuid::createUuid().toString(QUuid::Id128);
    if(!query.exec("SAVEPOINT "+savepoint)) queryFailure(query);bool completed=false;const auto rollback=qScopeGuard([&] {if(!completed) {query.exec("ROLLBACK TO "+savepoint);query.exec("RELEASE "+savepoint);}});
    auto catalog=std::make_shared<NativeRelationCatalog>(db);const auto relations=catalog->read(table);(void)readNativeIndexCatalog(db,table);checkNativeForeignKeys(db);
    for(const auto& relation:relations) for(const auto& pair:relation.columns) if(((relation.parent==table && pair.first==column.name) || (relation.child==table && pair.second==column.name)) && relation.source!="DAO_NATIVE") throw VbaRuntimeError(3251,"DAO: Original relation field names remain protected");
    QStringList previousNames,nextNames;for(const auto& item:columns) {previousNames.append(item.name);nextNames.append(item.name==column.name ? renamed : item.name);}
    const auto digest=nativeTableValueDigest(db,table,previousNames);
    struct Metadata {QString table;QVariant previous;QByteArray replacement;};QVector<Metadata> metadata;
    for(const auto& name:{QString("_la_table_schema"),QString("_la_access_indexes")}) {
        query.prepare("SELECT definition FROM main."+quote(name)+" WHERE table_name=? COLLATE NOCASE");query.addBindValue(table);if(!query.exec() || !query.next()) throw VbaRuntimeError(3251,"DAO: Rename metadata unavailable");
        const auto previous=query.value(0);auto document=QJsonDocument::fromJson(previous.toByteArray());if(query.next()) throw VbaRuntimeError(3251,"DAO: Rename metadata ambiguous");query.finish();
        if(name=="_la_table_schema") {
            if(!document.isArray()) throw VbaRuntimeError(3251,"DAO: Rename field metadata invalid");auto fields=document.array();bool found=false;
            for(int i=0;i<fields.size();++i) {auto field=fields[i].toObject();if(field.value("source_origin").toString()!="DAO_NATIVE") throw VbaRuntimeError(3251,"DAO: Original field names remain protected");if(field.value("name").toString()==column.name) {field.insert("name",renamed);fields[i]=field;found=true;}}
            if(!found) throw VbaRuntimeError(3251,"DAO: Rename field declaration missing");document=QJsonDocument(fields);
        } else {
            auto root=document.object();auto names=root.value("columns").toArray();for(int i=0;i<names.size();++i) if(names[i].toString()==column.name) names[i]=renamed;root.insert("columns",names);
            auto indexes=root.value("indexes").toArray();for(int i=0;i<indexes.size();++i) {auto index=indexes[i].toObject();auto fields=index.value("fields").toArray();for(int j=0;j<fields.size();++j) {auto field=fields[j].toObject();if(field.value("name").toString()==column.name) {field.insert("name",renamed);fields[j]=field;}}index.insert("fields",fields);indexes[i]=index;}root.insert("indexes",indexes);document=QJsonDocument(root);
            (void)checkedAccessSourceIndexes(document.toJson(QJsonDocument::Compact),nextNames);
        }
        metadata.append({name,previous,document.toJson(QJsonDocument::Compact)});
    }
    if(workspace->tableFieldRevisions.value(table.toCaseFolded())==std::numeric_limits<quint64>::max()) throw VbaRuntimeError(6,"DAO: Field schema revision overflow");
    if(!query.exec("ALTER TABLE main."+quote(table)+" RENAME COLUMN "+quote(column.name)+" TO "+quote(renamed))) queryFailure(query);
    for(const auto& item:metadata) {query.prepare("UPDATE main."+quote(item.table)+" SET definition=? WHERE table_name=? COLLATE NOCASE AND definition=?");query.addBindValue(item.replacement);query.addBindValue(table);query.addBindValue(item.previous);if(!query.exec() || query.numRowsAffected()!=1) throw VbaRuntimeError(3251,"DAO: Rename metadata changed during write");}
    if(db.tables().contains("_la_access_relations")) for(const auto& side:{QString("parent"),QString("child")}) {
        query.prepare("UPDATE main._la_access_relations SET "+side+"_column=? WHERE "+side+"_table=? COLLATE NOCASE AND "+side+"_column=? COLLATE NOCASE");query.addBindValue(renamed);query.addBindValue(table);query.addBindValue(column.name);if(!query.exec()) queryFailure(query);
    }
    if(nativeTableValueDigest(db,table,nextNames)!=digest) throw VbaRuntimeError(3251,"DAO: Rename changed stored values/row identity");
    (void)readTableSchemaColumns(db,table);(void)readNativeIndexCatalog(db,table);(void)catalog->read(table);checkNativeForeignKeys(db);auto reader=currentNativeFieldReader(db,table);(void)reader();
    if(!query.exec("RELEASE "+savepoint)) queryFailure(query);completed=true;++workspace->tableFieldRevisions[table.toCaseFolded()];
    if(!settings.restore()) throw DaoStoredSchemaChange("DAO: Field rename stored; connection settings failed to restore, reopen the project");return reader;
}

TableSchemaReader reorderNativeField(QSqlDatabase db,const QString& table,const QString& field,int destination) {
    const auto workspace=workspaceState(db);const auto columns=readTableSchemaColumns(db,table);QStringList names;int source=-1;
    for(const auto& column:columns) {if(column.name.compare(field,Qt::CaseInsensitive)==0) source=names.size();names.append(column.name);}
    if(source<0) throw VbaRuntimeError(3265,"DAO: Field unavailable");
    if(destination<0 || destination>=names.size()) throw VbaRuntimeError(3251,"DAO: Native OrdinalPosition requires an existing zero-based field position");
    if(source==destination) return currentNativeFieldReader(db,table);
    const auto version=nativeSchemaVersion(db),generation=workspace->tableGenerations.value(table.toCaseFolded());
    auto& revision=workspace->tableFieldRevisions[table.toCaseFolded()];if(revision==std::numeric_limits<quint64>::max()) throw VbaRuntimeError(6,"DAO: Field revision overflow");
    const auto moved=names.takeAt(source);names.insert(destination,moved);
    checkReservedCounterSchemaAction(db);NativeSchemaSettings settings(db);settings.begin(false);
    QSqlQuery query(db);const auto savepoint="la_field_order_"+QUuid::createUuid().toString(QUuid::Id128);
    if(!query.exec("SAVEPOINT "+savepoint)) queryFailure(query);bool completed=false;const auto rollback=qScopeGuard([&] {if(!completed) {query.exec("ROLLBACK TO "+savepoint);query.exec("RELEASE "+savepoint);}});
    if(version!=nativeSchemaVersion(db) || generation!=workspace->tableGenerations.value(table.toCaseFolded())) throw VbaRuntimeError(3420,"DAO: Field schema changed before reorder");
    auto catalog=std::make_shared<NativeRelationCatalog>(db);const auto relations=catalog->read(table);(void)readNativeIndexCatalog(db,table);
    query.prepare("SELECT definition FROM main._la_table_schema WHERE table_name=? COLLATE NOCASE");query.addBindValue(table);if(!query.exec() || !query.next()) throw VbaRuntimeError(3251,"DAO: Native field metadata unavailable");const auto previousFields=query.value(0);if(query.next()) throw VbaRuntimeError(3251,"DAO: Native field metadata ambiguous");query.finish();
    const auto metadata=QJsonDocument::fromJson(previousFields.toByteArray()).array();QJsonArray ordered,columnNames;for(const auto& name:names) {bool found=false;for(const auto& item:metadata) if(item.toObject().value("name").toString()==name) {if(found) throw VbaRuntimeError(3251,"DAO: Duplicate field declaration");ordered.append(item);found=true;}if(!found) throw VbaRuntimeError(3251,"DAO: Missing field declaration");columnNames.append(name);}
    query.prepare("SELECT definition FROM main._la_access_indexes WHERE table_name=? COLLATE NOCASE");query.addBindValue(table);if(!query.exec() || !query.next()) throw VbaRuntimeError(3251,"DAO: Native index metadata unavailable");const auto previousIndexes=query.value(0);if(query.next()) throw VbaRuntimeError(3251,"DAO: Native index metadata ambiguous");query.finish();
    auto indexMetadata=QJsonDocument::fromJson(previousIndexes.toByteArray()).object();indexMetadata.insert("columns",columnNames);const auto encodedIndexes=QJsonDocument(indexMetadata).toJson(QJsonDocument::Compact);(void)checkedAccessSourceIndexes(encodedIndexes,names);
    rebuildNativeRelationTable(db,table,relations,relations,{}, {}, {},names);
    query.prepare("UPDATE main._la_table_schema SET definition=? WHERE table_name=? COLLATE NOCASE AND definition=?");query.addBindValue(QJsonDocument(ordered).toJson(QJsonDocument::Compact));query.addBindValue(table);query.addBindValue(previousFields);if(!query.exec() || query.numRowsAffected()!=1) throw VbaRuntimeError(3251,"DAO: Field metadata changed during reorder");
    query.prepare("UPDATE main._la_access_indexes SET definition=? WHERE table_name=? COLLATE NOCASE AND definition=?");query.addBindValue(QString::fromUtf8(encodedIndexes));query.addBindValue(table);query.addBindValue(previousIndexes);if(!query.exec() || query.numRowsAffected()!=1) throw VbaRuntimeError(3251,"DAO: Index metadata changed during reorder");
    (void)readTableSchemaColumns(db,table);(void)readNativeIndexCatalog(db,table);(void)catalog->read(table);checkNativeForeignKeys(db);auto reader=currentNativeFieldReader(db,table);(void)reader();
    if(!query.exec("RELEASE "+savepoint)) queryFailure(query);completed=true;++revision;
    if(!settings.restore()) throw DaoStoredSchemaChange("DAO: Field order stored; connection settings failed to restore, reopen the project");return reader;
}

TableSchemaReader updateNativeFieldDisplayProperties(QSqlDatabase db,const QString& table,const QString& field,const QMap<QString,VbaValue>& properties) {
    const auto workspace=workspaceState(db);workspace->check();if(!connectionUpdatable(db)) throw VbaRuntimeError(3027,"DAO: Database is read-only");const auto version=nativeSchemaVersion(db),generation=workspace->tableGenerations.value(table.toCaseFolded());QMap<QString,std::optional<QString>> values;
    for(auto it=properties.cbegin();it!=properties.cend();++it) {if(it.key()!="caption" && it.key()!="description" && it.key()!="format") throw VbaRuntimeError(3251,"DAO: Display property change must be separate");if(it.value().kind==VbaValue::Kind::Null) values.insert(it.key(),{});else {if(it.value().kind!=VbaValue::Kind::Text) throw VbaRuntimeError(13,"DAO: Field display property requires text");const auto text=QString::fromStdString(it.value().text);if(text.size()>255 || text.contains(QChar(0)) || !text.isValidUtf16() || (it.key()=="caption" && (text.contains('\n') || text.contains('\r')))) throw VbaRuntimeError(3251,"DAO: Field caption/description exceeds supported text limits");values.insert(it.key(),text);}}
    QSqlQuery query(db);const auto savepoint="la_field_display_"+QUuid::createUuid().toString(QUuid::Id128);if(!query.exec("SAVEPOINT "+savepoint)) queryFailure(query);bool complete=false;const auto rollback=qScopeGuard([&] {if(!complete) {query.exec("ROLLBACK TO "+savepoint);query.exec("RELEASE "+savepoint);}});
    if(version!=nativeSchemaVersion(db) || generation!=workspace->tableGenerations.value(table.toCaseFolded())) throw VbaRuntimeError(3420,"DAO: Field changed before display property update");query.prepare("SELECT definition FROM main._la_table_schema WHERE table_name=? COLLATE NOCASE");query.addBindValue(table);if(!query.exec() || !query.next()) throw VbaRuntimeError(3251,"DAO: Native field metadata unavailable");const auto previous=query.value(0);if(query.next()) throw VbaRuntimeError(3251,"DAO: Ambiguous native metadata");query.finish();auto metadata=QJsonDocument::fromJson(previous.toByteArray()).array();bool found=false;
    for(int index=0;index<metadata.size();++index) {auto object=metadata[index].toObject();if(object.value("name").toString()==field) {if(found || object.value("source_origin").toString()!="DAO_NATIVE") throw VbaRuntimeError(3251,"DAO: Display property requires an own native field");for(auto it=values.cbegin();it!=values.cend();++it) {if(it.value()) {if(!object.contains("native_"+it.key())) object.insert("native_"+it.key()+"_id",QUuid::createUuid().toString(QUuid::WithoutBraces));object.insert("native_"+it.key(),*it.value());}else {object.remove("native_"+it.key());object.remove("native_"+it.key()+"_id");}}metadata[index]=object;found=true;}}
    if(!found) throw VbaRuntimeError(3420,"DAO: Field display metadata missing");query.prepare("UPDATE main._la_table_schema SET definition=? WHERE table_name=? COLLATE NOCASE AND definition=?");query.addBindValue(QJsonDocument(metadata).toJson(QJsonDocument::Compact));query.addBindValue(table);query.addBindValue(previous);if(!query.exec() || query.numRowsAffected()!=1) throw VbaRuntimeError(3251,"DAO: Field metadata changed during display update");auto reader=currentNativeFieldReader(db,table);(void)reader();if(!query.exec("RELEASE "+savepoint)) queryFailure(query);complete=true;return reader;
}

TableSchemaReader updateNativeFieldPropertiesImpl(QSqlDatabase db,const QString& requestedTable,const QString& requestedField,const QMap<QString,VbaValue>& properties) {
    const auto workspace=workspaceState(db);QString table;
    for(const auto& name:nativeTableNames(db)) if(name.compare(requestedTable,Qt::CaseInsensitive)==0) {table=name;break;}
    if(table.isEmpty()) throw VbaRuntimeError(3265,"DAO: Field table unavailable");
    const auto version=nativeSchemaVersion(db),generation=workspace->tableGenerations.value(table.toCaseFolded());
    QMap<QString,VbaValue> evaluated;
    for(auto it=properties.cbegin();it!=properties.cend();++it) {const auto key=it.key().toCaseFolded();if(key!="required" && key!="defaultvalue" && key!="size" && key!="name" && key!="allowzerolength" && key!="validationrule" && key!="validationtext" && key!="ordinalposition" && key!="datatype" && key!="caption" && key!="description" && key!="format") throw VbaRuntimeError(383,"DAO: Field property is read-only");if(evaluated.contains(key)) throw VbaRuntimeError(450,"DAO: Duplicate field property");evaluated.insert(key,scalarVbaValue(it.value()));}
    workspace->check();if(version!=nativeSchemaVersion(db) || generation!=workspace->tableGenerations.value(table.toCaseFolded())) throw VbaRuntimeError(3420,"DAO: Field identity changed while evaluating properties");
    const auto columns=readTableSchemaColumns(db,table);std::optional<TableSchemaColumn> selected;
    for(const auto& column:columns) if(column.name.compare(requestedField,Qt::CaseInsensitive)==0) {selected=column;break;}
    if(!selected) throw VbaRuntimeError(3265,"DAO: Field not found");
    const auto& column=*selected;
    if(!column.nativeAttributes || !column.required || column.sourceType<=0) throw VbaRuntimeError(3251,"DAO: Only explicitly declared native field properties can be changed");
    if(evaluated.contains("caption") || evaluated.contains("description") || evaluated.contains("format")) return updateNativeFieldDisplayProperties(db,table,column.name,evaluated);
    if(evaluated.contains("ordinalposition")) {if(evaluated.size()!=1) throw VbaRuntimeError(3251,"DAO: OrdinalPosition must be a separate field operation");const auto value=coerceVbaModuleValue(evaluated.value("ordinalposition"),"long");return reorderNativeField(db,table,column.name,static_cast<int>(value.asNumber()));}
    if(*column.nativeAttributes&16) throw VbaRuntimeError(3251,"DAO: Counter property alteration requires dedicated migration");
    if(evaluated.contains("name")) {if(evaluated.size()!=1) throw VbaRuntimeError(3251,"DAO: Rename must be a separate field operation");if(evaluated.value("name").kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null field name");return renameNativeField(db,table,column,text(evaluated.value("name").asText()));}
    auto owner=std::make_shared<DetachedTableState>();owner->db=db;owner->workspace=workspace;owner->name=table;
    DetachedField field(owner,column.name,column.sourceType,column.sourceSize);field.attributes_=*column.nativeAttributes;field.required_=*column.required;field.allowZeroLength_=column.allowZeroLength.value_or(true);field.canonicalDate_=column.canonicalDate.value_or(false);field.validationRule_=column.validationRule;field.validationText_=column.validationText;
    const QMap<int,QString> numericTypes{{1,"BOOLEAN"},{2,"INTEGER1"},{3,"INTEGER2"},{4,"INTEGER4"},{6,"SINGLE_REAL"},{7,"REAL"}};
    if(evaluated.contains("datatype")) {if(evaluated.size()!=1 || !numericTypes.contains(column.sourceType)) throw VbaRuntimeError(3251,"DAO: Numeric type conversion must be a separate native numeric field operation");const int type=static_cast<int>(coerceVbaModuleValue(evaluated.value("datatype"),"long").asNumber());if(!numericTypes.contains(type)) throw VbaRuntimeError(3251,"DAO: Numeric target type unsupported");field.type_=type;}
    const auto lossless=[&](const QVariant& value) {if(value.isNull()) return;const auto type=value.metaType().id();if(type!=QMetaType::Int && type!=QMetaType::LongLong && type!=QMetaType::Double && type!=QMetaType::Float && type!=QMetaType::Bool) throw VbaRuntimeError(3251,"DAO: Non-numeric value/default prevents numeric type conversion");const auto converted=sqlFieldValue(fromSql(value),numericTypes.value(field.type_));if(std::bit_cast<quint64>(value.toDouble())!=std::bit_cast<quint64>(converted.toDouble())) throw VbaRuntimeError(3251,"DAO: Numeric type conversion would change a value or default");};
    if(column.hasDefault) {if(column.sourceType==8 && nativeDateDefaultExpression(column.storedDefault)) field.default_=*nativeDateDefaultExpression(column.storedDefault);else {if(column.literalDefault.isEmpty()) throw VbaRuntimeError(3251,"DAO: Nonliteral default alteration unsupported");field.default_=column.literalDefault;}}
    for(auto it=evaluated.cbegin();it!=evaluated.cend();++it) {
        if(it.key()=="size" && field.type_!=10) throw VbaRuntimeError(3251,"DAO: Only Text capacities can be changed");
        if(it.key()!="datatype") field.assign(it.key().toStdString(),it.value());
    }
    if(evaluated.contains("datatype") && !field.default_.isEmpty()) {QSqlQuery value(db);if(!value.exec("SELECT "+*accessLiteralDefaultSql(field.default_)) || !value.next()) queryFailure(value);lossless(value.value(0));}
    (void)field.definition();
    if(evaluated.contains("defaultvalue") && !field.default_.isEmpty() && !(field.type_==8 && nativeDateDefaultSql(field.default_))) {
        QSqlQuery literal(db);if(!literal.exec("SELECT "+*(field.type_==8 && nativeDateLiteralSql(field.default_) ? nativeDateLiteralSql(field.default_) : accessLiteralDefaultSql(field.default_))) || !literal.next()) queryFailure(literal);const auto scalar=literal.value(0);literal.finish();
        if(field.type_==11 && !scalar.isNull()) throw VbaRuntimeError(3251,"DAO: Binary fields support only Null defaults");
        const auto stored=sqlFieldValue(fromSql(scalar),column.type);
        if(stored.isNull()) field.default_="NULL";
        else if(stored.metaType().id()==QMetaType::QString) {auto value=stored.toString();if(field.type_==10 && value.size()>field.size_) throw VbaRuntimeError(3163,"DAO: Default exceeds Text capacity");field.default_="'"+value.replace("'","''")+"'";}
        else if(stored.metaType().id()==QMetaType::Double) field.default_=QString::number(stored.toDouble(),'g',17);
        else field.default_=QString::number(stored.toLongLong());
    }
    if(field.type_==10 && !field.default_.isEmpty()) {QSqlQuery literal(db);if(!literal.exec("SELECT "+*accessLiteralDefaultSql(field.default_)) || !literal.next()) queryFailure(literal);if(!literal.value(0).isNull() && literal.value(0).toString().size()>field.size_) throw VbaRuntimeError(3163,"DAO: Existing default exceeds new Text capacity");}
    const auto oldDefault=column.hasDefault ? (column.sourceType==8 && nativeDateDefaultExpression(column.storedDefault) ? *nativeDateDefaultExpression(column.storedDefault) : column.literalDefault) : QString{};
    if(field.type_==column.sourceType && field.required_==*column.required && field.size_==(field.type_==10 ? column.sourceSize : static_cast<int>(fixedDaoFieldSize(field.type_).number)) && field.default_==oldDefault && field.validationRule_==column.validationRule && field.validationText_==column.validationText && field.allowZeroLength_==column.allowZeroLength.value_or(true) && (!evaluated.contains("allowzerolength") || column.allowZeroLength.has_value())) return currentNativeFieldReader(db,table);
    checkReservedCounterSchemaAction(db);
    NativeSchemaSettings settings(db);settings.begin(false);
    QSqlQuery query(db);const auto savepoint="la_field_properties_"+QUuid::createUuid().toString(QUuid::Id128);
    if(!query.exec("SAVEPOINT "+savepoint)) queryFailure(query);bool completed=false;
    const auto rollback=qScopeGuard([&] {if(!completed) {query.exec("ROLLBACK TO "+savepoint);query.exec("RELEASE "+savepoint);}});
    if(version!=nativeSchemaVersion(db) || generation!=workspace->tableGenerations.value(table.toCaseFolded())) throw VbaRuntimeError(3420,"DAO: Field schema changed before edit");
    auto catalog=std::make_shared<NativeRelationCatalog>(db);const auto relations=catalog->read(table);
    if(field.type_!=column.sourceType || (field.type_==10 && field.size_!=column.sourceSize) || (*column.required && !field.required_)) for(const auto& relation:relations) for(const auto& pair:relation.columns)
        if((relation.parent==table && pair.first==column.name) || (relation.child==table && pair.second==column.name)) throw VbaRuntimeError(3251,"DAO: Remove dependent relations before changing key capacity/relaxing Required");
    if(field.type_!=column.sourceType) {QSqlQuery values(db);if(!values.exec("SELECT "+quote(column.name)+" FROM main."+quote(table))) queryFailure(values);qint64 count=0;while(values.next()) {if(++count>1000000) throw VbaRuntimeError(3251,"DAO: Numeric type conversion row bound exceeded");lossless(values.value(0));}if(values.lastError().isValid()) queryFailure(values);values.finish();}
    query.prepare("SELECT definition FROM main._la_table_schema WHERE table_name=? COLLATE NOCASE");query.addBindValue(table);if(!query.exec() || !query.next()) throw VbaRuntimeError(3251,"DAO: Native field metadata unavailable");
    const auto originalMetadataValue=query.value(0);const auto originalMetadata=originalMetadataValue.toByteArray();if(query.next()) throw VbaRuntimeError(3251,"DAO: Native field metadata ambiguous");query.finish();
    auto metadata=QJsonDocument::fromJson(originalMetadata).array();bool found=false;
    for(int i=0;i<metadata.size();++i) {auto object=metadata[i].toObject();if(object.value("name").toString().compare(column.name,Qt::CaseInsensitive)==0) {if(object.value("source_origin").toString()!="DAO_NATIVE") throw VbaRuntimeError(3251,"DAO: Original field alteration unsupported");object.insert("native_required",field.required_);if(field.type_!=column.sourceType) {object.insert("source_type",field.metadata().value("source_type"));object.insert("source_size",0);}if(!field.validationRule_.isEmpty()) object.insert("native_validation_rule",field.validationRule_);else object.remove("native_validation_rule");if(!field.validationText_.isEmpty()) object.insert("native_validation_text",field.validationText_);else object.remove("native_validation_text");if(field.type_==10 || field.type_==12) object.insert("native_allow_zero_length",field.allowZeroLength_);if(field.type_==10) object.insert("source_size",field.size_);metadata[i]=object;found=true;break;}}
    if(!found) throw VbaRuntimeError(3251,"DAO: Native field declaration missing");
    NativeFieldChange change;if(field.type_!=column.sourceType) change.type=field.type_;change.required=field.required_;change.allowZeroLength=field.allowZeroLength_;if(field.type_==10) change.size=field.size_;change.defaultValue=field.default_;change.validationRule=field.validationRule_;change.validationText=field.validationText_;
    if(workspace->tableFieldRevisions.value(table.toCaseFolded())==std::numeric_limits<quint64>::max()) throw VbaRuntimeError(6,"DAO: Field schema revision overflow");
    rebuildNativeRelationTable(db,table,relations,relations,{{column.name.toCaseFolded(),change}});
    query.prepare("UPDATE main._la_table_schema SET definition=? WHERE table_name=? COLLATE NOCASE AND definition=?");query.addBindValue(QJsonDocument(metadata).toJson(QJsonDocument::Compact));query.addBindValue(table);query.addBindValue(originalMetadataValue);if(!query.exec() || query.numRowsAffected()!=1) throw VbaRuntimeError(3251,"DAO: Field metadata changed during edit");
    (void)readTableSchemaColumns(db,table);(void)readNativeIndexCatalog(db,table);(void)catalog->read(table);auto reader=currentNativeFieldReader(db,table);(void)reader();
    if(!query.exec("RELEASE "+savepoint)) queryFailure(query);completed=true;++workspace->tableFieldRevisions[table.toCaseFolded()];
    if(!settings.restore()) throw DaoStoredSchemaChange("DAO: Field changes stored; connection settings failed to restore, reopen the project");
    return reader;
}

void mutateNativeRelation(std::shared_ptr<NativeRelationCatalog> catalog,const VbaValue& value,bool erase) {
    const auto db=catalog->db;checkReservedCounterSchemaAction(db);
    const auto preliminary=catalog->read();std::shared_ptr<DetachedRelation> candidate;bool rebuild=false;
    if(erase) {
        const auto scalar=scalarVbaValue(value);if(scalar.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null relation name");
        const auto requested=text(scalar.asText());bool found=false;for(const auto& relation:preliminary) if(relation.name.compare(requested,Qt::CaseInsensitive)==0) {rebuild=!(relation.attributes&2);found=true;break;}if(!found) throw VbaRuntimeError(3265,"DAO: Relation not found");
    } else {
        if(value.kind!=VbaValue::Kind::Object) throw VbaRuntimeError(450,"DAO: Detached relation required");candidate=std::dynamic_pointer_cast<DetachedRelation>(value.object);
        if(!candidate || candidate->state_->workspace!=catalog->workspace || candidate->state_->attached) throw VbaRuntimeError(3219,"DAO: Relation belongs to a different or attached database");
        const auto attributes=candidate->state_->attributes;
        if((attributes&2) ? (attributes!=2 && attributes!=3) : (attributes&~(1u|256u|4096u))) throw VbaRuntimeError(3251,"DAO: Relation attributes unsupported");
        rebuild=!(attributes&2);
    }
    NativeSchemaSettings settings(db);if(rebuild) settings.begin(true);
    QSqlQuery query(db);
    const auto savepoint="la_relation_mutation_"+QUuid::createUuid().toString(QUuid::Id128);
    if(!query.exec("SAVEPOINT "+savepoint)) queryFailure(query);
    bool completed=false;const auto rollback=qScopeGuard([&] {if(!completed) {query.exec("ROLLBACK TO "+savepoint);query.exec("RELEASE "+savepoint);}});
    const auto existing=catalog->read();QString name,rebuiltTable;std::shared_ptr<DetachedRelation> builder;
    if(erase) {
        const auto requested=text(scalarVbaValue(value).asText());std::optional<NativeRelationData> selected;for(const auto& relation:existing) if(relation.name.compare(requested,Qt::CaseInsensitive)==0) {selected=relation;break;}
        if(!selected || rebuild==static_cast<bool>(selected->attributes&2)) throw VbaRuntimeError(3251,"DAO: Relation changed before deletion");
        name=selected->name;
        if(rebuild) {auto desired=existing;for(int i=0;i<desired.size();++i) if(desired[i].name==name) {desired.removeAt(i);break;}rebuiltTable=selected->child;rebuildNativeRelationTable(db,rebuiltTable,existing,desired);}
        query.prepare("DELETE FROM main._la_access_relations WHERE name=? COLLATE NOCASE");query.addBindValue(name);if(!query.exec() || query.numRowsAffected()!=selected->columns.size()) throw VbaRuntimeError(3251,"DAO: Relation changed during deletion");
    } else {
        builder=candidate;builder->state_->mutableCheck();name=builder->state_->name;
        if(name.isEmpty() || name.trimmed()!=name || name.size()>64 || name.contains(QChar(0)) || !name.isValidUtf16()) throw VbaRuntimeError(5,"DAO: Invalid relation name");
        for(const auto& relation:existing) if(relation.name.compare(name,Qt::CaseInsensitive)==0) throw VbaRuntimeError(3251,"DAO: Relation name already exists");
        const auto canonicalTable=[&](const QString& requested) {for(const auto& table:nativeTableNames(db)) if(table.compare(requested,Qt::CaseInsensitive)==0) return table;throw VbaRuntimeError(3265,"DAO: Relation table not found");};
        const auto parent=canonicalTable(builder->state_->parent),child=canonicalTable(builder->state_->child);
        const auto parentFields=readTableSchemaColumns(db,parent),childFields=readTableSchemaColumns(db,child);
        const auto findField=[](const QVector<TableSchemaColumn>& columns,const QString& requested) {for(const auto& column:columns) if(column.name.compare(requested,Qt::CaseInsensitive)==0) return column;throw VbaRuntimeError(3265,"DAO: Relation field not found");};
        if(builder->fields_->fields_.isEmpty()) throw VbaRuntimeError(3264,"DAO: Relation requires fields");
        QVector<QPair<QString,QString>> pairs;QSet<QString> primaryNames,foreignNames;QStringList parentNames;
        for(const auto& field:builder->fields_->fields_) {
            const auto p=findField(parentFields,field->name_),c=findField(childFields,field->foreign_);
            const auto pType=p.sourceType ? p.sourceType : provenDaoFieldType(p.type),cType=c.sourceType ? c.sourceType : provenDaoFieldType(c.type);
            if(pType<=0 || pType!=cType || pType==11 || pType==12 || pType==15 || (pType==10 && (p.sourceSize<=0 || p.sourceSize!=c.sourceSize))) throw VbaRuntimeError(3251,"DAO: Relation field types/capacities are not compatible and proven");
            if((rebuild || (builder->state_->attributes&1)) && pType!=1 && pType!=2 && pType!=3 && pType!=4 && pType!=6 && pType!=7 && pType!=8) throw VbaRuntimeError(3251,"DAO: Enforced relation requires matching numeric or canonical native Date fields");
            if(pType==6 && (rebuild || (builder->state_->attributes&1))) {
                const auto nativeSingle=[&](const QString& table,const TableSchemaColumn& column) {if(!column.nativeAttributes || *column.nativeAttributes!=1 || !column.required) return false;QSqlQuery declared(db);if(!declared.exec("PRAGMA main.table_info("+quote(table)+")")) queryFailure(declared);while(declared.next()) if(declared.value(1).toString()==column.name) return declared.value(2).toString().trimmed().toUpper()=="SINGLE_REAL";return false;};
                if(!nativeSingle(parent,p) || !nativeSingle(child,c)) throw VbaRuntimeError(3251,"DAO: Single relationship keys require native quantized Single fields on both sides");
            }
            if(pType==8 && (rebuild || (builder->state_->attributes&1)) && (!canonicalNativeDateColumn(db,parent,p) || !canonicalNativeDateColumn(db,child,c))) throw VbaRuntimeError(3251,"DAO: Date relationship keys require canonical native Date fields on both sides");
            if(primaryNames.contains(p.name.toCaseFolded()) || foreignNames.contains(c.name.toCaseFolded())) throw VbaRuntimeError(3191,"DAO: Duplicate relation field");primaryNames.insert(p.name.toCaseFolded());foreignNames.insert(c.name.toCaseFolded());pairs.append({p.name,c.name});parentNames.append(p.name);
        }
        for(const auto& relation:existing) if(relation.parent==parent && relation.child==child && relation.columns==pairs) throw VbaRuntimeError(3251,"DAO: Relation field mapping already exists");
        if(builder->state_->attributes&1) {QStringList childNames;for(const auto& pair:pairs) childNames.append(pair.second);if(!checkedRelationKey(db,parent,parentNames) || !checkedRelationKey(db,child,childNames)) throw VbaRuntimeError(3251,"DAO: One-to-one relation requires checked unique parent and child keys");}
        if(rebuild) {
            bool unique=false;for(const auto& index:readNativeIndexCatalog(db,parent)) if((index.source.flags&1) && index.source.columns==parentNames) unique=true;
            if(!unique) throw VbaRuntimeError(3251,"DAO: Enforced relation requires a checked unique parent key");
            if(!query.exec("PRAGMA main.table_info("+quote(parent)+")")) queryFailure(query);QSet<QString> required;while(query.next()) if(query.value(3).toBool()) required.insert(query.value(1).toString().toCaseFolded());query.finish();
            for(const auto& field:parentNames) if(!required.contains(field.toCaseFolded())) throw VbaRuntimeError(3251,"DAO: Enforced parent key must be structurally non-null");
            auto desired=existing;desired.append({name,"DAO_NATIVE",parent,child,builder->state_->attributes,pairs});
            std::sort(desired.begin(),desired.end(),[](const auto& a,const auto& b) {return a.name.toCaseFolded()<b.name.toCaseFolded();});
            rebuiltTable=child;rebuildNativeRelationTable(db,child,existing,desired);
        }
        if(!query.exec("CREATE TABLE IF NOT EXISTS main._la_access_relations(source_file TEXT NOT NULL,name TEXT NOT NULL,flags INTEGER NOT NULL,ordinal INTEGER NOT NULL,child_table TEXT NOT NULL,child_column TEXT NOT NULL,parent_table TEXT NOT NULL,parent_column TEXT NOT NULL,PRIMARY KEY(name,ordinal))")) queryFailure(query);
        query.prepare("INSERT INTO main._la_access_relations(source_file,name,flags,ordinal,child_table,child_column,parent_table,parent_column) VALUES(?,?,?,?,?,?,?,?)");
        for(int i=0;i<pairs.size();++i) {query.bindValue(0,QStringLiteral("DAO_NATIVE"));query.bindValue(1,name);query.bindValue(2,static_cast<qlonglong>(builder->state_->attributes));query.bindValue(3,i);query.bindValue(4,child);query.bindValue(5,pairs[i].second);query.bindValue(6,parent);query.bindValue(7,pairs[i].first);if(!query.exec()) queryFailure(query);}
    }
    auto& generation=catalog->workspace->relationGenerations[name.toCaseFolded()];if(generation==std::numeric_limits<quint64>::max()) throw VbaRuntimeError(6,"DAO: Relation generation overflow");
    if(!rebuiltTable.isEmpty() && catalog->workspace->tableGenerations.value(rebuiltTable.toCaseFolded())==std::numeric_limits<quint64>::max()) throw VbaRuntimeError(6,"DAO: Table generation overflow");
    const auto rows=catalog->read();std::shared_ptr<NativeRelation> attached;
    if(builder) {for(const auto& relation:rows) if(relation.name==name) {attached=std::make_shared<NativeRelation>(catalog,relation,catalog->version(),catalog->workspace->rollbackRevision);break;}if(!attached) throw VbaRuntimeError(3251,"DAO: Appended relation missing");(void)attached->member("Name",{});}
    if(!query.exec("RELEASE "+savepoint)) queryFailure(query);completed=true;++generation;
    if(!rebuiltTable.isEmpty()) ++catalog->workspace->tableGenerations[rebuiltTable.toCaseFolded()];
    if(builder) {attached->publishGeneration(generation);builder->state_->attached=std::move(attached);}
    if(!settings.restore()) throw DaoStoredSchemaChange("DAO: Relation change stored; connection settings failed to restore, reopen the project");
}

class NativeRelations final : public VbaObject {
public:
    explicit NativeRelations(QSqlDatabase db) : catalog_(std::make_shared<NativeRelationCatalog>(std::move(db))) {}
    std::string_view className() const override { return "Relations"; }
    bool isNativeType(std::string_view type) const override { return type=="dao.relations"; }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        const auto key=lower(name);if(key=="item") return VbaArgumentSignature{{"Index"},1};if(key=="append") return VbaArgumentSignature{{"Object"},1};if(key=="delete") return VbaArgumentSignature{{"Name"},1};return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        const auto rows=catalog_->read(); const auto key=lower(name);
        if(key=="count" || key=="refresh") {
            if(!args.empty()) throw VbaRuntimeError(450,"DAO: Relation collection property takes no arguments");
            return key=="count" ? integerCount(rows.size()) : VbaValue{};
        }
        if(key=="append" || key=="delete") {if(args.size()!=1) throw VbaRuntimeError(450,"DAO: Relation/name required");mutateNativeRelation(catalog_,args[0],key=="delete");return {};}
        if(key!="item") throw VbaRuntimeError(3251,"DAO: Relation catalog member unsupported");
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
        if (key=="createproperty") return VbaArgumentSignature{{"Name","Type","Value","DDL"},3};
        if (key=="openrecordset") return VbaArgumentSignature{{"Name","Type"},1};
        if (key=="execute") return VbaArgumentSignature{{"Query","Options"},1};
        if (key=="createquerydef") return VbaArgumentSignature{{"Name","SQLText"},1};
        if(key=="createrelation") return VbaArgumentSignature{{"Name","Table","ForeignTable","Attributes"},0};
        if(key=="createtabledef") return VbaArgumentSignature{{"Name","Attributes","SourceTableName","Connect"},0};
        if (key=="tabledefs" || key=="relations") return VbaArgumentSignature{{"Index"},0};
        if (key=="querydefs") return VbaArgumentSignature{{"Index"},0};
        return {};
    }
    explicit Database(QSqlDatabase db) : db_(std::move(db)),workspace_(workspaceState(db_)) {}
    VbaValue member(std::string_view name, const std::vector<VbaValue>& args) override {
        workspace_->check(); require(db_.isOpen(), "DAO: Database is closed");
        const auto key = lower(name);
        if(key=="updatable") {if(!args.empty()) throw VbaRuntimeError(450,"DAO: Database.Updatable takes no arguments");return VbaValue::logical(connectionUpdatable(db_));}
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
        if(key=="createrelation") {
            if(args.size()>4) throw VbaRuntimeError(450,"DAO: CreateRelation argument count");
            QString values[3];quint32 attributes=0;
            for(std::size_t i=0;i<args.size();++i) if(args[i].kind!=VbaValue::Kind::Missing) {const auto value=scalarVbaValue(args[i]);if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null relation argument");if(i<3) values[i]=text(value.asText());else attributes=static_cast<quint32>(static_cast<qint32>(coerceVbaModuleValue(value,"long").number));}
            return VbaValue::reference(std::make_shared<DetachedRelation>(db_,values[0],values[1],values[2],attributes));
        }
        if(key=="createproperty") {if(args.size()<3 || args.size()>4) throw VbaRuntimeError(450,"DAO: CreateProperty requires Name,Type,Value");const auto key=nativeTextPropertyKey(text(scalarVbaValue(args[0]).asText()));if(coerceVbaModuleValue(scalarVbaValue(args[1]),"long").asNumber()!=10 || (args.size()==4 && args[3].kind!=VbaValue::Kind::Missing && coerceVbaModuleValue(scalarVbaValue(args[3]),"boolean").truthy())) throw VbaRuntimeError(3251,"DAO: Only non-DDL native Caption/Description/Format Text properties are implemented");return VbaValue::reference(std::make_shared<DetachedTextProperty>(key,nativeTextPropertyValue(key,args[2])));}
        if(key=="createtabledef") {
            if(args.size()>4) throw VbaRuntimeError(450,"DAO: CreateTableDef argument count");
            QString tableName;
            if(!args.empty() && args[0].kind!=VbaValue::Kind::Missing) {if(args[0].kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null table name");tableName=text(args[0].asText());}
            for(std::size_t i=1;i<args.size();++i) if(args[i].kind!=VbaValue::Kind::Missing) {
                const auto value=scalarVbaValue(args[i]);if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"DAO: Null table definition option");
                if(i==1 ? coerceVbaModuleValue(value,"long").number!=0 : !value.asText().empty()) throw VbaRuntimeError(3251,"DAO: Linked/attribute TableDef creation unsupported");
            }
            return VbaValue::reference(std::make_shared<DetachedTableDefinition>(db_,tableName));
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
            const auto source = text(args[0].asText());
            const bool physical=nativeTableNames(db_).contains(source,Qt::CaseInsensitive);const int mode=nativeRecordsetMode(args,1,physical ? 1 : 2);
            if (db_.tables().contains(source,Qt::CaseInsensitive)) {if(mode==1 && !physical) throw VbaRuntimeError(3251,"DAO: Table-type Recordset requires a physical main user table");return VbaValue::reference(std::make_shared<Recordset>(db_,source,mode == 4,QMap<QString,QVariant>{},AccessQueryPlan{},mode));}
            if(mode==1) throw VbaRuntimeError(3251,"DAO: Table-type Recordset requires a physical table");
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
    void assign(std::string_view name,const VbaValue& value) override {workspace_->check();if(lower(name)=="updatable") throw VbaRuntimeError(383,"DAO: Database.Updatable is read-only");VbaObject::assign(name,value);}
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
bool nativeFieldHasCanonicalDateStorage(QSqlDatabase db,const QString& table,const QString& field) {
    workspaceState(db)->check();for(const auto& column:readTableSchemaColumns(db,table)) if(column.name.compare(field,Qt::CaseInsensitive)==0) return canonicalNativeDateColumn(db,table,column);return false;
}
QMap<QString,NativeFieldDisplayProperties> nativeTableFieldDisplayProperties(QSqlDatabase db,const QString& requested) {
    workspaceState(db)->check();QString table;for(const auto& name:nativeTableNames(db)) if(name.compare(requested,Qt::CaseInsensitive)==0) {table=name;break;}if(table.isEmpty()) return {};QMap<QString,NativeFieldDisplayProperties> result;for(const auto& field:readTableSchemaColumns(db,table)) if(field.nativeAttributes || field.sourceDisplayProven) result.insert(field.name.toCaseFolded(),{field.caption,field.description,field.nativeAttributes.has_value(),field.format});return result;
}
std::optional<QString> nativeDateLiteralSql(const QString& expression) {
    auto token=expression.trimmed();if(token.startsWith('=')) token=token.mid(1).trimmed();
    static const QRegularExpression shape(R"(^#[0-9]{4}-[0-9]{2}-[0-9]{2}(?:[ T][0-9]{2}:[0-9]{2}(?::[0-9]{2})?)?#$)");
    if(!shape.match(token).hasMatch()) return {};
    try {const auto stored=sqlFieldValue(qtFieldValue(token.mid(1,token.size()-2),"DATETIME"),"DATETIME").toString();return "'"+stored+"'";}catch(const VbaRuntimeError&) {throw VbaRuntimeError(3251,"Native Date literal requires a valid ISO whole-second date");}
}
void changeNativeNumericFieldType(QSqlDatabase db,const QString& table,const QString& field,int type) {
    (void)updateNativeFieldPropertiesImpl(db,table,field,{{"datatype",VbaValue::longInteger(type)}});
}
void updateNativeFieldProperties(QSqlDatabase db,const QString& table,const QString& field,const QMap<QString,VbaValue>& properties) {
    (void)updateNativeFieldPropertiesImpl(db,table,field,properties);
}
void trackAccessCounterReservations(QSqlDatabase db,const QVector<AccessCounterReservation>& reservations) {
    const auto state=workspaceState(db);
    if(state->savepoints.isEmpty()) return;
    for(const auto& reservation:reservations) {
        const auto key=reservation.table+QChar(0)+reservation.column;
        auto& retained=state->counterReservations[key];
        if(reservation.jetLong || retained.table.isEmpty() || reservation.value>retained.value) retained=reservation;
    }
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
    return literalDefaultValues(database,readTableSchemaColumns(database,canonical));
}
QMap<QString,QVariant> nativeTableDateDefaults(QSqlDatabase database,const QString& table) {
    workspaceState(database)->check();if(!nativeTableNames(database).contains(table,Qt::CaseInsensitive)) return {};
    QSqlQuery shadow(database);shadow.prepare("SELECT 1 FROM temp.sqlite_schema WHERE type IN ('table','view') AND (name=? COLLATE NOCASE OR name='_la_table_schema' COLLATE NOCASE)");shadow.addBindValue(table);if(!shadow.exec()) queryFailure(shadow);if(shadow.next()) throw VbaRuntimeError(3251,"Native Date defaults have a TEMP shadow");shadow.finish();
    const auto columns=readTableSchemaColumns(database,table);QVector<TableSchemaColumn> dates;for(const auto& column:columns) if(column.sourceType==8 || column.type.contains("DATE",Qt::CaseInsensitive)) dates.append(column);return literalDefaultValues(database,dates,true);
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
