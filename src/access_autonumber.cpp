#include "access_autonumber.hpp"
#include "jet_sql_functions.hpp"
#include "vba_runtime.hpp"
#include <QCryptographicHash>
#include <QSqlDriver>
#include <QSqlQuery>
#include <sqlite3.h>
#include <limits>
#include <memory>

namespace linaccess {
namespace {
QString identifier(QString value) { return '"'+value.replace('"',"\"\"")+'"'; }
QString literal(QString value) { return '\''+value.replace('\'',"''")+'\''; }
using Statement=std::unique_ptr<sqlite3_stmt,decltype(&sqlite3_finalize)>;
void counterError(sqlite3_context* context,int number) noexcept {
    const auto message=number==6 ? "LinAccess SQL function error 6" : number==94 ? "LinAccess SQL function error 94" : number==3035 ? "LinAccess SQL function error 3035" : "LinAccess SQL function error 3251";
    sqlite3_result_error(context,message,-1);
}
}
QString accessCounterDefault(const QString& table,const QString& column) {
    return "linaccess_next_counter("+literal(table)+','+literal(column)+')';
}
void nextAccessCounter(sqlite3_context* context,int count,sqlite3_value** values) noexcept {
    try {
        if (count!=2) { counterError(context,3251); return; }
        for (int i=0;i<2;++i) {
            if (sqlite3_value_type(values[i])==SQLITE_NULL) { counterError(context,94); return; }
            if (sqlite3_value_type(values[i])!=SQLITE_TEXT || sqlite3_value_bytes(values[i])>1024) { counterError(context,3251); return; }
        }
        auto* connection=sqlite3_context_db_handle(context); sqlite3_stmt* raw=nullptr;
        const auto prepare=sqlite3_prepare_v2(connection,"UPDATE _la_counter_sequences SET value=value+1 WHERE table_name=? AND column_name=? AND value<9223372036854775807 RETURNING value",-1,&raw,nullptr);
        Statement statement(raw,&sqlite3_finalize);
        if (prepare!=SQLITE_OK) { counterError(context,prepare==SQLITE_INTERRUPT ? 3035 : 3251); return; }
        for (int i=0;i<2;++i) if (sqlite3_bind_text(statement.get(),i+1,reinterpret_cast<const char*>(sqlite3_value_text(values[i])),sqlite3_value_bytes(values[i]),SQLITE_TRANSIENT)!=SQLITE_OK) { counterError(context,3251); return; }
        auto result=sqlite3_step(statement.get());
        if (result==SQLITE_ROW) {
            const auto next=sqlite3_column_int64(statement.get(),0);
            result=sqlite3_step(statement.get());
            if (result!=SQLITE_DONE) { counterError(context,result==SQLITE_INTERRUPT ? 3035 : 3251); return; }
            sqlite3_result_int64(context,next); return;
        }
        if (result!=SQLITE_DONE) { counterError(context,result==SQLITE_INTERRUPT ? 3035 : 3251); return; }
        statement.reset(); raw=nullptr;
        if (sqlite3_prepare_v2(connection,"SELECT value FROM _la_counter_sequences WHERE table_name=? AND column_name=?",-1,&raw,nullptr)!=SQLITE_OK) { counterError(context,3251); return; }
        statement=Statement(raw,&sqlite3_finalize);
        for (int i=0;i<2;++i) if (sqlite3_bind_text(statement.get(),i+1,reinterpret_cast<const char*>(sqlite3_value_text(values[i])),sqlite3_value_bytes(values[i]),SQLITE_TRANSIENT)!=SQLITE_OK) { counterError(context,3251); return; }
        const bool exhausted=sqlite3_step(statement.get())==SQLITE_ROW && sqlite3_column_int64(statement.get(),0)==std::numeric_limits<sqlite3_int64>::max();
        counterError(context,exhausted ? 6 : 3251);
    } catch (...) { counterError(context,3251); }
}
void createAccessCounter(QSqlDatabase db,const QString& table,const QString& column) {
    QSqlQuery query(db);
    if (!query.exec("CREATE TABLE IF NOT EXISTS _la_counter_sequences(table_name TEXT NOT NULL,column_name TEXT NOT NULL,value INTEGER NOT NULL CHECK(typeof(value)='integer' AND value>=0),PRIMARY KEY(table_name,column_name))"))
        throw VbaRuntimeError(3146,"Counter storage creation failed");
    query.prepare("INSERT INTO _la_counter_sequences VALUES(?,?,0)"); query.addBindValue(table); query.addBindValue(column);
    if (!query.exec()) throw VbaRuntimeError(3146,"Counter initialization failed");
    const auto key=literal(table)+" AND column_name="+literal(column);
    QByteArray identity=table.toUtf8(); identity.append('\0'); identity+=column.toUtf8();
    const auto prefix="_la_counter_"+QString::fromLatin1(QCryptographicHash::hash(identity,QCryptographicHash::Sha256).toHex());
    const auto body=" BEGIN SELECT CASE WHEN typeof(NEW."+identifier(column)+")!='integer' THEN RAISE(ABORT,'LinAccess SQL function error 13') END; UPDATE _la_counter_sequences SET value=MAX(value,NEW."+identifier(column)+") WHERE table_name="+key+"; END";
    if (!query.exec("CREATE TRIGGER "+identifier(prefix+"_insert")+" AFTER INSERT ON "+identifier(table)+body)
        || !query.exec("CREATE TRIGGER "+identifier(prefix+"_update")+" AFTER UPDATE OF "+identifier(column)+" ON "+identifier(table)+body))
        throw VbaRuntimeError(3146,"Counter seed trigger creation failed");
}
void protectAccessCounterUpdates(QSqlDatabase db,const QString& table,const QString& column) {
    QByteArray identity=table.toUtf8(); identity.append('\0'); identity+=column.toUtf8();
    const auto name="_la_counter_readonly_"+QString::fromLatin1(QCryptographicHash::hash(identity,QCryptographicHash::Sha256).toHex());
    QSqlQuery query(db);
    if(!query.exec("CREATE TRIGGER "+identifier(name)+" BEFORE UPDATE OF "+identifier(column)+" ON "+identifier(table)
        +" BEGIN SELECT RAISE(ABORT,'LinAccess SQL function error 3113'); END"))
        throw VbaRuntimeError(3146,"Source counter update protection failed");
}
void restoreAccessCounter(QSqlDatabase db,const QString& table,const QString& column,qint32 last,bool native) {
    if(last<0) throw VbaRuntimeError(3251,"Unsupported source counter state");
    QSqlQuery query(db);
    if(!query.exec("SELECT MAX("+identifier(column)+") FROM "+identifier(table)) || !query.next())
        throw VbaRuntimeError(3146,"Counter row validation failed");
    if(!query.value(0).isNull() && query.value(0).toLongLong()>last)
        throw VbaRuntimeError(3251,"Source counter state contradicts active rows");
    query.finish();
    if(native) {
        query.prepare("UPDATE main.sqlite_sequence SET seq=? WHERE name=?");
        query.addBindValue(last); query.addBindValue(table);
        if(!query.exec()) throw VbaRuntimeError(3146,"Native counter restoration failed");
        if(query.numRowsAffected()==0) {
            query.prepare("INSERT INTO main.sqlite_sequence(name,seq) VALUES(?,?)");
            query.addBindValue(table); query.addBindValue(last);
            if(!query.exec()) throw VbaRuntimeError(3146,"Empty native counter restoration failed");
        }
    } else {
        query.prepare("UPDATE main._la_counter_sequences SET value=? WHERE table_name=? AND column_name=?");
        query.addBindValue(last); query.addBindValue(table); query.addBindValue(column);
        if(!query.exec() || query.numRowsAffected()!=1) throw VbaRuntimeError(3146,"Managed counter restoration failed");
    }
}
QSet<QString> generatedAccessColumns(QSqlDatabase db,const QString& requested) {
    installJetSqlFunctions(db);
    QString table;
    for (const auto& candidate:db.tables()) if (candidate.compare(requested,Qt::CaseInsensitive)==0) { table=candidate; break; }
    if (table.isEmpty()) throw VbaRuntimeError(3078,"Counter table unavailable");
    auto handle=db.driver()->handle();
    if (!handle.isValid() || QByteArray(handle.typeName())!="sqlite3*") throw VbaRuntimeError(3251,"Counter metadata adapter unavailable");
    auto* connection=*static_cast<sqlite3**>(handle.data());
    if (!connection) throw VbaRuntimeError(3251,"Counter connection unavailable");
    QSqlQuery query(db);
    if (!query.exec("PRAGMA table_info("+identifier(table)+")")) throw VbaRuntimeError(3146,"Counter column metadata unavailable");
    QSet<QString> result;
    while (query.next()) {
        const auto name=query.value(1).toString(),initial=query.value(4).toString(); int generated=0;
        if (sqlite3_table_column_metadata(connection,"main",table.toUtf8().constData(),name.toUtf8().constData(),nullptr,nullptr,nullptr,nullptr,&generated)!=SQLITE_OK)
            throw VbaRuntimeError(3146,"Counter column metadata unavailable");
        const auto expected=accessCounterDefault(table,name);
        if (generated || initial==expected || initial=='('+expected+')') result.insert(name.toCaseFolded());
    }
    return result;
}
}
