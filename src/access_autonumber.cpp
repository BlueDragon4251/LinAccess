#include "access_autonumber.hpp"
#include "jet_sql_functions.hpp"
#include "vba_runtime.hpp"
#include <QCryptographicHash>
#include <QSqlDriver>
#include <QSqlQuery>
#include <QSqlError>
#include <QMap>
#include <QUuid>
#include <sqlite3.h>
#include <limits>
#include <memory>
#include <algorithm>
#include <cstring>

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
struct AccessCounterSqlScope::State {
    struct Capture {AccessCounterReservation reservation;qlonglong previous{};};
    QSqlDatabase db;sqlite3* connection{};State* parent{};bool completed{};
    QMap<QString,Capture> captures;
};
namespace {
thread_local AccessCounterSqlScope::State* activeCounterSql=nullptr;
void captureCounterSql(sqlite3* connection,sqlite3_value** names,qlonglong value,qlonglong increment,bool jetLong) {
    auto* frame=activeCounterSql;
    while(frame && frame->connection!=connection) frame=frame->parent;
    if(!frame) return;
    const auto table=QString::fromUtf8(reinterpret_cast<const char*>(sqlite3_value_text(names[0])),sqlite3_value_bytes(names[0]));
    const auto column=QString::fromUtf8(reinterpret_cast<const char*>(sqlite3_value_text(names[1])),sqlite3_value_bytes(names[1]));
    const auto key=table+QChar(0)+column;
    auto previous=value-increment;if(jetLong && previous<std::numeric_limits<qint32>::min()) previous+=4294967296LL;
    auto it=frame->captures.find(key);
    if(it==frame->captures.end()) frame->captures.insert(key,{{table,column,value,false,jetLong},previous});
    else it->reservation.value=value;
}
}
AccessCounterSqlScope::AccessCounterSqlScope(QSqlDatabase db):state_(std::make_unique<State>()) {
    if(!db.isOpen() || !db.driver()) throw VbaRuntimeError(3420,"Counter SQL connection unavailable");
    const auto handle=db.driver()->handle();
    if(!handle.isValid() || QByteArray(handle.typeName())!="sqlite3*") throw VbaRuntimeError(3251,"Counter SQL backend unavailable");
    state_->db=db;state_->connection=*static_cast<sqlite3* const*>(handle.constData());
    state_->parent=activeCounterSql;activeCounterSql=state_.get();
}
AccessCounterSqlScope::~AccessCounterSqlScope() {
    if(!state_->completed) {
        try {complete(false);} catch(...) {if(state_->db.driver()) state_->db.driver()->setProperty("la_counter_sql_sync_failed",true);}
    }
    activeCounterSql=state_->parent;
}
void AccessCounterSqlScope::complete(bool succeeded) {
    if(state_->completed) return;
    QVector<AccessCounterReservation> consumed;
    for(const auto& capture:state_->captures) {
        const auto& reservation=capture.reservation;
        if(!succeeded) {
            QSqlQuery query(state_->db);
            const auto storage=reservation.jetLong ? QString("_la_counter32_sequences") : QString("_la_counter_sequences");
            query.prepare("UPDATE main."+identifier(storage)+" SET value=? WHERE table_name=? AND column_name=? AND value=?");
            query.addBindValue(reservation.value);query.addBindValue(reservation.table);query.addBindValue(reservation.column);query.addBindValue(capture.previous);
            if(!query.exec()) throw VbaRuntimeError(3146,"Consumed SQL counter persistence failed");
            // A changed frontier belongs to another completed allocation. Never
            // rewind it after a failed statement releases its SQLite lock.
        }
        consumed.append(reservation);
        auto* parent=state_->parent;while(parent && parent->connection!=state_->connection) parent=parent->parent;
        if(parent) {
            const auto key=reservation.table+QChar(0)+reservation.column;
            auto it=parent->captures.find(key);
            if(it==parent->captures.end()) parent->captures.insert(key,capture);else it->reservation.value=reservation.value;
        }
    }
    if(!consumed.isEmpty()) trackAccessCounterReservations(state_->db,consumed);
    state_->completed=true;
    if(activeCounterSql==state_.get()) activeCounterSql=state_->parent;
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
        sqlite3_int64 increment=1;
        const auto stepPrepare=sqlite3_prepare_v2(connection,"SELECT increment,typeof(increment) FROM main._la_counter32_steps WHERE table_name=? AND column_name=?",-1,&raw,nullptr);
        Statement step(raw,&sqlite3_finalize);
        if(stepPrepare==SQLITE_OK) {
            for(int i=0;i<2;++i) if(sqlite3_bind_text(step.get(),i+1,reinterpret_cast<const char*>(sqlite3_value_text(values[i])),sqlite3_value_bytes(values[i]),SQLITE_TRANSIENT)!=SQLITE_OK) {counterError(context,3251);return;}
            const auto result=sqlite3_step(step.get());
            if(result==SQLITE_ROW) {
                const auto* storage=sqlite3_column_text(step.get(),1);increment=sqlite3_column_int64(step.get(),0);
                if(!storage || strcmp(reinterpret_cast<const char*>(storage),"integer") || increment<1 || increment>2147483647 || sqlite3_step(step.get())!=SQLITE_DONE) {counterError(context,3251);return;}
            } else if(result!=SQLITE_DONE) {counterError(context,3251);return;}
        } else {
            step.reset();raw=nullptr;
            if(sqlite3_prepare_v2(connection,"SELECT 1 FROM main.sqlite_schema WHERE name='_la_counter32_steps'",-1,&raw,nullptr)!=SQLITE_OK) {counterError(context,3251);return;}
            Statement exists(raw,&sqlite3_finalize);if(sqlite3_step(exists.get())==SQLITE_ROW) {counterError(context,3251);return;}
        }
        step.reset();raw=nullptr;
        // Source-backed Jet Long counters wrap in signed32; legacy SQLite-only
        // generators keep their existing64-bit contract.
        const auto jetPrepare=sqlite3_prepare_v2(connection,"UPDATE main._la_counter32_sequences SET value=CASE WHEN value+?3>2147483647 THEN value+?3-4294967296 ELSE value+?3 END WHERE table_name=?1 AND column_name=?2 RETURNING value",-1,&raw,nullptr);
        Statement jet(raw,&sqlite3_finalize);
        if(jetPrepare==SQLITE_OK) {
            for(int i=0;i<2;++i) if(sqlite3_bind_text(jet.get(),i+1,reinterpret_cast<const char*>(sqlite3_value_text(values[i])),sqlite3_value_bytes(values[i]),SQLITE_TRANSIENT)!=SQLITE_OK) {counterError(context,3251);return;}
            if(sqlite3_bind_int64(jet.get(),3,increment)!=SQLITE_OK) {counterError(context,3251);return;}
            const auto result=sqlite3_step(jet.get());
            if(result==SQLITE_ROW) {
                const auto value=sqlite3_column_int64(jet.get(),0);
                if(sqlite3_step(jet.get())!=SQLITE_DONE) {counterError(context,3251);return;}
                captureCounterSql(connection,values,value,increment,true);
                sqlite3_result_int64(context,value);return;
            }
            if(result!=SQLITE_DONE) {counterError(context,result==SQLITE_INTERRUPT ? 3035 : 3251);return;}
        } else {
            jet.reset();raw=nullptr;
            if(sqlite3_prepare_v2(connection,"SELECT 1 FROM main.sqlite_schema WHERE name='_la_counter32_sequences'",-1,&raw,nullptr)!=SQLITE_OK) {counterError(context,3251);return;}
            Statement exists(raw,&sqlite3_finalize);
            if(sqlite3_step(exists.get())==SQLITE_ROW) {counterError(context,3251);return;}
        }
        jet.reset();raw=nullptr;
        const auto prepare=sqlite3_prepare_v2(connection,"UPDATE _la_counter_sequences SET value=value+1 WHERE table_name=? AND column_name=? AND value<9223372036854775807 RETURNING value",-1,&raw,nullptr);
        Statement statement(raw,&sqlite3_finalize);
        if (prepare!=SQLITE_OK) { counterError(context,prepare==SQLITE_INTERRUPT ? 3035 : 3251); return; }
        for (int i=0;i<2;++i) if (sqlite3_bind_text(statement.get(),i+1,reinterpret_cast<const char*>(sqlite3_value_text(values[i])),sqlite3_value_bytes(values[i]),SQLITE_TRANSIENT)!=SQLITE_OK) { counterError(context,3251); return; }
        auto result=sqlite3_step(statement.get());
        if (result==SQLITE_ROW) {
            const auto next=sqlite3_column_int64(statement.get(),0);
            result=sqlite3_step(statement.get());
            if (result!=SQLITE_DONE) { counterError(context,result==SQLITE_INTERRUPT ? 3035 : 3251); return; }
            captureCounterSql(connection,values,next,1,false);
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
QSet<QString> jetLongAccessColumns(QSqlDatabase db,const QString& table) {
    QSet<QString> result;
    if(!db.tables().contains("_la_counter32_sequences",Qt::CaseInsensitive)) return result;
    QSqlQuery query(db);
    query.prepare("SELECT column_name,value,typeof(value) FROM main._la_counter32_sequences WHERE table_name=? COLLATE NOCASE");query.addBindValue(table);
    if(!query.exec()) throw VbaRuntimeError(3146,"Jet Long counter catalog unavailable");
    while(query.next()) {
        const auto name=query.value(0).toString().toCaseFolded();const auto value=query.value(1).toLongLong();
        if(name.isEmpty() || result.contains(name) || query.value(2).toString()!="integer" || value<std::numeric_limits<qint32>::min() || value>std::numeric_limits<qint32>::max())
            throw VbaRuntimeError(3251,"Jet Long counter catalog is not proven");
        result.insert(name);
    }
    query.finish();
    if(!result.isEmpty()) {
        const auto generators=generatedAccessColumns(db,table);
        for(const auto& name:result) if(!generators.contains(name)) throw VbaRuntimeError(3251,"Jet Long generator identity changed");
    }
    return result;
}
void configureNewAccessCounterStep(QSqlDatabase db,const QString& table,const QString& column,qint32 seed,qint32 increment) {
    if(increment<=0 || !jetLongAccessColumns(db,table).contains(column.toCaseFolded())) throw VbaRuntimeError(3251,"Counter step configuration unavailable");
    QSqlQuery query(db);
    if(!query.exec("SELECT COUNT(*) FROM main."+identifier(table)) || !query.next() || query.value(0).toLongLong()!=0)
        throw VbaRuntimeError(3251,"Counter seed configuration requires a new empty table");
    query.finish();
    if(!query.exec("CREATE TABLE IF NOT EXISTS main._la_counter32_steps(table_name TEXT NOT NULL,column_name TEXT NOT NULL,increment INTEGER NOT NULL CHECK(typeof(increment)='integer' AND increment BETWEEN 1 AND 2147483647),PRIMARY KEY(table_name,column_name))"))
        throw VbaRuntimeError(3146,"Counter step storage unavailable");
    query.prepare("INSERT INTO main._la_counter32_steps VALUES(?,?,?)");query.addBindValue(table);query.addBindValue(column);query.addBindValue(increment);
    if(!query.exec()) throw VbaRuntimeError(3146,"Counter step initialization failed");
    qlonglong previous=static_cast<qlonglong>(seed)-increment;
    if(previous<std::numeric_limits<qint32>::min()) previous+=4294967296LL;
    if(previous>std::numeric_limits<qint32>::max()) previous-=4294967296LL;
    query.prepare("UPDATE main._la_counter32_sequences SET value=? WHERE table_name=? AND column_name=?");
    query.addBindValue(previous);query.addBindValue(table);query.addBindValue(column);
    if(!query.exec() || query.numRowsAffected()!=1) throw VbaRuntimeError(3146,"Counter seed initialization failed");
}
void configureNewAccessCounters(QSqlDatabase db,const QString& table) {
    const auto generators=generatedAccessColumns(db,table);
    if(generators.isEmpty()) return;
    QSqlQuery query(db);
    if(!query.exec("SELECT COUNT(*) FROM main."+identifier(table)) || !query.next() || query.value(0).toLongLong()!=0)
        throw VbaRuntimeError(3251,"Counter setup requires a newly created empty table");
    query.finish();
    if(!query.exec("PRAGMA main.table_info("+identifier(table)+")")) throw VbaRuntimeError(3146,"New counter fields unavailable");
    QStringList names;
    while(query.next()) if(generators.contains(query.value(1).toString().toCaseFolded())) names.append(query.value(1).toString());
    query.finish();
    for(const auto& column:names) {createJetLongCounter(db,table,column);protectAccessCounterUpdates(db,table,column);}
}
void createJetLongCounter(QSqlDatabase db,const QString& table,const QString& column) {
    QSqlQuery query(db);
    if(!query.exec("CREATE TABLE IF NOT EXISTS main._la_counter32_sequences(table_name TEXT NOT NULL,column_name TEXT NOT NULL,value INTEGER NOT NULL CHECK(typeof(value)='integer' AND value BETWEEN -2147483648 AND 2147483647),PRIMARY KEY(table_name,column_name))"))
        throw VbaRuntimeError(3146,"Jet Long counter storage unavailable");
    query.prepare("INSERT INTO main._la_counter32_sequences VALUES(?,?,0)");query.addBindValue(table);query.addBindValue(column);
    if(!query.exec()) throw VbaRuntimeError(3146,"Jet Long counter initialization failed");
    QByteArray identity=table.toUtf8();identity.append('\0');identity+=column.toUtf8();
    const auto name="_la_counter32_"+QString::fromLatin1(QCryptographicHash::hash(identity,QCryptographicHash::Sha256).toHex());
    if(!query.exec("CREATE TRIGGER "+identifier(name)+" AFTER INSERT ON "+identifier(table)
        +" BEGIN SELECT CASE WHEN typeof(NEW."+identifier(column)+")!='integer' OR NEW."+identifier(column)+" NOT BETWEEN -2147483648 AND 2147483647 THEN RAISE(ABORT,'LinAccess SQL function error 3251: native counter SQL wrap requires explicit Jet allocation') END; UPDATE _la_counter32_sequences SET value=MAX(value,NEW."
        +identifier(column)+") WHERE table_name="+literal(table)+" AND column_name="+literal(column)+"; END"))
        throw VbaRuntimeError(3146,"Jet Long counter observation unavailable");
    const auto handle=db.driver()->handle();auto* connection=*static_cast<sqlite3* const*>(handle.constData());int native=0;
    if(sqlite3_table_column_metadata(connection,"main",table.toUtf8().constData(),column.toUtf8().constData(),nullptr,nullptr,nullptr,nullptr,&native)!=SQLITE_OK)
        throw VbaRuntimeError(3251,"Jet Long field identity unavailable");
    if(!native && !query.exec("CREATE TRIGGER "+identifier(name+"_null")+" BEFORE INSERT ON "+identifier(table)
        +" WHEN NEW."+identifier(column)+" IS NULL BEGIN SELECT RAISE(ABORT,'LinAccess SQL function error 3162'); END"))
        throw VbaRuntimeError(3146,"Jet Long Null guard unavailable");
}
AccessCounterReservation currentAccessCounterReservation(QSqlDatabase db,AccessCounterReservation reservation) {
    QSqlQuery query(db);
    if(reservation.jetLong) query.prepare("SELECT value FROM main._la_counter32_sequences WHERE table_name=? AND column_name=?");
    else if(reservation.native) query.prepare("SELECT seq FROM main.sqlite_sequence WHERE name=?");
    else query.prepare("SELECT value FROM main._la_counter_sequences WHERE table_name=? AND column_name=?");
    query.addBindValue(reservation.table);if(!reservation.native || reservation.jetLong) query.addBindValue(reservation.column);
    if(!query.exec() || !query.next()) throw VbaRuntimeError(3146,"Current reserved counter state unavailable");
    const auto value=query.value(0).toLongLong();
    reservation.value=reservation.jetLong ? value : std::max(value,reservation.value);
    return reservation;
}
AccessCounterReservation reserveAccessCounter(QSqlDatabase db,const QString& requestedTable,const QString& column) {
    QString table;
    for(const auto& name:db.tables()) if(name.compare(requestedTable,Qt::CaseInsensitive)==0) {table=name;break;}
    if(table.isEmpty()) throw VbaRuntimeError(3078,"Counter table unavailable");
    if(!generatedAccessColumns(db,table).contains(column.toCaseFolded())) throw VbaRuntimeError(3251,"Counter generator unavailable");
    auto handle=db.driver()->handle();
    auto* connection=*static_cast<sqlite3**>(handle.data());int native=0;
    if(sqlite3_table_column_metadata(connection,"main",table.toUtf8().constData(),column.toUtf8().constData(),nullptr,nullptr,nullptr,nullptr,&native)!=SQLITE_OK)
        throw VbaRuntimeError(3251,"Counter generator identity unavailable");
    const auto scope="la_counter_reserve_"+QUuid::createUuid().toString(QUuid::Id128);
    QSqlQuery guard(db),query(db);
    if(!guard.exec("SAVEPOINT "+scope)) throw VbaRuntimeError(3146,"Counter reservation unavailable");
    try {
        qlonglong value{};bool jetLong=false;
        if(db.tables().contains("_la_counter32_sequences",Qt::CaseInsensitive)) {
            query.prepare("SELECT 1 FROM main._la_counter32_sequences WHERE table_name=? AND column_name=?");query.addBindValue(table);query.addBindValue(column);
            if(!query.exec()) throw VbaRuntimeError(3146,"Jet Long counter metadata unavailable");
            jetLong=query.next();query.finish();
        }
        if(jetLong) {
            query.prepare("SELECT linaccess_next_counter(?,?)");query.addBindValue(table);query.addBindValue(column);
            if(!query.exec() || !query.next()) throw VbaRuntimeError(3146,"Jet Long counter reservation failed");
            value=query.value(0).toLongLong();query.finish();
            retainAccessCounterReservation(db,{table,column,value,native!=0,true});
        } else if(native) {
            query.prepare("SELECT seq,typeof(seq) FROM main.sqlite_sequence WHERE name=?");query.addBindValue(table);
            if(!query.exec()) throw VbaRuntimeError(3146,"Native counter state unavailable");
            const bool stored=query.next();
            if(stored) {
                bool valid=false;value=query.value(0).toLongLong(&valid);
                if(!valid || query.value(1).toString()!="integer" || value<0 || query.next()) throw VbaRuntimeError(3251,"Native counter state is not proven");
            }
            query.finish();
            if(!query.exec("SELECT MAX("+identifier(column)+") FROM main."+identifier(table)) || !query.next())
                throw VbaRuntimeError(3146,"Native counter row state unavailable");
            if(!query.value(0).isNull()) value=std::max(value,query.value(0).toLongLong());
            query.finish();
            if(value==std::numeric_limits<qlonglong>::max()) throw VbaRuntimeError(6,"Native counter exhausted");
            ++value;
            query.prepare(stored ? "UPDATE main.sqlite_sequence SET seq=? WHERE name=?" : "INSERT INTO main.sqlite_sequence(seq,name) VALUES(?,?)");
            query.addBindValue(value);query.addBindValue(table);
            if(!query.exec() || query.numRowsAffected()!=1) throw VbaRuntimeError(3146,"Native counter reservation failed");
        } else {
            query.prepare("SELECT linaccess_next_counter(?,?)");query.addBindValue(table);query.addBindValue(column);
            if(!query.exec() || !query.next()) {
                const auto error=jetSqlFunctionErrorNumber(query.lastError().text().toLower().toStdString());
                throw VbaRuntimeError(error.value_or(3146),"Managed counter reservation failed");
            }
            value=query.value(0).toLongLong();query.finish();
        }
        if(!guard.exec("RELEASE "+scope)) throw VbaRuntimeError(3146,"Counter reservation commit failed");
        return {table,column,value,native!=0,jetLong};
    } catch(...) {guard.exec("ROLLBACK TO "+scope);guard.exec("RELEASE "+scope);throw;}
}
void retainAccessCounterReservation(QSqlDatabase db,const AccessCounterReservation& reservation) {
    // Called after ROLLBACK TO but before RELEASE: the original transaction keeps
    // its write lock while business changes roll back and consumed IDs survive.
    if(!db.tables().contains(reservation.table,Qt::CaseInsensitive)) return; // Created inside the rolled-back transaction.
    if(!generatedAccessColumns(db,reservation.table).contains(reservation.column.toCaseFolded()))
        throw VbaRuntimeError(3251,"Reserved counter identity changed");
    QSqlQuery query(db);
    if(reservation.jetLong) {
        query.prepare("UPDATE main._la_counter32_sequences SET value=? WHERE table_name=? AND column_name=?");
        query.addBindValue(reservation.value);query.addBindValue(reservation.table);query.addBindValue(reservation.column);
        if(!query.exec() || query.numRowsAffected()!=1) throw VbaRuntimeError(3146,"Reserved Jet Long counter persistence failed");
        if(!reservation.native || reservation.value<0) return;
    }
    if(reservation.native) {
        query.prepare("SELECT 1 FROM main.sqlite_sequence WHERE name=?");query.addBindValue(reservation.table);
        if(!query.exec()) throw VbaRuntimeError(3146,"Reserved native counter unavailable");
        const bool exists=query.next();query.finish();
        query.prepare(exists ? "UPDATE main.sqlite_sequence SET seq=MAX(seq,?) WHERE name=?" : "INSERT INTO main.sqlite_sequence(seq,name) VALUES(?,?)");
        query.addBindValue(reservation.value);query.addBindValue(reservation.table);
    } else {
        query.prepare("UPDATE main._la_counter_sequences SET value=MAX(value,?) WHERE table_name=? AND column_name=?");
        query.addBindValue(reservation.value);query.addBindValue(reservation.table);query.addBindValue(reservation.column);
    }
    if(!query.exec() || query.numRowsAffected()!=1) throw VbaRuntimeError(3146,"Consumed counter persistence failed");
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
    if(db.tables().contains("_la_counter32_sequences",Qt::CaseInsensitive)) {
        query.prepare("UPDATE main._la_counter32_sequences SET value=? WHERE table_name=? AND column_name=?");
        query.addBindValue(last);query.addBindValue(table);query.addBindValue(column);
        if(!query.exec()) throw VbaRuntimeError(3146,"Source Jet Long counter restoration failed");
    }
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
