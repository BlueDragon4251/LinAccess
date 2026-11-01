#include "access_query.hpp"
#include "dao_database.hpp"
#include "vba_runtime.hpp"
#include <QCoreApplication>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>

int main(int argc,char** argv) {
    QCoreApplication application(argc,argv); using namespace linaccess;
    auto db=QSqlDatabase::addDatabase("QSQLITE"); db.setDatabaseName(":memory:"); assert(db.open()); QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE target(id INTEGER PRIMARY KEY,stock REAL,note TEXT,flag INTEGER,bytes BLOB)"));
    assert(sql.exec("CREATE TABLE source(id INTEGER,delta REAL,note TEXT,flag INTEGER,bytes BLOB)"));
    assert(sql.exec("INSERT INTO target VALUES(1,10,'old',0,NULL),(2,20,'old',-1,NULL),(3,30,'untouched',0,NULL)"));
    assert(sql.exec("INSERT INTO source VALUES(1,2,NULL,-1,x'00ff'),(2,5,'new',0,x'7f80')"));
    auto database=makeDaoDatabase(db);
    const auto execute=[&](const char* text) { return database->member("Execute",{VbaValue::string(text),VbaValue::numeric(128)}); };
    execute("UPDATE target AS d INNER JOIN source AS s ON d.id=s.id SET d.stock=d.stock+s.delta,d.note=s.note,d.flag=s.flag,d.bytes=s.bytes WHERE d.id<=2");
    assert(database->member("RecordsAffected",{}).asNumber()==2);
    assert(sql.exec("SELECT stock,note,flag,bytes FROM target ORDER BY id") && sql.next());
    assert(sql.value(0).toDouble()==12 && sql.value(1).isNull() && sql.value(2).toInt()==-1 && sql.value(3).toByteArray()==QByteArray::fromHex("00ff"));
    assert(sql.next() && sql.value(0).toDouble()==25 && sql.value(1).toString()=="new" && sql.value(2).toInt()==0 && sql.value(3).toByteArray()==QByteArray::fromHex("7f80"));
    assert(sql.next() && sql.value(0).toDouble()==30 && sql.value(1).toString()=="untouched"); sql.finish();
    auto definition=database->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string("PARAMETERS [step] Long; UPDATE target AS d INNER JOIN source AS s ON d.id=s.id SET d.stock=d.stock+[step] WHERE s.id=1")}).object;
    definition->member("Parameters",{VbaValue::string("step")}).object->assign("Value",VbaValue::numeric(3)); definition->member("Execute",{});
    assert(definition->member("RecordsAffected",{}).asNumber()==1);
    assert(sql.exec("SAVEPOINT identical_source"));
    assert(sql.exec("INSERT INTO source SELECT * FROM source WHERE id=2"));
    execute("UPDATE target AS d INNER JOIN source AS s ON d.id=s.id SET d.stock=d.stock+s.delta,d.note=s.note,d.bytes=s.bytes");
    assert(database->member("RecordsAffected",{}).asNumber()==2);
    assert(sql.exec("SELECT stock,note,bytes FROM target WHERE id=2") && sql.next() && sql.value(0).toDouble()==30 && sql.value(1).toString()=="new" && sql.value(2).toByteArray()==QByteArray::fromHex("7f80")); sql.finish();
    assert(sql.exec("ROLLBACK TO identical_source") && sql.exec("RELEASE identical_source"));
    assert(sql.exec("INSERT INTO source VALUES(2,7,'ambiguous',0,NULL)"));
    int error=0; try { execute("UPDATE target AS d INNER JOIN source AS s ON d.id=s.id SET d.stock=d.stock+s.delta"); }
    catch (const VbaRuntimeError& failure) { error=failure.number; } assert(error==3251);
    assert(sql.exec("SELECT stock FROM target ORDER BY id") && sql.next() && sql.value(0).toDouble()==15);
    assert(sql.next() && sql.value(0).toDouble()==25 && sql.next() && sql.value(0).toDouble()==30); sql.finish();
    execute("UPDATE target AS d INNER JOIN target AS s ON d.id=s.id+1 SET d.stock=s.stock");
    assert(database->member("RecordsAffected",{}).asNumber()==2);
    assert(sql.exec("SELECT stock FROM target ORDER BY id") && sql.next() && sql.value(0).toDouble()==15);
    assert(sql.next() && sql.value(0).toDouble()==15 && sql.next() && sql.value(0).toDouble()==25); sql.finish();
    execute("UPDATE source AS s INNER JOIN target AS t ON s.id=t.id SET t.note='comma, WHERE; JOIN' WHERE s.id=1");
    assert(database->member("RecordsAffected",{}).asNumber()==1);
    assert(sql.exec("SELECT note FROM target WHERE id=1") && sql.next() && sql.value(0).toString()=="comma, WHERE; JOIN"); sql.finish();
    assert(sql.exec("CREATE TABLE shadow(rowid TEXT,id INTEGER PRIMARY KEY,value INTEGER)"));
    assert(sql.exec("INSERT INTO shadow VALUES('text identity',1,0)"));
    execute("UPDATE shadow AS d INNER JOIN source AS s ON d.id=s.id SET d.value=42 WHERE s.id=1");
    assert(sql.exec("SELECT rowid,value FROM shadow") && sql.next() && sql.value(0).toString()=="text identity" && sql.value(1).toInt()==42); sql.finish();
    assert(sql.exec("CREATE VIEW readonly_target AS SELECT * FROM target"));
    assert(sql.exec("CREATE TABLE no_rowid(id INTEGER PRIMARY KEY,value INTEGER) WITHOUT ROWID"));
    assert(sql.exec("INSERT INTO no_rowid VALUES(1,0)"));
    for (const auto* text:{
        "UPDATE target AS d LEFT JOIN source AS s ON d.id=s.id SET d.stock=s.delta",
        "UPDATE target AS d INNER JOIN source AS s ON d.id=s.id SET d.stock=1,s.delta=2",
        "UPDATE readonly_target AS d INNER JOIN source AS s ON d.id=s.id SET d.stock=1",
        "UPDATE no_rowid AS d INNER JOIN source AS s ON d.id=s.id SET d.value=1"}) {
        error=0; try { execute(text); } catch (const VbaRuntimeError& failure) { error=failure.number; } assert(error==3251);
    }
    execute("UPDATE target AS d INNER JOIN source AS s ON d.id=s.id SET d.stock=1 WHERE s.id=999");
    assert(database->member("RecordsAffected",{}).asNumber()==0);
    std::cout<<"Joined UPDATE: native arithmetic, parameters, typed NULL/binary values, frozen self-join, ambiguity rollback and writable-target boundaries passed\n";
}
