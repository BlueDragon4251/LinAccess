#include "access_query.hpp"
#include "vba_runtime.hpp"
#include <QCoreApplication>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
int main(int argc, char** argv) {
    QCoreApplication app(argc,argv);
    auto db = QSqlDatabase::addDatabase("QSQLITE");
    db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE samples(id INTEGER, score INTEGER, label TEXT)"));
    assert(query.exec("INSERT INTO samples VALUES(1,10,'one'),(2,20,'two'),(3,20,NULL),(4,30,'four')"));
    linaccess::AccessQueryModel model;
    model.execute(db,"SELECT TOP 2 id,score FROM samples ORDER BY score");
    assert(model.rowCount()==3 && model.columnCount()==2);
    assert(model.headerData(1,Qt::Horizontal).toString()=="score");
    assert(model.data(model.index(2,1)).toInt()==20);
    assert(!(model.flags(model.index(0,0)) & Qt::ItemIsEditable));
    model.execute(db,"PARAMETERS [minimum] Long; SELECT id,label FROM samples WHERE score >= [minimum] ORDER BY id",{{"minimum",20}});
    assert(model.rowCount()==3 && model.data(model.index(0,0)).toInt()==2);
    assert(model.data(model.index(1,1)).isNull());
    bool rejected=false;
    try { model.execute(db,"DELETE FROM samples"); } catch(const linaccess::VbaRuntimeError&) { rejected=true; }
    assert(rejected && model.rowCount()==0 && model.columnCount()==0);
    assert(query.exec("SELECT COUNT(*) FROM samples") && query.next() && query.value(0).toInt()==4);
    rejected=false;
    try { model.execute(db,"SELECT * FROM missing_dependency"); } catch(const linaccess::VbaRuntimeError&) { rejected=true; }
    assert(rejected && model.rowCount()==0);
    model.execute(db,"SELECT id FROM samples WHERE id=4");
    assert(model.rowCount()==1 && model.data(model.index(0,0)).toInt()==4);
    model.execute(db,"SELECT IIf(True, 7, 9) AS chosen, IIf(Null,1,2) AS null_condition, IIf(False,1,Null) AS null_result");
    assert(model.data(model.index(0,0)).toInt()==7 && model.data(model.index(0,1)).toInt()==2);
    assert(model.data(model.index(0,2)).isNull());
    model.execute(db,"SELECT IIf(True, 'kept', 9), IIf(False, 'discarded', 12)");
    assert(model.data(model.index(0,0)).metaType().id()==QMetaType::QString);
    assert(model.data(model.index(0,1)).toInt()==12);
    // Access evaluates both IIf arguments. SQLite's built-in IIf short-circuits.
    assert(query.exec("SELECT IIf(1,7,abs(-9223372036854775808))"));
    assert(query.next() && query.value(0).toInt()==7); query.finish();
    rejected = false;
    try { model.execute(db,"SELECT IIf(True,7,abs(-9223372036854775808))"); }
    catch(const linaccess::VbaRuntimeError&) { rejected = true; }
    assert(rejected && model.rowCount()==0);
    model.execute(db,"SELECT IIf('True',1,2), IIf('False',1,2)");
    assert(model.data(model.index(0,0)).toInt()==1 && model.data(model.index(0,1)).toInt()==2);
    rejected = false;
    try { model.execute(db,"SELECT IIf('invalid',1,2)"); }
    catch(const linaccess::VbaRuntimeError&) { rejected = true; }
    assert(rejected);

    assert(query.exec("CREATE TABLE aggregate_values(group_name TEXT,sequence INTEGER,value TEXT)"));
    assert(query.exec("INSERT INTO aggregate_values VALUES('a',1,NULL),('a',2,'last'),('b',1,'first'),('b',2,NULL)"));
    query.finish();
    model.execute(db,"SELECT group_name,First(value),Last(value) FROM (SELECT * FROM aggregate_values ORDER BY sequence) GROUP BY group_name ORDER BY group_name");
    assert(model.rowCount()==2 && model.columnCount()==3);
    assert(model.data(model.index(0,1)).isNull() && model.data(model.index(0,2)).toString()=="last");
    assert(model.data(model.index(1,1)).toString()=="first" && model.data(model.index(1,2)).isNull());
    model.execute(db,"SELECT First(value),Last(value) FROM aggregate_values WHERE 1=0");
    assert(model.rowCount()==1 && model.data(model.index(0,0)).isNull() && model.data(model.index(0,1)).isNull());
    model.execute(db,"SELECT First(9007199254740993),Last(7.5)");
    assert(model.data(model.index(0,0)).toLongLong()==9007199254740993LL);
    assert(model.data(model.index(0,1)).toDouble()==7.5);
    assert(query.exec("CREATE TABLE aggregate_blobs(value BLOB)"));
    query.prepare("INSERT INTO aggregate_blobs VALUES(?)");
    const auto binary = QByteArray::fromHex("0001ff00"); query.addBindValue(binary); assert(query.exec()); query.finish();
    model.execute(db,"SELECT First(value),Last(value) FROM aggregate_blobs");
    assert(model.data(model.index(0,0)).toByteArray()==binary && model.data(model.index(0,1)).toByteArray()==binary);
    query.clear(); db.close(); assert(db.open());
    model.execute(db,"SELECT IIf(True,5,6)");
    assert(model.data(model.index(0,0)).toInt()==5);
    model.execute(db,"SELECT First(12),Last('restored')");
    assert(model.data(model.index(0,0)).toInt()==12 && model.data(model.index(0,1)).toString()=="restored");

}
