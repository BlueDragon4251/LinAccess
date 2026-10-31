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
}
