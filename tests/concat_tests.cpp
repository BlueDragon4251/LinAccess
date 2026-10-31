#include "access_query.hpp"
#include "vba_runtime.hpp"
#include <QCoreApplication>
#include <QLocale>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
int main(int argc,char** argv) {
    QCoreApplication app(argc,argv);
    auto db=QSqlDatabase::addDatabase("QSQLITE"); db.setDatabaseName(":memory:"); assert(db.open());
    linaccess::AccessQueryModel model;
    model.execute(db,"SELECT 1+2 & 3*4 AS value, 'a' & 1+2 AS text, (1+2) & 4 AS grouped");
    assert(model.data(model.index(0,0)).toString()=="312");
    assert(model.data(model.index(0,1)).toString()=="a3");
    assert(model.data(model.index(0,2)).toString()=="34");
    model.execute(db,"SELECT IIf(1+2 & 3='33','yes','no'), 'a' & 'b' & 'c', 'März' & '!'");
    assert(model.data(model.index(0,0)).toString()=="yes");
    assert(model.data(model.index(0,1)).toString()=="abc");
    assert(model.data(model.index(0,2)).toString()=="März!");
    model.execute(db,"SELECT Null & Null, Null & 'x', 'x' & Null, Null & '', '&' & '&'");
    assert(model.data(model.index(0,0)).isNull());
    assert(model.data(model.index(0,1)).toString()=="x" && model.data(model.index(0,2)).toString()=="x");
    assert(!model.data(model.index(0,3)).isNull() && model.data(model.index(0,3)).toString().isEmpty());
    assert(model.data(model.index(0,4)).toString()=="&&");
    model.execute(db,"SELECT 'accepted' WHERE (1+2) & 4='34'");
    assert(model.rowCount()==1);
    model.execute(db,"SELECT 'ab' & 'c' LIKE('a*'), 'ab' & 'c' IN('abc'), 'a_b' LIKE('a_' & '*')");
    assert(model.data(model.index(0,0)).toBool() && model.data(model.index(0,1)).toBool() && model.data(model.index(0,2)).toBool());
    model.execute(db,"SELECT 'x' WHERE ('a' & 'b')='ab' AND 'abc' LIKE 'a' & 'b' & '*'"); assert(model.rowCount()==1);
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE samples(id INTEGER,label TEXT)"));
    assert(query.exec("INSERT INTO samples VALUES(1,'ab%_one'),(2,'abZZone'),(3,'second')")); query.finish();
    model.execute(db,"PARAMETERS prefix Text(255); SELECT label FROM samples WHERE label LIKE [prefix] & '*' ORDER BY id",{{"prefix","ab%_"}});
    assert(model.rowCount()==1 && model.data(model.index(0,0)).toString()=="ab%_one");
    model.execute(db,"SELECT 'abc' LIKE (SELECT 'a*' WHERE 'yes' LIKE 'y*')");
    assert(model.rowCount()==1 && model.data(model.index(0,0)).toBool());
    model.execute(db,"SELECT 'ab' LIKE 'a' & '?', Null LIKE 'a' & '*'");
    assert(model.data(model.index(0,0)).toBool() && model.data(model.index(0,1)).isNull());
    bool unsupported=false;
    try { model.execute(db,"SELECT 'abc' LIKE 'a' & '[bc]'"); }
    catch(const linaccess::VbaRuntimeError& error) { unsupported=error.number==3085; }
    assert(unsupported);
    const auto original=QLocale(); QLocale::setDefault(QLocale(QLocale::German,QLocale::Germany));
    model.execute(db,"SELECT 1.5 & '', 1234567 & '', 9007199254740993 & ''");
    assert(model.data(model.index(0,0)).toString()=="1,5");
    assert(model.data(model.index(0,1)).toString()=="1234567");
    assert(model.data(model.index(0,2)).toString()=="9007199254740993"); QLocale::setDefault(original);
    bool malformed=false;
    try { model.execute(db,"SELECT 'a' &"); } catch(const linaccess::VbaRuntimeError&) { malformed=true; }
    assert(malformed && model.rowCount()==0);
    assert(query.exec("CREATE TABLE binary_values(value BLOB)"));
    query.prepare("INSERT INTO binary_values VALUES(?)"); query.addBindValue(QByteArray::fromHex("00ff")); assert(query.exec()); query.finish();
    bool rejectedBlob=false;
    try { model.execute(db,"SELECT value & 'x' FROM binary_values"); }
    catch(const linaccess::VbaRuntimeError& error) { rejectedBlob=error.number==13; }
    assert(rejectedBlob);
}
