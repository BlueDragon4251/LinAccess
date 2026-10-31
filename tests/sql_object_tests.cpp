#include "access_query.hpp"
#include "jet_sql_functions.hpp"
#include "vba_runtime.hpp"
#include <QCoreApplication>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
int main(int argc,char** argv) {
    QCoreApplication app(argc,argv);
    auto db=QSqlDatabase::addDatabase("QSQLITE"); db.setDatabaseName(":memory:"); assert(db.open());
    linaccess::AccessQueryModel model;
    bool missing=false;
    try { model.execute(db,"SELECT Formulare!filter!year"); }
    catch(const linaccess::VbaRuntimeError& error) { missing=error.number==3061; }
    assert(missing && model.rowCount()==0);
    auto year=linaccess::VbaValue::numeric(2003);
    linaccess::setJetSqlObjectResolver(db,[&](std::string_view collection,std::string_view object,std::string_view field) {
        assert(collection=="forms" || collection=="reports");
        assert(object=="filter");
        if (field=="year") return year;
        if (field=="flag") return linaccess::VbaValue::logical(true);
        if (field=="null") return linaccess::VbaValue{};
        throw linaccess::VbaRuntimeError(2465,"Missing test field");
    });
    model.execute(db,"SELECT [Formulare]![filter]![year].Value, Forms!filter!flag, Berichte!filter!null");
    assert(model.data(model.index(0,0)).toInt()==2003);
    assert(model.data(model.index(0,1)).toInt()==-1);
    assert(model.data(model.index(0,2)).isNull());
    year=linaccess::VbaValue::string("002004");
    model.execute(db,"SELECT Forms!filter!year & '-x'");
    assert(model.data(model.index(0,0)).toString()=="002004-x");
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('query','live_values',?)");
    query.addBindValue(QByteArray("SELECT Formulare!filter!year AS year")); assert(query.exec()); query.finish();
    assert(linaccess::installAccessQueryViews(db)==1);
    model.execute(db,"SELECT year FROM live_values"); assert(model.data(model.index(0,0)).toString()=="002004");
    year=linaccess::VbaValue::numeric(2005);
    model.execute(db,"SELECT year FROM live_values"); assert(model.data(model.index(0,0)).toInt()==2005);
    model.execute(db,"PARAMETERS item Value; SELECT [item]",{{"item",42}});
    assert(model.data(model.index(0,0)).toInt()==42 && model.data(model.index(0,0)).metaType().id()!=QMetaType::QString);
    model.execute(db,"PARAMETERS item Value; SELECT [item]",{{"item","0042"}});
    assert(model.data(model.index(0,0)).toString()=="0042");
    model.execute(db,"PARAMETERS item Value; SELECT [item]",{{"item",true}});
    assert(model.data(model.index(0,0)).toInt()==-1);
    model.execute(db,"PARAMETERS item Value; SELECT [item]",{{"item",QVariant{}}});
    assert(model.data(model.index(0,0)).isNull());
    model.execute(db,"PARAMETERS item Value; SELECT [item]",{{"item",QByteArray("binary")}});
    assert(model.data(model.index(0,0)).toByteArray()==QByteArray("binary"));
    bool invalid=false;
    try { model.execute(db,"PARAMETERS item Value; SELECT [item]",{{"item",QStringList{"unsupported"}}}); }
    catch(const linaccess::VbaRuntimeError& error) { invalid=error.number==13; }
    assert(invalid);
    query.clear(); db.close(); assert(db.open());
    year=linaccess::VbaValue::numeric(2006);
    model.execute(db,"SELECT Forms!filter!year"); assert(model.data(model.index(0,0)).toInt()==2006);
    auto separate=QSqlDatabase::addDatabase("QSQLITE","separate-objects"); separate.setDatabaseName(":memory:"); assert(separate.open());
    missing=false;
    try { model.execute(separate,"SELECT Forms!filter!year"); }
    catch(const linaccess::VbaRuntimeError& error) { missing=error.number==3061; }
    assert(missing);
    missing=false;
    try { model.execute(db,"SELECT Forms!filter!missing"); }
    catch(const linaccess::VbaRuntimeError& error) { missing=error.number==2465; }
    assert(missing);
}
