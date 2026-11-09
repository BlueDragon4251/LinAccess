#include "access_format.hpp"
#include "vba_decimal.hpp"
#include "access_query.hpp"
#include <QCoreApplication>
#include <QLocale>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <limits>
int main(int argc,char** argv) {
    QCoreApplication app(argc,argv);
    using namespace linaccess;
    const auto format=[](VbaValue value,const char* mask,bool strict=true) { return formatAccessValue(value,mask,strict); };
    const auto date=VbaValue::string("2003-12-31 18:05:06");
    assert(format(date,"yyyy").text=="2003" && format(date,"mm").text=="12" && format(date,"dd").text=="31");
    assert(format(VbaValue::string("2004-02-29"),"dd").text=="29");
    assert(format(VbaValue::string("2024-03-31 02:30:00"),"dd").text=="31");
    assert(format(VbaValue::numeric(0),"yyyy").text=="1899");
    assert(format(VbaValue::numeric(2),"yyyy").text=="1900");
    assert(format(VbaValue::numeric(-0.75),"dd").text=="30");
    assert(format(VbaValue::numeric(-1.75),"dd").text=="29");
    assert(format(VbaValue::empty(),"yyyy").text=="1899");
    assert(format(VbaValue{},"yyyy",false).kind==VbaValue::Kind::Null);
    const auto rejects=[&](VbaValue value,const char* mask,int expected) {
        bool rejected=false;
        try { (void)format(value,mask); } catch(const VbaRuntimeError& error) { rejected=error.number==expected; }
        assert(rejected);
    };
    rejects({},"yyyy",94);
    rejects(VbaValue::string("2003-02-29"),"yyyy",13);
    rejects(VbaValue::string("2003-12-31 25:00:00"),"yyyy",13);
    rejects(VbaValue::string("2003-12-31 suffix"),"yyyy",13);
    rejects(VbaValue::numeric(std::numeric_limits<double>::infinity()),"yyyy",6);
    rejects(VbaValue::numeric(1e100),"yyyy",6);
    rejects(VbaValue::numeric(1),"0.00E+00",5);
    const auto original=QLocale();
    QLocale::setDefault(QLocale(QLocale::German,QLocale::Germany));
    assert(format(date,"Long Date").text=="Mittwoch, 31. Dezember 2003");
    assert(format(date,"yyyy-mm-dd hh:nn:ss").asText()=="2003-12-31 18:05:06");assert(format(date,"hh:mm:ss").asText()=="18:05:06");assert(format(date,"yyyy \"Jahr\" q").asText()=="2003 Jahr 4");
    assert(format(decimalLiteral("1.245"),"0.00").asText()=="1,24");assert(format(decimalLiteral("0.25"),"0.##%").asText()=="25%");assert(format(decimalLiteral("1.2"),"0.00##").asText()=="1,20");assert(format(decimalLiteral("79228162514264337593543950335"),"0").asText()=="79228162514264337593543950335");
    assert(format(decimalLiteral("-1234.5"),"#,##0.00;(#,##0.00);\"zero\"").asText()=="(1.234,50)");
    assert(format(decimalLiteral("0"),"0.00;(0.00);\"zero\"").asText()=="zero");
    assert(format(decimalLiteral("-1.5"),"0.00;;\"zero\"").asText()=="-1,50");
    assert(format(decimalLiteral("12.5"),"\"EUR \"0.00\" net\"").asText()=="EUR 12,50 net");
    assert(format(decimalLiteral("12.5"),"0.00\\x").asText()=="12,50x");
    assert(format(decimalLiteral("1234500"),"0,.00\" k\"").asText()=="1234,50 k");
    assert(format(decimalLiteral("1234500000"),"0,,.00\" m\"").asText()=="1234,50 m");
    assert(format(decimalLiteral("-0.125"),"0.0%;(0.0%)").asText()=="(12,5%)");
    assert(format(decimalLiteral("1.2"),"\"0; \"0.00").asText()=="0; 1,20");
    assert(format(decimalLiteral("0.5"),".00").asText()==",50");
    assert(format(decimalLiteral("-12.5"),"\"EUR \"0.00").asText()=="-EUR 12,50");
    rejects(decimalLiteral("1"),"0.00;0.00;0.00;0.00",5);
    rejects(decimalLiteral("1"),"0.00[Red]",5);
    rejects(decimalLiteral("1e-28"),"0,.00",3251);
    const auto options=parseVbaNumericFormatOptions({decimalLiteral("-0.5"),VbaValue::integer(2),VbaValue::integer(0),VbaValue::integer(-1),VbaValue::integer(0)});assert(formatAccessNumericValue(decimalLiteral("-0.5"),false,options).asText()=="(,50)");

    auto db=QSqlDatabase::addDatabase("QSQLITE"); db.setDatabaseName(":memory:"); assert(db.open());
    AccessQueryModel model;
    model.execute(db,"SELECT Format$(#12/31/2003#,\"yyyy\"),Format$(#12/31/2003#,\"mm\"),Format(#12/31/2003#,\"dd\")");
    assert(model.data(model.index(0,0)).toString()=="2003" && model.data(model.index(0,1)).toString()=="12" && model.data(model.index(0,2)).toString()=="31");
    assert(model.data(model.index(0,1)).metaType().id()==QMetaType::QString);
    model.execute(db,"SELECT FormatNumber(9007199254740993,0,-1,0,0),FormatCurrency(1.25,2),Format(1.245,\"0.00\")");assert(model.data(model.index(0,0)).toString()=="9007199254740993" && model.data(model.index(0,2)).toString()=="1,24");

    model.execute(db,"SELECT Format(-1234.5,'#,##0.00;(#,##0.00)'),Format(1234500,'0,.00'),Format(0,'0.00;(0.00);\"zero\"')");
    assert(model.data(model.index(0,0)).toString()=="(1.234,50)" && model.data(model.index(0,1)).toString()=="1234,50" && model.data(model.index(0,2)).toString()=="zero");
    model.execute(db,"SELECT Format$(#12/31/2003#,\"Long Date\"),Format(Null,\"yyyy\")");
    assert(model.data(model.index(0,0)).toString()=="Mittwoch, 31. Dezember 2003");
    assert(model.data(model.index(0,1)).isNull());
    bool nullRejected=false;
    try { model.execute(db,"SELECT Format$(Null,\"yyyy\")"); } catch(const VbaRuntimeError& error) { nullRejected=error.number==94; }
    assert(nullRejected && model.rowCount()==0);
    QLocale::setDefault(QLocale(QLocale::English,QLocale::UnitedStates));
    model.execute(db,"SELECT Format$(#12/31/2003#,\"Long Date\")");
    assert(model.data(model.index(0,0)).toString()=="Wednesday, December 31, 2003");

    auto dependencies=QSqlDatabase::addDatabase("QSQLITE","format-dependencies");
    dependencies.setDatabaseName(":memory:"); assert(dependencies.open());
    QSqlQuery definition(dependencies);
    assert(definition.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    definition.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('query','child',?)");
    definition.addBindValue(QByteArray("SELECT Format$(#12/31/2003#,\"yyyy\") AS year")); assert(definition.exec()); definition.finish();
    definition.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('query','parent',?)");
    definition.addBindValue(QByteArray("SELECT year FROM child")); assert(definition.exec()); definition.finish();
    assert(installAccessQueryViews(dependencies)==2);
    model.execute(dependencies,"SELECT year FROM parent");
    assert(model.rowCount()==1 && model.data(model.index(0,0)).toString()=="2003");
    QLocale::setDefault(original);
}
