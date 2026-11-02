#include "access_query.hpp"
#include "access_table_import.hpp"
#include "dao_database.hpp"
#include "vba_currency.hpp"
#include <QCoreApplication>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <optional>

class CurrencyScalar final:public linaccess::VbaObject {
public:
    int reads{};
    linaccess::VbaValue scalarValue() override { ++reads; return linaccess::VbaValue::currency(5000); }
    linaccess::VbaValue member(std::string_view,const std::vector<linaccess::VbaValue>&) override { throw linaccess::VbaRuntimeError(438,"No explicit member"); }
};
int main(int argc,char** argv) {
    QCoreApplication app(argc,argv);
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","currency-parameters"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    const auto schema=linaccess::parseMdbSchema("CREATE TABLE `cash`\n(\n`amount` REAL\n);\n-- LINACCESS_SOURCE_TYPE 616d6f756e74 CURRENCY\n");
    linaccess::importAccessTable(db,"synthetic-parameters","cash",schema,linaccess::parseCsvDetailed("amount\n"));
    const auto definition=linaccess::makeDaoDatabase(db)->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string("PARAMETERS p Currency; INSERT INTO cash VALUES(p)")}).object;
    auto parameter=definition->member("Parameters",{VbaValue::string("p")}).object;
    const auto expect=[](auto action,int wanted) { int actual=0; try { action(); } catch(const linaccess::VbaRuntimeError& error) { actual=error.number; } assert(actual==wanted); };
    assert(parameter->member("Type",{}).number==5 && parameter->member("Type",{}).numericType==VbaValue::NumericType::Integer);
    assert(parameter->member("Value",{}).kind==VbaValue::Kind::Null);
    expect([&] { definition->member("Execute",{}); },3061);
    expect([&] { parameter->assign("Type",VbaValue::numeric(7)); },383);
    parameter->assign("Value",VbaValue::string("1.23455"));
    assert(parameter->member("Value",{}).kind==VbaValue::Kind::Currency && parameter->scalarValue().currencyUnits==12346);
    definition->member("Execute",{});
    const auto probe=std::make_shared<CurrencyScalar>(); parameter->assign("Value",VbaValue::reference(probe));
    assert(probe->reads==1 && parameter->member("Value",{}).currencyUnits==5000);
    expect([&] { parameter->assign("Value",VbaValue::currency(5497558138880000LL)); },3251);
    expect([&] { parameter->assign("Value",VbaValue::errorValue(11)); },13);
    assert(parameter->member("Value",{}).currencyUnits==5000);
    parameter->assign("Value",VbaValue::currency(5497558138879999LL)); definition->member("Execute",{});
    parameter->assign("Value",{}); assert(parameter->member("Value",{}).kind==VbaValue::Kind::Null); definition->member("Execute",{});
    auto rows=linaccess::makeDaoDatabase(db)->member("OpenRecordset",{VbaValue::string("cash")}).object;
    auto field=rows->member("Fields",{VbaValue::string("amount")}).object;
    for(const auto expected:std::vector<std::optional<std::int64_t>>{12346,5497558138879999LL,{}}) {
        const auto value=field->member("Value",{}); assert(expected ? value.kind==VbaValue::Kind::Currency && value.currencyUnits==*expected : value.kind==VbaValue::Kind::Null); rows->member("MoveNext",{});
    }
    assert(rows->member("EOF",{}).truthy()); rows->member("Close",{});
    for(const auto* type:{"Currency","Money"}) {
        const auto plan=linaccess::compileAccessQuery(QString("PARAMETERS p ")+type+"; SELECT p < 1 AS small");
        assert(plan.parameters.size()==1 && plan.parameters[0].type=="currency");
        const auto bindings=linaccess::bindAccessQueryParameters(plan,{{"p",QString("0.5")}});
        assert(bindings.value(plan.parameters[0].binding).metaType().id()==QMetaType::Double && bindings.value(plan.parameters[0].binding).toDouble()==0.5);
        const auto result=linaccess::executeAccessQueryPlan(db,plan,bindings); assert(result.rows.size()==1 && result.rows[0][0].toInt()==1);
        expect([&] { (void)linaccess::bindAccessQueryParameters(plan,{{"p",QString("549755813888")}}); },3251);
        expect([&] { (void)linaccess::bindAccessQueryParameters(plan,{{"p",QString("invalid")}}); },13);
        expect([&] { (void)linaccess::bindAccessQueryParameters(plan,{}); },3061);
    }
    const auto like=linaccess::compileAccessQuery("PARAMETERS p Currency; SELECT amount FROM cash WHERE amount Like p");
    expect([&] { (void)linaccess::bindAccessQueryParameters(like,{{"p",0.5}}); },3251);
    expect([&] { (void)linaccess::compileAccessQuery("PARAMETERS p Currency(8); SELECT p"); },3085);
    definition->assign("SQL",VbaValue::string("PARAMETERS p Money; SELECT p"));
    expect([&] { (void)parameter->member("Value",{}); },3420);
    auto fresh=definition->member("Parameters",{VbaValue::string("p")}).object; fresh->assign("Value",VbaValue::currency(1000));
    assert(fresh->scalarValue().kind==VbaValue::Kind::Currency && fresh->scalarValue().currencyUnits==1000);
    definition->member("Close",{}); expect([&] { (void)fresh->member("Type",{}); },3420);
    assert(sql.exec("SELECT COUNT(*) FROM cash") && sql.next() && sql.value(0).toInt()==3); sql.finish();
}
