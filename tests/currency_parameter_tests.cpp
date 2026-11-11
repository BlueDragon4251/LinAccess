#include "access_query.hpp"
#include "access_table_import.hpp"
#include "dao_database.hpp"
#include "vba_currency.hpp"
#include "vba_qt_value.hpp"
#include <QCoreApplication>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <optional>
#include <functional>
#include <limits>

class CurrencyScalar final:public linaccess::VbaObject {
public:
    int reads{};
    linaccess::VbaValue scalarValue() override { ++reads; return linaccess::VbaValue::currency(5000); }
    linaccess::VbaValue member(std::string_view,const std::vector<linaccess::VbaValue>&) override { throw linaccess::VbaRuntimeError(438,"No explicit member"); }
};
class MutatingCurrencyScalar final:public linaccess::VbaObject {
public:
    std::function<void()> callback;int reads{};
    linaccess::VbaValue scalarValue() override {++reads;callback();return linaccess::VbaValue::currency(2000);}
    linaccess::VbaValue member(std::string_view,const std::vector<linaccess::VbaValue>&) override {throw linaccess::VbaRuntimeError(438,"No explicit member");}
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
    int errorCheck=0;const auto expect=[&](auto action,int wanted) { ++errorCheck;int actual=0; try { action(); } catch(const linaccess::VbaRuntimeError& error) { actual=error.number; } if(actual!=wanted) qFatal("Currency parameter error check %d: actual=%d expected=%d",errorCheck,actual,wanted); };
    assert(parameter->member("Type",{}).number==5 && parameter->member("Type",{}).numericType==VbaValue::NumericType::Integer);
    assert(parameter->member("Value",{}).kind==VbaValue::Kind::Null);
    expect([&] { definition->member("Execute",{}); },3061);
    expect([&] { parameter->assign("Type",VbaValue::numeric(7)); },383);
    parameter->assign("Value",VbaValue::string("1.23455"));
    assert(parameter->member("Value",{}).kind==VbaValue::Kind::Currency && parameter->scalarValue().currencyUnits==12346);
    definition->member("Execute",{});
    const auto probe=std::make_shared<CurrencyScalar>(); parameter->assign("Value",VbaValue::reference(probe));
    assert(probe->reads==1 && parameter->member("Value",{}).currencyUnits==5000);
    for(const auto units:{5497558138880000LL,std::numeric_limits<long long>::max(),std::numeric_limits<long long>::min()}) {
        parameter->assign("Value",VbaValue::currency(units));assert(parameter->member("Value",{}).kind==VbaValue::Kind::Currency && parameter->scalarValue().currencyUnits==units);
        expect([&] {definition->member("Execute",{});},3251);expect([&] {parameter->assign("Value",VbaValue::errorValue(11));},13);assert(parameter->scalarValue().currencyUnits==units);
    }
    parameter->assign("Value",VbaValue::currency(5000));assert(parameter->member("Value",{}).currencyUnits==5000);
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
        const auto result=linaccess::executeAccessQueryPlan(db,plan,bindings); assert(result.rows.size()==1 && result.rows[0][0].toInt()==-1); // Access Boolean True.
        expect([&] { (void)linaccess::bindAccessQueryParameters(plan,{{"p",QString("549755813888")}}); },3251);
        expect([&] { (void)linaccess::bindAccessQueryParameters(plan,{{"p",QString("invalid")}}); },13);
        expect([&] { (void)linaccess::bindAccessQueryParameters(plan,{}); },3061);
    }
    // Full signed64 values pass through table-free projections without REAL conversion.
    const QString projection="PARAMETERS p Currency,q Money; SELECT p AS amount,q AS other";
    const auto exact=linaccess::makeDaoDatabase(db)->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string(projection.toStdString())}).object;
    auto exactP=exact->member("Parameters",{VbaValue::string("p")}).object;
    auto exactQ=exact->member("Parameters",{VbaValue::string("q")}).object;
    const auto maximum=std::numeric_limits<std::int64_t>::max(),minimum=std::numeric_limits<std::int64_t>::min();
    exactP->assign("Value",VbaValue::currency(maximum));exactQ->assign("Value",VbaValue::currency(minimum));
    auto exactRows=exact->member("OpenRecordset",{}).object;
    auto exactField=exactRows->member("Fields",{VbaValue::string("amount")}).object;
    assert(exactField->member("Value",{}).kind==VbaValue::Kind::Currency && exactField->scalarValue().currencyUnits==maximum);
    assert(exactField->member("Type",{}).number==5 && exactField->member("Size",{}).number==8);
    assert(exactRows->member("Fields",{VbaValue::string("other")}).object->scalarValue().currencyUnits==minimum);
    expect([&] {exactRows->member("FindFirst",{VbaValue::string("amount > 0")});},3251);
    assert(exactField->scalarValue().currencyUnits==maximum);
    exactRows->member("FindFirst",{VbaValue::string("amount > CDec('922337203685477.5806') AND other < CCur(0)")});assert(!exactRows->member("NoMatch",{}).truthy() && exactField->scalarValue().currencyUnits==maximum);
    exactRows->assign("Filter",VbaValue::string("amount > CDec('922337203685477.5806') AND other < CCur(0)"));exactRows->assign("Sort",VbaValue::string("amount DESC,other ASC"));auto exactSnapshot=exactRows->member("OpenRecordset",{VbaValue::integer(4)}).object;assert(exactSnapshot->member("RecordCount",{}).asNumber()==1 && exactSnapshot->item("amount").currencyUnits==maximum && exactSnapshot->item("other").currencyUnits==minimum && !exactSnapshot->member("Updatable",{}).truthy());exactSnapshot->member("FindLast",{VbaValue::string("CStr(amount)='922337203685477.5807'")});assert(!exactSnapshot->member("NoMatch",{}).truthy());exactSnapshot->member("Close",{});
    exactRows->member("Close",{});
    exactP->assign("Value",{});exactRows=exact->member("OpenRecordset",{}).object;
    assert(exactRows->member("Fields",{VbaValue::string("amount")}).object->scalarValue().kind==VbaValue::Kind::Null);exactRows->member("Close",{});
    const QMap<QString,QVariant> exactValues{{"p",QString("922337203685477.5807")},{"q",QString("-922337203685477.5808")}};
    const auto exactPlan=linaccess::compileAccessQuery(projection);
    const auto exactBindings=linaccess::bindAccessQueryParameters(exactPlan,exactValues);
    assert(exactBindings.value(exactPlan.parameters[0].binding).metaType().id()==QMetaType::LongLong);
    assert(exactBindings.value(exactPlan.parameters[0].binding).toLongLong()==maximum);
    const auto exactResult=linaccess::executeAccessQueryPlan(db,exactPlan,exactBindings);
    assert(exactResult.rows[0][1].toLongLong()==minimum && exactResult.types.value("amount")=="CURRENCY_TICKS");
    linaccess::AccessQueryModel exactModel;exactModel.execute(db,projection,exactValues);
    assert(exactModel.data(exactModel.index(0,0)).toString()=="922337203685477.5807");
    assert(exactModel.clipboardText(exactModel.index(0,1))=="-922337203685477.5808");
    expect([&] {(void)linaccess::filterAccessQueryPlan(db,exactPlan,"amount > 1");},3251);
    linaccess::AccessQueryModel sameName; sameName.execute(db,"PARAMETERS p Currency; SELECT p AS p",{{"p",exactValues.value("p")}});
    assert(sameName.clipboardText(sameName.index(0,0))=="922337203685477.5807");
    auto derived=linaccess::filterAccessQueryPlan(db,exactPlan,"amount>other AND amount>=CDec('922337203685477.5807') AND CStr(CVar(amount))='922337203685477.5807'");derived=linaccess::sortAccessQueryPlan(db,derived,"other ASC,amount DESC");const auto derivedResult=linaccess::executeAccessQueryPlan(db,derived,linaccess::bindAccessQueryParameters(derived,exactValues));assert(derivedResult.rows.size()==1 && derivedResult.rows[0][0].toLongLong()==maximum && derivedResult.rows[0][1].toLongLong()==minimum);
    auto nested=linaccess::filterAccessQueryPlan(db,derived,"amount IN(CDec('922337203685477.5807'))");assert(linaccess::executeAccessQueryPlan(db,nested,linaccess::bindAccessQueryParameters(nested,exactValues)).rows.size()==1);auto tampered=derived;tampered.sql+=" LIMIT 1";expect([&]{(void)linaccess::bindAccessQueryParameters(tampered,exactValues);},3251);auto changedRoot=derived;changedRoot.source="PARAMETERS p Currency,q Money; SELECT q AS amount,p AS other";expect([&]{(void)linaccess::bindAccessQueryParameters(changedRoot,exactValues);},3251);
    auto changedPlan=exactPlan;changedPlan.sql+=" WHERE 1";
    expect([&] {(void)linaccess::bindAccessQueryParameters(changedPlan,exactValues);},3251);
    for(const auto* unsafe:{"SELECT p+0 AS amount","SELECT p AS amount WHERE 1","SELECT p AS amount FROM cash","SELECT p AS same,p AS same"}) {
        const auto bounded=linaccess::compileAccessQuery(QString("PARAMETERS p Currency; ")+unsafe);
        assert(!bounded.parameters[0].currencyTicks);
        expect([&] {(void)linaccess::bindAccessQueryParameters(bounded,{{"p",exactValues.value("p")}});},3251);
    }
    assert(sql.exec("CREATE TABLE currency_tick_order(amount CURRENCY_TICKS)"));for(const auto units:{minimum,maximum,maximum-1}) {assert(sql.prepare("INSERT INTO currency_tick_order VALUES(?)"));sql.addBindValue(QVariant::fromValue(static_cast<qlonglong>(units)));assert(sql.exec());sql.finish();}
    const auto typed=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT amount,CVar(amount) AS retained,CCur(amount) AS converted,CStr(amount) AS rendered,VarType(amount) AS subtype FROM currency_tick_order WHERE amount>=CCur(0) ORDER BY amount"),{});assert(typed.rows.size()==2 && typed.columnTypes==QVector<QString>({"CURRENCY_TICKS","CURRENCY_TICKS","CURRENCY_TICKS","TEXT","INTEGER2"}) && typed.rows[0][0].toLongLong()==maximum-1 && typed.rows[1][1].toLongLong()==maximum && typed.rows[1][2].toLongLong()==maximum && typed.rows[0][3].toString()=="922337203685477.5806" && typed.rows[1][4].toInt()==6);
    for(const auto* unsupported:{"SELECT CAST(amount AS REAL) FROM currency_tick_order","SELECT Avg(amount) FROM currency_tick_order"}) expect([&]{(void)linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery(unsupported),{});},3251);
    assert(sql.exec("INSERT INTO currency_tick_order VALUES(NULL)"));sql.finish();
    const auto fieldMath=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT amount+CCur(0) AS added,amount-CCur(0) AS subtracted,amount*CCur(1) AS multiplied,-amount AS negative FROM currency_tick_order WHERE amount>=CCur(0) ORDER BY amount"),{});assert(fieldMath.rows.size()==2 && fieldMath.columnTypes==QVector<QString>(4,"CURRENCY_TICKS") && fieldMath.rows[0][0].toLongLong()==maximum-1 && fieldMath.rows[1][1].toLongLong()==maximum && fieldMath.rows[1][2].toLongLong()==maximum && fieldMath.rows[1][3].toLongLong()==-maximum);
    const auto selectedMoney=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT Min(amount) AS lo,Max(DISTINCT amount) AS hi,First(amount) AS initial,Last(amount) AS final FROM currency_tick_order"),{});assert(selectedMoney.columnTypes==QVector<QString>(4,"CURRENCY_TICKS") && selectedMoney.rows.size()==1 && selectedMoney.rows[0][0].toLongLong()==minimum && selectedMoney.rows[0][1].toLongLong()==maximum && selectedMoney.rows[0][2].toLongLong()==minimum && selectedMoney.rows[0][3].isNull());
    const auto emptyMoney=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT Min(amount),Max(amount),First(amount),Last(amount) FROM currency_tick_order WHERE amount IS NULL"),{});assert(emptyMoney.rows.size()==1 && emptyMoney.columnTypes==QVector<QString>(4,"CURRENCY_TICKS"));for(const auto& cell:emptyMoney.rows[0]) assert(cell.isNull());
    const auto moneySelectorTypes=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT VarType(Min(amount)) AS subtype,CStr(Max(amount)) AS rendered FROM currency_tick_order"),{});assert(moneySelectorTypes.rows.size()==1 && moneySelectorTypes.rows[0][0].toInt()==6 && moneySelectorTypes.rows[0][1].toString()=="922337203685477.5807");
    const auto groupedMoney=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT amount IS NULL AS missing,Min(amount) AS lo,Max(amount) AS hi FROM currency_tick_order GROUP BY amount IS NULL ORDER BY missing DESC"),{});assert(groupedMoney.rows.size()==2 && groupedMoney.rows[0][1].toLongLong()==minimum && groupedMoney.rows[0][2].toLongLong()==maximum && groupedMoney.rows[1][1].isNull());
    assert(sql.exec("CREATE TABLE currency_integral(amount CURRENCY_TICKS,qty INTEGER4,byte_value INTEGER1,small_value INTEGER2,flag BOOL)"));assert(sql.exec("INSERT INTO currency_integral VALUES(10001,2,255,-2,1),(NULL,NULL,NULL,NULL,NULL)"));sql.finish();
    const auto integralMoney=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT amount+qty AS added,qty-amount AS subtracted,amount*qty AS multiplied,byte_value+amount AS byte_added,small_value*amount AS short_product,amount+flag AS logical_added FROM currency_integral"),{});assert(integralMoney.columnTypes==QVector<QString>(6,"CURRENCY_TICKS") && integralMoney.rows.size()==2);const QVector<qint64> integralExpected{30001,9999,20002,2560001,-20002,1};for(int col=0;col<6;++col) assert(integralMoney.rows[0][col].toLongLong()==integralExpected[col] && integralMoney.rows[1][col].isNull());
    const auto integralPredicates=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT amount<qty AS smaller,qty>amount AS larger,amount BETWEEN CLng(1) AND CInt(2) AS inside,amount IN(CLng(1),CCur('1.0001')) AS included FROM currency_integral WHERE qty=2"),{});assert(integralPredicates.rows.size()==1);for(const auto& value:integralPredicates.rows[0]) assert(value.toInt()==-1);
    expect([&]{(void)linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT amount*2 FROM currency_integral"),{});},3251);
    linaccess::makeDaoDatabase(db)->member("Execute",{VbaValue::string("UPDATE currency_integral SET amount=amount*qty WHERE qty=2")});assert(sql.exec("SELECT amount FROM currency_integral WHERE qty=2") && sql.next() && sql.value(0).toLongLong()==20002);sql.finish();
    assert(sql.exec("INSERT INTO currency_integral VALUES(9223372036854775807,2,0,0,0)"));sql.finish();expect([&]{linaccess::makeDaoDatabase(db)->member("Execute",{VbaValue::string("UPDATE currency_integral SET amount=amount*qty WHERE qty=2")});},6);assert(sql.exec("SELECT amount FROM currency_integral WHERE rowid=1") && sql.next() && sql.value(0).toLongLong()==20002);sql.finish();
    assert(sql.exec("CREATE TABLE currency_branches(id INTEGER4,exact CURRENCY_TICKS,legacy CURRENCY_REAL,selector CURRENCY_TICKS)"));assert(sql.exec("INSERT INTO currency_branches VALUES(1,9223372036854775807,1.25,10000),(2,NULL,2.5,20000)"));sql.finish();
    const auto moneyBranches=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT Nz(exact,legacy) AS replaced,IIf(id=1,exact,legacy) AS conditional,Choose(id,exact,legacy) AS chosen,Switch(id=1,exact,True,legacy) AS switched FROM currency_branches ORDER BY id"),{});assert(moneyBranches.columnTypes==QVector<QString>(4,"CURRENCY_TICKS") && moneyBranches.rows.size()==2);for(int col=0;col<4;++col) assert(moneyBranches.rows[0][col].toLongLong()==maximum && moneyBranches.rows[1][col].toLongLong()==25000);
    const auto moneyControls=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT Choose(selector,legacy,CCur('3.75')) AS chosen,IIf(selector,legacy,CCur('3.75')) AS conditional,Switch(selector,legacy,True,CCur('3.75')) AS switched,Round(legacy,selector) AS rounded FROM currency_branches ORDER BY id"),{});assert(moneyControls.columnTypes==QVector<QString>(4,"CURRENCY_REAL") && moneyControls.rows.size()==2 && moneyControls.rows[0][0].toDouble()==1.25 && moneyControls.rows[1][0].toDouble()==3.75 && moneyControls.rows[0][1].toDouble()==1.25 && moneyControls.rows[1][1].toDouble()==2.5 && moneyControls.rows[0][2].toDouble()==1.25 && moneyControls.rows[1][2].toDouble()==2.5 && moneyControls.rows[0][3].toDouble()==1.2 && moneyControls.rows[1][3].toDouble()==2.5);
    const auto moneyBranchText=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT CStr(Nz(exact,legacy)) AS text_value FROM currency_branches ORDER BY id"),{});assert(moneyBranchText.rows[0][0].toString()=="922337203685477.5807" && moneyBranchText.rows[1][0].toString()=="2.5");
    linaccess::makeDaoDatabase(db)->member("Execute",{VbaValue::string("UPDATE currency_branches SET exact=Nz(exact,legacy) WHERE id=2")});assert(sql.exec("SELECT exact FROM currency_branches WHERE id=2") && sql.next() && sql.value(0).toLongLong()==25000);sql.finish();
    const auto fullSum=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT Sum(amount) AS total FROM currency_tick_order"),{});assert(fullSum.rows.size()==1 && fullSum.columnTypes==QVector<QString>({"CURRENCY_TICKS"}) && fullSum.rows[0][0].toLongLong()==maximum-2);
    const auto nullSum=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT Sum(amount) AS total FROM currency_tick_order WHERE amount IS NULL"),{});assert(nullSum.rows.size()==1 && nullSum.rows[0][0].isNull() && nullSum.columnTypes[0]=="CURRENCY_TICKS");
    for(const auto* overflowing:{"SELECT amount+CCur('0.0001') FROM currency_tick_order WHERE amount=CDec('922337203685477.5807')","SELECT -amount FROM currency_tick_order WHERE amount<CCur(0)","SELECT Sum(amount) FROM currency_tick_order WHERE amount>=CCur(0)"}) expect([&]{(void)linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery(overflowing),{});},6);
    const auto arithmeticFilter=linaccess::filterAccessQueryPlan(db,exactPlan,"amount+other < CCur(0)");assert(linaccess::executeAccessQueryPlan(db,arithmeticFilter,linaccess::bindAccessQueryParameters(arithmeticFilter,exactValues)).rows.size()==1);
    assert(sql.exec("CREATE TABLE full_currency_target(id INTEGER4,amount CURRENCY_TICKS DEFAULT 12500,note TEXT)"));assert(sql.exec("CREATE TABLE legacy_currency_target(amount CURRENCY_REAL)"));sql.finish();const auto database=linaccess::makeDaoDatabase(db);const auto execute=[&](const char* command) {database->member("Execute",{VbaValue::string(command)});};
    execute("INSERT INTO full_currency_target(id,amount,note) VALUES(1,CDec('922337203685477.5807'),'maximum')");execute("INSERT INTO full_currency_target(id,amount) VALUES(2,'-922337203685477.5808')");
    execute("INSERT INTO full_currency_target(id,amount,note) SELECT id+10,amount,note FROM full_currency_target WHERE id IN(1,2)");execute("UPDATE full_currency_target SET amount=amount-CCur('0.0001') WHERE id=1");
    assert(sql.exec("SELECT amount FROM full_currency_target WHERE id=1") && sql.next() && sql.value(0).toLongLong()==maximum-1);sql.finish();assert(sql.exec("SELECT amount FROM full_currency_target WHERE id=12") && sql.next() && sql.value(0).toLongLong()==minimum);sql.finish();
    execute("INSERT INTO full_currency_target DEFAULT VALUES");assert(sql.exec("SELECT amount FROM full_currency_target WHERE id IS NULL") && sql.next() && sql.value(0).toLongLong()==12500);sql.finish();
    execute("UPDATE full_currency_target SET amount='1.2500' WHERE id=2");execute("INSERT INTO legacy_currency_target(amount) SELECT amount FROM full_currency_target WHERE id=2");assert(sql.exec("SELECT amount FROM legacy_currency_target") && sql.next() && sql.value(0).toDouble()==1.25);sql.finish();
    execute("INSERT INTO full_currency_target(id,amount) VALUES(3,CDec('1.23455'))");assert(sql.exec("SELECT amount FROM full_currency_target WHERE id=3") && sql.next() && sql.value(0).toLongLong()==12346);sql.finish();
    assert(sql.exec("CREATE TABLE currency_binary_input(id INTEGER4,payload BLOB)"));assert(sql.prepare("INSERT INTO currency_binary_input VALUES(4,?)"));sql.addBindValue(QByteArray(8,0));assert(sql.exec());sql.finish();
    expect([&]{execute("INSERT INTO full_currency_target(id,amount) SELECT id,payload FROM currency_binary_input");},13);expect([&]{execute("INSERT INTO full_currency_target(id,amount) VALUES(4,CDec('922337203685477.5808'))");},6);
    expect([&]{execute("UPDATE full_currency_target SET amount=IIf(id=11,'not-money','2.50') WHERE id IN(1,11)");},13);assert(sql.exec("SELECT amount FROM full_currency_target WHERE id=1") && sql.next() && sql.value(0).toLongLong()==maximum-1);sql.finish();
    expect([&]{execute("INSERT INTO legacy_currency_target(amount) SELECT amount FROM full_currency_target WHERE id=11");},3251);assert(sql.exec("SELECT COUNT(*) FROM legacy_currency_target") && sql.next() && sql.value(0).toInt()==1);sql.finish();
    expect([&]{execute("INSERT INTO full_currency_target(id,amount) SELECT 99,t.amount FROM full_currency_target AS t INNER JOIN currency_binary_input AS b ON t.id=b.id");},3251);expect([&]{execute("INSERT INTO full_currency_target DEFAULT VALUES RETURNING amount");},3251);
    execute("UPDATE full_currency_target SET amount=NULL WHERE id=11");assert(sql.exec("SELECT amount FROM full_currency_target WHERE id=11") && sql.next() && sql.value(0).isNull());sql.finish();assert(sql.exec("SELECT COUNT(*) FROM full_currency_target") && sql.next() && sql.value(0).toInt()==6);sql.finish();
    // Exact parameter-only arithmetic never passes through the REAL money adapter.
    const QString arithmetic="PARAMETERS p Currency,q Money; SELECT ([p]+[q]) AS amount,[p]-[q] AS difference,[p]*[q] AS product";
    const auto arithmeticPlan=linaccess::compileAccessQuery(arithmetic);assert(arithmeticPlan.resultTypes==QVector<QString>({"CURRENCY_TICKS","CURRENCY_TICKS","CURRENCY_TICKS"}));
    const QMap<QString,QVariant> arithmeticValues{{"p",QString("922337203685477.5806")},{"q",QString("0.0001")}};
    auto calculated=linaccess::executeAccessQueryPlan(db,arithmeticPlan,linaccess::bindAccessQueryParameters(arithmeticPlan,arithmeticValues));
    assert(calculated.rows.size()==1 && calculated.rows[0][0].toLongLong()==maximum && calculated.rows[0][1].toLongLong()==maximum-2 && calculated.rows[0][2].toLongLong()==922337203685478LL);
    linaccess::AccessQueryModel arithmeticModel;arithmeticModel.execute(db,arithmetic,arithmeticValues);assert(arithmeticModel.data(arithmeticModel.index(0,0)).toString()=="922337203685477.5807" && arithmeticModel.clipboardText(arithmeticModel.index(0,2))=="92233720368.5478");
    const auto arithmeticDefinition=linaccess::makeDaoDatabase(db)->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string(arithmetic.toStdString())}).object;
    arithmeticDefinition->member("Parameters",{VbaValue::string("p")}).object->assign("Value",VbaValue::currency(maximum-1));arithmeticDefinition->member("Parameters",{VbaValue::string("q")}).object->assign("Value",VbaValue::currency(1));
    auto arithmeticRows=arithmeticDefinition->member("OpenRecordset",{}).object;auto arithmeticField=arithmeticRows->member("Fields",{VbaValue::string("amount")}).object;assert(arithmeticField->scalarValue().kind==VbaValue::Kind::Currency && arithmeticField->scalarValue().currencyUnits==maximum && arithmeticField->member("Type",{}).number==5 && arithmeticField->member("Size",{}).number==8);arithmeticRows->member("MoveNext",{});assert(arithmeticRows->member("EOF",{}).truthy() && arithmeticField->member("Type",{}).number==5);arithmeticRows->member("Close",{});arithmeticDefinition->member("Close",{});
    const auto exactRun=[&](const QString& expression,const QMap<QString,QVariant>& values){const auto plan=linaccess::compileAccessQuery("PARAMETERS p Currency,q Currency; SELECT "+expression);assert(plan.parameters[0].currencyTicks && plan.parameters[1].currencyTicks);return linaccess::executeAccessQueryPlan(db,plan,linaccess::bindAccessQueryParameters(plan,values));};
    calculated=exactRun("p+q AS value",exactValues);assert(calculated.rows[0][0].toLongLong()==-1);
    calculated=exactRun("p+q*q AS precedence,(p+q)*q AS grouped,-p AS negative,+q AS positive",{{"p",QString("1")},{"q",QString("2")}});assert(calculated.rows[0][0].toLongLong()==50000 && calculated.rows[0][1].toLongLong()==60000 && calculated.rows[0][2].toLongLong()==-10000 && calculated.rows[0][3].toLongLong()==20000);
    for(const auto& item:std::vector<std::pair<QString,std::int64_t>>{{"0.0001",0},{"0.0003",2},{"-0.0001",0},{"-0.0003",-2}}){calculated=exactRun("p*q AS value",{{"p",item.first},{"q",QString("0.5")}});assert(calculated.rows[0][0].toLongLong()==item.second);}
    calculated=linaccess::executeAccessQueryPlan(db,arithmeticPlan,linaccess::bindAccessQueryParameters(arithmeticPlan,{{"p",arithmeticValues.value("p")},{"q",QVariant{}}}));assert(calculated.rows[0][0].isNull() && calculated.rows[0][1].isNull() && calculated.rows[0][2].isNull() && calculated.types.value("product")=="CURRENCY_TICKS");
    expect([&]{(void)exactRun("p+q AS value",{{"p",exactValues.value("p")},{"q",QString("0.0001")}});},6);
    expect([&]{(void)exactRun("p-q AS value",{{"p",exactValues.value("q")},{"q",QString("0.0001")}});},6);
    expect([&]{(void)exactRun("p*q AS value",{{"p",exactValues.value("p")},{"q",QString("2")}});},6);
    expect([&]{(void)exactRun("-p AS value",{{"p",exactValues.value("q")},{"q",QString("0")}});},6);
    auto changedArithmetic=arithmeticPlan;changedArithmetic.source="PARAMETERS p Currency,q Money; SELECT ([p]-[q]) AS amount,[p]-[q] AS difference,[p]*[q] AS product";expect([&]{(void)linaccess::bindAccessQueryParameters(changedArithmetic,arithmeticValues);},3251);
    expect([&]{(void)linaccess::filterAccessQueryPlan(db,arithmeticPlan,"amount > 1");},3251);
    exact->member("Close",{});
    const auto like=linaccess::compileAccessQuery("PARAMETERS p Currency; SELECT amount FROM cash WHERE amount Like p");
    expect([&] { (void)linaccess::bindAccessQueryParameters(like,{{"p",0.5}}); },3251);
    expect([&] { (void)linaccess::compileAccessQuery("PARAMETERS p Currency(8); SELECT p"); },3085);
    definition->assign("SQL",VbaValue::string("PARAMETERS p Money; SELECT p"));
    expect([&] { (void)parameter->member("Value",{}); },3420);
    auto fresh=definition->member("Parameters",{VbaValue::string("p")}).object; fresh->assign("Value",VbaValue::currency(1000));
    assert(fresh->scalarValue().kind==VbaValue::Kind::Currency && fresh->scalarValue().currencyUnits==1000);
    auto mutating=std::make_shared<MutatingCurrencyScalar>();mutating->callback=[definition] {definition->assign("SQL",VbaValue::string("PARAMETERS p Currency; SELECT p AS changed"));};expect([&] {fresh->assign("Value",VbaValue::reference(mutating));},3420);assert(mutating->reads==1);fresh=definition->member("Parameters",{VbaValue::string("p")}).object;assert(fresh->scalarValue().kind==VbaValue::Kind::Null);fresh->assign("Value",VbaValue::currency(1000));assert(fresh->scalarValue().currencyUnits==1000);
    definition->member("Close",{}); expect([&] { (void)fresh->member("Type",{}); },3420);
    assert(sql.exec("SELECT COUNT(*) FROM cash") && sql.next() && sql.value(0).toInt()==3); sql.finish();
}
