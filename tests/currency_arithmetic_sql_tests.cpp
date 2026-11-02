#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_query.hpp"
#include "access_table_import.hpp"
#include "dao_database.hpp"
#include "vba_currency.hpp"
#include <QApplication>
#include <QLineEdit>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

int main(int argc,char** argv) {
    QApplication app(argc,argv);
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","currency-arithmetic-sql"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)")); assert(sql.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    const auto schema=linaccess::parseMdbSchema("CREATE TABLE `cash`\n(\n`id` INTEGER,\n`a` REAL,\n`b` REAL\n);\n-- LINACCESS_SOURCE_TYPE 6964 INT\n-- LINACCESS_SOURCE_TYPE 61 CURRENCY\n-- LINACCESS_SOURCE_TYPE 62 CURRENCY\n");
    linaccess::importAccessTable(db,"synthetic-money-arithmetic","cash",schema,linaccess::parseCsvDetailed("id,a,b\n1,0.1,0.2\n2,1.2345,0.0001\n3,\x1e,0.2\n"));
    const auto run=[&](const QString& source) { const auto plan=linaccess::compileAccessQuery(source); return linaccess::executeAccessQueryPlan(db,plan,linaccess::bindAccessQueryParameters(plan,{})); };
    auto result=run("SELECT a+b AS total,a-b AS difference,a*b AS product FROM cash WHERE id=1");
    for(int i=0;i<3;++i) assert(result.types.value(result.columns[i].toCaseFolded())=="CURRENCY_REAL");
    assert(linaccess::currencyFromRealStorage(result.rows[0][0].toDouble()).currencyUnits==3000 && linaccess::currencyFromRealStorage(result.rows[0][1].toDouble()).currencyUnits==-1000 && linaccess::currencyFromRealStorage(result.rows[0][2].toDouble()).currencyUnits==200);
    for(const auto& [expression,units]:std::vector<std::pair<QString,std::int64_t>>{{"a+b*b",1400},{"a-b-b",-3000},{"(a+b)*b",600},{"-(a+b)",-3000},{"CCur('1.2345')*CCur('0.1')",1234},{"CCur('1.2355')*CCur('0.1')",1236},{"CCur('-1.2345')*CCur('0.1')",-1234},{"CCur(a+b)",3000}}) {
        result=run("SELECT "+expression+" AS value FROM cash WHERE id=1"); assert(result.types.value("value")=="CURRENCY_REAL" && linaccess::currencyFromRealStorage(result.rows[0][0].toDouble()).currencyUnits==units);
    }
    result=run("SELECT a+b AS value FROM cash WHERE id=3"); assert(result.types.value("value")=="CURRENCY_REAL" && result.rows[0][0].isNull());
    result=run("SELECT a+0.2 AS value FROM cash WHERE id=1"); assert(result.types.isEmpty());
    result=run("SELECT Sum(a+b) AS value FROM cash"); assert(result.types.value("value")=="CURRENCY_REAL" && linaccess::currencyFromRealStorage(result.rows[0][0].toDouble()).currencyUnits==15346);
    auto definition=linaccess::makeDaoDatabase(db)->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string("PARAMETERS p Currency; SELECT a+p AS value FROM cash WHERE id=1")}).object;
    auto parameter=definition->member("Parameters",{VbaValue::string("p")}).object; parameter->assign("Value",VbaValue::currency(2000));
    auto rows=definition->member("OpenRecordset",{}).object; auto field=rows->member("Fields",{VbaValue::string("value")}).object;
    assert(field->member("Type",{}).number==5 && field->member("Value",{}).kind==VbaValue::Kind::Currency && field->member("Value",{}).currencyUnits==3000); rows->member("Close",{});
    parameter->assign("Value",{}); rows=definition->member("OpenRecordset",{}).object; assert(rows->member("Fields",{VbaValue::string("value")}).object->member("Value",{}).kind==VbaValue::Kind::Null); rows->member("Close",{});
    for(const auto& [expression,wanted]:std::vector<std::pair<QString,int>>{{"CCur('549755813887.9999')+CCur('0.0001')",3251},{"CCur('549755813887')*CCur('549755813887')",6}}) {
        int error=0; try { (void)run("SELECT "+expression); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==wanted);
    }
    const QByteArray layout="Begin Report\nRecordSource =\"SELECT a+b AS total FROM cash WHERE id=1\"\nBegin Section\nBegin TextBox\nName =\"total\"\nControlSource =\"total\"\nEnd\nEnd\nEnd\n";
    AccessFormPreview report("money_arithmetic",layout,{},db); report.show(); QApplication::processEvents(); auto model=makePreviewFormObject(report);
    auto* editor=report.findChild<QLineEdit*>("access-control-total"); assert(editor && editor->isReadOnly() && editor->text()=="0.3" && model->item("total").kind==VbaValue::Kind::Currency && model->item("total").currencyUnits==3000); assert(report.close());
    assert(sql.exec("SELECT a,b FROM cash WHERE id=1") && sql.next() && sql.value(0).toDouble()==0.1 && sql.value(1).toDouble()==0.2); sql.finish();
}
