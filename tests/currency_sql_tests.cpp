#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_query.hpp"
#include "dao_database.hpp"
#include "access_table_import.hpp"
#include "vba_currency.hpp"
#include <QApplication>
#include <QLineEdit>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>

int main(int argc,char** argv) {
    QApplication app(argc,argv);
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","currency-sql"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    const auto schema=linaccess::parseMdbSchema("CREATE TABLE `cash`\n(\n`id` INTEGER,\n`amount` REAL\n);\n-- LINACCESS_SOURCE_TYPE 6964 INT\n-- LINACCESS_SOURCE_TYPE 616d6f756e74 CURRENCY\n");
    linaccess::importAccessTable(db,"synthetic-currency-sql","cash",schema,linaccess::parseCsvDetailed("id,amount\n1,1.25\n2,2.5\n3,\x1e\n"));
    assert(sql.exec("CREATE TABLE binary_input(payload BLOB)")); assert(sql.exec("INSERT INTO binary_input VALUES(X'00')"));
    const auto run=[&](const QString& source) { const auto plan=linaccess::compileAccessQuery(source); return linaccess::executeAccessQueryPlan(db,plan,linaccess::bindAccessQueryParameters(plan,{})); };
    for(const auto& [expression,units]:std::vector<std::pair<QString,std::int64_t>>{{"CCur('1.23455')",12346},{"CCur('1.23445')",12344},{"CCur('-1.23455')",-12346},{"CCur('0.00005')",0},{"CCur(True)",-10000},{"CCur(False)",0},{"CCur(0.1+0.2)",3000},{"CCur('549755813887.9999')",5497558138879999LL}}) {
        const auto result=run("SELECT "+expression+" AS v"); assert(result.rows.size()==1 && linaccess::currencyFromRealStorage(result.rows[0][0].toDouble()).currencyUnits==units);
    }
    const auto literal=linaccess::compileAccessQuery("SELECT 'CCur(' AS literal,CCur('1.23455') AS v");
    assert(literal.sql.contains("'CCur('") && literal.sql.contains("linaccess_jet_ccur"));
    const auto aliased=run("SELECT CCur('1.23455') AS converted,converted AS duplicate FROM cash WHERE id=1");
    assert(aliased.rows.size()==1 && aliased.rows[0][0].toDouble()==1.2346 && aliased.rows[0][1].toDouble()==1.2346);
    for(const auto& [expression,wanted]:std::vector<std::pair<QString,int>>{{"CCur()",450},{"CCur(1,2)",450},{"CCur(Null)",94},{"CCur('invalid')",13},{"CCur('549755813888')",3251},{"CCur('922337203685477.5808')",6}}) {
        int error=0; try { (void)run("SELECT "+expression); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==wanted);
    }
    int binaryError=0; try { (void)run("SELECT CCur(payload) FROM binary_input"); } catch(const linaccess::VbaRuntimeError& failure) { binaryError=failure.number; } assert(binaryError==13);
    int error=0; try {
        const auto update=linaccess::makeDaoDatabase(db)->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string("UPDATE cash SET amount=CCur(IIf(id=1,'3.33335','invalid')) WHERE id IN(1,2)")}).object;
        update->member("Execute",{});
    } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } if(error!=13) std::cerr<<"Synthetic mixed Currency UPDATE error "<<error<<'\n'; assert(error==13);
    assert(sql.exec("SELECT amount FROM cash ORDER BY id"));
    assert(sql.next() && sql.value(0).toDouble()==1.25); assert(sql.next() && sql.value(0).toDouble()==2.5); assert(sql.next() && sql.value(0).isNull()); sql.finish();
    const QByteArray layout="Begin Form\nRecordSource =\"cash\"\nBegin Section\nBegin TextBox\nName =\"amount\"\nControlSource =\"amount\"\nWidth =5000\nEnd\nEnd\nEnd\n";
    const QByteArray code=R"VBA(Function ConvertSQL(ByVal text As String)
Dim q As DAO.QueryDef
Set q=CurrentDb.CreateQueryDef("","PARAMETERS p Text(255); UPDATE cash SET amount=CCur(p) WHERE id=1")
q.Parameters("p").Value=text
q.Execute
Me.Requery
ConvertSQL=VarType(Me!amount)
End Function
)VBA";
    AccessFormPreview form("currency sql",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
    auto model=makePreviewFormObject(form); auto* editor=form.findChild<QLineEdit*>("access-control-amount"); assert(editor);
    std::vector<VbaValue> args{VbaValue::string("1.23455")};
    const auto converted=form.executeProcedure("ConvertSQL",args); assert(converted.status==linaccess::VbaExecutionStatus::Completed && converted.typedReturnValue->number==6 && model->item("amount").currencyUnits==12346 && !form.hasUnsavedChanges());
    args[0]=VbaValue::string("549755813887.9999");
    assert(form.executeProcedure("ConvertSQL",args).status==linaccess::VbaExecutionStatus::Completed && model->item("amount").currencyUnits==5497558138879999LL && editor->text()=="549755813887.9999");
    args[0]=VbaValue::string("invalid"); const auto rejected=form.executeProcedure("ConvertSQL",args);
    assert(rejected.status!=linaccess::VbaExecutionStatus::Completed && rejected.errorNumber==13 && model->item("amount").currencyUnits==5497558138879999LL && !form.hasUnsavedChanges());
    assert(sql.exec("SELECT amount FROM cash WHERE id=1") && sql.next() && linaccess::currencyFromRealStorage(sql.value(0).toDouble()).currencyUnits==5497558138879999LL); sql.finish();
    assert(form.close());
}
