#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_query.hpp"
#include "access_table_import.hpp"
#include "dao_database.hpp"
#include "vba_currency.hpp"
#include <QApplication>
#include <QLineEdit>
#include <QSqlQuery>
#include <QTest>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

int main(int argc,char** argv) {
    QApplication app(argc,argv); assert(argc==2);
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","currency-query-types"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)")); assert(sql.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    const auto schema=linaccess::parseMdbSchema("CREATE TABLE `cash`\n(\n`amount` REAL\n);\n-- LINACCESS_SOURCE_TYPE 616d6f756e74 CURRENCY\n");
    linaccess::importAccessTable(db,"synthetic-query-types","cash",schema,linaccess::parseCsvDetailed("amount\n0.1\n"));
    if(QString::fromLatin1(argv[1])=="objects") {
        assert(sql.exec("CREATE TABLE object_audit(amount REAL)"));
        const QByteArray layout="Begin Form\nRecordSource =\"cash\"\nBegin Section\nBegin TextBox\nName =\"amount\"\nControlSource =\"amount\"\nWidth =4000\nEnd\nBegin TextBox\nName =\"free\"\nTop =500\nWidth =4000\nEnd\nEnd\nEnd\n";
        const QByteArray code=R"VBA(Sub CaptureReferences()
CurrentDb.Execute "INSERT INTO object_audit SELECT Forms!currency_types!amount"
CurrentDb.Execute "INSERT INTO object_audit SELECT Formulare!currency_types!amount"
CurrentDb.Execute "INSERT INTO object_audit SELECT Forms![currency_types]![amount].Value"
End Sub
Sub CaptureReports()
CurrentDb.Execute "INSERT INTO object_audit SELECT Reports!currency_report!amount"
CurrentDb.Execute "INSERT INTO object_audit SELECT Berichte!currency_report!amount"
End Sub
Sub CaptureFree()
CurrentDb.Execute "INSERT INTO object_audit SELECT Forms!currency_types!free"
End Sub
Function NumericReference()
Dim r As DAO.Recordset
Set r=CurrentDb.OpenRecordset("SELECT Forms!currency_types!amount<1 AS small")
NumericReference=r.Fields("small").Value
r.Close
End Function
)VBA";
        AccessFormPreview form("currency_types",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
        auto model=makePreviewFormObject(form); std::vector<VbaValue> none;
        assert(form.executeProcedure("CaptureReferences",none).status==linaccess::VbaExecutionStatus::Completed);
        const auto numeric=form.executeProcedure("NumericReference",none); assert(numeric.status==linaccess::VbaExecutionStatus::Completed && numeric.typedReturnValue->number==1);
        auto* editor=form.findChild<QLineEdit*>("access-control-amount"); assert(editor); editor->setFocus(); editor->selectAll(); QTest::keyClicks(editor,"0.6");
        assert(form.executeProcedure("CaptureReferences",none).status==linaccess::VbaExecutionStatus::Completed && editor->text()=="0.6" && model->item("amount").currencyUnits==1000 && form.hasUnsavedChanges());
        model->assignItem("free",VbaValue::currency(5497558138879999LL)); assert(form.executeProcedure("CaptureFree",none).status==linaccess::VbaExecutionStatus::Completed);
        model->assignItem("free",VbaValue::currency(5497558138880000LL)); const auto rejected=form.executeProcedure("CaptureFree",none); assert(rejected.status!=linaccess::VbaExecutionStatus::Completed && rejected.errorNumber==3251);
        const QByteArray reportLayout="Begin Report\nRecordSource =\"SELECT CCur('0.25') AS amount\"\nBegin Section\nBegin TextBox\nName =\"amount\"\nControlSource =\"amount\"\nEnd\nEnd\nEnd\n";
        AccessFormPreview report("currency_report",reportLayout,{},db); report.show(); QApplication::processEvents();
        assert(makePreviewFormObject(report)->item("amount").kind==VbaValue::Kind::Currency);
        assert(form.executeProcedure("CaptureReports",none).status==linaccess::VbaExecutionStatus::Completed);
        assert(sql.exec("SELECT amount FROM object_audit ORDER BY rowid"));
        for(int i=0;i<9;++i) assert(sql.next() && linaccess::currencyFromRealStorage(sql.value(0).toDouble()).currencyUnits==(i<6 ? 1000 : i==6 ? 5497558138879999LL : 2500));
        assert(!sql.next()); sql.finish(); assert(sql.exec("SELECT amount FROM cash") && sql.next() && sql.value(0).toDouble()==0.1); sql.finish();
        model->member("Undo",{}); assert(!form.hasUnsavedChanges() && report.close() && form.close()); return 0;
    }
    assert(QString::fromLatin1(argv[1])=="results");
    auto database=linaccess::makeDaoDatabase(db);
    const auto open=[&](const std::string& source) { return database->member("OpenRecordset",{VbaValue::string(source)}).object; };
    for(const auto* expression:{"CCur('549755813887.9999')","(CCur('549755813887.9999'))","+CCur('549755813887.9999')","-CCur('-549755813887.9999')"}) {
        auto rows=open(std::string("SELECT ")+expression+" AS amount"); auto field=rows->member("Fields",{VbaValue::string("amount")}).object;
        assert(field->member("Type",{}).number==5 && field->member("Value",{}).kind==VbaValue::Kind::Currency && field->member("Value",{}).currencyUnits==5497558138879999LL);
        rows->member("MoveNext",{}); assert(rows->member("EOF",{}).truthy() && field->member("Type",{}).number==5); rows->member("Close",{});
    }
    for(const auto* source:{"SELECT TOP 1 CCur('1.2345') AS amount FROM cash ORDER BY amount","SELECT DISTINCT CCur('1.2345') AS amount FROM cash","SELECT CCur('1.2345') AS amount WHERE 0"}) {
        auto rows=open(source); auto field=rows->member("Fields",{VbaValue::string("amount")}).object;
        assert(field->member("Type",{}).number==5);
        if(!rows->member("EOF",{}).truthy()) assert(field->member("Value",{}).currencyUnits==12345);
        rows->member("Close",{});
    }
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('query','money_projection',?)"); sql.addBindValue(QByteArray("SELECT CCur('1.2345') AS amount")); assert(sql.exec()); sql.finish();
    auto saved=open("money_projection"); assert(saved->member("Fields",{VbaValue::string("amount")}).object->member("Value",{}).currencyUnits==12345); saved->member("Close",{});
    auto definition=database->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string("PARAMETERS p Money; SELECT [p] AS amount")}).object;
    auto parameter=definition->member("Parameters",{VbaValue::string("p")}).object; parameter->assign("Value",VbaValue::currency(12345));
    auto parameterRows=definition->member("OpenRecordset",{}).object; auto parameterField=parameterRows->member("Fields",{VbaValue::string("amount")}).object;
    assert(parameterField->member("Type",{}).number==5 && parameterField->member("Value",{}).currencyUnits==12345); parameterRows->member("Close",{});
    parameter->assign("Value",{}); parameterRows=definition->member("OpenRecordset",{}).object;
    assert(parameterRows->member("Fields",{VbaValue::string("amount")}).object->member("Value",{}).kind==VbaValue::Kind::Null); parameterRows->member("Close",{});
    const auto base=linaccess::compileAccessQuery("SELECT CCur('1.2345') AS amount,2 AS plain");
    const auto filtered=linaccess::filterAccessQueryPlan(db,base,"amount>1");
    const auto result=linaccess::executeAccessQueryPlan(db,filtered,{});
    assert(result.rows.size()==1 && result.types.value("amount")=="CURRENCY_REAL" && !result.types.contains("plain"));
    const auto empty=linaccess::executeAccessQueryPlan(db,linaccess::filterAccessQueryPlan(db,base,"amount<0"),{}); assert(empty.rows.empty() && empty.types.value("amount")=="CURRENCY_REAL");
    for(const auto* source:{"SELECT CCur('1.2') AS amount UNION ALL SELECT 2 AS amount","SELECT *,CCur('1.2') AS extra FROM cash","SELECT CCur('1.2') AS amount,2 AS amount"}) {
        const auto plan=linaccess::compileAccessQuery(source); const auto unknown=linaccess::executeAccessQueryPlan(db,plan,{}); assert(unknown.types.isEmpty());
    }
    linaccess::AccessQueryModel table; table.execute(db,"SELECT CCur('549755813887.9999') AS amount");
    assert(table.data(table.index(0,0),Qt::DisplayRole).toString()=="549755813887.9999" && linaccess::currencyFromRealStorage(table.data(table.index(0,0),Qt::EditRole).toDouble()).currencyUnits==5497558138879999LL);
    const QByteArray layout="Begin Form\nRecordSource =\"SELECT CCur('549755813887.9999') AS amount\"\nBegin Section\nBegin TextBox\nName =\"amount\"\nControlSource =\"amount\"\nWidth =5000\nEnd\nEnd\nEnd\n";
    AccessFormPreview form("currency_projection",layout,{},db); form.show(); QApplication::processEvents(); auto model=makePreviewFormObject(form);
    auto* editor=form.findChild<QLineEdit*>("access-control-amount"); assert(editor && editor->isReadOnly() && editor->text()=="549755813887.9999" && model->item("amount").kind==VbaValue::Kind::Currency && model->item("amount").currencyUnits==5497558138879999LL);
    model->assign("Filter",VbaValue::string("amount>1")); model->assign("FilterOn",VbaValue::logical(true)); assert(model->item("amount").currencyUnits==5497558138879999LL && !form.hasUnsavedChanges()); assert(form.close());
}
