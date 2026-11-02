#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include <QLineEdit>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

int main(int argc,char** argv) {
    QApplication app(argc,argv); using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","current-record"); db.setDatabaseName(":memory:"); assert(db.open()); QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,label TEXT)")); assert(sql.exec("INSERT INTO records VALUES(10,'first'),(20,'second'),(30,'third')"));
    assert(sql.exec("CREATE TABLE positions(position INTEGER,type INTEGER)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="records"
OnCurrent ="[Event Procedure]"
Begin Section
Begin TextBox
Name ="label"
ControlSource ="label"
Width =4000
End
Begin TextBox
Name ="position"
ControlSource ="=CurrentRecord"
Top =500
Width =4000
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub Form_Current()
CurrentDb.Execute "INSERT INTO positions VALUES(" & CStr(Me.CurrentRecord) & "," & CStr(VarType(Me.CurrentRecord)) & ")"
End Sub
Function PositionValue()
PositionValue=CurrentRecord
End Function
Function ThroughCollections()
ThroughCollections=Forms("current_record").CurrentRecord+Screen.ActiveForm.CurrentRecord
End Function
Sub NextRow()
DoCmd.RunCommand 65
End Sub
Sub LastRow()
DoCmd.RunCommand 68
End Sub
Sub NewRow()
DoCmd.RunCommand 28
End Sub
Function SortRows()
OrderBy="[records].[id] DESC"
OrderByOn=True
SortRows=Me.CurrentRecord
End Function
Function SortAscending()
DoCmd.RunCommand acCmdSortAscending
SortAscending=Me.OrderBy
End Function
Function SortDescending()
DoCmd.RunCommand acCmdSortDescending
SortDescending=Me.OrderBy
End Function
Sub WritePosition()
Me.CurrentRecord=1
End Sub
Function RunSqlNull() As Long
On Error Resume Next
DoCmd.RunSQL SQLStatement:=Null
If Err.Number=94 Then
RunSqlNull=RunSqlNull+1
End If
Err.Clear
DoCmd.RunSQL Null
If Err.Number=94 Then
RunSqlNull=RunSqlNull+1
End If
End Function
Function RunSqlTransactions() As Long
On Error Resume Next
DoCmd.RunSQL SQLStatement:="INSERT INTO named_unique VALUES(2),(1)",UseTransaction:=True
RunSqlTransactions=Err.Number
On Error GoTo 0
DoCmd.RunSQL "INSERT INTO named_unique VALUES(3)",False
DBEngine(0).BeginTrans
DoCmd.RunSQL "INSERT INTO named_unique VALUES(4)"
On Error Resume Next
DoCmd.RunSQL SQLStatement:="INSERT INTO named_unique VALUES(5),(1)",UseTransaction:=True
On Error GoTo 0
DBEngine(0).Rollback
End Function
Function NamedCommands() As String
DoCmd.RunSQL SQLStatement:="INSERT INTO named_calls VALUES(1)",UseTransaction:=True
DoCmd.RunSQL "INSERT INTO named_calls VALUES(2)",False
DoCmd.OpenForm WhereCondition:="[id]=20", FormName:="named_child", OpenArgs:="payload", DataMode:=acFormPropertySettings, WindowMode:=acWindowNormal
NamedCommands=Forms("named_child").OpenArgs & ":" & Forms("named_child")!label
DoCmd.Close Save:=2,ObjectName:="named_child",ObjectType:=2
DoCmd.OpenForm FormName:="named_child",OpenArgs:="new"
NamedCommands=NamedCommands & "|" & Forms("named_child").OpenArgs
DoCmd.Close Save:=2
DoCmd.OpenReport OpenArgs:="report-value",ReportName:="named_report",View:=acViewPreview
NamedCommands=NamedCommands & "|" & Reports("named_report").OpenArgs
DoCmd.Close Save:=acSaveNo,ObjectName:="named_report",ObjectType:=acReport
End Function
)VBA";
    AccessFormPreview form("current_record",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents(); auto model=makePreviewFormObject(form);
    const auto position=[&](int wanted) { const auto value=model->member("CurrentRecord",{}); assert(value.kind==VbaValue::Kind::Number && value.numericType==VbaValue::NumericType::Long && value.asNumber()==wanted); auto* editor=form.findChild<QLineEdit*>("access-control-position"); assert(editor && editor->text()==QString::number(wanted)); };
    const auto run=[&](const char* name) { std::vector<VbaValue> args; return form.executeProcedure(name,args); };
    position(1); auto result=run("PositionValue"); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->kind==VbaValue::Kind::Number && result.typedReturnValue->numericType==VbaValue::NumericType::Long && result.typedReturnValue->asNumber()==1);
    result=run("ThroughCollections"); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==2);
    assert(run("NextRow").status==linaccess::VbaExecutionStatus::Completed); position(2); assert(model->item("id").asNumber()==20);
    model->assignItem("label",VbaValue::string("pending")); position(2); assert(form.hasUnsavedChanges());
    assert(sql.exec("SELECT label FROM records WHERE id=20") && sql.next() && sql.value(0).toString()=="second"); sql.finish();
    result=run("WritePosition"); assert(result.status==linaccess::VbaExecutionStatus::Error && result.errorNumber==438); position(2); assert(form.hasUnsavedChanges()); model->member("Undo",{});
    assert(run("LastRow").status==linaccess::VbaExecutionStatus::Completed); position(3);
    model->assign("OrderBy",VbaValue::string("[id] DESC")); assert(!model->member("OrderByOn",{}).truthy()); position(3);
    result=run("SortRows"); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==1); position(1); assert(model->item("id").asNumber()==30);
    int badOrder=0; try { model->assign("OrderBy",VbaValue::string("missing DESC")); } catch(const linaccess::VbaRuntimeError& error) { badOrder=error.number; }
    assert(badOrder==3251 && model->member("OrderBy",{}).asText()=="[records].[id] DESC" && model->member("OrderByOn",{}).truthy()); position(1);
    model->assign("OrderByOn",VbaValue::logical(false)); position(3);
    auto* activeField=form.findChild<QLineEdit*>("access-control-label"); activeField->setFocus(); QApplication::processEvents(); assert(QApplication::focusWidget()==activeField);
    result=run("SortAscending"); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->asText()=="[label] ASC"); position(3);
    result=run("SortDescending"); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->asText()=="[label] DESC"); position(1);
    model->assign("OrderByOn",VbaValue::logical(false)); position(3);
    model->assign("Filter",VbaValue::string("id>=20")); model->assign("FilterOn",VbaValue::logical(true)); position(2);
    assert(run("NewRow").status==linaccess::VbaExecutionStatus::Completed); position(3); assert(model->member("NewRecord",{}).truthy() && !form.hasUnsavedChanges());
    assert(sql.exec("SELECT COUNT(*) FROM records") && sql.next() && sql.value(0).toInt()==3); sql.finish();
    assert(sql.exec("SELECT COUNT(*),MIN(type),MAX(type) FROM positions") && sql.next() && sql.value(0).toInt()>=5 && sql.value(1).toInt()==3 && sql.value(2).toInt()==3); sql.finish();
    assert(form.close()); int closed=0; try { model->member("CurrentRecord",{}); } catch(const linaccess::VbaRuntimeError& error) { closed=error.number; } assert(closed==2450);
    const QByteArray sortedLayout="Begin Form\nRecordSource =\"records\"\nOrderBy =\"[id] DESC\"\nOrderByOn =-1\nEnd\n";
    AccessFormPreview sorted("sorted_on_open",sortedLayout,{},db); auto sortedModel=makePreviewFormObject(sorted);
    assert(sortedModel->member("OrderByOn",{}).truthy() && sortedModel->item("id").asNumber()==30 && sorted.close());
    const QByteArray queryLayout="Begin Form\nRecordSource =\"SELECT id,label FROM records ORDER BY id DESC\"\nEnd\n";
    AccessFormPreview queryForm("sorted_query",queryLayout,{},db); auto queryModel=makePreviewFormObject(queryForm);
    assert(queryModel->item("id").asNumber()==30); queryModel->assign("OrderBy",VbaValue::string("id ASC")); queryModel->assign("OrderByOn",VbaValue::logical(true));
    assert(queryModel->item("id").asNumber()==10 && queryModel->member("CurrentRecord",{}).asNumber()==1 && queryForm.close());
    assert(sql.exec("CREATE TABLE named_calls(value INTEGER)"));
    assert(sql.exec("CREATE TABLE named_unique(value INTEGER UNIQUE)")); assert(sql.exec("INSERT INTO named_unique VALUES(1)"));
    assert(sql.exec("CREATE TABLE open_arg_events(phase TEXT,value TEXT)"));
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('form','named_child',?)");
    auto childLayout=layout; childLayout.replace("OnCurrent =\"[Event Procedure]\"","OnOpen =\"[Event Procedure]\"\nOnLoad =\"[Event Procedure]\"");
    sql.addBindValue(childLayout); assert(sql.exec()); sql.finish();
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('vba','Form_named_child',?)");
    sql.addBindValue(QByteArray(R"VBA(Sub Form_Open(Cancel As Integer)
CurrentDb.Execute "INSERT INTO open_arg_events VALUES('Open','" & Me.OpenArgs & "')"
End Sub
Sub Form_Load()
CurrentDb.Execute "INSERT INTO open_arg_events VALUES('Load','" & OpenArgs & "')"
End Sub
)VBA")); assert(sql.exec()); sql.finish();
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('report','named_report',?)");
    sql.addBindValue(QByteArray("Begin Report\nRecordSource =\"records\"\nOnOpen =\"[Event Procedure]\"\nEnd\n")); assert(sql.exec()); sql.finish();
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('vba','Report_named_report',?)");
    sql.addBindValue(QByteArray(R"VBA(Sub Report_Open(Cancel As Integer)
CurrentDb.Execute "INSERT INTO open_arg_events VALUES('ReportOpen','" & Me.OpenArgs & "')"
End Sub
)VBA")); assert(sql.exec()); sql.finish();
    AccessFormPreview namedForm("named_parent",layout,code,db); namedForm.show(); namedForm.activateWindow(); QApplication::processEvents();
    std::vector<VbaValue> namedArgs; const auto named=namedForm.executeProcedure("NamedCommands",namedArgs);
    if(named.status!=linaccess::VbaExecutionStatus::Completed) qWarning("Named DoCmd error=%d line=%d",named.errorNumber,named.line);
    assert(named.status==linaccess::VbaExecutionStatus::Completed && named.typedReturnValue->asText()=="payload:second|new|report-value");
    assert(sql.exec("SELECT COUNT(*) FROM named_calls") && sql.next() && sql.value(0).toInt()==2); sql.finish();
    const auto nullSql=namedForm.executeProcedure("RunSqlNull",namedArgs);
    assert(nullSql.status==linaccess::VbaExecutionStatus::Completed && nullSql.typedReturnValue->asNumber()==2);
    const auto transactions=namedForm.executeProcedure("RunSqlTransactions",namedArgs);
    if(transactions.status!=linaccess::VbaExecutionStatus::Completed) qWarning("RunSQL transaction error=%d line=%d",transactions.errorNumber,transactions.line);
    assert(transactions.status==linaccess::VbaExecutionStatus::Completed && transactions.typedReturnValue->asNumber()==3022);
    assert(sql.exec("SELECT COUNT(*),SUM(value) FROM named_unique") && sql.next() && sql.value(0).toInt()==2 && sql.value(1).toInt()==4); sql.finish();
    assert(makePreviewFormObject(namedForm)->member("OpenArgs",{}).kind==VbaValue::Kind::Null);
    assert(sql.exec("SELECT COUNT(*) FROM open_arg_events WHERE value IN('payload','new')") && sql.next() && sql.value(0).toInt()==4); sql.finish();
    assert(sql.exec("SELECT COUNT(*) FROM open_arg_events WHERE phase='ReportOpen' AND value='report-value'") && sql.next() && sql.value(0).toInt()==1); sql.finish();
    int readOnly=0; try { makePreviewFormObject(namedForm)->assign("OpenArgs",VbaValue::string("changed")); } catch(const linaccess::VbaRuntimeError& failure) { readOnly=failure.number; } assert(readOnly==438);
    assert(namedForm.close());
    const QByteArray report="Begin Report\nRecordSource =\"records\"\nEnd\n"; AccessFormPreview reportForm("position_report",report,{},db); auto reportModel=makePreviewFormObject(reportForm);
    int unsupported=0; try { reportModel->member("CurrentRecord",{}); } catch(const linaccess::VbaRuntimeError& error) { unsupported=error.number; } assert(unsupported==438); assert(reportForm.close());
}
