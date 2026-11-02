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
Sub WritePosition()
Me.CurrentRecord=1
End Sub
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
    model->assign("Filter",VbaValue::string("id>=20")); model->assign("FilterOn",VbaValue::logical(true)); position(2);
    assert(run("NewRow").status==linaccess::VbaExecutionStatus::Completed); position(3); assert(model->member("NewRecord",{}).truthy() && !form.hasUnsavedChanges());
    assert(sql.exec("SELECT COUNT(*) FROM records") && sql.next() && sql.value(0).toInt()==3); sql.finish();
    assert(sql.exec("SELECT COUNT(*),MIN(type),MAX(type) FROM positions") && sql.next() && sql.value(0).toInt()>=5 && sql.value(1).toInt()==3 && sql.value(2).toInt()==3); sql.finish();
    assert(form.close()); int closed=0; try { model->member("CurrentRecord",{}); } catch(const linaccess::VbaRuntimeError& error) { closed=error.number; } assert(closed==2450);
    const QByteArray report="Begin Report\nRecordSource =\"records\"\nEnd\n"; AccessFormPreview reportForm("position_report",report,{},db); auto reportModel=makePreviewFormObject(reportForm);
    int unsupported=0; try { reportModel->member("CurrentRecord",{}); } catch(const linaccess::VbaRuntimeError& error) { unsupported=error.number; } assert(unsupported==438); assert(reportForm.close());
}
