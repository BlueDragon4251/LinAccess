#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include <QLineEdit>
#include <QSqlQuery>
#include <QTest>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

int main(int argc,char** argv) {
    QApplication app(argc,argv);
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","form-visibility"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,label TEXT)")); assert(sql.exec("INSERT INTO records VALUES(1,'keep')"));
    assert(sql.exec("CREATE TABLE visible_events(name TEXT,event TEXT)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="records"
OnActivate ="[Event Procedure]"
OnDeactivate ="[Event Procedure]"
OnUnload ="[Event Procedure]"
OnClose ="[Event Procedure]"
Begin Section
Begin TextBox
Name ="label"
ControlSource ="label"
Width =4000
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub HideSelf()
Me.Visible=False
End Sub
Sub ShowSelf()
Me.Visible=True
End Sub
Sub SaveDefault()
DoCmd.RunCommand acCmdSaveRecord
End Sub
Sub CloseDefault()
DoCmd.Close
End Sub
Sub SendDefault()
SendKeys "X", True
End Sub
Function VisibleState() As Boolean
VisibleState=Me.Visible
End Function
Function ActiveName() As String
ActiveName=Screen.ActiveForm.Name
End Function
Function OpenCount() As Long
OpenCount=Forms.Count
End Function
Sub Form_Activate()
CurrentDb.Execute "INSERT INTO visible_events VALUES('" & Me.Name & "','activate')"
End Sub
Sub Form_Deactivate()
CurrentDb.Execute "INSERT INTO visible_events VALUES('" & Me.Name & "','deactivate')"
If Me.Tag="close" Then DoCmd.Close acForm,Me.Name
End Sub
Sub Form_Unload(Cancel As Integer)
CurrentDb.Execute "INSERT INTO visible_events VALUES('" & Me.Name & "','unload')"
If Me.Tag="blockclose" Then Cancel=True
End Sub
Sub Form_Close()
CurrentDb.Execute "INSERT INTO visible_events VALUES('" & Me.Name & "','close')"
End Sub
)VBA";
    std::vector<VbaValue> none;
    const auto eventCount=[&](const QString& name,const QString& event) {
        QSqlQuery read(db); read.prepare("SELECT COUNT(*) FROM visible_events WHERE name=? AND event=?"); read.addBindValue(name); read.addBindValue(event); assert(read.exec() && read.next()); return read.value(0).toInt();
    };
    AccessFormPreview a("visible_a",layout,code,db),b("visible_b",layout,code,db);
    auto first=makePreviewFormObject(a),second=makePreviewFormObject(b);
    int error=0; try { second->assign("Visible",VbaValue::logical(false)); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==3251);
    a.show(); QApplication::setActiveWindow(&a); QApplication::processEvents();
    b.show(); QApplication::setActiveWindow(&b); QApplication::processEvents();
    assert(second->member("Visible",{}).kind==VbaValue::Kind::Boolean && second->member("Visible",{}).boolean);
    auto state=b.executeProcedure("VisibleState",none); assert(state.status==linaccess::VbaExecutionStatus::Completed && state.typedReturnValue->boolean);
    assert(b.executeProcedure("OpenCount",none).typedReturnValue->number==2);
    auto* editor=b.findChild<QLineEdit*>("access-control-label"); assert(editor); editor->setFocus(); editor->selectAll(); QTest::keyClicks(editor,"draft");
    const auto deactivated=eventCount("visible_b","deactivate");
    auto result=b.executeProcedure("HideSelf",none); assert(result.status!=linaccess::VbaExecutionStatus::Completed && result.errorNumber==3251);
    assert(b.isVisible() && editor->text()=="draft" && second->item("label").asText()=="keep" && b.hasUnsavedChanges());
    assert(eventCount("visible_b","deactivate")==deactivated);
    second->member("Undo",{});
    int finished=0; QObject::connect(&b,&QDialog::finished,[&] { ++finished; });
    assert(b.executeProcedure("HideSelf",none).status==linaccess::VbaExecutionStatus::Completed); QApplication::processEvents();
    assert(b.isHidden() && !second->member("Visible",{}).boolean && finished==0 && eventCount("visible_b","unload")==0 && eventCount("visible_b","close")==0);
    assert(eventCount("visible_b","deactivate")==deactivated+1);
    const auto active=a.executeProcedure("ActiveName",none); assert(active.status!=linaccess::VbaExecutionStatus::Completed || active.typedReturnValue->asText()!="visible_b");
    assert(a.executeProcedure("OpenCount",none).typedReturnValue->number==2);
    error=0; try { second->assign("Visible",{}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==94 && b.isHidden());
    first->assign("Visible",VbaValue::logical(false)); QApplication::processEvents(); assert(a.isHidden() && b.isHidden());
    const auto saveHidden=b.executeProcedure("SaveDefault",none); assert(saveHidden.status!=linaccess::VbaExecutionStatus::Completed && saveHidden.errorNumber==2046);
    const auto sendHidden=b.executeProcedure("SendDefault",none); assert(sendHidden.status!=linaccess::VbaExecutionStatus::Completed && sendHidden.errorNumber==2110);
    const auto closeHidden=b.executeProcedure("CloseDefault",none); assert(closeHidden.status!=linaccess::VbaExecutionStatus::Completed && closeHidden.errorNumber==2046);
    assert(a.executeProcedure("OpenCount",none).typedReturnValue->number==2);
    first->assign("Visible",VbaValue::logical(true)); QApplication::processEvents();
    second->assignItem("label",VbaValue::string("program draft")); assert(b.hasUnsavedChanges());
    class Flag final:public linaccess::VbaObject {
    public:
        int reads{};
        VbaValue scalarValue() override { ++reads; return VbaValue::logical(true); }
        VbaValue member(std::string_view,const std::vector<VbaValue>&) override { throw linaccess::VbaRuntimeError(438,"No member"); }
    };
    const auto flag=std::make_shared<Flag>(); second->assign("Visible",VbaValue::reference(flag)); QApplication::processEvents();
    assert(flag->reads==1 && b.isVisible() && b.hasUnsavedChanges() && second->item("label").asText()=="program draft");
    assert(sql.exec("SELECT label FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="keep"); sql.finish(); second->member("Undo",{});
    assert(b.executeProcedure("HideSelf",none).status==linaccess::VbaExecutionStatus::Completed); QApplication::processEvents();
    second->assign("Tag",VbaValue::string("blockclose")); assert(!AccessFormPreview::closeProjectDocuments(db));
    assert(a.isVisible() && b.isHidden() && finished==0 && a.executeProcedure("OpenCount",none).typedReturnValue->number==2);
    second->assign("Tag",VbaValue::string("")); assert(AccessFormPreview::closeProjectDocuments(db));
    assert(!a.isVisible() && !b.isVisible() && finished==1 && eventCount("visible_a","close")==1 && eventCount("visible_b","close")==1);
    AccessFormPreview c("visible_c",layout,code,db); c.show(); QApplication::setActiveWindow(&c); QApplication::processEvents();
    auto third=makePreviewFormObject(c); third->assign("Tag",VbaValue::string("close"));
    third->assign("Visible",VbaValue::logical(false)); QApplication::processEvents(); assert(!c.isVisible() && eventCount("visible_c","close")==1);
    AccessFormPreview modal("visible_modal",layout,code,db); modal.setModal(true); modal.show(); QApplication::setActiveWindow(&modal); QApplication::processEvents();
    error=0; try { makePreviewFormObject(modal)->assign("Visible",VbaValue::logical(false)); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==3251 && modal.isVisible()); assert(modal.close());
    assert(sql.exec("SELECT label FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="keep"); sql.finish();
}
