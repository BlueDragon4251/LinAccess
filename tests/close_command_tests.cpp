#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include <QComboBox>
#include <QAbstractItemView>
#include <QSqlQuery>
#include <QTest>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

void testCloseCommands() {
    auto db=QSqlDatabase::addDatabase("QSQLITE","native-close-commands"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db); assert(query.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(query.exec("INSERT INTO records VALUES(1,'original')"));
    assert(query.exec("CREATE TABLE events(name TEXT)"));
    const QByteArray definition=R"FORM(Begin Form
RecordSource ="records"
OnUnload ="[Event Procedure]"
OnClose ="[Event Procedure]"
Begin Section
Begin TextBox
Name ="value"
ControlSource ="value"
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub Form_Unload(Cancel As Integer)
If Me.Tag="stay" Then
Cancel=1
End If
If Me.Tag="recursive" Then
DoCmd.Close
End If
End Sub
Sub Form_Close()
CurrentDb.Execute "INSERT INTO events VALUES('close')"
End Sub
)VBA";
    const QByteArray ownerCode=R"VBA(Sub CloseForm(name As String)
DoCmd.Close acForm,name,acSaveNo
End Sub
Sub CloseReport(name As String)
Application.DoCmd.Close acReport,name,acSavePrompt
End Sub
Sub CloseActive()
DoCmd.Close
End Sub
)VBA";
    AccessFormPreview owner("controller","Begin Form\nEnd\n",ownerCode,db);
    auto commands=makePreviewApplicationObject(owner)->member("DoCmd",{}).object;
    assert(commands==makePreviewApplicationObject(owner)->member("DoCmd",{}).object);
    AccessFormPreview child("target",definition,code,db);
    auto childModel=makePreviewFormObject(child);
    childModel->assign("Tag",linaccess::VbaValue::string("stay"));
    childModel->assignItem("value",linaccess::VbaValue::string("saved despite SaveNo"));
    std::vector<linaccess::VbaValue> args{linaccess::VbaValue::string("TARGET")};
    auto result=owner.executeProcedure("CloseForm",args);
    assert(result.status==linaccess::VbaExecutionStatus::Error && result.errorNumber==2501);
    assert(childModel->member("Name",{}).asText()=="target");
    assert(query.exec("SELECT value FROM records") && query.next() && query.value(0).toString()=="saved despite SaveNo"); query.finish();
    assert(query.exec("SELECT COUNT(*) FROM events") && query.next() && query.value(0).toInt()==0); query.finish();
    childModel->assign("Tag",linaccess::VbaValue::string("recursive"));
    result=owner.executeProcedure("CloseForm",args);
    assert(result.status==linaccess::VbaExecutionStatus::Error && result.errorNumber==2501);
    childModel->assign("Tag",linaccess::VbaValue::string(""));
    assert(owner.executeProcedure("CloseForm",args).status==linaccess::VbaExecutionStatus::Completed);
    bool invalid=false;
    try { (void)childModel->member("Name",{}); } catch(const linaccess::VbaRuntimeError& error) { invalid=error.number==2450; }
    assert(invalid);
    assert(query.exec("SELECT COUNT(*) FROM events") && query.next() && query.value(0).toInt()==1); query.finish();
    assert(owner.executeProcedure("CloseForm",args).errorNumber==2450);
    auto reportDefinition=definition; reportDefinition.replace("Begin Form","Begin Report");
    AccessFormPreview report("target",reportDefinition,{},db);
    assert(owner.executeProcedure("CloseForm",args).errorNumber==2450);
    assert(owner.executeProcedure("CloseReport",args).status==linaccess::VbaExecutionStatus::Completed);
    assert(owner.executeProcedure("CloseReport",args).errorNumber==2451);
    for (const auto& [arguments,error]:std::vector<std::pair<std::vector<linaccess::VbaValue>,int>>{
        {{linaccess::VbaValue{}},94},{{linaccess::VbaValue::numeric(0)},3251},
        {{linaccess::VbaValue::numeric(4294967298.)},6},
        {{linaccess::VbaValue::numeric(2),linaccess::VbaValue::string("controller"),linaccess::VbaValue::numeric(99)},5}}) {
        try { (void)commands->member("Close",arguments); assert(false); }
        catch(const linaccess::VbaRuntimeError& failure) { assert(failure.number==error); }
    }
    std::vector<linaccess::VbaValue> none;
    AccessFormPreview active("active","Begin Form\nEnd\n",{},db);
    active.show(); active.activateWindow(); QApplication::processEvents();
    assert(owner.executeProcedure("CloseActive",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(makePreviewFormsObject(owner)->member("Count",{}).asNumber()==1);
    assert(owner.close());
}

void testComboDropdown() {
    auto db=QSqlDatabase::addDatabase("QSQLITE","native-dropdown"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE choices(id INTEGER PRIMARY KEY,label TEXT)"));
    assert(query.exec("INSERT INTO choices VALUES(1,'first'),(2,'second')"));
    const QByteArray definition=R"FORM(Begin Form
Begin Section
Begin ComboBox
Name ="selection"
RowSourceType ="Table/Query"
RowSource ="SELECT id,label FROM choices ORDER BY id"
ColumnCount =2
End
Begin TextBox
Name ="plain"
End
End
End
)FORM";
    const QByteArray code="Sub OpenList()\nselection.Dropdown\nEnd Sub\nSub Invalid()\nplain.Dropdown\nEnd Sub\n";
    AccessFormPreview form("dropdown",definition,code,db); std::vector<linaccess::VbaValue> none;
    assert(form.executeProcedure("OpenList",none).errorNumber==2110);
    form.show(); QApplication::processEvents();
    assert(form.executeProcedure("OpenList",none).status==linaccess::VbaExecutionStatus::Completed);
    auto* combo=form.findChild<QComboBox*>("access-control-selection"); assert(combo && combo->view()->isVisible());
    assert(combo->count()==2);
    const auto position=combo->view()->visualRect(combo->model()->index(1,0)).center();
    QTest::mouseClick(combo->view()->viewport(),Qt::LeftButton,Qt::NoModifier,position);
    assert(makePreviewFormObject(form)->item("selection").asNumber()==2);
    combo->hidePopup(); combo->setEnabled(false);
    assert(form.executeProcedure("OpenList",none).errorNumber==2110);
    assert(form.executeProcedure("Invalid",none).errorNumber==438);
    assert(form.close());
}
