#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include <QComboBox>
#include <QAbstractItemView>
#include <QSqlQuery>
#include <QTest>
#include <QTimer>
#include <QMessageBox>
#include <QGroupBox>
#include <QButtonGroup>
#include <QPushButton>
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

void testOpenFormCommands() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","native-openform-commands"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(query.exec("INSERT INTO records VALUES(1,'first'),(2,'second')"));
    assert(query.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    auto store=[&](const QString& kind,const QString& name,const QByteArray& bytes) {
        query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES(?,?,?)");
        query.addBindValue(kind); query.addBindValue(name); query.addBindValue(bytes); assert(query.exec()); query.finish();
    };
    const QByteArray definition="Begin Form\nRecordSource =\"records\"\nOnOpen =\"[Event Procedure]\"\nBegin Section\nBegin TextBox\nName =\"value\"\nControlSource =\"value\"\nEnd\nEnd\nEnd\n";
    store("form","target",definition);
    store("vba","Form_target","Sub Form_Open(Cancel As Integer)\nMe.Tag=Me.Filter\nEnd Sub\n");
    const QByteArray code=R"VBA(Sub OpenFiltered(criteria As String)
DoCmd.OpenForm "TARGET",acNormal,,criteria
Me.Tag=Forms("target").Name
If Not (Screen.ActiveForm Is Forms("target")) Then Err.Raise 5
Forms("target").Controls("value").SetFocus
If Not (Screen.ActiveControl Is Forms("target").Controls("value")) Then Err.Raise 5
End Sub
Sub OpenObject(criteria As String)
Application.DoCmd.OpenForm "target",,,criteria
End Sub
Sub OpenAndWrite(criteria As String)
DoCmd.OpenForm "target",,,criteria
Forms("target")!value="continued"
Forms("target").Dirty=False
Me.Tag=Forms("target")!value
End Sub
Sub OpenMissing()
DoCmd.OpenForm "missing"
End Sub
Sub OpenDesign()
DoCmd.OpenForm "target",acDesign
End Sub
Sub OpenNamed()
DoCmd.OpenForm "target",,"named"
End Sub
Sub OpenAdd()
DoCmd.OpenForm "target",,,,acFormAdd
End Sub
)VBA";
    AccessFormPreview owner("controller","Begin Form\nEnd\n",code,db);
    for (const auto* procedure:{"OpenFiltered","OpenObject"}) {
        std::vector<VbaValue> args{VbaValue::string("[id]=2")};
        assert(owner.executeProcedure(procedure,args).status==linaccess::VbaExecutionStatus::Completed);
        AccessFormPreview* child=nullptr;
        for (auto* widget:QApplication::topLevelWidgets()) if (auto* form=dynamic_cast<AccessFormPreview*>(widget);form && form!=&owner && form->isVisible()) { assert(!child); child=form; }
        assert(child && !child->isModal());
        auto object=makePreviewFormObject(*child);
        assert(object->member("Filter",{}).asText()=="[id]=2");
        assert(object->member("Tag",{}).asText()=="[id]=2");
        assert(object->member("FilterOn",{}).truthy());
        assert(object->item("value").asText()=="second");
        auto* selector=child->findChild<QComboBox*>("access-record-selector"); assert(selector && selector->count()==2);
        assert(makePreviewFormsObject(owner)->member("Count",{}).asNumber()==2);
        object->assign("Tag",VbaValue::string("reused instance"));
        assert(owner.executeProcedure("OpenObject",args).status==linaccess::VbaExecutionStatus::Completed);
        assert(makePreviewFormsObject(owner)->item("target").object==object && object->member("Tag",{}).asText()=="reused instance");
        assert(child->close());
        bool invalidated=false;
        try { (void)object->member("Name",{}); } catch(const linaccess::VbaRuntimeError& error) { invalidated=error.number==2450; }
        assert(invalidated);
    }
    std::vector<VbaValue> write{VbaValue::string("[id]=2")};
    assert(owner.executeProcedure("OpenAndWrite",write).status==linaccess::VbaExecutionStatus::Completed);
    assert(makePreviewFormObject(owner)->member("Tag",{}).asText()=="continued");
    assert(query.exec("SELECT value FROM records WHERE id=2") && query.next() && query.value(0).toString()=="continued"); query.finish();
    makePreviewApplicationObject(owner)->member("DoCmd",{}).object->member("Close",{VbaValue::numeric(2),VbaValue::string("target")});

    std::vector<VbaValue> none;
    for (const auto& [procedure,number]:std::vector<std::pair<QString,int>>{{"OpenMissing",2102},{"OpenDesign",3251},{"OpenNamed",3251},{"OpenAdd",3251}}) {
        const auto result=owner.executeProcedure(procedure,none);
        assert(result.status==linaccess::VbaExecutionStatus::Error && result.errorNumber==number);
    }
    std::vector<VbaValue> malformed{VbaValue::string("[id]>")};
    assert(owner.executeProcedure("OpenFiltered",malformed).errorNumber==3075);
    store("vba","Form_target","Sub Form_Open(Cancel As Integer)\nErr.Raise 11\nEnd Sub\n");
    std::vector<VbaValue> valid{VbaValue::string("[id]=2")};
    assert(owner.executeProcedure("OpenFiltered",valid).errorNumber==11);
    store("vba","Form_target","Sub Form_Open(Cancel As Integer)\nCancel=1\nEnd Sub\n");
    std::vector<VbaValue> args{VbaValue::string("[id]=2")};
    assert(owner.executeProcedure("OpenFiltered",args).errorNumber==2501);

    // Modal controls user interaction, without blocking the VBA continuation.
    auto modal=definition; modal.replace("Begin Form\n","Begin Form\nModal =-1\nOnUnload =\"[Event Procedure]\"\n");
    store("form","target",modal);
    store("vba","Form_target",R"VBA(Sub Form_Open(Cancel As Integer)
Me.Tag=Me.Filter
End Sub
Sub Form_Unload(Cancel As Integer)
If Me.Tag="block" Then Cancel=1
End Sub
Sub CloseThenMessage()
DoCmd.Close acForm,"target"
MsgBox "closed window remains alive during execution"
End Sub
)VBA");
    assert(owner.executeProcedure("OpenFiltered",args).status==linaccess::VbaExecutionStatus::Completed);
    QPointer<AccessFormPreview> modalChild;
    for (auto* widget:QApplication::topLevelWidgets()) if (auto* form=dynamic_cast<AccessFormPreview*>(widget);form && form!=&owner && form->isVisible()) { assert(!modalChild); modalChild=form; }
    assert(modalChild); assert(modalChild->isModal());
    assert(makePreviewFormObject(owner)->member("Tag",{}).asText()=="target");
    auto retained=makePreviewFormObject(*modalChild);
    retained->assign("Tag",VbaValue::string("block")); assert(!AccessFormPreview::closeProjectDocuments(db) && modalChild->isVisible());
    retained->assign("Tag",VbaValue::string("allowed"));
    bool observedMessage=false; QTimer answer; answer.setInterval(1);
    QObject::connect(&answer,&QTimer::timeout,[&] {
        if (auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget())) {
            assert(modalChild); observedMessage=true; answer.stop(); box->accept();
        }
    }); answer.start();
    QTimer messageDeadline; messageDeadline.setSingleShot(true);
    QObject::connect(&messageDeadline,&QTimer::timeout,[] { assert(false && "Modeless self-close MsgBox timed out"); }); messageDeadline.start(3000);
    assert(modalChild->executeProcedure("CloseThenMessage",none).status==linaccess::VbaExecutionStatus::Completed);
    messageDeadline.stop();
    assert(observedMessage && modalChild); QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete); assert(modalChild.isNull());
    bool invalidated=false;
    try { (void)retained->member("Name",{}); } catch(const linaccess::VbaRuntimeError& error) { invalidated=error.number==2450; }
    assert(invalidated);
    assert(owner.executeProcedure("OpenFiltered",args).status==linaccess::VbaExecutionStatus::Completed);
    auto remaining=makePreviewFormsObject(owner)->item("target").object;
    assert(owner.close());
    assert(remaining->member("Name",{}).asText()=="target");
    assert(AccessFormPreview::closeProjectDocuments(db));
    invalidated=false;
    try { (void)remaining->member("Name",{}); } catch(const linaccess::VbaRuntimeError& error) { invalidated=error.number==2450; }
    assert(invalidated);
}

void testOptionGroups() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","native-option-groups"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db); assert(query.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,choice INTEGER)"));
    assert(query.exec("INSERT INTO records VALUES(1,1)"));
    assert(query.exec("CREATE TABLE events(value INTEGER)"));
    const QByteArray definition=R"FORM(Begin Form
RecordSource ="records"
Begin Section
Height =3000
Begin OptionGroup
Name ="selection"
ControlSource ="choice"
DefaultValue ="1"
Width =3000
Height =2000
AfterUpdate ="[Event Procedure]"
OnClick ="[Event Procedure]"
Begin
Begin ToggleButton
Name ="first"
OptionValue =1
Top =300
End
Begin ToggleButton
Name ="second"
OptionValue =2
Top =800
OnClick ="[Event Procedure]"
End
End
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub BracketValue()
Dim selection As Boolean
selection=False
If [selection]=2 Then
Me.Tag="selected"
Else
Me.Tag="clear"
End If
If [selection].ControlType<>acOptionGroup Then
Err.Raise 5
End If
End Sub
Sub selection_AfterUpdate()
CurrentDb.Execute "INSERT INTO events VALUES(" & Me.selection & ")"
End Sub
Sub selection_Click()
If Me.selection<>2 Then
Err.Raise 5
End If
End Sub
)VBA";
    const auto layout=linaccess::parseAccessFormLayout(QString::fromUtf8(definition));
    assert(layout.controls.size()==3 && layout.controls[1].optionGroupIndex==0 && layout.controls[2].optionGroupIndex==0);
    AccessFormPreview bound("bound options",definition,code,db);
    auto object=makePreviewFormObject(bound);
    auto* first=bound.findChild<QPushButton*>("access-control-first");
    auto* second=bound.findChild<QPushButton*>("access-control-second"); assert(first && second);
    assert(first->isChecked() && !second->isChecked() && object->item("selection").asNumber()==1);
    std::vector<VbaValue> noArguments;
    assert(bound.executeProcedure("BracketValue",noArguments).status==linaccess::VbaExecutionStatus::Completed && object->member("Tag",{}).asText()=="clear");
    bound.show(); QApplication::processEvents();
    QTest::mouseClick(second,Qt::LeftButton);
    assert(!first->isChecked() && second->isChecked() && object->item("selection").asNumber()==2);
    assert(bound.executeProcedure("BracketValue",noArguments).status==linaccess::VbaExecutionStatus::Completed && object->member("Tag",{}).asText()=="selected");
    const auto clicked=bound.eventResult("control:selection:Click");
    assert(!bound.eventResult("control:second:Click"));
    assert(clicked && clicked->status==linaccess::VbaExecutionStatus::Completed);
    assert(query.exec("SELECT value FROM events") && query.next() && query.value(0).toInt()==2 && !query.next()); query.finish();
    assert(bound.close());
    assert(query.exec("SELECT choice FROM records WHERE id=1") && query.next() && query.value(0).toInt()==2); query.finish();
    auto unbound=definition; unbound.replace("RecordSource =\"records\"",""); unbound.replace("ControlSource =\"choice\"","");
    AccessFormPreview plain("unbound options",unbound,{},db); auto plainObject=makePreviewFormObject(plain);
    auto group=plainObject->member("Controls",{VbaValue::string("selection")}).object;
    assert(group->member("ControlType",{}).asNumber()==103 && plainObject->item("selection").asNumber()==1);
    auto* buttons=plain.findChild<QButtonGroup*>("access-options"); assert(buttons && buttons->checkedButton());
    group->assign("Value",VbaValue{}); assert(plainObject->item("selection").kind==VbaValue::Kind::Null && !buttons->checkedButton());
    group->assign("Value",VbaValue::numeric(2)); assert(buttons->checkedButton()==plain.findChild<QPushButton*>("access-control-second"));
    try { group->assign("Value",VbaValue::numeric(1.5)); assert(false); } catch(const linaccess::VbaRuntimeError& error) { assert(error.number==3251); }
    assert(plainObject->item("selection").asNumber()==2);
    assert(plain.close());
}
