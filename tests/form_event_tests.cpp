#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include <QComboBox>
#include <QCheckBox>
#include <QLineEdit>
#include <QPushButton>
#include <QSqlQuery>
#include <QTest>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

void testInsertionEvents();
void testChangeEvents();

void testFormEvents() {
    testInsertionEvents();
    testChangeEvents();
    auto db = QSqlDatabase::addDatabase("QSQLITE","form-events");
    db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,name TEXT)"));
    assert(query.exec("INSERT INTO records VALUES (1,'first'),(2,'second')"));
    assert(query.exec("CREATE TABLE events(id INTEGER PRIMARY KEY,name TEXT)"));
    const QByteArray form = R"FORM(Begin Form
RecordSource ="records"
OnOpen ="[Event Procedure]"
OnLoad ="[Event Procedure]"
OnCurrent ="[Event Procedure]"
OnDirty ="[Event Procedure]"
BeforeUpdate ="[Event Procedure]"
AfterUpdate ="[Event Procedure]"
OnUnload ="[Event Procedure]"
OnClose ="[Event Procedure]"
Begin Section
Height =2000
Begin TextBox
Name ="Name"
ControlSource ="name"
BeforeUpdate ="[Event Procedure]"
AfterUpdate ="[Event Procedure]"
End
Begin TextBox
Name ="Mirror"
ControlSource ="name"
Top =400
End
Begin CommandButton
Name ="Unused"
Top =800
End
End
End
)FORM";
    const QByteArray code = R"VBA(
Sub Form_Open(Cancel As Integer)
CurrentDb.Execute "INSERT INTO events(name) VALUES ('open')"
End Sub
Sub Form_Load()
CurrentDb.Execute "INSERT INTO events(name) VALUES ('load')"
End Sub
Sub Form_Current()
CurrentDb.Execute "INSERT INTO events(name) VALUES ('current')"
End Sub
Sub Form_Dirty(Cancel As Integer)
CurrentDb.Execute "INSERT INTO events(name) VALUES ('dirty')"
End Sub
Sub Name_BeforeUpdate(Cancel As Integer)
CurrentDb.Execute "INSERT INTO events(name) VALUES ('control_before')"
If Me!Name = "blocked" Then
Cancel = 1
End If
End Sub
Sub Name_AfterUpdate()
CurrentDb.Execute "INSERT INTO events(name) VALUES ('control_after')"
End Sub
Sub Form_BeforeUpdate(Cancel As Integer)
CurrentDb.Execute "INSERT INTO events(name) VALUES ('form_before')"
If Me!Name = "cancel" Then
Cancel = 1
End If
End Sub
Sub Form_AfterUpdate()
CurrentDb.Execute "INSERT INTO events(name) VALUES ('form_after')"
End Sub
Sub Form_Unload(Cancel As Integer)
CurrentDb.Execute "INSERT INTO events(name) VALUES ('unload')"
If Me!Name = "stay" Then
Cancel = 1
End If
End Sub
Sub Form_Close()
CurrentDb.Execute "INSERT INTO events(name) VALUES ('close')"
End Sub
Sub Unused_Click()
CurrentDb.Execute "INSERT INTO events(name) VALUES ('unbound_event_must_not_run')"
End Sub
)VBA";
    const auto events = [&] {
        QStringList result;
        assert(query.exec("SELECT name FROM events ORDER BY id"));
        while (query.next()) result << query.value(0).toString();
        return result;
    };
    const auto value = [&](int id) {
        query.prepare("SELECT name FROM records WHERE id=?"); query.addBindValue(id); assert(query.exec());
        assert(query.next()); const auto text = query.value(0).toString(); query.finish(); return text;
    };
    AccessFormPreview dialog("events",form,code,db);
    assert(events() == QStringList({"open","load","current"}));
    auto* edit = dialog.findChild<QLineEdit*>("access-control-Name");
    auto* mirror = dialog.findChild<QLineEdit*>("access-control-Mirror");
    auto* selector = dialog.findChild<QComboBox*>("access-record-selector");
    auto* save = dialog.findChild<QPushButton*>("access-save-record");
    assert(edit && mirror && selector && save && edit->text() == "first");
    dialog.show(); dialog.activateWindow(); QCoreApplication::processEvents();
    const auto type = [&](const char* text) {
        edit->setFocus(); edit->selectAll(); QTest::keyClicks(edit,text);
        assert(mirror->text() == QString::fromLatin1(text));
    };
    const auto commit = [&] {
        edit->clearFocus(); QCoreApplication::processEvents();
    };
    type("updated");
    assert(dialog.hasUnsavedChanges());
    commit();
    assert(events().mid(3) == QStringList({"dirty","control_before","control_after"}));
    save->click(); assert(!dialog.hasUnsavedChanges() && value(1) == "updated");
    assert(events().mid(6) == QStringList({"form_before","form_after"}));
    auto* unused = dialog.findChild<QPushButton*>("access-control-Unused"); assert(unused);
    const int count = events().size(); unused->click(); assert(events().size() == count);

    selector->setCurrentIndex(2); assert(edit->text() == "second");
    type("cancel"); commit(); selector->setCurrentIndex(1);
    assert(selector->currentIndex() == 2 && dialog.hasUnsavedChanges() && value(2) == "second");
    assert(events().last() == "form_before");
    type("accepted"); commit(); selector->setCurrentIndex(1);
    assert(selector->currentIndex() == 1 && !dialog.hasUnsavedChanges() && value(2) == "accepted");
    assert(events().mid(events().size()-3) == QStringList({"form_before","form_after","current"}));

    type("blocked"); commit(); save->click();
    assert(dialog.hasUnsavedChanges() && value(1) == "updated");
    assert(events().last() == "control_before");
    type("final"); commit(); save->click(); assert(value(1) == "final");
    type("stay"); commit();
    assert(!dialog.close() && dialog.isVisible() && value(1) == "stay");
    assert(events().last() == "unload");
    type("leave"); commit();
    assert(dialog.close()); assert(events().mid(events().size()-2) == QStringList({"unload","close"}));

    assert(query.exec("DELETE FROM events"));
    const QByteArray cancelled = R"VBA(Sub Form_Open(Cancel As Integer)
CurrentDb.Execute "INSERT INTO events(name) VALUES ('cancelled_open')"
Cancel = 1
End Sub
Sub Form_Load()
CurrentDb.Execute "INSERT INTO events(name) VALUES ('unexpected_load')"
End Sub
)VBA";
    AccessFormPreview noOpen("cancelled",form,cancelled,db);
    assert(noOpen.cancelledByOpen() && events() == QStringList({"cancelled_open"}));
    AccessFormPreview insertion("new",form,code,db);
    auto* newSelector = insertion.findChild<QComboBox*>("access-record-selector");
    auto* newEdit = insertion.findChild<QLineEdit*>("access-control-Name");
    auto* newSave = insertion.findChild<QPushButton*>("access-save-record");
    assert(newSelector && newEdit && newSave);
    insertion.show(); insertion.activateWindow(); QCoreApplication::processEvents();
    newSelector->setCurrentIndex(0);
    newEdit->setFocus(); QTest::keyClicks(newEdit,"new record");
    newEdit->clearFocus(); QCoreApplication::processEvents(); newSave->click();
    assert(!insertion.hasUnsavedChanges() && newSelector->currentIndex() == 3);
    assert(value(3) == "new record"); assert(insertion.close());

    const QByteArray cancelledDirty = R"VBA(Sub Form_Dirty(Cancel As Integer)
CurrentDb.Execute "INSERT INTO events(name) VALUES ('cancelled_dirty')"
Cancel = 1
End Sub
)VBA";
    AccessFormPreview dirtyCancel("dirty_cancel",form,cancelledDirty,db);
    auto* cancelledEdit = dirtyCancel.findChild<QLineEdit*>("access-control-Name"); assert(cancelledEdit);
    dirtyCancel.show(); dirtyCancel.activateWindow(); QCoreApplication::processEvents();
    cancelledEdit->setFocus(); cancelledEdit->selectAll(); QTest::keyClicks(cancelledEdit,"ignored");
    assert(!dirtyCancel.hasUnsavedChanges() && cancelledEdit->text() == "leave" && value(1) == "leave");
    assert(dirtyCancel.close());

    assert(query.exec("CREATE TABLE typed_records(id INTEGER PRIMARY KEY,quantity INTEGER,price REAL,label TEXT,flag BOOLEAN)"));
    assert(query.exec("INSERT INTO typed_records VALUES (1,12,3.5,'',-1)"));
    const QByteArray typedForm = R"FORM(Begin Form
RecordSource ="typed_records"
Begin Section
Height =3000
Begin TextBox
Name ="Quantity"
ControlSource ="quantity"
End
Begin TextBox
Name ="Price"
ControlSource ="price"
Top =400
End
Begin TextBox
Name ="LabelValue"
ControlSource ="label"
Top =800
End
Begin CheckBox
Name ="Flag"
ControlSource ="flag"
Top =1200
End
Begin TextBox
Name ="Result"
Top =1600
End
Begin TextBox
Name ="Temporary"
End
Begin TextBox
Name ="BoolText"
End
Begin TextBox
Name ="NullText"
End
Begin CommandButton
Name ="Compute"
OnClick ="[Event Procedure]"
End
Begin CommandButton
Name ="Assign"
OnClick ="[Event Procedure]"
End
End
End
)FORM";
    const QByteArray typedCode = R"VBA(Sub Compute_Click()
Me!Result = Me!Quantity + Me!Price
Me!BoolText = CStr(Me!Flag)
If IsNull(Me!LabelValue) Then
Me!NullText = "null"
Else
Me!NullText = "text"
End If
End Sub
Sub Assign_Click()
Me!Quantity = 2.5
Me!Flag = False
Me!LabelValue = ""
Me!Temporary = 7
Me!Result = Me!Temporary + Me!Quantity
End Sub
)VBA";
    AccessFormPreview typed("typed",typedForm,typedCode,db);
    auto* compute = typed.findChild<QPushButton*>("access-control-Compute");
    auto* assign = typed.findChild<QPushButton*>("access-control-Assign");
    auto* result = typed.findChild<QLineEdit*>("access-control-Result");
    auto* boolean = typed.findChild<QLineEdit*>("access-control-BoolText");
    auto* null = typed.findChild<QLineEdit*>("access-control-NullText");
    auto* typedSave = typed.findChild<QPushButton*>("access-save-record");
    assert(compute && assign && result && boolean && null && typedSave);
    compute->click(); assert(result->text() == "15.5" && boolean->text() == "True" && null->text() == "text");
    assign->click(); assert(result->text() == "9"); typedSave->click();
    assert(query.exec("SELECT quantity,typeof(quantity),flag,label,label IS NULL FROM typed_records WHERE id=1"));
    assert(query.next() && query.value(0).toInt() == 2 && query.value(1).toString() == "integer");
    assert(query.value(2).toInt() == 0 && query.value(3).toString().isEmpty() && query.value(4).toInt() == 0);
    query.finish();
    compute->click(); assert(result->text() == "5.5" && boolean->text() == "False");
    typed.show(); typed.activateWindow(); QCoreApplication::processEvents();
    auto* quantity = typed.findChild<QLineEdit*>("access-control-Quantity"); assert(quantity);
    quantity->setFocus(); quantity->selectAll(); QTest::keyClicks(quantity,"bad"); quantity->clearFocus();
    typedSave->click(); assert(typed.hasUnsavedChanges());
    assert(query.exec("SELECT quantity FROM typed_records WHERE id=1")); assert(query.next() && query.value(0).toInt() == 2); query.finish();
    quantity->setFocus(); quantity->selectAll(); QTest::keyClicks(quantity,"3"); quantity->clearFocus(); typedSave->click();
    assert(!typed.hasUnsavedChanges()); assert(typed.close());

    assert(query.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('query','ranked',?)");
    query.addBindValue(QByteArray("SELECT TOP 1 id,name FROM records ORDER BY id")); assert(query.exec()); query.finish();
    const QByteArray queryForm = R"FORM(Begin Form
RecordSource ="ranked"
Begin Section
Height =1000
Begin TextBox
Name ="QueryName"
ControlSource ="name"
End
End
End
)FORM";
    AccessFormPreview queryPreview("query",queryForm,{},db);
    auto* queryName = queryPreview.findChild<QLineEdit*>("access-control-QueryName");
    auto* querySave = queryPreview.findChild<QPushButton*>("access-save-record");
    assert(queryName && queryName->isReadOnly() && !queryName->text().isEmpty());
    assert(querySave && !querySave->isEnabled());
    const auto selectors = queryPreview.findChildren<QComboBox*>();
    bool foundRows = false;
    for (auto* selector : selectors) if (selector->count() == 2) foundRows = true;
    assert(foundRows);

    assert(queryPreview.close());
    assert(query.exec("DELETE FROM events")); query.finish();
    const QByteArray activationForm=R"FORM(Begin Form
RecordSource ="records"
OnOpen ="[Event Procedure]"
OnLoad ="[Event Procedure]"
OnActivate ="[Event Procedure]"
OnDeactivate ="[Event Procedure]"
OnCurrent ="[Event Procedure]"
OnUnload ="[Event Procedure]"
OnClose ="[Event Procedure]"
Begin Section
Height =300
End
End
)FORM";
    const QByteArray activationCode=R"VBA(
Sub Form_Open(Cancel As Integer)
CurrentDb.Execute "INSERT INTO events(name) VALUES ('open')"
End Sub
Sub Form_Load()
CurrentDb.Execute "INSERT INTO events(name) VALUES ('load')"
End Sub
Sub Form_Activate()
CurrentDb.Execute "INSERT INTO events(name) VALUES ('activate')"
End Sub
Sub Form_Deactivate()
CurrentDb.Execute "INSERT INTO events(name) VALUES ('deactivate')"
End Sub
Sub Form_Current()
CurrentDb.Execute "INSERT INTO events(name) VALUES ('current')"
End Sub
Sub Form_Unload(Cancel As Integer)
CurrentDb.Execute "INSERT INTO events(name) VALUES ('unload')"
If Me.Tag = "stay" Then
Cancel = 1
End If
End Sub
Sub Form_Close()
CurrentDb.Execute "INSERT INTO events(name) VALUES ('close')"
End Sub
)VBA";
    AccessFormPreview firstWindow("first activation",activationForm,activationCode,db);
    assert(events()==QStringList({"open","load"}));
    firstWindow.findChild<QComboBox*>("access-record-selector")->setCurrentIndex(2);
    assert(events()==QStringList({"open","load"}));
    firstWindow.show(); QApplication::setActiveWindow(&firstWindow); QApplication::processEvents();
    assert(events()==QStringList({"open","load","activate","current"}));
    QWidget nonDocument;
    nonDocument.show(); QApplication::setActiveWindow(&nonDocument); QApplication::processEvents();
    assert(events().size()==4);
    QApplication::setActiveWindow(&firstWindow); QApplication::processEvents();
    assert(events().size()==4);
    AccessFormPreview secondWindow("second activation",activationForm,activationCode,db);
    assert(events().mid(4)==QStringList({"open","load"}));
    secondWindow.show(); QApplication::setActiveWindow(&secondWindow); QApplication::processEvents();
    assert(events().mid(6)==QStringList({"deactivate","activate","current"}));
    QApplication::setActiveWindow(&firstWindow); QApplication::processEvents();
    assert(events().mid(9)==QStringList({"deactivate","activate"}));
    firstWindow.setProperty("la_access_tag","stay");
    assert(!firstWindow.close());
    assert(events().mid(11)==QStringList({"unload"}));
    firstWindow.setProperty("la_access_tag","");
    assert(firstWindow.close());
    assert(events().mid(12)==QStringList({"unload","deactivate","close"}));
    QApplication::setActiveWindow(&secondWindow); QApplication::processEvents();
    assert(events().mid(15)==QStringList({"activate"}));
    assert(secondWindow.close());
    assert(events().mid(16)==QStringList({"unload","deactivate","close"}));
    assert(nonDocument.close());
    assert(query.exec("DELETE FROM events")); query.finish();
    QWidget container;
    AccessFormPreview embedded("embedded activation",activationForm,activationCode,db,&container,true);
    container.show(); embedded.show(); QApplication::setActiveWindow(&container); QApplication::processEvents();
    assert(events()==QStringList({"open","load","current"}));
    assert(!embedded.eventResult("OnActivate"));
    assert(embedded.close()); assert(container.close());


}

void testInsertionEvents() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","insertion-events"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,name TEXT UNIQUE,flag INTEGER)"));
    assert(sql.exec("CREATE TABLE trace(id INTEGER PRIMARY KEY,event TEXT)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="records"
BeforeInsert ="[Event Procedure]"
AfterInsert ="[Event Procedure]"
OnDirty ="[Event Procedure]"
BeforeUpdate ="[Event Procedure]"
AfterUpdate ="[Event Procedure]"
Begin Section
Begin TextBox
Name ="name"
ControlSource ="name"
End
Begin CheckBox
Name ="flag"
ControlSource ="flag"
Top =500
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub Form_BeforeInsert(Cancel As Integer)
CurrentDb.Execute "INSERT INTO trace(event) VALUES('before_insert')"
If Me.Tag="cancel_insert" Then Cancel=1
If Me.Tag="fail_insert" Then Err.Raise 5
If Me.Tag="undo_insert" Then Me.Undo
If Me.Tag="initialize" Then Me.flag=True
End Sub
Sub Form_Dirty(Cancel As Integer)
CurrentDb.Execute "INSERT INTO trace(event) VALUES('dirty')"
End Sub
Sub Form_BeforeUpdate(Cancel As Integer)
CurrentDb.Execute "INSERT INTO trace(event) VALUES('before_update')"
If Me.Tag="cancel_save" Then Cancel=1
End Sub
Sub Form_AfterUpdate()
CurrentDb.Execute "INSERT INTO trace(event) VALUES('after_update')"
End Sub
Sub Form_AfterInsert()
If DCount("*","records")<1 Then Err.Raise 5
CurrentDb.Execute "INSERT INTO trace(event) VALUES('after_insert')"
End Sub
)VBA";
    AccessFormPreview form("insertion",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
    auto root=makePreviewFormObject(form);
    auto* edit=form.findChild<QLineEdit*>("access-control-name");
    auto* selector=form.findChild<QComboBox*>("access-record-selector");
    auto* save=form.findChild<QPushButton*>("access-save-record"); assert(edit && selector && save);
    const auto events=[&] { QStringList out; assert(sql.exec("SELECT event FROM trace ORDER BY id")); while(sql.next()) out<<sql.value(0).toString(); sql.finish(); return out; };
    const auto clear=[&] { assert(sql.exec("DELETE FROM trace")); };
    root->assign("Tag",VbaValue::string("cancel_insert"));
    edit->setFocus(); QTest::keyClicks(edit,"a");
    assert(edit->text().isEmpty() && !form.hasUnsavedChanges());
    assert(events()==QStringList{"before_insert"}); clear();
    root->assign("Tag",VbaValue::string("initialize")); QTest::keyClicks(edit,"typed");
    assert(events()==QStringList({"before_insert","dirty"}) && form.hasUnsavedChanges());
    assert(root->item("flag").truthy());
    save->click();
    assert(events()==QStringList({"before_insert","dirty","before_update","after_update","after_insert"}));
    assert(!form.hasUnsavedChanges());
    assert(sql.exec("SELECT name,flag FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="typed" && sql.value(1).toInt()==-1); sql.finish();
    clear(); selector->setCurrentIndex(0);
    root->assign("Tag",VbaValue::string("allow")); root->assignItem("name",VbaValue::string("programmatic")); save->click();
    assert(events()==QStringList({"dirty","before_update","after_update"}));
    assert(sql.exec("SELECT name FROM records WHERE id=2") && sql.next() && sql.value(0).toString()=="programmatic"); sql.finish();
    clear(); selector->setCurrentIndex(0); edit->setFocus(); QTest::keyClicks(edit,"retry");
    root->assign("Tag",VbaValue::string("cancel_save")); save->click();
    assert(form.hasUnsavedChanges() && root->member("NewRecord",{}).truthy());
    assert(events()==QStringList({"before_insert","dirty","before_update"}));
    root->assign("Tag",VbaValue::string("allow")); save->click();
    assert(!form.hasUnsavedChanges());
    assert(events()==QStringList({"before_insert","dirty","before_update","before_update","after_update","after_insert"}));
    assert(sql.exec("SELECT name FROM records WHERE id=3") && sql.next() && sql.value(0).toString()=="retry"); sql.finish();
    clear(); selector->setCurrentIndex(0); edit->setFocus(); QTest::keyClicks(edit,"undo"); root->member("Undo",{});
    assert(!form.hasUnsavedChanges()); clear();
    QTest::keyClicks(edit,"again"); save->click();
    assert(events()==QStringList({"before_insert","dirty","before_update","after_update","after_insert"}));
    clear(); selector->setCurrentIndex(0);
    auto* checkbox=form.findChild<QCheckBox*>("access-control-flag"); assert(checkbox);
    root->assign("Tag",VbaValue::string("cancel_insert")); checkbox->click();
    assert(!form.hasUnsavedChanges() && events()==QStringList{"before_insert"}); clear();
    root->assign("Tag",VbaValue::string("allow")); checkbox->click();
    assert(form.hasUnsavedChanges() && events()==QStringList({"before_insert","dirty"}));
    save->click(); assert(events()==QStringList({"before_insert","dirty","before_update","after_update","after_insert"}));
    clear(); selector->setCurrentIndex(0); root->assign("Tag",VbaValue::string("allow"));
    edit->setFocus(); QTest::keyClicks(edit,"typed"); save->click();
    assert(form.hasUnsavedChanges() && root->member("NewRecord",{}).truthy());
    assert(events()==QStringList({"before_insert","dirty","before_update"}));
    edit->setFocus(); edit->selectAll(); QTest::keyClicks(edit,"constraint retry"); save->click();
    assert(!form.hasUnsavedChanges());
    assert(events()==QStringList({"before_insert","dirty","before_update","before_update","after_update","after_insert"}));
    clear(); selector->setCurrentIndex(0); root->assign("Tag",VbaValue::string("fail_insert"));
    edit->setFocus(); QTest::keyClicks(edit,"f");
    assert(!form.hasUnsavedChanges() && edit->text().isEmpty() && events()==QStringList{"before_insert"});
    clear(); root->assign("Tag",VbaValue::string("undo_insert")); QTest::keyClicks(edit,"u");
    assert(!form.hasUnsavedChanges() && edit->text().isEmpty() && events()==QStringList{"before_insert"});
    assert(form.close());
}

void testChangeEvents() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","control-change-events"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,name TEXT)"));
    assert(sql.exec("CREATE TABLE trace(id INTEGER PRIMARY KEY,event TEXT)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="records"
BeforeInsert ="[Event Procedure]"
Begin Section
Begin TextBox
Name ="input"
EventProcPrefix ="Editor"
ControlSource ="name"
OnChange ="[Event Procedure]"
OnKeyDown ="[Event Procedure]"
OnKeyPress ="[Event Procedure]"
OnKeyUp ="[Event Procedure]"
End
Begin TextBox
Name ="mirror"
OnChange ="[Event Procedure]"
Top =400
End
Begin ComboBox
Name ="combo"
ControlSource ="name"
OnChange ="[Event Procedure]"
Top =800
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub Form_BeforeInsert(Cancel As Integer)
CurrentDb.Execute "INSERT INTO trace(event) VALUES('insert')"
End Sub
Sub Editor_KeyDown(KeyCode As Integer,Shift As Integer)
CurrentDb.Execute "INSERT INTO trace(event) VALUES('down')"
End Sub
Sub Editor_KeyPress(KeyAscii As Integer)
CurrentDb.Execute "INSERT INTO trace(event) VALUES('press')"
If Me.Tag="block" Then KeyAscii=0
End Sub
Sub Editor_KeyUp(KeyCode As Integer,Shift As Integer)
CurrentDb.Execute "INSERT INTO trace(event) VALUES('up')"
End Sub
Sub Editor_Change()
CurrentDb.Execute "INSERT INTO trace(event) VALUES('change')"
Me!mirror=Me!input
If Me.Tag="fail_change" Then Err.Raise 5
End Sub
Sub mirror_Change()
CurrentDb.Execute "INSERT INTO trace(event) VALUES('unexpected_programmatic_change')"
End Sub
Sub combo_Change()
CurrentDb.Execute "INSERT INTO trace(event) VALUES('combo_change')"
End Sub
)VBA";
    AccessFormPreview form("change",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
    auto root=makePreviewFormObject(form);
    auto* edit=form.findChild<QLineEdit*>("access-control-input");
    auto* selector=form.findChild<QComboBox*>("access-record-selector");
    auto* combo=form.findChild<QComboBox*>("access-control-combo");
    auto* save=form.findChild<QPushButton*>("access-save-record"); assert(edit && selector && combo && combo->lineEdit() && save);
    const auto events=[&] { QStringList out; assert(sql.exec("SELECT event FROM trace ORDER BY id")); while(sql.next()) out<<sql.value(0).toString(); sql.finish(); return out; };
    const auto clear=[&] { assert(sql.exec("DELETE FROM trace")); };
    edit->setFocus(); QTest::keyClicks(edit,"a");
    assert(events()==QStringList({"down","press","insert","change","up"}));
    assert(root->item("mirror").asText()=="a"); save->click();
    assert(sql.exec("SELECT name FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="a"); sql.finish();
    selector->setCurrentIndex(0); clear(); root->assignItem("input",VbaValue::string("programmatic"));
    assert(events().isEmpty() && root->item("mirror").asText()=="a"); save->click();
    selector->setCurrentIndex(0); root->assign("Tag",VbaValue::string("block")); clear(); edit->setFocus(); QTest::keyClicks(edit,"z");
    assert(events()==QStringList({"down","press","up"}) && !form.hasUnsavedChanges() && edit->text().isEmpty());
    root->assign("Tag",VbaValue::string("allow")); selector->setCurrentIndex(1); clear(); edit->setFocus(); edit->selectAll(); QTest::keyClicks(edit,"b");
    assert(events()==QStringList({"down","press","change","up"}) && root->item("mirror").asText()=="b"); save->click();
    clear(); combo->lineEdit()->setFocus(); combo->lineEdit()->selectAll(); QTest::keyClicks(combo->lineEdit(),"d");
    assert(events()==QStringList{"combo_change"} && root->item("input").asText()=="d"); save->click();
    assert(sql.exec("SELECT name FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="d"); sql.finish();
    root->assign("Tag",VbaValue::string("fail_change")); clear(); edit->setFocus(); edit->selectAll(); QTest::keyClicks(edit,"e");
    const auto failed=form.eventResult("control:input:Change"); assert(failed && failed->errorNumber==5);
    assert(form.hasUnsavedChanges() && root->item("input").asText()=="e");
    assert(sql.exec("SELECT name FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="d"); sql.finish();
    root->member("Undo",{}); assert(form.close());
}
