#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include "vba_source_codec.hpp"
#include <QDir>
#include <QJsonArray>
#include <QJsonObject>
#include <QFile>
#include <QCryptographicHash>
#include <QDateTime>
#include <QRegularExpression>

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
void testTimerEvents();
void testTextBuffering();
void testControlUndo();
void testWorkspaceEvents();

void testFormEvents() {
    testInsertionEvents();
    testChangeEvents();
    testTimerEvents();
    testTextBuffering();
    testControlUndo();
    testWorkspaceEvents();
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
        const auto previous=mirror->text();
        edit->setFocus(); edit->selectAll(); QTest::keyClicks(edit,text);
        assert(mirror->text()==previous);
    };
    const auto commit = [&] {
        edit->clearFocus(); QCoreApplication::processEvents();
    };
    type("updated");
    assert(dialog.hasUnsavedChanges());
    commit();
    assert(events().mid(3) == QStringList({"dirty","control_before","control_after"}));
    assert(mirror->text()=="updated");
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
Me!mirror=Me.input.Text
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
    assert(events()==QStringList{"combo_change"} && root->item("input").asText()=="b"); save->click();
    assert(sql.exec("SELECT name FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="d"); sql.finish();
    root->assign("Tag",VbaValue::string("fail_change")); clear(); edit->setFocus(); edit->selectAll(); QTest::keyClicks(edit,"e");
    const auto failed=form.eventResult("control:input:Change"); assert(failed && failed->errorNumber==5);
    assert(form.hasUnsavedChanges() && root->item("input").asText()=="d");
    assert(root->objectMember("input",{}).object->member("Text",{}).asText()=="e");
    assert(sql.exec("SELECT name FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="d"); sql.finish();
    root->member("Undo",{}); assert(form.close());
}

void testTimerEvents() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","form-timer-events"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE ticks(value INTEGER)"));
    const QByteArray layout=R"FORM(Begin Form
Tag ="0"
TimerInterval =15
OnTimer ="[Event Procedure]"
OnLoad ="[Event Procedure]"
OnUnload ="[Event Procedure]"
Begin Section
Begin TextBox
Name ="closeGuard"
End
End
End
)FORM";
    assert(linaccess::parseAccessFormLayout(QString::fromUtf8(layout)).timerInterval==15);
    const QByteArray code=R"VBA(Sub Form_Load()
Me.TimerInterval=5
End Sub
Sub Form_Timer()
CurrentDb.Execute "INSERT INTO ticks VALUES(1)"
Me.Tag=CLng(Me.Tag)+1
If CLng(Me.Tag)=3 Then Me.TimerInterval=0
End Sub
Sub Form_Unload(Cancel As Integer)
If Me!closeGuard="stay" Then Cancel=1
End Sub
)VBA";
    AccessFormPreview form("timer",layout,code,db); form.show();
    auto root=makePreviewFormObject(form);
    assert(root->member("TimerInterval",{}).asNumber()==5);
    assert(QTest::qWaitFor([&] { return root->member("Tag",{}).asText()=="3"; },1500));
    assert(root->member("TimerInterval",{}).asNumber()==0);
    const auto count=[&] { assert(sql.exec("SELECT COUNT(*) FROM ticks") && sql.next()); const int n=sql.value(0).toInt(); sql.finish(); return n; };
    assert(count()==3); QTest::qWait(30); assert(count()==3);
    for (const auto& [value,expected]:std::vector<std::pair<VbaValue,int>>{{VbaValue::numeric(-1),5},{VbaValue::numeric(2147483648.0),6},{VbaValue{},94}}) {
        int error=0; try { root->assign("TimerInterval",value); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==expected);
        assert(root->member("TimerInterval",{}).asNumber()==0);
    }
    root->assign("Tag",VbaValue::string("0")); form.hide(); root->assign("TimerInterval",VbaValue::numeric(5));
    assert(QTest::qWaitFor([&] { return root->member("Tag",{}).asText()=="3"; },1500)); assert(count()==6 && !form.isVisible());
    root->assign("Tag",VbaValue::string("0")); root->assignItem("closeGuard",VbaValue::string("stay"));
    form.show(); root->assign("TimerInterval",VbaValue::numeric(5)); assert(!form.close());
    assert(QTest::qWaitFor([&] { return root->member("Tag",{}).asText()=="3"; },1500)); assert(count()==9);
    root->assignItem("closeGuard",VbaValue::string("allow")); root->assign("TimerInterval",VbaValue::numeric(5));
    assert(form.close()); QTest::qWait(30); assert(count()==9);
    const QByteArray selfCode=R"VBA(Sub Form_Timer()
CurrentDb.Execute "INSERT INTO ticks VALUES(2)"
DoCmd.Close acForm,Me.Name
End Sub
)VBA";
    auto selfLayout=layout; selfLayout.replace("OnLoad =\"[Event Procedure]\"",""); selfLayout.replace("OnUnload =\"[Event Procedure]\"","");
    AccessFormPreview self("self closing timer",selfLayout,selfCode,db); self.show();
    assert(QTest::qWaitFor([&] { return !self.isVisible(); },1500)); assert(count()==10); QTest::qWait(30); assert(count()==10);
    const QByteArray errorCode="Sub Form_Timer()\nErr.Raise 17\nEnd Sub\n";
    AccessFormPreview failure("failed timer",selfLayout,errorCode,db); auto failureRoot=makePreviewFormObject(failure);
    assert(QTest::qWaitFor([&] { return failure.eventResult("OnTimer").has_value(); },1500));
    assert(failure.eventResult("OnTimer")->errorNumber==17); failureRoot->assign("TimerInterval",VbaValue::numeric(0)); assert(failure.close());
    auto expressionLayout=selfLayout; expressionLayout.replace("OnTimer =\"[Event Procedure]\"","OnTimer =\"=Pulse()\"");
    const QByteArray expressionCode="Public Function Pulse()\nCurrentDb.Execute \"INSERT INTO ticks VALUES(3)\"\nMe.TimerInterval=0\nPulse=1\nEnd Function\n";
    AccessFormPreview expression("expression timer",expressionLayout,expressionCode,db);
    assert(QTest::qWaitFor([&] { return expression.eventResult("OnTimer").has_value(); },1500));
    assert(expression.eventResult("OnTimer")->status==linaccess::VbaExecutionStatus::Completed && count()==11); assert(expression.close());
    AccessFormPreview report("timer report","Begin Report\nEnd\n",{},db); auto reportRoot=makePreviewFormObject(report);
    int reportError=0; try { reportRoot->member("TimerInterval",{}); } catch(const linaccess::VbaRuntimeError& error) { reportError=error.number; } assert(reportError==438);
    reportError=0; try { reportRoot->assign("TimerInterval",VbaValue::numeric(5)); } catch(const linaccess::VbaRuntimeError& error) { reportError=error.number; } assert(reportError==438); assert(report.close());
    for (const auto* invalid:{"-1","2147483648","invalid"}) {
        bool rejected=false;
        try { (void)linaccess::parseAccessFormLayout("Begin Form\nTimerInterval ="+QString::fromLatin1(invalid)+"\nEnd\n"); }
        catch(const std::runtime_error&) { rejected=true; }
        assert(rejected);
    }

}

void testOriginalTimer(const QDir& directory,const QJsonArray& objects,QSqlDatabase db) {
    const QString expected="265db230b1d253b070d2a58047e7db80ccc1d4b9f4e7ff38739e195371bc7deb";
    int executions=0;
    for (const auto& entry:objects) {
        const auto module=entry.toObject(); if (module["kind"]!="vba" || module["sha256"].toString()!=expected) continue;
        QFile source(directory.filePath(module["file"].toString())); assert(source.open(QIODevice::ReadOnly));
        const auto code=source.readAll(); assert(QCryptographicHash::hash(code,QCryptographicHash::Sha256).toHex()==expected.toLatin1());
        const auto name=module["name"].toString().mid(5); QByteArray original;
        for (const auto& candidate:objects) {
            const auto metadata=candidate.toObject(); if (metadata["kind"]!="form" || metadata["name"].toString()!=name) continue;
            QFile layout(directory.filePath(metadata["file"].toString())); assert(layout.open(QIODevice::ReadOnly)); original=layout.readAll();
            assert(QCryptographicHash::hash(original,QCryptographicHash::Sha256).toHex()==metadata["sha256"].toString().toLatin1());
        }
        assert(!original.isEmpty());
        auto layout=linaccess::decodeVbaSource(original);
        const auto imported=linaccess::parseAccessFormLayout(layout);
        assert(imported.events.value("ontimer").compare("[Event Procedure]",Qt::CaseInsensitive)==0 && imported.timerInterval>0);
        // Isolate the original timer from unrelated startup/shutdown workflows.
        layout.remove(QRegularExpression("^\\s*(?:OnOpen|OnLoad|OnCurrent|OnActivate|OnDeactivate|OnUnload|OnClose|OnDirty|BeforeUpdate|AfterUpdate|BeforeInsert|AfterInsert)\\s*=.*(?:\\r?\\n|$)",QRegularExpression::MultilineOption));
        const auto decoded=linaccess::decodeVbaSource(code);
        const QRegularExpression procedure("^\\s*(?:(?:Private|Public)\\s+)?Sub\\s+Form_Timer\\s*\\([^\\n]*\\)([\\s\\S]*?)^\\s*End\\s+Sub",QRegularExpression::MultilineOption|QRegularExpression::CaseInsensitiveOption);
        const auto body=procedure.match(decoded); assert(body.hasMatch());
        const QRegularExpression assignment("^\\s*Me[.!](?:\\[([^\\]]+)\\]|([\\p{L}_][\\p{L}\\p{N}_]*))(?:\\.(Caption))?\\s*=\\s*Now\\s*(?:\\(\\))?\\s*$",QRegularExpression::MultilineOption|QRegularExpression::CaseInsensitiveOption);
        QVector<QPair<QString,bool>> targets; auto matches=assignment.globalMatch(body.captured(1));
        while(matches.hasNext()) { const auto match=matches.next(); targets.append({match.captured(1).isEmpty() ? match.captured(2) : match.captured(1),!match.captured(3).isEmpty()}); }
        assert(targets.size()==2);
        QSqlQuery guard(db); assert(guard.exec("SAVEPOINT original_timer"));
        {
            const auto before=QDateTime::currentDateTime().addSecs(-1);
            AccessFormPreview document(name,layout.toUtf8(),code,db); auto root=makePreviewFormObject(document);
            root->assign("TimerInterval",linaccess::VbaValue::numeric(0));
            const auto sentinel=linaccess::VbaValue::string("2000-01-01 00:00:00");
            for (const auto& target:targets) {
                if (target.second) root->objectMember(target.first.toStdString(),{}).object->assign("Caption",sentinel);
                else if (target.first.compare("Caption",Qt::CaseInsensitive)==0) root->assign("Caption",sentinel);
                else root->assignItem(target.first.toStdString(),sentinel);
            }
            root->assign("TimerInterval",linaccess::VbaValue::numeric(5));
            assert(QTest::qWaitFor([&] { return document.eventResult("OnTimer").has_value(); },1500));
            root->assign("TimerInterval",linaccess::VbaValue::numeric(0));
            assert(document.eventResult("OnTimer")->status==linaccess::VbaExecutionStatus::Completed);
            const auto after=QDateTime::currentDateTime().addSecs(1);
            for (const auto& target:targets) {
                const auto value=target.second ? root->objectMember(target.first.toStdString(),{}).object->member("Caption",{})
                    : target.first.compare("Caption",Qt::CaseInsensitive)==0 ? root->member("Caption",{}) : root->item(target.first.toStdString());
                const auto date=QDateTime::fromString(QString::fromStdString(value.asText()),"yyyy-MM-dd HH:mm:ss");
                assert(date.isValid() && date>=before && date<=after);
            }
            assert(document.close()); ++executions;
        }
        assert(guard.exec("ROLLBACK TO original_timer") && guard.exec("RELEASE original_timer"));
    }
    assert(executions==1);
}

void testTextBuffering() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","text-buffer-events"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(sql.exec("INSERT INTO records VALUES(1,'old')"));
    assert(sql.exec("CREATE TABLE trace(event TEXT,edited TEXT,stored TEXT)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="records"
Begin Section
Begin TextBox
Name ="editor"
ControlSource ="value"
OnChange ="[Event Procedure]"
BeforeUpdate ="[Event Procedure]"
AfterUpdate ="[Event Procedure]"
End
Begin TextBox
Name ="mirror"
ControlSource ="value"
Top =400
End
Begin ComboBox
Name ="combo"
ControlSource ="value"
OnChange ="[Event Procedure]"
BeforeUpdate ="[Event Procedure]"
AfterUpdate ="[Event Procedure]"
Top =800
End
Begin TextBox
Name ="derived"
ControlSource ="=[editor]"
Top =1200
End
Begin CommandButton
Name ="button"
Top =1600
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub editor_Change()
CurrentDb.Execute "INSERT INTO trace VALUES('change','" & Me.editor.Text & "','" & Nz(Me.editor.Value,"NULL") & "')"
If Me.Tag="cascade" And Me.editor.Text="trigger" Then Me.editor.Text="done"
End Sub
Sub editor_BeforeUpdate(Cancel As Integer)
CurrentDb.Execute "INSERT INTO trace VALUES('before','" & Me.editor.Text & "','" & Nz(Me.editor.Value,"NULL") & "')"
If Me.Tag="block" Then Cancel=1
End Sub
Sub editor_AfterUpdate()
CurrentDb.Execute "INSERT INTO trace VALUES('after','" & Me.editor.Text & "','" & Nz(Me.editor.Value,"NULL") & "')"
End Sub
Sub ReplaceText()
Me.editor.Text="new"
End Sub
Sub combo_Change()
CurrentDb.Execute "INSERT INTO trace VALUES('combo_change','" & Me.combo.Text & "','" & Nz(Me.combo.Value,"NULL") & "')"
End Sub
Sub combo_BeforeUpdate(Cancel As Integer)
CurrentDb.Execute "INSERT INTO trace VALUES('combo_before','" & Me.combo.Text & "','" & Nz(Me.combo.Value,"NULL") & "')"
If Me.Tag="block" Then Cancel=1
End Sub
Sub combo_AfterUpdate()
CurrentDb.Execute "INSERT INTO trace VALUES('combo_after','" & Me.combo.Text & "','" & Nz(Me.combo.Value,"NULL") & "')"
End Sub
)VBA";
    AccessFormPreview form("text-buffer",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
    auto root=makePreviewFormObject(form),editor=root->objectMember("editor",{}).object;
    auto* line=form.findChild<QLineEdit*>("access-control-editor");
    auto* mirror=form.findChild<QLineEdit*>("access-control-mirror");
    auto* selector=form.findChild<QComboBox*>("access-record-selector"); assert(line && mirror && selector);
    const auto stored=[&] { assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next()); const auto out=sql.value(0).toString(); sql.finish(); return out; };
    const auto events=[&] { QStringList out; assert(sql.exec("SELECT event,edited,stored FROM trace ORDER BY rowid")); while(sql.next()) out<<sql.value(0).toString()+":"+sql.value(1).toString()+":"+sql.value(2).toString(); sql.finish(); return out; };
    const auto clear=[&] { assert(sql.exec("DELETE FROM trace")); };
    line->setFocus(); QApplication::processEvents(); std::vector<VbaValue> none;
    assert(form.executeProcedure("ReplaceText",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(editor->member("Text",{}).asText()=="new" && editor->member("Value",{}).asText()=="old");
    assert(mirror->text()=="old" && root->item("derived").asText()=="old" && stored()=="old" && form.hasUnsavedChanges());
    assert(events()==QStringList{"change:new:old"});
    root->assign("Dirty",VbaValue::logical(false));
    assert(events()==QStringList({"change:new:old","before:new:new","after:new:new"}));
    assert(stored()=="new" && mirror->text()=="new" && editor->member("Value",{}).asText()=="new");
    clear(); root->assign("Tag",VbaValue::string("block")); editor->assign("Text",VbaValue::string("blocked"));
    int error=0; try { root->assign("Dirty",VbaValue::logical(false)); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
    assert(error==2501 && stored()=="new" && mirror->text()=="new");
    assert(editor->member("Text",{}).asText()=="blocked" && editor->member("Value",{}).asText()=="new");
    root->assign("Tag",VbaValue::string("allow")); root->assign("Dirty",VbaValue::logical(false));
    assert(stored()=="blocked" && mirror->text()=="blocked");
    clear(); editor->assign("Text",VbaValue::string("pending"));
    editor->assign("Value",VbaValue::string("assigned"));
    assert(events()==QStringList{"change:pending:blocked"});
    assert(editor->member("Text",{}).asText()=="assigned" && editor->member("Value",{}).asText()=="assigned");
    QTest::keyClicks(line,"x");
    assert(editor->member("Text",{}).asText()=="assignedx" && editor->member("Value",{}).asText()=="assigned");
    QTest::keyClick(line,Qt::Key_Escape);
    assert(editor->member("Text",{}).asText()=="assigned" && editor->member("Value",{}).asText()=="assigned");
    root->member("Undo",{}); assert(stored()=="blocked" && line->text()=="blocked");
    editor->assign("Value",VbaValue::string("unsaved")); editor->assign("Text",VbaValue::string("draft"));
    assert(root->item("derived").asText()=="unsaved");
    root->member("Undo",{}); assert(root->item("derived").asText()=="blocked");
    clear(); line->setFocus(); editor->assign("Text",VbaValue::string("focus")); mirror->setFocus(); QApplication::processEvents();
    assert(editor->member("Value",{}).asText()=="focus" && stored()=="blocked");
    error=0; try { editor->member("Text",{}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==2185);
    line->setFocus(); QApplication::processEvents();
    error=0; try { editor->assign("Text",VbaValue{}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==94);
    error=0; try { root->objectMember("button",{}).object->member("Text",{}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==438);
    root->member("Undo",{}); selector->setCurrentIndex(0); line->setFocus(); clear();
    editor->assign("Text",VbaValue::string("fresh"));
    assert(editor->member("Value",{}).kind==VbaValue::Kind::Null && events()==QStringList{"change:fresh:NULL"});
    root->assign("Dirty",VbaValue::logical(false));
    assert(sql.exec("SELECT value FROM records WHERE id=2") && sql.next() && sql.value(0).toString()=="fresh"); sql.finish();
    auto combo=root->objectMember("combo",{}).object;
    auto* comboWidget=form.findChild<QComboBox*>("access-control-combo"); assert(comboWidget && comboWidget->lineEdit());
    comboWidget->lineEdit()->setFocus(); QApplication::processEvents(); clear();
    combo->assign("Text",VbaValue::string("combo"));
    assert(combo->member("Value",{}).asText()=="fresh" && line->text()=="fresh");
    assert(events()==QStringList{"combo_change:combo:fresh"});
    root->assign("Tag",VbaValue::string("block")); error=0;
    try { root->assign("Dirty",VbaValue::logical(false)); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
    assert(error==2501 && combo->member("Value",{}).asText()=="fresh");
    root->assign("Tag",VbaValue::string("allow")); clear();
    line->setFocus(); QApplication::processEvents();
    assert(events()==QStringList({"combo_before:combo:combo","combo_after:combo:combo"}));
    assert(combo->member("Value",{}).asText()=="combo" && line->text()=="combo");
    root->assign("Dirty",VbaValue::logical(false));
    assert(sql.exec("SELECT value FROM records WHERE id=2") && sql.next() && sql.value(0).toString()=="combo"); sql.finish();
    clear(); root->assign("Tag",VbaValue::string("cascade")); editor->assign("Text",VbaValue::string("trigger"));
    assert(events()==QStringList({"change:trigger:combo","change:done:combo"}));
    assert(editor->member("Text",{}).asText()=="done" && editor->member("Value",{}).asText()=="combo");
    root->assign("Dirty",VbaValue::logical(false));
    assert(sql.exec("SELECT value FROM records WHERE id=2") && sql.next() && sql.value(0).toString()=="done"); sql.finish();
    assert(form.close());
}

void testControlUndo() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","control-undo-events"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,value TEXT,other TEXT,price REAL)"));
    assert(sql.exec("INSERT INTO records VALUES(1,'old','original',1.2345)"));
    assert(sql.exec("CREATE TABLE trace(event TEXT,oldvalue TEXT,newvalue TEXT)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="records"
AfterUpdate ="[Event Procedure]"
Begin Section
Begin TextBox
Name ="editor"
ControlSource ="value"
OnChange ="[Event Procedure]"
BeforeUpdate ="[Event Procedure]"
AfterUpdate ="[Event Procedure]"
End
Begin TextBox
Name ="other"
ControlSource ="other"
Top =400
End
Begin TextBox
Name ="price"
ControlSource ="price"
Top =800
End
Begin ComboBox
Name ="combo"
ControlSource ="value"
Top =1200
End
Begin TextBox
Name ="unbound"
Top =1600
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub editor_Change()
If Me.Tag="undo_change" Then Me.editor.Undo
End Sub
Sub editor_BeforeUpdate(Cancel As Integer)
CurrentDb.Execute "INSERT INTO trace VALUES('before','" & Nz(Me.editor.OldValue,"NULL") & "','" & Nz(Me.editor.Value,"NULL") & "')"
If Me.Tag="undo_before" Then Me.editor.Undo
If Me.Tag="set_value" Then Me.editor.Value="invalid"
If Me.Tag="set_text" Then Me.editor.Text="invalid"
If Me.Tag="set_selection" Then Me.editor.SelText="invalid"
End Sub
Sub editor_AfterUpdate()
CurrentDb.Execute "INSERT INTO trace VALUES('control_after','" & Nz(Me.editor.OldValue,"NULL") & "','" & Nz(Me.editor.Value,"NULL") & "')"
If Me.Tag="after_assign" Then Me.editor.Value="after"
If Me.Tag="after_text" Then
Me.Tag="allow"
Me.editor.Text="followup"
End If
End Sub
Sub Form_AfterUpdate()
CurrentDb.Execute "INSERT INTO trace VALUES('form_after','" & Nz(Me.editor.OldValue,"NULL") & "','" & Nz(Me.editor.Value,"NULL") & "')"
End Sub
)VBA";
    AccessFormPreview form("control undo",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
    auto root=makePreviewFormObject(form),editor=root->objectMember("editor",{}).object;
    auto price=root->objectMember("price",{}).object,combo=root->objectMember("combo",{}).object;
    auto* line=form.findChild<QLineEdit*>("access-control-editor");
    auto* comboWidget=form.findChild<QComboBox*>("access-control-combo");
    auto* selector=form.findChild<QComboBox*>("access-record-selector"); assert(line && comboWidget && selector);
    const auto stored=[&] { assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next()); const auto out=sql.value(0).toString(); sql.finish(); return out; };
    const auto events=[&] { QStringList out; assert(sql.exec("SELECT event,oldvalue,newvalue FROM trace ORDER BY rowid")); while(sql.next()) out<<sql.value(0).toString()+":"+sql.value(1).toString()+":"+sql.value(2).toString(); sql.finish(); return out; };
    const auto clear=[&] { assert(sql.exec("DELETE FROM trace")); };
    assert(editor->member("OldValue",{}).asText()=="old");
    assert(price->member("OldValue",{}).kind==VbaValue::Kind::Number && price->member("OldValue",{}).asNumber()==1.2345);
    editor->assign("Value",VbaValue::string("programmatic")); root->assignItem("other",VbaValue::string("changed"));
    assert(editor->member("OldValue",{}).asText()=="old");
    line->setFocus(); editor->assign("Text",VbaValue::string("draft")); editor->member("Undo",{});
    assert(line->text()=="programmatic" && root->item("other").asText()=="changed" && form.hasUnsavedChanges() && stored()=="old");
    root->assign("Tag",VbaValue::string("undo_change")); editor->assign("Text",VbaValue::string("discard"));
    assert(line->text()=="programmatic" && root->item("other").asText()=="changed");
    root->assign("Tag",VbaValue::string("allow")); clear(); root->assign("Dirty",VbaValue::logical(false));
    assert(stored()=="programmatic" && events()==QStringList{"form_after:old:programmatic"});
    assert(editor->member("OldValue",{}).asText()=="programmatic");
    clear(); root->assign("Tag",VbaValue::string("undo_before")); editor->assign("Text",VbaValue::string("discard"));
    root->assign("Dirty",VbaValue::logical(false));
    assert(!form.hasUnsavedChanges() && line->text()=="programmatic" && events()==QStringList{"before:programmatic:discard"});
    for (const auto* mode:{"set_value","set_text","set_selection"}) {
        clear(); root->assign("Tag",VbaValue::string(mode)); editor->assign("Text",VbaValue::string("draft"));
        if (QString::fromLatin1(mode)=="set_text") {
            auto* other=form.findChild<QLineEdit*>("access-control-other"); assert(other); other->setFocus(); QApplication::processEvents();
            const auto focusResult=form.eventResult("control:editor:BeforeUpdate"); assert(focusResult && focusResult->errorNumber==2115);
        }
        int error=0; try { root->assign("Dirty",VbaValue::logical(false)); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
        const auto result=form.eventResult("control:editor:BeforeUpdate");
        assert(error==2501 && result && result->errorNumber==2115 && line->text()=="draft" && stored()=="programmatic");
        root->member("Undo",{});
    }
    clear(); root->assign("Tag",VbaValue::string("after_assign")); editor->assign("Text",VbaValue::string("updated"));
    root->assign("Dirty",VbaValue::logical(false));
    assert(events()==QStringList({"before:programmatic:updated","control_after:programmatic:updated","form_after:programmatic:after"}));
    assert(stored()=="after" && editor->member("OldValue",{}).asText()=="after");
    comboWidget->lineEdit()->setFocus(); QApplication::processEvents(); combo->assign("Text",VbaValue::string("draft")); combo->member("Undo",{});
    assert(combo->member("Value",{}).asText()=="after" && combo->member("OldValue",{}).asText()=="after" && !form.hasUnsavedChanges());
    line->setFocus(); QApplication::processEvents(); root->assign("Tag",VbaValue::string("after_text"));
    editor->assign("Text",VbaValue::string("trigger"));
    int retryError=0; try { root->assign("Dirty",VbaValue::logical(false)); } catch(const linaccess::VbaRuntimeError& failure) { retryError=failure.number; }
    assert(retryError==2501 && stored()=="after" && editor->member("Value",{}).asText()=="trigger");
    assert(editor->member("Text",{}).asText()=="followup" && form.hasUnsavedChanges());
    root->assign("Dirty",VbaValue::logical(false));
    assert(stored()=="followup" && editor->member("OldValue",{}).asText()=="followup");
    int error=0; try { editor->assign("OldValue",VbaValue::string("bad")); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==383);
    error=0; try { root->objectMember("unbound",{}).object->member("OldValue",{}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==3251);
    selector->setCurrentIndex(0); line->setFocus(); QApplication::processEvents(); clear(); root->assign("Tag",VbaValue::string("allow"));
    assert(editor->member("OldValue",{}).kind==VbaValue::Kind::Null && price->member("OldValue",{}).kind==VbaValue::Kind::Null);
    editor->assign("Text",VbaValue::string("fresh")); root->assign("Dirty",VbaValue::logical(false));
    assert(events()==QStringList({"before:NULL:fresh","control_after:NULL:fresh","form_after:NULL:fresh"}));
    assert(editor->member("OldValue",{}).asText()=="fresh");
    assert(sql.exec("SELECT value FROM records WHERE id=2") && sql.next() && sql.value(0).toString()=="fresh"); sql.finish();
    assert(form.close());
}

void testWorkspaceEvents() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","workspace-button-events"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE balances(id INTEGER PRIMARY KEY,amount INTEGER)"));
    assert(sql.exec("INSERT INTO balances VALUES(1,100),(2,200)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="balances"
Begin Section
Begin TextBox
Name ="amount"
ControlSource ="amount"
End
Begin TextBox
Name ="lastError"
Top =400
End
Begin CommandButton
Name ="Transfer"
OnClick ="[Event Procedure]"
Top =800
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub Transfer_Click()
Dim session As DAO.Workspace
Set session=DBEngine(0)
If Me.Tag="rollback" Then Set session=Application.DBEngine.Workspaces(0)
If Me.Tag="fail" Then Set session=DBEngine.Workspaces.Item(0)
If Not (Application.DBEngine Is DBEngine) Then Err.Raise 5
If Not (DBEngine(0)(0) Is DBEngine(0)(0)) Then Err.Raise 5
Me!lastError=0
On Error GoTo failed
session.BeginTrans
CurrentDb.Execute "UPDATE balances SET amount=amount-10 WHERE id=1",dbFailOnError
CurrentDb.Execute "UPDATE balances SET amount=amount+10 WHERE id=2",dbFailOnError
If Me.Tag="fail" Then CurrentDb.Execute "INSERT INTO balances VALUES(1,5)",dbFailOnError
If Me.Tag="rollback" Then
session.Rollback
Else
session.CommitTrans
End If
Me.Requery
Exit Sub
failed:
Me!lastError=Err.Number
session.Rollback
Me.Requery
End Sub)VBA";
    AccessFormPreview form("workspace events",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
    auto root=makePreviewFormObject(form); auto* button=form.findChild<QPushButton*>("access-control-Transfer"); assert(button);
    const auto amount=[&](int id) { sql.prepare("SELECT amount FROM balances WHERE id=?"); sql.addBindValue(id); assert(sql.exec() && sql.next()); const auto value=sql.value(0).toInt(); sql.finish(); return value; };
    button->click(); assert(amount(1)==90 && amount(2)==210 && root->item("amount").asNumber()==90 && root->item("lastError").asNumber()==0);
    root->assign("Tag",VbaValue::string("rollback")); button->click();
    assert(amount(1)==90 && amount(2)==210 && root->item("amount").asNumber()==90);
    root->assign("Tag",VbaValue::string("fail")); button->click();
    assert(amount(1)==90 && amount(2)==210 && root->item("amount").asNumber()==90 && root->item("lastError").asNumber()==3022);
    assert(form.close());
}
