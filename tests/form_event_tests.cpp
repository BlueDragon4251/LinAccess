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
#include <QLabel>
#include <QFontMetricsF>
#include <QAbstractItemView>
#include <QPushButton>
#include <QSqlQuery>
#include <QTest>
#include <QTemporaryDir>
#include <QMessageBox>
#include <QTimer>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <algorithm>
#include <iostream>

void testInsertionEvents();
void testChangeEvents();
void testTimerEvents();
void testTextBuffering();
void testControlUndo();
void testWorkspaceEvents();
void testImportedControlStates();
void testPasswordMask();
void testFormCycle();
void testCycleClose();
void testControlFonts();

void testFormEvents() {
    testInsertionEvents();
    testChangeEvents();
    testTimerEvents();
    testTextBuffering();
    testControlUndo();
    testWorkspaceEvents();
    testImportedControlStates();
    testPasswordMask();
    testFormCycle();
    testCycleClose();
    testControlFonts();
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
Begin TextBox
Name ="branch"
Top =800
End
Begin CommandButton
Name ="Transfer"
OnClick ="[Event Procedure]"
Top =1200
End
End
End
)FORM";
    const QByteArray code=R"VBA(Option Compare Text
Sub Transfer_Click()
Dim session As DAO.Workspace
Set session=DBEngine(0)
If Me.Tag="rollback" Then
Set session=Application.DBEngine.Workspaces(0)
Me!branch=1
ElseIf Me.Tag="fail" Then
Set session=DBEngine.Workspaces.Item(0)
Me!branch=2
Else
Me!branch=3
End If
If Not (Application.DBEngine Is DBEngine) Then Err.Raise 5
If Not (DBEngine(0)(0) Is DBEngine(0)(0)) Then Err.Raise 5
Me!lastError=0
On Error GoTo failed
session.BeginTrans
Dim movement As Long
movement=1
Do Until movement>2
Select Case movement
Case 1
CurrentDb.Execute "UPDATE balances SET amount=amount-10 WHERE id=1",dbFailOnError
Case Else
CurrentDb.Execute "UPDATE balances SET amount=amount+10 WHERE id=2",dbFailOnError
End Select
movement=movement+1
Loop
If Me.Tag="fail" Then CurrentDb.Execute "INSERT INTO balances VALUES(1,5)",dbFailOnError
Select Case Me.Tag
Case "rollback"
session.Rollback
Case Else
session.CommitTrans
End Select
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
    button->click(); assert(amount(1)==90 && amount(2)==210 && root->item("amount").asNumber()==90 && root->item("lastError").asNumber()==0 && root->item("branch").asNumber()==3);
    root->assign("Tag",VbaValue::string("ROLLBACK")); button->click();
    assert(amount(1)==90 && amount(2)==210 && root->item("amount").asNumber()==90 && root->item("branch").asNumber()==1);
    root->assign("Tag",VbaValue::string("fail")); button->click();
    assert(amount(1)==90 && amount(2)==210 && root->item("amount").asNumber()==90 && root->item("lastError").asNumber()==3022 && root->item("branch").asNumber()==2);
    assert(form.close());
}

void testImportedControlStates() {
    using linaccess::VbaValue;
    const auto defaults=linaccess::parseAccessFormLayout(R"FORM(Begin Form
AllowAdditions =NotDefault
AllowDeletions =NotDefault
Modal =NotDefault
KeyPreview =NotDefault
FilterOnLoad =NotDefault
Begin
Begin TextBox
Visible =NotDefault
Locked =NotDefault
End
End
Begin Section
Begin TextBox
Name ="actual"
End
End
End
)FORM");
    assert(!defaults.allowAdditions && !defaults.allowDeletions && defaults.modal && defaults.keyPreview && defaults.filterOnLoad);
    assert(defaults.controls.size()==1 && defaults.controls[0].visible && !defaults.controls[0].locked);
    auto db=QSqlDatabase::addDatabase("QSQLITE","imported-control-states"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,value TEXT)")); assert(sql.exec("INSERT INTO records VALUES(1,'alpha'),(2,'beta')"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="records"
Begin Section
Begin TextBox
Name ="last"
TabIndex =6
Top =2400
End
Begin TextBox
Name ="second"
TabIndex =2
Top =800
End
Begin TextBox
Name ="hidden"
Visible =NotDefault
TabIndex =4
Top =1600
End
Begin TextBox
Name ="first"
ControlSource ="value"
TabIndex =0
End
Begin TextBox
Name ="disabled"
Enabled =NotDefault
TabIndex =5
Top =2000
End
Begin TextBox
Name ="skip"
TabStop =NotDefault
TabIndex =1
Top =400
End
Begin TextBox
Name ="locked"
ControlSource ="value"
Locked =NotDefault
TabIndex =3
Top =1200
End
End
End
)FORM";
    const QByteArray code="Sub FocusHidden()\nMe.hidden.SetFocus\nEnd Sub\nSub FocusDisabled()\nMe.disabled.SetFocus\nEnd Sub\nSub InvalidVisible()\nMe.hidden.Visible=Null\nEnd Sub\nSub Reorder()\nMe.last.TabIndex=0\nEnd Sub\n";
    AccessFormPreview form("control states",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
    auto root=makePreviewFormObject(form);
    const auto object=[&](const char* name) { return root->objectMember(name,{}).object; };
    const auto editor=[&](const char* name) { auto* line=form.findChild<QLineEdit*>(QString("access-control-")+name); assert(line); return line; };
    auto* first=editor("first"); auto* second=editor("second"); auto* locked=editor("locked");
    auto* selector=form.findChild<QComboBox*>("access-record-selector"); assert(selector);
    assert(!editor("hidden")->isVisible() && !editor("disabled")->isEnabled() && locked->isEnabled() && locked->isReadOnly());
    assert(!object("skip")->member("TabStop",{}).truthy() && object("last")->member("TabIndex",{}).asNumber()==6);
    std::vector<VbaValue> noArguments;
    assert(form.executeProcedure("FocusHidden",noArguments).errorNumber==2110);
    assert(form.executeProcedure("FocusDisabled",noArguments).errorNumber==2110);
    assert(form.executeProcedure("InvalidVisible",noArguments).errorNumber==94 && !editor("hidden")->isVisible());
    first->setFocus(); QTest::keyClick(first,Qt::Key_Tab); assert(second->hasFocus());
    QTest::keyClick(second,Qt::Key_Backtab,Qt::ShiftModifier); assert(first->hasFocus());
    first->selectAll(); QTest::keyClicks(first,"changed"); QTest::keyClick(first,Qt::Key_Tab);
    assert(second->hasFocus() && locked->text()=="changed"); root->assign("Dirty",VbaValue::logical(false));
    locked->setFocus(); QTest::keyClicks(locked,"X"); assert(locked->text()=="changed" && !form.hasUnsavedChanges());
    object("locked")->assign("Value",VbaValue::string("programmatic")); root->assign("Dirty",VbaValue::logical(false));
    assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="programmatic"); sql.finish();
    object("first")->assign("Enabled",VbaValue::logical(false)); selector->setCurrentIndex(2);
    assert(!first->isEnabled() && first->text()=="beta");
    object("first")->assign("Enabled",VbaValue::logical(true)); object("first")->assign("Locked",VbaValue::logical(true)); selector->setCurrentIndex(1);
    assert(first->isEnabled() && first->isReadOnly() && object("first")->member("Locked",{}).truthy());
    object("first")->assign("Locked",VbaValue::logical(false)); object("hidden")->assign("Visible",VbaValue::logical(true)); selector->setCurrentIndex(2);
    assert(!first->isReadOnly() && editor("hidden")->isVisible());
    object("skip")->assign("TabStop",VbaValue::logical(true)); first->setFocus(); QTest::keyClick(first,Qt::Key_Tab); assert(editor("skip")->hasFocus());
    object("skip")->assign("TabStop",VbaValue::logical(false)); selector->setCurrentIndex(1);
    assert(!object("skip")->member("TabStop",{}).truthy());
    first->setFocus(); assert(form.executeProcedure("Reorder",noArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(object("last")->member("TabIndex",{}).asNumber()==0 && object("second")->member("TabIndex",{}).asNumber()==3 && first->hasFocus());
    QTest::keyClick(first,Qt::Key_Backtab,Qt::ShiftModifier); assert(editor("last")->hasFocus());
    QTest::keyClick(editor("last"),Qt::Key_Tab); assert(first->hasFocus());
    object("last")->assign("TabIndex",VbaValue::numeric(6)); QTest::keyClick(first,Qt::Key_Tab); assert(second->hasFocus());
    for(const auto& [value,expected]:std::vector<std::pair<VbaValue,int>>{{VbaValue::numeric(-1),5},{VbaValue::numeric(7),5},{VbaValue::numeric(32768),6},{VbaValue{},94}}) {
        int error=0; try { object("last")->assign("TabIndex",value); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
        assert(error==expected && object("last")->member("TabIndex",{}).asNumber()==6);
    }
    assert(!form.hasUnsavedChanges());
    assert(form.close());
    assert(sql.exec("CREATE TABLE readonly_records(id INTEGER PRIMARY KEY,flag INTEGER,pick INTEGER)"));
    assert(sql.exec("INSERT INTO readonly_records VALUES(1,0,1)"));
    const QByteArray readonlyLayout=R"FORM(Begin Form
RecordSource ="SELECT * FROM readonly_records"
Begin Section
Begin CheckBox
Name ="flag"
ControlSource ="flag"
End
Begin ComboBox
Name ="pick"
ControlSource ="pick"
Top =400
End
Begin OptionGroup
Name ="group"
ControlSource ="pick"
Top =800
Begin
Begin OptionButton
Name ="one"
OptionValue =1
Top =850
End
Begin OptionButton
Name ="two"
OptionValue =2
Top =1200
End
End
End
End
End
)FORM";
    AccessFormPreview readonlyForm("readonly states",readonlyLayout,{},db); readonlyForm.show(); QApplication::setActiveWindow(&readonlyForm); QApplication::processEvents();
    auto readonlyRoot=makePreviewFormObject(readonlyForm);
    auto* box=readonlyForm.findChild<QCheckBox*>("access-control-flag");
    auto* combo=readonlyForm.findChild<QComboBox*>("access-control-pick");
    auto* one=readonlyForm.findChild<QAbstractButton*>("access-control-one"); auto* two=readonlyForm.findChild<QAbstractButton*>("access-control-two");
    assert(box && combo && one && two && box->isEnabled() && combo->isEnabled() && combo->lineEdit()->isReadOnly());
    box->click(); two->click(); assert(!box->isChecked() && one->isChecked() && !two->isChecked());
    const auto previous=combo->currentText(); QTest::keyClick(combo,Qt::Key_Down); assert(combo->currentText()==previous);
    int error=0; try { readonlyRoot->objectMember("pick",{}).object->member("Dropdown",{}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==2448);
    error=0; try { readonlyRoot->assignItem("flag",VbaValue::logical(true)); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==2448);
    assert(!readonlyForm.hasUnsavedChanges()); assert(readonlyForm.close());
}

void testOriginalTabOrder(const QDir& directory,const QJsonArray& objects,QSqlDatabase db) {
    int checked=0;
    for (const auto& item:objects) {
        const auto artifact=item.toObject(); if (artifact["kind"]!="form") continue;
        QFile input(directory.filePath(artifact["file"].toString())); assert(input.open(QIODevice::ReadOnly)); const auto bytes=input.readAll();
        assert(QCryptographicHash::hash(bytes,QCryptographicHash::Sha256).toHex()==artifact["sha256"].toString().toLatin1().toLower());
        const auto layout=linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(bytes));
        AccessFormPreview form(artifact["name"].toString(),bytes,{},db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
        struct Entry { QWidget* widget; QString section; int index; };
        QVector<Entry> entries; bool flat=true;
        for (const auto& control:layout.controls) {
            auto* widget=form.findChild<QWidget*>("access-control-"+control.name);
            if (!widget || !widget->isVisible() || !widget->isEnabled() || !(widget->focusPolicy()&Qt::TabFocus)) continue;
            if (!qobject_cast<QLineEdit*>(widget) && !qobject_cast<QComboBox*>(widget) && !qobject_cast<QPushButton*>(widget) && !qobject_cast<QCheckBox*>(widget)) { flat=false; break; }
            entries.append({widget,control.section,control.tabIndex});
        }
        if (!flat || entries.size()<3) { assert(form.close()); continue; }
        const auto section=[](const QString& name) { return name=="FormHeader" ? 0 : name=="FormFooter" ? 2 : 1; };
        std::stable_sort(entries.begin(),entries.end(),[&](const auto& a,const auto& b) { return section(a.section)!=section(b.section) ? section(a.section)<section(b.section) : a.index<b.index; });
        bool distinct=false; for(int i=1;i<entries.size();++i) if(entries[i].index!=entries[0].index) distinct=true;
        if (!distinct) { assert(form.close()); continue; }
        const auto focused=[](QWidget* widget) { if(auto* combo=qobject_cast<QComboBox*>(widget)) return combo->hasFocus() || (combo->lineEdit() && combo->lineEdit()->hasFocus()); return widget->hasFocus(); };
        entries[0].widget->setFocus(); QApplication::processEvents(); assert(focused(entries[0].widget));
        const int steps=std::min<int>(entries.size()-1,4);
        for(int index=1;index<=steps;++index) { assert(QApplication::focusWidget()); QTest::keyClick(QApplication::focusWidget(),Qt::Key_Tab); assert(focused(entries[index].widget)); }
        for(int index=steps-1;index>=0;--index) { QTest::keyClick(QApplication::focusWidget(),Qt::Key_Backtab,Qt::ShiftModifier); assert(focused(entries[index].widget)); }
        assert(form.close()); ++checked;
        std::cout<<"Original imported tab traversal: flat layouts="<<checked<<", forward/backward steps="<<steps*2<<"\n";
        break;
    }
    assert(checked==1);
}

void testPasswordMask() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","password-mask"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,value TEXT)")); assert(sql.exec("INSERT INTO records VALUES(1,'probe'),(2,NULL)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="records"
Begin Section
Begin TextBox
Name ="masked"
ControlSource ="value"
InputMask ="Password"
End
Begin TextBox
Name ="unsupported"
InputMask ="00/00"
Top =400
End
End
End
)FORM";
    const QByteArray code="Sub Reveal()\nMe.masked.InputMask=\"\"\nEnd Sub\nSub Mask()\nMe.masked.InputMask=\"Password\"\nEnd Sub\n";
    AccessFormPreview form("mask",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
    auto root=makePreviewFormObject(form),masked=root->objectMember("masked",{}).object;
    auto* editor=form.findChild<QLineEdit*>("access-control-masked"); auto* unsupported=form.findChild<QLineEdit*>("access-control-unsupported");
    auto* selector=form.findChild<QComboBox*>("access-record-selector"); assert(editor && unsupported && selector);
    assert(editor->echoMode()==QLineEdit::Password && editor->text()=="probe" && editor->displayText()=="*****");
    assert(masked->member("InputMask",{}).asText()=="Password" && root->item("masked").asText()=="probe");
    assert(unsupported->isReadOnly() && !unsupported->toolTip().isEmpty());
    editor->setFocus(); editor->selectAll(); QTest::keyClicks(editor,"sample");
    assert(editor->displayText()=="******" && masked->member("Text",{}).asText()=="sample" && masked->member("Value",{}).asText()=="probe");
    root->assign("Dirty",VbaValue::logical(false));
    assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="sample"); sql.finish();
    std::vector<VbaValue> none; assert(form.executeProcedure("Reveal",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(editor->echoMode()==QLineEdit::Normal && editor->displayText()=="sample" && !form.hasUnsavedChanges());
    assert(form.executeProcedure("Mask",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(editor->echoMode()==QLineEdit::Password && editor->displayText()=="******");
    QTemporaryDir output; assert(output.isValid()); const auto pdf=output.filePath("masked.pdf"); assert(form.exportLayoutPdf(pdf));
    QFile exported(pdf); assert(exported.open(QIODevice::ReadOnly) && exported.read(5)=="%PDF-");
    for (const auto& value:std::vector<VbaValue>{VbaValue{},VbaValue::string("00/00")}) {
        int error=0; try { masked->assign("InputMask",value); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
        assert(error==(value.kind==VbaValue::Kind::Null ? 94 : 3251) && editor->echoMode()==QLineEdit::Password && editor->displayText()=="******");
    }
    masked->assign("Value",VbaValue::string("A😀Z")); root->assign("Dirty",VbaValue::logical(false));
    assert(editor->text()==QString::fromUtf8("A😀Z") && !editor->displayText().contains(QString::fromUtf8("😀")));
    selector->setCurrentIndex(2); assert(editor->echoMode()==QLineEdit::Password && editor->displayText().isEmpty() && masked->member("Value",{}).kind==VbaValue::Kind::Null);
    selector->setCurrentIndex(1); assert(editor->echoMode()==QLineEdit::Password && root->item("masked").asText()=="A😀Z");
    assert(form.close());
}

void testOriginalPasswordMask(const QDir& directory,const QJsonArray& objects) {
    int checked=0;
    auto db=QSqlDatabase::addDatabase("QSQLITE","original-mask-rendering"); db.setDatabaseName(":memory:"); assert(db.open());
    for (const auto& entry:objects) {
        const auto artifact=entry.toObject(); if (artifact["kind"]!="form") continue;
        QFile input(directory.filePath(artifact["file"].toString())); assert(input.open(QIODevice::ReadOnly)); const auto bytes=input.readAll();
        assert(QCryptographicHash::hash(bytes,QCryptographicHash::Sha256).toHex()==artifact["sha256"].toString().toLatin1().toLower());
        const auto layout=linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(bytes));
        for (const auto& control:layout.controls) if (control.type=="TextBox" && control.inputMask.compare("Password",Qt::CaseInsensitive)==0) {
            // No original table data or credential is loaded; only synthetic echo input.
            AccessFormPreview form(artifact["name"].toString(),bytes,{},db);
            auto* editor=form.findChild<QLineEdit*>("access-control-"+control.name); assert(editor && editor->echoMode()==QLineEdit::Password);
            editor->setText("probe"); assert(editor->displayText()=="*****");
            assert(form.close()); ++checked;
        }
    }
    assert(checked==1);
}

void testFormCycle() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","form-tab-cycle"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,a TEXT,b TEXT)"));
    assert(sql.exec("INSERT INTO records VALUES(1,'a1','b1'),(2,'a2','b2')")); assert(sql.exec("CREATE TABLE trace(id INTEGER PRIMARY KEY,event TEXT)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="records"
Cycle =1
BeforeUpdate ="[Event Procedure]"
AfterUpdate ="[Event Procedure]"
OnCurrent ="[Event Procedure]"
Begin Section
Begin TextBox
Name ="a"
ControlSource ="a"
TabIndex =0
BeforeUpdate ="[Event Procedure]"
AfterUpdate ="[Event Procedure]"
End
Begin TextBox
Name ="b"
ControlSource ="b"
TabIndex =1
BeforeUpdate ="[Event Procedure]"
AfterUpdate ="[Event Procedure]"
Top =400
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub a_BeforeUpdate(Cancel As Integer)
CurrentDb.Execute "INSERT INTO trace(event) VALUES('a_before')"
If Me.Tag="control_cancel" Then Cancel=1
End Sub
Sub a_AfterUpdate()
CurrentDb.Execute "INSERT INTO trace(event) VALUES('a_after')"
End Sub
Sub b_BeforeUpdate(Cancel As Integer)
CurrentDb.Execute "INSERT INTO trace(event) VALUES('b_before')"
If Me.Tag="control_cancel" Then Cancel=1
End Sub
Sub b_AfterUpdate()
CurrentDb.Execute "INSERT INTO trace(event) VALUES('b_after')"
If Me.Tag="handler_focus" Then Me.a.SetFocus
End Sub
Sub Form_BeforeUpdate(Cancel As Integer)
CurrentDb.Execute "INSERT INTO trace(event) VALUES('form_before')"
If Me.Tag="record_cancel" Then Cancel=1
End Sub
Sub Form_AfterUpdate()
CurrentDb.Execute "INSERT INTO trace(event) VALUES('form_after')"
End Sub
Sub Form_Current()
CurrentDb.Execute "INSERT INTO trace(event) VALUES('current')"
End Sub
Sub AllRecords()
Me.Cycle=0
End Sub
)VBA";
    AccessFormPreview form("tab cycle",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
    auto root=makePreviewFormObject(form); auto* a=form.findChild<QLineEdit*>("access-control-a"); auto* b=form.findChild<QLineEdit*>("access-control-b");
    auto* selector=form.findChild<QComboBox*>("access-record-selector"); assert(a && b && selector && selector->currentIndex()==1);
    const auto value=[&](int id) { sql.prepare("SELECT b FROM records WHERE id=?"); sql.addBindValue(id); assert(sql.exec() && sql.next()); const auto result=sql.value(0).toString(); sql.finish(); return result; };
    const auto count=[&] { assert(sql.exec("SELECT COUNT(*) FROM records") && sql.next()); const auto result=sql.value(0).toInt(); sql.finish(); return result; };
    const auto events=[&] { QStringList result; assert(sql.exec("SELECT event FROM trace ORDER BY id")); while(sql.next()) result<<sql.value(0).toString(); sql.finish(); return result; };
    const auto clear=[&] { assert(sql.exec("DELETE FROM trace")); };
    assert(root->member("Cycle",{}).asNumber()==1); b->setFocus(); b->selectAll(); QTest::keyClicks(b,"changed"); clear();
    QTest::keyClick(b,Qt::Key_Tab);
    assert(a->hasFocus() && selector->currentIndex()==1 && value(1)=="b1" && form.hasUnsavedChanges());
    assert(events()==QStringList({"b_before","b_after"})); root->assign("Dirty",VbaValue::logical(false)); assert(value(1)=="changed");
    std::vector<VbaValue> none; assert(form.executeProcedure("AllRecords",none).status==linaccess::VbaExecutionStatus::Completed && root->member("Cycle",{}).asNumber()==0);
    b->setFocus(); b->selectAll(); QTest::keyClicks(b,"row1"); clear(); QTest::keyClick(b,Qt::Key_Tab);
    assert(a->hasFocus() && selector->currentIndex()==2 && value(1)=="row1");
    assert(events()==QStringList({"b_before","b_after","form_before","form_after","current"}));
    clear(); QTest::keyClick(a,Qt::Key_Backtab,Qt::ShiftModifier); assert(b->hasFocus() && selector->currentIndex()==1 && events()==QStringList{"current"});
    a->setFocus(); QTest::keyClick(a,Qt::Key_Backtab,Qt::ShiftModifier); assert(a->hasFocus() && selector->currentIndex()==1);
    b->setFocus(); b->selectAll(); QTest::keyClicks(b,"blocked"); root->assign("Tag",VbaValue::string("control_cancel")); clear();
    QTest::keyClick(b,Qt::Key_Tab); assert(b->hasFocus() && b->text()=="blocked" && selector->currentIndex()==1 && value(1)=="row1");
    assert(events()==QStringList{"b_before"});
    root->assign("Tag",VbaValue::string("allow")); QTest::keyClick(b,Qt::Key_Tab); assert(a->hasFocus() && selector->currentIndex()==2 && value(1)=="blocked");
    b->setFocus(); b->selectAll(); QTest::keyClicks(b,"pending"); root->assign("Tag",VbaValue::string("record_cancel")); clear();
    QTest::keyClick(b,Qt::Key_Tab); assert(b->hasFocus() && selector->currentIndex()==2 && value(2)=="b2" && form.hasUnsavedChanges());
    assert(events()==QStringList({"b_before","b_after","form_before"}));
    root->assign("Tag",VbaValue::string("allow")); clear(); QTest::keyClick(b,Qt::Key_Tab);
    assert(a->hasFocus() && selector->currentIndex()==0 && value(2)=="pending" && root->member("NewRecord",{}).truthy());
    assert(events()==QStringList({"form_before","form_after","current"}));
    QTest::keyClicks(a,"new_a"); QTest::keyClick(a,Qt::Key_Tab); assert(b->hasFocus()); QTest::keyClicks(b,"new_b"); QTest::keyClick(b,Qt::Key_Tab);
    assert(a->hasFocus() && selector->currentIndex()==0 && count()==3 && value(3)=="new_b");
    b->setFocus(); clear(); QTest::keyClick(b,Qt::Key_Tab); assert(a->hasFocus() && selector->currentIndex()==0 && count()==3 && events().isEmpty());
    QTest::keyClick(a,Qt::Key_Backtab,Qt::ShiftModifier); assert(b->hasFocus() && selector->currentIndex()==3 && count()==3);
    root->assign("AllowAdditions",VbaValue::logical(false)); QTest::keyClick(b,Qt::Key_Tab); assert(b->hasFocus() && selector->currentIndex()==3 && count()==3);
    root->assign("Tag",VbaValue::string("handler_focus")); b->selectAll(); QTest::keyClicks(b,"focus_b"); QTest::keyClick(b,Qt::Key_Tab);
    assert(a->hasFocus() && selector->currentIndex()==3 && value(3)=="new_b" && form.hasUnsavedChanges()); root->member("Undo",{});
    root->assign("Tag",VbaValue::string("allow"));
    b->setFocus(); b->selectAll(); QTest::keyClicks(b,"hide_probe");
    auto bObject=root->objectMember("b",{}).object; bObject->assign("Visible",VbaValue::logical(false));
    assert(selector->currentIndex()==3 && value(3)=="new_b");
    bObject->assign("Visible",VbaValue::logical(true)); root->member("Undo",{});
    b->setFocus(); b->selectAll(); QTest::keyClicks(b,"disable_probe"); bObject->assign("Enabled",VbaValue::logical(false));
    assert(selector->currentIndex()==3 && value(3)=="new_b");
    bObject->assign("Enabled",VbaValue::logical(true)); root->member("Undo",{});
    for(const auto& [mode,expected]:std::vector<std::pair<VbaValue,int>>{{VbaValue{},94},{VbaValue::numeric(-1),5},{VbaValue::numeric(3),5},{VbaValue::numeric(2),3251}}) {
        int error=0; try { root->assign("Cycle",mode); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==expected && root->member("Cycle",{}).asNumber()==0);
    }
    assert(form.close());
    auto unboundLayout=layout; unboundLayout.replace("RecordSource =\"records\"",""); unboundLayout.replace("ControlSource =\"a\"",""); unboundLayout.replace("ControlSource =\"b\"",""); unboundLayout.replace("Cycle =1","Cycle =0");
    AccessFormPreview unbound("unbound tab",unboundLayout,{},db); unbound.show(); QApplication::setActiveWindow(&unbound); QApplication::processEvents();
    auto* first=unbound.findChild<QLineEdit*>("access-control-a"); auto* last=unbound.findChild<QLineEdit*>("access-control-b"); assert(first && last);
    last->setFocus(); QTest::keyClicks(last,"unbound"); QTest::keyClick(last,Qt::Key_Tab); assert(first->hasFocus() && !unbound.hasUnsavedChanges() && count()==3);
    assert(makePreviewFormObject(unbound)->item("b").asText()=="unbound"); assert(unbound.close());
}

void testCycleClose() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","tab-cycle-close"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,value TEXT)")); assert(sql.exec("INSERT INTO records VALUES(1,'old')"));
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    const QByteArray child=R"FORM(Begin Form
RecordSource ="records"
Begin Section
Begin TextBox
Name ="entry"
ControlSource ="value"
AfterUpdate ="[Event Procedure]"
End
End
End
)FORM";
    const QByteArray childCode="Sub entry_AfterUpdate()\nDoCmd.Close acForm,Me.Name\nMsgBox \"close probe\"\nEnd Sub\n";
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('form','cycle child',?)"); sql.addBindValue(child); assert(sql.exec());
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('vba','Form_cycle child',?)"); sql.addBindValue(childCode); assert(sql.exec()); sql.finish();
    AccessFormPreview caller("cycle caller","Begin Form\nBegin Section\nEnd\nEnd","Sub OpenChild()\nDoCmd.OpenForm \"cycle child\"\nEnd Sub\n",db);
    caller.show(); QApplication::setActiveWindow(&caller); QApplication::processEvents(); std::vector<VbaValue> none;
    assert(caller.executeProcedure("OpenChild",none).status==linaccess::VbaExecutionStatus::Completed);
    QPointer<AccessFormPreview> target;
    for (auto* widget:QApplication::topLevelWidgets()) if(auto* form=dynamic_cast<AccessFormPreview*>(widget);form && form!=&caller && form->objectName()=="access-form-preview-cycle child") target=form;
    if(!target) for (auto* widget:QApplication::topLevelWidgets()) if(auto* form=dynamic_cast<AccessFormPreview*>(widget);form && form!=&caller && form->findChild<QLineEdit*>("access-control-entry")) target=form;
    assert(target); QApplication::setActiveWindow(target); QApplication::processEvents();
    auto* entry=target->findChild<QLineEdit*>("access-control-entry"); assert(entry); entry->setFocus(); entry->selectAll(); QTest::keyClicks(entry,"closing");
    bool dialogSeen=false;
    QTimer::singleShot(0,&caller,[&] {
        for(auto* widget:QApplication::topLevelWidgets()) if(auto* box=qobject_cast<QMessageBox*>(widget)) {
            assert(target && !target->isVisible()); dialogSeen=true; box->accept(); return;
        }
        assert(false);
    });
    QTest::keyPress(entry,Qt::Key_Tab);
    assert(dialogSeen && QTest::qWaitFor([&] { return target.isNull(); },1500));
    assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="closing"); sql.finish();
    assert(sql.exec("SELECT COUNT(*) FROM records") && sql.next() && sql.value(0).toInt()==1); sql.finish();
    assert(caller.close());
}

void testControlFonts() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","control-fonts"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(sql.exec("INSERT INTO records VALUES(1,'one'),(2,'two')"));
    const QByteArray source=R"FORM(Begin Form
RecordSource ="records"
Begin
Begin TextBox
FontName ="DejaVu Sans"
FontSize =11
FontWeight =400
FontItalic =0
End
Begin Label
FontName ="DejaVu Sans"
FontSize =10
FontWeight =700
FontItalic =NotDefault
End
Begin CommandButton
FontSize =13
FontWeight =600
End
Begin ComboBox
FontSize =12
FontUnderline =NotDefault
End
End
Begin Section
Begin TextBox
Name ="editor"
ControlSource ="value"
InputMask ="Password"
FontSize =18
FontItalic =NotDefault
Height =600
End
Begin TextBox
Name ="sibling"
ControlSource ="value"
Top =700
End
Begin Label
Name ="small"
Caption ="Sample"
FontWeight =400
FontItalic =0
FontUnderline =NotDefault
Top =1100
Height =300
Width =4000
End
Begin Label
Name ="large"
Caption ="Sample"
FontSize =24
FontItalic =0
Top =1600
Height =900
Width =4000
End
Begin CommandButton
Name ="button"
OnClick ="[Event Procedure]"
Top =2600
Height =600
End
Begin ComboBox
Name ="combo"
Top =3300
Height =600
RowSourceType ="Value List"
RowSource ="one;two"
End
Begin Rectangle
Name ="shape"
FontSize =18
Left =8000
End
End
End
)FORM";
    auto layout=linaccess::parseAccessFormLayout(QString::fromUtf8(source)); assert(layout.controls.size()==7);
    assert(layout.controls[0].font.size==18 && layout.controls[1].font.size==11);
    assert(layout.controls[0].font.italic==true && layout.controls[1].font.italic==false);
    assert(layout.controls[2].font.weight==400 && layout.controls[3].font.weight==700);
    const QByteArray code=R"VBA(Sub button_Click()
Me!editor.FontSize=22
Me!editor.FontName="DejaVu Sans Mono"
Me!editor.FontWeight=600
If Me!editor.FontBold Then Err.Raise 5
Me!editor.FontBold=True
Me!editor.FontItalic=False
Me!editor.FontUnderline=True
Me!combo.FontSize=16
Me!combo.FontWeight=300
Me!combo.FontItalic=True
Me!combo.FontUnderline=False
Me!large.FontSize=28
Me!button.FontBold=True
Me.Tag=Me!editor.FontName & ":" & CStr(Me!editor.FontSize) & ":" & CStr(Me!editor.FontBold) & ":" & CStr(Me!editor.FontItalic) & ":" & CStr(Me!editor.FontUnderline)
End Sub
Sub FontError()
On Error Resume Next
Me!editor.FontSize=128
Me.Tag=CStr(Err.Number)
End Sub
)VBA";
    AccessFormPreview form("font regression",source,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
    auto root=makePreviewFormObject(form);
    auto* editor=form.findChild<QLineEdit*>("access-control-editor"); auto* sibling=form.findChild<QLineEdit*>("access-control-sibling");
    auto* small=form.findChild<QLabel*>("access-control-small"); auto* large=form.findChild<QLabel*>("access-control-large");
    auto* button=form.findChild<QPushButton*>("access-control-button"); auto* combo=form.findChild<QComboBox*>("access-control-combo");
    auto* selector=form.findChild<QComboBox*>("access-record-selector"); assert(editor && sibling && small && large && button && combo && selector);
    assert(editor->font().family()=="DejaVu Sans" && editor->font().pointSize()==18 && editor->font().italic());
    assert(sibling->font().pointSize()==11 && !sibling->font().italic());
    assert(small->font().pointSize()==10 && small->font().weight()==400 && small->font().underline() && !small->font().italic());
    assert(large->font().pointSize()==24 && large->font().weight()==700 && !large->font().italic());
    assert(button->font().pointSize()==13 && button->font().weight()==600);
    if (!(combo->font().pointSize()==12 && combo->font().underline() && combo->lineEdit()->font().pointSize()==12))
        std::cerr<<"Synthetic combo font: control size="<<combo->font().pointSize()<<", underline="<<combo->font().underline()<<", editor size="<<combo->lineEdit()->font().pointSize()<<"\n";
    assert(combo->font().pointSize()==12 && combo->font().underline() && combo->lineEdit()->font().pointSize()==12);
    root->objectMember("combo",{}).object->member("Requery",{});
    assert(combo->count()==2 && combo->view()->font().pointSize()==12 && combo->view()->font().underline());
    assert(combo->lineEdit()->font().pointSize()==12 && combo->lineEdit()->font().underline());
    assert(QFontMetricsF(large->font()).horizontalAdvance("Sample")>QFontMetricsF(small->font()).horizontalAdvance("Sample")*2);
    auto editorObject=root->objectMember("editor",{}).object;
    assert(editorObject->member("FontName",{}).asText()=="DejaVu Sans" && editorObject->member("FontSize",{}).asNumber()==18);
    assert(editorObject->member("FontWeight",{}).asNumber()==400 && !editorObject->member("FontBold",{}).truthy());
    assert(editorObject->member("FontItalic",{}).truthy() && !editorObject->member("FontUnderline",{}).truthy());
    editorObject->assign("FontSize",VbaValue::numeric(19.5)); assert(editor->font().pointSize()==20);
    editorObject->assign("FontSize",VbaValue::numeric(18));
    assert(!form.hasUnsavedChanges()); editor->setFocus(); editor->selectAll(); QTest::keyClicks(editor,"fresh");
    editor->setSelection(1,3); assert(editor->hasFocus());
    button->click();
    const auto click=form.eventResult("control:button:Click"); assert(click && click->status==linaccess::VbaExecutionStatus::Completed);
    assert(root->member("Tag",{}).asText()=="DejaVu Sans Mono:22:True:False:True");
    assert(editor->hasFocus() && editor->selectionStart()==1 && editor->selectedText()=="res" && editor->displayText()=="*****");
    assert(editorObject->member("Text",{}).asText()=="fresh" && editorObject->member("Value",{}).asText()=="one" && form.hasUnsavedChanges());
    assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="one"); sql.finish();
    assert(editor->font().pointSize()==22 && editor->font().bold() && !editor->font().italic() && editor->font().underline());
    assert(editor->font().family()=="DejaVu Sans Mono" && sibling->font().pointSize()==11 && !sibling->font().bold());
    assert(combo->font().pointSize()==16 && combo->font().weight()==300 && combo->font().italic() && !combo->font().underline());
    assert(combo->lineEdit()->font().pointSize()==16 && combo->view()->font().pointSize()==16 && !combo->view()->font().underline());
    editorObject->assign("FontWeight",VbaValue::numeric(699)); assert(!editorObject->member("FontBold",{}).truthy());
    editorObject->assign("FontWeight",VbaValue::numeric(700)); assert(editorObject->member("FontBold",{}).truthy());
    root->objectMember("combo",{}).object->member("Requery",{});
    assert(combo->view()->font().pointSize()==16 && combo->view()->font().italic() && editor->hasFocus());
    const auto rejected=[&](const char* property,const VbaValue& value,int expected) {
        const auto before=editor->font(); int number=0;
        try { editorObject->assign(property,value); } catch(const linaccess::VbaRuntimeError& failure) { number=failure.number; }
        assert(number==expected && editor->font()==before && editor->hasFocus() && editor->selectionStart()==1 && editor->selectedText()=="res");
    };
    for(const auto* property:{"FontName","FontSize","FontWeight","FontBold","FontItalic","FontUnderline"}) rejected(property,VbaValue{},94);
    rejected("FontName",VbaValue::string(""),5); rejected("FontSize",VbaValue::numeric(0),5);
    rejected("FontSize",VbaValue::numeric(128),5); rejected("FontSize",VbaValue::numeric(32768),6);
    rejected("FontWeight",VbaValue::numeric(1001),5); rejected("FontWeight",VbaValue::numeric(-1),5);
    rejected("FontSize",VbaValue::string("invalid"),13); rejected("FontBold",VbaValue::errorValue(5),13);
    int number=0; try { editorObject->member("FontSize",{VbaValue::numeric(1)}); } catch(const linaccess::VbaRuntimeError& failure) { number=failure.number; } assert(number==450);
    auto shape=root->objectMember("shape",{}).object;
    number=0; try { shape->member("FontSize",{}); } catch(const linaccess::VbaRuntimeError& failure) { number=failure.number; } assert(number==438);
    number=0; try { shape->assign("FontSize",VbaValue::numeric(12)); } catch(const linaccess::VbaRuntimeError& failure) { number=failure.number; } assert(number==438);
    std::vector<VbaValue> noArgs; auto failure=form.executeProcedure("FontError",noArgs);
    assert(failure.status==linaccess::VbaExecutionStatus::Completed && root->member("Tag",{}).asText()=="5" && editor->font().pointSize()==22);
    class FontScalar final:public linaccess::VbaObject {
    public:
        int reads{};
        VbaValue scalarValue() override { ++reads; return VbaValue::numeric(22.5); }
        VbaValue member(std::string_view,const std::vector<VbaValue>&) override { throw linaccess::VbaRuntimeError(438,"No font scalar members"); }
    };
    auto scalar=std::make_shared<FontScalar>(); editorObject->assign("FontSize",VbaValue::reference(scalar));
    assert(scalar->reads==1 && editor->font().pointSize()==22 && editor->selectedText()=="res" && editor->displayText()=="*****");
    root->assign("Dirty",VbaValue::logical(false));
    assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="fresh"); sql.finish();
    assert(sibling->text()=="fresh" && editor->font().pointSize()==22 && sibling->font().pointSize()==11);
    selector->setCurrentIndex(2); assert(root->item("editor").asText()=="two" && !form.hasUnsavedChanges());
    assert(editor->font().pointSize()==22 && !editor->font().italic() && !combo->lineEdit()->font().underline());
    assert(form.close());
    number=0; try { editorObject->member("FontSize",{}); } catch(const linaccess::VbaRuntimeError& failure) { number=failure.number; } assert(number==2450);
    assert(sql.exec("CREATE VIEW readonly_records AS SELECT * FROM records")); sql.finish();
    auto readonlySource=source; readonlySource.replace("RecordSource =\"records\"","RecordSource =\"readonly_records\"");
    AccessFormPreview readonlyForm("readonly font regression",readonlySource,{},db);
    auto readonlyRoot=makePreviewFormObject(readonlyForm); auto readonlyObject=readonlyRoot->objectMember("editor",{}).object;
    auto* readonlyEditor=readonlyForm.findChild<QLineEdit*>("access-control-editor"); assert(readonlyEditor && readonlyEditor->isReadOnly() && readonlyEditor->font().pointSize()==18);
    readonlyObject->assign("FontSize",VbaValue::numeric(24)); readonlyObject->assign("FontBold",VbaValue::logical(true));
    assert(readonlyEditor->font().pointSize()==24 && readonlyEditor->font().bold() && readonlyEditor->isReadOnly() && !readonlyForm.hasUnsavedChanges());
    assert(sql.exec("SELECT COUNT(*) FROM records") && sql.next() && sql.value(0).toInt()==2); sql.finish();
    assert(readonlyForm.close());
    const QByteArray defaults=R"FORM(Begin Form
Begin Section
Begin TextBox
Name ="defaultText"
End
Begin CommandButton
Name ="defaultButton"
Top =400
End
End
End
)FORM";
    AccessFormPreview defaultForm("default font regression",defaults,{},db);
    auto defaultRoot=makePreviewFormObject(defaultForm);
    assert(defaultForm.findChild<QLineEdit*>("access-control-defaultText")->font().pointSize()==8);
    assert(defaultForm.findChild<QPushButton*>("access-control-defaultButton")->font().pointSize()==10);
    assert(defaultRoot->objectMember("defaultText",{}).object->member("FontSize",{}).asNumber()==8);
    assert(defaultRoot->objectMember("defaultButton",{}).object->member("FontSize",{}).asNumber()==10);
    assert(defaultForm.close());
    for (const auto& property:QStringList{"FontSize =0","FontSize =128","FontWeight =-1","FontWeight =1001","FontItalic =invalid","FontName =\"\""}) {
        bool rejected=false;
        try { (void)linaccess::parseAccessFormLayout("Begin Form\nBegin Section\nBegin TextBox\nName =\"bad\"\n"+property+"\nEnd\nEnd\nEnd"); }
        catch(const std::runtime_error&) { rejected=true; }
        assert(rejected);
    }
}

void testOriginalControlFonts(const QDir& directory,const QJsonArray& objects) {
    auto db=QSqlDatabase::addDatabase("QSQLITE","original-control-fonts-empty"); db.setDatabaseName(":memory:"); assert(db.open());
    int forms=0,reports=0;
    for (const auto& entry:objects) {
        const auto artifact=entry.toObject(); const auto kind=artifact["kind"].toString();
        if ((kind!="form" && kind!="report") || (kind=="form" ? forms : reports)) continue;
        QFile input(directory.filePath(artifact["file"].toString())); assert(input.open(QIODevice::ReadOnly)); const auto bytes=input.readAll();
        assert(QCryptographicHash::hash(bytes,QCryptographicHash::Sha256).toHex()==artifact["sha256"].toString().toLatin1().toLower());
        const auto layout=linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(bytes));
        for (const auto& control:layout.controls) if (control.type=="Label" && control.font.size) {
            // Original layout only: no original VBA, table data or credentials.
            AccessFormPreview form(artifact["name"].toString(),bytes,{},db);
            auto* label=form.findChild<QLabel*>("access-control-"+control.name); assert(label && label->font().pointSize()==*control.font.size);
            if (control.font.weight) assert(label->font().weight()==(*control.font.weight ? *control.font.weight : 400));
            if (control.font.italic) assert(label->font().italic()==*control.font.italic);
            if (control.font.underline) assert(label->font().underline()==*control.font.underline);
            auto root=makePreviewFormObject(form); auto object=root->objectMember(control.name.toUtf8().toStdString(),{}).object;
            const auto changed=*control.font.size==127 ? 126 : *control.font.size+1;
            if (kind=="form") {
                object->assign("FontSize",linaccess::VbaValue::numeric(changed));
                assert(label->font().pointSize()==changed && object->member("FontSize",{}).asNumber()==changed && !form.hasUnsavedChanges());
            } else {
                int number=0; const auto before=label->font();
                try { object->assign("FontSize",linaccess::VbaValue::numeric(changed)); } catch(const linaccess::VbaRuntimeError& failure) { number=failure.number; }
                assert(number==3251 && label->font()==before && !form.hasUnsavedChanges());
            }
            assert(form.close()); if (kind=="form") ++forms; else ++reports; break;
        }
        if (forms && reports) break;
    }
    assert(forms==1 && reports==1);
}
