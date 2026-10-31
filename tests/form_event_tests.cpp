#include "access_form_preview.hpp"
#include <QApplication>
#include <QComboBox>
#include <QLineEdit>
#include <QPushButton>
#include <QSqlQuery>
#include <QTest>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

void testFormEvents() {
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
If Me.Name = "blocked" Then
Cancel = 1
End If
End Sub
Sub Name_AfterUpdate()
CurrentDb.Execute "INSERT INTO events(name) VALUES ('control_after')"
End Sub
Sub Form_BeforeUpdate(Cancel As Integer)
CurrentDb.Execute "INSERT INTO events(name) VALUES ('form_before')"
If Me.Name = "cancel" Then
Cancel = 1
End If
End Sub
Sub Form_AfterUpdate()
CurrentDb.Execute "INSERT INTO events(name) VALUES ('form_after')"
End Sub
Sub Form_Unload(Cancel As Integer)
CurrentDb.Execute "INSERT INTO events(name) VALUES ('unload')"
If Me.Name = "stay" Then
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

}
