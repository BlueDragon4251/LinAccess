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

}
