#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include <QLineEdit>
#include <QPushButton>
#include <QSqlQuery>
#include <QTest>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

void testEventPrefixes() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","event-prefixes"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(query.exec("INSERT INTO records VALUES(1,'initial')"));
    assert(query.exec("CREATE TABLE trace(id INTEGER PRIMARY KEY,event TEXT)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="records"
Begin Section
Height =2400
Begin TextBox
Name ="CurrentText"
EventProcPrefix ="OriginalText"
ControlSource ="value"
OnKeyDown ="[Event Procedure]"
OnKeyPress ="[Event Procedure]"
OnKeyUp ="[Event Procedure]"
BeforeUpdate ="[Event Procedure]"
AfterUpdate ="[Event Procedure]"
Width =3000
Height =300
End
Begin CommandButton
Name ="CurrentButton"
EventProcPrefix ="OriginalButton"
OnClick ="[Event Procedure]"
OnDblClick ="[Event Procedure]"
Top =450
Width =2500
Height =400
End
Begin CommandButton
Name ="Silent"
EventProcPrefix ="OriginalButton"
Top =1000
End
Begin CommandButton
Name ="Missing"
EventProcPrefix ="Absent"
OnClick ="[Event Procedure]"
Top =1500
End
End
End
)FORM";
    const QByteArray code=R"VBA(
Sub OriginalButton_Click()
CurrentDb.Execute "INSERT INTO trace(event) VALUES('click')"
End Sub
Sub OriginalButton_DblClick(Cancel As Integer)
CurrentDb.Execute "INSERT INTO trace(event) VALUES('double')"
Cancel = 1
End Sub
Sub CurrentButton_Click()
CurrentDb.Execute "INSERT INTO trace(event) VALUES('wrong')"
End Sub
Sub Missing_Click()
CurrentDb.Execute "INSERT INTO trace(event) VALUES('wrong')"
End Sub
Sub OriginalText_KeyDown(KeyCode As Integer, Shift As Integer)
CurrentDb.Execute "INSERT INTO trace(event) VALUES('down')"
End Sub
Sub OriginalText_KeyPress(KeyAscii As Integer)
CurrentDb.Execute "INSERT INTO trace(event) VALUES('press')"
If KeyAscii = 97 Then KeyAscii = 98
End Sub
Sub OriginalText_KeyUp(KeyCode As Integer, Shift As Integer)
CurrentDb.Execute "INSERT INTO trace(event) VALUES('up')"
End Sub
Sub OriginalText_BeforeUpdate(Cancel As Integer)
CurrentDb.Execute "INSERT INTO trace(event) VALUES('before')"
If Me.CurrentText = "blocked" Then Cancel = 1
End Sub
Sub OriginalText_AfterUpdate()
CurrentDb.Execute "INSERT INTO trace(event) VALUES('after')"
End Sub
Sub Alternate_Click()
CurrentDb.Execute "INSERT INTO trace(event) VALUES('alternate')"
End Sub
Sub SwitchPrefix()
Me.CurrentButton.EventProcPrefix = "Alternate"
End Sub
)VBA";
    const auto parsed=linaccess::parseAccessFormLayout(QString::fromUtf8(layout));
    assert(parsed.controls.size()==4 && parsed.controls[0].eventProcPrefix=="OriginalText");
    AccessFormPreview form("prefix regression",layout,code,db); form.show(); form.activateWindow(); QApplication::processEvents();
    auto model=makePreviewFormObject(form);
    auto control=model->member("Controls",{VbaValue::string("CurrentButton")}).object;
    assert(control && control->member("Name",{}).asText()=="CurrentButton");
    assert(control->member("EventProcPrefix",{}).asText()=="OriginalButton");
    auto* button=form.findChild<QPushButton*>("access-control-CurrentButton");
    auto* silent=form.findChild<QPushButton*>("access-control-Silent");
    auto* missing=form.findChild<QPushButton*>("access-control-Missing");
    auto* edit=form.findChild<QLineEdit*>("access-control-CurrentText");
    auto* save=form.findChild<QPushButton*>("access-save-record"); assert(button && silent && missing && edit && save);
    const auto trace=[&] {
        QStringList result; assert(query.exec("SELECT event FROM trace ORDER BY id"));
        while (query.next()) result.append(query.value(0).toString()); query.finish(); return result;
    };
    const auto clear=[&] { assert(query.exec("DELETE FROM trace")); };
    button->click(); silent->click(); missing->click(); assert(trace()==QStringList({"click"}));
    const auto absent=form.eventResult("control:Missing:Click");
    assert(absent && absent->status==linaccess::VbaExecutionStatus::ProcedureMissing);
    clear(); QTest::mouseDClick(button,Qt::LeftButton); QApplication::processEvents();
    assert(trace()==QStringList({"double"}));
    clear(); edit->setFocus(); edit->selectAll(); QTest::keyClick(edit,Qt::Key_A);
    assert(edit->text()=="b" && trace()==QStringList({"down","press","up"}));
    clear(); save->click();
    assert(trace()==QStringList({"before","after"}) && !form.hasUnsavedChanges());
    assert(query.exec("SELECT value FROM records WHERE id=1") && query.next() && query.value(0).toString()=="b"); query.finish();
    edit->setFocus(); edit->selectAll(); QTest::keyClicks(edit,"blocked"); clear(); save->click();
    assert(trace()==QStringList({"before"}) && form.hasUnsavedChanges());
    assert(query.exec("SELECT value FROM records WHERE id=1") && query.next() && query.value(0).toString()=="b"); query.finish();
    model->member("Undo",{}); assert(!form.hasUnsavedChanges()); clear();
    std::vector<VbaValue> arguments;
    assert(form.executeProcedure("SwitchPrefix",arguments).status==linaccess::VbaExecutionStatus::Completed);
    button->click(); assert(trace()==QStringList({"alternate"}));
    assert(control->member("Name",{}).asText()=="CurrentButton");
    assert(control->member("EventProcPrefix",{}).asText()=="Alternate");
    int number=0;
    try { control->assign("EventProcPrefix",VbaValue::string("bad prefix")); }
    catch (const linaccess::VbaRuntimeError& error) { number=error.number; }
    assert(number==5 && control->member("EventProcPrefix",{}).asText()=="Alternate");
    number=0;
    try { control->assign("EventProcPrefix",VbaValue{}); }
    catch (const linaccess::VbaRuntimeError& error) { number=error.number; }
    assert(number==94 && control->member("EventProcPrefix",{}).asText()=="Alternate");
    assert(form.close());
    auto defaultLayout=layout; defaultLayout.replace("EventProcPrefix =\"Absent\"","");
    AccessFormPreview defaults("default prefixes",defaultLayout,code,db);
    auto defaultControl=makePreviewFormObject(defaults)->member("Controls",{VbaValue::string("Missing")}).object;
    assert(defaultControl->member("EventProcPrefix",{}).asText()=="Missing");
    clear(); defaults.findChild<QPushButton*>("access-control-Missing")->click();
    assert(trace()==QStringList({"wrong"}) && defaults.close());
}
