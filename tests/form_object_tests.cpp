#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QLineEdit>
#include <QComboBox>
#include <QToolBar>
#include <QAction>
#include <QPushButton>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

void testFormObjects() {
    auto db = QSqlDatabase::addDatabase("QSQLITE","form-object-test");
    db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,name TEXT)"));
    assert(query.exec("INSERT INTO records VALUES(1,'original')"));
    const QByteArray form = R"FORM(Begin Form
RecordSource ="records"
Begin Section
Height =2000
Begin TextBox
Name ="Name.Field"
ControlSource ="name"
End
Begin TextBox
Name ="Summary"
Top =400
End
Begin CommandButton
Name ="Scan"
OnClick ="[Event Procedure]"
Top =800
End
Begin CommandButton
Name ="Write"
OnClick ="[Event Procedure]"
Top =1200
End
End
End
)FORM";
    const QByteArray code = R"VBA(
Sub Scan_Click()
Dim ctl As Control
Dim frm As Form
Dim count As Long
Dim formVisits As Long
For Each frm In Forms
formVisits = formVisits + 1
Next frm
If formVisits <> Forms.Count Then
Err.Raise 5
End If
Set frm = Me
If Not (frm Is Forms("first form")) Then
Err.Raise 5
End If
If Not (Me.Controls("Name.Field") Is frm.Controls("Name.Field")) Then
Err.Raise 5
End If
For Each ctl In Me.Controls
If ctl.ControlType = acTextBox Then
count = count + 1
End If
Next ctl
Me!Summary = frm.Name & ":" & CStr(count) & ":" & CStr(Forms.Count)
End Sub
Sub Write_Click()
Dim ctl As Control
Forms![second form]![Name.Field] = "changed"
Forms("second form").Dirty = False
Set ctl = Forms("second form").Controls("Name.Field")
Me!Summary = ctl.Value
End Sub
)VBA";
    AccessFormPreview first("first form",form,code,db);
    AccessFormPreview second("second form",form,{},db);
    auto* scan = first.findChild<QPushButton*>("access-control-Scan");
    auto* write = first.findChild<QPushButton*>("access-control-Write");
    auto* summary = first.findChild<QLineEdit*>("access-control-Summary");
    assert(scan && write && summary);
    scan->click(); assert(summary->text()=="first form:2:2");
    write->click(); assert(summary->text()=="changed");
    assert(query.exec("SELECT name FROM records WHERE id=1") && query.next());
    assert(query.value(0).toString()=="changed"); query.finish();
    auto forms = makePreviewFormsObject(first);
    assert(forms->enumerate().size()==2);
    assert(forms->member("Item",{linaccess::VbaValue::numeric(0)}).object->member("Name",{}).asText()=="first form");
    auto secondObject = forms->member("Item",{linaccess::VbaValue::string("SECOND FORM")}).object;
    auto controls = secondObject->member("Controls",{}).object;
    assert(controls->member("Count",{}).asNumber()==4);
    auto control = controls->member("Item",{linaccess::VbaValue::string("Name.Field")}).object;
    assert(control->member("ControlType",{}).asNumber()==109);
    control->assign("Visible",linaccess::VbaValue::logical(false));
    assert(!control->member("Visible",{}).boolean);
    control->assign("Visible",linaccess::VbaValue::logical(true));
    assert(control->member("Visible",{}).boolean);
    bool failed = false;
    try { (void)forms->member("Item",{linaccess::VbaValue::string("missing")}); }
    catch(const linaccess::VbaRuntimeError& error) { failed = error.number==2450; }
    assert(failed);
    assert(second.close());
    assert(forms->member("Count",{}).asNumber()==1);
    failed = false;
    try { (void)control->member("Value",{}); }
    catch(const linaccess::VbaRuntimeError& error) { failed = error.number==2450; }
    assert(failed);
    auto* temporary = new AccessFormPreview("temporary",form,{},db);
    auto retained = makePreviewFormObject(*temporary);
    delete temporary;
    failed = false;
    try { (void)retained->member("Name",{}); }
    catch(const linaccess::VbaRuntimeError& error) { failed = error.number==2450; }
    assert(failed);
    auto separate = QSqlDatabase::addDatabase("QSQLITE","form-object-other-project");
    separate.setDatabaseName(":memory:"); assert(separate.open());
    AccessFormPreview other("other project",form,{},separate);
    assert(forms->member("Count",{}).asNumber()==1);

    auto comboForm = form;
    comboForm.replace("Begin TextBox\nName =\"Summary\"", "Begin ComboBox\nName =\"Summary\"\nRowSource =\"SELECT name FROM records ORDER BY id\"");
    AccessFormPreview comboPreview("combo form",comboForm,{},db);
    auto comboObject = makePreviewFormObject(comboPreview)->member("Controls",{linaccess::VbaValue::string("Summary")}).object;
    auto* combo = comboPreview.findChild<QComboBox*>("access-control-Summary"); assert(combo);
    comboObject->member("Requery",{}); assert(combo->count()==1 && combo->itemText(0)=="changed");
    assert(query.exec("INSERT INTO records VALUES(2,'second')")); query.finish();
    comboObject->member("Requery",{}); assert(combo->count()==2 && !comboPreview.hasUnsavedChanges());
    assert(comboPreview.close());

    auto conflictForm = form;
    conflictForm.replace("Name.Field","Name");
    const QByteArray conflictCode = R"VBA(
Sub Scan_Click()
Me!Summary = Me.Name & ":" & Me!Name.Value & ":" & Forms!conflict!Name
End Sub
Sub Write_Click()
Forms!conflict!Name = "via bang"
Forms!conflict.Dirty = False
On Error Resume Next
Me.Name = "forbidden"
Me!Summary = CStr(Err.Number)
End Sub
)VBA";
    AccessFormPreview conflict("conflict",conflictForm,conflictCode,db);
    auto* conflictSummary = conflict.findChild<QLineEdit*>("access-control-Summary"); assert(conflictSummary);
    conflict.findChild<QPushButton*>("access-control-Scan")->click();
    assert(conflictSummary->text()=="conflict:changed:changed");
    conflict.findChild<QPushButton*>("access-control-Write")->click();
    assert(conflictSummary->text()=="438");
    assert(query.exec("SELECT name FROM records WHERE id=1") && query.next());
    assert(query.value(0).toString()=="via bang"); query.finish();
    assert(conflict.close());

    AccessFormPreview countNamed("Count",form,{},db);
    assert(forms->item("Count").object->member("Name",{}).asText()=="Count");
    assert(forms->member("Count",{}).asNumber()==2);
    assert(countNamed.close());

    auto firstObject = makePreviewFormObject(first);
    firstObject->assign("Caption",linaccess::VbaValue::string("Changed caption"));
    assert(first.windowTitle().startsWith("Changed caption"));
    assert(firstObject->member("Caption",{}).asText()=="Changed caption");
    auto scanObject = firstObject->member("Controls",{linaccess::VbaValue::string("Scan")}).object;
    scanObject->assign("Caption",linaccess::VbaValue::string("New button caption"));
    assert(scan->text()=="New button caption" && scanObject->member("Caption",{}).asText()=="New button caption");
    bool invalidCaption = false;
    try { (void)firstObject->member("Controls",{linaccess::VbaValue::string("Name.Field")}).object->member("Caption",{}); }
    catch(const linaccess::VbaRuntimeError& error) { invalidCaption = error.number==438; }
    assert(invalidCaption);

    auto reportSource = form;
    reportSource.replace("Begin Form","Begin Report");
    AccessFormPreview report("toolbar report",reportSource,{},db);
    auto application = makePreviewApplicationObject(report);
    auto bars = application->member("CommandBars",{}).object;
    assert(bars->member("Count",{}).asNumber()==1);
    auto bar = bars->member("Item",{linaccess::VbaValue::numeric(1)}).object;
    auto commandControls = bar->member("Controls",{}).object;
    auto command = commandControls->member("Item",{linaccess::VbaValue::numeric(1)}).object;
    auto* toolbar = report.findChild<QToolBar*>("access-print-preview-toolbar"); assert(toolbar);
    assert(command->member("Caption",{}).asText()==toolbar->actions()[0]->text().toUtf8().toStdString());
    bool zeroRejected = false;
    try { (void)commandControls->member("Item",{linaccess::VbaValue::numeric(0)}); }
    catch(const linaccess::VbaRuntimeError& error) { zeroRejected = error.number==5; }
    assert(zeroRejected);
    command->assign("Caption",linaccess::VbaValue::string("Renamed command"));
    assert(toolbar->actions()[0]->text()=="Renamed command");
    assert(commandControls->member("Item",{linaccess::VbaValue::string("Renamed command")}).object==command);
    bar->assign("Enabled",linaccess::VbaValue::logical(false)); assert(!toolbar->isEnabled());
    bar->assign("Enabled",linaccess::VbaValue::logical(true)); assert(toolbar->isEnabled());
    delete toolbar->actions()[0];
    bool deletedRejected = false;
    try { (void)command->member("Caption",{}); }
    catch(const linaccess::VbaRuntimeError& error) { deletedRejected = error.number==91; }
    assert(deletedRejected); assert(report.close());


    const QByteArray formatCode=R"VBA(
Sub Scan_Click()
Me!Summary = Format$("2003-12-31 18:05:06", "yyyy") & ":" & Format("2003-12-31", "mm")
End Sub
Sub Write_Click()
On Error Resume Next
Me!Summary = Format$(Null, "yyyy")
Me!Summary = CStr(Err.Number)
End Sub
)VBA";
    AccessFormPreview formatted("formatted",form,formatCode,db);
    formatted.findChild<QPushButton*>("access-control-Scan")->click();
    assert(formatted.findChild<QLineEdit*>("access-control-Summary")->text()=="2003:12");
    formatted.findChild<QPushButton*>("access-control-Write")->click();
    assert(formatted.findChild<QLineEdit*>("access-control-Summary")->text()=="94");
    assert(formatted.close());
    auto retainedField = firstObject->member("Controls",{linaccess::VbaValue::string("Name.Field")}).object;
    delete first.findChild<QLineEdit*>("access-control-Name.Field");
    bool deletedFieldRejected = false;
    try { (void)retainedField->member("Value",{}); }
    catch(const linaccess::VbaRuntimeError& error) { deletedFieldRejected = error.number==2465; }
    assert(deletedFieldRejected);
    assert(firstObject->member("Controls",{}).object->member("Count",{}).asNumber()==3);
    assert(first.close());
}
