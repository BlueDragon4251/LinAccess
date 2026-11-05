#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_query.hpp"
#include <QLineEdit>
#include <QComboBox>
#include <QToolBar>
#include <QAction>
#include <QMainWindow>
#include <QStatusBar>
#include <QProgressBar>
#include <QMessageBox>
#include <QMenuBar>
#include <QMenu>
#include <QPushButton>
#include <QSqlQuery>
#include <QApplication>
#include <QFileDialog>
#include <QFile>
#include <QTemporaryDir>
#include <QTimer>
#include <QFileInfo>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

namespace {
void testNativeTableDefaults() {
    auto db=QSqlDatabase::addDatabase("QSQLITE","form-table-defaults"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE defaults_rows(id INTEGER PRIMARY KEY AUTOINCREMENT,qty INTEGER DEFAULT 7,label TEXT DEFAULT 'table value',flag BOOLEAN DEFAULT 0)"));
    const QByteArray layout=R"LAYOUT(Begin Form
RecordSource ="defaults_rows"
OnCurrent ="[Event Procedure]"
Begin Section
Height =1200
Begin TextBox
Name ="qty"
ControlSource ="qty"
DefaultValue ="11"
End
Begin TextBox
Name ="label"
ControlSource ="label"
End
Begin CheckBox
Name ="flag"
ControlSource ="flag"
End
End
End
)LAYOUT";
    const QByteArray code=R"VBA(Sub Form_Current()
Me.Tag=CStr(Me!qty) & ":" & Me!label & ":" & CStr(Me!flag)
End Sub
Sub StoreDefaults()
Me!label="saved value"
Me.Dirty=False
End Sub
)VBA";
    AccessFormPreview preview("table defaults",layout,code,db); preview.show(); QApplication::processEvents();
    auto object=makePreviewFormObject(preview);
    assert(preview.hasDataSource());
    assert(object->member("Tag",{}).asText()=="11:table value:False");
    assert(!object->member("Dirty",{}).truthy());
    assert(sql.exec("SELECT COUNT(*) FROM defaults_rows") && sql.next() && sql.value(0).toInt()==0); sql.finish();
    std::vector<linaccess::VbaValue> args;
    assert(preview.executeProcedure("StoreDefaults",args).status==linaccess::VbaExecutionStatus::Completed);
    assert(sql.exec("SELECT qty,label,flag FROM defaults_rows") && sql.next());
    assert(sql.value(0).toInt()==11 && sql.value(1).toString()=="saved value" && sql.value(2).toInt()==0); sql.finish();
    assert(preview.close());
}
void testNativeStatusMeter() {
    auto db=QSqlDatabase::addDatabase("QSQLITE","native-status-meter"); db.setDatabaseName(":memory:"); assert(db.open());
    QMainWindow window; window.setCentralWidget(new QWidget(&window));
    const QByteArray source="Begin Form\nBegin Section\nHeight =600\nEnd\nEnd\n";
    const QByteArray module=R"VBA(Sub StartMeter()
If Not IsNull(SysCmd(Action:=acSysCmdInitMeter,Argument2:="Loading rows",Argument3:=200)) Then Err.Raise 5
End Sub
Sub AdvanceMeter()
SysCmd acSysCmdUpdateMeter,Argument2:=100
End Sub
Sub StatusMessage()
Application.SysCmd acSysCmdSetStatus,"Sorting rows"
End Sub
Sub ClearMessage()
SysCmd acSysCmdClearStatus
End Sub
Sub BusyProcedure()
DoCmd.Hourglass HourglassOn:=True
If Screen.MousePointer<>vbHourglass Then Err.Raise 5
NestedBusy
If Screen.MousePointer<>vbHourglass Then Err.Raise 5
End Sub
Sub NestedBusy()
If Screen.MousePointer<>vbHourglass Then Err.Raise 5
End Sub
Sub BusyFailure()
DoCmd.Hourglass True
Err.Raise 5
End Sub
Sub EchoOff()
Application.Echo EchoOn:=False,bstrStatusBarText:="Importing rows"
End Sub
Sub EchoOn()
DoCmd.Echo EchoOn:=True,StatusBarText:="Ready"
End Sub


)VBA";
    AccessFormPreview preview("native status",source,module,db,&window);
    window.show(); preview.show(); QApplication::processEvents();
    std::vector<linaccess::VbaValue> args;
    assert(preview.executeProcedure("StartMeter",args).status==linaccess::VbaExecutionStatus::Completed);
    auto* status=window.statusBar();
    auto* meter=status->findChild<QProgressBar*>("access-syscmd-meter");
    assert(meter && meter->maximum()==200 && meter->value()==0 && status->currentMessage()=="Loading rows");
    assert(preview.executeProcedure("AdvanceMeter",args).status==linaccess::VbaExecutionStatus::Completed && meter->value()==100);
    const auto application=makePreviewApplicationObject(preview);
    assert(!QApplication::overrideCursor());
    assert(preview.executeProcedure("BusyProcedure",args).status==linaccess::VbaExecutionStatus::Completed && !QApplication::overrideCursor());
    assert(preview.executeProcedure("BusyFailure",args).errorNumber==5 && !QApplication::overrideCursor());
    const auto screen=application->member("Screen",{}).object;
    screen->assign("MousePointer",linaccess::VbaValue::integer(11));
    assert(QApplication::overrideCursor() && screen->member("MousePointer",{}).asNumber()==11);
    screen->assign("MousePointer",linaccess::VbaValue::integer(0));
    assert(!QApplication::overrideCursor());
    QApplication::setOverrideCursor(QCursor(Qt::CrossCursor));
    int foreignCursorError=0;
    try { application->member("DoCmd",{}).object->member("Hourglass",{linaccess::VbaValue::logical(true)}); }
    catch(const linaccess::VbaRuntimeError& failure) { foreignCursorError=failure.number; }
    assert(foreignCursorError==3251 && QApplication::overrideCursor()->shape()==Qt::CrossCursor);
    QApplication::restoreOverrideCursor();

    assert(preview.executeProcedure("EchoOff",args).status==linaccess::VbaExecutionStatus::Completed);
    auto* selector=preview.findChild<QComboBox*>("access-record-selector"); assert(selector);
    assert(!window.centralWidget()->updatesEnabled() && !selector->updatesEnabled() && status->updatesEnabled());
    assert(status->currentMessage()=="Importing rows" && meter->value()==100);
    AccessFormPreview late("echo late",source,{},db,&window); late.show(); QApplication::processEvents();
    auto* lateSelector=late.findChild<QComboBox*>("access-record-selector"); assert(lateSelector && !lateSelector->updatesEnabled());
    QMessageBox message(QMessageBox::Information,"Native modal message","Synthetic modal message",QMessageBox::Ok,&window);
    QTimer::singleShot(0,&message,[&] { assert(message.updatesEnabled()); message.accept(); });
    message.exec();
    assert(preview.executeProcedure("EchoOn",args).status==linaccess::VbaExecutionStatus::Completed);
    assert(window.centralWidget()->updatesEnabled() && selector->updatesEnabled() && lateSelector->updatesEnabled());
    assert(status->currentMessage()=="Ready" && meter->value()==100);
    assert(late.close());
    window.centralWidget()->setUpdatesEnabled(false);
    application->member("Echo",{linaccess::VbaValue::logical(false)});
    application->member("Echo",{linaccess::VbaValue::logical(true)});
    assert(!window.centralWidget()->updatesEnabled()); window.centralWidget()->setUpdatesEnabled(true);

    int error=0;
    try { application->member("SysCmd",{linaccess::VbaValue::integer(2),linaccess::VbaValue::integer(201)}); }
    catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
    assert(error==5 && meter->value()==100);
    assert(preview.executeProcedure("StatusMessage",args).status==linaccess::VbaExecutionStatus::Completed);
    assert(!status->findChild<QProgressBar*>("access-syscmd-meter") && status->currentMessage()=="Sorting rows");
    assert(preview.executeProcedure("ClearMessage",args).status==linaccess::VbaExecutionStatus::Completed && status->currentMessage().isEmpty());
    assert(preview.close()); error=0;
    try { application->member("SysCmd",{linaccess::VbaValue::integer(5)}); }
    catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
    assert(error==2450);
}

void testNativeFilePicker() {
    auto db=QSqlDatabase::addDatabase("QSQLITE","native-file-picker"); db.setDatabaseName(":memory:"); assert(db.open());
    QTemporaryDir directory; assert(directory.isValid()); const auto path=directory.filePath("picker.txt");
    QFile file(path); assert(file.open(QIODevice::WriteOnly)); assert(file.write("synthetic picker fixture")==24); file.close();
    const QByteArray source="Begin Form\nBegin Section\nHeight =600\nEnd\nEnd\n";
    const QByteArray module=R"VBA(Sub PickFile()
Dim picker As Office.FileDialog, filters As Office.FileDialogFilters, entry As Office.FileDialogFilter
Dim selected As Office.FileDialogSelectedItems, item As Variant, count As Long
Set picker=Application.FileDialog(DialogType:=msoFileDialogFilePicker)
If Not TypeOf picker Is Office.FileDialog Then Err.Raise 5
picker.Title="Native file selection"
picker.AllowMultiSelect=False
picker.InitialFileName=Me.Tag
Set filters=picker.Filters
filters.Clear
filters.Add "Second","*.csv"
filters.Add "First","*.txt",Position:=1
count=0
For Each entry In filters
count=count+1
If Not TypeOf entry Is Office.FileDialogFilter Then Err.Raise 5
If count=1 And entry.Description<>"First" Then Err.Raise 5
If count=2 And entry.Description<>"Second" Then Err.Raise 5
Next
If count<>2 Then Err.Raise 5
picker.FilterIndex=2
filters.Delete filter:=1
If filters.Count<>1 Or filters(1).Description<>"Second" Or picker.FilterIndex<>1 Then Err.Raise 5
On Error Resume Next
filters.Delete filter:=0
count=Err.Number
On Error GoTo 0
If count<>5 Then Err.Raise 5
If filters.Count<>1 Then Err.Raise 5
filters.Delete filter:=1
If filters.Count<>0 Or picker.FilterIndex<>1 Then Err.Raise 5
filters.Clear
count=0
For Each entry In filters
Err.Raise 5
Next
filters.Add Description:="Text files",Extensions:="*.txt"
filters.Add Description:="Remove before Show",Extensions:="*.csv"
filters.Delete filter:=2
If filters.Count<>1 Then Err.Raise 5
Set entry=filters.Item(1)
If entry.Description<>"Text files" Or entry.Extensions<>"*.txt" Then Err.Raise 5
If picker.Show<>-1 Then Err.Raise 5
Set selected=picker.SelectedItems
If selected.Count<>1 Then Err.Raise 5
For Each item In selected
count=count+1
Me.Tag=item
Next
If count<>1 Or picker.SelectedItems(1)<>Me.Tag Then Err.Raise 5
End Sub
Sub CancelFile()
Dim picker As FileDialog
Set picker=Application.FileDialog(msoFileDialogFilePicker)
If picker.Title<>"Native file selection" Then Err.Raise 5
If picker.Show<>0 Then Err.Raise 5
If picker.SelectedItems.Count<>0 Then Err.Raise 5
End Sub
Sub PickFolder()
Dim picker As Office.FileDialog
Set picker=Application.FileDialog(msoFileDialogFolderPicker)
picker.InitialFileName=Me.Tag
If picker.DialogType<>4 Then Err.Raise 5
If picker.Show<>-1 Then Err.Raise 5
Me.Tag=picker.SelectedItems(1)
End Sub
)VBA";
    AccessFormPreview preview("native picker",source,module,db); preview.show(); QApplication::processEvents();
    auto object=makePreviewFormObject(preview); object->assign("Tag",linaccess::VbaValue::string(path.toStdString()));
    QTimer::singleShot(0,&preview,[path] {
        auto* dialog=qobject_cast<QFileDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="access-file-dialog");
        assert(dialog->windowTitle()=="Native file selection" && dialog->fileMode()==QFileDialog::ExistingFile);
        assert(dialog->nameFilters().size()==1 && dialog->nameFilters().front().contains("*.txt"));
        dialog->selectFile(path); assert(QMetaObject::invokeMethod(dialog,"accept",Qt::DirectConnection));
    });
    std::vector<linaccess::VbaValue> args;
    const auto picked=preview.executeProcedure("PickFile",args);
    assert(picked.status==linaccess::VbaExecutionStatus::Completed);
    assert(object->member("Tag",{}).asText()==path.toStdString());
    QTimer::singleShot(0,&preview,[] {
        auto* dialog=qobject_cast<QFileDialog*>(QApplication::activeModalWidget()); assert(dialog); dialog->reject();
    });
    assert(preview.executeProcedure("CancelFile",args).status==linaccess::VbaExecutionStatus::Completed);
    object->assign("Tag",linaccess::VbaValue::string(directory.path().toStdString()));
    QTimer::singleShot(0,&preview,[folder=directory.path()] {
        auto* dialog=qobject_cast<QFileDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->fileMode()==QFileDialog::Directory);
        dialog->setDirectory(QFileInfo(folder).absolutePath()); dialog->selectFile(folder);
        assert(QMetaObject::invokeMethod(dialog,"accept",Qt::DirectConnection));
    });
    assert(preview.executeProcedure("PickFolder",args).status==linaccess::VbaExecutionStatus::Completed);
    assert(QFileInfo(QString::fromStdString(object->member("Tag",{}).asText())).isDir());
    auto picker=makePreviewApplicationObject(preview)->member("FileDialog",{linaccess::VbaValue::longInteger(3)}).object;
    auto selected=picker->member("SelectedItems",{}).object;
    assert(preview.close()); int error=0;
    try { selected->member("Count",{}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
    assert(error==2450);
}
}

void testFormObjects() {
    testNativeTableDefaults();
    testNativeStatusMeter();
    testNativeFilePicker();
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
Set ctl = Forms("second form").Controls("Name.Field")
Let ctl = "changed"
Forms("second form").Dirty = False
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

    const QByteArray referenceCode=R"VBA(
Sub Scan_Click()
Dim ctl As Control
Dim same As Control
Set ctl = Me.Summary
Set same = (Me!Summary)
If Not (ctl Is same) Then
Err.Raise 5
End If
Me.Summary.Enabled = False
Me("Sum" & "mary").Visible = False
Forms![reference context]!Summary.Value = "live"
Set ctl = Forms![reference context]!Summary
ctl.Value = ctl.Value & ":reference"
Me!Summary.Visible = True
Me.Scan.Requery
End Sub
Sub Write_Click()
On Error Resume Next
Dim ctl As Control
Set ctl = Me.Caption
Me!Summary = CStr(Err.Number)
End Sub
)VBA";
    AccessFormPreview referenceContext("reference context",form,referenceCode,db);
    std::vector<linaccess::VbaValue> referenceArgs;
    const auto references=referenceContext.executeProcedure("Scan_Click",referenceArgs);
    assert(references.status==linaccess::VbaExecutionStatus::Completed);
    auto* referenceInput=referenceContext.findChild<QLineEdit*>("access-control-Summary");
    assert(referenceInput && referenceInput->text()=="live:reference" && !referenceInput->isEnabled() && !referenceInput->isHidden());
    assert(referenceContext.executeProcedure("Write_Click",referenceArgs).status==linaccess::VbaExecutionStatus::Completed);
    assert(referenceInput->text()=="13");
    assert(referenceContext.close());

    QMainWindow projectWindow;
    auto* projectToolbar=projectWindow.addToolBar("Project");
    auto* projectAction=projectToolbar->addAction("Open"); projectToolbar->setEnabled(false);
    projectToolbar->addSeparator();
    const QByteArray commandCode=R"VBA(
Sub Scan_Click()
Dim bar As Variant
Dim command As Variant
Dim count As Long
For Each bar In CommandBars
bar.Enabled = True
For Each command In bar.Controls
command.Caption = command.Caption & "!"
count = count + 1
Next command
Next bar
If Not (CommandBars(1) Is Application.CommandBars("Project")) Then
Err.Raise 5
End If
Me!Summary = CStr(count)
End Sub
Sub Write_Click()
CommandBars("Project").Visible = False
End Sub
)VBA";
    AccessFormPreview projectForm("toolbar context",form,commandCode,db,&projectWindow);
    assert(projectForm.executeProcedure("Scan_Click",referenceArgs).status==linaccess::VbaExecutionStatus::Completed);
    assert(projectToolbar->isEnabled() && projectAction->text()=="Open!");
    assert(projectForm.findChild<QLineEdit*>("access-control-Summary")->text()=="1");
    assert(projectForm.executeProcedure("Write_Click",referenceArgs).status==linaccess::VbaExecutionStatus::Completed);
    assert(projectToolbar->isHidden());
    auto projectBar=makePreviewApplicationObject(projectForm)->member("CommandBars",{linaccess::VbaValue::string("Project")}).object;
    auto projectCommands=projectBar->member("Controls",{}).object;
    delete projectToolbar;
    bool removedBar=false,removedControls=false;
    try { (void)projectBar->member("Enabled",{}); } catch(const linaccess::VbaRuntimeError& error) { removedBar=error.number==91; }
    try { (void)projectCommands->enumerate(); } catch(const linaccess::VbaRuntimeError& error) { removedControls=error.number==91; }
    assert(removedBar && removedControls); assert(projectForm.close());

    QMainWindow menuWindow;
    auto* menuBar=menuWindow.menuBar();
    auto* fileMenu=menuBar->addMenu("&Datei");
    auto* openAction=fileMenu->addAction("Öffnen");
    auto* sharedToolbar=menuWindow.addToolBar("Project"); sharedToolbar->addAction(openAction);
    int menuCommands=0;
    QObject::connect(openAction,&QAction::triggered,[&] { ++menuCommands; });
    const QByteArray menuCode=R"VBA(
Sub Scan_Click()
Dim bar As Variant
Dim popup As Variant
Set bar = CommandBars.ActiveMenuBar
If Not (bar Is CommandBars("Menu Bar")) Then
Err.Raise 5
End If
bar.Enabled = False
Set popup = bar.Controls(1)
If popup.Type <> 10 Then
Err.Raise 5
End If
popup.Controls(1).Caption = "Start"
popup.Controls(1).Enabled = False
Me!Summary = popup.Controls(1).Caption & ":" & CStr(CommandBars.Count)
End Sub
Sub Write_Click()
CommandBars.ActiveMenuBar.Enabled = True
CommandBars.ActiveMenuBar.Controls(1).Controls(1).Enabled = True
End Sub
)VBA";
    AccessFormPreview menuForm("native menus",form,menuCode,db,&menuWindow);
    assert(menuForm.executeProcedure("Scan_Click",referenceArgs).status==linaccess::VbaExecutionStatus::Completed);
    assert(!menuBar->isEnabled() && !openAction->isEnabled() && openAction->text()=="Start");
    assert(menuForm.findChild<QLineEdit*>("access-control-Summary")->text()=="Start:2");
    assert(menuForm.executeProcedure("Write_Click",referenceArgs).status==linaccess::VbaExecutionStatus::Completed);
    assert(menuBar->isEnabled() && openAction->isEnabled()); openAction->trigger(); assert(menuCommands==1);
    auto nativeBar=makePreviewApplicationObject(menuForm)->member("CommandBars",{}).object->member("ActiveMenuBar",{}).object;
    auto popup=nativeBar->member("Controls",{linaccess::VbaValue::numeric(1)}).object;
    auto popupItems=popup->member("Controls",{}).object;
    delete fileMenu;
    bool removedPopup=false;
    try { (void)popupItems->enumerate(); } catch(const linaccess::VbaRuntimeError& error) { removedPopup=error.number==91; }
    assert(removedPopup);
    delete menuBar;
    bool removedMenu=false;
    try { (void)nativeBar->member("Enabled",{}); } catch(const linaccess::VbaRuntimeError& error) { removedMenu=error.number==91; }
    assert(removedMenu && menuForm.close());

    AccessFormPreview parameters("sql parameters",form,{},db);
    auto input=makePreviewFormObject(parameters);
    input->assignItem("Summary",linaccess::VbaValue::numeric(2003));
    linaccess::AccessQueryModel sqlValues;
    sqlValues.execute(db,"SELECT Formulare![sql parameters]!Summary");
    assert(sqlValues.data(sqlValues.index(0,0)).toInt()==2003);
    input->assignItem("Summary",linaccess::VbaValue::string("002004"));
    sqlValues.execute(db,"SELECT Forms![sql parameters]!Summary.Value");
    assert(sqlValues.data(sqlValues.index(0,0)).toString()=="002004");
    AccessFormPreview otherParameters("sql parameters",form,{},separate);
    makePreviewFormObject(otherParameters)->assignItem("Summary",linaccess::VbaValue::numeric(9999));
    sqlValues.execute(db,"SELECT Forms![sql parameters]!Summary");
    assert(sqlValues.data(sqlValues.index(0,0)).toString()=="002004");
    sqlValues.execute(separate,"SELECT Forms![sql parameters]!Summary");
    assert(sqlValues.data(sqlValues.index(0,0)).toInt()==9999);
    assert(parameters.close());
    bool closedSqlRejected=false;
    try { sqlValues.execute(db,"SELECT Formulare![sql parameters]!Summary"); }
    catch(const linaccess::VbaRuntimeError& error) { closedSqlRejected=error.number==2450; }
    assert(closedSqlRejected); assert(otherParameters.close());
    auto retainedField = firstObject->member("Controls",{linaccess::VbaValue::string("Name.Field")}).object;
    delete first.findChild<QLineEdit*>("access-control-Name.Field");
    bool deletedFieldRejected = false;
    try { (void)retainedField->member("Value",{}); }
    catch(const linaccess::VbaRuntimeError& error) { deletedFieldRejected = error.number==2465; }
    assert(deletedFieldRejected);
    assert(firstObject->member("Controls",{}).object->member("Count",{}).asNumber()==3);
    assert(first.close());
}
