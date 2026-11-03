#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include <QLineEdit>
#include <QSqlQuery>
#include <QTimer>
#include <QPushButton>
#include <QTest>
#include <QLabel>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

int main(int argc,char** argv) {
    QApplication app(argc,argv); using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","current-record"); db.setDatabaseName(":memory:"); assert(db.open()); QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,label TEXT)")); assert(sql.exec("INSERT INTO records VALUES(10,'first'),(20,'second'),(30,'third')"));
    assert(sql.exec("CREATE TABLE positions(position INTEGER,type INTEGER)"));
    {
        const QByteArray buttons=R"FORM(Begin Form
Begin Section
Begin TextBox
Name ="input"
Begin Label
Name ="input_label"
Caption ="&Input"
End
End
Begin CommandButton
Name ="accept"
Caption ="Accept"
Default =-1
OnClick ="[Event Procedure]"
Top =500
End
Begin CommandButton
Name ="other"
Cancel =-1
Caption ="Other"
OnClick ="[Event Procedure]"
Top =1000
End
End
End
)FORM";
        const QByteArray handlers=R"VBA(Sub accept_Click()
CurrentDb.Execute "INSERT INTO positions VALUES(101,0)"
End Sub
Sub other_Click()
CurrentDb.Execute "INSERT INTO positions VALUES(102,0)"
End Sub
Sub SwitchDefault()
Me.other.Default=True
End Sub
Sub SwitchCancel()
Me.accept.Cancel=True
End Sub
)VBA";
        AccessFormPreview defaults("default_buttons",buttons,handlers,db); defaults.show(); QApplication::processEvents();
        auto* input=defaults.findChild<QLineEdit*>("access-control-input");
        auto* accept=defaults.findChild<QPushButton*>("access-control-accept");
        auto* other=defaults.findChild<QPushButton*>("access-control-other");
        assert(input && accept && other && accept->isDefault());
        auto* label=defaults.findChild<QLabel*>("access-control-input_label"); assert(label && label->buddy()==input);
        defaults.activateWindow(); other->setFocus(); QApplication::processEvents();
        QTest::keyClick(other,Qt::Key_I,Qt::AltModifier); QApplication::processEvents();
        assert(input->hasFocus());
        QTest::keyClicks(input,"discard"); QTest::keyClick(input,Qt::Key_Escape);
        assert(input->text().isEmpty() && defaults.isVisible());
        assert(sql.exec("SELECT COUNT(*) FROM positions WHERE position=102") && sql.next() && sql.value(0).toInt()==1); sql.finish();
        input->setFocus(); QTest::keyClick(input,Qt::Key_Return);
        assert(sql.exec("SELECT COUNT(*) FROM positions WHERE position=101") && sql.next() && sql.value(0).toInt()==1); sql.finish();
        auto controls=makePreviewFormObject(defaults)->member("Controls",{}).object;
        auto second=controls->member("Item",{VbaValue::string("other")}).object;
        std::vector<VbaValue> buttonArgs;
        assert(defaults.executeProcedure("SwitchDefault",buttonArgs).status==linaccess::VbaExecutionStatus::Completed);
        assert(!accept->isDefault() && other->isDefault());
        input->setFocus(); QTest::keyClick(input,Qt::Key_Return);
        assert(sql.exec("SELECT COUNT(*) FROM positions WHERE position=102") && sql.next() && sql.value(0).toInt()==2); sql.finish();
        accept->setFocus(); QTest::keyClick(accept,Qt::Key_Return);
        assert(sql.exec("SELECT COUNT(*) FROM positions WHERE position=101") && sql.next() && sql.value(0).toInt()==2); sql.finish();
        assert(defaults.executeProcedure("SwitchCancel",buttonArgs).status==linaccess::VbaExecutionStatus::Completed);
        assert(!second->member("Cancel",{}).truthy());
        other->setFocus(); QTest::keyClick(other,Qt::Key_Escape);
        assert(sql.exec("SELECT COUNT(*) FROM positions WHERE position=101") && sql.next() && sql.value(0).toInt()==3); sql.finish();
        assert(defaults.close()); assert(sql.exec("DELETE FROM positions"));
    }
    {
        assert(sql.exec("CREATE TABLE mode_records AS SELECT * FROM records"));
        const QByteArray readOnlyLayout=R"FORM(Begin Form
RecordSource ="mode_records"
AllowEdits =0
AllowAdditions =0
AllowDeletions =0
OnOpen ="[Event Procedure]"
Begin Section
Begin TextBox
Name ="label"
ControlSource ="label"
End
End
End
)FORM";
        sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('form','read_only_form',?)");
        sql.addBindValue(readOnlyLayout); assert(sql.exec()); sql.finish();
        sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('vba','Form_read_only_form',?)");
        sql.addBindValue(QByteArray("Sub Form_Open(Cancel As Integer)\nCurrentDb.Execute \"INSERT INTO positions VALUES(201,\" & CLng(Me.AllowEdits Or Me.AllowAdditions Or Me.AllowDeletions) & \")\"\nEnd Sub\n"));
        assert(sql.exec()); sql.finish();
        const QByteArray opener=R"VBA(Sub ReadOnlyOpen()
DoCmd.OpenForm FormName:="read_only_form",DataMode:=acFormReadOnly
End Sub
Sub EditableOpen()
DoCmd.OpenForm FormName:="read_only_form",DataMode:=acFormEdit
End Sub
Sub ReadOnlyClose()
DoCmd.Close ObjectType:=acForm,ObjectName:="read_only_form",Save:=acSaveNo
End Sub
)VBA";
        AccessFormPreview caller("readonly_caller","Begin Form\nBegin Section\nEnd\nEnd\n",opener,db);
        std::vector<VbaValue> args; assert(caller.executeProcedure("ReadOnlyOpen",args).status==linaccess::VbaExecutionStatus::Completed);
        assert(sql.exec("SELECT type FROM positions WHERE position=201") && sql.next() && sql.value(0).toInt()==0); sql.finish();
        AccessFormPreview* opened=nullptr;
        for(auto* widget:QApplication::topLevelWidgets())
            if(auto* candidate=dynamic_cast<AccessFormPreview*>(widget);candidate && candidate!=&caller && candidate->isVisible() && candidate->findChild<QLineEdit*>("access-control-label")) opened=candidate;
        assert(opened && opened->findChild<QLineEdit*>("access-control-label")->isReadOnly());
        auto form=makePreviewFormObject(*opened);
        assert(!form->member("AllowEdits",{}).truthy() && !form->member("AllowAdditions",{}).truthy() && !form->member("AllowDeletions",{}).truthy());
        assert(caller.executeProcedure("ReadOnlyClose",args).status==linaccess::VbaExecutionStatus::Completed);
        assert(caller.executeProcedure("EditableOpen",args).status==linaccess::VbaExecutionStatus::Completed);
        opened=nullptr;
        for(auto* widget:QApplication::topLevelWidgets())
            if(auto* candidate=dynamic_cast<AccessFormPreview*>(widget);candidate && candidate!=&caller && candidate->isVisible() && candidate->findChild<QLineEdit*>("access-control-label")) opened=candidate;
        assert(opened); auto* editor=opened->findChild<QLineEdit*>("access-control-label"); assert(!editor->isReadOnly());
        form=makePreviewFormObject(*opened);
        assert(form->member("AllowEdits",{}).truthy() && form->member("AllowAdditions",{}).truthy() && form->member("AllowDeletions",{}).truthy());
        editor->setFocus(); editor->selectAll(); QTest::keyClicks(editor,"edited");
        form->assign("Dirty",VbaValue::logical(false));
        assert(sql.exec("SELECT label FROM mode_records WHERE id=10") && sql.next() && sql.value(0).toString()=="edited"); sql.finish();
        assert(caller.executeProcedure("ReadOnlyClose",args).status==linaccess::VbaExecutionStatus::Completed);
        assert(caller.close()); assert(sql.exec("DELETE FROM positions"));
    }
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="records"
OnCurrent ="[Event Procedure]"
Begin Section
Begin TextBox
Name ="label"
ControlSource ="label"
Width =4000
End
Begin TextBox
Name ="position"
ControlSource ="=CurrentRecord"
Top =500
Width =4000
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub Form_Current()
CurrentDb.Execute "INSERT INTO positions VALUES(" & CStr(Me.CurrentRecord) & "," & CStr(VarType(Me.CurrentRecord)) & ")"
End Sub
Function PositionValue()
PositionValue=CurrentRecord
End Function
Function ThroughCollections()
ThroughCollections=Forms("current_record").CurrentRecord+Screen.ActiveForm.CurrentRecord
End Function
Sub NextRow()
DoCmd.RunCommand 65
End Sub
Sub LastRow()
DoCmd.RunCommand 68
End Sub
Sub NewRow()
DoCmd.RunCommand 28
End Sub
Function SortRows()
OrderBy="[records].[id] DESC"
OrderByOn=True
SortRows=Me.CurrentRecord
End Function
Function SortAscending()
DoCmd.RunCommand acCmdSortAscending
SortAscending=Me.OrderBy
End Function
Function SortDescending()
DoCmd.RunCommand acCmdSortDescending
SortDescending=Me.OrderBy
End Function
Sub WritePosition()
Me.CurrentRecord=1
End Sub
Sub OpenIcon()
DoCmd.OpenForm FormName:="named_icon",WindowMode:=acIcon,OpenArgs:="icon"
End Sub
Sub CloseIcon()
DoCmd.Close ObjectName:="named_icon",ObjectType:=acForm,Save:=acSaveNo
End Sub
Sub MinimizeWindow()
DoCmd.Minimize
End Sub
Sub RestoreWindow()
DoCmd.Restore
End Sub
Sub MaximizeWindow()
DoCmd.Maximize
End Sub
Function DialogSequence() As String
CurrentDb.Execute "INSERT INTO dialog_events(phase,mode) VALUES('Before','close')"
DoCmd.OpenForm FormName:="named_dialog",WindowMode:=acDialog,OpenArgs:="close"
CurrentDb.Execute "INSERT INTO dialog_events(phase,mode) VALUES('After','close')"
CurrentDb.Execute "INSERT INTO dialog_events(phase,mode) VALUES('Before','hide')"
DoCmd.OpenForm FormName:="named_dialog",WindowMode:=acDialog,OpenArgs:="hide"
DialogSequence=CStr(Forms("named_dialog").Modal) & ":" & CStr(Forms("named_dialog").PopUp) & ":" & CStr(Forms("named_dialog").Visible)
CurrentDb.Execute "INSERT INTO dialog_events(phase,mode) VALUES('After','hide')"
DoCmd.Close ObjectName:="named_dialog",ObjectType:=acForm,Save:=acSaveNo
CurrentDb.Execute "INSERT INTO dialog_events(phase,mode) VALUES('Before','cancel')"
DoCmd.OpenForm FormName:="named_dialog",WindowMode:=acDialog,OpenArgs:="cancel"
CurrentDb.Execute "INSERT INTO dialog_events(phase,mode) VALUES('After','cancel')"
End Function
Function RunSqlNull() As Long
On Error Resume Next
DoCmd.RunSQL SQLStatement:=Null
If Err.Number=94 Then
RunSqlNull=RunSqlNull+1
End If
Err.Clear
DoCmd.RunSQL Null
If Err.Number=94 Then
RunSqlNull=RunSqlNull+1
End If
End Function
Function RunSqlTransactions() As Long
On Error Resume Next
DoCmd.RunSQL SQLStatement:="INSERT INTO named_unique VALUES(2),(1)",UseTransaction:=True
RunSqlTransactions=Err.Number
On Error GoTo 0
DoCmd.RunSQL "INSERT INTO named_unique VALUES(3)",False
DBEngine(0).BeginTrans
DoCmd.RunSQL "INSERT INTO named_unique VALUES(4)"
On Error Resume Next
DoCmd.RunSQL SQLStatement:="INSERT INTO named_unique VALUES(5),(1)",UseTransaction:=True
On Error GoTo 0
DBEngine(0).Rollback
End Function
Function NamedCommands() As String
DoCmd.RunSQL SQLStatement:="INSERT INTO named_calls VALUES(1)",UseTransaction:=True
DoCmd.RunSQL "INSERT INTO named_calls VALUES(2)",False
DoCmd.OpenForm WhereCondition:="[id]=20", FormName:="named_child", OpenArgs:="payload", DataMode:=acFormPropertySettings, WindowMode:=acWindowNormal
NamedCommands=Forms("named_child").OpenArgs & ":" & Forms("named_child")!label
DoCmd.Close Save:=2,ObjectName:="named_child",ObjectType:=2
DoCmd.OpenForm FormName:="named_child",OpenArgs:="new"
NamedCommands=NamedCommands & "|" & Forms("named_child").OpenArgs
DoCmd.Close Save:=2
DoCmd.OpenReport OpenArgs:="report-value",ReportName:="named_report",View:=acViewPreview
NamedCommands=NamedCommands & "|" & Reports("named_report").OpenArgs
DoCmd.Close Save:=acSaveNo,ObjectName:="named_report",ObjectType:=acReport
DoCmd.OpenForm FormName:="named_hidden",WindowMode:=acHidden,OpenArgs:="hidden"
NamedCommands=NamedCommands & "|" & CStr(Forms("named_hidden").Visible) & ":" & Forms("named_hidden").OpenArgs
CurrentDb.Execute "INSERT INTO open_arg_events VALUES('BeforeReveal','hidden')"
Forms("named_hidden").Visible=True
NamedCommands=NamedCommands & ":" & CStr(Forms("named_hidden").Visible)
DoCmd.OpenForm FormName:="named_hidden",WindowMode:=acHidden
NamedCommands=NamedCommands & ":" & CStr(Forms("named_hidden").Visible)
DoCmd.Close ObjectName:="named_hidden",ObjectType:=acForm,Save:=acSaveNo
End Function
)VBA";
    AccessFormPreview form("current_record",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents(); auto model=makePreviewFormObject(form);
    const auto position=[&](int wanted) { const auto value=model->member("CurrentRecord",{}); assert(value.kind==VbaValue::Kind::Number && value.numericType==VbaValue::NumericType::Long && value.asNumber()==wanted); auto* editor=form.findChild<QLineEdit*>("access-control-position"); assert(editor && editor->text()==QString::number(wanted)); };
    const auto run=[&](const char* name) { std::vector<VbaValue> args; return form.executeProcedure(name,args); };
    position(1); auto result=run("PositionValue"); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->kind==VbaValue::Kind::Number && result.typedReturnValue->numericType==VbaValue::NumericType::Long && result.typedReturnValue->asNumber()==1);
    result=run("ThroughCollections"); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==2);
    assert(run("NextRow").status==linaccess::VbaExecutionStatus::Completed); position(2); assert(model->item("id").asNumber()==20);
    model->assignItem("label",VbaValue::string("pending")); position(2); assert(form.hasUnsavedChanges());
    assert(sql.exec("SELECT label FROM records WHERE id=20") && sql.next() && sql.value(0).toString()=="second"); sql.finish();
    result=run("WritePosition"); assert(result.status==linaccess::VbaExecutionStatus::Error && result.errorNumber==438); position(2); assert(form.hasUnsavedChanges()); model->member("Undo",{});
    assert(run("LastRow").status==linaccess::VbaExecutionStatus::Completed); position(3);
    model->assign("OrderBy",VbaValue::string("[id] DESC")); assert(!model->member("OrderByOn",{}).truthy()); position(3);
    result=run("SortRows"); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->asNumber()==1); position(1); assert(model->item("id").asNumber()==30);
    int badOrder=0; try { model->assign("OrderBy",VbaValue::string("missing DESC")); } catch(const linaccess::VbaRuntimeError& error) { badOrder=error.number; }
    assert(badOrder==3251 && model->member("OrderBy",{}).asText()=="[records].[id] DESC" && model->member("OrderByOn",{}).truthy()); position(1);
    model->assign("OrderByOn",VbaValue::logical(false)); position(3);
    auto* activeField=form.findChild<QLineEdit*>("access-control-label"); activeField->setFocus(); QApplication::processEvents(); assert(QApplication::focusWidget()==activeField);
    result=run("SortAscending"); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->asText()=="[label] ASC"); position(3);
    result=run("SortDescending"); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->asText()=="[label] DESC"); position(1);
    model->assign("OrderByOn",VbaValue::logical(false)); position(3);
    model->assign("Filter",VbaValue::string("id>=20")); model->assign("FilterOn",VbaValue::logical(true)); position(2);
    assert(run("NewRow").status==linaccess::VbaExecutionStatus::Completed); position(3); assert(model->member("NewRecord",{}).truthy() && !form.hasUnsavedChanges());
    assert(sql.exec("SELECT COUNT(*) FROM records") && sql.next() && sql.value(0).toInt()==3); sql.finish();
    assert(sql.exec("SELECT COUNT(*),MIN(type),MAX(type) FROM positions") && sql.next() && sql.value(0).toInt()>=5 && sql.value(1).toInt()==3 && sql.value(2).toInt()==3); sql.finish();
    assert(form.close()); int closed=0; try { model->member("CurrentRecord",{}); } catch(const linaccess::VbaRuntimeError& error) { closed=error.number; } assert(closed==2450);
    const QByteArray sortedLayout="Begin Form\nRecordSource =\"records\"\nOrderBy =\"[id] DESC\"\nOrderByOn =-1\nEnd\n";
    AccessFormPreview sorted("sorted_on_open",sortedLayout,{},db); auto sortedModel=makePreviewFormObject(sorted);
    assert(sortedModel->member("OrderByOn",{}).truthy() && sortedModel->item("id").asNumber()==30 && sorted.close());
    const QByteArray queryLayout="Begin Form\nRecordSource =\"SELECT id,label FROM records ORDER BY id DESC\"\nEnd\n";
    AccessFormPreview queryForm("sorted_query",queryLayout,{},db); auto queryModel=makePreviewFormObject(queryForm);
    assert(queryModel->item("id").asNumber()==30); queryModel->assign("OrderBy",VbaValue::string("id ASC")); queryModel->assign("OrderByOn",VbaValue::logical(true));
    assert(queryModel->item("id").asNumber()==10 && queryModel->member("CurrentRecord",{}).asNumber()==1 && queryForm.close());
    assert(sql.exec("CREATE TABLE named_calls(value INTEGER)"));
    assert(sql.exec("CREATE TABLE named_unique(value INTEGER UNIQUE)")); assert(sql.exec("INSERT INTO named_unique VALUES(1)"));
    assert(sql.exec("CREATE TABLE open_arg_events(phase TEXT,value TEXT)"));
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('form','named_child',?)");
    auto childLayout=layout; childLayout.replace("OnCurrent =\"[Event Procedure]\"","OnOpen =\"[Event Procedure]\"\nOnLoad =\"[Event Procedure]\"");
    sql.addBindValue(childLayout); assert(sql.exec()); sql.finish();
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('vba','Form_named_child',?)");
    sql.addBindValue(QByteArray(R"VBA(Sub Form_Open(Cancel As Integer)
CurrentDb.Execute "INSERT INTO open_arg_events VALUES('Open','" & Me.OpenArgs & "')"
End Sub
Sub Form_Load()
CurrentDb.Execute "INSERT INTO open_arg_events VALUES('Load','" & OpenArgs & "')"
End Sub
)VBA")); assert(sql.exec()); sql.finish();
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('report','named_report',?)");
    sql.addBindValue(QByteArray("Begin Report\nRecordSource =\"records\"\nOnOpen =\"[Event Procedure]\"\nEnd\n")); assert(sql.exec()); sql.finish();
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('vba','Report_named_report',?)");
    sql.addBindValue(QByteArray(R"VBA(Sub Report_Open(Cancel As Integer)
CurrentDb.Execute "INSERT INTO open_arg_events VALUES('ReportOpen','" & Me.OpenArgs & "')"
End Sub
)VBA")); assert(sql.exec()); sql.finish();
    auto hiddenLayout=childLayout;
    hiddenLayout.replace("OnLoad =\"[Event Procedure]\"","OnLoad =\"[Event Procedure]\"\nOnCurrent =\"[Event Procedure]\"\nOnActivate =\"[Event Procedure]\"");
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('form','named_hidden',?)");
    sql.addBindValue(hiddenLayout); assert(sql.exec()); sql.finish();
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('vba','Form_named_hidden',?)");
    sql.addBindValue(QByteArray(R"VBA(Sub Form_Open(Cancel As Integer)
CurrentDb.Execute "INSERT INTO open_arg_events VALUES('HiddenOpen','" & Me.OpenArgs & "')"
End Sub
Sub Form_Load()
CurrentDb.Execute "INSERT INTO open_arg_events VALUES('HiddenLoad','" & Me.OpenArgs & "')"
End Sub
Sub Form_Current()
CurrentDb.Execute "INSERT INTO open_arg_events VALUES('HiddenCurrent','" & Me.OpenArgs & "')"
End Sub
Sub Form_Activate()
CurrentDb.Execute "INSERT INTO open_arg_events VALUES('HiddenActivate','" & Me.OpenArgs & "')"
End Sub
)VBA")); assert(sql.exec()); sql.finish();
    assert(sql.exec("CREATE TABLE dialog_events(phase TEXT,mode TEXT,modal INTEGER,popup INTEGER)"));
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('form','named_dialog',?)");
    sql.addBindValue(QByteArray("Begin Form\nRecordSource =\"records\"\nOnOpen =\"[Event Procedure]\"\nOnLoad =\"[Event Procedure]\"\nOnActivate =\"[Event Procedure]\"\nOnCurrent =\"[Event Procedure]\"\nOnTimer =\"[Event Procedure]\"\nTimerInterval =25\nOnUnload =\"[Event Procedure]\"\nEnd\n")); assert(sql.exec()); sql.finish();
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('vba','Form_named_dialog',?)");
    sql.addBindValue(QByteArray(R"VBA(Private attempts As Long
Sub Form_Open(Cancel As Integer)
CurrentDb.Execute "INSERT INTO dialog_events VALUES('Open','" & OpenArgs & "'," & CStr(CLng(Me.Modal)) & "," & CStr(CLng(Me.PopUp)) & ")"
End Sub
Sub Form_Load()
CurrentDb.Execute "INSERT INTO dialog_events(phase,mode) VALUES('Load','" & OpenArgs & "')"
End Sub
Sub Form_Activate()
CurrentDb.Execute "INSERT INTO dialog_events(phase,mode) VALUES('Activate','" & OpenArgs & "')"
End Sub
Sub Form_Current()
CurrentDb.Execute "INSERT INTO dialog_events(phase,mode) VALUES('Current','" & OpenArgs & "')"
End Sub
Sub Form_Timer()
On Error Resume Next
attempts=attempts+1
CurrentDb.Execute "INSERT INTO dialog_events(phase,mode) VALUES('Timer','" & OpenArgs & "')"
If OpenArgs="hide" Then
Me.Visible=False
Else
DoCmd.Close
End If
End Sub
Sub Form_Unload(Cancel As Integer)
If OpenArgs="cancel" And attempts=1 Then
Cancel=1
CurrentDb.Execute "INSERT INTO dialog_events(phase,mode) VALUES('UnloadCancel','" & OpenArgs & "')"
Else
CurrentDb.Execute "INSERT INTO dialog_events(phase,mode) VALUES('Unload','" & OpenArgs & "')"
End If
End Sub
)VBA")); assert(sql.exec()); sql.finish();
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('form','named_icon',?)");
    sql.addBindValue(QByteArray("Begin Form\nRecordSource =\"records\"\nEnd\n")); assert(sql.exec()); sql.finish();
    AccessFormPreview namedForm("named_parent",layout,code,db); namedForm.show(); namedForm.activateWindow(); QApplication::processEvents();
    std::vector<VbaValue> namedArgs; const auto named=namedForm.executeProcedure("NamedCommands",namedArgs);
    if(named.status!=linaccess::VbaExecutionStatus::Completed) qWarning("Named DoCmd error=%d line=%d",named.errorNumber,named.line);
    assert(named.status==linaccess::VbaExecutionStatus::Completed && named.typedReturnValue->asText()=="payload:second|new|report-value|False:hidden:True:False");
    assert(sql.exec("SELECT COUNT(*) FROM named_calls") && sql.next() && sql.value(0).toInt()==2); sql.finish();
    const auto nullSql=namedForm.executeProcedure("RunSqlNull",namedArgs);
    assert(nullSql.status==linaccess::VbaExecutionStatus::Completed && nullSql.typedReturnValue->asNumber()==2);
    const auto transactions=namedForm.executeProcedure("RunSqlTransactions",namedArgs);
    if(transactions.status!=linaccess::VbaExecutionStatus::Completed) qWarning("RunSQL transaction error=%d line=%d",transactions.errorNumber,transactions.line);
    assert(transactions.status==linaccess::VbaExecutionStatus::Completed && transactions.typedReturnValue->asNumber()==3022);
    assert(sql.exec("SELECT COUNT(*),SUM(value) FROM named_unique") && sql.next() && sql.value(0).toInt()==2 && sql.value(1).toInt()==4); sql.finish();
    assert(makePreviewFormObject(namedForm)->member("OpenArgs",{}).kind==VbaValue::Kind::Null);
    assert(sql.exec("SELECT COUNT(*) FROM open_arg_events WHERE value IN('payload','new')") && sql.next() && sql.value(0).toInt()==4); sql.finish();
    assert(sql.exec("SELECT COUNT(*) FROM open_arg_events WHERE phase='ReportOpen' AND value='report-value'") && sql.next() && sql.value(0).toInt()==1); sql.finish();
    assert(sql.exec("SELECT group_concat(phase,',') FROM (SELECT phase FROM open_arg_events WHERE value='hidden' ORDER BY rowid)") && sql.next() && sql.value(0).toString()=="HiddenOpen,HiddenLoad,HiddenCurrent,BeforeReveal,HiddenActivate"); sql.finish();
    int readOnly=0; try { makePreviewFormObject(namedForm)->assign("OpenArgs",VbaValue::string("changed")); } catch(const linaccess::VbaRuntimeError& failure) { readOnly=failure.number; } assert(readOnly==438);
    QTimer deadline; deadline.setSingleShot(true); QObject::connect(&deadline,&QTimer::timeout,[] { qFatal("Native dialog execution deadline exceeded"); }); deadline.start(3000);
    const auto dialogs=namedForm.executeProcedure("DialogSequence",namedArgs); deadline.stop();
    if(dialogs.status!=linaccess::VbaExecutionStatus::Completed) qWarning("Dialog error=%d line=%d",dialogs.errorNumber,dialogs.line);
    assert(dialogs.status==linaccess::VbaExecutionStatus::Completed && dialogs.typedReturnValue->asText()=="True:True:False");
    for(const auto& mode:{QString("close"),QString("hide"),QString("cancel")}) {
        sql.prepare("SELECT group_concat(phase,',') FROM (SELECT phase FROM dialog_events WHERE mode=? ORDER BY rowid)"); sql.addBindValue(mode); assert(sql.exec() && sql.next());
        const auto wanted=mode=="close" ? "Before,Open,Load,Activate,Current,Timer,Unload,After" : mode=="hide" ? "Before,Open,Load,Activate,Current,Timer,After,Unload" : "Before,Open,Load,Activate,Current,Timer,UnloadCancel,Timer,Unload,After";
        assert(sql.value(0).toString()==wanted); sql.finish();
    }
    assert(sql.exec("SELECT COUNT(*) FROM dialog_events WHERE phase='Open' AND modal=-1 AND popup=-1") && sql.next() && sql.value(0).toInt()==3); sql.finish();
    namedForm.activateWindow(); QApplication::setActiveWindow(&namedForm); QApplication::processEvents();
    assert(namedForm.executeProcedure("OpenIcon",namedArgs).status==linaccess::VbaExecutionStatus::Completed);
    QPointer<AccessFormPreview> icon;
    for(auto* widget:QApplication::topLevelWidgets()) if(auto* candidate=dynamic_cast<AccessFormPreview*>(widget); candidate && candidate->isMinimized() && candidate!=&namedForm) {
        if(makePreviewFormObject(*candidate)->member("Name",{}).asText()=="named_icon") icon=candidate;
    }
    assert(icon && icon->isMinimized() && !namedForm.isMinimized());
    assert(makePreviewFormObject(*icon)->member("OpenArgs",{}).asText()=="icon");
    assert(namedForm.executeProcedure("RestoreWindow",namedArgs).status==linaccess::VbaExecutionStatus::Completed && !icon->isMinimized());
    assert(namedForm.executeProcedure("MinimizeWindow",namedArgs).status==linaccess::VbaExecutionStatus::Completed && icon->isMinimized() && !namedForm.isMinimized());
    assert(namedForm.executeProcedure("CloseIcon",namedArgs).status==linaccess::VbaExecutionStatus::Completed);
    namedForm.activateWindow(); QApplication::setActiveWindow(&namedForm); QApplication::processEvents();
    assert(namedForm.executeProcedure("MinimizeWindow",namedArgs).status==linaccess::VbaExecutionStatus::Completed && namedForm.isMinimized());
    assert(namedForm.executeProcedure("RestoreWindow",namedArgs).status==linaccess::VbaExecutionStatus::Completed && !namedForm.isMinimized() && !namedForm.isMaximized());
    assert(namedForm.executeProcedure("MaximizeWindow",namedArgs).status==linaccess::VbaExecutionStatus::Completed && namedForm.isMaximized());
    assert(namedForm.executeProcedure("RestoreWindow",namedArgs).status==linaccess::VbaExecutionStatus::Completed && !namedForm.isMinimized() && !namedForm.isMaximized());
    assert(namedForm.close());
    const QByteArray report="Begin Report\nRecordSource =\"records\"\nEnd\n"; AccessFormPreview reportForm("position_report",report,{},db); auto reportModel=makePreviewFormObject(reportForm);
    int unsupported=0; try { reportModel->member("CurrentRecord",{}); } catch(const linaccess::VbaRuntimeError& error) { unsupported=error.number; } assert(unsupported==438); assert(reportForm.close());
    assert(sql.exec("CREATE TABLE quit_steps(phase TEXT,code INTEGER)"));
    const QByteArray quitChildLayout="Begin Form\nOnTimer =\"[Event Procedure]\"\nTimerInterval =25\nEnd\n";
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('form','quit_cancel_child',?)"); sql.addBindValue(quitChildLayout); assert(sql.exec()); sql.finish();
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('vba','Form_quit_cancel_child',?)");
    sql.addBindValue(QByteArray(R"VBA(Private attempts As Long
Sub Form_Timer()
On Error Resume Next
attempts=attempts+1
If attempts=1 Then
DoCmd.Quit
CurrentDb.Execute "INSERT INTO quit_steps VALUES('cancel-number'," & CStr(Err.Number) & ")"
Else
DoCmd.Close
End If
End Sub
)VBA")); assert(sql.exec()); sql.finish();
    const QByteArray quitParentLayout="Begin Form\nOnUnload =\"[Event Procedure]\"\nEnd\n";
    const QByteArray cancelQuitCode=R"VBA(Private preventQuit As Boolean
Sub Form_Unload(Cancel As Integer)
If preventQuit Then
Cancel=1
End If
End Sub
Sub CancelledQuit()
preventQuit=True
DoCmd.OpenForm FormName:="quit_cancel_child",WindowMode:=acDialog
preventQuit=False
CurrentDb.Execute "INSERT INTO quit_steps VALUES('after-cancel',0)"
End Sub
)VBA";
    AccessFormPreview cancelQuitParent("quit_cancel_parent",quitParentLayout,cancelQuitCode,db); cancelQuitParent.show();
    QTimer quitDeadline; quitDeadline.setSingleShot(true); QObject::connect(&quitDeadline,&QTimer::timeout,[] { qFatal("Native Quit dialog execution deadline exceeded"); });
    quitDeadline.start(3000); const auto cancelledQuit=cancelQuitParent.executeProcedure("CancelledQuit",namedArgs); quitDeadline.stop();
    if(cancelledQuit.status!=linaccess::VbaExecutionStatus::Completed) qWarning("Cancelled Quit error=%d line=%d",cancelledQuit.errorNumber,cancelledQuit.line);
    assert(cancelledQuit.status==linaccess::VbaExecutionStatus::Completed && !cancelledQuit.exitSub);
    assert(sql.exec("SELECT COUNT(*) FROM quit_steps WHERE phase='cancel-number' AND code=2501") && sql.next() && sql.value(0).toInt()==1); sql.finish();
    assert(sql.exec("SELECT COUNT(*) FROM quit_steps WHERE phase='after-cancel'") && sql.next() && sql.value(0).toInt()==1); sql.finish();
    assert(cancelQuitParent.close());
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('form','quit_child',?)"); sql.addBindValue(quitChildLayout); assert(sql.exec()); sql.finish();
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('vba','Form_quit_child',?)");
    sql.addBindValue(QByteArray(R"VBA(Sub Form_Timer()
Application.Quit(Option:=acQuitSaveAll)
CurrentDb.Execute "INSERT INTO quit_steps VALUES('child-after',0)"
End Sub
)VBA")); assert(sql.exec()); sql.finish();
    const QByteArray quitCode=R"VBA(Sub QuitFromDialog()
On Error Resume Next
CurrentDb.Execute "INSERT INTO quit_steps VALUES('before',0)"
DoCmd.OpenForm FormName:="quit_child",WindowMode:=acDialog
CurrentDb.Execute "INSERT INTO quit_steps VALUES('after',0)"
End Sub
)VBA";
    AccessFormPreview quitParent("quit_parent",QByteArray("Begin Form\nEnd\n"),quitCode,db); quitParent.show();
    quitDeadline.start(3000); const auto quit=quitParent.executeProcedure("QuitFromDialog",namedArgs); quitDeadline.stop();
    if(quit.status!=linaccess::VbaExecutionStatus::Completed) qWarning("Quit propagation error=%d line=%d",quit.errorNumber,quit.line);
    assert(quit.status==linaccess::VbaExecutionStatus::Completed && quit.exitSub);
    assert(sql.exec("SELECT COUNT(*) FROM quit_steps WHERE phase='before'") && sql.next() && sql.value(0).toInt()==1); sql.finish();
    assert(sql.exec("SELECT COUNT(*) FROM quit_steps WHERE phase IN('after','child-after')") && sql.next() && sql.value(0).toInt()==0); sql.finish();

    const QByteArray startupLayout="Begin Form\nOnOpen =\"[Event Procedure]\"\nOnLoad =\"[Event Procedure]\"\nEnd\n";
    const QByteArray openQuit=R"VBA(Sub Form_Open(Cancel As Integer)
CurrentDb.Execute "INSERT INTO quit_steps VALUES('startup-open-before',0)"
DoCmd.Quit
CurrentDb.Execute "INSERT INTO quit_steps VALUES('startup-open-after',0)"
End Sub
Sub Form_Load()
CurrentDb.Execute "INSERT INTO quit_steps VALUES('startup-load-forbidden',0)"
End Sub
)VBA";
    assert(!AccessFormPreview::showDocument(linaccess::AccessDocumentKind::Form,"startup_open_quit",startupLayout,openQuit,db));
    const QByteArray loadQuit=R"VBA(Sub Form_Open(Cancel As Integer)
CurrentDb.Execute "INSERT INTO quit_steps VALUES('startup-open-normal',0)"
End Sub
Sub Form_Load()
CurrentDb.Execute "INSERT INTO quit_steps VALUES('startup-load-before',0)"
DoCmd.Quit()
CurrentDb.Execute "INSERT INTO quit_steps VALUES('startup-load-after',0)"
End Sub
)VBA";
    AccessFormPreview startupLoad("startup_load_quit",startupLayout,loadQuit,db);
    assert(startupLoad.cancelledByOpen() && startupLoad.isHidden());
    assert(sql.exec("SELECT COUNT(*) FROM quit_steps WHERE phase IN('startup-open-before','startup-open-normal','startup-load-before')") && sql.next() && sql.value(0).toInt()==3); sql.finish();
    assert(sql.exec("SELECT COUNT(*) FROM quit_steps WHERE phase IN('startup-open-after','startup-load-forbidden','startup-load-after')") && sql.next() && sql.value(0).toInt()==0); sql.finish();

}
