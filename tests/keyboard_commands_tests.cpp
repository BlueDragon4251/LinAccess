#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "vba_analyzer.hpp"
#include "vba_source_codec.hpp"
#include <QCryptographicHash>
#include <QDir>
#include <QFile>
#include <QJsonArray>
#include <QJsonObject>
#include <QMainWindow>
#include <QMenuBar>
#include <QToolBar>
#include <QTimer>
#include <QMessageBox>
#include <QKeyEvent>
#include <iostream>
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
#include <algorithm>

void testKeyboardEvents();
void testTextSelection();

void testKeyboardCommands() {
    testTextSelection();
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","keyboard-commands"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,first TEXT,second TEXT)"));
    assert(query.exec("INSERT INTO records VALUES(1,'before',NULL)"));
    assert(query.exec("CREATE TABLE events(kind TEXT)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="records"
BeforeUpdate ="[Event Procedure]"
OnUnload ="[Event Procedure]"
Begin Section
Height =3000
Begin TextBox
Name ="first"
ControlSource ="first"
End
Begin TextBox
Name ="second"
Top =400
ControlSource ="second"
End
Begin TextBox
Name ="counter"
ControlSource ="=DCount(""*"",""records"")"
End
Begin CommandButton
Name ="action"
OnClick ="[Event Procedure]"
Top =1000
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub Form_BeforeUpdate(Cancel As Integer)
If Me.Tag="blocksave" Then Cancel=1
End Sub
Sub Form_Unload(Cancel As Integer)
If Me.Tag="blockclose" Then Cancel=1
End Sub
Sub action_Click()
CurrentDb.Execute "INSERT INTO events VALUES('clicked')"
End Sub
Sub Recalculate()
Me.Recalc
End Sub
Sub Discard()
Me.Undo
End Sub
)VBA";
    AccessFormPreview form("keyboard",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
    auto object=makePreviewFormObject(form);
    auto* first=form.findChild<QLineEdit*>("access-control-first"); auto* second=form.findChild<QLineEdit*>("access-control-second"); assert(first && second);
    first->setFocus(); QApplication::processEvents(); QTest::keyClick(first,Qt::Key_Escape);
    assert(form.isVisible()); // Escape must never fall through to QDialog reject.
    first->selectAll(); QTest::keyClicks(first,"edited"); second->setFocus(); QApplication::processEvents();
    QTest::keyClicks(second,"transient"); assert(form.hasUnsavedChanges());
    QTest::keyClick(second,Qt::Key_Escape);
    assert(first->text()=="edited" && second->property("la_null").toBool() && form.hasUnsavedChanges());
    QTest::keyClick(second,Qt::Key_Escape);
    assert(first->text()=="before" && !form.hasUnsavedChanges() && form.isVisible());
    first->setFocus(); QApplication::processEvents(); first->selectAll(); QTest::keyClicks(first,"persisted");
    object->assign("Tag",VbaValue::string("blocksave")); QTest::keyClick(first,Qt::Key_Return,Qt::ShiftModifier);
    assert(form.hasUnsavedChanges()); assert(query.exec("SELECT first FROM records WHERE id=1") && query.next() && query.value(0).toString()=="before"); query.finish();
    object->assign("Tag",VbaValue::string("")); QTest::keyClick(first,Qt::Key_Return,Qt::ShiftModifier);
    assert(!form.hasUnsavedChanges()); assert(query.exec("SELECT first FROM records WHERE id=1") && query.next() && query.value(0).toString()=="persisted"); query.finish();
    assert(query.exec("UPDATE records SET first='external' WHERE id=1"));
    QTest::keyClick(first,Qt::Key_F9,Qt::ShiftModifier); assert(first->text()=="external");
    assert(object->item("counter").asNumber()==1); assert(query.exec("INSERT INTO records VALUES(2,'new',NULL)"));
    QTest::keyClick(first,Qt::Key_F9); assert(object->item("counter").asNumber()==2);
    auto* selector=form.findChild<QComboBox*>("access-record-selector"); assert(selector && selector->count()==2); // Recalc does not reload records.
    assert(query.exec("INSERT INTO records VALUES(3,'third',NULL)")); std::vector<VbaValue> none;
    assert(form.executeProcedure("Recalculate",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(object->item("counter").asNumber()==3 && selector->count()==2);
    first->setFocus(); first->setSelection(0,first->text().size()); QTest::keyClick(first,Qt::Key_F2); assert(!first->hasSelectedText());
    QTest::keyClick(first,Qt::Key_F2); assert(first->selectedText()==first->text());
    QTest::keyClick(first,Qt::Key_Return); QApplication::processEvents(); assert(form.isVisible() && second->hasFocus());
    assert(query.exec("SELECT COUNT(*) FROM events") && query.next() && query.value(0).toInt()==0); query.finish();
    object->assignItem("first",VbaValue::string("discarded by VBA"));
    assert(form.executeProcedure("Discard",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(first->text()=="external" && !form.hasUnsavedChanges());
    object->assign("Tag",VbaValue::string("blockclose")); QTest::keyClick(second,Qt::Key_W,Qt::ControlModifier); assert(form.isVisible());
    object->assign("Tag",VbaValue::string("")); QTest::keyClick(second,Qt::Key_F4,Qt::ControlModifier); assert(!form.isVisible());
    assert(query.exec("CREATE TABLE numbers(id INTEGER PRIMARY KEY,value INTEGER)"));
    assert(query.exec("INSERT INTO numbers VALUES(1,10)"));
    const QByteArray numericLayout="Begin Form\nRecordSource =\"numbers\"\nBegin Section\nBegin TextBox\nName =\"value\"\nControlSource =\"value\"\nEnd\nBegin TextBox\nName =\"alias\"\nControlSource =\"value\"\nTop =400\nEnd\nEnd\nEnd\n";
    AccessFormPreview numbers("numbers",numericLayout,{},db); numbers.show(); QApplication::setActiveWindow(&numbers); QApplication::processEvents();
    auto* number=numbers.findChild<QLineEdit*>("access-control-value"); auto* alias=numbers.findChild<QLineEdit*>("access-control-alias"); assert(number && alias);
    number->setFocus(); QApplication::processEvents(); number->selectAll(); QTest::keyClicks(number,"invalid number");
    QTest::keyClick(number,Qt::Key_Escape);
    assert(number->text()=="10" && alias->text()=="10" && !numbers.hasUnsavedChanges());
    auto* records=numbers.findChild<QComboBox*>("access-record-selector"); assert(records); records->setCurrentIndex(0);
    number->setFocus(); QApplication::processEvents(); QTest::keyClicks(number,"22"); QTest::keyClick(number,Qt::Key_Escape);
    assert(number->property("la_null").toBool() && alias->property("la_null").toBool() && !numbers.hasUnsavedChanges());
    assert(query.exec("SELECT COUNT(*) FROM numbers") && query.next() && query.value(0).toInt()==1); query.finish();
    assert(numbers.close());
    testKeyboardEvents();
}

void testKeyboardEvents() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","keyboard-events"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE records(id INTEGER)")); assert(query.exec("INSERT INTO records VALUES(1)"));
    assert(query.exec("CREATE TABLE events(kind TEXT,code INTEGER,shift INTEGER)"));
    const QByteArray layout=R"FORM(Begin Form
KeyPreview =-1
OnKeyDown ="[Event Procedure]"
OnKeyPress ="[Event Procedure]"
OnKeyUp ="[Event Procedure]"
Begin Section
Begin TextBox
Name ="editor"
OnKeyDown ="[Event Procedure]"
OnKeyPress ="[Event Procedure]"
OnKeyUp ="[Event Procedure]"
End
Begin TextBox
Name ="counter"
ControlSource ="=DCount(""*"",""records"")"
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub Form_KeyDown(KeyCode As Integer,Shift As Integer)
CurrentDb.Execute "INSERT INTO events VALUES('FD'," & KeyCode & "," & Shift & ")"
If KeyCode=vbKeyW And Shift=acCtrlMask Then KeyCode=0
If KeyCode=vbKeyF9 And Me.Tag="cancelrecalc" Then KeyCode=0
If KeyCode=vbKeyF8 Then Err.Raise 11
End Sub
Sub Form_KeyPress(KeyAscii As Integer)
CurrentDb.Execute "INSERT INTO events VALUES('FP'," & KeyAscii & ",0)"
If KeyAscii=97 Then KeyAscii=65
If KeyAscii=122 Then KeyAscii=0
End Sub
Sub Form_KeyUp(KeyCode As Integer,Shift As Integer)
CurrentDb.Execute "INSERT INTO events VALUES('FU'," & KeyCode & "," & Shift & ")"
End Sub
Sub editor_KeyDown(KeyCode As Integer,Shift As Integer)
CurrentDb.Execute "INSERT INTO events VALUES('CD'," & KeyCode & "," & Shift & ")"
End Sub
Sub editor_KeyPress(KeyAscii As Integer)
CurrentDb.Execute "INSERT INTO events VALUES('CP'," & KeyAscii & ",0)"
If KeyAscii=65 Then KeyAscii=66
End Sub
Sub editor_KeyUp(KeyCode As Integer,Shift As Integer)
CurrentDb.Execute "INSERT INTO events VALUES('CU'," & KeyCode & "," & Shift & ")"
End Sub
Sub ControlFirst()
Me.KeyPreview=False
End Sub
)VBA";
    AccessFormPreview form("key events",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
    auto root=makePreviewFormObject(form); auto* editor=form.findChild<QLineEdit*>("access-control-editor"); assert(editor);
    editor->setFocus(); QApplication::processEvents();
    assert(root->member("KeyPreview",{}).truthy());
    QTest::keyClicks(editor,"a"); assert(editor->text()=="B");
    const auto kinds=[&] {
        QStringList result; assert(query.exec("SELECT kind FROM events ORDER BY rowid"));
        while (query.next()) result.append(query.value(0).toString()); query.finish(); return result;
    };
    assert(kinds()==QStringList({"FD","CD","FP","CP","FU","CU"}));
    assert(query.exec("SELECT code FROM events WHERE kind='CP'") && query.next() && query.value(0).toInt()==65); query.finish();
    assert(query.exec("DELETE FROM events"));
    std::vector<VbaValue> none; assert(form.executeProcedure("ControlFirst",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(!root->member("KeyPreview",{}).truthy()); QTest::keyClicks(editor,"a"); assert(editor->text()=="BA");
    assert(kinds()==QStringList({"CD","FD","CP","FP","CU","FU"}));
    root->assign("KeyPreview",VbaValue::logical(true)); assert(query.exec("DELETE FROM events"));
    QTest::keyClicks(editor,"z"); assert(editor->text()=="BA");
    assert(query.exec("SELECT COUNT(*) FROM events WHERE kind='CP'") && query.next() && query.value(0).toInt()==0); query.finish();
    QTest::keyClick(editor,Qt::Key_W,Qt::ControlModifier); assert(form.isVisible());
    assert(query.exec("SELECT COUNT(*) FROM events WHERE kind='CD' AND code=87") && query.next() && query.value(0).toInt()==0); query.finish();
    QTest::keyClick(editor,Qt::Key_X,Qt::ShiftModifier);
    assert(query.exec("SELECT COUNT(*) FROM events WHERE kind='FD' AND code=88 AND shift=1") && query.next() && query.value(0).toInt()==1); query.finish();
    assert(query.exec("INSERT INTO records VALUES(2)")); root->assign("Tag",VbaValue::string("cancelrecalc"));
    QTest::keyClick(editor,Qt::Key_F9); assert(root->item("counter").asNumber()==1);
    root->assign("Tag",VbaValue::string("")); QTest::keyClick(editor,Qt::Key_F9); assert(root->item("counter").asNumber()==2);
    QTest::keyClick(editor,Qt::Key_F8); const auto error=form.eventResult("OnKeyDown");
    assert(error && error->status==linaccess::VbaExecutionStatus::Error && error->errorNumber==11 && form.isVisible());
    bool invalid=false; try { root->assign("KeyPreview",VbaValue{}); } catch(const linaccess::VbaRuntimeError& e) { invalid=e.number==94; } assert(invalid);
    const auto beforeUnsupported=editor->text();
    QKeyEvent unsupported(QEvent::KeyPress,Qt::Key_unknown,Qt::NoModifier,QString::fromUtf8("ü"));
    QApplication::sendEvent(editor,&unsupported);
    const auto unsupportedResult=form.eventResult("keyboarderror");
    assert(unsupportedResult && unsupportedResult->status==linaccess::VbaExecutionStatus::Unsupported && unsupportedResult->errorNumber==3251);
    assert(editor->text()==beforeUnsupported);
    assert(form.close());
    AccessFormPreview report("key report","Begin Report\nEnd\n",{},db); report.show(); QApplication::setActiveWindow(&report); QApplication::processEvents();
    QTest::keyClick(&report,Qt::Key_Escape); assert(!report.isVisible());
    assert(query.exec("CREATE TABLE parent_records(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(query.exec("INSERT INTO parent_records VALUES(1,'parent before')"));
    assert(query.exec("CREATE TABLE child_records(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(query.exec("INSERT INTO child_records VALUES(1,'child before')"));
    assert(query.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    const QByteArray child="Begin Form\nRecordSource =\"child_records\"\nBegin Section\nBegin TextBox\nName =\"value\"\nControlSource =\"value\"\nEnd\nEnd\nEnd\n";
    query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('form','child',?)"); query.addBindValue(child); assert(query.exec()); query.finish();
    const QByteArray parent="Begin Form\nRecordSource =\"parent_records\"\nKeyPreview =-1\nOnKeyDown =\"[Event Procedure]\"\nBegin Section\nBegin TextBox\nName =\"value\"\nControlSource =\"value\"\nEnd\nBegin Subform\nName =\"detail\"\nSourceObject =\"Form.child\"\nTop =500\nWidth =5000\nHeight =2500\nEnd\nEnd\nEnd\n";
    const QByteArray parentCode="Sub Form_KeyDown(KeyCode As Integer,Shift As Integer)\nCurrentDb.Execute \"INSERT INTO events VALUES('parent',0,0)\"\nEnd Sub\n";
    AccessFormPreview parentForm("parent keys",parent,parentCode,db); parentForm.show(); QApplication::setActiveWindow(&parentForm); QApplication::processEvents();
    auto* parentLine=parentForm.findChild<QLineEdit*>("access-control-value"); assert(parentLine);
    parentLine->setFocus(); parentLine->selectAll(); QTest::keyClicks(parentLine,"parent pending");
    auto* childForm=dynamic_cast<AccessFormPreview*>(parentForm.findChild<QWidget*>("access-control-detail")); assert(childForm);
    auto* childLine=childForm->findChild<QLineEdit*>("access-control-value"); assert(childLine);
    childLine->setFocus(); QApplication::processEvents(); assert(query.exec("DELETE FROM events"));
    childLine->selectAll(); QTest::keyClicks(childLine,"child saved"); QTest::keyClick(childLine,Qt::Key_Return,Qt::ShiftModifier);
    assert(query.exec("SELECT value FROM child_records WHERE id=1") && query.next() && query.value(0).toString()=="child saved"); query.finish();
    assert(query.exec("SELECT value FROM parent_records WHERE id=1") && query.next() && query.value(0).toString()=="parent before"); query.finish();
    assert(parentForm.hasUnsavedChanges() && !childForm->hasUnsavedChanges());
    assert(query.exec("SELECT COUNT(*) FROM events") && query.next() && query.value(0).toInt()==0); query.finish();
    makePreviewFormObject(parentForm)->member("Undo",{});
    QTest::keyClick(childLine,Qt::Key_W,Qt::ControlModifier); assert(!parentForm.isVisible());

}

void testOriginalKeyboardEvents(const QDir& directory,const QJsonArray& objects,QSqlDatabase db) {
    QSqlQuery query(db); assert(query.exec("SAVEPOINT original_keyboard_events"));
    QMainWindow project;
    auto* toolbar=project.addToolBar("Project"); toolbar->addAction("Test command");
    project.menuBar()->addMenu("Project")->addAction(toolbar->actions().first());
    int unexpected=0; QTimer dismiss; dismiss.setInterval(1);
    QObject::connect(&dismiss,&QTimer::timeout,[&] { if (auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget())) { ++unexpected; box->accept(); } }); dismiss.start();
    int executions=0,directExecutions=0,missingBindings=0;
    for (const auto& entry:objects) {
        const auto module=entry.toObject();
        if (module["kind"]!="vba") continue;
        const auto hash=module["sha256"].toString();
        if (hash!="6bde4373d4584a1626a6a7f5edbb14d072f3f01083bd797faf9fee914f04f7f8"
            && hash!="6e8e0a0e41e1a1768cff4a0b79a07205c7b0f799eb8b465af5397843ce048ed8") continue;
        QFile file(directory.filePath(module["file"].toString())); assert(file.open(QIODevice::ReadOnly));
        const auto code=file.readAll(); assert(QCryptographicHash::hash(code,QCryptographicHash::Sha256).toHex()==hash.toLatin1());
        const auto name=module["name"].toString().mid(5); QByteArray layout;
        for (const auto& candidate:objects) {
            const auto form=candidate.toObject(); if (form["kind"]!="form" || form["name"].toString()!=name) continue;
            QFile source(directory.filePath(form["file"].toString())); assert(source.open(QIODevice::ReadOnly)); layout=source.readAll();
            assert(QCryptographicHash::hash(layout,QCryptographicHash::Sha256).toHex()==form["sha256"].toString().toLatin1());
        }
        assert(!layout.isEmpty());
        AccessFormPreview document(name,layout,code,db,&project); document.show(); QApplication::setActiveWindow(&document); QApplication::processEvents();
        const auto procedures=linaccess::analyzeVba(linaccess::decodeVbaSource(code).toStdString()).procedures;
        const auto parsed=linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(layout));
        for (const auto& control:parsed.controls) {
            if (control.onKeyPress.compare("[Event Procedure]",Qt::CaseInsensitive)!=0) continue;
            auto* editor=document.findChild<QLineEdit*>("access-control-"+control.name); assert(editor);
            editor->setFocus(); QApplication::processEvents();
            const auto before=editor->text(); const bool beforeNull=editor->property("la_null").toBool();
            QKeyEvent down(QEvent::KeyPress,Qt::Key_A,Qt::NoModifier,"a"); QApplication::sendEvent(editor,&down);
            QKeyEvent up(QEvent::KeyRelease,Qt::Key_A,Qt::NoModifier,"a"); QApplication::sendEvent(editor,&up);
            const auto result=document.eventResult("control:"+control.name+":KeyPress"); assert(result);
            const auto handler=control.name+"_KeyPress";
            const bool present=std::any_of(procedures.begin(),procedures.end(),[&](const auto& procedure) {
                return QString::fromStdString(procedure.name).compare(handler,Qt::CaseInsensitive)==0;
            });
            if (present) { assert(result->status==linaccess::VbaExecutionStatus::Completed); ++executions; }
            else { assert(result->status==linaccess::VbaExecutionStatus::ProcedureMissing); ++missingBindings; }
            assert(editor->text()==before && editor->property("la_null").toBool()==beforeNull);
        }
        for (const auto& procedure:procedures) {
            const auto name=QString::fromStdString(procedure.name);
            if (!name.endsWith("_KeyPress",Qt::CaseInsensitive)) continue;
            std::vector<linaccess::VbaValue> arguments{linaccess::VbaValue::numeric(97)};
            assert(document.executeProcedure(name,arguments).status==linaccess::VbaExecutionStatus::Completed && arguments[0].asNumber()==0);
            ++directExecutions;
        }
        assert(document.close());
    }
    dismiss.stop(); assert(unexpected==0 && executions==1 && directExecutions==3 && missingBindings==2);
    assert(query.exec("ROLLBACK TO original_keyboard_events")); assert(query.exec("RELEASE original_keyboard_events"));
    std::cout<<"Original KeyPress: "<<executions<<" bound execution, "<<directExecutions<<" explicit calls, "<<missingBindings<<" missing-handler bindings; unchanged native control values verified\n";
}

void testTextSelection() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","text-selection"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(sql.exec("INSERT INTO records VALUES(1,'abcdef')"));
    assert(sql.exec("CREATE TABLE trace(value TEXT)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="records"
Begin Section
Begin TextBox
Name ="editor"
ControlSource ="value"
OnChange ="[Event Procedure]"
End
Begin ComboBox
Name ="combo"
ControlSource ="value"
Top =400
End
Begin CommandButton
Name ="button"
Top =800
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub ReplaceSelection()
Me.editor.SetFocus
Me.editor.SelStart=2
Me.editor.SelLength=2
If Me.editor.SelText<>"cd" Then Err.Raise 5
Me.editor.SelText="XY"
End Sub
Sub editor_Change()
CurrentDb.Execute "INSERT INTO trace VALUES('change')"
End Sub
)VBA";
    AccessFormPreview form("selection",layout,code,db); form.show(); form.activateWindow(); QApplication::setActiveWindow(&form); QApplication::processEvents();
    auto root=makePreviewFormObject(form);
    auto editor=root->objectMember("editor",{}).object,combo=root->objectMember("combo",{}).object;
    auto* line=form.findChild<QLineEdit*>("access-control-editor"); auto* comboWidget=form.findChild<QComboBox*>("access-control-combo"); assert(line && comboWidget && comboWidget->lineEdit());
    std::vector<VbaValue> none; assert(form.executeProcedure("ReplaceSelection",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(line->text()=="abXYef" && editor->member("SelStart",{}).asNumber()==4);
    assert(editor->member("SelLength",{}).asNumber()==0 && editor->member("SelText",{}).kind==VbaValue::Kind::Null);
    root->assign("Dirty",VbaValue::logical(false));
    assert(sql.exec("SELECT value FROM records") && sql.next() && sql.value(0).toString()=="abXYef"); sql.finish();
    assert(sql.exec("SELECT COUNT(*) FROM trace") && sql.next() && sql.value(0).toInt()==1); sql.finish();
    editor->assign("SelStart",VbaValue::numeric(1)); editor->assign("SelLength",VbaValue::numeric(99));
    assert(editor->member("SelText",{}).asText()=="bXYef");
    editor->assign("SelStart",VbaValue::numeric(99));
    assert(editor->member("SelStart",{}).asNumber()==6 && editor->member("SelLength",{}).asNumber()==0);
    for (const auto& [value,expected]:std::vector<std::pair<VbaValue,int>>{{VbaValue::numeric(-1),5},{VbaValue::numeric(32768),6},{VbaValue{},94}}) {
        int error=0; try { editor->assign("SelLength",value); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==expected);
    }
    assert(!form.hasUnsavedChanges());
    comboWidget->lineEdit()->setFocus(); QApplication::processEvents();
    int error=0; try { editor->member("SelStart",{}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==2185);
    combo->assign("SelStart",VbaValue::numeric(0)); combo->assign("SelLength",VbaValue::numeric(2)); combo->assign("SelText",VbaValue::string("12"));
    assert(comboWidget->currentText()=="12XYef" && root->item("editor").asText()=="abXYef");
    root->assign("Dirty",VbaValue::logical(false));
    assert(sql.exec("SELECT value FROM records") && sql.next() && sql.value(0).toString()=="12XYef"); sql.finish();
    error=0; try { root->objectMember("button",{}).object->member("SelText",{}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==438);
    line->setFocus(); QApplication::processEvents();
    root->assignItem("editor",VbaValue::string("A😀Z"));
    editor->assign("SelStart",VbaValue::numeric(1)); editor->assign("SelLength",VbaValue::numeric(2));
    assert(editor->member("SelText",{}).asText()=="😀");
    editor->assign("SelText",VbaValue::string("ö"));
    assert(editor->member("Text",{}).asText()=="AöZ" && root->item("editor").asText()=="A😀Z");
    root->assign("Dirty",VbaValue::logical(false));
    assert(sql.exec("SELECT value FROM records") && sql.next() && sql.value(0).toString()==QString::fromUtf8("AöZ")); sql.finish();
    assert(form.close());
}
