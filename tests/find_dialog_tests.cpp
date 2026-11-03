#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include <QCheckBox>
#include <QComboBox>
#include <QLabel>
#include <QLineEdit>
#include <QPushButton>
#include <QPixmap>
#include <QSqlQuery>
#include <QTest>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

int main(int argc,char** argv) {
    QApplication app(argc,argv);
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","find-form-documents"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,name TEXT,notes TEXT)"));
    assert(sql.exec("INSERT INTO records VALUES(1,'Alpha','first'),(2,'NeedleÖ','other'),(3,'beta','Needle note'),(4,'prefix needleÖ suffix',NULL)"));
    const QString longPattern=QString(40000,'x')+"special";
    sql.prepare("INSERT INTO records VALUES(5,'long',?)"); sql.addBindValue(longPattern); assert(sql.exec()); sql.finish();
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="records"
OnUnload ="[Event Procedure]"
Begin Section
Begin TextBox
Name ="name"
ControlSource ="name"
BeforeUpdate ="[Event Procedure]"
Width =5000
Height =400
End
Begin TextBox
Name ="notes"
ControlSource ="notes"
Top =500
Width =5000
Height =400
End
Begin TextBox
Name ="id"
ControlSource ="id"
Locked =-1
TabStop =0
Top =1000
Width =1000
Height =400
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub name_BeforeUpdate(Cancel As Integer)
If Me.Tag="blockupdate" Then Cancel=True
End Sub
Sub Form_Unload(Cancel As Integer)
If Me.Tag="blockclose" Then Cancel=True
End Sub
)VBA";
    AccessFormPreview form("find records",layout,code,db); form.show(); form.activateWindow();
    auto model=makePreviewFormObject(form); auto* editor=form.findChild<QLineEdit*>("access-control-name"); assert(editor);
    const auto focus=[&] { form.activateWindow(); editor->setFocus(); assert(QTest::qWaitFor([&] { return QApplication::focusWidget()==editor; },2000)); };
    const auto stored=[&](int id) { sql.prepare("SELECT name FROM records WHERE id=?"); sql.addBindValue(id); assert(sql.exec() && sql.next()); const auto value=sql.value(0).toString(); sql.finish(); return value; };
    auto nameControl=model->member("Controls",{VbaValue::string("name")}).object;
    nameControl->assign("InputMask",VbaValue::string("Password"));
    focus(); editor->selectAll(); QTest::keyClick(editor,Qt::Key_F,Qt::ControlModifier);
    auto* dialog=form.findChild<QDialog*>("access-find-dialog"); assert(dialog && dialog->isVisible() && !dialog->isModal() && !QApplication::activeModalWidget());
    auto* query=dialog->findChild<QLineEdit*>("access-find-text"); auto* field=dialog->findChild<QComboBox*>("access-find-field");
    auto* match=dialog->findChild<QComboBox*>("access-find-match"); auto* sensitive=dialog->findChild<QCheckBox*>("access-find-case");
    auto* next=dialog->findChild<QPushButton*>("access-find-next"); auto* close=dialog->findChild<QPushButton*>("access-find-close"); auto* result=dialog->findChild<QLabel*>("access-find-result");
    assert(query && field && match && sensitive && next && close && result);
    assert(query->text().isEmpty()); nameControl->assign("InputMask",VbaValue::string(""));
    QCoreApplication::processEvents();
    const auto screenshot=qEnvironmentVariable("LINACCESS_PRIVATE_FIND_SCREENSHOT"); if(!screenshot.isEmpty()) assert(dialog->grab().save(screenshot));
    const QByteArray callerCode=R"VBA(Function ActiveName() As String
ActiveName=Screen.ActiveForm.Name
End Function
Sub FindActive()
DoCmd.RunCommand acCmdFind
End Sub
Sub FindNextActive()
DoCmd.RunCommand acCmdFindNext
End Sub
Sub RepeatViaSendKeys()
SendKeys "+{F4}", True
End Sub
Sub FindBeta()
DoCmd.FindRecord FindWhat:="BETA",MatchCase:=False
Me.Tag="findrecord-beta"
End Sub
Sub FindAlpha()
DoCmd.FindRecord "Alpha"
Me.Tag="findrecord-alpha"
End Sub
Sub FindNeedleAfterCurrent()
DoCmd.FindRecord "needle",acAnywhere,,,,,False
Me.Tag="findrecord-after"
End Sub
Sub FindUpBeta()
DoCmd.FindRecord "beta",Search:=acUp,FindFirst:=False
End Sub
Sub FindDownAlphaNoWrap()
DoCmd.FindRecord "Alpha",Search:=acDown,FindFirst:=False
End Sub
Sub FindUpFromFirstAlpha()
DoCmd.FindRecord "Alpha",Search:=acUp,FindFirst:=True
End Sub
Sub FindAbsent()
DoCmd.FindRecord "absent"
Me.Tag="findrecord-nohit"
End Sub
)VBA";
    AccessFormPreview caller("hidden find caller","Begin Form\nEnd",callerCode,db);
    std::vector<VbaValue> none;
    auto active=caller.executeProcedure("ActiveName",none); assert(active.status==linaccess::VbaExecutionStatus::Completed && active.typedReturnValue->asText()=="find records");
    assert(caller.executeProcedure("FindActive",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(form.findChildren<QDialog*>("access-find-dialog").size()==1 && caller.findChildren<QDialog*>("access-find-dialog").isEmpty());
    const auto noHistory=caller.executeProcedure("FindNextActive",none); assert(noHistory.status==linaccess::VbaExecutionStatus::Error && noHistory.errorNumber==2046);
    field->setCurrentIndex(field->findData("name")); match->setCurrentIndex(1); query->setText("needleö"); query->setFocus(); QTest::keyClick(query,Qt::Key_Return); assert(model->item("id").number==2);
    auto* direction=dialog->findChild<QComboBox*>("access-find-direction"); assert(direction);
    direction->setCurrentIndex(direction->findData(1)); query->setText("prefix needleÖ suffix"); next->click(); assert(model->item("id").number==4);
    direction->setCurrentIndex(direction->findData(0)); query->setText("beta"); next->click(); assert(model->item("id").number==3);
    query->setText("prefix needleÖ suffix"); next->click(); assert(model->item("id").number==3 && result->text()=="Keine Übereinstimmung gefunden.");
    direction->setCurrentIndex(direction->findData(2)); query->setText("needleö"); next->click(); assert(model->item("id").number==2);
    field->setCurrentIndex(0); match->setCurrentIndex(0); query->setText("Needle"); next->click(); assert(model->item("id").number==2);
    next->click(); assert(model->item("id").number==3); next->click(); assert(model->item("id").number==4);
    field->setCurrentIndex(field->findData("name")); match->setCurrentIndex(2); sensitive->setChecked(true); next->click(); assert(model->item("id").number==2);
    field->setCurrentIndex(field->findData("notes")); match->setCurrentIndex(1); query->setText(longPattern); next->click(); assert(model->item("id").number==5);
    field->setCurrentIndex(field->findData("name")); sensitive->setChecked(false); query->setText("NeedleÖ"); next->click(); assert(model->item("id").number==2);
    query->setText("Needle*"); next->click(); assert(model->item("id").number==2 && result->text()=="Keine Übereinstimmung gefunden.");
    query->setText(""); assert(!next->isEnabled());
    field->setCurrentIndex(0); match->setCurrentIndex(0); sensitive->setChecked(false); query->setText("Needle"); next->click(); assert(model->item("id").number==2);
    close->click(); QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete);
    assert(!form.findChild<QDialog*>("access-find-dialog"));
    focus(); QTest::keyClick(editor,Qt::Key_F4,Qt::ShiftModifier); assert(model->item("id").number==3 && !form.findChild<QDialog*>("access-find-dialog"));
    assert(caller.executeProcedure("FindNextActive",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(model->item("id").number==4 && !caller.findChild<QDialog*>("access-find-dialog"));
    focus(); assert(caller.executeProcedure("RepeatViaSendKeys",none).status==linaccess::VbaExecutionStatus::Completed); assert(model->item("id").number==2);
    QTest::keyClick(editor,Qt::Key_F3); assert(model->item("id").number==2);
    editor->selectAll(); QTest::keyClicks(editor,"blocked repeat"); model->assign("Tag",VbaValue::string("blockupdate"));
    QTest::keyClick(editor,Qt::Key_F4,Qt::ShiftModifier); assert(model->item("id").number==2 && stored(2)=="NeedleÖ" && model->member("Dirty",{}).boolean);
    assert(!form.findChild<QDialog*>("access-find-dialog")); model->assign("Tag",VbaValue::string("")); model->member("Undo",{});
    QTest::keyClick(editor,Qt::Key_F,Qt::ControlModifier); dialog=form.findChild<QDialog*>("access-find-dialog"); assert(dialog);
    query=dialog->findChild<QLineEdit*>("access-find-text"); field=dialog->findChild<QComboBox*>("access-find-field"); match=dialog->findChild<QComboBox*>("access-find-match"); sensitive=dialog->findChild<QCheckBox*>("access-find-case");
    next=dialog->findChild<QPushButton*>("access-find-next"); close=dialog->findChild<QPushButton*>("access-find-close"); result=dialog->findChild<QLabel*>("access-find-result");
    focus(); QTest::keyClick(editor,Qt::Key_F,Qt::ControlModifier); assert(form.findChildren<QDialog*>("access-find-dialog").size()==1);
    focus(); editor->selectAll(); QTest::keyClicks(editor,"draft0"); model->assign("Tag",VbaValue::string("blockupdate"));
    query->setText("NeedleÖ"); match->setCurrentIndex(1); sensitive->setChecked(false); next->click();
    assert(model->item("id").number==2 && stored(2)=="NeedleÖ" && result->text()=="Suche durch Formularvalidierung abgebrochen.");
    model->assign("Tag",VbaValue::string("")); model->member("Undo",{});
    close->click(); QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete);
    focus(); editor->selectAll(); QTest::keyClicks(editor,"draft"); model->assign("Tag",VbaValue::string("blockupdate"));
    QTest::keyClick(editor,Qt::Key_F,Qt::ControlModifier); assert(!form.findChild<QDialog*>("access-find-dialog") && stored(2)=="NeedleÖ" && model->member("Dirty",{}).boolean);
    model->assign("Tag",VbaValue::string("")); QTest::keyClick(editor,Qt::Key_F,Qt::ControlModifier);
    dialog=form.findChild<QDialog*>("access-find-dialog"); assert(dialog && dialog->isVisible());
    query=dialog->findChild<QLineEdit*>("access-find-text"); field=dialog->findChild<QComboBox*>("access-find-field"); match=dialog->findChild<QComboBox*>("access-find-match"); next=dialog->findChild<QPushButton*>("access-find-next");
    field->setCurrentIndex(field->findData("name")); match->setCurrentIndex(1); query->setText("draft"); next->click();
    assert(model->item("id").number==2 && stored(2)=="NeedleÖ" && model->member("Dirty",{}).boolean);
    query->setText("missing"); next->click(); assert(model->item("id").number==2 && stored(2)=="NeedleÖ");
    query->setText("Alpha"); next->click(); assert(model->item("id").number==1 && stored(2)=="draft" && !model->member("Dirty",{}).boolean);
    close=dialog->findChild<QPushButton*>("access-find-close"); close->click(); QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete);
    focus(); assert(caller.executeProcedure("FindBeta",none).status==linaccess::VbaExecutionStatus::Completed && model->item("id").number==3);
    assert(makePreviewFormObject(caller)->member("Tag",{}).asText()=="findrecord-beta");
    assert(caller.executeProcedure("FindAlpha",none).status==linaccess::VbaExecutionStatus::Completed && model->item("id").number==1);
    assert(caller.executeProcedure("FindNeedleAfterCurrent",none).status==linaccess::VbaExecutionStatus::Completed && model->item("id").number==4);
    assert(caller.executeProcedure("FindUpBeta",none).status==linaccess::VbaExecutionStatus::Completed && model->item("id").number==3);
    assert(caller.executeProcedure("FindDownAlphaNoWrap",none).status==linaccess::VbaExecutionStatus::Completed && model->item("id").number==3);
    assert(caller.executeProcedure("FindUpFromFirstAlpha",none).status==linaccess::VbaExecutionStatus::Completed && model->item("id").number==1);
    assert(caller.executeProcedure("FindBeta",none).status==linaccess::VbaExecutionStatus::Completed && model->item("id").number==3);
    assert(caller.executeProcedure("FindAbsent",none).status==linaccess::VbaExecutionStatus::Completed && model->item("id").number==3);
    focus(); editor->selectAll(); QTest::keyClicks(editor,"blocked API search"); model->assign("Tag",VbaValue::string("blockupdate"));
    assert(caller.executeProcedure("FindAlpha",none).errorNumber==2501 && model->item("id").number==3 && stored(3)=="beta" && model->member("Dirty",{}).boolean);
    assert(makePreviewFormObject(caller)->member("Tag",{}).asText()=="findrecord-nohit");
    model->assign("Tag",VbaValue::string("")); model->member("Undo",{});
    focus(); assert(caller.executeProcedure("FindAlpha",none).status==linaccess::VbaExecutionStatus::Completed && model->item("id").number==1);
    nameControl->assign("Locked",VbaValue::logical(true)); focus();
    assert(caller.executeProcedure("FindBeta",none).status==linaccess::VbaExecutionStatus::Completed && model->item("id").number==3 && stored(3)=="beta");
    nameControl->assign("Locked",VbaValue::logical(false));
    focus(); QTest::keyClick(editor,Qt::Key_F,Qt::ControlModifier); dialog=form.findChild<QDialog*>("access-find-dialog"); assert(dialog && dialog->isVisible());
    assert(caller.close());
    model->assign("Tag",VbaValue::string("blockclose")); assert(!form.close() && dialog->isVisible());
    model->assign("Tag",VbaValue::string("")); const QPointer<QDialog> guarded(dialog); assert(form.close());
    QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete); assert(!guarded || !guarded->isVisible());
    assert(sql.exec("SELECT COUNT(*) FROM records") && sql.next() && sql.value(0).toInt()==5); sql.finish();
}
