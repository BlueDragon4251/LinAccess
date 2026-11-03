#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include <QComboBox>
#include <QLineEdit>
#include <QSqlQuery>
#include <QTest>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

int main(int argc,char** argv) {
    QApplication app(argc,argv);
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","record-navigation-keys"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,label TEXT)"));
    assert(sql.exec("INSERT INTO records VALUES(1,'first'),(2,'second'),(3,'third')"));
    assert(sql.exec("CREATE TABLE events(kind TEXT)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="records"
KeyPreview =-1
OnKeyDown ="[Event Procedure]"
BeforeUpdate ="[Event Procedure]"
OnCurrent ="[Event Procedure]"
Begin Section
Begin TextBox
Name ="label"
ControlSource ="label"
BeforeUpdate ="[Event Procedure]"
OnKeyDown ="[Event Procedure]"
Width =4000
End
Begin ComboBox
Name ="choice"
ControlSource ="label"
RowSourceType ="Value List"
RowSource ="first;second;third;edited"
Top =500
Width =4000
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub Form_KeyDown(KeyCode As Integer,Shift As Integer)
If Me.Tag="blockkey" And Shift=acCtrlMask Then KeyCode=0
End Sub
Sub label_KeyDown(KeyCode As Integer,Shift As Integer)
If Me.Tag="controlblock" And Shift=acCtrlMask Then KeyCode=0
End Sub
Sub label_BeforeUpdate(Cancel As Integer)
If Me.Tag="blockfield" Then Cancel=True
End Sub
Sub Form_BeforeUpdate(Cancel As Integer)
If Me.Tag="blocksave" Then Cancel=True
End Sub
Sub Form_Current()
CurrentDb.Execute "INSERT INTO events VALUES('current')"
If Me.Tag="closecurrent" Then DoCmd.Close
End Sub
Sub SendNext()
SendKeys "^{PGDN}",True
End Sub
Sub SendPrevious()
SendKeys "^{PGUP}",True
End Sub
)VBA";
    AccessFormPreview form("record navigation keys",layout,code,db); form.show(); QApplication::setActiveWindow(&form);
    auto model=makePreviewFormObject(form);
    auto* editor=form.findChild<QLineEdit*>("access-control-label"); assert(editor);
    const auto focus=[&] { editor->setFocus(); assert(QTest::qWaitFor([&] { return QApplication::focusWidget()==editor; },2000)); };
    const auto tag=[&](const char* value) { model->assign("Tag",VbaValue::string(value)); };
    const auto key=[&](bool next) { focus(); QTest::keyClick(editor,next ? Qt::Key_PageDown : Qt::Key_PageUp,Qt::ControlModifier); };
    const auto stored=[&] { assert(sql.exec("SELECT label FROM records WHERE id=1") && sql.next()); const auto value=sql.value(0).toString(); sql.finish(); return value; };
    const auto count=[&] { assert(sql.exec("SELECT COUNT(*) FROM records") && sql.next()); const auto value=sql.value(0).toInt(); sql.finish(); return value; };
    const auto currents=[&] { assert(sql.exec("SELECT COUNT(*) FROM events") && sql.next()); const auto value=sql.value(0).toInt(); sql.finish(); return value; };
    focus(); assert(editor->text()=="first" && currents()==1);
    tag("blockkey"); key(true); assert(editor->text()=="first" && !form.hasUnsavedChanges());
    tag("controlblock"); key(true); assert(editor->text()=="first");
    tag("blockfield"); editor->selectAll(); QTest::keyClicks(editor,"edited"); key(true);
    assert(editor->text()=="edited" && model->item("label").asText()=="first" && stored()=="first" && form.hasUnsavedChanges());
    tag("blocksave"); key(true);
    assert(editor->text()=="edited" && model->item("label").asText()=="edited" && stored()=="first" && form.hasUnsavedChanges() && currents()==1);
    tag(""); key(true);
    assert(editor->text()=="second" && editor->hasFocus() && stored()=="edited" && !form.hasUnsavedChanges() && currents()==2);
    auto* choice=form.findChild<QComboBox*>("access-control-choice"); assert(choice);
    choice->setFocus(); assert(QTest::qWaitFor([&] { return QApplication::focusWidget()==choice; },2000));
    QTest::keyClick(choice,Qt::Key_PageUp,Qt::ControlModifier); assert(editor->text()=="edited" && choice->currentText()=="edited" && choice->hasFocus());
    QTest::keyClick(choice,Qt::Key_PageDown,Qt::ControlModifier); assert(editor->text()=="second" && choice->currentText()=="second" && choice->hasFocus());
    focus();
    QTest::keyClick(editor,Qt::Key_PageDown); assert(editor->text()=="second");
    QTest::keyClick(editor,Qt::Key_PageUp); assert(editor->text()=="second");
    QTest::keyClick(editor,Qt::Key_PageDown,Qt::ControlModifier|Qt::ShiftModifier); assert(editor->text()=="second");
    key(false); assert(editor->text()=="edited"); key(false); assert(editor->text()=="edited" && count()==3);
    std::vector<VbaValue> none;
    assert(form.executeProcedure("SendNext",none).status==linaccess::VbaExecutionStatus::Completed && editor->text()=="second");
    assert(form.executeProcedure("SendPrevious",none).status==linaccess::VbaExecutionStatus::Completed && editor->text()=="edited");
    key(true); key(true); assert(editor->text()=="third");
    key(true); assert(model->member("NewRecord",{}).truthy() && count()==3 && !form.hasUnsavedChanges());
    key(false); assert(editor->text()=="third" && !model->member("NewRecord",{}).truthy());
    model->assign("AllowAdditions",VbaValue::logical(false)); key(true);
    assert(editor->text()=="third" && count()==3 && !model->member("NewRecord",{}).truthy());
    int finished=0; QObject::connect(&form,&QDialog::finished,[&] { ++finished; });
    tag("closecurrent"); key(false); assert(form.isHidden() && finished==1 && count()==3 && stored()=="edited");
}
