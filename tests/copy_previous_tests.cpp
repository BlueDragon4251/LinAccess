#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include <QClipboard>
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
    auto db=QSqlDatabase::addDatabase("QSQLITE","copy-previous-field"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE entries(id INTEGER PRIMARY KEY,label TEXT)"));
    sql.prepare("INSERT INTO entries VALUES(?,?)");
    const QString first="ö 'first' 😀";
    for(int i=1;i<=5;++i) { sql.bindValue(0,i); sql.bindValue(1,i==1 ? QVariant(first) : i==2 ? QVariant("second") : i==3 ? QVariant{} : i==4 ? QVariant(QString("")) : QVariant("last")); assert(sql.exec()); } sql.finish();
    assert(sql.exec("CREATE TABLE events(kind TEXT)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="entries"
KeyPreview =-1
OnKeyDown ="[Event Procedure]"
Begin Section
Begin TextBox
Name ="label"
ControlSource ="label"
OnChange ="[Event Procedure]"
BeforeUpdate ="[Event Procedure]"
AfterUpdate ="[Event Procedure]"
Width =4000
Height =400
End
Begin TextBox
Name ="id"
ControlSource ="id"
Locked =-1
TabStop =0
Left =4200
Width =1000
Height =400
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub Navigate(ByVal target As Long)
DoCmd.GoToRecord , , acGoTo, target
End Sub
Sub RunNavigation(ByVal command As Long)
DoCmd.RunCommand command
End Sub
Sub NewViaCommand()
DoCmd.RunCommand acCmdRecordsGoToNew
End Sub
Sub CopyViaSendKeys()
SendKeys "^'", True
End Sub
Sub NewViaSendKeys()
SendKeys "^{+}", True
End Sub
Sub NewRow()
DoCmd.GoToRecord , , acNewRec
End Sub
Sub Form_KeyDown(KeyCode As Integer,Shift As Integer)
If KeyCode=222 And Shift=acCtrlMask And Me.Tag="blockkey" Then KeyCode=0
If KeyCode=187 And Me.Tag="blocknew" Then KeyCode=0
End Sub
Sub label_Change()
CurrentDb.Execute "INSERT INTO events VALUES('change')"
End Sub
Sub label_BeforeUpdate(Cancel As Integer)
CurrentDb.Execute "INSERT INTO events VALUES('before')"
If Me.Tag="blockupdate" Then Cancel=True
End Sub
Sub label_AfterUpdate()
CurrentDb.Execute "INSERT INTO events VALUES('after')"
End Sub
)VBA";
    AccessFormPreview form("copy previous",layout,code,db); form.show(); form.activateWindow();
    auto model=makePreviewFormObject(form);
    auto* editor=form.findChild<QLineEdit*>("access-control-label"); assert(editor);
    const auto focus=[&] { editor->setFocus(); assert(QTest::qWaitFor([&] { return QApplication::focusWidget()==editor; },2000)); };
    const auto navigate=[&](int row) { std::vector<VbaValue> args{VbaValue::numeric(row)}; assert(form.executeProcedure("Navigate",args).status==linaccess::VbaExecutionStatus::Completed); focus(); };
    const auto copy=[&] { QTest::keyClick(editor,Qt::Key_Apostrophe,Qt::ControlModifier); };
    const auto stored=[&](int row) { sql.prepare("SELECT label FROM entries WHERE id=?"); sql.addBindValue(row); assert(sql.exec() && sql.next()); const auto result=sql.value(0); sql.finish(); return result; };
    QApplication::clipboard()->setText("unchanged");
    navigate(2); model->assign("Tag",VbaValue::string("blockkey")); copy(); assert(editor->text()=="second" && !model->member("Dirty",{}).boolean);
    model->assign("Tag",VbaValue::string("")); copy();
    assert(editor->text()==first && model->item("label").asText()=="second" && model->member("Dirty",{}).boolean);
    assert(stored(2).toString()=="second" && QApplication::clipboard()->text()=="unchanged");
    model->assign("Tag",VbaValue::string("blockupdate"));
    int cancelled=0; try { model->assign("Dirty",VbaValue::logical(false)); } catch(const linaccess::VbaRuntimeError& error) { cancelled=error.number; } assert(cancelled==2501);
    assert(stored(2).toString()=="second" && editor->text()==first && model->member("Dirty",{}).boolean);
    model->assign("Tag",VbaValue::string("")); model->assign("Dirty",VbaValue::logical(false)); assert(stored(2).toString()==first);
    copy(); assert(!model->member("Dirty",{}).boolean);
    assert(sql.exec("SELECT kind FROM events ORDER BY rowid"));
    QStringList sequence; while(sql.next()) sequence<<sql.value(0).toString(); sql.finish();
    assert(sequence==QStringList({"change","before","before","after"}));
    navigate(4); copy(); assert(editor->text().isEmpty() && !model->item("label").asText().size());
    model->assign("Dirty",VbaValue::logical(false)); assert(stored(4).isNull());
    // Restore the previous row to an actual empty string, distinct from NULL.
    assert(sql.exec("UPDATE entries SET label='' WHERE id=4")); model->member("Requery",{});
    navigate(5); copy(); assert(editor->text().isEmpty() && stored(5).toString()=="last");
    model->assign("Dirty",VbaValue::logical(false)); assert(!stored(5).isNull() && stored(5).toString().isEmpty());
    navigate(1); copy(); assert(editor->text()==first && !model->member("Dirty",{}).boolean);
    navigate(2); model->assignItem("label",VbaValue::string("readonly target")); model->assign("Dirty",VbaValue::logical(false));
    editor->setReadOnly(true); copy(); assert(editor->text()=="readonly target" && !model->member("Dirty",{}).boolean); editor->setReadOnly(false);
    model->assignItem("label",VbaValue::string("ok")); model->assign("Dirty",VbaValue::logical(false));
    const auto editorCapacity=editor->maxLength(); editor->setMaxLength(2); copy(); assert(editor->text()=="ok" && !model->member("Dirty",{}).boolean && stored(2).toString()=="ok"); editor->setMaxLength(editorCapacity);
    std::vector<VbaValue> none;
    assert(form.executeProcedure("CopyViaSendKeys",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(editor->text()==first && stored(2).toString()=="ok" && model->member("Dirty",{}).boolean);
    QTest::keyClick(editor,Qt::Key_Escape); assert(editor->text()=="ok" && !model->member("Dirty",{}).boolean);
    model->assign("Tag",VbaValue::string("blocknew")); QTest::keyClick(editor,Qt::Key_Plus,Qt::ControlModifier);
    assert(!model->member("NewRecord",{}).truthy() && !model->member("Dirty",{}).boolean);
    model->assign("Tag",VbaValue::string("")); model->assign("AllowAdditions",VbaValue::logical(false));
    QTest::keyClick(editor,Qt::Key_Plus,Qt::ControlModifier); assert(!model->member("NewRecord",{}).truthy() && stored(2).toString()=="ok");
    model->assign("AllowAdditions",VbaValue::logical(true));
    editor->selectAll(); QTest::keyClicks(editor,"pending");
    model->assign("AllowAdditions",VbaValue::logical(false)); QTest::keyClick(editor,Qt::Key_Plus,Qt::ControlModifier);
    assert(stored(2).toString()=="ok" && model->member("Dirty",{}).boolean && editor->text()=="pending");
    model->assign("AllowAdditions",VbaValue::logical(true)); model->assign("Tag",VbaValue::string("blockupdate"));
    QTest::keyClick(editor,Qt::Key_Plus,Qt::ControlModifier);
    assert(!model->member("NewRecord",{}).truthy() && model->member("Dirty",{}).boolean && editor->text()=="pending" && stored(2).toString()=="ok");
    model->assign("Tag",VbaValue::string("")); assert(form.executeProcedure("NewViaSendKeys",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(model->member("NewRecord",{}).truthy() && !model->member("Dirty",{}).boolean && stored(2).toString()=="pending");
    focus(); copy();
    assert(model->member("Dirty",{}).boolean);
    QTest::keyClick(editor,Qt::Key_Escape); assert(!model->member("Dirty",{}).boolean);
    assert(sql.exec("SELECT COUNT(*) FROM entries") && sql.next() && sql.value(0).toInt()==5); sql.finish();
    sql.prepare("UPDATE entries SET label=? WHERE id=1"); sql.addBindValue(QByteArray("binary")); assert(sql.exec()); sql.finish();
    model->member("Requery",{}); navigate(2); copy();
    assert(editor->text()=="pending" && stored(2).toString()=="pending" && !model->member("Dirty",{}).boolean);
    QTest::keyClick(editor,Qt::Key_Equal,Qt::ControlModifier|Qt::ShiftModifier);
    assert(model->member("NewRecord",{}).truthy() && !model->member("Dirty",{}).boolean);
    for(const auto& step:std::vector<std::pair<int,int>>{{67,1},{65,2},{68,5},{66,4}}) {
        std::vector<VbaValue> command{VbaValue::numeric(step.first)};
        assert(form.executeProcedure("RunNavigation",command).status==linaccess::VbaExecutionStatus::Completed);
        assert(model->item("id").asNumber()==step.second && !model->member("Dirty",{}).boolean);
    }
    for(const auto incorrect:{24,25,26,27}) {
        std::vector<VbaValue> command{VbaValue::numeric(incorrect)};
        const auto rejected=form.executeProcedure("RunNavigation",command);
        assert(rejected.status==linaccess::VbaExecutionStatus::Error && rejected.errorNumber==3251);
        assert(model->item("id").asNumber()==4 && !model->member("Dirty",{}).boolean);
    }
    assert(form.executeProcedure("NewViaCommand",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(model->member("NewRecord",{}).truthy() && !model->member("Dirty",{}).boolean);
    assert(sql.exec("SELECT COUNT(*) FROM entries") && sql.next() && sql.value(0).toInt()==5); sql.finish();
    assert(form.close());
}
