#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include <QAbstractButton>
#include <QLineEdit>
#include <QMessageBox>
#include <QPushButton>
#include <QSqlQuery>
#include <QTest>
#include <QTimer>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <cstdlib>

int main(int argc,char** argv) {
    QApplication app(argc,argv);
    QTimer::singleShot(10000,&app,[] { std::abort(); });
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","delete-record-key"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("PRAGMA foreign_keys=ON"));
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(sql.exec("INSERT INTO records VALUES(1,'first'),(2,'second'),(3,'third')"));
    assert(sql.exec("CREATE TABLE refs(parent INTEGER REFERENCES records(id))")); assert(sql.exec("INSERT INTO refs VALUES(3)"));
    assert(sql.exec("CREATE TABLE events(kind TEXT,status INTEGER)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="records"
KeyPreview =-1
OnKeyDown ="[Event Procedure]"
OnDelete ="[Event Procedure]"
BeforeDelConfirm ="[Event Procedure]"
AfterDelConfirm ="[Event Procedure]"
Begin Section
Begin TextBox
Name ="value"
ControlSource ="value"
BeforeUpdate ="[Event Procedure]"
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
    const QByteArray code=R"VBA(Sub Form_KeyDown(KeyCode As Integer,Shift As Integer)
If KeyCode=189 And Me.Tag="blockkey" Then KeyCode=0
End Sub
Sub value_BeforeUpdate(Cancel As Integer)
If Me.Tag="blockupdate" Then Cancel=True
End Sub
Sub Form_Delete(Cancel As Integer)
CurrentDb.Execute "INSERT INTO events VALUES('delete',0)"
If Me.Tag="canceldelete" Then Cancel=True
End Sub
Sub Form_BeforeDelConfirm(Cancel As Integer,Response As Integer)
CurrentDb.Execute "INSERT INTO events VALUES('beforeconfirm',0)"
If Me.Tag="cancelconfirm" Then Cancel=True
End Sub
Sub Form_AfterDelConfirm(Status As Integer)
CurrentDb.Execute "INSERT INTO events VALUES('after',"&CStr(Status)&")"
End Sub
Sub DeleteViaSendKeys()
SendKeys "^-", True
End Sub
)VBA";
    AccessFormPreview form("delete record key",layout,code,db); form.show(); form.activateWindow();
    auto model=makePreviewFormObject(form); auto* editor=form.findChild<QLineEdit*>("access-control-value"); assert(editor);
    const auto focus=[&] { form.activateWindow(); editor->setFocus(); assert(QTest::qWaitFor([&] { return QApplication::focusWidget()==editor; },2000)); };
    const auto count=[&] { assert(sql.exec("SELECT COUNT(*) FROM records") && sql.next()); const int result=sql.value(0).toInt(); sql.finish(); return result; };
    const auto lastStatus=[&] { assert(sql.exec("SELECT status FROM events WHERE kind='after' ORDER BY rowid DESC LIMIT 1") && sql.next()); const int result=sql.value(0).toInt(); sql.finish(); return result; };
    bool expectDialog=false,seen=false,unexpected=false; QMessageBox::StandardButton answer=QMessageBox::No;
    QTimer responder;
    QObject::connect(&responder,&QTimer::timeout,&app,[&] {
        auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget()); if(!box) return;
        if(!expectDialog || box->objectName()!="access-delete-confirm") unexpected=true;
        assert(box->defaultButton()==box->button(QMessageBox::No));
        seen=true; box->button(expectDialog ? answer : QMessageBox::No)->click();
    }); responder.start(1);
    const auto press=[&](bool confirmation=false,QMessageBox::StandardButton response=QMessageBox::No,bool viaSendKeys=false) {
        focus(); expectDialog=confirmation; answer=response; seen=false;
        if(viaSendKeys) { std::vector<VbaValue> none; assert(form.executeProcedure("DeleteViaSendKeys",none).status==linaccess::VbaExecutionStatus::Completed); }
        else QTest::keyClick(editor,Qt::Key_Minus,Qt::ControlModifier);
        assert(!unexpected && seen==confirmation); expectDialog=false;
    };
    model->assign("Tag",VbaValue::string("blockkey")); press(); assert(count()==3 && model->item("id").number==1);
    model->assign("Tag",VbaValue::string("")); model->assign("AllowDeletions",VbaValue::logical(false)); press(); assert(count()==3);
    model->assign("AllowDeletions",VbaValue::logical(true)); model->assign("AllowEdits",VbaValue::logical(false)); press(); assert(count()==3);
    model->assign("AllowEdits",VbaValue::logical(true)); focus(); editor->selectAll(); QTest::keyClicks(editor,"edited");
    model->assign("Tag",VbaValue::string("blockupdate")); press(); assert(count()==3 && editor->text()=="edited" && model->member("Dirty",{}).boolean);
    assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="first"); sql.finish();
    model->assign("Tag",VbaValue::string("")); model->member("Undo",{});
    model->assign("Tag",VbaValue::string("canceldelete")); press(); assert(count()==3 && model->item("id").number==1);
    model->assign("Tag",VbaValue::string("cancelconfirm")); press(); assert(count()==3 && model->item("id").number==1 && lastStatus()==1);
    model->assign("Tag",VbaValue::string("")); press(true,QMessageBox::No); assert(count()==3 && model->item("id").number==1 && lastStatus()==2);
    press(true,QMessageBox::Yes); assert(count()==2 && model->item("id").number==2 && lastStatus()==0);
    press(true,QMessageBox::Yes,true); assert(count()==1 && model->item("id").number==3 && lastStatus()==0);
    assert(sql.exec("DELETE FROM events"));
    press(); assert(count()==1 && model->item("id").number==3 && !model->member("Dirty",{}).boolean);
    assert(sql.exec("SELECT COUNT(*) FROM events WHERE kind='delete'") && sql.next() && sql.value(0).toInt()==1); sql.finish();
    assert(sql.exec("SELECT COUNT(*) FROM events WHERE kind IN ('beforeconfirm','after')") && sql.next() && sql.value(0).toInt()==0); sql.finish();
    assert(sql.exec("SELECT parent FROM refs") && sql.next() && sql.value(0).toInt()==3); sql.finish();
    responder.stop(); assert(form.close());
}
