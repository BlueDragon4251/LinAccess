#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include <QAbstractButton>
#include <QComboBox>
#include <QLineEdit>
#include <QMessageBox>
#include <QSqlQuery>
#include <QTest>
#include <QTimer>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

void testRecordDelete() {
    using namespace linaccess;
    auto db=QSqlDatabase::addDatabase("QSQLITE","record-delete"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("PRAGMA foreign_keys=ON"));
    assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(sql.exec("INSERT INTO records VALUES(1,'first'),(2,NULL),(3,'third')"));
    const QByteArray definition=R"FORM(Begin Form
RecordSource ="records"
OnCurrent ="[Event Procedure]"
BeforeUpdate ="[Event Procedure]"
OnDelete ="[Event Procedure]"
BeforeDelConfirm ="[Event Procedure]"
AfterDelConfirm ="[Event Procedure]"
Begin Section
Begin TextBox
Name ="value"
ControlSource ="value"
End
Begin TextBox
Name ="mode"
Top =400
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub SelectRow()
DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
End Sub
Sub DeleteRow()
Application.DoCmd.RunCommand acCmdDeleteRecord
End Sub
Sub LegacyDelete()
DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
End Sub
Sub Form_Current()
Me.Tag=Me.Tag & "C"
End Sub
Sub Form_BeforeUpdate(Cancel As Integer)
If Me.mode=4 Then
Cancel=True
End If
End Sub
Sub Form_Delete(Cancel As Integer)
Me.Tag=Me.Tag & "D"
If Me.mode=3 Then
Cancel=True
End If
If Me.mode=5 Then
Me.value="forbidden edit"
End If
End Sub
Sub Form_BeforeDelConfirm(Cancel As Integer,Response As Integer)
Me.Tag=Me.Tag & "B"
If Me.mode=1 Then
Cancel=True
End If
If Me.mode=2 Then
Response=acDataErrContinue
End If
If Me.mode=6 Then
Response=2
End If
If Me.mode=7 Then
Response=Null
End If
End Sub
Sub Form_AfterDelConfirm(Status As Integer)
Me.Tag=Me.Tag & "A" & Status
End Sub
)VBA";
    AccessFormPreview form("record deletion",definition,code,db); form.show(); form.activateWindow(); QApplication::processEvents();
    auto root=makePreviewFormObject(form),app=makePreviewApplicationObject(form);
    auto command=app->member("DoCmd",{}).object;
    auto clone=root->member("Recordset",{}).object->member("Clone",{}).object;
    clone->member("MoveFirst",{}); const auto removed=clone->member("Bookmark",{});
    clone->member("MoveNext",{}); const auto survivor=clone->member("Bookmark",{});
    const auto count=[&] { assert(sql.exec("SELECT count(*) FROM records") && sql.next()); const auto count=sql.value(0).toInt(); sql.finish(); return count; };
    const auto execute=[&](const char* name,int mode) {
        root->assignItem("mode",VbaValue::numeric(mode)); root->assign("Tag",VbaValue::string(""));
        std::vector<VbaValue> arguments; return form.executeProcedure(name,arguments);
    };
    assert(execute("SelectRow",0).status==VbaExecutionStatus::Completed && root->member("SelHeight",{}).asNumber()==1);
    assert(count()==3 && !root->member("Dirty",{}).truthy());
    auto* selector=form.findChild<QComboBox*>("access-record-selector"); assert(selector && selector->hasFocus());
    assert(!selector->styleSheet().isEmpty());
    for (const auto& args:std::vector<std::vector<VbaValue>>{
        {VbaValue::numeric(0),VbaValue::numeric(1),VbaValue::numeric(8)},
        {VbaValue::numeric(0),VbaValue::numeric(1),VbaValue::numeric(8),VbaValue::empty(),VbaValue::numeric(20)},
        {VbaValue::numeric(0),VbaValue::numeric(1),VbaValue::numeric(2),VbaValue::empty(),VbaValue::numeric(70)}}) {
        int error=0; try { command->member("DoMenuItem",args); } catch (const VbaRuntimeError& e) { error=e.number; }
        assert(error==3251 && count()==3);
    }
    auto* editor=form.findChild<QLineEdit*>("access-control-value"); assert(editor); editor->setFocus(); QApplication::processEvents();
    assert(root->member("SelHeight",{}).asNumber()==0 && selector->styleSheet().isEmpty());
    root->assignItem("mode",VbaValue::numeric(4)); root->assignItem("value",VbaValue::string("unsaved"));
    auto result=execute("DeleteRow",4);
    assert(result.errorNumber==2501 && count()==3 && root->member("Dirty",{}).truthy());
    assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="first"); sql.finish();
    root->member("Undo",{});
    result=execute("DeleteRow",5);
    assert(result.errorNumber==3251 && count()==3 && root->item("value").asText()=="first");
    for (int mode:{6,7}) {
        result=execute("DeleteRow",mode);
        assert(result.errorNumber==(mode==6 ? 5 : 94) && count()==3 && root->member("Bookmark",{}).bytes==removed.bytes);
    }
    result=execute("DeleteRow",3);
    assert(result.errorNumber==2501 && count()==3 && root->member("Tag",{}).asText()=="D");
    assert(root->member("Bookmark",{}).bytes==removed.bytes);
    result=execute("DeleteRow",1);
    assert(result.errorNumber==2501 && count()==3 && root->member("Tag",{}).asText()=="DCBA1");
    assert(root->member("Bookmark",{}).bytes==removed.bytes);
    int confirmations=0;
    const auto confirm=[&](QMessageBox::StandardButton button) {
        const int expected=count()-1;
        auto* timer=new QTimer(&form); timer->setInterval(1);
        QObject::connect(timer,&QTimer::timeout,&form,[&,timer,button,expected] {
            auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());
            if (!box) return;
            assert(box->objectName()=="access-delete-confirm" && count()==expected);
            int number=0;
            try { command->member("RunCommand",{VbaValue::numeric(223)}); } catch (const VbaRuntimeError& e) { number=e.number; } assert(number==2046);
            number=0; try { command->member("GoToRecord",{VbaValue::empty(),VbaValue::empty(),VbaValue::numeric(1)}); }
            catch (const VbaRuntimeError& e) { number=e.number; } assert(number==3251);
            assert(!form.close());
            timer->stop(); timer->deleteLater(); ++confirmations;
            QTest::mouseClick(box->button(button),Qt::LeftButton);
        });
        timer->start();
    };
    confirm(QMessageBox::No); result=execute("DeleteRow",0);
    assert(result.errorNumber==2501 && count()==3 && root->member("Tag",{}).asText()=="DCBA2");
    assert(root->member("Bookmark",{}).bytes==removed.bytes);
    result=execute("LegacyDelete",2);
    assert(result.status==VbaExecutionStatus::Completed && count()==2 && root->member("Tag",{}).asText()=="DCBA0");
    assert(root->member("Bookmark",{}).bytes==survivor.bytes);
    int error=0; try { clone->assign("Bookmark",removed); } catch (const VbaRuntimeError& e) { error=e.number; } assert(error==3159);
    clone->assign("Bookmark",survivor); assert(clone->item("id").asNumber()==2 && clone->member("RecordCount",{}).asNumber()==2);
    confirm(QMessageBox::Yes); result=execute("DeleteRow",0);
    assert(result.status==VbaExecutionStatus::Completed && count()==1 && root->item("value").asText()=="third");
    assert(execute("SelectRow",0).status==VbaExecutionStatus::Completed);
    confirm(QMessageBox::Yes); QTest::keyClick(selector,Qt::Key_Delete); QApplication::processEvents();
    assert(count()==0 && root->member("NewRecord",{}).truthy() && root->member("SelHeight",{}).asNumber()==0);
    assert(confirmations==3);
    result=execute("DeleteRow",2); assert(result.errorNumber==2046 && count()==0);
    assert(form.close());

    assert(sql.exec("INSERT INTO records VALUES(10,'referenced')"));
    assert(sql.exec("CREATE TABLE children(parent INTEGER REFERENCES records(id))"));
    assert(sql.exec("INSERT INTO children VALUES(10)"));
    auto protectedDefinition=definition; protectedDefinition.replace("RecordSource =", "AllowDeletions =0\nRecordSource =");
    AccessFormPreview protectedForm("protected record",protectedDefinition,code,db);
    auto protectedRoot=makePreviewFormObject(protectedForm); protectedRoot->assignItem("mode",VbaValue::numeric(2));
    std::vector<VbaValue> arguments; result=protectedForm.executeProcedure("DeleteRow",arguments);
    assert(result.errorNumber==2046 && count()==1 && !protectedRoot->member("AllowDeletions",{}).truthy());
    protectedRoot->assign("AllowDeletions",VbaValue::logical(true));
    result=protectedForm.executeProcedure("DeleteRow",arguments);
    assert(result.errorNumber==3200 && count()==1 && protectedRoot->item("value").asText()=="referenced");
    assert(protectedForm.close());

    AccessFormPreview report("read-only report","Begin Report\nRecordSource =\"records\"\nEnd\n",{},db);
    auto reportRoot=makePreviewFormObject(report);
    for (const auto* property:{"SelHeight","AllowDeletions"}) {
        int number=0; try { reportRoot->member(property,{}); } catch (const VbaRuntimeError& e) { number=e.number; } assert(number==438);
    }
    error=0; try { makePreviewApplicationObject(report)->member("DoCmd",{}).object->member("RunCommand",{VbaValue::numeric(223)}); }
    catch (const VbaRuntimeError& e) { error=e.number; } assert(error==2046 && count()==1 && report.close());

    assert(sql.exec("CREATE TABLE other(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(sql.exec("INSERT INTO other VALUES(42,'target')"));
    const QByteArray callerCode="Sub DeleteActive()\nDoCmd.RunCommand acCmdDeleteRecord\nMe.Tag=\"caller\"\nEnd Sub\n";
    AccessFormPreview caller("delete caller","Begin Form\nEnd\n",callerCode,db); caller.show();
    auto targetDefinition=definition; targetDefinition.replace("RecordSource =\"records\"","RecordSource =\"other\"");
    AccessFormPreview target("delete target",targetDefinition,code,db); target.show(); target.activateWindow(); QApplication::processEvents();
    auto targetRoot=makePreviewFormObject(target); targetRoot->assignItem("mode",VbaValue::numeric(2)); targetRoot->assign("Tag",VbaValue::string(""));
    result=caller.executeProcedure("DeleteActive",arguments);
    assert(result.status==VbaExecutionStatus::Completed && count()==1);
    assert(sql.exec("SELECT count(*) FROM other") && sql.next() && sql.value(0).toInt()==0); sql.finish();
    assert(targetRoot->member("Tag",{}).asText()=="DCBA0" && makePreviewFormObject(caller)->member("Tag",{}).asText()=="caller");
    assert(target.close() && caller.close());
}
