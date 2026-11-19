#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include <QLineEdit>
#include <QSqlQuery>
#include <QTimer>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <cstdio>

namespace {
void testPeerOpenDuringClose(QSqlDatabase db, const QByteArray& layout) {
    const QByteArray targetCode=R"VBA(Sub Form_Unload(Cancel As Integer)
If Me.Tag="cancel" Then
DoEvents
Cancel=True
End If
End Sub
Sub Form_Close()
DoEvents
End Sub
)VBA";
    const QByteArray peerCode=R"VBA(Sub OpenTarget()
DoCmd.OpenForm "closing_target", , , "[id]=1", , acHidden
End Sub
)VBA";
    AccessFormPreview target("closing_target",layout,targetCode,db);
    AccessFormPreview peer("opening_peer","Begin Form\nEnd\n",peerCode,db);
    target.show();
    const auto targetObject=makePreviewFormObject(target);

    const auto openFromPeerDuringClose=[&](bool cancel) {
        bool attempted=false;
        QTimer::singleShot(0,&target,[&] {
            attempted=true;
            std::vector<linaccess::VbaValue> arguments;
            const auto result=peer.executeProcedure("OpenTarget",arguments);
            assert(result.status==linaccess::VbaExecutionStatus::Error && result.errorNumber==2501);
        });
        const bool closed=target.close();
        assert(attempted && closed!=cancel);
    };

    targetObject->assign("Tag",linaccess::VbaValue::string("cancel"));
    openFromPeerDuringClose(true);
    assert(target.isVisible() && targetObject->member("Filter",{}).asText().empty());
    assert(!targetObject->member("FilterOn",{}).truthy());

    targetObject->assign("Tag",linaccess::VbaValue::string("allow"));
    openFromPeerDuringClose(false);
    assert(target.isHidden());
    assert(makePreviewFormsObject(peer)->member("Count",{}).asNumber()==1);
}
}

void testDocumentLifetimes() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","document-lifetimes"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(query.exec("INSERT INTO records VALUES(1,'before')"));
    assert(query.exec("CREATE TABLE events(value TEXT)"));
    assert(query.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    const QByteArray layout="Begin Form\nRecordSource =\"records\"\nOnUnload =\"[Event Procedure]\"\nOnClose =\"[Event Procedure]\"\nBegin Section\nBegin TextBox\nName =\"value\"\nControlSource =\"value\"\nEnd\nEnd\nEnd\n";
    testPeerOpenDuringClose(db,layout);
    const QByteArray code=R"VBA(Sub Form_Unload(Cancel As Integer)
If Me.Tag="block" Then Cancel=1
If Me.Tag="reenter" Then
On Error Resume Next
DoCmd.OpenForm "target"
If Err.Number<>2501 Then Err.Raise 5
Cancel=1
End If
End Sub
Sub Form_Close()
CurrentDb.Execute "INSERT INTO events VALUES('close')"
End Sub
Sub QuitNow()
DoCmd.Quit
CurrentDb.Execute "INSERT INTO events VALUES('unreachable')"
End Sub
Sub PumpEvents()
Dim result As Integer
result=VBA.DoEvents()
If result<>0 Then Err.Raise 5
DoEvents
End Sub
Sub PumpQuit()
On Error Resume Next
DoEvents
CurrentDb.Execute "INSERT INTO events VALUES('unreachable')"
End Sub
Sub QuitObject()
Application.DoCmd.Quit
End Sub
)VBA";
    query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('form','target',?)"); query.addBindValue(layout); assert(query.exec());
    query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('vba','Form_target',?)"); query.addBindValue(code); assert(query.exec()); query.finish();
    QPointer<AccessFormPreview> target;
    std::shared_ptr<linaccess::VbaObject> object;
    {
        AccessFormPreview caller("caller","Begin Form\nEnd\n","Sub Start()\nDoCmd.OpenForm \"target\"\nDoCmd.Close acForm,\"caller\"\nEnd Sub\n",db);
        std::vector<VbaValue> none;
        assert(caller.executeProcedure("Start",none).status==linaccess::VbaExecutionStatus::Completed);
        for (auto* widget:QApplication::topLevelWidgets())
            if (auto* form=dynamic_cast<AccessFormPreview*>(widget);form && form!=&caller && form->isVisible()) { assert(!target); target=form; }
        assert(target);
        object=makePreviewFormsObject(*target)->item("target").object;
        assert(!target->parentWidget());
    }
    QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete);
    assert(target && target->isVisible() && object->member("Name",{}).asText()=="target");
    auto foreignDb=QSqlDatabase::addDatabase("QSQLITE","lifetime-foreign"); foreignDb.setDatabaseName(":memory:"); assert(foreignDb.open());
    AccessFormPreview foreign("foreign","Begin Report\nEnd\n",{},foreignDb); foreign.show();
    AccessFormPreview peer("peer","Begin Report\nEnd\n",{},db); peer.show();
    auto peerObject=makePreviewFormObject(peer);
    object->assignItem("value",VbaValue::string("saved before cancelled unload"));
    std::vector<VbaValue> none;
    bool pumped=false;
    QTimer::singleShot(0,target,[&] { pumped=true; assert(query.exec("INSERT INTO events VALUES('pumped')")); });
    const auto pumpResult=target->executeProcedure("PumpEvents",none);
    if(pumpResult.status!=linaccess::VbaExecutionStatus::Completed) std::fprintf(stderr,"DoEvents regression status=%d error=%d line=%zu\n",static_cast<int>(pumpResult.status),pumpResult.errorNumber,static_cast<std::size_t>(pumpResult.line));
    assert(pumpResult.status==linaccess::VbaExecutionStatus::Completed && pumped);
    assert(query.exec("SELECT COUNT(*) FROM events WHERE value='pumped'") && query.next() && query.value(0).toInt()==1); query.finish();
    assert(query.exec("DELETE FROM events")); query.finish();
    object->assign("Tag",VbaValue::string("block"));
    for (const auto* procedure : {"QuitNow","QuitObject"}) {
        const auto result=target->executeProcedure(procedure,none);
        assert(result.status==linaccess::VbaExecutionStatus::Error && result.errorNumber==2501);
        assert(target->isVisible() && peer.isVisible() && foreign.isVisible());
        assert(object->member("Name",{}).asText()=="target" && peerObject->member("Name",{}).asText()=="peer");
    }
    assert(query.exec("SELECT value FROM records WHERE id=1") && query.next() && query.value(0).toString()=="saved before cancelled unload"); query.finish();
    assert(query.exec("SELECT COUNT(*) FROM events") && query.next() && query.value(0).toInt()==0); query.finish();
    object->assign("Tag",VbaValue::string("reenter"));
    assert(!AccessFormPreview::closeProjectDocuments(db));
    assert(makePreviewFormsObject(*target)->member("Count",{}).asNumber()==1);
    assert(target->eventResult("OnUnload")->status==linaccess::VbaExecutionStatus::Completed);
    object->assign("Tag",VbaValue::string("allowed"));
    assert(AccessFormPreview::closeProjectDocuments(db));
    assert(!target->isVisible() && !peer.isVisible() && foreign.isVisible());
    assert(query.exec("SELECT COUNT(*) FROM events") && query.next() && query.value(0).toInt()==1); query.finish();
    bool invalid=false; try { (void)object->member("Name",{}); } catch(const linaccess::VbaRuntimeError& error) { invalid=error.number==2450; } assert(invalid);
    AccessFormPreview once("once","Begin Form\nEnd\n",{},db); once.show();
    int finished=0; QObject::connect(&once,&QDialog::finished,[&] { ++finished; });
    assert(once.close() && finished==1);
    // Successful Quit validates all connections before invalidating any document.
    AccessFormPreview quitter("quitter",layout,code,db); quitter.show();
    bool quitCompleted=false;
    QTimer deadline; deadline.setSingleShot(true);
    QObject::connect(&deadline,&QTimer::timeout,[] { QApplication::exit(42); }); deadline.start(3000);
    QTimer::singleShot(0,&quitter,[&] {
        QTimer::singleShot(0,&quitter,[&] { assert(quitter.executeProcedure("QuitNow",none).status==linaccess::VbaExecutionStatus::Completed); });
        quitCompleted=quitter.executeProcedure("PumpQuit",none).status==linaccess::VbaExecutionStatus::Completed;
    });
    assert(QApplication::exec()==0 && quitCompleted); deadline.stop();
    assert(!quitter.isVisible() && !foreign.isVisible());
    assert(query.exec("SELECT COUNT(*) FROM events WHERE value='unreachable'") && query.next() && query.value(0).toInt()==0); query.finish();
    assert(AccessFormPreview::closeProjectDocuments(db)); // repeated success leaves no guard behind
    QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete); assert(!target);
}
