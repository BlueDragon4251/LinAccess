#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include <QLineEdit>
#include <QSqlQuery>
#include <QTest>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <tuple>

void testActiveDocumentCommands() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","active-document-commands"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db); assert(query.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(query.exec("INSERT INTO records VALUES(1,'first'),(2,'second'),(3,'third')"));
    assert(query.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    const QByteArray definition="Begin Form\nRecordSource =\"records\"\nBegin Section\nBegin TextBox\nName =\"value\"\nControlSource =\"value\"\nEnd\nEnd\nEnd\n";
    query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('form','target',?)"); query.addBindValue(definition); assert(query.exec()); query.finish();
    auto callerDefinition=definition; callerDefinition.replace("RecordSource =\"records\"",""); callerDefinition.replace("ControlSource =\"value\"","");
    callerDefinition.replace("End\nEnd\nEnd\n","End\nBegin TextBox\nName =\"unset\"\nVisible =0\nEnd\nEnd\nEnd\n");
    const QByteArray code=R"VBA(Sub OpenAndType()
DoCmd.OpenForm "target"
DoCmd.GoToControl "value"
SendKeys "^anative",True
Application.DoCmd.RunCommand acCmdSaveRecord
Me.Tag=Forms("target")!value
End Sub
Sub NextRecord()
Application.DoCmd.GoToRecord , , acNext
Me.Tag=Forms("target")!value
End Sub
Sub LastRecord()
DoCmd.GoToRecord , , acLast
Me.Tag=Forms("target")!value
End Sub
Sub AddRecord()
DoCmd.GoToRecord , , acNewRec
Forms("target")!value="added"
DoCmd.RunCommand acCmdSaveRecord
Me.Tag=Forms("target")!value
End Sub
Sub RefreshActive()
Application.DoCmd.Requery
Me.Tag=Forms("target")!value
End Sub
Sub MaxActive()
Application.DoCmd.Maximize
End Sub
Sub ExplicitMe()
Me!value="explicit caller"
Me.Dirty=False
End Sub
Sub TypeActive()
Application.DoCmd.GoToControl "value"
SendKeys "^asibling input",True
DoCmd.RunCommand acCmdSaveRecord
Me.Tag=Screen.ActiveForm.Name
End Sub
Sub CallerOnly()
DoCmd.GoToControl "value"
SendKeys "^acaller typed",True
End Sub
Sub QueuedThenOpen()
DoCmd.GoToControl "value"
SendKeys "^alater",False
DoCmd.OpenForm "target"
DoCmd.GoToControl "value"
End Sub
Sub SaveActive()
DoCmd.RunCommand acCmdSaveRecord
End Sub
Sub ObjectSql()
Application.DoCmd.RunSQL "CREATE TABLE command_actions(id INTEGER,value TEXT)"
Application.DoCmd.RunSQL "INSERT INTO command_actions VALUES(1,'inserted')"
Application.DoCmd.RunSQL "UPDATE command_actions SET value='updated' WHERE id=1"
End Sub
Sub ObjectDelete()
Application.DoCmd.RunSQL "DELETE FROM command_actions WHERE id=1"
End Sub
Sub ObjectDrop()
Application.DoCmd.RunSQL "DROP TABLE command_actions"
End Sub
)VBA";
    AccessFormPreview owner("controller",callerDefinition,code,db); owner.show(); QApplication::setActiveWindow(&owner); QApplication::processEvents();
    auto root=makePreviewFormObject(owner); root->assignItem("value",VbaValue::string("caller"));
    auto* caller=owner.findChild<QLineEdit*>("access-control-value"); assert(caller); caller->setFocus();
    std::vector<VbaValue> none;
    assert(owner.executeProcedure("OpenAndType",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(query.exec("SELECT value FROM records WHERE id=1") && query.next() && query.value(0).toString()=="native"); query.finish();
    assert(caller->text()=="caller" && root->member("Tag",{}).asText()=="native");
    assert(owner.executeProcedure("NextRecord",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(root->member("Tag",{}).asText()=="second");
    assert(owner.executeProcedure("LastRecord",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(root->member("Tag",{}).asText()=="third");
    assert(owner.executeProcedure("AddRecord",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(root->member("Tag",{}).asText()=="added");
    assert(query.exec("SELECT id,value FROM records WHERE id=4") && query.next() && query.value(1).toString()=="added"); query.finish();
    assert(query.exec("UPDATE records SET value='refreshed' WHERE id=4")); query.finish();
    assert(owner.executeProcedure("RefreshActive",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(root->member("Tag",{}).asText()=="refreshed");
    assert(owner.executeProcedure("ExplicitMe",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(caller->text()=="explicit caller" && makePreviewFormsObject(owner)->item("target").object->item("value").asText()=="refreshed");
    assert(owner.executeProcedure("MaxActive",none).status==linaccess::VbaExecutionStatus::Completed);
    AccessFormPreview* maximized=nullptr;
    for (auto* widget:QApplication::topLevelWidgets()) if (auto* candidate=dynamic_cast<AccessFormPreview*>(widget);candidate && candidate!=&owner && candidate->isVisible()) { assert(!maximized); maximized=candidate; }
    assert(maximized && maximized->isMaximized() && !owner.isMaximized());
    auto application=makePreviewApplicationObject(owner);
    auto commands=application->member("DoCmd",{}).object;
    for (const auto& [name,args,error]:std::vector<std::tuple<std::string,std::vector<VbaValue>,int>>{
        {"RunCommand",{},450},{"RunCommand",{VbaValue{}},94},{"RunCommand",{root->objectMember("unset",{})},94},
        {"RunCommand",{VbaValue::errorValue(2042)},13},{"RunCommand",{VbaValue::numeric(2147483648.0)},6},
        {"RunCommand",{VbaValue::numeric(42)},3251},{"Requery",{root->objectMember("unset",{})},94},
        {"RunSQL",{VbaValue{}},94},{"RunSQL",{},450},{"RunSQL",{VbaValue::string("SELECT 1"),VbaValue::logical(false)},3251},
        {"GoToRecord",{VbaValue::empty(),VbaValue::empty(),VbaValue::numeric(1),VbaValue::numeric(2)},3251},
        {"Maximize",{VbaValue::numeric(1)},450},{"Requery",{VbaValue::string("value")},3251}}) {
        bool rejected=false;
        try { (void)commands->member(name,args); } catch(const linaccess::VbaRuntimeError& failure) { rejected=failure.number==error; }
        assert(rejected);
    }
    assert(owner.executeProcedure("ObjectSql",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(query.exec("SELECT value FROM command_actions WHERE id=1") && query.next() && query.value(0).toString()=="updated"); query.finish();
    assert(owner.executeProcedure("ObjectDelete",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(query.exec("SELECT COUNT(*) FROM command_actions") && query.next() && query.value(0).toInt()==0); query.finish();
    assert(owner.executeProcedure("ObjectDrop",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(!db.tables().contains("command_actions"));
    application->member("DoCmd",{}).object->member("Close",{VbaValue::numeric(2),VbaValue::string("target")});
    QApplication::setActiveWindow(&owner); caller->setFocus(); QApplication::processEvents();
    assert(owner.executeProcedure("CallerOnly",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(caller->text()=="caller typed");
    assert(owner.executeProcedure("QueuedThenOpen",none).status==linaccess::VbaExecutionStatus::Completed);
    auto target=makePreviewFormsObject(owner)->item("target").object;
    assert(target->item("value").asText()=="native" && caller->text()=="caller typed");
    QApplication::processEvents();
    assert(target->item("value").asText()=="later" && caller->text()=="caller typed");
    assert(owner.executeProcedure("SaveActive",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(query.exec("SELECT value FROM records WHERE id=1") && query.next() && query.value(0).toString()=="later"); query.finish();
    AccessFormPreview sibling("sibling",definition,{},db,nullptr,false,"[id]=3"); sibling.show(); QApplication::setActiveWindow(&sibling); QApplication::processEvents();
    assert(owner.executeProcedure("TypeActive",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(root->member("Tag",{}).asText()=="sibling");
    assert(query.exec("SELECT value FROM records WHERE id=3") && query.next() && query.value(0).toString()=="sibling input"); query.finish();
    assert(query.exec("SELECT value FROM records WHERE id=1") && query.next() && query.value(0).toString()=="later"); query.finish();
    auto foreignDb=QSqlDatabase::addDatabase("QSQLITE","foreign-active-document"); foreignDb.setDatabaseName(":memory:"); assert(foreignDb.open());
    QSqlQuery foreignQuery(foreignDb); assert(foreignQuery.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(foreignQuery.exec("INSERT INTO records VALUES(1,'foreign')"));
    AccessFormPreview foreign("foreign",definition,{},foreignDb); foreign.show(); QApplication::setActiveWindow(&foreign); QApplication::processEvents();
    makePreviewFormObject(foreign)->assignItem("value",VbaValue::string("unsaved foreign"));
    assert(owner.executeProcedure("SaveActive",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(foreignQuery.exec("SELECT value FROM records WHERE id=1") && foreignQuery.next() && foreignQuery.value(0).toString()=="foreign"); foreignQuery.finish();
    assert(owner.close() && sibling.isVisible()); assert(sibling.close());
    assert(AccessFormPreview::closeProjectDocuments(db)); assert(foreign.isVisible()); assert(foreign.close());
}
