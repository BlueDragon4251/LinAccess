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

int main(int argc,char** argv) {
    QApplication app(argc,argv);
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","isobject-functions"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,label TEXT)")); assert(sql.exec("INSERT INTO records VALUES(1,'keep')"));
    assert(sql.exec("CREATE TABLE object_audit(kind TEXT,answer INTEGER)"));
    const QByteArray layout="Begin Form\nRecordSource =\"records\"\nBegin Section\nBegin TextBox\nName =\"label\"\nControlSource =\"label\"\nWidth =4000\nEnd\nEnd\nEnd\n";
    const QByteArray code=R"VBA(Sub RecordResult(ByVal kind As String,ByVal answer As Boolean)
Dim q As DAO.QueryDef
Set q=CurrentDb.CreateQueryDef("","PARAMETERS n Text(255),v Bit; INSERT INTO object_audit VALUES(n,v)")
q.Parameters("n").Value=kind
q.Parameters("v").Value=answer
q.Execute
End Sub
Sub Capture()
Dim c As Object
Dim rs As DAO.Recordset
Dim f As Object
Set c=Me!label
Call RecordResult("form",IsObject(Me))
Call RecordResult("control",IsObject(c))
Call RecordResult("direct",IsObject(Me!label))
Call RecordResult("collection",IsObject(Me.Controls("label")))
Call RecordResult("control-value",IsObject(c.Value))
Call RecordResult("nothing",IsObject(Nothing))
Call RecordResult("empty",IsObject(Empty))
Call RecordResult("null",IsObject(Null))
Call RecordResult("error-object",IsObject(Err))
Call RecordResult("error-number",IsObject(Err.Number))
Set rs=CurrentDb.OpenRecordset("records")
Set f=rs.Fields("label")
Call RecordResult("recordset",IsObject(rs))
Call RecordResult("field",IsObject(f))
rs.Close
Call RecordResult("closed-rs",IsObject(rs))
Call RecordResult("closed-field",IsObject(f))
Set f=Nothing
Call RecordResult("reset",IsObject(f))
Call RecordResult("scalar",IsObject(42))
End Sub
Sub RejectExplicitValue()
Dim rs As DAO.Recordset
Dim f As Object
Dim answer As Boolean
Set rs=CurrentDb.OpenRecordset("records")
Set f=rs.Fields("label")
rs.Close
answer=IsObject(f.Value)
Call RecordResult("bad",answer)
End Sub
)VBA";
    AccessFormPreview form("isobject functions",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
    auto model=makePreviewFormObject(form); std::vector<VbaValue> none;
    assert(form.executeProcedure("Capture",none).status==linaccess::VbaExecutionStatus::Completed);
    auto* editor=form.findChild<QLineEdit*>("access-control-label"); assert(editor); editor->setFocus(); editor->selectAll(); QTest::keyClicks(editor,"draft");
    assert(form.executeProcedure("Capture",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(editor->text()=="draft" && model->item("label").asText()=="keep" && form.hasUnsavedChanges());
    const auto failed=form.executeProcedure("RejectExplicitValue",none); assert(failed.status!=linaccess::VbaExecutionStatus::Completed && failed.errorNumber==3420);
    const std::vector<std::pair<QString,bool>> cases{{"form",true},{"control",true},{"direct",true},{"collection",true},{"control-value",false},{"nothing",true},{"empty",false},{"null",false},{"error-object",true},{"error-number",false},{"recordset",true},{"field",true},{"closed-rs",true},{"closed-field",true},{"reset",true},{"scalar",false}};
    for(const auto& [name,expected]:cases) {
        sql.prepare("SELECT answer FROM object_audit WHERE kind=? ORDER BY rowid"); sql.addBindValue(name); assert(sql.exec());
        for(int i=0;i<2;++i) assert(sql.next() && sql.value(0).toInt()==(expected ? -1 : 0));
        assert(!sql.next()); sql.finish();
    }
    assert(sql.exec("SELECT COUNT(*) FROM object_audit") && sql.next() && sql.value(0).toInt()==32); sql.finish();
    assert(sql.exec("SELECT label FROM records") && sql.next() && sql.value(0).toString()=="keep"); sql.finish();
    model->member("Undo",{}); assert(editor->text()=="keep" && !form.hasUnsavedChanges()); assert(form.close());
}
