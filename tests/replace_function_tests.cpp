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
    auto db=QSqlDatabase::addDatabase("QSQLITE","replace-functions"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,label TEXT)"));
    assert(sql.exec("INSERT INTO records VALUES(1,'red/red/red')"));
    assert(sql.exec("CREATE TABLE replace_audit(whole TEXT,tail TEXT,unicode TEXT,comparison TEXT)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="records"
Begin Section
Begin TextBox
Name ="label"
ControlSource ="label"
Width =4000
End
Begin TextBox
Name ="replaced"
ControlSource ="=Replace([label],""red"",""blue"")"
Top =500
Width =4000
End
End
End
)FORM";
    const QByteArray code=R"VBA(Option Compare Text
Sub Capture()
Dim q As DAO.QueryDef
Set q=CurrentDb.CreateQueryDef("","PARAMETERS p Text(255),t Text(255),u Text(255),c Text(255); INSERT INTO replace_audit VALUES(p,t,u,c)")
q.Parameters("p").Value=Replace(Me!label,"red","blue")
q.Parameters("t").Value=Replace(Me!label,"red","blue",5,1)
q.Parameters("u").Value=Replace("😀/😀","😀","Ω")
q.Parameters("c").Value=Replace("a-A","a","X",1,-1,-1)
q.Execute
End Sub
Function DefaultComparison()
DefaultComparison=Replace("a-A","a","X")
End Function
Function UnicodeComparison()
UnicodeComparison=Replace("Ä/ä","ä","Ω",1,-1,1)
End Function
Sub ChangeValue()
Me!label=Replace(Me!label,"red","green",1,1)
End Sub
Sub RejectStart()
Dim value As String
value=Replace(Me!label,"red","blue",0)
CurrentDb.Execute "INSERT INTO replace_audit VALUES('bad','bad','bad','bad')"
End Sub
Sub RejectNull()
Dim value As String
value=Replace(Me!label,Null,"blue")
CurrentDb.Execute "INSERT INTO replace_audit VALUES('bad','bad','bad','bad')"
End Sub
)VBA";
    AccessFormPreview form("replace functions",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
    auto model=makePreviewFormObject(form); std::vector<VbaValue> none;
    auto* editor=form.findChild<QLineEdit*>("access-control-label");
    auto* calculated=form.findChild<QLineEdit*>("access-control-replaced"); assert(editor && calculated && calculated->text()=="blue/blue/blue");
    assert(form.executeProcedure("Capture",none).status==linaccess::VbaExecutionStatus::Completed);
    const auto binary=form.executeProcedure("DefaultComparison",none); assert(binary.status==linaccess::VbaExecutionStatus::Completed && binary.typedReturnValue->asText()=="X-A");
    const auto unicode=form.executeProcedure("UnicodeComparison",none); assert(unicode.status==linaccess::VbaExecutionStatus::Completed && unicode.typedReturnValue->asText()=="Ω/Ω");
    editor->setFocus(); editor->selectAll(); QTest::keyClicks(editor,"red/pending");
    assert(editor->text()=="red/pending" && model->item("label").asText()=="red/red/red" && form.hasUnsavedChanges());
    assert(form.executeProcedure("Capture",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(sql.exec("SELECT label FROM records") && sql.next() && sql.value(0).toString()=="red/red/red"); sql.finish();
    model->assign("Dirty",VbaValue::logical(false)); assert(calculated->text()=="blue/pending");
    assert(form.executeProcedure("Capture",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(form.executeProcedure("ChangeValue",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(editor->text()=="green/pending" && form.hasUnsavedChanges());
    assert(sql.exec("SELECT label FROM records") && sql.next() && sql.value(0).toString()=="red/pending"); sql.finish();
    model->member("Undo",{}); assert(editor->text()=="red/pending" && !form.hasUnsavedChanges());
    const auto badStart=form.executeProcedure("RejectStart",none),badNull=form.executeProcedure("RejectNull",none);
    assert(badStart.errorNumber==5 && badNull.errorNumber==94);
    assert(sql.exec("SELECT whole,tail,unicode,comparison FROM replace_audit ORDER BY rowid"));
    for(int i=0;i<3;++i) {
        assert(sql.next()); assert(sql.value(0).toString()==(i<2 ? "blue/blue/blue" : "blue/pending"));
        assert(sql.value(1).toString()==(i<2 ? "blue/red" : "pending"));
        assert(sql.value(2).toString()=="Ω/Ω" && sql.value(3).toString()=="X-X");
    }
    assert(!sql.next()); sql.finish(); assert(form.close());
}
