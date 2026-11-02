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
    auto db=QSqlDatabase::addDatabase("QSQLITE","character-functions"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE letters(id INTEGER PRIMARY KEY,label TEXT,unicode TEXT)"));
    assert(sql.exec("INSERT INTO letters VALUES(1,'A','Ω')"));
    assert(sql.exec("CREATE TABLE character_audit(code INTEGER,ascii TEXT,unicode TEXT)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="letters"
Begin Section
Begin TextBox
Name ="label"
ControlSource ="label"
End
Begin TextBox
Name ="unicode"
ControlSource ="unicode"
End
Begin TextBox
Name ="asciicode"
ControlSource ="=Asc([label])"
End
Begin TextBox
Name ="unicodecode"
ControlSource ="=AscW([unicode])"
End
Begin TextBox
Name ="unicodechar"
ControlSource ="=ChrW(937)"
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub Capture()
Dim q As DAO.QueryDef
Set q=CurrentDb.CreateQueryDef("","PARAMETERS p Long,c Text(255),u Text(255); INSERT INTO character_audit VALUES(p,c,u)")
q.Parameters("p").Value=Asc(Me!label)
q.Parameters("c").Value=Chr(Asc(Me!label))
q.Parameters("u").Value=ChrW(AscW(Me!unicode))
q.Execute
End Sub
Sub RejectANSI()
Dim value As Integer
value=Asc(Me!unicode)
CurrentDb.Execute "INSERT INTO character_audit VALUES(99,'bad','bad')"
End Sub
Sub RejectCodepage()
Dim value As String
value=Chr(128)
CurrentDb.Execute "INSERT INTO character_audit VALUES(99,'bad','bad')"
End Sub
Sub RejectSurrogate()
Dim value As String
value=ChrW(55357)
CurrentDb.Execute "INSERT INTO character_audit VALUES(99,'bad','bad')"
End Sub
Sub RejectRange()
Dim value As String
value=ChrW(65536)
CurrentDb.Execute "INSERT INTO character_audit VALUES(99,'bad','bad')"
End Sub
)VBA";
    AccessFormPreview form("character functions",layout,code,db); form.show(); form.activateWindow(); QApplication::processEvents();
    auto model=makePreviewFormObject(form); std::vector<VbaValue> none;
    assert(model->item("asciicode").number==65 && model->item("asciicode").numericType==VbaValue::NumericType::Integer);
    assert(model->item("unicodecode").number==937 && model->item("unicodechar").asText()=="Ω");
    assert(form.executeProcedure("Capture",none).status==linaccess::VbaExecutionStatus::Completed);
    auto* editor=form.findChild<QLineEdit*>("access-control-label"); assert(editor); editor->setFocus(); editor->selectAll(); QTest::keyClicks(editor,"B");
    assert(editor->text()=="B" && model->item("label").asText()=="A" && form.hasUnsavedChanges());
    assert(form.executeProcedure("Capture",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(sql.exec("SELECT label FROM letters WHERE id=1") && sql.next() && sql.value(0).toString()=="A"); sql.finish();
    model->assign("Dirty",VbaValue::logical(false)); assert(model->item("asciicode").number==66);
    assert(form.executeProcedure("Capture",none).status==linaccess::VbaExecutionStatus::Completed);
    for(const auto* name:{"RejectANSI","RejectCodepage","RejectSurrogate","RejectRange"}) {
        const auto result=form.executeProcedure(name,none); assert(result.status!=linaccess::VbaExecutionStatus::Completed && result.errorNumber==(QString(name)=="RejectRange" ? 5 : 3251));
    }
    assert(sql.exec("SELECT code,ascii,unicode FROM character_audit ORDER BY rowid"));
    for(const auto* expected:{"A","A","B"}) assert(sql.next() && sql.value(0).toInt()==QString(expected)[0].unicode() && sql.value(1).toString()==expected && sql.value(2).toString()=="Ω");
    assert(!sql.next()); sql.finish();
    assert(sql.exec("SELECT label,unicode FROM letters WHERE id=1") && sql.next() && sql.value(0).toString()=="B" && sql.value(1).toString()=="Ω"); sql.finish();
    assert(form.close());
}
