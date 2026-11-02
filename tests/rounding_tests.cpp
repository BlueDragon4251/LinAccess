#include "access_query.hpp"
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
#include <cmath>

int main(int argc,char** argv) {
    QApplication app(argc,argv);
    auto db=QSqlDatabase::addDatabase("QSQLITE","round-documents"); db.setDatabaseName(":memory:"); assert(db.open());
    const auto result=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT Round(2.5) AS a,Round(3.5) AS b,Round(-2.5) AS c,Round(-3.5) AS d,Round(0.12345,4) AS e,Round(NULL) AS n"),{});
    assert(result.rows.size()==1 && result.rows[0][0].toDouble()==2 && result.rows[0][1].toDouble()==4 && result.rows[0][2].toDouble()==-2 && result.rows[0][3].toDouble()==-4);
    assert(std::abs(result.rows[0][4].toDouble()-0.1234)<1e-12 && result.rows[0][5].isNull());
    using linaccess::VbaValue;
    QSqlQuery sql(db);
    assert(sql.exec("SELECT round(2.5)") && sql.next() && sql.value(0).toDouble()==3); sql.finish();
    const auto literal=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT 'Round(2.5)' AS literal"),{}); assert(literal.rows[0][0].toString()=="Round(2.5)");
    for(const auto& [query,code]:std::vector<std::pair<QString,int>>{
        {"SELECT Round(1,-1)",5},{"SELECT Round(1,16)",3251},{"SELECT Round('invalid')",13},
        {"SELECT Round(9007199254740993)",3251},{"SELECT Round()",450},{"SELECT Round(1,2,3)",450}}) {
        int actual=0; try { (void)linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery(query),{}); } catch(const linaccess::VbaRuntimeError& failure) { actual=failure.number; } assert(actual==code);
    }
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE round_rows(id INTEGER PRIMARY KEY,amount REAL)"));
    assert(sql.exec("INSERT INTO round_rows VALUES(1,2.5),(2,3.5)"));
    assert(sql.exec("CREATE TABLE round_audit(vba REAL,sql REAL)"));
    assert(sql.exec("CREATE TABLE round_fixed(value TEXT)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="round_rows"
Begin Section
Begin TextBox
Name ="amount"
ControlSource ="amount"
Width =4000
End
Begin TextBox
Name ="rounded"
ControlSource ="=Round([amount])"
Top =500
Width =4000
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub Capture()
Dim q As DAO.QueryDef
Set q=CurrentDb.CreateQueryDef("","PARAMETERS p Double,v Double; INSERT INTO round_audit(vba,sql) VALUES(v,Round(p))")
q.Parameters("p").Value=Me!amount
q.Parameters("v").Value=Round(Me!amount)
q.Execute
End Sub
Sub CaptureFixed()
Dim value As Currency
Dim q As DAO.QueryDef
value=Round(CCur("922337203685477.1234"),3)
Set q=CurrentDb.CreateQueryDef("","PARAMETERS p Text(255); INSERT INTO round_fixed VALUES(p)")
q.Parameters("p").Value=value
q.Execute
End Sub
Sub RejectPrecision()
Dim value As Double
value=Round(Me!amount,-1)
CurrentDb.Execute "INSERT INTO round_audit VALUES(99,99)"
End Sub
Sub RejectPartialSQL()
CurrentDb.Execute "UPDATE round_rows SET amount=Round(amount,IIf(id=1,0,-1))"
End Sub
Sub RejectFixedOverflow()
Dim value As Currency
value=Round(CCur("922337203685477.5807"),0)
CurrentDb.Execute "INSERT INTO round_fixed VALUES('bad')"
End Sub
)VBA";
    AccessFormPreview form("round documents",layout,code,db); form.show(); form.activateWindow(); QApplication::processEvents();
    auto model=makePreviewFormObject(form); std::vector<VbaValue> none;
    assert(model->item("rounded").number==2);
    assert(form.executeProcedure("Capture",none).status==linaccess::VbaExecutionStatus::Completed);
    auto* editor=form.findChild<QLineEdit*>("access-control-amount"); assert(editor); editor->setFocus(); editor->selectAll(); QTest::keyClicks(editor,"4.5");
    assert(editor->text()=="4.5" && model->item("amount").number==2.5 && form.hasUnsavedChanges());
    assert(form.executeProcedure("Capture",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(sql.exec("SELECT amount FROM round_rows WHERE id=1") && sql.next() && sql.value(0).toDouble()==2.5); sql.finish();
    model->assign("Dirty",VbaValue::logical(false)); assert(model->item("rounded").number==4);
    assert(form.executeProcedure("Capture",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(form.executeProcedure("CaptureFixed",none).status==linaccess::VbaExecutionStatus::Completed);
    for(const auto* name:{"RejectPrecision","RejectPartialSQL","RejectFixedOverflow"}) {
        const auto failure=form.executeProcedure(name,none); assert(failure.status!=linaccess::VbaExecutionStatus::Completed && failure.errorNumber==(QString(name)=="RejectFixedOverflow" ? 6 : 5));
    }
    assert(sql.exec("SELECT vba,sql FROM round_audit ORDER BY rowid"));
    for(const auto expected:{2,2,4}) assert(sql.next() && sql.value(0).toDouble()==expected && sql.value(1).toDouble()==expected);
    assert(!sql.next()); sql.finish();
    assert(sql.exec("SELECT amount FROM round_rows ORDER BY id") && sql.next() && sql.value(0).toDouble()==4.5 && sql.next() && sql.value(0).toDouble()==3.5 && !sql.next()); sql.finish();
    assert(sql.exec("SELECT value,typeof(value) FROM round_fixed") && sql.next() && sql.value(0).toString()=="922337203685477.123" && sql.value(1).toString()=="text" && !sql.next()); sql.finish();
    assert(!form.hasUnsavedChanges() && form.close());
}
