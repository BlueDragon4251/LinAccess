#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <cmath>

int main(int argc,char** argv) {
    QApplication app(argc,argv);
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","integer-arithmetic-documents"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE arithmetic_rows(id INTEGER PRIMARY KEY,qty REAL)"));
    assert(sql.exec("INSERT INTO arithmetic_rows VALUES(1,12.6)"));
    assert(sql.exec("CREATE TABLE arithmetic_audit(remainder INTEGER,quotient INTEGER)"));
    assert(sql.exec("CREATE TABLE power_audit(amount REAL)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="arithmetic_rows"
Begin Section
Begin TextBox
Name ="qty"
ControlSource ="qty"
End
Begin TextBox
Name ="remainder"
ControlSource ="=[qty] Mod 5"
End
Begin TextBox
Name ="quotient"
ControlSource ="=[qty] \ 5"
End
Begin TextBox
Name ="power"
ControlSource ="=[qty] ^ 2"
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub CaptureMath()
CurrentDb.Execute "INSERT INTO arithmetic_audit VALUES("&CStr(Me!remainder)&","&CStr(Me!quotient)&")"
End Sub
Sub CapturePower()
CurrentDb.Execute "INSERT INTO power_audit VALUES("&CStr(Me!power)&")"
End Sub
Sub RejectPowerDomain()
Dim value As Double
value=(-2)^0.5
CurrentDb.Execute "INSERT INTO power_audit VALUES(99)"
End Sub
Sub RejectPowerOverflow()
Dim value As Double
value=1e200^2
CurrentDb.Execute "INSERT INTO power_audit VALUES(99)"
End Sub
Sub RejectZero()
Dim value As Long
value=7 \ 0.5
CurrentDb.Execute "INSERT INTO arithmetic_audit VALUES(99,99)"
End Sub
Sub RejectOverflow()
Dim value As Long
value=CLng(-2147483648) \ CLng(-1)
CurrentDb.Execute "INSERT INTO arithmetic_audit VALUES(99,99)"
End Sub
)VBA";
    AccessFormPreview form("integer arithmetic",layout,code,db);
    auto model=makePreviewFormObject(form);
    assert(model->item("remainder").number==3 && model->item("quotient").number==2);
    std::vector<VbaValue> args;
    auto result=form.executeProcedure("CaptureMath",args); assert(result.status==linaccess::VbaExecutionStatus::Completed);
    assert(model->item("power").kind==VbaValue::Kind::Number && std::abs(model->item("power").number-158.76)<1e-10);
    result=form.executeProcedure("CapturePower",args); assert(result.status==linaccess::VbaExecutionStatus::Completed);
    model->assignItem("qty",VbaValue::numeric(9.5));
    assert(model->member("Dirty",{}).boolean);
    assert(model->item("remainder").kind==VbaValue::Kind::Number && model->item("remainder").number==0 && model->item("quotient").number==2);
    assert(sql.exec("SELECT qty FROM arithmetic_rows") && sql.next() && sql.value(0).toDouble()==12.6); sql.finish();
    result=form.executeProcedure("CaptureMath",args); assert(result.status==linaccess::VbaExecutionStatus::Completed);
    assert(model->item("power").number==90.25);
    result=form.executeProcedure("CapturePower",args); assert(result.status==linaccess::VbaExecutionStatus::Completed);
    result=form.executeProcedure("RejectPowerDomain",args); assert(result.status==linaccess::VbaExecutionStatus::Error && result.errorNumber==5);
    result=form.executeProcedure("RejectPowerOverflow",args); assert(result.status==linaccess::VbaExecutionStatus::Error && result.errorNumber==6);
    assert(sql.exec("SELECT amount FROM power_audit ORDER BY rowid") && sql.next() && std::abs(sql.value(0).toDouble()-158.76)<1e-10);
    assert(sql.next() && sql.value(0).toDouble()==90.25 && !sql.next()); sql.finish();
    result=form.executeProcedure("RejectZero",args); assert(result.status==linaccess::VbaExecutionStatus::Error && result.errorNumber==11);
    result=form.executeProcedure("RejectOverflow",args); assert(result.status==linaccess::VbaExecutionStatus::Error && result.errorNumber==6);
    assert(sql.exec("SELECT remainder,quotient FROM arithmetic_audit ORDER BY rowid") && sql.next() && sql.value(0).toInt()==3 && sql.value(1).toInt()==2);
    assert(sql.next() && sql.value(0).toInt()==0 && sql.value(1).toInt()==2 && !sql.next()); sql.finish();
    model->assign("Dirty",VbaValue::logical(false));
    assert(sql.exec("SELECT qty FROM arithmetic_rows") && sql.next() && sql.value(0).toDouble()==9.5); sql.finish();
    assert(form.close());
}
