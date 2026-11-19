#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

int main(int argc,char** argv) {
    QApplication app(argc,argv);
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","byte-values-documents"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE byte_source(id INTEGER PRIMARY KEY,amount REAL)"));
    assert(sql.exec("INSERT INTO byte_source VALUES(1,12.5)"));
    assert(sql.exec("CREATE TABLE byte_audit(value INTEGER NOT NULL CHECK(value BETWEEN 0 AND 255))"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="byte_source"
Begin Section
Begin TextBox
Name ="amount"
ControlSource ="amount"
End
Begin TextBox
Name ="converted"
ControlSource ="=CByte([amount])"
End
Begin TextBox
Name ="inverted"
ControlSource ="=Not CByte([amount])"
End
Begin TextBox
Name ="typecode"
ControlSource ="=VarType(CByte([amount]))"
End
Begin TextBox
Name ="codetype"
ControlSource ="=TypeName(VarType(CByte([amount])))"
End
End
End
)FORM";
    const QByteArray code=R"VBA(Public tracked As Byte
Sub CaptureByte()
Dim value As Byte
Dim q As DAO.QueryDef
If VarType(value)<>vbByte Then Err.Raise(5)
value=CByte(Me!amount)
tracked=value
Set q=CurrentDb.CreateQueryDef("","PARAMETERS p Byte; INSERT INTO byte_audit(value) VALUES(p)")
If q.Parameters("p").Type<>dbByte Then Err.Raise(5)
q.Parameters("p").Value=value
q.Execute
End Sub
Sub CaptureNot()
Dim value As Byte
value=Not CByte(Me!amount)
CurrentDb.Execute "INSERT INTO byte_audit VALUES(" & CStr(value) & ")"
End Sub
Sub RejectNotRange()
Dim value As Long
value=Not 1e100
CurrentDb.Execute "INSERT INTO byte_audit VALUES(99)"
End Sub
Sub CaptureTypeCode()
Dim code As Variant
code=VarType(CByte(Me!amount))
If TypeName(code)<>"Integer" Then Err.Raise(5)
CurrentDb.Execute "INSERT INTO byte_audit VALUES(" & CStr(VarType(code)) & ")"
End Sub
Sub RejectByte()
Dim value As Byte
value=256
CurrentDb.Execute "INSERT INTO byte_audit VALUES(99)"
End Sub
Sub RejectByteArgument()
Call StoreByte(256)
End Sub
Sub StoreByte(ByVal value As Byte)
CurrentDb.Execute "INSERT INTO byte_audit VALUES(" & CStr(value) & ")"
End Sub
Sub CaptureByteLoop()
Dim value As Byte
For value=1 To 2
CurrentDb.Execute "INSERT INTO byte_audit VALUES(" & CStr(value) & ")"
Next value
End Sub
)VBA";
    AccessFormPreview form("byte values",layout,code,db); form.show(); form.activateWindow();
    auto model=makePreviewFormObject(form); assert(model->item("converted").asNumber()==12);
    std::vector<VbaValue> none;
    assert(form.executeProcedure("CaptureByte",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(model->member("tracked",{}).numericType==VbaValue::NumericType::Byte && model->member("tracked",{}).number==12);
    model->assignItem("amount",VbaValue::numeric(3.5)); assert(model->item("converted").asNumber()==4);
    assert(form.executeProcedure("CaptureByte",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(model->item("inverted").asNumber()==251);
    assert(form.executeProcedure("CaptureNot",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(model->item("typecode").number==17 && model->item("typecode").numericType==VbaValue::NumericType::Integer);
    assert(model->item("codetype").asText()=="Integer");
    assert(form.executeProcedure("CaptureTypeCode",none).status==linaccess::VbaExecutionStatus::Completed);
    for(const auto& name:{"RejectByte","RejectByteArgument","RejectNotRange"}) {
        const auto failure=form.executeProcedure(name,none); assert(failure.status==linaccess::VbaExecutionStatus::Error && failure.errorNumber==6);
    }
    assert(form.executeProcedure("CaptureByteLoop",none).status==linaccess::VbaExecutionStatus::Completed);
    int error=0; try { model->assign("tracked",VbaValue::numeric(256)); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
    assert(error==6 && model->member("tracked",{}).number==4);
    assert(sql.exec("SELECT value,typeof(value) FROM byte_audit ORDER BY rowid"));
    assert(sql.next() && sql.value(0).toInt()==12 && sql.value(1).toString()=="integer");
    assert(sql.next() && sql.value(0).toInt()==4 && sql.value(1).toString()=="integer");
    assert(sql.next() && sql.value(0).toInt()==251 && sql.value(1).toString()=="integer");
    assert(sql.next() && sql.value(0).toInt()==2 && sql.value(1).toString()=="integer");
    assert(sql.next() && sql.value(0).toInt()==1 && sql.value(1).toString()=="integer");
    assert(sql.next() && sql.value(0).toInt()==2 && sql.value(1).toString()=="integer" && !sql.next()); sql.finish();
    assert(sql.exec("SELECT amount FROM byte_source WHERE id=1") && sql.next() && sql.value(0).toDouble()==12.5); sql.finish();
    model->assign("Dirty",VbaValue::logical(false));
    assert(sql.exec("SELECT amount FROM byte_source WHERE id=1") && sql.next() && sql.value(0).toDouble()==3.5); sql.finish();
    assert(form.close());
}
