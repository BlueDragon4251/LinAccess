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
    auto db=QSqlDatabase::addDatabase("QSQLITE","default-type-documents"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE defaults(id INTEGER PRIMARY KEY,label TEXT,notes TEXT)"));
    assert(sql.exec("INSERT INTO defaults VALUES(1,'synthetic',NULL)"));
    assert(sql.exec("CREATE TABLE type_audit(control INTEGER,nullable INTEGER,typed INTEGER,field INTEGER,parameter INTEGER,errorcode INTEGER)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="defaults"
Begin Section
Begin TextBox
Name ="label"
ControlSource ="label"
End
Begin TextBox
Name ="nullable"
ControlSource ="notes"
End
Begin TextBox
Name ="typed"
ControlSource ="=CByte(4)"
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub CaptureTypes()
Dim c As Control
Dim n As Control
Dim b As Control
Dim r As DAO.Recordset
Dim f As DAO.Field
Dim q As DAO.QueryDef
Dim p As DAO.Parameter
Set c=Me.Controls("label")
Set n=Me.Controls("nullable")
Set b=Me.Controls("typed")
Set r=CurrentDb.OpenRecordset("defaults")
Set f=r.Fields("label")
Set q=CurrentDb.CreateQueryDef("","PARAMETERS p Text(255); SELECT label FROM defaults WHERE label=p")
Set p=q.Parameters("p")
p.Value="synthetic"
CurrentDb.Execute "INSERT INTO type_audit VALUES(" & CStr(VarType(c)) & "," & CStr(VarType(n)) & "," & CStr(VarType(b)) & "," & CStr(VarType(f)) & "," & CStr(VarType(p)) & "," & CStr(VarType(Err)) & ")"
r.Close
q.Close
End Sub
Function ClosedFieldType() As Long
Dim r As DAO.Recordset
Dim f As DAO.Field
Dim code As Integer
Set r=CurrentDb.OpenRecordset("defaults")
Set f=r.Fields("label")
r.Close
On Error Resume Next
code=VarType(f)
ClosedFieldType=Err.Number
End Function
Function ClosedParameterType() As Long
Dim q As DAO.QueryDef
Dim p As DAO.Parameter
Dim code As Integer
Set q=CurrentDb.CreateQueryDef("","PARAMETERS p Text(255); SELECT label FROM defaults WHERE label=p")
Set p=q.Parameters("p")
q.Close
On Error Resume Next
code=VarType(p)
ClosedParameterType=Err.Number
End Function
Function EndOfFileType() As Long
Dim r As DAO.Recordset
Dim f As DAO.Field
Dim code As Integer
Set r=CurrentDb.OpenRecordset("defaults")
Set f=r.Fields("label")
r.MoveNext
On Error Resume Next
code=VarType(f)
EndOfFileType=Err.Number
r.Close
End Function
Function ClosedFieldsCount() As Long
Dim r As DAO.Recordset
Dim count As Long
Set r=CurrentDb.OpenRecordset("defaults")
r.Close
On Error Resume Next
count=r.Fields.Count
ClosedFieldsCount=Err.Number
End Function
Function ClosedFieldWrite() As Long
Dim r As DAO.Recordset
Dim f As DAO.Field
Set r=CurrentDb.OpenRecordset("defaults")
Set f=r.Fields("label")
r.Close
On Error Resume Next
f.Value="rejected"
ClosedFieldWrite=Err.Number
End Function
)VBA";
    AccessFormPreview form("default value types",layout,code,db); form.show(); QApplication::processEvents();
    std::vector<VbaValue> none;
    assert(form.executeProcedure("CaptureTypes",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(sql.exec("SELECT control,nullable,typed,field,parameter,errorcode FROM type_audit") && sql.next());
    assert(sql.value(0).toInt()==8 && sql.value(1).toInt()==1 && sql.value(2).toInt()==17 && sql.value(3).toInt()==8 && sql.value(4).toInt()==8 && sql.value(5).toInt()==3 && !sql.next()); sql.finish();
    auto result=form.executeProcedure("ClosedFieldType",none); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->number==3420);
    result=form.executeProcedure("ClosedParameterType",none); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->number==3420);
    result=form.executeProcedure("EndOfFileType",none); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->number==3021);
    for(const auto* name:{"ClosedFieldsCount","ClosedFieldWrite"}) {
        result=form.executeProcedure(name,none); assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->number==3420);
    }
    assert(!form.hasUnsavedChanges());
    assert(sql.exec("SELECT label,notes FROM defaults WHERE id=1") && sql.next() && sql.value(0).toString()=="synthetic" && sql.value(1).isNull()); sql.finish();
    assert(form.close());
}
