#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "dao_database.hpp"
#include "vba_qt_value.hpp"
#include <QApplication>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
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
    auto db=QSqlDatabase::addDatabase("QSQLITE","integer-field-documents"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE _la_table_schema(table_name TEXT PRIMARY KEY,definition TEXT)"));
    assert(sql.exec("CREATE TABLE numbers(id INTEGER PRIMARY KEY,small INTEGER,wide INTEGER,octet INTEGER,large INTEGER)"));
    assert(sql.exec("INSERT INTO numbers VALUES(1,-32768,2147483647,255,9007199254740991)"));
    const QJsonArray columns{QJsonObject{{"name","small"},{"source_type","SMALLINT"}},QJsonObject{{"name","wide"},{"source_type","INT"}},QJsonObject{{"name","octet"},{"source_type","TINYINT"}}};
    sql.prepare("INSERT INTO _la_table_schema VALUES('numbers',?)"); sql.addBindValue(QString::fromUtf8(QJsonDocument(columns).toJson(QJsonDocument::Compact))); assert(sql.exec()); sql.finish();
    assert(sql.exec("CREATE TABLE field_audit(st INTEGER,wt INTEGER,bt INTEGER,sv INTEGER,wv INTEGER,bv INTEGER)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="numbers"
Begin Section
Begin TextBox
Name ="small"
ControlSource ="small"
Width =4000
End
Begin TextBox
Name ="wide"
ControlSource ="wide"
Top =500
Width =4000
End
Begin TextBox
Name ="octet"
ControlSource ="octet"
Top =1000
Width =4000
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub CaptureTypes()
Dim r As DAO.Recordset
Set r=CurrentDb.OpenRecordset("numbers")
CurrentDb.Execute "INSERT INTO field_audit VALUES(" & CStr(r.Fields("small").Type) & "," & CStr(r.Fields("wide").Type) & "," & CStr(r.Fields("octet").Type) & "," & CStr(VarType(r.Fields("small"))) & "," & CStr(VarType(r.Fields("wide"))) & "," & CStr(VarType(r.Fields("octet"))) & ")"
r.Close
End Sub
Function RejectOverflow() As Long
Dim r As DAO.Recordset
Set r=CurrentDb.OpenRecordset("numbers")
r.Edit
On Error Resume Next
r.Fields("wide").Value=2147483648
RejectOverflow=Err.Number
r.CancelUpdate
r.Close
End Function
)VBA";
    AccessFormPreview form("integer fields",layout,code,db); form.show(); form.activateWindow(); QApplication::processEvents();
    auto model=makePreviewFormObject(form);
    assert(model->item("small").numericType==VbaValue::NumericType::Integer && model->item("small").number==-32768);
    assert(model->item("wide").numericType==VbaValue::NumericType::Long && model->item("wide").number==2147483647);
    assert(model->item("octet").numericType==VbaValue::NumericType::Byte && model->item("octet").number==255);
    assert(model->item("large").numericType==VbaValue::NumericType::Double && model->item("large").number==9007199254740991.0);
    int error=0; try { model->assignItem("wide",VbaValue::numeric(2147483648.0)); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
    assert(error==6 && model->item("wide").number==2147483647 && !form.hasUnsavedChanges());
    model->assignItem("small",VbaValue::numeric(2.5)); model->assignItem("wide",VbaValue::numeric(-3.5)); model->assignItem("octet",VbaValue::numeric(254.5));
    assert(model->item("small").number==2 && model->item("wide").number==-4 && model->item("octet").number==254 && form.hasUnsavedChanges());
    assert(sql.exec("SELECT small,wide,octet FROM numbers WHERE id=1") && sql.next() && sql.value(0).toInt()==-32768 && sql.value(1).toLongLong()==2147483647 && sql.value(2).toInt()==255); sql.finish();
    model->assign("Dirty",VbaValue::logical(false));
    auto* editor=form.findChild<QLineEdit*>("access-control-wide"); assert(editor); editor->setFocus(); editor->selectAll(); QTest::keyClicks(editor,"42");
    assert(editor->text()=="42" && model->item("wide").number==-4 && model->item("wide").numericType==VbaValue::NumericType::Long);
    model->assign("Dirty",VbaValue::logical(false)); assert(model->item("wide").number==42);
    std::vector<VbaValue> none; assert(form.executeProcedure("CaptureTypes",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(sql.exec("SELECT st,wt,bt,sv,wv,bv FROM field_audit") && sql.next());
    const int expectedTypes[]{3,4,2,2,3,17};
    for(int i=0;i<6;++i) assert(sql.value(i).toInt()==expectedTypes[i]); sql.finish();
    const auto rejected=form.executeProcedure("RejectOverflow",none); assert(rejected.status==linaccess::VbaExecutionStatus::Completed && rejected.typedReturnValue->number==6);
    const auto rows=linaccess::makeDaoDatabase(db)->member("OpenRecordset",{VbaValue::string("numbers")}).object;
    const auto field=rows->member("Fields",{VbaValue::string("wide")}).object;
    error=0; try { field->assign("Type",VbaValue::numeric(3)); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==383 && field->member("Type",{}).number==4);
    error=0; try { field->member("Type",{VbaValue::numeric(1)}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==450);
    const auto large=rows->member("Fields",{VbaValue::string("large")}).object;
    error=0; try { large->member("Type",{}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==3251);
    rows->member("MoveNext",{}); assert(field->member("Type",{}).number==4); rows->member("Close",{});
    error=0; try { field->member("Type",{}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==3420);
    assert(sql.exec("SAVEPOINT invalid_source") && sql.exec("UPDATE numbers SET wide=2147483648 WHERE id=1"));
    const auto invalid=linaccess::makeDaoDatabase(db)->member("OpenRecordset",{VbaValue::string("numbers")}).object;
    error=0; try { invalid->member("Fields",{VbaValue::string("wide")}).object->member("Value",{}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==6);
    invalid->member("Close",{}); assert(sql.exec("ROLLBACK TO invalid_source") && sql.exec("RELEASE invalid_source"));
    assert(sql.exec("SELECT small,wide,octet,large FROM numbers WHERE id=1") && sql.next() && sql.value(0).toInt()==2 && sql.value(1).toInt()==42 && sql.value(2).toInt()==254 && sql.value(3).toLongLong()==9007199254740991LL); sql.finish();
    for(const auto& type:{"INTEGER1","INTEGER2","INTEGER4"}) {
        error=0; try { (void)linaccess::qtFieldValue(3.5,type); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==13);
    }
    assert(!form.hasUnsavedChanges() && form.close());
}
