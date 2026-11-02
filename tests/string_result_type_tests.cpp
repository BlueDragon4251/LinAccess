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
    auto db=QSqlDatabase::addDatabase("QSQLITE","string-result-types"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE words(id INTEGER PRIMARY KEY,label TEXT)")); assert(sql.exec("INSERT INTO words VALUES(1,'A😀b')"));
    assert(sql.exec("CREATE TABLE string_audit(length INTEGER,position INTEGER,comparison INTEGER,lengthtype INTEGER,positiontype INTEGER,comparisontype INTEGER)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="words"
Begin Section
Begin TextBox
Name ="label"
ControlSource ="label"
Width =4000
End
Begin TextBox
Name ="length"
ControlSource ="=Len([label])"
End
Begin TextBox
Name ="position"
ControlSource ="=InStr([label],""b"")"
End
Begin TextBox
Name ="comparison"
ControlSource ="=StrComp([label],""A😀b"",0)"
End
Begin TextBox
Name ="lengthname"
ControlSource ="=TypeName(Len([label]))"
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub Capture()
Dim q As DAO.QueryDef
Set q=CurrentDb.CreateQueryDef("","PARAMETERS l Long,p Long,c Integer,lt Integer,pt Integer,ct Integer; INSERT INTO string_audit VALUES(l,p,c,lt,pt,ct)")
q.Parameters("l").Value=Len(Me!label)
q.Parameters("p").Value=InStr(Me!label,"b")
q.Parameters("c").Value=StrComp(Me!label,"A😀b",0)
q.Parameters("lt").Value=VarType(Len(Me!label))
q.Parameters("pt").Value=VarType(InStr(Me!label,"b"))
q.Parameters("ct").Value=VarType(StrComp(Me!label,"A😀b",0))
q.Execute
End Sub
)VBA";
    AccessFormPreview form("string result types",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
    auto model=makePreviewFormObject(form); std::vector<VbaValue> none;
    assert(model->item("length").numericType==VbaValue::NumericType::Long && model->item("length").number==4);
    assert(model->item("position").numericType==VbaValue::NumericType::Long && model->item("position").number==4);
    assert(model->item("comparison").numericType==VbaValue::NumericType::Integer && model->item("comparison").number==0);
    assert(model->item("lengthname").asText()=="Long");
    assert(form.executeProcedure("Capture",none).status==linaccess::VbaExecutionStatus::Completed);
    auto* editor=form.findChild<QLineEdit*>("access-control-label"); assert(editor);
    editor->setFocus(); editor->selectAll(); QTest::keyClicks(editor,"other");
    assert(editor->text()=="other" && model->item("label").asText()=="A😀b" && form.hasUnsavedChanges());
    assert(form.executeProcedure("Capture",none).status==linaccess::VbaExecutionStatus::Completed);
    model->assign("Dirty",VbaValue::logical(false));
    assert(model->item("length").number==5 && model->item("length").numericType==VbaValue::NumericType::Long && model->item("position").number==0);
    assert(form.executeProcedure("Capture",none).status==linaccess::VbaExecutionStatus::Completed);
    model->assignItem("label",VbaValue{}); assert(form.hasUnsavedChanges());
    assert(model->item("length").kind==VbaValue::Kind::Null && model->item("position").kind==VbaValue::Kind::Null && model->item("comparison").kind==VbaValue::Kind::Null);
    assert(form.executeProcedure("Capture",none).status==linaccess::VbaExecutionStatus::Completed);
    model->member("Undo",{}); assert(!form.hasUnsavedChanges() && editor->text()=="other");
    assert(sql.exec("SELECT length,position,comparison,lengthtype,positiontype,comparisontype FROM string_audit ORDER BY rowid"));
    for(int row=0;row<4;++row) {
        assert(sql.next());
        if(row==3) { for(int i=0;i<3;++i) assert(sql.value(i).isNull() && sql.value(i+3).toInt()==1); }
        else {
            assert(sql.value(0).toInt()==(row<2 ? 4 : 5) && sql.value(1).toInt()==(row<2 ? 4 : 0) && sql.value(2).toInt()==(row<2 ? 0 : 1));
            assert(sql.value(3).toInt()==3 && sql.value(4).toInt()==3 && sql.value(5).toInt()==2);
        }
    }
    assert(!sql.next()); sql.finish();
    assert(sql.exec("SELECT label FROM words") && sql.next() && sql.value(0).toString()=="other"); sql.finish();
    assert(form.close());
}
