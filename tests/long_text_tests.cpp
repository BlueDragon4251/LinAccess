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
    auto db=QSqlDatabase::addDatabase("QSQLITE","long-text-documents"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE documents(id INTEGER PRIMARY KEY,memo TEXT,label TEXT)"));
    const QString original=QString("head\r\n")+QString(40000,'x')+QString("ö😀\nend");
    sql.prepare("INSERT INTO documents VALUES(1,?,'one')"); sql.addBindValue(original); assert(sql.exec()); sql.finish();
    assert(sql.exec("INSERT INTO documents VALUES(2,'short','two')"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="documents"
Begin Section
Begin TextBox
Name ="memo"
ControlSource ="memo"
Width =5000
Height =500
End
Begin TextBox
Name ="label"
ControlSource ="label"
Top =600
Width =5000
Height =400
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub ChangeOther()
Me!label="changed"
Me.Dirty=False
End Sub
Sub AppendMemo()
Me!memo=Me!memo & " tail"
Me.Dirty=False
End Sub
Sub Second()
DoCmd.GoToRecord , , acGoTo, 2
End Sub
)VBA";
    AccessFormPreview form("long text",layout,code,db); form.show(); form.activateWindow();
    auto model=makePreviewFormObject(form); auto* editor=form.findChild<QLineEdit*>("access-control-memo"); assert(editor);

    std::vector<VbaValue> none;
    assert(form.executeProcedure("ChangeOther",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(sql.exec("SELECT memo,label FROM documents WHERE id=1") && sql.next());
    assert(sql.value(1).toString()=="changed");
    assert(sql.value(0).toString()==original); sql.finish();
    assert(editor->text()==original && QString::fromStdString(model->item("memo").asText())==original);
    assert(form.executeProcedure("AppendMemo",none).status==linaccess::VbaExecutionStatus::Completed);
    const QString appended=original+" tail";
    assert(sql.exec("SELECT memo FROM documents WHERE id=1") && sql.next() && sql.value(0).toString()==appended); sql.finish();
    assert(form.executeProcedure("Second",none).status==linaccess::VbaExecutionStatus::Completed);
    editor->setFocus(); assert(QTest::qWaitFor([&] { return QApplication::focusWidget()==editor; },2000));
    QTest::keyClick(editor,Qt::Key_Apostrophe,Qt::ControlModifier);
    assert(editor->text()==appended && model->item("memo").asText()=="short");
    assert(sql.exec("SELECT memo FROM documents WHERE id=2") && sql.next() && sql.value(0).toString()=="short"); sql.finish();
    model->assign("Dirty",VbaValue::logical(false));
    assert(sql.exec("SELECT memo FROM documents WHERE id=2") && sql.next() && sql.value(0).toString()==appended); sql.finish();
    assert(form.close());
}
