#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_table_import.hpp"
#include "dao_database.hpp"
#include "vba_qt_value.hpp"
#include "access_table_model.hpp"
#include "access_autonumber.hpp"
#include "vba_currency.hpp"
#include "access_table_delegate.hpp"
#include "access_table_view.hpp"
#include <QHeaderView>
#include <QTableView>
#include <QDateTime>
#include <QApplication>
#include <QLineEdit>
#include <QSqlQuery>
#include <QTest>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>

int main(int argc,char** argv) {
    QApplication app(argc,argv);
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","currency-fields"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    const auto schema=linaccess::parseMdbSchema("CREATE TABLE `cash`\n(\n`id` INTEGER PRIMARY KEY,\n`amount` REAL\n);\n-- LINACCESS_SOURCE_TYPE 6964 INT\n-- LINACCESS_SOURCE_TYPE 616d6f756e74 CURRENCY\n");
    linaccess::importAccessTable(db,"synthetic-currency","cash",schema,linaccess::parseCsvDetailed("id,amount\n1,0.1\n2,1.2345\n3,\x1e\n4,549755813887.9999\n"));
    assert(sql.exec("CREATE TABLE currency_audit(value TEXT,source TEXT,type INTEGER,fieldtype INTEGER)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="cash"
Begin Section
Begin TextBox
Name ="amount"
ControlSource ="amount"
Width =4000
End
Begin TextBox
Name ="incremented"
ControlSource ="=[amount]+CCur(""0.0001"")"
Top =500
Width =4000
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub Capture()
Dim q As DAO.QueryDef
Dim r As DAO.Recordset
Set r=CurrentDb.OpenRecordset("cash")
Set q=CurrentDb.CreateQueryDef("","PARAMETERS v Text(255),s Text(255),t Integer,f Integer; INSERT INTO currency_audit VALUES(v,s,t,f)")
q.Parameters("v").Value=Me!amount
q.Parameters("s").Value=r.Fields("amount").Value
q.Parameters("t").Value=VarType(Me!amount)
q.Parameters("f").Value=r.Fields("amount").Type
q.Execute
r.Close
End Sub
Sub ChangeValue()
Me!amount=CCur("0.1")+CCur("0.2")
End Sub
Function ParameterUpdate(ByVal value As Currency)
Dim q As DAO.QueryDef
Set q=CurrentDb.CreateQueryDef("","PARAMETERS p Currency; UPDATE cash SET amount=p WHERE id=1")
q.Parameters("p").Value=value
If q.Parameters("p").Type<>dbCurrency Then Err.Raise 5
If VarType(q.Parameters("p"))<>vbCurrency Then Err.Raise 5
q.Execute
Me.Requery
ParameterUpdate=VarType(q.Parameters("p").Value)
End Function
Sub GoLast()
DoCmd.GoToRecord , , acLast
End Sub
Sub NewRow()
DoCmd.GoToRecord , , acNewRec
End Sub
Sub RejectLarge()
Me!amount=CCur("549755813888")
CurrentDb.Execute "INSERT INTO currency_audit VALUES('bad','bad',99,99)"
End Sub
)VBA";
    AccessFormPreview form("currency fields",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
    auto model=makePreviewFormObject(form); std::vector<VbaValue> none;
    assert(model->item("amount").kind==VbaValue::Kind::Currency && model->item("amount").currencyUnits==1000);
    assert(model->item("incremented").kind==VbaValue::Kind::Currency && model->item("incremented").currencyUnits==1001);
    const auto firstCapture=form.executeProcedure("Capture",none); if(firstCapture.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Synthetic capture error "<<firstCapture.errorNumber<<" at line "<<firstCapture.line<<'\n';
    assert(firstCapture.status==linaccess::VbaExecutionStatus::Completed);
    assert(form.executeProcedure("ChangeValue",none).status==linaccess::VbaExecutionStatus::Completed && model->item("amount").currencyUnits==3000 && form.hasUnsavedChanges());
    assert(form.executeProcedure("Capture",none).status==linaccess::VbaExecutionStatus::Completed);
    auto* editor=form.findChild<QLineEdit*>("access-control-amount"); assert(editor);
    model->member("Undo",{}); assert(model->item("amount").currencyUnits==1000 && !form.hasUnsavedChanges());
    editor->setFocus(); editor->selectAll(); QTest::keyClicks(editor,"1.23456");
    assert(editor->text()=="1.23456" && model->item("amount").currencyUnits==1000 && form.hasUnsavedChanges());
    model->assign("Dirty",VbaValue::logical(false)); assert(model->item("amount").currencyUnits==12346 && !form.hasUnsavedChanges());
    assert(form.executeProcedure("Capture",none).status==linaccess::VbaExecutionStatus::Completed);
    const auto rejected=form.executeProcedure("RejectLarge",none); assert(rejected.status!=linaccess::VbaExecutionStatus::Completed && rejected.errorNumber==3251);
    assert(model->item("amount").currencyUnits==12346 && !form.hasUnsavedChanges());
    assert(db.transaction());
    auto edit=linaccess::makeDaoDatabase(db)->member("OpenRecordset",{VbaValue::string("cash")}).object;
    auto editField=edit->member("Fields",{VbaValue::string("amount")}).object;
    edit->member("Edit",{}); editField->assign("Value",VbaValue::currency(12347));
    assert(editField->member("Value",{}).currencyUnits==12347);
    assert(sql.exec("SELECT amount FROM cash WHERE id=1") && sql.next() && sql.value(0).toDouble()==1.2346); sql.finish();
    edit->member("Update",{});
    assert(sql.exec("SELECT amount FROM cash WHERE id=1") && sql.next() && sql.value(0).toDouble()==1.2347); sql.finish();
    edit->member("Edit",{}); int assignmentError=0;
    try { editField->assign("Value",VbaValue::currency(5497558138880000LL)); } catch(const linaccess::VbaRuntimeError& failure) { assignmentError=failure.number; }
    assert(assignmentError==3251 && editField->member("Value",{}).currencyUnits==12347);
    edit->member("CancelUpdate",{}); edit->member("Close",{}); assert(db.rollback());
    model->assignItem("amount",VbaValue::currency(5497558138879999LL));
    assert(editor->text()=="549755813887.9999" && model->item("amount").currencyUnits==5497558138879999LL && form.hasUnsavedChanges());
    const auto old=model->objectItem("amount").object->member("OldValue",{}); assert(old.kind==VbaValue::Kind::Currency && old.currencyUnits==12346);
    model->member("Undo",{}); assert(model->item("amount").currencyUnits==12346 && !form.hasUnsavedChanges());
    auto rows=linaccess::makeDaoDatabase(db)->member("OpenRecordset",{VbaValue::string("cash")}).object;
    auto field=rows->member("Fields",{VbaValue::string("amount")}).object;
    assert(field->member("Type",{}).number==5 && field->member("Type",{}).numericType==VbaValue::NumericType::Integer);
    for(const auto expected:std::vector<std::optional<std::int64_t>>{12346,12345,{},5497558138879999LL}) {
        const auto value=field->member("Value",{});
        assert(expected ? value.kind==VbaValue::Kind::Currency && value.currencyUnits==*expected : value.kind==VbaValue::Kind::Null);
        rows->member("MoveNext",{});
    }
    assert(rows->member("EOF",{}).truthy() && field->member("Type",{}).number==5); rows->member("Close",{});
    assert(sql.exec("SELECT value,source,type,fieldtype FROM currency_audit ORDER BY rowid"));
    for(int row=0;row<3;++row) {
        assert(sql.next() && sql.value(0).toString()==(row==0 ? "0.1" : row==1 ? "0.3" : "1.2346"));
        assert(sql.value(1).toString()==(row<2 ? "0.1" : "1.2346") && sql.value(2).toInt()==6 && sql.value(3).toInt()==5);
    }
    assert(!sql.next()); sql.finish();
    assert(db.transaction()); assert(sql.exec("UPDATE cash SET amount=1.234567 WHERE id=1"));
    auto invalid=linaccess::makeDaoDatabase(db)->member("OpenRecordset",{VbaValue::string("cash")}).object;
    int error=0; try { (void)invalid->member("Fields",{VbaValue::string("amount")}).object->member("Value",{}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==13);
    invalid->member("Close",{}); assert(db.rollback());
    assert(sql.exec("SELECT amount FROM cash WHERE id=1") && sql.next() && sql.value(0).toDouble()==1.2346); sql.finish();
    std::vector<VbaValue> parameterArguments{VbaValue::currency(5000)};
    const auto parameterUpdate=form.executeProcedure("ParameterUpdate",parameterArguments);
    assert(parameterUpdate.status==linaccess::VbaExecutionStatus::Completed && parameterUpdate.typedReturnValue->number==6);
    assert(model->item("amount").currencyUnits==5000 && model->item("incremented").currencyUnits==5001 && !form.hasUnsavedChanges());
    assert(form.executeProcedure("GoLast",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(editor->text()=="549755813887.9999" && model->item("amount").currencyUnits==5497558138879999LL && !form.hasUnsavedChanges());
    assert(form.executeProcedure("NewRow",none).status==linaccess::VbaExecutionStatus::Completed && model->member("NewRecord",{}).truthy());
    editor->setFocus(); assert(QTest::qWaitFor([&] { return QApplication::focusWidget()==editor; },2000));
    QTest::keyClick(editor,Qt::Key_Apostrophe,Qt::ControlModifier);
    assert(editor->text()=="549755813887.9999" && model->item("amount").kind==VbaValue::Kind::Null && form.hasUnsavedChanges());
    QTest::keyClick(editor,Qt::Key_Return);
    assert(editor->text()=="549755813887.9999" && model->item("amount").currencyUnits==5497558138879999LL && form.hasUnsavedChanges());
    model->member("Undo",{}); assert(model->item("amount").kind==VbaValue::Kind::Null && !form.hasUnsavedChanges());
    assert(sql.exec("SELECT COUNT(*) FROM cash") && sql.next() && sql.value(0).toInt()==4); sql.finish();
    assert(form.close());
    assert(sql.exec("CREATE TABLE grid(id INTEGER PRIMARY KEY AUTOINCREMENT,amount REAL,memo TEXT,small INTEGER,stamp DATETIME,secondary INTEGER NOT NULL DEFAULT("+linaccess::accessCounterDefault("grid","secondary")+"),flag INTEGER)"));
    linaccess::createAccessCounter(db,"grid","secondary");
    sql.prepare("INSERT INTO _la_table_schema VALUES('grid',?)"); sql.addBindValue(QStringLiteral(R"JSON([{"name":"amount","source_type":"CURRENCY"},{"name":"small","source_type":"SMALLINT"},{"name":"stamp","source_type":"DATETIME"},{"name":"flag","source_type":"BOOLEAN"}])JSON")); assert(sql.exec());
    sql.prepare("INSERT INTO grid(amount,memo,small,stamp,flag) VALUES(?,?,?,?,?)"); sql.addBindValue(549755813887.9999); sql.addBindValue(QString(50000,'x')); sql.addBindValue(32767); sql.addBindValue("2024-02-29 12:34:56"); sql.addBindValue(0); assert(sql.exec()); sql.finish();
    linaccess::AccessTableModel grid(nullptr,db); grid.setTable("grid"); grid.setEditStrategy(QSqlTableModel::OnManualSubmit); assert(grid.select());
    assert(!(grid.flags(grid.index(0,0))&Qt::ItemIsEditable));
    assert(grid.data(grid.index(0,1),Qt::EditRole).toString()=="549755813887.9999");
    linaccess::AccessTableView view; view.setModel(&grid); linaccess::AccessTableDelegate delegate; QStyleOptionViewItem option;
    view.setItemDelegate(&delegate); view.setSortingEnabled(true); view.show(); QApplication::setActiveWindow(&view); view.setFocus(); QApplication::processEvents();
    view.edit(grid.index(0,1)); QApplication::processEvents(); auto* live=qobject_cast<QLineEdit*>(QApplication::focusWidget()); assert(live && view.isAncestorOf(live)); live->setText("1.23455");
    const auto priorSection=view.horizontalHeader()->sortIndicatorSection(); const auto priorOrder=view.horizontalHeader()->sortIndicatorOrder();
    view.horizontalHeader()->setSortIndicator(2,Qt::AscendingOrder);
    assert(QApplication::focusWidget()==live && live->text()=="1.23455" && !grid.isDirty() && view.horizontalHeader()->sortIndicatorSection()==priorSection && view.horizontalHeader()->sortIndicatorOrder()==priorOrder);
    QTest::keyClick(live,Qt::Key_Escape); QApplication::processEvents();
    auto* moneyEditor=delegate.createEditor(&view,option,grid.index(0,1)); auto* moneyText=qobject_cast<QLineEdit*>(moneyEditor); assert(moneyText);
    delegate.setEditorData(moneyEditor,grid.index(0,1)); assert(moneyText->text()=="549755813887.9999"); moneyText->setText("1.23455"); delegate.setModelData(moneyEditor,&grid,grid.index(0,1)); delete moneyEditor;
    assert(grid.data(grid.index(0,1)).toString()=="1.2346"); int gridRejected=0; QObject::connect(&grid,&linaccess::AccessTableModel::editRejected,[&](int code,const QString&) { gridRejected=code; });
    grid.sort(2,Qt::DescendingOrder); assert(gridRejected==2501 && grid.isDirty() && grid.data(grid.index(0,1)).toString()=="1.2346");
    assert(!grid.setData(grid.index(0,1),"549755813888") && gridRejected==3251 && grid.data(grid.index(0,1)).toString()=="1.2346");
    assert(!grid.setData(grid.index(0,3),32768) && gridRejected==6 && grid.data(grid.index(0,3)).toInt()==32767);
    view.setCurrentIndex(grid.index(0,1)); view.edit(grid.index(0,1)); QApplication::processEvents();
    auto* pending=qobject_cast<QLineEdit*>(QApplication::focusWidget()); assert(pending); pending->setText("549755813888");
    assert(!view.finishEditing() && QApplication::focusWidget()==pending && pending->text()=="549755813888" && grid.data(grid.index(0,1)).toString()=="1.2346");
    const auto guardedRows=grid.rowCount();
    assert(!view.insertRecord() && !view.deleteCurrentRecord() && grid.rowCount()==guardedRows && QApplication::focusWidget()==pending);
    pending->setText("1.23455"); assert(view.finishEditing() && grid.data(grid.index(0,1)).toString()=="1.2346");
    auto* memoEditor=delegate.createEditor(&view,option,grid.index(0,2)); auto* memoText=qobject_cast<QLineEdit*>(memoEditor); assert(memoText);
    delegate.setEditorData(memoEditor,grid.index(0,2)); assert(memoText->text().size()==50000); memoText->insert("z"); delegate.setModelData(memoEditor,&grid,grid.index(0,2)); delete memoEditor;
    assert(grid.setData(grid.index(0,4),QDateTime(QDate(2024,3,1),QTime(1,2,3))));
    assert(db.transaction() && grid.submitAll() && db.commit());
    assert(sql.exec("SELECT amount,length(memo),stamp FROM grid") && sql.next()); assert(linaccess::currencyFromRealStorage(sql.value(0).toDouble()).currencyUnits==12346 && sql.value(1).toInt()==50001 && sql.value(2).toDateTime()==QDateTime(QDate(2024,3,1),QTime(1,2,3))); sql.finish();
    assert(grid.insertRow(grid.rowCount())); const auto inserted=grid.rowCount()-1; assert(grid.data(grid.index(inserted,0)).toString()=="(AutoWert)" && grid.data(grid.index(inserted,5)).toString()=="(AutoWert)" && !(grid.flags(grid.index(inserted,5))&Qt::ItemIsEditable));
    assert(grid.setData(grid.index(inserted,1),"0.00015") && grid.setData(grid.index(inserted,2),"new")); assert(db.transaction() && grid.submitAll() && db.commit());
    assert(sql.exec("SELECT id,amount,secondary FROM grid ORDER BY id DESC") && sql.next() && sql.value(0).toInt()==2 && sql.value(2).toInt()==2 && linaccess::currencyFromRealStorage(sql.value(1).toDouble()).currencyUnits==2); sql.finish();

    const int boolRow=grid.data(grid.index(0,0)).toInt()==1 ? 0 : 1;
    assert(grid.flags(grid.index(boolRow,6))&Qt::ItemIsUserCheckable); assert(grid.data(grid.index(boolRow,6),Qt::CheckStateRole).toInt()==Qt::Unchecked && grid.data(grid.index(1-boolRow,6),Qt::CheckStateRole).toInt()==Qt::PartiallyChecked);
    view.setCurrentIndex(grid.index(boolRow,6)); view.setFocus(); QApplication::processEvents(); QTest::keyClick(&view,Qt::Key_Space);
    assert(grid.data(grid.index(boolRow,6),Qt::CheckStateRole).toInt()==Qt::Checked && grid.isDirty());
    assert(sql.exec("SELECT flag FROM grid WHERE id=1") && sql.next() && sql.value(0).toInt()==0); sql.finish();
    assert(db.transaction() && grid.submitAll() && db.commit()); assert(sql.exec("SELECT flag FROM grid WHERE id=1") && sql.next() && sql.value(0).toInt()==-1); sql.finish();

    AccessFormPreview dateForm("date-values","Begin Form\nRecordSource =\"grid\"\nBegin Section\nBegin TextBox\nName =\"stamp\"\nControlSource =\"stamp\"\nEnd\nEnd\nEnd\n",R"VBA(Function ShiftStamp() As Date
If VarType(Me!stamp)<>vbDate Then Err.Raise 5
Me!stamp=Me!stamp+1
Me.Dirty=False
ShiftStamp=Me!stamp
End Function
Function ParameterStamp() As Date
Dim definition As DAO.QueryDef
Dim rows As DAO.Recordset
Set definition=CurrentDb.CreateQueryDef("","PARAMETERS p DateTime; UPDATE grid SET stamp=[p] WHERE id=1")
definition.Parameters("p").Value=Me!stamp+1
If VarType(definition.Parameters("p").Value)<>vbDate Then Err.Raise 5
definition.Execute
Me.Requery
Set rows=CurrentDb.OpenRecordset("SELECT stamp FROM grid WHERE id=1")
If VarType(rows!stamp)<>vbDate Then Err.Raise 5
ParameterStamp=rows!stamp
rows.Close
End Function
)VBA",db);
    auto dateObject=makePreviewFormObject(dateForm); assert(dateObject->member("CurrentRecord",{}).asNumber()==1);
    std::vector<linaccess::VbaValue> dateArguments;
    auto shifted=dateForm.executeProcedure("ShiftStamp",dateArguments);
    assert(shifted.status==linaccess::VbaExecutionStatus::Completed && shifted.typedReturnValue->kind==linaccess::VbaValue::Kind::Date && shifted.typedReturnValue->asText()=="2024-03-02 01:02:03");
    shifted=dateForm.executeProcedure("ParameterStamp",dateArguments);
    assert(shifted.status==linaccess::VbaExecutionStatus::Completed && shifted.typedReturnValue->asText()=="2024-03-03 01:02:03");
    assert(sql.exec("SELECT stamp FROM grid WHERE id=1") && sql.next() && sql.value(0).toString()=="2024-03-03 01:02:03"); sql.finish();
    assert(dateForm.close());

}
