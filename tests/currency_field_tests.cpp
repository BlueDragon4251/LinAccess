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
#include <QClipboard>
#include <QTableView>
#include <QDateTime>
#include <QApplication>
#include <QLineEdit>
#include <QSqlQuery>
#include <QSqlError>
#include <QTest>
#include <QTimer>
#include <QContextMenuEvent>
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
    view.setCurrentIndex(grid.index(0,1)); assert(view.copySelection() && QApplication::clipboard()->text()=="549755813887.9999");
    assert(grid.clipboardText(grid.index(0,4))=="2024-02-29 12:34:56" && grid.clipboardText(grid.index(0,6))=="False");

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
    QTest::keyClick(pending,Qt::Key_F2); QApplication::processEvents();
    assert(QApplication::focusWidget()==pending && pending->text()=="549755813888" && grid.data(grid.index(0,1)).toString()=="1.2346");
    for(const auto key:{Qt::Key_Tab,Qt::Key_Backtab,Qt::Key_Return,Qt::Key_Enter}) {
        QTest::keyClick(pending,key); QApplication::processEvents();
        assert(QApplication::focusWidget()==pending && pending->text()=="549755813888" && grid.data(grid.index(0,1)).toString()=="1.2346");
    }
    QTest::keyClick(pending,Qt::Key_Tab,Qt::ShiftModifier); QApplication::processEvents();
    assert(QApplication::focusWidget()==pending && pending->text()=="549755813888" && grid.data(grid.index(0,1)).toString()=="1.2346");
    QTest::mouseClick(view.viewport(),Qt::LeftButton,Qt::NoModifier,view.visualRect(grid.index(0,0)).center()); QApplication::processEvents();
    assert(QApplication::focusWidget()==pending && pending->text()=="549755813888" && view.currentIndex()==grid.index(0,1));
    QApplication::clipboard()->setText("unchanged-invalid-editor");
    assert(!view.copySelection() && QApplication::clipboard()->text()=="unchanged-invalid-editor" && QApplication::focusWidget()==pending);
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
    dateObject->assignItem("stamp",linaccess::VbaValue::numeric(2.5)); dateObject->assign("Dirty",linaccess::VbaValue::logical(false));
    assert(dateObject->item("stamp").kind==linaccess::VbaValue::Kind::Date && dateObject->item("stamp").asNumber()==2.5);
    assert(sql.exec("SELECT stamp FROM grid WHERE id=1") && sql.next() && sql.value(0).toString()=="1900-01-01 12:00:00"); sql.finish();
    int datePrecision=0; try { dateObject->assignItem("stamp",linaccess::VbaValue::dateSerial(2.123456789)); } catch(const linaccess::VbaRuntimeError& error) { datePrecision=error.number; }
    assert(datePrecision==3251 && dateObject->item("stamp").asNumber()==2.5 && !dateForm.hasUnsavedChanges());
    dateObject->assignItem("stamp",linaccess::VbaValue::string("2024-03-31 02:30:00")); dateObject->assign("Dirty",linaccess::VbaValue::logical(false));
    assert(dateObject->item("stamp").asText()=="2024-03-31 02:30:00");
    assert(sql.exec("SELECT stamp FROM grid WHERE id=1") && sql.next() && sql.value(0).toString()=="2024-03-31 02:30:00"); sql.finish();
    assert(dateForm.close());
    assert(grid.select() && grid.rowCount()==2); view.setCurrentIndex(grid.index(0,1)); QApplication::setActiveWindow(&view); view.setFocus(); QApplication::processEvents();
    const auto firstMoney=grid.clipboardText(grid.index(0,1)),secondMoney=grid.clipboardText(grid.index(1,1));
    QApplication::clipboard()->setText("1.25\r\nbad"); assert(!view.pasteSelection() && !grid.isDirty());
    assert(grid.clipboardText(grid.index(0,1))==firstMoney && grid.clipboardText(grid.index(1,1))==secondMoney);
    QApplication::clipboard()->setText("1.25\r\n2.5"); QTest::keyClick(&view,Qt::Key_V,Qt::ControlModifier);
    assert(grid.isDirty() && grid.clipboardText(grid.index(0,1))=="1.25" && grid.clipboardText(grid.index(1,1))=="2.5");
    QApplication::clipboard()->setText("9"); assert(view.pasteSelection() && grid.clipboardText(grid.index(0,1))=="9" && grid.clipboardText(grid.index(1,1))=="2.5");
    QApplication::clipboard()->setText("bad");assert(!view.pasteSelection() && grid.isDirty() && grid.clipboardText(grid.index(0,1))=="9" && grid.canUndoBlock());
    assert(grid.undoBlock() && grid.isDirty() && grid.clipboardText(grid.index(0,1))=="1.25" && grid.clipboardText(grid.index(1,1))=="2.5" && grid.canRedoBlock());
    assert(grid.redoBlock() && grid.clipboardText(grid.index(0,1))=="9" && grid.undoBlock());
    assert(db.transaction() && grid.submitAll() && db.commit());
    assert(sql.exec("SELECT SUM(amount) FROM grid") && sql.next() && sql.value(0).toDouble()==3.75); sql.finish();

    {
        assert(sql.exec("CREATE TABLE paste_append(id INTEGER PRIMARY KEY AUTOINCREMENT,value INTEGER,stamp TEXT DEFAULT '2003-12-31 18:05:06')"));
        assert(sql.exec("INSERT INTO paste_append(value) VALUES(10)"));
        linaccess::AccessTableModel cells(nullptr,db);cells.setTable("paste_append");cells.setEditStrategy(QSqlTableModel::OnManualSubmit);assert(cells.select());
        linaccess::AccessTableView sheet;sheet.setModel(&cells);sheet.setCurrentIndex(cells.index(0,1));
        QApplication::clipboard()->setText("11\r\n12\r\nbad");assert(!sheet.pasteSelection() && cells.rowCount()==1 && !cells.isDirty());
        QApplication::clipboard()->setText("11\r\n12\r\n13");assert(sheet.pasteSelection() && cells.rowCount()==3 && cells.isDirty());
        assert(!sheet.canPasteSelection() && !sheet.pasteSelection() && cells.rowCount()==3);
        assert(cells.clipboardText(cells.index(2,1))=="13");
        assert(sql.exec("SELECT count(*),sum(value) FROM paste_append") && sql.next() && sql.value(0).toInt()==1 && sql.value(1).toInt()==10);sql.finish();
        assert(cells.undoPendingRecord(2) && cells.rowCount()==2);
        cells.revertAll();assert(cells.rowCount()==1 && !cells.isDirty());
        sheet.setCurrentIndex(cells.index(0,1));assert(sheet.pasteSelection() && cells.submitTransactionally());
        assert(sql.exec("SELECT count(*),sum(value),count(DISTINCT id),min(stamp) FROM paste_append") && sql.next() && sql.value(0).toInt()==3 && sql.value(1).toInt()==36 && sql.value(2).toInt()==3 && sql.value(3).toString()=="2003-12-31 18:05:06");sql.finish();
        assert(sql.exec("CREATE UNIQUE INDEX paste_append_unique ON paste_append(value)"));
        sheet.setCurrentIndex(cells.index(2,1));QApplication::clipboard()->setText("20\r\n11");
        assert(sheet.pasteSelection() && cells.rowCount()==4 && !cells.submitTransactionally());
        assert(sql.exec("SELECT count(*),sum(value) FROM paste_append") && sql.next() && sql.value(0).toInt()==3 && sql.value(1).toInt()==36);sql.finish();
        cells.revertAll();assert(cells.select() && cells.rowCount()==3 && !cells.isDirty());
        sheet.setCurrentIndex(cells.index(2,0));QApplication::clipboard()->setText("21\r\n22");
        assert(!sheet.pasteSelection() && cells.rowCount()==3 && !cells.isDirty());
        assert(sql.exec("DELETE FROM paste_append"));assert(cells.select() && cells.rowCount()==0);sheet.setCurrentIndex({});
        QApplication::clipboard()->setText("31\r\n32");assert(sheet.canPasteSelection() && sheet.pasteSelection() && cells.rowCount()==2 && sheet.currentIndex().column()==1);
        assert(cells.undoBlock() && cells.rowCount()==0 && cells.canRedoBlock());assert(cells.redoBlock() && cells.rowCount()==2);

        assert(cells.submitTransactionally());sheet.setCurrentIndex(cells.index(0,1));
        QApplication::clipboard()->setText("33\r\n34");assert(sheet.pasteSelection(true) && cells.rowCount()==4 && sheet.currentIndex().row()==2 && cells.clipboardText(cells.index(0,1))=="31");
        assert(cells.submitTransactionally());assert(sql.exec("SELECT count(*),sum(value) FROM paste_append") && sql.next() && sql.value(0).toInt()==4 && sql.value(1).toInt()==130);sql.finish();


    }

    {
        assert(sql.exec("CREATE TABLE copy_cells(id INTEGER PRIMARY KEY,value TEXT)"));
        assert(sql.exec("WITH RECURSIVE n(i) AS (VALUES(1) UNION ALL SELECT i+1 FROM n WHERE i<300) INSERT INTO copy_cells SELECT i,'row-'||i FROM n"));
        linaccess::AccessTableModel cells(nullptr,db); cells.setTable("copy_cells"); cells.setEditStrategy(QSqlTableModel::OnManualSubmit); assert(cells.select());
        linaccess::AccessTableView sheet; sheet.setModel(&cells); sheet.show(); QApplication::setActiveWindow(&sheet); sheet.setFocus(); QApplication::processEvents();
        QTest::keyClick(&sheet,Qt::Key_End,Qt::ControlModifier);
        assert(cells.rowCount()==300 && sheet.currentIndex().row()==299 && sheet.currentIndex().column()==1);
        QTest::keyClick(&sheet,Qt::Key_Home,Qt::ControlModifier); assert(sheet.currentIndex()==cells.index(0,0));
        QTest::keyClick(&sheet,Qt::Key_A,Qt::ControlModifier);
        assert(cells.rowCount()==300 && sheet.selectionModel()->selectedIndexes().size()==600 && sheet.copySelection());
        const auto all=QApplication::clipboard()->text(); assert(all.split("\r\n").size()==300 && all.startsWith("1\trow-1\r\n") && all.endsWith("300\trow-300") && !cells.isDirty());
        assert(cells.setData(cells.index(0,1),"a\tb\"c\n")); sheet.setCurrentIndex(cells.index(0,1)); sheet.selectionModel()->select(cells.index(0,1),QItemSelectionModel::ClearAndSelect);
        QTimer::singleShot(0,&sheet,[&] {
            auto* menu=qobject_cast<QMenu*>(QApplication::activePopupWidget()); assert(menu && menu->objectName()=="native-datasheet-context-menu");
            auto* copy=menu->findChild<QAction*>("datasheet-copy"); auto* paste=menu->findChild<QAction*>("datasheet-paste");
            assert(copy && copy->isEnabled() && paste && paste->isEnabled()); menu->setActiveAction(copy); QTest::keyClick(menu,Qt::Key_Return);
        });
        const auto contextPoint=sheet.visualRect(cells.index(0,1)).center();
        QContextMenuEvent context(QContextMenuEvent::Mouse,contextPoint,sheet.viewport()->mapToGlobal(contextPoint)); QApplication::sendEvent(sheet.viewport(),&context);
        assert(QApplication::clipboard()->text()=="\"a\tb\"\"c\n\"");
        const auto quoted=QApplication::clipboard()->text(); cells.revertAll();
        sheet.setCurrentIndex(cells.index(0,1)); QTest::keyClick(&sheet,Qt::Key_V,Qt::ControlModifier);
        assert(cells.data(cells.index(0,1),Qt::EditRole).toString()=="a\tb\"c\n" && cells.isDirty());
        assert(sql.exec("SELECT value FROM copy_cells WHERE id=1") && sql.next() && sql.value(0).toString()=="row-1"); sql.finish();
        assert(db.transaction() && cells.submitAll() && db.commit());

        sheet.selectionModel()->select(cells.index(0,0),QItemSelectionModel::ClearAndSelect);
        sheet.selectionModel()->select(cells.index(1,1),QItemSelectionModel::Select);
        assert(!sheet.copySelection() && QApplication::clipboard()->text()==quoted);
        assert(sql.exec("ALTER TABLE copy_cells ADD COLUMN payload BLOB")); assert(sql.exec("UPDATE copy_cells SET payload=x'0102' WHERE id=1")); cells.setTable("copy_cells"); assert(cells.select());
        sheet.setCurrentIndex(cells.index(0,2)); sheet.selectionModel()->select(cells.index(0,2),QItemSelectionModel::ClearAndSelect); assert(sheet.copySelection() && QApplication::clipboard()->text()=="0x0102" && QApplication::clipboard()->mimeData()->hasFormat(linaccess::nativeGridMime));
        assert(sql.exec("SELECT value FROM copy_cells WHERE id=1") && sql.next() && sql.value(0).toString()=="a\tb\"c\n"); sql.finish();
    }
    {
        assert(sql.exec("CREATE TABLE typed_grid(id INTEGER PRIMARY KEY AUTOINCREMENT,n BIGINT,note TEXT,payload BLOB,amount CURRENCY_REAL,flag BOOLEAN,stamp DATETIME)"));
        assert(sql.exec("INSERT INTO typed_grid VALUES(1,9223372036854775806,NULL,x'0001ff80',1.25,-1,'2003-12-31 18:05:06'),(2,-9223372036854775807,'',x'',NULL,NULL,NULL)"));
        linaccess::AccessTableModel cells(nullptr,db);cells.setTable("typed_grid");cells.setEditStrategy(QSqlTableModel::OnManualSubmit);assert(cells.select());
        linaccess::AccessTableView sheet;sheet.setModel(&cells);sheet.setCurrentIndex(cells.index(0,1));
        sheet.selectionModel()->select(QItemSelection(cells.index(0,1),cells.index(1,6)),QItemSelectionModel::ClearAndSelect);
        assert(sheet.copySelection() && QApplication::clipboard()->mimeData()->hasFormat(linaccess::nativeGridMime));
        const auto encoded=QApplication::clipboard()->mimeData()->data(linaccess::nativeGridMime);
        sheet.setColumnHidden(2,true);assert(!sheet.pasteSelection() && !cells.isDirty());sheet.setColumnHidden(2,false);
        const bool pasted=sheet.pasteSelection(true);if(!pasted) std::cerr<<"typed paste: "<<cells.lastError().text().toStdString()<<" code="<<cells.lastError().nativeErrorCode().toStdString()<<"\n";assert(pasted && cells.rowCount()==4);
        assert(cells.canUndoBlock());QTest::keyClick(&sheet,Qt::Key_Z,Qt::ControlModifier);assert(cells.rowCount()==2 && !cells.isDirty() && cells.canRedoBlock());
        QTest::keyClick(&sheet,Qt::Key_Y,Qt::ControlModifier);assert(cells.rowCount()==4 && cells.isDirty() && cells.canUndoBlock());
        const bool saved=cells.submitTransactionally();if(!saved) std::cerr<<"typed save: "<<cells.lastError().text().toStdString()<<"\n";assert(saved);
        assert(sql.exec("SELECT n,note,payload,amount,flag,stamp FROM typed_grid ORDER BY id"));
        QVector<QVariantList> before;for(int row=0;row<2;++row) {assert(sql.next());QVariantList values;for(int column=0;column<6;++column) values<<sql.value(column);before<<values;}
        for(int row=0;row<2;++row) {assert(sql.next());for(int column=0;column<6;++column) assert(sql.value(column)==before[row][column] && sql.value(column).isNull()==before[row][column].isNull());}assert(!sql.next());sql.finish();
        assert(!before[1][1].isNull() && !before[1][2].isNull() && before[1][2].toByteArray().isEmpty());
        sheet.setCurrentIndex(cells.index(0,3));sheet.selectionModel()->select(cells.index(0,3),QItemSelectionModel::ClearAndSelect);assert(sheet.copySelection());
        sheet.setCurrentIndex(cells.index(1,3));assert(sheet.canPasteSelection() && sheet.pasteSelection() && cells.submitTransactionally());
        assert(sql.exec("SELECT hex(payload) FROM typed_grid WHERE id=2") && sql.next() && sql.value(0).toString()=="0001FF80");sql.finish();
        auto* invalid=new QMimeData;invalid->setData(linaccess::nativeGridMime,"{\"version\":1,\"rows\":[[{\"kind\":\"binary\",\"value\":\"!\"}]]}");QApplication::clipboard()->setMimeData(invalid);
        assert(!sheet.pasteSelection() && !cells.isDirty() && cells.rowCount()==4);
        auto* mismatch=new QMimeData;mismatch->setData(linaccess::nativeGridMime,encoded);QApplication::clipboard()->setMimeData(mismatch);
        sheet.setCurrentIndex(cells.index(0,0));assert(!sheet.pasteSelection() && !cells.isDirty());
        assert(cells.setTypedDataBlock(0,1,{{QVariant(2.5)}}) && cells.QSqlTableModel::data(cells.index(0,1),Qt::EditRole).toLongLong()==2 && cells.canUndoBlock());
        assert(cells.setData(cells.index(0,2),"later") && !cells.canUndoBlock() && !cells.canRedoBlock());cells.revertAll();
        sheet.setCurrentIndex(cells.index(0,5));sheet.selectionModel()->select(cells.index(0,5),QItemSelectionModel::ClearAndSelect);assert(sheet.copySelection() && QApplication::clipboard()->text()=="True");
        sheet.setCurrentIndex(cells.index(0,0));sheet.selectionModel()->select(QItemSelection(cells.index(0,0),cells.index(0,6)),QItemSelectionModel::ClearAndSelect);
        assert(sheet.copySelection() && sheet.canAppendRecords());const auto full=QApplication::clipboard()->mimeData()->data(linaccess::nativeGridMime);
        assert(sheet.pasteSelection(true) && cells.rowCount()==5 && cells.canUndoBlock());assert(cells.undoBlock() && cells.rowCount()==4 && cells.redoBlock() && cells.rowCount()==5 && cells.submitTransactionally());
        assert(sql.exec("SELECT id,n,hex(payload) FROM typed_grid ORDER BY id DESC LIMIT 1") && sql.next() && sql.value(0).toInt()==5 && sql.value(1).toLongLong()==9223372036854775806LL && sql.value(2).toString()=="0001FF80");sql.finish();
        auto forged=QJsonDocument::fromJson(full).object();auto fields=forged.value("fields").toArray();fields[0]="different_id";forged["fields"]=fields;
        auto* wrong=new QMimeData;wrong->setData(linaccess::nativeGridMime,QJsonDocument(forged).toJson());QApplication::clipboard()->setMimeData(wrong);assert(!sheet.pasteSelection(true) && !cells.isDirty() && cells.rowCount()==5);
        sheet.setCurrentIndex(cells.index(0,3));sheet.selectionModel()->select(cells.index(0,3),QItemSelectionModel::ClearAndSelect);assert(sheet.canCutSelection());QTest::keyClick(&sheet,Qt::Key_X,Qt::ControlModifier);
        assert(cells.isDirty() && cells.QSqlTableModel::data(cells.index(0,3),Qt::EditRole).isNull() && QApplication::clipboard()->text()=="0x0001ff80" && cells.canUndoBlock());
        assert(sql.exec("SELECT hex(payload) FROM typed_grid WHERE id=1") && sql.next() && sql.value(0).toString()=="0001FF80");sql.finish();
        assert(cells.undoBlock() && cells.redoBlock() && cells.undoBlock() && !cells.isDirty());
        assert(cells.setData(cells.index(0,2),"prior-draft"));sheet.setCurrentIndex(cells.index(3,3));
        assert(sheet.pasteSelection() && cells.isDirty() && cells.QSqlTableModel::data(cells.index(3,3),Qt::EditRole).toByteArray()==QByteArray::fromHex("0001ff80"));
        assert(cells.undoBlock() && cells.isDirty() && cells.clipboardText(cells.index(0,2))=="prior-draft" && cells.QSqlTableModel::data(cells.index(3,3),Qt::EditRole).toByteArray().isEmpty() && !cells.QSqlTableModel::data(cells.index(3,3),Qt::EditRole).isNull());
        assert(cells.redoBlock() && cells.clipboardText(cells.index(0,2))=="prior-draft");cells.revertAll();assert(cells.select());

        sheet.setCurrentIndex(cells.index(1,2));sheet.selectionModel()->select(cells.index(1,2),QItemSelectionModel::ClearAndSelect);assert(sheet.cutSelection() && cells.submitTransactionally());
        assert(sql.exec("SELECT note IS NULL FROM typed_grid WHERE id=2") && sql.next() && sql.value(0).toBool());sql.finish();
        sheet.setCurrentIndex(cells.index(0,0));sheet.selectionModel()->select(cells.index(0,0),QItemSelectionModel::ClearAndSelect);QApplication::clipboard()->setText("protected-cut");assert(!sheet.canCutSelection() && !sheet.cutSelection() && !cells.isDirty() && QApplication::clipboard()->text()=="protected-cut");



    }

    {
        assert(sql.exec("CREATE TABLE repeat_cells(id INTEGER PRIMARY KEY AUTOINCREMENT,amount CURRENCY_REAL,stamp DATETIME,flag BOOLEAN,note TEXT,payload BLOB)"));
        assert(sql.exec("INSERT INTO repeat_cells VALUES(1,549755813887.9999,'2024-02-29 12:34:56',1,NULL,x'0102'),(2,0,NULL,0,'old',NULL)"));
        linaccess::AccessTableModel repeated(nullptr,db); repeated.setTable("repeat_cells"); repeated.setEditStrategy(QSqlTableModel::OnManualSubmit); assert(repeated.select());
        for(int column=1;column<=4;++column) assert(repeated.repeatPreviousValue(repeated.index(1,column)));
        assert(repeated.clipboardText(repeated.index(1,1))=="549755813887.9999" && repeated.clipboardText(repeated.index(1,2))=="2024-02-29 12:34:56" && repeated.clipboardText(repeated.index(1,3))=="True");
        assert(repeated.data(repeated.index(1,4),Qt::EditRole).isNull());
        assert(!repeated.repeatPreviousValue(repeated.index(0,1)) && !repeated.repeatPreviousValue(repeated.index(1,0)) && !repeated.repeatPreviousValue(repeated.index(1,5)));
        assert(sql.exec("SELECT amount,note FROM repeat_cells WHERE id=2") && sql.next() && sql.value(0).toDouble()==0 && sql.value(1).toString()=="old"); sql.finish();
        linaccess::AccessTableView repeatedView; linaccess::AccessTableDelegate repeatedDelegate;
        repeatedView.setModel(&repeated); repeatedView.setItemDelegate(&repeatedDelegate); repeatedView.show(); QApplication::setActiveWindow(&repeatedView);
        repeatedView.setCurrentIndex(repeated.index(1,4)); repeatedView.edit(repeatedView.currentIndex()); QApplication::processEvents();
        auto* nullEditor=qobject_cast<QLineEdit*>(QApplication::focusWidget()); assert(nullEditor && repeatedView.isAncestorOf(nullEditor));
        assert(!repeatedView.hasPendingEditor() && repeatedView.finishEditing() && repeated.data(repeated.index(1,4),Qt::EditRole).isNull());
        assert(repeated.setData(repeated.index(1,4),"replace")); repeatedView.edit(repeatedView.currentIndex()); QApplication::processEvents();
        nullEditor=qobject_cast<QLineEdit*>(QApplication::focusWidget()); assert(nullEditor); QTest::keyClick(nullEditor,Qt::Key_Apostrophe,Qt::ControlModifier);
        assert(nullEditor->text().isEmpty() && repeated.data(repeated.index(1,4),Qt::EditRole).isNull());
        assert(repeatedView.finishEditing() && repeated.data(repeated.index(1,4),Qt::EditRole).isNull());
        assert(repeated.submitTransactionally());
        assert(sql.exec("SELECT amount,stamp,flag,note,payload FROM repeat_cells WHERE id=2") && sql.next());
        assert(linaccess::currencyFromRealStorage(sql.value(0).toDouble()).currencyUnits==5497558138879999LL && sql.value(1).toString()=="2024-02-29 12:34:56" && sql.value(2).toInt()==-1 && sql.value(3).isNull() && sql.value(4).isNull()); sql.finish();
        repeatedView.setCurrentIndex(repeated.index(1,4)); repeatedView.edit(repeatedView.currentIndex()); QApplication::processEvents();
        nullEditor=qobject_cast<QLineEdit*>(QApplication::focusWidget()); assert(nullEditor); QTest::keyClicks(nullEditor,"x"); QTest::keyClick(nullEditor,Qt::Key_Backspace);
        assert(repeatedView.hasPendingEditor() && repeatedView.finishEditing() && !repeated.data(repeated.index(1,4),Qt::EditRole).isNull());
        assert(repeated.submitTransactionally()); assert(sql.exec("SELECT note IS NULL,length(note) FROM repeat_cells WHERE id=2") && sql.next() && !sql.value(0).toBool() && sql.value(1).toInt()==0); sql.finish();
    }
    {
        assert(sql.exec("CREATE TABLE paste_constraints(id INTEGER PRIMARY KEY,amount INTEGER UNIQUE)"));
        assert(sql.exec("INSERT INTO paste_constraints VALUES(1,1),(2,2)"));
        linaccess::AccessTableModel constrained(nullptr,db); constrained.setTable("paste_constraints"); constrained.setEditStrategy(QSqlTableModel::OnManualSubmit); assert(constrained.select());
        assert(constrained.setDataBlock(constrained.index(0,1),{{"3"},{"3"}}));
        assert(!constrained.submitTransactionally());
        assert(sql.exec("SELECT SUM(amount) FROM paste_constraints") && sql.next() && sql.value(0).toInt()==3); sql.finish();
        assert(constrained.isDirty(constrained.index(0,1)) && constrained.isDirty(constrained.index(1,1)));
        assert(constrained.setData(constrained.index(1,1),4));
        assert(constrained.submitTransactionally());
        assert(sql.exec("SELECT SUM(amount) FROM paste_constraints") && sql.next() && sql.value(0).toInt()==7); sql.finish();
    }
    {
        assert(sql.exec("CREATE TABLE draft_mix(id INTEGER PRIMARY KEY AUTOINCREMENT,label TEXT UNIQUE,note TEXT DEFAULT 'seed')"));
        assert(sql.exec("INSERT INTO draft_mix(label) VALUES('old-a'),('old-b')"));
        linaccess::AccessTableModel drafts(nullptr,db); drafts.setTable("draft_mix"); drafts.setEditStrategy(QSqlTableModel::OnManualSubmit); assert(drafts.select());
        assert(drafts.setData(drafts.index(0,1),"updated") && drafts.insertRow(1));
        assert(drafts.setData(drafts.index(1,1),"old-b") && drafts.removeRow(2));
        assert(!drafts.submitTransactionally() && drafts.isDirty(drafts.index(0,1)) && drafts.isDirty(drafts.index(1,1)) && drafts.isDirty(drafts.index(2,1)));
        assert(sql.exec("SELECT COUNT(*) FROM draft_mix WHERE label IN ('old-a','old-b')") && sql.next() && sql.value(0).toInt()==2); sql.finish();
        assert(drafts.setData(drafts.index(1,1),"new") && drafts.submitTransactionally());
        assert(sql.exec("SELECT COUNT(*) FROM draft_mix WHERE label IN ('updated','new') AND note='seed'") && sql.next() && sql.value(0).toInt()==2); sql.finish();
        assert(drafts.insertRow(0) && drafts.setData(drafts.index(0,1),"temporary") && drafts.setData(drafts.index(1,1),"again") && drafts.removeRow(2));
        drafts.revertRow(0); assert(drafts.submitTransactionally());
        assert(sql.exec("SELECT COUNT(*),MIN(label) FROM draft_mix") && sql.next() && sql.value(0).toInt()==1 && sql.value(1).toString()=="again"); sql.finish();
    }

    {
        assert(sql.exec("CREATE TABLE refresh_commit(id INTEGER PRIMARY KEY AUTOINCREMENT,label TEXT)"));
        assert(sql.exec("INSERT INTO _la_table_schema(table_name,definition) VALUES('refresh_commit','[]')"));
        linaccess::AccessTableModel committed(nullptr,db); committed.setTable("refresh_commit"); committed.setEditStrategy(QSqlTableModel::OnManualSubmit); assert(committed.select());
        assert(committed.insertRow(0) && committed.setData(committed.index(0,1),"committed-once"));
        QObject::connect(&committed,&QSqlTableModel::beforeInsert,[&](QSqlRecord&) {
            QSqlQuery metadata(db); metadata.prepare("UPDATE _la_table_schema SET definition=? WHERE table_name='refresh_commit'");
            metadata.addBindValue("[{\"name\":\"label\",\"source_type\":\"CURRENCY\"}]"); assert(metadata.exec());
        });
        assert(!committed.submitTransactionally()); // SQL commits; the metadata refresh is refused.
        assert(committed.lastError().databaseText().startsWith("Änderungen gespeichert"));
        assert(sql.exec("SELECT COUNT(*) FROM refresh_commit") && sql.next() && sql.value(0).toInt()==1); sql.finish();
        assert(!committed.submitTransactionally());
        assert(sql.exec("SELECT COUNT(*) FROM refresh_commit") && sql.next() && sql.value(0).toInt()==1); sql.finish();
        assert(!committed.isDirty());
        assert(sql.exec("UPDATE _la_table_schema SET definition='[]' WHERE table_name='refresh_commit'"));
        assert(committed.submitTransactionally() && !committed.isDirty());
        assert(sql.exec("SELECT COUNT(*),MIN(label) FROM refresh_commit") && sql.next() && sql.value(0).toInt()==1 && sql.value(1).toString()=="committed-once"); sql.finish();
    }

    {
        assert(sql.exec("CREATE TABLE floating_cells(id INTEGER PRIMARY KEY,value REAL,money CURRENCY_REAL,stamp DATETIME)"));
        assert(sql.exec("INSERT INTO floating_cells VALUES(1,1.2345678901234567,NULL,NULL),(2,NULL,NULL,NULL)"));
        assert(sql.exec("SELECT value FROM floating_cells") && sql.next()); const auto original=sql.value(0).toDouble(); sql.finish();
        linaccess::AccessTableModel floating(nullptr,db); floating.setTable("floating_cells"); floating.setEditStrategy(QSqlTableModel::OnManualSubmit); assert(floating.select());
        linaccess::AccessTableView sheet; linaccess::AccessTableDelegate floatDelegate;
        sheet.setModel(&floating); sheet.setItemDelegate(&floatDelegate); sheet.show(); QApplication::setActiveWindow(&sheet);
        sheet.setCurrentIndex(floating.index(0,1)); sheet.edit(sheet.currentIndex()); QApplication::processEvents();
        assert(sheet.finishEditing() && floating.submitTransactionally());
        assert(sql.exec("SELECT value FROM floating_cells") && sql.next() && sql.value(0).toDouble()==original); sql.finish();
        sheet.setCurrentIndex(floating.index(0,1)); sheet.edit(sheet.currentIndex()); QApplication::processEvents();
        auto* editor=qobject_cast<QLineEdit*>(QApplication::focusWidget()); assert(editor && sheet.isAncestorOf(editor)); editor->setText("9.876543210987654");
        QTest::keyClick(editor,Qt::Key_Tab); QApplication::processEvents(); assert(floating.isDirty());
        assert(sql.exec("SELECT value FROM floating_cells") && sql.next() && sql.value(0).toDouble()==original); sql.finish();
        assert(floating.submitTransactionally());
        assert(sql.exec("SELECT value FROM floating_cells") && sql.next() && sql.value(0).toDouble()==9.876543210987654); sql.finish();
        for(int column=1;column<=3;++column) {
            sheet.setCurrentIndex(floating.index(1,column)); sheet.edit(sheet.currentIndex()); QApplication::processEvents();
            assert(sheet.finishEditing() && !floating.isDirty());
        }
        assert(floating.submitTransactionally());
        assert(sql.exec("SELECT value IS NULL,money IS NULL,stamp IS NULL FROM floating_cells WHERE id=2") && sql.next() && sql.value(0).toBool() && sql.value(1).toBool() && sql.value(2).toBool()); sql.finish();
    }

}
