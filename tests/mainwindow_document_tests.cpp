#include "mainwindow.hpp"
#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_table_view.hpp"
#include <QAction>
#include <QApplication>
#include <QCryptographicHash>
#include <QClipboard>
#include <QDialog>
#include <QFileDialog>
#include <QFile>
#include <QFileInfo>
#include <QLineEdit>
#include <QLabel>
#include <QPlainTextEdit>
#include <QTableView>
#include <QTabWidget>
#include <QTabBar>
#include <QContextMenuEvent>
#include <QInputDialog>
#include <QMessageBox>
#include <QPointer>
#include <QMenu>
#include <QPushButton>
#include <QSqlQuery>
#include <QSplitter>
#include <QStandardPaths>
#include <QTemporaryDir>
#include <QTest>
#include <QTimer>
#include <QDateTime>
#include <QTreeWidget>
#include <QTreeWidgetItemIterator>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <cstdlib>
#include <iostream>

namespace {
QAction* action(MainWindow& window,const QString& text) {
    for(auto* candidate:window.findChildren<QAction*>()) if(candidate->text()==text) return candidate;
    assert(false); return nullptr;
}
void openProject(MainWindow& window,const QString& path,QWidget* shortcutTarget=nullptr) {
    bool selected=false;
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QFileDialog*>(QApplication::activeModalWidget()); assert(dialog);
        dialog->selectFile(path); selected=true;
        static_cast<QDialog*>(dialog)->accept();
    });
    if(shortcutTarget) { QTest::keyPress(shortcutTarget,Qt::Key_O,Qt::ControlModifier); QTest::keyRelease(&window,Qt::Key_O,Qt::ControlModifier); }
    else action(window,"Öffnen")->trigger();
    assert(selected);
}
void cancelProjectShortcut(QWidget* target,Qt::Key key) {
    bool rejected=false;
    QTimer::singleShot(0,target,[&] { auto* dialog=qobject_cast<QFileDialog*>(QApplication::activeModalWidget()); assert(dialog); rejected=true; dialog->reject(); });
    QTest::keyClick(target,key,Qt::ControlModifier); assert(rejected);
}
int artifact(QSqlDatabase db,const QString& kind,const QString& name,const QByteArray& bytes) {
    QSqlQuery sql(db);
    sql.prepare("INSERT INTO _la_access_artifacts(source_sha256,source_filename,kind,name,status,format,detail,source_bytes) VALUES(?,'synthetic',?,?,'source_only','synthetic','',?)");
    sql.addBindValue(QString::fromLatin1(QCryptographicHash::hash(bytes,QCryptographicHash::Sha256).toHex()));
    sql.addBindValue(kind); sql.addBindValue(name); sql.addBindValue(bytes); assert(sql.exec()); return sql.lastInsertId().toInt();
}
void openArtifact(MainWindow& window,int id) {
    auto* explorer=window.findChild<QTreeWidget*>("database-explorer"); assert(explorer);
    QTreeWidgetItem* selected=nullptr;
    for(QTreeWidgetItemIterator item(explorer);*item;++item)
        if((*item)->data(0,Qt::UserRole).toString()=="access_artifact" && (*item)->data(0,Qt::UserRole+1).toInt()==id) { selected=*item; break; }
    assert(selected); bool opened=false;
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());
        assert(dialog && dialog->objectName()=="access-artifact-details");
        for(auto* button:dialog->findChildren<QPushButton*>())
            if(button->text().startsWith("Access-Formular öffnen") || button->text().startsWith("Access-Bericht öffnen")) { opened=true; button->click(); return; }
        assert(false);
    });
    assert(QMetaObject::invokeMethod(explorer,"itemDoubleClicked",Qt::DirectConnection,Q_ARG(QTreeWidgetItem*,selected),Q_ARG(int,0)));
    assert(opened);
}
AccessFormPreview* document(const std::string& name) {
    for(auto* widget:QApplication::topLevelWidgets()) if(auto* form=dynamic_cast<AccessFormPreview*>(widget);form && form->isVisible())
        if(makePreviewFormObject(*form)->member("Name",{}).asText()==name) return form;
    return nullptr;
}
void exerciseTableSave(MainWindow& window,QSqlDatabase db) {
    auto* explorer=window.findChild<QTreeWidget*>("database-explorer"); assert(explorer); QTreeWidgetItem* selected=nullptr;
    for(QTreeWidgetItemIterator item(explorer);*item;++item) if((*item)->data(0,Qt::UserRole).toString()=="table" && (*item)->text(0)=="shortcut_records") { selected=*item; break; }
    assert(selected && QMetaObject::invokeMethod(explorer,"itemDoubleClicked",Qt::DirectConnection,Q_ARG(QTreeWidgetItem*,selected),Q_ARG(int,0)));
    linaccess::AccessTableView* table=nullptr;
    for(auto* candidate:window.findChildren<QTableView*>()) if(auto* typed=dynamic_cast<linaccess::AccessTableView*>(candidate)) table=typed;
    assert(table && table->model()->rowCount()>0); const auto id=table->model()->index(0,0).data().toInt();
    auto* documentTabs=window.findChild<QTabWidget*>(); assert(documentTabs && documentTabs->count()==1);
    const QPointer<QWidget> firstPage=documentTabs->currentWidget();
    auto* initialModel=qobject_cast<linaccess::AccessTableModel*>(table->model()); assert(initialModel && initialModel->setData(initialModel->index(0,1),"preserved draft"));
    assert(QMetaObject::invokeMethod(explorer,"itemDoubleClicked",Qt::DirectConnection,Q_ARG(QTreeWidgetItem*,selected),Q_ARG(int,0)));
    assert(documentTabs->count()==1 && documentTabs->currentWidget()==firstPage && table->model()==initialModel && initialModel->index(0,1).data().toString()=="preserved draft");
    QSqlQuery unchanged(db); assert(unchanged.exec("SELECT value FROM shortcut_records WHERE id=1") && unchanged.next() && unchanged.value(0).toString()=="before"); unchanged.finish(); initialModel->revertAll();
    QTreeWidgetItem* auditItem=nullptr;
    for(QTreeWidgetItemIterator item(explorer);*item;++item) if((*item)->data(0,Qt::UserRole).toString()=="table" && (*item)->text(0)=="sql_editor_audit") { auditItem=*item; break; }
    assert(auditItem && QMetaObject::invokeMethod(explorer,"itemDoubleClicked",Qt::DirectConnection,Q_ARG(QTreeWidgetItem*,auditItem),Q_ARG(int,0)));
    const QPointer<QWidget> secondPage=documentTabs->currentWidget(); assert(secondPage && secondPage!=firstPage && documentTabs->count()==2);
    linaccess::AccessTableView* auditView=nullptr;
    for(auto* candidate:secondPage->findChildren<QTableView*>()) if(auto* typed=dynamic_cast<linaccess::AccessTableView*>(candidate)) auditView=typed;
    assert(auditView); QApplication::setActiveWindow(&window); auditView->setFocus(); QApplication::processEvents();
    auto* documentsMenu=window.findChild<QMenu*>("window-menu"); assert(documentsMenu && QMetaObject::invokeMethod(documentsMenu,"aboutToShow",Qt::DirectConnection));
    QAction* firstEntry=nullptr; QAction* secondEntry=nullptr;
    for(auto* item:documentsMenu->actions()) {
        if(item->property("la_tab_title").toString()=="shortcut_records") firstEntry=item;
        if(item->property("la_tab_title").toString()=="sql_editor_audit") secondEntry=item;
    }
    assert(firstEntry && secondEntry && secondEntry->isChecked() && !firstEntry->isChecked());
    firstEntry->trigger(); QApplication::processEvents(); assert(documentTabs->currentWidget()==firstPage && table->hasFocus());
    QTest::keyClick(table,Qt::Key_F6,Qt::ControlModifier); QApplication::processEvents(); assert(documentTabs->currentWidget()==secondPage && auditView->hasFocus());
    QTest::keyClick(auditView,Qt::Key_F6,Qt::ControlModifier|Qt::ShiftModifier); QApplication::processEvents(); assert(documentTabs->currentWidget()==firstPage && table->hasFocus());
    QTest::keyClick(table,Qt::Key_F6,Qt::ControlModifier|Qt::ShiftModifier); QApplication::processEvents(); assert(explorer->hasFocus());
    QTest::keyClick(explorer,Qt::Key_F6,Qt::ControlModifier); QApplication::processEvents(); assert(documentTabs->currentWidget()==firstPage && table->hasFocus());
    secondEntry->trigger(); QApplication::processEvents(); assert(documentTabs->currentWidget()==secondPage && auditView->hasFocus());
    auditView->setCurrentIndex(auditView->model()->index(0,0)); auditView->edit(auditView->currentIndex()); QApplication::processEvents();
    auto* invalidDocumentEdit=qobject_cast<QLineEdit*>(QApplication::focusWidget()); assert(invalidDocumentEdit); invalidDocumentEdit->setText("invalid number");
    QTest::keyClick(invalidDocumentEdit,Qt::Key_F6,Qt::ControlModifier); QApplication::processEvents();
    assert(documentTabs->currentWidget()==secondPage && invalidDocumentEdit->hasFocus() && invalidDocumentEdit->text()=="invalid number");
    QTest::mouseClick(documentTabs->tabBar(),Qt::LeftButton,Qt::NoModifier,documentTabs->tabBar()->tabRect(0).center()); QApplication::processEvents();
    assert(documentTabs->currentWidget()==secondPage && invalidDocumentEdit->hasFocus());
    firstEntry->trigger(); QApplication::processEvents(); assert(documentTabs->currentWidget()==secondPage && invalidDocumentEdit->hasFocus());
    QApplication::clipboard()->setText("invalid-editor-sentinel");
    assert(window.dispatchTableCommand(190).value("error").toInt()==13);
    assert(window.dispatchTableCommand(50).value("error").toInt()==13);
    assert(QApplication::clipboard()->text()=="invalid-editor-sentinel" && invalidDocumentEdit->text()=="invalid number");
    const QPointer<QLineEdit> hiddenEditor=invalidDocumentEdit;
    documentTabs->setCurrentWidget(firstPage); QApplication::processEvents();
    assert(hiddenEditor && auditView->hasPendingEditor() && hiddenEditor->text()=="invalid number");
    const auto hiddenClose=window.closeTableDocument(0,"sql_editor_audit",2);
    assert(hiddenClose.value("error").toInt()==13 && documentTabs->count()==2 && auditView->hasPendingEditor());
    documentTabs->setCurrentWidget(secondPage); QApplication::setActiveWindow(&window); hiddenEditor->setFocus(); QApplication::processEvents();
    QTest::keyClick(invalidDocumentEdit,Qt::Key_Escape); QApplication::processEvents();
    const auto mixedId=artifact(db,"form","tab_cycle_form","Begin Form\nBegin Section\nBegin TextBox\nName =\"focus\"\nEnd\nEnd\nEnd\n");
    artifact(db,"vba","Form_tab_cycle_form",R"VBA(Sub SelectOpenTable()
DoCmd.SelectObject acTable,"SHORTCUT_RECORDS",False
Me.Tag="table-selected"
End Sub
Sub OpenReadonlyTable()
DoCmd.OpenTable TableName:="command_readonly",DataMode:=acReadOnly
Me.Tag="readonly-opened"
End Sub
Sub OpenEditableTable()
DoCmd.OpenTable "command_edit"
Me.Tag="edit-opened"
End Sub
Sub SaveActiveTable()
DoCmd.RunCommand acCmdSaveRecord
Me.Tag="table-saved"
End Sub
Sub SaveActiveTableViaApplication()
Application.RunCommand Command:=acCmdSaveRecord
Me.Tag="application-table-saved"
End Sub
Sub OpenNavigationTable()
DoCmd.OpenTable "command_navigation"
End Sub
Sub FirstNamedTableRecord()
DoCmd.GoToRecord acTable,"command_navigation",acFirst
End Sub
Sub NamedTableRecord300()
DoCmd.GoToRecord acTable,"command_navigation",acGoTo,300
End Sub
Sub LastTableRecord()
Application.RunCommand Command:=acCmdRecordsGoToLast
End Sub
Sub PreviousTableRecord()
DoCmd.RunCommand acCmdRecordsGoToPrevious
End Sub
Sub NextTableRecord()
DoCmd.GoToRecord
End Sub
Sub CloseNavigationTable()
DoCmd.Close acTable,"command_navigation"
End Sub
Sub ZoomTableCell()
DoCmd.RunCommand acCmdZoomBox
End Sub
Sub SelectTableRecord()
DoCmd.RunCommand acCmdSelectRecord
End Sub
Sub CopyTableSelection()
Application.RunCommand Command:=acCmdCopy
End Sub
Sub CloseEditableTable()
DoCmd.Close acTable,"command_edit",acSaveNo
Me.Tag="table-closed"
End Sub
Sub CloseActiveTable()
DoCmd.Close
Me.Tag="active-table-closed"
End Sub
Sub RejectAppendTable()
DoCmd.OpenTable "command_edit",,acAdd
Me.Tag="must-not-run"
End Sub
Sub RejectEmptyMode()
DoCmd.OpenTable TableName:="command_edit",DataMode:=Empty
Me.Tag="must-not-run"
End Sub
)VBA");
    assert(unchanged.exec("CREATE TABLE command_readonly(id INTEGER PRIMARY KEY AUTOINCREMENT,value REAL,flag BOOLEAN)"));
    assert(unchanged.exec("INSERT INTO command_readonly VALUES(1,1.25,-1),(2,NULL,NULL)"));
    assert(unchanged.exec("CREATE TABLE command_edit(id INTEGER PRIMARY KEY AUTOINCREMENT,value TEXT CHECK(value!='blocked'))"));
    assert(unchanged.exec("INSERT INTO command_edit VALUES(1,'initial')"));
    assert(unchanged.exec("CREATE TABLE command_navigation(id INTEGER PRIMARY KEY,value TEXT CHECK(value!='blocked'))"));
    assert(unchanged.exec("WITH RECURSIVE n(x) AS (VALUES(1) UNION ALL SELECT x+1 FROM n WHERE x<300) INSERT INTO command_navigation SELECT x,'row-'||x FROM n"));
    action(window,"Aktualisieren")->trigger(); openArtifact(window,mixedId); QPointer<AccessFormPreview> mixed=document("tab_cycle_form"); assert(mixed);
    auto* mixedEditor=mixed->findChild<QLineEdit*>("access-control-focus"); assert(mixedEditor); QApplication::setActiveWindow(mixed); mixedEditor->setFocus(); QApplication::processEvents();
    QTest::keyClick(mixedEditor,Qt::Key_F6,Qt::ControlModifier); QApplication::processEvents(); assert(explorer->hasFocus());
    QTest::keyClick(explorer,Qt::Key_F6,Qt::ControlModifier|Qt::ShiftModifier); QApplication::processEvents(); assert(activePreviewDocument()==mixed);
    std::vector<linaccess::VbaValue> selectArguments;
    assert(mixed->executeProcedure("SelectOpenTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    QApplication::processEvents(); assert(documentTabs->currentWidget()==firstPage && table->hasFocus() && !activePreviewDocument() && mixed->isVisible());
    assert(makePreviewFormObject(*mixed)->member("Tag",{}).asText()=="table-selected");
    const auto nativeTable=[&](QWidget* page) {
        linaccess::AccessTableView* result=nullptr;
        for(auto* candidate:page->findChildren<QTableView*>()) if(auto* typed=dynamic_cast<linaccess::AccessTableView*>(candidate)) result=typed;
        assert(result); return result;
    };
    assert(mixed->executeProcedure("OpenReadonlyTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed); QApplication::processEvents();
    const QPointer<QWidget> readonlyPage=documentTabs->currentWidget(); auto* readonlyView=nativeTable(readonlyPage);
    auto* readonlyModel=qobject_cast<linaccess::AccessTableModel*>(readonlyView->model()); assert(readonlyModel && readonlyModel->isReadOnly() && readonlyModel->rowCount()==2);
    const int readonlyRow=readonlyModel->index(0,0).data().toInt()==1 ? 0 : 1;
    assert(!readonlyModel->flags(readonlyModel->index(readonlyRow,1)).testFlag(Qt::ItemIsEditable) && !readonlyModel->flags(readonlyModel->index(readonlyRow,2)).testFlag(Qt::ItemIsUserCheckable));
    assert(!readonlyModel->setData(readonlyModel->index(readonlyRow,1),9) && !readonlyModel->insertRow(2) && !readonlyModel->removeRow(0) && !readonlyModel->isDirty());
    readonlyView->setCurrentIndex(readonlyModel->index(readonlyRow,2)); QTest::keyClick(readonlyView,Qt::Key_Space); QTest::keyClick(readonlyView,Qt::Key_Minus,Qt::ControlModifier);
    assert(readonlyModel->rowCount()==2 && !readonlyModel->isDirty() && readonlyModel->index(readonlyRow,2).data(Qt::CheckStateRole).toInt()==Qt::Checked);
    for(auto* button:readonlyPage->findChildren<QPushButton*>()) if(button->text()=="Speichern" || button->text()=="Datensatz löschen") assert(!button->isEnabled());
    assert(window.openTableDocument("command_readonly",0,1)==3251 && readonlyModel->isReadOnly());
    assert(mixed->executeProcedure("SaveActiveTable",selectArguments).errorNumber==2046 && !readonlyModel->isDirty());
    assert(makePreviewFormObject(*mixed)->member("Tag",{}).asText()=="readonly-opened");
    readonlyView->setCurrentIndex(readonlyModel->index(readonlyRow,1));
    assert(mixed->executeProcedure("SelectTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(readonlyView->selectionModel()->selectedIndexes().size()==3 && readonlyView->currentIndex()==readonlyModel->index(readonlyRow,1));
    QApplication::clipboard()->setText("unchanged-sentinel");
    assert(mixed->executeProcedure("CopyTableSelection",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(QApplication::clipboard()->text()=="1\t1.25\tTrue" && !readonlyModel->isDirty());
    QTimer::singleShot(0,[] {
        auto* zoom=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(zoom && zoom->objectName()=="native-cell-zoom");
        auto* editor=zoom->findChild<QPlainTextEdit*>("native-cell-zoom-text"); assert(editor && editor->isReadOnly() && editor->toPlainText()=="1.25"); zoom->reject();
    });
    assert(mixed->executeProcedure("ZoomTableCell",selectArguments).status==linaccess::VbaExecutionStatus::Completed && !readonlyModel->isDirty());
    QApplication::setActiveWindow(&window); readonlyView->setFocus(); QApplication::processEvents();
    assert(unchanged.exec("SELECT value,flag FROM command_readonly WHERE id=1") && unchanged.next() && unchanged.value(0).toDouble()==1.25 && unchanged.value(1).toInt()==-1); unchanged.finish();
    assert(mixed->executeProcedure("OpenNavigationTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed); QApplication::processEvents();
    auto* navigationView=nativeTable(documentTabs->currentWidget());
    auto* navigationModel=qobject_cast<linaccess::AccessTableModel*>(navigationView->model()); assert(navigationModel);
    navigationView->sortByColumn(0,Qt::AscendingOrder);
    navigationView->setCurrentIndex(navigationModel->index(0,1));
    assert(mixed->executeProcedure("LastTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(navigationView->currentIndex().row()==299 && navigationView->currentIndex().column()==1 && navigationModel->rowCount()==300);
    assert(mixed->executeProcedure("PreviousTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==298);
    assert(mixed->executeProcedure("NextTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==299);
    assert(mixed->executeProcedure("NextTableRecord",selectArguments).errorNumber==2105 && navigationView->currentIndex().row()==299);
    assert(window.activateOpenTable("command_readonly")==0);
    assert(mixed->executeProcedure("FirstNamedTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed && documentTabs->currentWidget()==navigationView->parentWidget());
    assert(navigationView->currentIndex().row()==0);
    assert(mixed->executeProcedure("NamedTableRecord300",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==299);
    navigationView->setCurrentIndex(navigationModel->index(0,1));
    assert(navigationModel->setData(navigationModel->index(0,1),"navigation-saved"));
    assert(navigationModel->setData(navigationModel->index(1,1),"other-draft"));
    assert(mixed->executeProcedure("LastTableRecord",selectArguments).errorNumber==3251 && navigationView->currentIndex().row()==0 && navigationModel->isDirty());
    assert(unchanged.exec("SELECT value FROM command_navigation WHERE id=1") && unchanged.next() && unchanged.value(0).toString()=="row-1"); unchanged.finish();
    assert(mixed->executeProcedure("SaveActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(mixed->executeProcedure("LastTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==299);
    assert(navigationModel->setData(navigationModel->index(299,1),"automatic-navigation-save"));
    assert(mixed->executeProcedure("PreviousTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(navigationView->currentIndex().row()==298 && !navigationModel->isDirty());
    assert(unchanged.exec("SELECT value FROM command_navigation WHERE id=300") && unchanged.next() && unchanged.value(0).toString()=="automatic-navigation-save"); unchanged.finish();

    assert(navigationModel->setData(navigationModel->index(298,1),"blocked"));
    assert(mixed->executeProcedure("NextTableRecord",selectArguments).errorNumber==3146 && navigationView->currentIndex().row()==298 && navigationModel->isDirty());
    assert(unchanged.exec("SELECT value FROM command_navigation WHERE id=299") && unchanged.next() && unchanged.value(0).toString()=="row-299"); unchanged.finish();
    assert(navigationModel->setData(navigationModel->index(298,1),"navigation-retry"));
    assert(mixed->executeProcedure("NextTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==299 && !navigationModel->isDirty());
    assert(mixed->executeProcedure("CloseNavigationTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed); QApplication::processEvents();
    assert(mixed->executeProcedure("OpenEditableTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed); QApplication::processEvents();
    const QPointer<QWidget> editablePage=documentTabs->currentWidget(); auto* editableView=nativeTable(editablePage);
    auto* editableModel=qobject_cast<linaccess::AccessTableModel*>(editableView->model()); assert(editableModel && !editableModel->isReadOnly() && editableModel->setData(editableModel->index(0,1),"changed"));
    assert(!editableModel->setReadOnly(true) && !editableModel->isReadOnly() && editableModel->isDirty());
    editableView->setCurrentIndex(editableModel->index(0,1));
    assert(mixed->executeProcedure("SelectTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(mixed->executeProcedure("CopyTableSelection",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(QApplication::clipboard()->text()=="1\tchanged" && editableModel->isDirty());
    QTimer::singleShot(0,[] {
        auto* zoom=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(zoom);
        auto* editor=zoom->findChild<QPlainTextEdit*>("native-cell-zoom-text"); assert(editor && !editor->isReadOnly() && editor->toPlainText()=="changed"); editor->setPlainText("zoom-draft"); zoom->accept();
    });
    assert(mixed->executeProcedure("ZoomTableCell",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(editableModel->index(0,1).data(Qt::EditRole).toString()=="zoom-draft" && editableModel->isDirty());
    QApplication::setActiveWindow(&window); editableView->setFocus(); QApplication::processEvents();
    QTimer::singleShot(0,[] {
        auto* zoom=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(zoom);
        zoom->findChild<QPlainTextEdit*>("native-cell-zoom-text")->setPlainText("discarded zoom"); zoom->reject();
    });
    assert(mixed->executeProcedure("ZoomTableCell",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(editableModel->index(0,1).data(Qt::EditRole).toString()=="zoom-draft");
    QApplication::setActiveWindow(&window); editableView->setFocus(); QApplication::processEvents();
    assert(editableModel->setData(editableModel->index(0,1),"changed"));
    assert(unchanged.exec("SELECT value FROM command_edit") && unchanged.next() && unchanged.value(0).toString()=="initial"); unchanged.finish();
    {
        AccessFormPreview initializing("table-background-init","Begin Form\nRecordSource =\"sql_editor_audit\"\nOnCurrent =\"[Event Procedure]\"\nBegin Section\nBegin TextBox\nName =\"value\"\nControlSource =\"value\"\nEnd\nEnd\nEnd\n",
            "Sub Form_Current()\nDoCmd.RunCommand acCmdSaveRecord\nMe.Tag=\"initialized\"\nEnd Sub\n",db,&window);
        assert(makePreviewFormObject(initializing)->member("Tag",{}).asText()=="initialized" && editableModel->isDirty());
        assert(unchanged.exec("SELECT value FROM command_edit") && unchanged.next() && unchanged.value(0).toString()=="initial"); unchanged.finish();
        assert(initializing.close());
    }
    QPushButton* commitButton=nullptr; for(auto* button:editablePage->findChildren<QPushButton*>()) if(button->text()=="Speichern") commitButton=button;
    assert(commitButton && commitButton->isEnabled());
    const auto saveResult=mixed->executeProcedure("SaveActiveTable",selectArguments);
    assert(saveResult.status==linaccess::VbaExecutionStatus::Completed && !editableModel->isDirty());
    assert(unchanged.exec("SELECT value FROM command_edit") && unchanged.next() && unchanged.value(0).toString()=="changed"); unchanged.finish();
    assert(makePreviewFormObject(*mixed)->member("Tag",{}).asText()=="table-saved");
    assert(editableModel->setData(editableModel->index(0,1),"application-changed"));
    assert(mixed->executeProcedure("SaveActiveTableViaApplication",selectArguments).status==linaccess::VbaExecutionStatus::Completed && !editableModel->isDirty());
    assert(unchanged.exec("SELECT value FROM command_edit") && unchanged.next() && unchanged.value(0).toString()=="application-changed"); unchanged.finish();
    assert(editableModel->setData(editableModel->index(0,1),"blocked"));
    const auto beforeCloseFailure=documentTabs->count();
    assert(mixed->executeProcedure("CloseEditableTable",selectArguments).errorNumber==3146 && documentTabs->count()==beforeCloseFailure && editableModel->isDirty());
    assert(unchanged.exec("SELECT value FROM command_edit") && unchanged.next() && unchanged.value(0).toString()=="application-changed"); unchanged.finish();
    assert(editableModel->setData(editableModel->index(0,1),"closed-by-native"));
    assert(mixed->executeProcedure("CloseEditableTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed); QApplication::processEvents();
    assert(documentTabs->count()==beforeCloseFailure-1);
    assert(unchanged.exec("SELECT value FROM command_edit") && unchanged.next() && unchanged.value(0).toString()=="closed-by-native"); unchanged.finish();
    const auto beforeUnsupported=documentTabs->count();
    const auto rejectedAppend=mixed->executeProcedure("RejectAppendTable",selectArguments);
    assert(rejectedAppend.errorNumber==3251 && documentTabs->count()==beforeUnsupported);
    assert(mixed->executeProcedure("RejectEmptyMode",selectArguments).errorNumber==3251 && documentTabs->count()==beforeUnsupported);
    assert(makePreviewFormObject(*mixed)->member("Tag",{}).asText()=="table-closed");
    assert(window.activateOpenTable("command_readonly")==0); QApplication::processEvents();
    QTimer::singleShot(0,&window,[&] {
        auto* popup=qobject_cast<QMenu*>(QApplication::activePopupWidget()); assert(popup && popup->objectName()=="native-datasheet-context-menu");
        assert(mixed->executeProcedure("CloseActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed && mixed->isVisible());
    });
    const auto popupPoint=readonlyView->visualRect(readonlyView->currentIndex()).center();
    QContextMenuEvent closeContext(QContextMenuEvent::Mouse,popupPoint,readonlyView->viewport()->mapToGlobal(popupPoint));
    QApplication::sendEvent(readonlyView->viewport(),&closeContext); QApplication::processEvents();
    assert(!readonlyPage && documentTabs->count()==2 && makePreviewFormObject(*mixed)->member("Tag",{}).asText()=="active-table-closed");
    assert(mixed->close()); QApplication::processEvents(); window.focusDatabaseWindow(); QApplication::processEvents();
    QTest::keyClick(explorer,Qt::Key_F6,Qt::ControlModifier|Qt::ShiftModifier); QApplication::processEvents(); assert(documentTabs->currentWidget()==secondPage && auditView->hasFocus());
    QTest::keyClick(auditView,Qt::Key_W,Qt::ControlModifier); QApplication::processEvents(); assert(documentTabs->count()==1 && documentTabs->currentWidget()==firstPage);
    table->setCurrentIndex(table->model()->index(0,1)); QApplication::setActiveWindow(&window); table->setFocus(); table->edit(table->currentIndex()); QApplication::processEvents();
    auto* editor=qobject_cast<QLineEdit*>(QApplication::focusWidget()); assert(editor && table->isAncestorOf(editor)); editor->setText("saved-through-shortcut");
    auto* editMenu=window.findChild<QMenu*>("native-edit-menu"); assert(editMenu);
    const auto editAction=[&](const char* name) { auto* result=editMenu->findChild<QAction*>(QString("edit-")+name); assert(result); return result; };
    const auto refreshEdit=[&] { assert(QMetaObject::invokeMethod(editMenu,"aboutToShow",Qt::DirectConnection)); };
    editor->selectAll(); refreshEdit(); editAction("copy")->trigger(); assert(QApplication::clipboard()->text()=="saved-through-shortcut");
    refreshEdit(); editAction("cut")->trigger(); assert(editor->text().isEmpty());
    refreshEdit(); editAction("paste")->trigger(); assert(editor->text()=="saved-through-shortcut");
    refreshEdit(); editAction("undo")->trigger(); assert(editor->text().isEmpty());
    refreshEdit(); editAction("redo")->trigger(); assert(editor->text()=="saved-through-shortcut");
    editor->selectAll(); editor->setReadOnly(true); refreshEdit(); assert(editAction("copy")->isEnabled() && !editAction("cut")->isEnabled() && !editAction("paste")->isEnabled()); editor->setReadOnly(false);
    editor->setEchoMode(QLineEdit::Password); refreshEdit(); assert(!editAction("copy")->isEnabled() && !editAction("cut")->isEnabled()); editor->setEchoMode(QLineEdit::Normal);
    QTest::keyPress(editor,Qt::Key_S,Qt::ControlModifier); QTest::keyRelease(table,Qt::Key_S,Qt::ControlModifier); QApplication::processEvents();
    QSqlQuery query(db); query.prepare("SELECT value FROM shortcut_records WHERE id=?"); query.addBindValue(id); assert(query.exec() && query.next() && query.value(0).toString()=="saved-through-shortcut");
    auto* model=qobject_cast<linaccess::AccessTableModel*>(table->model()); assert(!model->isDirty());
    assert(table->currentIndex().isValid() && model->index(table->currentIndex().row(),0).data().toInt()==id);
    assert(query.exec("INSERT INTO shortcut_records VALUES(999,'m-middle')")); query.finish(); assert(model->select());
    table->sortByColumn(1,Qt::AscendingOrder);
    int clockRow=-1; for(int row=0;row<model->rowCount();++row) if(model->index(row,0).data().toInt()==id) clockRow=row;
    assert(clockRow==1); table->setCurrentIndex(model->index(clockRow,1));
    QApplication::setActiveWindow(&window); table->setFocus();
    QTest::keyClick(table,Qt::Key_Space,Qt::ShiftModifier);
    assert(table->selectionModel()->selectedIndexes().size()==model->columnCount());
    QTest::keyClick(table,Qt::Key_C,Qt::ControlModifier); assert(QApplication::clipboard()->text()==QString::number(id)+"\tsaved-through-shortcut");
    QTest::keyClick(table,Qt::Key_Space,Qt::ControlModifier);
    assert(table->selectionModel()->selectedIndexes().size()==model->rowCount());
    refreshEdit(); assert(editAction("copy")->isEnabled()); editAction("copy")->trigger();
    assert(QApplication::clipboard()->text()=="m-middle\r\nsaved-through-shortcut" && !model->isDirty());
    table->setCurrentIndex(model->index(0,1)); QApplication::clipboard()->setText("first-range\r\nsecond-range");
    refreshEdit(); assert(editAction("paste")->isEnabled()); editAction("paste")->trigger();
    assert(model->isDirty() && model->index(0,1).data().toString()=="first-range" && model->index(1,1).data().toString()=="second-range");
    refreshEdit(); assert(!editAction("paste")->isEnabled());
    query.prepare("SELECT value FROM shortcut_records WHERE id=?"); query.addBindValue(id); assert(query.exec() && query.next() && query.value(0).toString()=="saved-through-shortcut"); query.finish();
    QTest::keyClick(table,Qt::Key_S,Qt::ControlModifier); QApplication::processEvents(); assert(!model->isDirty());
    assert(query.exec("SELECT COUNT(*) FROM shortcut_records WHERE value IN ('first-range','second-range')") && query.next() && query.value(0).toInt()==2); query.finish();
    assert(table->pasteSelection()); if(model->isDirty()) model->revertAll();
    QTest::keyClick(table,Qt::Key_A,Qt::ControlModifier); assert(table->selectionModel()->selectedIndexes().size()==model->rowCount()*model->columnCount());
    table->setCurrentIndex(model->index(1,1));
    const auto previous=model->index(0,1).data().toString();
    QTest::keyClick(table,Qt::Key_Apostrophe,Qt::ControlModifier);
    assert(model->index(1,1).data().toString()==previous && model->isDirty());
    assert(query.exec("SELECT COUNT(DISTINCT value) FROM shortcut_records") && query.next() && query.value(0).toInt()==2); query.finish();
    model->revertAll(); table->setCurrentIndex(model->index(clockRow,1));
    QTimer::singleShot(0,&window,[&] { auto* zoom=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(zoom); zoom->accept(); });
    const auto unchangedZoom=linaccess::nativeCellZoom(table,"line one\r\nline two",false);
    assert(unchangedZoom && !unchangedZoom->modified && unchangedZoom->text=="line one\r\nline two");
    QApplication::setActiveWindow(&window); table->setFocus();
    const auto beforeZoom=model->index(clockRow,1).data().toString();
    QTimer::singleShot(0,&window,[&] {
        auto* zoom=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(zoom && zoom->objectName()=="native-cell-zoom");
        auto* text=zoom->findChild<QPlainTextEdit*>("native-cell-zoom-text"); assert(text && text->toPlainText()==beforeZoom); text->setPlainText("cancelled zoom"); zoom->reject();
    });
    QTest::keyClick(table,Qt::Key_F2,Qt::ShiftModifier); assert(!model->isDirty() && model->index(clockRow,1).data().toString()==beforeZoom);
    QTimer::singleShot(0,&window,[&] {
        auto* zoom=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(zoom);
        zoom->findChild<QPlainTextEdit*>("native-cell-zoom-text")->setPlainText("zoom line one\nzoom line two"); zoom->accept();
    });
    QTest::keyClick(table,Qt::Key_F2,Qt::ShiftModifier); assert(model->isDirty() && model->index(clockRow,1).data().toString()=="zoom line one\nzoom line two");
    model->revertAll(); table->setCurrentIndex(model->index(clockRow,1)); QApplication::setActiveWindow(&window); table->setFocus();
    QTest::keyClick(table,Qt::Key_F2); QApplication::processEvents();
    auto* modeEditor=qobject_cast<QLineEdit*>(QApplication::focusWidget());
    assert(modeEditor && table->isAncestorOf(modeEditor));
    assert(!modeEditor->hasSelectedText() && modeEditor->cursorPosition()==modeEditor->text().size());
    QTimer::singleShot(0,&window,[&] {
        auto* zoom=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(zoom);
        zoom->findChild<QPlainTextEdit*>("native-cell-zoom-text")->setPlainText("editor zoom"); zoom->accept();
    });
    QTest::keyClick(modeEditor,Qt::Key_F2,Qt::ShiftModifier); assert(modeEditor->text()=="editor zoom" && !model->isDirty());
    QApplication::setActiveWindow(&window); modeEditor->setFocus();
    QTest::keyClick(modeEditor,Qt::Key_F2); QApplication::processEvents(); assert(QApplication::focusWidget()==table && model->isDirty());
    model->revertAll(); table->setCurrentIndex(model->index(clockRow,1));
    QTest::keyClick(table,Qt::Key_Left); assert(table->currentIndex().column()==0);
    table->setCurrentIndex(model->index(clockRow,1));
    const auto beforeDate=QDate::currentDate(); QTest::keyClick(table,Qt::Key_Semicolon,Qt::ControlModifier);
    auto* clockEditor=qobject_cast<QLineEdit*>(QApplication::focusWidget()); assert(clockEditor && table->isAncestorOf(clockEditor));
    const auto clockDate=clockEditor->text(); assert(clockDate==beforeDate.toString(Qt::ISODate) || clockDate==QDate::currentDate().toString(Qt::ISODate));
    QTest::keyPress(clockEditor,Qt::Key_S,Qt::ControlModifier); QTest::keyRelease(table,Qt::Key_S,Qt::ControlModifier); QApplication::processEvents();
    query.prepare("SELECT value FROM shortcut_records WHERE id=?"); query.addBindValue(id); assert(query.exec() && query.next() && query.value(0).toString()==clockDate); query.finish();
    assert(table->currentIndex().row()==0 && model->index(table->currentIndex().row(),0).data().toInt()==id);
    table->edit(table->currentIndex()); QApplication::processEvents(); clockEditor=qobject_cast<QLineEdit*>(QApplication::focusWidget()); assert(clockEditor);
    clockEditor->setText("prefix-suffix"); clockEditor->setSelection(7,6); QTest::keyClick(clockEditor,Qt::Key_Colon,Qt::ControlModifier|Qt::ShiftModifier);
    const auto clockTime=clockEditor->text(); assert(clockTime.startsWith("prefix-") && QTime::fromString(clockTime.mid(7),"HH:mm:ss").isValid());
    QTest::keyPress(clockEditor,Qt::Key_S,Qt::ControlModifier); QTest::keyRelease(table,Qt::Key_S,Qt::ControlModifier); QApplication::processEvents();
    query.prepare("SELECT value FROM shortcut_records WHERE id=?"); query.addBindValue(id); assert(query.exec() && query.next() && query.value(0).toString()==clockTime); query.finish();
    assert(table->currentIndex().row()==1 && model->index(table->currentIndex().row(),0).data().toInt()==id);
    assert(query.exec("DELETE FROM shortcut_records WHERE id=999")); query.finish(); assert(model->select());
    table->sortByColumn(0,Qt::AscendingOrder); table->setCurrentIndex(model->index(0,1));
    QApplication::setActiveWindow(&window); table->setFocus(); const auto originalRows=model->rowCount();
    QTest::keyClick(table,Qt::Key_Plus,Qt::ControlModifier|Qt::ShiftModifier); QApplication::processEvents();
    assert(model->rowCount()==originalRows+1 && table->currentIndex().row()==originalRows);
    assert(model->setData(model->index(originalRows,1),"keyboard-insert"));
    QTest::keyClick(table,Qt::Key_S,Qt::ControlModifier); QApplication::processEvents();
    assert(query.exec("SELECT COUNT(*) FROM shortcut_records WHERE value='keyboard-insert'") && query.next() && query.value(0).toInt()==1);
    int added=-1; for(int row=0;row<model->rowCount();++row) if(model->index(row,1).data().toString()=="keyboard-insert") added=row;
    assert(added>=0); table->setCurrentIndex(model->index(added,1));
    QTest::keyClick(table,Qt::Key_Minus,Qt::ControlModifier); QApplication::processEvents();
    QTest::keyClick(table,Qt::Key_S,Qt::ControlModifier); QApplication::processEvents();
    assert(query.exec("SELECT COUNT(*) FROM shortcut_records WHERE value='keyboard-insert'") && query.next() && query.value(0).toInt()==0);
    assert(model->rowCount()==originalRows && !model->isDirty());
    table->setCurrentIndex(model->index(0,1)); table->edit(table->currentIndex()); QApplication::processEvents();
    auto* live=qobject_cast<QLineEdit*>(QApplication::focusWidget()); assert(live && window.allowProjectClose(db.connectionName())); live->setText("unsaved-live");
    assert(!window.close() && live->isVisible() && live->text()=="unsaved-live"); QTest::keyClick(live,Qt::Key_Escape); QApplication::processEvents();
    assert(model->setData(model->index(0,1),"pending-close")); auto* tabs=window.findChild<QTabWidget*>(); assert(tabs); const auto count=tabs->count();
    QApplication::setActiveWindow(&window); table->setFocus(); QTest::keyClick(table,Qt::Key_W,Qt::ControlModifier); QApplication::processEvents();
    assert(tabs->count()==count && model->isDirty());
    assert(!window.close() && model->isDirty());
    const auto original=db.databaseName(),alternate=original+".other.linaccess"; assert(QFile::copy(original,alternate));
    openProject(window,alternate); assert(db.databaseName()==original && model->isDirty());
    const QByteArray layout="Begin Form\nRecordSource =\"shortcut_records\"\nEnd\n",code="Sub RequestQuit()\nDoCmd.Quit\nEnd Sub\n";
    auto* preview=AccessFormPreview::showDocument(linaccess::AccessDocumentKind::Form,"pending_data_quit",layout,code,db,&window); assert(preview);
    std::vector<linaccess::VbaValue> arguments; const auto quit=preview->executeProcedure("RequestQuit",arguments);
    assert(quit.status==linaccess::VbaExecutionStatus::Error && quit.errorNumber==2501 && preview->isVisible() && model->isDirty()); assert(preview->close()); QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete); QApplication::processEvents();
    const auto guardedTabs=tabs->count(); action(window,"Alle schließen")->trigger();
    assert(model->isDirty() && tabs->count()==guardedTabs);
    model->revertAll(); assert(!model->isDirty());
    QApplication::setActiveWindow(&window); table->setFocus(); QTest::keyPress(table,Qt::Key_F4,Qt::ControlModifier); QTest::keyRelease(&window,Qt::Key_F4,Qt::ControlModifier); QApplication::processEvents();
    assert(tabs->count()==count-1);
}
void exerciseQueryEditor(MainWindow& window,QSqlDatabase db) {
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE sql_editor_audit(value REAL)")); bool saved=false;
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->windowTitle()=="SQL-Abfrage");
        auto* source=dialog->findChild<QPlainTextEdit*>(); auto* table=dialog->findChild<QTableView*>(); assert(source && table);
        QPushButton* run=nullptr; QPushButton* save=nullptr;
        for(auto* button:dialog->findChildren<QPushButton*>()) { if(button->text()=="Ausführen") run=button; if(button->text()=="Speichern") save=button; } assert(run && save);
        source->setPlainText("SELECT CCur('1.23455') AS money,DateDiff('d',#2/28/2024#,#3/1/2024#) AS days"); run->click();
        assert(table->model()->rowCount()==1 && table->model()->index(0,0).data().toString()=="1.2346" && table->model()->index(0,1).data().toInt()==2);
        source->setPlainText("CREATE TABLE shortcut_records(id INTEGER PRIMARY KEY,value TEXT)"); run->click();
        bool zeroRows=false; for(auto* label:dialog->findChildren<QLabel*>()) zeroRows|=label->text()=="0 Datensätze geändert"; assert(zeroRows);
        source->setPlainText("INSERT INTO shortcut_records VALUES(1,'before')"); run->click();
        source->setPlainText("INSERT INTO sql_editor_audit VALUES(CCur('1.23455'))"); run->click(); assert(table->isHidden());
        assert(sql.exec("SELECT value FROM sql_editor_audit") && sql.next() && sql.value(0).toDouble()==1.2346); sql.finish();
        source->setPlainText("PARAMETERS p Currency; SELECT p AS amount");
        QTimer::singleShot(0,dialog,[] { auto* prompt=qobject_cast<QInputDialog*>(QApplication::activeModalWidget()); assert(prompt); prompt->setTextValue("0.1"); prompt->accept(); });
        run->click(); assert(!table->isHidden() && table->model()->index(0,0).data().toString()=="0.1");
        source->setPlainText("INSERT INTO sql_editor_audit VALUES(CCur('0.2'))"); auto* name=dialog->findChild<QLineEdit*>(); assert(name); name->setText("native_editor_action"); save->click(); saved=dialog->result()==QDialog::Accepted;
    });
    action(window,"+ Abfrage")->trigger(); assert(saved);
}
void dispose() { QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete); QApplication::processEvents(); }
}

int main(int argc,char** argv) {
    QCoreApplication::setAttribute(Qt::AA_DontUseNativeDialogs);
    QApplication app(argc,argv); app.setQuitOnLastWindowClosed(false); app.setApplicationName("LinAccessSyntheticGuiTest");
    QStandardPaths::setTestModeEnabled(true);
    QTimer watchdog; watchdog.setSingleShot(true);
    QObject::connect(&watchdog,&QTimer::timeout,[] { std::cerr<<"Synthetic main-window document test timed out\n"; std::abort(); }); watchdog.start(20000);
    QTemporaryDir temporary; assert(temporary.isValid());
    const auto first=temporary.filePath("first.linaccess"),second=temporary.filePath("second.linaccess");
    auto seed=QSqlDatabase::addDatabase("QSQLITE","mainwindow-seed"); seed.setDatabaseName(first); assert(seed.open());
    QSqlQuery initial(seed); assert(initial.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(initial.exec("INSERT INTO records VALUES(1,'old')")); assert(initial.exec("CREATE TABLE events(id INTEGER PRIMARY KEY,kind TEXT)")); initial.finish(); seed.close();
    seed.setDatabaseName(second); assert(seed.open()); seed.close();
    MainWindow window; window.show(); QApplication::setActiveWindow(&window); QApplication::processEvents();
    openProject(window,first); auto db=QSqlDatabase::database("linaccess-main"); assert(db.isOpen() && db.databaseName()==first);
    const QByteArray form="Begin Form\nRecordSource =\"records\"\nKeyPreview =NotDefault\nOnKeyDown =\"[Event Procedure]\"\nOnActivate =\"[Event Procedure]\"\nOnDeactivate =\"[Event Procedure]\"\nOnLoad =\"[Event Procedure]\"\nOnUnload =\"[Event Procedure]\"\nBegin Section\nBegin TextBox\nName =\"editor\"\nControlSource =\"value\"\nBeforeUpdate =\"[Event Procedure]\"\nEnd\nEnd\nEnd";
    const QByteArray code="Sub Form_Load()\nCurrentDb.Execute \"INSERT INTO events(kind) VALUES('load')\"\nEnd Sub\nSub Form_Unload(Cancel As Integer)\nIf Me.Tag=\"hold\" Then Cancel=True\nEnd Sub\nSub CloseSelf()\nDoCmd.Close\nMsgBox \"Synthetic retained document\"\nEnd Sub\n"
        "Sub Form_Activate()\nCurrentDb.Execute \"INSERT INTO events(kind) VALUES('activate')\"\nEnd Sub\nSub Form_Deactivate()\nCurrentDb.Execute \"INSERT INTO events(kind) VALUES('deactivate')\"\nEnd Sub\n"
        "Sub Form_KeyDown(KeyCode As Integer,Shift As Integer)\nIf KeyCode=vbKeyF11 And Me.Tag=\"block-f11\" Then KeyCode=0\nIf KeyCode=vbKeyF6 And Me.Tag=\"block-cycle\" Then KeyCode=0\nIf Me.Tag=\"block-project\" And (Shift And acCtrlMask)<>0 Then\nIf KeyCode=vbKeyO Or KeyCode=vbKeyN Then KeyCode=0\nEnd If\nEnd Sub\n"
        "Sub editor_BeforeUpdate(Cancel As Integer)\nIf Me.Tag=\"cancel-focus\" Then Cancel=True\nIf Me.Tag=\"close-cycle\" Then DoCmd.Close acReport,\"report\"\nEnd Sub\n"
        "Sub BackgroundSave()\nOn Error Resume Next\nDoCmd.RunCommand 97\nMe.Tag=CStr(Err.Number)\nEnd Sub\n"
        "Sub BackgroundKeys()\nOn Error Resume Next\nSendKeys \"x\",True\nMe.Tag=CStr(Err.Number)\nEnd Sub\n"
        "Sub BackgroundClose()\nOn Error Resume Next\nDoCmd.Close\nMe.Tag=CStr(Err.Number)\nEnd Sub\n"
        "Sub ExplicitSave()\nMe.Dirty=False\nEnd Sub\n"
        "Sub SelectTable()\nDoCmd.SelectObject acTable,\"records\",True\nEnd Sub\n"
        "Sub SelectTableGroup()\nDoCmd.SelectObject acTable,,True\nEnd Sub\n"
        "Sub SelectSourceForm()\nDoCmd.SelectObject acForm,\"ENTRY\",True\nEnd Sub\n"
        "Sub SelectQuery()\nDoCmd.SelectObject acQuery,\"source_query\",True\nEnd Sub\n"
        "Sub SelectEntry()\nDoCmd.SelectObject acForm,\"entry\",False\nEnd Sub\n"
        "Sub SelectReport()\nDoCmd.SelectObject acReport,\"report\",False\nEnd Sub\n"
        "Sub ProjectKeysWhileRunning()\nSendKeys \"^o\",True\nMe.Tag=\"returned\"\nEnd Sub\n";
    const int formId=artifact(db,"form","entry",form); artifact(db,"vba","Form_entry",code);
    const int queryId=artifact(db,"query","source_query","SELECT value FROM records");
    auto report=form; report.replace("Begin Form","Begin Report"); report.replace("OnLoad =\"[Event Procedure]\"\n",""); report.replace("OnUnload =\"[Event Procedure]\"\n","");
    const int reportId=artifact(db,"report","report",report);
    auto modal=form; modal.replace("Begin Form\n","Begin Form\nModal =NotDefault\n"); const int modalId=artifact(db,"form","modal",modal); artifact(db,"vba","Form_modal",code);
    auto cancel=form; cancel.replace("OnLoad =\"[Event Procedure]\"","OnOpen =\"[Event Procedure]\"");
    const int cancelId=artifact(db,"form","cancel",cancel); artifact(db,"vba","Form_cancel","Sub Form_Open(Cancel As Integer)\nCancel=True\nEnd Sub\n");
    const int initializationId=artifact(db,"form","initial_current","Begin Form\nRecordSource =\"records\"\nOnCurrent =\"[Event Procedure]\"\nBegin Section\nBegin TextBox\nName =\"editor\"\nControlSource =\"value\"\nEnd\nEnd\nEnd");
    artifact(db,"vba","Form_initial_current","Sub Form_Current()\nMe.Tag=\"before\"\nDoCmd.RunCommand 97\nMe.Tag=\"initialized\"\nEnd Sub\n");
    action(window,"Aktualisieren")->trigger();
    window.focusDatabaseWindow();
    openArtifact(window,initializationId); QPointer<AccessFormPreview> initializing=document("initial_current"); assert(initializing);
    assert(makePreviewFormObject(*initializing)->member("Tag",{}).asText()=="initialized"); assert(initializing->close()); dispose(); assert(!initializing);
    openArtifact(window,formId); assert(QTest::qWaitFor([&] { return document("entry")!=nullptr; },3000));
    QPointer<AccessFormPreview> entry=document("entry"); assert(entry->parentWidget()==&window && !entry->isModal() && QApplication::activeModalWidget()==nullptr);
    openArtifact(window,formId); assert(document("entry")==entry);
    QSqlQuery sql(db); assert(sql.exec("SELECT COUNT(*) FROM events WHERE kind='load'") && sql.next() && sql.value(0).toInt()==1); sql.finish();
    auto root=makePreviewFormObject(*entry); auto* editor=entry->findChild<QLineEdit*>("access-control-editor"); assert(editor);
    std::vector<linaccess::VbaValue> arguments;
    QApplication::setActiveWindow(entry); editor->setFocus(); root->assign("Tag",linaccess::VbaValue::string("block-project"));
    QTest::keyClick(editor,Qt::Key_O,Qt::ControlModifier); QTest::keyClick(editor,Qt::Key_N,Qt::ControlModifier);
    assert(editor->hasFocus() && !QApplication::activeModalWidget() && db.databaseName()==first);
    root->assign("Tag",linaccess::VbaValue::string(""));
    for(const auto key:{Qt::Key_O,Qt::Key_N}) {
        QApplication::setActiveWindow(entry); editor->setFocus(); cancelProjectShortcut(editor,key);
        assert(entry && db.databaseName()==first && db.isOpen());
    }
    QApplication::setActiveWindow(entry); editor->setFocus(); QApplication::processEvents();
    assert(entry->executeProcedure("ProjectKeysWhileRunning",arguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(entry && db.databaseName()==first && root->member("Tag",{}).asText()=="returned" && !QApplication::activeModalWidget());
    root->assign("Tag",linaccess::VbaValue::string(""));
    QApplication::setActiveWindow(entry); editor->setFocus(); editor->selectAll(); QTest::keyClicks(editor,"fresh");
    assert(entry->hasUnsavedChanges());
    auto* explorer=window.findChild<QTreeWidget*>("database-explorer"); assert(explorer);
    for(const auto* tag:{"block-cycle","cancel-focus"}) {
        root->assign("Tag",linaccess::VbaValue::string(tag));
        for(const auto modifiers:{Qt::KeyboardModifiers(Qt::ControlModifier),Qt::ControlModifier|Qt::ShiftModifier}) {
            QTest::keyClick(editor,Qt::Key_F6,modifiers); QApplication::processEvents();
            assert(activePreviewDocument()==entry && editor->hasFocus() && !AccessFormPreview::databaseWindowActive());
        }
    }
    root->assign("Tag",linaccess::VbaValue::string("block-f11")); QTest::keyClick(editor,Qt::Key_F11); QApplication::processEvents();
    assert(editor->hasFocus() && !AccessFormPreview::databaseWindowActive());
    root->assign("Tag",linaccess::VbaValue::string("cancel-focus")); QTest::keyClick(editor,Qt::Key_F11); QApplication::processEvents();
    assert(editor->hasFocus() && !AccessFormPreview::databaseWindowActive() && root->objectMember("editor",{}).object->member("Text",{}).asText()=="fresh");
    root->assign("Tag",linaccess::VbaValue::string(""));
    auto* splitter=qobject_cast<QSplitter*>(explorer->parentWidget()); assert(splitter); splitter->setSizes({0,window.width()});
    assert(sql.exec("SELECT COUNT(*) FROM events WHERE kind='deactivate'") && sql.next()); const int beforeDatabase=sql.value(0).toInt(); sql.finish();
    QTest::keyClick(editor,Qt::Key_F11); assert(QTest::qWaitFor([&] { return explorer->hasFocus(); },3000));
    assert(AccessFormPreview::databaseWindowActive(&window) && QApplication::activeWindow()==&window && splitter->sizes()[0]>0);
    assert(root->objectMember("editor",{}).object->member("Value",{}).asText()=="fresh" && entry->hasUnsavedChanges());
    assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="old"); sql.finish();
    assert(sql.exec("SELECT COUNT(*) FROM events WHERE kind='deactivate'") && sql.next() && sql.value(0).toInt()==beforeDatabase+1); sql.finish();
    QTest::keyClick(explorer,Qt::Key_F11); assert(explorer->hasFocus() && explorer->isVisible() && splitter->sizes()[0]>0);
    assert(sql.exec("SELECT COUNT(*) FROM events WHERE kind='deactivate'") && sql.next() && sql.value(0).toInt()==beforeDatabase+1); sql.finish();
    auto screen=makePreviewApplicationObject(*entry)->member("Screen",{}).object;
    for(const auto& property:std::vector<std::pair<const char*,int>>{{"ActiveForm",2450},{"ActiveControl",2474}}) {
        int error=0; try { screen->member(property.first,{}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==property.second);
    }
    for(const auto& procedure:std::vector<std::pair<const char*,const char*>>{{"BackgroundSave","2046"},{"BackgroundKeys","3251"},{"BackgroundClose","3251"}}) {
        assert(entry->executeProcedure(procedure.first,arguments).status==linaccess::VbaExecutionStatus::Completed);
        assert(root->member("Tag",{}).asText()==procedure.second && entry->isVisible());
    }
    assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="old"); sql.finish();
    assert(entry->executeProcedure("ExplicitSave",arguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="fresh"); sql.finish();
    QApplication::setActiveWindow(entry); editor->setFocus(); QApplication::processEvents();
    assert(!AccessFormPreview::databaseWindowActive() && screen->member("ActiveForm",{}).object->member("Name",{}).asText()=="entry");
    editor->selectAll(); QTest::keyClicks(editor,"second"); root->assign("Tag",linaccess::VbaValue::string("cancel-focus"));
    auto selection=entry->executeProcedure("SelectTable",arguments); assert(selection.status==linaccess::VbaExecutionStatus::Error && selection.errorNumber==2501 && editor->hasFocus());
    root->assign("Tag",linaccess::VbaValue::string("")); selection=entry->executeProcedure("SelectTable",arguments); assert(selection.status==linaccess::VbaExecutionStatus::Completed);
    assert(QTest::qWaitFor([&] { return explorer->hasFocus(); },3000));
    assert(explorer->currentItem()->data(0,Qt::UserRole+2).toString()=="table" && explorer->currentItem()->data(0,Qt::UserRole+3).toString()=="records");
    assert(entry->hasUnsavedChanges() && sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="fresh"); sql.finish();
    assert(entry->executeProcedure("SelectTableGroup",arguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(explorer->currentItem()->data(0,Qt::UserRole+4).toBool() && explorer->currentItem()->data(0,Qt::UserRole+2).toString()=="table");
    assert(entry->executeProcedure("SelectSourceForm",arguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(explorer->currentItem()->data(0,Qt::UserRole+1).toInt()==formId && explorer->currentItem()->parent()->data(0,Qt::UserRole+2).toString()=="form");
    entry->hide(); assert(entry->executeProcedure("SelectEntry",arguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(entry->isVisible() && activePreviewDocument()==entry && !AccessFormPreview::databaseWindowActive());
    assert(entry->executeProcedure("SelectQuery",arguments).status==linaccess::VbaExecutionStatus::Completed && explorer->currentItem()->data(0,Qt::UserRole+1).toInt()==queryId);
    selection=entry->executeProcedure("SelectReport",arguments); assert(selection.status==linaccess::VbaExecutionStatus::Error && selection.errorNumber==2451 && AccessFormPreview::databaseWindowActive());
    openArtifact(window,reportId); QPointer<AccessFormPreview> selectedReport=document("report"); assert(selectedReport);
    assert(entry->executeProcedure("SelectEntry",arguments).status==linaccess::VbaExecutionStatus::Completed && activePreviewDocument()==entry);
    assert(entry->executeProcedure("SelectReport",arguments).status==linaccess::VbaExecutionStatus::Completed && activePreviewDocument()==selectedReport);
    assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="fresh" && entry->hasUnsavedChanges()); sql.finish();
    auto* windowMenu=window.findChild<QMenu*>("window-menu"); assert(windowMenu);
    assert(QMetaObject::invokeMethod(windowMenu,"aboutToShow",Qt::DirectConnection));
    QAction* entryItem=nullptr; QAction* reportItem=nullptr;
    for(auto* item:windowMenu->actions()) if(item->isCheckable() && item->property("la_document_entry").toBool()) {
        if(item->text().contains(entry->windowTitle())) entryItem=item;
        if(item->text().contains(selectedReport->windowTitle())) reportItem=item;
    }
    assert(entryItem && reportItem && reportItem->isChecked() && !entryItem->isChecked());
    entryItem->trigger(); QApplication::processEvents(); assert(activePreviewDocument()==entry);
    action(window,"Datenbankfenster\tF11")->trigger(); QApplication::processEvents();
    assert(AccessFormPreview::databaseWindowActive(&window) && explorer->hasFocus());
    reportItem->trigger(); QApplication::processEvents(); assert(activePreviewDocument()==selectedReport);
    const auto cycle=[&](QWidget* source,bool backwards,AccessFormPreview* expected) {
        QTest::keyClick(source,Qt::Key_F6,backwards ? Qt::ControlModifier|Qt::ShiftModifier : Qt::KeyboardModifiers(Qt::ControlModifier));
        assert(QTest::qWaitFor([&] { return expected ? activePreviewDocument()==expected && QApplication::activeWindow()==expected : AccessFormPreview::databaseWindowActive(&window) && explorer->hasFocus(); },3000));
    };
    auto* reportEditor=selectedReport->findChild<QLineEdit*>("access-control-editor"); assert(reportEditor);
    QApplication::setActiveWindow(selectedReport); reportEditor->setFocus(); QApplication::processEvents();
    cycle(reportEditor,false,nullptr);
    action(window,"Nächstes Fenster\tStrg+F6")->trigger(); QApplication::processEvents();
    assert(activePreviewDocument()==entry && editor->hasFocus());

    cycle(editor,false,selectedReport); assert(reportEditor->hasFocus());
    action(window,"Vorheriges Fenster\tStrg+Umschalt+F6")->trigger(); QApplication::processEvents();
    assert(activePreviewDocument()==entry && editor->hasFocus());

    cycle(editor,true,nullptr);
    cycle(explorer,true,selectedReport);
    {
        assert(seed.open());
        AccessFormPreview unrelated("unrelated","Begin Form\nBegin Section\nBegin TextBox\nName =\"other\"\nEnd\nEnd\nEnd",{},seed,&window);
        unrelated.show(); window.focusDatabaseWindow(); cycle(explorer,true,selectedReport);
        assert(unrelated.close()); seed.close();
    }
    selectedReport->hide(); window.focusDatabaseWindow();
    cycle(explorer,true,entry); assert(!selectedReport->isVisible());
    selectedReport->showMinimized(); window.focusDatabaseWindow();
    cycle(explorer,true,selectedReport); assert(!selectedReport->isMinimized());
    assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="fresh" && entry->hasUnsavedChanges()); sql.finish();
    assert(entry->executeProcedure("SelectEntry",arguments).status==linaccess::VbaExecutionStatus::Completed);
    QApplication::setActiveWindow(entry); editor->setFocus(); editor->selectAll(); QTest::keyClicks(editor,"temporary");
    root->assign("Tag",linaccess::VbaValue::string("close-cycle")); cycle(editor,false,entry);
    dispose(); assert(!selectedReport);
    reportItem->trigger(); QApplication::processEvents(); assert(activePreviewDocument()==entry);
    root->assign("Tag",linaccess::VbaValue::string("")); editor->selectAll(); QTest::keyClicks(editor,"second");
    window.focusDatabaseWindow(); cycle(explorer,true,entry); cycle(editor,false,nullptr); cycle(explorer,false,entry);
    assert(entry->executeProcedure("SelectEntry",arguments).status==linaccess::VbaExecutionStatus::Completed);
    auto command=makePreviewApplicationObject(*entry)->member("DoCmd",{}).object;
    using linaccess::VbaValue;
    const auto badSelection=[&](const std::vector<VbaValue>& values,int expected) {
        int error=0; try { command->member("SelectObject",values); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==expected);
    };
    badSelection({},450); badSelection({VbaValue{}},94); badSelection({VbaValue::numeric(4)},3251);
    badSelection({VbaValue::numeric(0)},450);
    badSelection({VbaValue::numeric(0),VbaValue::string("records"),VbaValue::logical(false)},3251);
    badSelection({VbaValue::numeric(2),VbaValue::string("absent")},2450);
    badSelection({VbaValue::numeric(0),VbaValue::string("absent"),VbaValue::logical(true)},7874);
    badSelection({VbaValue::numeric(0),VbaValue::string("records"),VbaValue::string("bad")},13);
    badSelection({VbaValue::numeric(0),VbaValue::string("records"),VbaValue::logical(true),VbaValue::numeric(0)},450);
    class SelectionName final:public linaccess::VbaObject {
    public:
        int reads{};
        VbaValue scalarValue() override { ++reads; return VbaValue::string("records"); }
        VbaValue member(std::string_view,const std::vector<VbaValue>&) override { throw linaccess::VbaRuntimeError(438,"No members"); }
    };
    auto scalarName=std::make_shared<SelectionName>(); command->member("SelectObject",{VbaValue::numeric(0),VbaValue::reference(scalarName),VbaValue::logical(true)});
    assert(scalarName->reads==1 && explorer->currentItem()->data(0,Qt::UserRole+3).toString()=="records");
    assert(sql.exec("INSERT INTO _la_objects(kind,name,definition) VALUES('form','ENTRY','{}')")); action(window,"Aktualisieren")->trigger();
    badSelection({VbaValue::numeric(2),VbaValue::string("entry"),VbaValue::logical(true)},5);
    assert(sql.exec("DELETE FROM _la_objects WHERE kind='form' AND name='ENTRY'")); action(window,"Aktualisieren")->trigger();
    assert(entry->executeProcedure("SelectEntry",arguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(entry->executeProcedure("ExplicitSave",arguments).status==linaccess::VbaExecutionStatus::Completed);
    root->assign("Tag",linaccess::VbaValue::string("hold"));
    action(window,"Alle schließen")->trigger(); dispose(); assert(entry && entry->isVisible());

    QApplication::setActiveWindow(entry); editor->setFocus(); openProject(window,second,editor); assert(db.databaseName()==first && db.isOpen() && entry && entry->isVisible());
    root->assign("Tag",linaccess::VbaValue::string(""));
    QApplication::setActiveWindow(entry); editor->setFocus(); QApplication::processEvents();
    bool retained=false;
    QTimer::singleShot(0,&window,[&] {
        auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget()); assert(box);
        dispose(); retained=entry!=nullptr; assert(retained); box->button(QMessageBox::Ok)->click();
    });
    const auto closed=entry->executeProcedure("CloseSelf",arguments); assert(closed.status==linaccess::VbaExecutionStatus::Completed && retained);
    dispose(); assert(!entry);
    assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="second"); sql.finish();
    openArtifact(window,reportId); QPointer<AccessFormPreview> reportWindow=document("report"); assert(reportWindow && !reportWindow->isModal() && QApplication::activeModalWidget()==nullptr);
    assert(reportWindow->findChild<QLineEdit*>("access-control-editor")->isReadOnly()); action(window,"Alle schließen")->trigger(); dispose(); assert(!reportWindow && window.findChild<QTabWidget*>()->count()==0);
    openArtifact(window,modalId); QPointer<AccessFormPreview> modalWindow=document("modal"); assert(modalWindow && modalWindow->isModal() && QApplication::activeModalWidget()==modalWindow);
    auto* modalEditor=modalWindow->findChild<QLineEdit*>("access-control-editor"); assert(modalEditor); modalEditor->setFocus(); QTest::keyClick(modalEditor,Qt::Key_F11);
    assert(QApplication::activeModalWidget()==modalWindow && !AccessFormPreview::databaseWindowActive());
    QTest::keyClick(modalEditor,Qt::Key_O,Qt::ControlModifier); QTest::keyClick(modalEditor,Qt::Key_N,Qt::ControlModifier); assert(QApplication::activeModalWidget()==modalWindow);
    QTest::keyClick(modalEditor,Qt::Key_F6,Qt::ControlModifier); QTest::keyClick(modalEditor,Qt::Key_F6,Qt::ControlModifier|Qt::ShiftModifier);
    assert(QApplication::activeModalWidget()==modalWindow && !AccessFormPreview::databaseWindowActive());
    selection=modalWindow->executeProcedure("SelectTable",arguments); assert(selection.status==linaccess::VbaExecutionStatus::Error && selection.errorNumber==2046 && QApplication::activeModalWidget()==modalWindow);
    assert(modalWindow->close()); dispose(); assert(!modalWindow && QApplication::activeModalWidget()==nullptr);
    openArtifact(window,cancelId); assert(!document("cancel") && QApplication::activeModalWidget()==nullptr);
    openProject(window,second,explorer); assert(QSqlDatabase::database("linaccess-main").databaseName()==second);
    const auto third=temporary.filePath("third.linaccess"); bool created=false;
    QTimer::singleShot(0,&window,[&] { auto* dialog=qobject_cast<QFileDialog*>(QApplication::activeModalWidget()); assert(dialog); created=true; dialog->selectFile(third); static_cast<QDialog*>(dialog)->accept(); });
    QTest::keyClick(explorer,Qt::Key_N,Qt::ControlModifier);
    assert(created && QFileInfo::exists(third) && QSqlDatabase::database("linaccess-main").databaseName()==third);
    const auto hash=[&] { QFile input(first); assert(input.open(QIODevice::ReadOnly)); return QCryptographicHash::hash(input.readAll(),QCryptographicHash::Sha256); };
    const auto beforeRefusal=hash(); bool refused=false;
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QFileDialog*>(QApplication::activeModalWidget()); assert(dialog);
        dialog->setOption(QFileDialog::DontConfirmOverwrite,true); dialog->selectFile(first);
        QTimer::singleShot(0,&window,[&] { auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget()); assert(box); refused=true; box->accept(); });
        static_cast<QDialog*>(dialog)->accept();
    });
    QTest::keyClick(explorer,Qt::Key_N,Qt::ControlModifier);
    assert(refused && hash()==beforeRefusal && QSqlDatabase::database("linaccess-main").databaseName()==third);
    exerciseQueryEditor(window,db);
    exerciseTableSave(window,db);
    assert(window.close()); dispose(); watchdog.stop();
    std::cout<<"Main window: native project/explorer opening, independent nonmodal form/report lifetime, deduplication, modal metadata, cancellation, SQL save and retained close/MsgBox passed\n";
}
