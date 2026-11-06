#include "mainwindow.hpp"
#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_table_view.hpp"
#include "access_query.hpp"
#include "native_datasheet_print.hpp"
#include "access_table_import.hpp"
#include "access_catalog.hpp"
#include <QProcess>
#include <QPrintDialog>
#include <QPrinter>
#include <QRegularExpression>
#include <QAction>
#include <QApplication>
#include <QCryptographicHash>
#include <QClipboard>
#include <QCheckBox>
#include <QComboBox>
#include <QDialog>
#include <QDialogButtonBox>
#include <QDoubleSpinBox>
#include <QFileDialog>
#include <QFile>
#include <QFileInfo>
#include <QLineEdit>
#include <QListWidget>
#include <QSpinBox>
#include <QSettings>
#include "native_classic_style.hpp"
#include <QToolBar>
#include <QJsonArray>
#include <QJsonDocument>
#include <QLabel>
#include <QPlainTextEdit>
#include <QTableView>
#include <QTableWidget>
#include <QTabWidget>
#include <QTabBar>
#include <QContextMenuEvent>
#include <QInputDialog>
#include <QMessageBox>
#include <QPointer>
#include <QMenu>
#include <QPushButton>
#include <QSqlQuery>
#include <QSqlDriver>
#include <QLibrary>
#include <QSplitter>
#include <QStandardPaths>
#include <QTemporaryDir>
#include <QTest>
#include <QPrintPreviewWidget>
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

struct sqlite3;
struct sqlite3_context;
struct sqlite3_value;

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
void dispose();
void exerciseReferenceDatasheet() {
    const auto source=qEnvironmentVariable("LINACCESS_DATASHEET_REFERENCE");
    if(source.isEmpty()) return;
    const auto digest=[&] { QFile file(source); assert(file.open(QIODevice::ReadOnly)); return QCryptographicHash::hash(file.readAll(),QCryptographicHash::Sha256); };
    const auto before=digest();
    QTemporaryDir temporary; assert(temporary.isValid());
    {
        auto db=QSqlDatabase::addDatabase("QSQLITE","datasheet-reference"); db.setDatabaseName(temporary.filePath("copy.linaccess")); assert(db.open());
        QSqlQuery setup(db); assert(setup.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
        assert(linaccess::importAccessDatabase(db,source)==2);
        int tables=0,rows=0,texts=0;
        for(const auto& object:linaccess::AccessCatalog(source).objects()) {
            if(object.kind!="table" || object.internal) continue;
            QTableView view; linaccess::AccessTableModel model(nullptr,db);
            model.setTable(object.name); assert(model.select()); view.setModel(&model);
            const auto snapshot=linaccess::nativeDatasheetSnapshot(view,object.name);
            const auto quoted='"'+QString(object.name).replace('"',"\"\"")+'"';
            QSqlQuery actual(db); assert(actual.exec("SELECT * FROM "+quoted)); int row=0;
            while(actual.next()) {
                assert(row<snapshot.rows.size());
                for(int col=0;col<actual.record().count();++col) {
                    const auto value=actual.value(col);
                    if(value.isNull()) assert(snapshot.rows[row][col].isEmpty());
                    else if(value.metaType().id()==QMetaType::QString) { assert(snapshot.rows[row][col]==value.toString()); ++texts; }
                }
                ++row;
            }
            assert(row==snapshot.rows.size()); rows+=row;
            linaccess::exportNativeDatasheetPdf(snapshot,temporary.filePath(QString::number(tables++)+".pdf"));
        }
        assert(tables==2 && rows==39 && texts>0);
        db.close();
    }
    QSqlDatabase::removeDatabase("datasheet-reference"); assert(digest()==before);
    std::cout<<"Original recipe datasheet PDF: 2 tables, 39 rows, actual text/NULL values and unchanged original hash verified\n";
}

void exerciseDatasheetOutput(MainWindow& window,QSqlDatabase db) {
    QTemporaryDir output; assert(output.isValid());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE output_rows(id INTEGER PRIMARY KEY, value TEXT, hidden TEXT, bytes BLOB)"));
    assert(sql.exec("WITH RECURSIVE n(x) AS (VALUES(1) UNION ALL SELECT x+1 FROM n WHERE x<320) INSERT INTO output_rows SELECT x, 'row_'||printf('%03d',x), 'excluded', X'0102' FROM n"));
    assert(window.openTableDocument("output_rows",0,1)==0);
    auto* tabs=window.findChild<QTabWidget*>(); assert(tabs);
    const QPointer<QWidget> page=tabs->currentWidget();
    auto* view=page->findChild<QTableView*>(); assert(view);
    auto* model=qobject_cast<linaccess::AccessTableModel*>(view->model()); assert(model);
    view->hideColumn(2); view->hideColumn(3);
    model->setFilter("id>=10"); view->sortByColumn(0,Qt::DescendingOrder);
    view->horizontalHeader()->moveSection(view->horizontalHeader()->visualIndex(1),0);
    view->setFocus(); window.activateWindow(); QApplication::processEvents();
    const auto snapshot=linaccess::nativeDatasheetSnapshot(*view,"output_rows");
    assert(snapshot.rows.size()==311 && snapshot.headers==QStringList({"value","id"}));
    assert(snapshot.rows.front()==QStringList({"row_320","320"}) && snapshot.rows.back()==QStringList({"row_010","10"}));
    const auto path=output.filePath("data.pdf");
    assert(window.outputDatasheet(true,path).value("error").toInt()==0);
    const auto read=[&] { QFile file(path); assert(file.open(QIODevice::ReadOnly)); return file.readAll(); };
    const auto original=read(); assert(original.startsWith("%PDF-"));
    QRegularExpression pages("/Type\\s*/Page(?:\\s|/|>)");
    auto matches=pages.globalMatch(QString::fromLatin1(original)); int count=0;
    while(matches.hasNext()) { matches.next(); ++count; } assert(count>1);
    const auto reader=QStandardPaths::findExecutable("pdftotext");
    if(!reader.isEmpty()) {
        QProcess text; text.start(reader,{"-layout",path,"-"}); assert(text.waitForFinished(10000) && text.exitCode()==0);
        const auto content=QString::fromUtf8(text.readAllStandardOutput());
        assert(content.count("row_320")==1 && content.count("row_010")==1 && !content.contains("row_009") && !content.contains("excluded"));
        assert(content.indexOf("row_320")<content.indexOf("row_010") && content.count("value")==count);
        const auto lastPage=content.split(QChar(12),Qt::SkipEmptyParts).last(); assert(lastPage.contains("row_010"));
    }
    // A late rendering/range failure must preserve an existing destination too.
    bool rangeRejected=false;
    try { linaccess::exportNativeDatasheetPdf(snapshot,path,QPageLayout(QPageSize(QPageSize::A4),QPageLayout::Portrait,QMarginsF(15,15,15,15)),999999,999999); }
    catch(const linaccess::VbaRuntimeError& error) { rangeRejected=error.number==5; }
    assert(rangeRejected && read()==original);
    bool printPdf=false;
    const auto dialogPdf=output.filePath("dialog.pdf");
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QPrintDialog*>(QApplication::activeModalWidget()); assert(dialog);
        auto* printer=dialog->printer(); printer->setOutputFormat(QPrinter::PdfFormat); printer->setOutputFileName(dialogPdf);
        printer->setPrintRange(QPrinter::PageRange); printer->setFromTo(2,2); printPdf=true;
        dialog->QDialog::done(QDialog::Accepted);
    });
    const auto printResult=window.outputDatasheet(false);
    if(printResult.value("error").toInt()) std::cerr<<"Datasheet print: "<<printResult.value("error").toInt()<<" "<<printResult.value("description").toString().toStdString()<<"\n";
    assert(printResult.value("error").toInt()==0 && printPdf);
    QFile dialogFile(dialogPdf); assert(dialogFile.open(QIODevice::ReadOnly));
    auto printedPages=pages.globalMatch(QString::fromLatin1(dialogFile.readAll())); int printedCount=0;
    while(printedPages.hasNext()) { printedPages.next(); ++printedCount; } assert(printedCount==1);
    bool previewed=false;
    const auto inspectPreview=[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="native-datasheet-preview");
        auto* pages=dialog->findChild<QPrintPreviewWidget*>("datasheet-preview-pages"); assert(pages); pages->updatePreview();
        assert(pages->property("la_datasheet_pages_valid").toBool() && pages->pageCount()>1);
        auto* number=dialog->findChild<QSpinBox*>("datasheet-preview-page"); assert(number && number->maximum()==pages->pageCount());
        number->setValue(pages->pageCount()); assert(pages->currentPage()==pages->pageCount());
        assert(!window.allowProjectClose(db.connectionName()) && window.previewDatasheet().value("error").toInt()==2046);
        previewed=true; dialog->reject();
    };
    QTimer::singleShot(0,&window,inspectPreview);
    window.findChild<QAction*>("datasheet-preview")->trigger(); assert(previewed);
    if(!reader.isEmpty()) {
        QProcess text; text.start(reader,{"-bbox",path,"-"}); assert(text.waitForFinished(10000) && text.exitCode()==0);
        QRegularExpression word("<word[^>]*yMin=\"([^\"]+)\"[^>]*yMax=\"([^\"]+)\"[^>]*>value</word>");
        const auto bounds=word.match(QString::fromUtf8(text.readAllStandardOutput())); assert(bounds.hasMatch());
        const double height=bounds.captured(2).toDouble()-bounds.captured(1).toDouble(); assert(height>6 && height<11);
    }
    assert(model->setData(model->index(0,1),"draft"));
    assert(window.outputDatasheet(true,path).value("error").toInt()==2501 && read()==original && model->isDirty()); model->revertAll();
    view->showColumn(3);
    assert(window.outputDatasheet(true,path).value("error").toInt()==3251 && read()==original); view->hideColumn(3);
    assert(QFile::link(path,output.filePath("alias.pdf")));
    assert(window.outputDatasheet(true,output.filePath("alias.pdf")).value("error").toInt()!=0 && read()==original);
    bool printed=false;
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QPrintDialog*>(QApplication::activeModalWidget()); assert(dialog);
        printed=true; assert(!window.allowProjectClose(db.connectionName()));
        assert(window.outputDatasheet(true,path).value("error").toInt()==2046); dialog->reject();
    });
    QTest::keyClick(view,Qt::Key_P,Qt::ControlModifier); assert(printed && read()==original);
    // The modal dialog freezes actual values before another independent SQL write.
    bool selected=false;
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QFileDialog*>(QApplication::activeModalWidget()); assert(dialog);
        selected=true; assert(!window.allowProjectClose(db.connectionName()));
        assert(sql.exec("UPDATE output_rows SET value='changed_after_snapshot' WHERE id=320"));
        dialog->selectFile(output.filePath("frozen.pdf")); static_cast<QDialog*>(dialog)->accept();
    });
    assert(window.outputDatasheet(true).value("error").toInt()==0 && selected);
    if(!reader.isEmpty()) {
        QProcess text; text.start(reader,{"-layout",output.filePath("frozen.pdf"),"-"}); assert(text.waitForFinished(10000) && text.exitCode()==0);
        const auto content=text.readAllStandardOutput(); assert(content.contains("row_320") && !content.contains("changed_after_snapshot"));
    }
    {
        const QByteArray source="Begin Form\nEnd\n";
        const QByteArray code=R"vb(Sub PrintData()
DoCmd.SelectObject acTable, "output_rows", False
DoCmd.RunCommand acCmdPrint
Me.Tag = "printed"
End Sub
Sub PreviewData()
DoCmd.SelectObject acTable, "output_rows", False
Application.RunCommand Command:=acCmdPrintPreview
Me.Tag = "previewed"
End Sub
)vb";
        AccessFormPreview caller("output-command-caller",source,code,db,&window);
        std::vector<linaccess::VbaValue> arguments;
        const auto nativePdf=output.filePath("vba.pdf");
        QTimer::singleShot(0,&window,[&] {
            auto* dialog=qobject_cast<QPrintDialog*>(QApplication::activeModalWidget()); assert(dialog);
            auto* printer=dialog->printer(); printer->setOutputFormat(QPrinter::PdfFormat); printer->setOutputFileName(nativePdf);
            dialog->QDialog::done(QDialog::Accepted);
        });
        const auto printed=caller.executeProcedure("PrintData",arguments);
        if(printed.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Native VBA print: "<<printed.errorNumber<<" "<<printed.explanation<<"\n";
        assert(printed.status==linaccess::VbaExecutionStatus::Completed && QFileInfo::exists(nativePdf));
        assert(makePreviewFormObject(caller)->member("Tag",{}).asText()=="printed");
        previewed=false; QTimer::singleShot(0,&window,inspectPreview);
        const auto preview=caller.executeProcedure("PreviewData",arguments);
        assert(preview.status==linaccess::VbaExecutionStatus::Completed && previewed && makePreviewFormObject(caller)->member("Tag",{}).asText()=="previewed");
        assert(caller.close());
    }
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QFileDialog*>(QApplication::activeModalWidget()); assert(dialog);
        const int index=tabs->indexOf(page); assert(index>=0);
        emit tabs->tabCloseRequested(index); assert(tabs->indexOf(page)<0);
        dialog->selectFile(path); dialog->setOption(QFileDialog::DontConfirmOverwrite,true); static_cast<QDialog*>(dialog)->accept();
    });
    assert(window.outputDatasheet(true).value("error").toInt()==2501 && read()==original);
    assert(sql.exec("INSERT INTO _la_objects(kind,name,definition) VALUES('query','output_query','{\"sql\":\"SELECT value,id FROM output_rows WHERE id IN (10,20) ORDER BY id DESC\"}')"));
    assert(window.openQueryDocument("output_query",0,2).value("error").toInt()==0);
    auto* queryView=tabs->currentWidget()->findChild<QTableView*>(); assert(queryView);
    queryView->setFocus(); window.activateWindow(); QApplication::processEvents();
    const auto query=linaccess::nativeDatasheetSnapshot(*queryView,"output_query");
    assert(query.rows==QVector<QStringList>({{"row_020","20"},{"row_010","10"}}));
    assert(window.outputDatasheet(true,output.filePath("query.pdf")).value("error").toInt()==0);
    assert(window.closeQueryDocument(1,"output_query",2).value("error").toInt()==0);
    assert(sql.exec("DELETE FROM _la_objects WHERE kind='query' AND name='output_query'"));
    assert(sql.exec("DROP TABLE output_rows"));
}

void exerciseTableReplace(MainWindow& window,QSqlDatabase db) {
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE command_replace(id INTEGER PRIMARY KEY,value TEXT UNIQUE)"));
    assert(sql.exec("INSERT INTO command_replace VALUES(1,'alpha'),(2,'beta')"));
    assert(window.openTableDocument("command_replace",0,1)==0);
    auto* tabs=window.findChild<QTabWidget*>(); assert(tabs && tabs->currentWidget());
    linaccess::AccessTableView* view=nullptr;
    for(auto* candidate:tabs->currentWidget()->findChildren<QTableView*>()) if(auto* typed=dynamic_cast<linaccess::AccessTableView*>(candidate)) view=typed;
    assert(view); auto* model=qobject_cast<linaccess::AccessTableModel*>(view->model()); assert(model);
    view->sortByColumn(0,Qt::AscendingOrder); assert(model->index(0,0).data().toInt()==1);
    view->setCurrentIndex(model->index(0,1)); QApplication::setActiveWindow(&window); view->setFocus(); QApplication::processEvents();
    QTest::keyClick(view,Qt::Key_H,Qt::ControlModifier); QApplication::processEvents();
    QPointer<QDialog> dialog=view->findChild<QDialog*>("native-table-replace-dialog"); assert(dialog && dialog->isVisible() && !dialog->isModal());
    auto* sought=dialog->findChild<QLineEdit*>("native-table-find-text");
    auto* replacement=dialog->findChild<QLineEdit*>("native-table-replace-with");
    auto* apply=dialog->findChild<QPushButton*>("native-table-replace-current");
    auto* next=dialog->findChild<QPushButton*>("native-table-find-next");
    auto* match=dialog->findChild<QComboBox*>("native-table-find-match");
    assert(sought && replacement && apply && next && match && match->currentIndex()==1 && !match->isEnabled());
    sought->setText("alpha"); replacement->setText("gamma"); next->click(); apply->click();
    assert(model->isDirty() && model->index(0,1).data(Qt::EditRole).toString()=="gamma");
    assert(sql.exec("SELECT value FROM command_replace WHERE id=1") && sql.next() && sql.value(0).toString()=="alpha"); sql.finish();
    assert(window.dispatchTableCommand(97).value("error").toInt()==0 && !model->isDirty());
    assert(sql.exec("SELECT value FROM command_replace WHERE id=1") && sql.next() && sql.value(0).toString()=="gamma"); sql.finish();
    sought->setText("gamma"); replacement->setText("beta"); apply->click();
    assert(model->isDirty() && window.dispatchTableCommand(97).value("error").toInt()==3146);
    assert(model->isDirty() && model->index(0,1).data(Qt::EditRole).toString()=="beta");
    assert(sql.exec("SELECT value FROM command_replace WHERE id=1") && sql.next() && sql.value(0).toString()=="gamma"); sql.finish();
    model->revertAll(); view->setCurrentIndex(model->index(0,1)); assert(model->setReadOnly(true));
    assert(window.dispatchTableCommand(29).value("error").toInt()==2046);
    sought->setText("gamma"); replacement->setText("forbidden"); apply->click();
    assert(!model->isDirty() && model->index(0,1).data(Qt::EditRole).toString()=="gamma");
    dialog->reject(); dispose(); assert(!dialog);
    assert(model->setReadOnly(false));
    QApplication::setActiveWindow(&window); view->setFocus(); QApplication::processEvents();
    auto* edit=window.findChild<QMenu*>("native-edit-menu"); auto* menu=window.findChild<QAction*>("table-replace"); assert(edit && menu);
    assert(QMetaObject::invokeMethod(edit,"aboutToShow",Qt::DirectConnection) && menu->isEnabled()); menu->trigger();
    dialog=view->findChild<QDialog*>("native-table-replace-dialog"); assert(dialog && dialog->isVisible());
    view->setCurrentIndex(model->index(0,1));
    view->setRecordNavigationGuard([view,model] { view->setCurrentIndex(model->index(1,1)); return true; });
    assert(!view->replaceCurrentText({"beta",1,1},"unintended"));
    assert(!model->isDirty() && model->lastError().nativeErrorCode()=="2501" && model->index(1,1).data().toString()=="beta");
    assert(window.closeTableDocument(0,"command_replace",2).value("error").toInt()==0); dispose(); assert(!dialog);
    assert(sql.exec("CREATE TABLE command_replace_all(id INTEGER PRIMARY KEY,value TEXT CHECK(id<>2 OR value<>'blocked'))"));
    assert(sql.exec("INSERT INTO command_replace_all VALUES(1,'Äpfel'),(2,'äpfel'),(3,NULL),(4,'other')"));
    assert(window.openTableDocument("command_replace_all",0,1)==0);
    view=nullptr;
    for(auto* candidate:tabs->currentWidget()->findChildren<QTableView*>()) if(auto* typed=dynamic_cast<linaccess::AccessTableView*>(candidate)) view=typed;
    assert(view); model=qobject_cast<linaccess::AccessTableModel*>(view->model()); assert(model);
    view->sortByColumn(1,Qt::DescendingOrder);
    for(int row=0;row<model->rowCount();++row) if(model->record(row).value("id").toInt()==2) view->setCurrentIndex(model->index(row,1));
    assert(view->openReplaceDialog()); dialog=view->findChild<QDialog*>("native-table-replace-dialog"); assert(dialog);
    sought=dialog->findChild<QLineEdit*>("native-table-find-text"); replacement=dialog->findChild<QLineEdit*>("native-table-replace-with");
    auto* all=dialog->findChild<QPushButton*>("native-table-replace-all"); auto* direction=dialog->findChild<QComboBox*>("native-table-find-direction");
    assert(sought && replacement && all && direction);
    sought->setText("ÄPFEL"); replacement->setText("Obst");
    direction->setCurrentIndex(direction->findData(1)); assert(!all->isEnabled());
    direction->setCurrentIndex(direction->findData(2)); assert(all->isEnabled()); all->click();
    assert(!model->isDirty() && model->record(view->currentIndex().row()).value("id").toInt()==2);
    assert(sql.exec("SELECT COUNT(*) FROM command_replace_all WHERE value='Obst'") && sql.next() && sql.value(0).toInt()==2); sql.finish();
    assert(sql.exec("SELECT value FROM command_replace_all WHERE id=3") && sql.next() && sql.value(0).isNull()); sql.finish();
    sought->setText("Obst"); replacement->setText("blocked"); all->click();
    assert(!model->isDirty());
    assert(sql.exec("SELECT COUNT(*) FROM command_replace_all WHERE value='Obst'") && sql.next() && sql.value(0).toInt()==2); sql.finish();
    assert(sql.exec("SELECT COUNT(*) FROM command_replace_all WHERE value='blocked'") && sql.next() && sql.value(0).toInt()==0); sql.finish();
    int draft=-1;
    for(int row=0;row<model->rowCount();++row) if(model->record(row).value("id").toInt()==4) draft=row;
    assert(draft>=0 && model->setData(model->index(draft,1),"unsaved",Qt::EditRole));
    replacement->setText("unintended"); all->click(); assert(model->isDirty() && model->index(draft,1).data(Qt::EditRole).toString()=="unsaved");
    assert(sql.exec("SELECT COUNT(*) FROM command_replace_all WHERE value='Obst'") && sql.next() && sql.value(0).toInt()==2); sql.finish();
    model->revertAll(); assert(window.closeTableDocument(0,"command_replace_all",2).value("error").toInt()==0); dispose(); assert(!dialog);
}
void exerciseTableMultiDelete(MainWindow& window,QSqlDatabase db) {
    QSqlQuery sql(db); assert(sql.exec("PRAGMA foreign_keys=ON"));
    assert(sql.exec("CREATE TABLE delete_many(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(sql.exec("INSERT INTO delete_many VALUES(1,'one'),(2,'two'),(3,'three')"));
    assert(sql.exec("CREATE TABLE delete_many_child(parent INTEGER REFERENCES delete_many(id))"));
    assert(sql.exec("INSERT INTO delete_many_child VALUES(3)"));
    assert(window.openTableDocument("delete_many",0,1)==0);
    auto* tabs=window.findChild<QTabWidget*>(); linaccess::AccessTableView* view=nullptr;
    for(auto* candidate:tabs->currentWidget()->findChildren<QTableView*>()) if(auto* typed=dynamic_cast<linaccess::AccessTableView*>(candidate)) view=typed;
    assert(view); auto* model=qobject_cast<linaccess::AccessTableModel*>(view->model()); assert(model);
    view->sortByColumn(0,Qt::AscendingOrder); QApplication::setActiveWindow(&window); view->setFocus(); QApplication::processEvents();
    QItemSelection cells(model->index(0,1),model->index(2,1));
    view->selectionModel()->select(cells,QItemSelectionModel::ClearAndSelect);
    view->selectionModel()->setCurrentIndex(model->index(0,1),QItemSelectionModel::NoUpdate);
    QTest::keyClick(view,Qt::Key_Minus,Qt::ControlModifier); assert(!model->isDirty());
    const auto selectRecords=[&] {
        QItemSelection rows; rows.select(model->index(0,0),model->index(0,1)); rows.select(model->index(2,0),model->index(2,1));
        view->selectionModel()->select(rows,QItemSelectionModel::ClearAndSelect|QItemSelectionModel::Rows);
        view->selectionModel()->setCurrentIndex(model->index(0,1),QItemSelectionModel::NoUpdate);
    };
    selectRecords(); QTest::keyClick(view,Qt::Key_Minus,Qt::ControlModifier); assert(model->isDirty());
    assert(sql.exec("SELECT COUNT(*) FROM delete_many") && sql.next() && sql.value(0).toInt()==3); sql.finish();
    assert(window.dispatchTableCommand(97).value("error").toInt()==3146);
    assert(sql.exec("SELECT COUNT(*) FROM delete_many") && sql.next() && sql.value(0).toInt()==3); sql.finish();
    model->revertAll(); assert(!model->isDirty());
    assert(sql.exec("DELETE FROM delete_many_child")); sql.finish();
    selectRecords(); QTest::keyClick(view,Qt::Key_Minus,Qt::ControlModifier);
    assert(window.dispatchTableCommand(97).value("error").toInt()==0 && !model->isDirty());
    assert(sql.exec("SELECT id FROM delete_many") && sql.next() && sql.value(0).toInt()==2 && !sql.next()); sql.finish();
    assert(window.closeTableDocument(0,"delete_many",2).value("error").toInt()==0); dispose();
}
void exerciseStoredForm(MainWindow& window,QSqlDatabase db) {
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE native_designed_records(id INTEGER PRIMARY KEY,first TEXT,\"second value\" TEXT)"));
    assert(sql.exec("INSERT INTO native_designed_records VALUES(1,'before','initial')"));
    const QString module=R"VBA(Private state As Long
Private mouseTrace As String
Private heldOwner As Access.Form
Private heldControl As Control
Function TypedSelf() As Form
Set TypedSelf=Me
End Function
Sub ClearForm(ByRef value As Access.Form)
Set value=Nothing
End Sub
Sub MarkControl(ByVal value As Access.Control)
value.Tag="typed-control"
End Sub
Function IsFirstControl(ByVal unused As Long, ByVal value As Control) As Boolean
IsFirstControl=value.Name="first"
End Function
Function Len(ByVal value As Control) As Long
Len=VBA.Len(value.Name)
End Function
Function EmptyControl(Optional ByVal value As Control=Nothing) As Boolean
EmptyControl=value Is Nothing
End Function
Sub EditTypedField()
Dim target As TextBox
Set target=Me!first
target="typed edit"
End Sub
Private Sub Form_Load()
state=1
Dim owner As Form, field As Access.TextBox, wrong As Object, slotError As Long, argumentError As Long
Set owner=TypedSelf()
Set heldOwner=owner
Set field=Me!first
Set heldControl=field
MarkControl Me!first
If Len(Me!first)<>5 Then Err.Raise 5
If Not IsFirstControl(value:=Me!first,unused:=1) Then Err.Raise 5
ClearForm owner
If Not owner Is Nothing Then Err.Raise 5
If Not EmptyControl() Then Err.Raise 5
Set wrong=CurrentDb
On Error Resume Next
Set heldOwner=wrong
slotError=Err.Number
Err.Clear
MarkControl wrong
argumentError=Err.Number
On Error GoTo 0
If slotError<>13 Or argumentError<>13 Then Err.Raise 5
If Not heldOwner Is Me Then Err.Raise 5
If field.Name<>"first" Then Err.Raise 5
If Not TypeOf Me Is Access.Form Then Err.Raise 5
If Not TypeOf Me!first Is TextBox Then Err.Raise 5
If Not TypeOf Me!btnSpeichern Is CommandButton Then Err.Raise 5
If Not TypeOf Me.Controls Is Access.Controls Then Err.Raise 5
If TypeOf Me Is Report Then Err.Raise 5
End Sub
Private Sub first_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button=1 Then mouseTrace=CStr(Button) & ":" & CStr(Shift) & ":" & CStr(X) & ":" & CStr(Y)
End Sub
Private Sub first_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button=1 Then mouseTrace=mouseTrace & "|up"
End Sub
Private Sub first_Click()
Me.Tag=mouseTrace & "|click"
Me!first.SelStart=0
Me!first.SelLength=0
End Sub
Private Sub first_DblClick(Cancel As Integer)
Me.Tag="text-double"
Cancel=True
End Sub
Private Sub btnSpeichern_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Tag="save-down"
End Sub
Private Sub btnSpeichern_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.Tag<>"save-down" Then Err.Raise 5
Me!first=Me!first & "|up"
End Sub
Private Sub btnAktion_Click()
state=state+1
Me![second value]=CStr(state)
DoCmd.RunCommand acCmdSaveRecord
End Sub
Sub OpenOwnedChild()
DoCmd.OpenForm "Owned Child"
DoCmd.Close acForm,"Owned Native Form"
End Sub
)VBA";
    auto* toolbar=window.findChild<QToolBar*>(); assert(toolbar);
    QAction* create=nullptr; for(auto* action:toolbar->actions()) if(action->text()=="+ Formular") create=action;
    assert(create);
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="native-form-designer");
        dialog->findChild<QLineEdit*>("native-form-name")->setText("Owned Native Form");
        dialog->findChild<QLineEdit*>("native-form-caption")->setText("Owned form caption");
        dialog->findChild<QComboBox*>("native-form-source")->setCurrentText("native_designed_records");
        auto* fields=dialog->findChild<QListWidget*>("native-form-fields"); assert(fields && fields->count()==2);
        auto* moved=fields->takeItem(1); assert(moved->data(Qt::UserRole).toString()=="second value");
        moved->setText("Second \"caption\""); fields->insertItem(0,moved);
        fields->setCurrentItem(moved);
        dialog->findChild<QSpinBox*>("native-form-field-left")->setValue(2400);
        dialog->findChild<QSpinBox*>("native-form-field-width")->setValue(2400);
        auto* code=dialog->findChild<QPlainTextEdit*>("native-form-vba"); code->setPlainText(module);
        QApplication::setActiveWindow(dialog); code->setFocus(); QApplication::processEvents();
        QTest::keyClick(code,Qt::Key_S,Qt::ControlModifier);
        assert(!dialog->isWindowModified());
        assert(dialog->isVisible() && QApplication::activeModalWidget()==dialog);
        assert(sql.exec("SELECT definition FROM main._la_objects WHERE kind='form' AND name='Owned Native Form'") && sql.next());
        assert(QJsonDocument::fromJson(sql.value(0).toByteArray()).object().value("vba").toString()==module); sql.finish();
        dialog->findChild<QPushButton*>("native-form-save")->click();
    });
    create->trigger(); QApplication::processEvents();
    assert(sql.exec("SELECT definition FROM _la_objects WHERE kind='form' AND name='Owned Native Form'") && sql.next());
    const auto saved=QJsonDocument::fromJson(sql.value(0).toByteArray()).object(); sql.finish();
    assert(saved.value("fields").toArray().first().toString()=="second value" && saved.value("vba").toString()==module);
    assert(saved.value("field_layout").toObject().value("second value").toObject().value("width").toInt()==2400);
    auto child=saved; child.insert("vba",""); child.insert("caption","Owned child caption");
    sql.prepare("INSERT INTO _la_objects(kind,name,definition) VALUES('form','Owned Child',?)");
    sql.addBindValue(QString::fromUtf8(QJsonDocument(child).toJson(QJsonDocument::Compact))); assert(sql.exec()); sql.finish();
    auto* explorer=window.findChild<QTreeWidget*>(); assert(explorer);
    for(auto* action:toolbar->actions()) if(action->text()=="Aktualisieren") action->trigger();
    QTreeWidgetItem* selected=nullptr;
    for(QTreeWidgetItemIterator item(explorer);*item;++item)
        if((*item)->data(0,Qt::UserRole).toString()=="form" && (*item)->text(0)=="Owned Native Form") selected=*item;
    assert(selected); explorer->setCurrentItem(selected);
    const auto open=[&] { assert(QMetaObject::invokeMethod(explorer,"itemDoubleClicked",Qt::DirectConnection,Q_ARG(QTreeWidgetItem*,selected),Q_ARG(int,0))); QApplication::processEvents(); };
    const auto locate=[&](const QString& name) {
        QPointer<AccessFormPreview> result;
        for(const auto& document:AccessFormPreview::projectDocuments(db)) if(document && document->documentName()==name) result=document;
        return result;
    };
    assert(sql.exec("CREATE TEMP TABLE _la_objects(kind TEXT,name TEXT,definition TEXT)"));
    assert(sql.exec("INSERT INTO temp._la_objects VALUES('form','Owned Native Form','{}')"));
    open(); auto form=locate("Owned Native Form"); assert(form && form->hasDataSource() && form->windowTitle().contains("Owned form caption"));
    assert(sql.exec("DROP TABLE temp._la_objects"));
    auto* first=form->findChild<QLineEdit*>("access-control-first"); auto* second=form->findChild<QLineEdit*>("access-control-second value");
    assert(first && second && second->y()<first->y() && second->x()>first->x() && second->width()<first->width() && first->text()=="before");
    bool caption=false; for(auto* label:form->findChildren<QLabel*>()) if(label->text()=="Second \"caption\"") caption=true;
    assert(caption && AccessFormPreview::projectObjectLoaded(db,linaccess::AccessDocumentKind::Form,"Owned Native Form"));
    auto formObject=makePreviewFormObject(*form);
    const auto loadResult=form->eventResult("OnLoad");
    if(loadResult && loadResult->status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Typed Access load error="<<loadResult->errorNumber<<" line="<<loadResult->line<<" explanation="<<loadResult->explanation<<"\n";
    assert(loadResult && loadResult->status==linaccess::VbaExecutionStatus::Completed);
    assert(first->property("la_access_tag").toString()=="typed-control");
    std::vector<linaccess::VbaValue> typedArguments;
    assert(form->executeProcedure("EditTypedField",typedArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(first->text()=="typed edit" && form->hasUnsavedChanges());
    assert(sql.exec("SELECT first FROM native_designed_records WHERE id=1") && sql.next() && sql.value(0).toString()=="before"); sql.finish();
    (void)formObject->member("Undo",{}); assert(first->text()=="before" && !form->hasUnsavedChanges());
    QApplication::setActiveWindow(form); first->setFocus();
    QTest::mouseClick(first,Qt::LeftButton,Qt::ShiftModifier|Qt::ControlModifier,QPoint(3,4));
    assert(formObject->member("Tag",{}).asText()=="1:3:45:60|up|click" && !formObject->member("Dirty",{}).boolean && first->cursorPosition()==0);
    formObject->assign("Tag",linaccess::VbaValue::string("right-click"));
    QTest::mouseClick(first,Qt::RightButton);
    assert(formObject->member("Tag",{}).asText()=="right-click");
    QTest::mouseDClick(first,Qt::LeftButton);
    assert(formObject->member("Tag",{}).asText()=="text-double");
    first->setFocus(); first->selectAll(); QTest::keyClicks(first,"native save");
    auto* save=form->findChild<QPushButton*>("access-control-btnSpeichern"); auto* action=form->findChild<QPushButton*>("access-control-btnAktion"); assert(save && action);
    QTest::mouseClick(save,Qt::LeftButton); QApplication::processEvents();
    assert(sql.exec("SELECT first FROM native_designed_records WHERE id=1") && sql.next() && sql.value(0).toString()=="native save|up"); sql.finish();
    QTest::mouseClick(action,Qt::LeftButton); QTest::mouseClick(action,Qt::LeftButton); QApplication::processEvents();
    assert(sql.exec("SELECT \"second value\" FROM native_designed_records WHERE id=1") && sql.next() && sql.value(0).toString()=="3"); sql.finish();
    std::vector<linaccess::VbaValue> arguments;
    assert(form->executeProcedure("OpenOwnedChild",arguments).status==linaccess::VbaExecutionStatus::Completed); dispose();
    auto ownedChild=locate("Owned Child"); assert(ownedChild && ownedChild->isVisible() && !AccessFormPreview::projectObjectLoaded(db,linaccess::AccessDocumentKind::Form,"Owned Native Form"));
    assert(ownedChild->close()); dispose();
    {
        const QByteArray layout="Begin Form\nBegin Section\nBegin Subform\nName =\"child\"\nSourceObject =\"Form.Owned Child\"\nWidth =6000\nHeight =3000\nEnd\nEnd\nEnd\n";
        AccessFormPreview parent("Owned Parent",layout,{},db,&window);
        AccessFormPreview* embedded=nullptr;
        for(auto* dialog:parent.findChildren<QDialog*>()) if(auto* document=dynamic_cast<AccessFormPreview*>(dialog)) embedded=document;
        assert(embedded && embedded->documentName()=="Owned Child" && !embedded->isWindow() && embedded->hasDataSource());
        auto childObject=makePreviewFormObject(*embedded);
        childObject->assign("first",linaccess::VbaValue::string("embedded save"));
        assert(childObject->member("Dirty",{}).boolean);
        childObject->assign("Dirty",linaccess::VbaValue::logical(false));
        assert(sql.exec("SELECT first FROM native_designed_records WHERE id=1") && sql.next() && sql.value(0).toString()=="embedded save"); sql.finish();
        assert(AccessFormPreview::projectObjectLoaded(db,linaccess::AccessDocumentKind::Form,"Owned Child"));
        assert(parent.close() && !AccessFormPreview::projectObjectLoaded(db,linaccess::AccessDocumentKind::Form,"Owned Child"));
    }
    open(); form=locate("Owned Native Form"); assert(form);
    auto* edit=window.findChild<QAction*>("native-form-edit"); assert(edit);
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="native-form-designer");
        assert(!AccessFormPreview::projectObjectLoaded(db,linaccess::AccessDocumentKind::Form,"Owned Native Form"));
        const auto* fields=dialog->findChild<QListWidget*>("native-form-fields");
        assert(fields->item(0)->data(Qt::UserRole).toString()=="second value" && fields->item(0)->text()=="Second \"caption\"");
        dialog->findChild<QListWidget*>("native-form-fields")->setCurrentRow(0);
        assert(dialog->findChild<QSpinBox*>("native-form-field-left")->value()==2400);
        assert(dialog->findChild<QSpinBox*>("native-form-field-width")->value()==2400);
        assert(dialog->findChild<QPlainTextEdit*>("native-form-vba")->toPlainText()==module);
        dialog->reject();
    });
    edit->trigger(); dispose();
    const QString totalsSql="SELECT COUNT(*) AS total FROM [Owned Form ZBase]";
    sql.prepare("INSERT INTO main._la_objects(kind,name,definition) VALUES('query','Owned Form ZBase',?)");
    sql.addBindValue(QString::fromUtf8(QJsonDocument(QJsonObject{{"sql","SELECT id FROM native_designed_records"}}).toJson(QJsonDocument::Compact)));
    assert(sql.exec()); sql.finish();
    sql.prepare("INSERT INTO main._la_objects(kind,name,definition) VALUES('query','Owned Form Totals',?)");
    sql.addBindValue(QString::fromUtf8(QJsonDocument(QJsonObject{{"sql",totalsSql}}).toJson(QJsonDocument::Compact)));
    assert(sql.exec()); sql.finish();
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog);
        dialog->findChild<QLineEdit*>("native-form-name")->setText("Owned Query Form");
        auto* source=dialog->findChild<QComboBox*>("native-form-source"); assert(source->isEditable());
        source->setCurrentText("SELECT 42 AS probe");
        auto* fields=dialog->findChild<QListWidget*>("native-form-fields");
        assert(fields->count()==1 && fields->item(0)->data(Qt::UserRole).toString()=="probe");
        source->setCurrentText("Owned Form Totals");
        assert(fields->count()==1 && fields->item(0)->data(Qt::UserRole).toString()=="total");
        dialog->findChild<QPushButton*>("native-form-save")->click();
    });
    create->trigger(); QApplication::processEvents();
    selected=nullptr;
    for(QTreeWidgetItemIterator item(explorer);*item;++item)
        if((*item)->data(0,Qt::UserRole).toString()=="form" && (*item)->text(0)=="Owned Query Form") selected=*item;
    assert(selected); explorer->setCurrentItem(selected); open();
    auto queryForm=locate("Owned Query Form"); assert(queryForm && queryForm->hasDataSource());
    auto* total=queryForm->findChild<QLineEdit*>("access-control-total"); assert(total && total->isReadOnly());
    assert(sql.exec(totalsSql) && sql.next()); const int beforeCount=sql.value(0).toInt(); sql.finish();
    assert(total->text().toInt()==beforeCount);
    assert(!makePreviewFormObject(*queryForm)->member("AllowAdditions",{}).boolean);
    assert(!queryForm->findChild<QPushButton*>("access-control-btnNeu"));
    assert(!queryForm->findChild<QPushButton*>("access-control-btnSpeichern"));
    assert(sql.exec("INSERT INTO native_designed_records(first) VALUES('query change')")); sql.finish();
    (void)makePreviewFormObject(*queryForm)->member("Requery",{});
    assert(total->text().toInt()==beforeCount+1);
    assert(makePreviewApplicationObject(*queryForm)->member("Eval",{linaccess::VbaValue::string(R"VBA(DLookup("id","Owned Form ZBase","id=1"))VBA")}).asNumber()==1);
    sql.prepare("UPDATE main._la_objects SET definition=? WHERE kind='query' AND name='Owned Form ZBase'");
    sql.addBindValue(QString::fromUtf8(QJsonDocument(QJsonObject{{"sql","SELECT id FROM native_designed_records WHERE id<0"}}).toJson(QJsonDocument::Compact)));
    assert(sql.exec()); sql.finish();
    (void)makePreviewFormObject(*queryForm)->member("Requery",{});
    assert(total->text()=="0");
    assert(queryForm->close()); dispose();
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog);
        assert(dialog->findChild<QComboBox*>("native-form-source")->currentText()=="Owned Form Totals");
        assert(dialog->findChild<QListWidget*>("native-form-fields")->count()==1);
        dialog->reject();
    });
    explorer->setCurrentItem(selected); edit->trigger(); dispose();
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="native-query-designer");
        dialog->findChild<QLineEdit*>("native-query-name")->setText("Shortcut Saved Query");
        auto* code=dialog->findChild<QPlainTextEdit*>("native-query-sql"); code->setPlainText("SELECT 42 AS value");
        assert(dialog->isWindowModified());
        QApplication::setActiveWindow(dialog); code->setFocus(); QApplication::processEvents();
        QTest::keyClick(code,Qt::Key_S,Qt::ControlModifier);
        assert(!dialog->isWindowModified());
        assert(dialog->isVisible() && QApplication::activeModalWidget()==dialog);
        assert(sql.exec("SELECT definition FROM main._la_objects WHERE kind='query' AND name='Shortcut Saved Query'") && sql.next());
        auto definition=QJsonDocument::fromJson(sql.value(0).toByteArray()).object(); sql.finish();
        assert(!dialog->isWindowModified() && !dialog->property("la_designer_new").toBool());
        assert(definition.value("sql").toString()=="SELECT 42 AS value"); definition.insert("archive","preserve");
        sql.prepare("UPDATE main._la_objects SET definition=? WHERE kind='query' AND name='Shortcut Saved Query'");
        sql.addBindValue(QString::fromUtf8(QJsonDocument(definition).toJson(QJsonDocument::Compact))); assert(sql.exec()); sql.finish();
        code->setPlainText("SELECT 43 AS value");
        assert(dialog->isWindowModified());
        assert(dialog->findChild<QLabel*>("native-designer-save-status")->text()=="Nicht gespeicherte Änderungen");
        QTest::keyClick(code,Qt::Key_S,Qt::ControlModifier);
        assert(!dialog->isWindowModified());
        assert(dialog->isVisible() && QApplication::activeModalWidget()==dialog);
        assert(sql.exec("SELECT definition FROM main._la_objects WHERE kind='query' AND name='Shortcut Saved Query'") && sql.next());
        definition=QJsonDocument::fromJson(sql.value(0).toByteArray()).object(); sql.finish();
        assert(definition.value("sql").toString()=="SELECT 43 AS value" && definition.value("archive").toString()=="preserve");
        dialog->findChild<QPushButton*>("native-query-run")->click();
        auto* results=dialog->findChild<QTableView*>(); assert(results && results->model()->rowCount()==1 && results->model()->index(0,0).data().toInt()==43);
        dialog->reject();
    });
    QAction* createQuery=nullptr;
    for(auto* candidate:toolbar->actions()) if(candidate->text()=="+ Abfrage") createQuery=candidate;
    assert(createQuery); createQuery->trigger(); dispose();
    assert(sql.exec("DELETE FROM main._la_objects WHERE kind='query' AND name='Shortcut Saved Query'")); sql.finish();
}

void exerciseStoredReport(MainWindow& window,QSqlDatabase db) {
    using namespace linaccess;
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE native_report_rows(id INTEGER PRIMARY KEY,name TEXT)"));
    assert(sql.exec("INSERT INTO native_report_rows VALUES(1,'first'),(2,'second'),(3,'third')"));
    assert(sql.exec("CREATE TABLE native_report_events(kind TEXT)"));
    const QString source="SELECT TOP 2 id,\n name AS [Display value]\n FROM native_report_rows ORDER BY id";
    const QString module=R"VBA(Private counter As Long
Sub Report_Open(Cancel As Integer)
Dim typed As Report
Set typed=Me
If Not TypeOf typed Is Access.Report Then Err.Raise 5
CurrentDb.Execute "INSERT INTO native_report_events VALUES('open')"
End Sub
Sub Report_Load()
counter=counter+1
Me.Tag=CStr(counter)
Me![Display value].Tag="bound"
If Not Me![Display value].CanGrow Then Err.Raise 5
If Me.Section(acHeader).ForceNewPage<>2 Then Err.Raise 5
End Sub
Sub Increment()
counter=counter+1
Me.Tag=CStr(counter)
End Sub
Sub Report_Close()
CurrentDb.Execute "INSERT INTO native_report_events VALUES('close')"
End Sub
)VBA";
    auto* create=action(window,"+ Bericht"); assert(create);
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="native-report-designer");
        dialog->findChild<QLineEdit*>("native-report-name")->setText("Owned Native Report");
        dialog->findChild<QLineEdit*>("native-report-title")->setText("Owned report title");
        dialog->findChild<QPlainTextEdit*>("native-report-sql")->setPlainText(source);
        dialog->findChild<QPlainTextEdit*>("native-report-vba")->setPlainText(module);
        dialog->findChild<QCheckBox*>("native-report-cangrow")->setChecked(true);
        dialog->findChild<QComboBox*>("native-report-page-break")->setCurrentIndex(3);
        dialog->findChild<QComboBox*>("native-report-header-break")->setCurrentIndex(2);
        auto* code=dialog->findChild<QPlainTextEdit*>("native-report-vba");
        QApplication::setActiveWindow(dialog); code->setFocus(); QApplication::processEvents();
        QTest::keyClick(code,Qt::Key_S,Qt::ControlModifier);
        assert(!dialog->isWindowModified());
        assert(dialog->isVisible() && QApplication::activeModalWidget()==dialog);
        assert(sql.exec("SELECT definition FROM main._la_objects WHERE kind='report' AND name='Owned Native Report'") && sql.next());
        assert(QJsonDocument::fromJson(sql.value(0).toByteArray()).object().value("sql").toString()==source); sql.finish();
        dialog->findChild<QPushButton*>("native-report-save")->click();
    });
    create->trigger(); QApplication::processEvents();
    assert(sql.exec("SELECT definition FROM main._la_objects WHERE kind='report' AND name='Owned Native Report'") && sql.next());
    const auto saved=QJsonDocument::fromJson(sql.value(0).toByteArray()).object(); sql.finish();
    assert(saved.value("sql").toString()==source && saved.value("vba").toString()==module);
    assert(saved.value("can_grow").toBool());
    assert(saved.value("detail_force_new_page").toInt()==3 && saved.value("header_force_new_page").toInt()==2);
    assert(sql.exec("SELECT COUNT(*) FROM native_report_events") && sql.next() && sql.value(0).toInt()==0); sql.finish();
    auto* explorer=window.findChild<QTreeWidget*>("database-explorer"); assert(explorer);
    QTreeWidgetItem* selected=nullptr;
    for(QTreeWidgetItemIterator item(explorer);*item;++item)
        if((*item)->data(0,Qt::UserRole).toString()=="report" && (*item)->text(0)=="Owned Native Report") selected=*item;
    assert(selected); explorer->setCurrentItem(selected);
    assert(sql.exec("CREATE TEMP TABLE _la_objects(kind TEXT,name TEXT,definition TEXT)"));
    assert(sql.exec("INSERT INTO temp._la_objects VALUES('report','Owned Native Report','{}')"));
    assert(QMetaObject::invokeMethod(explorer,"itemDoubleClicked",Qt::DirectConnection,Q_ARG(QTreeWidgetItem*,selected),Q_ARG(int,0)));
    QApplication::processEvents();
    const auto locate=[&](const QString& name=QStringLiteral("Owned Native Report")) {
        QPointer<AccessFormPreview> found;
        for(const auto& document:AccessFormPreview::projectDocuments(db))
            if(document && document->documentKind()==AccessDocumentKind::Report && document->documentName()==name) found=document;
        return found;
    };
    auto report=locate(); assert(report && report->parentWidget()==&window && report->hasDataSource());
    assert(report->eventResult("OnLoad") && report->eventResult("OnLoad")->status==VbaExecutionStatus::Completed);
    auto object=makePreviewFormObject(*report);
    assert(object->member("RecordSource",{}).asText()==source.toStdString() && object->member("Tag",{}).asText()=="1");
    auto* value=report->findChild<QLineEdit*>("access-control-Display value"); assert(value && value->isReadOnly() && value->text()=="first" && value->property("la_access_tag").toString()=="bound");
    assert(report->findChild<QComboBox*>("access-record-selector")->count()==3);
    assert(report->findChild<QPrintPreviewWidget*>("access-report-pages")->pageCount()==3);
    QTimer::singleShot(0,report,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="access-report-page-setup-dialog");
        auto* save=dialog->findChild<QCheckBox*>("report-page-save"); assert(save && save->isEnabled()); save->setChecked(true);
        dialog->findChild<QComboBox*>("report-page-orientation")->setCurrentIndex(1);
        dialog->findChild<QDoubleSpinBox*>("report-page-left")->setValue(25);
        dialog->findChild<QDialogButtonBox*>("report-page-buttons")->button(QDialogButtonBox::Ok)->click();
    });
    report->findChild<QAction*>("access-report-page-setup")->trigger();
    assert(sql.exec("SELECT definition FROM main._la_objects WHERE kind='report' AND name='Owned Native Report'") && sql.next());
    const auto withPage=QJsonDocument::fromJson(sql.value(0).toByteArray()).object(); sql.finish();
    assert(withPage.value("sql").toString()==source && withPage.value("vba").toString()==module);
    assert(withPage.value("page_setup").toObject().value("orientation").toInt()==1);
    assert(sql.exec("SELECT definition FROM temp._la_objects WHERE kind='report'") && sql.next() && sql.value(0).toString()=="{}"); sql.finish();
    std::vector<VbaValue> none;
    assert(report->executeProcedure("Increment",none).status==VbaExecutionStatus::Completed && object->member("Tag",{}).asText()=="2");
    QTemporaryDir output; assert(output.isValid() && report->exportReportPdf(output.filePath("owned.pdf")));
    assert(report->close()); dispose(); assert(!locate());
    auto filtered=saved; auto filteredSource=source; filteredSource.replace("TOP 2 ",""); filtered.insert("sql",filteredSource);
    sql.prepare("INSERT INTO main._la_objects(kind,name,definition) VALUES('report','Owned Filtered Report',?)");
    sql.addBindValue(QString::fromUtf8(QJsonDocument(filtered).toJson(QJsonDocument::Compact))); assert(sql.exec()); sql.finish();
    const QByteArray callerSource="Begin Form\nBegin Section\nHeight =300\nEnd\nEnd";
    const QByteArray callerCode="Sub OpenStored()\nDoCmd.OpenReport \"owned filtered report\", acViewPreview, , \"id=2\"\nEnd Sub";
    AccessFormPreview caller("Stored report caller",callerSource,callerCode,db,&window);
    caller.show(); QApplication::setActiveWindow(&caller); QApplication::processEvents();
    const auto opened=caller.executeProcedure("OpenStored",none);
    if (opened.status!=VbaExecutionStatus::Completed) std::cerr<<"Stored report OpenReport: error="<<opened.errorNumber<<", line="<<opened.line<<", "<<opened.explanation<<'\n';
    assert(opened.status==VbaExecutionStatus::Completed);
    report=locate("Owned Filtered Report"); assert(report && report->findChild<QLineEdit*>("access-control-Display value")->text()=="second");
    assert(caller.close() && AccessFormPreview::projectObjectLoaded(db,AccessDocumentKind::Report,"Owned Filtered Report"));
    assert(report->close()); dispose();
    assert(sql.exec("DROP TABLE temp._la_objects"));
    assert(QMetaObject::invokeMethod(explorer,"itemDoubleClicked",Qt::DirectConnection,Q_ARG(QTreeWidgetItem*,selected),Q_ARG(int,0)));
    assert(locate());
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="access-report-page-setup-dialog");
        assert(dialog->findChild<QComboBox*>("report-page-orientation")->currentIndex()==1);
        assert(dialog->findChild<QDoubleSpinBox*>("report-page-left")->value()==25);
        dialog->reject();
    });
    locate()->findChild<QAction*>("access-report-page-setup")->trigger();
    assert(locate()->close()); dispose();
    auto invalidDesign=withPage; auto invalidPage=invalidDesign.value("page_setup").toObject();
    invalidPage.insert("orientation",2); invalidDesign.insert("page_setup",invalidPage);
    sql.prepare("UPDATE main._la_objects SET definition=? WHERE kind='report' AND name='Owned Native Report'");
    sql.addBindValue(QString::fromUtf8(QJsonDocument(invalidDesign).toJson(QJsonDocument::Compact))); assert(sql.exec()); sql.finish();
    assert(QMetaObject::invokeMethod(explorer,"itemDoubleClicked",Qt::DirectConnection,Q_ARG(QTreeWidgetItem*,selected),Q_ARG(int,0)));
    auto invalidReport=locate(); assert(invalidReport && !invalidReport->findChild<QTabWidget*>("access-report-view-tabs")->isTabEnabled(0));
    QFile protectedPdf(output.filePath("owned.pdf")); assert(protectedPdf.open(QIODevice::ReadOnly));
    const auto previousPdf=protectedPdf.readAll(); protectedPdf.close();
    assert(!invalidReport->exportReportPdf(output.filePath("owned.pdf")));
    assert(protectedPdf.open(QIODevice::ReadOnly) && protectedPdf.readAll()==previousPdf); protectedPdf.close();
    QTimer::singleShot(0,invalidReport,[] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="access-report-page-setup-dialog");
        dialog->findChild<QComboBox*>("report-page-orientation")->setCurrentIndex(1);
        dialog->findChild<QDoubleSpinBox*>("report-page-left")->setValue(25);
        dialog->findChild<QCheckBox*>("report-page-save")->setChecked(true);
        dialog->findChild<QDialogButtonBox*>("report-page-buttons")->button(QDialogButtonBox::Ok)->click();
    });
    invalidReport->findChild<QAction*>("access-report-page-setup")->trigger();
    assert(invalidReport->findChild<QTabWidget*>("access-report-view-tabs")->isTabEnabled(0));
    assert(invalidReport->exportReportPdf(output.filePath("repaired.pdf")));
    auto* edit=window.findChild<QAction*>("native-report-edit"); assert(edit);
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="native-report-designer");
        assert(!AccessFormPreview::projectObjectLoaded(db,AccessDocumentKind::Report,"Owned Native Report"));
        assert(dialog->findChild<QPlainTextEdit*>("native-report-sql")->toPlainText()==source);
        assert(dialog->findChild<QPlainTextEdit*>("native-report-vba")->toPlainText()==module);
        assert(dialog->findChild<QComboBox*>("native-report-page-break")->currentIndex()==3);
        assert(dialog->findChild<QComboBox*>("native-report-header-break")->currentIndex()==2);
        bool previewed=false;
        QTimer::singleShot(0,dialog,[&] {
            auto* preview=dynamic_cast<AccessFormPreview*>(QApplication::activeModalWidget());
            assert(preview && preview->documentKind()==AccessDocumentKind::Report && preview->hasDataSource());
            assert(preview->findChild<QPrintPreviewWidget*>("access-report-pages")->pageCount()==3);
            QTimer::singleShot(0,preview,[] {
                auto* setup=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(setup && setup->objectName()=="access-report-page-setup-dialog");
                setup->findChild<QDoubleSpinBox*>("report-page-left")->setValue(30);
                setup->findChild<QCheckBox*>("report-page-save")->setChecked(true);
                setup->findChild<QDialogButtonBox*>("report-page-buttons")->button(QDialogButtonBox::Ok)->click();
            });
            preview->findChild<QAction*>("access-report-page-setup")->trigger();
            previewed=true; preview->reject();
        });
        dialog->findChild<QPushButton*>("native-report-preview")->click();
        assert(previewed && QApplication::activeModalWidget()==dialog);
        assert(dialog->findChild<QPlainTextEdit*>("native-report-sql")->toPlainText()==source);
        dialog->findChild<QPushButton*>("native-report-save")->click();
    });
    explorer->setCurrentItem(selected); edit->trigger(); dispose();
    assert(sql.exec("SELECT definition FROM main._la_objects WHERE kind='report' AND name='Owned Native Report'") && sql.next());
    const auto finalDesign=QJsonDocument::fromJson(sql.value(0).toByteArray()).object(); sql.finish();
    assert(finalDesign.value("page_setup").toObject().value("margins_mm").toArray()[0].toDouble()==30);
    assert(finalDesign.value("sql").toString()==source && finalDesign.value("vba").toString()==module);
    assert(sql.exec("SELECT COUNT(*) FROM native_report_events WHERE kind='open'") && sql.next() && sql.value(0).toInt()==5); sql.finish();
    assert(sql.exec("SELECT COUNT(*) FROM native_report_events WHERE kind='close'") && sql.next() && sql.value(0).toInt()==5); sql.finish();
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog);
        dialog->findChild<QLineEdit*>("native-report-name")->setText("Owned Summary Report");
        dialog->findChild<QPlainTextEdit*>("native-report-sql")->setPlainText("SELECT id FROM native_report_rows ORDER BY id");
        dialog->findChild<QLineEdit*>("native-report-footer-expression")->setText("=Sum([id])");
        dialog->findChild<QComboBox*>("native-report-footer-break")->setCurrentIndex(1);
        dialog->findChild<QPushButton*>("native-report-save")->click();
    });
    create->trigger(); QApplication::processEvents();
    assert(sql.exec("SELECT definition FROM main._la_objects WHERE kind='report' AND name='Owned Summary Report'") && sql.next());
    const auto summaryDesign=QJsonDocument::fromJson(sql.value(0).toByteArray()).object(); sql.finish();
    assert(summaryDesign.value("footer_expression").toString()=="=Sum([id])" && summaryDesign.value("footer_force_new_page").toInt()==1);
    selected=nullptr;
    for(QTreeWidgetItemIterator item(explorer);*item;++item)
        if((*item)->data(0,Qt::UserRole).toString()=="report" && (*item)->text(0)=="Owned Summary Report") selected=*item;
    assert(selected); explorer->setCurrentItem(selected);
    assert(QMetaObject::invokeMethod(explorer,"itemDoubleClicked",Qt::DirectConnection,Q_ARG(QTreeWidgetItem*,selected),Q_ARG(int,0)));
    QApplication::processEvents();
    auto summaryReport=locate("Owned Summary Report"); assert(summaryReport);
    const auto total=makePreviewControlsObject(*summaryReport)->member("Item",{VbaValue::string("native_report_summary")}).object->member("Value",{});
    assert(total.asNumber()==6);
    assert(summaryReport->findChild<QPrintPreviewWidget*>("access-report-pages")->pageCount()==2);
    assert(summaryReport->exportReportPdf(output.filePath("owned-summary.pdf")));
    assert(summaryReport->close()); dispose();
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog);
        assert(dialog->findChild<QLineEdit*>("native-report-footer-expression")->text()=="=Sum([id])");
        assert(dialog->findChild<QComboBox*>("native-report-footer-break")->currentIndex()==1);
        dialog->reject();
    });
    explorer->setCurrentItem(selected); edit->trigger(); dispose();
}

void exerciseRecordEscape(MainWindow& window,QSqlDatabase db) {
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE command_record_undo(id INTEGER PRIMARY KEY AUTOINCREMENT,value TEXT UNIQUE)"));
    assert(sql.exec("INSERT INTO command_record_undo(value) VALUES('alpha'),('beta')"));
    assert(window.openTableDocument("command_record_undo",0,1)==0);
    auto* tabs=window.findChild<QTabWidget*>(); assert(tabs && tabs->currentWidget());
    linaccess::AccessTableView* view=nullptr;
    for(auto* candidate:tabs->currentWidget()->findChildren<QTableView*>()) if(auto* typed=dynamic_cast<linaccess::AccessTableView*>(candidate)) view=typed;
    assert(view); auto* model=qobject_cast<linaccess::AccessTableModel*>(view->model()); assert(model);
    view->sortByColumn(0,Qt::AscendingOrder); view->setCurrentIndex(model->index(0,1));
    assert(model->setData(model->index(0,1),"first-draft") && model->setData(model->index(1,1),"other-draft"));
    QApplication::setActiveWindow(&window); view->setFocus(); view->edit(view->currentIndex()); QApplication::processEvents();
    auto* editor=qobject_cast<QLineEdit*>(QApplication::focusWidget()); assert(editor); editor->setText("uncommitted-editor");
    QTest::keyClick(editor,Qt::Key_Escape); QApplication::processEvents();
    assert(model->index(0,1).data().toString()=="first-draft" && model->index(1,1).data().toString()=="other-draft");
    view->setFocus(); QTest::keyClick(view,Qt::Key_Escape);
    assert(model->index(0,1).data().toString()=="alpha" && model->index(1,1).data().toString()=="other-draft" && model->isDirty());
    assert(!model->hasPendingRecord(0) && model->hasPendingRecord(1));
    assert(sql.exec("SELECT value FROM command_record_undo WHERE id=2") && sql.next() && sql.value(0).toString()=="beta"); sql.finish();
    assert(view->insertRecord() && model->rowCount()==3 && model->setData(model->index(2,1),"new-draft"));
    view->setFocus(); QTest::keyClick(view,Qt::Key_Escape);
    assert(model->rowCount()==2 && model->index(1,1).data().toString()=="other-draft" && model->hasPendingRecord(1));
    view->setCurrentIndex(model->index(1,1)); QTest::keyClick(view,Qt::Key_Escape); assert(!model->isDirty());
    view->setCurrentIndex(model->index(0,1)); assert(model->setData(model->index(0,1),"saved"));
    assert(window.dispatchTableCommand(97).value("error").toInt()==0 && !model->isDirty());
    view->setFocus(); QTest::keyClick(view,Qt::Key_Escape);
    assert(sql.exec("SELECT value FROM command_record_undo WHERE id=1") && sql.next() && sql.value(0).toString()=="saved"); sql.finish();
    assert(model->setData(model->index(0,1),"saved-by-down"));
    view->setCurrentIndex(model->index(0,1)); QTest::keyClick(view,Qt::Key_Down);
    assert(view->currentIndex().row()==1 && view->currentIndex().column()==1 && !model->isDirty());
    assert(sql.exec("SELECT value FROM command_record_undo WHERE id=1") && sql.next() && sql.value(0).toString()=="saved-by-down"); sql.finish();
    assert(model->setData(model->index(1,1),"saved-by-down"));
    QTest::keyClick(view,Qt::Key_Up);
    assert(view->currentIndex().row()==1 && model->isDirty() && model->lastError().nativeErrorCode()=="3146");
    assert(sql.exec("SELECT value FROM command_record_undo WHERE id=2") && sql.next() && sql.value(0).toString()=="beta"); sql.finish();
    QTest::keyClick(view,Qt::Key_Escape); assert(!model->isDirty());
    QTest::keyClick(view,Qt::Key_Up); assert(view->currentIndex().row()==0 && view->currentIndex().column()==1);
    assert(model->setData(model->index(0,1),"saved-by-tab"));
    view->setFocus(); QTest::keyClick(view,Qt::Key_Tab);
    assert(view->currentIndex().row()==1 && view->currentIndex().column()==0 && !model->isDirty());
    assert(sql.exec("SELECT value FROM command_record_undo WHERE id=1") && sql.next() && sql.value(0).toString()=="saved-by-tab"); sql.finish();
    QTest::keyClick(view,Qt::Key_Backtab,Qt::ShiftModifier);
    assert(view->currentIndex().row()==0 && view->currentIndex().column()==1);
    view->edit(view->currentIndex()); QApplication::processEvents();
    editor=qobject_cast<QLineEdit*>(QApplication::focusWidget()); assert(editor); editor->setText("saved-by-editor-tab");
    QTest::keyClick(editor,Qt::Key_Tab); QApplication::processEvents();
    assert(view->currentIndex().row()==1 && view->currentIndex().column()==0 && !model->isDirty());
    assert(sql.exec("SELECT value FROM command_record_undo WHERE id=1") && sql.next() && sql.value(0).toString()=="saved-by-editor-tab"); sql.finish();
    view->setCurrentIndex(model->index(0,1)); assert(model->setData(model->index(0,1),"beta"));
    view->setFocus(); QTest::keyClick(view,Qt::Key_Tab);
    assert(view->currentIndex().row()==0 && view->currentIndex().column()==1 && model->isDirty() && model->lastError().nativeErrorCode()=="3146");
    assert(sql.exec("SELECT value FROM command_record_undo WHERE id=1") && sql.next() && sql.value(0).toString()=="saved-by-editor-tab"); sql.finish();
    if(auto* live=qobject_cast<QLineEdit*>(QApplication::focusWidget());live && view->isAncestorOf(live)) QTest::keyClick(live,Qt::Key_Escape);
    view->setFocus(); QTest::keyClick(view,Qt::Key_Escape); assert(!model->isDirty());
    assert(model->setData(model->index(0,1),"saved-by-ctrl-end"));
    QTest::keyClick(view,Qt::Key_End,Qt::ControlModifier);
    assert(view->currentIndex().row()==1 && view->currentIndex().column()==1 && !model->isDirty());
    assert(sql.exec("SELECT value FROM command_record_undo WHERE id=1") && sql.next() && sql.value(0).toString()=="saved-by-ctrl-end"); sql.finish();
    assert(model->setData(model->index(1,1),"saved-by-ctrl-home"));
    QTest::keyClick(view,Qt::Key_Home,Qt::ControlModifier);
    assert(view->currentIndex().row()==0 && view->currentIndex().column()==0 && !model->isDirty());
    assert(sql.exec("SELECT value FROM command_record_undo WHERE id=2") && sql.next() && sql.value(0).toString()=="saved-by-ctrl-home"); sql.finish();
    assert(model->setData(model->index(0,1),"same-record-draft"));
    QTest::keyClick(view,Qt::Key_Home,Qt::ControlModifier); assert(model->isDirty() && view->currentIndex().row()==0);
    assert(sql.exec("SELECT value FROM command_record_undo WHERE id=1") && sql.next() && sql.value(0).toString()=="saved-by-ctrl-end"); sql.finish();
    QTest::keyClick(view,Qt::Key_Escape); QTest::keyClick(view,Qt::Key_End,Qt::ControlModifier);
    assert(model->setData(model->index(1,1),"saved-by-ctrl-end")); QTest::keyClick(view,Qt::Key_Home,Qt::ControlModifier);
    assert(view->currentIndex().row()==1 && model->isDirty() && model->lastError().nativeErrorCode()=="3146");
    QTest::keyClick(view,Qt::Key_Escape); assert(!model->isDirty());
    assert(window.closeTableDocument(0,"command_record_undo",2).value("error").toInt()==0); dispose();
}
void exercisePageNavigation(MainWindow& window,QSqlDatabase db) {
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE command_page_navigation(id INTEGER PRIMARY KEY,value TEXT UNIQUE)"));
    for(int row=1;row<=60;++row) {
        sql.prepare("INSERT INTO command_page_navigation VALUES(?,?)"); sql.addBindValue(row); sql.addBindValue(QString("page-row-%1").arg(row)); assert(sql.exec());
    }
    assert(window.openTableDocument("command_page_navigation",0,1)==0);
    auto* tabs=window.findChild<QTabWidget*>(); assert(tabs && tabs->currentWidget());
    linaccess::AccessTableView* view=nullptr;
    for(auto* candidate:tabs->currentWidget()->findChildren<QTableView*>()) if(auto* typed=dynamic_cast<linaccess::AccessTableView*>(candidate)) view=typed;
    assert(view); auto* model=qobject_cast<linaccess::AccessTableModel*>(view->model()); assert(model);
    window.resize(1100,700); QApplication::setActiveWindow(&window); QApplication::processEvents();
    view->sortByColumn(0,Qt::AscendingOrder); view->setCurrentIndex(model->index(0,1)); view->setFocus();
    assert(model->setData(model->index(0,1),"saved-by-page-down")); QTest::keyClick(view,Qt::Key_PageDown);
    const int destination=view->currentIndex().row(); assert(destination>1 && view->currentIndex().column()==1 && !model->isDirty());
    assert(sql.exec("SELECT value FROM command_page_navigation WHERE id=1") && sql.next() && sql.value(0).toString()=="saved-by-page-down"); sql.finish();
    const int destinationId=model->index(destination,0).data().toInt(); const auto original=model->index(destination,1).data().toString();
    assert(model->setData(model->index(destination,1),"saved-by-page-down")); QTest::keyClick(view,Qt::Key_PageUp);
    assert(view->currentIndex().row()==destination && model->isDirty() && model->lastError().nativeErrorCode()=="3146");
    sql.prepare("SELECT value FROM command_page_navigation WHERE id=?"); sql.addBindValue(destinationId);
    assert(sql.exec() && sql.next() && sql.value(0).toString()==original); sql.finish();
    QTest::keyClick(view,Qt::Key_Escape); QTest::keyClick(view,Qt::Key_PageUp);
    assert(!model->isDirty() && view->currentIndex().row()<destination && view->currentIndex().column()==1);
    view->setCurrentIndex(model->index(0,1)); assert(model->setData(model->index(0,1),"same-page-draft")); QTest::keyClick(view,Qt::Key_PageUp);
    assert(view->currentIndex().row()==0 && model->isDirty());
    assert(sql.exec("SELECT value FROM command_page_navigation WHERE id=1") && sql.next() && sql.value(0).toString()=="saved-by-page-down"); sql.finish();
    QTest::keyClick(view,Qt::Key_Escape); assert(!model->isDirty());
    assert(window.closeTableDocument(0,"command_page_navigation",2).value("error").toInt()==0); dispose();
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
    invalidDocumentEdit->deselect();
    assert(window.dispatchTableCommand(190).value("error").toInt()==2046 && QApplication::clipboard()->text()=="invalid-editor-sentinel");
    invalidDocumentEdit->selectAll();
    assert(window.dispatchTableCommand(190).value("error").toInt()==0 && QApplication::clipboard()->text()=="invalid number");
    QApplication::clipboard()->setText("invalid-editor-sentinel");
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
Sub OpenEditableQuery()
DoCmd.OpenQuery "command_edit_query"
Me.Tag="query-edit-opened"
End Sub
Sub EditableQueryAsReadonly()
DoCmd.OpenQuery "command_edit_query",,acReadOnly
End Sub
Sub CloseEditableQuery()
DoCmd.Close acQuery,"command_edit_query",acSaveNo
End Sub
Sub OpenReadQuery()
DoCmd.OpenQuery QueryName:="command_query_doc",DataMode:=acReadOnly
Me.Tag="query-opened"
End Sub
Sub QueryLastRecord()
DoCmd.GoToRecord , , acLast
Me.Tag="query-last"
End Sub
Sub QueryNamedRecord()
DoCmd.GoToRecord ObjectType:=acDataQuery,ObjectName:="COMMAND_QUERY_DOC",Record:=acGoTo,Offset:=2
Me.Tag="query-named"
End Sub
Sub QueryAbsentRecord()
DoCmd.GoToRecord acDataQuery,"absent_query",acFirst
End Sub
Sub QueryCopyRecord()
DoCmd.RunCommand acCmdSelectRecord
DoCmd.RunCommand acCmdCopy
End Sub
Sub QuerySelectAll()
DoCmd.RunCommand acCmdSelectAllRecords
End Sub
Sub QueryCopyByKeys()
DoCmd.SendKeys "^c",True
Me.Tag="query-keys"
End Sub
Sub QueryZoom()
DoCmd.RunCommand acCmdZoomBox
Me.Tag="query-zoomed"
End Sub
Sub QueryFindNumber()
DoCmd.FindRecord 1.25
Me.Tag="query-found"
End Sub
Sub QueryFindText()
DoCmd.FindRecord "needle",acStart
Me.Tag="query-text-found"
End Sub
Sub QueryFindCurrency()
DoCmd.FindRecord CCur("1.2346")
End Sub
Sub QueryFirstCommand()
DoCmd.RunCommand acCmdRecordsGoToFirst
End Sub
Sub QueryNextRecord()
DoCmd.GoToRecord
End Sub
Sub QueryNewRecord()
DoCmd.GoToRecord , , acNewRec
End Sub
Sub SelectReadQuery()
DoCmd.SelectObject acQuery,"COMMAND_QUERY_DOC",False
Me.Tag="query-selected"
End Sub
Sub SelectParameterQuery()
DoCmd.SelectObject acQuery,"command_parameter_doc",False
End Sub
Sub SelectAbsentQuery()
DoCmd.SelectObject acQuery,"absent_query",False
Me.Tag="unreachable-query"
End Sub
Sub OpenParameterQuery()
DoCmd.OpenQuery "command_parameter_doc",,acReadOnly
Me.Tag="parameter-opened"
End Sub
Sub OpenActionQuery()
DoCmd.OpenQuery "command_action_doc",,acReadOnly
Me.Tag="action-opened"
End Sub
Sub UnsupportedQueryEdit()
DoCmd.OpenQuery "command_query_doc"
Me.Tag="edit-opened"
End Sub
Sub MissingQuery()
DoCmd.OpenQuery "absent_query",,acReadOnly
End Sub
Sub NullQuery()
DoCmd.OpenQuery Null,,acReadOnly
End Sub
Sub CloseReadQuery()
DoCmd.Close acQuery,"command_query_doc",acSaveNo
Me.Tag="query-closed"
End Sub
Sub CloseActiveQuery()
DoCmd.Close
Me.Tag="active-query-closed"
End Sub
Sub CloseParameterQuery()
DoCmd.Close acQuery,"command_parameter_doc"
End Sub
Sub OpenReadonlyTable()
DoCmd.OpenTable TableName:="command_readonly",DataMode:=acReadOnly
Me.Tag="readonly-opened"
End Sub
Sub OpenEditableTable()
DoCmd.OpenTable "command_edit"
Me.Tag="edit-opened"
End Sub
Sub FindTableByKeys()
SendKeys "^f^ainitial{ENTER}",True
End Sub
Sub ZoomTableByKeys()
SendKeys "+{F2}{ESC}",True
End Sub
Sub CloseTableFindByKeys()
DoCmd.SendKeys "{ESC}",True
End Sub
Sub SendTableKeys()
SendKeys "{F2}^akey-typed{TAB}",True
Me.Tag="table-keys"
End Sub
Sub SendTableKeysDoCmd()
DoCmd.SendKeys Keys:="{F2}^amethod-keys{TAB}",Wait:=True
Me.Tag="docmd-table-keys"
End Sub
Sub QueueTableKeys()
DoCmd.SendKeys Keys:="{F2}^aqueued-keys{TAB}"
Me.Tag="queued-table-keys"
End Sub
Sub RejectTableKeySequence()
DoCmd.SendKeys "^a{UNSUPPORTED_KEY}",True
Me.Tag="must-not-run"
End Sub
Sub RejectNullTableKeys()
DoCmd.SendKeys Null,True
Me.Tag="must-not-run"
End Sub
Sub SaveActiveTable()
DoCmd.RunCommand acCmdSaveRecord
Me.Tag="table-saved"
End Sub
Sub SaveActiveTableViaApplication()
Application.RunCommand Command:=acCmdSaveRecord
Me.Tag="application-table-saved"
End Sub
Sub OpenTypedFindTable()
DoCmd.OpenTable "command_typed_find",DataMode:=acReadOnly
End Sub
Sub OpenTypedPasteTable()
DoCmd.OpenTable "command_typed_find",DataMode:=acEdit
End Sub
Sub CloseTypedFindTable()
DoCmd.Close acTable,"command_typed_find"
End Sub
Sub FindTypedNumber()
DoCmd.FindRecord 1.2345678901234567
End Sub
Sub FindTypedAdjacent()
DoCmd.FindRecord 1.2345678901234565
End Sub
Sub FindTypedNumberAbsent()
DoCmd.FindRecord 1.2345678901234563
End Sub
Sub FindTypedMoney()
DoCmd.FindRecord CCur("0.1001")
End Sub
Sub FindTypedMoneySmall()
DoCmd.FindRecord CCur("-0.0001")
End Sub
Sub FindTypedStamp()
DoCmd.FindRecord #2020-02-29 12:34:56#
End Sub
Sub FindTypedNegativeStamp()
DoCmd.FindRecord CDate(-0.25)
End Sub
Sub FindTypedFalse()
DoCmd.FindRecord False
End Sub
Sub FindTypedTrue()
DoCmd.FindRecord True
End Sub
Sub FindTypedPartial()
DoCmd.FindRecord 42,acAnywhere
End Sub
Sub FindTypedSubsecond()
DoCmd.FindRecord CDate(0.250001)
End Sub
Sub FindRecordDefault()
DoCmd.FindRecord FindWhat:="ROW-250",MatchCase:=False
Me.Tag="findrecord-default"
End Sub
Sub FindRecordNext()
DoCmd.FindRecord "row-2",acStart,,,,,False
End Sub
Sub FindRecordEmptyFirst()
DoCmd.FindRecord "row-2",acStart,,,,,Empty
End Sub
Sub FindRecordEmptyMatch()
DoCmd.FindRecord "ow-25",Empty
End Sub
Sub FindRecordNoMatch()
DoCmd.FindRecord "not-present"
Me.Tag="findrecord-nohit"
End Sub
Sub FindRecordReset()
DoCmd.FindRecord "row-280"
End Sub
Sub FindRecordAllFields()
DoCmd.FindRecord "row-250",OnlyCurrentField:=acAll
Me.Tag="must-not-run"
End Sub
Sub FindRecordFormatted()
DoCmd.FindRecord "row-250",SearchAsFormatted:=True
Me.Tag="must-not-run"
End Sub
Sub FindRecordUp()
DoCmd.FindRecord "row-249",Search:=acUp,FindFirst:=False
Me.Tag="findrecord-up"
End Sub
Sub FindRecordUpNoWrap()
DoCmd.FindRecord "row-280",Search:=acUp,FindFirst:=False
End Sub
Sub FindRecordDown()
DoCmd.FindRecord "row-280",Search:=acDown,FindFirst:=False
End Sub
Sub FindRecordDownNoWrap()
DoCmd.FindRecord "row-249",Search:=acDown,FindFirst:=False
End Sub
Sub FindRecordInvalidDirection()
DoCmd.FindRecord "row-250",Search:=7
Me.Tag="must-not-run"
End Sub
Sub FindRecordNull()
DoCmd.FindRecord Null
Me.Tag="must-not-run"
End Sub
Sub FindRecordNumber()
DoCmd.FindRecord 42
Me.Tag="must-not-run"
End Sub
Sub FindRecordWildcard()
DoCmd.FindRecord "row-*"
Me.Tag="must-not-run"
End Sub
Sub FindTableText()
DoCmd.RunCommand acCmdFind
End Sub
Sub RepeatTableFind()
Application.RunCommand Command:=acCmdFindNext
End Sub
Sub SortTableAscending()
DoCmd.RunCommand acCmdSortAscending
End Sub
Sub SortTableDescending()
Application.RunCommand Command:=acCmdSortDescending
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
Sub SelectAllTableRecords()
Application.RunCommand Command:=acCmdSelectAllRecords
End Sub
Sub CopyTableSelection()
Application.RunCommand Command:=acCmdCopy
End Sub
Sub CutTableSelection()
DoCmd.RunCommand acCmdCut
End Sub
Sub UndoTableText()
DoCmd.RunCommand acCmdUndo
End Sub
Sub RedoTableText()
Application.RunCommand acCmdRedo
End Sub
Sub PasteTableSelection()
DoCmd.RunCommand acCmdPaste
End Sub
Sub RequeryActiveTable()
DoCmd.Requery
End Sub
Sub OpenRecordEditTable()
DoCmd.OpenTable "command_record_edits"
End Sub
Sub AddTableRecord()
DoCmd.RunCommand acCmdRecordsGoToNew
End Sub
Sub AddNamedTableRecord()
DoCmd.GoToRecord ObjectType:=acTable,ObjectName:="command_record_edits",Record:=acNewRec
End Sub
Sub AddActiveTableRecord()
DoCmd.GoToRecord , , acNewRec
End Sub
Sub AddNavigationTableRecord()
DoCmd.GoToRecord acTable,"command_navigation",acNewRec
End Sub
Sub DeleteTableRecord()
DoCmd.RunCommand acCmdDeleteRecord
End Sub
Sub CloseRecordEditTable()
DoCmd.Close acTable,"command_record_edits"
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
    assert(unchanged.exec(R"SQL(INSERT INTO _la_objects(kind,name,definition) VALUES('query','command_query_doc','{"sql":"SELECT id,value,flag FROM command_readonly ORDER BY id"}'))SQL"));
    assert(unchanged.exec(R"SQL(INSERT INTO _la_objects(kind,name,definition) VALUES('query','command_edit_query','{"sql":"SELECT * FROM command_edit"}'))SQL"));
    artifact(db,"query","command_parameter_doc","PARAMETERS p Currency; SELECT p AS amount");
    artifact(db,"query","command_action_doc","DELETE FROM command_readonly");
    const int queryCallerId=artifact(db,"form","query_lifetime_caller","Begin Form\nEnd\n");
    artifact(db,"vba","Form_query_lifetime_caller","Sub OpenAndClose()\nDoCmd.OpenQuery \"command_query_doc\",,acReadOnly\nDoCmd.Close acForm,\"query_lifetime_caller\"\nEnd Sub\nSub OpenEditableAndClose()\nDoCmd.OpenQuery \"command_edit_query\"\nDoCmd.Close acForm,\"query_lifetime_caller\"\nEnd Sub\n");
    assert(unchanged.exec("CREATE TABLE command_typed_find(id INTEGER PRIMARY KEY,value REAL,money CURRENCY_REAL,stamp DATETIME,flag BOOLEAN,payload BLOB)"));
    assert(unchanged.exec("INSERT INTO command_typed_find VALUES(1,1.2345678901234567,0.1001,'2020-02-29 12:34:56',-1,x'0102'),(2,1.2345678901234565,0.1,'1899-12-30 06:00:00',NULL,x'0102'),(3,42,-0.0001,'2020-02-29 12:34:57',0,x'0102'),(4,NULL,NULL,NULL,0,x'0102')"));
    assert(unchanged.exec("CREATE TABLE command_record_edits(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(unchanged.exec("INSERT INTO command_record_edits VALUES(1,'kept')"));
    assert(unchanged.exec("CREATE TABLE command_navigation(id INTEGER PRIMARY KEY,value TEXT CHECK(value!='blocked'),payload BLOB)"));
    assert(unchanged.exec("WITH RECURSIVE n(x) AS (VALUES(1) UNION ALL SELECT x+1 FROM n WHERE x<300) INSERT INTO command_navigation SELECT x,'row-'||x,x'0102' FROM n"));
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
    const int beforeQueries=documentTabs->count();
    const auto callerBeforeQueries=mixedEditor->text();
    assert(mixed->executeProcedure("OpenEditableQuery",selectArguments).status==linaccess::VbaExecutionStatus::Completed && documentTabs->count()==beforeQueries+1);
    QPointer<QWidget> editableQueryPage=documentTabs->currentWidget();
    auto* editableQueryView=dynamic_cast<linaccess::AccessTableView*>(editableQueryPage->findChild<QTableView*>("native-editable-query-view")); assert(editableQueryView);
    auto* editableQueryModel=qobject_cast<linaccess::AccessTableModel*>(editableQueryView->model()); assert(editableQueryModel && !editableQueryModel->isReadOnly());
    const int queryValueColumn=editableQueryModel->fieldIndex("value"); assert(queryValueColumn>=0 && editableQueryView->currentIndex().isValid());
    auto editableDescriptor=makePreviewApplicationObject(*mixed)->member("CurrentData",{}).object->member("AllQueries",{}).object->member("Item",{linaccess::VbaValue::string("command_edit_query")}).object;
    auto sourceTableDescriptor=makePreviewApplicationObject(*mixed)->member("CurrentData",{}).object->member("AllTables",{}).object->member("Item",{linaccess::VbaValue::string("command_edit")}).object;
    assert(editableDescriptor->member("IsLoaded",{}).boolean && !sourceTableDescriptor->member("IsLoaded",{}).boolean && window.activateOpenTable("command_edit")==3251);
    assert(mixed->executeProcedure("EditableQueryAsReadonly",selectArguments).errorNumber==3251 && documentTabs->currentWidget()==editableQueryPage);
    assert(window.openTableDocument("command_edit",0,1)==3251);
    assert(editableQueryModel->setData(editableQueryModel->index(0,queryValueColumn),"query-updated"));
    assert(unchanged.exec("SELECT value FROM command_edit WHERE id=1") && unchanged.next() && unchanged.value(0).toString()=="initial"); unchanged.finish();
    QSqlQuery sourceShadow(editableQueryModel->database());
    assert(sourceShadow.exec("CREATE TEMP TABLE command_edit(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(mixed->executeProcedure("SaveActiveTable",selectArguments).errorNumber==3251 && editableQueryModel->isDirty());
    assert(mixed->executeProcedure("RequeryActiveTable",selectArguments).errorNumber==3251 && editableQueryModel->isDirty());
    assert(sourceShadow.exec("DROP TABLE temp.command_edit"));
    assert(mixed->executeProcedure("SaveActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed && !editableQueryModel->isDirty());
    assert(unchanged.exec("SELECT value FROM command_edit WHERE id=1") && unchanged.next() && unchanged.value(0).toString()=="query-updated"); unchanged.finish();
    assert(unchanged.exec("UPDATE command_edit SET value='query-refreshed' WHERE id=1"));
    assert(mixed->executeProcedure("RequeryActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed && editableQueryModel->index(0,queryValueColumn).data().toString()=="query-refreshed");
    assert(unchanged.exec(R"SQL(UPDATE _la_objects SET definition='{"sql":"SELECT * FROM command_edit WHERE id=1"}' WHERE kind='query' AND name='command_edit_query')SQL"));
    assert(mixed->executeProcedure("RequeryActiveTable",selectArguments).errorNumber==3251 && documentTabs->currentWidget()==editableQueryPage);
    assert(unchanged.exec(R"SQL(UPDATE _la_objects SET definition='{"sql":"SELECT * FROM command_edit"}' WHERE kind='query' AND name='command_edit_query')SQL"));
    assert(editableQueryModel->setData(editableQueryModel->index(0,queryValueColumn),"blocked"));
    assert(mixed->executeProcedure("CloseEditableQuery",selectArguments).status==linaccess::VbaExecutionStatus::Error && editableDescriptor->member("IsLoaded",{}).boolean && editableQueryModel->isDirty());
    assert(makePreviewApplicationObject(*mixed)->member("SysCmd",{linaccess::VbaValue::integer(10),linaccess::VbaValue::integer(1),linaccess::VbaValue::string("command_edit_query")}).asNumber()==1);
    assert(unchanged.exec("SELECT value FROM command_edit WHERE id=1") && unchanged.next() && unchanged.value(0).toString()=="query-refreshed"); unchanged.finish();
    editableQueryModel->revertAll(); assert(editableQueryModel->setData(editableQueryModel->index(0,queryValueColumn),"query-closed"));
    assert(mixed->executeProcedure("CloseEditableQuery",selectArguments).status==linaccess::VbaExecutionStatus::Completed && !editableDescriptor->member("IsLoaded",{}).boolean);
    QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete); assert(!editableQueryPage && mixed->isVisible() && mixedEditor->text()==callerBeforeQueries);
    assert(unchanged.exec("SELECT value FROM command_edit WHERE id=1") && unchanged.next() && unchanged.value(0).toString()=="query-closed"); unchanged.finish();
    assert(unchanged.exec("INSERT INTO command_edit VALUES(2,'filtered-out')"));
    assert(unchanged.exec(R"SQL(INSERT INTO _la_objects(kind,name,definition) VALUES('query','command_filtered_query','{"sql":"SELECT * FROM command_edit WHERE value <> ''filtered-out''"}'))SQL"));
    assert(window.openQueryDocument("command_filtered_query",0,1).value("error").toInt()==0);
    auto* filteredView=documentTabs->currentWidget()->findChild<QTableView*>("native-editable-query-view"); assert(filteredView);
    auto* filteredModel=qobject_cast<linaccess::AccessTableModel*>(filteredView->model()); assert(filteredModel && filteredModel->rowCount()==1 && filteredModel->index(0,0).data().toInt()==1);
    assert(filteredModel->setData(filteredModel->index(0,filteredModel->fieldIndex("value")),"filtered-out"));
    assert(mixed->executeProcedure("SaveActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed && filteredModel->rowCount()==0 && !filteredModel->isDirty());
    assert(unchanged.exec("SELECT value FROM command_edit WHERE id=1") && unchanged.next() && unchanged.value(0).toString()=="filtered-out"); unchanged.finish();
    assert(window.closeQueryDocument(1,"command_filtered_query",0).value("error").toInt()==0);
    QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete);
    assert(unchanged.exec("DELETE FROM command_edit WHERE id=2"));
    assert(unchanged.exec("UPDATE command_edit SET value='initial' WHERE id=1"));
    assert(mixed->executeProcedure("OpenReadQuery",selectArguments).status==linaccess::VbaExecutionStatus::Completed && documentTabs->count()==beforeQueries+1);
    QPointer<QWidget> queryPage=documentTabs->currentWidget();
    auto* queryView=queryPage->findChild<QTableView*>("native-query-view"); assert(queryView && dynamic_cast<linaccess::AccessQueryModel*>(queryView->model()));
    assert(queryView->model()->rowCount()==2 && queryView->model()->index(0,1).data().toDouble()==1.25 && !queryView->model()->index(0,1).flags().testFlag(Qt::ItemIsEditable));
    assert(mixedEditor->text()==callerBeforeQueries && makePreviewFormObject(*mixed)->member("Tag",{}).asText()=="query-opened");
    queryView->setCurrentIndex(queryView->model()->index(0,1));
    assert(mixed->executeProcedure("QueryLastRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed && queryView->currentIndex().row()==1 && queryView->currentIndex().column()==1);
    assert(mixed->executeProcedure("QueryNextRecord",selectArguments).errorNumber==2105 && queryView->currentIndex().row()==1);
    assert(mixed->executeProcedure("QueryFirstCommand",selectArguments).status==linaccess::VbaExecutionStatus::Completed && queryView->currentIndex().row()==0);
    assert(mixed->executeProcedure("QueryNextRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed && queryView->currentIndex().row()==1);
    assert(mixed->executeProcedure("QueryNewRecord",selectArguments).errorNumber==2448 && queryView->model()->rowCount()==2);
    queryView->setCurrentIndex(queryView->model()->index(0,1));
    QApplication::setActiveWindow(&window); queryView->setFocus(); QApplication::processEvents();
    QApplication::clipboard()->setText("query-copy-sentinel"); QTest::keyClick(queryView,Qt::Key_C,Qt::ControlModifier);
    assert(QApplication::clipboard()->text()=="1.25");
    QApplication::clipboard()->setText("query-copy-sentinel"); QTest::keyClick(queryView,Qt::Key_Insert,Qt::ControlModifier);
    assert(QApplication::clipboard()->text()=="1.25");
    queryView->setColumnHidden(1,true); QApplication::clipboard()->setText("query-copy-sentinel");
    assert(window.dispatchTableCommand(190).value("error").toInt()==3251 && QApplication::clipboard()->text()=="query-copy-sentinel");
    queryView->setColumnHidden(1,false);
    assert(mixed->executeProcedure("QueryCopyByKeys",selectArguments).status==linaccess::VbaExecutionStatus::Completed && QApplication::clipboard()->text()=="1.25");
    assert(makePreviewFormObject(*mixed)->member("Tag",{}).asText()=="query-keys");
    assert(mixed->executeProcedure("QueryCopyRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(queryView->selectionModel()->selectedIndexes().size()==3 && QApplication::clipboard()->text().startsWith("1\t1.25\t"));
    assert(mixed->executeProcedure("QuerySelectAll",selectArguments).status==linaccess::VbaExecutionStatus::Completed && queryView->selectionModel()->selectedIndexes().size()==6);
    queryView->setCurrentIndex(queryView->model()->index(0,1));


    assert(window.navigateTableDocument({},2,1,true).value("error").toInt()==3251);
    queryView->selectionModel()->select(queryView->currentIndex(),QItemSelectionModel::ClearAndSelect);
    auto* queryEditMenu=window.findChild<QMenu*>("native-edit-menu"); assert(queryEditMenu);
    assert(QMetaObject::invokeMethod(queryEditMenu,"aboutToShow",Qt::DirectConnection));
    auto* queryMenuCopy=queryEditMenu->findChild<QAction*>("edit-copy"); auto* queryMenuAll=queryEditMenu->findChild<QAction*>("edit-select-all");
    assert(queryMenuCopy && queryMenuCopy->isEnabled() && queryMenuAll && queryMenuAll->isEnabled());
    assert(!queryEditMenu->findChild<QAction*>("edit-paste")->isEnabled() && !queryEditMenu->findChild<QAction*>("edit-cut")->isEnabled());
    QApplication::clipboard()->setText("query-menu-sentinel"); queryMenuCopy->trigger(); assert(QApplication::clipboard()->text()=="1.25");
    queryMenuAll->trigger(); assert(queryView->selectionModel()->selectedIndexes().size()==6);
    queryView->setCurrentIndex(queryView->model()->index(0,1));
    bool queryZoomSeen=false;
    const auto closeQueryZoom=[&] {
        auto* zoom=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(zoom && zoom->objectName()=="native-cell-zoom");
        auto* text=zoom->findChild<QPlainTextEdit*>("native-cell-zoom-text"); assert(text && text->isReadOnly() && text->toPlainText()=="1.25");
        queryZoomSeen=true; zoom->reject();
    };
    QTimer::singleShot(0,&window,closeQueryZoom);
    assert(mixed->executeProcedure("QueryZoom",selectArguments).status==linaccess::VbaExecutionStatus::Completed && queryZoomSeen);
    assert(makePreviewFormObject(*mixed)->member("Tag",{}).asText()=="query-zoomed" && queryView->currentIndex().data().toDouble()==1.25);
    queryZoomSeen=false; QApplication::setActiveWindow(&window); queryView->setFocus(); QApplication::processEvents();
    QTimer::singleShot(0,&window,closeQueryZoom); QTest::keyClick(queryView,Qt::Key_F2,Qt::ShiftModifier); assert(queryZoomSeen && queryView->currentIndex().data().toDouble()==1.25);

    queryView->setCurrentIndex(queryView->model()->index(1,1));
    assert(mixed->executeProcedure("QueryFindNumber",selectArguments).status==linaccess::VbaExecutionStatus::Completed && queryView->currentIndex().row()==0);
    QApplication::setActiveWindow(&window); queryView->setFocus(); QApplication::processEvents();
    QTest::keyClick(queryView,Qt::Key_F,Qt::ControlModifier); QApplication::processEvents();
    QPointer<QDialog> queryFind=queryView->findChild<QDialog*>("native-table-find-dialog"); assert(queryFind && queryFind->isVisible());
    auto* queryFindText=queryFind->findChild<QLineEdit*>("native-table-find-text"); assert(queryFindText);
    queryFindText->setText("invalid-number"); QTest::mouseClick(queryFind->findChild<QPushButton*>("native-table-find-next"),Qt::LeftButton);
    assert(queryView->currentIndex().row()==0 && queryFind->findChild<QLabel*>("native-table-find-result")->text().contains("13"));
    queryFindText->setFocus(); queryFindText->selectAll();
    assert(window.dispatchTableCommand(190).value("error").toInt()==0 && QApplication::clipboard()->text()=="invalid-number");
    QApplication::clipboard()->setText("1.25"); assert(window.dispatchTableCommand(191).value("error").toInt()==0 && queryFindText->text()=="1.25");
    assert(window.dispatchTableCommand(292).value("error").toInt()==0 && queryFindText->text()=="invalid-number");
    assert(window.dispatchTableCommand(199).value("error").toInt()==0 && queryFindText->text()=="1.25");
    QTest::mouseClick(queryFind->findChild<QPushButton*>("native-table-find-next"),Qt::LeftButton);
    assert(queryView->currentIndex().row()==0 && queryFind->findChild<QLabel*>("native-table-find-result")->text().contains("Datensatz 1"));
    queryFind->close(); QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete); assert(!queryFind);
    QApplication::setActiveWindow(&window); queryView->setFocus(); QApplication::processEvents(); QTest::keyClick(queryView,Qt::Key_F4,Qt::ShiftModifier);
    assert(queryView->currentIndex().row()==0);


    auto queryDescriptor=makePreviewApplicationObject(*mixed)->member("CurrentData",{}).object->member("AllQueries",{}).object->member("Item",{linaccess::VbaValue::string("command_query_doc")}).object;
    assert(queryDescriptor->member("IsLoaded",{}).boolean);
    assert(mixed->executeProcedure("OpenReadQuery",selectArguments).status==linaccess::VbaExecutionStatus::Completed && documentTabs->currentWidget()==queryPage && documentTabs->count()==beforeQueries+1);
    assert(unchanged.exec("UPDATE command_readonly SET value=2.5 WHERE id=1"));
    assert(queryView->model()->index(0,1).data().toDouble()==1.25);
    assert(mixed->executeProcedure("RequeryActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed && queryView->model()->index(0,1).data().toDouble()==2.5);
    assert(unchanged.exec("UPDATE command_readonly SET value=1.25 WHERE id=1"));
    QApplication::setActiveWindow(&window); queryView->setFocus(); QApplication::processEvents(); QTest::keyClick(queryView,Qt::Key_F9,Qt::ShiftModifier);
    assert(queryView->model()->index(0,1).data().toDouble()==1.25);
    QPointer<QAbstractItemModel> oldQueryResult=queryView->model();
    assert(unchanged.exec(R"SQL(UPDATE _la_objects SET definition='{"sql":"SELECT unavailable_column FROM command_readonly"}' WHERE kind='query' AND name='command_query_doc')SQL"));
    assert(mixed->executeProcedure("RequeryActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Error);
    assert(queryView->model()==oldQueryResult && queryView->model()->index(0,1).data().toDouble()==1.25 && queryDescriptor->member("IsLoaded",{}).boolean);
    assert(unchanged.exec(R"SQL(UPDATE _la_objects SET definition='{"sql":"SELECT id FROM command_readonly ORDER BY id DESC"}' WHERE kind='query' AND name='command_query_doc')SQL"));
    assert(mixed->executeProcedure("RequeryActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed && queryView->model()->columnCount()==1 && queryView->model()->index(0,0).data().toInt()==2);
    assert(unchanged.exec(R"SQL(UPDATE _la_objects SET definition='{"sql":"SELECT id,value,flag FROM command_readonly ORDER BY id"}' WHERE kind='query' AND name='command_query_doc')SQL"));
    assert(mixed->executeProcedure("RequeryActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed && queryView->model()->columnCount()==3);
    assert(unchanged.exec("INSERT INTO command_edit VALUES(2,'needle-value')"));
    assert(unchanged.exec(R"SQL(UPDATE _la_objects SET definition='{"sql":"SELECT value FROM command_edit ORDER BY id"}' WHERE kind='query' AND name='command_query_doc')SQL"));
    assert(mixed->executeProcedure("RequeryActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(mixed->executeProcedure("QueryFindText",selectArguments).status==linaccess::VbaExecutionStatus::Completed && queryView->currentIndex().row()==1 && queryView->currentIndex().data().toString()=="needle-value");
    QApplication::setActiveWindow(&window); queryView->setFocus(); QApplication::processEvents(); QTest::keyClick(queryView,Qt::Key_F,Qt::ControlModifier); QApplication::processEvents();
    QPointer<QDialog> refreshedFind=queryView->findChild<QDialog*>("native-table-find-dialog"); assert(refreshedFind && refreshedFind->isVisible());
    refreshedFind->reject(); QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete);
    assert(unchanged.exec("DELETE FROM command_edit WHERE id=2"));
    assert(unchanged.exec(R"SQL(UPDATE _la_objects SET definition='{"sql":"SELECT id,value,flag FROM command_readonly ORDER BY id"}' WHERE kind='query' AND name='command_query_doc')SQL"));
    QApplication::setActiveWindow(&window); queryView->setFocus(); QApplication::processEvents();
    assert(mixed->executeProcedure("RequeryActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed && queryView->model()->columnCount()==3);
    // Execute against the actual driver connection; the callback closes the old
    // page while its replacement is being materialized, exercising Qt lifetime.
    QLibrary sqliteLibrary("sqlite3",0); assert(sqliteLibrary.load());
    using Scalar=void (*)(sqlite3_context*,int,sqlite3_value**);
    using Create=int (*)(sqlite3*,const char*,int,int,void*,Scalar,Scalar,void (*)(sqlite3_context*),void (*)(void*));
    const auto createScalar=reinterpret_cast<Create>(sqliteLibrary.resolve("sqlite3_create_function_v2"));
    static auto callbackData=reinterpret_cast<void* (*)(sqlite3_context*)>(sqliteLibrary.resolve("sqlite3_user_data"));
    static auto callbackResult=reinterpret_cast<void (*)(sqlite3_context*,int)>(sqliteLibrary.resolve("sqlite3_result_int"));
    assert(createScalar && callbackData && callbackResult);
    struct QueryCloseState { MainWindow* window; QString connection; bool called{}; } closeState{&window,db.connectionName()};
    auto nativeHandle=db.driver()->handle(); assert(QByteArray(nativeHandle.typeName())=="sqlite3*");
    auto* sqliteConnection=*static_cast<sqlite3**>(nativeHandle.data());
    assert(createScalar(sqliteConnection,"la_test_close_query",0,1,&closeState,[](sqlite3_context* context,int,sqlite3_value**) {
        auto& state=*static_cast<QueryCloseState*>(callbackData(context)); state.called=true;
        assert(!state.window->allowProjectClose(state.connection));
        assert(state.window->requeryTableDocument().value("error").toInt()==3251);
        assert(state.window->closeQueryDocument(1,"command_query_doc",0).value("error").toInt()==0);
        QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete);
        callbackResult(context,1);
    },nullptr,nullptr,nullptr)==0);
    assert(unchanged.exec(R"SQL(UPDATE _la_objects SET definition='{"sql":"SELECT la_test_close_query() AS result"}' WHERE kind='query' AND name='command_query_doc')SQL"));
    assert(mixed->executeProcedure("RequeryActiveTable",selectArguments).errorNumber==2501);
    assert(closeState.called && !queryPage && !queryDescriptor->member("IsLoaded",{}).boolean);
    assert(createScalar(sqliteConnection,"la_test_close_query",0,1,nullptr,nullptr,nullptr,nullptr,nullptr)==0);
    assert(unchanged.exec(R"SQL(UPDATE _la_objects SET definition='{"sql":"SELECT id,value,flag FROM command_readonly ORDER BY id"}' WHERE kind='query' AND name='command_query_doc')SQL"));
    assert(mixed->executeProcedure("OpenReadQuery",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    queryPage=documentTabs->currentWidget();
    assert(mixed->executeProcedure("UnsupportedQueryEdit",selectArguments).errorNumber==3251);
    assert(mixed->executeProcedure("MissingQuery",selectArguments).errorNumber==7874 && mixed->executeProcedure("NullQuery",selectArguments).errorNumber==94);
    assert(mixed->executeProcedure("OpenActionQuery",selectArguments).errorNumber==3251);
    assert(unchanged.exec("SELECT count(*) FROM command_readonly") && unchanged.next() && unchanged.value(0).toInt()==2); unchanged.finish();
    assert(mixed->executeProcedure("CloseActiveQuery",selectArguments).status==linaccess::VbaExecutionStatus::Completed && !queryDescriptor->member("IsLoaded",{}).boolean && mixed->isVisible());
    QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete); assert(!queryPage && mixed->isVisible());
    QTimer::singleShot(0,&window,[] { auto* dialog=qobject_cast<QInputDialog*>(QApplication::activeModalWidget()); assert(dialog); dialog->reject(); });
    assert(mixed->executeProcedure("OpenParameterQuery",selectArguments).errorNumber==2501 && documentTabs->count()==beforeQueries);
    QTimer::singleShot(0,&window,[] { auto* dialog=qobject_cast<QInputDialog*>(QApplication::activeModalWidget()); assert(dialog); dialog->setTextValue("1.23455"); dialog->accept(); });
    assert(mixed->executeProcedure("OpenParameterQuery",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    auto* parameterView=documentTabs->currentWidget()->findChild<QTableView*>("native-query-view"); assert(parameterView && parameterView->model()->index(0,0).data().toString()=="1.2346");
    const auto queryCurrencyFind=mixed->executeProcedure("QueryFindCurrency",selectArguments);
    assert(queryCurrencyFind.status==linaccess::VbaExecutionStatus::Completed && parameterView->currentIndex().row()==0);
    QPointer<QAbstractItemModel> oldParameterResult=parameterView->model();
    QTimer::singleShot(0,&window,[] { auto* dialog=qobject_cast<QInputDialog*>(QApplication::activeModalWidget()); assert(dialog); dialog->reject(); });
    assert(mixed->executeProcedure("RequeryActiveTable",selectArguments).errorNumber==2501 && parameterView->model()==oldParameterResult && parameterView->model()->index(0,0).data().toString()=="1.2346");
    QTimer::singleShot(0,&window,[] { auto* dialog=qobject_cast<QInputDialog*>(QApplication::activeModalWidget()); assert(dialog); dialog->setTextValue("2.5"); dialog->accept(); });
    assert(mixed->executeProcedure("RequeryActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed && parameterView->model()->index(0,0).data().toDouble()==2.5);
    const QPointer<QWidget> parameterPage=documentTabs->currentWidget();
    assert(mixed->executeProcedure("OpenReadQuery",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    const QPointer<QWidget> selectQueryPage=documentTabs->currentWidget();
    assert(mixed->executeProcedure("SelectParameterQuery",selectArguments).status==linaccess::VbaExecutionStatus::Completed && documentTabs->currentWidget()==parameterPage);
    const int beforeSelect=documentTabs->count();
    assert(mixed->executeProcedure("QueryNamedRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed && documentTabs->currentWidget()==selectQueryPage);
    assert(selectQueryPage->findChild<QTableView*>("native-query-view")->currentIndex().row()==1);
    assert(mixed->executeProcedure("QueryAbsentRecord",selectArguments).errorNumber==3251 && documentTabs->currentWidget()==selectQueryPage);
    assert(mixed->executeProcedure("SelectParameterQuery",selectArguments).status==linaccess::VbaExecutionStatus::Completed && documentTabs->currentWidget()==parameterPage);

    assert(mixed->executeProcedure("SelectReadQuery",selectArguments).status==linaccess::VbaExecutionStatus::Completed && documentTabs->currentWidget()==selectQueryPage && documentTabs->count()==beforeSelect);
    assert(makePreviewFormObject(*mixed)->member("Tag",{}).asText()=="query-selected");
    assert(mixed->executeProcedure("SelectAbsentQuery",selectArguments).errorNumber==3251 && documentTabs->currentWidget()==selectQueryPage && documentTabs->count()==beforeSelect);
    assert(mixed->executeProcedure("CloseReadQuery",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(mixed->executeProcedure("SelectParameterQuery",selectArguments).status==linaccess::VbaExecutionStatus::Completed && documentTabs->currentWidget()==parameterPage);
    assert(mixed->executeProcedure("CloseParameterQuery",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    openArtifact(window,queryCallerId); QPointer<AccessFormPreview> editQueryCaller=document("query_lifetime_caller"); assert(editQueryCaller);
    assert(editQueryCaller->executeProcedure("OpenEditableAndClose",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete); QApplication::processEvents();
    assert(!editQueryCaller && editableDescriptor->member("IsLoaded",{}).boolean);
    assert(window.closeQueryDocument(1,"command_edit_query",0).value("error").toInt()==0 && !editableDescriptor->member("IsLoaded",{}).boolean);
    openArtifact(window,queryCallerId); QPointer<AccessFormPreview> queryCaller=document("query_lifetime_caller"); assert(queryCaller);
    assert(queryCaller->executeProcedure("OpenAndClose",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete); QApplication::processEvents();
    assert(!queryCaller && queryDescriptor->member("IsLoaded",{}).boolean && documentTabs->count()==beforeQueries+1);
    auto* independentQuery=documentTabs->currentWidget()->findChild<QTableView*>("native-query-view"); assert(independentQuery);
    QApplication::setActiveWindow(&window); independentQuery->setFocus(); QApplication::processEvents();
    QTest::keyClick(independentQuery,Qt::Key_F4,Qt::ControlModifier); QApplication::processEvents();
    assert(!queryDescriptor->member("IsLoaded",{}).boolean && documentTabs->count()==beforeQueries);
    assert(mixed->executeProcedure("OpenReadonlyTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed); QApplication::processEvents();
    const QPointer<QWidget> readonlyPage=documentTabs->currentWidget(); auto* readonlyView=nativeTable(readonlyPage);
    auto* readonlyModel=qobject_cast<linaccess::AccessTableModel*>(readonlyView->model()); assert(readonlyModel && readonlyModel->isReadOnly() && readonlyModel->rowCount()==2);
    const int readonlyRow=readonlyModel->index(0,0).data().toInt()==1 ? 0 : 1;
    auto nativeData=makePreviewApplicationObject(*mixed)->member("CurrentData",{}).object;
    auto readonlyDescriptor=nativeData->member("AllTables",{}).object->member("Item",{linaccess::VbaValue::string("command_readonly")}).object;
    assert(readonlyDescriptor->member("IsLoaded",{}).boolean);
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
    QApplication::setActiveWindow(&window); readonlyView->setFocus(); QApplication::processEvents();
    QApplication::clipboard()->setText("insert-copy-sentinel");
    QTest::keyClick(readonlyView,Qt::Key_Insert,Qt::ControlModifier);
    assert(QApplication::clipboard()->text()=="1\t1.25\tTrue" && !readonlyModel->isDirty());
    assert(mixed->executeProcedure("PasteTableSelection",selectArguments).errorNumber==2448 && !readonlyModel->isDirty());
    QTimer::singleShot(0,[] {
        auto* zoom=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(zoom && zoom->objectName()=="native-cell-zoom");
        auto* editor=zoom->findChild<QPlainTextEdit*>("native-cell-zoom-text"); assert(editor && editor->isReadOnly() && editor->toPlainText()=="1.25"); zoom->reject();
    });
    assert(mixed->executeProcedure("ZoomTableCell",selectArguments).status==linaccess::VbaExecutionStatus::Completed && !readonlyModel->isDirty());
    QApplication::setActiveWindow(&window); readonlyView->setFocus(); QApplication::processEvents();
    readonlyView->setCurrentIndex(readonlyModel->index(readonlyRow,0));
    assert(mixed->executeProcedure("SortTableAscending",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(readonlyModel->isReadOnly() && !readonlyModel->isDirty() && readonlyModel->index(0,0).data().toInt()==1 && readonlyView->currentIndex().row()==0);
    assert(mixed->executeProcedure("SortTableDescending",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(readonlyModel->index(0,0).data().toInt()==2 && readonlyView->currentIndex().row()==1);
    assert(unchanged.exec("SELECT value,flag FROM command_readonly WHERE id=1") && unchanged.next() && unchanged.value(0).toDouble()==1.25 && unchanged.value(1).toInt()==-1); unchanged.finish();
    assert(mixed->executeProcedure("AddTableRecord",selectArguments).errorNumber==2448 && !readonlyModel->isDirty());
    readonlyView->clearSelection();
    assert(mixed->executeProcedure("DeleteTableRecord",selectArguments).errorNumber==2448 && !readonlyModel->isDirty());
    assert(mixed->executeProcedure("AddActiveTableRecord",selectArguments).errorNumber==2448 && !readonlyModel->isDirty());
    assert(mixed->executeProcedure("OpenRecordEditTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    auto* recordEditView=nativeTable(documentTabs->currentWidget());
    auto* recordEditModel=qobject_cast<linaccess::AccessTableModel*>(recordEditView->model()); assert(recordEditModel);
    recordEditView->sortByColumn(0,Qt::AscendingOrder);
    assert(mixed->executeProcedure("AddTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed && recordEditModel->rowCount()==2);
    const int addedRow=recordEditView->currentIndex().row(); assert(addedRow==1);
    assert(mixed->executeProcedure("AddActiveTableRecord",selectArguments).errorNumber==3251 && recordEditModel->rowCount()==2);
    assert(recordEditModel->setData(recordEditModel->index(addedRow,1),"new-record"));
    assert(mixed->executeProcedure("AddTableRecord",selectArguments).errorNumber==3251 && recordEditModel->rowCount()==2);
    assert(unchanged.exec("SELECT count(*) FROM command_record_edits") && unchanged.next() && unchanged.value(0).toInt()==1); unchanged.finish();
    recordEditView->setCurrentIndex(recordEditModel->index(addedRow,1));
    assert(mixed->executeProcedure("SaveActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(recordEditView->currentIndex().column()==1 && recordEditModel->index(recordEditView->currentIndex().row(),0).data().toInt()==2);
    assert(unchanged.exec("SELECT id FROM command_record_edits WHERE value='new-record'") && unchanged.next() && unchanged.value(0).toInt()==2); unchanged.finish();
    assert(mixed->executeProcedure("SelectAllTableRecords",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    recordEditView->selectionModel()->setCurrentIndex(recordEditModel->index(1,1),QItemSelectionModel::NoUpdate);
    assert(mixed->executeProcedure("DeleteTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed && recordEditModel->isDirty());
    assert(unchanged.exec("SELECT count(*) FROM command_record_edits") && unchanged.next() && unchanged.value(0).toInt()==2); unchanged.finish();
    recordEditModel->revertAll(); assert(!recordEditModel->isDirty());
    recordEditView->clearSelection();
    recordEditView->setCurrentIndex(recordEditModel->index(1,1));
    assert(mixed->executeProcedure("DeleteTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed && recordEditModel->isDirty());
    assert(unchanged.exec("SELECT count(*) FROM command_record_edits") && unchanged.next() && unchanged.value(0).toInt()==2); unchanged.finish();
    assert(mixed->executeProcedure("SaveActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed && !recordEditModel->isDirty());
    assert(unchanged.exec("SELECT id,value FROM command_record_edits") && unchanged.next() && unchanged.value(0).toInt()==1 && unchanged.value(1).toString()=="kept" && !unchanged.next()); unchanged.finish();
    assert(unchanged.exec("CREATE TRIGGER ignore_record_edit BEFORE INSERT ON command_record_edits WHEN NEW.value='ignored-insert' BEGIN SELECT RAISE(IGNORE); END"));
    assert(mixed->executeProcedure("AddTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    const int ignoredRow=recordEditView->currentIndex().row();
    assert(recordEditModel->setData(recordEditModel->index(ignoredRow,1),"ignored-insert"));
    assert(mixed->executeProcedure("SaveActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(recordEditModel->submittedPrimaryKey(ignoredRow).isEmpty());
    assert(unchanged.exec("SELECT count(*) FROM command_record_edits") && unchanged.next() && unchanged.value(0).toInt()==1); unchanged.finish();
    assert(unchanged.exec("DELETE FROM command_record_edits"));
    assert(mixed->executeProcedure("RequeryActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed && recordEditModel->rowCount()==0);
    assert(window.activateOpenTable("command_readonly")==0);
    assert(mixed->executeProcedure("AddNamedTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed && recordEditModel->rowCount()==1);
    recordEditView->setCurrentIndex(recordEditModel->index(0,1));
    assert(recordEditModel->setData(recordEditModel->index(0,1),"first-record"));
    assert(mixed->executeProcedure("SaveActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed && recordEditView->currentIndex()==recordEditModel->index(0,1));
    assert(unchanged.exec("SELECT value FROM command_record_edits") && unchanged.next() && unchanged.value(0).toString()=="first-record"); unchanged.finish();
    assert(mixed->executeProcedure("CloseRecordEditTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(mixed->executeProcedure("OpenNavigationTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed); QApplication::processEvents();
    auto* navigationView=nativeTable(documentTabs->currentWidget());
    assert(!readonlyPage->isVisible() && readonlyDescriptor->member("IsLoaded",{}).boolean);
    auto* navigationModel=qobject_cast<linaccess::AccessTableModel*>(navigationView->model()); assert(navigationModel);
    navigationView->sortByColumn(0,Qt::AscendingOrder);
    navigationView->setCurrentIndex(navigationModel->index(0,1));
    assert(navigationModel->rowCount()<300); // Native SQL model still has another fetch batch.
    assert(mixed->executeProcedure("SelectAllTableRecords",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(navigationModel->rowCount()==300 && navigationView->selectionModel()->selectedIndexes().size()==300*navigationModel->columnCount());
    assert(navigationView->currentIndex()==navigationModel->index(0,1) && !navigationModel->isDirty());
    navigationView->clearSelection();
    assert(mixed->executeProcedure("LastTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(navigationView->currentIndex().row()==299 && navigationView->currentIndex().column()==1 && navigationModel->rowCount()==300);
    assert(unchanged.exec("UPDATE command_navigation SET value='requeried' WHERE id=300"));
    assert(navigationModel->index(299,1).data().toString()=="row-300");
    assert(mixed->executeProcedure("RequeryActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(!navigationModel->isDirty() && navigationView->currentIndex()==navigationModel->index(0,1));
    assert(mixed->executeProcedure("LastTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationModel->index(299,1).data().toString()=="requeried");
    assert(unchanged.exec("UPDATE command_navigation SET value='row-300' WHERE id=300"));
    QApplication::setActiveWindow(&window); navigationView->setFocus(); QApplication::processEvents();
    QTest::keyClick(navigationView,Qt::Key_F9,Qt::ShiftModifier); QApplication::processEvents();
    assert(navigationView->currentIndex()==navigationModel->index(0,1));
    assert(mixed->executeProcedure("LastTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationModel->index(299,1).data().toString()=="row-300");
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
    assert(mixed->executeProcedure("SortTableDescending",selectArguments).errorNumber==3251 && navigationView->currentIndex().row()==0 && navigationModel->isDirty());
    assert(mixed->executeProcedure("RequeryActiveTable",selectArguments).errorNumber==3251 && navigationView->currentIndex().row()==0 && navigationModel->isDirty());
    assert(unchanged.exec("SELECT value FROM command_navigation WHERE id=1") && unchanged.next() && unchanged.value(0).toString()=="row-1"); unchanged.finish();
    assert(mixed->executeProcedure("SaveActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(mixed->executeProcedure("LastTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==299);
    navigationView->setCurrentIndex(navigationModel->index(299,0));
    assert(mixed->executeProcedure("SortTableDescending",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(navigationModel->index(0,0).data().toInt()==300 && navigationView->currentIndex().row()==0 && navigationView->currentIndex().column()==0);
    assert(navigationModel->setData(navigationModel->index(0,1),"sort-saved"));
    assert(mixed->executeProcedure("SortTableAscending",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(navigationView->currentIndex().row()==299 && navigationModel->index(299,0).data().toInt()==300 && !navigationModel->isDirty());
    assert(unchanged.exec("SELECT value FROM command_navigation WHERE id=300") && unchanged.next() && unchanged.value(0).toString()=="sort-saved"); unchanged.finish();
    navigationView->setCurrentIndex(navigationModel->index(299,1));
    assert(navigationModel->setData(navigationModel->index(299,1),"shift-enter-saved"));
    QApplication::setActiveWindow(&window); navigationView->setFocus(); QApplication::processEvents();
    QTest::keyClick(navigationView,Qt::Key_Return,Qt::ShiftModifier);
    assert(!navigationModel->isDirty() && navigationView->currentIndex()==navigationModel->index(299,1));
    assert(unchanged.exec("SELECT value FROM command_navigation WHERE id=300") && unchanged.next() && unchanged.value(0).toString()=="shift-enter-saved"); unchanged.finish();
    navigationView->edit(navigationModel->index(299,1));
    auto* saveEditor=qobject_cast<QLineEdit*>(QApplication::focusWidget()); assert(saveEditor);
    saveEditor->setText("keypad-enter-saved");
    QTest::keyClick(saveEditor,Qt::Key_Enter,Qt::ShiftModifier|Qt::KeypadModifier);
    assert(!navigationModel->isDirty() && navigationView->currentIndex()==navigationModel->index(299,1));
    assert(unchanged.exec("SELECT value FROM command_navigation WHERE id=300") && unchanged.next() && unchanged.value(0).toString()=="keypad-enter-saved"); unchanged.finish();
    assert(navigationModel->setData(navigationModel->index(299,1),"automatic-navigation-save"));
    assert(mixed->executeProcedure("PreviousTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(navigationView->currentIndex().row()==298 && !navigationModel->isDirty());
    assert(unchanged.exec("SELECT value FROM command_navigation WHERE id=300") && unchanged.next() && unchanged.value(0).toString()=="automatic-navigation-save"); unchanged.finish();

    assert(navigationModel->setData(navigationModel->index(298,1),"blocked"));
    assert(mixed->executeProcedure("NextTableRecord",selectArguments).errorNumber==3146 && navigationView->currentIndex().row()==298 && navigationModel->isDirty());
    assert(unchanged.exec("SELECT value FROM command_navigation WHERE id=299") && unchanged.next() && unchanged.value(0).toString()=="row-299"); unchanged.finish();
    assert(navigationModel->setData(navigationModel->index(298,1),"navigation-retry"));
    assert(mixed->executeProcedure("NextTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==299 && !navigationModel->isDirty());
    navigationView->setCurrentIndex(navigationModel->index(299,2));
    QApplication::clipboard()->setText("binary-copy-sentinel");
    assert(mixed->executeProcedure("CopyTableSelection",selectArguments).errorNumber==3251);
    assert(QApplication::clipboard()->text()=="binary-copy-sentinel" && !navigationModel->isDirty());
    assert(unchanged.exec("SELECT hex(payload) FROM command_navigation WHERE id=300") && unchanged.next() && unchanged.value(0).toString()=="0102"); unchanged.finish();
    assert(mixed->executeProcedure("RepeatTableFind",selectArguments).errorNumber==2046);
    assert(mixed->executeProcedure("FindTableText",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(window.activateOpenTable("command_readonly")==0);
    QPointer<QDialog> tableFind=navigationView->findChild<QDialog*>("native-table-find-dialog"); assert(tableFind && tableFind->isVisible() && !tableFind->isModal());
    auto* findText=tableFind->findChild<QLineEdit*>("native-table-find-text");
    auto* findMatch=tableFind->findChild<QComboBox*>("native-table-find-match");
    auto* findCase=tableFind->findChild<QCheckBox*>("native-table-find-case");
    auto* findNext=tableFind->findChild<QPushButton*>("native-table-find-next");
    auto* findResult=tableFind->findChild<QLabel*>("native-table-find-result"); assert(findText && findMatch && findCase && findNext && findResult);
    findText->setText("row-250"); findNext->click();
    assert(navigationView->currentIndex().row()==249 && navigationView->currentIndex().column()==1 && documentTabs->currentWidget()==navigationView->parentWidget());
    auto* findDirection=tableFind->findChild<QComboBox*>("native-table-find-direction"); assert(findDirection);
    findDirection->setCurrentIndex(findDirection->findData(0)); findMatch->setCurrentIndex(1); findText->setText("row-249"); findNext->click(); assert(navigationView->currentIndex().row()==248);
    findText->setText("row-280"); findNext->click(); assert(navigationView->currentIndex().row()==248 && findResult->text()=="Keine Übereinstimmung gefunden.");
    findDirection->setCurrentIndex(findDirection->findData(1)); findNext->click(); assert(navigationView->currentIndex().row()==279);
    findText->setText("row-249"); findNext->click(); assert(navigationView->currentIndex().row()==279);
    findDirection->setCurrentIndex(findDirection->findData(2)); findText->setText("row-250"); findNext->click(); assert(navigationView->currentIndex().row()==249);
    findMatch->setCurrentIndex(0);
    findText->setText("row-2"); findNext->click();
    assert(mixed->executeProcedure("RepeatTableFind",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==250);
    QTest::keyClick(navigationView,Qt::Key_F4,Qt::ShiftModifier); assert(navigationView->currentIndex().row()==251);
    auto* findEditMenu=window.findChild<QMenu*>("native-edit-menu"); auto* menuFindNext=window.findChild<QAction*>("table-find-next"); assert(findEditMenu && menuFindNext);
    QMetaObject::invokeMethod(findEditMenu,"aboutToShow",Qt::DirectConnection); assert(menuFindNext->isEnabled()); menuFindNext->trigger(); assert(navigationView->currentIndex().row()==252);
    findMatch->setCurrentIndex(1); findText->setText("ROW-250"); findNext->click(); assert(navigationView->currentIndex().row()==249);
    findCase->setChecked(true); findNext->click(); assert(navigationView->currentIndex().row()==249 && findResult->text()=="Keine Übereinstimmung gefunden.");
    findCase->setChecked(false); findMatch->setCurrentIndex(2); findText->setText("row-29"); findNext->click(); assert(navigationView->currentIndex().row()==289);
    assert(navigationModel->setData(navigationModel->index(289,1),"find-saved"));
    findText->setText("row-280"); findNext->click();
    assert(navigationView->currentIndex().row()==279 && !navigationModel->isDirty());
    assert(unchanged.exec("SELECT value FROM command_navigation WHERE id=290") && unchanged.next() && unchanged.value(0).toString()=="find-saved"); unchanged.finish();
    assert(navigationModel->setData(navigationModel->index(279,1),"pending-current") && navigationModel->setData(navigationModel->index(278,1),"pending-other"));
    findText->setText("row-270"); findNext->click();
    assert(navigationView->currentIndex().row()==279 && navigationModel->isDirty() && navigationModel->lastError().nativeErrorCode()=="3251");
    assert(mixed->executeProcedure("RepeatTableFind",selectArguments).errorNumber==3251);
    assert(unchanged.exec("SELECT value FROM command_navigation WHERE id=280") && unchanged.next() && unchanged.value(0).toString()=="row-280"); unchanged.finish();
    navigationModel->revertAll(); assert(navigationModel->setReadOnly(true));
    findText->setText("row-280"); findNext->click(); assert(navigationView->currentIndex().row()==279 && navigationModel->isReadOnly() && !navigationModel->isDirty());
    assert(navigationModel->setReadOnly(false));
    tableFind->reject(); QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete); QApplication::processEvents(); assert(!tableFind);
    assert(mixed->executeProcedure("FindRecordDefault",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==249);
    assert(makePreviewFormObject(*mixed)->member("Tag",{}).asText()=="findrecord-default");
    assert(mixed->executeProcedure("FindRecordUp",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==248);
    assert(mixed->executeProcedure("FindRecordUpNoWrap",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==248);
    assert(mixed->executeProcedure("FindRecordDown",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==279);
    assert(mixed->executeProcedure("FindRecordDownNoWrap",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==279);
    assert(mixed->executeProcedure("FindRecordDefault",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==249);
    assert(mixed->executeProcedure("FindRecordNext",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==250);
    assert(mixed->executeProcedure("FindRecordEmptyFirst",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==251);
    assert(mixed->executeProcedure("FindRecordEmptyMatch",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==24);
    assert(mixed->executeProcedure("FindRecordNoMatch",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==24);
    for(const char* procedure:{"FindRecordAllFields","FindRecordFormatted","FindRecordNumber","FindRecordWildcard"})
        assert(mixed->executeProcedure(procedure,selectArguments).errorNumber==3251 && navigationView->currentIndex().row()==24);
    assert(mixed->executeProcedure("FindRecordInvalidDirection",selectArguments).errorNumber==5);
    assert(mixed->executeProcedure("FindRecordNull",selectArguments).errorNumber==94);
    assert(makePreviewFormObject(*mixed)->member("Tag",{}).asText()=="findrecord-nohit");
    navigationView->setCurrentIndex(navigationModel->index(24,0));
    assert(mixed->executeProcedure("FindRecordDefault",selectArguments).errorNumber==3251 && navigationView->currentIndex().row()==24);
    navigationView->setCurrentIndex(navigationModel->index(24,1));
    assert(mixed->executeProcedure("FindRecordReset",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==279);
    window.activateWindow(); QApplication::processEvents(); QApplication::setActiveWindow(&window);
    navigationView->setCurrentIndex(navigationModel->index(279,1)); navigationView->setFocus();
    assert(navigationView->hasFocus());
    QTest::keyClick(navigationView,Qt::Key_F2);
    auto* findSourceEditor=qobject_cast<QLineEdit*>(QApplication::focusWidget()); assert(findSourceEditor); findSourceEditor->setText("keyboard-find-saved");
    QTest::keyClick(findSourceEditor,Qt::Key_F,Qt::ControlModifier); QApplication::processEvents();
    tableFind=navigationView->findChild<QDialog*>("native-table-find-dialog"); assert(tableFind && tableFind->isVisible() && navigationModel->isDirty());
    assert(unchanged.exec("SELECT value FROM command_navigation WHERE id=280") && unchanged.next() && unchanged.value(0).toString()=="row-280"); unchanged.finish();
    assert(mixed->executeProcedure("RequeryActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationModel->canFetchMore());
    assert(mixed->executeProcedure("AddNavigationTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(navigationModel->rowCount()==301 && navigationView->currentIndex().row()==300 && !navigationModel->canFetchMore());
    assert(unchanged.exec("SELECT count(*) FROM command_navigation") && unchanged.next() && unchanged.value(0).toInt()==300); unchanged.finish();
    navigationModel->revertAll();
    assert(mixed->executeProcedure("CloseNavigationTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed); QApplication::processEvents();
    QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete); assert(!tableFind);

    assert(mixed->executeProcedure("OpenTypedFindTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed); QApplication::processEvents();
    auto* typedView=nativeTable(documentTabs->currentWidget()); auto* typedModel=qobject_cast<linaccess::AccessTableModel*>(typedView->model()); assert(typedModel && typedModel->isReadOnly());
    typedView->sortByColumn(0,Qt::AscendingOrder); typedView->setCurrentIndex(typedModel->index(3,1));
    assert(mixed->executeProcedure("FindTypedNumber",selectArguments).status==linaccess::VbaExecutionStatus::Completed && typedView->currentIndex().row()==0);
    assert(mixed->executeProcedure("FindTypedAdjacent",selectArguments).status==linaccess::VbaExecutionStatus::Completed && typedView->currentIndex().row()==1);
    assert(mixed->executeProcedure("FindTypedNumberAbsent",selectArguments).status==linaccess::VbaExecutionStatus::Completed && typedView->currentIndex().row()==1);
    assert(mixed->executeProcedure("FindTypedPartial",selectArguments).errorNumber==3251);
    typedView->setCurrentIndex(typedModel->index(1,2));
    assert(mixed->executeProcedure("FindTypedMoney",selectArguments).status==linaccess::VbaExecutionStatus::Completed && typedView->currentIndex().row()==0);
    assert(mixed->executeProcedure("FindTypedMoneySmall",selectArguments).status==linaccess::VbaExecutionStatus::Completed && typedView->currentIndex().row()==2);
    assert(mixed->executeProcedure("FindTypedNumber",selectArguments).errorNumber==3251 && typedView->currentIndex().row()==2);
    typedView->setCurrentIndex(typedModel->index(2,3));
    assert(mixed->executeProcedure("FindTypedStamp",selectArguments).status==linaccess::VbaExecutionStatus::Completed && typedView->currentIndex().row()==0);
    assert(mixed->executeProcedure("FindTypedNegativeStamp",selectArguments).status==linaccess::VbaExecutionStatus::Completed && typedView->currentIndex().row()==1);
    assert(mixed->executeProcedure("FindTypedSubsecond",selectArguments).errorNumber==3251 && typedView->currentIndex().row()==1);
    typedView->setCurrentIndex(typedModel->index(1,4));
    assert(mixed->executeProcedure("FindTypedFalse",selectArguments).status==linaccess::VbaExecutionStatus::Completed && typedView->currentIndex().row()==2);
    QMetaObject::invokeMethod(findEditMenu,"aboutToShow",Qt::DirectConnection); assert(menuFindNext->isEnabled()); menuFindNext->trigger(); assert(typedView->currentIndex().row()==3);
    assert(mixed->executeProcedure("FindTypedTrue",selectArguments).status==linaccess::VbaExecutionStatus::Completed && typedView->currentIndex().row()==0);
    assert(mixed->executeProcedure("FindTableText",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    QPointer<QDialog> typedFind=typedView->findChild<QDialog*>("native-table-find-dialog"); assert(typedFind && typedFind->isVisible());
    auto* typedQuery=typedFind->findChild<QLineEdit*>("native-table-find-text"); auto* typedScope=typedFind->findChild<QComboBox*>("native-table-find-field");
    auto* typedMatch=typedFind->findChild<QComboBox*>("native-table-find-match"); auto* typedCase=typedFind->findChild<QCheckBox*>("native-table-find-case");
    auto* typedNext=typedFind->findChild<QPushButton*>("native-table-find-next"); assert(typedQuery && typedScope && typedMatch && typedCase && typedNext);
    assert(typedScope->findData(-1)<0 && typedScope->currentData().toInt()==4 && typedMatch->currentIndex()==1 && !typedMatch->isEnabled() && !typedCase->isEnabled());
    typedScope->setCurrentIndex(typedScope->findData(1)); typedQuery->setText("1.2345678901234565"); typedNext->click(); assert(typedView->currentIndex().row()==1);
    typedScope->setCurrentIndex(typedScope->findData(2)); typedQuery->setText("0.1001"); typedNext->click(); assert(typedView->currentIndex().row()==0);
    typedScope->setCurrentIndex(typedScope->findData(3)); typedQuery->setText("1899-12-30 06:00:00"); typedNext->click(); assert(typedView->currentIndex().row()==1);
    typedQuery->setText("not-a-date"); typedNext->click(); assert(typedView->currentIndex().row()==1 && typedModel->lastError().nativeErrorCode()=="13");
    typedScope->setCurrentIndex(typedScope->findData(4)); typedQuery->setText("False"); typedNext->click(); assert(typedView->currentIndex().row()==2);
    assert(mixed->executeProcedure("RepeatTableFind",selectArguments).status==linaccess::VbaExecutionStatus::Completed && typedView->currentIndex().row()==3);
    QApplication::setActiveWindow(typedFind); typedQuery->setFocus(); assert(typedQuery->hasFocus());
    typedQuery->selectAll(); QApplication::clipboard()->setText("True");
    assert(mixed->executeProcedure("PasteTableSelection",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(typedQuery->text()=="True" && !typedModel->isDirty());
    typedQuery->selectAll();
    assert(mixed->executeProcedure("CopyTableSelection",selectArguments).status==linaccess::VbaExecutionStatus::Completed && QApplication::clipboard()->text()=="True");
    assert(mixed->executeProcedure("CutTableSelection",selectArguments).status==linaccess::VbaExecutionStatus::Completed && typedQuery->text().isEmpty() && !typedModel->isDirty());

    typedFind->reject(); QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete); assert(!typedFind);
    assert(!typedModel->isDirty());
    assert(unchanged.exec("SELECT value,typeof(money),stamp,flag,hex(payload) FROM command_typed_find WHERE id=2") && unchanged.next() && unchanged.value(0).toDouble()==1.2345678901234565 && unchanged.value(1).toString()=="real" && unchanged.value(2).toString()=="1899-12-30 06:00:00" && unchanged.value(3).isNull() && unchanged.value(4).toString()=="0102"); unchanged.finish();
    assert(mixed->executeProcedure("CloseTypedFindTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(mixed->executeProcedure("OpenTypedPasteTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    typedView=nativeTable(documentTabs->currentWidget());
    typedModel=qobject_cast<linaccess::AccessTableModel*>(typedView->model()); assert(typedModel && !typedModel->isReadOnly());
    typedView->setCurrentIndex(typedModel->index(0,1));
    const auto originalFirstNumber=typedModel->index(0,1).data(Qt::EditRole).toString();
    QApplication::clipboard()->setText("2.5\nnot-a-number");
    const auto invalidPaste=mixed->executeProcedure("PasteTableSelection",selectArguments);
    assert(invalidPaste.errorNumber==13 && !typedModel->isDirty());
    assert(typedModel->index(0,1).data(Qt::EditRole).toString()==originalFirstNumber);
    QApplication::clipboard()->setText("2.5\n3.75");
    assert(mixed->executeProcedure("PasteTableSelection",selectArguments).status==linaccess::VbaExecutionStatus::Completed && typedModel->isDirty());
    assert(typedModel->index(0,1).data(Qt::EditRole).toDouble()==2.5 && typedModel->index(1,1).data(Qt::EditRole).toDouble()==3.75);
    assert(unchanged.exec("SELECT value FROM command_typed_find WHERE id=2") && unchanged.next() && unchanged.value(0).toDouble()==1.2345678901234565); unchanged.finish();
    QApplication::clipboard()->setText("9");
    assert(mixed->executeProcedure("PasteTableSelection",selectArguments).errorNumber==2501 && typedModel->index(0,1).data(Qt::EditRole).toDouble()==2.5);
    typedModel->revertAll();
    QApplication::setActiveWindow(&window); typedView->setCurrentIndex(typedModel->index(0,1)); typedView->setFocus(); QApplication::processEvents();
    QApplication::clipboard()->setText("4.5\n6.75"); QTest::keyClick(typedView,Qt::Key_Insert,Qt::ShiftModifier);
    assert(typedModel->isDirty() && typedModel->index(0,1).data(Qt::EditRole).toDouble()==4.5 && typedModel->index(1,1).data(Qt::EditRole).toDouble()==6.75);
    typedModel->revertAll();
    assert(mixed->executeProcedure("CloseTypedFindTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed); QApplication::processEvents();
    assert(mixed->executeProcedure("OpenEditableTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed); QApplication::processEvents();
    const QPointer<QWidget> editablePage=documentTabs->currentWidget(); auto* editableView=nativeTable(editablePage);
    auto* editableModel=qobject_cast<linaccess::AccessTableModel*>(editableView->model()); assert(editableModel && !editableModel->isReadOnly());
    const auto callerTextBeforeKeys=mixedEditor->text();
    const auto focusEditable=[&] {
        assert(editableView->finishEditing()); QApplication::processEvents(); QApplication::setActiveWindow(&window);
        editableView->setCurrentIndex(editableModel->index(0,1)); editableView->setFocus(); assert(editableView->hasFocus());
    };
    focusEditable();
    editableView->edit(editableModel->index(0,1));
    auto* pastedEditor=qobject_cast<QLineEdit*>(QApplication::focusWidget()); assert(pastedEditor);
    pastedEditor->deselect(); pastedEditor->setCursorPosition(pastedEditor->text().size());
    QApplication::clipboard()->setText("-pasted");
    assert(mixed->executeProcedure("PasteTableSelection",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(pastedEditor->text()=="initial-pasted" && editableModel->index(0,1).data(Qt::EditRole).toString()=="initial" && !editableModel->isDirty());
    pastedEditor->setSelection(0,7);
    assert(mixed->executeProcedure("CopyTableSelection",selectArguments).status==linaccess::VbaExecutionStatus::Completed && QApplication::clipboard()->text()=="initial" && pastedEditor->text()=="initial-pasted");
    assert(mixed->executeProcedure("CutTableSelection",selectArguments).status==linaccess::VbaExecutionStatus::Completed && pastedEditor->text()=="-pasted" && !editableModel->isDirty());
    assert(mixed->executeProcedure("PasteTableSelection",selectArguments).status==linaccess::VbaExecutionStatus::Completed && pastedEditor->text()=="initial-pasted");
    assert(mixed->executeProcedure("UndoTableText",selectArguments).status==linaccess::VbaExecutionStatus::Completed && pastedEditor->text()=="-pasted" && !editableModel->isDirty());
    assert(mixed->executeProcedure("RedoTableText",selectArguments).status==linaccess::VbaExecutionStatus::Completed && pastedEditor->text()=="initial-pasted" && !editableModel->isDirty());
    QApplication::clipboard()->setText("no-selection-sentinel"); pastedEditor->deselect();
    assert(mixed->executeProcedure("CopyTableSelection",selectArguments).errorNumber==2046 && QApplication::clipboard()->text()=="no-selection-sentinel");

    assert(editableView->finishEditing() && editableModel->index(0,1).data(Qt::EditRole).toString()=="initial-pasted" && editableModel->isDirty());
    editableModel->revertAll(); focusEditable();
    const auto keyFindResult=mixed->executeProcedure("FindTableByKeys",selectArguments);
    assert(keyFindResult.status==linaccess::VbaExecutionStatus::Completed);
    QPointer<QDialog> keysFind=editableView->findChild<QDialog*>("native-table-find-dialog"); assert(keysFind && keysFind->isVisible());
    assert(keysFind->findChild<QLineEdit*>("native-table-find-text")->text()=="initial" && keysFind->findChild<QLabel*>("native-table-find-result")->text().contains("Übereinstimmung"));
    assert(!editableModel->isDirty() && mixedEditor->text()==callerTextBeforeKeys);
    assert(mixed->executeProcedure("CloseTableFindByKeys",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete); assert(!keysFind);
    focusEditable();
    bool modalFallback=false; QTimer modalWatch; modalWatch.setSingleShot(true);
    QObject::connect(&modalWatch,&QTimer::timeout,[&] { modalFallback=true; if(auto* modal=qobject_cast<QDialog*>(QApplication::activeModalWidget())) modal->reject(); }); modalWatch.start(2000);
    assert(mixed->executeProcedure("ZoomTableByKeys",selectArguments).status==linaccess::VbaExecutionStatus::Completed); modalWatch.stop();
    assert(!modalFallback && !QApplication::activeModalWidget() && !editableModel->isDirty());
    focusEditable();
    assert(mixed->executeProcedure("SendTableKeys",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(editableModel->index(0,1).data(Qt::EditRole).toString()=="key-typed" && editableModel->isDirty() && mixedEditor->text()==callerTextBeforeKeys);
    focusEditable();
    assert(mixed->executeProcedure("SendTableKeysDoCmd",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(editableModel->index(0,1).data(Qt::EditRole).toString()=="method-keys" && mixedEditor->text()==callerTextBeforeKeys);
    assert(mixed->executeProcedure("RejectTableKeySequence",selectArguments).errorNumber==3251);
    assert(mixed->executeProcedure("RejectNullTableKeys",selectArguments).errorNumber==94);
    assert(editableModel->index(0,1).data(Qt::EditRole).toString()=="method-keys" && makePreviewFormObject(*mixed)->member("Tag",{}).asText()=="docmd-table-keys");
    focusEditable();
    assert(mixed->executeProcedure("QueueTableKeys",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    window.focusDatabaseWindow(); QApplication::processEvents();
    assert(editableModel->index(0,1).data(Qt::EditRole).toString()=="method-keys" && mixedEditor->text()==callerTextBeforeKeys);
    assert(QTest::qWaitFor([&] {
        for(const auto* label:mixed->findChildren<QLabel*>()) if(label->text().contains("SendKeys: Fehler 2110")) return true;
        return false;
    },2000));
    assert(window.activateOpenTable("command_edit")==0); focusEditable();
    assert(mixed->executeProcedure("QueueTableKeys",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(QTest::qWaitFor([&] { return editableModel->index(0,1).data(Qt::EditRole).toString()=="queued-keys"; },2000));
    assert(mixedEditor->text()==callerTextBeforeKeys);
    assert(unchanged.exec("SELECT value FROM command_edit") && unchanged.next() && unchanged.value(0).toString()=="initial"); unchanged.finish();
    assert(editableView->finishEditing() && editableModel->setData(editableModel->index(0,1),"changed"));
    makePreviewFormObject(*mixed)->assign("Tag",linaccess::VbaValue::string("edit-opened"));
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
    assert(!readonlyDescriptor->member("IsLoaded",{}).boolean);
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
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="native-query-designer");
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
    QApplication app(argc,argv); app.setQuitOnLastWindowClosed(false); app.setApplicationName("LinAccessSyntheticGuiTest"); app.setOrganizationName("LinAccessSyntheticTests");
    QStandardPaths::setTestModeEnabled(true);
    QTimer watchdog; watchdog.setSingleShot(true);
    QObject::connect(&watchdog,&QTimer::timeout,[] { std::cerr<<"Synthetic main-window document test timed out\n"; std::abort(); }); watchdog.start(20000);
    QTemporaryDir temporary; assert(temporary.isValid());
    QSettings::setDefaultFormat(QSettings::IniFormat);
    QSettings::setPath(QSettings::IniFormat,QSettings::UserScope,temporary.path());
    const auto first=temporary.filePath("first.linaccess"),second=temporary.filePath("second.linaccess");
    auto seed=QSqlDatabase::addDatabase("QSQLITE","mainwindow-seed"); seed.setDatabaseName(first); assert(seed.open());
    QSqlQuery initial(seed); assert(initial.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(initial.exec("INSERT INTO records VALUES(1,'old')")); assert(initial.exec("CREATE TABLE events(id INTEGER PRIMARY KEY,kind TEXT)")); initial.finish(); seed.close();
    seed.setDatabaseName(second); assert(seed.open()); seed.close();
    MainWindow window; window.show(); QApplication::setActiveWindow(&window); QApplication::processEvents();
    {
        const auto originalFont=app.font();
        auto* options=window.findChild<QAction*>("application-options"); assert(options);
        QTimer::singleShot(0,&window,[&] {
            auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="application-options-dialog");
            auto* size=dialog->findChild<QSpinBox*>("options-font-size"); assert(size); size->setValue(13);
            dialog->reject();
        });
        options->trigger(); assert(app.font()==originalFont && !QSettings().contains("ui/applicationFont"));
        QTimer::singleShot(0,&window,[&] {
            auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog);
            auto* size=dialog->findChild<QSpinBox*>("options-font-size"); assert(size); size->setValue(10);
            auto* buttons=dialog->findChild<QDialogButtonBox*>(); assert(buttons);
            buttons->button(QDialogButtonBox::Apply)->click();
            assert(app.font().pointSize()==10 && QSettings().value("ui/applicationFont").value<QFont>().pointSize()==10);
            size->setValue(13); dialog->reject();
        });
        options->trigger(); assert(app.font().pointSize()==10);
        app.setFont(originalFont); linaccess::initializeClassicStyle(); assert(app.font().pointSize()==10);
        QTimer::singleShot(0,&window,[&] {
            auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog);
            auto* size=dialog->findChild<QSpinBox*>("options-font-size"); assert(size && size->value()==10); size->setValue(9);
            dialog->findChild<QDialogButtonBox*>()->button(QDialogButtonBox::Ok)->click();
        });
        options->trigger(); assert(app.font().pointSize()==9);
        QSettings().remove("ui/applicationFont"); QSettings().remove("ui/classicStyle");
        app.setFont(originalFont);
    }
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
        "Sub ProjectKeysWhileRunning()\nSendKeys \"^o\",True\nMe.Tag=\"returned\"\nEnd Sub\n"
        "Sub ReadProjectMetadata()\nMe.Tag=CurrentProject.FullName\nEnd Sub\n"
        "Sub ProjectIdentity()\nIf CurrentProject Is Application.CurrentProject Then Me.Tag=\"same-project\"\nEnd Sub\n"
        "Sub RenameProjectMetadata()\nCurrentProject.Name=\"renamed.linaccess\"\nMe.Tag=\"renamed\"\nEnd Sub\n"
        "Sub ShadowProjectMetadata()\nDim CurrentProject As String\nCurrentProject=\"local-project-name\"\nMe.Tag=CurrentProject\nEnd Sub\n"
        "Sub ReadProjectCatalog()\nDim item As Object\nDim found As Boolean\nFor Each item In CurrentProject.AllForms\nIf item.Name=\"entry\" Then found=True\nNext item\nIf found And CurrentProject.AllForms.Item(Index:=\"entry\").IsLoaded And Not CurrentProject.AllReports(Index:=\"report\").IsLoaded Then Me.Tag=\"catalog-ready\"\nEnd Sub\n"
        "Sub ReadDataCatalog()\nDim item As Object\nDim found As Boolean\nFor Each item In CurrentData.AllTables\nIf item.Name=\"records\" Then found=True\nNext item\nIf found And Application.CurrentData.AllQueries(Index:=\"source_query\").Name=\"source_query\" Then Me.Tag=\"data-catalog-ready\"\nEnd Sub\n";
    const int formId=artifact(db,"form","entry",form); artifact(db,"vba","Form_entry",code);
    const int queryId=artifact(db,"query","source_query","SELECT value FROM records");
    auto report=form; report.replace("Begin Form","Begin Report"); report.replace("OnLoad =\"[Event Procedure]\"\n",""); report.replace("OnUnload =\"[Event Procedure]\"\n","");
    const int reportId=artifact(db,"report","report",report);
    artifact(db,"vba","Report_report",QByteArray("Sub Report_Activate()\nCurrentDb.Execute \"INSERT INTO events(kind) VALUES('report-activate')\"\nEnd Sub\nSub Report_Deactivate()\nCurrentDb.Execute \"INSERT INTO events(kind) VALUES('report-deactivate')\"\nEnd Sub\nSub HideSelf()\nMe.Visible=False\nMe.Tag=\"hidden-report\"\nEnd Sub\nSub ShowSelf()\nMe.Visible=True\nMe.Tag=\"shown-report\"\nEnd Sub\n"));
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
    auto projectMetadata=makePreviewApplicationObject(*entry)->member("CurrentProject",{}).object; assert(projectMetadata);
    const auto statusApplication=makePreviewApplicationObject(*entry);
    const auto objectState=[&](int type,const std::string& name) {
        return statusApplication->member("SysCmd",{linaccess::VbaValue::integer(10),linaccess::VbaValue::integer(type),linaccess::VbaValue::string(name)}).asNumber();
    };
    assert(objectState(2,"ENTRY")==1 && objectState(3,"report")==0 && objectState(2,"missing state object")==0);
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="native-query-designer");
        dialog->findChild<QLineEdit*>("native-query-name")->setText("Native State Query");
        auto* code=dialog->findChild<QPlainTextEdit*>("native-query-sql"); code->setPlainText("SELECT 1 AS value");
        QApplication::setActiveWindow(dialog); code->setFocus(); QApplication::processEvents();
        QTest::keyClick(code,Qt::Key_S,Qt::ControlModifier);
        assert(objectState(1,"Native State Query")==1);
        code->setPlainText("SELECT 2 AS value"); assert(objectState(1,"Native State Query")==3);
        QTest::keyClick(code,Qt::Key_S,Qt::ControlModifier); assert(objectState(1,"Native State Query")==1);
        dialog->reject();
    });
    action(window,"+ Abfrage")->trigger(); assert(objectState(1,"Native State Query")==0);

    assert(projectMetadata->member("Name",{}).asText()=="first.linaccess" && projectMetadata->member("Path",{}).asText()==temporary.path().toStdString());
    assert(entry->executeProcedure("ReadProjectMetadata",arguments).status==linaccess::VbaExecutionStatus::Completed && root->member("Tag",{}).asText()==first.toStdString());
    assert(entry->executeProcedure("RenameProjectMetadata",arguments).errorNumber==438 && root->member("Tag",{}).asText()==first.toStdString() && QFileInfo::exists(first));
    assert(entry->executeProcedure("ProjectIdentity",arguments).status==linaccess::VbaExecutionStatus::Completed && root->member("Tag",{}).asText()=="same-project");
    assert(entry->executeProcedure("ShadowProjectMetadata",arguments).status==linaccess::VbaExecutionStatus::Completed && root->member("Tag",{}).asText()=="local-project-name");
    auto allForms=projectMetadata->member("AllForms",{}).object;
    auto allReports=projectMetadata->member("AllReports",{}).object;
    auto dataMetadata=makePreviewApplicationObject(*entry)->member("CurrentData",{}).object;
    auto allTables=dataMetadata->member("AllTables",{}).object;
    auto allQueries=dataMetadata->member("AllQueries",{}).object;
    auto catalogTable=allTables->member("Item",{linaccess::VbaValue::string("RECORDS")}).object;
    auto catalogQuery=allQueries->member("Item",{linaccess::VbaValue::string("source_query")}).object;
    assert(catalogTable->member("Name",{}).asText()=="records" && !catalogTable->member("IsLoaded",{}).boolean && !catalogQuery->member("IsLoaded",{}).boolean);
    assert(dataMetadata==makePreviewApplicationObject(*entry)->member("CurrentData",{}).object);
    assert(entry->executeProcedure("ReadDataCatalog",arguments).status==linaccess::VbaExecutionStatus::Completed && root->member("Tag",{}).asText()=="data-catalog-ready");
    int internalTable=0; try { (void)allTables->member("Item",{linaccess::VbaValue::string("_la_objects")}); } catch(const linaccess::VbaRuntimeError& error) { internalTable=error.number; } assert(internalTable==3265);
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="access-artifact-details");
        assert(catalogQuery->member("IsLoaded",{}).boolean); dialog->reject();
    });
    auto* catalogExplorer=window.findChild<QTreeWidget*>("database-explorer"); QTreeWidgetItem* catalogQueryItem=nullptr;
    for(QTreeWidgetItemIterator it(catalogExplorer);*it;++it)
        if((*it)->data(0,Qt::UserRole).toString()=="access_artifact" && (*it)->data(0,Qt::UserRole+1).toInt()==queryId) catalogQueryItem=*it;
    assert(catalogQueryItem && QMetaObject::invokeMethod(catalogExplorer,"itemDoubleClicked",Qt::DirectConnection,Q_ARG(QTreeWidgetItem*,catalogQueryItem),Q_ARG(int,0)));
    assert(!catalogQuery->member("IsLoaded",{}).boolean);
    auto catalogEntry=allForms->member("Item",{linaccess::VbaValue::string("ENTRY")}).object;
    auto catalogReport=allReports->member("Item",{linaccess::VbaValue::string("report")}).object;
    assert(catalogEntry->member("Name",{}).asText()=="entry" && catalogEntry->member("IsLoaded",{}).boolean && !catalogReport->member("IsLoaded",{}).boolean);
    assert(allForms->member("Count",{}).asNumber()>0 && allReports->member("Count",{}).asNumber()>0);
    assert(allForms->member("Item",{linaccess::VbaValue::string("entry")}).object==catalogEntry);
    const auto catalogCount=allForms->member("Count",{}).asNumber();
    assert(sql.exec("CREATE TEMP TABLE _la_objects(kind TEXT,name TEXT,definition TEXT)"));
    assert(sql.exec("INSERT INTO temp._la_objects VALUES('form','temporary_shadow','{}')"));
    assert(allForms->member("Count",{}).asNumber()==catalogCount);
    int shadowItem=0; try { (void)allForms->member("Item",{linaccess::VbaValue::string("temporary_shadow")}); } catch(const linaccess::VbaRuntimeError& error) { shadowItem=error.number; } assert(shadowItem==3265);
    assert(sql.exec("DROP TABLE temp._la_objects"));
    assert(entry->executeProcedure("ReadProjectCatalog",arguments).status==linaccess::VbaExecutionStatus::Completed && root->member("Tag",{}).asText()=="catalog-ready");
    artifact(db,"form","catalog_subform",QByteArray("Begin Form\nEnd\n"));
    auto catalogChild=allForms->member("Item",{linaccess::VbaValue::string("catalog_subform")}).object;
    {
        AccessFormPreview parent("catalog_parent",QByteArray("Begin Form\nBegin Section\nBegin Subform\nName =\"child\"\nSourceObject =\"Form.catalog_subform\"\nEnd\nEnd\nEnd\n"),{},db,&window);
        assert(catalogChild->member("IsLoaded",{}).boolean);
        assert(parent.close() && !catalogChild->member("IsLoaded",{}).boolean);
    }
    for(const auto& invalid:std::vector<std::pair<linaccess::VbaValue,int>>{{{},94},{linaccess::VbaValue::longInteger(-1),3265},{linaccess::VbaValue::string("absent"),3265}}) {
        int number=0; try { (void)allForms->member("Item",{invalid.first}); } catch(const linaccess::VbaRuntimeError& error) { number=error.number; } assert(number==invalid.second);
    }
    {
        const auto foreignFile=temporary.filePath("foreign.linaccess");
        auto foreignDb=QSqlDatabase::addDatabase("QSQLITE","mainwindow-project-metadata-foreign"); foreignDb.setDatabaseName(foreignFile); assert(foreignDb.open());
        std::shared_ptr<linaccess::VbaObject> foreignProject;
        {
            AccessFormPreview foreignPreview("foreign_metadata",QByteArray("Begin Form\nEnd\n"),{},foreignDb,&window);
            foreignProject=makePreviewApplicationObject(foreignPreview)->member("CurrentProject",{}).object;
            assert(foreignProject!=projectMetadata && foreignProject->member("FullName",{}).asText()==foreignFile.toStdString());
            assert(makePreviewApplicationObject(*entry)->member("CurrentProject",{}).object==projectMetadata);
        }
        assert(foreignProject->member("FullName",{}).asText()==foreignFile.toStdString());
        int noCatalog=0; try { (void)foreignProject->member("AllForms",{}).object->member("Count",{}); } catch(const linaccess::VbaRuntimeError& error) { noCatalog=error.number; } assert(noCatalog==3251);
        foreignDb.close(); bool closedProject=false;
        try { (void)foreignProject->member("Name",{}); } catch(const linaccess::VbaRuntimeError& error) { closedProject=error.number==3420; }
        assert(closedProject); foreignProject.reset(); foreignDb=QSqlDatabase{}; QSqlDatabase::removeDatabase("mainwindow-project-metadata-foreign");
    }
    assert(sql.exec(R"SQL(INSERT INTO _la_objects(kind,name,definition) VALUES('form','catalog_native_form','{"table":"records","fields":["value"]}'))SQL"));
    action(window,"Aktualisieren")->trigger();
    auto nativeCatalog=allForms->member("Item",{linaccess::VbaValue::string("catalog_native_form")}).object;
    QTreeWidgetItem* nativeFormItem=nullptr;
    for(QTreeWidgetItemIterator it(window.findChild<QTreeWidget*>("database-explorer"));*it;++it)
        if((*it)->data(0,Qt::UserRole).toString()=="form" && (*it)->text(0)=="catalog_native_form") nativeFormItem=*it;
    assert(nativeFormItem && !nativeCatalog->member("IsLoaded",{}).boolean);
    assert(QMetaObject::invokeMethod(window.findChild<QTreeWidget*>("database-explorer"),"itemDoubleClicked",Qt::DirectConnection,Q_ARG(QTreeWidgetItem*,nativeFormItem),Q_ARG(int,0)));
    assert(nativeCatalog->member("IsLoaded",{}).boolean && !QApplication::activeModalWidget());
    QPointer<AccessFormPreview> nativeDocument;
    for(const auto& document:AccessFormPreview::projectDocuments(db))
        if(document && document->documentName()=="catalog_native_form") nativeDocument=document;
    assert(nativeDocument && nativeDocument->hasDataSource() && nativeDocument->close());
    dispose();
    assert(!nativeCatalog->member("IsLoaded",{}).boolean);
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
    int ambiguousCatalog=0; try { (void)allForms->member("Count",{}); } catch(const linaccess::VbaRuntimeError& error) { ambiguousCatalog=error.number; } assert(ambiguousCatalog==5);
    badSelection({VbaValue::numeric(2),VbaValue::string("entry"),VbaValue::logical(true)},5);
    assert(sql.exec("DELETE FROM _la_objects WHERE kind='form' AND name='ENTRY'")); action(window,"Aktualisieren")->trigger();
    assert(entry->executeProcedure("SelectEntry",arguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(entry->executeProcedure("ExplicitSave",arguments).status==linaccess::VbaExecutionStatus::Completed);
    root->assign("Tag",linaccess::VbaValue::string("hold"));
    action(window,"Alle schließen")->trigger(); dispose(); assert(entry && entry->isVisible());

    QApplication::setActiveWindow(entry); editor->setFocus(); openProject(window,second,editor); assert(db.databaseName()==first && db.isOpen() && entry && entry->isVisible());
    assert(projectMetadata->member("FullName",{}).asText()==first.toStdString());
    root->assign("Tag",linaccess::VbaValue::string(""));
    QApplication::setActiveWindow(entry); editor->setFocus(); QApplication::processEvents();
    bool retained=false;
    QTimer::singleShot(0,&window,[&] {
        auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget()); assert(box);
        dispose(); retained=entry!=nullptr; assert(retained); box->button(QMessageBox::Ok)->click();
    });
    const auto closed=entry->executeProcedure("CloseSelf",arguments); assert(closed.status==linaccess::VbaExecutionStatus::Completed && retained);
    dispose(); assert(!entry);
    assert(projectMetadata->member("FullName",{}).asText()==first.toStdString());
    assert(!catalogEntry->member("IsLoaded",{}).boolean);
    assert(catalogTable->member("Name",{}).asText()=="records" && !catalogTable->member("IsLoaded",{}).boolean);
    assert(sql.exec("SELECT value FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="second"); sql.finish();
    openArtifact(window,reportId); QPointer<AccessFormPreview> reportWindow=document("report"); assert(reportWindow && !reportWindow->isModal() && QApplication::activeModalWidget()==nullptr);
    assert(catalogReport->member("IsLoaded",{}).boolean);
    assert(reportWindow->executeProcedure("HideSelf",arguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(reportWindow->isHidden() && !makePreviewFormObject(*reportWindow)->member("Visible",{}).boolean && catalogReport->member("IsLoaded",{}).boolean);
    assert(makePreviewFormObject(*reportWindow)->member("Tag",{}).asText()=="hidden-report");
    assert(sql.exec("SELECT count(*) FROM events WHERE kind='report-deactivate'") && sql.next() && sql.value(0).toInt()>0); sql.finish();
    assert(reportWindow->executeProcedure("ShowSelf",arguments).status==linaccess::VbaExecutionStatus::Completed);
    QApplication::setActiveWindow(reportWindow); QApplication::processEvents();
    assert(makePreviewFormObject(*reportWindow)->member("Visible",{}).boolean && makePreviewFormObject(*reportWindow)->member("Tag",{}).asText()=="shown-report");
    assert(sql.exec("SELECT count(*) FROM events WHERE kind='report-activate'") && sql.next() && sql.value(0).toInt()>0); sql.finish();
    assert(reportWindow->findChild<QLineEdit*>("access-control-editor")->isReadOnly()); action(window,"Alle schließen")->trigger(); dispose(); assert(!reportWindow && window.findChild<QTabWidget*>()->count()==0);
    openArtifact(window,modalId); QPointer<AccessFormPreview> modalWindow=document("modal"); assert(modalWindow && modalWindow->isModal() && QApplication::activeModalWidget()==modalWindow);
    assert(allForms->member("Item",{linaccess::VbaValue::string("modal")}).object->member("IsLoaded",{}).boolean);
    auto* modalEditor=modalWindow->findChild<QLineEdit*>("access-control-editor"); assert(modalEditor); modalEditor->setFocus(); QTest::keyClick(modalEditor,Qt::Key_F11);
    assert(QApplication::activeModalWidget()==modalWindow && !AccessFormPreview::databaseWindowActive());
    QTest::keyClick(modalEditor,Qt::Key_O,Qt::ControlModifier); QTest::keyClick(modalEditor,Qt::Key_N,Qt::ControlModifier); assert(QApplication::activeModalWidget()==modalWindow);
    QTest::keyClick(modalEditor,Qt::Key_F6,Qt::ControlModifier); QTest::keyClick(modalEditor,Qt::Key_F6,Qt::ControlModifier|Qt::ShiftModifier);
    assert(QApplication::activeModalWidget()==modalWindow && !AccessFormPreview::databaseWindowActive());
    selection=modalWindow->executeProcedure("SelectTable",arguments); assert(selection.status==linaccess::VbaExecutionStatus::Error && selection.errorNumber==2046 && QApplication::activeModalWidget()==modalWindow);
    assert(modalWindow->close()); dispose(); assert(!modalWindow && QApplication::activeModalWidget()==nullptr);
    openArtifact(window,cancelId); assert(!document("cancel") && QApplication::activeModalWidget()==nullptr);
    openProject(window,first,explorer); // Same filename still starts a new project generation.
    bool staleProject=false;
    try { (void)projectMetadata->member("FullName",{}); } catch(const linaccess::VbaRuntimeError& error) { staleProject=error.number==3420; }
    assert(staleProject);
    int staleData=0; try { (void)allTables->member("Count",{}); } catch(const linaccess::VbaRuntimeError& error) { staleData=error.number; } assert(staleData==3420);
    int staleCatalog=0; try { (void)catalogEntry->member("IsLoaded",{}); } catch(const linaccess::VbaRuntimeError& error) { staleCatalog=error.number; } assert(staleCatalog==3420);
    openArtifact(window,formId); auto* reopenedEntry=document("entry"); assert(reopenedEntry);
    auto reopenedProject=makePreviewApplicationObject(*reopenedEntry)->member("CurrentProject",{}).object;
    assert(reopenedProject!=projectMetadata && reopenedProject->member("FullName",{}).asText()==first.toStdString());
    assert(reopenedEntry->close()); dispose();
    openProject(window,second,explorer); assert(QSqlDatabase::database("linaccess-main").databaseName()==second);
    staleProject=false;
    try { (void)reopenedProject->member("FullName",{}); } catch(const linaccess::VbaRuntimeError& error) { staleProject=error.number==3420; }
    assert(staleProject);
    const auto third=temporary.filePath("third.linaccess"); bool created=false;
    QTimer::singleShot(0,&window,[&] { auto* dialog=qobject_cast<QFileDialog*>(QApplication::activeModalWidget()); assert(dialog); created=true; dialog->selectFile(third); static_cast<QDialog*>(dialog)->accept(); });
    QTest::keyClick(explorer,Qt::Key_N,Qt::ControlModifier);
    assert(created && QFileInfo::exists(third) && QSqlDatabase::database("linaccess-main").databaseName()==third);
    {
        AccessFormPreview probe("empty_catalog_probe",QByteArray("Begin Form\nEnd\n"),{},QSqlDatabase::database("linaccess-main"),&window);
        auto emptyProject=makePreviewApplicationObject(probe)->member("CurrentProject",{}).object;
        assert(emptyProject->member("AllForms",{}).object->member("Count",{}).asNumber()==0 && emptyProject->member("AllReports",{}).object->member("Count",{}).asNumber()==0);
        auto emptyData=makePreviewApplicationObject(probe)->member("CurrentData",{}).object;
        assert(emptyData->member("AllTables",{}).object->member("Count",{}).asNumber()==0 && emptyData->member("AllQueries",{}).object->member("Count",{}).asNumber()==0);
    }
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
    exerciseReferenceDatasheet();
    exerciseDatasheetOutput(window,db);
    exerciseTableReplace(window,db);
    exerciseTableMultiDelete(window,db);
    exerciseRecordEscape(window,db);
    exercisePageNavigation(window,db);
    exerciseStoredForm(window,db);
    exerciseStoredReport(window,db);
    {
        QTimer::singleShot(0,&window,[&] {
            auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="native-table-designer");
            dialog->findChild<QLineEdit*>("native-table-name")->setText("authored_defaults");
            auto* grid=dialog->findChild<QTableWidget*>("native-table-fields"); assert(grid && grid->columnCount()==4);
            grid->item(0,0)->setText("qty"); qobject_cast<QComboBox*>(grid->cellWidget(0,1))->setCurrentText("INTEGER");
            grid->item(0,2)->setText("ja"); grid->item(0,3)->setText("Now()");
            auto* buttons=dialog->findChild<QDialogButtonBox*>(); assert(buttons);
            QTimer::singleShot(0,dialog,[] { auto* error=qobject_cast<QMessageBox*>(QApplication::activeModalWidget()); assert(error); error->accept(); });
            buttons->button(QDialogButtonBox::Save)->click();
            assert(dialog->isVisible() && grid->item(0,0)->text()=="qty" && !db.tables().contains("authored_defaults"));
            grid->item(0,3)->setText("7"); dialog->findChild<QPushButton*>("native-table-add-field")->click();
            grid->item(1,0)->setText("label"); grid->item(1,3)->setText("\"ui value\"");
            dialog->findChild<QPushButton*>("native-table-add-field")->click();
            grid->item(2,0)->setText("enabled"); qobject_cast<QComboBox*>(grid->cellWidget(2,1))->setCurrentText("BOOLEAN"); grid->item(2,3)->setText("True");
            buttons->button(QDialogButtonBox::Save)->click();
        });
        action(window,"+ Tabelle")->trigger();
        QSqlQuery authored(db); assert(authored.exec("INSERT INTO authored_defaults DEFAULT VALUES"));
        assert(authored.exec("SELECT qty,label,enabled FROM authored_defaults") && authored.next());
        assert(authored.value(0).toInt()==7 && authored.value(1).toString()=="ui value" && authored.value(2).toInt()==-1); authored.finish();
    }
    assert(window.close()); dispose(); watchdog.stop();
    std::cout<<"Main window: native project/explorer opening, independent nonmodal form/report lifetime, deduplication, modal metadata, cancellation, SQL save and retained close/MsgBox passed\n";
}
