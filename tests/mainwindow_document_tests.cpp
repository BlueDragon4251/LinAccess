#include "native_project_gui_extensions.hpp"
#include "native_macro_storage.hpp"
#include "native_access_cursor.hpp"
#include "native_access_warnings.hpp"
#include <QStatusBar>
#include <QProgressBar>
#include "native_standard_module_host.hpp"
#include <QTextBrowser>
#include "native_schema_editor.hpp"
#include "native_macro.hpp"
#include "native_module_definition.hpp"
#include "vba_module_context.hpp"
#include "native_stored_form.hpp"
#include "native_stored_report.hpp"
#include <QDateTime>
#include <QLocale>
#include <QScopeGuard>
#include <bit>
#include <limits>
#include "native_relation_designer.hpp"
#include "native_datasheet_page_store.hpp"
#include <QInputDialog>
#include "mainwindow.hpp"
#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_table_view.hpp"
#include "native_query_view.hpp"
#include "access_query.hpp"
#include "native_datasheet_print.hpp"
#include "access_table_import.hpp"
#include "access_autonumber.hpp"
#include "dao_database.hpp"
#include "native_relation_designer.hpp"
#include "access_catalog.hpp"
#include <QProcess>
#include <QPrintDialog>
#include <QPrintPreviewWidget>
#include <QPrinter>
#include <QRegularExpression>
#include <QAction>
#include <QApplication>
#include <QCryptographicHash>
#include <QClipboard>
#include <QCheckBox>
#include <QStyleOptionButton>
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
#include <chrono>
#include <cstdlib>
#include <iostream>

struct sqlite3;
struct sqlite3_context;
struct sqlite3_value;

void exerciseStoredForm(MainWindow& window,QSqlDatabase db);
void exerciseVisualQueryDesigner(MainWindow& window,QSqlDatabase db);
void exerciseNativeDebugOutput(MainWindow& window,QSqlDatabase db);
void exerciseNativeProperties(QSqlDatabase db);
void exerciseNativeStaticProcedures(QSqlDatabase db);
void exerciseNativeClasses(MainWindow& window,QSqlDatabase db);

namespace {
QAction* action(MainWindow& window,const QString& text) {
    for(auto* candidate:window.findChildren<QAction*>()) if(candidate->text()==text) return candidate;
    assert(false); return nullptr;
}
void openProject(MainWindow& window,const QString& path,QWidget* shortcutTarget=nullptr,bool holdShift=false) {
    bool selected=false;
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QFileDialog*>(QApplication::activeModalWidget()); assert(dialog);
        dialog->selectFile(path); selected=true;
        if(holdShift) QTest::keyPress(dialog,Qt::Key_Shift,Qt::ShiftModifier);
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

void exerciseDatasheetPageSetup(MainWindow& window,QSqlDatabase db) {
    QTemporaryDir directory; assert(directory.isValid()); QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE page_layout_rows(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(sql.exec("INSERT INTO page_layout_rows VALUES(1,'geometry')")); sql.finish();
    const auto open=[&] { assert(window.openTableDocument("page_layout_rows",0,1)==0); window.activateWindow(); QApplication::processEvents(); };
    const auto close=[&] { assert(window.closeTableDocument(0,"page_layout_rows",2).value("error").toInt()==0); QApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete); };
    const auto accept=[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="native-datasheet-page-setup");
        assert(!window.allowProjectClose(db.connectionName()) && window.setupDatasheetPage().value("error").toInt()==2046);
        dialog->findChild<QComboBox*>("report-page-orientation")->setCurrentIndex(1);
        dialog->findChild<QDoubleSpinBox*>("report-page-left")->setValue(25);
        auto* save=dialog->findChild<QCheckBox*>("report-page-save"); assert(save && save->isEnabled()); save->setChecked(true);
        dialog->findChild<QDialogButtonBox*>()->button(QDialogButtonBox::Ok)->click();
    };
    open(); QTimer::singleShot(0,&window,accept); assert(window.setupDatasheetPage().value("error").toInt()==0);
    const auto path=directory.filePath("geometry.pdf");
    const auto read=[&] { QFile file(path); assert(file.open(QIODevice::ReadOnly)); return file.readAll(); };
    assert(window.outputDatasheet(true,path).value("error").toInt()==0);
    const QRegularExpression media("/MediaBox\\s*\\[\\s*0\\s+0\\s+([0-9.]+)\\s+([0-9.]+)");
    auto match=media.match(QString::fromLatin1(read())); assert(match.hasMatch() && match.captured(1).toDouble()>match.captured(2).toDouble());
    const auto reader=QStandardPaths::findExecutable("pdftotext");
    if(!reader.isEmpty()) {
        QProcess text; text.start(reader,{"-bbox",path,"-"}); assert(text.waitForFinished(10000) && text.exitCode()==0);
        const auto position=QRegularExpression("<word xMin=\"([0-9.]+)\"").match(QString::fromUtf8(text.readAllStandardOutput()));
        assert(position.hasMatch() && position.captured(1).toDouble()>70 && position.captured(1).toDouble()<90);
    }
    assert(db.transaction()); assert(sql.exec("UPDATE page_layout_rows SET value='pending'"));
    QTimer::singleShot(0,&window,accept); assert(window.setupDatasheetPage().value("error").toInt()==0);
    assert(db.rollback()); assert(sql.exec("SELECT value FROM page_layout_rows") && sql.next() && sql.value(0).toString()=="geometry"); sql.finish();
    close(); open();
    assert(sql.exec("CREATE TEMP TABLE _la_datasheet_page_setup(incorrect TEXT)"));
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog);
        assert(dialog->findChild<QComboBox*>("report-page-orientation")->currentIndex()==1);
        assert(dialog->findChild<QDoubleSpinBox*>("report-page-left")->value()==25);
        dialog->findChild<QComboBox*>("report-page-orientation")->setCurrentIndex(0); dialog->reject();
    });
    assert(window.setupDatasheetPage().value("error").toInt()==2501);
    assert(window.outputDatasheet(true,path).value("error").toInt()==0);
    match=media.match(QString::fromLatin1(read())); assert(match.hasMatch() && match.captured(1).toDouble()>match.captured(2).toDouble());
    assert(sql.exec("DROP TABLE temp._la_datasheet_page_setup"));
    assert(sql.exec("SELECT layout FROM main._la_datasheet_page_setup WHERE kind='table' AND name='page_layout_rows'") && sql.next());
    auto changed=QJsonDocument::fromJson(sql.value(0).toString().toUtf8()).object(); sql.finish();
    auto edges=changed.value("margins_mm").toArray(); edges[0]=30; changed.insert("margins_mm",edges);
    const auto external=QString::fromUtf8(QJsonDocument(changed).toJson(QJsonDocument::Compact));
    QTimer::singleShot(0,&window,[&] {
        sql.prepare("UPDATE main._la_datasheet_page_setup SET layout=? WHERE kind='table' AND name='page_layout_rows'"); sql.addBindValue(external); assert(sql.exec()); sql.finish(); accept();
    });
    assert(window.setupDatasheetPage().value("error").toInt()==3251);
    assert(sql.exec("SELECT layout FROM main._la_datasheet_page_setup WHERE kind='table' AND name='page_layout_rows'") && sql.next() && sql.value(0).toString()==external); sql.finish();
    close();
    assert(sql.exec("UPDATE main._la_datasheet_page_setup SET layout='{invalid' WHERE kind='table' AND name='page_layout_rows'"));
    open(); const auto prior=read(); assert(window.outputDatasheet(true,path).value("error").toInt()==3251 && read()==prior);
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->windowTitle().contains("reparieren")); accept();
    });
    assert(window.setupDatasheetPage().value("error").toInt()==0 && window.outputDatasheet(true,path).value("error").toInt()==0);
    assert(sql.exec("SELECT COUNT(*) FROM page_layout_rows") && sql.next() && sql.value(0).toInt()==1); sql.finish(); close();
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
    bool copiesRejected=false;
    try { linaccess::exportNativeDatasheetPdf(snapshot,path,QPageLayout(QPageSize(QPageSize::A4),QPageLayout::Portrait,QMarginsF(15,15,15,15)),1,2,false,101); }
    catch(const linaccess::VbaRuntimeError& error) { copiesRejected=error.number==3251; }
    assert(copiesRejected && read()==original);
    for(const bool collated:{true,false}) {
        const auto copiesPath=output.filePath(collated ? "collated.pdf" : "uncollated.pdf");
        QTimer::singleShot(0,&window,[&] {
            auto* dialog=qobject_cast<QPrintDialog*>(QApplication::activeModalWidget()); assert(dialog);
            auto* printer=dialog->printer(); printer->setOutputFormat(QPrinter::PdfFormat); printer->setOutputFileName(copiesPath);
            printer->setPrintRange(QPrinter::PageRange); printer->setFromTo(1,2);
            printer->setCopyCount(2); printer->setCollateCopies(collated);
            printer->setPageOrder(collated ? QPrinter::FirstPageFirst : QPrinter::LastPageFirst); dialog->QDialog::done(QDialog::Accepted);
        });
        assert(window.outputDatasheet(false).value("error").toInt()==0);
        QFile file(copiesPath); assert(file.open(QIODevice::ReadOnly)); auto found=pages.globalMatch(QString::fromLatin1(file.readAll()));
        int actual=0; while(found.hasNext()) { found.next(); ++actual; } assert(actual==4);
        if(!reader.isEmpty()) {
            QProcess text; text.start(reader,{"-layout",copiesPath,"-"}); assert(text.waitForFinished(10000) && text.exitCode()==0);
            const auto leaves=QString::fromUtf8(text.readAllStandardOutput()).split(QChar(12),Qt::SkipEmptyParts); assert(leaves.size()==4);
            assert(leaves[0].contains("row_320")==collated && leaves[0]!=leaves[collated ? 1 : 2]);
            if(collated) assert(leaves[0]==leaves[2] && leaves[1]==leaves[3]);
            else assert(leaves[0]==leaves[1] && leaves[2]==leaves[3]);
        }
    }
    const auto selectedPdf=output.filePath("selection.pdf");
    view->selectionModel()->select(model->index(0,0),QItemSelectionModel::ClearAndSelect|QItemSelectionModel::Rows);
    view->selectionModel()->select(model->index(2,0),QItemSelectionModel::Select|QItemSelectionModel::Rows);
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QPrintDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->testOption(QAbstractPrintDialog::PrintSelection));
        auto* printer=dialog->printer(); printer->setOutputFormat(QPrinter::PdfFormat); printer->setOutputFileName(selectedPdf);
        printer->setCopyCount(1); printer->setPrintRange(QPrinter::Selection);
        view->clearSelection(); view->selectRow(1); dialog->QDialog::done(QDialog::Accepted);
    });
    assert(window.outputDatasheet(false).value("error").toInt()==0);
    if(!reader.isEmpty()) {
        QProcess text; text.start(reader,{"-layout",selectedPdf,"-"}); assert(text.waitForFinished(10000) && text.exitCode()==0);
        const auto content=text.readAllStandardOutput(); assert(content.contains("row_320") && content.contains("row_318") && !content.contains("row_319") && !content.contains("row_317"));
    }
    for(const bool partial:{false,true}) {
        view->clearSelection();
        if(partial) view->selectionModel()->select(model->index(1,0),QItemSelectionModel::ClearAndSelect);
        QTimer::singleShot(0,&window,[&] {
            auto* dialog=qobject_cast<QPrintDialog*>(QApplication::activeModalWidget()); assert(dialog && !dialog->testOption(QAbstractPrintDialog::PrintSelection));
            auto* printer=dialog->printer(); printer->setOutputFormat(QPrinter::PdfFormat); printer->setOutputFileName(path);
            printer->setPrintRange(QPrinter::Selection); dialog->QDialog::done(QDialog::Accepted);
        });
        assert(window.outputDatasheet(false).value("error").toInt()==3251 && read()==original);
    }
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
Sub SetupData()
DoCmd.SelectObject acTable, "output_rows", False
Application.RunCommand Command:=acCmdPageSetup
Me.Tag = "configured"
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
        QTimer::singleShot(0,&window,[&] {
            auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="native-datasheet-page-setup");
            dialog->findChild<QDialogButtonBox*>()->button(QDialogButtonBox::Ok)->click();
        });
        const auto configured=caller.executeProcedure("SetupData",arguments);
        assert(configured.status==linaccess::VbaExecutionStatus::Completed && makePreviewFormObject(caller)->member("Tag",{}).asText()=="configured");
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
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="native-datasheet-page-setup");
        dialog->findChild<QComboBox*>("report-page-orientation")->setCurrentIndex(1);
        dialog->findChild<QCheckBox*>("report-page-save")->setChecked(true);
        dialog->findChild<QDialogButtonBox*>()->button(QDialogButtonBox::Ok)->click();
    });
    assert(window.setupDatasheetPage().value("error").toInt()==0);
    assert(window.outputDatasheet(true,output.filePath("query.pdf")).value("error").toInt()==0);
    queryView->selectRow(1);
    const auto querySelection=output.filePath("query-selection.pdf");
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QPrintDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->testOption(QAbstractPrintDialog::PrintSelection));
        auto* printer=dialog->printer(); printer->setOutputFormat(QPrinter::PdfFormat); printer->setOutputFileName(querySelection);
        printer->setPrintRange(QPrinter::Selection); dialog->QDialog::done(QDialog::Accepted);
    });
    assert(window.outputDatasheet(false).value("error").toInt()==0);
    if(!reader.isEmpty()) {
        QProcess text; text.start(reader,{"-layout",querySelection,"-"}); assert(text.waitForFinished(10000) && text.exitCode()==0);
        const auto content=text.readAllStandardOutput(); assert(content.contains("row_010") && !content.contains("row_020"));
    }
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
    assert(sought && replacement && apply && next && match && match->currentIndex()==1 && match->isEnabled());
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
    assert(sql.exec("CREATE TABLE command_replace_parts(id INTEGER PRIMARY KEY,value TEXT CHECK(id<>2 OR value<>'blocked2'))"));
    assert(sql.exec("INSERT INTO command_replace_parts VALUES(1,'prefix-aa-aa-aa'),(2,'prefix-AA'),(3,'other'),(4,NULL)"));
    assert(window.openTableDocument("command_replace_parts",0,1)==0);view=nullptr;
    for(auto* candidate:tabs->currentWidget()->findChildren<QTableView*>()) if(auto* typed=dynamic_cast<linaccess::AccessTableView*>(candidate)) view=typed;
    assert(view);model=qobject_cast<linaccess::AccessTableModel*>(view->model());assert(model);view->sortByColumn(0,Qt::AscendingOrder);view->setCurrentIndex(model->index(0,1));
    assert(view->openReplaceDialog());dialog=view->findChild<QDialog*>("native-table-replace-dialog");assert(dialog);
    sought=dialog->findChild<QLineEdit*>("native-table-find-text");replacement=dialog->findChild<QLineEdit*>("native-table-replace-with");match=dialog->findChild<QComboBox*>("native-table-find-match");all=dialog->findChild<QPushButton*>("native-table-replace-all");apply=dialog->findChild<QPushButton*>("native-table-replace-current");
    match->setCurrentIndex(0);sought->setText("aa");replacement->setText("X");all->click();assert(!model->isDirty());
    assert(sql.exec("SELECT value FROM command_replace_parts ORDER BY id") && sql.next() && sql.value(0).toString()=="prefix-X-X-X" && sql.next() && sql.value(0).toString()=="prefix-X");sql.finish();
    match->setCurrentIndex(2);sought->setText("prefix-");replacement->clear();apply->click();assert(model->isDirty() && model->index(0,1).data(Qt::EditRole).toString()=="X-X-X");
    assert(sql.exec("SELECT value FROM command_replace_parts WHERE id=1") && sql.next() && sql.value(0).toString()=="prefix-X-X-X");sql.finish();assert(window.dispatchTableCommand(97).value("error").toInt()==0);
    match->setCurrentIndex(0);sought->setText("-");replacement->setText(".");apply->click();assert(model->isDirty() && model->index(0,1).data(Qt::EditRole).toString()=="X.X-X");assert(window.dispatchTableCommand(97).value("error").toInt()==0);
    match->setCurrentIndex(2);sought->setText("prefix-");replacement->setText("start-");all->click();assert(!model->isDirty());
    assert(sql.exec("SELECT value FROM command_replace_parts WHERE id=2") && sql.next() && sql.value(0).toString()=="start-X");sql.finish();
    assert(view->replaceAllText({"other",1,0,false,0,2},QString{}));assert(sql.exec("SELECT value IS NULL,length(value) FROM command_replace_parts WHERE id=3") && sql.next() && !sql.value(0).toBool() && sql.value(1).toInt()==0);sql.finish();
    assert(sql.exec("UPDATE command_replace_parts SET value='seed'||id WHERE id IN(1,2)"));assert(model->select());view->setCurrentIndex(model->index(0,1));
    assert(!view->replaceAllText({"seed",1,0,false,0,2},"blocked") && !model->isDirty());
    assert(sql.exec("SELECT value FROM command_replace_parts WHERE id IN(1,2) ORDER BY id") && sql.next() && sql.value(0).toString()=="seed1" && sql.next() && sql.value(0).toString()=="seed2");sql.finish();
    assert(!view->replaceAllText({"seed",1,0,false,0,2},QString("x")+QChar(0)+"y") && model->lastError().nativeErrorCode()=="3251" && !model->isDirty());
    dialog->reject();dispose();assert(window.closeTableDocument(0,"command_replace_parts",2).value("error").toInt()==0);

    assert(sql.exec("CREATE TABLE command_replace_fields(id TEXT PRIMARY KEY,a TEXT,b TEXT CHECK(id<>'seed-two' OR b<>'blocked2'),hidden TEXT,n INTEGER,derived TEXT GENERATED ALWAYS AS(a||'-calculated'),raw BLOB)"));
    assert(sql.exec("INSERT INTO command_replace_fields(id,a,b,hidden,n,raw) VALUES('seed-one','seed1','SEED1','seed-hidden',7,X'0102'),('seed-two','seed2','SEED2','seed-hidden',8,X'0304'),('third',NULL,'other','seed-hidden',9,NULL)"));
    assert(window.openTableDocument("command_replace_fields",0,1)==0);view=nullptr;
    for(auto* candidate:tabs->currentWidget()->findChildren<QTableView*>()) if(auto* typed=dynamic_cast<linaccess::AccessTableView*>(candidate)) view=typed;
    assert(view);model=qobject_cast<linaccess::AccessTableModel*>(view->model());assert(model);view->hideColumn(3);view->sortByColumn(1,Qt::DescendingOrder);
    int selectedRow=-1;for(int row=0;row<model->rowCount();++row) if(model->record(row).value("id").toString()=="seed-two") selectedRow=row;assert(selectedRow>=0);view->setCurrentIndex(model->index(selectedRow,1));
    assert(view->openReplaceDialog());dialog=view->findChild<QDialog*>("native-table-replace-dialog");assert(dialog);
    auto* scope=dialog->findChild<QComboBox*>("native-table-find-field");assert(scope && scope->findData(-1)>=0);scope->setCurrentIndex(scope->findData(-1));
    sought=dialog->findChild<QLineEdit*>("native-table-find-text");replacement=dialog->findChild<QLineEdit*>("native-table-replace-with");match=dialog->findChild<QComboBox*>("native-table-find-match");all=dialog->findChild<QPushButton*>("native-table-replace-all");apply=dialog->findChild<QPushButton*>("native-table-replace-current");
    auto* feedback=dialog->findChild<QLabel*>("native-table-find-result");assert(feedback);
    sought->setText("seed");replacement->setText("done");match->setCurrentIndex(0);all->click();assert(!model->isDirty() && feedback->text().startsWith("4 Feldwerte"));assert(model->record(view->currentIndex().row()).value("id").toString()=="seed-two" && view->currentIndex().column()==1);
    assert(sql.exec("SELECT id,a,b,hidden,n,derived,hex(raw) FROM command_replace_fields WHERE id='seed-two'") && sql.next() && sql.value(1).toString()=="done2" && sql.value(2).toString()=="done2" && sql.value(3).toString()=="seed-hidden" && sql.value(4).toInt()==8 && sql.value(5).toString()=="done2-calculated" && sql.value(6).toString()=="0304");sql.finish();
    assert(sql.exec("SELECT a IS NULL,raw IS NULL FROM command_replace_fields WHERE id='third'") && sql.next() && sql.value(0).toBool() && sql.value(1).toBool());sql.finish();
    sought->setText("done");replacement->setText("draft");apply->click();assert(model->isDirty() && model->index(view->currentIndex().row(),1).data(Qt::EditRole).toString()=="draft2");model->revertAll();
    assert(sql.exec("UPDATE command_replace_fields SET a=CASE id WHEN 'seed-one' THEN 'seed1' ELSE 'seed2' END,b=CASE id WHEN 'seed-one' THEN 'seed1' ELSE 'seed2' END WHERE id IN('seed-one','seed-two')"));assert(model->select());view->setCurrentIndex(model->index(0,1));
    sought->setText("seed");replacement->setText("blocked");all->click();assert(!model->isDirty());
    assert(sql.exec("SELECT COUNT(*) FROM command_replace_fields WHERE a LIKE 'blocked%' OR b LIKE 'blocked%'") && sql.next() && sql.value(0).toInt()==0);sql.finish();
    assert(sql.exec("SELECT a,b FROM command_replace_fields WHERE id='seed-one'") && sql.next() && sql.value(0).toString()=="seed1" && sql.value(1).toString()=="seed1");sql.finish();
    assert(model->setData(model->index(1,4),99,Qt::EditRole));replacement->setText("unintended");all->click();assert(model->isDirty() && model->index(1,4).data(Qt::EditRole).toInt()==99);model->revertAll();
    view->setRecordNavigationGuard([view]{view->hideColumn(2);return true;});assert(!view->replaceAllText({"seed",-1,0,false,0,2},"unexpected") && model->lastError().nativeErrorCode()=="2501" && !model->isDirty());view->showColumn(2);view->setRecordNavigationGuard({});
    assert(sql.exec("SELECT COUNT(*) FROM command_replace_fields WHERE a LIKE 'unexpected%' OR b LIKE 'unexpected%' OR a LIKE 'unintended%' OR b LIKE 'unintended%'") && sql.next() && sql.value(0).toInt()==0);sql.finish();
    int untouchedCount=9;assert(!model->replaceAllLiteralTextColumns({1,1},"seed","x",0,Qt::CaseInsensitive,&untouchedCount) && untouchedCount==0 && !model->isDirty());
    dialog->reject();dispose();assert(window.closeTableDocument(0,"command_replace_fields",2).value("error").toInt()==0);

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
    assert(!model->stageRecordDeletions({0,0}) && !model->isDirty());
    assert(!model->stageRecordDeletions({0,3}) && !model->isDirty());
    assert(model->setReadOnly(true) && !model->stageRecordDeletions({1}) && !model->isDirty() && model->setReadOnly(false));
    view->clearSelection();view->setCurrentIndex(model->index(1,1));
    QTest::keyClick(view,Qt::Key_Minus,Qt::ControlModifier);
    assert(model->canUndoBlock() && model->hasPendingRecord(1) && !model->hasPendingRecord(0) && !model->hasPendingRecord(2));
    QTest::keyClick(view,Qt::Key_Z,Qt::ControlModifier);
    assert(!model->isDirty() && model->canRedoBlock() && model->index(1,1).data().toString()=="two");
    QTest::keyClick(view,Qt::Key_Y,Qt::ControlModifier);
    assert(model->hasPendingRecord(1) && model->undoBlock() && !model->isDirty());
    assert(model->setData(model->index(1,1),"changed") && !model->canRedoBlock() && !model->canUndoBlock());
    assert(model->undoPendingRecord(1) && !model->isDirty() && model->index(1,1).data().toString()=="two");
    assert(!model->canRedoBlock());
    const auto selectRecords=[&] {
        QItemSelection rows; rows.select(model->index(0,0),model->index(0,1)); rows.select(model->index(2,0),model->index(2,1));
        view->selectionModel()->select(rows,QItemSelectionModel::ClearAndSelect|QItemSelectionModel::Rows);
        view->selectionModel()->setCurrentIndex(model->index(0,1),QItemSelectionModel::NoUpdate);
    };
    selectRecords(); QTest::keyClick(view,Qt::Key_Minus,Qt::ControlModifier); assert(model->isDirty());
    assert(sql.exec("SELECT COUNT(*) FROM delete_many") && sql.next() && sql.value(0).toInt()==3); sql.finish();
    assert(model->canUndoBlock() && !model->canRedoBlock());
    QTest::keyClick(view,Qt::Key_Z,Qt::ControlModifier);
    assert(!model->isDirty() && !model->canUndoBlock() && model->canRedoBlock());
    assert(model->rowCount()==3 && model->index(0,0).data().toInt()==1 && model->index(2,0).data().toInt()==3);
    QTest::keyClick(view,Qt::Key_Y,Qt::ControlModifier);
    assert(model->isDirty() && model->hasPendingRecord(0) && !model->hasPendingRecord(1) && model->hasPendingRecord(2));
    assert(window.dispatchTableCommand(97).value("error").toInt()==3146);
    assert(sql.exec("SELECT COUNT(*) FROM delete_many") && sql.next() && sql.value(0).toInt()==3); sql.finish();
    assert(model->canUndoBlock() && model->undoBlock() && !model->isDirty());
    assert(model->redoBlock() && model->hasPendingRecord(0) && model->hasPendingRecord(2));
    model->revertAll(); assert(!model->isDirty());
    assert(sql.exec("DELETE FROM delete_many_child")); sql.finish();
    selectRecords(); QTest::keyClick(view,Qt::Key_Minus,Qt::ControlModifier);
    assert(window.dispatchTableCommand(97).value("error").toInt()==0 && !model->isDirty());
    assert(sql.exec("SELECT id FROM delete_many") && sql.next() && sql.value(0).toInt()==2 && !sql.next()); sql.finish();
    assert(window.closeTableDocument(0,"delete_many",2).value("error").toInt()==0); dispose();
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
        dialog->findChild<QLineEdit*>("native-report-running-field")->setText("id");
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
    assert(saved.value("running_sum_field").toString()=="id");
    {const auto packet=nativeStoredReport("validation",saved,db);const auto layout=parseAccessFormLayout(QString::fromUtf8(packet.layout));bool found=false;for(const auto& control:layout.controls) if(control.name=="id") {assert(control.runningSum==2 && !control.canGrow);found=true;}assert(found);
        for(const auto& field:{QString("missing"),QString("Display value")}) {auto bad=saved;bad.insert("running_sum_field",field);int code=0;try {(void)nativeStoredReport("invalid",bad,db);}catch(const VbaRuntimeError& error){code=error.number;}assert(code==(field=="missing" ? 3265 : 3251));}}
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
        assert(dialog->findChild<QLineEdit*>("native-report-running-field")->text()=="id");
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
    assert(finalDesign.value("running_sum_field").toString()=="id");
    assert(sql.exec("SELECT COUNT(*) FROM native_report_events WHERE kind='open'") && sql.next() && sql.value(0).toInt()==5); sql.finish();
    assert(sql.exec("SELECT COUNT(*) FROM native_report_events WHERE kind='close'") && sql.next() && sql.value(0).toInt()==5); sql.finish();
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog);
        dialog->findChild<QLineEdit*>("native-report-name")->setText("Owned Summary Report");
        dialog->findChild<QPlainTextEdit*>("native-report-sql")->setPlainText("SELECT id FROM native_report_rows ORDER BY id");
        dialog->findChild<QLineEdit*>("native-report-footer-expression")->setText("=Sum([id])");
        QTimer::singleShot(0,dialog,[&] {auto* editor=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(editor && editor->objectName()=="native-report-group-designer");auto* list=editor->findChild<QTableWidget*>("native-report-groups-list");assert(list && list->rowCount()==0);editor->findChild<QPushButton*>("native-report-group-add")->click();assert(list->rowCount()==1);qobject_cast<QComboBox*>(list->cellWidget(0,0))->setCurrentText("id");qobject_cast<QLineEdit*>(list->cellWidget(0,3))->setText("=Count(*)");qobject_cast<QCheckBox*>(list->cellWidget(0,4))->setChecked(true);editor->findChild<QDialogButtonBox*>("native-report-group-buttons")->button(QDialogButtonBox::Ok)->click();});
        dialog->findChild<QPushButton*>("native-report-group-design")->click();
        dialog->findChild<QLineEdit*>("native-report-running-field")->setText("id");
        auto* runningMode=dialog->findChild<QComboBox*>("native-report-running-mode");assert(runningMode);runningMode->setCurrentIndex(runningMode->findData(1));

        dialog->findChild<QComboBox*>("native-report-footer-break")->setCurrentIndex(0);
        dialog->findChild<QComboBox*>("native-report-page-header-mode")->setCurrentIndex(3);
        dialog->findChild<QComboBox*>("native-report-page-footer-mode")->setCurrentIndex(2);
        dialog->findChild<QPushButton*>("native-report-save")->click();
    });
    create->trigger(); QApplication::processEvents();
    assert(sql.exec("SELECT definition FROM main._la_objects WHERE kind='report' AND name='Owned Summary Report'") && sql.next());
    const auto summaryDesign=QJsonDocument::fromJson(sql.value(0).toByteArray()).object(); sql.finish();
    assert(summaryDesign.value("groups").toArray()[0].toObject().value("repeat_header").toBool());
    assert(summaryDesign.value("running_sum_mode").toInt()==1 && summaryDesign.value("running_sum_field").toString()=="id");
    assert(summaryDesign.value("groups").toArray().size()==1 && summaryDesign.value("groups").toArray()[0].toObject().value("footer_expression").toString()=="=Count(*)");
    assert(summaryDesign.value("footer_expression").toString()=="=Sum([id])" && summaryDesign.value("footer_force_new_page").toInt()==0 && summaryDesign.value("page_header").toInt()==3 && summaryDesign.value("page_footer").toInt()==2);
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
        assert(dialog->findChild<QComboBox*>("native-report-running-mode")->currentData().toInt()==1 && dialog->findChild<QLineEdit*>("native-report-running-field")->text()=="id");
        QTimer::singleShot(0,dialog,[&] {auto* editor=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(editor && editor->objectName()=="native-report-group-designer");auto* list=editor->findChild<QTableWidget*>("native-report-groups-list");assert(list && list->rowCount()==1 && qobject_cast<QComboBox*>(list->cellWidget(0,0))->currentText()=="id" && qobject_cast<QCheckBox*>(list->cellWidget(0,2))->isChecked() && qobject_cast<QLineEdit*>(list->cellWidget(0,3))->text()=="=Count(*)" && qobject_cast<QCheckBox*>(list->cellWidget(0,4))->isChecked());editor->reject();});
        dialog->findChild<QPushButton*>("native-report-group-design")->click();

        assert(dialog->findChild<QComboBox*>("native-report-footer-break")->currentIndex()==0 && dialog->findChild<QComboBox*>("native-report-page-header-mode")->currentIndex()==3 && dialog->findChild<QComboBox*>("native-report-page-footer-mode")->currentIndex()==2);
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
    assert(window.dispatchTableCommand(28).value("error").toInt()==0 && model->rowCount()==3 && model->hasPendingRecord(1));
    assert(sql.exec("SELECT COUNT(*) FROM command_record_undo") && sql.next() && sql.value(0).toInt()==2);sql.finish();
    assert(model->setData(model->index(2,1),"new-draft"));
    assert(window.dispatchTableCommand(28).value("error").toInt()==3251 && model->rowCount()==3 && model->hasPendingRecord(1) && model->hasPendingRecord(2));
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
    view->setCurrentIndex(model->index(0,1));
    assert(model->setData(model->index(0,1),"new-row-saved-current") && model->setData(model->index(1,1),"new-row-other-draft"));
    assert(window.dispatchTableCommand(28).value("error").toInt()==0 && model->rowCount()==3 && view->currentIndex().row()==2 && !model->hasPendingRecord(0) && model->hasPendingRecord(1));
    assert(model->index(1,1).data(Qt::EditRole).toString()=="new-row-other-draft");
    assert(sql.exec("SELECT value FROM command_record_undo WHERE id=1") && sql.next() && sql.value(0).toString()=="new-row-saved-current");sql.finish();
    assert(sql.exec("SELECT value FROM command_record_undo WHERE id=2") && sql.next() && sql.value(0).toString()=="saved-by-ctrl-home");sql.finish();
    view->setFocus();QTest::keyClick(view,Qt::Key_Escape);assert(model->rowCount()==2 && model->hasPendingRecord(1));
    view->setCurrentIndex(model->index(1,1));QTest::keyClick(view,Qt::Key_Escape);assert(!model->isDirty());
    assert(window.closeTableDocument(0,"command_record_undo",2).value("error").toInt()==0); dispose();
}
void exercisePageNavigation(MainWindow& window,QSqlDatabase db) {
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE command_page_navigation(id INTEGER PRIMARY KEY,value TEXT UNIQUE)"));
    assert(sql.exec("WITH RECURSIVE n(i) AS (VALUES(1) UNION ALL SELECT i+1 FROM n WHERE i<600) INSERT INTO command_page_navigation SELECT i,'page-row-'||i FROM n"));
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
    assert(model->canFetchMore() && model->setData(model->index(0,1),"unrelated-page-draft"));
    const int loadedBoundary=model->rowCount();view->setCurrentIndex(model->index(loadedBoundary-2,1));
    QTest::keyClick(view,Qt::Key_PageDown);
    assert(view->currentIndex().row()>=loadedBoundary && model->rowCount()>loadedBoundary && model->hasPendingRecord(0));
    assert(model->index(0,1).data(Qt::EditRole).toString()=="unrelated-page-draft");
    assert(sql.exec("SELECT value FROM command_page_navigation WHERE id=1") && sql.next() && sql.value(0).toString()=="saved-by-page-down");sql.finish();
    assert(model->undoPendingRecord(0) && !model->isDirty());
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
Me.Tag="findrecord-all-raw"
End Sub
Sub FindRecordAllNumber()
DoCmd.FindRecord 42,Match:=1,OnlyCurrentField:=acAll
Me.Tag="findrecord-all-number"
End Sub
Sub FindRecordFormatted()
DoCmd.FindRecord "row-250",SearchAsFormatted:=True
Me.Tag="findrecord-formatted"
End Sub
Sub FindRecordAllFormatted()
DoCmd.FindRecord "row-250",SearchAsFormatted:=True,OnlyCurrentField:=acAll
Me.Tag="findrecord-all-formatted"
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
DoCmd.FindRecord "row-2[45]#",1,False,2,False,-1,True
Me.Tag="wildcard-found"
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
Sub AppendTableSelection()
DoCmd.RunCommand acCmdPasteAppend
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

    assert(unchanged.exec("CREATE TABLE command_alias_sort_rows(id INTEGER PRIMARY KEY,value TEXT,active INTEGER,priority INTEGER)"));assert(unchanged.exec("INSERT INTO command_alias_sort_rows VALUES(1,'alpha',1,1),(2,'zulu',1,2),(3,'excluded',0,9),(4,'beta',1,2)"));
    unchanged.prepare("INSERT INTO _la_objects(kind,name,definition) VALUES('query','command_alias_sort_query',?)");unchanged.addBindValue(QString::fromUtf8(QJsonDocument(QJsonObject{{"sql","SELECT sourceRow.* FROM command_alias_sort_rows AS sourceRow WHERE sourceRow.active=1 ORDER BY sourceRow.priority DESC,sourceRow.value ASC"}}).toJson(QJsonDocument::Compact)));assert(unchanged.exec());unchanged.finish();
    assert(window.openQueryDocument("command_alias_sort_query",0,1).value("error").toInt()==0);auto* aliasView=documentTabs->currentWidget()->findChild<QTableView*>("native-editable-query-view");assert(aliasView);auto* aliasModel=qobject_cast<linaccess::AccessTableModel*>(aliasView->model());assert(aliasModel && !aliasModel->isReadOnly() && aliasModel->rowCount()==3 && aliasModel->record(0).value("id").toInt()==4 && aliasModel->record(1).value("id").toInt()==2 && aliasView->horizontalHeader()->sortIndicatorSection()==3 && aliasView->horizontalHeader()->sortIndicatorOrder()==Qt::DescendingOrder);
    aliasView->setCurrentIndex(aliasModel->index(1,1));assert(aliasModel->setData(aliasModel->index(1,1),"aardvark"));assert(mixed->executeProcedure("SaveActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed && !aliasModel->isDirty());
    assert(aliasModel->record(0).value("id").toInt()==2 && aliasModel->record(1).value("id").toInt()==4 && aliasModel->record(aliasView->currentIndex().row()).value("id").toInt()==2);assert(unchanged.exec("SELECT value FROM command_alias_sort_rows WHERE id=2") && unchanged.next() && unchanged.value(0).toString()=="aardvark");unchanged.finish();
    assert(aliasModel->setData(aliasModel->index(aliasView->currentIndex().row(),2),0));assert(mixed->executeProcedure("SaveActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed && aliasModel->rowCount()==2 && aliasModel->record(0).value("id").toInt()==4 && !aliasModel->isDirty());
    aliasView->sortByColumn(1,Qt::AscendingOrder);assert(aliasModel->record(0).value("id").toInt()==1);assert(mixed->executeProcedure("RequeryActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed && aliasModel->rowCount()==2 && aliasModel->record(0).value("id").toInt()==1);assert(unchanged.exec("SELECT active FROM command_alias_sort_rows WHERE id=2") && unchanged.next() && unchanged.value(0).toInt()==0);unchanged.finish();assert(window.closeQueryDocument(1,"command_alias_sort_query",0).value("error").toInt()==0);QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete);
    assert(unchanged.exec("CREATE TABLE command_projection_rows(id INTEGER PRIMARY KEY,value TEXT,hidden TEXT,note TEXT,priority INTEGER)"));assert(unchanged.exec("INSERT INTO command_projection_rows VALUES(1,'match','keep1','match',2),(2,'alpha','keep2','second',1),(3,'excluded','keep3','third',0)"));
    unchanged.prepare("INSERT INTO _la_objects(kind,name,definition) VALUES('query','command_projection_query',?)");unchanged.addBindValue(QString::fromUtf8(QJsonDocument(QJsonObject{{"sql","SELECT p.note AS [Notiz],p.value AS [Bezeichnung] FROM command_projection_rows p WHERE p.priority>=1 ORDER BY Bezeichnung ASC"}}).toJson(QJsonDocument::Compact)));assert(unchanged.exec());unchanged.finish();
    assert(window.openQueryDocument("command_projection_query",0,1).value("error").toInt()==0);auto* projectionView=dynamic_cast<linaccess::AccessTableView*>(documentTabs->currentWidget()->findChild<QTableView*>("native-editable-query-view"));assert(projectionView);auto* projectionModel=qobject_cast<linaccess::AccessTableModel*>(projectionView->model());assert(projectionModel && projectionModel->rowCount()==2 && projectionView->currentIndex().column()==3 && projectionView->horizontalHeader()->logicalIndex(0)==3 && projectionView->horizontalHeader()->logicalIndex(1)==1 && projectionView->isColumnHidden(0) && projectionView->isColumnHidden(2) && projectionView->isColumnHidden(4));assert(projectionModel->headerData(3,Qt::Horizontal).toString()=="Notiz" && projectionModel->headerData(1,Qt::Horizontal).toString()=="Bezeichnung" && projectionModel->record(0).value("id").toInt()==2);
    linaccess::NativeTableFindArguments searchProjection;searchProjection.text="match";searchProjection.column=-1;searchProjection.origin=1;searchProjection.mode=1;assert(projectionView->findText(searchProjection)==1 && projectionView->currentIndex().column()==3);
    projectionView->setCurrentIndex(projectionModel->index(0,1));assert(projectionModel->setData(projectionModel->index(0,1),"updated"));assert(mixed->executeProcedure("SaveActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed && !projectionModel->isDirty());assert(unchanged.exec("SELECT value,hidden,note FROM command_projection_rows WHERE id=2") && unchanged.next() && unchanged.value(0).toString()=="updated" && unchanged.value(1).toString()=="keep2" && unchanged.value(2).toString()=="second");unchanged.finish();
    assert(mixed->executeProcedure("RequeryActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed && projectionModel->headerData(3,Qt::Horizontal).toString()=="Notiz" && projectionView->isColumnHidden(0) && projectionView->horizontalHeader()->logicalIndex(0)==3);
    projectionView->setCurrentIndex(projectionModel->index(0,3));QApplication::clipboard()->setText("one\ttwo");assert(projectionView->pasteSelection() && projectionModel->index(0,3).data().toString()=="one" && projectionModel->index(0,1).data().toString()=="two" && projectionModel->index(0,2).data().toString()=="keep1");assert(projectionModel->undoBlock() && !projectionModel->isDirty() && projectionModel->index(0,3).data().toString()=="match");assert(projectionModel->redoBlock() && projectionModel->index(0,3).data().toString()=="one" && projectionModel->index(0,1).data().toString()=="two");assert(mixed->executeProcedure("SaveActiveTable",selectArguments).status==linaccess::VbaExecutionStatus::Completed);assert(unchanged.exec("SELECT note,value,hidden FROM command_projection_rows WHERE id=1") && unchanged.next() && unchanged.value(0).toString()=="one" && unchanged.value(1).toString()=="two" && unchanged.value(2).toString()=="keep1");unchanged.finish();
    projectionView->selectionModel()->select(QItemSelection(projectionModel->index(0,0),projectionModel->index(0,4)),QItemSelectionModel::ClearAndSelect);assert(projectionView->copySelection() && QApplication::clipboard()->text()=="one\ttwo");const auto copiedProjection=QJsonDocument::fromJson(QApplication::clipboard()->mimeData()->data(linaccess::nativeGridMime)).object();assert(copiedProjection.value("fields").toArray()==QJsonArray({"note","value"}));assert(projectionView->canCutSelection() && projectionView->cutSelection() && projectionModel->index(0,3).data(Qt::EditRole).isNull() && projectionModel->index(0,1).data(Qt::EditRole).isNull() && projectionModel->index(0,2).data().toString()=="keep1");assert(projectionModel->undoBlock() && !projectionModel->isDirty() && projectionModel->index(0,3).data().toString()=="one" && projectionModel->index(0,1).data().toString()=="two");assert(window.closeQueryDocument(1,"command_projection_query",0).value("error").toInt()==0);QCoreApplication::sendPostedEvents(nullptr,QEvent::DeferredDelete);
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
    assert(mixed->executeProcedure("PreviousTableRecord",selectArguments).errorNumber==3251 && recordEditModel->rowCount()==2 && recordEditView->currentIndex().row()==1);
    assert(recordEditModel->setData(recordEditModel->index(addedRow,1),"new-record"));
    assert(mixed->executeProcedure("AddTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed && recordEditModel->rowCount()==3 && recordEditView->currentIndex().row()==2);
    assert(unchanged.exec("SELECT count(*) FROM command_record_edits") && unchanged.next() && unchanged.value(0).toInt()==2); unchanged.finish();
    assert(recordEditModel->undoPendingRecord(2) && recordEditModel->rowCount()==2 && !recordEditModel->isDirty());
    recordEditView->setCurrentIndex(recordEditModel->index(addedRow,1));
    assert(mixed->executeProcedure("PreviousTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed && recordEditView->currentIndex().row()==0 && !recordEditModel->isDirty());
    assert(mixed->executeProcedure("LastTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
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
    navigationView->setCurrentIndex(navigationModel->index(2,1));QTest::keyClick(navigationView,Qt::Key_Down);assert(navigationView->currentIndex().row()==3 && navigationModel->isDirty());
    assert(mixed->executeProcedure("LastTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==299 && navigationModel->isDirty());
    assert(navigationModel->index(0,1).data(Qt::EditRole).toString()=="navigation-saved" && navigationModel->index(1,1).data(Qt::EditRole).toString()=="other-draft");
    navigationView->setCurrentIndex(navigationModel->index(0,1));

    assert(mixed->executeProcedure("LastTableRecord",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==299 && navigationModel->isDirty());
    assert(!navigationModel->hasPendingRecord(0) && navigationModel->hasPendingRecord(1) && navigationModel->index(1,1).data(Qt::EditRole).toString()=="other-draft");
    assert(unchanged.exec("SELECT value FROM command_navigation WHERE id=2") && unchanged.next() && unchanged.value(0).toString()=="row-2");unchanged.finish();
    navigationView->setCurrentIndex(navigationModel->index(0,1));
    assert(mixed->executeProcedure("SortTableDescending",selectArguments).errorNumber==3251 && navigationView->currentIndex().row()==0 && navigationModel->isDirty());
    assert(mixed->executeProcedure("RequeryActiveTable",selectArguments).errorNumber==3251 && navigationView->currentIndex().row()==0 && navigationModel->isDirty());
    assert(unchanged.exec("SELECT value FROM command_navigation WHERE id=1") && unchanged.next() && unchanged.value(0).toString()=="navigation-saved"); unchanged.finish();
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
    assert(mixed->executeProcedure("CopyTableSelection",selectArguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(QApplication::clipboard()->text()=="0x0102" && QApplication::clipboard()->mimeData()->hasFormat(linaccess::nativeGridMime) && !navigationModel->isDirty());
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
    assert(navigationView->currentIndex().row()==269 && navigationModel->isDirty() && !navigationModel->hasPendingRecord(279) && navigationModel->hasPendingRecord(278));
    assert(unchanged.exec("SELECT value FROM command_navigation WHERE id=280") && unchanged.next() && unchanged.value(0).toString()=="pending-current");unchanged.finish();
    assert(unchanged.exec("SELECT value FROM command_navigation WHERE id=279") && unchanged.next() && unchanged.value(0).toString()=="row-279");unchanged.finish();
    assert(unchanged.exec("UPDATE command_navigation SET value='row-280' WHERE id=280"));
    navigationModel->revertAll();assert(navigationModel->select()); assert(navigationModel->setReadOnly(true));
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
    assert(mixed->executeProcedure("FindRecordFormatted",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==249);
    navigationView->setCurrentIndex(navigationModel->index(24,0));
    assert(mixed->executeProcedure("FindRecordAllFormatted",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==249 && navigationView->currentIndex().column()==1);
    navigationView->setCurrentIndex(navigationModel->index(24,0));assert(mixed->executeProcedure("FindRecordAllFields",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==249 && navigationView->currentIndex().column()==1);
    assert(mixed->executeProcedure("FindRecordAllNumber",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==41 && navigationView->currentIndex().column()==0);
    navigationView->setCurrentIndex(navigationModel->index(24,1));

    assert(mixed->executeProcedure("FindRecordEmptyMatch",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==24);
    assert(mixed->executeProcedure("FindRecordNoMatch",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==24);
    for(const char* procedure:{"FindRecordNumber"})
        assert(mixed->executeProcedure(procedure,selectArguments).errorNumber==3251 && navigationView->currentIndex().row()==24);
    assert(mixed->executeProcedure("FindRecordWildcard",selectArguments).status==linaccess::VbaExecutionStatus::Completed && navigationView->currentIndex().row()==239 && makePreviewFormObject(*mixed)->member("Tag",{}).asText()=="wildcard-found");
    assert(mixed->executeProcedure("FindRecordInvalidDirection",selectArguments).errorNumber==5);
    assert(mixed->executeProcedure("FindRecordNull",selectArguments).errorNumber==94);
    assert(makePreviewFormObject(*mixed)->member("Tag",{}).asText()=="wildcard-found");
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
    assert(mixed->executeProcedure("PasteTableSelection",selectArguments).status==linaccess::VbaExecutionStatus::Completed && typedModel->index(0,1).data(Qt::EditRole).toDouble()==9 && typedModel->index(1,1).data(Qt::EditRole).toDouble()==3.75);
    assert(mixed->executeProcedure("UndoTableText",selectArguments).status==linaccess::VbaExecutionStatus::Completed && typedModel->isDirty() && typedModel->index(0,1).data(Qt::EditRole).toDouble()==2.5 && typedModel->index(1,1).data(Qt::EditRole).toDouble()==3.75);
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
    refreshEdit();assert(editAction("cut")->isEnabled());editAction("cut")->trigger();assert(model->isDirty() && model->canUndoBlock());refreshEdit();assert(editAction("undo")->isEnabled());editAction("undo")->trigger();assert(!model->isDirty());
    table->setCurrentIndex(model->index(0,1));table->selectionModel()->select(model->index(0,1),QItemSelectionModel::ClearAndSelect);
    const auto cutRange=window.dispatchTableCommand(189);assert(cutRange.value("handled").toBool() && cutRange.value("error").toInt()==0 && model->isDirty());assert(window.dispatchTableCommand(292).value("error").toInt()==0 && !model->isDirty());

    table->setCurrentIndex(model->index(0,1)); QApplication::clipboard()->setText("first-range\r\nsecond-range");
    refreshEdit(); assert(editAction("paste")->isEnabled()); editAction("paste")->trigger();
    assert(model->isDirty() && model->index(0,1).data().toString()=="first-range" && model->index(1,1).data().toString()=="second-range");
    refreshEdit(); assert(editAction("paste")->isEnabled());
    query.prepare("SELECT value FROM shortcut_records WHERE id=?"); query.addBindValue(id); assert(query.exec() && query.next() && query.value(0).toString()=="saved-through-shortcut"); query.finish();
    QTest::keyClick(table,Qt::Key_S,Qt::ControlModifier); QApplication::processEvents(); assert(!model->isDirty());
    assert(query.exec("SELECT COUNT(*) FROM shortcut_records WHERE value IN ('first-range','second-range')") && query.next() && query.value(0).toInt()==2); query.finish();
    assert(table->pasteSelection()); if(model->isDirty()) model->revertAll();
    table->setCurrentIndex(model->index(0,1));QApplication::clipboard()->setText("appended-range");
    const int beforeAppend=model->rowCount();const auto appended=window.dispatchTableCommand(38);
    assert(appended.value("handled").toBool() && appended.value("error").toInt()==0 && model->rowCount()==beforeAppend+1 && model->data(model->index(beforeAppend,1),Qt::EditRole).toString()=="appended-range");
    assert(model->canUndoBlock());const auto undone=window.dispatchTableCommand(292);assert(undone.value("error").toInt()==0 && model->rowCount()==beforeAppend && model->canRedoBlock());
    const auto redone=window.dispatchTableCommand(199);assert(redone.value("error").toInt()==0 && model->rowCount()==beforeAppend+1 && model->canUndoBlock());
    model->revertAll();

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

void exerciseNativeRelationDesigner() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","native-relation-designer-test");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery query(db);assert(query.exec("PRAGMA foreign_keys=ON"));query.finish();
    const auto database=linaccess::makeDaoDatabase(db);
    for(const auto& name:{"GuiParent","GuiChild"}) {
        auto table=database->member("CreateTableDef",{VbaValue::string(name)}).object;
        table->member("Fields",{}).object->member("Append",{table->member("CreateField",{VbaValue::string("ID"),VbaValue::integer(4)})});
        if(std::string(name)=="GuiChild") table->member("Fields",{}).object->member("Append",{table->member("CreateField",{VbaValue::string("ParentID"),VbaValue::integer(4)})});
        auto key=table->member("CreateIndex",{VbaValue::string("PrimaryKey")}).object;key->assign("Primary",VbaValue::logical(true));key->member("Fields",{}).object->member("Append",{key->member("CreateField",{VbaValue::string("ID")})});table->member("Indexes",{}).object->member("Append",{VbaValue::reference(key)});
        database->member("TableDefs",{}).object->member("Append",{VbaValue::reference(table)});
    }
    auto child=database->member("TableDefs",{VbaValue::string("GuiChild")}).object;auto unique=child->member("CreateIndex",{VbaValue::string("ParentUnique")}).object;unique->assign("Unique",VbaValue::logical(true));unique->member("Fields",{}).object->member("Append",{unique->member("CreateField",{VbaValue::string("ParentID")})});child->member("Indexes",{}).object->member("Append",{VbaValue::reference(unique)});
    database->member("Execute",{VbaValue::string("INSERT INTO GuiParent VALUES(7)")});database->member("Execute",{VbaValue::string("INSERT INTO GuiChild VALUES(1,7)")});
    QTimer::singleShot(0,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-relation-designer");
        auto* list=dialog->findChild<QTableWidget*>("native-relation-list");assert(list && list->rowCount()==0);
        dialog->findChild<QLineEdit*>("native-relation-name")->setText("GuiForeignKey");
        dialog->findChild<QComboBox*>("native-relation-parent")->setCurrentText("GuiParent");dialog->findChild<QComboBox*>("native-relation-child")->setCurrentText("GuiChild");
        auto* fields=dialog->findChild<QTableWidget*>("native-relation-fields");assert(fields && fields->rowCount()==1);
        qobject_cast<QComboBox*>(fields->cellWidget(0,0))->setCurrentText("ID");qobject_cast<QComboBox*>(fields->cellWidget(0,1))->setCurrentText("ParentID");
        dialog->findChild<QCheckBox*>("native-relation-unique")->setChecked(true);
        assert(dialog->findChild<QCheckBox*>("native-relation-integrity")->isChecked());dialog->findChild<QCheckBox*>("native-relation-update-cascade")->setChecked(true);dialog->findChild<QCheckBox*>("native-relation-delete-cascade")->setChecked(true);
        QTimer::singleShot(0,dialog,[] {auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());if(box && box->windowTitle()=="Beziehung konnte nicht geändert werden") {std::cerr<<"Relation GUI error: "<<box->text().toStdString()<<"\n";box->accept();}});
        dialog->findChild<QPushButton*>("native-relation-create")->click();assert(list->rowCount()==1 && list->item(0,4)->text()=="ja" && list->item(0,5)->text()=="ja" && list->item(0,6)->text()=="ja" && list->item(0,7)->text()=="ja");
        assert(query.exec("PRAGMA foreign_key_list(GuiChild)") && query.next() && query.value(2).toString()=="GuiParent" && query.value(5).toString()=="CASCADE" && query.value(6).toString()=="CASCADE");query.finish();
        assert(query.exec("SELECT ID,ParentID FROM GuiChild") && query.next() && query.value(0).toInt()==1 && query.value(1).toInt()==7);query.finish();
        QTimer::singleShot(0,dialog,[] {auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(box);box->button(QMessageBox::Yes)->click();});
        dialog->findChild<QPushButton*>("native-relation-delete")->click();assert(list->rowCount()==0);dialog->reject();
    });
    linaccess::editNativeRelations(nullptr,db,"GuiParent");
    assert(query.exec("PRAGMA foreign_key_list(GuiChild)") && !query.next());query.finish();
    assert(query.exec("SELECT ID,ParentID FROM GuiChild") && query.next() && query.value(0).toInt()==1 && query.value(1).toInt()==7);query.finish();
    database->member("Execute",{VbaValue::string("INSERT INTO GuiChild VALUES(2,999)")});
    db.close();
}

void exerciseScalarRelationDesigner(QWidget& owner,int type) {
    using linaccess::VbaValue;auto db=QSqlDatabase::addDatabase("QSQLITE","scalar-relation-designer-"+QString::number(type));db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);assert(sql.exec("PRAGMA foreign_keys=ON"));auto database=linaccess::makeDaoDatabase(db);
    for(const auto* name:{"Scalar UI parent","Scalar UI child"}) {auto table=database->member("CreateTableDef",{VbaValue::string(name)}).object;table->member("Fields",{}).object->member("Append",{table->member("CreateField",{VbaValue::string("KeyValue"),VbaValue::integer(type)})});if(QString(name).endsWith("parent")) {auto index=table->member("CreateIndex",{VbaValue::string("PrimaryKey")}).object;index->assign("Primary",VbaValue::logical(true));index->member("Fields",{}).object->member("Append",{index->member("CreateField",{VbaValue::string("KeyValue")})});table->member("Indexes",{}).object->member("Append",{VbaValue::reference(index)});}database->member("TableDefs",{}).object->member("Append",{VbaValue::reference(table)});database->member("Execute",{VbaValue::string((QString("INSERT INTO [")+name+"] VALUES("+(type==5 ? QString("1.2345") : type==6 ? QString("0.1") : QString("'2026-11-08'"))+")").toStdString())});}
    QTimer::singleShot(0,&owner,[&] {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-relation-designer");dialog->findChild<QComboBox*>("native-relation-parent")->setCurrentText("Scalar UI parent");dialog->findChild<QComboBox*>("native-relation-child")->setCurrentText("Scalar UI child");dialog->findChild<QLineEdit*>("native-relation-name")->setText("Scalar UI mapping");dialog->findChild<QCheckBox*>("native-relation-update-cascade")->setChecked(true);dialog->findChild<QCheckBox*>("native-relation-delete-cascade")->setChecked(true);auto* fields=dialog->findChild<QTableWidget*>("native-relation-fields");assert(fields && fields->rowCount()==1);for(int column=0;column<2;++column) qobject_cast<QComboBox*>(fields->cellWidget(0,column))->setCurrentText("KeyValue");
        QTimer::singleShot(0,dialog,[] {auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());if(box && box->windowTitle()=="Beziehung konnte nicht geändert werden") {std::cerr<<"Scalar relation GUI error: "<<box->text().toStdString()<<"\n";box->accept();}});dialog->findChild<QPushButton*>("native-relation-create")->click();assert(dialog->findChild<QTableWidget*>("native-relation-list")->rowCount()==1);dialog->reject();});
    linaccess::editNativeRelations(&owner,db,"Scalar UI parent");assert(sql.exec("PRAGMA foreign_key_list('Scalar UI child')") && sql.next() && sql.value(2).toString()=="Scalar UI parent");sql.finish();database->member("Execute",{VbaValue::string(type==5 ? "UPDATE [Scalar UI parent] SET KeyValue=-2.3456" : type==6 ? "UPDATE [Scalar UI parent] SET KeyValue=0.2" : "UPDATE [Scalar UI parent] SET KeyValue='2026-11-09'")});assert(sql.exec("SELECT KeyValue FROM [Scalar UI child]") && sql.next() && (type==5 ? sql.value(0).toDouble()==-2.3456 : type==6 ? sql.value(0).toDouble()==static_cast<double>(0.2f) : sql.value(0).toString()=="2026-11-09 00:00:00"));sql.finish();database->member("Execute",{VbaValue::string("DELETE FROM [Scalar UI parent]")});assert(sql.exec("SELECT COUNT(*) FROM [Scalar UI child]") && sql.next() && sql.value(0).toInt()==0);sql.finish();
}

void exerciseDateDefaultBuffers() {
    using linaccess::VbaValue;auto db=QSqlDatabase::addDatabase("QSQLITE","date-default-buffers");db.setDatabaseName(":memory:");assert(db.open());auto database=linaccess::makeDaoDatabase(db);auto table=database->member("CreateTableDef",{VbaValue::string("Date UI rows")}).object;
    auto id=table->member("CreateField",{VbaValue::string("ID"),VbaValue::integer(4)}).object;id->assign("Attributes",VbaValue::longInteger(16));table->member("Fields",{}).object->member("Append",{VbaValue::reference(id)});auto stamp=table->member("CreateField",{VbaValue::string("Moment"),VbaValue::integer(8)}).object;stamp->assign("DefaultValue",VbaValue::string("Now()"));table->member("Fields",{}).object->member("Append",{VbaValue::reference(stamp)});table->member("Fields",{}).object->member("Append",{table->member("CreateField",{VbaValue::string("Label"),VbaValue::integer(10),VbaValue::longInteger(8)})});auto primary=table->member("CreateIndex",{VbaValue::string("PrimaryKey")}).object;primary->assign("Primary",VbaValue::logical(true));primary->member("Fields",{}).object->member("Append",{primary->member("CreateField",{VbaValue::string("ID")})});table->member("Indexes",{}).object->member("Append",{VbaValue::reference(primary)});database->member("TableDefs",{}).object->member("Append",{VbaValue::reference(table)});
    QSqlQuery sql(db);QString frozen;
    {linaccess::AccessTableModel model(nullptr,db);model.setTable("Date UI rows");model.setEditStrategy(QSqlTableModel::OnManualSubmit);assert(model.select() && model.insertRow(0));const int moment=model.fieldIndex("Moment");frozen=model.record(0).value(moment).toString();assert(QDateTime::fromString(frozen,"yyyy-MM-dd HH:mm:ss").isValid() && model.record(0).isGenerated(moment));assert(model.setData(model.index(0,model.fieldIndex("Label")),"sheet"));assert(model.submitTransactionally());}
    assert(sql.exec("SELECT Moment FROM [Date UI rows] WHERE ID=1") && sql.next() && sql.value(0).toString()==frozen);sql.finish();
    const QJsonObject definition{{"table","Date UI rows"},{"fields",QJsonArray{"Moment","Label"}},{"vba","Sub StoreDate()\nIf IsNull(Me!Moment) Then Err.Raise 5\nIf VarType(Me!Moment)<>vbDate Then Err.Raise 5\nMe!Label=\"form\"\nDoCmd.RunCommand acCmdSaveRecord\nEnd Sub\n"}};const auto stored=linaccess::nativeStoredForm("Date buffer form",definition,db);
    AccessFormPreview form(stored.name,stored.layout,stored.code,db,nullptr,false,{},std::nullopt,0,0);form.show();QApplication::processEvents();assert(form.hasDataSource());auto* moment=form.findChild<QLineEdit*>("access-control-Moment");assert(moment && !moment->text().isEmpty());const auto formDefault=moment->text();std::vector<VbaValue> arguments;const auto result=form.executeProcedure("StoreDate",arguments);if(result.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Date form: "<<result.explanation<<"\n";assert(result.status==linaccess::VbaExecutionStatus::Completed);assert(sql.exec("SELECT Moment FROM [Date UI rows] WHERE ID=2") && sql.next() && sql.value(0).toString()==formDefault);sql.finish();assert(form.close());
}

void testNativeMacros(MainWindow& window,QSqlDatabase db) {
    const QString isolatedConnection="missing-native-macro-catalog";
    {auto isolated=QSqlDatabase::addDatabase("QSQLITE",isolatedConnection);isolated.setDatabaseName(":memory:");assert(isolated.open());QSqlQuery probe(isolated);
        const auto expect=[&](int wanted){int actual=0;try{(void)linaccess::loadNativeMacro(isolated,"MissingMacro");}catch(const linaccess::VbaRuntimeError& error){actual=error.number;}assert(actual==wanted);};
        expect(2485);assert(probe.exec("SELECT COUNT(*) FROM main.sqlite_schema WHERE name='_la_objects'") && probe.next() && probe.value(0).toInt()==0);probe.finish();
        assert(probe.exec("CREATE TEMP TABLE _la_objects(kind TEXT,name TEXT,definition TEXT)"));assert(probe.exec("INSERT INTO temp._la_objects VALUES('macro','MissingMacro','{}')"));expect(2485);
        assert(probe.exec("CREATE TABLE main._la_objects(name TEXT)"));expect(3146);assert(probe.exec("DROP TABLE main._la_objects"));assert(probe.exec("CREATE TABLE main._la_objects(kind TEXT,name TEXT,definition TEXT)"));expect(2485);probe.finish();isolated.close();}
    QSqlDatabase::removeDatabase(isolatedConnection);
    using linaccess::VbaValue;QSqlQuery sql(db);assert(sql.exec("CREATE TABLE native_macro_effects(n INTEGER)"));
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-macro-designer");
        dialog->findChild<QLineEdit*>("native-macro-name")->setText("Native Macro");auto* rows=dialog->findChild<QTableWidget*>("native-macro-actions");assert(rows && rows->rowCount()==1);
        rows->item(0,1)->setText("INSERT INTO native_macro_effects VALUES(1)");rows->item(0,2)->setText("Ja");rows->item(0,4)->setText("True");dialog->findChild<QPushButton*>("native-macro-save")->click();
    });action(window,"+ Makro")->trigger();
    assert(linaccess::loadNativeMacro(db,"native macro").size()==1 && linaccess::loadNativeMacro(db,"native macro").front().condition==QString("True"));
    auto* explorer=window.findChild<QTreeWidget*>("database-explorer");bool listed=false;QTreeWidgetItem* macroItem=nullptr;
    for(int group=0;group<explorer->topLevelItemCount();++group) for(int index=0;index<explorer->topLevelItem(group)->childCount();++index) {auto* item=explorer->topLevelItem(group)->child(index);if(item->data(0,Qt::UserRole).toString()=="macro" && item->text(0)=="Native Macro") {listed=true;macroItem=item;}}
    assert(listed && macroItem);QTimer::singleShot(0,&window,[&] {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-macro-designer");auto* rows=dialog->findChild<QTableWidget*>("native-macro-actions");assert(rows && rows->item(0,1)->text()=="INSERT INTO native_macro_effects VALUES(1)");dialog->reject();});QMetaObject::invokeMethod(explorer,"itemDoubleClicked",Qt::DirectConnection,Q_ARG(QTreeWidgetItem*,macroItem),Q_ARG(int,0));
    const QByteArray code=R"VBA(Sub RepeatMacro()
DoCmd.RunMacro RepeatCount:=3,MacroName:="Native Macro"
End Sub
Sub FalseExpression()
DoCmd.RunMacro "Native Macro",1,False
End Sub
Sub MissingMacro()
DoCmd.RunMacro "No Such Macro"
End Sub)VBA";
    artifact(db,"vba","MacroConditions",R"VBA(Public Function ConditionOnce() As Boolean
CurrentDb.Execute "INSERT INTO native_macro_effects VALUES(4)"
ConditionOnce=True
End Function)VBA");
    AccessFormPreview form("native_macro_context","Begin Form\nBegin Section\nEnd\nEnd",code,db,&window);form.show();QApplication::processEvents();std::vector<VbaValue> args;
    assert(form.executeProcedure("RepeatMacro",args).status==linaccess::VbaExecutionStatus::Completed);
    assert(sql.exec("SELECT COUNT(*),SUM(n) FROM native_macro_effects") && sql.next() && sql.value(0).toInt()==3 && sql.value(1).toInt()==3);sql.finish();
    assert(form.executeProcedure("FalseExpression",args).status==linaccess::VbaExecutionStatus::Completed && form.executeProcedure("MissingMacro",args).errorNumber==2485);
    auto application=makePreviewApplicationObject(form);auto macros=application->member("CurrentProject",{}).object->member("AllMacros",{}).object;
    assert(macros->member("Count",{}).number==1 && macros->enumerate().size()==1);
    auto descriptor=macros->member("Item",{VbaValue::string("native macro")}).object;
    assert(descriptor->member("Name",{}).asText()=="Native Macro" && !descriptor->member("IsLoaded",{}).truthy());
    assert(application->member("SysCmd",{VbaValue::integer(10),VbaValue::integer(4),VbaValue::string("Native Macro")}).number==0);
    auto commands=makePreviewApplicationObject(form)->member("DoCmd",{}).object;
    const auto store=[&](const QString& name,const QJsonArray& actions) {QSqlQuery insert(db);insert.prepare("INSERT INTO _la_objects(kind,name,definition) VALUES('macro',?,?)");insert.addBindValue(name);insert.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"actions",actions}}).toJson(QJsonDocument::Compact));assert(insert.exec());};
    const QJsonObject write{{"action","RunSQL"},{"arguments",QJsonArray{"INSERT INTO native_macro_effects VALUES(2)"}}};
    store("stopped",{write,QJsonObject{{"action","StopMacro"},{"arguments",QJsonArray{}}},write});commands->member("RunMacro",{VbaValue::string("stopped"),VbaValue::integer(3)});
    assert(sql.exec("SELECT COUNT(*),SUM(n) FROM native_macro_effects") && sql.next() && sql.value(0).toInt()==4 && sql.value(1).toInt()==5);sql.finish();
    store("invalid",{write,QJsonObject{{"action","UnsupportedAction"},{"arguments",QJsonArray{}}}});
    int error=0;try {commands->member("RunMacro",{VbaValue::string("invalid")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==3251);
    store("conditional",{write,QJsonObject{{"action","RunSQL"},{"arguments",QJsonArray{"INSERT INTO native_macro_effects VALUES(3)"}},{"condition",false}}});
    error=0;try {commands->member("RunMacro",{VbaValue::string("conditional")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==13);
    store("typed_invalid",{write,QJsonObject{{"action","RunCommand"},{"arguments",QJsonArray{1.5}}}});
    error=0;try {commands->member("RunMacro",{VbaValue::string("typed_invalid")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==13);
    assert(sql.exec("SELECT COUNT(*) FROM native_macro_effects") && sql.next() && sql.value(0).toInt()==4);sql.finish();
    store("sql_failure",{write,QJsonObject{{"action","RunSQL"},{"arguments",QJsonArray{"INSERT INTO native_macro_missing VALUES(1)"}}}});
    error=0;QString description;try {commands->member("RunMacro",{VbaValue::string("sql_failure")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;description=QString::fromUtf8(failure.what());}assert(error==3078 && description.contains("Aktion 2 (RunSQL)"));
    assert(sql.exec("SELECT COUNT(*),SUM(n) FROM native_macro_effects") && sql.next() && sql.value(0).toInt()==5 && sql.value(1).toInt()==7);sql.finish();
    error=0;try {commands->member("RunMacro",{VbaValue::string("Native Macro"),VbaValue::longInteger(1001)});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==3251);
    store(QString::fromUtf8("Änderung"),{QJsonObject{{"action","RunSQL"},{"arguments",QJsonArray{"INSERT INTO native_macro_effects VALUES(3)"}}}});
    commands->member("RunMacro",{VbaValue::string("änderung")});assert(linaccess::nativeMacroCatalogName(db,QString::fromUtf8("ÄNDERUNG"))==QString::fromUtf8("Änderung"));
    store(QString::fromUtf8("ÄNDERUNG"),{write});error=0;try {commands->member("RunMacro",{VbaValue::string("änderung")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==3251);
    assert(sql.exec("SELECT COUNT(*) FROM native_macro_effects WHERE n=3") && sql.next() && sql.value(0).toInt()==1);sql.finish();
    sql.prepare("DELETE FROM _la_objects WHERE kind='macro' AND name=? COLLATE BINARY");sql.addBindValue(QString::fromUtf8("ÄNDERUNG"));assert(sql.exec());sql.finish();
    const auto effect=[](int number,const QString& condition=QString{}) {QJsonObject value{{"action","RunSQL"},{"arguments",QJsonArray{QString("INSERT INTO native_macro_effects VALUES(%1)").arg(number)}}};if(!condition.isEmpty()) value.insert("condition",condition);return value;};
    store("condition_values",{effect(5,"True"),effect(6,"False"),effect(7,"Null")});commands->member("RunMacro",{VbaValue::string("condition_values")});
    assert(sql.exec("SELECT COUNT(*) FROM native_macro_effects WHERE n=5") && sql.next() && sql.value(0).toInt()==1);sql.finish();assert(sql.exec("SELECT COUNT(*) FROM native_macro_effects WHERE n IN (6,7)") && sql.next() && sql.value(0).toInt()==0);sql.finish();
    store("condition_once",{effect(8,"ConditionOnce()")});commands->member("RunMacro",{VbaValue::string("condition_once"),VbaValue::integer(3)});
    assert(sql.exec("SELECT COUNT(*) FROM native_macro_effects WHERE n=4") && sql.next() && sql.value(0).toInt()==3);sql.finish();assert(sql.exec("SELECT COUNT(*) FROM native_macro_effects WHERE n=8") && sql.next() && sql.value(0).toInt()==3);sql.finish();
    assert(sql.exec("SELECT COUNT(*) FROM native_macro_effects") && sql.next());const int beforeSyntax=sql.value(0).toInt();sql.finish();store("condition_syntax",{write,effect(9,"(")});error=0;try {commands->member("RunMacro",{VbaValue::string("condition_syntax")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==3251);
    assert(sql.exec("SELECT COUNT(*) FROM native_macro_effects") && sql.next() && sql.value(0).toInt()==beforeSyntax);sql.finish();
    store("nested_child",{effect(10)});store("nested_parent",{QJsonObject{{"action","RunMacro"},{"arguments",QJsonArray{"nested_child",2}}},effect(11)});commands->member("RunMacro",{VbaValue::string("nested_parent")});
    assert(sql.exec("SELECT COUNT(*) FROM native_macro_effects WHERE n=10") && sql.next() && sql.value(0).toInt()==2);sql.finish();assert(sql.exec("SELECT COUNT(*) FROM native_macro_effects WHERE n=11") && sql.next() && sql.value(0).toInt()==1);sql.finish();
    store("nested_stop",{effect(12),QJsonObject{{"action","StopMacro"},{"arguments",QJsonArray{}}},effect(14)});store("nested_continue",{QJsonObject{{"action","RunMacro"},{"arguments",QJsonArray{"nested_stop",3}}},effect(13)});commands->member("RunMacro",{VbaValue::string("nested_continue")});
    assert(sql.exec("SELECT COUNT(*) FROM native_macro_effects WHERE n IN (12,13)") && sql.next() && sql.value(0).toInt()==2);sql.finish();assert(sql.exec("SELECT COUNT(*) FROM native_macro_effects WHERE n=14") && sql.next() && sql.value(0).toInt()==0);sql.finish();
    store("recursive_macro",{QJsonObject{{"action","RunMacro"},{"arguments",QJsonArray{"recursive_macro"}}}});error=0;try {commands->member("RunMacro",{VbaValue::string("recursive_macro")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==3251);commands->member("RunMacro",{VbaValue::string("condition_values")});
    assert(form.close());
    assert(sql.exec("CREATE TABLE macro_event_rows(id INTEGER PRIMARY KEY,value TEXT)"));assert(sql.exec("INSERT INTO macro_event_rows VALUES(1,'before')"));assert(sql.exec("CREATE TABLE macro_event_audit(kind TEXT)"));
    for(const auto* kind:{"load","current","click","double","change","before","after","timer","close"}) store(QStringLiteral("event_")+kind,{QJsonObject{{"action","RunSQL"},{"arguments",QJsonArray{QStringLiteral("INSERT INTO macro_event_audit VALUES('")+kind+"')"}}}});
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="macro_event_rows"
OnLoad ="event_load"
OnCurrent ="event_current"
OnTimer ="event_timer"
OnClose ="event_close"
Width =5000
Begin Section
Height =1800
Begin TextBox
Name ="Value"
ControlSource ="value"
OnChange ="event_change"
BeforeUpdate ="event_before"
AfterUpdate ="event_after"
OnDblClick ="event_double"
Width =2000
Height =400
End
Begin CommandButton
Name ="Run"
OnClick ="event_click"
Top =600
Width =1800
Height =400
End
Begin CommandButton
Name ="Missing"
OnClick ="No Such Macro"
Left =2200
Top =600
Width =1800
Height =400
End
End
End)FORM";
    AccessFormPreview bound("macro_event_form",layout,{},db,&window);bound.show();QApplication::processEvents();assert(bound.hasDataSource());
    const auto observed=[&](const QString& kind) {QSqlQuery audit(db);audit.prepare("SELECT COUNT(*) FROM macro_event_audit WHERE kind=?");audit.addBindValue(kind);assert(audit.exec() && audit.next());return audit.value(0).toInt();};
    assert(observed("load")==1 && observed("current")==1);
    auto* button=bound.findChild<QPushButton*>("access-control-Run");auto* missing=bound.findChild<QPushButton*>("access-control-Missing");auto* edit=bound.findChild<QLineEdit*>("access-control-Value");assert(button && missing && edit);
    QTest::mouseClick(button,Qt::LeftButton);assert(observed("click")==1 && bound.eventResult("control:Run:Click")->status==linaccess::VbaExecutionStatus::Completed);
    QTest::mouseClick(missing,Qt::LeftButton);assert(bound.eventResult("control:Missing:Click")->errorNumber==2485);
    QTest::mouseDClick(edit,Qt::LeftButton);assert(observed("double")==1);
    edit->setFocus();QTest::keyClick(edit,Qt::Key_A,Qt::ControlModifier);QTest::keyClicks(edit,"after");assert(observed("change")>0);
    assert(sql.exec("SELECT value FROM macro_event_rows WHERE id=1") && sql.next() && sql.value(0).toString()=="before");sql.finish();
    makePreviewFormObject(bound)->assign("Dirty",VbaValue::logical(false));assert(observed("before")==1 && observed("after")==1);
    assert(sql.exec("SELECT value FROM macro_event_rows WHERE id=1") && sql.next() && sql.value(0).toString()=="after");sql.finish();
    auto* timer=bound.findChild<QTimer*>("access-form-timer");assert(timer && QMetaObject::invokeMethod(timer,"timeout",Qt::DirectConnection));assert(observed("timer")==1);
    assert(bound.close() && observed("close")==1);
    auto refused=layout;refused.replace("BeforeUpdate =\"event_before\"","BeforeUpdate =\"No Such Macro\"");
    AccessFormPreview guarded("macro_event_guarded",refused,{},db,&window);guarded.show();QApplication::processEvents();edit=guarded.findChild<QLineEdit*>("access-control-Value");assert(edit);edit->setFocus();QTest::keyClick(edit,Qt::Key_A,Qt::ControlModifier);QTest::keyClicks(edit,"refused");
    error=0;try {makePreviewFormObject(guarded)->assign("Dirty",VbaValue::logical(false));}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error && guarded.eventResult("control:Value:BeforeUpdate")->errorNumber==2485);
    assert(sql.exec("SELECT value FROM macro_event_rows WHERE id=1") && sql.next() && sql.value(0).toString()=="after");sql.finish();makePreviewFormObject(guarded)->member("Undo",{});assert(guarded.close());
    store("event_cancel",{QJsonObject{{"action","CancelEvent"},{"arguments",QJsonArray{}}}});
    auto cancelLayout=layout;cancelLayout.replace("BeforeUpdate =\"event_before\"","BeforeUpdate =\"event_cancel\"");
    AccessFormPreview validation("macro_event_cancel",cancelLayout,{},db,&window);validation.show();QApplication::processEvents();edit=validation.findChild<QLineEdit*>("access-control-Value");assert(edit);edit->setFocus();QTest::keyClick(edit,Qt::Key_A,Qt::ControlModifier);QTest::keyClicks(edit,"cancelled");
    error=0;try {makePreviewFormObject(validation)->assign("Dirty",VbaValue::logical(false));}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==2501);
    const auto cancellation=validation.eventResult("control:Value:BeforeUpdate");assert(cancellation && cancellation->status==linaccess::VbaExecutionStatus::Completed && cancellation->cancelled && !cancellation->errorNumber);
    assert(sql.exec("SELECT value FROM macro_event_rows WHERE id=1") && sql.next() && sql.value(0).toString()=="after");sql.finish();
    auto outside=makePreviewApplicationObject(validation)->member("DoCmd",{}).object;error=0;try {outside->member("CancelEvent",{});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==3251);makePreviewFormObject(validation)->member("Undo",{});assert(validation.close());
    AccessFormPreview cancelledOpen("macro_cancel_open","Begin Form\nOnOpen =\"event_cancel\"\nBegin Section\nEnd\nEnd",{},db,&window);cancelledOpen.show();QApplication::processEvents();assert(cancelledOpen.cancelledByOpen());cancelledOpen.close();
    AccessFormPreview cancelledClose("macro_cancel_close","Begin Form\nOnUnload =\"event_cancel\"\nBegin Section\nEnd\nEnd",{},db,&window);cancelledClose.show();QApplication::processEvents();assert(!cancelledClose.close() && cancelledClose.isVisible());
    assert(cancelledClose.eventResult("OnUnload")->cancelled);cancelledClose.hide();
    store("event_validate",{QJsonObject{{"action","CancelEvent"},{"arguments",QJsonArray{}},{"condition","[Value]=\"blocked\""}}});auto conditionalLayout=layout;conditionalLayout.replace("BeforeUpdate =\"event_before\"","BeforeUpdate =\"event_validate\"");
    AccessFormPreview conditional("conditional_validation",conditionalLayout,{},db,&window);conditional.show();QApplication::processEvents();edit=conditional.findChild<QLineEdit*>("access-control-Value");assert(edit);edit->setFocus();QTest::keyClick(edit,Qt::Key_A,Qt::ControlModifier);QTest::keyClicks(edit,"blocked");error=0;try {makePreviewFormObject(conditional)->assign("Dirty",VbaValue::logical(false));}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==2501 && conditional.eventResult("control:Value:BeforeUpdate")->cancelled);
    assert(sql.exec("SELECT value FROM macro_event_rows WHERE id=1") && sql.next() && sql.value(0).toString()=="after");sql.finish();edit->setFocus();QTest::keyClick(edit,Qt::Key_A,Qt::ControlModifier);QTest::keyClicks(edit,"allowed");makePreviewFormObject(conditional)->assign("Dirty",VbaValue::logical(false));assert(!conditional.eventResult("control:Value:BeforeUpdate")->cancelled);
    assert(sql.exec("SELECT value FROM macro_event_rows WHERE id=1") && sql.next() && sql.value(0).toString()=="allowed");sql.finish();assert(conditional.close());assert(sql.exec("UPDATE macro_event_rows SET value='after' WHERE id=1"));
}

void testNativeMacroDesigner(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;
    QTimer::singleShot(0,&window,[&] {
        auto* parent=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(parent && parent->objectName()=="native-form-designer");parent->findChild<QLineEdit*>("native-form-name")->setText("Authored Macro Form");parent->findChild<QComboBox*>("native-form-source")->setCurrentText("macro_event_rows");
        auto* edit=parent->findChild<QPushButton*>("native-form-event-bindings");assert(edit);
        QTimer::singleShot(0,&window,[&] {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-form-event-editor");dialog->findChild<QComboBox*>("native-form-event-event")->setCurrentText("OnLoad");dialog->findChild<QComboBox*>("native-form-event-mode")->setCurrentIndex(1);dialog->findChild<QComboBox*>("native-form-event-target")->setCurrentText("event_cancel");dialog->findChild<QDialogButtonBox*>()->button(QDialogButtonBox::Cancel)->click();});edit->click();
        QTimer::singleShot(0,&window,[&] {
            auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-form-event-editor");auto* event=dialog->findChild<QComboBox*>("native-form-event-event");event->setCurrentText("OnLoad");assert(dialog->findChild<QComboBox*>("native-form-event-mode")->currentIndex()==0);
            dialog->findChild<QComboBox*>("native-form-event-mode")->setCurrentIndex(1);dialog->findChild<QComboBox*>("native-form-event-target")->setCurrentText("event_load");
            dialog->findChild<QComboBox*>("native-field-event-field")->setCurrentText("value");dialog->findChild<QComboBox*>("native-field-event-event")->setCurrentText("BeforeUpdate");dialog->findChild<QComboBox*>("native-field-event-mode")->setCurrentIndex(1);dialog->findChild<QComboBox*>("native-field-event-target")->setCurrentText("event_cancel");dialog->findChild<QDialogButtonBox*>()->button(QDialogButtonBox::Ok)->click();
        });edit->click();parent->findChild<QPushButton*>("native-form-save")->click();
    });action(window,"+ Formular")->trigger();
    QSqlQuery sql(db);assert(sql.exec("SELECT definition FROM _la_objects WHERE kind='form' AND name='Authored Macro Form'") && sql.next());const auto definition=QJsonDocument::fromJson(sql.value(0).toByteArray()).object();sql.finish();
    assert(definition.value("events").toObject().value("OnLoad")=="event_load" && definition.value("field_events").toObject().value("value").toObject().value("BeforeUpdate")=="event_cancel");
    const auto stored=linaccess::loadNativeStoredForm(db,"Authored Macro Form");assert(stored);const auto parsed=linaccess::parseAccessFormLayout(QString::fromUtf8(stored->layout));assert(parsed.events.value("onload")=="event_load");
    assert(sql.exec("SELECT COUNT(*) FROM macro_event_audit WHERE kind='load'") && sql.next());const int before=sql.value(0).toInt();sql.finish();
    AccessFormPreview authored(stored->name,stored->layout,stored->code,db,&window);authored.show();QApplication::processEvents();assert(sql.exec("SELECT COUNT(*) FROM macro_event_audit WHERE kind='load'") && sql.next() && sql.value(0).toInt()==before+1);sql.finish();
    auto* editor=authored.findChild<QLineEdit*>("access-control-value");assert(editor);editor->setFocus();QTest::keyClick(editor,Qt::Key_A,Qt::ControlModifier);QTest::keyClicks(editor,"designer-cancel");int error=0;try {makePreviewFormObject(authored)->assign("Dirty",VbaValue::logical(false));}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==2501 && authored.eventResult("control:value:BeforeUpdate")->cancelled);
    assert(sql.exec("SELECT value FROM macro_event_rows WHERE id=1") && sql.next() && sql.value(0).toString()=="after");sql.finish();makePreviewFormObject(authored)->member("Undo",{});assert(authored.close());
    auto disabled=definition;disabled.insert("events",QJsonObject{{"OnLoad",""}});disabled.insert("field_events",QJsonObject{});disabled.insert("vba","Private Sub Form_Load()\nCurrentDb.Execute \"INSERT INTO macro_event_audit VALUES('automatic')\"\nEnd Sub");
    const auto without=linaccess::nativeStoredForm("Disabled Auto Event",disabled,db);AccessFormPreview suppressed(without.name,without.layout,without.code,db,&window);suppressed.show();QApplication::processEvents();assert(sql.exec("SELECT COUNT(*) FROM macro_event_audit WHERE kind='automatic'") && sql.next() && sql.value(0).toInt()==0);sql.finish();assert(suppressed.close());
    auto unknown=definition;unknown.insert("events",QJsonObject{{"UnknownEvent","event_load"}});error=0;try {(void)linaccess::nativeStoredForm("Unknown event",unknown,db);}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==3251);
    unknown=definition;unknown.insert("field_events",QJsonObject{{"missing",QJsonObject{{"OnClick","event_load"}}}});error=0;try {(void)linaccess::nativeStoredForm("Unknown field event",unknown,db);}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==3251);
    unknown=definition;unknown.insert("timer_interval",-1);error=0;try {(void)linaccess::nativeStoredForm("Invalid timer",unknown,db);}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==5);
}


void testNativeMacroSetValue(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;QSqlQuery sql(db);assert(sql.exec("CREATE TABLE set_value_rows(id INTEGER PRIMARY KEY,n INTEGER CHECK(n>=0))"));assert(sql.exec("INSERT INTO set_value_rows VALUES(1,10)"));assert(sql.exec("CREATE TABLE set_value_audit(n INTEGER)"));sql.finish();const auto action=[](QString name,QJsonArray args) {return QJsonObject{{"action",name},{"arguments",args}};};const auto macro=[&](QString name,QJsonArray actions) {linaccess::storeNativeMacro(db,name,{{"version",1},{"actions",actions}});};
    macro("SetValueCurrent",{action("SetValue",{"Value","Forms!SetValueForm!Value+1"}),action("SetValue",{"Me.Tag","\"from macro\""})});macro("SetValueExplicit",{action("SetValue",{"Forms!SetValueForm!Value","25"}),action("SetValue",{"Forms!SetValueForm.Visible","True"})});macro("SetValueFailure",{action("SetValue",{"Value","1/0"})});macro("SetValueNull",{action("SetValue",{"Value","Null"})});macro("SetValueCalculated",{action("SetValue",{"Calc","9"})});
    const auto source=QByteArray("Begin Form\nRecordSource =\"set_value_rows\"\nWidth =2400\nBegin Section\nHeight =1000\nBegin TextBox\nName =\"Value\"\nControlSource =\"n\"\nWidth =1800\nHeight =300\nEnd\nBegin TextBox\nName =\"Calc\"\nControlSource =\"=[n]+1\"\nTop =400\nWidth =1800\nHeight =300\nEnd\nEnd\nEnd");AccessFormPreview form("SetValueForm",source,{},db,&window);form.show();QApplication::processEvents();auto root=makePreviewFormObject(form);auto value=root->member("Controls",{VbaValue::string("Value")}).object;linaccess::NativeStandardModuleHost host(&window,db,[] {});auto commands=host.applicationObject()->member("DoCmd",{}).object;const auto stored=[&] {QSqlQuery query(db);assert(query.exec("SELECT n FROM set_value_rows WHERE id=1") && query.next());return query.value(0);};
    commands->member("RunMacro",{VbaValue::string("SetValueCurrent")});assert(value->scalarValue().number==11 && root->member("Tag",{}).asText()=="from macro" && root->member("Dirty",{}).boolean && stored().toInt()==10);root->assign("Dirty",VbaValue::logical(false));assert(stored().toInt()==11);auto formCommands=makePreviewApplicationObject(form)->member("DoCmd",{}).object;formCommands->member("RunMacro",{VbaValue::string("SetValueCurrent")});assert(value->scalarValue().number==12);root->assign("Dirty",VbaValue::logical(false));assert(stored().toInt()==12);
    root->assign("Visible",VbaValue::logical(false));window.focusDatabaseWindow();commands->member("RunMacro",{VbaValue::string("SetValueExplicit")});assert(root->member("Visible",{}).boolean && value->scalarValue().number==25);root->assign("Dirty",VbaValue::logical(false));assert(stored().toInt()==25);commands->member("SelectObject",{VbaValue::integer(2),VbaValue::string("SetValueForm"),VbaValue::logical(false)});const auto failure=[&](const char* name,int expected) {int error=0;try {commands->member("RunMacro",{VbaValue::string(name)});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}if(error!=expected) qFatal("SetValue %s: expected=%d actual=%d",name,expected,error);};failure("SetValueFailure",11);assert(value->scalarValue().number==25 && !root->member("Dirty",{}).boolean);failure("SetValueCalculated",2448);assert(value->scalarValue().number==25);commands->member("RunMacro",{VbaValue::string("SetValueNull")});assert(value->scalarValue().kind==VbaValue::Kind::Null);root->assign("Dirty",VbaValue::logical(false));assert(stored().isNull());
    const auto invalid=QJsonObject{{"version",1},{"actions",QJsonArray{action("RunSQL",{"INSERT INTO set_value_audit VALUES(1)"}),action("SetValue",{"Value","=1+"})}}};sql.prepare("INSERT INTO _la_objects VALUES('macro','SetValueInvalidPrefix',?)");sql.addBindValue(QJsonDocument(invalid).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();failure("SetValueInvalidPrefix",3251);int error=0;assert(sql.exec("SELECT COUNT(*) FROM set_value_audit") && sql.next() && sql.value(0).toInt()==0);sql.finish();error=0;try {commands->member("SetValue",{VbaValue::string("Value"),VbaValue::integer(1)});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==438);
    sql.prepare("INSERT INTO _la_objects VALUES('module','SetValueRuntimeCode',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba","Public Function CloseSetValueTarget() As Long\nApplication.DoCmd.Close acForm,\"SetValueForm\"\nCloseSetValueTarget=99\nEnd Function"}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();macro("SetValueClosedDuringExpression",{action("SetValue",{"Value","CloseSetValueTarget()"})});failure("SetValueClosedDuringExpression",2450);assert(stored().isNull());assert(sql.exec("DELETE FROM _la_objects WHERE kind='module' AND name='SetValueRuntimeCode'"));
}

void testMacroArgumentPresence(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;QSqlQuery sql(db);assert(sql.exec("CREATE TABLE macro_argument_events(n INTEGER)"));sql.prepare("INSERT INTO _la_objects VALUES('module','MacroArgumentCode',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba",R"VBA(Public Function ObserveMacroOpening() As Long
If IsNull(Forms!StandaloneOpenForm.OpenArgs) Then Err.Raise 7101
If Forms!StandaloneOpenForm.OpenArgs<>"" Then Err.Raise 7102
If Forms!StandaloneOpenForm!n<>20 Then Err.Raise 7103
CurrentDb.Execute "INSERT INTO macro_argument_events VALUES(1)"
ObserveMacroOpening=1
End Function)VBA"}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();
    const auto press=[&](QDialog* dialog,const QString& text) {for(auto* button:dialog->findChildren<QPushButton*>()) if(button->text()==text) {button->click();return;}assert(false);};
    QTimer::singleShot(0,&window,[&] {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-macro-designer");dialog->findChild<QLineEdit*>("native-macro-name")->setText("OpeningArguments");auto* list=dialog->findChild<QTableWidget*>("native-macro-actions");qobject_cast<QComboBox*>(list->cellWidget(0,0))->setCurrentText("OpenForm");QTimer::singleShot(0,dialog,[] {auto* editor=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(editor && editor->objectName()=="native-macro-argument-dialog");editor->findChild<QPlainTextEdit*>("native-macro-argument-1")->setPlainText("StandaloneOpenForm");editor->findChild<QPlainTextEdit*>("native-macro-argument-4")->setPlainText("id=2");editor->findChild<QLineEdit*>("native-macro-argument-5")->setText("2");editor->findChild<QLineEdit*>("native-macro-argument-6")->setText("1");editor->findChild<QCheckBox*>("native-macro-argument-supplied-7")->setChecked(true);assert(editor->findChild<QPlainTextEdit*>("native-macro-argument-7")->toPlainText().isEmpty());editor->accept();});dialog->findChild<QPushButton*>("native-macro-arguments")->click();list->item(0,4)->setText("True");
        dialog->findChild<QPushButton*>("native-macro-add")->click();qobject_cast<QComboBox*>(list->cellWidget(1,0))->setCurrentText("RunCode");list->item(1,1)->setText("ObserveMacroOpening()");dialog->findChild<QPushButton*>("native-macro-add")->click();qobject_cast<QComboBox*>(list->cellWidget(2,0))->setCurrentText("Close");list->item(2,1)->setText("2");list->item(2,2)->setText("StandaloneOpenForm");list->item(2,3)->setText("2");list->setCurrentCell(0,1);press(dialog,"Nach unten");press(dialog,"Nach oben");dialog->findChild<QCheckBox*>("native-macro-group-mode")->setChecked(true);dialog->findChild<QLineEdit*>("native-macro-group-name")->setText("Open");dialog->findChild<QPushButton*>("native-macro-group-add")->click();dialog->findChild<QLineEdit*>("native-macro-group-name")->setText("Other");qobject_cast<QComboBox*>(list->cellWidget(0,0))->setCurrentText("StopMacro");dialog->findChild<QComboBox*>("native-macro-group")->setCurrentIndex(0);dialog->findChild<QPushButton*>("native-macro-save")->click();assert(!dialog->isWindowModified());dialog->reject();});action(window,"+ Makro")->trigger();
    const auto definition=[&] {auto revision=linaccess::readNativeMacroRevision(db,"OpeningArguments");assert(revision);return QJsonDocument::fromJson(revision->definition.toByteArray()).object();};auto original=definition();const auto arguments=original.value("groups").toArray()[0].toObject().value("actions").toArray()[0].toObject().value("arguments").toArray();assert(arguments.size()==7 && arguments[1].isNull() && arguments[2].isNull() && arguments[6].isString() && arguments[6].toString().isEmpty());auto* explorer=window.findChild<QTreeWidget*>("database-explorer");QTreeWidgetItem* item=nullptr;for(auto* candidate:explorer->findItems("OpeningArguments",Qt::MatchExactly|Qt::MatchRecursive,0)) if(candidate->data(0,Qt::UserRole).toString()=="macro") item=candidate;assert(item);QTimer::singleShot(0,&window,[&] {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());auto* list=dialog->findChild<QTableWidget*>("native-macro-actions");list->setCurrentCell(0,1);QTimer::singleShot(0,dialog,[] {auto* editor=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(editor->findChild<QCheckBox*>("native-macro-argument-supplied-7")->isChecked() && editor->findChild<QPlainTextEdit*>("native-macro-argument-7")->toPlainText().isEmpty());editor->accept();});dialog->findChild<QPushButton*>("native-macro-arguments")->click();dialog->findChild<QPushButton*>("native-macro-save")->click();assert(!dialog->isWindowModified());dialog->reject();});QMetaObject::invokeMethod(explorer,"itemDoubleClicked",Qt::DirectConnection,Q_ARG(QTreeWidgetItem*,item),Q_ARG(int,0));assert(definition()==original);
    linaccess::NativeStandardModuleHost host(&window,db,[] {});auto commands=host.applicationObject()->member("DoCmd",{}).object;commands->member("RunMacro",{VbaValue::string("OpeningArguments.Open")});{AccessFormPreview caller("macro_argument_caller","Begin Form\nBegin Section\nEnd\nEnd",{},db,&window);caller.show();QApplication::processEvents();makePreviewApplicationObject(caller)->member("DoCmd",{}).object->member("RunMacro",{VbaValue::string("OpeningArguments.Open")});assert(caller.close());}
    const auto row=[](QString action,QJsonArray args) {return QJsonObject{{"action",action},{"arguments",args}};};const auto missing=QJsonValue(QJsonValue::Null);linaccess::storeNativeMacro(db,"LegacyArgumentDefaults",{{"version",1},{"actions",QJsonArray{row("OpenTable",{"opening_rows",missing,2}),row("GoToRecord",{0,"opening_rows",4,2}),row("RunSQL",{"INSERT INTO macro_argument_events VALUES(3)",missing}),row("Requery",{missing}),row("SelectObject",{0,"opening_rows",missing}),row("MessageBox",{"optional defaults",false,missing,"omitted type"}),row("Close",{0,"opening_rows",missing})}}});int dialogs=0;QTimer close;QObject::connect(&close,&QTimer::timeout,&window,[&] {auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());if(!box) return;assert(box->objectName()=="native-macro-message-box" && box->windowTitle()=="omitted type" && box->icon()==QMessageBox::NoIcon);++dialogs;box->accept();});close.start(1);commands->member("RunMacro",{VbaValue::string("LegacyArgumentDefaults")});{AccessFormPreview caller("macro_default_caller","Begin Form\nBegin Section\nEnd\nEnd",{},db,&window);caller.show();QApplication::processEvents();makePreviewApplicationObject(caller)->member("DoCmd",{}).object->member("RunMacro",{VbaValue::string("LegacyArgumentDefaults")});assert(caller.close());}close.stop();assert(dialogs==2);assert(sql.exec("SELECT COUNT(*) FROM macro_argument_events") && sql.next() && sql.value(0).toInt()==4);sql.finish();assert(sql.exec("DELETE FROM _la_objects WHERE kind='module' AND name='MacroArgumentCode'"));
}

void testStandaloneDocumentOpening(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;QSqlQuery sql(db);assert(sql.exec("CREATE TABLE opening_rows(id INTEGER PRIMARY KEY,n INTEGER)"));assert(sql.exec("INSERT INTO opening_rows VALUES(1,10),(2,20),(3,30)"));sql.finish();const auto store=[&](QString kind,QString name,QJsonObject definition) {sql.prepare("INSERT INTO _la_objects VALUES(?,?,?)");sql.addBindValue(kind);sql.addBindValue(name);sql.addBindValue(QJsonDocument(definition).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();};
    auto definition=QJsonObject{{"table","opening_rows"},{"fields",QJsonArray{"n"}},{"vba","Private Sub Form_Load()\nMe.Tag=Nz(Me.OpenArgs,\"no args\")\nEnd Sub"}};store("form","StandaloneOpenForm",definition);store("report","StandaloneOpenReport",{{"sql","SELECT id,n FROM opening_rows"},{"title","Native opening reference"}});store("module","StandaloneOpeningCode",{{"version",1},{"vba",R"VBA(Public Sub OpenStandaloneFiltered()
Application.DoCmd.OpenForm FormName:="StandaloneOpenForm",WhereCondition:="id=2",DataMode:=acFormReadOnly,WindowMode:=acHidden,OpenArgs:="module argument"
End Sub)VBA"}});
    linaccess::NativeStandardModuleHost host(&window,db,[] {});auto application=host.applicationObject();auto commands=application->member("DoCmd",{}).object;auto forms=application->member("Forms",{}).object;const auto before=forms->member("Count",{}).number;application->member("Run",{VbaValue::string("OpenStandaloneFiltered")});auto root=forms->member("Item",{VbaValue::string("StandaloneOpenForm")}).object;auto field=root->member("Controls",{VbaValue::string("n")}).object;assert(!root->member("Visible",{}).boolean && !root->member("AllowEdits",{}).boolean && root->member("Tag",{}).asText()=="module argument" && field->scalarValue().number==20);
    commands->member("OpenForm",{VbaValue::string("StandaloneOpenForm"),VbaValue::missing(),VbaValue::missing(),VbaValue::string("id=3"),VbaValue::missing(),VbaValue::integer(1)});assert(field->scalarValue().number==30 && root->member("OpenArgs",{}).asText()=="module argument");int error=0;try {commands->member("OpenForm",{VbaValue::string("StandaloneOpenForm"),VbaValue::missing(),VbaValue::missing(),VbaValue::string("id=1"),VbaValue::integer(1)});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==3251 && field->scalarValue().number==30);commands->member("Close",{VbaValue::integer(2),VbaValue::string("StandaloneOpenForm"),VbaValue::integer(2)});
    commands->member("OpenForm",{VbaValue::string("StandaloneOpenForm"),VbaValue::missing(),VbaValue::missing(),VbaValue::empty(),VbaValue::integer(0),VbaValue::integer(1),VbaValue{}});root=forms->member("Item",{VbaValue::string("StandaloneOpenForm")}).object;assert(root->member("DataEntry",{}).boolean && root->member("NewRecord",{}).number==-1 && root->member("OpenArgs",{}).kind==VbaValue::Kind::Null);root->member("Controls",{VbaValue::string("n")}).object->assign("Value",VbaValue::integer(40));root->assign("Dirty",VbaValue::logical(false));assert(sql.exec("SELECT COUNT(*) FROM opening_rows WHERE n=40") && sql.next() && sql.value(0).toInt()==1);sql.finish();commands->member("Close",{VbaValue::integer(2),VbaValue::string("StandaloneOpenForm"),VbaValue::integer(2)});
    bool waited=false;QTimer::singleShot(0,&window,[&] {auto* form=dynamic_cast<AccessFormPreview*>(QApplication::activeModalWidget());assert(form && form->documentName()=="StandaloneOpenForm");assert(makePreviewFormObject(*form)->member("OpenArgs",{}).asText()=="modal argument");waited=true;assert(form->close());});commands->member("OpenForm",{VbaValue::string("StandaloneOpenForm"),VbaValue::missing(),VbaValue::missing(),VbaValue::string("id=1"),VbaValue::integer(2),VbaValue::integer(3),VbaValue::string("modal argument")});assert(waited && forms->member("Count",{}).number==before);
    commands->member("OpenReport",{VbaValue::string("StandaloneOpenReport"),VbaValue::integer(2),VbaValue::missing(),VbaValue::string("id=2"),VbaValue::integer(1),VbaValue::string("report argument")});auto report=application->member("Reports",{VbaValue::string("StandaloneOpenReport")}).object;assert(report->member("OpenArgs",{}).asText()=="report argument");commands->member("Close",{VbaValue::integer(3),VbaValue::string("StandaloneOpenReport"),VbaValue::integer(2)});
    definition.insert("vba","Private Sub Form_Load()\nErr.Raise 7006,\"OpenSurface\",\"init marker\",\"opening.chm\",9\nEnd Sub");store("form","StandaloneOpenError",definition);error=0;try {commands->member("OpenForm",{VbaValue::string("StandaloneOpenError")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;assert(failure.source=="OpenSurface" && failure.helpFile=="opening.chm" && failure.helpContext==9);}assert(error==7006 && forms->member("Count",{}).number==before);
    artifact(db,"form","StandaloneImportedOpen","Begin Form\nOnLoad =\"[Event Procedure]\"\nBegin Section\nEnd\nEnd");artifact(db,"vba","Form_StandaloneImportedOpen","Private Sub Form_Load()\nMe.Tag=\"imported native opening\"\nEnd Sub");commands->member("OpenForm",{VbaValue::string("StandaloneImportedOpen")});assert(forms->member("Item",{VbaValue::string("StandaloneImportedOpen")}).object->member("Tag",{}).asText()=="imported native opening");commands->member("Close",{VbaValue::integer(2),VbaValue::string("StandaloneImportedOpen"),VbaValue::integer(2)});assert(sql.exec("DELETE FROM _la_objects WHERE kind='module' AND name='StandaloneOpeningCode'"));assert(sql.exec("DELETE FROM opening_rows WHERE n=40"));
}

void testStandaloneApplicationServices(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;QSqlQuery sql(db);sql.prepare("INSERT INTO _la_objects VALUES('module','StandaloneServicesCode',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba",R"VBA(Public Sub StandaloneServices()
Dim database As DAO.Database
Set database=Application.CurrentDb()
If database.OpenRecordset("SELECT Count(*) AS n FROM module_nav_rows").Fields(0).Value<>3 Then Err.Raise 5
If Application.DSum(Expr:="id",Domain:="module_nav_rows")<>6 Then Err.Raise 5
If Application.DLookup(Expr:="id",Domain:="module_nav_rows",Criteria:="id=2")<>2 Then Err.Raise 5
If Application.DCount("*","module_nav_rows")<>3 Then Err.Raise 5
If Application.DMin("id","module_nav_rows")<>1 Then Err.Raise 5
If Application.DMax("id","module_nav_rows")<>3 Then Err.Raise 5
SysCmd Action:=acSysCmdSetStatus,Argument2:="native project status"
Application.SysCmd Action:=acSysCmdInitMeter,Argument2:="native project meter",Argument3:=4
Application.SysCmd acSysCmdUpdateMeter,3
End Sub)VBA"}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();linaccess::NativeStandardModuleHost host(&window,db,[] {});auto application=host.applicationObject();application->member("Run",{VbaValue::string("StandaloneServices")});auto* meter=window.statusBar()->findChild<QProgressBar*>("access-syscmd-meter",Qt::FindDirectChildrenOnly);assert(window.statusBar()->currentMessage()=="native project meter" && meter && meter->maximum()==4 && meter->value()==3);assert(!window.findChild<QStatusBar*>("access-status-bar",Qt::FindDirectChildrenOnly));
    application->member("SysCmd",{VbaValue::integer(3)});assert(window.statusBar()->currentMessage().isEmpty() && !window.statusBar()->findChild<QProgressBar*>("access-syscmd-meter"));assert(application->member("SysCmd",{VbaValue::integer(10),VbaValue::integer(0),VbaValue::string("module_nav_rows")}).number==0);auto commands=application->member("DoCmd",{}).object;commands->member("OpenTable",{VbaValue::string("module_nav_rows")});assert(static_cast<int>(application->member("SysCmd",{VbaValue::integer(10),VbaValue::integer(0),VbaValue::string("module_nav_rows")}).number)&1);commands->member("Close",{VbaValue::integer(0),VbaValue::string("module_nav_rows")});
    class Action final : public linaccess::VbaObject {public:int reads{};bool hasScalarDefault() const override {return true;}VbaValue scalarValue() override {++reads;return VbaValue::integer(4);}VbaValue member(std::string_view,const std::vector<VbaValue>&) override {throw linaccess::VbaRuntimeError(438,"No explicit member");}};auto action=std::make_shared<Action>();application->member("SysCmd",{VbaValue::reference(action),VbaValue::string("single evaluation")});assert(action->reads==1 && window.statusBar()->currentMessage()=="single evaluation");{AccessFormPreview context("StatusActionContext","Begin Form\nBegin Section\nEnd\nEnd",{},db,&window);action->reads=0;makePreviewApplicationObject(context)->member("SysCmd",{VbaValue::reference(action),VbaValue::string("form single evaluation")});assert(action->reads==1 && window.statusBar()->currentMessage()=="form single evaluation");assert(context.close());}application->member("Echo",{VbaValue::logical(true),VbaValue::string("actual project echo status")});assert(window.statusBar()->currentMessage()=="actual project echo status");application->member("SysCmd",{VbaValue::integer(5)});assert(window.statusBar()->currentMessage().isEmpty());assert(sql.exec("DELETE FROM _la_objects WHERE kind='module' AND name='StandaloneServicesCode'"));
}

void testStandaloneFileDialog(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;QTemporaryDir directory;assert(directory.isValid());const auto path=directory.filePath("native-picker-reference.txt");{QFile file(path);assert(file.open(QIODevice::WriteOnly) && file.write("native picker reference\n")>0);}
    linaccess::NativeStandardModuleHost host(&window,db,[] {});auto application=host.applicationObject();auto picker=application->member("FileDialog",{VbaValue::integer(3)}).object;picker->assign("Title",VbaValue::string("Native module file selection"));picker->assign("InitialFileName",VbaValue::string(path.toStdString()));picker->assign("AllowMultiSelect",VbaValue::logical(false));auto filters=picker->member("Filters",{}).object;filters->member("Clear",{});filters->member("Add",{VbaValue::string("Technical reference"),VbaValue::string("*.txt")});auto selected=picker->member("SelectedItems",{}).object;
    const auto choose=[&](bool accept,bool folder) {QTimer::singleShot(0,&window,[&,accept,folder] {auto* dialog=qobject_cast<QFileDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="access-file-dialog");assert(dialog->fileMode()==(folder ? QFileDialog::Directory : QFileDialog::ExistingFile));if(!accept) {dialog->reject();return;}dialog->selectFile(folder ? directory.path() : path);QMetaObject::invokeMethod(dialog,"accept",Qt::DirectConnection);});};
    choose(true,false);assert(picker->member("Show",{}).number==-1 && selected->member("Count",{}).number==1 && selected->member("Item",{VbaValue::integer(1)}).asText()==path.toStdString());choose(false,false);assert(picker->member("Show",{}).number==0 && selected->member("Count",{}).number==0);assert(application->member("FileDialog",{VbaValue::integer(4)}).object==picker);picker->assign("InitialFileName",VbaValue::string(directory.path().toStdString()));choose(true,true);assert(picker->member("Show",{}).number==-1 && selected->member("Item",{VbaValue::integer(1)}).asText()==directory.path().toStdString());
    QSqlQuery sql(db);sql.prepare("INSERT INTO _la_objects VALUES('module','StandalonePickerCode',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba",R"VBA(Public Function StandalonePick(ByVal initial As String) As String
Dim picker As Office.FileDialog
Set picker=Application.FileDialog(DialogType:=msoFileDialogFilePicker)
picker.InitialFileName=initial
picker.AllowMultiSelect=False
If picker.Show=-1 Then
StandalonePick=picker.SelectedItems(1)
End If
End Function)VBA"}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();choose(true,false);assert(application->member("Run",{VbaValue::string("StandalonePick"),VbaValue::string(path.toStdString())}).asText()==path.toStdString());choose(false,false);assert(application->member("Run",{VbaValue::string("StandalonePick"),VbaValue::string(path.toStdString())}).asText().empty());
    int error=0;try {(void)application->member("FileDialog",{VbaValue::integer(1)});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==3251);error=0;try {(void)application->member("FileDialog",{VbaValue{}});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==94);assert(sql.exec("DELETE FROM _la_objects WHERE kind='module' AND name='StandalonePickerCode'"));
}

void testProjectDocumentLifetime(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;linaccess::NativeStandardModuleHost host(&window,db,[] {});auto application=host.applicationObject();auto forms=application->member("Forms",{}).object;auto reports=application->member("Reports",{}).object;auto screen=application->member("Screen",{}).object;const auto before=forms->member("Count",{}).number;
    application.reset();const auto result=linaccess::executeVbaProcedure(R"VBA(Public Function GlobalDocumentIdentity() As Boolean
Dim first As Object
Dim second As Object
Dim display As Object
Set first=Forms
Set second=Application.Forms
Set display=Screen
GlobalDocumentIdentity=(first Is second) And (display Is Application.Screen)
End Function)VBA","GlobalDocumentIdentity",host);assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue && result.typedReturnValue->kind==VbaValue::Kind::Boolean && result.typedReturnValue->boolean);application=host.applicationObject();assert(application->member("Forms",{}).object==forms && application->member("Reports",{}).object==reports && application->member("Screen",{}).object==screen);
    AccessFormPreview pending("LoadedUnshownDocument","Begin Form\nBegin Section\nEnd\nEnd",{},db,&window);assert(pending.isHidden() && forms->member("Count",{}).number==before+1);assert(pending.close() && forms->member("Count",{}).number==before);AccessFormPreview hidden("HiddenNavigationDocument","Begin Form\nRecordSource =\"module_nav_rows\"\nBegin Section\nEnd\nEnd",{},db,&window);hidden.show();QApplication::processEvents();hidden.hide();window.focusDatabaseWindow();assert(hidden.isHidden() && hidden.hasDataSource() && forms->member("Count",{}).number==before+1);auto root=forms->member("Item",{VbaValue::string("HiddenNavigationDocument")}).object;auto commands=application->member("DoCmd",{}).object;commands->member("GoToRecord",{VbaValue::integer(2),VbaValue::string("hiddennavigationdocument"),VbaValue::integer(4),VbaValue::integer(2)});assert(root->member("CurrentRecord",{}).number==2 && hidden.isHidden());commands->member("Close",{VbaValue::integer(2),VbaValue::string("hiddennavigationdocument"),VbaValue::integer(2)});assert(forms->member("Count",{}).number==before);int error=0;try {(void)root->member("CurrentRecord",{});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==2450);
}

void testStandaloneDocuments(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;linaccess::NativeStandardModuleHost host(&window,db,[] {});auto application=host.applicationObject();auto forms=application->member("Forms",{}).object;auto reports=application->member("Reports",{}).object;auto screen=application->member("Screen",{}).object;const auto formCount=forms->member("Count",{}).number,reportCount=reports->member("Count",{}).number;
    AccessFormPreview first("StandaloneFirstForm","Begin Form\nWidth =2400\nBegin Section\nHeight =600\nBegin TextBox\nName =\"Value\"\nWidth =1800\nHeight =300\nEnd\nEnd\nEnd",{},db,&window);first.show();AccessFormPreview hidden("StandaloneHiddenForm","Begin Form\nBegin Section\nEnd\nEnd",{},db,&window);hidden.show();QApplication::processEvents();hidden.hide();AccessFormPreview report("StandaloneReport","Begin Report\nBegin Section\nEnd\nEnd",{},db,&window);report.show();QApplication::processEvents();assert(forms->member("Count",{}).number==formCount+2 && reports->member("Count",{}).number==reportCount+1);auto firstObject=makePreviewFormObject(first);assert(application->member("Forms",{VbaValue::string("standalonefirstform")}).object==firstObject && forms->member("Item",{VbaValue::string("StandaloneHiddenForm")}).object==makePreviewFormObject(hidden));assert(forms->member("Item",{VbaValue::numeric(formCount)}).object==firstObject);assert(reports->member("Item",{VbaValue::string("StandaloneReport")}).object==makePreviewFormObject(report));
    auto commands=application->member("DoCmd",{}).object;commands->member("SelectObject",{VbaValue::integer(2),VbaValue::string("StandaloneFirstForm"),VbaValue::logical(false)});commands->member("GoToControl",{VbaValue::string("Value")});assert(screen->member("ActiveForm",{}).object==firstObject && screen->member("ActiveControl",{}).object->member("Name",{}).asText()=="Value");
    QSqlQuery sql(db);sql.prepare("INSERT INTO _la_objects VALUES('module','StandaloneDocumentCode',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba",R"VBA(Public Function StandaloneDocumentsCount() As Long
Dim document As Object
Dim count As Long
For Each document In Forms
count=count+1
Next document
For Each document In Reports
count=count+1
Next document
StandaloneDocumentsCount=count
End Function
Public Function StandaloneScreenName() As String
StandaloneScreenName=Screen.ActiveForm.Name & ":" & Screen.ActiveControl.Name
End Function
Public Function StandaloneMousePointer() As Long
Screen.MousePointer=vbHourglass
StandaloneMousePointer=Screen.MousePointer
End Function)VBA"}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();assert(application->member("Run",{VbaValue::string("StandaloneDocumentsCount")}).number==formCount+reportCount+3);assert(application->member("Run",{VbaValue::string("StandaloneScreenName")}).asText()=="StandaloneFirstForm:Value");assert(application->member("Run",{VbaValue::string("StandaloneMousePointer")}).number==11 && !QApplication::overrideCursor());
    assert(first.close());int error=0;try {(void)firstObject->member("Name",{});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==2450 && forms->member("Count",{}).number==formCount+1);assert(hidden.close() && report.close());assert(forms->member("Count",{}).number==formCount && reports->member("Count",{}).number==reportCount);window.focusDatabaseWindow();error=0;try {(void)screen->member("ActiveForm",{});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==2450);assert(sql.exec("DELETE FROM _la_objects WHERE kind='module' AND name='StandaloneDocumentCode'"));
}

void testStandaloneProjectCatalog(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;QSqlQuery sql(db);sql.prepare("INSERT INTO _la_objects VALUES('module','StandaloneCatalogCode',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba","Public Function StandaloneCatalogProbe() As Long\nDim project As Object\nSet project=CurrentProject\nStandaloneCatalogProbe=CurrentData.AllTables.Count+project.AllModules.Count\nEnd Function"}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();linaccess::NativeStandardModuleHost host(&window,db,[] {});auto application=host.applicationObject();auto project=application->member("CurrentProject",{}).object;auto data=application->member("CurrentData",{}).object;assert(project && data && project==application->member("CurrentProject",{}).object && data==application->member("CurrentData",{}).object);auto modules=project->member("AllModules",{}).object;auto tables=data->member("AllTables",{}).object;const auto count=modules->member("Count",{}).number+tables->member("Count",{}).number;assert(count>0 && application->member("Run",{VbaValue::string("StandaloneCatalogProbe")}).number==count);auto descriptor=modules->member("Item",{VbaValue::string("standalonecatalogcode")}).object;assert(descriptor->member("Name",{}).asText()=="StandaloneCatalogCode");assert(tables->member("Item",{VbaValue::string("module_find_rows")}).object->member("Name",{}).asText()=="module_find_rows");
    {AccessFormPreview form("StandaloneCatalogIdentity","Begin Form\nBegin Section\nEnd\nEnd",{},db,&window);form.show();QApplication::processEvents();auto formApplication=makePreviewApplicationObject(form);assert(formApplication->member("CurrentProject",{}).object==project && formApplication->member("CurrentData",{}).object==data);assert(form.close());}assert(modules->member("Count",{}).number+tables->member("Count",{}).number==count);
    int error=0;try {(void)application->member("CurrentProject",{VbaValue::integer(1)});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==450);error=0;try {project->assign("Name",VbaValue::string("renamed"));}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==438);assert(sql.exec("DELETE FROM _la_objects WHERE kind='module' AND name='StandaloneCatalogCode'"));assert(modules->member("Count",{}).number+tables->member("Count",{}).number==count-1);
}

void testModuleUiCommands(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;QSqlQuery sql(db);sql.prepare("INSERT INTO _la_objects VALUES('module','UiExecutionModule',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba",R"VBA(Public Function UiCursorNested() As Long
UiCursorNested=MsgBox("module cursor",vbOKOnly)
End Function
Public Function UiCursor() As Long
Application.DoCmd.Hourglass HourglassOn:=True
UiCursor=Application.Run("UiCursorNested")
End Function
Public Function UiCursorFail() As Long
Application.DoCmd.Hourglass True
Err.Raise 7004,"UiExecutionModule","cursor failure"
End Function
Public Function UiCursorCheck() As Long
UiCursorCheck=MsgBox("macro cursor",vbOKOnly)
End Function
Public Sub UiSelectFocus()
Application.DoCmd.GoToControl ControlName:="[Value]"
End Sub)VBA"}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();
    linaccess::NativeStandardModuleHost host(&window,db,[] {});auto application=host.applicationObject();auto commands=application->member("DoCmd",{}).object;int seen=0;QTimer probe;QObject::connect(&probe,&QTimer::timeout,&window,[&] {auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());if(!box) return;assert((box->text()=="module cursor" || box->text()=="macro cursor" || box->text()=="macro cursor after") && linaccess::nativeBusyCursorOnTop());++seen;box->accept();});probe.start(1);
    assert(application->member("Run",{VbaValue::string("UiCursor")}).number==1 && !QApplication::overrideCursor());int error=0;try {application->member("Run",{VbaValue::string("UiCursorFail")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==7004 && !QApplication::overrideCursor());
    AccessFormPreview form("UiExecutionForm","Begin Form\nWidth =2400\nBegin Section\nHeight =600\nBegin TextBox\nName =\"Value\"\nWidth =1800\nHeight =300\nEnd\nEnd\nEnd",{},db,&window);form.show();QApplication::processEvents();auto* editor=form.findChild<QLineEdit*>("access-control-Value");assert(editor);editor->setText("focus text");application->member("Run",{VbaValue::string("UiSelectFocus")});assert(editor->hasFocus() && editor->selectedText()=="focus text");
    const auto action=[](QString name,QJsonArray arguments=QJsonArray{}) {return QJsonObject{{"action",name},{"arguments",arguments}};};linaccess::storeNativeMacro(db,"UiCursorMacro",{{"version",1},{"actions",QJsonArray{action("GoToControl",{"[Value]"}),action("Hourglass",{true}),action("RunCode",{"UiCursorCheck()"}),action("MessageBox",{"macro cursor after",false}),action("StopMacro")}}});commands->member("RunMacro",{VbaValue::string("UiCursorMacro")});assert(!QApplication::overrideCursor());makePreviewApplicationObject(form)->member("DoCmd",{}).object->member("RunMacro",{VbaValue::string("UiCursorMacro")});assert(!QApplication::overrideCursor());probe.stop();assert(seen==5);
    commands->member("Hourglass",{VbaValue::logical(true)});assert(linaccess::nativeBusyCursorOnTop());commands->member("Hourglass",{VbaValue::logical(false)});assert(!QApplication::overrideCursor());QApplication::setOverrideCursor(Qt::CrossCursor);error=0;try {commands->member("Hourglass",{VbaValue::logical(true)});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==3251 && QApplication::overrideCursor()->shape()==Qt::CrossCursor);QApplication::restoreOverrideCursor();assert(form.close());window.focusDatabaseWindow();error=0;try {commands->member("GoToControl",{VbaValue::string("Value")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==2046);assert(sql.exec("DELETE FROM _la_objects WHERE kind='module' AND name='UiExecutionModule'"));
}

void testModuleFindRecord(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;QSqlQuery sql(db);assert(sql.exec("CREATE TABLE module_find_rows(id INTEGER PRIMARY KEY,value TEXT,money CURRENCY_REAL)"));assert(sql.exec("INSERT INTO module_find_rows VALUES(1,'Alpha',1.2345),(2,'needle',2.0),(3,'needle',3.0)"));sql.finish();
    QTimer::singleShot(0,&window,[&] {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-macro-designer");dialog->findChild<QLineEdit*>("native-macro-name")->setText("FindArguments");auto* list=dialog->findChild<QTableWidget*>("native-macro-actions");qobject_cast<QComboBox*>(list->cellWidget(0,0))->setCurrentText("FindRecord");QTimer::singleShot(0,dialog,[] {auto* editor=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(editor && editor->objectName()=="native-macro-argument-dialog");editor->findChild<QPlainTextEdit*>("native-macro-argument-1")->setPlainText("needle");const QStringList options{"","Nein","2","Nein","0","Nein"};for(int index=0;index<options.size();++index) editor->findChild<QLineEdit*>("native-macro-argument-"+QString::number(index+2))->setText(options[index]);editor->accept();});dialog->findChild<QPushButton*>("native-macro-arguments")->click();dialog->findChild<QPushButton*>("native-macro-save")->click();assert(!dialog->isWindowModified());dialog->reject();});action(window,"+ Makro")->trigger();
    auto revision=linaccess::readNativeMacroRevision(db,"FindArguments");assert(revision);const auto args=QJsonDocument::fromJson(revision->definition.toByteArray()).object().value("actions").toArray()[0].toObject().value("arguments").toArray();assert(args.size()==7 && args[1].isNull() && args[6]==false);
    linaccess::NativeStandardModuleHost host(&window,db,[] {});auto application=host.applicationObject();auto commands=application->member("DoCmd",{}).object;commands->member("OpenTable",{VbaValue::string("module_find_rows")});auto* tabs=window.findChild<QTabWidget*>();const auto view=[&] {auto* table=tabs->currentWidget()->findChild<QTableView*>();assert(table);return table;};assert(view()->currentIndex().isValid());view()->sortByColumn(0,Qt::AscendingOrder);view()->setCurrentIndex(view()->model()->index(0,0));
    commands->member("FindRecord",{VbaValue::integer(3)});assert(view()->currentIndex().row()==2);view()->setCurrentIndex(view()->model()->index(0,1));commands->member("FindRecord",{VbaValue::string("need*")});assert(view()->currentIndex().row()==1);commands->member("FindRecord",{VbaValue::string("needle"),VbaValue::missing(),VbaValue::missing(),VbaValue::missing(),VbaValue::missing(),VbaValue::missing(),VbaValue::logical(false)});assert(view()->currentIndex().row()==2);
    view()->setCurrentIndex(view()->model()->index(2,2));commands->member("FindRecord",{VbaValue::currency(12345)});assert(view()->currentIndex().row()==0);view()->setCurrentIndex(view()->model()->index(0,0));commands->member("RunMacro",{VbaValue::string("FindArguments")});assert(view()->currentIndex().row()==1 && view()->currentIndex().column()==1);
    int error=0;try {commands->member("FindRecord",{VbaValue::string("[invalid")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==93 && view()->currentIndex().row()==1);error=0;try {commands->member("FindRecord",{VbaValue{}});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==94 && view()->currentIndex().row()==1);
    auto database=linaccess::makeDaoDatabase(db);database->member("CreateQueryDef",{VbaValue::string("ModuleFindQuery"),VbaValue::string("SELECT id,value FROM module_find_rows ORDER BY id DESC")});commands->member("OpenQuery",{VbaValue::string("ModuleFindQuery"),VbaValue::integer(0),VbaValue::integer(2)});commands->member("FindRecord",{VbaValue::integer(1)});assert(view()->currentIndex().row()==2);
    sql.prepare("INSERT INTO _la_objects VALUES('module','ModuleFindProcedure',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba","Public Sub FindNamed()\nApplication.DoCmd.FindRecord FindWhat:=\"Alpha\",OnlyCurrentField:=acAll,FindFirst:=True\nEnd Sub"}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();application->member("Run",{VbaValue::string("FindNamed")});assert(view()->currentIndex().row()==2 && view()->currentIndex().column()==1);
    AccessFormPreview form("ModuleFindForm","Begin Form\nRecordSource =\"module_find_rows\"\nBegin Section\nBegin TextBox\nName =\"Value\"\nControlSource =\"value\"\nWidth =1800\nHeight =300\nEnd\nEnd\nEnd",{},db,&window);form.show();QApplication::processEvents();auto root=makePreviewFormObject(form);assert(form.hasDataSource());commands->member("RunMacro",{VbaValue::string("FindArguments")});assert(root->member("CurrentRecord",{}).number==2);makePreviewApplicationObject(form)->member("DoCmd",{}).object->member("RunMacro",{VbaValue::string("FindArguments")});assert(root->member("CurrentRecord",{}).number==3);assert(form.close());commands->member("Close",{VbaValue::integer(1),VbaValue::string("ModuleFindQuery")});commands->member("Close",{VbaValue::integer(0),VbaValue::string("module_find_rows")});assert(sql.exec("DELETE FROM _la_objects WHERE kind='module' AND name='ModuleFindProcedure'"));
}

void testModuleRecordNavigation(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;QSqlQuery sql(db);assert(sql.exec("CREATE TABLE module_nav_rows(id INTEGER PRIMARY KEY)"));assert(sql.exec("INSERT INTO module_nav_rows VALUES(1),(2),(3)"));sql.finish();auto database=linaccess::makeDaoDatabase(db);database->member("CreateQueryDef",{VbaValue::string("ModuleNavigationQuery"),VbaValue::string("SELECT id FROM module_nav_rows ORDER BY id DESC")});
    linaccess::NativeStandardModuleHost host(&window,db,[] {});auto application=host.applicationObject();auto commands=application->member("DoCmd",{}).object;commands->member("OpenTable",{VbaValue::string("module_nav_rows")});auto* tabs=window.findChild<QTabWidget*>();const auto currentRow=[&] {auto* view=tabs->currentWidget()->findChild<QTableView*>();assert(view);return view->currentIndex().row();};
    commands->member("GoToRecord",{VbaValue::integer(0),VbaValue::string("MODULE_NAV_ROWS"),VbaValue::integer(4),VbaValue::integer(3)});assert(currentRow()==2);commands->member("GoToRecord",{VbaValue::missing(),VbaValue::missing(),VbaValue::integer(2)});assert(currentRow()==0);commands->member("GoToRecord",{});assert(currentRow()==1);
    int error=0;try {commands->member("GoToRecord",{VbaValue::integer(0),VbaValue::string("module_nav_rows"),VbaValue::integer(4),VbaValue::integer(0)});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==2105 && currentRow()==1);
    commands->member("OpenQuery",{VbaValue::string("ModuleNavigationQuery"),VbaValue::integer(0),VbaValue::integer(2)});commands->member("GoToRecord",{VbaValue::integer(1),VbaValue::string("modulenavigationquery"),VbaValue::integer(3)});assert(currentRow()==2);commands->member("GoToRecord",{VbaValue::missing(),VbaValue::missing(),VbaValue::integer(0)});assert(currentRow()==1);
    sql.prepare("INSERT INTO _la_objects VALUES('module','NavigationProcedure',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba","Public Sub NavigateNamed()\nApplication.DoCmd.GoToRecord ObjectType:=acDataTable,ObjectName:=\"module_nav_rows\",Record:=acGoTo,Offset:=3\nEnd Sub"}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();application->member("Run",{VbaValue::string("NavigateNamed")});assert(currentRow()==2);
    const auto definition=QJsonObject{{"version",1},{"actions",QJsonArray{QJsonObject{{"action","GoToRecord"},{"arguments",QJsonArray{QJsonValue::Null,QJsonValue::Null,4,2}}}}}};linaccess::storeNativeMacro(db,"NavigateMacro",definition);commands->member("RunMacro",{VbaValue::string("NavigateMacro")});assert(currentRow()==1);
    AccessFormPreview form("ModuleNavigationForm","Begin Form\nRecordSource =\"module_nav_rows\"\nBegin Section\nEnd\nEnd",{},db,&window);form.show();QApplication::processEvents();auto root=makePreviewFormObject(form);assert(form.hasDataSource());commands->member("GoToRecord",{VbaValue::integer(2),VbaValue::string("modulenavigationform"),VbaValue::integer(4),VbaValue::integer(3)});assert(root->member("CurrentRecord",{}).number==3);commands->member("GoToRecord",{VbaValue::missing(),VbaValue::missing(),VbaValue::integer(2)});assert(root->member("CurrentRecord",{}).number==1);commands->member("RunMacro",{VbaValue::string("NavigateMacro")});assert(root->member("CurrentRecord",{}).number==2);
    makePreviewApplicationObject(form)->member("DoCmd",{}).object->member("RunMacro",{VbaValue::string("NavigateMacro")});assert(root->member("CurrentRecord",{}).number==2);assert(form.close());
    commands->member("Close",{VbaValue::integer(1),VbaValue::string("ModuleNavigationQuery")});commands->member("Close",{VbaValue::integer(0),VbaValue::string("module_nav_rows")});error=0;try {commands->member("GoToRecord",{VbaValue::integer(2),VbaValue::string("ModuleNavigationForm")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==2450);assert(sql.exec("DELETE FROM _la_objects WHERE kind='module' AND name='NavigationProcedure'"));
}

void testMacroMessageArguments(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-macro-designer");dialog->findChild<QLineEdit*>("native-macro-name")->setText("MessageArguments");auto* list=dialog->findChild<QTableWidget*>("native-macro-actions");qobject_cast<QComboBox*>(list->cellWidget(0,0))->setCurrentText("MessageBox");
        QTimer::singleShot(0,dialog,[] {auto* editor=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(editor && editor->objectName()=="native-macro-argument-dialog");editor->findChild<QPlainTextEdit*>("native-macro-argument-1")->setPlainText("=\"<Heading>@Body@Tail\"");editor->findChild<QLineEdit*>("native-macro-argument-2")->setText("Nein");editor->findChild<QPlainTextEdit*>("native-macro-argument-3")->setPlainText("Information");editor->findChild<QPlainTextEdit*>("native-macro-argument-4")->setPlainText("Four arguments");editor->accept();});dialog->findChild<QPushButton*>("native-macro-arguments")->click();
        QTimer::singleShot(0,dialog,[] {auto* editor=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(editor->findChild<QPlainTextEdit*>("native-macro-argument-4")->toPlainText()=="Four arguments");editor->findChild<QPlainTextEdit*>("native-macro-argument-4")->setPlainText("discarded");editor->reject();});dialog->findChild<QPushButton*>("native-macro-arguments")->click();
        auto* mode=dialog->findChild<QCheckBox*>("native-macro-group-mode");mode->setChecked(true);dialog->findChild<QLineEdit*>("native-macro-group-name")->setText("Message");dialog->findChild<QPushButton*>("native-macro-group-add")->click();dialog->findChild<QLineEdit*>("native-macro-group-name")->setText("Other");qobject_cast<QComboBox*>(list->cellWidget(0,0))->setCurrentText("RunSQL");list->item(0,1)->setText("DELETE FROM unused_reference\nWHERE 1=0");QTimer::singleShot(0,dialog,[] {auto* editor=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(editor->findChild<QPlainTextEdit*>("native-macro-argument-1")->toPlainText()=="DELETE FROM unused_reference\nWHERE 1=0");editor->accept();});dialog->findChild<QPushButton*>("native-macro-arguments")->click();dialog->findChild<QComboBox*>("native-macro-group")->setCurrentIndex(0);dialog->findChild<QPushButton*>("native-macro-save")->click();assert(!dialog->isWindowModified());dialog->reject();
    });action(window,"+ Makro")->trigger();
    auto revision=linaccess::readNativeMacroRevision(db,"MessageArguments");assert(revision);auto definition=QJsonDocument::fromJson(revision->definition.toByteArray()).object();auto args=definition.value("groups").toArray()[0].toObject().value("actions").toArray()[0].toObject().value("arguments").toArray();assert(args.size()==4 && args[3]=="Four arguments" && args[1]==false);
    int shown=0;QTimer close;QObject::connect(&close,&QTimer::timeout,&window,[&] {auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());if(!box || box->objectName()!="native-macro-message-box") return;assert(box->windowTitle()=="Four arguments" && box->icon()==QMessageBox::Information && box->textFormat()==Qt::RichText && box->text().contains("&lt;Heading&gt;") && box->text().contains("<br><br>Tail"));++shown;box->accept();});close.start(1);
    linaccess::NativeStandardModuleHost host(&window,db,[] {});auto commands=host.applicationObject()->member("DoCmd",{}).object;commands->member("RunMacro",{VbaValue::string("MessageArguments")});
    {AccessFormPreview context("message_argument_context","Begin Form\nBegin Section\nEnd\nEnd",{},db,&window);makePreviewApplicationObject(context)->member("DoCmd",{}).object->member("RunMacro",{VbaValue::string("MessageArguments")});assert(context.close());}close.stop();assert(shown==2);
    int error=0;try {(void)commands->member("MessageBox",{});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==438);
    auto invalid=definition;auto groups=invalid.value("groups").toArray();auto group=groups[0].toObject();auto actions=group.value("actions").toArray();auto entry=actions[0].toObject();args[2]="Invented";entry.insert("arguments",args);actions[0]=entry;group.insert("actions",actions);groups[0]=group;invalid.insert("groups",groups);error=0;try {(void)linaccess::parseNativeMacro(invalid);}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==5);
}

void testStopAllMacros(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;QSqlQuery sql(db);assert(sql.exec("CREATE TABLE stop_all_effects(n INTEGER)"));const auto action=[](QString name,QJsonArray args=QJsonArray{}) {return QJsonObject{{"action",name},{"arguments",args}};};
    const auto macro=[&](QString name,QJsonArray actions) {(void)linaccess::storeNativeMacro(db,name,{{"version",1},{"actions",actions}});};
    sql.prepare("INSERT INTO _la_objects VALUES('module','StopAllFunctions',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba",R"VBA(Public Function StopThroughFunction() As Long
StopThroughFunction=MsgBox("Stop all echo probe",vbOKOnly)
Application.DoCmd.RunMacro "StopAllChild"
CurrentDb.Execute "INSERT INTO stop_all_effects VALUES(3)"
End Function
Public Function StopFromPredicate() As Boolean
Application.DoCmd.RunMacro "StopAllChild"
StopFromPredicate=True
End Function
Public Function EchoConditionalProbe() As Long
EchoConditionalProbe=MsgBox("Echo conditional probe",vbOKOnly)
End Function
Public Function NamedApplicationEcho() As Long
Application.Echo EchoOn:=False,bstrStatusBarText:="named echo"
Application.Echo EchoOn:=True
NamedApplicationEcho=1
End Function)VBA"}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();
    macro("StopAllChild",{action("RunSQL",{"INSERT INTO stop_all_effects VALUES(2)"}),action("StopAllMacros"),action("RunSQL",{"INSERT INTO stop_all_effects VALUES(99)"})});macro("StopAllParent",{action("Echo",{false,"echo suppressed"}),action("RunSQL",{"INSERT INTO stop_all_effects VALUES(1)"}),action("RunCode",{"StopThroughFunction()"}),action("RunSQL",{"INSERT INTO stop_all_effects VALUES(4)"}),action("Echo",{true})});
    linaccess::NativeStandardModuleHost host(&window,db,[] {});auto application=host.applicationObject();auto commands=application->member("DoCmd",{}).object;assert(window.centralWidget()->updatesEnabled());int observed=0;QTimer probe;probe.setInterval(1);QObject::connect(&probe,&QTimer::timeout,&window,[&] {if(auto* message=qobject_cast<QMessageBox*>(QApplication::activeModalWidget())) {assert(message->text()=="Stop all echo probe" || message->text()=="Echo conditional probe");assert(!window.centralWidget()->updatesEnabled());++observed;message->button(QMessageBox::Ok)->click();}});probe.start();commands->member("RunMacro",{VbaValue::string("StopAllParent"),VbaValue::integer(5)});assert(window.centralWidget()->updatesEnabled() && !linaccess::nativeMacroStopAllPending() && !linaccess::nativeProjectMacrosActive(db));
    {AccessFormPreview caller("stop_all_form","Begin Form\nBegin Section\nEnd\nEnd",{},db,&window);caller.show();QApplication::processEvents();makePreviewApplicationObject(caller)->member("DoCmd",{}).object->member("RunMacro",{VbaValue::string("StopAllParent"),VbaValue::integer(5)});assert(window.centralWidget()->updatesEnabled() && !linaccess::nativeMacroStopAllPending());assert(caller.close());}
    macro("AfterStopAll",{action("RunSQL",{"INSERT INTO stop_all_effects VALUES(5)"})});commands->member("RunMacro",{VbaValue::string("AfterStopAll")});macro("PredicateStopBody",{action("RunSQL",{"INSERT INTO stop_all_effects VALUES(7)"})});commands->member("RunMacro",{VbaValue::string("PredicateStopBody"),VbaValue::missing(),VbaValue::string("StopFromPredicate()")});
    auto conditional=action("StopAllMacros");conditional.insert("condition","False");macro("FalseStopAll",{action("Echo",{false}),conditional,action("RunCode",{"EchoConditionalProbe()"}),action("Echo",{true}),action("RunSQL",{"INSERT INTO stop_all_effects VALUES(6)"})});commands->member("RunMacro",{VbaValue::string("FalseStopAll")});probe.stop();assert(observed==3 && window.centralWidget()->updatesEnabled());
    assert(sql.exec("SELECT n FROM stop_all_effects"));QVector<int> effects;while(sql.next()) effects.push_back(sql.value(0).toInt());assert((effects==QVector<int>{1,2,3,1,2,3,5,2,6}));sql.finish();assert(application->member("Run",{VbaValue::string("NamedApplicationEcho")}).number==1 && window.centralWidget()->updatesEnabled());int error=0;try {(void)commands->member("Echo",{VbaValue::logical(false),VbaValue::string("invalid\nstatus")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==5 && window.centralWidget()->updatesEnabled());error=0;try {(void)commands->member("StopAllMacros",{});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==438 && !linaccess::nativeMacroStopAllPending());assert(sql.exec("DELETE FROM _la_objects WHERE kind='module' AND name='StopAllFunctions'"));
}

void testMacroSelectObject(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;auto database=linaccess::makeDaoDatabase(db);(void)database->member("CreateQueryDef",{VbaValue::string("SelectionQuery"),VbaValue::string("SELECT id,n FROM CommandKeep")});QSqlQuery sql(db);sql.prepare("INSERT INTO _la_objects VALUES('module','SelectionModule',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba","Public Function SelectionMarker() As Long\nSelectionMarker=1\nEnd Function"}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();(void)linaccess::storeNativeMacro(db,"SelectionTarget",{{"version",1},{"actions",QJsonArray{QJsonObject{{"action","StopMacro"},{"arguments",QJsonArray{}}}}}});action(window,"Aktualisieren")->trigger();
    const auto generation=window.property("la_project_generation");linaccess::NativeStandardModuleHost host(&window,db,[&] {if(window.property("la_project_generation")!=generation) throw linaccess::VbaRuntimeError(3420,"Project changed");});auto commands=host.applicationObject()->member("DoCmd",{}).object;const auto select=[&](int type,const QString& name,bool inDatabase) {commands->member("SelectObject",{VbaValue::integer(static_cast<std::int16_t>(type)),VbaValue::string(name.toStdString()),VbaValue::logical(inDatabase)});};
    commands->member("OpenTable",{VbaValue::string("CommandKeep"),VbaValue::integer(0),VbaValue::integer(1)});commands->member("OpenQuery",{VbaValue::string("SelectionQuery"),VbaValue::integer(0),VbaValue::integer(2)});select(0,"commandkeep",false);auto* tabs=window.findChild<QTabWidget*>();assert(tabs->currentWidget()->property("la_project_object_kind").toString()=="table" && tabs->currentWidget()->property("la_project_object_name").toString()=="CommandKeep");select(1,"selectionquery",false);assert(tabs->currentWidget()->property("la_project_object_kind").toString()=="query" && tabs->currentWidget()->property("la_project_object_name").toString()=="SelectionQuery");
    commands->member("OpenForm",{VbaValue::string("CommandForm"),VbaValue::integer(0)});commands->member("OpenReport",{VbaValue::string("CommandReport"),VbaValue::integer(2)});select(2,"commandform",false);assert(activePreviewDocument() && activePreviewDocument()->documentName()=="CommandForm");select(3,"commandreport",false);assert(activePreviewDocument() && activePreviewDocument()->documentName()=="CommandReport");commands->member("Close",{VbaValue::integer(2),VbaValue::string("CommandForm"),VbaValue::integer(2)});commands->member("Close",{VbaValue::integer(3),VbaValue::string("CommandReport"),VbaValue::integer(2)});
    auto* explorer=window.findChild<QTreeWidget*>("database-explorer");select(4,"SelectionTarget",true);assert(explorer->currentItem() && explorer->currentItem()->data(0,Qt::UserRole).toString()=="macro" && explorer->currentItem()->text(0)=="SelectionTarget");action(window,"Aktualisieren")->trigger();assert(explorer->currentItem() && explorer->currentItem()->text(0)=="SelectionTarget");select(5,"SelectionModule",true);assert(explorer->currentItem()->data(0,Qt::UserRole).toString()=="module" && explorer->currentItem()->text(0)=="SelectionModule");select(4,"",true);assert(explorer->currentItem()->data(0,Qt::UserRole+4).toBool() && explorer->currentItem()->data(0,Qt::UserRole+2).toString()=="macro");
    (void)linaccess::storeNativeMacro(db,"SelectionStarter",{{"version",1},{"actions",QJsonArray{QJsonObject{{"action","SelectObject"},{"arguments",QJsonArray{5,"SelectionModule",true}}}}}});action(window,"Aktualisieren")->trigger();QTreeWidgetItem* starter=nullptr;for(auto* item:explorer->findItems("SelectionStarter",Qt::MatchExactly|Qt::MatchRecursive,0)) if(item->data(0,Qt::UserRole).toString()=="macro") starter=item;assert(starter);explorer->setCurrentItem(starter);window.focusDatabaseWindow();auto* run=window.findChild<QAction*>("native-macro-run");assert(run && run->isEnabled());run->trigger();assert(window.statusBar()->currentMessage()=="Makro ausgeführt: SelectionStarter" && explorer->currentItem() && explorer->currentItem()->data(0,Qt::UserRole).toString()=="module" && explorer->currentItem()->text(0)=="SelectionModule");
    {AccessFormPreview caller("selection_real_caller","Begin Form\nBegin Section\nEnd\nEnd",{},db,&window);caller.show();QApplication::processEvents();auto formCommands=makePreviewApplicationObject(caller)->member("DoCmd",{}).object;formCommands->member("SelectObject",{VbaValue::integer(4),VbaValue::string("SelectionTarget"),VbaValue::logical(true)});assert(explorer->currentItem() && explorer->currentItem()->text(0)=="SelectionTarget");formCommands->member("SelectObject",{VbaValue::integer(5),VbaValue::string("SelectionModule"),VbaValue::logical(true)});assert(explorer->currentItem() && explorer->currentItem()->text(0)=="SelectionModule");assert(caller.close());}
    const auto expect=[&](int type,const QString& name,bool inDatabase,int wanted) {int actual=0;try {select(type,name,inDatabase);}catch(const linaccess::VbaRuntimeError& failure) {actual=failure.number;}assert(actual==wanted);};expect(4,"MissingSelection",true,7874);expect(0,"NotOpenSelection",false,3251);expect(4,"SelectionTarget",false,3251);expect(99,"SelectionTarget",true,3251);commands->member("Close",{VbaValue::integer(1),VbaValue::string("SelectionQuery"),VbaValue::integer(2)});commands->member("Close",{VbaValue::integer(0),VbaValue::string("CommandKeep"),VbaValue::integer(2)});assert(sql.exec("DELETE FROM _la_objects WHERE kind='module' AND name='SelectionModule'"));
}

void testMacroProjectCommands(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;auto database=linaccess::makeDaoDatabase(db);QSqlQuery sql(db);assert(sql.exec("CREATE TABLE command_events(n INTEGER)"));
    const auto createTable=[&](const QString& name) {auto table=database->member("CreateTableDef",{VbaValue::string(name.toStdString())}).object;for(const auto* field:{"id","n"}) table->member("Fields",{}).object->member("Append",{table->member("CreateField",{VbaValue::string(field),VbaValue::integer(4)})});auto primary=table->member("CreateIndex",{VbaValue::string("PrimaryKey")}).object;primary->assign("Primary",VbaValue::logical(true));primary->member("Fields",{}).object->member("Append",{primary->member("CreateField",{VbaValue::string("id")})});table->member("Indexes",{}).object->member("Append",{VbaValue::reference(primary)});database->member("TableDefs",{}).object->member("Append",{VbaValue::reference(table)});};createTable("CommandBefore");createTable("CommandKeep");database->member("Execute",{VbaValue::string("INSERT INTO CommandBefore VALUES(1,1)")});database->member("Execute",{VbaValue::string("INSERT INTO CommandKeep VALUES(1,5)")});
    (void)database->member("CreateQueryDef",{VbaValue::string("CommandQueryBefore"),VbaValue::string("SELECT id,n FROM CommandKeep")});
    const auto store=[&](QString kind,QString name,QJsonObject definition) {QSqlQuery save(db);save.prepare("INSERT INTO _la_objects(kind,name,definition) VALUES(?,?,?)");save.addBindValue(kind);save.addBindValue(name);save.addBindValue(QJsonDocument(definition).toJson(QJsonDocument::Compact));assert(save.exec());};
    store("module","CommandCode",{{"version",1},{"vba",R"VBA(Public Function MacroCommandProbe() As Long
MacroCommandProbe=MsgBox("Macro command probe",vbOKOnly)
CurrentDb.Execute "INSERT INTO command_events VALUES(10)"
End Function)VBA"}});
    const auto formDefinition=QJsonObject{{"table","CommandKeep"},{"fields",QJsonArray{"n"}},{"vba","Private Sub Form_Unload(Cancel As Integer)\nCurrentDb.Execute \"INSERT INTO command_events VALUES(11)\"\nEnd Sub"}};
    store("form","CommandForm",formDefinition);store("report","CommandReport",{{"sql","SELECT n FROM CommandKeep"},{"title","Macro command output"},{"vba","Private Sub Report_Close()\nCurrentDb.Execute \"INSERT INTO command_events VALUES(12)\"\nEnd Sub"}});
    const auto macroAction=[](QString name,QJsonArray args=QJsonArray{}) {return QJsonObject{{"action",name},{"arguments",args}};};
    (void)linaccess::storeNativeMacro(db,"CommandDisposable",{{"version",1},{"actions",QJsonArray{macroAction("StopMacro")}}});
    const QJsonArray actions{macroAction("Rename",{"CommandAfter",0,"CommandBefore"}),macroAction("OpenTable",{"CommandAfter",0,1}),macroAction("Close",{0,"CommandAfter",2}),macroAction("DeleteObject",{0,"CommandAfter"}),macroAction("Rename",{"CommandQueryAfter",1,"CommandQueryBefore"}),macroAction("OpenQuery",{"CommandQueryAfter",0,2}),macroAction("Close"),macroAction("DeleteObject",{1,"CommandQueryAfter"}),macroAction("DeleteObject",{4,"CommandDisposable"}),macroAction("OpenTable",{"CommandKeep",0,1}),macroAction("RunCode",{"MacroCommandProbe()"}),macroAction("Close",{0,"CommandKeep",2}),macroAction("OpenForm",{"CommandForm",0}),macroAction("Close",{2,"CommandForm",2}),macroAction("OpenReport",{"CommandReport",2}),macroAction("Close",{3,"CommandReport",2}),macroAction("RunSQL",{"INSERT INTO command_events VALUES(13)"})};
    (void)linaccess::storeNativeMacro(db,"CommandProject",{{"version",1},{"actions",actions}});::action(window,"Aktualisieren")->trigger();auto* explorer=window.findChild<QTreeWidget*>("database-explorer");QTreeWidgetItem* item=nullptr;for(auto* candidate:explorer->findItems("CommandProject",Qt::MatchExactly|Qt::MatchRecursive,0)) if(candidate->data(0,Qt::UserRole).toString()=="macro") item=candidate;assert(item);explorer->setCurrentItem(item);window.focusDatabaseWindow();
    const auto generation=window.property("la_project_generation");bool protectedClose=false,edited=false;QTimer probe;probe.setInterval(1);QObject::connect(&probe,&QTimer::timeout,&window,[&] {auto* message=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());if(!message) return;assert(message->text()=="Macro command probe");assert(linaccess::nativeProjectMacrosActive(db));auto* page=window.findChild<QTabWidget*>()->currentWidget();assert(page && page->property("la_project_object_name").toString()=="CommandKeep");auto* view=page->findChild<QTableView*>();auto* model=qobject_cast<linaccess::AccessTableModel*>(view->model());assert(model && model->setData(model->index(0,model->fieldIndex("n")),43));edited=true;protectedClose=!window.close();assert(protectedClose && db.isOpen() && window.property("la_project_generation")==generation);message->button(QMessageBox::Ok)->click();});probe.start();auto* run=window.findChild<QAction*>("native-macro-run");assert(run && run->isEnabled());run->trigger();probe.stop();QApplication::processEvents();assert(protectedClose && edited && window.isVisible() && window.property("la_project_generation")==generation && window.statusBar()->currentMessage()=="Makro ausgeführt: CommandProject");
    assert(sql.exec("SELECT n FROM CommandKeep WHERE id=1") && sql.next() && sql.value(0).toInt()==43);sql.finish();assert(sql.exec("SELECT COUNT(*) FROM sqlite_schema WHERE type='table' AND name IN ('CommandBefore','CommandAfter')") && sql.next() && sql.value(0).toInt()==0);sql.finish();assert(sql.exec("SELECT COUNT(*) FROM _la_objects WHERE kind='query' AND name IN ('CommandQueryBefore','CommandQueryAfter')") && sql.next() && sql.value(0).toInt()==0);sql.finish();assert(!linaccess::readNativeMacroRevision(db,"CommandDisposable") && !linaccess::nativeProjectMacrosActive(db));
    assert(sql.exec("SELECT n FROM command_events"));QVector<int> events;while(sql.next()) events.push_back(sql.value(0).toInt());assert((events==QVector<int>{10,11,12,13}));sql.finish();for(const auto& document:AccessFormPreview::projectDocuments(db)) assert(document->documentName()!="CommandForm" && document->documentName()!="CommandReport");
    auto failing=formDefinition;failing.insert("vba","Private Sub Form_Load()\nErr.Raise 7003,\"CommandForm\",\"initialization marker\"\nEnd Sub");store("form","CommandHistorical",failing);store("form","CommandNewError",failing);const auto existing=linaccess::nativeStoredForm("CommandHistorical",failing,db);QPointer<AccessFormPreview> document=AccessFormPreview::showDocument(linaccess::AccessDocumentKind::Form,existing.name,existing.layout,existing.code,db,&window);assert(document && document->eventResult("OnLoad")->errorNumber==7003);
    linaccess::NativeStandardModuleHost host(&window,db,[] {});auto commands=host.applicationObject()->member("DoCmd",{}).object;commands->member("OpenForm",{VbaValue::string("CommandHistorical"),VbaValue::integer(0)});assert(document && document->eventResult("OnLoad")->errorNumber==7003);commands->member("Close",{VbaValue::integer(2),VbaValue::string("CommandHistorical"),VbaValue::integer(2)});int error=0;try {(void)commands->member("OpenForm",{VbaValue::string("CommandNewError"),VbaValue::integer(0)});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==7003);for(const auto& opened:AccessFormPreview::projectDocuments(db)) assert(opened->documentName()!="CommandHistorical" && opened->documentName()!="CommandNewError");assert(sql.exec("DELETE FROM _la_objects WHERE kind='module' AND name='CommandCode'"));
}

void testNativeMacroStorage(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;QSqlQuery sql(db);assert(sql.exec("CREATE TABLE managed_macro_effects(n INTEGER)"));
    const auto program=[](int value) {return QJsonObject{{"version",1},{"actions",QJsonArray{QJsonObject{{"action","RunSQL"},{"arguments",QJsonArray{QString("INSERT INTO managed_macro_effects VALUES(%1)").arg(value)}}}}}};};
    const auto archive=[&] {QSqlQuery query(db);assert(query.exec("SELECT id,name,source_bytes FROM _la_access_artifacts ORDER BY id"));QCryptographicHash hash(QCryptographicHash::Sha256);while(query.next()) {const auto name=query.value(1).toString().toUtf8(),bytes=query.value(2).toByteArray();hash.addData(QByteArray::number(query.value(0).toLongLong())+":"+QByteArray::number(name.size())+":"+name+":"+QByteArray::number(bytes.size())+":"+bytes);}return hash.result();};
    artifact(db,"macro","ArchivedManaged","archive_only_macro");const auto originalArchive=archive();
    auto external=QSqlDatabase::addDatabase("QSQLITE","macro-editor-concurrent");external.setDatabaseName(db.databaseName());assert(external.open());
    const auto revise=[&](int value) {QSqlQuery update(external);update.prepare("UPDATE _la_objects SET definition=? WHERE kind='macro' AND name='ManagedMacro'");update.addBindValue(QJsonDocument(program(value)).toJson(QJsonDocument::Compact));assert(update.exec() && update.numRowsAffected()==1);};
    QTimer::singleShot(0,&window,[&] {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-macro-designer");auto* name=dialog->findChild<QLineEdit*>("native-macro-name");auto* list=dialog->findChild<QTableWidget*>("native-macro-actions");name->setText("ManagedMacro");list->item(0,1)->setText("INSERT INTO managed_macro_effects VALUES(1)");QApplication::setActiveWindow(dialog);dialog->activateWindow();list->setFocus();QApplication::processEvents();QTest::keyClick(list,Qt::Key_S,Qt::ControlModifier);QApplication::processEvents();assert(name->isReadOnly() && dialog->findChild<QPushButton*>("native-macro-delete")->isEnabled());
        revise(9);list->item(0,1)->setText("INSERT INTO managed_macro_effects VALUES(2)");QTimer::singleShot(0,dialog,[] {auto* message=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(message && message->text().contains("zwischenzeitlich"));message->accept();});dialog->findChild<QPushButton*>("native-macro-save")->click();assert(dialog->isVisible() && dialog->isWindowModified());dialog->reject();});action(window,"+ Makro")->trigger();
    auto revision=linaccess::readNativeMacroRevision(db,"ManagedMacro");assert(revision && QJsonDocument::fromJson(revision->definition.toByteArray()).object()==program(9));
    const auto collision=[&](const QString& target) {QTimer::singleShot(0,&window,[&] {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog);auto* name=dialog->findChild<QLineEdit*>("native-macro-name");name->setText(target);dialog->findChild<QTableWidget*>("native-macro-actions")->item(0,1)->setText("INSERT INTO managed_macro_effects VALUES(99)");QTimer::singleShot(0,dialog,[] {auto* message=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(message && message->text().contains("bereits vorhanden"));message->accept();});dialog->findChild<QPushButton*>("native-macro-save")->click();assert(dialog->isVisible() && !name->isReadOnly() && !dialog->findChild<QPushButton*>("native-macro-delete")->isEnabled());dialog->reject();});action(window,"+ Makro")->trigger();};collision("managedmacro");collision("archivedmanaged");assert(archive()==originalArchive);
    auto* explorer=window.findChild<QTreeWidget*>("database-explorer");const auto macroItem=[&](const QString& name) {for(auto* item:explorer->findItems(name,Qt::MatchExactly|Qt::MatchRecursive,0)) if(item->data(0,Qt::UserRole).toString()=="macro") return item;return static_cast<QTreeWidgetItem*>(nullptr);};
    auto* item=macroItem("ManagedMacro");assert(item);QTimer::singleShot(0,&window,[&] {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->findChild<QLineEdit*>("native-macro-name")->isReadOnly());auto* remove=dialog->findChild<QPushButton*>("native-macro-delete");QTimer::singleShot(0,dialog,[] {auto* message=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(message);message->button(QMessageBox::No)->click();});remove->click();assert(dialog->isVisible());
        QTimer::singleShot(0,dialog,[&] {auto* question=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(question);revise(10);QTimer::singleShot(0,dialog,[] {auto* message=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(message && message->text().contains("zwischenzeitlich"));message->accept();});question->button(QMessageBox::Yes)->click();});remove->click();assert(dialog->isVisible());dialog->reject();});QMetaObject::invokeMethod(explorer,"itemDoubleClicked",Qt::DirectConnection,Q_ARG(QTreeWidgetItem*,item),Q_ARG(int,0));revision=linaccess::readNativeMacroRevision(db,"ManagedMacro");assert(revision && QJsonDocument::fromJson(revision->definition.toByteArray()).object()==program(10));
    int error=0;assert(sql.exec("CREATE TEMP TABLE _la_objects(kind,name,definition)"));try {(void)linaccess::storeNativeMacro(db,"ShadowedMacro",program(1));}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==3251);assert(sql.exec("DROP TABLE temp._la_objects"));assert(!linaccess::readNativeMacroRevision(db,"ShadowedMacro"));
    assert(sql.exec("CREATE TRIGGER managed_catalog_guard AFTER DELETE ON _la_objects BEGIN SELECT 1; END"));error=0;try {linaccess::deleteNativeMacro(db,*revision);}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==3251);assert(sql.exec("DROP TRIGGER managed_catalog_guard"));assert(linaccess::readNativeMacroRevision(db,"ManagedMacro"));
    action(window,"Aktualisieren")->trigger();item=macroItem("ManagedMacro");assert(item);explorer->setCurrentItem(item);auto* remove=window.findChild<QAction*>("native-object-delete");assert(remove);QMetaObject::invokeMethod(remove->parent(),"aboutToShow",Qt::DirectConnection);assert(remove->isEnabled());QTimer::singleShot(0,&window,[] {auto* question=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(question);question->button(QMessageBox::No)->click();});remove->trigger();assert(linaccess::readNativeMacroRevision(db,"ManagedMacro"));QTimer::singleShot(0,&window,[] {auto* question=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(question);question->button(QMessageBox::Yes)->click();});remove->trigger();assert(!linaccess::readNativeMacroRevision(db,"ManagedMacro") && archive()==originalArchive);
    assert(sql.exec("INSERT INTO _la_objects VALUES('macro','InvalidManaged','not-json')"));action(window,"Aktualisieren")->trigger();item=macroItem("InvalidManaged");assert(item);explorer->setCurrentItem(item);QMetaObject::invokeMethod(remove->parent(),"aboutToShow",Qt::DirectConnection);assert(remove->isEnabled());QTimer::singleShot(0,&window,[] {auto* question=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(question);question->button(QMessageBox::Yes)->click();});remove->trigger();assert(!linaccess::readNativeMacroRevision(db,"InvalidManaged"));
    sql.prepare("INSERT INTO _la_objects VALUES('module','ManagedDeletionCode',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba",R"VBA(Public Function DeleteRunningMacro() As Long
Dim number As Long
On Error Resume Next
Application.DoCmd.DeleteObject 4,"ManagedRunning"
number=Err.Number
On Error GoTo 0
CurrentDb.Execute "INSERT INTO managed_macro_effects VALUES(" & number & ")"
DeleteRunningMacro=number
End Function)VBA"}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();
    const auto running=QJsonObject{{"version",1},{"actions",QJsonArray{QJsonObject{{"action","RunCode"},{"arguments",QJsonArray{"DeleteRunningMacro()"}}},QJsonObject{{"action","RunSQL"},{"arguments",QJsonArray{"INSERT INTO managed_macro_effects VALUES(3)"}}}}}};(void)linaccess::storeNativeMacro(db,"ManagedRunning",running);
    linaccess::NativeStandardModuleHost host(&window,db,[] {});auto commands=host.applicationObject()->member("DoCmd",{}).object;commands->member("RunMacro",{VbaValue::string("ManagedRunning")});assert(linaccess::readNativeMacroRevision(db,"ManagedRunning") && !linaccess::nativeMacroObjectActive(db,"ManagedRunning"));
    {AccessFormPreview context("managed_macro_context","Begin Form\nBegin Section\nEnd\nEnd",{},db,&window);makePreviewApplicationObject(context)->member("DoCmd",{}).object->member("RunMacro",{VbaValue::string("ManagedRunning")});assert(context.close());}
    assert(sql.exec("SELECT n FROM managed_macro_effects"));QVector<int> effects;while(sql.next()) effects.push_back(sql.value(0).toInt());assert((effects==QVector<int>{3251,3,3251,3}));sql.finish();commands->member("DeleteObject",{VbaValue::integer(4),VbaValue::string("ManagedRunning")});assert(!linaccess::readNativeMacroRevision(db,"ManagedRunning"));error=0;try {(void)commands->member("DeleteObject",{VbaValue::integer(4),VbaValue::string("ArchivedManaged")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==3251 && archive()==originalArchive);
    assert(sql.exec("DELETE FROM _la_objects WHERE kind='module' AND name='ManagedDeletionCode'"));external.close();external=QSqlDatabase{};QSqlDatabase::removeDatabase("macro-editor-concurrent");
}

void testNativeMacroGroups(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;QSqlQuery sql(db);assert(sql.exec("CREATE TABLE group_effects(id INTEGER PRIMARY KEY,n INTEGER)"));
    const auto execute=[&](QString name,int repeat=1) {linaccess::NativeStandardModuleHost host(&window,db,[] {});(void)host.applicationObject()->member("DoCmd",{}).object->member("RunMacro",{VbaValue::string(name.toStdString()),VbaValue::longInteger(repeat)});};
    const auto saveShortcut=[&](QDialog* dialog,QTableWidget* list) {QApplication::setActiveWindow(dialog);dialog->activateWindow();list->setFocus();QApplication::processEvents();QTest::keyClick(list,Qt::Key_S,Qt::ControlModifier);QApplication::processEvents();assert(dialog->isVisible() && !dialog->isWindowModified());};
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-macro-designer");dialog->findChild<QLineEdit*>("native-macro-name")->setText("Grouped");auto* mode=dialog->findChild<QCheckBox*>("native-macro-group-mode");auto* selector=dialog->findChild<QComboBox*>("native-macro-group");auto* label=dialog->findChild<QLineEdit*>("native-macro-group-name");auto* list=dialog->findChild<QTableWidget*>("native-macro-actions");assert(mode && selector && label && list);mode->setChecked(true);label->setText("z_First");list->item(0,1)->setText("INSERT INTO group_effects(n) VALUES(1)");
        dialog->findChild<QPushButton*>("native-macro-group-add")->click();label->setText(QString::fromUtf8("Änderung"));list->item(0,1)->setText("INSERT INTO group_effects(n) VALUES(2)");dialog->findChild<QPushButton*>("native-macro-add")->click();qobject_cast<QComboBox*>(list->cellWidget(1,0))->setCurrentText("StopMacro");dialog->findChild<QPushButton*>("native-macro-add")->click();list->item(2,1)->setText("INSERT INTO group_effects(n) VALUES(999)");
        dialog->findChild<QPushButton*>("native-macro-group-add")->click();label->setText("a_Last");list->item(0,1)->setText("INSERT INTO group_effects(n) VALUES(3)");saveShortcut(dialog,list);
        selector->setCurrentIndex(0);assert(!dialog->isWindowModified() && list->rowCount()==1 && list->item(0,1)->text().endsWith("VALUES(1)"));selector->setCurrentIndex(1);assert(!dialog->isWindowModified() && list->rowCount()==3 && qobject_cast<QComboBox*>(list->cellWidget(1,0))->currentText()=="StopMacro");selector->setCurrentIndex(2);assert(!dialog->isWindowModified() && list->item(0,1)->text().endsWith("VALUES(3)"));
        dialog->findChild<QPushButton*>("native-macro-group-up")->click();assert(selector->currentIndex()==1);dialog->findChild<QPushButton*>("native-macro-group-down")->click();assert(selector->currentIndex()==2);saveShortcut(dialog,list);
        QTimer::singleShot(0,dialog,[] {auto* message=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(message);message->accept();});mode->setChecked(false);assert(mode->isChecked() && selector->count()==3 && !dialog->isWindowModified());dialog->reject();
    });action(window,"+ Makro")->trigger();
    assert(sql.exec("SELECT definition FROM _la_objects WHERE kind='macro' AND name='Grouped'") && sql.next());const auto source=sql.value(0).toByteArray();const auto definition=QJsonDocument::fromJson(source).object();assert(definition.value("version")==QJsonValue(2) && definition.value("groups").toArray()[0].toObject().value("name").toString()=="z_First");sql.finish();
    execute("grouped");execute(QString::fromUtf8("GROUPED.änderung"));execute("Grouped.a_Last");execute(QString::fromUtf8("Grouped.Änderung"),2);
    const auto macroAction=[](QString name,QJsonArray arguments) {return QJsonObject{{"action",name},{"arguments",arguments}};};
    const auto store=[&](QString name,QJsonObject program) {QSqlQuery save(db);save.prepare("INSERT INTO _la_objects VALUES('macro',?,?)");save.addBindValue(name);save.addBindValue(QJsonDocument(program).toJson(QJsonDocument::Compact));assert(save.exec());};
    store("GroupParent",{{"version",1},{"actions",QJsonArray{macroAction("RunMacro",{QString::fromUtf8("Grouped.Änderung")}),macroAction("RunSQL",{"INSERT INTO group_effects(n) VALUES(4)"})}}});execute("GroupParent");
    assert(sql.exec("SELECT n FROM group_effects ORDER BY id"));QVector<int> effects;while(sql.next()) effects.push_back(sql.value(0).toInt());assert((effects==QVector<int>{1,2,3,2,2,4}));sql.finish();
    const auto stored=linaccess::nativeStoredForm("group_event",{{"table","group_effects"},{"fields",QJsonArray{"n"}},{"events",QJsonObject{{"OnLoad",QString::fromUtf8("Grouped.änderung")}}}},db);AccessFormPreview event(stored.name,stored.layout,stored.code,db,&window);event.show();QApplication::processEvents();assert(event.eventResult("OnLoad")->status==linaccess::VbaExecutionStatus::Completed);assert(sql.exec("SELECT COUNT(*) FROM group_effects") && sql.next() && sql.value(0).toInt()==7);sql.finish();auto macros=makePreviewApplicationObject(event)->member("CurrentProject",{}).object->member("AllMacros",{}).object;assert(macros->member("Item",{VbaValue::string("Grouped")}).object->member("Name",{}).asText()=="Grouped");assert(event.close());
    const auto expect=[&](QString name,int wanted) {int actual=0;try {execute(name);}catch(const linaccess::VbaRuntimeError& error) {actual=error.number;}if(actual!=wanted) qFatal("Macro group expected %d, received %d",wanted,actual);};expect("Grouped.Missing",2485);expect("Grouped.a_Last.extra",3251);
    store("GroupInvalid",{{"version",2},{"groups",QJsonArray{QJsonObject{{"name","First"},{"actions",QJsonArray{macroAction("RunSQL",{"INSERT INTO group_effects(n) VALUES(8)"})}}},QJsonObject{{"name","Bad"},{"actions",QJsonArray{macroAction("UnsupportedGroupAction",{})}}}}}});expect("GroupInvalid.First",3251);
    store("GroupDuplicate",{{"version",2},{"groups",QJsonArray{QJsonObject{{"name",QString::fromUtf8("Änderung")},{"actions",QJsonArray{macroAction("RunSQL",{"INSERT INTO group_effects(n) VALUES(8)"})}}},QJsonObject{{"name",QString::fromUtf8("änderung")},{"actions",QJsonArray{macroAction("StopMacro",{})}}}}}});expect("GroupDuplicate",3251);assert(sql.exec("SELECT COUNT(*) FROM group_effects") && sql.next() && sql.value(0).toInt()==7);sql.finish();
    auto* explorer=window.findChild<QTreeWidget*>("database-explorer");QTreeWidgetItem* item=nullptr;for(auto* candidate:explorer->findItems("Grouped",Qt::MatchExactly|Qt::MatchRecursive,0)) if(candidate->data(0,Qt::UserRole).toString()=="macro") item=candidate;assert(item);
    QTimer::singleShot(0,&window,[&] {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-macro-designer");auto* selector=dialog->findChild<QComboBox*>("native-macro-group");auto* list=dialog->findChild<QTableWidget*>("native-macro-actions");assert(selector->count()==3 && !dialog->isWindowModified());selector->setCurrentIndex(2);dialog->findChild<QPushButton*>("native-macro-group-up")->click();dialog->findChild<QPushButton*>("native-macro-group-up")->click();saveShortcut(dialog,list);assert(selector->currentIndex()==0);dialog->reject();});QMetaObject::invokeMethod(explorer,"itemDoubleClicked",Qt::DirectConnection,Q_ARG(QTreeWidgetItem*,item),Q_ARG(int,0));execute("Grouped");assert(sql.exec("SELECT n FROM group_effects ORDER BY id DESC LIMIT 1") && sql.next() && sql.value(0).toInt()==3);sql.finish();
}

void testMacroRepeatExpressions(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;QSqlQuery sql(db);assert(sql.exec("CREATE TABLE repeat_effects(n INTEGER)"));assert(sql.exec("CREATE TABLE repeat_checks(n INTEGER)"));
    const auto store=[&](QString kind,QString name,QJsonObject definition) {QSqlQuery save(db);save.prepare("INSERT INTO _la_objects(kind,name,definition) VALUES(?,?,?)");save.addBindValue(kind);save.addBindValue(name);save.addBindValue(QJsonDocument(definition).toJson(QJsonDocument::Compact));assert(save.exec());};
    const auto action=[](QString name,QJsonArray args=QJsonArray{}) {return QJsonObject{{"action",name},{"arguments",args}};};
    const auto macro=[&](QString name,QJsonArray actions) {store("macro",name,{{"version",1},{"actions",actions}});};
    store("module","RepeatPredicates",{{"version",1},{"vba",R"VBA(Public Function RepeatPredicate() As Boolean
CurrentDb.Execute "INSERT INTO repeat_checks VALUES(1)"
RepeatPredicate=(DCount("*","repeat_effects")<3)
End Function
Public Function RepeatFailure() As Boolean
Err.Raise 7002,"RepeatPredicate","deliberate repeat failure"
End Function)VBA"}});
    macro("RepeatBody",{action("RunSQL",{"INSERT INTO repeat_effects VALUES(1)"})});
    const auto generation=window.property("la_project_generation");linaccess::NativeStandardModuleHost host(&window,db,[&] {if(window.property("la_project_generation")!=generation) throw linaccess::VbaRuntimeError(3420,"Project changed");});auto commands=host.applicationObject()->member("DoCmd",{}).object;
    const auto count=[&](QString table) {QSqlQuery query(db);assert(query.exec("SELECT COUNT(*) FROM "+table) && query.next());return query.value(0).toInt();};
    commands->member("RunMacro",{VbaValue::string("RepeatBody"),VbaValue::missing(),VbaValue::logical(false)});assert(count("repeat_effects")==0);
    commands->member("RunMacro",{VbaValue::string("RepeatBody"),VbaValue::longInteger(2),VbaValue::logical(true)});assert(count("repeat_effects")==2);assert(sql.exec("DELETE FROM repeat_effects"));
    commands->member("RunMacro",{VbaValue::string("RepeatBody"),VbaValue::missing(),VbaValue::string("RepeatPredicate()")});assert(count("repeat_effects")==3 && count("repeat_checks")==4);
    commands->member("RunMacro",{VbaValue::string("RepeatBody"),VbaValue::longInteger(2),VbaValue::string("DCount(\"*\",\"repeat_effects\")<10")});assert(count("repeat_effects")==5);
    int error=0;try {(void)commands->member("RunMacro",{VbaValue::string("RepeatBody"),VbaValue::missing(),VbaValue::string("RepeatFailure()")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;assert(failure.source=="RepeatPredicate" && std::string(failure.what()).find("Wiederholung 1")!=std::string::npos);}assert(error==7002 && count("repeat_effects")==5);
    macro("RepeatInvalid",{action("RunSQL",{"INSERT INTO repeat_effects VALUES(999)"}),action("RunMacro",{"RepeatBody",QJsonValue(QJsonValue::Null),"("})});error=0;try {(void)commands->member("RunMacro",{VbaValue::string("RepeatInvalid")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==3251 && count("repeat_effects")==5);
    macro("RepeatChildStop",{action("StopMacro")});macro("RepeatUnbounded",{action("RunMacro",{"RepeatChildStop"})});error=0;try {(void)commands->member("RunMacro",{VbaValue::string("RepeatUnbounded"),VbaValue::missing(),VbaValue::logical(true)});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;assert(std::string(failure.what()).find("iteration budget")!=std::string::npos);}assert(error==3251 && count("repeat_effects")==5);
    assert(sql.exec("DELETE FROM repeat_effects") && sql.exec("DELETE FROM repeat_checks"));
    QTimer::singleShot(0,&window,[&] {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-macro-designer");dialog->findChild<QLineEdit*>("native-macro-name")->setText("RepeatUi");auto* list=dialog->findChild<QTableWidget*>("native-macro-actions");qobject_cast<QComboBox*>(list->cellWidget(0,0))->setCurrentText("RunMacro");list->item(0,1)->setText("RepeatBody");list->item(0,2)->setText("");list->item(0,3)->setText("RepeatPredicate()");QApplication::setActiveWindow(dialog);dialog->activateWindow();list->setFocus();QApplication::processEvents();QTest::keyClick(list,Qt::Key_S,Qt::ControlModifier);QApplication::processEvents();assert(dialog->isVisible() && !dialog->isWindowModified());dialog->reject();});
    ::action(window,"+ Makro")->trigger();
    assert(sql.exec("SELECT definition FROM _la_objects WHERE kind='macro' AND name='RepeatUi'") && sql.next());const auto arguments=QJsonDocument::fromJson(sql.value(0).toByteArray()).object().value("actions").toArray()[0].toObject().value("arguments").toArray();assert(arguments.size()==3 && arguments[1].isNull() && arguments[2].toString()=="RepeatPredicate()");sql.finish();
    auto* explorer=window.findChild<QTreeWidget*>("database-explorer");QTreeWidgetItem* item=nullptr;for(auto* candidate:explorer->findItems("RepeatUi",Qt::MatchExactly|Qt::MatchRecursive,0)) if(candidate->data(0,Qt::UserRole).toString()=="macro") item=candidate;assert(item);
    QTimer::singleShot(0,&window,[&] {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-macro-designer");auto* list=dialog->findChild<QTableWidget*>("native-macro-actions");assert(list->item(0,2)->text().isEmpty() && list->item(0,3)->text()=="RepeatPredicate()");dialog->reject();});QMetaObject::invokeMethod(explorer,"itemDoubleClicked",Qt::DirectConnection,Q_ARG(QTreeWidgetItem*,item),Q_ARG(int,0));
    commands->member("RunMacro",{VbaValue::string("RepeatUi")});assert(count("repeat_effects")==3 && count("repeat_checks")==4);
    assert(sql.exec("DELETE FROM _la_objects WHERE kind='module' AND name='RepeatPredicates'"));
}

void testNativeAutoExec(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;const auto original=db.databaseName();assert(original!=":memory:");QTemporaryDir temporary;assert(temporary.isValid());const auto path=temporary.filePath("startup.linaccess");
    {
        auto setup=QSqlDatabase::addDatabase("QSQLITE","native-autoexec-setup");setup.setDatabaseName(path);assert(setup.open());
        {QSqlQuery sql(setup);assert(sql.exec("CREATE TABLE _la_objects(kind TEXT NOT NULL,name TEXT NOT NULL,definition TEXT NOT NULL,PRIMARY KEY(kind,name))"));assert(sql.exec("CREATE TABLE startup_effects(id INTEGER PRIMARY KEY,n INTEGER)"));
        const auto store=[&](QString kind,QString name,QJsonObject definition) {sql.prepare("INSERT INTO _la_objects VALUES(?,?,?)");sql.addBindValue(kind);sql.addBindValue(name);sql.addBindValue(QJsonDocument(definition).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();};
        store("module","StartupCode",{{"version",1},{"vba",R"VBA(Public Function StartupFunction() As Long
CurrentDb.Execute "INSERT INTO startup_effects(n) VALUES(2)"
StartupFunction=2
End Function)VBA"}});
        store("macro","aUtOeXeC",{{"version",1},{"actions",QJsonArray{QJsonObject{{"action","RunSQL"},{"arguments",QJsonArray{"INSERT INTO startup_effects(n) VALUES(1)"}}},QJsonObject{{"action","RunCode"},{"arguments",QJsonArray{"StartupFunction()"}}},QJsonObject{{"action","OpenTable"},{"arguments",QJsonArray{"startup_effects",0,1}}}}}});}
        setup.close();
    }QSqlDatabase::removeDatabase("native-autoexec-setup");
    window.setProperty("la_native_module_applications",QVariant{});window.setProperty("la_native_file_dialog",QVariant{});linaccess::NativeStandardModuleHost host(&window,db,[] {});auto heldApplication=host.applicationObject();auto heldProject=heldApplication->member("CurrentProject",{}).object;auto heldForms=heldApplication->member("Forms",{}).object;auto heldScreen=heldApplication->member("Screen",{}).object;auto heldPicker=heldApplication->member("FileDialog",{VbaValue::integer(3)}).object;auto heldTables=heldApplication->member("CurrentData",{}).object->member("AllTables",{}).object;
    openProject(window,path);assert(window.property("la_startup_macro_status").toString()=="completed");
    const auto count=[&] {QSqlQuery query(db);assert(query.exec("SELECT COUNT(*) FROM startup_effects") && query.next());return query.value(0).toInt();};assert(count()==2);
    int stale=0;try {(void)heldApplication->member("Eval",{VbaValue::string("1")});}catch(const linaccess::VbaRuntimeError& error) {stale=error.number;}assert(stale==3420);stale=0;try {(void)heldProject->member("AllModules",{});}catch(const linaccess::VbaRuntimeError& error) {stale=error.number;}assert(stale==3420);stale=0;try {(void)heldTables->member("Count",{});}catch(const linaccess::VbaRuntimeError& error) {stale=error.number;}assert(stale==3420);
    stale=0;try {(void)heldForms->member("Count",{});}catch(const linaccess::VbaRuntimeError& error) {stale=error.number;}assert(stale==3420);stale=0;try {(void)heldScreen->member("MousePointer",{});}catch(const linaccess::VbaRuntimeError& error) {stale=error.number;}assert(stale==3420);
    stale=0;try {(void)heldPicker->member("Title",{});}catch(const linaccess::VbaRuntimeError& error) {stale=error.number;}assert(stale==3420);{linaccess::NativeStandardModuleHost current(&window,db,[] {});auto fresh=current.applicationObject()->member("FileDialog",{VbaValue::integer(3)}).object;assert(fresh!=heldPicker && fresh->member("DialogType",{}).number==3);}
    openProject(window,path);assert(window.property("la_startup_macro_status").toString()=="completed" && count()==4);
    openProject(window,path,nullptr,true);QTest::keyRelease(&window,Qt::Key_Shift);assert(window.property("la_startup_macro_status").toString()=="bypassed" && count()==4);
    {QSqlQuery sql(db);assert(sql.exec("DELETE FROM _la_objects WHERE kind='macro'"));}
    const auto source=QByteArray("archived_macro_source_only");const int archived=artifact(db,"macro","AutoExec",source);openProject(window,path);assert(window.property("la_startup_macro_status").toString()=="none" && count()==4);
    {QSqlQuery sql(db);sql.prepare("SELECT source_bytes FROM _la_access_artifacts WHERE id=?");sql.addBindValue(archived);assert(sql.exec() && sql.next() && sql.value(0).toByteArray()==source);sql.finish();sql.prepare("INSERT INTO _la_objects VALUES('macro','AutoExec',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"actions",QJsonArray{QJsonObject{{"action","RunSQL"},{"arguments",QJsonArray{"INSERT INTO startup_effects(n) VALUES(7)"}}},QJsonObject{{"action","RunCode"},{"arguments",QJsonArray{"MissingStartupFunction()"}}}}}}).toJson(QJsonDocument::Compact));assert(sql.exec());}
    bool observedError=false;QTimer dismiss;dismiss.setInterval(1);QObject::connect(&dismiss,&QTimer::timeout,&window,[&] {if(auto* message=qobject_cast<QMessageBox*>(QApplication::activeModalWidget())) {assert(message->text().contains("AutoExec") && message->text().contains("2517"));observedError=true;message->accept();}});dismiss.start();openProject(window,path);dismiss.stop();assert(observedError && window.property("la_startup_macro_status").toString()=="failed" && db.isOpen() && count()==5);
    openProject(window,path,nullptr,true);QTest::keyRelease(&window,Qt::Key_Shift);assert(window.property("la_startup_macro_status").toString()=="bypassed" && count()==5);
    openProject(window,original);assert(db.databaseName()==original && window.property("la_startup_macro_status").toString()=="none");
}

void testStandaloneMacroExecution(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;QSqlQuery sql(db);assert(sql.exec("CREATE TABLE standalone_macro_effects(id INTEGER PRIMARY KEY,n INTEGER)"));
    const auto store=[&](const QString& kind,const QString& name,const QJsonObject& definition) {QSqlQuery save(db);save.prepare("INSERT INTO _la_objects(kind,name,definition) VALUES(?,?,?)");save.addBindValue(kind);save.addBindValue(name);save.addBindValue(QJsonDocument(definition).toJson(QJsonDocument::Compact));assert(save.exec());};
    const auto macroAction=[](QString name,QJsonArray arguments=QJsonArray{}) {return QJsonObject{{"action",name},{"arguments",arguments}};};
    const auto macro=[&](QString name,QJsonArray actions) {store("macro",name,{{"version",1},{"actions",actions}});};
    store("module","StandaloneCode",{{"version",1},{"vba",R"VBA(Public Function StandaloneMacroCode() As Long
CurrentDb.Execute "INSERT INTO standalone_macro_effects(n) VALUES(2)"
StandaloneMacroCode=2
End Function)VBA"}});
    macro("StandaloneChild",{macroAction("RunCode",{"StandaloneMacroCode()"}),macroAction("StopMacro"),macroAction("RunSQL",{"INSERT INTO standalone_macro_effects(n) VALUES(999)"})});
    macro("StandaloneLoad",{macroAction("RunSQL",{"INSERT INTO standalone_macro_effects(n) VALUES(4)"})});
    store("query","StandaloneQuery",{{"sql","SELECT id,n FROM standalone_macro_effects ORDER BY id"}});
    store("form","StandaloneForm",{{"table","standalone_macro_effects"},{"fields",QJsonArray{"n"}},{"events",QJsonObject{{"OnLoad","StandaloneLoad"}}}});
    store("report","StandaloneReport",{{"sql","SELECT n FROM standalone_macro_effects ORDER BY id"},{"title","Module macro output"}});
    auto first=macroAction("RunSQL",{"INSERT INTO standalone_macro_effects(n) VALUES(1)"});first.insert("condition","DCount(\"*\",\"standalone_macro_effects\")=0");
    auto skipped=macroAction("RunSQL",{"INSERT INTO standalone_macro_effects(n) VALUES(3)"});skipped.insert("condition","False");
    macro("StandaloneDispatch",{first,macroAction("RunMacro",{"StandaloneChild"}),skipped,macroAction("OpenTable",{"standalone_macro_effects",0,1}),macroAction("RunCommand",{97}),macroAction("Requery"),macroAction("OpenQuery",{"StandaloneQuery",0,2}),macroAction("Requery"),macroAction("OpenForm",{"StandaloneForm",0}),macroAction("Requery"),macroAction("OpenReport",{"StandaloneReport",2})});
    action(window,"Aktualisieren")->trigger();auto* explorer=window.findChild<QTreeWidget*>("database-explorer");QTreeWidgetItem* item=nullptr;for(auto* candidate:explorer->findItems("StandaloneDispatch",Qt::MatchExactly|Qt::MatchRecursive,0)) if(candidate->data(0,Qt::UserRole).toString()=="macro") item=candidate;assert(item);explorer->setCurrentItem(item);window.focusDatabaseWindow();
    auto* run=window.findChild<QAction*>("native-macro-run");assert(run && run->isEnabled());run->trigger();QApplication::processEvents();
    assert(window.statusBar()->currentMessage()=="Makro ausgeführt: StandaloneDispatch");
    assert(sql.exec("SELECT n FROM standalone_macro_effects ORDER BY id"));QVector<int> effects;while(sql.next()) effects.push_back(sql.value(0).toInt());assert((effects==QVector<int>{1,2,4}));sql.finish();
    QPointer<AccessFormPreview> openedForm,openedReport;for(const auto& document:AccessFormPreview::projectDocuments(db)) {if(document->documentName()=="StandaloneForm") openedForm=document;if(document->documentName()=="StandaloneReport") openedReport=document;}assert(openedForm && openedReport && openedForm->hasDataSource() && openedReport->hasDataSource());assert(makePreviewFormObject(*openedForm)->member("Recordset",{}).object->member("RecordCount",{}).number==3);
    assert(openedReport->close() && openedForm->close());QApplication::processEvents();window.focusDatabaseWindow();
    const auto generation=window.property("la_project_generation");linaccess::NativeStandardModuleHost host(&window,db,[&] {if(window.property("la_project_generation")!=generation) throw linaccess::VbaRuntimeError(3420,"Project changed");});auto commands=host.applicationObject()->member("DoCmd",{}).object;
    macro("StandaloneFailure",{macroAction("RunSQL",{"INSERT INTO standalone_macro_effects(n) VALUES(5)"}),macroAction("RunCode",{"MissingStandaloneFunction()"})});int error=0;try {(void)commands->member("RunMacro",{VbaValue::string("StandaloneFailure")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;assert(std::string(failure.what()).find("Aktion 2")!=std::string::npos);}assert(error==2517);
    assert(sql.exec("SELECT COUNT(*) FROM standalone_macro_effects WHERE n=5") && sql.next() && sql.value(0).toInt()==1);sql.finish();
    macro("StandaloneRecursive",{macroAction("RunMacro",{"StandaloneRecursive"})});error=0;try {(void)commands->member("RunMacro",{VbaValue::string("StandaloneRecursive")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==3251);
    macro("StandaloneAfterRecursion",{macroAction("RunSQL",{"INSERT INTO standalone_macro_effects(n) VALUES(6)",false})});commands->member("RunMacro",{VbaValue::string("StandaloneAfterRecursion")});
    error=0;try {(void)commands->member("RunSQL",{VbaValue::string("SELECT * FROM standalone_macro_effects")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==2342);
    macro("StandaloneCancel",{macroAction("CancelEvent")});error=0;try {(void)commands->member("RunMacro",{VbaValue::string("StandaloneCancel")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==3251);
    store("form","StandaloneCancelled",{{"table","standalone_macro_effects"},{"fields",QJsonArray{"n"}},{"events",QJsonObject{{"OnOpen","StandaloneCancel"}}}});error=0;try {(void)commands->member("OpenForm",{VbaValue::string("StandaloneCancelled")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==2501);
    for(const auto& document:AccessFormPreview::projectDocuments(db)) assert(document->documentName()!="StandaloneCancelled");
    assert(sql.exec("SELECT COUNT(*) FROM standalone_macro_effects WHERE n=6") && sql.next() && sql.value(0).toInt()==1);sql.finish();
    assert(sql.exec("DELETE FROM _la_objects WHERE kind='module' AND name='StandaloneCode'"));
}

void testNativeModuleEditor(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;
    const auto archive=[&] {QSqlQuery query(db);assert(query.exec("SELECT id,name,source_bytes FROM _la_access_artifacts ORDER BY id"));QCryptographicHash hash(QCryptographicHash::Sha256);while(query.next()) {const auto name=query.value(1).toString().toUtf8(),bytes=query.value(2).toByteArray();hash.addData(QByteArray::number(query.value(0).toLongLong())+":"+QByteArray::number(name.size())+":"+name+":"+QByteArray::number(bytes.size())+":"+bytes);}return hash.result();};const auto originalArchive=archive();
    const QByteArray body=R"VBA(Public Counter As Long
Private Function NativeAdjust(value As Long) As Long
NativeAdjust=value+3
End Function
Public Function NativeAdd(value As Long) As Long
Counter=Counter+1
NativeAdd=NativeAdjust(value)
End Function
Public Function NativeCounter() As Long
NativeCounter=Counter
End Function
Public Sub NativeSetCounter(ByVal value As Long)
Counter=value
End Sub
Public Function NativeOptional(Optional value As Long=9) As Long
NativeOptional=value
End Function
Public Function NativeText(ByVal value As String) As String
Counter=Counter+1
NativeText=value
End Function
Public Function NativeFailure() As Long
Err.Raise 7001, "NativeRun", "deliberate failure"
End Function
Public Sub NativeHaltProbe()
Dim n As Long
n=17
Stop
Debug.Assert n=18
n=n+1
CurrentDb.Execute "INSERT INTO editor_module_effects VALUES(" & CStr(n) & ")"
End Sub
Public Sub NativeExecute()
CurrentDb.Execute "INSERT INTO editor_module_effects VALUES(" & Application.Run("NativeSupport",NativeAdd(1)) & ")"
End Sub
Public Function NativeUseField(ByVal value As DAO.Field) As Long
NativeUseField=value.Value
End Function
Public Function NativeApplicationObjectArg() As Long
Dim rs As DAO.Recordset
Set rs=CurrentDb.OpenRecordset("SELECT n FROM editor_module_effects")
NativeApplicationObjectArg=Application.Run("NativeUseField",rs.Fields(0))
rs.Close
End Function
Public Function NativeApplicationEval() As Long
NativeApplicationEval=Application.Eval("NativeOptional(6)")+Application.Eval("DCount(""*"",""editor_module_effects"")")
End Function
Public Sub NativeKeepApplication()
SaveModuleApplication
End Sub
Public Sub NativeInvalidContext()
Me.Tag=1
End Sub
Public Function NativePrompt() As Long
NativePrompt=MsgBox("Module execution",vbYesNo,"Module")
End Function
Public Sub NativePreflight()
CurrentDb.Execute "INSERT INTO editor_module_effects VALUES(999)"
GoSub MissingLabel
End Sub
Public Function NativeRecursive() As Long
NativeRecursive=Application.Run("NativeRecursive")
End Function)VBA";
    AccessFormPreview context("owned_module_context","Begin Form\nBegin Section\nEnd\nEnd",R"VBA(Sub UseOwnedModule()
Me.Tag=NativeAdd(4)
End Sub
Sub UseDynamicModule()
Me.Tag=Application.Run(Procedure:="NativeOptional",Arg1:=12)
End Sub
Sub OpenOwnedEditor()
DoCmd.OpenModule ProcedureName:="NativeHaltProbe"
Me.Tag="module-call-returned"
End Sub
Sub SelectOwnedEditor()
DoCmd.SelectObject ObjectType:=acModule,ObjectName:="nativemath",InDatabaseWindow:=False
DoCmd.RepaintObject acModule,"NativeMath"
Me.Tag="module-selected"
End Sub)VBA",db,&window);context.show();QApplication::processEvents();
    auto application=makePreviewApplicationObject(context);auto modules=application->member("CurrentProject",{}).object->member("AllModules",{}).object;const auto baseline=modules->member("Count",{}).number;
    std::shared_ptr<linaccess::VbaObject> descriptor;
    QTimer::singleShot(0,&window,[&] {
        auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-module-editor");auto* name=dialog->findChild<QLineEdit*>("native-module-name");auto* editor=dialog->findChild<QPlainTextEdit*>("native-module-source");assert(name && editor);name->setText("NativeMath");editor->setPlainText(QString::fromUtf8(body));QApplication::setActiveWindow(dialog);dialog->activateWindow();editor->setFocus();QApplication::processEvents();QTest::keyClick(editor,Qt::Key_S,Qt::ControlModifier);QApplication::processEvents();
        assert(dialog->isVisible() && name->isReadOnly() && modules->member("Count",{}).number==baseline+1);descriptor=modules->member("Item",{VbaValue::string("nativemath")}).object;assert(descriptor->member("Name",{}).asText()=="NativeMath" && descriptor->member("IsLoaded",{}).truthy());assert(application->member("SysCmd",{VbaValue::integer(10),VbaValue::integer(5),VbaValue::string("NativeMath")}).number==1);dialog->reject();
    });action(window,"+ Modul")->trigger();assert(descriptor && !descriptor->member("IsLoaded",{}).truthy() && archive()==originalArchive);
    {
        std::vector<VbaValue> openArguments;const auto opened=context.executeProcedure("OpenOwnedEditor",openArguments);if(opened.status!=linaccess::VbaExecutionStatus::Completed) qFatal("OpenModule actual VBA error %d",opened.errorNumber);
        assert(opened.status==linaccess::VbaExecutionStatus::Completed && makePreviewFormObject(context)->member("Tag",{}).asText()=="module-call-returned");QApplication::processEvents();
        const auto editorFor=[&](const QString& name) -> QDialog* {QDialog* result=nullptr;for(auto* editor:window.findChildren<QDialog*>("native-module-editor")) if(editor->isVisible() && editor->property("la_project_object_name").toString()==name) {assert(!result);result=editor;}return result;};
        QPointer<QDialog> editor=editorFor("NativeMath");assert(editor && !editor->isModal() && !QApplication::activeModalWidget() && descriptor->member("IsLoaded",{}).truthy());
        auto* text=editor->findChild<QPlainTextEdit*>("native-module-source");assert(text && text->toPlainText()==QString::fromUtf8(body));
        int line=0;for(const auto& procedure:linaccess::analyzeVba(body.toStdString()).procedures) if(procedure.name=="NativeHaltProbe") line=procedure.line;assert(line>0 && text->textCursor().blockNumber()+1==line);
        editor->showMaximized();QApplication::processEvents();assert(editor->isMaximized());QApplication::setActiveWindow(&context);const auto selected=context.executeProcedure("SelectOwnedEditor",openArguments);assert(selected.status==linaccess::VbaExecutionStatus::Completed && editor->isMaximized() && QApplication::activeWindow()==editor && makePreviewFormObject(context)->member("Tag",{}).asText()=="module-selected");
        text->appendPlainText("' pending editor draft");assert(editor->isWindowModified());
        linaccess::NativeStandardModuleHost moduleHost(&window,db,[]{});auto moduleCommands=moduleHost.applicationObject()->member("DoCmd",{}).object;
        moduleCommands->member("OpenModule",{VbaValue::string("nativemath"),VbaValue::string("NativeCounter")});assert(editorFor("NativeMath")==editor && text->toPlainText().endsWith("' pending editor draft"));
        editor->showMinimized();QApplication::processEvents();assert(editor->isMinimized());QApplication::setActiveWindow(&context);moduleCommands->member("SelectObject",{VbaValue::integer(5),VbaValue::string("NativeMath"),VbaValue::logical(false)});assert(!editor->isMinimized() && editor->isMaximized() && QApplication::activeWindow()==editor);moduleCommands->member("RepaintObject",{});assert(text->toPlainText().endsWith("' pending editor draft") && editor->isWindowModified());

        QTimer::singleShot(0,&window,[] {auto* question=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(question);question->button(QMessageBox::Cancel)->click();});assert(!editor->close() && editor->isVisible() && editor->isWindowModified());
        QSqlQuery revision(db);assert(revision.exec("SELECT definition FROM main._la_objects WHERE kind='module' AND name='NativeMath'") && revision.next());const auto previous=revision.value(0);revision.finish();
        const auto external=QString::fromUtf8(QJsonDocument(QJsonObject{{"version",1},{"vba",QString::fromUtf8(body)+"\n' external revision"}}).toJson(QJsonDocument::Compact));
        revision.prepare("UPDATE main._la_objects SET definition=? WHERE kind='module' AND name='NativeMath'");revision.addBindValue(external);assert(revision.exec());revision.finish();
        QTimer::singleShot(0,&window,[] {auto* message=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(message);message->button(QMessageBox::Ok)->click();});editor->findChild<QPushButton*>("native-module-save")->click();
        assert(editor->isWindowModified() && revision.exec("SELECT definition FROM main._la_objects WHERE kind='module' AND name='NativeMath'") && revision.next() && revision.value(0).toString()==external);revision.finish();
        revision.prepare("UPDATE main._la_objects SET definition=? WHERE kind='module' AND name='NativeMath'");revision.addBindValue(previous);assert(revision.exec());revision.finish();
        moduleCommands->member("Close",{VbaValue::integer(5),VbaValue::string("NativeMath"),VbaValue::integer(1)});QApplication::processEvents();assert(!descriptor->member("IsLoaded",{}).truthy());
        assert(revision.exec("SELECT definition FROM main._la_objects WHERE kind='module' AND name='NativeMath'") && revision.next() && QJsonDocument::fromJson(revision.value(0).toByteArray()).object().value("vba").toString().endsWith("' pending editor draft"));revision.finish();
        revision.prepare("UPDATE main._la_objects SET definition=? WHERE kind='module' AND name='NativeMath'");revision.addBindValue(previous);assert(revision.exec());revision.finish();
        moduleCommands->member("OpenModule",{VbaValue::string("NativeMath")});auto* discard=editorFor("NativeMath");assert(discard);discard->findChild<QPlainTextEdit*>("native-module-source")->appendPlainText("' discard this draft");moduleCommands->member("Close",{VbaValue::integer(5),VbaValue::string("NativeMath"),VbaValue::integer(2)});QApplication::processEvents();assert(!descriptor->member("IsLoaded",{}).truthy());

        int closedSelection=0;try{moduleCommands->member("SelectObject",{VbaValue::integer(5),VbaValue::string("NativeMath"),VbaValue::logical(false)});}catch(const linaccess::VbaRuntimeError& error){closedSelection=error.number;}assert(closedSelection==3251 && !editorFor("NativeMath"));
        const auto failure=[&](std::vector<VbaValue> args,int expected) {int actual=0;try{moduleCommands->member("OpenModule",args);}catch(const linaccess::VbaRuntimeError& error){actual=error.number;}assert(actual==expected);};failure({},5);failure({VbaValue{}},94);failure({VbaValue::string("missing module")},3251);
        // A temporary technical import is viewed without modifying its archived source.
        const QByteArray imported="Public Sub ViewerOnly()\nErr.Raise 7099\nEnd Sub";const int id=artifact(db,"vba","ImportedModuleViewer",imported);const auto importedArchive=archive();
        moduleCommands->member("OpenModule",{VbaValue::string("ImportedModuleViewer"),VbaValue::string("ViewerOnly")});auto* viewer=editorFor("ImportedModuleViewer");assert(viewer && viewer->findChild<QPlainTextEdit*>("native-module-source")->isReadOnly() && !viewer->findChild<QPushButton*>("native-module-save")->isEnabled() && !viewer->findChild<QPushButton*>("native-module-run")->isEnabled());
        assert(archive()==importedArchive);moduleCommands->member("Close",{VbaValue::integer(5),VbaValue::string("ImportedModuleViewer"),VbaValue::integer(2)});QSqlQuery remove(db);remove.prepare("DELETE FROM main._la_access_artifacts WHERE id=?");remove.addBindValue(id);assert(remove.exec());remove.finish();assert(archive()==originalArchive);QApplication::processEvents();
    }
    std::vector<VbaValue> arguments;const auto initialOwned=context.executeProcedure("UseOwnedModule",arguments);
    if(initialOwned.status!=linaccess::VbaExecutionStatus::Completed) {
        std::cerr<<"Owned module status="<<static_cast<int>(initialOwned.status)<<" error="<<initialOwned.errorNumber<<" line="<<initialOwned.line<<"\n";
    }
    assert(initialOwned.status==linaccess::VbaExecutionStatus::Completed && makePreviewFormObject(context)->member("Tag",{}).asText()=="7");
    const auto first=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT NativeAdd(5) AS computed,NativeCounter() AS total"),{});assert(first.rows.size()==1 && first.rows[0][0].toInt()==8 && first.rows[0][1].toInt()==2);
    QSqlQuery sql(db);assert(!sql.exec("SELECT NativeAdjust(1)"));
    QSqlQuery prepared(db);assert(prepared.prepare("SELECT NativeAdd(5)"));
    const auto open=[&](const std::function<void(QDialog*)>& callback) {auto* explorer=window.findChild<QTreeWidget*>("database-explorer");QTreeWidgetItem* module=nullptr;for(int group=0;group<explorer->topLevelItemCount();++group) for(int row=0;row<explorer->topLevelItem(group)->childCount();++row) {auto* item=explorer->topLevelItem(group)->child(row);if(item->data(0,Qt::UserRole).toString()=="module" && item->text(0)=="NativeMath") module=item;}assert(module);QTimer::singleShot(0,&window,[&] {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-module-editor");callback(dialog);});QMetaObject::invokeMethod(explorer,"itemDoubleClicked",Qt::DirectConnection,Q_ARG(QTreeWidgetItem*,module),Q_ARG(int,0));};
    auto changed=body;changed.replace("value+3","value+10");
    open([&](QDialog* dialog) {auto* editor=dialog->findChild<QPlainTextEdit*>("native-module-source");assert(editor && editor->toPlainText()==QString::fromUtf8(body));editor->setPlainText(QString::fromUtf8(changed));dialog->findChild<QPushButton*>("native-module-save")->click();});
    assert(prepared.exec() && prepared.next() && prepared.value(0).toInt()==15);prepared.finish();assert(context.executeProcedure("UseOwnedModule",arguments).status==linaccess::VbaExecutionStatus::Completed && makePreviewFormObject(context)->member("Tag",{}).asText()=="14");
    assert(sql.exec("SELECT NativeCounter()") && sql.next() && sql.value(0).toInt()==2);sql.finish();
    assert(sql.exec("SAVEPOINT owned_module_revision"));auto temporary=changed;temporary.replace("value+10","value+20");sql.prepare("UPDATE _la_objects SET definition=? WHERE kind='module' AND name='NativeMath'");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba",QString::fromUtf8(temporary)}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();assert(prepared.exec() && prepared.next() && prepared.value(0).toInt()==25);prepared.finish();assert(sql.exec("ROLLBACK TO owned_module_revision") && sql.exec("RELEASE owned_module_revision"));assert(prepared.exec() && prepared.next() && prepared.value(0).toInt()==15);prepared.finish();
    assert(application->member("Run",{VbaValue::string("nativeadd"),VbaValue::longInteger(2)}).number==12);
    assert(application->member("Run",{VbaValue::string("NativeSetCounter"),VbaValue::longInteger(41)}).kind==VbaValue::Kind::Empty);
    assert(context.executeProcedure("UseDynamicModule",arguments).status==linaccess::VbaExecutionStatus::Completed && makePreviewFormObject(context)->member("Tag",{}).asText()=="12");
    assert(application->member("Run",{VbaValue::string("NativeCounter")}).number==41);
    assert(application->member("Run",{VbaValue::string("NativeOptional")}).number==9);
    assert(application->member("Run",{VbaValue::string("NativeOptional"),VbaValue::longInteger(6)}).number==6);
    const auto runError=[&](std::vector<VbaValue> args,int expected) {int actual=0;try {(void)application->member("Run",args);}catch(const linaccess::VbaRuntimeError& failure) {actual=failure.number;}if(actual!=expected) qFatal("Application.Run expected %d, received %d",expected,actual);};
    runError({VbaValue::string("NativeAdjust"),VbaValue::longInteger(1)},2517);
    runError({VbaValue::string("NoSuchNativeProcedure")},2517);
    runError({VbaValue::string("OtherProject.NativeAdd")},3251);
    runError({VbaValue::string("NativeAdd")},450);
    runError({VbaValue::string("NativeRecursive")},28);
    runError(std::vector<VbaValue>(32,VbaValue::longInteger(0)),450);
    int unsupportedRun=0;try {(void)application->member("DoCmd",{}).object->member("Run",{VbaValue::string("NativeCounter")});}catch(const linaccess::VbaRuntimeError& failure) {unsupportedRun=failure.number;}assert(unsupportedRun==438);
    try {(void)application->member("Run",{VbaValue::string("NativeFailure")});assert(false);}catch(const linaccess::VbaRuntimeError& failure) {assert(failure.number==7001 && failure.source=="NativeRun");}
    assert(application->member("Run",{VbaValue::string("NativeOptional")}).number==9);
    const auto storeRunCode=[&](QString name,QJsonArray actions) {QSqlQuery save(db);save.prepare("INSERT INTO _la_objects(kind,name,definition) VALUES('macro',?,?)");save.addBindValue(name);save.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"actions",actions}}).toJson(QJsonDocument::Compact));assert(save.exec());};
    const auto codeAction=[](QString expression) {return QJsonObject{{"action","RunCode"},{"arguments",QJsonArray{expression}}};};
    auto commands=application->member("DoCmd",{}).object;
    storeRunCode("native_run_code",{codeAction("NativeOptional(3)")});
    commands->member("RunMacro",{VbaValue::string("native_run_code")});
    storeRunCode("native_run_nested_code",{QJsonObject{{"action","RunMacro"},{"arguments",QJsonArray{"native_run_code"}}}});
    AccessFormPreview eventCode("run_code_event","Begin Form\nOnLoad =\"native_run_nested_code\"\nBegin Section\nEnd\nEnd",R"VBA(Private Function NativeOptional(value As Long) As Long
Me.Tag=value+70
NativeOptional=value
End Function)VBA",db,&window);eventCode.show();QApplication::processEvents();
    assert(makePreviewFormObject(eventCode)->member("Tag",{}).asText()=="73");
    assert(!eventCode.property("la_bound_event_active").toBool());
    commands->member("RunMacro",{VbaValue::string("native_run_code")});
    assert(makePreviewFormObject(context)->member("Tag",{}).asText()=="12");
    storeRunCode("native_run_text",{codeAction("NativeText(\"a)\"\"b(\")"),codeAction("NativeAdd(NativeOptional(2))")});
    commands->member("RunMacro",{VbaValue::string("native_run_text")});
    assert(application->member("Run",{VbaValue::string("NativeCounter")}).number==43);
    const auto macroFailure=[&](QString name,int expected) {int actual=0;try {(void)commands->member("RunMacro",{VbaValue::string(name.toStdString())});}catch(const linaccess::VbaRuntimeError& failure) {actual=failure.number;}if(actual!=expected) qFatal("RunCode expected %d, received %d",expected,actual);};
    storeRunCode("native_run_sub",{codeAction("NativeSetCounter(99)")});macroFailure("native_run_sub",3251);
    storeRunCode("native_run_private",{codeAction("NativeAdjust(2)")});macroFailure("native_run_private",2517);
    assert(sql.exec("CREATE TABLE run_code_effects(n INTEGER)"));
    storeRunCode("native_run_shape",{QJsonObject{{"action","RunSQL"},{"arguments",QJsonArray{"INSERT INTO run_code_effects VALUES(1)"}}},codeAction("NativeCounter()+NativeCounter()")});macroFailure("native_run_shape",3251);
    assert(sql.exec("SELECT COUNT(*) FROM run_code_effects") && sql.next() && sql.value(0).toInt()==0);sql.finish();
    storeRunCode("native_run_error",{codeAction("NativeFailure()")});
    try {(void)commands->member("RunMacro",{VbaValue::string("native_run_error")});assert(false);}catch(const linaccess::VbaRuntimeError& failure) {assert(failure.number==7001 && failure.source=="NativeRun" && std::string(failure.what()).find("RunCode")!=std::string::npos);}
    assert(application->member("Run",{VbaValue::string("NativeCounter")}).number==43);
    int directRunCode=0;try {(void)commands->member("RunCode",{VbaValue::string("NativeCounter()")});}catch(const linaccess::VbaRuntimeError& failure) {directRunCode=failure.number;}assert(directRunCode==438);
    assert(eventCode.close());
    assert(sql.exec("CREATE TEMP TABLE _la_objects(kind,name,definition)"));int error=0;try {(void)linaccess::standardVbaModuleSources(db);}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==3251 && !sql.exec("SELECT NativeAdd(5)"));assert(sql.exec("DROP TABLE temp._la_objects"));
    const int collision=artifact(db,"vba","NATIVEMATH",body);error=0;try {(void)linaccess::standardVbaModuleSources(db);}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==5 && !sql.exec("SELECT NativeAdd(5)"));sql.prepare("DELETE FROM _la_access_artifacts WHERE id=?");sql.addBindValue(collision);assert(sql.exec());sql.finish();assert(archive()==originalArchive);
    QTimer::singleShot(0,&window,[&] {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-module-editor");dialog->findChild<QLineEdit*>("native-module-name")->setText("MacroConditions");dialog->findChild<QPlainTextEdit*>("native-module-source")->setPlainText(QString::fromUtf8(body));QTimer::singleShot(0,&window,[&] {auto* message=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(message);message->accept();});dialog->findChild<QPushButton*>("native-module-save")->click();assert(dialog->isVisible());dialog->reject();});action(window,"+ Modul")->trigger();assert(archive()==originalArchive);
    error=0;try {(void)linaccess::nativeModuleSource(QJsonObject{{"version",1},{"vba","VERSION 1.0 CLASS\n"}});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==3251);
    assert(sql.exec("CREATE TABLE editor_module_effects(n INTEGER)"));
    sql.prepare("INSERT INTO _la_objects(kind,name,definition) VALUES('module','NativeSupportModule',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba","Public Function NativeSupport(value As Long) As Long\nNativeSupport=value*2\nEnd Function"}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();
    sql.prepare("INSERT INTO _la_objects(kind,name,definition) VALUES('module','NativeApplicationHolder',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba",R"VBA(Public SavedApp As Access.Application
Public Sub SaveModuleApplication()
Set SavedApp=Application
End Sub
Public Function SavedModuleApplication() As Access.Application
Set SavedModuleApplication=SavedApp
End Function
Public Function ModuleApplicationIdentity() As Boolean
ModuleApplicationIdentity=(SavedApp Is Application)
End Function)VBA"}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();
    open([&](QDialog* dialog) {
        auto* choices=dialog->findChild<QComboBox*>("native-module-procedure");auto* run=dialog->findChild<QPushButton*>("native-module-run");auto* result=dialog->findChild<QTextBrowser*>("native-module-result");auto* editor=dialog->findChild<QPlainTextEdit*>("native-module-source");assert(choices && run && result && editor);
        const auto select=[&](const QString& name) {const int index=choices->findData(name);assert(index>=0);choices->setCurrentIndex(index);};
        bool sawStop=false,sawAssertion=false;
        QTimer debugErrors;debugErrors.setInterval(1);QObject::connect(&debugErrors,&QTimer::timeout,dialog,[&] {auto* debugger=dialog->findChild<QDialog*>("native-module-debugger");if(!debugger || !debugger->isVisible()) return;
            if(auto* status=debugger->findChild<QLabel*>("native-debugger-status")) {
                if(status->text().startsWith("Stop-Anweisung") || status->text().startsWith("Debug.Assert")) {
                    auto* locals=debugger->findChild<QTreeWidget*>("native-debugger-locals");assert(locals);bool found=false;
                    for(int i=0;i<locals->topLevelItemCount();++i) if(locals->topLevelItem(i)->text(0)=="n") {assert(locals->topLevelItem(i)->text(3)=="17");found=true;}assert(found);
                    sawStop|=status->text().startsWith("Stop-Anweisung");sawAssertion|=status->text().startsWith("Debug.Assert");
                }
            }
            if(auto* resume=debugger->findChild<QPushButton*>("native-debugger-continue");resume && resume->isEnabled()) resume->click();else if(auto* close=debugger->findChild<QPushButton*>("native-debugger-close");close && close->isEnabled()) close->click();});debugErrors.start();
        select("NativeExecute");QApplication::setActiveWindow(dialog);dialog->activateWindow();editor->setFocus();QApplication::processEvents();QTest::keyClick(editor,Qt::Key_F5);QApplication::processEvents();assert(result->toPlainText().startsWith("Prozedur ausgeführt."));
        assert(sql.exec("SELECT n FROM editor_module_effects") && sql.next() && sql.value(0).toInt()==22 && !sql.next());sql.finish();
        assert(sql.exec("SAVEPOINT native_halt_probe"));select("NativeHaltProbe");run->click();
        assert(sawStop && sawAssertion && result->toPlainText().startsWith("Prozedur ausgeführt."));
        assert(sql.exec("SELECT COUNT(*) FROM editor_module_effects WHERE n=18") && sql.next() && sql.value(0).toInt()==1);sql.finish();
        assert(sql.exec("ROLLBACK TO native_halt_probe") && sql.exec("RELEASE native_halt_probe"));
        assert(application->member("Run",{VbaValue::string("NativeCounter")}).number==44);
        select("NativeCounter");run->click();assert(result->toPlainText().contains("Rückgabewert: 44"));
        select("NativeApplicationEval");run->click();assert(result->toPlainText().contains("Rückgabewert: 7"));
        select("NativeApplicationObjectArg");run->click();assert(result->toPlainText().contains("Rückgabewert: 22"));
        select("NativeRecursive");run->click();assert(result->toPlainText().startsWith("Fehler 28"));
        select("NativeKeepApplication");run->click();assert(result->toPlainText().startsWith("Prozedur ausgeführt."));
        select("NativePrompt");QTimer promptAnswer;promptAnswer.setInterval(1);QObject::connect(&promptAnswer,&QTimer::timeout,dialog,[] {if(auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget())) box->button(QMessageBox::Yes)->click();});promptAnswer.start();run->click();promptAnswer.stop();assert(result->toPlainText().contains("Rückgabewert: 6"));
        select("NativeAdd");run->click();assert(result->toPlainText().startsWith("Fehler 450"));
        select("NativeInvalidContext");run->click();assert(result->toPlainText().startsWith("Fehler 438"));
        select("NativePreflight");run->click();assert(result->toPlainText().startsWith("Fehler 3251"));
        assert(sql.exec("SELECT COUNT(*) FROM editor_module_effects") && sql.next() && sql.value(0).toInt()==1);sql.finish();
        editor->appendPlainText("'unsaved");assert(!run->isEnabled());dialog->findChild<QPushButton*>("native-module-analyze")->click();select("NativeExecute");run->click();assert(result->toPlainText().startsWith("Fehler 3251"));
        assert(application->member("Run",{VbaValue::string("NativeCounter")}).number==44);debugErrors.stop();dialog->reject();
    });
    const auto heldApplication=application->member("Run",{VbaValue::string("SavedModuleApplication")}).object;assert(heldApplication);
    assert(heldApplication->member("Run",{VbaValue::string("NativeOptional"),VbaValue::longInteger(8)}).number==8);
    assert(heldApplication->member("Run",{VbaValue::string("ModuleApplicationIdentity")}).truthy());
    try {(void)heldApplication->member("Run",{VbaValue::string("NativeFailure")});assert(false);}catch(const linaccess::VbaRuntimeError& failure) {assert(failure.number==7001 && failure.source=="NativeRun");}
    const auto generation=window.property("la_project_generation");window.setProperty("la_project_generation",generation.toULongLong()+1);int staleApplication=0;try {(void)heldApplication->member("Run",{VbaValue::string("NativeOptional")});}catch(const linaccess::VbaRuntimeError& failure) {staleApplication=failure.number;}assert(staleApplication==3420);window.setProperty("la_project_generation",generation);
    assert(sql.exec("DELETE FROM _la_objects WHERE kind='module' AND name='NativeApplicationHolder'"));
    assert(sql.exec("DELETE FROM _la_objects WHERE kind='module' AND name='NativeSupportModule'"));
    open([&](QDialog* dialog) {QTimer::singleShot(0,&window,[&] {auto* message=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(message);message->button(QMessageBox::Yes)->click();});dialog->findChild<QPushButton*>("native-module-delete")->click();});
    assert(modules->member("Count",{}).number==baseline && !descriptor->member("IsLoaded",{}).truthy() && archive()==originalArchive);assert(!prepared.exec());assert(context.close());
}

void testEditableControlFormats(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;
    auto database=linaccess::makeDaoDatabase(db);auto table=database->member("CreateTableDef",{VbaValue::string("format_edit")}).object;
    for(const auto& [name,type]:std::vector<std::pair<const char*,int>>{{"id",4},{"amount",7},{"label",10},{"moment",8}}) table->member("Fields",{}).object->member("Append",{table->member("CreateField",{VbaValue::string(name),VbaValue::integer(type),VbaValue::longInteger(type==10 ? 255 : 0)})});
    database->member("TableDefs",{}).object->member("Append",{VbaValue::reference(table)});
    database->member("Execute",{VbaValue::string("INSERT INTO format_edit VALUES(1,0.12345678901234567,'before',#2003-12-31 18:05:06#)")});
    database->member("Execute",{VbaValue::string("INSERT INTO format_edit VALUES(2,2.5,'second',#2004-02-29 00:00:00#)")});
    QSqlQuery sql(db);assert(sql.exec("CREATE TABLE format_events(kind TEXT)"));
    assert(sql.exec("SELECT amount FROM format_edit WHERE id=1") && sql.next());const double initial=sql.value(0).toDouble();sql.finish();
    const QByteArray source=R"FORM(Begin Form
RecordSource ="format_edit"
Width =7000
Begin Section
Height =2000
Begin TextBox
Name ="Amount"
ControlSource ="amount"
Format ="0.00"
OnChange ="[Event Procedure]"
BeforeUpdate ="[Event Procedure]"
AfterUpdate ="[Event Procedure]"
Width =3000
Height =300
End
Begin TextBox
Name ="Duplicate"
ControlSource ="amount"
Format ="0.0000"
Top =400
Width =3000
Height =300
End
Begin TextBox
Name ="Label"
ControlSource ="label"
Format =">"
Top =800
Width =3000
Height =300
End
Begin TextBox
Name ="Moment"
ControlSource ="moment"
Format ="yyyy-mm-dd"
Top =1200
Width =3000
Height =300
End
End
End
)FORM";
    const QByteArray code=R"VBA(Private Sub Amount_Change()
CurrentDb.Execute "INSERT INTO format_events VALUES('change')"
End Sub
Private Sub Amount_BeforeUpdate(Cancel As Integer)
If Me!Amount<0 Then Cancel=True
End Sub
Private Sub Amount_AfterUpdate()
Me!Amount.Format="0.00"
End Sub
Sub Store()
DoCmd.RunCommand acCmdSaveRecord
End Sub
Sub FindDateSecond()
Me!Moment.SetFocus
DoCmd.FindRecord FindWhat:="2004-02-29",Match:=1,Search:=2,SearchAsFormatted:=True,OnlyCurrentField:=-1,FindFirst:=True
End Sub
Sub FindAllFormatted()
DoCmd.FindRecord "MIXED",1,True,2,True,0,True
End Sub
Sub FindDuplicateFormatted()
Me!Duplicate.SetFocus
DoCmd.FindRecord "0.1235",1,False,2,True,-1,True
End Sub
Sub FindAllRawLabel()
Me!Amount.SetFocus
DoCmd.FindRecord "mixed",1,True,2,False,0,True
End Sub
Sub FindAllRawNumber()
Me!Label.SetFocus
DoCmd.FindRecord 2.5,1,False,2,False,0,True
End Sub
Sub FindAllRawDate()
Me!Label.SetFocus
DoCmd.FindRecord CDate("2003-12-31 18:05:06"),1,False,2,False,0,True
End Sub
Sub FindWildcardLabel()
Me!Label.SetFocus
DoCmd.FindRecord "m[!a]x??",1,False,2,False,-1,True
End Sub
Sub InvalidWildcardLabel()
Me!Label.SetFocus
DoCmd.FindRecord "[",1,False,2,False,-1,True
End Sub
Sub OpenSearch()
Me!Moment.SetFocus
DoCmd.RunCommand acCmdFind
End Sub
Sub FindLabelPlainWhenMasked()
Me!Label.SetFocus
DoCmd.FindRecord "mixed",1,False,2,True,-1,True
End Sub
Sub OpenLabelSearch()
Me!Label.SetFocus
DoCmd.RunCommand acCmdFind
End Sub


)VBA";
    AccessFormPreview form("editable formatted controls",source,code,db,&window);
    auto root=makePreviewFormObject(form);auto controls=root->member("Controls",{}).object;auto amount=controls->item("Amount").object;
    auto* input=form.findChild<QLineEdit*>("access-control-Amount");auto* duplicate=form.findChild<QLineEdit*>("access-control-Duplicate");auto* label=form.findChild<QLineEdit*>("access-control-Label");auto* moment=form.findChild<QLineEdit*>("access-control-Moment");assert(input && duplicate && label && moment);
    assert(input->text()==QLocale().toString(0.12,'f',2) && label->text()=="BEFORE" && moment->text()=="2003-12-31");
    assert(amount->scalarValue().number==initial && controls->item("Label").object->scalarValue().asText()=="before" && controls->item("Moment").object->scalarValue().kind==VbaValue::Kind::Date);
    form.show();QApplication::setActiveWindow(&form);input->setFocus();QApplication::processEvents();
    assert(input->hasFocus() && input->text()==QString::number(initial,'g',std::numeric_limits<double>::max_digits10));
    label->setFocus();QApplication::processEvents();assert(input->text()==QLocale().toString(0.12,'f',2) && label->text()=="before");
    moment->setFocus();QApplication::processEvents();assert(moment->text()=="2003-12-31 18:05:06");label->setFocus();QApplication::processEvents();
    assert(moment->text()=="2003-12-31" && !root->member("Dirty",{}).truthy());
    assert(sql.exec("SELECT COUNT(*) FROM format_events") && sql.next() && sql.value(0).toInt()==0);sql.finish();
    input->setFocus();input->selectAll();QTest::keyClicks(input,"-1");label->setFocus();QApplication::processEvents();
    assert(input->hasFocus() && input->text()=="-1" && amount->scalarValue().number==initial);
    int rejected=0;try {amount->assign("Format",VbaValue::string("0.000"));}catch(const linaccess::VbaRuntimeError& failure) {rejected=failure.number;}assert(rejected==3251 && input->text()=="-1");
    amount->member("Undo",{});assert(!root->member("Dirty",{}).truthy() && amount->scalarValue().number==initial);
    input->selectAll();QTest::keyClicks(input,"0.1234");label->setFocus();QApplication::processEvents();
    assert(!input->hasFocus() && input->text()==QLocale().toString(0.12,'f',2) && duplicate->text()==QLocale().toString(0.1234,'f',4) && amount->scalarValue().number==0.1234);
    assert(sql.exec("SELECT amount FROM format_edit WHERE id=1") && sql.next() && sql.value(0).toDouble()==initial);sql.finish();
    std::vector<VbaValue> arguments;auto stored=form.executeProcedure("Store",arguments);if(stored.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Formatted save "<<stored.errorNumber<<" "<<stored.explanation<<"\n";assert(stored.status==linaccess::VbaExecutionStatus::Completed);
    assert(sql.exec("SELECT amount,moment FROM format_edit WHERE id=1") && sql.next() && sql.value(0).toDouble()==0.1234 && sql.value(1).toString()=="2003-12-31 18:05:06");sql.finish();
    for(const double next:{0.12345,0.12346}) {
        amount->assign("Value",VbaValue::numeric(next));assert(root->member("Dirty",{}).truthy() && amount->scalarValue().number==next);
        assert(form.executeProcedure("Store",arguments).status==linaccess::VbaExecutionStatus::Completed);
        assert(sql.exec("SELECT amount FROM format_edit WHERE id=1") && sql.next() && sql.value(0).toDouble()==next);sql.finish();
    }
    label->setFocus();label->selectAll();QTest::keyClicks(label,"mixed");moment->setFocus();QApplication::processEvents();
    assert(label->text()=="MIXED" && controls->item("Label").object->scalarValue().asText()=="mixed");
    assert(form.executeProcedure("Store",arguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(sql.exec("SELECT label,moment FROM format_edit WHERE id=1") && sql.next() && sql.value(0).toString()=="mixed" && sql.value(1).toString()=="2003-12-31 18:05:06");sql.finish();
    assert(form.executeProcedure("FindDateSecond",arguments).status==linaccess::VbaExecutionStatus::Completed && root->member("CurrentRecord",{}).number==2);
    assert(form.executeProcedure("FindAllRawLabel",arguments).status==linaccess::VbaExecutionStatus::Completed && root->member("CurrentRecord",{}).number==1);
    assert(form.executeProcedure("FindAllRawNumber",arguments).status==linaccess::VbaExecutionStatus::Completed && root->member("CurrentRecord",{}).number==2);
    assert(form.executeProcedure("FindAllRawDate",arguments).status==linaccess::VbaExecutionStatus::Completed && root->member("CurrentRecord",{}).number==1);
    assert(form.executeProcedure("FindDateSecond",arguments).status==linaccess::VbaExecutionStatus::Completed && root->member("CurrentRecord",{}).number==2);
    assert(form.executeProcedure("FindWildcardLabel",arguments).status==linaccess::VbaExecutionStatus::Completed && root->member("CurrentRecord",{}).number==1);
    assert(form.executeProcedure("InvalidWildcardLabel",arguments).errorNumber==93 && root->member("CurrentRecord",{}).number==1);
    assert(form.executeProcedure("FindDateSecond",arguments).status==linaccess::VbaExecutionStatus::Completed && root->member("CurrentRecord",{}).number==2);
    assert(form.executeProcedure("FindAllFormatted",arguments).status==linaccess::VbaExecutionStatus::Completed && root->member("CurrentRecord",{}).number==1);
    assert(form.executeProcedure("FindDateSecond",arguments).status==linaccess::VbaExecutionStatus::Completed && root->member("CurrentRecord",{}).number==2);
    assert(form.executeProcedure("FindDuplicateFormatted",arguments).status==linaccess::VbaExecutionStatus::Completed && root->member("CurrentRecord",{}).number==1);
    assert(form.executeProcedure("FindDateSecond",arguments).status==linaccess::VbaExecutionStatus::Completed);
    assert(form.executeProcedure("OpenSearch",arguments).status==linaccess::VbaExecutionStatus::Completed);auto* search=form.findChild<QDialog*>("access-find-dialog");assert(search);
    search->findChild<QCheckBox*>("access-find-formatted")->setChecked(true);search->findChild<QCheckBox*>("access-find-wildcards")->setChecked(true);search->findChild<QLineEdit*>("access-find-text")->setText("2003-##-##");search->findChild<QComboBox*>("access-find-match")->setCurrentIndex(1);search->findChild<QPushButton*>("access-find-next")->click();
    assert(root->member("CurrentRecord",{}).number==1 && search->findChild<QLabel*>("access-find-result")->text().contains("gefunden"));search->reject();
    assert(sql.exec("SELECT label,moment FROM format_edit WHERE id=1") && sql.next() && sql.value(0).toString()=="mixed" && sql.value(1).toString()=="2003-12-31 18:05:06");sql.finish();
    assert(form.executeProcedure("FindDateSecond",arguments).status==linaccess::VbaExecutionStatus::Completed && root->member("CurrentRecord",{}).number==2);
    controls->item("Label").object->assign("InputMask",VbaValue::string("Password"));
    assert(form.executeProcedure("FindLabelPlainWhenMasked",arguments).status==linaccess::VbaExecutionStatus::Completed && root->member("CurrentRecord",{}).number==2);
    assert(form.executeProcedure("OpenLabelSearch",arguments).status==linaccess::VbaExecutionStatus::Completed);search=nullptr;for(auto* candidate:form.findChildren<QDialog*>("access-find-dialog")) if(candidate->isVisible()) search=candidate;assert(search);
    search->findChild<QCheckBox*>("access-find-formatted")->setChecked(true);search->findChild<QLineEdit*>("access-find-text")->setText("*****");search->findChild<QComboBox*>("access-find-match")->setCurrentIndex(1);search->findChild<QPushButton*>("access-find-next")->click();
    assert(root->member("CurrentRecord",{}).number==1 && controls->item("Label").object->scalarValue().asText()=="mixed");search->reject();
    controls->item("Label").object->assign("InputMask",VbaValue::string(""));
    assert(form.close());
    linaccess::AccessTableModel records(nullptr,db);records.setTable("format_edit");records.setEditStrategy(QSqlTableModel::OnManualSubmit);assert(records.select());
    linaccess::AccessTableView sheet;sheet.setModel(&records);sheet.setCurrentIndex(records.index(1,2));
    linaccess::NativeTableFindArguments find{"m[!a]x??",2,1,false,1,2};find.wildcards=true;
    assert(sheet.findText(find)==1 && sheet.currentIndex().row()==0 && !records.isDirty());
    find.text="*";find.wildcards=false;assert(sheet.findText(find)==0);find.wildcards=true;assert(sheet.findText(find)==1);
    find.text="[";assert(sheet.findText(find)==-1 && records.lastError().nativeErrorCode()=="93");
    find.text="s*c?nd";assert(sheet.findText(find)==1 && sheet.currentIndex().row()==1);
    assert(sheet.openFindDialog());auto* dialog=sheet.findChild<QDialog*>("native-table-find-dialog");assert(dialog);dialog->findChild<QCheckBox*>("native-table-find-wildcards")->setChecked(true);dialog->findChild<QLineEdit*>("native-table-find-text")->setText("M?XED");dialog->findChild<QComboBox*>("native-table-find-match")->setCurrentIndex(1);dialog->findChild<QPushButton*>("native-table-find-next")->click();assert(sheet.currentIndex().row()==0);dialog->reject();
    linaccess::AccessQueryModel results;results.execute(db,"SELECT label FROM format_edit ORDER BY id");linaccess::NativeQueryView query;query.setModel(&results);query.setCurrentIndex(results.index(1,0));find.column=0;find.text="m[!a]x??";assert(query.findText(find)==1 && query.currentIndex().row()==0);
    find.text=QString(16385,'*');bool bounded=false;try{query.findText(find);}catch(const linaccess::VbaRuntimeError& error){bounded=error.number==3251;}assert(bounded);

}

static void exerciseNativeDecimalClipboard() {
    using namespace linaccess;
    auto db=QSqlDatabase::addDatabase("QSQLITE","decimal-grid-clipboard");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);assert(sql.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));sql.finish();
    const auto schema=parseMdbSchema("CREATE TABLE `decimal_copy_probe`\n(\n`v` NUMERIC(28,4)\n);\n");importAccessTable(db,"technical-decimal","decimal_copy_probe",schema,parseCsvDetailed("v\n9007199254740993.0001\n"));
    assert(sql.exec("CREATE TABLE decimal_text_destination(value TEXT)"));assert(sql.exec("INSERT INTO decimal_text_destination VALUES('')"));sql.finish();
    {
        AccessTableModel source(nullptr,db);source.setEditStrategy(QSqlTableModel::OnManualSubmit);source.setTable("decimal_copy_probe");assert(source.select());AccessTableView sourceView;sourceView.setModel(&source);sourceView.setCurrentIndex(source.index(0,0));sourceView.selectionModel()->select(source.index(0,0),QItemSelectionModel::ClearAndSelect);assert((source.flags(source.index(0,0))&Qt::ItemIsEditable) && sourceView.copySelection() && QApplication::clipboard()->text()=="9007199254740993.0001");
        NativeTableFindArguments exactFind;exactFind.text="9007199254740993.0001";exactFind.column=0;exactFind.origin=1;exactFind.mode=1;exactFind.value=decimalLiteral("9007199254740993.0001");
        assert(nativeFindEqual(nativeFindValueFromPayload(nativeFindValuePayload(*exactFind.value)),*exactFind.value));assert(sourceView.findText(exactFind)==1);exactFind.value=decimalLiteral("9007199254740993.0002");assert(sourceView.findText(exactFind)==0);
        AccessQueryModel decimalResults;decimalResults.execute(db,"SELECT v FROM decimal_copy_probe");NativeQueryView decimalQuery;decimalQuery.setModel(&decimalResults);exactFind.value=decimalLiteral("9007199254740993.0001");assert(decimalQuery.findText(exactFind)==1);exactFind.value.reset();exactFind.formatted=true;assert(decimalQuery.findText(exactFind)==1);
        AccessTableModel target(nullptr,db);target.setEditStrategy(QSqlTableModel::OnManualSubmit);target.setTable("decimal_text_destination");assert(target.select());AccessTableView targetView;targetView.setModel(&target);targetView.setCurrentIndex(target.index(0,0));assert(targetView.pasteSelection() && target.index(0,0).data().toString()=="9007199254740993.0001" && target.submitAll());
        assert(sql.exec("SELECT value,typeof(value) FROM decimal_text_destination") && sql.next() && sql.value(0).toString()=="9007199254740993.0001" && sql.value(1).toString()=="text");sql.finish();
        sourceView.resize(400,150);sourceView.show();sourceView.edit(source.index(0,0));QCoreApplication::processEvents();auto* decimalEditor=sourceView.findChild<QLineEdit*>();assert(decimalEditor && decimalEditor->text()=="9007199254740993.0001");decimalEditor->setText("9007199254740994.0001");QTest::keyClick(decimalEditor,Qt::Key_Return);QCoreApplication::processEvents();assert(source.index(0,0).data().toString()=="9007199254740994.0001" && source.isDirty());assert(source.undoPendingRecord(0) && !source.isDirty() && source.index(0,0).data().toString()=="9007199254740993.0001");
        QApplication::clipboard()->setText("9007199254740994.0001");assert(sourceView.pasteSelection() && source.canUndoBlock() && source.undoBlock() && source.index(0,0).data().toString()=="9007199254740993.0001");assert(source.redoBlock() && source.index(0,0).data().toString()=="9007199254740994.0001" && source.submitTransactionally());
        assert(!source.setData(source.index(0,0),QString("1.00001")) && !source.isDirty() && source.index(0,0).data().toString()=="9007199254740994.0001");
        sourceView.setCurrentIndex(source.index(0,0));sourceView.selectionModel()->select(source.index(0,0),QItemSelectionModel::ClearAndSelect);assert(sourceView.clearSelectedCells() && source.index(0,0).data(Qt::EditRole).isNull() && source.submitTransactionally());assert(sql.exec("SELECT v FROM decimal_copy_probe") && sql.next() && sql.value(0).isNull());sql.finish();sourceView.close();
    }
    {
        const QByteArray layout("Version =20\nBegin Form\nRecordSource =\"decimal_copy_probe\"\nWidth =3000\nBegin Section\nHeight =600\nBegin TextBox\nName =\"Value\"\nControlSource =\"v\"\nWidth =2400\nHeight =300\nEnd\nEnd\nEnd");
        AccessFormPreview form("DecimalForm",layout,{},db);form.show();QApplication::processEvents();auto root=makePreviewFormObject(form);auto value=root->member("Controls",{VbaValue::string("Value")}).object;
        value->assign("Value",VbaValue::string("9007199254740993.0001"));assert(value->scalarValue().asText()=="9007199254740993.0001");root->assign("Dirty",VbaValue::logical(false));
        assert(sql.exec("SELECT v FROM decimal_copy_probe") && sql.next() && qtFieldValue(sql.value(0),"DECIMAL_EXACT").asText()=="9007199254740993.0001");sql.finish();
        auto* field=form.findChild<QLineEdit*>("access-control-Value");assert(field);field->setFocus();QApplication::processEvents();makePreviewApplicationObject(form)->member("DoCmd",{}).object->member("FindRecord",{decimalLiteral("9007199254740993.0001")});
        const auto count=[&] {return root->member("Recordset",{}).object->member("RecordCount",{}).asNumber();};
        root->assign("Filter",VbaValue::string("[v] > CDec(\"9007199254740993.0002\")"));root->assign("FilterOn",VbaValue::logical(true));assert(count()==0);root->assign("FilterOn",VbaValue::logical(false));assert(count()==1);
        root->assign("Filter",VbaValue::string("[v] = CDec(\"9007199254740993.0001\")"));root->assign("FilterOn",VbaValue::logical(true));assert(count()==1);root->assign("FilterOn",VbaValue::logical(false));
        auto first=root->member("RecordsetClone",{}).object;first->member("MoveFirst",{});root->assign("Bookmark",first->member("Bookmark",{}));
        field->setFocus();QTest::keyClick(field,Qt::Key_A,Qt::ControlModifier);QTest::keyClicks(field,"9007199254740993.00010");root->assign("Dirty",VbaValue::logical(false));assert(value->scalarValue().asText()=="9007199254740993.0001" && nativeFindEqual(nativeFindInput("9007199254740993.00010","DECIMAL_EXACT"),value->scalarValue()));
        int error=0;try {value->assign("Value",VbaValue::string("1.00001"));}catch(const VbaRuntimeError& failure) {error=failure.number;}assert(error==6 && !root->member("Dirty",{}).boolean && value->scalarValue().asText()=="9007199254740993.0001");
        value->assign("Value",VbaValue{});root->assign("Dirty",VbaValue::logical(false));assert(sql.exec("SELECT v FROM decimal_copy_probe") && sql.next() && sql.value(0).isNull());sql.finish();assert(form.close());
    }
    QApplication::clipboard()->clear();sql.clear();db.close();
}

static void exercisePhysicalFormIdentity() {
    using namespace linaccess;
    auto db=QSqlDatabase::addDatabase("QSQLITE","physical-form-identity");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE form_identity_probe(rowid INTEGER,__la_rowid INTEGER,v TEXT)"));assert(sql.exec("INSERT INTO form_identity_probe VALUES(7,91,'first'),(7,91,'second')"));sql.finish();
    const QByteArray layout("Begin Form\nRecordSource =\"form_identity_probe\"\nWidth =3000\nBegin Section\nHeight =600\nBegin TextBox\nName =\"Value\"\nControlSource =\"v\"\nWidth =2400\nHeight =300\nEnd\nEnd\nEnd");
    {
        AccessFormPreview form("PhysicalIdentity",layout,{},db);form.show();QApplication::processEvents();assert(form.hasDataSource());auto root=makePreviewFormObject(form);auto value=root->member("Controls",{VbaValue::string("Value")}).object;
        value->assign("Value",VbaValue::string("changed"));root->assign("Dirty",VbaValue::logical(false));
        assert(sql.exec("SELECT v FROM form_identity_probe ORDER BY _rowid_") && sql.next() && sql.value(0).toString()=="changed" && sql.next() && sql.value(0).toString()=="second");sql.finish();
        QTimer::singleShot(0,[] {auto* dialog=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="access-delete-confirm");dialog->done(QMessageBox::Yes);});makePreviewApplicationObject(form)->member("DoCmd",{}).object->member("RunCommand",{VbaValue::integer(223)});
        assert(sql.exec("SELECT v FROM form_identity_probe") && sql.next() && sql.value(0).toString()=="second" && !sql.next());sql.finish();
        root->assign("DataEntry",VbaValue::logical(true));value->assign("Value",VbaValue::string("added"));root->assign("Dirty",VbaValue::logical(false));assert(root->member("Recordset",{}).object->member("RecordCount",{}).asNumber()==1 && value->scalarValue().asText()=="added");assert(form.close());
    }
    assert(sql.exec("CREATE TABLE identity_unavailable(rowid INTEGER,_rowid_ INTEGER,oid INTEGER,v TEXT)"));assert(sql.exec("INSERT INTO identity_unavailable VALUES(1,1,1,'unchanged')"));sql.finish();
    {
        auto blockedLayout=layout;blockedLayout.replace("form_identity_probe","identity_unavailable");AccessFormPreview blocked("IdentityUnavailable",blockedLayout,{},db);assert(!blocked.hasDataSource() && blocked.dataSourceErrorNumber()==3251);assert(blocked.close());
    }
    assert(sql.exec("SELECT COUNT(*) FROM form_identity_probe") && sql.next() && sql.value(0).toInt()==2);sql.finish();sql.clear();db.close();
}

static void exerciseNativeWarnings() {
    using namespace linaccess;assert(nativeWarningsEnabled());const auto restore=qScopeGuard([] {restoreNativeWarnings();});
    auto db=QSqlDatabase::addDatabase("QSQLITE","native-warning-state");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE _la_objects(kind TEXT NOT NULL,name TEXT NOT NULL,definition TEXT NOT NULL,PRIMARY KEY(kind,name))"));
    assert(sql.exec("CREATE TABLE warnings_rows(id INTEGER PRIMARY KEY,v TEXT)"));assert(sql.exec("INSERT INTO warnings_rows VALUES(1,'one'),(2,'two'),(3,'three')"));sql.finish();
    const QByteArray layout("Begin Form\nRecordSource =\"warnings_rows\"\nWidth =2400\nBegin Section\nHeight =600\nBegin TextBox\nName =\"Value\"\nControlSource =\"v\"\nWidth =1800\nHeight =300\nEnd\nEnd\nEnd");
    {
        AccessFormPreview form("WarningsForm",layout,{},db);form.show();QApplication::processEvents();assert(form.hasDataSource());NativeStandardModuleHost host(&form,db,[] {});auto commands=host.applicationObject()->member("DoCmd",{}).object;auto formCommands=makePreviewApplicationObject(form)->member("DoCmd",{}).object;
        const auto failure=executeVbaProcedure("Sub DisableWarningsAndFail()\nDoCmd.SetWarnings WarningsOn:=False\nErr.Raise 5\nEnd Sub","DisableWarningsAndFail",host);assert(failure.status==VbaExecutionStatus::Error && failure.errorNumber==5 && !nativeWarningsEnabled());
        for(const auto& rejected:std::vector<std::pair<std::vector<VbaValue>,int>>{{{},450},{{VbaValue{}},94}}) {int error=0;try {commands->member("SetWarnings",rejected.first);}catch(const VbaRuntimeError& failure) {error=failure.number;}assert(error==rejected.second && !nativeWarningsEnabled());}
        formCommands->member("RunCommand",{VbaValue::integer(223)});assert(sql.exec("SELECT COUNT(*) FROM warnings_rows WHERE id=1") && sql.next() && sql.value(0).toInt()==0);sql.finish();formCommands->member("SetWarnings",{VbaValue::logical(true)});assert(nativeWarningsEnabled());
        const auto action=[](QString name,QJsonArray arguments=QJsonArray{}) {return QJsonObject{{"action",name},{"arguments",arguments}};};const auto store=[&](QString name,QJsonArray actions) {(void)storeNativeMacro(db,name,{{"version",1},{"actions",actions}});};
        store("WarningsSilentDelete",{action("SetWarnings",{false}),action("RunCommand",{223})});commands->member("RunMacro",{VbaValue::string("WarningsSilentDelete")});assert(nativeWarningsEnabled() && sql.exec("SELECT COUNT(*) FROM warnings_rows") && sql.next() && sql.value(0).toInt()==1);sql.finish();
        store("WarningsStop",{action("SetWarnings",{false}),action("StopAllMacros")});formCommands->member("RunMacro",{VbaValue::string("WarningsStop")});assert(nativeWarningsEnabled() && !nativeMacroStopAllPending());
        store("WarningsError",{action("SetWarnings",{false}),action("RunSQL",{"UPDATE missing_warning_source SET n=1"})});int error=0;try {commands->member("RunMacro",{VbaValue::string("WarningsError")});}catch(const VbaRuntimeError& failure) {error=failure.number;}assert(error==3078 && nativeWarningsEnabled());
        commands->member("RunSQL",{VbaValue::string("UPDATE warnings_rows SET v='updated' WHERE id=3")});assert(sql.exec("SELECT v FROM warnings_rows WHERE id=3") && sql.next() && sql.value(0).toString()=="updated");sql.finish();
        commands->member("SetWarnings",{VbaValue::logical(false)});bool explicitMessage=false;QTimer::singleShot(0,[&] {auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(box && box->text()=="Explicit warnings probe");explicitMessage=true;box->done(QMessageBox::Ok);});const auto message=executeVbaProcedure("Sub ExplicitMessage()\nMsgBox \"Explicit warnings probe\"\nEnd Sub","ExplicitMessage",host);assert(message.status==VbaExecutionStatus::Completed && explicitMessage && !nativeWarningsEnabled());formCommands->member("SetWarnings",{VbaValue::logical(true)});assert(form.close());
    }
    sql.clear();db.close();
}

int main(int argc,char** argv) {
    QCoreApplication::setAttribute(Qt::AA_DontUseNativeDialogs);
    QApplication app(argc,argv); app.setQuitOnLastWindowClosed(false); app.setApplicationName("LinAccessSyntheticGuiTest"); app.setOrganizationName("LinAccessSyntheticTests");
    QStandardPaths::setTestModeEnabled(true);
    const auto profileStart=std::chrono::steady_clock::now();
    const auto phase=[&](const char* name) {if(qEnvironmentVariableIsSet("LINACCESS_GUI_PROFILE")) std::cerr<<"GUI phase "<<name<<" ms="<<std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now()-profileStart).count()<<'\n';};
    QTimer watchdog; watchdog.setSingleShot(true); watchdog.setTimerType(Qt::PreciseTimer);
    QObject::connect(&watchdog,&QTimer::timeout,[] {
        auto* modal=QApplication::activeModalWidget();
        std::cerr<<"Synthetic main-window document test timed out; modal="<<(modal ? modal->metaObject()->className() : "none")<<"\n";
        std::abort();
    }); watchdog.start(20000);
    QTemporaryDir temporary; assert(temporary.isValid());
    QSettings::setDefaultFormat(QSettings::IniFormat);
    QSettings::setPath(QSettings::IniFormat,QSettings::UserScope,temporary.path());
    exerciseNativeRelationDesigner();
    exerciseNativeDecimalClipboard();
    exercisePhysicalFormIdentity();
    exerciseNativeWarnings();
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
    exerciseVisualQueryDesigner(window,db);
    exerciseNativeDebugOutput(window,db);
    exerciseNativeProperties(db);
    exerciseNativeStaticProcedures(db);
    exerciseNativeClasses(window,db);
    exerciseTableSave(window,db);
    exerciseReferenceDatasheet();
    exerciseDatasheetOutput(window,db);
    exerciseDatasheetPageSetup(window,db);
    exerciseTableReplace(window,db);
    exerciseTableMultiDelete(window,db);
    exerciseRecordEscape(window,db);
    exercisePageNavigation(window,db);
    exerciseDateDefaultBuffers();
    exerciseScalarRelationDesigner(window,5);
    exerciseScalarRelationDesigner(window,6);
    exerciseScalarRelationDesigner(window,8);
    exerciseStoredForm(window,db);
    exerciseStoredReport(window,db);phase("before-table-wizard");
    {
        QTimer::singleShot(0,&window,[&] {
            auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget()); assert(dialog && dialog->objectName()=="native-table-designer");
            dialog->findChild<QLineEdit*>("native-table-name")->setText("authored_defaults");
            auto* grid=dialog->findChild<QTableWidget*>("native-table-fields"); assert(grid && grid->columnCount()==6);
            grid->item(0,0)->setText("qty"); qobject_cast<QComboBox*>(grid->cellWidget(0,1))->setCurrentText("INTEGER");
            grid->item(0,2)->setText("ja"); grid->item(0,3)->setText("Now()");
            auto* buttons=dialog->findChild<QDialogButtonBox*>(); assert(buttons);
            QTimer::singleShot(0,dialog,[] { auto* error=qobject_cast<QMessageBox*>(QApplication::activeModalWidget()); assert(error); error->accept(); });
            buttons->button(QDialogButtonBox::Save)->click();
            assert(dialog->isVisible() && grid->item(0,0)->text()=="qty" && !db.tables().contains("authored_defaults"));
            grid->item(0,3)->setText("=7"); dialog->findChild<QPushButton*>("native-table-add-field")->click();
            grid->item(1,0)->setText("label"); grid->item(1,3)->setText("\"ui value\"");grid->item(1,4)->setText("12");
            dialog->findChild<QPushButton*>("native-table-add-field")->click();
            grid->item(2,0)->setText("enabled"); qobject_cast<QComboBox*>(grid->cellWidget(2,1))->setCurrentText("BOOLEAN"); grid->item(2,3)->setText("True");
            dialog->findChild<QPushButton*>("native-table-add-field")->click();grid->item(3,0)->setText("sample");qobject_cast<QComboBox*>(grid->cellWidget(3,1))->setCurrentText("SINGLE");grid->item(3,3)->setText("0.1");dialog->findChild<QPushButton*>("native-table-add-field")->click();grid->item(4,0)->setText("moment");qobject_cast<QComboBox*>(grid->cellWidget(4,1))->setCurrentText("DATE");grid->item(4,3)->setText("=#2026-11-08#");
            dialog->findChild<QPushButton*>("native-table-add-field")->click();grid->item(5,0)->setText("wizard_currency");auto* currencyType=qobject_cast<QComboBox*>(grid->cellWidget(5,1));currencyType->setCurrentText("CURRENCY");assert(currencyType->currentText()=="CURRENCY");grid->item(5,3)->setText("1.23455");
            grid->item(1,4)->setText("0");
            QTimer::singleShot(0,dialog,[] {auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(box);box->accept();});
            buttons->button(QDialogButtonBox::Save)->click();
            assert(dialog->isVisible() && !db.tables().contains("authored_defaults") && grid->item(1,3)->text()=="\"ui value\"");
            grid->item(1,4)->setText("12");qobject_cast<QCheckBox*>(grid->cellWidget(1,5))->setChecked(false);
            QTimer::singleShot(0,dialog,[] {auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(box);box->accept();});buttons->button(QDialogButtonBox::Save)->click();assert(dialog->isVisible() && !db.tables().contains("authored_defaults") && grid->item(5,3)->text()=="1.23455");
            grid->item(5,3)->setText("1.2345");buttons->button(QDialogButtonBox::Save)->click();
        });
        action(window,"+ Tabelle")->trigger();phase("after-table-wizard");
        QSqlQuery authored(db); assert(authored.exec("INSERT INTO authored_defaults DEFAULT VALUES"));
        assert(authored.exec("SELECT qty,label,enabled FROM authored_defaults") && authored.next());
        assert(authored.value(0).toInt()==7 && authored.value(1).toString()=="ui value" && authored.value(2).toInt()==-1); authored.finish();
        auto wizardTable=linaccess::makeDaoDatabase(db)->member("TableDefs",{linaccess::VbaValue::string("authored_defaults")}).object;auto wizardMoney=wizardTable->member("Fields",{linaccess::VbaValue::string("wizard_currency")}).object;assert(wizardMoney->member("Type",{}).asNumber()==5 && wizardMoney->member("Size",{}).asNumber()==8);
        assert(authored.exec("SELECT wizard_currency FROM authored_defaults") && authored.next() && linaccess::qtFieldValue(authored.value(0),"CURRENCY_REAL").currencyUnits==12345);authored.finish();wizardTable->member("Fields",{}).object->member("Delete",{linaccess::VbaValue::string("wizard_currency")});
        assert(linaccess::jetLongAccessColumns(db,"authored_defaults").contains("id"));
        assert(authored.exec("SELECT sample FROM authored_defaults") && authored.next() && authored.value(0).toDouble()==static_cast<double>(0.1f));authored.finish();assert(authored.exec("SELECT moment FROM authored_defaults") && authored.next() && authored.value(0).toString()=="2026-11-08 00:00:00");authored.finish();
        {linaccess::AccessTableModel model(nullptr,db);model.setTable("authored_defaults");model.setEditStrategy(QSqlTableModel::OnManualSubmit);assert(model.select());const auto cell=model.index(0,model.fieldIndex("sample"));assert(model.setData(cell,"0.2") && !model.setData(cell,"1e40"));assert(model.submitTransactionally());}
        assert(authored.exec("SELECT sample FROM authored_defaults") && authored.next() && authored.value(0).toDouble()==static_cast<double>(0.2f));authored.finish();
        auto definition=linaccess::makeDaoDatabase(db)->member("TableDefs",{linaccess::VbaValue::string("authored_defaults")}).object;
        auto labelField=definition->member("Fields",{linaccess::VbaValue::string("label")}).object;
        assert(labelField->member("Type",{}).asNumber()==10 && labelField->member("Size",{}).asNumber()==12);assert(!labelField->member("AllowZeroLength",{}).truthy());
        assert(definition->member("Fields",{linaccess::VbaValue::string("qty")}).object->member("Type",{}).asNumber()==4);
        auto primary=definition->member("Indexes",{linaccess::VbaValue::string("PrimaryKey")}).object;
        assert(primary->member("Primary",{}).truthy());
        auto extraIndex=definition->member("CreateIndex",{linaccess::VbaValue::string("LabelIndex")}).object;
        extraIndex->member("Fields",{}).object->member("Append",{extraIndex->member("CreateField",{linaccess::VbaValue::string("label")})});
        definition->member("Indexes",{}).object->member("Append",{linaccess::VbaValue::reference(extraIndex)});
        assert(definition->member("Indexes",{}).object->member("Count",{}).asNumber()==2);
        auto cursor=linaccess::makeDaoDatabase(db)->member("OpenRecordset",{linaccess::VbaValue::string("authored_defaults")}).object;
        cursor->member("AddNew",{});assert(cursor->item("id").asNumber()==2 && cursor->item("id").numericType==linaccess::VbaValue::NumericType::Long);cursor->member("CancelUpdate",{});
        cursor->member("AddNew",{});assert(cursor->item("id").asNumber()==3);cursor->member("Update",{});cursor->member("Close",{});
        assert(authored.exec("SELECT COUNT(*) FROM authored_defaults WHERE id=2") && authored.next() && authored.value(0).toInt()==0);authored.finish();
        const auto matching=explorer->findItems("authored_defaults",Qt::MatchExactly|Qt::MatchRecursive,0);assert(matching.size()==1);explorer->setCurrentItem(matching.front());
        QTimer::singleShot(0,&window,[&] {
            auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-index-designer");
            auto* list=dialog->findChild<QListWidget*>("native-index-list");assert(list && list->count()==2);
            dialog->findChild<QLineEdit*>("native-index-name")->setText("QtyDescending");
            auto* grid=dialog->findChild<QTableWidget*>("native-index-fields");assert(grid && grid->rowCount()==1);
            qobject_cast<QComboBox*>(grid->cellWidget(0,0))->setCurrentText("qty");qobject_cast<QComboBox*>(grid->cellWidget(0,1))->setCurrentIndex(1);
            dialog->findChild<QPushButton*>("native-index-create")->click();assert(list->count()==3);
            for(int row=0;row<list->count();++row) if(list->item(row)->data(Qt::UserRole).toString()=="QtyDescending") list->setCurrentRow(row);
            QTimer::singleShot(0,dialog,[] {auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(box);box->button(QMessageBox::Yes)->click();});
            dialog->findChild<QPushButton*>("native-index-delete")->click();assert(list->count()==2);
            for(int row=0;row<list->count();++row) if(list->item(row)->data(Qt::UserRole).toString()=="PrimaryKey") list->setCurrentRow(row);
            QTimer::singleShot(0,dialog,[] {auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(box);box->button(QMessageBox::Yes)->click();});dialog->findChild<QPushButton*>("native-index-delete")->click();assert(list->count()==1);
            dialog->findChild<QCheckBox*>("native-index-primary")->setChecked(true);dialog->findChild<QLineEdit*>("native-index-name")->setText("PrimaryKey");qobject_cast<QComboBox*>(grid->cellWidget(0,0))->setCurrentText("id");qobject_cast<QComboBox*>(grid->cellWidget(0,1))->setCurrentIndex(0);
            QTimer::singleShot(0,dialog,[] {auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());if(box && box->windowTitle()=="Index konnte nicht geändert werden") {std::cerr<<"Primary GUI error: "<<box->text().toStdString()<<"\n";box->accept();}});dialog->findChild<QPushButton*>("native-index-create")->click();assert(list->count()==2);
            dialog->findChild<QCheckBox*>("native-index-primary")->setChecked(false);dialog->findChild<QCheckBox*>("native-index-unique")->setChecked(false);dialog->findChild<QCheckBox*>("native-index-required")->setChecked(true);dialog->findChild<QLineEdit*>("native-index-name")->setText("LabelRequired");qobject_cast<QComboBox*>(grid->cellWidget(0,0))->setCurrentText("label");dialog->findChild<QPushButton*>("native-index-create")->click();assert(list->count()==3);
            QSqlQuery requiredState(db);assert(requiredState.exec("PRAGMA table_info(authored_defaults)"));bool labelRequired=false;while(requiredState.next()) if(requiredState.value(1).toString()=="label") labelRequired=requiredState.value(3).toBool();assert(labelRequired);requiredState.finish();
            for(int row=0;row<list->count();++row) if(list->item(row)->data(Qt::UserRole).toString()=="LabelRequired") list->setCurrentRow(row);QTimer::singleShot(0,dialog,[] {auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(box);box->button(QMessageBox::Yes)->click();});dialog->findChild<QPushButton*>("native-index-delete")->click();assert(list->count()==2);
            assert(requiredState.exec("PRAGMA table_info(authored_defaults)"));labelRequired=false;while(requiredState.next()) if(requiredState.value(1).toString()=="label") labelRequired=requiredState.value(3).toBool();assert(!labelRequired);requiredState.finish();
            dialog->reject();
        });
        auto* editIndexes=window.findChild<QAction*>("native-table-indexes");assert(editIndexes);editIndexes->trigger();
        assert(authored.exec("SELECT COUNT(*) FROM authored_defaults") && authored.next() && authored.value(0).toInt()==2);authored.finish();
        QTimer::singleShot(0,&window,[&] {
            auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-field-designer");
            auto* list=dialog->findChild<QTableWidget*>("native-field-list");assert(list && list->rowCount()==6);
            dialog->findChild<QLineEdit*>("native-field-name")->setText("ExtraNote");
            dialog->findChild<QSpinBox*>("native-field-size")->setValue(8);
            dialog->findChild<QLineEdit*>("native-field-default")->setText("\"note\"");
            dialog->findChild<QPushButton*>("native-field-create")->click();assert(list->rowCount()==7);
            QSqlQuery values(db);assert(values.exec("SELECT COUNT(*) FROM authored_defaults WHERE ExtraNote='note'") && values.next() && values.value(0).toInt()==2);values.finish();
            for(int row=0;row<list->rowCount();++row) if(list->item(row,0)->text()=="ExtraNote") list->setCurrentCell(row,0);
            dialog->findChild<QTabWidget*>("native-field-editor-tabs")->setCurrentIndex(1);
            dialog->findChild<QCheckBox*>("native-field-property-allow-zero")->setChecked(false);dialog->findChild<QSpinBox*>("native-field-property-size")->setValue(12);dialog->findChild<QCheckBox*>("native-field-property-required")->setChecked(true);dialog->findChild<QLineEdit*>("native-field-property-default")->setText("\"future\"");
            QTimer::singleShot(0,dialog,[] {auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());if(box && box->windowTitle()=="Feld konnte nicht geändert werden") {std::cerr<<"Field property GUI error: "<<box->text().toStdString()<<"\n";box->accept();}});
            dialog->findChild<QPushButton*>("native-field-property-save")->click();
            auto propertyField=linaccess::makeDaoDatabase(db)->member("TableDefs",{linaccess::VbaValue::string("authored_defaults")}).object->member("Fields",{linaccess::VbaValue::string("ExtraNote")}).object;
            assert(propertyField->member("Size",{}).asNumber()==12 && propertyField->member("Required",{}).truthy());assert(!propertyField->member("AllowZeroLength",{}).truthy());
            assert(values.exec("SELECT COUNT(*) FROM authored_defaults WHERE ExtraNote='note'") && values.next() && values.value(0).toInt()==2);values.finish();
            dialog->findChild<QLineEdit*>("native-field-rename-name")->setText("RenamedNote");
            QTimer::singleShot(0,dialog,[] {auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());if(box && box->windowTitle()=="Feld konnte nicht geändert werden") {std::cerr<<"Field rename GUI error: "<<box->text().toStdString()<<"\n";box->accept();}});
            dialog->findChild<QPushButton*>("native-field-rename")->click();assert(list->item(list->currentRow(),0)->text()=="RenamedNote");
            assert(values.exec("SELECT COUNT(*) FROM authored_defaults WHERE RenamedNote='note'") && values.next() && values.value(0).toInt()==2);values.finish();assert(!values.exec("INSERT INTO authored_defaults(RenamedNote) VALUES('')"));
            QTimer::singleShot(0,dialog,[] {auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(box);box->button(QMessageBox::Yes)->click();});
            dialog->findChild<QPushButton*>("native-field-delete")->click();assert(list->rowCount()==6);
            for(int row=0;row<list->rowCount();++row) if(list->item(row,0)->text()=="qty") list->setCurrentCell(row,0);auto* rule=dialog->findChild<QLineEdit*>("native-field-property-validation");assert(rule && rule->isEnabled());rule->setText("Between 0 And 10");dialog->findChild<QLineEdit*>("native-field-property-validation-text")->setText("Bereich für Menge beachten");dialog->findChild<QPushButton*>("native-field-property-save")->click();assert(list->item(list->currentRow(),5)->text()=="Between 0 And 10");assert(!values.exec("UPDATE authored_defaults SET qty=11") && values.lastError().databaseText().contains("Bereich für Menge beachten"));assert(values.exec("SELECT SUM(qty) FROM authored_defaults") && values.next() && values.value(0).toInt()==14);values.finish();dialog->findChild<QSpinBox*>("native-field-property-position")->setValue(0);dialog->findChild<QPushButton*>("native-field-move")->click();assert(list->item(0,0)->text()=="qty");assert(values.exec("SELECT * FROM authored_defaults") && values.record().fieldName(0)=="qty");values.finish();auto* numericType=dialog->findChild<QComboBox*>("native-field-property-type");assert(numericType && numericType->isEnabled());numericType->setCurrentIndex(numericType->findData(7));dialog->findChild<QPushButton*>("native-field-change-type")->click();assert(list->item(0,1)->text().contains("Double"));for(int row=0;row<list->rowCount();++row) if(list->item(row,0)->text()=="moment") list->setCurrentCell(row,0);assert(rule->isEnabled());rule->setText(">=#2000-01-01# And <=Now()");dialog->findChild<QPushButton*>("native-field-property-save")->click();assert(!values.exec("UPDATE authored_defaults SET moment='1999-12-31 00:00:00'"));for(int row=0;row<list->rowCount();++row) if(list->item(row,0)->text()=="qty") list->setCurrentCell(row,0);auto* caption=dialog->findChild<QLineEdit*>("native-field-property-caption");assert(caption && caption->isEnabled());caption->setText("Anzahl Ä");dialog->findChild<QLineEdit*>("native-field-property-description")->setText("Technischer Hinweis");auto* format=dialog->findChild<QLineEdit*>("native-field-property-format");assert(format && format->isEnabled());format->setText("0.00\" units\"");dialog->findChild<QPushButton*>("native-field-display-save")->click();assert(caption->text()=="Anzahl Ä");
            dialog->findChild<QTabWidget*>("native-field-editor-tabs")->setCurrentIndex(0);dialog->findChild<QLineEdit*>("native-field-name")->setText("ui_currency");auto* newType=dialog->findChild<QComboBox*>("native-field-type");assert(newType->findData(5)>=0);newType->setCurrentIndex(newType->findData(5));dialog->findChild<QLineEdit*>("native-field-default")->setText("1.2500");dialog->findChild<QLineEdit*>("native-field-validation")->setText("Between 0.0001 And 2.0001");dialog->findChild<QPushButton*>("native-field-create")->click();assert(list->rowCount()==7);
            for(int row=0;row<list->rowCount();++row) if(list->item(row,0)->text()=="ui_currency") list->setCurrentCell(row,0);dialog->findChild<QTabWidget*>("native-field-editor-tabs")->setCurrentIndex(1);assert(list->item(list->currentRow(),1)->text().contains("Währung") && rule->isEnabled());
            numericType->setCurrentIndex(numericType->findData(7));dialog->findChild<QPushButton*>("native-field-change-type")->click();assert(list->item(list->currentRow(),1)->text().contains("Double"));numericType->setCurrentIndex(numericType->findData(5));dialog->findChild<QPushButton*>("native-field-change-type")->click();assert(list->item(list->currentRow(),1)->text().contains("Währung"));
            rule->setText(">=1.2500 And <=2.0001");dialog->findChild<QLineEdit*>("native-field-property-validation-text")->setText("Währungsgrenze");
            dialog->findChild<QLineEdit*>("native-field-property-default")->setText("2.0001");dialog->findChild<QPushButton*>("native-field-property-save")->click();assert(dialog->findChild<QLineEdit*>("native-field-property-default")->text()=="2.0001");assert(values.exec("SELECT COUNT(*) FROM authored_defaults WHERE ui_currency=1.25") && values.next() && values.value(0).toInt()==2);values.finish();
            assert(!values.exec("UPDATE authored_defaults SET ui_currency=2.0002") && values.lastError().databaseText().contains("Währungsgrenze"));values.finish();
            QTimer::singleShot(0,dialog,[] {auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(box);box->button(QMessageBox::Yes)->click();});dialog->findChild<QPushButton*>("native-field-delete")->click();assert(list->rowCount()==6);dialog->reject();
        });
        auto* editFields=window.findChild<QAction*>("native-table-fields-edit");assert(editFields);editFields->trigger();phase("after-field-editor");
        assert(authored.exec("SELECT COUNT(*) FROM authored_defaults WHERE qty=7 AND label='ui value'") && authored.next() && authored.value(0).toInt()==2);authored.finish();
        {linaccess::AccessTableModel model(nullptr,db);model.setTable("authored_defaults");assert(model.select());const int column=model.fieldIndex("qty");assert(column==0 && model.headerData(column,Qt::Horizontal).toString()=="Anzahl Ä" && model.headerData(column,Qt::Horizontal,Qt::ToolTipRole).toString().contains("Technischer Hinweis") && model.record().fieldName(column)=="qty");}
        {
            linaccess::updateNativeFieldProperties(db,"authored_defaults","moment",{{"Format",linaccess::VbaValue::string("yyyy-mm-dd")}});
            linaccess::updateNativeFieldProperties(db,"authored_defaults","label",{{"Format",linaccess::VbaValue::string(">")}});
            linaccess::AccessTableModel model(nullptr,db);model.setTable("authored_defaults");assert(model.select());
            const int quantity=model.fieldIndex("qty"),moment=model.fieldIndex("moment");
            const auto expected=QLocale().toString(7.0,'f',2)+" units";
            assert(model.data(model.index(0,quantity)).toString()==expected);
            const int label=model.fieldIndex("label");assert(model.data(model.index(0,label)).toString()=="UI VALUE" && model.clipboardText(model.index(0,label))=="ui value");
            assert(model.data(model.index(0,quantity),Qt::EditRole).toDouble()==7 && model.clipboardText(model.index(0,quantity))=="7");
            QTableView grid;grid.setModel(&model);
            const auto snapshot=linaccess::nativeDatasheetSnapshot(grid,"Formatted fields");
            assert(snapshot.rows.size()==2 && snapshot.rows[0][quantity]==expected);
            linaccess::AccessTableView search;search.setModel(&model);search.setCurrentIndex(model.index(0,quantity));
            linaccess::NativeTableFindArguments find{expected,quantity,1,true,1,2};find.formatted=true;assert(search.findText(find)==1 && search.currentIndex().row()==0);
            assert(search.repeatFind() && search.currentIndex().row()==1);
            find={"ui value",label,1,false,1,2};find.formatted=true;assert(search.findText(find)==1 && search.currentIndex().row()==0);
            find.formatted=false;find.origin=0;assert(search.findText(find)==1 && search.currentIndex().row()==0);
            assert(search.openFindDialog());auto* dialog=search.findChild<QDialog*>("native-table-find-dialog");assert(dialog);
            dialog->findChild<QComboBox*>("native-table-find-field")->setCurrentIndex(dialog->findChild<QComboBox*>("native-table-find-field")->findData(quantity));
            dialog->findChild<QCheckBox*>("native-table-find-formatted")->setChecked(true);dialog->findChild<QLineEdit*>("native-table-find-text")->setText(expected);dialog->findChild<QComboBox*>("native-table-find-match")->setCurrentIndex(1);dialog->findChild<QPushButton*>("native-table-find-next")->click();
            assert(search.currentIndex().column()==quantity && dialog->findChild<QLabel*>("native-table-find-result")->text().contains("Übereinstimmung"));dialog->reject();
            linaccess::AccessQueryModel results;results.execute(db,"SELECT 1.2345 AS number, 'shown' AS label UNION ALL SELECT 2.5,'other'");linaccess::NativeQueryView queryView;queryView.setModel(&results);queryView.setCurrentIndex(results.index(0,0));
            linaccess::NativeTableFindArguments queryFind{"2.5",0,1,false,1,2};queryFind.formatted=true;assert(queryView.findText(queryFind)==1 && queryView.currentIndex().row()==1);

            assert(authored.exec("SELECT moment FROM authored_defaults LIMIT 1") && authored.next());const auto date=authored.value(0).toString().left(10);authored.finish();
            assert(snapshot.rows[0][moment]==date && model.data(model.index(0,moment)).toString()==date);
            QTemporaryDir output;assert(output.isValid());const auto path=output.filePath("formatted.pdf");linaccess::exportNativeDatasheetPdf(snapshot,path);
            const auto reader=QStandardPaths::findExecutable("pdftotext");if(!reader.isEmpty()) {QProcess text;text.start(reader,{"-layout",path,"-"});assert(text.waitForFinished(10000) && text.exitCode()==0);assert(QString::fromUtf8(text.readAllStandardOutput()).contains(expected));}
            linaccess::updateNativeFieldProperties(db,"authored_defaults","qty",{{"Format",linaccess::VbaValue::string("[Red]0.00")}});assert(model.select());
            assert(model.data(model.index(0,quantity)).toString()=="#Fehler (5)" && model.clipboardText(model.index(0,quantity))=="7");
            int error=0;try {(void)linaccess::nativeDatasheetSnapshot(grid,"Unsupported format");}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==5);
            linaccess::updateNativeFieldProperties(db,"authored_defaults","qty",{{"Format",linaccess::VbaValue::string("0.00\" units\"")}});
        }
        const auto pageBefore=linaccess::datasheet_page_store::read(db,"table","authored_defaults");linaccess::datasheet_page_store::save(db,"table","authored_defaults",pageBefore,linaccess::datasheet_page_store::defaults());
        assert(authored.exec("SELECT value FROM _la_counter32_sequences WHERE table_name='authored_defaults'") && authored.next());const auto beforeRenameCounter=authored.value(0).toInt();authored.finish();
        QTimer::singleShot(0,&window,[] {auto* input=qobject_cast<QInputDialog*>(QApplication::activeModalWidget());assert(input && input->objectName()=="native-table-rename");input->setTextValue("authored_renamed");input->accept();});window.findChild<QAction*>("native-table-rename")->trigger();
        assert(db.tables().contains("authored_renamed") && !db.tables().contains("authored_defaults"));assert(linaccess::datasheet_page_store::read(db,"table","authored_renamed").present && !linaccess::datasheet_page_store::read(db,"table","authored_defaults").present);
        assert(authored.exec("INSERT INTO authored_renamed DEFAULT VALUES"));assert(authored.exec("SELECT COUNT(*),MAX(id) FROM authored_renamed") && authored.next() && authored.value(0).toInt()==3 && authored.value(1).toInt()==beforeRenameCounter+1);authored.finish();
    }
    QTimer::singleShot(0,&window,[&] {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-form-designer");dialog->findChild<QLineEdit*>("native-form-name")->setText("Caption form");dialog->findChild<QComboBox*>("native-form-source")->setCurrentText("authored_renamed");auto* fields=dialog->findChild<QListWidget*>("native-form-fields");bool caption=false;for(int row=0;row<fields->count();++row) if(fields->item(row)->data(Qt::UserRole).toString()=="qty") {caption=fields->item(row)->text()=="Anzahl Ä";}assert(caption);dialog->findChild<QPushButton*>("native-form-save")->click();dialog->reject();});action(window,"+ Formular")->trigger();
    {const auto stored=linaccess::loadNativeStoredForm(db,"Caption form");assert(stored && stored->layout.contains(QString("Anzahl Ä").toUtf8()) && stored->layout.contains("Format ="));}
    QTimer::singleShot(0,&window,[&] {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-report-designer");dialog->findChild<QLineEdit*>("native-report-name")->setText("Caption report");dialog->findChild<QPlainTextEdit*>("native-report-sql")->setPlainText("SELECT * FROM authored_renamed");dialog->findChild<QPushButton*>("native-report-save")->click();dialog->reject();});action(window,"+ Bericht")->trigger();
    {
        const auto stored=linaccess::loadNativeStoredReport(db,"Caption report");assert(stored && stored->layout.contains(QString("Anzahl Ä").toUtf8()) && stored->layout.contains("Format ="));
        AccessFormPreview report(stored->name,stored->layout,stored->code,db,&window,false,{},std::nullopt,0,-1,stored->recordSource);
        assert(report.hasDataSource());auto controls=makePreviewFormObject(report)->member("Controls",{}).object;
        auto amount=controls->item("qty").object;assert(amount->scalarValue().kind==linaccess::VbaValue::Kind::Number && amount->scalarValue().number==7);
        assert(controls->item("moment").object->scalarValue().kind==linaccess::VbaValue::Kind::Date);
        const auto expected=QLocale().toString(7.0,'f',2)+" units";
        assert(report.findChild<QLineEdit*>("access-control-qty")->text()==expected);
        assert(report.findChild<QLineEdit*>("access-control-label")->text()=="UI VALUE" && controls->item("label").object->scalarValue().asText()=="ui value");
        QTemporaryDir output;assert(output.isValid());const auto path=output.filePath("formatted-report.pdf");assert(report.exportReportPdf(path));
        QFile pdf(path);assert(pdf.open(QIODevice::ReadOnly));const auto original=pdf.readAll();pdf.close();
        const auto reader=QStandardPaths::findExecutable("pdftotext");if(!reader.isEmpty()) {QProcess text;text.start(reader,{"-layout",path,"-"});assert(text.waitForFinished(10000) && text.exitCode()==0);assert(QString::fromUtf8(text.readAllStandardOutput()).contains(expected));}
        assert(amount->scalarValue().number==7);assert(report.close());
        const QByteArray variantSource="Begin Report\nWidth =3000\nBegin Section\nHeight =600\nBegin TextBox\nName =\"Typed\"\nFormat =\"0.00\"\nWidth =2500\nHeight =300\nEnd\nEnd\nEnd\n";
        AccessFormPreview variants("unbound formatted variants",variantSource,{},db,&window,false,{},std::nullopt,0,-1,stored->recordSource);
        auto typed=makePreviewFormObject(variants)->member("Controls",{linaccess::VbaValue::string("Typed")}).object;
        typed->assign("Value",linaccess::VbaValue::longInteger(7));assert(typed->scalarValue().numericType==linaccess::VbaValue::NumericType::Long);
        typed->assign("Format",linaccess::VbaValue::string(""));assert(typed->scalarValue().numericType==linaccess::VbaValue::NumericType::Long);
        typed->assign("Value",linaccess::VbaValue::currency(std::numeric_limits<std::int64_t>::max()));typed->assign("Format",linaccess::VbaValue::string("0.0000"));assert(typed->scalarValue().currencyUnits==std::numeric_limits<std::int64_t>::max());
        typed->assign("Format",linaccess::VbaValue::string(""));assert(typed->scalarValue().kind==linaccess::VbaValue::Kind::Currency && typed->scalarValue().currencyUnits==std::numeric_limits<std::int64_t>::max());assert(variants.close());

        auto invalid=linaccess::parseAccessFormLayout(QString::fromUtf8(stored->layout));assert(!invalid.controls.isEmpty());
        auto badSource=stored->layout;badSource.replace("Format =\"0.00\"\" units\"\"\"","Format =\"[Red]0.00\"");
        AccessFormPreview bad("Unsupported report format",badSource,{},db,&window,false,{},std::nullopt,0,-1,stored->recordSource);
        assert(bad.findChild<QLineEdit*>("access-control-qty")->property("la_format_error").toInt()==5 && !bad.exportReportPdf(path));
        assert(makePreviewFormObject(bad)->member("Controls",{linaccess::VbaValue::string("qty")}).object->scalarValue().number==7);
        assert(pdf.open(QIODevice::ReadOnly) && pdf.readAll()==original);assert(bad.close());
    }
    QTimer::singleShot(0,&window,[&] {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-query-designer");dialog->findChild<QLineEdit*>("native-query-name")->setText("Owned rename query");dialog->findChild<QPlainTextEdit*>("native-query-sql")->setPlainText("SELECT qty FROM authored_renamed");dialog->findChild<QPushButton*>("native-query-save")->click();});action(window,"+ Abfrage")->trigger();
    {const auto before=linaccess::datasheet_page_store::read(db,"query","Owned rename query");linaccess::datasheet_page_store::save(db,"query","Owned rename query",before,linaccess::datasheet_page_store::defaults());}
    {auto* tree=window.findChild<QTreeWidget*>();bool selected=false;for(auto* item:tree->findItems("Owned rename query",Qt::MatchExactly|Qt::MatchRecursive,0)) if(item->data(0,Qt::UserRole).toString()=="query") {tree->setCurrentItem(item);selected=true;break;}assert(selected);}
    QTimer::singleShot(0,&window,[] {auto* input=qobject_cast<QInputDialog*>(QApplication::activeModalWidget());assert(input && input->objectName()=="native-query-rename");input->setTextValue("Renamed query Ä");input->accept();});window.findChild<QAction*>("native-query-rename")->trigger();
    {bool owned=false;assert(linaccess::accessQuerySource(db,"Renamed query Ä",nullptr,&owned)=="SELECT qty FROM authored_renamed" && owned);assert(linaccess::datasheet_page_store::read(db,"query","Renamed query Ä").present && !linaccess::datasheet_page_store::read(db,"query","Owned rename query").present);const auto result=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT qty FROM [Renamed query Ä]"),{});assert(result.rows.size()==3);}
    {
        using namespace linaccess;auto database=makeDaoDatabase(db);auto table=database->member("CreateTableDef",{VbaValue::string("rename_command_table")}).object;auto field=table->member("CreateField",{VbaValue::string("value"),VbaValue::integer(4)}).object;table->member("Fields",{}).object->member("Append",{VbaValue::reference(field)});database->member("TableDefs",{}).object->member("Append",{VbaValue::reference(table)});
        const QByteArray source="Begin Form\nBegin Section\nEnd\nEnd\n";const QByteArray code=R"VBA(Sub RenameQuery()
DoCmd.Rename OldName:="Renamed query Ä", NewName:="Command query", ObjectType:=acQuery
End Sub
Sub RenameTable()
DoCmd.Rename "renamed_command_table", acTable, "rename_command_table"
End Sub
Sub WrongType()
DoCmd.Rename "blocked", acForm, "Caption form"
End Sub
Sub MissingType()
DoCmd.Rename NewName:="blocked", OldName:="Command query"
End Sub
Sub OmittedType()
DoCmd.Rename "blocked", , "Command query"
End Sub
Sub NullName()
DoCmd.Rename Null, acQuery, "Command query"
End Sub
Sub DeleteQuery()
DoCmd.DeleteObject ObjectName:="Command query", ObjectType:=acQuery
End Sub
Sub MissingDeleteType()
DoCmd.DeleteObject ObjectName:="Command query"
End Sub
Sub NullDeleteName()
DoCmd.DeleteObject acQuery, Null
End Sub
)VBA";
        AccessFormPreview caller("rename-command-caller",source,code,db,&window);std::vector<VbaValue> arguments;
        const auto renamed=caller.executeProcedure("RenameQuery",arguments);if(renamed.status!=VbaExecutionStatus::Completed) std::cerr<<"DoCmd.Rename error="<<renamed.errorNumber<<" "<<renamed.explanation<<"\n";assert(renamed.status==VbaExecutionStatus::Completed);
        assert(caller.executeProcedure("RenameTable",arguments).status==VbaExecutionStatus::Completed);assert(db.tables().contains("renamed_command_table") && !db.tables().contains("rename_command_table"));
        assert(caller.executeProcedure("WrongType",arguments).errorNumber==3251);assert(caller.executeProcedure("MissingType",arguments).errorNumber==3251);assert(caller.executeProcedure("OmittedType",arguments).errorNumber==3251);assert(caller.executeProcedure("NullName",arguments).errorNumber==94);
        assert(accessQuerySource(db,"Command query")=="SELECT qty FROM authored_renamed" && datasheet_page_store::read(db,"query","Command query").present && !datasheet_page_store::read(db,"query","Renamed query Ä").present);
        QSqlQuery values(db);assert(values.exec("INSERT INTO renamed_command_table VALUES(7)"));assert(values.exec("SELECT value FROM renamed_command_table") && values.next() && values.value(0).toInt()==7);values.finish();
        assert(caller.executeProcedure("MissingDeleteType",arguments).errorNumber==3251 && caller.executeProcedure("NullDeleteName",arguments).errorNumber==94);
        assert(window.openQueryDocument("Command query",0,2).value("error").toInt()==0);assert(caller.executeProcedure("DeleteQuery",arguments).errorNumber==3251);assert(window.closeQueryDocument(1,"Command query",2).value("error").toInt()==0);
        assert(caller.executeProcedure("DeleteQuery",arguments).status==VbaExecutionStatus::Completed);assert(!db.tables(QSql::Views).contains("Command query",Qt::CaseInsensitive) && !datasheet_page_store::read(db,"query","Command query").present);assert(caller.close());
    }
    {const auto before=linaccess::datasheet_page_store::read(db,"table","renamed_command_table");linaccess::datasheet_page_store::save(db,"table","renamed_command_table",before,linaccess::datasheet_page_store::defaults());auto* tree=window.findChild<QTreeWidget*>();bool selected=false;for(auto* item:tree->findItems("renamed_command_table",Qt::MatchExactly|Qt::MatchRecursive,0)) if(item->data(0,Qt::UserRole).toString()=="table") {tree->setCurrentItem(item);selected=true;break;}assert(selected);}
    QTimer::singleShot(0,&window,[] {auto* question=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(question && question->objectName()=="native-object-delete-confirm");question->done(QMessageBox::No);});window.findChild<QAction*>("native-object-delete")->trigger();assert(db.tables().contains("renamed_command_table"));
    QTimer::singleShot(0,&window,[] {auto* question=qobject_cast<QMessageBox*>(QApplication::activeModalWidget());assert(question && question->objectName()=="native-object-delete-confirm");question->done(QMessageBox::Yes);});window.findChild<QAction*>("native-object-delete")->trigger();assert(!db.tables().contains("renamed_command_table") && !linaccess::datasheet_page_store::read(db,"table","renamed_command_table").present);
    {QSqlQuery unchanged(db);assert(unchanged.exec("SELECT COUNT(*) FROM authored_renamed") && unchanged.next() && unchanged.value(0).toInt()==3);}
    {
        const QByteArray source="Begin Form\nBegin Section\nBegin TextBox\nName =\"DecimalProbe\"\nEnd\nEnd\nEnd\n";
        const QByteArray code=R"VBA(Sub DecimalControl()
Me!DecimalProbe=Application.Eval("0.1+0.2")
If VarType(Me!DecimalProbe)<>vbDecimal Then Err.Raise 5
If CStr(Me!DecimalProbe)<>"0.3" Then Err.Raise 5
Me!DecimalProbe=CDec("79228162514264337593543950335")
If VarType(Me!DecimalProbe)<>vbDecimal Then Err.Raise 5
If CStr(Me!DecimalProbe)<>"79228162514264337593543950335" Then Err.Raise 5
End Sub
)VBA";
        AccessFormPreview caller("decimal-control-caller",source,code,db,&window);std::vector<linaccess::VbaValue> arguments;const auto result=caller.executeProcedure("DecimalControl",arguments);if(result.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Decimal control error="<<result.errorNumber<<" "<<result.explanation<<"\n";assert(result.status==linaccess::VbaExecutionStatus::Completed);assert(caller.findChild<QLineEdit*>("access-control-DecimalProbe")->text()=="79228162514264337593543950335");assert(caller.close());
    }
    {
        const auto original=QLocale();QLocale::setDefault(QLocale(QLocale::German,QLocale::Austria));const auto restore=qScopeGuard([original] {QLocale::setDefault(original);});
        const QByteArray source="Begin Form\nBegin Section\nEnd\nEnd\n";const QByteArray code=R"VBA(Sub NamedNumberFormat()
Me.Tag=FormatNumber(GroupDigits:=vbFalse,Expression:=CDec("1234.5"),NumDigitsAfterDecimal:=2)
If FormatPercent(Expression:=CCur("-0.00125"),NumDigitsAfterDecimal:=2,IncludeLeadingDigit:=vbFalse,UseParensForNegativeNumbers:=vbTrue,GroupDigits:=vbFalse)<>"(,12%)" Then Err.Raise 5
If VBA.FormatPercent(CCur("0.125"),,vbUseDefault)<>"12,50%" Then Err.Raise 5
If Format(CCur("0.125"),"Percent")<>"12,50%" Then Err.Raise 5
If FormatDateTime(Date:=#12/31/2003 18:05:06#,NamedFormat:=vbLongTime)<>"18:05:06" Then Err.Raise 5
If VBA.FormatDateTime(CDate(0.75),)<>"18:00:00" Then Err.Raise 5
If FormatDateTime(#12/31/2003 18:05:06#,vbShortTime)<>"18:05" Then Err.Raise 5
If vbGeneralDate<>0 Or vbLongDate<>1 Or vbShortDate<>2 Then Err.Raise 5


End Sub
)VBA";AccessFormPreview caller("format-reference-caller",source,code,db,&window);const auto application=makePreviewApplicationObject(caller);const auto accessEval=[&](const QString& expression) {return application->member("Eval",{linaccess::VbaValue::string(expression.toStdString())});};
        QFile reference(QFileInfo(QString::fromUtf8(__FILE__)).absoluteDir().filePath("access2003_reference/expression-observations.json"));assert(reference.open(QIODevice::ReadOnly));const auto root=QJsonDocument::fromJson(reference.readAll()).object();int checked=0;
        for(const auto& entry:root.value("observations").toArray()) {
            const auto row=entry.toObject();if(row.value("verification_group").toString()!="desktop") continue;const auto expression=row.value("expression").toString();linaccess::VbaValue value;
            try {value=accessEval(expression);}catch(const linaccess::VbaRuntimeError& failure) {std::cerr<<"Desktop Eval oracle "<<row.value("case").toString().toStdString()<<": error="<<failure.number<<"\n";assert(false);}
            const auto type=accessEval("VarType("+expression+")").asNumber();if(type!=row.value("vartype").toInt()) std::cerr<<"Desktop Eval oracle "<<row.value("case").toString().toStdString()<<": type="<<type<<"\n";assert(type==row.value("vartype").toInt());const auto expected=row.value("result").toObject();
            if(type==8) {if(value.asText()!=expected.value("value").toString().toStdString()) std::cerr<<"Desktop Eval oracle "<<row.value("case").toString().toStdString()<<": string differs\n";assert(value.asText()==expected.value("value").toString().toStdString());}
            else if(type==1) assert(value.kind==linaccess::VbaValue::Kind::Null);
            else if(type==7) {auto text=expected.value("value").toString();text.replace('T',' ');const auto date=QDate::fromString(text.left(10),Qt::ISODate);const auto time=QTime::fromString(text.mid(11,8),"HH:mm:ss");assert(date.isValid() && time.isValid());const auto serial=QDate(1899,12,30).daysTo(date)+QTime(0,0).secsTo(time)/86400.0;assert(std::abs(value.number-serial)<1e-9);}
            else if(type==5) {const auto bytes=QByteArray::fromHex(expected.value("bits").toString().toLatin1());assert(bytes.size()==8);std::uint64_t bits=0;for(int i=0;i<8;++i) bits|=static_cast<std::uint64_t>(static_cast<unsigned char>(bytes[i]))<<(8*i);assert(value.asNumber()==std::bit_cast<double>(bits));}
            else assert(value.asNumber()==expected.value("value").toDouble());++checked;
        }
        assert(checked==48);std::vector<linaccess::VbaValue> arguments;const auto named=caller.executeProcedure("NamedNumberFormat",arguments);if(named.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Named format: "<<named.errorNumber<<" line="<<named.line<<" "<<named.explanation<<"\n";assert(named.status==linaccess::VbaExecutionStatus::Completed && makePreviewFormObject(caller)->member("Tag",{}).asText()=="1234,50");assert(caller.close());
    }
    {
        using linaccess::VbaValue;
        const QByteArray source=R"FORM(Begin Form
Width =6000
Begin Section
Height =1800
Begin TextBox
Name ="Amount"
ControlSource ="=CCur(""1.2345"")"
Format ="0.00"" units"""
Width =3000
Height =300
End
Begin TextBox
Name ="Invalid"
ControlSource ="=1"
Format ="[Red]0.00"
Top =400
Width =3000
Height =300
End
Begin TextBox
Name ="Writable"
DefaultValue ="1.25"
Format ="0.00"
Top =800
Width =3000
Height =300
End
End
End
)FORM";
        const QByteArray code=R"VBA(Sub ChangeMask()
If VarType(Me!Amount)<>vbCurrency Then Err.Raise 5
Me!Amount.Format="0.000"
If Me!Amount<>CCur("1.2345") Then Err.Raise 5
End Sub
)VBA";
        AccessFormPreview form("calculated format",source,code,db,&window);
        auto controls=makePreviewFormObject(form)->member("Controls",{}).object;auto amount=controls->item("Amount").object;
        auto* line=form.findChild<QLineEdit*>("access-control-Amount");
        assert(line && line->text()==QLocale().toString(1.23,'f',2)+" units" && amount->scalarValue().kind==VbaValue::Kind::Currency && amount->scalarValue().currencyUnits==12345);
        auto invalid=controls->item("Invalid").object;assert(invalid->scalarValue().asNumber()==1 && form.findChild<QLineEdit*>("access-control-Invalid")->text()=="#Fehler (5)");
        QTemporaryDir output;assert(output.isValid() && !form.exportLayoutPdf(output.filePath("invalid.pdf")));
        invalid->assign("Format",VbaValue::string("0.0"));assert(form.exportLayoutPdf(output.filePath("valid.pdf")));
        std::vector<VbaValue> arguments;assert(form.executeProcedure("ChangeMask",arguments).status==linaccess::VbaExecutionStatus::Completed);
        assert(line->text()==QLocale().toString(1.234,'f',3) && amount->scalarValue().currencyUnits==12345);
        const auto expect=[](auto action,int wanted) {int actual=0;try {action();}catch(const linaccess::VbaRuntimeError& failure) {actual=failure.number;}assert(actual==wanted);};
        expect([&] {amount->assign("Format",VbaValue::string("[Red]0.00"));},5);assert(amount->member("Format",{}).asText()=="0.000");
        expect([&] {amount->assign("Format",{});},94);controls->item("Writable").object->assign("Format",VbaValue::string("0.000"));
        assert(form.findChild<QLineEdit*>("access-control-Writable")->text()==QLocale().toString(1.25,'f',3));
        class ScalarFormatValue final:public linaccess::VbaObject {
        public:int reads{};std::function<void()> callback;VbaValue scalarValue() override {++reads;if(callback) callback();return VbaValue::currency(12345);}VbaValue member(std::string_view,const std::vector<VbaValue>&) override {throw linaccess::VbaRuntimeError(438,"No explicit member");}
        };
        const auto probe=std::make_shared<ScalarFormatValue>();auto writable=controls->item("Writable").object;writable->assign("Value",VbaValue::reference(probe));
        assert(probe->reads==1 && writable->scalarValue().kind==VbaValue::Kind::Currency && writable->scalarValue().currencyUnits==12345);

        amount->assign("Format",VbaValue::string(""));assert(line->text()=="1.2345" && amount->scalarValue().currencyUnits==12345);
        const auto closing=std::make_shared<ScalarFormatValue>();closing->callback=[&form] {assert(form.close());};expect([&] {writable->assign("Value",VbaValue::reference(closing));},2450);assert(closing->reads==1);
    }
    testEditableControlFormats(window,db);
    testNativeMacros(window,db);phase("testNativeMacros");
    testNativeMacroDesigner(window,db);phase("testNativeMacroDesigner");
    testNativeModuleEditor(window,db);phase("testNativeModuleEditor");
    testStandaloneMacroExecution(window,db);phase("testStandaloneMacroExecution");
    testMacroRepeatExpressions(window,db);phase("testMacroRepeatExpressions");
    testNativeMacroGroups(window,db);phase("testNativeMacroGroups");
    testNativeMacroStorage(window,db);phase("testNativeMacroStorage");
    testMacroProjectCommands(window,db);phase("testMacroProjectCommands");
    testMacroSelectObject(window,db);phase("testMacroSelectObject");
    testStopAllMacros(window,db);phase("testStopAllMacros");
    testMacroMessageArguments(window,db);phase("testMacroMessageArguments");
    testModuleRecordNavigation(window,db);phase("testModuleRecordNavigation");
    testModuleFindRecord(window,db);phase("testModuleFindRecord");
    testModuleUiCommands(window,db);phase("testModuleUiCommands");
    testStandaloneProjectCatalog(window,db);phase("testStandaloneProjectCatalog");
    testStandaloneDocuments(window,db);phase("testStandaloneDocuments");
    testProjectDocumentLifetime(window,db);phase("testProjectDocumentLifetime");
    testStandaloneFileDialog(window,db);phase("testStandaloneFileDialog");
    testStandaloneApplicationServices(window,db);phase("testStandaloneApplicationServices");
    testStandaloneDocumentOpening(window,db);phase("testStandaloneDocumentOpening");
    testMacroArgumentPresence(window,db);phase("testMacroArgumentPresence");
    testNativeMacroSetValue(window,db);phase("testNativeMacroSetValue");
    testNativeProjectGuiExtensions(window,db);phase("testNativeProjectGuiExtensions");
    testNativeAutoExec(window,db);phase("testNativeAutoExec");
    assert(window.close()); dispose();phase("finished"); watchdog.stop();
    std::cout<<"Main window: native project/explorer opening, independent nonmodal form/report lifetime, deduplication, modal metadata, cancellation, SQL save and retained close/MsgBox passed\n";
}
