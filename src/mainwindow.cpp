#include "native_schema_editor.hpp"
#include "native_field_defaults.hpp"
#include "mainwindow.hpp"
#include "native_macro.hpp"
#include "native_macro_editor_draft.hpp"
#include "native_macro_storage.hpp"
#include "native_module_definition.hpp"
#include "native_standard_module_host.hpp"
#include "native_form_event_editor.hpp"
#include "native_grid_clipboard.hpp"
#include "native_datasheet_print.hpp"
#include "native_datasheet_page_store.hpp"
#include "native_report_page_setup.hpp"
#include "native_query_view.hpp"
#include "native_input_box.hpp"
#include "native_index_designer.hpp"
#include "native_field_designer.hpp"
#include "native_relation_designer.hpp"
#include "native_edit_menu.hpp"
#include "native_classic_style.hpp"
#include "native_options_dialog.hpp"
#include <QCloseEvent>
#include <QKeyEvent>
#include <QShortcut>
#include <QScopeGuard>
#include <QScopedValueRollback>

#include "access_inspector.hpp"
#include "access_table_import.hpp"
#include "access_autonumber.hpp"
#include "access_form_import.hpp"
#include "access_form_preview.hpp"
#include "native_stored_form.hpp"
#include "native_stored_report.hpp"
#include "vba_module_context.hpp"
#include "access_form_objects.hpp"
#include "bridge_importer.hpp"
#include "vba_analyzer.hpp"
#include "vba_runtime.hpp"
#include "dao_database.hpp"
#include "dao_schema_change.hpp"
#include "vba_qt_value.hpp"
#include "access_query.hpp"
#include "jet_sql_functions.hpp"
#include "access_table_model.hpp"
#include "access_table_delegate.hpp"
#include "access_table_view.hpp"
#include "access_format.hpp"
#include "vba_source_codec.hpp"

#include <cmath>
#include <QAction>
#include <QApplication>
#include <QCheckBox>
#include <QComboBox>
#include <QDialog>
#include <QDialogButtonBox>
#include <QFileDialog>
#include <QFile>
#include <QFileInfo>
#include <QFormLayout>
#include <QSpinBox>
#include <QSignalBlocker>
#include <QHeaderView>
#include <QHBoxLayout>
#include <QInputDialog>
#include <QJsonArray>
#include <QJsonDocument>
#include <QLabel>
#include <QLineEdit>
#include <QClipboard>
#include <QMimeData>
#include <QListWidget>
#include <QMap>
#include <QMessageBox>
#include <QPlainTextEdit>
#include <QPrinter>
#include <QPrintDialog>
#include <QPrintPreviewWidget>
#include <QPushButton>
#include <QSqlError>
#include <QSqlQuery>
#include <QSqlQueryModel>
#include <QSqlRecord>
#include <QSqlIndex>
#include <QSqlTableModel>
#include <QStatusBar>
#include <QSplitter>
#include <QTabWidget>
#include <QTabBar>
#include <QMouseEvent>
#include <QTableView>
#include <QTableWidget>
#include <QTextBrowser>
#include <QTextDocument>
#include <QToolBar>
#include <QMenuBar>
#include <QMenu>
#include <QTreeWidget>
#include <QTreeWidgetItem>
#include <QVBoxLayout>

#include <set>
#include <memory>
#include <functional>
#include <optional>
#include <stdexcept>

namespace {
void tagProjectObject(QWidget* widget,QObject* owner,const QSqlDatabase& database,const QString& kind,const QString& name) {
    if(name.isEmpty()) return;
    widget->setProperty("la_project_object_kind",kind);
    widget->setProperty("la_project_object_name",name);
    widget->setProperty("la_project_object_connection",database.connectionName());
    widget->setProperty("la_project_object_database",database.databaseName());
    widget->setProperty("la_project_object_generation",owner->property("la_project_generation"));
}
QString quoteSql(QString name) {
    name.replace('"', "\"\"");
    return '"' + name + '"';
}
QVariantMap tableCommandFailure(int number,const QString& description) {
    return {{"handled",true},{"error",number},{"description",description}};
}
QString editableSourceSchema(QSqlDatabase db,const QString& table) {
    QSqlQuery schema(db);
    schema.prepare("SELECT sql FROM main.sqlite_schema WHERE type='table' AND name=? COLLATE NOCASE");
    schema.addBindValue(table);
    if(!schema.exec() || !schema.next()) throw linaccess::VbaRuntimeError(3251,"Die Quelltabelle ist nicht mehr verfügbar.");
    return schema.value(0).toString();
}
QVariantMap validateEditableQuerySource(QWidget* page,linaccess::AccessTableModel* model) {
    if(!page || !page->property("la_editable_query").toBool()) return {};
    try {
        const auto table=linaccess::editableAccessQueryTable(model->database(),page->property("la_editable_query_source").toString());
        if(table!=model->tableName() || editableSourceSchema(model->database(),table)!=page->property("la_editable_query_schema").toString())
            return tableCommandFailure(3251,"Die Quellstruktur wurde geändert. Abfrage schließen und erneut öffnen.");
    } catch(const linaccess::VbaRuntimeError&) {
        return tableCommandFailure(3251,"Die Quellstruktur wurde geändert. Abfrage schließen und erneut öffnen.");
    }
    return {};
}
QVariantMap focusedDataTextCommand(QWidget* scope,int command) {
    if(!scope || (command!=189 && command!=190 && command!=191 && command!=199 && command!=292)) return {};
    QPointer<QLineEdit> editor=qobject_cast<QLineEdit*>(QApplication::focusWidget());
    QWidget* parent=editor.data();
    while(parent && parent!=scope) parent=parent->parentWidget();
    if(editor && parent) {
        if(!editor->isVisible() || !editor->isEnabled()) return tableCommandFailure(2046,"Kein verfügbares aktives Textfeld.");
        if(command==199 || command==292) {
            if(editor->isReadOnly()) return tableCommandFailure(2448,"Das aktive Textfeld ist schreibgeschützt.");
            if(!(command==292 ? editor->isUndoAvailable() : editor->isRedoAvailable())) return tableCommandFailure(2046,"Kein entsprechender Textbearbeitungsschritt verfügbar.");
            if(command==292) editor->undo(); else editor->redo();
            return {{"handled",true},{"error",0}};
        }
        if(command==189 || command==190) {
            if(editor->echoMode()!=QLineEdit::Normal || !editor->hasSelectedText()) return tableCommandFailure(2046,"Kein kopierbarer Text ausgewählt.");
            if(command==189 && editor->isReadOnly()) return tableCommandFailure(2448,"Das aktive Textfeld ist schreibgeschützt.");
            const auto selected=editor->selectedText();
            if(selected.size()>4*1024*1024 || selected.contains(QChar(0))) return tableCommandFailure(3251,"Dieser Zwischenablageumfang wird noch nicht unterstützt.");
            if(command==189) editor->cut(); else editor->copy();
            return {{"handled",true},{"error",0}};
        }
        if(editor->isReadOnly()) return tableCommandFailure(2448,"Das aktive Textfeld ist schreibgeschützt.");
        const auto* mime=QApplication::clipboard()->mimeData();
        if(!mime || !mime->hasText()) return tableCommandFailure(2046,"Die Zwischenablage enthält keinen einfügbaren Text.");
        const auto text=QApplication::clipboard()->text();
        if(text.size()>4*1024*1024 || text.contains(QChar(0))) return tableCommandFailure(3251,"Dieser Zwischenablageumfang wird noch nicht unterstützt.");
        editor->paste();
        return {{"handled",true},{"error",0}};
    }
    return {};
}

QVariantMap tableModelFailure(linaccess::AccessTableModel* model,int fallback) {
    if(!model) return tableCommandFailure(2501,"Tabellendokument wurde während des Befehls geschlossen.");
    const auto problem=model->lastError();
    bool valid=false; const int native=problem.nativeErrorCode().toInt(&valid);
    int number=problem.driverText()=="LinAccess" && valid && native>0 ? native : fallback;
    if(const auto function=linaccess::jetSqlFunctionErrorNumber(problem.text().toLower().toStdString())) number=*function;
    return tableCommandFailure(number,problem.isValid() ? problem.text() : QStringLiteral("Tabellenbefehl abgebrochen."));
}
void messageError(QWidget* parent, const QString& explanation) {
    QMessageBox::critical(parent, "LinAccess — Fehler", explanation);
}
QPushButton* addButton(QBoxLayout* layout, const QString& name) {
    auto* button = new QPushButton(name);
    layout->addWidget(button);
    return button;
}
void installDesignerSaveShortcut(QDialog& dialog,QPushButton* save) {
    auto* status=new QLabel(&dialog); status->setObjectName("native-designer-save-status");
    dialog.layout()->addWidget(status);
    auto* shortcut=new QShortcut(QKeySequence::Save,&dialog); shortcut->setContext(Qt::WidgetWithChildrenShortcut);
    const QPointer<QDialog> guarded=&dialog; const QPointer<QPushButton> button=save;
    QObject::connect(shortcut,&QShortcut::activated,&dialog,[guarded,button] {
        if(!guarded || !button) return;
        const auto previous=guarded->property("la_editor_keep_open"); guarded->setProperty("la_editor_keep_open",true);
        const auto reset=qScopeGuard([guarded,previous] { if(guarded) guarded->setProperty("la_editor_keep_open",previous); });
        button->click();
    });
}
void trackDesignerChanges(QDialog& dialog) {
    dialog.setProperty("la_designer_state_available",true);
    const bool fresh=dialog.property("la_project_object_name").toString().isEmpty();
    dialog.setProperty("la_designer_new",fresh);
    if(!dialog.windowTitle().contains("[*]")) dialog.setWindowTitle(dialog.windowTitle()+" [*]");
    dialog.setWindowModified(fresh);
    const QPointer<QDialog> guarded=&dialog;
    const auto changed=[guarded] {
        if(!guarded) return;
        guarded->setWindowModified(true);
        if(auto* status=guarded->findChild<QLabel*>("native-designer-save-status")) status->setText("Nicht gespeicherte Änderungen");
    };
    for(auto* editor:dialog.findChildren<QLineEdit*>()) QObject::connect(editor,&QLineEdit::textChanged,&dialog,changed);
    for(auto* editor:dialog.findChildren<QPlainTextEdit*>()) QObject::connect(editor,&QPlainTextEdit::textChanged,&dialog,changed);
    for(auto* editor:dialog.findChildren<QComboBox*>()) if(!editor->property("la_designer_transient").toBool()) QObject::connect(editor,&QComboBox::currentTextChanged,&dialog,changed);
    for(auto* editor:dialog.findChildren<QCheckBox*>()) if(!editor->property("la_designer_transient").toBool()) QObject::connect(editor,&QCheckBox::toggled,&dialog,changed);
    for(auto* editor:dialog.findChildren<QSpinBox*>()) QObject::connect(editor,&QSpinBox::valueChanged,&dialog,changed);
    const auto trackModel=[&](QAbstractItemModel* model) {
        QObject::connect(model,&QAbstractItemModel::dataChanged,&dialog,changed);
        QObject::connect(model,&QAbstractItemModel::rowsInserted,&dialog,changed);
        QObject::connect(model,&QAbstractItemModel::rowsRemoved,&dialog,changed);
        QObject::connect(model,&QAbstractItemModel::rowsMoved,&dialog,changed);
        QObject::connect(model,&QAbstractItemModel::modelReset,&dialog,changed);
    };
    for(auto* list:dialog.findChildren<QListWidget*>()) trackModel(list->model());
    for(auto* table:dialog.findChildren<QTableWidget*>()) trackModel(table->model());
}
void completeDesignerSave(QDialog& dialog,const QString& name) {
    dialog.setProperty("la_editor_saved_name",name);
    dialog.setProperty("la_designer_new",false); dialog.setWindowModified(false);
    if(auto* status=dialog.findChild<QLabel*>("native-designer-save-status")) status->setText("Gespeichert: "+name);
    if(!dialog.property("la_editor_keep_open").toBool()) dialog.accept();
}
QString asText(const QVariant& value) {
    return value.isNull() ? QString{} : value.toString();
}
bool isSelect(const QString& sql) {
    return sql.trimmed().startsWith("SELECT ", Qt::CaseInsensitive)
        || sql.trimmed().startsWith("SELECT\n", Qt::CaseInsensitive);
}


}

MainWindow::~MainWindow() { disconnect(tabs_,nullptr,this,nullptr); linaccess::resetVbaProjectState(db_); }

MainWindow::MainWindow(QWidget* parent) : QMainWindow(parent) {
    setWindowTitle("LinAccess");
    resize(1260, 790);
    db_ = QSqlDatabase::addDatabase("QSQLITE", "linaccess-main");

    auto* splitter = new QSplitter(this);
    explorer_ = new QTreeWidget(splitter);
    explorer_->setObjectName("database-explorer");
    explorer_->installEventFilter(this);
    explorer_->setHeaderHidden(true);
    explorer_->setMinimumWidth(220);
    tabs_ = new QTabWidget(splitter);
    tabs_->setTabsClosable(true);
    tabs_->installEventFilter(this); tabs_->tabBar()->installEventFilter(this);
    connect(tabs_,&QTabWidget::currentChanged,this,[this] {
        auto* focus=QApplication::focusWidget();
        if(!tabs_->currentWidget()) activeDataPage_.clear();
        else if(focus && (focus==tabs_ || tabs_->isAncestorOf(focus))) activeDataPage_=tabs_->currentWidget();
    });
    splitter->addWidget(explorer_);
    splitter->addWidget(tabs_);
    splitter->setStretchFactor(1, 1);
    setCentralWidget(splitter);
    connect(tabs_, &QTabWidget::tabCloseRequested, this, [this](int index) {
        QWidget* page = tabs_->widget(index);
        if(!canCloseDataPage(page)) return;
        tabs_->removeTab(index); page->hide();
        for(auto* model:page->findChildren<QSqlTableModel*>()) model->clear();
        page->deleteLater();
    });
    connect(explorer_, &QTreeWidget::itemDoubleClicked,
            this, &MainWindow::openExplorerItem);

    auto* toolbar = addToolBar("Projekt");
    auto action = [this, toolbar](const QString& label, auto callback) {
        auto* item = toolbar->addAction(label);
        connect(item, &QAction::triggered, this, callback);
    };
    action("Neu", [this] { newDatabase(); });
    action("Öffnen", [this] { chooseDatabase(); });
    action("Access untersuchen", [this] { inspectAccess(); });
    action("Bridge-Export öffnen", [this] { inspectBridgeExport(); });
    toolbar->addSeparator();
    action("+ Tabelle", [this] { newTable(); });
    action("+ Abfrage", [this] { editQuery(); });
    action("+ Formular", [this] { editForm(); });
    action("+ Bericht", [this] { editReport(); });
    action("+ Makro", [this] { editMacro(); });
    action("+ Modul", [this] { editModule(); });
    action("Aktualisieren", [this] { refreshExplorer(); });
    auto* fileMenu=menuBar()->addMenu("&Datei");
    const auto projectActions=toolbar->actions();
    projectActions[0]->setObjectName("project-new"); projectActions[0]->setShortcut(QKeySequence::New);
    projectActions[1]->setObjectName("project-open"); projectActions[1]->setShortcut(QKeySequence::Open);
    for(int index=0;index<2;++index) projectActions[index]->setShortcutContext(Qt::WindowShortcut);
    for (int i=0;i<4;++i) fileMenu->addAction(projectActions[i]);
    fileMenu->addSeparator();
    auto* printData=fileMenu->addAction("Datenblatt &drucken...\tStrg+P"); printData->setObjectName("datasheet-print");
    auto* pdfData=fileMenu->addAction("Datenblatt als PDF..."); pdfData->setObjectName("datasheet-pdf");
    auto* previewData=fileMenu->addAction("Datenblatt-Seitenansicht..."); previewData->setObjectName("datasheet-preview");
    auto* setupData=fileMenu->addAction("Datenblatt: Seite einrichten..."); setupData->setObjectName("datasheet-page-setup");
    connect(setupData,&QAction::triggered,this,[this] {
        const QPointer<MainWindow> self=this; const auto result=setupDatasheetPage();
        if(self) if(const int error=result.value("error").toInt()) self->statusBar()->showMessage(result.value("description").toString()+" ("+QString::number(error)+")",10000);
    });
    connect(previewData,&QAction::triggered,this,[this] {
        const QPointer<MainWindow> self=this; const auto result=previewDatasheet();
        if(self) if(const int error=result.value("error").toInt()) self->statusBar()->showMessage(result.value("description").toString()+" ("+QString::number(error)+")",10000);
    });
    const auto outputData=[this](bool pdf) {
        const QPointer<MainWindow> self=this;
        const auto result=outputDatasheet(pdf);
        if(self) if(const int error=result.value("error").toInt()) self->statusBar()->showMessage(result.value("description").toString()+" ("+QString::number(error)+")",10000);
    };
    connect(printData,&QAction::triggered,this,[outputData] { outputData(false); });
    connect(pdfData,&QAction::triggered,this,[outputData] { outputData(true); });
    connect(fileMenu,&QMenu::aboutToShow,this,[this,printData,pdfData,previewData,setupData] {
        const bool enabled=activeDataPage_ && tabs_->currentWidget()==activeDataPage_ && !activePreviewDocument()
            && AccessFormPreview::databaseWindowActive(this) && !QApplication::activeModalWidget() && !datasheetOutputActive_;
        printData->setEnabled(enabled); pdfData->setEnabled(enabled); previewData->setEnabled(enabled); setupData->setEnabled(enabled);
    });
    fileMenu->addAction("Beenden",this,[this] { close(); });
    auto* editMenu=linaccess::addNativeEditMenu(menuBar(),this);
    editMenu->addSeparator();
    auto* findTable=editMenu->addAction("&Suchen...\tStrg+F"); findTable->setObjectName("table-find"); findTable->setEnabled(false);
    auto* replaceTable=editMenu->addAction("&Ersetzen...\tStrg+H"); replaceTable->setObjectName("table-replace"); replaceTable->setEnabled(false);
    connect(replaceTable,&QAction::triggered,this,[this] { (void)dispatchTableCommand(29); });
    auto* findNext=editMenu->addAction("&Weitersuchen\tUmschalt+F4"); findNext->setObjectName("table-find-next"); findNext->setEnabled(false);
    connect(findTable,&QAction::triggered,this,[this] { (void)dispatchTableCommand(30); });
    connect(findNext,&QAction::triggered,this,[this] { (void)dispatchTableCommand(341); });
    connect(editMenu,&QMenu::aboutToShow,this,[this,findTable,findNext,replaceTable] {
        linaccess::AccessTableView* view=nullptr;
        if(!activePreviewDocument() && AccessFormPreview::databaseWindowActive(this) && activeDataPage_ && tabs_->currentWidget()==activeDataPage_)
            for(auto* candidate:activeDataPage_->findChildren<QTableView*>()) if((view=dynamic_cast<linaccess::AccessTableView*>(candidate))) break;
        const auto* model=view ? qobject_cast<linaccess::AccessTableModel*>(view->model()) : nullptr;
        bool textColumn=false;
        if(model) for(int column=0;column<model->columnCount();++column) if(!view->isColumnHidden(column) && model->isSearchableColumn(column)) { textColumn=true; break; }
        bool replaceColumn=false;
        if(model && !model->isReadOnly()) for(int column=0;column<model->columnCount();++column)
            if(!view->isColumnHidden(column) && model->isTextColumn(column)) { replaceColumn=true; break; }
        replaceTable->setEnabled(replaceColumn && !QApplication::activeModalWidget());
        findTable->setEnabled(textColumn && !QApplication::activeModalWidget());
        findNext->setEnabled(view && view->hasFindText() && !QApplication::activeModalWidget());
    });
    auto* objectMenu=menuBar()->addMenu("&Einfügen");
    auto* viewMenu=menuBar()->addMenu("&Ansicht");
    menuBar()->removeAction(viewMenu->menuAction()); menuBar()->insertMenu(objectMenu->menuAction(),viewMenu);
    for (int i=5;i<11;++i) objectMenu->addAction(projectActions[i]);
    auto* renameTable=viewMenu->addAction("Tabelle &umbenennen...");renameTable->setObjectName("native-table-rename");
    connect(renameTable,&QAction::triggered,this,[this] {
        const auto* item=explorer_->currentItem();if(!item || item->data(0,Qt::UserRole).toString()!="table" || QApplication::activeModalWidget() || !db_.isOpen()) return;const auto oldName=item->text(0);
        for(auto* model:findChildren<linaccess::AccessTableModel*>()) if(model->tableName().compare(oldName,Qt::CaseInsensitive)==0) {messageError(this,"Schließen Sie die Tabelle und zugehörige bearbeitbare Abfragen vor dem Umbenennen.");return;}
        QInputDialog input(this);input.setObjectName("native-table-rename");input.setWindowTitle("Tabelle umbenennen");input.setLabelText("Neuer Tabellenname. Gespeicherte Access-Abfragen und Formularbindungen müssen anschließend angepasst werden.");input.setTextValue(oldName);if(auto* editor=input.findChild<QLineEdit*>()) editor->setMaxLength(64);if(input.exec()!=QDialog::Accepted) return;const auto newName=input.textValue().trimmed();if(newName==oldName) return;
        try {linaccess::makeDaoDatabase(db_)->member("TableDefs",{linaccess::VbaValue::string(oldName.toStdString())}).object->assign("Name",linaccess::VbaValue::string(newName.toStdString()));}
        catch(const linaccess::DaoStoredSchemaChange& error) {refreshExplorer();messageError(this,"Tabelle wurde umbenannt. Bitte das Projekt neu öffnen.\n"+QString::fromUtf8(error.what()));return;}catch(const std::exception& error) {messageError(this,QString::fromUtf8(error.what()));return;}
        refreshExplorer();for(auto* entry:explorer_->findItems(newName,Qt::MatchExactly|Qt::MatchRecursive,0)) if(entry->data(0,Qt::UserRole).toString()=="table") {explorer_->setCurrentItem(entry);break;}statusBar()->showMessage("Tabelle umbenannt: "+newName);
    });
    connect(viewMenu,&QMenu::aboutToShow,this,[this,renameTable] {const auto* item=explorer_->currentItem();renameTable->setEnabled(db_.isOpen() && item && item->data(0,Qt::UserRole).toString()=="table" && !QApplication::activeModalWidget());});
    auto* renameQuery=viewMenu->addAction("Abfrage umbenennen...");renameQuery->setObjectName("native-query-rename");
    connect(renameQuery,&QAction::triggered,this,[this] {
        const auto* item=explorer_->currentItem();if(!item || item->data(0,Qt::UserRole).toString()!="query" || QApplication::activeModalWidget() || !db_.isOpen()) return;const auto oldName=item->text(0);
        if(findOpenQuery(oldName)) {messageError(this,"Schließen Sie die Abfrage vor dem Umbenennen.");return;}
        try {
            bool owned=false;(void)linaccess::accessQuerySource(db_,oldName,nullptr,&owned,true);if(!owned) throw linaccess::VbaRuntimeError(3251,"Importierte Abfragen können noch nicht umbenannt werden.");
            const auto definition=linaccess::makeDaoDatabase(db_)->member("QueryDefs",{linaccess::VbaValue::string(oldName.toStdString())}).object;
            QInputDialog input(this);input.setObjectName("native-query-rename");input.setWindowTitle("Abfrage umbenennen");input.setLabelText("Neuer Abfragename. Verweise in anderen Abfragen, Formularen, Berichten und VBA müssen anschließend angepasst werden.");input.setTextValue(oldName);if(auto* editor=input.findChild<QLineEdit*>()) editor->setMaxLength(64);if(input.exec()!=QDialog::Accepted) return;const auto newName=input.textValue().trimmed();if(newName==oldName) return;
            definition->assign("Name",linaccess::VbaValue::string(newName.toStdString()));refreshExplorer();for(auto* entry:explorer_->findItems(newName,Qt::MatchExactly|Qt::MatchRecursive,0)) if(entry->data(0,Qt::UserRole).toString()=="query") {explorer_->setCurrentItem(entry);break;}statusBar()->showMessage("Abfrage umbenannt: "+newName);
        }catch(const std::exception& error) {messageError(this,QString::fromUtf8(error.what()));}
    });
    connect(viewMenu,&QMenu::aboutToShow,this,[this,renameQuery] {const auto* item=explorer_->currentItem();bool owned=false;if(db_.isOpen() && item && item->data(0,Qt::UserRole).toString()=="query" && !QApplication::activeModalWidget()) try {(void)linaccess::accessQuerySource(db_,item->text(0),nullptr,&owned,true);}catch(const std::exception&) {}renameQuery->setEnabled(owned);});
    auto* deleteObject=viewMenu->addAction("Ausgewähltes Objekt löschen...");deleteObject->setObjectName("native-object-delete");
    const auto selectedOwned=[this] {const auto* item=explorer_->currentItem();if(!db_.isOpen() || !item || QApplication::activeModalWidget()) return false;const auto kind=item->data(0,Qt::UserRole).toString();try {if(kind=="macro") return linaccess::readNativeMacroRevision(db_,item->text(0)).has_value();if(kind=="query") {bool owned=false;(void)linaccess::accessQuerySource(db_,item->text(0),nullptr,&owned,true);return owned;}if(kind=="table") {const auto fields=linaccess::nativeTableFieldDisplayProperties(db_,item->text(0));return !fields.isEmpty() && std::all_of(fields.cbegin(),fields.cend(),[](const auto& field) {return field.nativeOwned;});}}catch(const std::exception&) {}return false;};
    connect(deleteObject,&QAction::triggered,this,[this,selectedOwned] {
        if(!selectedOwned()) return;const auto* item=explorer_->currentItem();const auto kind=item->data(0,Qt::UserRole).toString(),name=item->text(0);
        if(kind=="macro") {
            try {const auto revision=linaccess::readNativeMacroRevision(db_,name);if(!revision) throw linaccess::VbaRuntimeError(2485,"Eigenes Makro ist nicht verfügbar.");const auto identity=db_.databaseName();const auto generation=property("la_project_generation");
                QMessageBox question(QMessageBox::Question,"Makroobjekt löschen","Makroobjekt „"+name+"“ mit allen Teilmakros löschen? Gespeicherte Verweise werden nicht automatisch angepasst.",QMessageBox::Yes|QMessageBox::No,this);question.setObjectName("native-object-delete-confirm");question.setDefaultButton(QMessageBox::No);if(question.exec()!=QMessageBox::Yes) return;
                if(!db_.isOpen() || db_.databaseName()!=identity || property("la_project_generation")!=generation) throw linaccess::VbaRuntimeError(3420,"Das Projekt hat sich geändert.");linaccess::deleteNativeMacro(db_,*revision);refreshExplorer();statusBar()->showMessage("Makro gelöscht: "+name);
            }catch(const std::exception& error) {messageError(this,QString::fromUtf8(error.what()));}return;
        }
        const int type=kind=="table" ? 0 : 1;
        try {const auto held=linaccess::makeDaoDatabase(db_)->member(type==0 ? "TableDefs" : "QueryDefs",{linaccess::VbaValue::string(name.toStdString())}).object;
            QMessageBox question(QMessageBox::Question,"Datenbankobjekt löschen","Objekt „"+name+"“ und seine enthaltenen Daten löschen? Gespeicherte Verweise werden nicht automatisch angepasst.",QMessageBox::Yes|QMessageBox::No,this);question.setObjectName("native-object-delete-confirm");question.setDefaultButton(QMessageBox::No);if(question.exec()!=QMessageBox::Yes) return;
            if(QString::fromStdString(held->member("Name",{}).asText())!=name) throw linaccess::VbaRuntimeError(3420,"Das ausgewählte Objekt hat sich geändert.");const auto result=deleteDatabaseObject(type,name);if(const int error=result.value("error").toInt()) throw linaccess::VbaRuntimeError(error,result.value("description").toString().toStdString());statusBar()->showMessage("Objekt gelöscht: "+name);
        }catch(const std::exception& error) {messageError(this,QString::fromUtf8(error.what()));}
    });
    connect(viewMenu,&QMenu::aboutToShow,this,[deleteObject,selectedOwned] {deleteObject->setEnabled(selectedOwned());});
    auto* runMacro=viewMenu->addAction("Ausgewähltes Makro ausführen");runMacro->setObjectName("native-macro-run");
    const auto canRunMacro=[this] {const auto* item=explorer_->currentItem();return db_.isOpen() && !QApplication::activeModalWidget() && queryExecutionDepth_==0 && macroExecutionDepth_==0 && !linaccess::nativeProjectMacrosActive(db_) && item && item->data(0,Qt::UserRole).toString()=="macro";};
    const auto updateRunMacro=[runMacro,canRunMacro] {runMacro->setEnabled(canRunMacro());};updateRunMacro();
    connect(viewMenu,&QMenu::aboutToShow,this,updateRunMacro);connect(explorer_,&QTreeWidget::currentItemChanged,this,updateRunMacro);
    connect(runMacro,&QAction::triggered,this,[this,canRunMacro] {
        if(!canRunMacro()) return;const auto name=explorer_->currentItem()->text(0);const auto result=runProjectMacro(name);
        if(const int error=result.value("error").toInt()) messageError(this,QString("Fehler %1: %2").arg(error).arg(result.value("description").toString()));
        else statusBar()->showMessage(result.value("stopped").toBool() ? "Makroausführung beendet." : "Makro ausgeführt: "+name);
    });
    auto* editRelations=viewMenu->addAction("&Beziehungen bearbeiten...");editRelations->setObjectName("native-relations-edit");
    connect(editRelations,&QAction::triggered,this,[this] {
        if(!db_.isOpen() || QApplication::activeModalWidget()) return;
        if(!findChildren<linaccess::AccessTableModel*>().isEmpty()) {messageError(this,"Schließen Sie Tabellen und bearbeitbare Abfragen vor Änderungen an Beziehungen.");return;}
        const auto* item=explorer_->currentItem();linaccess::editNativeRelations(this,db_,item && item->data(0,Qt::UserRole).toString()=="table" ? item->text(0) : QString{});
    });
    connect(viewMenu,&QMenu::aboutToShow,this,[this,editRelations] {editRelations->setEnabled(db_.isOpen() && !QApplication::activeModalWidget());});
    auto* editFields=viewMenu->addAction("Tabellen&felder bearbeiten...");editFields->setObjectName("native-table-fields-edit");
    connect(editFields,&QAction::triggered,this,[this] {
        const auto* item=explorer_->currentItem();if(!item || item->data(0,Qt::UserRole).toString()!="table" || QApplication::activeModalWidget() || !db_.isOpen()) return;
        const auto table=item->text(0);
        for(auto* model:findChildren<linaccess::AccessTableModel*>()) if(model->tableName().compare(table,Qt::CaseInsensitive)==0) {messageError(this,"Schließen Sie die Tabelle und zugehörige bearbeitbare Abfragen vor der Strukturänderung.");return;}
        linaccess::editNativeFields(this,db_,table);
    });
    connect(viewMenu,&QMenu::aboutToShow,this,[this,editFields] {const auto* item=explorer_->currentItem();editFields->setEnabled(db_.isOpen() && item && item->data(0,Qt::UserRole).toString()=="table" && !QApplication::activeModalWidget());});
    auto* editIndexes=viewMenu->addAction("Tabellen&indizes bearbeiten...");editIndexes->setObjectName("native-table-indexes");
    connect(editIndexes,&QAction::triggered,this,[this] {
        const auto* item=explorer_->currentItem();if(!item || item->data(0,Qt::UserRole).toString()!="table" || QApplication::activeModalWidget() || !db_.isOpen()) return;
        const auto table=item->text(0);
        for(auto* model:findChildren<linaccess::AccessTableModel*>()) if(model->tableName().compare(table,Qt::CaseInsensitive)==0) {messageError(this,"Schließen Sie die Tabelle und zugehörige bearbeitbare Abfragen vor Änderungen an Indizes oder Primärschlüsseln.");return;}
        linaccess::editNativeIndexes(this,db_,table);
    });
    connect(viewMenu,&QMenu::aboutToShow,this,[this,editIndexes] {
        const auto* item=explorer_->currentItem();editIndexes->setEnabled(db_.isOpen() && item && item->data(0,Qt::UserRole).toString()=="table" && !QApplication::activeModalWidget());
    });
    auto* editStoredForm=viewMenu->addAction("Formular&entwurf bearbeiten..."); editStoredForm->setObjectName("native-form-edit");
    connect(editStoredForm,&QAction::triggered,this,[this] {
        const auto* item=explorer_->currentItem(); if(item && item->data(0,Qt::UserRole).toString()=="form") editForm(item->text(0));
    });
    connect(viewMenu,&QMenu::aboutToShow,this,[this,editStoredForm] {
        const auto* item=explorer_->currentItem(); editStoredForm->setEnabled(db_.isOpen() && item && item->data(0,Qt::UserRole).toString()=="form" && !QApplication::activeModalWidget());
    });
    auto* editStoredReport=viewMenu->addAction("Berichtsentwurf bearbeiten..."); editStoredReport->setObjectName("native-report-edit");
    connect(editStoredReport,&QAction::triggered,this,[this] {
        const auto* item=explorer_->currentItem(); if(item && item->data(0,Qt::UserRole).toString()=="report") editReport(item->text(0));
    });
    connect(viewMenu,&QMenu::aboutToShow,this,[this,editStoredReport] {
        const auto* item=explorer_->currentItem(); editStoredReport->setEnabled(db_.isOpen() && item && item->data(0,Qt::UserRole).toString()=="report" && !QApplication::activeModalWidget());
    });
    viewMenu->addSeparator();
    viewMenu->addAction(projectActions.last());
    viewMenu->addSeparator();
    auto* classicStyle=viewMenu->addAction("&Klassische Oberfläche");
    classicStyle->setObjectName("classic-interface"); classicStyle->setCheckable(true);
    classicStyle->setChecked(qApp->property("la_classic_style").toBool());
    classicStyle->setEnabled(linaccess::classicStyleAvailable());
    connect(classicStyle,&QAction::toggled,this,[classicStyle](bool enabled) {
        if (linaccess::setClassicStyle(enabled)) QSettings().setValue("ui/classicStyle",enabled);
        else {
            const QSignalBlocker blocked(classicStyle);
            classicStyle->setChecked(qApp->property("la_classic_style").toBool());
        }
    });
    auto* toolsMenu=menuBar()->addMenu("E&xtras");
    auto* options=toolsMenu->addAction("&Optionen..."); options->setObjectName("application-options");
    connect(options,&QAction::triggered,this,[this,classicStyle] { linaccess::showNativeOptions(this,classicStyle); });
    auto* windowMenu=menuBar()->addMenu("&Fenster");
    windowMenu->setObjectName("window-menu");
    auto* nextWindow=windowMenu->addAction("Nächstes Fenster\tStrg+F6");
    auto* previousWindow=windowMenu->addAction("Vorheriges Fenster\tStrg+Umschalt+F6");
    // The existing key handlers retain form KeyPreview/cancellation semantics.
    connect(nextWindow,&QAction::triggered,this,[this] { AccessFormPreview::cycleProjectDocuments(db_,this,false,activePreviewDocument()); });
    connect(previousWindow,&QAction::triggered,this,[this] { AccessFormPreview::cycleProjectDocuments(db_,this,true,activePreviewDocument()); });
    auto* databaseWindow=windowMenu->addAction("Datenbankfenster\tF11");
    connect(databaseWindow,&QAction::triggered,this,[this] { (void)AccessFormPreview::activateProjectDocument(db_,this,nullptr); });
    windowMenu->addSeparator();
    auto* closeAll=windowMenu->addAction("Alle schließen");
    connect(closeAll,&QAction::triggered,this,[this] {
        if(QApplication::activeModalWidget() || !AccessFormPreview::closeProjectDocuments(db_)) return;
        while(tabs_->count()) {
            const auto count=tabs_->count(); emit tabs_->tabCloseRequested(count-1);
            if(tabs_->count()==count) break;
        }
    });
    connect(windowMenu,&QMenu::aboutToShow,this,[this,windowMenu] {
        for(auto* item:windowMenu->actions()) if(item->property("la_document_entry").toBool()) { windowMenu->removeAction(item); delete item; }
        const auto documents=AccessFormPreview::projectDocuments(db_);
        if(documents.isEmpty() && tabs_->count()==0) return;
        windowMenu->addSeparator()->setProperty("la_document_entry",true);
        int number=0;
        for(int index=0;index<tabs_->count();++index) {
            const QPointer<QWidget> page=tabs_->widget(index);
            auto title=tabs_->tabText(index); title.replace('&',"&&"); ++number;
            auto* item=windowMenu->addAction((number<10 ? "&" : "")+QString::number(number)+" "+title);
            item->setProperty("la_document_entry",true); item->setProperty("la_tab_title",tabs_->tabText(index));
            item->setCheckable(true); item->setChecked(!activePreviewDocument() && activeDataPage_==page);
            connect(item,&QAction::triggered,this,[this,page] { if(page) (void)activateDataPage(page); });
        }
        for(const auto& document:documents) {
            ++number;
            auto title=document->windowTitle(); title.replace('&',"&&");
            auto* item=windowMenu->addAction((number<10 ? "&" : "")+QString::number(number)+" "+title);
            item->setProperty("la_document_entry",true); item->setCheckable(true); item->setChecked(activePreviewDocument()==document);
            connect(item,&QAction::triggered,this,[this,document] { if(document && finishActiveDataEditors()) (void)AccessFormPreview::activateProjectDocument(db_,this,document); });
        }
    });
    auto* helpMenu=menuBar()->addMenu("&?");
    auto* about=helpMenu->addAction("Über &LinAccess..."); about->setObjectName("application-about");
    connect(about,&QAction::triggered,this,[this] {
        QMessageBox::about(this,"Über LinAccess","LinAccess\nNative Datenbankanwendung für Linux.\nEntwicklungsstand: Access-Kompatibilität ist noch nicht vollständig.");
    });
    statusBar()->showMessage("Neue SQLite-Datenbank erstellen oder Access-Datei untersuchen");
}

int MainWindow::selectDatabaseObject(int objectType,const QString& name,bool activate) {
    if(!db_.isOpen()) return 3078;
    const QStringList kinds{"table","query","form","report","macro","module"};
    if(objectType<0 || objectType>=kinds.size()) return 3251;
    const auto locate=[&](int& error) -> QTreeWidgetItem* {
        QTreeWidgetItem* target=nullptr;
        const std::function<void(QTreeWidgetItem*)> visit=[&](QTreeWidgetItem* item) {
            if(item->data(0,Qt::UserRole+2).toString()==kinds[objectType]
                && (name.isEmpty() ? item->data(0,Qt::UserRole+4).toBool()
                    : !item->data(0,Qt::UserRole+4).toBool() && item->data(0,Qt::UserRole+3).toString().compare(name,Qt::CaseInsensitive)==0)) {
                if(target) error=5; else target=item;
            }
            for(int index=0;index<item->childCount();++index) visit(item->child(index));
        };
        for(int index=0;index<explorer_->topLevelItemCount();++index) visit(explorer_->topLevelItem(index));
        if(!target && !error) error=7874;
        return target;
    };
    int error=0; auto* target=locate(error); if(error) return error;
    if(QApplication::activeModalWidget()) return 2046;
    if(!activate) return 0;
    focusDatabaseWindow(); if(!AccessFormPreview::databaseWindowActive(this)) return 2501;
    target=locate(error); if(error) return error;
    explorer_->setCurrentItem(target); explorer_->scrollToItem(target); return 0;
}

bool MainWindow::finishActiveDataEditors() {
    if(activePreviewDocument() || !activeDataPage_ || tabs_->indexOf(activeDataPage_)<0) return true;
    for(auto* candidate:activeDataPage_->findChildren<QTableView*>())
        if(auto* view=dynamic_cast<linaccess::AccessTableView*>(candidate);view && !view->finishEditing()) return false;
    return true;
}

QVariantMap MainWindow::saveDataPage(QWidget* page,bool interactive,bool currentOnly) {
    QPointer<MainWindow> self=this;
    QPointer<linaccess::AccessTableView> table;
    if(page) for(auto* candidate:page->findChildren<QTableView*>()) if(auto* view=dynamic_cast<linaccess::AccessTableView*>(candidate)) { table=view; break; }
    QPointer<linaccess::AccessTableModel> model=table ? qobject_cast<linaccess::AccessTableModel*>(table->model()) : nullptr;
    if(!table || !model) return tableCommandFailure(2046,"Kein aktives natives Tabellendokument.");
    if(model->isReadOnly()) return tableCommandFailure(2046,"Speichern ist in der schreibgeschützten Tabelle nicht verfügbar.");
    const auto failure=[&](int fallback) { return tableModelFailure(model,fallback); };
    if(!table->finishEditing()) return failure(2501);
    if(!self || !model || !table) return tableCommandFailure(2501,"Tabellendokument wurde während der Eingabevalidierung geschlossen.");
    const auto sourceError=validateEditableQuerySource(page,model);
    if(!sourceError.isEmpty()) return sourceError;
    const auto selected=table->currentIndex();
    if(currentOnly && !model->canSubmitCurrentRecordAlone(selected.row(),true)) {
        if(!model->submitExistingRecord(selected.row())) return failure(3146);
        return {{"handled",true},{"error",0}};
    }
    const auto primary=model->primaryKey();
    QMap<QString,QVariant> selectedKey;
    if(selected.isValid() && primary.count()) {
        const auto record=model->record(selected.row());
        for(int i=0;i<primary.count();++i) {
            const auto field=primary.fieldName(i); const auto value=record.value(field);
            if(value.isNull()) { selectedKey.clear(); break; }
            selectedKey.insert(field,value);
        }
    }
    if(!model->submitTransactionally()) {
        const auto result=failure(3146);
        if(interactive && self) messageError(this,result.value("description").toString());
        return result;
    }
    if(!self || !model || !table) return tableCommandFailure(2501,"Tabellendokument wurde beim Speichern geschlossen.");
    if(selected.isValid()) {
        const auto insertedKey=model->submittedPrimaryKey(selected.row());
        if(!insertedKey.isEmpty()) selectedKey=insertedKey;
    }
    // Model submission/reset invalidates Qt indexes. Restore the same keyed
    // record rather than a row number that may now identify another row.
    if(!selectedKey.isEmpty()) {
        int row=0;
        while(row<1000000) {
            if(row>=model->rowCount()) {
                if(!model->canFetchMore()) break;
                const auto count=model->rowCount(); model->fetchMore();
                if(model->rowCount()==count) break;
            }
            const auto record=model->record(row);
            bool matches=true;
            for(auto key=selectedKey.cbegin();key!=selectedKey.cend();++key)
                if(record.value(key.key())!=key.value()) { matches=false; break; }
            if(matches) { table->setCurrentIndex(model->index(row,selected.column())); table->scrollTo(table->currentIndex()); break; }
            ++row;
        }
    }
    return {{"handled",true},{"error",0}};
}

QVariantMap MainWindow::runProjectMacro(const QString& name) {
    if(!db_.isOpen()) return tableCommandFailure(3420,"Projekt ist geschlossen.");
    if(QApplication::activeModalWidget() || queryExecutionDepth_>0 || macroExecutionDepth_>0 || linaccess::nativeProjectMacrosActive(db_)) return tableCommandFailure(2046,"Makro kann in diesem Kontext nicht gestartet werden.");
    const QPointer<MainWindow> owner=this;const auto identity=db_.databaseName(),connection=db_.connectionName();const auto generation=property("la_project_generation");
    const auto check=[owner,identity,connection,generation] {if(!owner || !owner->db_.isOpen() || owner->db_.databaseName()!=identity || owner->db_.connectionName()!=connection || owner->property("la_project_generation")!=generation) throw linaccess::VbaRuntimeError(3420,"Makro gehört nicht mehr zum aktuellen Projekt.");};
    QScopedValueRollback<int> execution(macroExecutionDepth_,macroExecutionDepth_+1);
    const auto refresh=qScopeGuard([owner] {if(owner && owner->db_.isOpen()) owner->refreshExplorer();});
    try {linaccess::NativeStandardModuleHost host(this,db_,check);(void)host.applicationObject()->member("DoCmd",{}).object->member("RunMacro",{linaccess::VbaValue::string(name.toStdString())});check();return {{"handled",true},{"error",0}};}
    catch(const linaccess::VbaExecutionStopped&) {return {{"handled",true},{"error",0},{"stopped",true}};}
    catch(const linaccess::VbaRuntimeError& error) {auto result=tableCommandFailure(error.number,QString::fromUtf8(error.what()));result.insert("source",QString::fromStdString(error.source));result.insert("help_file",QString::fromStdString(error.helpFile));result.insert("help_context",error.helpContext);return result;}
    catch(const std::exception& error) {return tableCommandFailure(3146,QString::fromUtf8(error.what()));}
}

QVariantMap MainWindow::renameDatabaseObject(const QString& newName,int objectType,const QString& oldName) {
    if(!db_.isOpen()) return tableCommandFailure(3420,"Das Projekt ist geschlossen.");
    if(QApplication::activeModalWidget() || queryExecutionDepth_>0) return tableCommandFailure(2046,"Umbenennen ist in diesem Kontext nicht verfügbar.");
    if(objectType!=0 && objectType!=1) return tableCommandFailure(3251,"DoCmd.Rename unterstützt hier eigene Tabellen und Abfragen mit ausdrücklichem Objekttyp und Namen.");
    try {
        const auto kind=objectType==0 ? QString("table") : QString("query");
        if(objectType==0) {for(auto* model:findChildren<linaccess::AccessTableModel*>()) if(model->tableName().compare(oldName,Qt::CaseInsensitive)==0) return tableCommandFailure(3251,"Tabelle und zugehörige bearbeitbare Abfragen vor dem Umbenennen schließen.");}
        else if(findOpenQuery(oldName)) return tableCommandFailure(3251,"Abfrage vor dem Umbenennen schließen.");
        const auto definition=linaccess::makeDaoDatabase(db_)->member(objectType==0 ? "TableDefs" : "QueryDefs",{linaccess::VbaValue::string(oldName.toStdString())}).object;
        definition->assign("Name",linaccess::VbaValue::string(newName.toStdString()));const auto canonical=QString::fromStdString(definition->member("Name",{}).asText());refreshExplorer();
        for(auto* entry:explorer_->findItems(canonical,Qt::MatchExactly|Qt::MatchRecursive,0)) if(entry->data(0,Qt::UserRole).toString()==kind) {explorer_->setCurrentItem(entry);break;}
        return {{"handled",true},{"error",0}};
    }catch(const linaccess::DaoStoredSchemaChange& error) {refreshExplorer();return tableCommandFailure(3251,"Änderung gespeichert; Projekt neu öffnen. "+QString::fromUtf8(error.what()));}
    catch(const linaccess::VbaRuntimeError& error) {return tableCommandFailure(error.number,QString::fromUtf8(error.what()));}
    catch(const std::exception& error) {return tableCommandFailure(3146,QString::fromUtf8(error.what()));}
}

QVariantMap MainWindow::deleteDatabaseObject(int objectType,const QString& name) {
    if(!db_.isOpen()) return tableCommandFailure(3420,"Das Projekt ist geschlossen.");
    if(QApplication::activeModalWidget() || queryExecutionDepth_>0) return tableCommandFailure(2046,"Löschen ist in diesem Kontext nicht verfügbar.");
    if(objectType!=0 && objectType!=1 && objectType!=4) return tableCommandFailure(3251,"DeleteObject unterstützt hier eigene Tabellen, Abfragen und Makros mit ausdrücklichem Objekttyp und Namen.");
    try {
        if(objectType==4) {const auto revision=linaccess::readNativeMacroRevision(db_,name);if(!revision) throw linaccess::VbaRuntimeError(linaccess::importedMacroNameExists(db_,name) ? 3251 : 2485,"Kein eigenes gespeichertes Makroobjekt verfügbar.");linaccess::deleteNativeMacro(db_,*revision);refreshExplorer();return {{"handled",true},{"error",0}};}
        if(objectType==0) {
            const auto fields=linaccess::nativeTableFieldDisplayProperties(db_,name);if(fields.isEmpty() || std::any_of(fields.cbegin(),fields.cend(),[](const auto& field) {return !field.nativeOwned;})) return tableCommandFailure(3251,"DeleteObject benötigt hier eine selbst angelegte native Tabelle.");
            for(auto* model:findChildren<linaccess::AccessTableModel*>()) if(model->tableName().compare(name,Qt::CaseInsensitive)==0) return tableCommandFailure(3251,"Tabelle und zugehörige bearbeitbare Abfragen vor dem Löschen schließen.");
        }else if(findOpenQuery(name)) return tableCommandFailure(3251,"Abfrage vor dem Löschen schließen.");
        linaccess::makeDaoDatabase(db_)->member(objectType==0 ? "TableDefs" : "QueryDefs",{}).object->member("Delete",{linaccess::VbaValue::string(name.toStdString())});refreshExplorer();return {{"handled",true},{"error",0}};
    }catch(const linaccess::VbaRuntimeError& error) {return tableCommandFailure(error.number,QString::fromUtf8(error.what()));}
    catch(const std::exception& error) {return tableCommandFailure(3146,QString::fromUtf8(error.what()));}
}

QWidget* MainWindow::findOpenQuery(const QString& name) const {
    for(int index=0;index<tabs_->count();++index) {
        auto* page=tabs_->widget(index);
        if(page->property("la_project_object_kind").toString()=="query"
            && page->property("la_project_object_name").toString().compare(name,Qt::CaseInsensitive)==0) return page;
    }
    return nullptr;
}

QVariantMap MainWindow::openQueryDocument(const QString& name,int view,int mode) {
    return loadQueryDocument(name,view,mode,nullptr);
}

QVariantMap MainWindow::loadQueryDocument(const QString& name,int view,int mode,QWidget* refresh) {
    if(!db_.isOpen() || QApplication::activeModalWidget()) return tableCommandFailure(2046,"Abfrage kann in diesem Kontext nicht geöffnet werden.");
    if(view!=0 || (mode!=1 && mode!=2)) return tableCommandFailure(3251,"Abfrage-Dokumente unterstützen bisher Normalansicht mit acEdit oder acReadOnly.");
    if(name.isEmpty()) return tableCommandFailure(7874,"Kein Abfragename angegeben.");
    if(auto* existing=findOpenQuery(name);existing && !refresh) {
        if(existing->property("la_query_data_mode").toInt()!=mode) return tableCommandFailure(3251,"Die geöffnete Abfrage hat einen anderen Datenmodus.");
        if(!activateDataPage(existing)) return tableCommandFailure(2501,"Abfrageaktivierung wurde abgebrochen.");
        return {{"handled",true},{"error",0}};
    }
    const QPointer<MainWindow> self=this;
    const QPointer<QWidget> refreshing=refresh;
    if(refreshing && tabs_->indexOf(refreshing)<0) return tableCommandFailure(2501,"Abfragedokument ist nicht mehr geöffnet.");
    if(executingQueries_.contains(name.toCaseFolded())) return tableCommandFailure(3251,"Rekursive Ausführung derselben Abfrage ist nicht implementiert.");
    const auto database=db_; const auto identity=db_.databaseName();
    const auto generation=property("la_project_generation").toULongLong();
    const auto alive=[self,identity,generation] {
        return self && self->db_.isOpen() && self->db_.databaseName()==identity && self->property("la_project_generation").toULongLong()==generation;
    };
    try {
        QString source,canonical; bool native=false,imported=false;
        QSqlQuery query(database);
        if(!query.prepare("SELECT name,definition FROM main._la_objects WHERE kind='query' AND name=? COLLATE NOCASE")) return tableCommandFailure(3146,"Abfragekatalog konnte nicht vorbereitet werden.");
        query.addBindValue(name);
        if(!query.exec()) return tableCommandFailure(3146,"Abfragekatalog konnte nicht gelesen werden.");
        while(query.next()) {
            if(native) return tableCommandFailure(5,"Mehrdeutiger Abfragename.");
            native=true; canonical=query.value(0).toString();
            source=QJsonDocument::fromJson(query.value(1).toString().toUtf8()).object().value("sql").toString();
        }
        if(query.lastError().isValid()) return tableCommandFailure(3146,"Abfragekatalog konnte nicht vollständig gelesen werden.");
        if(!query.prepare("SELECT name,source_bytes FROM main._la_access_artifacts WHERE kind='query' AND name=? COLLATE NOCASE ORDER BY id DESC LIMIT 1")) return tableCommandFailure(3146,"Importkatalog konnte nicht vorbereitet werden.");
        query.addBindValue(name);
        if(!query.exec()) return tableCommandFailure(3146,"Importkatalog konnte nicht gelesen werden.");
        if(query.next()) {
            imported=true;
            if(!native) { canonical=query.value(0).toString(); source=linaccess::decodeVbaSource(query.value(1).toByteArray()); }
        }
        if(query.lastError().isValid()) return tableCommandFailure(3146,"Importkatalog konnte nicht vollständig gelesen werden.");
        query.finish();
        if(native && imported) return tableCommandFailure(5,"Mehrdeutiger nativer/importierter Abfragename.");
        if(!native && !imported) return tableCommandFailure(7874,"Gespeicherte Abfrage nicht gefunden.");
        if(source.trimmed().isEmpty()) return tableCommandFailure(3251,"Gespeicherte Abfragedefinition ist nicht ausführbar.");
        const auto plan=linaccess::compileAccessQuery(source);
        if(!plan.makeTable.isEmpty() || !isSelect(plan.sql)) return tableCommandFailure(3251,"Nur schreibgeschützte SELECT-Abfragen werden als Dokument geöffnet.");
        if(mode==1) {
            const auto table=linaccess::editableAccessQueryTable(database,source);
            if(refresh) {
                if(!refreshing || !refreshing->property("la_editable_query").toBool() || refreshing->property("la_editable_query_source").toString()!=source)
                    return tableCommandFailure(3251,"Geänderte Definitionen geöffneter bearbeitbarer Abfragen sind noch nicht implementiert.");
                const auto key=canonical.toCaseFolded(); executingQueries_.insert(key); ++queryExecutionDepth_;
                const auto finished=qScopeGuard([self,key] { if(self) { self->executingQueries_.remove(key); --self->queryExecutionDepth_; } });
                return requeryTableDataPage(refreshing);
            }
            const int error=showTable(table,false,false,canonical,source);
            if(error) return tableCommandFailure(error,"Bearbeitbare Abfrage konnte nicht geöffnet werden.");
            return {{"handled",true},{"error",0}};
        }
        if(!refreshing && (!finishActiveDataEditors() || !AccessFormPreview::activateProjectDocument(database,this,nullptr))) return tableCommandFailure(2501,"Dokumentwechsel wurde abgebrochen.");
        if(!alive()) return tableCommandFailure(2501,"Projekt wurde während des Dokumentwechsels geschlossen.");
        QMap<QString,QVariant> parameters;
        for(const auto& parameter:plan.parameters) {
            bool accepted=false;
            const auto value=QInputDialog::getText(self,"Abfrageparameter",parameter.name+" ("+parameter.type+")",QLineEdit::Normal,{},&accepted);
            if(!alive() || !accepted) return tableCommandFailure(2501,"Abfrageparameter wurden abgebrochen.");
            parameters.insert(parameter.name.toCaseFolded(),value);
        }
        auto page=std::make_unique<QWidget>();
        auto* layout=new QVBoxLayout(page.get());
        layout->addWidget(new QLabel("Abfrage: "+canonical+" — Schreibgeschützt",page.get()));
        auto* model=new linaccess::AccessQueryModel(page.get());
        {
            const auto key=canonical.toCaseFolded();
            if(executingQueries_.contains(key)) return tableCommandFailure(3251,"Rekursive Ausführung derselben Abfrage ist nicht implementiert.");
            executingQueries_.insert(key); ++queryExecutionDepth_;
            const auto finished=qScopeGuard([self,key] { if(self) { self->executingQueries_.remove(key); --self->queryExecutionDepth_; } });
            model->execute(database,source,parameters);
        }
        if(!alive() || linaccess::vbaExecutionStopRequested()) return tableCommandFailure(2501,"Abfrageöffnung wurde beendet.");
        if(refresh) {
            if(!refreshing || tabs_->indexOf(refreshing)<0) return tableCommandFailure(2501,"Abfragedokument wurde während der Ausführung geschlossen.");
            const QPointer<QTableView> target=refreshing->findChild<QTableView*>("native-query-view");
            if(!target) return tableCommandFailure(2501,"Abfrageansicht ist nicht mehr verfügbar.");
            const int column=qMax(0,target->currentIndex().column());
            const QPointer<QAbstractItemModel> previous=target->model();
            const QPointer<linaccess::AccessQueryModel> replacement=model;
            model->setParent(refreshing); target->setModel(model);
            if(previous) previous->deleteLater();
            if(!alive() || !refreshing || !target || !replacement || tabs_->indexOf(refreshing)<0)
                return tableCommandFailure(2501,"Abfragedokument wurde beim Ergebnisaustausch geschlossen.");
            if(model->rowCount() && model->columnCount()) target->setCurrentIndex(model->index(0,qMin(column,model->columnCount()-1)));
            if(!alive() || !refreshing || !target || !replacement || linaccess::vbaExecutionStopRequested())
                return tableCommandFailure(2501,"Abfragedokument wurde während der Auswahl geschlossen.");
            return {{"handled",true},{"error",0}};
        }
        auto* table=new linaccess::NativeQueryView(page.get()); table->setObjectName("native-query-view"); table->setModel(model); table->setEditTriggers(QAbstractItemView::NoEditTriggers);
        if(model->rowCount() && model->columnCount()) table->setCurrentIndex(model->index(0,0));
        table->installEventFilter(this); layout->addWidget(table);
        auto* document=page.release(); document->setParent(tabs_);
        const QPointer<QWidget> registered=document;
        tagProjectObject(document,this,db_,"query",canonical); document->setProperty("la_project_object_scope","tab");
        document->setProperty("la_query_data_mode",2);
        for(const auto key:{QKeySequence(Qt::ControlModifier|Qt::Key_W),QKeySequence(Qt::ControlModifier|Qt::Key_F4)}) {
            auto* close=new QShortcut(key,document); close->setContext(Qt::WidgetWithChildrenShortcut);
            connect(close,&QShortcut::activated,document,[this,document] { const int index=tabs_->indexOf(document); if(index>=0) emit tabs_->tabCloseRequested(index); });
        }
        for(const auto key:{QKeySequence(Qt::ControlModifier|Qt::Key_C),QKeySequence(Qt::ControlModifier|Qt::Key_Insert)}) {
            auto* copy=new QShortcut(key,document); copy->setContext(Qt::WidgetWithChildrenShortcut);
            connect(copy,&QShortcut::activated,document,[this] {
                const auto result=dispatchTableCommand(190);
                if(result.value("error").toInt()) statusBar()->showMessage(result.value("description").toString(),10000);
            });
        }
        for(const auto key:{QKeySequence(Qt::ControlModifier|Qt::Key_F),QKeySequence(Qt::ShiftModifier|Qt::Key_F4)}) {
            auto* find=new QShortcut(key,document); find->setContext(Qt::WidgetWithChildrenShortcut);
            connect(find,&QShortcut::activated,document,[this,key] {
                const auto result=dispatchTableCommand(key==QKeySequence(Qt::ControlModifier|Qt::Key_F) ? 30 : 341);
                if(result.value("error").toInt()) statusBar()->showMessage(result.value("description").toString(),10000);
            });
        }
        auto* zoom=new QShortcut(QKeySequence(Qt::ShiftModifier|Qt::Key_F2),document); zoom->setContext(Qt::WidgetWithChildrenShortcut);
        connect(zoom,&QShortcut::activated,document,[this] {
            const auto result=dispatchTableCommand(179);
            if(result.value("error").toInt()) statusBar()->showMessage(result.value("description").toString(),10000);
        });
        auto* requery=new QShortcut(QKeySequence(Qt::ShiftModifier|Qt::Key_F9),document);
        requery->setContext(Qt::WidgetWithChildrenShortcut);
        connect(requery,&QShortcut::activated,document,[this] {
            const auto result=requeryTableDocument();
            if(const int error=result.value("error").toInt()) statusBar()->showMessage(result.value("description").toString()+" ("+QString::number(error)+")",10000);
        });
        tabs_->setCurrentIndex(tabs_->addTab(document,"Abfrage: "+canonical));
        const bool activated=activateDataPage(document);
        if(!alive() || !registered) return tableCommandFailure(2501,"Abfragedokument wurde während der Aktivierung geschlossen.");
        if(!activated) {
            const int index=tabs_->indexOf(document); if(index>=0) tabs_->removeTab(index);
            document->deleteLater(); return tableCommandFailure(2501,"Abfrageaktivierung wurde abgebrochen.");
        }
        table->setFocus();
        return {{"handled",true},{"error",0}};
    } catch(const linaccess::VbaExecutionStopped&) { return tableCommandFailure(2501,"Abfrageöffnung wurde beendet."); }
    catch(const linaccess::VbaRuntimeError& error) { return tableCommandFailure(error.number,QString::fromUtf8(error.what())); }
    catch(const std::exception&) { return tableCommandFailure(3146,"Abfrage konnte nicht ausgeführt werden."); }
}

QVariantMap MainWindow::closeQueryDocument(int type,const QString& name,int save) {
    const bool active=!activePreviewDocument() && AccessFormPreview::databaseWindowActive(this) && activeDataPage_
        && tabs_->currentWidget()==activeDataPage_ && activeDataPage_->property("la_project_object_kind").toString()=="query";
    if(type!=1 && !(type==-1 && active)) return {};
    if(save<0 || save>2) return tableCommandFailure(5,"Ungültige Speicheroption.");
    if(!db_.isOpen() || QApplication::activeModalWidget()) return tableCommandFailure(2046,"Abfrage kann in diesem Kontext nicht geschlossen werden.");
    const QPointer<QWidget> page=name.isEmpty() ? (active ? activeDataPage_.data() : nullptr) : findOpenQuery(name);
    if(!page || tabs_->indexOf(page)<0) return tableCommandFailure(3251,"Die angeforderte offene Abfrage ist nicht verfügbar.");
    if(page->property("la_editable_query").toBool()) return closeTableDataPage(page);
    tabs_->removeTab(tabs_->indexOf(page)); page->hide(); page->deleteLater();
    return {{"handled",true},{"error",0}};
}

QVariantMap MainWindow::dispatchTableCommand(int command) {
    const QPointer<QWidget> page=activeDataPage_;
    if(!db_.isOpen() || activePreviewDocument() || !AccessFormPreview::databaseWindowActive(this) || !page || tabs_->indexOf(page)<0 || tabs_->currentWidget()!=page) return {};
    if(QApplication::activeModalWidget()) return tableCommandFailure(2046,"Tabellenbefehl ist durch ein modales Fenster blockiert.");
    if(command==32) return setupDatasheetPage();
    if(command==340) return outputDatasheet(false);
    if(command==54) return previewDatasheet();
    if((page->property("la_project_object_kind").toString()=="query" && !page->property("la_editable_query").toBool())) {
        if(const auto focused=focusedDataTextCommand(page->findChild<QTableView*>("native-query-view"),command);focused.value("handled").toBool()) return focused;
    }
    if((page->property("la_project_object_kind").toString()=="query" && !page->property("la_editable_query").toBool()) && (command==50 || command==109)) {
        auto* view=page->findChild<QTableView*>("native-query-view");
        if(!view || !view->model() || !view->selectionModel()) return tableCommandFailure(2046,"Kein natives Abfrageergebnis verfügbar.");
        if(command==50) {
            if(!view->currentIndex().isValid()) return tableCommandFailure(2046,"Kein aktueller Abfragedatensatz.");
            view->selectionModel()->select(view->currentIndex(),QItemSelectionModel::ClearAndSelect|QItemSelectionModel::Rows);
        } else {
            if(static_cast<qint64>(view->model()->rowCount())*view->model()->columnCount()>1000000)
                return tableCommandFailure(3251,"Abfrageauswahl überschreitet die native Zellgrenze.");
            view->selectAll();
        }
        view->setFocus(); return {{"handled",true},{"error",0}};
    }
    if((page->property("la_project_object_kind").toString()=="query" && !page->property("la_editable_query").toBool()) && command==190) {
        auto* view=dynamic_cast<linaccess::NativeQueryView*>(page->findChild<QTableView*>("native-query-view"));
        if(!view) return tableCommandFailure(2046,"Kein natives Abfrageergebnis verfügbar.");
        try {
            view->copySelection();
            return {{"handled",true},{"error",0}};
        } catch(const linaccess::VbaRuntimeError& error) { return tableCommandFailure(error.number,"Die Abfrageauswahl kann nicht verlustfrei als Text kopiert werden."); }
    }
    if((page->property("la_project_object_kind").toString()=="query" && !page->property("la_editable_query").toBool()) && (command==30 || command==341)) {
        auto* view=dynamic_cast<linaccess::NativeQueryView*>(page->findChild<QTableView*>("native-query-view"));
        if(!view) return tableCommandFailure(2046,"Keine native Abfrageansicht.");
        try {
            if(command==30 || !view->hasFindText()) view->openFindDialog(); else view->repeatFind();
            return {{"handled",true},{"error",0}};
        } catch(const linaccess::VbaRuntimeError& error) { return tableCommandFailure(error.number,QString::fromUtf8(error.what())); }
    }
    if((page->property("la_project_object_kind").toString()=="query" && !page->property("la_editable_query").toBool()) && command==179) {
        auto* view=dynamic_cast<linaccess::NativeQueryView*>(page->findChild<QTableView*>("native-query-view"));
        if(!view || (QApplication::focusWidget() && QApplication::focusWidget()->window()!=view->window()))
            return tableCommandFailure(2046,"Kein aktives Abfragefeld für Zoom.");
        try { view->zoomField(); return {{"handled",true},{"error",0}}; }
        catch(const linaccess::VbaRuntimeError& error) { return tableCommandFailure(error.number,QString::fromUtf8(error.what())); }
    }
    if(command==97) return saveDataPage(page,false);
    if(command>=65 && command<=68) {
        const int actions[]{1,0,2,3};
        return navigateTableDocument({},actions[command-65],1,false);
    }
    if(command==28 || command==38 || command==223 || command==50 || command==109 || command==189 || command==190 || command==191 || command==199 || command==292 || command==179 || command==163 || command==164 || command==29 || command==30 || command==341) {
        QPointer<linaccess::AccessTableView> table;
        for(auto* candidate:page->findChildren<QTableView*>()) if(auto* view=dynamic_cast<linaccess::AccessTableView*>(candidate)) { table=view; break; }
        QPointer<linaccess::AccessTableModel> model=table ? qobject_cast<linaccess::AccessTableModel*>(table->model()) : nullptr;
        if(!table || !model) return tableCommandFailure(2046,"Keine aktive Tabellenzelle verfügbar.");
        if(const auto focused=focusedDataTextCommand(table,command);focused.value("handled").toBool()) return focused;
        if((command!=28 && command!=38 && command!=292 && command!=199 && command!=109 && command!=29 && command!=30 && command!=341 && !(command==191 && model->rowCount()==0)) && !table->currentIndex().isValid()) return tableCommandFailure(2046,"Keine aktive Tabellenzelle verfügbar.");
        if(command==28) {
            const QPointer<MainWindow> self=this;
            if(!table->finishEditing()) return tableModelFailure(model,2501);
            if(!self || !table || !model) return tableCommandFailure(2501,"Tabellendokument wurde während der Validierung geschlossen.");
            if(model->isDirty() && model->hasPendingRecord(table->currentIndex().row())) {
                const auto saved=saveDataPage(page,false,true);
                if(saved.value("error").toInt()) return saved;
            }
            if(!self || !table || !model) return tableCommandFailure(2501,"Tabellendokument wurde während des Speicherns geschlossen.");
            if(!table->insertRecord()) return tableModelFailure(model,2501);
        } else if(command==223) {
            if(!table->deleteCurrentRecord()) return tableModelFailure(model,2501);
        } else if(command==29) {
            if(!table->openReplaceDialog()) return tableModelFailure(model,2046);
        } else if(command==30 || command==341) {
            if(!(command==30 ? table->openFindDialog() : table->repeatFind())) return tableModelFailure(model,2046);
        } else if(command==109) {
            if(!table->selectAllRecords()) return tableModelFailure(model,2501);
        } else if(command==190) {
            if(!table->copySelection()) return tableModelFailure(model,2501);
        } else if(command==191 || command==38) {
            if(!table->pasteSelection(command==38)) return tableModelFailure(model,2046);
        } else if(command==199 || command==292) {
            if(!table->undoPastedBlock(command==199)) return tableModelFailure(model,2046);
        } else if(command==189) {
            if(!table->cutSelection()) return tableModelFailure(model,2046);
        } else if(command==179) {
            if(!table->zoomCurrentCell()) return tableModelFailure(model,2046);
        } else if(command==163 || command==164) {
            QPointer<MainWindow> self=this;
            if(!table->finishEditing()) return tableModelFailure(model,2501);
            if(!self || !table || !model) return tableCommandFailure(2501,"Tabellendokument wurde während der Validierung geschlossen.");
            if(model->isDirty()) {
                if(!model->canSubmitCurrentRecordAlone(table->currentIndex().row()))
                    return tableCommandFailure(3251,"Sortieren mit weiteren Tabellenentwürfen ist noch nicht implementiert.");
                const auto saved=saveDataPage(page,false);
                if(saved.value("error").toInt()) return saved;
            }
            if(!self || !table || !model || !table->currentIndex().isValid()) return tableCommandFailure(2501,"Keine aktive Tabellenzelle nach dem Speichern.");
            const int column=table->currentIndex().column();
            QMap<QString,QVariant> key;
            const auto primary=model->primaryKey();
            const auto record=model->record(table->currentIndex().row());
            for(int field=0;field<primary.count();++field) key.insert(primary.fieldName(field),record.value(primary.fieldName(field)));
            table->sortByColumn(column,command==164 ? Qt::DescendingOrder : Qt::AscendingOrder);
            if(!self || !table || !model) return tableCommandFailure(2501,"Tabellendokument wurde während der Sortierung geschlossen.");
            if(model->lastError().isValid()) return tableModelFailure(model,3146);
            if(!key.isEmpty()) {
                int row=0;
                for(;row<1000000;++row) {
                    while(row>=model->rowCount() && model->canFetchMore()) {
                        const int before=model->rowCount(); model->fetchMore();
                        if(!self || !table || !model) return tableCommandFailure(2501,"Tabellendokument wurde beim Laden geschlossen.");
                        if(model->rowCount()==before) break;
                    }
                    if(row>=model->rowCount()) break;
                    const auto candidate=model->record(row);
                    bool matches=true;
                    for(auto field=key.cbegin();field!=key.cend();++field)
                        if(candidate.value(field.key())!=field.value()) { matches=false; break; }
                    if(matches) { table->setCurrentIndex(model->index(row,column)); table->scrollTo(table->currentIndex()); break; }
                }
                if(row>=1000000) return tableCommandFailure(3251,"Aktueller Datensatz überschreitet die native Suchgrenze nach der Sortierung.");
            }
        } else {
            if(!table->finishEditing()) return tableModelFailure(model,2501);
            if(!table || !model || !table->currentIndex().isValid()) return tableCommandFailure(2501,"Tabellenauswahl ist nicht mehr verfügbar.");
            const QPersistentModelIndex current=table->currentIndex();
            table->selectRow(current.row());
            if(table && current.isValid() && table->selectionModel())
                table->selectionModel()->setCurrentIndex(current,QItemSelectionModel::NoUpdate);
        }
        return {{"handled",true},{"error",0}};
    }
    return tableCommandFailure(3251,"Dieser native Tabellenbefehl ist noch nicht implementiert.");
}

QVariantMap MainWindow::requeryTableDocument() {
    const QPointer<QWidget> page=activeDataPage_;
    if(!db_.isOpen() || activePreviewDocument() || !AccessFormPreview::databaseWindowActive(this) || !page || tabs_->indexOf(page)<0 || tabs_->currentWidget()!=page) return {};
    if(QApplication::activeModalWidget()) return tableCommandFailure(2046,"Neuabfrage ist durch ein modales Fenster blockiert.");
    if(page->property("la_project_object_kind").toString()=="query")
        return loadQueryDocument(page->property("la_project_object_name").toString(),0,page->property("la_editable_query").toBool() ? 1 : 2,page);
    return requeryTableDataPage(page);
}

QVariantMap MainWindow::requeryTableDataPage(QWidget* rawPage) {
    const QPointer<QWidget> page=rawPage;
    QPointer<linaccess::AccessTableView> table;
    for(auto* candidate:page->findChildren<QTableView*>()) if(auto* view=dynamic_cast<linaccess::AccessTableView*>(candidate)) { table=view; break; }
    QPointer<linaccess::AccessTableModel> model=table ? qobject_cast<linaccess::AccessTableModel*>(table->model()) : nullptr;
    if(!table || !model) return {};
    const QPointer<MainWindow> self=this;
    if(!table->finishEditing()) return tableModelFailure(model,2501);
    if(!self || !table || !model) return tableCommandFailure(2501,"Tabellendokument wurde während der Validierung geschlossen.");
    const auto sourceError=validateEditableQuerySource(page,model);
    if(!sourceError.isEmpty()) return sourceError;
    const int column=qMax(0,table->currentIndex().column());
    if(model->isDirty()) {
        if(!model->canSubmitCurrentRecordAlone(table->currentIndex().row()))
            return tableCommandFailure(3251,"Neuabfrage mit weiteren Tabellenentwürfen ist noch nicht implementiert.");
        const auto saved=saveDataPage(page,false);
        if(saved.value("error").toInt()) return saved;
    }
    if(!self || !table || !model) return tableCommandFailure(2501,"Tabellendokument wurde während des Speicherns geschlossen.");
    if(!model->select()) return tableModelFailure(model,3146);
    if(table && model && model->rowCount() && model->columnCount())
        table->setCurrentIndex(model->index(0,qMin(column,model->columnCount()-1)));
    return {{"handled",true},{"error",0}};
}

QVariantMap MainWindow::tableKeyboardContext() {
    const QPointer<QWidget> page=activeDataPage_;
    if(!db_.isOpen() || activePreviewDocument() || !AccessFormPreview::databaseWindowActive(this) || !page || tabs_->indexOf(page)<0 || tabs_->currentWidget()!=page) return {};
    bool native=false;
    for(auto* candidate:page->findChildren<QTableView*>()) if(auto* view=dynamic_cast<linaccess::AccessTableView*>(candidate);view && qobject_cast<linaccess::AccessTableModel*>(view->model())) { native=true; break; }
    if(!native && page->property("la_project_object_kind").toString()=="query") {
        auto* view=page->findChild<QTableView*>("native-query-view");
        native=view && dynamic_cast<linaccess::AccessQueryModel*>(view->model());
    }
    if(!native) return {};
    QVariantMap result{{"handled",true},{"connection",db_.connectionName()},{"database",db_.databaseName()}};
    auto* focus=QApplication::focusWidget();
    auto* parent=focus;
    while(parent && parent!=page) parent=parent->parentWidget();
    if(!focus || !parent) return result; // QWidget::isAncestorOf stops at a dialog window.

    if(auto* modal=QApplication::activeModalWidget(); modal && modal!=focus->window()) return result;
    result.insert("target",QVariant::fromValue(static_cast<QObject*>(focus)));
    return result;
}

QVariantMap MainWindow::findTableRecord(const QVariantMap& request) {
    QPointer<QWidget> page=activeDataPage_;
    if(!db_.isOpen() || activePreviewDocument() || !AccessFormPreview::databaseWindowActive(this) || !page || tabs_->indexOf(page)<0 || tabs_->currentWidget()!=page) return {};
    if(QApplication::activeModalWidget()) return tableCommandFailure(2046,"FindRecord ist durch ein modales Fenster blockiert.");
    try {
        const auto sought=linaccess::nativeFindValueFromPayload(request);
        const int mode=request.value("match").toInt(),direction=request.value("direction").toInt();
        const bool sensitive=request.value("sensitive").toBool(),first=request.value("first").toBool(),formatted=request.value("formatted").toBool(),all=request.value("allfields").toBool();
        if(mode<0 || mode>2 || direction<0 || direction>2) return tableCommandFailure(5,"Ungültiger FindRecord-Vergleich oder Suchrichtung.");
        const bool textValue=sought.kind==linaccess::VbaValue::Kind::Text;
        if(formatted && !textValue) return tableCommandFailure(3251,"Formatierte Suche benötigt wörtlichen Text.");
        const auto text=linaccess::nativeFindValueText(sought);
        if(textValue && text.isEmpty()) return tableCommandFailure(3251,"Leere FindRecord-Suche ist noch nicht implementiert.");
        const bool wildcards=textValue && linaccess::nativeFindHasWildcards(text);
        if(wildcards) (void)linaccess::nativeFindTextMatches({},text,mode,sensitive,true);
        if(!textValue && mode!=1) return tableCommandFailure(3251,"Typisierte FindRecord-Suche unterstützt nur ganze Feldwerte.");
        if(page->property("la_project_object_kind").toString()=="query" && !page->property("la_editable_query").toBool()) {
            auto* view=dynamic_cast<linaccess::NativeQueryView*>(page->findChild<QTableView*>("native-query-view"));
            if(!view || !view->currentIndex().isValid()) return tableCommandFailure(2046,"Kein aktuelles Abfragefeld für FindRecord.");
            const int column=view->currentIndex().column();
            auto* result=dynamic_cast<linaccess::AccessQueryModel*>(view->model());
            if(!result || (!all && (view->isColumnHidden(column) || (!formatted && !linaccess::nativeFindSupports(sought,result->fieldSqlType(column))))))
                return tableCommandFailure(3251,"FindRecord-Suchwert und Abfragefeldtyp passen nicht zum unterstützten Teilumfang.");
            linaccess::NativeTableFindArguments arguments{text,all ? -1 : column,mode,sensitive,first ? 1 : 2,direction};arguments.formatted=formatted;arguments.wildcards=wildcards;
            if(!textValue) arguments.value=sought;
            view->findText(std::move(arguments));
            return {{"handled",true},{"error",0}};
        }
        QPointer<linaccess::AccessTableView> table;
        for(auto* candidate:page->findChildren<QTableView*>()) if(auto* view=dynamic_cast<linaccess::AccessTableView*>(candidate)) { table=view; break; }
        QPointer<linaccess::AccessTableModel> model=table ? qobject_cast<linaccess::AccessTableModel*>(table->model()) : nullptr;
        if(!table || !model || !table->currentIndex().isValid()) return tableCommandFailure(2046,"Kein aktuelles Tabellenfeld für FindRecord.");
        const int column=table->currentIndex().column();
        if(!all && (table->isColumnHidden(column) || (!formatted && !linaccess::nativeFindSupports(sought,model->fieldSqlType(column)))))
            return tableCommandFailure(3251,"FindRecord-Suchwert und aktueller Feldtyp passen nicht zum unterstützten Teilumfang.");
        linaccess::NativeTableFindArguments arguments{text,all ? -1 : column,mode,sensitive,first ? 1 : 2,direction};arguments.formatted=formatted;arguments.wildcards=wildcards;
        if(!textValue) arguments.value=sought;
        if(table->findText(std::move(arguments))<0) return tableModelFailure(model,2501);
        return {{"handled",true},{"error",0}};
    } catch(const linaccess::VbaRuntimeError& error) { return tableCommandFailure(error.number,QString::fromUtf8(error.what())); }
}

QVariantMap MainWindow::navigateQueryDocument(const QString& name,int action,int offset) {
    if(!db_.isOpen() || QApplication::activeModalWidget()) return tableCommandFailure(2046,"Abfragenavigation ist in diesem Kontext nicht verfügbar.");
    const QPointer<QWidget> page=name.isEmpty() ? activeDataPage_.data() : findOpenQuery(name);
    if(!page || tabs_->indexOf(page)<0 || page->property("la_project_object_kind").toString()!="query")
        return tableCommandFailure(3251,"Die angeforderte offene Abfrage ist nicht verfügbar.");
    const QPointer<MainWindow> self=this;
    if(!activateDataPage(page)) return tableCommandFailure(2501,"Abfrageaktivierung wurde abgebrochen.");
    if(!self || !page) return tableCommandFailure(2501,"Abfragedokument wurde während der Aktivierung geschlossen.");
    return navigateTableDocument({},action,offset,false);
}

QVariantMap MainWindow::navigateTableDocument(const QString& name,int action,int offset,bool explicitTable) {
    const bool active= !activePreviewDocument() && AccessFormPreview::databaseWindowActive(this) && activeDataPage_
        && tabs_->indexOf(activeDataPage_)>=0 && tabs_->currentWidget()==activeDataPage_;
    if(!explicitTable && !active) return {};
    if(action<0 || action>5) return tableCommandFailure(5,"Ungültige Datensatznavigation.");
    if((action==0 || action==1 || action==4) && offset<1) return tableCommandFailure(2105,"Ungültiger Datensatzabstand.");
    if(!db_.isOpen() || QApplication::activeModalWidget()) return tableCommandFailure(2046,"Tabellennavigation ist in diesem Kontext nicht verfügbar.");
    QPointer<MainWindow> self=this;
    QPointer<QWidget> page=name.isEmpty() ? (active ? activeDataPage_.data() : nullptr) : findOpenTable(name);
    if(!page) return tableCommandFailure(3251,"Die angeforderte offene Tabelle ist nicht verfügbar.");
    if((!active || page!=activeDataPage_) && !activateDataPage(page)) return tableCommandFailure(2501,"Tabellenaktivierung wurde abgebrochen.");
    if(!self || !page) return tableCommandFailure(2501,"Tabellendokument wurde während der Aktivierung geschlossen.");
    if(page->property("la_project_object_kind").toString()=="query" && explicitTable) return tableCommandFailure(3251,"Die aktive Abfrage ist keine Tabelle.");
    if(page->property("la_project_object_kind").toString()=="query" && !page->property("la_editable_query").toBool()) {
        const QPointer<QTableView> view=page->findChild<QTableView*>("native-query-view");
        auto* result=view ? dynamic_cast<linaccess::AccessQueryModel*>(view->model()) : nullptr;
        if(!result) return tableCommandFailure(2046,"Kein natives Abfrageergebnis verfügbar.");
        if(action==5) return tableCommandFailure(2448,"Schreibgeschützte Abfrage erlaubt keinen neuen Datensatz.");
        const auto current=view->currentIndex();
        if((action==0 || action==1) && !current.isValid()) return tableCommandFailure(2105,"Kein aktueller Datensatz.");
        const qint64 target=action==2 ? 0 : action==3 ? static_cast<qint64>(result->rowCount())-1
            : action==4 ? static_cast<qint64>(offset)-1
            : static_cast<qint64>(current.row())+(action==0 ? -static_cast<qint64>(offset) : offset);
        if(target<0 || target>=result->rowCount()) return tableCommandFailure(2105,"Der angeforderte Datensatz ist nicht verfügbar.");
        const auto destination=result->index(static_cast<int>(target),current.isValid() ? current.column() : 0);
        if(!destination.isValid()) return tableCommandFailure(2105,"Der angeforderte Datensatz ist nicht verfügbar.");
        view->setCurrentIndex(destination); view->scrollTo(destination); view->setFocus();
        return {{"handled",true},{"error",0}};
    }
    QPointer<linaccess::AccessTableView> table;
    for(auto* candidate:page->findChildren<QTableView*>()) if(auto* view=dynamic_cast<linaccess::AccessTableView*>(candidate)) { table=view; break; }
    QPointer<linaccess::AccessTableModel> model=table ? qobject_cast<linaccess::AccessTableModel*>(table->model()) : nullptr;
    if(!table || !model) return tableCommandFailure(2046,"Kein natives Tabellendokument.");
    if(action==5) return dispatchTableCommand(28);
    if(!table->finishEditing()) return tableModelFailure(model,2501);
    if(!self || !table || !model) return tableCommandFailure(2501,"Tabellendokument wurde während der Validierung geschlossen.");
    // A clean current record can move without writing unrelated drafts. Save
    // only the current existing record, or a sole populated inserted record.
    if(model->isDirty() && model->hasPendingRecord(table->currentIndex().row())) {
        const auto saved=saveDataPage(page,false,true);
        if(saved.value("error").toInt()) return saved;
        if(!self || !page || !table || !model) return tableCommandFailure(2501,"Tabellendokument wurde beim Speichern geschlossen.");
    }
    const auto current=table->currentIndex();
    qint64 target=action==2 ? 0 : action==4 ? static_cast<qint64>(offset)-1
        : action==0 ? static_cast<qint64>(current.row())-offset : static_cast<qint64>(current.row())+offset;
    if((action==0 || action==1) && !current.isValid()) return tableCommandFailure(2105,"Kein aktueller Datensatz.");
    if(target<0 && action!=3) return tableCommandFailure(2105,"Der angeforderte Datensatz ist nicht verfügbar.");
    constexpr int rowLimit=1000000;
    if(target>=rowLimit && action!=3) return tableCommandFailure(3251,"Datensatzziel überschreitet die native Navigationsgrenze.");
    while(model->canFetchMore() && (action==3 || target>=model->rowCount())) {
        if(model->rowCount()>=rowLimit) return tableCommandFailure(3251,"Datensatzmenge überschreitet die native Navigationsgrenze.");
        const int before=model->rowCount(); model->fetchMore();
        if(!self || !table || !model) return tableCommandFailure(2501,"Tabellendokument wurde beim Laden geschlossen.");
        if(model->rowCount()>rowLimit) return tableCommandFailure(3251,"Datensatzmenge überschreitet die native Navigationsgrenze.");
        if(model->rowCount()==before) {
            if(model->lastError().isValid()) return tableModelFailure(model,3146);
            break;
        }
    }
    if(action==3) target=static_cast<qint64>(model->rowCount())-1;
    if(target<0 || target>=model->rowCount()) return tableCommandFailure(2105,"Der angeforderte Datensatz ist nicht verfügbar.");
    const auto destination=model->index(static_cast<int>(target),current.isValid() ? current.column() : 0);
    if(!destination.isValid()) return tableCommandFailure(2105,"Der angeforderte Datensatz ist nicht verfügbar.");
    table->setCurrentIndex(destination); table->scrollTo(destination); table->setFocus();
    return {{"handled",true},{"error",0}};
}

QVariantMap MainWindow::closeTableDocument(int type,const QString& name,int save) {
    const bool active= !activePreviewDocument() && AccessFormPreview::databaseWindowActive(this) && activeDataPage_
        && tabs_->indexOf(activeDataPage_)>=0 && tabs_->currentWidget()==activeDataPage_;
    if(type!=0 && !(type==-1 && active)) return {};
    if(save<0 || save>2) return tableCommandFailure(5,"Ungültige Speicheroption für das Schließen.");
    if(!db_.isOpen() || QApplication::activeModalWidget()) return tableCommandFailure(2046,"Tabelle kann in diesem Kontext nicht geschlossen werden.");
    QPointer<QWidget> page;
    if(name.isEmpty()) { if(active) page=activeDataPage_; }
    else page=findOpenTable(name);
    if(!page || page->property("la_project_object_kind").toString()!="table") return tableCommandFailure(3251,"Die angeforderte offene Tabelle ist nicht verfügbar.");
    return closeTableDataPage(page);
}

QVariantMap MainWindow::closeTableDataPage(QWidget* rawPage) {
    const QPointer<MainWindow> self=this;
    const QPointer<QWidget> page=rawPage;
    QPointer<linaccess::AccessTableView> table;
    for(auto* candidate:page->findChildren<QTableView*>()) if(auto* view=dynamic_cast<linaccess::AccessTableView*>(candidate)) { table=view; break; }
    QPointer<linaccess::AccessTableModel> model=table ? qobject_cast<linaccess::AccessTableModel*>(table->model()) : nullptr;
    if(!table || !model) return tableCommandFailure(2046,"Kein natives Tabellendokument.");
    if(!table->finishEditing()) return tableModelFailure(model,2501);
    if(!self || !page || !model || !table) return tableCommandFailure(2501,"Tabellendokument wurde während der Validierung geschlossen.");
    // Save options concern object design. Pending record data is validated and
    // committed independently, including acSaveNo, before disposing its view.
    if(model->isDirty()) {
        const auto result=saveDataPage(page,false);
        if(result.value("error").toInt()) return result;
    }
    if(!self || !page || !model || !table || tabs_->indexOf(page)<0) return tableCommandFailure(2501,"Tabellendokument ist nicht mehr verfügbar.");
    if(QApplication::activeModalWidget()) return tableCommandFailure(2046,"Schließen ist durch ein modales Fenster blockiert.");
    if(!canCloseDataPage(page)) return tableCommandFailure(2501,"Tabellenänderungen konnten nicht abgeschlossen werden.");
    tabs_->removeTab(tabs_->indexOf(page)); page->hide();
    for(auto* pending:page->findChildren<QSqlTableModel*>()) pending->clear();
    page->deleteLater();
    return {{"handled",true},{"error",0}};
}

QWidget* MainWindow::findOpenTable(const QString& name) const {
    for(int index=0;index<tabs_->count();++index) {
        auto* page=tabs_->widget(index);
        if(page->property("la_project_object_kind").toString()!="table") continue;
        for(auto* candidate:page->findChildren<QTableView*>()) if(auto* view=dynamic_cast<linaccess::AccessTableView*>(candidate)) {
            const auto* model=qobject_cast<linaccess::AccessTableModel*>(view->model());
            if(model && model->tableName().compare(name,Qt::CaseInsensitive)==0) return page;
        }
    }
    return nullptr;
}

int MainWindow::openTableDocument(const QString& name,int view,int mode) {
    if(!db_.isOpen() || QApplication::activeModalWidget()) return 2046;
    if(name.isEmpty() || name.contains(QChar(0))) return 5;
    if(view!=0 || (mode!=1 && mode!=2)) return 3251;
    QString canonical;
    for(const auto& candidate:tableNames()) if(candidate.compare(name,Qt::CaseInsensitive)==0) { canonical=candidate; break; }
    if(canonical.isEmpty()) return 7874;
    return showTable(canonical,mode==2,false);
}

int MainWindow::activateOpenQuery(const QString& name) {
    if(!db_.isOpen() || QApplication::activeModalWidget()) return 2046;
    if(name.isEmpty()) return 450;
    auto* page=findOpenQuery(name);
    if(!page) return 3251;
    return activateDataPage(page) ? 0 : 2501;
}

int MainWindow::activateOpenTable(const QString& name) {
    if(!db_.isOpen() || QApplication::activeModalWidget()) return 2046;
    if(name.isEmpty()) return 450;
    auto* page=findOpenTable(name);
    if(!page) return 3251; // Closed-table selection/error equivalence remains unsupported.
    return activateDataPage(page) ? 0 : 2501;
}

bool MainWindow::activateDataPage(QWidget* page) {
    const QPointer<QWidget> selected=page;
    if(!db_.isOpen() || !selected || tabs_->indexOf(selected)<0 || QApplication::activeModalWidget() || !finishActiveDataEditors()) return false;
    if(!AccessFormPreview::activateProjectDocument(db_,this,nullptr) || !selected || tabs_->indexOf(selected)<0) return false;
    tabs_->setCurrentWidget(selected); activeDataPage_=selected;
    if(isMinimized()) showNormal(); else show();
    raise(); activateWindow();
    for(auto* view:selected->findChildren<QTableView*>()) {
        view->setFocus(Qt::ShortcutFocusReason); return true;
    }
    selected->setFocus(Qt::ShortcutFocusReason); return true;
}

void MainWindow::cycleDocument(bool backwards) {
    if(!db_.isOpen() || QApplication::activeModalWidget()) return;
    QVector<QPointer<QWidget>> pages;
    for(int index=0;index<tabs_->count();++index) pages.append(tabs_->widget(index));
    const auto documents=AccessFormPreview::projectDocuments(db_);
    int current=0;
    if(auto* native=activePreviewDocument()) {
        const int index=documents.indexOf(native); if(index<0) return;
        current=1+pages.size()+index;
    } else if(activeDataPage_) {
        const int index=pages.indexOf(activeDataPage_); if(index>=0) current=1+index;
    }
    const int count=1+pages.size()+documents.size();
    const int target=(current+(backwards ? count-1 : 1))%count;
    if(target>0 && target<=pages.size()) { if(pages[target-1]) (void)activateDataPage(pages[target-1]); return; }
    if(!finishActiveDataEditors()) return;
    if(target==0) { (void)AccessFormPreview::activateProjectDocument(db_,this,nullptr); return; }
    const auto selected=documents[target-1-pages.size()];
    if(selected) (void)AccessFormPreview::activateProjectDocument(db_,this,selected);
}

void MainWindow::focusDatabaseWindow() {
    if(!db_.isOpen() || QApplication::activeModalWidget() || !finishActiveDataEditors()) return;
    AccessFormPreview::activateDatabaseWindow(db_,this);
    if(!AccessFormPreview::databaseWindowActive(this)) return; // A handler may activate another document.
    activeDataPage_.clear();
    if(isMinimized()) showNormal(); else show();
    raise(); activateWindow(); explorer_->show();
    if(auto* splitter=qobject_cast<QSplitter*>(explorer_->parentWidget())) {
        auto sizes=splitter->sizes();
        if(sizes.size()>1 && sizes[0]==0) { sizes[0]=qMax(200,splitter->width()/5); sizes[1]=qMax(1,splitter->width()-sizes[0]); splitter->setSizes(sizes); }
    }
    explorer_->setFocus(Qt::ShortcutFocusReason);
}

QVariantMap MainWindow::outputDatasheet(bool pdf,const QString& path) { return outputDatasheetInternal(pdf ? 1 : 0,path); }
QVariantMap MainWindow::previewDatasheet() { return outputDatasheetInternal(2,{}); }
QVariantMap MainWindow::setupDatasheetPage() { return outputDatasheetInternal(3,{}); }
QVariantMap MainWindow::outputDatasheetInternal(int mode,const QString& requestedPath) {
    const bool pdf=mode==1;
    if(datasheetOutputActive_ || queryExecutionDepth_ || QApplication::activeModalWidget() || !db_.isOpen()
        || activePreviewDocument() || !AccessFormPreview::databaseWindowActive(this)
        || !activeDataPage_ || tabs_->currentWidget()!=activeDataPage_)
        return tableCommandFailure(2046,"Kein aktives Datenblatt für die Ausgabe verfügbar.");
    const QPointer<MainWindow> self=this;
    const QPointer<QWidget> page=activeDataPage_;
    const QPointer<QTableView> view=page->findChild<QTableView*>();
    const auto generation=property("la_project_generation");
    const auto identity=db_.databaseName();
    const auto alive=[self,page,view,generation,identity] {
        return self && page && view && self->db_.isOpen() && self->db_.databaseName()==identity
            && self->property("la_project_generation")==generation && self->tabs_->indexOf(page)>=0
            && self->tabs_->currentWidget()==page && self->activeDataPage_==page
            && !activePreviewDocument() && AccessFormPreview::databaseWindowActive(self)
            && !linaccess::vbaExecutionStopRequested();
    };
    if(!view) return tableCommandFailure(2046,"Keine Datenblattansicht verfügbar.");
    datasheetOutputActive_=true; ++queryExecutionDepth_;
    const auto finished=qScopeGuard([self] { if(self) { self->datasheetOutputActive_=false; --self->queryExecutionDepth_; } });
    try {
        if(auto* table=dynamic_cast<linaccess::AccessTableView*>(view.data())) {
            if(!table->finishEditing()) return tableModelFailure(qobject_cast<linaccess::AccessTableModel*>(table->model()),2501);
            if(!alive()) return tableCommandFailure(2501,"Datenblatt wurde geschlossen.");
            const auto source=validateEditableQuerySource(page,qobject_cast<linaccess::AccessTableModel*>(table->model()));
            if(source.value("error").toInt()) return source;
        }
        QSet<int> selectedRows;
        if(mode==0 && view->selectionModel()) for(const auto& index:view->selectionModel()->selectedRows()) selectedRows.insert(index.row());
        const auto snapshot=linaccess::nativeDatasheetSnapshot(*view,page->property("la_project_object_name").toString());
        auto selected=snapshot; selected.rows.clear(); selected.sourceRows.clear();
        for(int i=0;i<snapshot.rows.size();++i) if(selectedRows.contains(snapshot.sourceRows[i])) {
            selected.rows.append(snapshot.rows[i]); selected.sourceRows.append(snapshot.sourceRows[i]);
        }
        const QPointer<QAbstractItemModel> capturedModel=view->model();
        if(!alive()) return tableCommandFailure(2501,"Datenblatt wurde während der Vorbereitung geschlossen.");
        const auto kind=page->property("la_project_object_kind").toString(),name=page->property("la_project_object_name").toString();
        const auto stored=linaccess::datasheet_page_store::read(db_,kind,name);
        QPageLayout layout; bool repair=false;
        if(page->property("la_datasheet_page_layout").isValid()) layout=page->property("la_datasheet_page_layout").value<QPageLayout>();
        else try { layout=linaccess::datasheet_page_store::decode(stored); }
        catch(const linaccess::VbaRuntimeError&) { if(mode!=3) throw; layout=linaccess::datasheet_page_store::defaults(); repair=true; }
        if(mode==3) {
            QPointer<NativeReportPageSetupDialog> dialog=new NativeReportPageSetupDialog(layout,this,!name.isEmpty() && (kind=="table" || kind=="query"));
            dialog->setObjectName("native-datasheet-page-setup");
            if(repair) dialog->setWindowTitle("Seite einrichten – Einstellungen reparieren");
            auto* save=dialog->findChild<QCheckBox*>("report-page-save"); save->setText("Im Projekt speichern"); save->setToolTip("Für benannte Tabellen und Abfragen verfügbar.");
            const int result=dialog->exec(); const auto selected=dialog ? dialog->pageLayout() : layout;
            const bool persist=dialog && dialog->saveToDesign(); if(dialog) delete dialog;
            if(!alive() || view->model()!=capturedModel) return tableCommandFailure(2501,"Datenblatt wurde während der Seiteneinrichtung geschlossen.");
            if(result!=QDialog::Accepted) return tableCommandFailure(2501,"Seiteneinrichtung abgebrochen.");
            if(persist) linaccess::datasheet_page_store::save(db_,kind,name,stored,selected);
            if(!alive() || view->model()!=capturedModel) return tableCommandFailure(2501,"Datenblatt wurde beim Speichern der Seiteneinrichtung geschlossen.");
            page->setProperty("la_datasheet_page_layout",QVariant::fromValue(selected));
        } else if(mode==2) {
            // Destroy the preview widget before its stack-owned printer.
            QPrinter printer(QPrinter::HighResolution); printer.setOutputFormat(QPrinter::PdfFormat);
            if(!printer.setPageLayout(layout)) throw linaccess::VbaRuntimeError(3251,"Seitenlayout kann nicht übertragen werden.");
            QPointer<QDialog> dialog=new QDialog(this); dialog->setObjectName("native-datasheet-preview");
            dialog->setWindowTitle("Datenblatt-Seitenansicht — "+snapshot.title); dialog->resize(900,700);
            auto* layout=new QVBoxLayout(dialog); auto* navigation=new QHBoxLayout;
            auto* number=new QSpinBox(dialog); number->setObjectName("datasheet-preview-page"); number->setRange(1,1); number->setKeyboardTracking(false);
            auto* total=new QLabel(dialog); total->setObjectName("datasheet-preview-total");
            navigation->addWidget(number); navigation->addWidget(total); navigation->addStretch();
            auto* smaller=addButton(navigation,"Verkleinern"); auto* larger=addButton(navigation,"Vergrößern"); layout->addLayout(navigation);
            auto* pages=new QPrintPreviewWidget(&printer,dialog); pages->setObjectName("datasheet-preview-pages");
            pages->setZoomMode(QPrintPreviewWidget::FitInView); layout->addWidget(pages,1);
            auto* close=new QDialogButtonBox(QDialogButtonBox::Close,dialog); layout->addWidget(close);
            connect(close,&QDialogButtonBox::rejected,dialog,&QDialog::reject);
            QVariantMap failure;
            connect(pages,&QPrintPreviewWidget::paintRequested,dialog,[&snapshot,&failure,alive,pages,total](QPrinter* target) {
                pages->setProperty("la_datasheet_pages_valid",false);
                if(!failure.isEmpty()) return;
                try {
                    if(!alive()) throw linaccess::VbaRuntimeError(2501,"Datenblatt wurde während der Seitenansicht geschlossen.");
                    linaccess::renderNativeDatasheet(snapshot,*target);
                    pages->setProperty("la_datasheet_pages_valid",true);
                } catch(const linaccess::VbaRuntimeError& error) {
                    failure=tableCommandFailure(error.number,QString::fromUtf8(error.what()));
                    total->setText(failure.value("description").toString()); pages->setEnabled(false);
                }
            });
            connect(pages,&QPrintPreviewWidget::previewChanged,dialog,[pages,number,total,&failure] {
                if(!failure.isEmpty()) return;
                const QSignalBlocker blocked(number); number->setMaximum(qMax(1,pages->pageCount())); number->setValue(qMax(1,pages->currentPage()));
                total->setText("von "+QString::number(pages->pageCount()));
            });
            connect(number,&QSpinBox::valueChanged,pages,&QPrintPreviewWidget::setCurrentPage);
            connect(smaller,&QPushButton::clicked,pages,[pages] { pages->zoomOut(); });
            connect(larger,&QPushButton::clicked,pages,[pages] { pages->zoomIn(); });
            dialog->exec(); if(dialog) delete dialog;
            if(!failure.isEmpty()) return failure;
            if(!alive() || view->model()!=capturedModel) return tableCommandFailure(2501,"Datenblatt wurde während der Seitenansicht geschlossen.");
        } else if(pdf) {
            auto path=requestedPath;
            if(path.isEmpty()) {
                // A heap dialog can be destroyed safely together with its owner.
                QPointer<QFileDialog> dialog=new QFileDialog(this,"Datenblatt als PDF speichern",{},"PDF (*.pdf)");
                dialog->setAcceptMode(QFileDialog::AcceptSave); dialog->setDefaultSuffix("pdf");
                const int accepted=dialog->exec();
                if(dialog && accepted==QDialog::Accepted && !dialog->selectedFiles().isEmpty()) path=dialog->selectedFiles().front();
                if(dialog) delete dialog;
                if(path.isEmpty()) return tableCommandFailure(2501,"PDF-Ausgabe abgebrochen.");
            }
            if(!alive() || view->model()!=capturedModel) return tableCommandFailure(2501,"Datenblatt ist nach dem Dialog nicht mehr verfügbar.");
            if(QFileInfo(path).canonicalFilePath()==QFileInfo(identity).canonicalFilePath())
                return tableCommandFailure(75,"Die geöffnete Projektdatei darf nicht durch PDF ersetzt werden.");
            linaccess::exportNativeDatasheetPdf(snapshot,path,layout);
        } else {
            QPrinter printer(QPrinter::HighResolution);
            const bool layoutApplied=printer.setPageLayout(layout);
            QPointer<QPrintDialog> dialog=new QPrintDialog(&printer,this);
            dialog->setOption(QAbstractPrintDialog::PrintSelection,!selected.rows.isEmpty());
            dialog->setOption(QAbstractPrintDialog::PrintPageRange,true);
            const int accepted=dialog->exec(); if(dialog) delete dialog;
            if(accepted!=QDialog::Accepted) return tableCommandFailure(2501,"Druckausgabe abgebrochen.");
            if(!alive() || view->model()!=capturedModel) return tableCommandFailure(2501,"Datenblatt ist nach dem Dialog nicht mehr verfügbar.");
            if(!layoutApplied && !printer.setPageLayout(layout)) throw linaccess::VbaRuntimeError(3251,"Seitenlayout kann nicht übertragen werden.");
            if(printer.printRange()!=QPrinter::AllPages && printer.printRange()!=QPrinter::PageRange && printer.printRange()!=QPrinter::Selection)
                return tableCommandFailure(3251,"Dieser Druckbereich wird noch nicht unterstützt.");
            if(printer.printRange()==QPrinter::Selection && selected.rows.isEmpty())
                return tableCommandFailure(3251,"Keine vollständig markierten sichtbaren Datensätze für die Ausgabe vorhanden.");
            const auto& output=printer.printRange()==QPrinter::Selection ? selected : snapshot;
            const int copies=printer.copyCount(); const bool collated=printer.collateCopies();
            if(copies<1 || copies>100) return tableCommandFailure(3251,"Datenblatt unterstützt 1 bis 100 Druckexemplare.");
            const int first=printer.printRange()==QPrinter::PageRange ? printer.fromPage() : 1;
            const int last=printer.printRange()==QPrinter::PageRange ? printer.toPage() : 0;
            const bool reverse=printer.pageOrder()==QPrinter::LastPageFirst;
            if(printer.outputFormat()==QPrinter::PdfFormat) {
                const auto path=printer.outputFileName();
                if(path.isEmpty() || QFileInfo(path).canonicalFilePath()==QFileInfo(identity).canonicalFilePath())
                    return tableCommandFailure(75,"Kein zulässiges PDF-Ziel ausgewählt.");
                linaccess::exportNativeDatasheetPdf(output,path,printer.pageLayout(),first,last,reverse,copies,collated);
            } else {
                if(!printer.outputFileName().isEmpty()) return tableCommandFailure(3251,"Native Druckdateien außer PDF werden noch nicht unterstützt.");
                const int renderedCopies=printer.supportsMultipleCopies() ? 1 : copies;
                if(renderedCopies>1) printer.setCopyCount(1);
                linaccess::renderNativeDatasheet(output,printer,first,last,reverse,renderedCopies,collated);
                if(printer.printerState()==QPrinter::Error) return tableCommandFailure(3146,"Druckgerät meldet einen Fehler.");
            }
        }
        return {{"handled",true},{"error",0}};
    } catch(const linaccess::VbaRuntimeError& error) {
        return tableCommandFailure(error.number,QString::fromUtf8(error.what()));
    }
}

bool MainWindow::eventFilter(QObject* watched,QEvent* event) {
    if(!tabs_) return QMainWindow::eventFilter(watched,event);
    if(watched==tabs_->tabBar() && (event->type()==QEvent::MouseButtonPress || event->type()==QEvent::MouseButtonRelease)) {
        const int index=tabs_->tabBar()->tabAt(static_cast<QMouseEvent*>(event)->position().toPoint());
        if(index>=0) {
            if(activePreviewDocument()) {
                if(!AccessFormPreview::activateProjectDocument(db_,this,nullptr)) return true;
            } else if(index!=tabs_->currentIndex() && !finishActiveDataEditors()) return true;
        }
    }
    if((watched==tabs_ || watched==tabs_->tabBar()) && event->type()==QEvent::KeyPress) {
        const auto* key=static_cast<QKeyEvent*>(event);
        const auto modifiers=key->modifiers() & ~Qt::KeypadModifier;
        const bool control=modifiers==Qt::ControlModifier || modifiers==(Qt::ControlModifier|Qt::ShiftModifier);
        const bool changeTab=(control && (key->key()==Qt::Key_Tab || key->key()==Qt::Key_Backtab || key->key()==Qt::Key_PageUp || key->key()==Qt::Key_PageDown))
            || (watched==tabs_->tabBar() && modifiers==Qt::NoModifier && (key->key()==Qt::Key_Left || key->key()==Qt::Key_Right || key->key()==Qt::Key_Home || key->key()==Qt::Key_End));
        if(changeTab && !finishActiveDataEditors()) return true;
    }
    if(event->type()==QEvent::FocusIn && db_.isOpen()) {
        const QPointer<MainWindow> self=this;
        if(watched==explorer_) {
            AccessFormPreview::activateDatabaseWindow(db_,this);
            if(!self) return true;
            if(AccessFormPreview::databaseWindowActive(this)) activeDataPage_.clear();
        } else if(auto* widget=qobject_cast<QWidget*>(watched);widget && (dynamic_cast<linaccess::AccessTableView*>(widget)
            || widget->property("la_cell_index").isValid() || widget==tabs_ || widget==tabs_->tabBar())) {
            QPointer<QWidget> guarded=widget;
            QPointer<QWidget> page=widget;
            while(page && tabs_->indexOf(page)<0) page=page->parentWidget();
            if(!page && (widget==tabs_ || widget==tabs_->tabBar())) page=tabs_->currentWidget();
            if(page) {
                AccessFormPreview::activateDatabaseWindow(db_,this);
                if(!self || !guarded) return true;
                if(page && tabs_->indexOf(page)>=0 && !activePreviewDocument()) activeDataPage_=page;
            }
        }
    }
    return QMainWindow::eventFilter(watched,event);
}

void MainWindow::keyPressEvent(QKeyEvent* event) {
    if(event->key()==Qt::Key_P && event->modifiers()==Qt::ControlModifier) {
        const QPointer<MainWindow> self=this;
        const auto result=outputDatasheet(false);
        if(self) if(const int error=result.value("error").toInt()) self->statusBar()->showMessage(result.value("description").toString()+" ("+QString::number(error)+")",10000);
        event->accept(); return;
    }
    if(event->key()==Qt::Key_F6 && (event->modifiers()==Qt::ControlModifier || event->modifiers()==(Qt::ControlModifier|Qt::ShiftModifier))) {
        AccessFormPreview::cycleProjectDocuments(db_,this,event->modifiers().testFlag(Qt::ShiftModifier)); event->accept(); return;
    }
    if(event->key()==Qt::Key_F11 && event->modifiers()==Qt::NoModifier) { focusDatabaseWindow(); event->accept(); return; }
    if((event->key()==Qt::Key_O || event->key()==Qt::Key_N) && event->modifiers()==Qt::ControlModifier) {
        if(!QApplication::activeModalWidget()) { if(event->key()==Qt::Key_O) chooseDatabase(); else newDatabase(); }
        event->accept(); return;
    }
    QMainWindow::keyPressEvent(event);
}

bool MainWindow::hasProject() {
    if (db_.isOpen()) return true;
    QMessageBox::information(this, "Projekt erforderlich", "Zuerst ein SQLite-Projekt öffnen oder erstellen.");
    return false;
}

void MainWindow::newDatabase() {
    const auto path = QFileDialog::getSaveFileName(this, "Projekt erstellen", "datenbank.linaccess", "LinAccess (*.linaccess)");
    if (path.isEmpty()) return;
    if (QFileInfo::exists(path)) {
        messageError(this, "Diese Datei existiert bereits und wird nicht überschrieben.");
        return;
    }
    openDatabase(path);
}

void MainWindow::chooseDatabase() {
    const auto path = QFileDialog::getOpenFileName(this, "Projekt öffnen", {},
                                                  "LinAccess / SQLite (*.linaccess *.sqlite *.db);;Alle Dateien (*)");
    if (!path.isEmpty()) openDatabase(path);
}

void MainWindow::openDatabase(const QString& path) {
    if (path.endsWith(".mdb", Qt::CaseInsensitive) || path.endsWith(".accdb", Qt::CaseInsensitive)) {
        messageError(this, "Access-Dateien bitte über 'Access untersuchen' öffnen. "
                           "Sie dürfen nicht als SQLite-Projekt geöffnet werden.");
        return;
    }
    const bool bypassStartup=QApplication::keyboardModifiers().testFlag(Qt::ShiftModifier);
    const auto projectPath=QFileInfo(path).absoluteFilePath();
    if (!AccessFormPreview::closeProjectDocuments(db_)) return;
    setProperty("la_project_generation",QVariant::fromValue<qulonglong>(property("la_project_generation").toULongLong()+1));
    setProperty("la_native_current_projects",QVariant{});
    setProperty("la_native_module_applications",QVariant{});
    setProperty("la_startup_macro_status","none");
    while (tabs_->count()) {
        auto* widget = tabs_->widget(0);
        tabs_->removeTab(0);
        delete widget;
    }
    linaccess::resetVbaProjectState(db_);
    db_.close();
    db_.setDatabaseName(projectPath);
    if (!db_.open()) {
        messageError(this, "SQLite: " + db_.lastError().text());
        return;
    }
    connectPreviewSqlObjects(db_);
    QSqlQuery query(db_);
    if (!query.exec("PRAGMA foreign_keys=ON") ||
        !query.exec("CREATE TABLE IF NOT EXISTS _la_objects ("
                    "kind TEXT NOT NULL, name TEXT NOT NULL, definition TEXT NOT NULL, "
                    "PRIMARY KEY (kind,name))") ||
        !query.exec("CREATE TABLE IF NOT EXISTS _la_import_log ("
                    "id INTEGER PRIMARY KEY AUTOINCREMENT, source_file TEXT, kind TEXT, "
                    "name TEXT, status TEXT, details TEXT)") ||
        !query.exec("CREATE TABLE IF NOT EXISTS _la_access_artifacts ("
                    "id INTEGER PRIMARY KEY AUTOINCREMENT, "
                    "source_sha256 TEXT NOT NULL, source_filename TEXT NOT NULL, "
                    "kind TEXT NOT NULL, name TEXT NOT NULL, status TEXT NOT NULL, "
                    "format TEXT NOT NULL, detail TEXT NOT NULL, source_bytes BLOB, "
                    "UNIQUE(source_sha256,kind,name,format))")) {
        messageError(this, "Projektinitialisierung: " + query.lastError().text());
        db_.close();
        return;
    }
    linaccess::installAccessQueryViews(db_,true);
    databasePath_ = projectPath;
    setWindowTitle("LinAccess — " + QFileInfo(projectPath).fileName());
    statusBar()->showMessage(projectPath);
    refreshExplorer();
    if(explorer_->hasFocus()) AccessFormPreview::activateDatabaseWindow(db_,this);
    try {
        const auto startup=linaccess::nativeMacroCatalogName(db_,"AutoExec");
        if(!startup.isEmpty()) {
            if(bypassStartup) {setProperty("la_startup_macro_status","bypassed");statusBar()->showMessage("AutoExec mit Umschalttaste übersprungen.");}
            else {
                const auto outcome=runProjectMacro(startup);
                if(const int error=outcome.value("error").toInt()) throw linaccess::VbaRuntimeError(error,outcome.value("description").toString().toStdString());
                const bool stopped=outcome.value("stopped").toBool();setProperty("la_startup_macro_status",stopped ? "stopped" : "completed");statusBar()->showMessage(stopped ? "AutoExec beendet." : "AutoExec ausgeführt.");
            }
        }
    }catch(const linaccess::VbaRuntimeError& error) {setProperty("la_startup_macro_status","failed");statusBar()->showMessage(QString("AutoExec fehlgeschlagen: Fehler %1").arg(error.number));messageError(this,QString("AutoExec: Fehler %1: %2").arg(error.number).arg(QString::fromUtf8(error.what())));}
}

QStringList MainWindow::tableNames() {
    QStringList names;
    if (!db_.isOpen()) return names;
    QSqlQuery query(db_);
    if (query.exec("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name COLLATE NOCASE")) {
        while (query.next()) {
            const auto name = query.value(0).toString();
            if (!name.startsWith("_la_", Qt::CaseInsensitive) &&
                !name.startsWith("sqlite_", Qt::CaseInsensitive)) names << name;
        }
    }
    return names;
}

QStringList MainWindow::objectNames(const QString& kind) {
    QStringList names;
    QSqlQuery query(db_);
    query.prepare("SELECT name FROM main._la_objects WHERE kind=? ORDER BY name COLLATE NOCASE");
    query.addBindValue(kind);
    if (query.exec()) while (query.next()) names << query.value(0).toString();
    return names;
}

QJsonObject MainWindow::objectDefinition(const QString& kind, const QString& name) {
    QSqlQuery query(db_);
    query.prepare("SELECT definition FROM main._la_objects WHERE kind=? AND name=?");
    query.addBindValue(kind);
    query.addBindValue(name);
    if (query.exec() && query.next())
        return QJsonDocument::fromJson(query.value(0).toByteArray()).object();
    return {};
}

bool MainWindow::saveObject(const QString& kind, const QString& name, const QJsonObject& definition) {
    if (name.trimmed().isEmpty()) {
        messageError(this, "Ein Name ist erforderlich.");
        return false;
    }
    QSqlQuery query(db_);
    query.prepare("INSERT INTO main._la_objects(kind,name,definition) VALUES(?,?,?) "
                  "ON CONFLICT(kind,name) DO UPDATE SET definition=excluded.definition");
    query.addBindValue(kind);
    query.addBindValue(name.trimmed());
    query.addBindValue(QString::fromUtf8(QJsonDocument(definition).toJson(QJsonDocument::Compact)));
    if (!query.exec()) {
        messageError(this, query.lastError().text());
        return false;
    }
    refreshExplorer();
    return true;
}

void MainWindow::refreshExplorer() {
    const auto* previous=explorer_->currentItem();const bool sameProject=explorer_->property("la_explorer_project_generation")==property("la_project_generation");
    const auto previousKind=previous && sameProject ? previous->data(0,Qt::UserRole).toString() : QString{};
    const auto previousFamily=previous && sameProject ? previous->data(0,Qt::UserRole+2).toString() : QString{};
    const auto previousName=previous && sameProject ? previous->data(0,Qt::UserRole+3).toString() : QString{};
    const int previousArtifact=previous && sameProject ? previous->data(0,Qt::UserRole+1).toInt() : 0;
    const bool previousGroup=previous && sameProject && previous->data(0,Qt::UserRole+4).toBool();
    explorer_->clear();explorer_->setProperty("la_explorer_project_generation",property("la_project_generation"));
    if (!db_.isOpen()) return;
    QMap<QString,QTreeWidgetItem*> groups;
    const auto group = [this,&groups](const QString& label, const QString& kind, const QStringList& names) {
        auto* parent = new QTreeWidgetItem(explorer_, {label});
        groups.insert(kind,parent); parent->setData(0,Qt::UserRole+2,kind); parent->setData(0,Qt::UserRole+4,true);
        for (const auto& name : names) {
            auto* item = new QTreeWidgetItem(parent, {name});
            item->setData(0, Qt::UserRole, kind);
            item->setData(0,Qt::UserRole+2,kind); item->setData(0,Qt::UserRole+3,name);
        }
        parent->setExpanded(true);
    };
    group("Tabellen", "table", tableNames());
    group("Abfragen", "query", objectNames("query"));
    group("Formulare", "form", objectNames("form"));
    group("Berichte", "report", objectNames("report"));
    group("Makros", "macro", objectNames("macro"));
    group("Module", "module", objectNames("module"));
    QSqlQuery artifacts(db_);
    if (artifacts.exec("SELECT id,kind,name,status FROM _la_access_artifacts "
                       "ORDER BY kind,name COLLATE NOCASE")) {
        auto* accessGroup = new QTreeWidgetItem(explorer_, {"Weitere Access-Quellobjekte"});
        while (artifacts.next()) {
            const auto kind=artifacts.value(1).toString();
            const bool selectable=kind=="form" || kind=="report" || kind=="query";
            auto* item = new QTreeWidgetItem(selectable ? groups.value(kind) : accessGroup, {
                artifacts.value(1).toString() + ": " + artifacts.value(2).toString()
                + " [" + artifacts.value(3).toString() + "]"});
            item->setData(0, Qt::UserRole, "access_artifact");
            item->setData(0, Qt::UserRole + 1, artifacts.value(0).toInt());
            if(selectable) { item->setData(0,Qt::UserRole+2,kind); item->setData(0,Qt::UserRole+3,artifacts.value(2).toString()); }
        }
        accessGroup->setExpanded(true);
    }
    if(previousKind.isEmpty() && !previousGroup) return;
    const std::function<void(QTreeWidgetItem*)> restore=[&](QTreeWidgetItem* item) {
        bool match=false;
        if(previousGroup) match=item->data(0,Qt::UserRole+4).toBool() && item->data(0,Qt::UserRole+2).toString()==previousFamily;
        else if(previousKind=="access_artifact") match=item->data(0,Qt::UserRole).toString()==previousKind && item->data(0,Qt::UserRole+1).toInt()==previousArtifact;
        else match=item->data(0,Qt::UserRole).toString()==previousKind && item->data(0,Qt::UserRole+3).toString()==previousName;
        if(match) {explorer_->setCurrentItem(item);explorer_->scrollToItem(item);return;}
        for(int index=0;index<item->childCount();++index) restore(item->child(index));
    };
    for(int index=0;index<explorer_->topLevelItemCount();++index) restore(explorer_->topLevelItem(index));

}

void MainWindow::openExplorerItem(QTreeWidgetItem* item, int) {
    if (!item || !hasProject()) return;
    const auto kind = item->data(0, Qt::UserRole).toString();
    const auto name = item->text(0);
    if (kind == "table") showTable(name);
    else if (kind == "query") editQuery(name);
    else if (kind == "form") runForm(name);
    else if (kind == "report") runReport(name);
    else if (kind == "macro") editMacro(name);
    else if (kind == "module") editModule(name);
    else if (kind == "access_artifact") showAccessArtifact(
        item->data(0, Qt::UserRole + 1).toInt());
}

void MainWindow::inspectAccess() {
    const auto path = QFileDialog::getOpenFileName(this, "Access-Datei auswählen", {},
                                                  "Microsoft Access (*.mdb *.accdb)");
    if (path.isEmpty()) return;
    AccessInspector dialog(path, db_, this);
    dialog.exec();
    refreshExplorer();
}

void MainWindow::inspectBridgeExport() {
    const auto directory = QFileDialog::getExistingDirectory(this, "Ordner mit Bridge-manifest.json auswählen");
    if (directory.isEmpty()) return;
    BridgeImporter dialog(directory, db_, this);
    dialog.exec();
    refreshExplorer();
}

void MainWindow::showAccessArtifact(int id) {
    QSqlQuery query(db_);
    query.prepare("SELECT kind,name,status,format,detail,source_bytes FROM _la_access_artifacts WHERE id=?");
    query.addBindValue(id);
    if (!query.exec() || !query.next()) {
        messageError(this, "Quellobjekt nicht gefunden");
        return;
    }
    QDialog dialog(this);
    dialog.setObjectName("access-artifact-details");
    const auto artifactName=query.value(1).toString();
    dialog.setWindowTitle(query.value(0).toString() + ": " + artifactName);
    dialog.resize(850, 680);
    auto* outer = new QVBoxLayout(&dialog);
    const auto kind = query.value(0).toString();
    if(kind=="query") tagProjectObject(&dialog,this,db_,kind,artifactName);
    const auto status = query.value(2).toString();
    const auto format=query.value(3).toString();
    const auto detail=query.value(4).toString();
    const auto bytes=query.value(5).toByteArray();
    query.finish();
    outer->addWidget(new QLabel("Status: " + status
        + " | Format: " + format
        + " | Access-Layouts und vollständiges VBA sind NICHT kompatibel."));
    if (!detail.isEmpty()) outer->addWidget(new QLabel(detail));
    auto* source = new QPlainTextEdit(&dialog);
    source->setReadOnly(true);
    source->setPlainText((kind == "vba"
        ? linaccess::decodeVbaSource(bytes.left(512 * 1024))
        : QString::fromUtf8(bytes.left(512 * 1024)))
        + (bytes.size() > 512 * 1024 ? "\n\n[Vorschau gekürzt; Originalbytes bleiben gespeichert]" : ""));
    outer->addWidget(source);
    if (kind == "query" && status == "source_only" && !bytes.isEmpty()) {
        auto* execute = new QPushButton("Abfrage ausführen", &dialog);
        outer->addWidget(execute);
        const auto text = linaccess::decodeVbaSource(bytes).trimmed();
        connect(execute, &QPushButton::clicked, &dialog, [this, text, &dialog] {
            try {
                const auto plan = linaccess::compileAccessQuery(text);
                QMap<QString,QVariant> parameters;
                for (const auto& parameter : plan.parameters) {
                    bool accepted = false;
                    const auto value = QInputDialog::getText(&dialog,"Abfrageparameter",
                        parameter.name + " (" + parameter.type + ")",QLineEdit::Normal,{},&accepted);
                    if (!accepted) return;
                    parameters.insert(parameter.name.toCaseFolded(),value);
                }
                if (isSelect(plan.sql)) {
                    QDialog result(&dialog);
                    result.setWindowTitle("Abfrageergebnis");
                    result.resize(900,620);
                    auto* layout = new QVBoxLayout(&result);
                    auto* model = new linaccess::AccessQueryModel(&result);
                    model->execute(db_,text,parameters);
                    auto* table = new QTableView(&result);
                    table->setModel(model);
                    layout->addWidget(table);
                    result.exec();
                } else {
                    auto database = linaccess::makeDaoDatabase(db_);
                    const auto definition = database->member("CreateQueryDef",{
                        linaccess::VbaValue::string(""),linaccess::VbaValue::string(text.toUtf8().toStdString())});
                    for (auto it = parameters.cbegin(); it != parameters.cend(); ++it) {
                        const auto parameter = definition.object->member("Parameters",{linaccess::VbaValue::string(it.key().toUtf8().toStdString())});
                        parameter.object->assign("Value",linaccess::qtFieldValue(it.value()));
                    }
                    definition.object->member("Execute",{});
                    QMessageBox::information(&dialog,"Abfrage ausgeführt","Die Projektkopie wurde aktualisiert.");
                }
            } catch (const std::exception& error) {
                messageError(&dialog,QString::fromUtf8(error.what()));
            }
        });
    }

    if (kind == "report" && status == "source_only" && !bytes.isEmpty()) {
        auto* open = new QPushButton("Access-Bericht öffnen (Layoutansicht)", &dialog);
        outer->addWidget(open);
        const auto name = artifactName;
        connect(open, &QPushButton::clicked, &dialog, [this, name, bytes, &dialog] {
            try {
                QSqlQuery vba(db_);
                vba.prepare("SELECT source_bytes FROM _la_access_artifacts WHERE kind='vba' AND name=? ORDER BY id DESC LIMIT 1");
                vba.addBindValue("Report_" + name);
                const auto module = vba.exec() && vba.next() ? vba.value(0).toByteArray() : QByteArray{};
                vba.finish(); dialog.accept();
                (void)AccessFormPreview::showDocument(linaccess::AccessDocumentKind::Report,name,bytes,module,db_,this);
            } catch(const std::exception& error) { messageError(this,QString::fromUtf8(error.what())); }
        });
    }
    if (kind == "form" && status == "source_only" && !bytes.isEmpty()) {
        const auto formName = artifactName;
        auto* open = new QPushButton("Access-Formular öffnen (Originalanordnung soweit unterstützt)", &dialog);
        outer->addWidget(open);
        connect(open, &QPushButton::clicked, &dialog, [this, formName, bytes, &dialog] {
            try {
                QSqlQuery vba(db_);
                vba.prepare("SELECT source_bytes FROM _la_access_artifacts WHERE kind='vba' "
                            "AND name=? ORDER BY id DESC LIMIT 1");
                vba.addBindValue("Form_" + formName);
                const auto module = vba.exec() && vba.next() ? vba.value(0).toByteArray() : QByteArray{};
                vba.finish(); dialog.accept();
                (void)AccessFormPreview::showDocument(linaccess::AccessDocumentKind::Form,formName,bytes,module,db_,this);
            } catch(const std::exception& error) { messageError(this,QString::fromUtf8(error.what())); }
        });
        auto* convert = new QPushButton("Als natives Formular übernehmen (Teilkonvertierung)", &dialog);
        outer->addWidget(convert);
        connect(convert, &QPushButton::clicked, &dialog,
                [this, &dialog, formName, bytes, id] {
            const auto draft = linaccess::parseAccessFormText(linaccess::decodeVbaSource(bytes));
            if (draft.recordSource.isEmpty()) {
                messageError(this, "Die Access-Formularquelle hat keine einfache Datenquelle (RecordSource).");
                return;
            }
            if (!tableNames().contains(draft.recordSource, Qt::CaseInsensitive)) {
                messageError(this, "Formular nicht konvertierbar: Die Datenquelle '"
                    + draft.recordSource + "' fehlt im LinAccess-Projekt. "
                      "Die benötigte Access-Tabelle erst separat importieren. "
                      "Für Abfragequellen das native Access-Formular öffnen.");
                return;
            }
            QSqlQuery existing(db_);
            existing.prepare("SELECT 1 FROM main._la_objects WHERE kind='form' AND name=?");
            existing.addBindValue(formName);
            if (!existing.exec()) { messageError(this, existing.lastError().text()); return; }
            if (existing.next()) {
                messageError(this, "Ein natives Formular mit diesem Namen existiert bereits. "
                                   "Es wird nicht überschrieben.");
                return;
            }

            QSqlQuery schema(db_);
            if (!schema.exec("PRAGMA table_info(" + quoteSql(draft.recordSource) + ")")) {
                messageError(this, schema.lastError().text());
                return;
            }
            QStringList available;
            while (schema.next()) available << schema.value(1).toString();
            QJsonArray chosen;
            for (const auto& sourceField : draft.controlSources) {
                for (const auto& column : available) {
                    if (column.compare(sourceField, Qt::CaseInsensitive) == 0
                        && !chosen.contains(column)) {
                        chosen.append(column);
                        break;
                    }
                }
            }
            if (chosen.isEmpty()) {
                messageError(this, "Keine in der SQLite-Tabelle vorhandenen Formularfelder gefunden.");
                return;
            }
            QSqlQuery vb(db_);
            vb.prepare("SELECT source_bytes FROM _la_access_artifacts "
                       "WHERE kind='vba' AND status='source_only' AND name=? "
                       "ORDER BY id DESC LIMIT 1");
            vb.addBindValue("Form_" + formName);
            if (!vb.exec()) { messageError(this, vb.lastError().text()); return; }
            const QByteArray originalVba = vb.next() ? vb.value(0).toByteArray() : QByteArray{};
            const QJsonObject converted = {
                {"table", draft.recordSource},
                {"fields", chosen},
                {"vba", linaccess::decodeVbaSource(originalVba)},
                {"original_vba_base64", QString::fromLatin1(originalVba.toBase64())},
                {"language", "vba_subset_experimental"},
                {"source_form_artifact_id", id},
                {"import_status", "partial_native_form"}
            };
            if (!saveObject("form", formName, converted)) return;
            QMessageBox::information(&dialog, "Formular teilkonvertiert",
                "Ein natives LinAccess-Formular mit " + QString::number(chosen.size())
                + " tatsächlich vorhandenen Tabellenfeldern wurde erstellt. "
                  "Die Originalgeometrie, Unterformulare, Kombinationslisten und "
                  "nicht vorhandene Felder werden NICHT rekonstruiert. "
                + (originalVba.isEmpty() ? "Kein passendes VBA-Modul vorhanden."
                                         : "Das originale VBA-Modul wurde zur eingeschränkten Laufzeit übernommen."));
            dialog.accept();
        });
    }
    if (kind == "vba" && status == "source_only" && !bytes.isEmpty()) {
        const auto text = linaccess::decodeVbaSource(bytes);
        const auto analysis = linaccess::analyzeVba(text.toUtf8().toStdString());
        auto* procedures = new QComboBox(&dialog);
        for (const auto& procedure : analysis.procedures)
            procedures->addItem(QString::fromStdString(procedure.kind) + " "
                + QString::fromStdString(procedure.name), QString::fromStdString(procedure.name));
        auto* run = new QPushButton("VBA-Prozedur ausführen (eingeschränkte Laufzeit)", &dialog);
        run->setEnabled(!analysis.procedures.empty() && bytes.size() <= 1024 * 1024);
        outer->addWidget(procedures);
        outer->addWidget(run);
        auto* runtimeResult = new QLabel(&dialog);
        runtimeResult->setWordWrap(true);
        outer->addWidget(runtimeResult);
        connect(run, &QPushButton::clicked, &dialog, [this, &dialog, text, procedures, runtimeResult] {
            class ModuleHost final : public linaccess::VbaHost {
            public:
                ModuleHost(QWidget* dialog, QSqlDatabase database)
                    : dialog_(dialog), database_(std::move(database)) {}
                std::optional<std::string> field(std::string_view) override {
                    throw std::runtime_error("Me! benötigt einen geöffneten Formularkontext");
                }
                void setField(std::string_view, std::optional<std::string>) override {
                    throw std::runtime_error("Kein Formularkontext für Me! vorhanden");
                }
                void message(std::string_view value) override {
                    QMessageBox::information(dialog_, "VBA MsgBox",
                        QString::fromUtf8(value.data(), static_cast<qsizetype>(value.size())));
                }
                void quitApplication() override { AccessFormPreview::quitApplication(); }
                void saveRecord() override {
                    throw std::runtime_error("DoCmd benötigt einen geöffneten Formularkontext");
                }
                linaccess::VbaValue formatValue(const linaccess::VbaValue& value,std::string_view mask,bool stringResult) override {
                    return linaccess::formatAccessValue(value,mask,stringResult);
                }
                std::shared_ptr<linaccess::VbaObject> currentDatabase() override { return linaccess::makeDaoDatabase(database_); }
                void executeSql(std::string_view sql) override {
                    auto database = linaccess::makeDaoDatabase(database_);
                    (void)database->member("Execute",{linaccess::VbaValue::string(std::string(sql))});
                }
            private:
                QWidget* dialog_;
                QSqlDatabase database_;
            } host(&dialog, db_);
            const auto result = linaccess::executeVbaProcedure(
                text.toUtf8().toStdString(), procedures->currentData().toString().toUtf8().toStdString(), host);
            if (result.status == linaccess::VbaExecutionStatus::Completed)
                runtimeResult->setText("Prozedur ausgeführt (eingeschränkte VBA-Teilmenge)."
                    + (result.returnValue ? " Rückgabewert: "
                        + QString::fromUtf8(result.returnValue->data(), result.returnValue->size()) : ""));
            else runtimeResult->setText("Nicht ausgeführt / Fehler bei Zeile "
                + QString::number(result.line) + ": " + QString::fromStdString(result.explanation));
        });
    }
    dialog.exec();
}

void MainWindow::newTable() {
    if (!hasProject()) return;
    QDialog dialog(this); dialog.setObjectName("native-table-designer");
    dialog.setWindowTitle("Neue Tabelle"); dialog.resize(760,440);
    auto* outer=new QVBoxLayout(&dialog);
    auto* name=new QLineEdit(&dialog); name->setObjectName("native-table-name"); name->setPlaceholderText("Tabellenname"); outer->addWidget(name);
    outer->addWidget(new QLabel("Felder (id wird automatisch als Primärschlüssel angelegt)",&dialog));
    auto* grid=new QTableWidget(0,6,&dialog); grid->setObjectName("native-table-fields");
    grid->setHorizontalHeaderLabels({"Feldname","Datentyp","Pflichtfeld (ja/nein)","Standardwert (optional)","Textgröße (1–255)","Leere Zeichenfolge zulassen"});
    grid->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch); outer->addWidget(grid);
    outer->addWidget(new QLabel("TEXT: wählbare Feldgröße bis 255 Zeichen; MEMO: längerer Text.",&dialog));
    outer->addWidget(new QLabel("Standardwerte: Literal, optional mit =; Datum auch #ISO-Datum#, Date(), Now() oder Time().",&dialog));
    const auto addRow=[grid] {
        const int row=grid->rowCount(); grid->insertRow(row);
        grid->setItem(row,0,new QTableWidgetItem());
        auto* type=new QComboBox(grid); type->addItems({"TEXT","MEMO","BYTE","SMALLINT","INTEGER","SINGLE","REAL","DATE","BOOLEAN","BLOB"}); grid->setCellWidget(row,1,type);
        grid->setItem(row,2,new QTableWidgetItem("nein")); grid->setItem(row,3,new QTableWidgetItem());
        auto* allowZero=new QCheckBox(grid);allowZero->setChecked(true);grid->setCellWidget(row,5,allowZero);
        auto* size=new QTableWidgetItem("255");grid->setItem(row,4,size);
        connect(type,&QComboBox::currentTextChanged,grid,[size,allowZero](const QString& value) {allowZero->setEnabled(value=="TEXT" || value=="MEMO");size->setFlags(value=="TEXT" ? Qt::ItemIsEnabled|Qt::ItemIsSelectable|Qt::ItemIsEditable : Qt::ItemFlags{});});
    };
    auto* plus=addButton(outer,"+ Feld"); plus->setObjectName("native-table-add-field"); connect(plus,&QPushButton::clicked,&dialog,addRow); addRow();
    auto* buttons=new QDialogButtonBox(QDialogButtonBox::Save|QDialogButtonBox::Cancel,&dialog); outer->addWidget(buttons);
    connect(buttons,&QDialogButtonBox::rejected,&dialog,&QDialog::reject);
    connect(buttons,&QDialogButtonBox::accepted,&dialog,[&] {
        const auto tableName=name->text().trimmed();
        if(tableName.isEmpty() || tableName.startsWith("_la_",Qt::CaseInsensitive) || tableName.startsWith("sqlite_",Qt::CaseInsensitive)) {
            messageError(&dialog,"Ungültiger/reservierter Tabellenname"); return;
        }
        struct FieldSpec {QString name,type,defaultValue;bool required;int size;bool allowZero;};QVector<FieldSpec> fields;std::set<QString> seen={"id"};
        for(int row=0;row<grid->rowCount();++row) {
            const auto field=grid->item(row,0)->text().trimmed(); if(field.isEmpty()) continue;
            if(field.startsWith("_la_",Qt::CaseInsensitive) || seen.contains(field.toCaseFolded())) {
                messageError(&dialog,"Reservierter oder doppelter Feldname: "+field); return;
            }
            seen.insert(field.toCaseFolded());
            const auto type=qobject_cast<QComboBox*>(grid->cellWidget(row,1))->currentText();
            int size=0;
            if(type=="TEXT") {bool valid=false;size=grid->item(row,4)->text().trimmed().toInt(&valid);if(!valid || size<1 || size>255) {messageError(&dialog,"Für Textfelder eine Feldgröße von 1 bis 255 eingeben.");return;}}
            const auto required=grid->item(row,2)->text().trimmed().toLower();
            if(required!="ja" && required!="nein" && !required.isEmpty()) { messageError(&dialog,"Pflichtfeld muss ja oder nein sein"); return; }
            const auto token=grid->item(row,3)->text().trimmed(); QString defaultSql;
            if(type=="DATE" && linaccess::nativeDateDefaultSql(token)) defaultSql=" DEFAULT "+token;
            else if(!token.isEmpty()) {
                std::optional<QString> literal;try {if(type=="DATE") literal=linaccess::nativeDateLiteralSql(token);if(!literal) literal=linaccess::accessLiteralDefaultSql(token);}catch(const linaccess::VbaRuntimeError&) {messageError(&dialog,"Datum-Standardwert als gültiges ISO-Datum eingeben, beispielsweise #2026-11-08#.");return;}
                if(!literal || token.size()>255) { messageError(&dialog,"Standardwert: Literal mit höchstens 255 Zeichen eingeben; allgemeine Ausdrücke sind noch nicht unterstützt."); return; }
                QSqlQuery value(db_);
                if(!value.exec("SELECT "+*literal) || !value.next()) { messageError(&dialog,"Standardwert konnte nicht gelesen werden."); return; }
                const auto scalar=value.value(0); defaultSql=*literal;
                if(!scalar.isNull()) {
                    const auto kind=scalar.metaType().id(); const bool numeric=kind==QMetaType::Int || kind==QMetaType::LongLong || kind==QMetaType::Double;
                    if((type=="TEXT" || type=="MEMO") && kind!=QMetaType::QString) { messageError(&dialog,"Text-Standardwert in Anführungszeichen eingeben."); return; }
                    if(type=="TEXT" && scalar.toString().size()>size) {messageError(&dialog,"Text-Standardwert überschreitet die gewählte Feldgröße.");return;}
                    if(type=="INTEGER" || type=="SMALLINT" || type=="BYTE") {
                        if(!numeric || (kind==QMetaType::Double && (std::floor(scalar.toDouble())!=scalar.toDouble() || std::abs(scalar.toDouble())>9007199254740991.0))) {
                            messageError(&dialog,"Ganzzahligen Standardwert eingeben."); return;
                        }
                        try {(void)linaccess::qtFieldValue(scalar,type=="INTEGER" ? "INTEGER4" : type=="SMALLINT" ? "INTEGER2" : "INTEGER1");}
                        catch(const linaccess::VbaRuntimeError&) {messageError(&dialog,"Standardwert liegt außerhalb des gewählten Ganzzahltyps.");return;}
                        defaultSql=QString::number(scalar.toLongLong());
                    }
                    if((type=="SINGLE" || type=="REAL") && !numeric) { messageError(&dialog,"Numerischen Standardwert eingeben."); return; }
                    if(type=="SINGLE") {try {(void)linaccess::sqlFieldValue(linaccess::qtFieldValue(scalar,{}),"SINGLE_REAL");}catch(const linaccess::VbaRuntimeError&) {messageError(&dialog,"Standardwert liegt außerhalb des unterstützten Single-Bereichs.");return;}}
                    if(type=="BOOLEAN") {
                        if(!numeric) { messageError(&dialog,"Boolean-Standardwert als True, False oder Zahl eingeben."); return; }
                        defaultSql=scalar.toDouble()==0 ? "0" : "-1";
                    }
                    if(type=="DATE") {
                        try { (void)linaccess::qtFieldValue(scalar,"DATE"); }
                        catch(const linaccess::VbaRuntimeError&) { messageError(&dialog,"Datum-Standardwert als gültigen ISO-Text in Anführungszeichen eingeben."); return; }
                    }
                    if(type=="BLOB") { messageError(&dialog,"Binärfelder unterstützen hier nur Null als Standardwert."); return; }
                }
                defaultSql=" DEFAULT "+defaultSql;
            }
            fields.append({field,type,defaultSql.isEmpty() ? QString{} : defaultSql.mid(9),required=="ja",size,qobject_cast<QCheckBox*>(grid->cellWidget(row,5))->isChecked()});
        }
        if(fields.isEmpty()) { messageError(&dialog,"Mindestens ein Feld anlegen"); return; }
        try {
            using linaccess::VbaValue;
            auto database=linaccess::makeDaoDatabase(db_);
            auto table=database->member("CreateTableDef",{VbaValue::string(tableName.toStdString())}).object;
            auto catalog=table->member("Fields",{}).object;
            auto id=table->member("CreateField",{VbaValue::string("id"),VbaValue::integer(4)}).object;
            id->assign("Attributes",VbaValue::longInteger(16));catalog->member("Append",{VbaValue::reference(id)});
            for(const auto& spec:fields) {
                const int type=spec.type=="TEXT" ? 10 : spec.type=="MEMO" ? 12 : spec.type=="BYTE" ? 2 : spec.type=="SMALLINT" ? 3 : spec.type=="INTEGER" ? 4 : spec.type=="SINGLE" ? 6 : spec.type=="REAL" ? 7 : spec.type=="DATE" ? 8 : spec.type=="BOOLEAN" ? 1 : 11;
                auto field=table->member("CreateField",{VbaValue::string(spec.name.toStdString()),VbaValue::integer(static_cast<std::int16_t>(type)),VbaValue::longInteger(spec.size)}).object;
                field->assign("Required",VbaValue::logical(spec.required));if(type==10 || type==12) field->assign("AllowZeroLength",VbaValue::logical(spec.allowZero));
                if(!spec.defaultValue.isEmpty()) field->assign("DefaultValue",VbaValue::string(spec.defaultValue.toStdString()));
                catalog->member("Append",{VbaValue::reference(field)});
            }
            auto primary=table->member("CreateIndex",{VbaValue::string("PrimaryKey")}).object;primary->assign("Primary",VbaValue::logical(true));
            primary->member("Fields",{}).object->member("Append",{primary->member("CreateField",{VbaValue::string("id")})});
            table->member("Indexes",{}).object->member("Append",{VbaValue::reference(primary)});
            database->member("TableDefs",{}).object->member("Append",{VbaValue::reference(table)});
        } catch(const std::exception& error) {
            messageError(&dialog,QString::fromUtf8(error.what()));return;
        }
        dialog.accept();
    });
    if(dialog.exec()!=QDialog::Accepted) return;
    refreshExplorer();
}

int MainWindow::showTable(const QString& name,bool readOnly,bool interactive,const QString& queryName,const QString& querySource) {
    if(!db_.isOpen() || QApplication::activeModalWidget()) return 2046;
    if(auto* existing=queryName.isEmpty() ? findOpenTable(name) : findOpenQuery(queryName)) {
        if(!interactive) for(auto* candidate:existing->findChildren<QTableView*>()) if(auto* view=dynamic_cast<linaccess::AccessTableView*>(candidate)) {
            const auto* model=qobject_cast<linaccess::AccessTableModel*>(view->model());
            if(model && model->isReadOnly()!=readOnly) return 3251;
        }
        return activateDataPage(existing) ? 0 : 2501;
    }
    if(!readOnly) for(int i=0;i<tabs_->count();++i) {
        for(auto* candidate:tabs_->widget(i)->findChildren<linaccess::AccessTableModel*>())
            if(!candidate->isReadOnly() && candidate->tableName().compare(name,Qt::CaseInsensitive)==0) return 3251;
    }
    const QPointer<MainWindow> self=this;
    const auto database=db_; const auto identity=db_.databaseName(); const auto generation=property("la_project_generation").toULongLong();
    const auto alive=[self,identity,generation] { return self && self->db_.isOpen() && self->db_.databaseName()==identity && self->property("la_project_generation").toULongLong()==generation; };
    if(!finishActiveDataEditors() || !alive() || !AccessFormPreview::activateProjectDocument(database,this,nullptr) || !alive()) return 2501;
    if(!readOnly) for(int i=0;i<tabs_->count();++i) {
        for(auto* candidate:tabs_->widget(i)->findChildren<linaccess::AccessTableModel*>())
            if(!candidate->isReadOnly() && candidate->tableName().compare(name,Qt::CaseInsensitive)==0) return 3251;
    }
    QString sourceSchema,sourceFilter;
    if(!queryName.isEmpty()) {
        try {
            if(linaccess::editableAccessQueryTable(database,querySource,&sourceFilter)!=name) return 3251;
            sourceSchema=editableSourceSchema(database,name);
        } catch(const linaccess::VbaRuntimeError&) { return 3251; }
    }
    auto ownedPage=std::make_unique<QWidget>(); auto* page=ownedPage.get();
    tagProjectObject(page,this,database,queryName.isEmpty() ? "table" : "query",queryName.isEmpty() ? name : queryName);
    if(!queryName.isEmpty()) {
        page->setProperty("la_editable_query",true); page->setProperty("la_query_data_mode",1);
        page->setProperty("la_editable_query_source",querySource);
        page->setProperty("la_editable_query_schema",sourceSchema);
    }
    page->setProperty("la_project_object_scope","tab");
    auto* outer = new QVBoxLayout(page);
    outer->addWidget(new QLabel((queryName.isEmpty() ? "Tabelle: " + name : "Abfrage: " + queryName) + (readOnly ? " — Schreibgeschützt" : " — Zelländerungen mit 'Speichern' festschreiben")));
    auto* actions = new QHBoxLayout();
    auto* add = addButton(actions, "+ Datensatz");
    auto* remove = addButton(actions, "Datensatz löschen");
    auto* save = addButton(actions, "Speichern");
    auto* discard = addButton(actions, "Verwerfen");
    for(auto* button:{add,remove,save,discard}) button->setEnabled(!readOnly);
    outer->addLayout(actions);
    auto* model = new linaccess::AccessTableModel(page, database);
    connect(model,&linaccess::AccessTableModel::editRejected,page,[self](int number,const QString& message) { if(self) self->statusBar()->showMessage(message+" ("+QString::number(number)+")",10000); });
    connect(model,&QAbstractItemModel::dataChanged,page,[self] { if(self) self->statusBar()->clearMessage(); });
    model->setTable(name);
    if(!sourceFilter.isEmpty()) model->setFilter(sourceFilter);
    model->setEditStrategy(QSqlTableModel::OnManualSubmit);
    (void)model->setReadOnly(readOnly);
    bool selected=false;
    {
        const auto key=queryName.toCaseFolded();
        if(!key.isEmpty()) { if(executingQueries_.contains(key)) return 3251; executingQueries_.insert(key); }
        ++queryExecutionDepth_;
        const auto finished=qScopeGuard([self,key] { if(self) { if(!key.isEmpty()) self->executingQueries_.remove(key); --self->queryExecutionDepth_; } });
        selected=model->select();
    }
    if(!alive()) return 2501;
    if(!selected) {
        const auto problem=model->lastError();
        bool valid=false; const int code=problem.nativeErrorCode().toInt(&valid);
        if(interactive) messageError(this,problem.text());
        return problem.driverText()=="LinAccess" && valid ? code : 3146;
    }
    auto* table = new linaccess::AccessTableView(page);
    if(!queryName.isEmpty()) table->setObjectName("native-editable-query-view");
    table->installEventFilter(this);
    table->setModel(model);
    table->setItemDelegate(new linaccess::AccessTableDelegate(table));
    table->setRecordNavigationHandler([self=QPointer<MainWindow>(this),page=QPointer<QWidget>(page),table=QPointer<linaccess::AccessTableView>(table),model=QPointer<linaccess::AccessTableModel>(model)](linaccess::AccessTableView::RecordMove action,int distance) {
        if(!self || !page || !table || !model) return false;
        if(self->tabs_->indexOf(page)<0) return model->rejectEdit(2501,"Das Datenblatt ist nicht mehr geöffnet.");
        if(!table->finishEditing() || !self || !page || !table || !model) return false;
        if(!self->activateDataPage(page)) return model ? model->rejectEdit(2501,"Datensatzwechsel wurde bei der Aktivierung abgebrochen.") : false;
        if(!self || !page || !table || !model) return false;
        const auto result=self->navigateTableDocument({},static_cast<int>(action),distance,false);
        if(!self || !page || !table || !model) return false;
        if(!result.value("handled").toBool()) return model->rejectEdit(2046,"Datensatzwechsel ist nicht verfügbar.");
        if(const int error=result.value("error").toInt()) return model->rejectEdit(error,result.value("description").toString());
        return true;
    });
    table->setRecordNavigationGuard([self=QPointer<MainWindow>(this),page=QPointer<QWidget>(page),table=QPointer<linaccess::AccessTableView>(table),model=QPointer<linaccess::AccessTableModel>(model)] {
        if(!self || !page || !table || !model) return false;
        if(QApplication::activeModalWidget()) return model->rejectEdit(2046,"Suche ist durch ein modales Fenster blockiert.");
        if(self->tabs_->indexOf(page)<0) return model->rejectEdit(2501,"Tabellendokument ist nicht mehr geöffnet.");
        if((self->tabs_->currentWidget()!=page || activePreviewDocument() || !AccessFormPreview::databaseWindowActive(self)) && !self->activateDataPage(page))
            return model ? model->rejectEdit(2501,"Tabellenaktivierung wurde vor der Suche abgebrochen.") : false;
        if(!self || !page || !table || !model) return false;
        if(!model->isDirty() || !model->hasPendingRecord(table->currentIndex().row())) return true;
        const auto saved=self->saveDataPage(page,false,true);
        if(const int error=saved.value("error").toInt()) {
            if(model) return model->rejectEdit(error,saved.value("description").toString());
            return false;
        }
        return bool(self && page && table && model);
    });
    {
        const auto key=queryName.toCaseFolded(); if(!key.isEmpty()) executingQueries_.insert(key); ++queryExecutionDepth_;
        const auto finished=qScopeGuard([self,key] { if(self) { if(!key.isEmpty()) self->executingQueries_.remove(key); --self->queryExecutionDepth_; } });
        table->setSortingEnabled(true);
    }
    if(!alive()) return 2501;
    table->horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);
    outer->addWidget(table);
    connect(add, &QPushButton::clicked, page, [table] { (void)table->insertRecord(); });
    connect(remove, &QPushButton::clicked, page, [table] { (void)table->deleteCurrentRecord(); });
    connect(discard, &QPushButton::clicked, page, [model] { model->revertAll(); });
    connect(save, &QPushButton::clicked, page, [this,page] { (void)saveDataPage(page,true); });
    auto* saveShortcut=new QShortcut(QKeySequence::Save,page);
    saveShortcut->setContext(Qt::WidgetWithChildrenShortcut);
    connect(saveShortcut,&QShortcut::activated,save,&QPushButton::click);
    for(const auto key:{Qt::Key_Return,Qt::Key_Enter}) {
        auto* recordSaveShortcut=new QShortcut(QKeySequence(Qt::ShiftModifier|key),page);
        recordSaveShortcut->setContext(Qt::WidgetWithChildrenShortcut);
        connect(recordSaveShortcut,&QShortcut::activated,save,&QPushButton::click);
    }
    auto* requeryShortcut=new QShortcut(QKeySequence(Qt::ShiftModifier|Qt::Key_F9),page);
    requeryShortcut->setContext(Qt::WidgetWithChildrenShortcut);
    connect(requeryShortcut,&QShortcut::activated,page,[this] {
        const auto result=requeryTableDocument();
        if(const int error=result.value("error").toInt())
            statusBar()->showMessage(result.value("description").toString()+" ("+QString::number(error)+")",10000);
    });
    for(const auto key:{QKeySequence(Qt::ControlModifier|Qt::Key_W),QKeySequence(Qt::ControlModifier|Qt::Key_F4)}) {
        auto* closeShortcut=new QShortcut(key,page); closeShortcut->setContext(Qt::WidgetWithChildrenShortcut);
        connect(closeShortcut,&QShortcut::activated,page,[this,page] {
            const auto index=tabs_->indexOf(page); if(index>=0) emit tabs_->tabCloseRequested(index);
        });
    }
    page=ownedPage.release(); page->setParent(tabs_);
    const QPointer<QWidget> registered=page;
    tabs_->setCurrentIndex(tabs_->addTab(page, queryName.isEmpty() ? name : "Abfrage: "+queryName));
    const QPointer<linaccess::AccessTableModel> registeredModel=model;
    const QPointer<linaccess::AccessTableView> registeredView=table;
    if(!queryName.isEmpty() && registeredModel && registeredView && model->rowCount() && model->columnCount()) table->setCurrentIndex(model->index(0,0));
    if(!alive() || !registered || !registeredModel || !registeredView) return 2501;
    const bool activated=activateDataPage(page);
    if(!alive() || !registered) return 2501;
    if(!activated) { const int index=tabs_->indexOf(registered); if(index>=0) tabs_->removeTab(index); registered->deleteLater(); return 2501; }
    return 0;
}

void MainWindow::editModule(const QString& existingName) {
    if(!hasProject()) return;const auto old=existingName.isEmpty() ? QJsonObject{} : objectDefinition("module",existingName);
    try {if(!existingName.isEmpty()) (void)linaccess::nativeModuleSource(old);}catch(const std::exception& error) {messageError(this,QString::fromUtf8(error.what()));return;}
    QDialog dialog(this);dialog.setObjectName("native-module-editor");dialog.setWindowTitle("Standardmodul bearbeiten");dialog.resize(900,650);tagProjectObject(&dialog,this,db_,"module",existingName);auto* layout=new QVBoxLayout(&dialog);
    const QPointer<MainWindow> owner=this;const auto identity=db_.databaseName();const auto generation=property("la_project_generation");
    const auto checkProject=[owner,identity,generation] {if(!owner || !owner->db_.isOpen() || owner->db_.databaseName()!=identity || owner->property("la_project_generation")!=generation) throw linaccess::VbaRuntimeError(3420,"Das Modul gehört nicht mehr zum aktuellen Projekt.");};

    auto* name=new QLineEdit(existingName,&dialog);name->setObjectName("native-module-name");name->setMaxLength(64);name->setReadOnly(!existingName.isEmpty());name->setPlaceholderText("Name des Standardmoduls");layout->addWidget(name);
    auto* source=new QPlainTextEdit(&dialog);source->setObjectName("native-module-source");source->setPlainText(old.value("vba").toString());source->setPlaceholderText("VBA-Quelltext für ein Standardmodul");layout->addWidget(source,1);
    auto* note=new QLabel("Öffentliche Funktionen stehen unterstützten VBA-, Eval- und SQL-Aufrufen zur Verfügung. Ausführen verwendet das gespeicherte Modul ohne Pflichtargumente. Die native VBA-Sprache ist eingeschränkt; Klassenmodule und ein Formularkontext werden hier nicht bereitgestellt.",&dialog);note->setWordWrap(true);layout->addWidget(note);
    auto* row=new QHBoxLayout;auto* analyze=addButton(row,"Analysieren");analyze->setObjectName("native-module-analyze");auto* save=addButton(row,"Speichern");save->setObjectName("native-module-save");auto* remove=addButton(row,"Modul löschen");remove->setObjectName("native-module-delete");remove->setEnabled(!existingName.isEmpty());layout->addLayout(row);installDesignerSaveShortcut(dialog,save);
    auto* execution=new QHBoxLayout;auto* procedures=new QComboBox(&dialog);procedures->setObjectName("native-module-procedure");procedures->setProperty("la_designer_transient",true);execution->addWidget(procedures,1);auto* run=addButton(execution,"Prozedur ausführen (F5)");run->setObjectName("native-module-run");layout->addLayout(execution);
    auto* runShortcut=new QShortcut(QKeySequence(Qt::Key_F5),&dialog);connect(runShortcut,&QShortcut::activated,run,&QPushButton::click);
    const auto updateProcedures=[&] {const auto selected=procedures->currentData();procedures->clear();for(const auto& procedure:linaccess::analyzeVba(source->toPlainText().toUtf8().toStdString()).procedures) procedures->addItem(QString::fromStdString(procedure.kind)+" "+QString::fromStdString(procedure.name),QString::fromStdString(procedure.name));const int index=procedures->findData(selected);if(index>=0) procedures->setCurrentIndex(index);run->setEnabled(procedures->count()>0);};
    updateProcedures();connect(analyze,&QPushButton::clicked,&dialog,updateProcedures);
    connect(source,&QPlainTextEdit::textChanged,&dialog,[run] {run->setEnabled(false);});
    auto* result=new QTextBrowser(&dialog);result->setObjectName("native-module-result");result->setMaximumHeight(140);layout->addWidget(result);
    connect(analyze,&QPushButton::clicked,&dialog,[&] {const auto analysis=linaccess::analyzeVba(source->toPlainText().toUtf8().toStdString());QStringList lines{QString("Syntaxinventar: %1 Prozeduren; kein Ausführungsnachweis.").arg(analysis.procedures.size())};for(const auto& flag:analysis.compatibilityFlags) lines.append(QString::fromStdString(flag));result->setPlainText(lines.join('\n'));});
    connect(run,&QPushButton::clicked,&dialog,[&] {
        try {
            checkProject();const auto saved=dialog.property("la_editor_saved_name").toString();const auto current=saved.isEmpty() ? existingName : saved;
            if(current.isEmpty() || dialog.isWindowModified()) throw linaccess::VbaRuntimeError(3251,"Modul vor der Ausführung speichern und analysieren.");
            const auto code=linaccess::nativeModuleSource(objectDefinition("module",current));
            const auto catalog=linaccess::standardVbaModuleSources(db_);
            if(!catalog.contains(current.toCaseFolded()) || catalog.value(current.toCaseFolded())!=code) throw linaccess::VbaRuntimeError(3420,"Gespeicherter Modulkatalog hat sich geändert.");
            if(code!=source->toPlainText().toUtf8()) throw linaccess::VbaRuntimeError(3251,"Die gespeicherte Modulrevision hat sich geändert. Modul neu öffnen.");
            run->setEnabled(false);const auto restore=qScopeGuard([run] {run->setEnabled(true);});
            linaccess::NativeStandardModuleHost host(&dialog,db_,checkProject);
            const auto outcome=linaccess::executeVbaProcedure(code.toStdString(),procedures->currentData().toString().toStdString(),host);
            if(outcome.status==linaccess::VbaExecutionStatus::Completed) result->setPlainText("Prozedur ausgeführt."+(outcome.returnValue ? " Rückgabewert: "+QString::fromStdString(*outcome.returnValue) : QString{}));
            else result->setPlainText(QString("Fehler %1 in Zeile %2: %3").arg(outcome.errorNumber ? outcome.errorNumber : 3251).arg(outcome.line).arg(QString::fromStdString(outcome.explanation)));
        }catch(const linaccess::VbaRuntimeError& error) {result->setPlainText(QString("Fehler %1: %2").arg(error.number).arg(QString::fromUtf8(error.what())));}
    });
    connect(save,&QPushButton::clicked,&dialog,[&] {
        try {
            checkProject();
            const auto moduleName=name->text().trimmed();if(!linaccess::nativeModuleNameValid(moduleName)) throw linaccess::VbaRuntimeError(5,"Ungültiger Standardmodulname.");
            const QJsonObject definition{{"version",1},{"vba",source->toPlainText()}};(void)linaccess::nativeModuleSource(definition);
            const auto saved=dialog.property("la_editor_saved_name").toString();const auto current=saved.isEmpty() ? existingName : saved;
            const auto modules=linaccess::standardVbaModuleSources(db_);if(current.isEmpty() && modules.contains(moduleName.toCaseFolded())) throw linaccess::VbaRuntimeError(3010,"Modulname ist bereits vorhanden. Öffnen Sie das vorhandene native Modul oder wählen Sie einen anderen Namen.");
            if(!current.isEmpty() && current!=moduleName) throw linaccess::VbaRuntimeError(3251,"Modulumbenennung ist hier nicht implementiert.");
            if(saveObject("module",moduleName,definition)) {name->setText(moduleName);name->setReadOnly(true);remove->setEnabled(true);tagProjectObject(&dialog,this,db_,"module",moduleName);completeDesignerSave(dialog,moduleName);updateProcedures();}
        }catch(const std::exception& error) {messageError(&dialog,QString::fromUtf8(error.what()));}
    });
    connect(remove,&QPushButton::clicked,&dialog,[&] {
        try {checkProject();
        const auto saved=dialog.property("la_editor_saved_name").toString();const auto moduleName=saved.isEmpty() ? existingName : saved;if(moduleName.isEmpty()) return;
        QSqlQuery snapshot(db_);snapshot.prepare("SELECT definition FROM main._la_objects WHERE kind='module' AND name=? COLLATE BINARY");snapshot.addBindValue(moduleName);if(!snapshot.exec() || !snapshot.next()) {messageError(&dialog,"Gespeichertes Modul ist nicht verfügbar.");return;}const auto previous=snapshot.value(0);snapshot.finish();
        if(QMessageBox::question(&dialog,"Modul löschen","Gespeichertes Modul und ungespeicherten Text wirklich verwerfen?",QMessageBox::Yes|QMessageBox::No,QMessageBox::No)!=QMessageBox::Yes) return;
        checkProject();
        QSqlQuery query(db_);query.prepare("DELETE FROM main._la_objects WHERE kind='module' AND name=? COLLATE BINARY AND definition=?");query.addBindValue(moduleName);query.addBindValue(previous);
        if(!query.exec() || query.numRowsAffected()!=1) {messageError(&dialog,"Modul wurde zwischenzeitlich geändert oder konnte nicht gelöscht werden.");return;}refreshExplorer();dialog.accept();
        }catch(const std::exception& error) {messageError(&dialog,QString::fromUtf8(error.what()));}
    });
    trackDesignerChanges(dialog);dialog.exec();
}

void MainWindow::editMacro(const QString& existingName) {
    if(!hasProject()) return;QJsonObject old;std::optional<linaccess::NativeMacroRevision> revision;
    try {if(!existingName.isEmpty()) {revision=linaccess::readNativeMacroRevision(db_,existingName);if(!revision) throw linaccess::VbaRuntimeError(2485,"Gespeichertes Makro ist nicht verfügbar.");QJsonParseError error;const auto document=QJsonDocument::fromJson(revision->definition.toByteArray(),&error);if(error.error!=QJsonParseError::NoError || !document.isObject()) throw linaccess::VbaRuntimeError(13,"Ungültige native Makrodefinition.");old=document.object();(void)linaccess::parseNativeMacro(old);}}catch(const std::exception& error) {messageError(this,QString::fromUtf8(error.what()));return;}
    QDialog dialog(this);dialog.setObjectName("native-macro-designer");dialog.setWindowTitle("Makro bearbeiten");dialog.resize(1000,600);tagProjectObject(&dialog,this,db_,"macro",existingName);
    const QPointer<MainWindow> owner=this;const auto identity=db_.databaseName();const auto generation=property("la_project_generation");
    const auto checkProject=[owner,identity,generation] {if(!owner || !owner->db_.isOpen() || owner->db_.databaseName()!=identity || owner->property("la_project_generation")!=generation) throw linaccess::VbaRuntimeError(3420,"Makro gehört nicht mehr zum aktuellen Projekt.");};
    auto* layout=new QVBoxLayout(&dialog);auto* name=new QLineEdit(existingName,&dialog);name->setObjectName("native-macro-name");name->setMaxLength(64);name->setReadOnly(!existingName.isEmpty());name->setPlaceholderText("Name des Makros / der Makrogruppe");layout->addWidget(name);
    auto* groupRow=new QHBoxLayout;auto* grouped=new QCheckBox("Makrogruppe",&dialog);grouped->setObjectName("native-macro-group-mode");grouped->setProperty("la_designer_transient",true);groupRow->addWidget(grouped);
    auto* selector=new QComboBox(&dialog);selector->setObjectName("native-macro-group");selector->setProperty("la_designer_transient",true);groupRow->addWidget(selector,1);
    auto* groupName=new QLineEdit(&dialog);groupName->setObjectName("native-macro-group-name");groupName->setMaxLength(64);groupName->setPlaceholderText("Name des Teilmakros");groupRow->addWidget(groupName,1);
    auto* groupAdd=addButton(groupRow,"Teilmakro hinzufügen");groupAdd->setObjectName("native-macro-group-add");auto* groupRemove=addButton(groupRow,"Teilmakro entfernen");groupRemove->setObjectName("native-macro-group-remove");auto* groupUp=addButton(groupRow,"↑");groupUp->setObjectName("native-macro-group-up");auto* groupDown=addButton(groupRow,"↓");groupDown->setObjectName("native-macro-group-down");layout->addLayout(groupRow);
    auto* list=new QTableWidget(0,5,&dialog);list->setObjectName("native-macro-actions");list->setHorizontalHeaderLabels({"Aktion","Argument 1 (Ziel / SQL / Befehl)","Argument 2 (optional)","Argument 3 (optional)","Bedingung (optional)"});list->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);list->setSelectionBehavior(QAbstractItemView::SelectRows);layout->addWidget(list);
    auto* help=new QLabel("Aktionen laufen von oben nach unten. Eine Makrogruppe startet ohne Teilnamen ihr erstes Teilmakro; gezielter Aufruf: Gruppe.Teilmakro. Verweise werden beim Umbenennen nicht automatisch geändert. RunMacro: Makroname, optionale Anzahl und Wiederholungsbedingung; bei Bedingung darf die Anzahl leer bleiben. RunSQL: SQL und optional Ja/Nein für Transaktion. RunCode: Funktion(...). MessageBox: Meldung (auch =Ausdruck), Ja/Nein, benannter Typ und Titel; alle Argumente über Argumente… bearbeiten. StopMacro beendet den aktuellen Ablauf.",&dialog);help->setWordWrap(true);layout->addWidget(help);
    auto* buttons=new QHBoxLayout;auto* add=addButton(buttons,"Aktion hinzufügen");add->setObjectName("native-macro-add");auto* arguments=addButton(buttons,"Argumente…");arguments->setObjectName("native-macro-arguments");auto* remove=addButton(buttons,"Aktion entfernen");auto* up=addButton(buttons,"Nach oben");auto* down=addButton(buttons,"Nach unten");auto* save=addButton(buttons,"Speichern");save->setObjectName("native-macro-save");auto* deleteMacro=addButton(buttons,"Makroobjekt löschen");deleteMacro->setObjectName("native-macro-delete");deleteMacro->setEnabled(revision.has_value());layout->addLayout(buttons);installDesignerSaveShortcut(dialog,save);
    QVector<linaccess::NativeMacroGroupDraft> groups;int current=0;
    const bool groupMode=old.value("version")==QJsonValue(2);grouped->setChecked(groupMode);
    if(groupMode) for(const auto& value:old.value("groups").toArray()) {const auto group=value.toObject();groups.push_back({group.value("name").toString(),linaccess::nativeMacroDraftRows(group.value("actions").toArray())});}
    else groups.push_back({{},old.isEmpty() ? QVector<linaccess::NativeMacroRowDraft>{linaccess::NativeMacroRowDraft{}} : linaccess::nativeMacroDraftRows(old.value("actions").toArray())});
    const auto dirty=[&] {dialog.setWindowModified(true);if(auto* status=dialog.findChild<QLabel*>("native-designer-save-status")) status->setText("Nicht gespeicherte Änderungen");};
    const auto append=[&](const linaccess::NativeMacroRowDraft& entry) {const int row=list->rowCount();if(row>=256) return;list->insertRow(row);auto* action=new QComboBox(list);for(const auto& signature:linaccess::nativeMacroSignatures()) action->addItem(signature.name);action->setCurrentText(entry.action);list->setCellWidget(row,0,action);for(int index=0;index<4;++index) list->setItem(row,index+1,new QTableWidgetItem(entry.cells[index]));list->item(row,1)->setData(Qt::UserRole,entry.extraArguments);QObject::connect(action,&QComboBox::currentTextChanged,&dialog,dirty);list->setCurrentCell(row,1);};
    const auto capture=[&] {auto& group=groups[current];group.rows.clear();if(grouped->isChecked()) group.name=groupName->text().trimmed();for(int row=0;row<list->rowCount();++row) {linaccess::NativeMacroRowDraft draft;draft.action=qobject_cast<QComboBox*>(list->cellWidget(row,0))->currentText();for(int index=0;index<4;++index) draft.cells[index]=list->item(row,index+1)->text();draft.extraArguments=list->item(row,1)->data(Qt::UserRole).toStringList();group.rows.push_back(draft);}};
    QObject::connect(arguments,&QPushButton::clicked,&dialog,[&] {
        arguments->setFocus(Qt::ShortcutFocusReason);const int row=list->currentRow();if(row<0) return;
        const auto action=qobject_cast<QComboBox*>(list->cellWidget(row,0))->currentText();const linaccess::NativeMacroSignature* signature=nullptr;for(const auto& entry:linaccess::nativeMacroSignatures()) if(entry.name==action) signature=&entry;if(!signature) return;
        QDialog editor(&dialog);editor.setObjectName("native-macro-argument-dialog");editor.setWindowTitle(action+" – Argumente");editor.resize(640,400);auto* fields=new QFormLayout(&editor);QVector<QWidget*> inputs;QStringList values;for(int index=0;index<3;++index) values.push_back(list->item(row,index+1)->text());values.append(list->item(row,1)->data(Qt::UserRole).toStringList());
        // Retain surplus drafts after changing actions so incompatible values cannot disappear silently.
        int count=signature->types.size();for(int index=count;index<values.size();++index) if(!values[index].isEmpty()) count=index+1;
        const QStringList messageLabels{"Meldung","Signalton (Ja/Nein)","Typ (None/Critical/Warning?/Warning!/Information)","Titel"};
        for(int index=0;index<count;++index) {QWidget* input=nullptr;if(signature->types.value(index)=="text") {auto* text=new QPlainTextEdit(values.value(index),&editor);text->setMaximumHeight(85);input=text;}else input=new QLineEdit(values.value(index),&editor);input->setObjectName("native-macro-argument-"+QString::number(index+1));inputs.push_back(input);const auto label=action=="MessageBox" ? messageLabels.value(index) : QString("Argument %1 (%2%3)").arg(index+1).arg(signature->types.value(index,"nicht zulässig"),index<signature->required ? ", erforderlich" : ", optional");fields->addRow(label,input);}
        auto* accept=new QDialogButtonBox(QDialogButtonBox::Ok|QDialogButtonBox::Cancel,&editor);fields->addRow(accept);QObject::connect(accept,&QDialogButtonBox::accepted,&editor,&QDialog::accept);QObject::connect(accept,&QDialogButtonBox::rejected,&editor,&QDialog::reject);
        if(editor.exec()!=QDialog::Accepted) return;QStringList edited;for(auto* input:inputs) edited.push_back(qobject_cast<QPlainTextEdit*>(input) ? qobject_cast<QPlainTextEdit*>(input)->toPlainText() : qobject_cast<QLineEdit*>(input)->text());for(int index=0;index<3;++index) list->item(row,index+1)->setText(edited.value(index));list->item(row,1)->setData(Qt::UserRole,edited.mid(3));dirty();
    });
    const auto totalRows=[&] {qsizetype count=list->rowCount();for(int index=0;index<groups.size();++index) if(index!=current) count+=groups[index].rows.size();return count;};
    const auto controls=[&] {const bool enabled=grouped->isChecked();add->setEnabled(totalRows()<256);selector->setEnabled(enabled);groupName->setEnabled(enabled);groupAdd->setEnabled(enabled && groups.size()<64 && totalRows()<256);groupRemove->setEnabled(enabled && groups.size()>1);groupUp->setEnabled(enabled && current>0);groupDown->setEnabled(enabled && current+1<groups.size());};
    const auto render=[&] {const QSignalBlocker modelSignals(list->model()),nameSignals(groupName),selectorSignals(selector);list->setRowCount(0);for(const auto& row:groups[current].rows) append(row);groupName->setText(groups[current].name);selector->clear();for(const auto& group:groups) selector->addItem(group.name.isEmpty() ? "(ohne Namen)" : group.name);selector->setCurrentIndex(current);controls();};render();
    QObject::connect(groupName,&QLineEdit::textChanged,&dialog,[&](const QString& value) {groups[current].name=value;const QSignalBlocker blocked(selector);selector->setItemText(current,value.isEmpty() ? "(ohne Namen)" : value);});
    QObject::connect(selector,&QComboBox::currentIndexChanged,&dialog,[&](int index) {if(index<0 || index>=groups.size() || index==current) return;capture();current=index;render();});
    QObject::connect(grouped,&QCheckBox::toggled,&dialog,[&](bool enabled) {if(!enabled && groups.size()>1) {const QSignalBlocker blocked(grouped);grouped->setChecked(true);messageError(&dialog,"Vor dem Umwandeln bis auf ein Teilmakro reduzieren.");return;}capture();controls();dirty();if(enabled) groupName->setFocus();});
    QObject::connect(groupAdd,&QPushButton::clicked,&dialog,[&] {if(groups.size()>=64 || totalRows()>=256) return;capture();groups.push_back({{},QVector<linaccess::NativeMacroRowDraft>{linaccess::NativeMacroRowDraft{}}});current=groups.size()-1;render();dirty();groupName->setFocus();});
    QObject::connect(groupRemove,&QPushButton::clicked,&dialog,[&] {if(groups.size()<2) return;groups.removeAt(current);current=qMin(current,static_cast<int>(groups.size())-1);render();dirty();});
    const auto moveGroup=[&](int delta) {const int target=current+delta;if(target<0 || target>=groups.size()) return;capture();groups.swapItemsAt(current,target);current=target;render();dirty();};QObject::connect(groupUp,&QPushButton::clicked,&dialog,[&] {moveGroup(-1);});QObject::connect(groupDown,&QPushButton::clicked,&dialog,[&] {moveGroup(1);});
    QObject::connect(list->model(),&QAbstractItemModel::rowsInserted,&dialog,controls);QObject::connect(list->model(),&QAbstractItemModel::rowsRemoved,&dialog,controls);
    QObject::connect(add,&QPushButton::clicked,&dialog,[&] {if(totalRows()<256) append({});});QObject::connect(remove,&QPushButton::clicked,&dialog,[&] {if(list->currentRow()>=0) list->removeRow(list->currentRow());});
    const auto move=[&](int distance) {const int row=list->currentRow(),target=row+distance;if(row<0 || target<0 || target>=list->rowCount()) return;auto* first=qobject_cast<QComboBox*>(list->cellWidget(row,0));auto* second=qobject_cast<QComboBox*>(list->cellWidget(target,0));const auto action=first->currentText();first->setCurrentText(second->currentText());second->setCurrentText(action);const auto extra=list->item(row,1)->data(Qt::UserRole);list->item(row,1)->setData(Qt::UserRole,list->item(target,1)->data(Qt::UserRole));list->item(target,1)->setData(Qt::UserRole,extra);for(int column=1;column<5;++column) {const auto value=list->item(row,column)->text();list->item(row,column)->setText(list->item(target,column)->text());list->item(target,column)->setText(value);}list->setCurrentCell(target,1);};QObject::connect(up,&QPushButton::clicked,&dialog,[&] {move(-1);});QObject::connect(down,&QPushButton::clicked,&dialog,[&] {move(1);});
    QObject::connect(save,&QPushButton::clicked,&dialog,[&] {
        try {checkProject();save->setFocus(Qt::ShortcutFocusReason);if(!linaccess::nativeMacroNameValid(name->text().trimmed())) throw linaccess::VbaRuntimeError(5,"Ungültiger Makro-Objektname.");capture();QJsonObject definition;
            if(grouped->isChecked()) {QJsonArray entries;for(const auto& group:groups) entries.append(QJsonObject{{"name",group.name.trimmed()},{"actions",linaccess::nativeMacroDraftActions(group.rows)}});definition={{"version",2},{"groups",entries}};}
            else definition={{"version",1},{"actions",linaccess::nativeMacroDraftActions(groups[0].rows)}};
            (void)linaccess::parseNativeMacro(definition);
            revision=linaccess::storeNativeMacro(db_,name->text().trimmed(),definition,revision);refreshExplorer();name->setText(revision->name);name->setReadOnly(true);deleteMacro->setEnabled(true);tagProjectObject(&dialog,this,db_,"macro",revision->name);completeDesignerSave(dialog,revision->name);
        }catch(const std::exception& error) {messageError(&dialog,QString::fromUtf8(error.what()));}
    });
    QObject::connect(deleteMacro,&QPushButton::clicked,&dialog,[&] {
        try {checkProject();if(!revision) return;const auto snapshot=*revision;
            if(QMessageBox::question(&dialog,"Makroobjekt löschen","Gespeichertes Makroobjekt mit allen Teilmakros und ungespeicherten Änderungen wirklich löschen?",QMessageBox::Yes|QMessageBox::No,QMessageBox::No)!=QMessageBox::Yes) return;
            checkProject();linaccess::deleteNativeMacro(db_,snapshot);refreshExplorer();dialog.accept();
        }catch(const std::exception& error) {messageError(&dialog,QString::fromUtf8(error.what()));}
    });trackDesignerChanges(dialog);dialog.exec();
}

void MainWindow::editQuery(const QString& existingName) {
    if (!hasProject()) return;
    const auto old=existingName.isEmpty() ? QJsonObject{} : objectDefinition("query",existingName);
    QDialog dialog(this); dialog.setObjectName("native-query-designer");
    dialog.setWindowTitle("SQL-Abfrage");
    tagProjectObject(&dialog,this,db_,"query",existingName);
    dialog.resize(920, 700);
    auto* outer = new QVBoxLayout(&dialog);
    auto* name = new QLineEdit(existingName, &dialog);
    name->setObjectName("native-query-name"); name->setPlaceholderText("Name der Abfrage");
    outer->addWidget(name);
    auto* sql = new QPlainTextEdit(&dialog);
    sql->setObjectName("native-query-sql"); sql->setPlaceholderText("SELECT * FROM kunden");
    sql->setMaximumHeight(190);
    if (!existingName.isEmpty()) sql->setPlainText(objectDefinition("query", existingName).value("sql").toString());
    outer->addWidget(sql);
    auto* buttons = new QHBoxLayout();
    auto* run = addButton(buttons, "Ausführen"); run->setObjectName("native-query-run");
    auto* save = addButton(buttons, "Speichern");
    outer->addLayout(buttons);
    save->setObjectName("native-query-save"); installDesignerSaveShortcut(dialog,save);
    auto* model = new linaccess::AccessQueryModel(&dialog);
    auto* view = new QTableView(&dialog);
    view->setModel(model);
    outer->addWidget(view);
    auto* feedback=new QLabel(&dialog); outer->addWidget(feedback);
    connect(run, &QPushButton::clicked, &dialog, [this, sql, model, view, feedback, &dialog] {
        try {
            const auto plan=linaccess::compileAccessQuery(sql->toPlainText());
            QMap<QString,QVariant> parameters;
            for(const auto& parameter:plan.parameters) {
                bool accepted=false;
                const auto value=QInputDialog::getText(&dialog,"Abfrageparameter",parameter.name+" ("+parameter.type+")",QLineEdit::Normal,{},&accepted);
                if(!accepted) return;
                parameters.insert(parameter.name.toCaseFolded(),value);
            }
            view->setVisible(false);
            if(isSelect(plan.sql)) {
                model->execute(db_,sql->toPlainText(),parameters); view->setVisible(true);
                feedback->setText(QString("%1 Datensätze").arg(model->rowCount()));
            } else {
                const auto definition=linaccess::makeDaoDatabase(db_)->member("CreateQueryDef",{linaccess::VbaValue::string(""),linaccess::VbaValue::string(sql->toPlainText().toUtf8().toStdString())}).object;
                for(auto it=parameters.cbegin();it!=parameters.cend();++it) definition->member("Parameters",{linaccess::VbaValue::string(it.key().toUtf8().toStdString())}).object->assign("Value",linaccess::qtFieldValue(it.value()));
                definition->member("Execute",{});
                feedback->setText(QString("%1 Datensätze geändert").arg(definition->member("RecordsAffected",{}).asNumber()));
                refreshExplorer();
            }
        } catch(const std::exception& error) { feedback->clear(); messageError(&dialog,QString::fromUtf8(error.what())); }
    });
    connect(save, &QPushButton::clicked, &dialog, [this, name, sql, &dialog,old,existingName] {
        try {
            (void)linaccess::compileAccessQuery(sql->toPlainText());
        } catch(const std::exception& error) { messageError(&dialog,QString::fromUtf8(error.what())); return; }
        auto content=old; auto savedName=dialog.property("la_editor_saved_name").toString();
        if(savedName.isEmpty()) savedName=existingName;
        if(!savedName.isEmpty()) { const auto latest=objectDefinition("query",savedName); if(!latest.isEmpty()) content=latest; }
        content.insert("sql",sql->toPlainText());
        if(saveObject("query",name->text(),content)) {
            tagProjectObject(&dialog,this,db_,"query",name->text()); completeDesignerSave(dialog,name->text());
        }
    });
    trackDesignerChanges(dialog);
    dialog.exec();
}

void MainWindow::editForm(const QString& existingName) {
    if (!hasProject()) return;
    const auto old = existingName.isEmpty() ? QJsonObject{} : objectDefinition("form", existingName);
    QJsonObject formEvents;QMap<QString,QMap<QString,QString>> priorFieldEvents;int formTimerInterval=0;
    try {
        formEvents=linaccess::nativeEventObject(linaccess::nativeEventOverrides(old.value("events"),linaccess::nativeFormMacroEvents()));formTimerInterval=linaccess::nativeFormTimerInterval(old.value("timer_interval"));
        QStringList names;for(const auto& value:old.value("fields").toArray()) names.append(value.toString());priorFieldEvents=linaccess::nativeFieldEventOverrides(old.value("field_events"),names);
    } catch(const std::exception& error) {messageError(this,QString::fromUtf8(error.what()));return;}

    if(!existingName.isEmpty()) for(const auto& document:AccessFormPreview::projectDocuments(db_))
        if(document && document->documentKind()==linaccess::AccessDocumentKind::Form && document->documentName().compare(existingName,Qt::CaseInsensitive)==0)
            if(!document->close()) return;
    QDialog dialog(this);
    dialog.setObjectName("native-form-designer");
    dialog.setWindowTitle("Formulardesigner — experimentelle VBA-Ereignislaufzeit");
    tagProjectObject(&dialog,this,db_,"form",existingName);
    dialog.resize(820, 690);
    auto* outer = new QVBoxLayout(&dialog);
    auto* name = new QLineEdit(existingName, &dialog);
    name->setObjectName("native-form-name"); name->setPlaceholderText("Formularname");
    outer->addWidget(name);
    auto* caption=new QLineEdit(old.value("caption").toString(existingName),&dialog);
    caption->setObjectName("native-form-caption"); caption->setPlaceholderText("Formularüberschrift");
    outer->addWidget(caption);
    auto* source = new QComboBox(&dialog); source->setObjectName("native-form-source");
    source->setEditable(true);
    auto sourceNames=tableNames(); sourceNames.append(objectNames("query")); sourceNames.removeDuplicates();
    source->addItems(sourceNames);
    outer->addWidget(new QLabel("Datenquelle: Tabelle, Abfrage oder einzeiliges SELECT"));
    outer->addWidget(source);
    auto* sourceStatus=new QLabel(&dialog); sourceStatus->setObjectName("native-form-source-status");
    sourceStatus->setWordWrap(true); outer->addWidget(sourceStatus);
    auto* fields = new QListWidget(&dialog); fields->setObjectName("native-form-fields");
    fields->setDragDropMode(QAbstractItemView::InternalMove); fields->setDefaultDropAction(Qt::MoveAction);
    fields->setMaximumHeight(170);
    outer->addWidget(new QLabel("Formularfelder — ziehen für Reihenfolge, F2 für Beschriftung"));
    outer->addWidget(fields);
    auto* geometryRow=new QHBoxLayout();
    outer->addWidget(new QLabel("Ausgewähltes Feld: Position und Größe in Twips"));
    QList<QSpinBox*> geometry;
    const QStringList geometryKeys{"left","top","width","height"};
    const QStringList geometryLabels{"Links","Oben","Breite","Höhe"};
    for(int part=0;part<4;++part) {
        geometryRow->addWidget(new QLabel(geometryLabels[part],&dialog));
        auto* value=new QSpinBox(&dialog); value->setRange(part<2 ? 0 : 1,30000);
        value->setObjectName("native-form-field-"+geometryKeys[part]); value->setEnabled(false);
        geometryRow->addWidget(value); geometry.push_back(value);
        connect(value,&QSpinBox::valueChanged,&dialog,[fields,part,geometryKeys](int number) {
            if(auto* item=fields->currentItem()) {
                auto properties=item->data(Qt::UserRole+1).toJsonObject();
                properties.insert(geometryKeys[part],number); item->setData(Qt::UserRole+1,properties);
            }
        });
    }
    outer->addLayout(geometryRow);
    const auto updateGeometry=[fields,geometry,geometryKeys] {
        auto* item=fields->currentItem(); int ordinal=0;
        for(int row=0;row<fields->row(item);++row) if(fields->item(row)->checkState()==Qt::Checked) ++ordinal;
        const auto properties=item ? item->data(Qt::UserRole+1).toJsonObject() : QJsonObject{};
        const QList<int> defaults{2160,240+ordinal*420,3300,300};
        for(int part=0;part<geometry.size();++part) {
            const QSignalBlocker blocked(geometry[part]);
            geometry[part]->setEnabled(item!=nullptr);
            geometry[part]->setValue(properties.value(geometryKeys[part]).toInt(defaults[part]));
        }
    };
    connect(fields,&QListWidget::currentItemChanged,&dialog,updateGeometry);
    connect(fields,&QListWidget::itemChanged,&dialog,updateGeometry);
    connect(fields->model(),&QAbstractItemModel::rowsMoved,&dialog,updateGeometry);
    auto populate = [this, fields, source, old,sourceStatus,priorFieldEvents] {
        fields->clear(); sourceStatus->clear(); const auto table=source->currentText(); if(table.isEmpty()) return;
        QStringList available;QMap<QString,linaccess::NativeFieldDisplayProperties> display;
        try {
            const bool physical=db_.tables(QSql::Tables).contains(table,Qt::CaseInsensitive);
            for(const auto& column:linaccess::nativeFormSourceColumns(db_,table))
                if(!physical || column!="id") available.append(column);
            if(physical) display=linaccess::nativeTableFieldDisplayProperties(db_,table);
            if(!physical) sourceStatus->setText("Abfrageformular: schreibgeschützt; keine Parameterbindung.");
        } catch(const linaccess::VbaRuntimeError& error) {
            sourceStatus->setText("Datenquelle nicht verfügbar (Fehler "+QString::number(error.number)+")."); return;
        } catch(...) { sourceStatus->setText("Datenquelle nicht verfügbar."); return; }
        const auto previous=old.value("fields").toArray(); const bool sameSource=old.value("table").toString()==table;
        const auto captions=sameSource ? old.value("field_captions").toObject() : QJsonObject{};
        QStringList ordered;
        if(sameSource) for(const auto& value:previous) if(available.contains(value.toString()) && !ordered.contains(value.toString())) ordered.push_back(value.toString());
        for(const auto& field:available) if(!ordered.contains(field)) ordered.push_back(field);
        for(const auto& field:ordered) {
            auto inherited=display.value(field.toCaseFolded()).caption.value_or(field);if(inherited.isEmpty()) inherited=field;auto* item=new QListWidgetItem(captions.value(field).toString(inherited),fields);
            item->setData(Qt::UserRole,field); item->setToolTip("Gebundenes Feld: "+field+(display.value(field.toCaseFolded()).description ? "\n"+*display.value(field.toCaseFolded()).description : QString{}));
            if(sameSource) item->setData(Qt::UserRole+1,old.value("field_layout").toObject().value(field).toObject());
            if(sameSource) item->setData(Qt::UserRole+2,linaccess::nativeEventObject(priorFieldEvents.value(field.toCaseFolded())));
            item->setFlags(item->flags()|Qt::ItemIsUserCheckable|Qt::ItemIsEditable);
            item->setCheckState(!sameSource || previous.isEmpty() || previous.contains(field) ? Qt::Checked : Qt::Unchecked);
        }
    };
    auto* eventEditor=new QPushButton("Ereignisse zuordnen...",&dialog);eventEditor->setObjectName("native-form-event-bindings");outer->addWidget(eventEditor);
    connect(eventEditor,&QPushButton::clicked,&dialog,[&] {
        QStringList chosen;QMap<QString,QJsonObject> bindings;
        for(int row=0;row<fields->count();++row) if(fields->item(row)->checkState()==Qt::Checked) {const auto field=fields->item(row)->data(Qt::UserRole).toString();chosen.append(field);const auto events=fields->item(row)->data(Qt::UserRole+2).toJsonObject();if(!events.isEmpty()) bindings.insert(field.toCaseFolded(),events);}
        const auto beforeForm=formEvents;const auto beforeFields=bindings;const int beforeInterval=formTimerInterval;
        try {if(linaccess::editNativeFormEvents(&dialog,formEvents,bindings,formTimerInterval,chosen,objectNames("macro"))) {
            for(int row=0;row<fields->count();++row) if(fields->item(row)->checkState()==Qt::Checked) {const auto events=bindings.value(fields->item(row)->data(Qt::UserRole).toString().toCaseFolded());if(events!=fields->item(row)->data(Qt::UserRole+2).toJsonObject()) fields->item(row)->setData(Qt::UserRole+2,events);}
            if(formEvents!=beforeForm || bindings!=beforeFields || formTimerInterval!=beforeInterval) {dialog.setWindowModified(true);if(auto* status=dialog.findChild<QLabel*>("native-designer-save-status")) status->setText("Nicht gespeicherte Änderungen");}
        }} catch(const std::exception& error) {messageError(&dialog,QString::fromUtf8(error.what()));}
    });
    connect(source, &QComboBox::currentTextChanged, &dialog, populate);
    if (!old.isEmpty()) source->setCurrentText(old.value("table").toString());
    populate();
    outer->addWidget(new QLabel(
        "VBA-Modul — experimenteller Teilinterpreter für wenige Form-/Button-Ereignisse; "
        "keine vollständige Microsoft-VBA-Kompatibilität"));
    auto* vba = new QPlainTextEdit(&dialog); vba->setObjectName("native-form-vba");
    vba->setPlaceholderText("Private Sub Form_Load()\n    MsgBox \"Hallo\"\nEnd Sub");
    vba->setPlainText(old.value("vba").toString());
    auto importedBytes = std::make_shared<QByteArray>(QByteArray::fromBase64(
        old.value("original_vba_base64").toString().toLatin1()));
    outer->addWidget(vba, 1);
    auto* row = new QHBoxLayout();
    auto* importBas = addButton(row, ".bas/.cls laden");
    auto* importBridge = addButton(row, "VBA aus Bridge übernehmen");
    auto* analyze = addButton(row, "VBA analysieren");
    auto* save = addButton(row, "Formular speichern"); save->setObjectName("native-form-save");
    outer->addLayout(row);
    installDesignerSaveShortcut(dialog,save);
    auto* findings = new QTextBrowser(&dialog);
    findings->setMaximumHeight(105);
    outer->addWidget(findings);
    connect(importBas, &QPushButton::clicked, &dialog, [this, vba, importedBytes] {
        const auto path = QFileDialog::getOpenFileName(this, "VBA-Modul laden", {}, "VBA (*.bas *.cls *.txt)");
        if (path.isEmpty()) return;
        QFile file(path);
        if (!file.open(QIODevice::ReadOnly)) { messageError(this, file.errorString()); return; }
        *importedBytes = file.readAll();
        vba->setPlainText(linaccess::decodeVbaSource(*importedBytes));
        QMessageBox::information(this, "Original gesichert",
            "Die Originalbytes des VBA-Moduls werden zusätzlich zur Textansicht gespeichert. "
            "UTF-8, UTF-16LE und Windows-1252 werden zur Anzeige erkannt. "
            "Andere historische Windows-Codepages müssen ggf. manuell geprüft werden.");
    });
    connect(importBridge, &QPushButton::clicked, &dialog, [this, vba, importedBytes] {
        QSqlQuery query(db_);
        if (!query.exec("SELECT id,name FROM _la_access_artifacts WHERE kind='vba' "
                        "AND status='source_only' AND length(source_bytes)>0 ORDER BY name")) {
            messageError(this, query.lastError().text());
            return;
        }
        QStringList choices;
        QVector<int> ids;
        while (query.next()) {
            ids.append(query.value(0).toInt());
            choices << query.value(1).toString() + " (#" + query.value(0).toString() + ")";
        }
        if (choices.isEmpty()) {
            QMessageBox::information(this, "VBA", "Noch keine VBA-Module aus einer Access-Bridge importiert.");
            return;
        }
        bool confirmed = false;
        const auto name = QInputDialog::getItem(this, "Access-VBA auswählen",
            "Importierte VBA-Quelle (Originalbytes bleiben erhalten)", choices, 0, false, &confirmed);
        if (!confirmed) return;
        const auto index = choices.indexOf(name);
        if (index < 0) return;
        query.prepare("SELECT source_bytes FROM _la_access_artifacts WHERE id=?");
        query.addBindValue(ids[index]);
        if (!query.exec() || !query.next()) { messageError(this, "VBA-Quelle verschwunden"); return; }
        *importedBytes = query.value(0).toByteArray();
        vba->setPlainText(linaccess::decodeVbaSource(*importedBytes));
    });
    connect(analyze, &QPushButton::clicked, &dialog, [vba, findings] {
        const auto result = linaccess::analyzeVba(vba->toPlainText().toUtf8().toStdString());
        QStringList lines;
        lines << "Prozeduren: " + QString::number(result.procedures.size());
        for (const auto& proc : result.procedures) {
            lines << QString::fromStdString(proc.kind) + " " + QString::fromStdString(proc.name)
                   + " (Zeile " + QString::number(proc.line) + ")";
        }
        for (const auto& issue : result.compatibilityFlags)
            lines << "Noch nicht implementierte Abhängigkeit: " + QString::fromStdString(issue);
        findings->setPlainText(lines.join('\n'));
    });
    connect(save, &QPushButton::clicked, &dialog, [this, &dialog, name, caption, source, fields, vba, importedBytes,old,&formEvents,&formTimerInterval] {
        QJsonArray chosen; QJsonObject captions,positions,eventBindings;
        for(int index=0;index<fields->count();++index) if(fields->item(index)->checkState()==Qt::Checked) {
            const auto field=fields->item(index)->data(Qt::UserRole).toString(); chosen.append(field); captions.insert(field,fields->item(index)->text());
            const auto position=fields->item(index)->data(Qt::UserRole+1).toJsonObject();
            if(!position.isEmpty()) positions.insert(field,position);
            const auto events=fields->item(index)->data(Qt::UserRole+2).toJsonObject();if(!events.isEmpty()) eventBindings.insert(field.toCaseFolded(),events);
        }
        if(chosen.isEmpty()) { messageError(&dialog,"Mindestens ein Feld auswählen."); return; }
        auto content=old; const auto savedName=dialog.property("la_editor_saved_name").toString();
        if(!savedName.isEmpty()) { const auto latest=objectDefinition("form",savedName); if(!latest.isEmpty()) content=latest; }
        const bool sameFormatSource=content.value("table").toString()==source->currentText();
        const auto previousFormats=sameFormatSource ? content.value("field_formats").toObject() : QJsonObject{};
        QJsonObject formats;QMap<QString,linaccess::NativeFieldDisplayProperties> inherited;
        try {inherited=linaccess::nativeTableFieldDisplayProperties(db_,source->currentText());}catch(const std::exception& error) {messageError(&dialog,QString::fromUtf8(error.what()));return;}
        for(const auto& value:chosen) {const auto key=value.toString().toCaseFolded();if(previousFormats.contains(key)) formats.insert(key,previousFormats.value(key));else if(inherited.value(key).format) formats.insert(key,*inherited.value(key).format);}
        content.insert("field_formats",formats);
        content.insert("table",source->currentText()); content.insert("fields",chosen);
        content.insert("caption",caption->text()); content.insert("field_captions",captions); content.insert("vba",vba->toPlainText());
        content.insert("field_layout",positions);
        content.insert("events",formEvents);content.insert("field_events",eventBindings);content.insert("timer_interval",formTimerInterval);
        content.insert("original_vba_base64",QString::fromLatin1(importedBytes->toBase64())); content.insert("language","vba_subset_experimental");
        try { (void)linaccess::nativeStoredForm(name->text(),content,db_); }
        catch(const std::exception& error) { messageError(&dialog,QString::fromUtf8(error.what())); return; }
        if(saveObject("form",name->text(),content)) {
            tagProjectObject(&dialog,this,db_,"form",name->text()); completeDesignerSave(dialog,name->text());
        }
    });
    trackDesignerChanges(dialog);
    dialog.exec();
}

void MainWindow::runForm(const QString& name) {
    if (!hasProject() || !finishActiveDataEditors()) return;
    try {
        const auto form=linaccess::loadNativeStoredForm(db_,name);
        if(!form) throw linaccess::VbaRuntimeError(2102,"Gespeichertes Formular ist nicht verfügbar.");
        (void)AccessFormPreview::showDocument(linaccess::AccessDocumentKind::Form,form->name,form->layout,form->code,db_,this);
    } catch(const std::exception& error) { messageError(this,QString::fromUtf8(error.what())); }
}

void MainWindow::runReport(const QString& name) {
    if (!hasProject() || !finishActiveDataEditors()) return;
    try {
        const auto report=linaccess::loadNativeStoredReport(db_,name);
        if (!report) throw linaccess::VbaRuntimeError(2103,"Gespeicherter Bericht ist nicht verfügbar.");
        (void)AccessFormPreview::showDocument(linaccess::AccessDocumentKind::Report,report->name,report->layout,report->code,db_,this,report->recordSource);
    } catch(const std::exception& error) { messageError(this,QString::fromUtf8(error.what())); }
}
void MainWindow::editReport(const QString& existingName) {
    if (!hasProject()) return;
    for (const auto& document:AccessFormPreview::projectDocuments(db_))
        if (document && document->documentKind()==linaccess::AccessDocumentKind::Report && document->documentName().compare(existingName,Qt::CaseInsensitive)==0)
            if (!document->close()) return;
    const auto old=existingName.isEmpty() ? QJsonObject{} : objectDefinition("report",existingName);
    QDialog dialog(this); dialog.setObjectName("native-report-designer");
    dialog.setWindowTitle("Berichtsentwurf"); tagProjectObject(&dialog,this,db_,"report",existingName); dialog.resize(930,720);
    auto* outer=new QVBoxLayout(&dialog);
    auto* name=new QLineEdit(existingName,&dialog); name->setObjectName("native-report-name"); name->setPlaceholderText("Berichtname"); outer->addWidget(name);
    auto* title=new QLineEdit(old.value("title").toString(),&dialog); title->setObjectName("native-report-title"); title->setPlaceholderText("Überschrift"); outer->addWidget(title);
    auto* sql=new QPlainTextEdit(&dialog); sql->setObjectName("native-report-sql"); sql->setPlaceholderText("SELECT * FROM kunden"); sql->setPlainText(old.value("sql").toString()); outer->addWidget(sql,1);
    auto* canGrow=new QCheckBox("Text vertikal erweitern (CanGrow)",&dialog); canGrow->setObjectName("native-report-cangrow");
    canGrow->setChecked(old.value("can_grow").toBool());
    canGrow->setToolTip("Ausgerichtete oder untereinanderliegende Textfelder mit begrenztem Seitenumbruch zwischen Textzeilen."); outer->addWidget(canGrow);
    auto* headerBreak=new QComboBox(&dialog); headerBreak->setObjectName("native-report-header-break");
    headerBreak->addItems({"Kein erzwungener Seitenwechsel","Vor dem Berichtskopf","Nach dem Berichtskopf","Vor und nach dem Berichtskopf"});
    headerBreak->setCurrentIndex(qBound(0,old.value("header_force_new_page").toInt(),3));
    outer->addWidget(new QLabel("Seitenwechsel im Berichtskopf",&dialog)); outer->addWidget(headerBreak);
    auto* pageBreak=new QComboBox(&dialog); pageBreak->setObjectName("native-report-page-break");
    pageBreak->addItems({"Kein erzwungener Seitenwechsel","Vor jedem Detailbereich","Nach jedem Detailbereich","Vor und nach jedem Detailbereich"});
    pageBreak->setCurrentIndex(qBound(0,old.value("detail_force_new_page").toInt(),3));
    outer->addWidget(new QLabel("Seitenwechsel im Detailbereich",&dialog)); outer->addWidget(pageBreak);
    auto* footerExpression=new QLineEdit(old.value("footer_expression").toString(),&dialog);
    footerExpression->setObjectName("native-report-footer-expression"); footerExpression->setPlaceholderText("=Sum([Betrag])");
    outer->addWidget(new QLabel("Berechnung im Berichtsfuß (optional, mit = beginnen)",&dialog)); outer->addWidget(footerExpression);
    auto* footerBreak=new QComboBox(&dialog); footerBreak->setObjectName("native-report-footer-break");
    footerBreak->addItems({"Kein erzwungener Seitenwechsel","Vor dem Berichtsfuß","Nach dem Berichtsfuß","Vor und nach dem Berichtsfuß"});
    footerBreak->setCurrentIndex(qBound(0,old.value("footer_force_new_page").toInt(),3));
    outer->addWidget(new QLabel("Seitenwechsel im Berichtsfuß",&dialog)); outer->addWidget(footerBreak);
    outer->addWidget(new QLabel("VBA-Ereignisse: Open, Load, NoData und Close. Druck-/Format-Ereignisse sind noch nicht unterstützt.",&dialog));
    auto* vba=new QPlainTextEdit(&dialog); vba->setObjectName("native-report-vba"); vba->setPlaceholderText("Private Sub Report_Load()\nEnd Sub"); vba->setPlainText(old.value("vba").toString()); outer->addWidget(vba,1);
    auto codeFont=QFontDatabase::systemFont(QFontDatabase::FixedFont); codeFont.setPointSize(10);
    sql->setFont(codeFont); vba->setFont(codeFont);
    auto* row=new QHBoxLayout; auto* preview=addButton(row,"Vorschau"); preview->setObjectName("native-report-preview");
    auto* pdf=addButton(row,"PDF exportieren"); auto* save=addButton(row,"Speichern"); save->setObjectName("native-report-save"); outer->addLayout(row);
    installDesignerSaveShortcut(dialog,save);
    const auto definition=[this,&dialog,existingName,old,title,sql,vba,canGrow,pageBreak,headerBreak,footerExpression,footerBreak] {
        auto content=old;
        auto savedName=dialog.property("la_editor_saved_name").toString();
        if(savedName.isEmpty()) savedName=existingName;
        if(!savedName.isEmpty()) {
            const auto latest=objectDefinition("report",savedName);
            if(!latest.isEmpty()) content=latest;
        }
        const bool changedSource=content.value("sql").toString()!=sql->toPlainText();
        if(changedSource || !content.contains("field_captions") || !content.contains("field_formats")) {
            QString table;try {table=linaccess::editableAccessQueryTable(db_,sql->toPlainText());}catch(const linaccess::VbaRuntimeError& error) {if(error.number!=3251) throw;}
            QJsonObject captions,formats;
            if(!table.isEmpty()) {const auto display=linaccess::nativeTableFieldDisplayProperties(db_,table);for(auto it=display.cbegin();it!=display.cend();++it) {
                if(it.value().caption && !it.value().caption->isEmpty()) captions.insert(it.key(),*it.value().caption);
                if(it.value().format) formats.insert(it.key(),*it.value().format);
            }}
            if(changedSource || !content.contains("field_captions")) content.insert("field_captions",captions);
            content.insert("field_formats",formats);
        }
        content.insert("title",title->text()); content.insert("sql",sql->toPlainText());
        content.insert("vba",vba->toPlainText()); content.insert("can_grow",canGrow->isChecked());
        content.insert("detail_force_new_page",pageBreak->currentIndex());
        content.insert("header_force_new_page",headerBreak->currentIndex());
        content.insert("footer_expression",footerExpression->text());
        content.insert("footer_force_new_page",footerBreak->currentIndex());
        content.insert("language","vba_subset_experimental"); return content;
    };
    const auto packet=[this,name,definition] { return linaccess::nativeStoredReport(name->text().isEmpty() ? QString("Berichtsvorschau") : name->text(),definition(),db_); };
    connect(preview,&QPushButton::clicked,&dialog,[this,&dialog,packet] {
        try {
            const auto report=packet();
            for (const auto& document:AccessFormPreview::projectDocuments(db_))
                if (document && document->documentKind()==linaccess::AccessDocumentKind::Report && document->documentName().compare(report.name,Qt::CaseInsensitive)==0)
                    throw linaccess::VbaRuntimeError(5,"Ein Bericht mit diesem Namen ist bereits geöffnet.");
            AccessFormPreview document(report.name,report.layout,report.code,db_,&dialog,false,{},std::nullopt,0,-1,report.recordSource);
            if (!document.cancelledByOpen() && AccessFormPreview::projectObjectLoaded(db_,linaccess::AccessDocumentKind::Report,report.name))
                document.exec();
        } catch(const std::exception& error) { messageError(&dialog,QString::fromUtf8(error.what())); }
    });
    connect(pdf,&QPushButton::clicked,&dialog,[this,&dialog,name,packet] {
        const auto path=QFileDialog::getSaveFileName(&dialog,"PDF speichern",name->text()+".pdf","PDF (*.pdf)"); if(path.isEmpty()) return;
        try {
            const auto report=packet();
            AccessFormPreview document(report.name,report.layout,report.code,db_,this,false,{},std::nullopt,0,-1,report.recordSource);
            if(!document.exportReportPdf(path)) throw linaccess::VbaRuntimeError(3251,"Bericht konnte nicht vollständig ausgegeben werden.");
            (void)document.close();
        } catch(const std::exception& error) { messageError(this,QString::fromUtf8(error.what())); }
    });
    connect(save,&QPushButton::clicked,&dialog,[this,&dialog,name,definition,packet] {
        try {
            (void)packet();
            for (const auto& document:AccessFormPreview::projectDocuments(db_))
                if (document && document->documentKind()==linaccess::AccessDocumentKind::Report && document->documentName().compare(name->text(),Qt::CaseInsensitive)==0)
                    if (!document->close()) return;
            if(saveObject("report",name->text(),definition())) {
                tagProjectObject(&dialog,this,db_,"report",name->text()); completeDesignerSave(dialog,name->text());
            }
        } catch(const std::exception& error) { messageError(this,QString::fromUtf8(error.what())); }
    });
    trackDesignerChanges(dialog);
    dialog.exec();
}

bool MainWindow::canCloseDataPage(QWidget* page) {
    if(!page) return true;
    for(auto* candidate:page->findChildren<QTableView*>()) if(auto* view=dynamic_cast<linaccess::AccessTableView*>(candidate)) {
        const auto* model=qobject_cast<linaccess::AccessTableModel*>(view->model());
        if((model && model->isDirty()) || view->hasPendingEditor()) {
            statusBar()->showMessage("Zelleingabe abschließen und Tabellenänderungen speichern oder verwerfen.",10000); return false;
        }
    }
    return true;
}
bool MainWindow::allowProjectClose(const QString& connection) {
    if(!connection.isEmpty() && connection!=db_.connectionName()) return true;
    if(queryExecutionDepth_>0) { statusBar()->showMessage("Projektwechsel während einer laufenden Abfrage ist noch nicht verfügbar.",10000); return false; }
    if(macroExecutionDepth_>0 || linaccess::nativeProjectMacrosActive(db_)) {statusBar()->showMessage("Projektwechsel während eines laufenden Makros ist nicht verfügbar.",10000);return false;}
    for(int i=0;i<tabs_->count();++i) if(!canCloseDataPage(tabs_->widget(i))) return false;
    return true;
}

void MainWindow::closeEvent(QCloseEvent* event) {
    if (!AccessFormPreview::closeProjectDocuments(db_)) event->ignore();
    else QMainWindow::closeEvent(event);
}
