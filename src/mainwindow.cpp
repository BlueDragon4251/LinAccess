#include "mainwindow.hpp"
#include "native_input_box.hpp"
#include "native_edit_menu.hpp"
#include <QCloseEvent>
#include <QKeyEvent>
#include <QShortcut>

#include "access_inspector.hpp"
#include "access_form_import.hpp"
#include "access_form_preview.hpp"
#include "vba_module_context.hpp"
#include "access_form_objects.hpp"
#include "bridge_importer.hpp"
#include "vba_analyzer.hpp"
#include "vba_runtime.hpp"
#include "dao_database.hpp"
#include "vba_qt_value.hpp"
#include "access_query.hpp"
#include "jet_sql_functions.hpp"
#include "access_table_model.hpp"
#include "access_table_delegate.hpp"
#include "access_table_view.hpp"
#include "access_format.hpp"
#include "vba_source_codec.hpp"

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
QString asText(const QVariant& value) {
    return value.isNull() ? QString{} : value.toString();
}
bool isSelect(const QString& sql) {
    return sql.trimmed().startsWith("SELECT ", Qt::CaseInsensitive)
        || sql.trimmed().startsWith("SELECT\n", Qt::CaseInsensitive);
}

class QtFormVbaHost final : public linaccess::VbaHost {
public:
    QtFormVbaHost(QWidget* parent,
                  const std::vector<std::pair<QString, QLineEdit*>>& edits,
                  std::function<bool()> persist, std::function<void()> gotoLast,
                  QSqlDatabase database, std::function<void(QString)> openForm)
        : parent_(parent), edits_(edits), persist_(std::move(persist)),
          gotoLast_(std::move(gotoLast)), database_(std::move(database)),
          openForm_(std::move(openForm)) {}

    std::optional<std::string> field(std::string_view name) override {
        auto* editor = lookup(name);
        if (editor->text().isEmpty()) return std::nullopt;
        return editor->text().toUtf8().toStdString();
    }
    void setField(std::string_view name, std::optional<std::string> value) override {
        lookup(name)->setText(value ? QString::fromUtf8(value->data(),
                                                       static_cast<qsizetype>(value->size())) : QString{});
    }
    linaccess::VbaValue fieldValue(std::string_view name) override {
        auto* editor = lookup(name);
        return linaccess::qtFieldValue(editor->property("la_null").toBool() ? QVariant{} : QVariant(editor->text()),
                                      editor->property("la_sql_type").toString());
    }
    void setFieldValue(std::string_view name,const linaccess::VbaValue& value) override {
        auto* editor = lookup(name);
        const auto converted = linaccess::sqlFieldValue(value,editor->property("la_sql_type").toString());
        editor->setText(converted.toString()); editor->setProperty("la_null",converted.isNull());
    }
    void message(std::string_view text) override {
        QMessageBox::information(parent_, "VBA MsgBox",
            QString::fromUtf8(text.data(), static_cast<qsizetype>(text.size())));
    }
    std::string inputBox(const linaccess::VbaInputBoxArguments& arguments) override {
        return linaccess::nativeInputBox(parent_,arguments);
    }
    void saveRecord() override {
        if (!persist_()) throw std::runtime_error("DoCmd: Datensatz konnte nicht gespeichert werden");
    }
    void quitApplication() override { AccessFormPreview::quitApplication(); }
    void maximizeForm() override { parent_->showMaximized(); }
    void gotoLastRecord() override { gotoLast_(); }
    linaccess::VbaValue formatValue(const linaccess::VbaValue& value,std::string_view mask,bool stringResult) override {
        return linaccess::formatAccessValue(value,mask,stringResult);
    }
    std::shared_ptr<linaccess::VbaObject> currentDatabase() override { return linaccess::makeDaoDatabase(database_); }
    void executeSql(std::string_view statement) override {
        auto database = linaccess::makeDaoDatabase(database_);
        (void)database->member("Execute",{linaccess::VbaValue::string(std::string(statement))});
    }
    void openForm(std::string_view name) override {
        openForm_(QString::fromUtf8(name.data(), name.size()));
    }
private:
    QLineEdit* lookup(std::string_view name) {
        const auto target = QString::fromUtf8(name.data(), static_cast<qsizetype>(name.size()));
        for (const auto& [fieldName, editor] : edits_)
            if (fieldName.compare(target, Qt::CaseInsensitive) == 0) return editor;
        throw std::runtime_error("Unbekanntes Me!-Formularfeld: " + std::string(name));
    }
    QWidget* parent_;
    const std::vector<std::pair<QString, QLineEdit*>>& edits_;
    std::function<bool()> persist_;
    std::function<void()> gotoLast_;
    QSqlDatabase database_;
    std::function<void(QString)> openForm_;
};
}

MainWindow::~MainWindow() { disconnect(tabs_,nullptr,this,nullptr); linaccess::resetVbaProjectState(db_); }

MainWindow::MainWindow(QWidget* parent) : QMainWindow(parent) {
    setWindowTitle("LinAccess — C++ / Qt 6");
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
    action("Aktualisieren", [this] { refreshExplorer(); });
    auto* fileMenu=menuBar()->addMenu("&Datei");
    const auto projectActions=toolbar->actions();
    for (int i=0;i<4;++i) fileMenu->addAction(projectActions[i]);
    fileMenu->addSeparator();
    fileMenu->addAction("Beenden",this,[this] { close(); });
    auto* editMenu=linaccess::addNativeEditMenu(menuBar(),this);
    editMenu->addSeparator();
    auto* findTable=editMenu->addAction("&Suchen...\tStrg+F"); findTable->setObjectName("table-find"); findTable->setEnabled(false);
    auto* findNext=editMenu->addAction("&Weitersuchen\tUmschalt+F4"); findNext->setObjectName("table-find-next"); findNext->setEnabled(false);
    connect(findTable,&QAction::triggered,this,[this] { (void)dispatchTableCommand(30); });
    connect(findNext,&QAction::triggered,this,[this] { (void)dispatchTableCommand(341); });
    connect(editMenu,&QMenu::aboutToShow,this,[this,findTable,findNext] {
        linaccess::AccessTableView* view=nullptr;
        if(!activePreviewDocument() && AccessFormPreview::databaseWindowActive(this) && activeDataPage_ && tabs_->currentWidget()==activeDataPage_)
            for(auto* candidate:activeDataPage_->findChildren<QTableView*>()) if((view=dynamic_cast<linaccess::AccessTableView*>(candidate))) break;
        const auto* model=view ? qobject_cast<linaccess::AccessTableModel*>(view->model()) : nullptr;
        bool textColumn=false;
        if(model) for(int column=0;column<model->columnCount();++column) if(!view->isColumnHidden(column) && model->isSearchableColumn(column)) { textColumn=true; break; }
        findTable->setEnabled(textColumn && !QApplication::activeModalWidget());
        findNext->setEnabled(view && view->hasFindText() && !QApplication::activeModalWidget());
    });
    auto* objectMenu=menuBar()->addMenu("&Objekte");
    for (int i=5;i<9;++i) objectMenu->addAction(projectActions[i]);
    auto* viewMenu=menuBar()->addMenu("&Ansicht");
    viewMenu->addAction(projectActions.last());
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
    statusBar()->showMessage("Neue SQLite-Datenbank erstellen oder Access-Datei untersuchen");
}

int MainWindow::selectDatabaseObject(int objectType,const QString& name,bool activate) {
    if(!db_.isOpen()) return 3078;
    const QStringList kinds{"table","query","form","report"};
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

QVariantMap MainWindow::saveDataPage(QWidget* page,bool interactive) {
    QPointer<MainWindow> self=this;
    QPointer<linaccess::AccessTableView> table;
    if(page) for(auto* candidate:page->findChildren<QTableView*>()) if(auto* view=dynamic_cast<linaccess::AccessTableView*>(candidate)) { table=view; break; }
    QPointer<linaccess::AccessTableModel> model=table ? qobject_cast<linaccess::AccessTableModel*>(table->model()) : nullptr;
    if(!table || !model) return tableCommandFailure(2046,"Kein aktives natives Tabellendokument.");
    if(model->isReadOnly()) return tableCommandFailure(2046,"Speichern ist in der schreibgeschützten Tabelle nicht verfügbar.");
    const auto failure=[&](int fallback) { return tableModelFailure(model,fallback); };
    if(!table->finishEditing()) return failure(2501);
    if(!self || !model || !table) return tableCommandFailure(2501,"Tabellendokument wurde während der Eingabevalidierung geschlossen.");
    const auto selected=table->currentIndex();
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

QVariantMap MainWindow::dispatchTableCommand(int command) {
    const QPointer<QWidget> page=activeDataPage_;
    if(!db_.isOpen() || activePreviewDocument() || !AccessFormPreview::databaseWindowActive(this) || !page || tabs_->indexOf(page)<0 || tabs_->currentWidget()!=page) return {};
    if(QApplication::activeModalWidget()) return tableCommandFailure(2046,"Tabellenbefehl ist durch ein modales Fenster blockiert.");
    if(command==97) return saveDataPage(page,false);
    if(command>=65 && command<=68) {
        const int actions[]{1,0,2,3};
        return navigateTableDocument({},actions[command-65],1,false);
    }
    if(command==28 || command==223 || command==50 || command==109 || command==189 || command==190 || command==191 || command==199 || command==292 || command==179 || command==163 || command==164 || command==30 || command==341) {
        QPointer<linaccess::AccessTableView> table;
        for(auto* candidate:page->findChildren<QTableView*>()) if(auto* view=dynamic_cast<linaccess::AccessTableView*>(candidate)) { table=view; break; }
        QPointer<linaccess::AccessTableModel> model=table ? qobject_cast<linaccess::AccessTableModel*>(table->model()) : nullptr;
        if(!table || !model) return tableCommandFailure(2046,"Keine aktive Tabellenzelle verfügbar.");
        if(command==189 || command==190 || command==191 || command==199 || command==292) {
            QPointer<QLineEdit> editor=qobject_cast<QLineEdit*>(QApplication::focusWidget());
            QWidget* parent=editor.data();
            while(parent && parent!=table.data()) parent=parent->parentWidget();
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
        }
        if((command!=28 && command!=109 && command!=30 && command!=341) && !table->currentIndex().isValid()) return tableCommandFailure(2046,"Keine aktive Tabellenzelle verfügbar.");
        if(command==28) {
            const QPointer<MainWindow> self=this;
            if(!table->finishEditing()) return tableModelFailure(model,2501);
            if(!self || !table || !model) return tableCommandFailure(2501,"Tabellendokument wurde während der Validierung geschlossen.");
            if(model->isDirty()) {
                if(!model->canSubmitCurrentRecordAlone(table->currentIndex().row()))
                    return tableCommandFailure(3251,"Neuer Datensatz mit weiteren Tabellenentwürfen ist noch nicht implementiert.");
                const auto saved=saveDataPage(page,false);
                if(saved.value("error").toInt()) return saved;
            }
            if(!self || !table || !model) return tableCommandFailure(2501,"Tabellendokument wurde während des Speicherns geschlossen.");
            if(!table->insertRecord()) return tableModelFailure(model,2501);
        } else if(command==223) {
            if(!table->deleteCurrentRecord()) return tableModelFailure(model,2501);
        } else if(command==30 || command==341) {
            if(!(command==30 ? table->openFindDialog() : table->repeatFind())) return tableModelFailure(model,2046);
        } else if(command==109) {
            if(!table->selectAllRecords()) return tableModelFailure(model,2501);
        } else if(command==190) {
            if(!table->copySelection()) return tableModelFailure(model,2501);
        } else if(command==191) {
            if(!table->pasteSelection()) return tableModelFailure(model,2046);
        } else if(command==189 || command==199 || command==292) {
            return tableCommandFailure(3251,"Dieser Bearbeitungsbefehl außerhalb eines aktiven Textfelds ist noch nicht implementiert.");
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
    QPointer<linaccess::AccessTableView> table;
    for(auto* candidate:page->findChildren<QTableView*>()) if(auto* view=dynamic_cast<linaccess::AccessTableView*>(candidate)) { table=view; break; }
    QPointer<linaccess::AccessTableModel> model=table ? qobject_cast<linaccess::AccessTableModel*>(table->model()) : nullptr;
    if(!table || !model) return {};
    const QPointer<MainWindow> self=this;
    if(!table->finishEditing()) return tableModelFailure(model,2501);
    if(!self || !table || !model) return tableCommandFailure(2501,"Tabellendokument wurde während der Validierung geschlossen.");
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
        const bool sensitive=request.value("sensitive").toBool(),first=request.value("first").toBool();
        if(mode<0 || mode>2 || direction<0 || direction>2) return tableCommandFailure(5,"Ungültiger FindRecord-Vergleich oder Suchrichtung.");
        const bool textValue=sought.kind==linaccess::VbaValue::Kind::Text;
        const auto text=linaccess::nativeFindValueText(sought);
        if(textValue && (text.isEmpty() || text.contains('*') || text.contains('?') || text.contains('#') || text.contains('[') || text.contains(']')))
            return tableCommandFailure(3251,"Leere und Wildcard-FindRecord-Suchen sind noch nicht implementiert.");
        if(!textValue && mode!=1) return tableCommandFailure(3251,"Typisierte FindRecord-Suche unterstützt nur ganze Feldwerte.");
        QPointer<linaccess::AccessTableView> table;
        for(auto* candidate:page->findChildren<QTableView*>()) if(auto* view=dynamic_cast<linaccess::AccessTableView*>(candidate)) { table=view; break; }
        QPointer<linaccess::AccessTableModel> model=table ? qobject_cast<linaccess::AccessTableModel*>(table->model()) : nullptr;
        if(!table || !model || !table->currentIndex().isValid()) return tableCommandFailure(2046,"Kein aktuelles Tabellenfeld für FindRecord.");
        const int column=table->currentIndex().column();
        if(table->isColumnHidden(column) || !linaccess::nativeFindSupports(sought,model->fieldSqlType(column)))
            return tableCommandFailure(3251,"FindRecord-Suchwert und aktueller Feldtyp passen nicht zum unterstützten Teilumfang.");
        linaccess::NativeTableFindArguments arguments{text,column,mode,sensitive,first ? 1 : 2,direction};
        if(!textValue) arguments.value=sought;
        if(table->findText(std::move(arguments))<0) return tableModelFailure(model,2501);
        return {{"handled",true},{"error",0}};
    } catch(const linaccess::VbaRuntimeError& error) { return tableCommandFailure(error.number,QString::fromUtf8(error.what())); }
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
    QPointer<linaccess::AccessTableView> table;
    for(auto* candidate:page->findChildren<QTableView*>()) if(auto* view=dynamic_cast<linaccess::AccessTableView*>(candidate)) { table=view; break; }
    QPointer<linaccess::AccessTableModel> model=table ? qobject_cast<linaccess::AccessTableModel*>(table->model()) : nullptr;
    if(!table || !model) return tableCommandFailure(2046,"Kein natives Tabellendokument.");
    if(action==5) return dispatchTableCommand(28);
    if(!table->finishEditing()) return tableModelFailure(model,2501);
    if(!self || !table || !model) return tableCommandFailure(2501,"Tabellendokument wurde während der Validierung geschlossen.");
    // Only the current existing record may be saved here. A manual-submit cache
    // can also contain unrelated drafts, insertions and deletions.
    if(model->isDirty()) {
        if(!model->canSubmitCurrentRecordAlone(table->currentIndex().row()))
            return tableCommandFailure(3251,"Navigation mit mehreren Tabellenentwürfen oder neuen/gelöschten Zeilen ist noch nicht implementiert.");
        const auto saved=saveDataPage(page,false);
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
    QPointer<MainWindow> self=this;
    QPointer<QWidget> page;
    if(name.isEmpty()) { if(active) page=activeDataPage_; }
    else page=findOpenTable(name);
    if(!page) return tableCommandFailure(3251,"Die angeforderte offene Tabelle ist nicht verfügbar.");
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
    for(auto* candidate:selected->findChildren<QTableView*>()) if(auto* view=dynamic_cast<linaccess::AccessTableView*>(candidate)) {
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
    const auto projectPath=QFileInfo(path).absoluteFilePath();
    if (!AccessFormPreview::closeProjectDocuments(db_)) return;
    setProperty("la_project_generation",QVariant::fromValue<qulonglong>(property("la_project_generation").toULongLong()+1));
    setProperty("la_native_current_projects",QVariant{});
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
    query.prepare("SELECT name FROM _la_objects WHERE kind=? ORDER BY name COLLATE NOCASE");
    query.addBindValue(kind);
    if (query.exec()) while (query.next()) names << query.value(0).toString();
    return names;
}

QJsonObject MainWindow::objectDefinition(const QString& kind, const QString& name) {
    QSqlQuery query(db_);
    query.prepare("SELECT definition FROM _la_objects WHERE kind=? AND name=?");
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
    query.prepare("INSERT INTO _la_objects(kind,name,definition) VALUES(?,?,?) "
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
    explorer_->clear();
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
}

void MainWindow::openExplorerItem(QTreeWidgetItem* item, int) {
    if (!item || !hasProject()) return;
    const auto kind = item->data(0, Qt::UserRole).toString();
    const auto name = item->text(0);
    if (kind == "table") showTable(name);
    else if (kind == "query") editQuery(name);
    else if (kind == "form") runForm(name);
    else if (kind == "report") editReport(name);
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
            existing.prepare("SELECT 1 FROM _la_objects WHERE kind='form' AND name=?");
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
    QDialog dialog(this);
    dialog.setWindowTitle("Neue Tabelle");
    dialog.resize(610, 440);
    auto* outer = new QVBoxLayout(&dialog);
    auto* name = new QLineEdit(&dialog);
    name->setPlaceholderText("Tabellenname");
    outer->addWidget(name);
    outer->addWidget(new QLabel("Felder (id wird automatisch als Primärschlüssel angelegt)"));
    auto* grid = new QTableWidget(0, 3, &dialog);
    grid->setHorizontalHeaderLabels({"Feldname", "Datentyp", "Pflichtfeld (ja/nein)"});
    grid->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);
    outer->addWidget(grid);
    auto addRow = [grid] {
        const int row = grid->rowCount();
        grid->insertRow(row);
        grid->setItem(row, 0, new QTableWidgetItem());
        auto* type = new QComboBox(grid);
        type->addItems({"TEXT", "INTEGER", "REAL", "DATE", "BOOLEAN", "BLOB"});
        grid->setCellWidget(row, 1, type);
        grid->setItem(row, 2, new QTableWidgetItem("nein"));
    };
    auto* plus = addButton(outer, "+ Feld");
    connect(plus, &QPushButton::clicked, &dialog, addRow);
    addRow();
    auto* buttons = new QDialogButtonBox(QDialogButtonBox::Save | QDialogButtonBox::Cancel, &dialog);
    outer->addWidget(buttons);
    connect(buttons, &QDialogButtonBox::accepted, &dialog, &QDialog::accept);
    connect(buttons, &QDialogButtonBox::rejected, &dialog, &QDialog::reject);
    if (dialog.exec() != QDialog::Accepted) return;
    const auto tableName = name->text().trimmed();
    if (tableName.isEmpty() || tableName.startsWith("_la_", Qt::CaseInsensitive) ||
        tableName.startsWith("sqlite_", Qt::CaseInsensitive)) {
        messageError(this, "Ungültiger/reservierter Tabellenname");
        return;
    }
    QStringList columns = {"\"id\" INTEGER PRIMARY KEY AUTOINCREMENT"};
    std::set<QString> seen = {"id"};
    for (int row = 0; row < grid->rowCount(); ++row) {
        const auto field = grid->item(row, 0)->text().trimmed();
        if (field.isEmpty()) continue;
        if (field.startsWith("_la_", Qt::CaseInsensitive) || seen.contains(field.toCaseFolded())) {
            messageError(this, "Reservierter oder doppelter Feldname: " + field);
            return;
        }
        seen.insert(field.toCaseFolded());
        const auto type = qobject_cast<QComboBox*>(grid->cellWidget(row, 1))->currentText();
        const auto required = grid->item(row, 2)->text().trimmed().toLower();
        if (required != "ja" && required != "nein" && !required.isEmpty()) {
            messageError(this, "Pflichtfeld muss ja oder nein sein");
            return;
        }
        columns << quoteSql(field) + " " + type + (required == "ja" ? " NOT NULL" : "");
    }
    if (columns.size() == 1) {
        messageError(this, "Mindestens ein Feld anlegen");
        return;
    }
    QSqlQuery query(db_);
    if (!query.exec("CREATE TABLE " + quoteSql(tableName) + " (" + columns.join(",") + ")"))
        messageError(this, query.lastError().text());
    refreshExplorer();
}

int MainWindow::showTable(const QString& name,bool readOnly,bool interactive) {
    if(!db_.isOpen() || QApplication::activeModalWidget()) return 2046;
    if(auto* existing=findOpenTable(name)) {
        if(!interactive) for(auto* candidate:existing->findChildren<QTableView*>()) if(auto* view=dynamic_cast<linaccess::AccessTableView*>(candidate)) {
            const auto* model=qobject_cast<linaccess::AccessTableModel*>(view->model());
            if(model && model->isReadOnly()!=readOnly) return 3251;
        }
        return activateDataPage(existing) ? 0 : 2501;
    }
    if(!finishActiveDataEditors() || !AccessFormPreview::activateProjectDocument(db_,this,nullptr)) return 2501;
    auto* page = new QWidget(tabs_);
    tagProjectObject(page,this,db_,"table",name);
    page->setProperty("la_project_object_scope","tab");
    auto* outer = new QVBoxLayout(page);
    outer->addWidget(new QLabel("Tabelle: " + name + (readOnly ? " — Schreibgeschützt" : " — Änderungen erst mit 'Speichern' festschreiben")));
    auto* actions = new QHBoxLayout();
    auto* add = addButton(actions, "+ Datensatz");
    auto* remove = addButton(actions, "Datensatz löschen");
    auto* save = addButton(actions, "Speichern");
    auto* discard = addButton(actions, "Verwerfen");
    for(auto* button:{add,remove,save,discard}) button->setEnabled(!readOnly);
    outer->addLayout(actions);
    auto* model = new linaccess::AccessTableModel(page, db_);
    connect(model,&linaccess::AccessTableModel::editRejected,page,[this](int number,const QString& message) { statusBar()->showMessage(message+" ("+QString::number(number)+")",10000); });
    connect(model,&QAbstractItemModel::dataChanged,page,[this] { statusBar()->clearMessage(); });
    model->setTable(name);
    model->setEditStrategy(QSqlTableModel::OnManualSubmit);
    (void)model->setReadOnly(readOnly);
    if(!model->select()) {
        const auto problem=model->lastError();
        bool valid=false; const int code=problem.nativeErrorCode().toInt(&valid);
        if(interactive) messageError(this,problem.text());
        delete page; return problem.driverText()=="LinAccess" && valid ? code : 3146;
    }
    auto* table = new linaccess::AccessTableView(page);
    table->installEventFilter(this);
    table->setModel(model);
    table->setItemDelegate(new linaccess::AccessTableDelegate(table));
    table->setRecordNavigationGuard([self=QPointer<MainWindow>(this),page=QPointer<QWidget>(page),table=QPointer<linaccess::AccessTableView>(table),model=QPointer<linaccess::AccessTableModel>(model)] {
        if(!self || !page || !table || !model) return false;
        if(QApplication::activeModalWidget()) return model->rejectEdit(2046,"Suche ist durch ein modales Fenster blockiert.");
        if(self->tabs_->indexOf(page)<0) return model->rejectEdit(2501,"Tabellendokument ist nicht mehr geöffnet.");
        if((self->tabs_->currentWidget()!=page || activePreviewDocument() || !AccessFormPreview::databaseWindowActive(self)) && !self->activateDataPage(page))
            return model ? model->rejectEdit(2501,"Tabellenaktivierung wurde vor der Suche abgebrochen.") : false;
        if(!self || !page || !table || !model) return false;
        if(!model->isDirty()) return true;
        if(!model->canSubmitCurrentRecordAlone(table->currentIndex().row())) return model->rejectEdit(3251,"Suche mit weiteren Tabellenentwürfen ist noch nicht implementiert.");
        const auto saved=self->saveDataPage(page,false);
        if(const int error=saved.value("error").toInt()) {
            if(model) return model->rejectEdit(error,saved.value("description").toString());
            return false;
        }
        return bool(self && page && table && model);
    });
    table->setSortingEnabled(true);
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
    tabs_->setCurrentIndex(tabs_->addTab(page, name));
    return activateDataPage(page) ? 0 : 2501;
}

void MainWindow::editQuery(const QString& existingName) {
    if (!hasProject()) return;
    QDialog dialog(this);
    dialog.setWindowTitle("SQL-Abfrage");
    tagProjectObject(&dialog,this,db_,"query",existingName);
    dialog.resize(920, 700);
    auto* outer = new QVBoxLayout(&dialog);
    auto* name = new QLineEdit(existingName, &dialog);
    name->setPlaceholderText("Name der Abfrage");
    outer->addWidget(name);
    auto* sql = new QPlainTextEdit(&dialog);
    sql->setPlaceholderText("SELECT * FROM kunden");
    sql->setMaximumHeight(190);
    if (!existingName.isEmpty()) sql->setPlainText(objectDefinition("query", existingName).value("sql").toString());
    outer->addWidget(sql);
    auto* buttons = new QHBoxLayout();
    auto* run = addButton(buttons, "Ausführen");
    auto* save = addButton(buttons, "Speichern");
    outer->addLayout(buttons);
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
    connect(save, &QPushButton::clicked, &dialog, [this, name, sql, &dialog] {
        try {
            (void)linaccess::compileAccessQuery(sql->toPlainText());
        } catch(const std::exception& error) { messageError(&dialog,QString::fromUtf8(error.what())); return; }
        if (saveObject("query", name->text(), {{"sql", sql->toPlainText()}})) dialog.accept();
    });
    dialog.exec();
}

void MainWindow::editForm(const QString& existingName) {
    if (!hasProject()) return;
    if (tableNames().isEmpty()) {
        QMessageBox::information(this, "Formular", "Zuerst eine Tabelle erstellen.");
        return;
    }
    const auto old = existingName.isEmpty() ? QJsonObject{} : objectDefinition("form", existingName);
    QDialog dialog(this);
    dialog.setWindowTitle("Formulardesigner — experimentelle VBA-Ereignislaufzeit");
    tagProjectObject(&dialog,this,db_,"form",existingName);
    dialog.resize(820, 690);
    auto* outer = new QVBoxLayout(&dialog);
    auto* name = new QLineEdit(existingName, &dialog);
    name->setPlaceholderText("Formularname");
    outer->addWidget(name);
    auto* source = new QComboBox(&dialog);
    source->addItems(tableNames());
    outer->addWidget(new QLabel("Datenquelle"));
    outer->addWidget(source);
    auto* fields = new QListWidget(&dialog);
    fields->setMaximumHeight(170);
    outer->addWidget(new QLabel("Formularfelder"));
    outer->addWidget(fields);
    auto populate = [this, fields, source, old] {
        fields->clear();
        const auto table = source->currentText();
        if (table.isEmpty()) return;
        QSqlQuery query(db_);
        query.exec("PRAGMA table_info(" + quoteSql(table) + ")");
        const auto previous = old.value("fields").toArray();
        const bool sameSource = old.value("table").toString() == table;
        while (query.next()) {
            const auto field = query.value(1).toString();
            if (field == "id") continue;
            auto* item = new QListWidgetItem(field, fields);
            item->setFlags(item->flags() | Qt::ItemIsUserCheckable);
            const bool checked = !sameSource || previous.isEmpty() || previous.contains(field);
            item->setCheckState(checked ? Qt::Checked : Qt::Unchecked);
        }
    };
    connect(source, &QComboBox::currentTextChanged, &dialog, populate);
    if (!old.isEmpty()) source->setCurrentText(old.value("table").toString());
    populate();
    outer->addWidget(new QLabel(
        "VBA-Modul — experimenteller Teilinterpreter für wenige Form-/Button-Ereignisse; "
        "keine vollständige Microsoft-VBA-Kompatibilität"));
    auto* vba = new QPlainTextEdit(&dialog);
    vba->setPlaceholderText("Private Sub Form_Load()\n    MsgBox \"Hallo\"\nEnd Sub");
    vba->setPlainText(old.value("vba").toString());
    auto importedBytes = std::make_shared<QByteArray>(QByteArray::fromBase64(
        old.value("original_vba_base64").toString().toLatin1()));
    outer->addWidget(vba, 1);
    auto* row = new QHBoxLayout();
    auto* importBas = addButton(row, ".bas/.cls laden");
    auto* importBridge = addButton(row, "VBA aus Bridge übernehmen");
    auto* analyze = addButton(row, "VBA analysieren");
    auto* save = addButton(row, "Formular speichern");
    outer->addLayout(row);
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
    connect(save, &QPushButton::clicked, &dialog, [this, &dialog, name, source, fields, vba, importedBytes] {
        QJsonArray chosen;
        for (int i = 0; i < fields->count(); ++i)
            if (fields->item(i)->checkState() == Qt::Checked) chosen.append(fields->item(i)->text());
        if (chosen.isEmpty()) { messageError(this, "Mindestens ein Feld auswählen."); return; }
        const QJsonObject content = {{"table", source->currentText()},
                                     {"fields", chosen}, {"vba", vba->toPlainText()},
                                     {"original_vba_base64", QString::fromLatin1(importedBytes->toBase64())},
                                     {"language", "vba_subset_experimental"}};
        if (saveObject("form", name->text(), content)) dialog.accept();
    });
    dialog.exec();
}

void MainWindow::runForm(const QString& name) {
    if (!hasProject()) return;
    const auto definition = objectDefinition("form", name);
    const auto tableName = definition.value("table").toString();
    const auto fieldsJson = definition.value("fields").toArray();
    if (tableName.isEmpty() || fieldsJson.isEmpty()) {
        messageError(this, "Formulardefinition unvollständig.");
        return;
    }
    QSqlQuery schema(db_);
    schema.exec("PRAGMA table_info(" + quoteSql(tableName) + ")");
    bool hasId = false;
    while (schema.next()) if (schema.value(1).toString() == "id") hasId = true;
    // Access-imported tables do not necessarily have a column literally named
    // "id". They are ordinary SQLite rowid tables in this import version.
    const QString keyExpr = hasId ? QStringLiteral("\"id\"") : QStringLiteral("rowid");
    QDialog dialog(this);
    dialog.setWindowTitle("Formular: " + name);
    tagProjectObject(&dialog,this,db_,"form",name);
    dialog.resize(620, 510);
    auto* outer = new QVBoxLayout(&dialog);
    outer->addWidget(new QLabel("<b>VBA: experimenteller Teilinterpreter. "
                                "DAO unterstützt nur einfache Tabellenrecordsets; Windows-API und COM sind nicht implementiert.</b>"));
    auto* records = new QComboBox(&dialog);
    outer->addWidget(records);
    auto* form = new QFormLayout();
    std::vector<std::pair<QString, QLineEdit*>> entries;
    const auto fieldTypes = linaccess::sqliteColumnTypes(db_,tableName);
    for (const auto& field : fieldsJson) {
        const auto column = field.toString();
        auto* edit = new QLineEdit(&dialog);
        edit->setProperty("la_sql_type",fieldTypes.value(column.toCaseFolded()));
        edit->setProperty("la_null",true);
        connect(edit,&QLineEdit::textEdited,edit,[edit] { edit->setProperty("la_null",false); });
        entries.emplace_back(column, edit);
        form->addRow(column, edit);
    }
    outer->addLayout(form);
    auto reload = [this, tableName, keyExpr, records, entries](const QVariant& selected) {
        records->blockSignals(true);
        records->clear();
        records->addItem("(Neuer Datensatz)", QVariant());
        QSqlQuery query(db_);
        if (query.exec("SELECT " + keyExpr + " FROM " + quoteSql(tableName) + " LIMIT 500")) {
            while (query.next()) records->addItem("Datensatz #" + query.value(0).toString(), query.value(0));
        }
        if (!selected.isNull()) {
            const auto idx = records->findData(selected);
            if (idx >= 0) records->setCurrentIndex(idx);
        }
        records->blockSignals(false);
        for (const auto& entry : entries) { entry.second->clear(); entry.second->setProperty("la_null",true); }
        if (records->currentData().isNull()) return;
        QSqlQuery row(db_);
        row.prepare("SELECT * FROM " + quoteSql(tableName) + " WHERE " + keyExpr + "=?");
        row.addBindValue(records->currentData());
        if (row.exec() && row.next())
            for (const auto& entry : entries)
                {
                    const auto value = row.value(row.record().indexOf(entry.first));
                    entry.second->setText(asText(value)); entry.second->setProperty("la_null",value.isNull());
                }
    };
    connect(records, &QComboBox::currentIndexChanged, &dialog, [this, tableName, keyExpr, records, entries] {
        for (const auto& entry : entries) { entry.second->clear(); entry.second->setProperty("la_null",true); }
        if (records->currentData().isNull()) return;
        QSqlQuery row(db_);
        row.prepare("SELECT * FROM " + quoteSql(tableName) + " WHERE " + keyExpr + "=?");
        row.addBindValue(records->currentData());
        if (row.exec() && row.next())
            for (const auto& entry : entries)
                {
                    const auto value = row.value(row.record().indexOf(entry.first));
                    entry.second->setText(asText(value)); entry.second->setProperty("la_null",value.isNull());
                }
    });
    auto* buttons = new QHBoxLayout();
    auto* create = addButton(buttons, "Neu");
    auto* save = addButton(buttons, "Speichern");
    auto* action = addButton(buttons, "Aktion (btnAktion_Click)");
    auto* designer = addButton(buttons, "Entwurf bearbeiten");
    outer->addLayout(buttons);
    connect(create, &QPushButton::clicked, records, [records] { records->setCurrentIndex(0); });
    const auto persist = [this, tableName, keyExpr, records, entries, reload]() -> bool {
        QStringList columns, placeholders, assignments;
        for (const auto& [column, _] : entries) {
            columns << quoteSql(column);
            placeholders << "?";
            assignments << quoteSql(column) + "=?";
        }
        const auto id = records->currentData();
        QSqlQuery query(db_);
        if (id.isNull())
            query.prepare("INSERT INTO " + quoteSql(tableName) + " (" + columns.join(",")
                          + ") VALUES (" + placeholders.join(",") + ")");
        else
            query.prepare("UPDATE " + quoteSql(tableName) + " SET " + assignments.join(",")
                          + " WHERE " + keyExpr + "=?");
        try {
            for (const auto& entry : entries) {
                auto* editor = entry.second; const auto type = editor->property("la_sql_type").toString();
                const auto value = linaccess::qtFieldValue(editor->property("la_null").toBool() ? QVariant{} : QVariant(editor->text()),type);
                query.addBindValue(linaccess::sqlFieldValue(value,type));
            }
        } catch (const std::exception& error) { messageError(this,QString::fromUtf8(error.what())); return false; }
        if (!id.isNull()) query.addBindValue(id);
        if (!query.exec()) { messageError(this, query.lastError().text()); return false; }
        reload(id.isNull() ? query.lastInsertId() : id);
        return true;
    };
    const auto vba = definition.value("vba").toString().toUtf8().toStdString();
    const bool enableVba = !vba.empty();
    const auto gotoLast = [records] {
        if (records->count() > 1) records->setCurrentIndex(records->count() - 1);
    };
    auto fire = [this, &dialog, entries, persist, gotoLast, enableVba, vba](std::string_view procedure) {
        if (!enableVba) return linaccess::VbaExecutionResult{
            linaccess::VbaExecutionStatus::ProcedureMissing, {}, 0};
        QtFormVbaHost host(&dialog, entries, persist, gotoLast, db_,
            [this](QString name) {
                QSqlQuery query(db_);
                query.prepare("SELECT source_bytes FROM _la_access_artifacts "
                              "WHERE kind='form' AND name=? ORDER BY id DESC LIMIT 1");
                query.addBindValue(name);
                if (!query.exec() || !query.next())
                    throw std::runtime_error("DoCmd.OpenForm: Formularquelle fehlt: " + name.toStdString());
                const auto form = query.value(0).toByteArray();
                query.finish();
                QSqlQuery code(db_);
                code.prepare("SELECT source_bytes FROM _la_access_artifacts "
                             "WHERE kind='vba' AND name=? ORDER BY id DESC LIMIT 1");
                code.addBindValue("Form_" + name);
                const auto module = code.exec() && code.next() ? code.value(0).toByteArray() : QByteArray{};
                code.finish();
                (void)AccessFormPreview::showDocument(linaccess::AccessDocumentKind::Form,name,form,module,db_,this);
            });
        auto result = linaccess::executeVbaProcedure(vba, procedure, host);
        if (result.status == linaccess::VbaExecutionStatus::Unsupported ||
            result.status == linaccess::VbaExecutionStatus::Error) {
            messageError(this, "VBA-Fehler in " + QString::fromUtf8(procedure.data(),
                        static_cast<qsizetype>(procedure.size())) + ", Zeile "
                        + QString::number(result.line) + ": "
                        + QString::fromStdString(result.explanation));
        }
        return result;
    };
    connect(save, &QPushButton::clicked, &dialog, [persist, fire] {
        const auto result = fire("btnSpeichern_Click");
        if (result.status == linaccess::VbaExecutionStatus::ProcedureMissing) (void)persist();
        // When a VBA handler exists, it controls saving through DoCmd.RunCommand.
    });
    connect(action, &QPushButton::clicked, &dialog, [fire, &dialog] {
        if (fire("btnAktion_Click").status == linaccess::VbaExecutionStatus::ProcedureMissing)
            QMessageBox::information(&dialog, "VBA", "Keine btnAktion_Click-Prozedur vorhanden oder freigegeben.");
    });
    connect(designer, &QPushButton::clicked, &dialog, [&dialog] { dialog.accept(); });
    reload({});
    if (fire("Form_Open").cancelled) return;
    (void)fire("Form_Load");
    const auto openDesigner = dialog.exec() == QDialog::Accepted;
    if (openDesigner) editForm(name);
}

void MainWindow::editReport(const QString& existingName) {
    if (!hasProject()) return;
    const auto old = existingName.isEmpty() ? QJsonObject{} : objectDefinition("report", existingName);
    QDialog dialog(this);
    dialog.setWindowTitle("Berichtdesigner — einfache SQL-Tabellenberichte");
    tagProjectObject(&dialog,this,db_,"report",existingName);
    dialog.resize(930, 720);
    auto* outer = new QVBoxLayout(&dialog);
    auto* name = new QLineEdit(existingName, &dialog);
    name->setPlaceholderText("Berichtname");
    outer->addWidget(name);
    auto* title = new QLineEdit(old.value("title").toString(), &dialog);
    title->setPlaceholderText("Überschrift");
    outer->addWidget(title);
    auto* sql = new QPlainTextEdit(&dialog);
    sql->setMaximumHeight(150);
    sql->setPlaceholderText("SELECT * FROM kunden");
    sql->setPlainText(old.value("sql").toString());
    outer->addWidget(sql);
    auto* row = new QHBoxLayout();
    auto* previewButton = addButton(row, "Vorschau");
    auto* pdfButton = addButton(row, "PDF exportieren");
    auto* saveButton = addButton(row, "Speichern");
    outer->addLayout(row);
    auto* preview = new QTextBrowser(&dialog);
    outer->addWidget(preview);

    auto makeHtml = [this, sql, title]() -> QString {
        if (!isSelect(sql->toPlainText())) throw QString("Nur SELECT-Abfragen in Berichten.");
        QSqlQuery query(db_);
        if (!query.exec(sql->toPlainText())) throw query.lastError().text();
        const auto record = query.record();
        QString result = "<html><head><style>body{font-family:sans-serif}"
                         "table{border-collapse:collapse;width:100%}"
                         "th,td{padding:5px;border:1px solid gray}"
                         "th{background:#ddd}</style></head><body><h1>"
                         + title->text().toHtmlEscaped() + "</h1><table><tr>";
        for (int i = 0; i < record.count(); ++i)
            result += "<th>" + record.fieldName(i).toHtmlEscaped() + "</th>";
        result += "</tr>";
        int count = 0;
        while (count < 1000 && query.next()) {
            result += "<tr>";
            for (int i = 0; i < record.count(); ++i)
                result += "<td>" + asText(query.value(i)).toHtmlEscaped() + "</td>";
            result += "</tr>";
            ++count;
        }
        return result + "</table><p>" + QString::number(count)
            + " Zeilen (max. 1000)</p></body></html>";
    };
    connect(previewButton, &QPushButton::clicked, &dialog, [this, preview, makeHtml] {
        try { preview->setHtml(makeHtml()); }
        catch (const QString& problem) { messageError(this, problem); }
    });
    connect(pdfButton, &QPushButton::clicked, &dialog, [this, name, makeHtml] {
        const auto path = QFileDialog::getSaveFileName(this, "PDF speichern", name->text() + ".pdf", "PDF (*.pdf)");
        if (path.isEmpty()) return;
        try {
            QTextDocument document;
            document.setHtml(makeHtml());
            QPrinter printer(QPrinter::HighResolution);
            printer.setOutputFormat(QPrinter::PdfFormat);
            printer.setOutputFileName(path);
            document.print(&printer);
            QMessageBox::information(this, "Export", "PDF gespeichert: " + path);
        } catch (const QString& problem) { messageError(this, problem); }
    });
    connect(saveButton, &QPushButton::clicked, &dialog, [this, name, title, sql, &dialog] {
        if (!isSelect(sql->toPlainText())) { messageError(this, "Nur SELECT-Abfragen speichern."); return; }
        if (saveObject("report", name->text(), {{"title", title->text()}, {"sql", sql->toPlainText()}}))
            dialog.accept();
    });
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
    for(int i=0;i<tabs_->count();++i) if(!canCloseDataPage(tabs_->widget(i))) return false;
    return true;
}

void MainWindow::closeEvent(QCloseEvent* event) {
    if (!AccessFormPreview::closeProjectDocuments(db_)) event->ignore();
    else QMainWindow::closeEvent(event);
}
