#include "mainwindow.hpp"

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
#include <QListWidget>
#include <QMessageBox>
#include <QPlainTextEdit>
#include <QPrinter>
#include <QPrintDialog>
#include <QPushButton>
#include <QSqlError>
#include <QSqlQuery>
#include <QSqlQueryModel>
#include <QSqlRecord>
#include <QSqlTableModel>
#include <QStatusBar>
#include <QSplitter>
#include <QTabWidget>
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
QString quoteSql(QString name) {
    name.replace('"', "\"\"");
    return '"' + name + '"';
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
    void saveRecord() override {
        if (!persist_()) throw std::runtime_error("DoCmd: Datensatz konnte nicht gespeichert werden");
    }
    void quitApplication() override { QApplication::quit(); }
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

MainWindow::~MainWindow() { linaccess::resetVbaProjectState(db_); }

MainWindow::MainWindow(QWidget* parent) : QMainWindow(parent) {
    setWindowTitle("LinAccess — C++ / Qt 6");
    resize(1260, 790);
    db_ = QSqlDatabase::addDatabase("QSQLITE", "linaccess-main");

    auto* splitter = new QSplitter(this);
    explorer_ = new QTreeWidget(splitter);
    explorer_->setHeaderHidden(true);
    explorer_->setMinimumWidth(220);
    tabs_ = new QTabWidget(splitter);
    tabs_->setTabsClosable(true);
    splitter->addWidget(explorer_);
    splitter->addWidget(tabs_);
    splitter->setStretchFactor(1, 1);
    setCentralWidget(splitter);
    connect(tabs_, &QTabWidget::tabCloseRequested, this, [this](int index) {
        QWidget* page = tabs_->widget(index);
        tabs_->removeTab(index);
        delete page;
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
    auto* objectMenu=menuBar()->addMenu("&Objekte");
    for (int i=5;i<9;++i) objectMenu->addAction(projectActions[i]);
    auto* viewMenu=menuBar()->addMenu("&Ansicht");
    viewMenu->addAction(projectActions.last());
    statusBar()->showMessage("Neue SQLite-Datenbank erstellen oder Access-Datei untersuchen");
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
    while (tabs_->count()) {
        auto* widget = tabs_->widget(0);
        tabs_->removeTab(0);
        delete widget;
    }
    linaccess::resetVbaProjectState(db_);
    db_.close();
    db_.setDatabaseName(path);
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
    databasePath_ = path;
    setWindowTitle("LinAccess — " + QFileInfo(path).fileName());
    statusBar()->showMessage(path);
    refreshExplorer();
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
    const auto group = [this](const QString& label, const QString& kind, const QStringList& names) {
        auto* parent = new QTreeWidgetItem(explorer_, {label});
        for (const auto& name : names) {
            auto* item = new QTreeWidgetItem(parent, {name});
            item->setData(0, Qt::UserRole, kind);
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
        auto* accessGroup = new QTreeWidgetItem(explorer_, {"Access-Quellobjekte (nicht ausführbar)"});
        while (artifacts.next()) {
            auto* item = new QTreeWidgetItem(accessGroup, {
                artifacts.value(1).toString() + ": " + artifacts.value(2).toString()
                + " [" + artifacts.value(3).toString() + "]"});
            item->setData(0, Qt::UserRole, "access_artifact");
            item->setData(0, Qt::UserRole + 1, artifacts.value(0).toInt());
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
    dialog.setWindowTitle(query.value(0).toString() + ": " + query.value(1).toString());
    dialog.resize(850, 680);
    auto* outer = new QVBoxLayout(&dialog);
    const auto kind = query.value(0).toString();
    const auto status = query.value(2).toString();
    outer->addWidget(new QLabel("Status: " + status
        + " | Format: " + query.value(3).toString()
        + " | Access-Layouts und vollständiges VBA sind NICHT kompatibel."));
    if (!query.value(4).toString().isEmpty())
        outer->addWidget(new QLabel(query.value(4).toString()));
    auto* source = new QPlainTextEdit(&dialog);
    source->setReadOnly(true);
    const auto bytes = query.value(5).toByteArray();
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
        const auto name = query.value(1).toString();
        connect(open, &QPushButton::clicked, &dialog, [this, name, bytes] {
            QSqlQuery vba(db_);
            vba.prepare("SELECT source_bytes FROM _la_access_artifacts WHERE kind='vba' AND name=? ORDER BY id DESC LIMIT 1");
            vba.addBindValue("Report_" + name);
            const auto module = vba.exec() && vba.next() ? vba.value(0).toByteArray() : QByteArray{};
            AccessFormPreview preview(name,bytes,module,db_,this);
            if (!preview.cancelledByOpen()) preview.exec();
        });
    }
    if (kind == "form" && status == "source_only" && !bytes.isEmpty()) {
        const auto formName = query.value(1).toString();
        auto* open = new QPushButton("Access-Formular öffnen (Originalanordnung soweit unterstützt)", &dialog);
        outer->addWidget(open);
        connect(open, &QPushButton::clicked, &dialog, [this, formName, bytes] {
            QSqlQuery vba(db_);
            vba.prepare("SELECT source_bytes FROM _la_access_artifacts WHERE kind='vba' "
                        "AND name=? ORDER BY id DESC LIMIT 1");
            vba.addBindValue("Form_" + formName);
            const auto module = vba.exec() && vba.next() ? vba.value(0).toByteArray() : QByteArray{};
            AccessFormPreview preview(formName, bytes, module, db_, this);
            if (!preview.cancelledByOpen()) preview.exec();
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
                      "Abfragen werden noch nicht automatisch nach SQLite übersetzt.");
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
                void quitApplication() override { QApplication::quit(); }
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

void MainWindow::showTable(const QString& name) {
    auto* page = new QWidget(tabs_);
    auto* outer = new QVBoxLayout(page);
    outer->addWidget(new QLabel("Tabelle: " + name + " — Änderungen erst mit 'Speichern' festschreiben"));
    auto* actions = new QHBoxLayout();
    auto* add = addButton(actions, "+ Datensatz");
    auto* remove = addButton(actions, "Datensatz löschen");
    auto* save = addButton(actions, "Speichern");
    auto* discard = addButton(actions, "Verwerfen");
    outer->addLayout(actions);
    auto* model = new QSqlTableModel(page, db_);
    model->setTable(name);
    model->setEditStrategy(QSqlTableModel::OnManualSubmit);
    model->select();
    auto* table = new QTableView(page);
    table->setModel(model);
    table->setSortingEnabled(true);
    table->horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);
    outer->addWidget(table);
    connect(add, &QPushButton::clicked, page, [model] { model->insertRow(model->rowCount()); });
    connect(remove, &QPushButton::clicked, page, [model, table] {
        if (table->currentIndex().isValid()) model->removeRow(table->currentIndex().row());
    });
    connect(discard, &QPushButton::clicked, page, [model] { model->revertAll(); });
    connect(save, &QPushButton::clicked, page, [this, model] {
        if (!db_.transaction()) { messageError(this, db_.lastError().text()); return; }
        if (!model->submitAll()) {
            const auto problem = model->lastError().text();
            db_.rollback();
            messageError(this, problem);
            return;
        }
        if (!db_.commit()) { messageError(this, db_.lastError().text()); return; }
        model->select();
    });
    tabs_->setCurrentIndex(tabs_->addTab(page, name));
}

void MainWindow::editQuery(const QString& existingName) {
    if (!hasProject()) return;
    QDialog dialog(this);
    dialog.setWindowTitle("SQL-Abfrage");
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
    auto* run = addButton(buttons, "SELECT ausführen");
    auto* save = addButton(buttons, "Speichern");
    outer->addLayout(buttons);
    auto* model = new QSqlQueryModel(&dialog);
    auto* view = new QTableView(&dialog);
    view->setModel(model);
    outer->addWidget(view);
    connect(run, &QPushButton::clicked, &dialog, [this, sql, model] {
        if (!isSelect(sql->toPlainText())) {
            messageError(this, "In diesem Editor sind nur SELECT-Abfragen erlaubt.");
            return;
        }
        model->setQuery(sql->toPlainText(), db_);
        if (model->lastError().isValid()) messageError(this, model->lastError().text());
    });
    connect(save, &QPushButton::clicked, &dialog, [this, name, sql, &dialog] {
        if (!isSelect(sql->toPlainText())) {
            messageError(this, "Nur SELECT-Abfragen speichern.");
            return;
        }
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
                QSqlQuery code(db_);
                code.prepare("SELECT source_bytes FROM _la_access_artifacts "
                             "WHERE kind='vba' AND name=? ORDER BY id DESC LIMIT 1");
                code.addBindValue("Form_" + name);
                const auto module = code.exec() && code.next() ? code.value(0).toByteArray() : QByteArray{};
                AccessFormPreview preview(name, form, module, db_, this);
                if (!preview.cancelledByOpen()) preview.exec();
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
