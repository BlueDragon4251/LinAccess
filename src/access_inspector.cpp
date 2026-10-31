#include "access_inspector.hpp"

#include "csv.hpp"
#include "access_table_import.hpp"

#include <QApplication>
#include <QCryptographicHash>
#include <QFile>
#include <QFileInfo>
#include <QHBoxLayout>
#include <QLabel>
#include <QListWidget>
#include <QMessageBox>
#include <QProcess>
#include <QProgressDialog>
#include <QPushButton>
#include <QPlainTextEdit>
#include <QSqlError>
#include <QSqlQuery>
#include <QStandardPaths>
#include <QTabWidget>
#include <QTableWidget>
#include <QTextBrowser>
#include <QVBoxLayout>

#include <algorithm>
#include <stdexcept>

namespace {
QString toQString(const std::string& value) {
    return QString::fromUtf8(value.data(), static_cast<qsizetype>(value.size()));
}
}

AccessInspector::AccessInspector(QString sourceFile, QSqlDatabase projectDb, QWidget* parent)
    : QDialog(parent), sourceFile_(std::move(sourceFile)), projectDb_(std::move(projectDb)) {
    setWindowTitle("Access-Datei untersuchen — read-only Quelle");
    resize(1000, 700);
    auto* outer = new QVBoxLayout(this);
    outer->addWidget(new QLabel("Quelle (wird niemals verändert): " + sourceFile_));
    outer->addWidget(new QLabel(
        "MDBTools liest Tabellen, SQL-Abfragen und den Jet-Objektkatalog. "
        "Formular-/Berichtlayout und eingebettetes VBA können noch nicht extrahiert werden."));
    auto* tabs = new QTabWidget(this);
    outer->addWidget(tabs, 1);
    auto* tableTab = new QWidget(tabs);
    auto* tableLayout = new QVBoxLayout(tableTab);
    auto* row = new QHBoxLayout();
    tables_ = new QListWidget(tableTab);
    tables_->setMinimumWidth(245);
    row->addWidget(tables_);
    preview_ = new QTableWidget(tableTab);
    row->addWidget(preview_, 1);
    tableLayout->addLayout(row, 1);
    auto* buttons = new QHBoxLayout();
    auto* previewButton = new QPushButton("Tabelle ansehen", tableTab);
    importButton_ = new QPushButton("Typisierte Kopie ins Projekt importieren", tableTab);
    auto* importAllButton = new QPushButton("Alle Tabellen in Projektkopie übernehmen", tableTab);
    importButton_->setEnabled(false);
    importAllButton->setEnabled(projectDb_.isOpen());
    buttons->addWidget(previewButton);
    buttons->addWidget(importButton_);
    buttons->addWidget(importAllButton);
    tableLayout->addLayout(buttons);
    tabs->addTab(tableTab, "Tabellen / Daten");

    auto* catalogTab = new QWidget(tabs);
    auto* catalogLayout = new QVBoxLayout(catalogTab);
    auto* objectRow = new QHBoxLayout();
    objects_ = new QListWidget(catalogTab);
    objects_->setMinimumWidth(330);
    objectRow->addWidget(objects_);
    objectPreview_ = new QPlainTextEdit(catalogTab);
    objectPreview_->setReadOnly(true);
    objectRow->addWidget(objectPreview_, 1);
    catalogLayout->addLayout(objectRow, 1);
    catalogImportButton_ = new QPushButton("Objekt-Metadaten ins Projekt übernehmen", catalogTab);
    catalogImportButton_->setEnabled(projectDb_.isOpen());
    catalogLayout->addWidget(catalogImportButton_);
    tabs->addTab(catalogTab, "Abfragen / Formulare / Berichte");

    status_ = new QTextBrowser(this);
    status_->setMaximumHeight(150);
    status_->setHtml(
        "<b>Kompatibilität:</b> Datentypen und NULLs werden beim Tabellenimport nach Möglichkeit erhalten. "
        "Primärschlüssel und exportierte Indizes werden übernommen; der Gesamtimport erhält auch Beziehungen. Access-Standardwerte und OLE-Nutzdaten fehlen. "
        "Der Katalog zeigt Objektnamen und SQL, <b>nicht</b> VBA-Code oder GUI-Layouts. "
        "Importierte Katalogeinträge sind <b>metadata_only</b> und nicht ausführbar. "
        "Die Access-Originaldatei bleibt unverändert.");
    outer->addWidget(status_);
    connect(previewButton, &QPushButton::clicked, this, &AccessInspector::previewTable);
    connect(tables_, &QListWidget::itemDoubleClicked, this, [this] { previewTable(); });
    connect(importButton_, &QPushButton::clicked, this, &AccessInspector::importTable);
    connect(importAllButton, &QPushButton::clicked, this, &AccessInspector::importAllTables);
    connect(objects_, &QListWidget::currentRowChanged, this, [this] { previewObject(); });
    connect(catalogImportButton_, &QPushButton::clicked, this, &AccessInspector::importCatalog);
    loadTables();
    loadCatalog();
}

QByteArray AccessInspector::runMdbTool(const QString& name, const QStringList& arguments) {
    const auto tool = QStandardPaths::findExecutable(name);
    if (tool.isEmpty())
        throw std::runtime_error((name + " fehlt. Unter Debian/Ubuntu: sudo apt install mdbtools").toStdString());
    QProcess proc;
    proc.setProgram(tool);
    proc.setArguments(arguments); // No shell: table names and paths are never interpreted as commands.
    proc.start();
    if (!proc.waitForStarted(5000))
        throw std::runtime_error("MDBTools-Prozess konnte nicht gestartet werden");
    if (!proc.waitForFinished(30000)) {
        proc.kill();
        proc.waitForFinished(2000);
        throw std::runtime_error("MDBTools-Zeitlimit (30 Sekunden) überschritten");
    }
    if (proc.exitStatus() != QProcess::NormalExit || proc.exitCode() != 0)
        throw std::runtime_error("MDBTools-Fehler: " + proc.readAllStandardError().toStdString());
    auto output = proc.readAllStandardOutput();
    if (output.size() > 20 * 1024 * 1024)
        throw std::runtime_error("Export überschreitet das 20-MB-Prototyp-Limit");
    return output;
}

void AccessInspector::loadTables() {
    try {
        const auto output = runMdbTool("mdb-tables", {"-1", sourceFile_});
        for (auto table : QString::fromUtf8(output).split('\n', Qt::SkipEmptyParts)) {
            table = table.trimmed();
            if (!table.isEmpty() && !table.startsWith("MSys", Qt::CaseInsensitive))
                tables_->addItem(table);
        }
        if (tables_->count() == 0)
            status_->append("Keine lesbaren Benutzertabellen gefunden. ACCDB-Unterstützung variiert je nach MDBTools-Version.");
    } catch (const std::exception& exc) {
        status_->append("<b>Importpfad nicht verfügbar:</b> " + QString::fromUtf8(exc.what()).toHtmlEscaped());
    }
}

void AccessInspector::previewTable() {
    if (!tables_->currentItem()) return;
    selectedTable_ = tables_->currentItem()->text();
    try {
        // The control-character marker is different from an empty text field;
        // quoted text identical to the marker is also not considered NULL.
        const QByteArray typedOutput = linaccess::AccessCatalog::exportTable(sourceFile_,selectedTable_);
        csvRows_ = linaccess::parseCsvDetailed(std::string_view(typedOutput.constData(),
                                                               static_cast<std::size_t>(typedOutput.size())));
        if (csvRows_.empty()) throw std::runtime_error("Leerer Export ohne Kopfzeile");
        const auto& header = csvRows_[0];
        if (header.empty()) throw std::runtime_error("Kopfzeile fehlt");
        preview_->setColumnCount(static_cast<int>(header.size()));
        QStringList headers;
        for (const auto& field : header) headers << toQString(field.value);
        preview_->setHorizontalHeaderLabels(headers);
        const auto shown = std::min<std::size_t>(200, csvRows_.size() - 1);
        preview_->setRowCount(static_cast<int>(shown));
        for (std::size_t row = 0; row < shown; ++row) {
            const auto& cells = csvRows_[row + 1];
            for (std::size_t col = 0; col < std::min(header.size(), cells.size()); ++col)
                preview_->setItem(static_cast<int>(row), static_cast<int>(col),
                                  new QTableWidgetItem(!cells[col].quoted && cells[col].value == "\x1e"
                                      ? QStringLiteral("<NULL>") : toQString(cells[col].value)));
        }
        importButton_->setEnabled(projectDb_.isOpen());
        status_->append("Vorschau: " + selectedTable_.toHtmlEscaped() + " — "
                        + QString::number(csvRows_.size() - 1) + " Datenzeilen im CSV-Export; bis zu 200 sichtbar.");
    } catch (const std::exception& exc) {
        csvRows_.clear();
        importButton_->setEnabled(false);
        QMessageBox::critical(this, "Access-Export fehlgeschlagen", QString::fromUtf8(exc.what()));
    }
}

void AccessInspector::importTable() {
    if (!projectDb_.isOpen() || csvRows_.empty() || selectedTable_.isEmpty()) return;
    if (QMessageBox::warning(this, "Verlustbehafteter Import",
            "Die Tabelle wird als neue SQLite-Tabelle mit INTEGER/REAL/TEXT und echten NULL-Werten importiert. "
            "Beziehungen, Access-Standardwerte, Formulare und VBA "
            "werden NICHT übernommen. "
            "Die Access-Quelldatei bleibt erhalten. Fortfahren?",
            QMessageBox::Yes | QMessageBox::No, QMessageBox::No) != QMessageBox::Yes) return;
    try {
        linaccess::AccessCatalog catalog(sourceFile_);
        const auto columns = linaccess::parseMdbSchema(catalog.tableSchema(selectedTable_));
        linaccess::importAccessTable(projectDb_, sourceFile_, selectedTable_, columns, csvRows_);
        importButton_->setEnabled(false);
        QMessageBox::information(this, "Importiert", "Tabelle mit Datentypen und NULLs kopiert. "
            "Primärschlüssel und Indizes erhalten. Beziehungen nur beim Gesamtimport. Original unverändert.");
    } catch (const std::exception& exc) {
        QMessageBox::critical(this, "Import abgebrochen", QString::fromUtf8(exc.what()));
    }
}

void AccessInspector::importAllTables() {
    if (!projectDb_.isOpen() || tables_->count() == 0) return;
    if (QMessageBox::question(this, "Tabellenkopie",
        "Alle Tabellen der MDB werden in das geöffnete LinAccess-Projekt importiert. "
        "Der gesamte Import ist atomar. Gleichnamige vorhandene Tabellen führen zum Abbruch ohne Änderungen. "
        "Fortfahren?", QMessageBox::Yes | QMessageBox::No, QMessageBox::Yes) != QMessageBox::Yes) return;
    QApplication::setOverrideCursor(Qt::WaitCursor);
    try {
        const int imported = linaccess::importAccessDatabase(projectDb_, sourceFile_);
        QApplication::restoreOverrideCursor();
        status_->append(QString::number(imported) + " Tabellen mit Schlüsseln, Indizes und Beziehungen übernommen.");
        QMessageBox::information(this, "Tabellenimport abgeschlossen",
            QString::number(imported) + " Tabellen erfolgreich übernommen. Original unverändert.");
    } catch (const std::exception& error) {
        QApplication::restoreOverrideCursor();
        QMessageBox::warning(this, "Import zurückgerollt", QString::fromUtf8(error.what()));
    }
}

void AccessInspector::loadCatalog() {
    try {
        linaccess::AccessCatalog inspect(sourceFile_);
        catalog_ = inspect.objects();
        int visible = 0;
        for (int i = 0; i < catalog_.size(); ++i) {
            const auto& object = catalog_[i];
            if (object.internal) continue;
            auto* item = new QListWidgetItem(object.kind + ": " + object.name, objects_);
            item->setData(Qt::UserRole, i);
            ++visible;
        }
        status_->append("MDB-Version: " + inspect.formatVersion().toHtmlEscaped()
                        + " | " + QString::number(visible)
                        + " sichtbare Objekte (inkl. Tabellen und Beziehungen). "
                        "Versteckte Access-Systemabfragen sind ausgeblendet.");
    } catch (const std::exception& exc) {
        status_->append("Katalog nicht lesbar: " + QString::fromUtf8(exc.what()).toHtmlEscaped());
    }
}

void AccessInspector::previewObject() {
    const auto* selected = objects_->currentItem();
    if (!selected) return;
    const int index = selected->data(Qt::UserRole).toInt();
    if (index < 0 || index >= catalog_.size()) return;
    const auto& object = catalog_[index];
    QString details = "Typ: " + object.kind + "\nName: " + object.name + "\n\n";
    try {
        linaccess::AccessCatalog inspect(sourceFile_);
        if (object.kind == "query") {
            details += "Access/Jet-SQL (nur Quelle; nicht direkt auf SQLite ausführbar):\n\n"
                       + inspect.querySql(object.name);
            const auto missing = inspect.missingQuerySources(inspect.querySql(object.name));
            if (!missing.isEmpty())
                details += "\n\nACHTUNG: Möglicherweise fehlende verknüpfte Tabellen/Abfragen "
                           "(Heuristik, nicht vollständig):\n - " + missing.join("\n - ");
        } else if (object.kind == "table") {
            details += "MDBTools-Tabellenschema:\n\n" + inspect.tableSchema(object.name);
        } else if (object.kind == "relationship") {
            for (const auto& relation : inspect.relationships())
                details += relation.childTable + '.' + relation.childColumn + " → "
                           + relation.parentTable + '.' + relation.parentColumn + '\n';
        } else {
            details += "Nur im Access-Jet-Katalog nachgewiesen.\n"
                       "Layout, eingebettete Ereignisprozeduren und VBA-Bytecode sind "
                       "NICHT extrahiert und NICHT ausführbar.\n"
                       "Für den Originalquelltext ist ein separater Export mit Access auf Windows nötig.";
        }
    } catch (const std::exception& exc) {
        details += "MDBTools-Fehler: " + QString::fromUtf8(exc.what());
    }
    objectPreview_->setPlainText(details);
}

void AccessInspector::importCatalog() {
    if (!projectDb_.isOpen()) return;
    if (QMessageBox::question(this, "Nur Metadaten",
          "Es werden ausschließlich die sichtbaren Objektnamen, Abfrage-SQL und Tabellenschemata "
          "in den Quellkatalog kopiert. Keine Formulare, Berichte oder VBA-Prozeduren werden "
          "dadurch ausführbar. Fortfahren?", QMessageBox::Yes | QMessageBox::No,
          QMessageBox::No) != QMessageBox::Yes) return;
    QFile file(sourceFile_);
    if (!file.open(QIODevice::ReadOnly)) {
        QMessageBox::critical(this, "Katalog", "MDB-Quelle kann nicht gelesen werden");
        return;
    }
    QCryptographicHash sourceHash(QCryptographicHash::Sha256);
    if (!sourceHash.addData(&file)) {
        QMessageBox::critical(this, "Katalog", "MDB-Datei konnte nicht vollständig gehasht werden");
        return;
    }
    const auto digest = QString::fromLatin1(sourceHash.result().toHex());
    file.close();
    if (!projectDb_.transaction()) {
        QMessageBox::critical(this, "Transaktion", projectDb_.lastError().text());
        return;
    }
    QString problem;
    try {
        linaccess::AccessCatalog inspect(sourceFile_);
        QSqlQuery insert(projectDb_);
        if (!insert.prepare("INSERT INTO _la_access_artifacts "
                            "(source_sha256,source_filename,kind,name,status,format,detail,source_bytes) "
                            "VALUES (?,?,?,?,?,?,?,?) "
                            "ON CONFLICT(source_sha256,kind,name,format) DO NOTHING"))
            throw std::runtime_error(insert.lastError().text().toStdString());
        for (const auto& item : catalog_) {
            if (item.internal) continue;
            QString source;
            if (item.kind == "query") source = inspect.querySql(item.name);
            if (item.kind == "table") source = inspect.tableSchema(item.name);
            if (item.kind == "relationship") {
                for (const auto& r : inspect.relationships())
                    source += r.childTable + '.' + r.childColumn + " -> "
                              + r.parentTable + '.' + r.parentColumn + '\n';
            }
            insert.bindValue(0, digest);
            insert.bindValue(1, QFileInfo(sourceFile_).fileName());
            insert.bindValue(2, item.kind);
            insert.bindValue(3, item.name);
            insert.bindValue(4, "metadata_only");
            insert.bindValue(5, "mdbtools_catalog_v1");
            insert.bindValue(6, "Objekt katalogisiert; SQL/Schema ggf. vorhanden. "
                                "Access-Formulare, Berichte und VBA nicht lauffähig.");
            insert.bindValue(7, source.toUtf8());
            if (!insert.exec()) throw std::runtime_error(insert.lastError().text().toStdString());
        }
        if (!projectDb_.commit())
            throw std::runtime_error(projectDb_.lastError().text().toStdString());
    } catch (const std::exception& exc) {
        problem = QString::fromUtf8(exc.what());
        projectDb_.rollback();
    }
    if (!problem.isEmpty()) {
        QMessageBox::critical(this, "Katalogimport abgebrochen", problem);
        return;
    }
    QMessageBox::information(this, "Katalog übernommen",
        "Objektnamen und SQL/Schema sind als nicht ausführbare Metadaten gespeichert.");
}
