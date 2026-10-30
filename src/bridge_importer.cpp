#include "bridge_importer.hpp"
#include "vba_analyzer.hpp"
#include "vba_source_codec.hpp"

#include <QDir>
#include <QCryptographicHash>
#include <QFile>
#include <QFileInfo>
#include <QHeaderView>
#include <QHBoxLayout>
#include <QJsonDocument>
#include <QJsonParseError>
#include <QLabel>
#include <QMap>
#include <QMessageBox>
#include <QPushButton>
#include <QRegularExpression>
#include <QSqlError>
#include <QSqlQuery>
#include <QTableWidget>
#include <QTextBrowser>
#include <QVBoxLayout>

#include <stdexcept>

#include <utility>
#include <vector>

BridgeImporter::BridgeImporter(QString directory, QSqlDatabase project, QWidget* parent)
    : QDialog(parent), directory_(std::move(directory)), project_(std::move(project)) {
    setWindowTitle("Access Bridge — Objektquellen prüfen und übernehmen");
    resize(1050, 760);
    auto* outer = new QVBoxLayout(this);
    status_ = new QLabel("Bridge-Paket wird geprüft…", this);
    outer->addWidget(status_);
    outer->addWidget(new QLabel(
        "Formulare, Berichte und VBA bleiben als Originalartefakte erhalten. "
        "Beim Import wird kein VBA ausgeführt; später ist nur die explizit "
        "eingeschränkte VBA-Teillaufzeit verfügbar.", this));
    catalog_ = new QTableWidget(this);
    catalog_->setColumnCount(4);
    catalog_->setHorizontalHeaderLabels({"Typ", "Name", "Status", "Fehler / Details"});
    catalog_->horizontalHeader()->setSectionResizeMode(QHeaderView::ResizeToContents);
    catalog_->horizontalHeader()->setStretchLastSection(true);
    catalog_->setSelectionBehavior(QAbstractItemView::SelectRows);
    catalog_->setEditTriggers(QAbstractItemView::NoEditTriggers);
    outer->addWidget(catalog_, 2);
    preview_ = new QTextBrowser(this);
    outer->addWidget(preview_, 1);
    auto* buttons = new QHBoxLayout();
    auto* importButton = new QPushButton("Alle Artefakte ins SQLite-Projekt übernehmen", this);
    auto* closeButton = new QPushButton("Schließen", this);
    buttons->addWidget(importButton);
    buttons->addWidget(closeButton);
    outer->addLayout(buttons);
    connect(catalog_, &QTableWidget::itemSelectionChanged, this, &BridgeImporter::previewSelected);
    connect(closeButton, &QPushButton::clicked, this, &QDialog::accept);
    connect(importButton, &QPushButton::clicked, this, &BridgeImporter::importAll);
    importButton->setEnabled(load() && project_.isOpen());
}

bool BridgeImporter::load() {
    QFile file(QDir(directory_).filePath("manifest.json"));
    if (!file.open(QIODevice::ReadOnly)) {
        status_->setText("manifest.json fehlt oder ist nicht lesbar.");
        return false;
    }
    if (file.size() > 5 * 1024 * 1024) {
        status_->setText("Manifest überschreitet 5 MB.");
        return false;
    }
    QJsonParseError problem;
    const auto doc = QJsonDocument::fromJson(file.readAll(), &problem);
    if (problem.error != QJsonParseError::NoError || !doc.isObject()) {
        status_->setText("Ungültiges JSON-Manifest: " + problem.errorString());
        return false;
    }
    manifest_ = doc.object();
    if (manifest_.value("format").toString() != "linaccess-bridge-v1" ||
        !manifest_.value("objects").isArray()) {
        status_->setText("Nicht unterstützte Bridge-Version oder Objektliste fehlt.");
        return false;
    }
    objects_ = manifest_.value("objects").toArray();
    if (objects_.size() > 10000) {
        status_->setText("Mehr als 10.000 Objekte werden in diesem Prototyp nicht angenommen.");
        return false;
    }
    const auto sha = manifest_.value("source").toObject().value("sha256").toString();
    const QRegularExpression shaPattern("^[a-fA-F0-9]{64}$");
    if (!shaPattern.match(sha).hasMatch()) {
        status_->setText("Quellfingerabdruck fehlt oder ist ungültig.");
        return false;
    }
    catalog_->setRowCount(objects_.size());
    for (int i = 0; i < objects_.size(); ++i) {
        if (!objects_.at(i).isObject()) {
            status_->setText("Ungültiger Objektdatensatz.");
            return false;
        }
        const auto object = objects_.at(i).toObject();
        const auto strings = QStringList{
            object.value("kind").toString(), object.value("name").toString(),
            object.value("status").toString(), object.value("detail").toString()};
        for (int col = 0; col < strings.size(); ++col)
            catalog_->setItem(i, col, new QTableWidgetItem(strings[col]));
    }
    status_->setText("Quelle: " + manifest_.value("source").toObject().value("filename").toString()
                     + " — " + QString::number(objects_.size()) + " Objekteinträge. "
                     "Nur Text-/Binärquellen, keine Ausführung.");
    if (objects_.size()) catalog_->selectRow(0);
    return true;
}

QString BridgeImporter::artifactPath(const QJsonObject& object) const {
    const QString relative = object.value("file").toString();
    if (relative.isEmpty()) return {};
    if (QDir::isAbsolutePath(relative)) throw std::runtime_error("Absolute artifact path rejected");
    const QString root = QFileInfo(directory_).canonicalFilePath();
    const QString candidate = QFileInfo(QDir(directory_).filePath(relative)).canonicalFilePath();
    if (root.isEmpty() || candidate.isEmpty() ||
        !candidate.startsWith(root + QDir::separator()) ||
        !QFileInfo(candidate).isFile()) {
        throw std::runtime_error("Artifact path is outside bridge package or does not exist");
    }
    if (QFileInfo(candidate).size() > 16 * 1024 * 1024)
        throw std::runtime_error("Artifact exceeds 16-MB limit");
    return candidate;
}

void BridgeImporter::previewSelected() {
    const int row = catalog_->currentRow();
    if (row < 0 || row >= objects_.size()) return;
    const auto object = objects_.at(row).toObject();
    try {
        const auto path = artifactPath(object);
        if (path.isEmpty()) {
            preview_->setPlainText("Keine Quelldatei. " + object.value("detail").toString());
            return;
        }
        QFile file(path);
        if (!file.open(QIODevice::ReadOnly)) throw std::runtime_error("Cannot read artifact");
        const auto bytes = file.read(128 * 1024);
        preview_->setPlainText(QString::fromUtf8(bytes)
            + (file.atEnd() ? "" : "\n\n[Vorschau nach 128 KB abgeschnitten]"));
    } catch (const std::exception& exc) {
        preview_->setPlainText("Vorschaufehler: " + QString::fromUtf8(exc.what()));
    }
}

void BridgeImporter::importAll() {
    if (!project_.isOpen()) return;
    if (QMessageBox::question(this, "Bridge-Import",
            "Originaldefinitionen und VBA-Quellartefakte werden ins SQLite-Projekt kopiert. "
            "Sie bleiben zunächst NICHT ausführbar. Fortfahren?",
            QMessageBox::Yes | QMessageBox::No, QMessageBox::No) != QMessageBox::Yes) return;
    const auto source = manifest_.value("source").toObject();
    const auto sourceHash = source.value("sha256").toString();
    const auto sourceName = source.value("filename").toString();
    struct Record {
        QString kind, name, status, format, detail;
        QByteArray bytes;
    };
    std::vector<Record> ready;
    qint64 total = 0;
    try {
        for (const auto& value : objects_) {
            const auto item = value.toObject();
            Record record{item.value("kind").toString(), item.value("name").toString(),
                          item.value("status").toString(), item.value("format").toString(),
                          item.value("detail").toString(), {}};
            if (record.kind.isEmpty() || record.kind.size() > 64 || record.name.size() > 1024 ||
                record.status.isEmpty() || record.format.size() > 64)
                throw std::runtime_error("Invalid artifact metadata");
            const auto path = artifactPath(item); // canonical path + size checks
            if (!path.isEmpty()) {
                QFile file(path);
                if (!file.open(QIODevice::ReadOnly)) throw std::runtime_error("Cannot open artifact");
                record.bytes = file.readAll();
                const auto actualHash = QString::fromLatin1(
                    QCryptographicHash::hash(record.bytes, QCryptographicHash::Sha256).toHex());
                const auto expectedHash = item.value("sha256").toString().toLower();
                if (expectedHash.size() != 64 || expectedHash != actualHash)
                    throw std::runtime_error("Artifact SHA-256 mismatch or missing digest");
                total += record.bytes.size();
                if (total > 100 * 1024 * 1024)
                    throw std::runtime_error("Bridge package exceeds 100-MB import limit");
            } else if (record.status == "source_only") {
                throw std::runtime_error("Source-only artifact has no source file");
            }
            ready.push_back(std::move(record));
        }
    } catch (const std::exception& exc) {
        QMessageBox::critical(this, "Ungültiges Bridge-Paket", QString::fromUtf8(exc.what()));
        return;
    }

    if (!project_.transaction()) {
        QMessageBox::critical(this, "Transaktion", project_.lastError().text());
        return;
    }
    QSqlQuery query(project_);
    const bool prepared = query.prepare(
        "INSERT INTO _la_access_artifacts "
        "(source_sha256, source_filename, kind, name, status, format, detail, source_bytes) "
        "VALUES (?,?,?,?,?,?,?,?)");
    bool ok = prepared;
    QString problem = query.lastError().text();
    for (const auto& record : ready) {
        if (!ok) break;
        query.bindValue(0, sourceHash);
        query.bindValue(1, sourceName);
        query.bindValue(2, record.kind);
        query.bindValue(3, record.name);
        query.bindValue(4, record.status);
        query.bindValue(5, record.format);
        query.bindValue(6, record.detail);
        query.bindValue(7, record.bytes);
        ok = query.exec();
        if (!ok) problem = query.lastError().text();
    }
    if (ok) ok = project_.commit();
    if (!ok) {
        if (problem.isEmpty()) problem = project_.lastError().text();
        project_.rollback();
        QMessageBox::critical(this, "Bridge-Import abgebrochen", problem);
        return;
    }
    int detectedProcedures = 0;
    int vbaFiles = 0;
    QMap<QString, int> dependencies;
    for (const auto& record : ready) {
        if (record.kind != "vba" || record.status != "source_only") continue;
        ++vbaFiles;
        const auto source = linaccess::decodeVbaSource(record.bytes).toUtf8().toStdString();
        const auto report = linaccess::analyzeVba(source);
        detectedProcedures += static_cast<int>(report.procedures.size());
        for (const auto& item : report.compatibilityFlags)
            dependencies[QString::fromStdString(item)]++;
    }
    QString summary = QString::number(ready.size()) + " geprüfte Quellobjekte gespeichert. "
        "Beim Import wurde KEIN VBA ausgeführt.";
    if (vbaFiles) {
        summary += "\n\nStatische VBA-Analyse: " + QString::number(vbaFiles)
            + " Module, " + QString::number(detectedProcedures) + " Prozeduren erkannt.";
        if (!dependencies.empty()) {
            summary += "\nAccess-/Windows-Abhängigkeiten (Anzahl Module):";
            for (auto it = dependencies.cbegin(); it != dependencies.cend(); ++it)
                summary += "\n• " + it.key() + ": " + QString::number(it.value());
        }
        summary += "\n\nDas sind Warnhinweise, keine Ausführungsgarantien. "
                   "Insbesondere DAO, Access-Formobjekte und fremde COM-Objekte "
                   "sind noch nicht mit der Linux-VBA-Teillaufzeit kompatibel.";
    }
    QMessageBox::information(this, "Bridge-Import abgeschlossen", summary);
    accept();
}
