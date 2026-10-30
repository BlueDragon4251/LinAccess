#include "bridge_importer.hpp"
#include "vba_runtime.hpp"
#include "vba_source_codec.hpp"

#include <QApplication>
#include <QDir>
#include <QFile>
#include <QMessageBox>
#include <QPushButton>
#include <QSqlDatabase>
#include <QSqlError>
#include <QSqlQuery>
#include <QTemporaryDir>
#include <QTimer>

// Tests must stay active in Release builds (which normally define NDEBUG).
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>

namespace {
class TestModuleHost final : public linaccess::VbaHost {
public:
    std::string lastMessage;
    std::optional<std::string> field(std::string_view) override { return std::nullopt; }
    void setField(std::string_view, std::optional<std::string>) override { assert(false); }
    void message(std::string_view text) override { lastMessage = std::string(text); }
    void saveRecord() override { assert(false); }
};
}

namespace {
bool makeCatalog(QSqlDatabase& db) {
    QSqlQuery query(db);
    return query.exec("CREATE TABLE _la_access_artifacts ("
                      "id INTEGER PRIMARY KEY AUTOINCREMENT, "
                      "source_sha256 TEXT NOT NULL, source_filename TEXT NOT NULL, "
                      "kind TEXT NOT NULL, name TEXT NOT NULL, status TEXT NOT NULL, "
                      "format TEXT NOT NULL, detail TEXT NOT NULL, source_bytes BLOB, "
                      "UNIQUE(source_sha256,kind,name,format))");
}

int importWithAutomaticMessageAnswers(BridgeImporter& dialog, int timeoutMs) {
    QPushButton* importButton = nullptr;
    for (auto* button : dialog.findChildren<QPushButton*>()) {
        if (button->text().contains("Alle Artefakte")) importButton = button;
    }
    assert(importButton != nullptr && importButton->isEnabled());
    QTimer answer;
    QObject::connect(&answer, &QTimer::timeout, &dialog, [] {
        for (auto* widget : QApplication::topLevelWidgets()) {
            auto* box = qobject_cast<QMessageBox*>(widget);
            if (!box || !box->isVisible()) continue;
            if (box->standardButtons() & QMessageBox::Yes)
                box->button(QMessageBox::Yes)->click();
            else
                box->accept();
        }
    });
    answer.start(10);
    QTimer::singleShot(0, importButton, &QPushButton::click);
    QTimer::singleShot(timeoutMs, &dialog, &QDialog::reject);
    return dialog.exec();
}

int count(QSqlDatabase& db) {
    QSqlQuery query(db);
    assert(query.exec("SELECT COUNT(*) FROM _la_access_artifacts"));
    assert(query.next());
    return query.value(0).toInt();
}
}

int main(int argc, char** argv) {
    QApplication application(argc, argv);
    QTemporaryDir temp;
    assert(temp.isValid());
    auto db = QSqlDatabase::addDatabase("QSQLITE", "bridge-test");
    db.setDatabaseName(temp.filePath("project.sqlite"));
    assert(db.open());
    assert(makeCatalog(db));

    const QString fixture = QStringLiteral(LINACCESS_FIXTURE_DIRECTORY);
    {
        BridgeImporter dialog(fixture, db);
        assert(importWithAutomaticMessageAnswers(dialog, 1500) == QDialog::Accepted);
        assert(count(db) == 2);
        QSqlQuery query(db);
        assert(query.exec("SELECT source_bytes, status FROM _la_access_artifacts WHERE kind='vba'"));
        assert(query.next() && !query.value(0).toByteArray().isEmpty());
        assert(query.value(1).toString() == "source_only");
        TestModuleHost host;
        const auto source = linaccess::decodeVbaSource(query.value(0).toByteArray());
        const auto result = linaccess::executeVbaProcedure(source.toUtf8().toStdString(),
                                                           "Form_Load", host);
        assert(result.status == linaccess::VbaExecutionStatus::Completed);
        assert(host.lastMessage == "Synthetic bridge fixture");
        const QByteArray legacy = QByteArray::fromHex("4d7367426f7820224be4736522");
        assert(linaccess::decodeVbaSource(legacy) == QString::fromUtf8("MsgBox \"Käse\""));
    }

    // Tampering with a fixture file must reject the entire import transaction.
    const auto modified = temp.filePath("tampered");
    assert(QDir().mkpath(modified + "/objects"));
    for (const auto& filename : {"manifest.json", "objects/000001_vba.bas", "objects/000002_query.sql"}) {
        assert(QFile::copy(QDir(fixture).filePath(filename), QDir(modified).filePath(filename)));
    }
    QFile changed(modified + "/objects/000002_query.sql");
    assert(changed.open(QIODevice::Append));
    assert(changed.write("TAMPERED") > 0);
    changed.close();
    {
        BridgeImporter dialog(modified, db);
        assert(importWithAutomaticMessageAnswers(dialog, 700) == QDialog::Rejected);
        assert(count(db) == 2);
    }
    db.close();
    std::cout << "Bridge GUI integration tests passed\n";
}
