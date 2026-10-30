#include "bridge_importer.hpp"
#include "access_form_import.hpp"
#include "vba_analyzer.hpp"
#include "vba_runtime.hpp"
#include "vba_source_codec.hpp"

#include <QApplication>
#include <QFile>
#include <QJsonDocument>
#include <QMessageBox>
#include <QPushButton>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QTemporaryDir>
#include <QTimer>

#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>

int main(int argc, char** argv) {
    QApplication app(argc, argv);
    assert(argc == 8);
    const QString input = QString::fromLocal8Bit(argv[1]);
    const int expectedTotal = QString::fromLocal8Bit(argv[2]).toInt();
    const int expectedForms = QString::fromLocal8Bit(argv[3]).toInt();
    const int expectedReports = QString::fromLocal8Bit(argv[4]).toInt();
    const int expectedQueries = QString::fromLocal8Bit(argv[5]).toInt();
    const int expectedVba = QString::fromLocal8Bit(argv[6]).toInt();
    const int expectedModules = QString::fromLocal8Bit(argv[7]).toInt();
    QFile manifest(input + "/manifest.json");
    assert(manifest.open(QIODevice::ReadOnly));
    const auto document = QJsonDocument::fromJson(manifest.readAll());
    assert(document.isObject());
    assert(document.object().value("objects").toArray().size() == expectedTotal);
    for (const auto& value : document.object().value("objects").toArray()) {
        assert(value.isObject());
        assert(value.toObject().value("status").toString() == "source_only");
    }

    QTemporaryDir tmp;
    assert(tmp.isValid());
    auto db = QSqlDatabase::addDatabase("QSQLITE", "real-bridge-test");
    db.setDatabaseName(tmp.filePath("source.linaccess"));
    assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE _la_access_artifacts ("
        "id INTEGER PRIMARY KEY AUTOINCREMENT, "
        "source_sha256 TEXT NOT NULL, source_filename TEXT NOT NULL, "
        "kind TEXT NOT NULL, name TEXT NOT NULL, status TEXT NOT NULL, "
        "format TEXT NOT NULL, detail TEXT NOT NULL, source_bytes BLOB, "
        "UNIQUE(source_sha256,kind,name,format))"));
    BridgeImporter importer(input, db);
    QPushButton* importButton = nullptr;
    for (auto* button : importer.findChildren<QPushButton*>()) {
        if (button->text().contains("Alle Artefakte")) importButton = button;
    }
    assert(importButton && importButton->isEnabled());
    QTimer answers;
    QObject::connect(&answers, &QTimer::timeout, &importer, [] {
        for (auto* top : QApplication::topLevelWidgets()) {
            auto* message = qobject_cast<QMessageBox*>(top);
            if (!message || !message->isVisible()) continue;
            if (auto* yes = message->button(QMessageBox::Yes)) yes->click();
            else if (auto* ok = message->button(QMessageBox::Ok)) ok->click();
        }
    });
    answers.start(15);
    QTimer::singleShot(0, importButton, &QPushButton::click);
    QTimer::singleShot(12000, &importer, &QDialog::reject);
    assert(importer.exec() == QDialog::Accepted);

    assert(query.exec("SELECT kind,COUNT(*) FROM _la_access_artifacts GROUP BY kind"));
    int total = 0, forms = 0, reports = 0, queries = 0, vba = 0, modules = 0;
    while (query.next()) {
        const auto kind = query.value(0).toString();
        const auto number = query.value(1).toInt();
        total += number;
        if (kind == "form") forms = number;
        if (kind == "report") reports = number;
        if (kind == "query") queries = number;
        if (kind == "vba") vba = number;
        if (kind == "module") modules = number;
    }
    assert(total == expectedTotal && forms == expectedForms && reports == expectedReports
        && queries == expectedQueries && vba == expectedVba && modules == expectedModules);
    assert(query.exec("SELECT name,source_bytes FROM _la_access_artifacts "
                      "WHERE kind='vba' AND status='source_only'"));
    int procedureCount = 0;
    int verifiedActions = 0;
    struct BridgeHost final : linaccess::VbaHost {
        int maximized = 0;
        int navigated = 0;
        void maximizeForm() override { ++maximized; }
        void gotoLastRecord() override { ++navigated; }
        void message(std::string_view) override {}
        void saveRecord() override {}
        std::optional<std::string> field(std::string_view) override { return {}; }
        void setField(std::string_view, std::optional<std::string>) override {}
    };
    while (query.next()) {
        const auto module = linaccess::decodeVbaSource(query.value(1).toByteArray());
        const auto analysis = linaccess::analyzeVba(module.toUtf8().toStdString());
        procedureCount += static_cast<int>(analysis.procedures.size());
        if (query.value(0).toString() == "Form_3_frm_Rezepte") {
            BridgeHost host;
            const auto executed = linaccess::executeVbaProcedure(
                module.toUtf8().toStdString(), "Form_Open", host);
            assert(executed.status == linaccess::VbaExecutionStatus::Completed);
            assert(!executed.cancelled);
            assert(host.maximized == 1 && host.navigated == 1);
            verifiedActions = 2;
        }
    }
    if (expectedTotal == 26) {
        assert(procedureCount == 8 && verifiedActions == 2);
        assert(query.exec("SELECT source_bytes FROM _la_access_artifacts "
                          "WHERE kind='form' AND name='3_frm_Rezepte'"));
        assert(query.next());
        const auto form = linaccess::parseAccessFormText(
            linaccess::decodeVbaSource(query.value(0).toByteArray()));
        assert(form.recordSource == "3_tbl_Rezepte");
        assert(form.controlSources.contains("Rezeptname"));
        assert(form.eventNames.contains("OnOpen"));
        assert(query.exec("SELECT source_bytes FROM _la_access_artifacts "
                          "WHERE kind='form' AND name='33_Ufo_frm_ZutatenRezept'"));
        assert(query.next());
        const auto subform = linaccess::parseAccessFormText(
            linaccess::decodeVbaSource(query.value(0).toByteArray()));
        assert(subform.recordSource == "33_qry_ZutatenRezept");
    } else if (expectedTotal == 192) {
        assert(procedureCount == 187);
        assert(verifiedActions == 0); // Untrusted complex VBA is never executed in this import test.
        assert(query.exec("SELECT source_bytes FROM _la_access_artifacts "
                          "WHERE kind='form' AND name='frmRegistrierkasse'"));
        assert(query.next());
        const auto registerForm = linaccess::parseAccessFormText(
            linaccess::decodeVbaSource(query.value(0).toByteArray()));
        assert(!registerForm.recordSource.isEmpty() || !registerForm.controlSources.isEmpty());
    }
    std::cout << "Real Access Bridge verified: " << total << " source-only artifacts, "
              << vba << " VBA files, " << procedureCount << " procedures; "
              << verifiedActions << " small, explicitly scoped host actions.\n";
}
