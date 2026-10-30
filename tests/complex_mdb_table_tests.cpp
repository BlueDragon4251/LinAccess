#include "access_catalog.hpp"
#include "access_table_import.hpp"
#include "csv.hpp"
#include "vba_source_codec.hpp"

#include <QCoreApplication>
#include <QDir>
#include <QFile>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlError>
#include <QTemporaryDir>

#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>

int main(int argc, char** argv) {
    QCoreApplication app(argc, argv);
    assert(argc == 2 || argc == 3);
    const QString source = QString::fromLocal8Bit(argv[1]);
    QTemporaryDir temp;
    assert(temp.isValid());
    auto db = QSqlDatabase::addDatabase("QSQLITE", "complex-mdb-table-test");
    db.setDatabaseName(temp.filePath("copy.linaccess"));
    assert(db.open());
    QSqlQuery init(db);
    assert(init.exec("CREATE TABLE _la_import_log ("
        "id INTEGER PRIMARY KEY AUTOINCREMENT, source_file TEXT, kind TEXT, "
        "name TEXT, status TEXT, details TEXT)"));
    const linaccess::AccessCatalog catalog(source);
    const auto tables = QString::fromUtf8(linaccess::AccessCatalog::tool(
        "mdb-tables", {"-1", source})).split('\n', Qt::SkipEmptyParts);
    int imported = 0;
    QStringList failures;
    for (const auto& table : tables) {
        if (table.startsWith("MSys", Qt::CaseInsensitive)) continue;
        try {
            const auto csv = linaccess::AccessCatalog::tool(
                "mdb-export", {"-0", QString(QChar(0x1e)), source, table});
            const auto rows = linaccess::parseCsvDetailed(std::string_view(csv.constData(), csv.size()));
            linaccess::importAccessTable(db, source, table,
                linaccess::parseMdbColumns(catalog.tableSchema(table)), rows);
            ++imported;
        } catch (const std::exception& e) {
            failures << table + ": " + QString::fromUtf8(e.what());
        }
    }
    std::cout << "Kassa real table copy: " << imported << "/" << tables.size() << " tables\n";
    for (const auto& issue : failures) std::cout << "  ERROR " << issue.toStdString() << '\n';
    // This test deliberately does not pretend 100% source data compatibility;
    // it makes regressions against an actual large Access 2003 file visible.
    assert(imported >= 20);
    if (argc == 3) {
        const QDir bridge(QString::fromLocal8Bit(argv[2]));
        QFile file(bridge.filePath("manifest.json"));
        assert(file.open(QIODevice::ReadOnly));
        const auto objects = QJsonDocument::fromJson(file.readAll()).object().value("objects").toArray();
        int selectQueries = 0, directlyRunnable = 0;
        for (const auto& objectValue : objects) {
            const auto item = objectValue.toObject();
            if (item.value("kind") != "query") continue;
            QFile sqlSource(bridge.filePath(item.value("file").toString()));
            assert(sqlSource.open(QIODevice::ReadOnly));
            auto sql = linaccess::decodeVbaSource(sqlSource.readAll()).trimmed();
            if (!sql.startsWith("SELECT", Qt::CaseInsensitive)) continue;
            ++selectQueries;
            QSqlQuery q(db);
            if (q.exec(sql)) ++directlyRunnable;
            else if (selectQueries <= 25) {
                std::cout << "  QUERY " << item.value("name").toString().toStdString()
                          << ": " << q.lastError().text().toStdString() << '\n';
            }
        }
        std::cout << "Kassa saved SELECTs directly runnable with SQLite: "
                  << directlyRunnable << "/" << selectQueries << '\n';
    }
    return 0;
}
