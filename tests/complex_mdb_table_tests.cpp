#include "access_catalog.hpp"
#include "dao_database.hpp"
#include "access_table_import.hpp"
#include "csv.hpp"

#include <QCoreApplication>
#include <QCryptographicHash>
#include <QFile>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlError>
#include <QTemporaryDir>

#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>
#include <QJsonArray>
#include <QJsonDocument>
#include <set>

int main(int argc, char** argv) {
    QCoreApplication app(argc, argv);
    assert(argc == 2);
    const QString source = QString::fromLocal8Bit(argv[1]);
    QFile original(source);
    assert(original.open(QIODevice::ReadOnly));
    const auto before = QCryptographicHash::hash(original.readAll(), QCryptographicHash::Sha256);
    original.close();
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
    int booleanColumns = 0;
    int trueBooleans = 0;
    int typedBooleans = 0;
    QStringList failures;
    for (const auto& table : tables) {
        if (table.startsWith("MSys", Qt::CaseInsensitive)) continue;
        try {
            const auto csv = linaccess::AccessCatalog::exportTable(source,table);
            const auto rows = linaccess::parseCsvDetailed(std::string_view(csv.constData(), csv.size()));
            const auto schema = linaccess::parseMdbSchema(catalog.tableSchema(table));
            for (const auto& column : schema.columns)
                if (column.sourceType == "BOOLEAN" || column.sourceType == "BIT") ++booleanColumns;
            linaccess::importAccessTable(db, source, table, schema, rows);
            bool hasBoolean = false;
            for (const auto& column : schema.columns) hasBoolean |= column.sourceType == "BOOLEAN" || column.sourceType == "BIT";
            if (hasBoolean && rows.size() > 1) {
                auto database = linaccess::makeDaoDatabase(db);
                auto recordset = database->member("OpenRecordset",{linaccess::VbaValue::string(table.toUtf8().toStdString()),linaccess::VbaValue::numeric(4)});
                for (const auto& column : schema.columns) {
                    if (column.sourceType != "BOOLEAN" && column.sourceType != "BIT") continue;
                    const auto value = recordset.object->member("Fields",{linaccess::VbaValue::string(column.name.toUtf8().toStdString())}).object->member("Value",{});
                    assert(value.kind == linaccess::VbaValue::Kind::Boolean); ++typedBooleans;
                }
                recordset.object->member("Close",{});
            }
            auto quote = [](QString name) { name.replace('"', "\"\""); return '"' + name + '"'; };
            QSqlQuery copied(db);
            assert(copied.exec("SELECT * FROM " + quote(table)));
            std::multiset<QByteArray> expectedRows, actualRows;
            for (std::size_t row = 1; row < rows.size(); ++row) {
                QJsonArray values;
                for (int col = 0; col < schema.columns.size(); ++col) {
                    const auto& cell = rows[row][col];
                    const auto text = QString::fromUtf8(cell.value);
                    if (!cell.quoted && cell.value == "\x1e") values.append(QJsonValue::Null);
                    else if (schema.columns[col].sourceType == "BOOLEAN" || schema.columns[col].sourceType == "BIT") {
                        const bool truth = text == "1" || text == "-1" || text.compare("TRUE",Qt::CaseInsensitive) == 0;
                        values.append(truth ? "-1" : "0"); if (truth) ++trueBooleans;
                    }
                    else if (schema.columns[col].sqliteType == "INTEGER") values.append(QString::number(text.toLongLong()));
                    else if (schema.columns[col].sqliteType == "REAL") values.append(QString::number(text.toDouble(), 'g', 17));
                    else values.append(text);
                }
                expectedRows.insert(QJsonDocument(values).toJson(QJsonDocument::Compact));
            }
            while (copied.next()) {
                QJsonArray values;
                for (int col = 0; col < schema.columns.size(); ++col) {
                    const auto cell = copied.value(col);
                    if (cell.isNull()) values.append(QJsonValue::Null);
                    else if (schema.columns[col].sqliteType == "INTEGER") values.append(QString::number(cell.toLongLong()));
                    else if (schema.columns[col].sqliteType == "REAL") values.append(QString::number(cell.toDouble(), 'g', 17));
                    else values.append(cell.toString());
                }
                actualRows.insert(QJsonDocument(values).toJson(QJsonDocument::Compact));
            }
            assert(expectedRows == actualRows);
            assert(copied.exec("PRAGMA table_info(" + quote(table) + ")"));
            int col = 0;
            while (copied.next()) {
                assert(copied.value(5).toInt() == schema.columns[col].primaryKeyOrder);
                ++col;
            }
            assert(col == schema.columns.size());
            assert(copied.exec(QString("SELECT COUNT(*) FROM sqlite_master WHERE type='index' AND tbl_name=")
                              + "'" + QString(table).replace("'", "''") + "' AND sql IS NOT NULL"));
            assert(copied.next() && copied.value(0).toInt() == schema.indexes.size());
            ++imported;
        } catch (const std::exception& e) {
            failures << QString::number(imported + failures.size() + 1);
            (void)e; // Keep source names, values and tool stderr out of CI logs.
        }
    }
    std::cout << "Kassa real table copy: " << imported << "/" << tables.size() << " tables\n";
    for (const auto& issue : failures) std::cout << "  ERROR " << issue.toStdString() << '\n';
    // This test deliberately does not pretend 100% source data compatibility;
    // it makes regressions against an actual large Access 2003 file visible.
    assert(booleanColumns == 16 && trueBooleans > 0 && typedBooleans > 0);
    assert(tables.size() == 29);
    assert(imported == 29 && failures.empty());
    assert(original.open(QIODevice::ReadOnly));
    assert(QCryptographicHash::hash(original.readAll(), QCryptographicHash::Sha256) == before);
    auto full = QSqlDatabase::addDatabase("QSQLITE", "full-mdb-test");
    full.setDatabaseName(temp.filePath("full.linaccess"));
    assert(full.open());
    QSqlQuery fullQuery(full);
    assert(fullQuery.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    assert(linaccess::importAccessDatabase(full, source) == 29);
    assert(fullQuery.exec("PRAGMA foreign_key_check"));
    assert(!fullQuery.next());
    assert(fullQuery.exec("SELECT COUNT(*) FROM _la_access_relations"));
    assert(fullQuery.next() && fullQuery.value(0).toInt() == catalog.relationships().size());
    return 0;
}
