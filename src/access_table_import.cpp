#include "access_table_import.hpp"
#include "access_catalog.hpp"
#include <QMap>
#include <QFile>
#include <QTemporaryDir>
#include <QCryptographicHash>
#include <QJsonArray>
#include <QJsonObject>
#include <QJsonDocument>
#include <algorithm>

#include <QRegularExpression>
#include <QSqlError>
#include <QSqlQuery>

#include <set>
#include <stdexcept>

namespace linaccess {
namespace {
QString sqlQuote(QString name) {
    name.replace('"', "\"\"");
    return '"' + name + '"';
}
std::runtime_error fail(const QString& error) {
    return std::runtime_error(error.toUtf8().toStdString());
}
}

AccessTableSchema parseMdbSchema(const QString& schema) {
    const QRegularExpression column(R"(^\s*`((?:``|[^`])+)`\s+([A-Za-z][A-Za-z0-9_]*)(.*)$)");
    const QRegularExpression identifier(R"(`((?:``|[^`])+)`)");
    const QRegularExpression primary(R"(^,?\s*PRIMARY KEY\s*\((.*)\)\s*$)", QRegularExpression::CaseInsensitiveOption);
    const QRegularExpression index(R"(^CREATE\s+(UNIQUE\s+)?INDEX\s+`((?:``|[^`])+)`\s+ON\s+`((?:``|[^`])+)`\s*\((.*)\);$)", QRegularExpression::CaseInsensitiveOption);
    const QRegularExpression alterKey(R"(^ALTER TABLE `(?:``|[^`])+` ADD PRIMARY KEY \((.*)\);$)", QRegularExpression::CaseInsensitiveOption);
    const QRegularExpression alterIndex(R"(^ALTER TABLE `(?:``|[^`])+` ADD (UNIQUE )?INDEX `((?:``|[^`])+)` \((.*)\);$)", QRegularExpression::CaseInsensitiveOption);
    const auto names = [&](const QString& text) {
        QStringList result;
        auto matches = identifier.globalMatch(text);
        while (matches.hasNext()) result << matches.next().captured(1).replace("``", "`");
        return result;
    };
    AccessTableSchema result;
    QStringList keys;
    bool inTable = false;
    for (const auto& raw : schema.split('\n')) {
        const auto text = raw.trimmed();
        if (text.startsWith("CREATE TABLE ", Qt::CaseInsensitive)) {
            if (!result.columns.empty()) return {}; // exactly one table
            inTable = true;
            continue;
        }
        if (inTable && text.startsWith(')')) { inTable = false; continue; }
        if (!inTable) {
            const auto pk = alterKey.match(text);
            if (pk.hasMatch()) keys = names(pk.captured(1));
            const auto idx = alterIndex.match(text);
            if (idx.hasMatch()) result.indexes.append({idx.captured(2).replace("``", "`"),
                names(idx.captured(3)), !idx.captured(1).isEmpty()});
            const auto match = index.match(text);
            if (match.hasMatch()) result.indexes.append({match.captured(2).replace("``", "`"),
                names(match.captured(4)), !match.captured(1).isEmpty()});
            continue;
        }
        if (text == "(" || text.isEmpty()) continue;
        const auto key = primary.match(text);
        if (key.hasMatch()) { keys = names(key.captured(1)); continue; }
        const auto match = column.match(raw);
        if (!match.hasMatch()) return {};
        const auto native = match.captured(2).toUpper();
        QString type;
        if (native == "INTEGER" || native == "INT" || native == "SMALLINT"
            || native == "TINYINT" || native == "BIT") type = "INTEGER";
        else if (native == "REAL" || native == "FLOAT" || native == "DOUBLE"
                 || native == "NUMERIC" || native == "DECIMAL") type = "REAL";
        else if (native == "VARCHAR" || native == "TEXT" || native == "CHAR"
                 || native == "DATETIME" || native == "DATE" || native == "TIME"
                 || native == "MEMO") type = "TEXT";
        else if (native == "BLOB") type = "BLOB";
        else return {};
        result.columns.append({match.captured(1).replace("``", "`"), type, native,
            match.captured(3).contains("NOT NULL", Qt::CaseInsensitive), 0});
    }
    for (int i = 0; i < keys.size(); ++i) {
        bool found = false;
        for (auto& col : result.columns) {
            if (col.name == keys[i]) { col.primaryKeyOrder = i + 1; found = true; }
        }
        if (!found) return {};
    }
    return result;
}

QVector<AccessColumn> parseMdbColumns(const QString& schema) {
    return parseMdbSchema(schema).columns;
}

static void importAccessTablePhase(QSqlDatabase db, const QString& sourceFile,
                       const QString& tableName, const AccessTableSchema& schema,
                       const std::vector<std::vector<CsvField>>& rows, bool insertOnly = false) {
    const auto& columns = schema.columns;
    if (!db.isOpen() || columns.isEmpty() || rows.empty())
        throw fail("Datenbank, Schema oder CSV-Daten fehlen");
    if (tableName.startsWith("_la_", Qt::CaseInsensitive)
        || tableName.startsWith("sqlite_", Qt::CaseInsensitive))
        throw fail("Reservierter Tabellenname");
    if (static_cast<std::size_t>(columns.size()) != rows.front().size())
        throw fail("Schema und CSV-Spaltenanzahl stimmen nicht überein");
    std::set<QString> seen;
    QStringList definitions;
    for (int i = 0; i < columns.size(); ++i) {
        const auto& column = columns[i];
        if (column.name.isEmpty() || seen.contains(column.name.toCaseFolded())
            || rows[0][i].value != column.name.toUtf8().toStdString()
            || (column.sqliteType != "TEXT" && column.sqliteType != "INTEGER"
                && column.sqliteType != "REAL" && column.sqliteType != "BLOB"))
            throw fail("Spaltenname/-typ stimmt nicht mit MDBTools überein");
        seen.insert(column.name.toCaseFolded());
        definitions << sqlQuote(column.name) + ' ' + column.sqliteType
            + (column.required || column.primaryKeyOrder ? " NOT NULL" : "");
    }

    QStringList primary;
    for (int order = 1; order <= columns.size(); ++order)
        for (const auto& col : columns)
            if (col.primaryKeyOrder == order) primary << sqlQuote(col.name);
    if (!primary.empty()) definitions << "PRIMARY KEY (" + primary.join(',') + ")";
    for (const auto& key : schema.foreignKeys) {
        QStringList child, parent;
        for (const auto& name : key.columns) child << sqlQuote(name);
        for (const auto& name : key.parentColumns) parent << sqlQuote(name);
        definitions << "FOREIGN KEY (" + child.join(',') + ") REFERENCES " + sqlQuote(key.parentTable)
            + " (" + parent.join(',') + ") ON UPDATE " + (key.cascadeUpdate ? "CASCADE" : "NO ACTION")
            + " ON DELETE " + (key.cascadeDelete ? "CASCADE" : "NO ACTION")
            + " DEFERRABLE INITIALLY DEFERRED";
    }
    QSqlQuery scope(db);
    if (!scope.exec("SAVEPOINT la_table_import")) throw fail(scope.lastError().text());
    try {
        QSqlQuery query(db);
        if (!insertOnly && !query.exec("CREATE TABLE " + sqlQuote(tableName) + " (" + definitions.join(",") + ")"))
            throw fail("CREATE TABLE: " + query.lastError().text());
        QStringList placeholders;
        for (int i = 0; i < columns.size(); ++i) placeholders << "?";
        if (rows.size() > 1 && !query.prepare("INSERT INTO " + sqlQuote(tableName) + " VALUES ("
                           + placeholders.join(',') + ")"))
            throw fail("INSERT: " + query.lastError().text());
        for (std::size_t r = 1; r < rows.size(); ++r) {
            if (rows[r].size() != static_cast<std::size_t>(columns.size()))
                throw fail("CSV-Zeile mit abweichender Feldanzahl: " + QString::number(r));
            for (int c = 0; c < columns.size(); ++c) {
                const auto& cell = rows[r][c];
                const QString value = QString::fromUtf8(cell.value.data(), cell.value.size());
                QVariant converted;
                // ASCII 0x1E is supplied by `mdb-export -0`: only unquoted
                // single-character markers mean NULL, not quoted literal text.
                if (!cell.quoted && cell.value == "\x1e") {
                    converted = QVariant();
                } else if (columns[c].sqliteType == "BLOB") {
                    // MDBTools CSV exports are not an unambiguous OLE binary
                    // serialization. Preserve NULLs, but do not reinterpret
                    // arbitrary bytes as images or silently corrupt them.
                    throw fail("Nicht-NULL-OLE/BLOB-Daten brauchen einen Binäradapter: "
                               + columns[c].name + " Zeile " + QString::number(r));
                } else if (columns[c].sqliteType == "INTEGER") {
                    bool ok = false;
                    const auto number = value.toLongLong(&ok);
                    if (!ok) throw fail("Ungültiger INTEGER-Wert in Zeile " + QString::number(r));
                    converted = number;
                } else if (columns[c].sqliteType == "REAL") {
                    bool ok = false;
                    const auto number = value.toDouble(&ok);
                    if (!ok) throw fail("Ungültiger REAL-Wert in Zeile " + QString::number(r));
                    converted = number;
                } else converted = value;
                query.bindValue(c, converted);
            }
            if (!query.exec()) throw fail("INSERT: " + query.lastError().text());
        }
        if (!insertOnly) {
            for (const auto& index : schema.indexes) {
                QStringList fields;
                for (const auto& name : index.columns) {
                    if (!seen.contains(name.toCaseFolded())) throw fail("Indexspalte fehlt");
                    fields << sqlQuote(name);
                }
                if (fields.empty()) throw fail("Leerer Index");
                // Access index names are table-scoped; SQLite names are database-scoped.
                const auto indexName = tableName + "_" + index.name;
                if (!query.exec("CREATE " + QString(index.unique ? "UNIQUE " : "") + "INDEX "
                                + sqlQuote(indexName) + " ON " + sqlQuote(tableName)
                                + " (" + fields.join(',') + ")"))
                    throw fail("CREATE INDEX: " + query.lastError().text());
            }
            if (!query.exec("CREATE TABLE IF NOT EXISTS _la_table_schema "
                            "(table_name TEXT PRIMARY KEY, definition TEXT NOT NULL)"))
                throw fail(query.lastError().text());
            QJsonArray metadata;
            for (const auto& col : columns) metadata.append(QJsonObject{
                {"name", col.name}, {"source_type", col.sourceType}, {"sqlite_type", col.sqliteType},
                {"required", col.required}, {"primary_key_order", col.primaryKeyOrder}});
            query.prepare("INSERT INTO _la_table_schema VALUES (?,?)");
            query.addBindValue(tableName);
            query.addBindValue(QString::fromUtf8(QJsonDocument(metadata).toJson(QJsonDocument::Compact)));
            if (!query.exec()) throw fail(query.lastError().text());
            QSqlQuery log(db);
            if (!log.prepare("INSERT INTO _la_import_log(source_file,kind,name,status,details) "
                             "VALUES (?,?,?,?,?)"))
                throw fail(log.lastError().text());
            log.addBindValue(sourceFile);
            log.addBindValue("table");
            log.addBindValue(tableName);
            log.addBindValue("partial_typed_nulls");
            log.addBindValue("MDBTools schema+CSV; SQLite types and NULLs retained; "
                             "PK and exported indices retained; FK retained by full-database import only; defaults/OLE/Access semantics NOT retained");
            if (!log.exec()) throw fail(log.lastError().text());
        }
        if (!scope.exec("RELEASE la_table_import")) throw fail(scope.lastError().text());
    } catch (...) {
        scope.exec("ROLLBACK TO la_table_import");
        scope.exec("RELEASE la_table_import");
        throw;
    }
}

void importAccessTable(QSqlDatabase db, const QString& sourceFile,
                       const QString& tableName, const AccessTableSchema& schema,
                       const std::vector<std::vector<CsvField>>& rows) {
    importAccessTablePhase(db, sourceFile, tableName, schema, rows);
}

void importAccessTable(QSqlDatabase db, const QString& sourceFile,
                       const QString& tableName, const QVector<AccessColumn>& columns,
                       const std::vector<std::vector<CsvField>>& rows) {
    importAccessTable(db, sourceFile, tableName, AccessTableSchema{columns, {}, {}}, rows);
}
int importAccessDatabase(QSqlDatabase db, const QString& sourceFile) {
    // Read every schema and row from one immutable working copy. In particular,
    // restoring/replacing a source while importing must never mix two versions.
    QTemporaryDir snapshot;
    if (!snapshot.isValid()) throw fail("Temporäre MDB-Kopie kann nicht angelegt werden");
    QFile source(sourceFile);
    if (!source.open(QIODevice::ReadOnly)) throw fail("MDB-Quelle kann nicht gelesen werden");
    QCryptographicHash hash(QCryptographicHash::Sha256);
    if (!hash.addData(&source)) throw fail("MDB-Quelle kann nicht vollständig gelesen werden");
    const auto before = hash.result();
    source.close();
    const auto readPath = snapshot.filePath("source.mdb");
    if (!QFile::copy(sourceFile, readPath)) throw fail("MDB-Kopie fehlgeschlagen");
    QFile copy(readPath);
    if (!copy.open(QIODevice::ReadOnly)) throw fail("MDB-Kopie kann nicht gelesen werden");
    hash.reset();
    if (!hash.addData(&copy) || hash.result() != before) throw fail("MDB-Quelle wurde während der Kopie verändert");
    copy.close();
    AccessCatalog catalog(readPath);
    const auto tables = QString::fromUtf8(AccessCatalog::tool("mdb-tables", {"-1", readPath}))
        .split('\n', Qt::SkipEmptyParts);
    QMap<QString, AccessTableSchema> schemas;
    for (const auto& table : tables) {
        if (table.startsWith("MSys", Qt::CaseInsensitive)) continue;
        if (db.tables().contains(table, Qt::CaseInsensitive))
            throw fail("Zieltabelle existiert bereits; vollständiger Import wurde nicht begonnen");
        auto schema = parseMdbSchema(catalog.tableSchema(table));
        if (schema.columns.empty()) throw fail("Nicht unterstütztes Tabellenschema");
        schemas.insert(table, schema);
    }
    QMap<QString, QVector<AccessRelation>> groups;
    for (const auto& relation : catalog.relationships()) groups[relation.name].append(relation);
    for (auto relations : groups) {
        std::sort(relations.begin(), relations.end(), [](const auto& a, const auto& b) { return a.ordinal < b.ordinal; });
        const auto& first = relations.first();
        if ((first.flags & 2) || !schemas.contains(first.childTable) || !schemas.contains(first.parentTable)) continue;
        AccessForeignKey key;
        key.parentTable = first.parentTable;
        key.cascadeUpdate = first.flags & 0x100;
        key.cascadeDelete = first.flags & 0x1000;
        for (const auto& relation : relations) {
            key.columns << relation.childColumn;
            key.parentColumns << relation.parentColumn;
        }
        schemas[first.childTable].foreignKeys.append(key);
    }
    QSqlQuery query(db);
    if (!query.exec("PRAGMA foreign_keys=ON")) throw fail(query.lastError().text());
    if (!query.exec("PRAGMA foreign_keys") || !query.next() || !query.value(0).toBool())
        throw fail("Fremdschlüsselprüfung muss vor der Importtransaktion aktiviert sein");
    query.finish();
    if (!db.transaction()) throw fail(db.lastError().text());
    try {
        if (!query.exec("CREATE TABLE IF NOT EXISTS _la_access_relations "
                        "(source_file TEXT, name TEXT, flags INTEGER, ordinal INTEGER, "
                        "child_table TEXT, child_column TEXT, parent_table TEXT, parent_column TEXT)"))
            throw fail(query.lastError().text());
        for (const auto& relations : groups) for (const auto& r : relations) {
            query.prepare("INSERT INTO _la_access_relations VALUES (?,?,?,?,?,?,?,?)");
            for (const auto& value : QVariantList{sourceFile, r.name, r.flags, r.ordinal,
                  r.childTable, r.childColumn, r.parentTable, r.parentColumn}) query.addBindValue(value);
            if (!query.exec()) throw fail(query.lastError().text());
        }
        // Create all referenced tables before preparing any INSERT. Deferred
        // constraints alone do not allow SQLite to reference a missing table.
        for (auto it = schemas.cbegin(); it != schemas.cend(); ++it) {
            std::vector<CsvField> header;
            for (const auto& col : it.value().columns) header.push_back({col.name.toUtf8().toStdString(), true});
            importAccessTablePhase(db, sourceFile, it.key(), it.value(), {header});
        }
        for (auto it = schemas.cbegin(); it != schemas.cend(); ++it) {
            const auto csv = AccessCatalog::tool("mdb-export", {"-0", QString(QChar(0x1e)), readPath, it.key()});
            importAccessTablePhase(db, sourceFile, it.key(), it.value(),
                parseCsvDetailed(std::string_view(csv.constData(), csv.size())), true);
        }
        if (!db.commit()) throw fail(db.lastError().text());
    } catch (...) { db.rollback(); throw; }
    return schemas.size();
}

}
