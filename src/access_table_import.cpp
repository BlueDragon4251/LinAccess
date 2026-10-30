#include "access_table_import.hpp"

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

QVector<AccessColumn> parseMdbColumns(const QString& schema) {
    const QRegularExpression column(R"(^\s*`([^`]+)`\s+([A-Za-z][A-Za-z0-9_]*))");
    QVector<AccessColumn> columns;
    bool inTable = false;
    for (const auto& raw : schema.split('\n')) {
        const auto text = raw.trimmed();
        if (text.startsWith("CREATE TABLE ", Qt::CaseInsensitive)) {
            inTable = true;
            continue;
        }
        if (!inTable) continue;
        if (text.startsWith(')')) break;
        if (text == "(" || text.isEmpty()) continue;
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
        else return {}; // Unknown column types require a dedicated adapter.
        columns.append({match.captured(1), type});
    }
    return columns;
}

void importAccessTable(QSqlDatabase db, const QString& sourceFile,
                       const QString& tableName, const QVector<AccessColumn>& columns,
                       const std::vector<std::vector<CsvField>>& rows) {
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
        definitions << sqlQuote(column.name) + ' ' + column.sqliteType;
    }

    if (!db.transaction()) throw fail("SQLite-Transaktion fehlgeschlagen: " + db.lastError().text());
    try {
        QSqlQuery query(db);
        if (!query.exec("CREATE TABLE " + sqlQuote(tableName) + " (" + definitions.join(",") + ")"))
            throw fail("CREATE TABLE: " + query.lastError().text());
        QStringList placeholders;
        for (int i = 0; i < columns.size(); ++i) placeholders << "?";
        if (!query.prepare("INSERT INTO " + sqlQuote(tableName) + " VALUES ("
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
        QSqlQuery log(db);
        if (!log.prepare("INSERT INTO _la_import_log(source_file,kind,name,status,details) "
                         "VALUES (?,?,?,?,?)"))
            throw fail(log.lastError().text());
        log.addBindValue(sourceFile);
        log.addBindValue("table");
        log.addBindValue(tableName);
        log.addBindValue("partial_typed_nulls");
        log.addBindValue("MDBTools schema+CSV; SQLite types and NULLs retained; "
                         "PK/indices/FK/defaults/OLE/Access semantics NOT retained");
        if (!log.exec()) throw fail(log.lastError().text());
        if (!db.commit()) throw fail("COMMIT: " + db.lastError().text());
    } catch (...) {
        db.rollback();
        throw;
    }
}
}
