#include "access_catalog.hpp"
#include "csv.hpp"

#include <QProcess>
#include <QProcessEnvironment>
#include <QRegularExpression>
#include <QStandardPaths>
#include <QSet>

#include <stdexcept>

namespace linaccess {
AccessCatalog::AccessCatalog(QString file) : file_(std::move(file)) {}

QByteArray AccessCatalog::tool(const QString& name, const QStringList& args,
                                const QByteArray& input) {
    const auto path = QStandardPaths::findExecutable(name);
    if (path.isEmpty())
        throw std::runtime_error((name + " fehlt. Installiere mdbtools.").toStdString());
    QProcess process;
    auto environment=QProcessEnvironment::systemEnvironment();
    // GLib parses string options using the child process locale. Keep UTF-8
    // database/table/query arguments valid even when the desktop starts in C.
    environment.insert("LC_ALL","C.UTF-8");
    environment.insert("LANG","C.UTF-8");
    process.setProcessEnvironment(environment);
    process.setProgram(path);
    process.setArguments(args);
    process.start(); // argv only, no shell or command interpolation.
    if (!process.waitForStarted(5000)) throw std::runtime_error("MDBTools konnte nicht starten");
    if (!input.isEmpty()) {
        if (process.write(input) != input.size() || !process.waitForBytesWritten(5000)) {
            process.kill();
            process.waitForFinished(2000);
            throw std::runtime_error("MDBTools: Abfrage konnte nicht übertragen werden");
        }
    }
    process.closeWriteChannel();
    if (!process.waitForFinished(30000)) {
        process.kill();
        process.waitForFinished(2000);
        throw std::runtime_error("MDBTools: 30-Sekunden-Zeitlimit");
    }
    if (process.exitStatus() != QProcess::NormalExit || process.exitCode() != 0)
        throw std::runtime_error("MDBTools: " + process.readAllStandardError().toStdString());
    const auto diagnostics = process.readAllStandardError();
    // MDBTools can report damaged pages while returning exit code zero.
    // Never accept a silently truncated schema or row set in that case.
    const QRegularExpression fatalDiagnostic("beyond EOF|error|failed|unable|could not",
        QRegularExpression::CaseInsensitiveOption);
    if (fatalDiagnostic.match(QString::fromUtf8(diagnostics)).hasMatch())
        throw std::runtime_error("MDBTools meldet einen Lesefehler; Import wurde abgebrochen");
    const auto result = process.readAllStandardOutput();
    if (result.size() > 20 * 1024 * 1024)
        throw std::runtime_error("MDBTools-Ausgabe überschreitet 20 MB");
    return result;
}

QByteArray AccessCatalog::exportTable(const QString& source, const QString& table) {
    return tool("mdb-export",{"-D","%Y-%m-%d %H:%M:%S","-T","%Y-%m-%d %H:%M:%S",
        "-0",QString(QChar(0x1e)),source,table});
}

QString AccessCatalog::formatVersion() const {
    return QString::fromUtf8(tool("mdb-ver", {file_})).trimmed();
}

QVector<AccessCatalogEntry> AccessCatalog::objects() const {
    // Query only small scalar system columns. MSysObjects contains binary large
    // object properties which MUST NOT be exported as TSV or interpreted as VBA.
    const auto output = tool("mdb-sql", {"-P", "-H", "-F", file_},
                             "SELECT Name, Type, Flags FROM MSysObjects;\n");
    QVector<AccessCatalogEntry> entries;
    for (const auto& row : QString::fromUtf8(output).split('\n', Qt::SkipEmptyParts)) {
        const auto columns = row.split('\t');
        if (columns.size() != 3) continue;
        bool validType = false;
        const int type = columns[1].toInt(&validType);
        if (!validType) continue;
        QString kind;
        switch (type) {
        case 1: kind = "table"; break;
        case 5: kind = "query"; break;
        case 8: kind = "relationship"; break;
        case -32768: kind = "form"; break;
        case -32764: kind = "report"; break;
        case -32761: kind = "module"; break;
        case -32766: kind = "macro"; break;
        default: continue;
        }
        const auto name = columns[0].trimmed();
        if (name.isEmpty()) continue;
        const bool internal = name.startsWith("MSys", Qt::CaseInsensitive)
                            || name.startsWith('~')
                            || columns[2].toInt() != 0;
        entries.append({kind, name, internal});
    }
    return entries;
}

QVector<AccessRelation> AccessCatalog::relationships() const {
    const auto csv = tool("mdb-export", {file_, "MSysRelationships"});
    const auto rows = parseCsv(std::string_view(csv.constData(), csv.size()));
    QVector<AccessRelation> result;
    if (rows.empty()) return result;
    // MDBTools exports: name, flags, columns, order, child, child col, parent, parent col.
    for (std::size_t i = 1; i < rows.size(); ++i) {
        if (rows[i].size() < 8) continue;
        const auto toString = [](const std::string& s) { return QString::fromUtf8(s); };
        result.append({toString(rows[i][4]), toString(rows[i][5]),
                       toString(rows[i][6]), toString(rows[i][7]), toString(rows[i][0]),
                       toString(rows[i][1]).toUInt(), toString(rows[i][3]).toInt()});
    }
    return result;
}

QString AccessCatalog::querySql(const QString& name) const {
    return QString::fromUtf8(tool("mdb-queries", {file_, name})).trimmed();
}

QStringList AccessCatalog::missingQuerySources(const QString& sql) const {
    QSet<QString> existing;
    for (const auto& object : objects()) {
        if (object.kind == "query" || object.kind == "table")
            existing.insert(object.name.toCaseFolded());
    }
    QSet<QString> missing;
    // Bracket-qualified identifiers in Jet SQL: [Object].[Field]. This is a
    // diagnostic heuristic, not a full Jet SQL parser (aliases may be ambiguous).
    const QRegularExpression sourceRef(R"(\[([^\]]+)\]\s*\.)");
    auto match = sourceRef.globalMatch(sql);
    while (match.hasNext()) {
        const auto name = match.next().captured(1);
        if ((name.contains("_tbl_", Qt::CaseInsensitive)
             || name.contains("_qry_", Qt::CaseInsensitive))
            && !existing.contains(name.toCaseFolded())) missing.insert(name);
    }
    auto result = missing.values();
    result.sort(Qt::CaseInsensitive);
    return result;
}

QString AccessCatalog::tableSchema(const QString& name) const {
    auto schema = QString::fromUtf8(tool("mdb-schema", {"-T", name, file_, "sqlite"}));
    // MDBTools 1.0.0's SQLite backend omits all keys and indexes. Its MySQL
    // backend exposes them. Read only structural ALTER lines; never execute them.
    {
        const auto metadata = QString::fromUtf8(tool("mdb-schema", {"-T", name, file_, "mysql"}));
        const bool needKeys = !schema.contains("PRIMARY KEY", Qt::CaseInsensitive);
        const bool needIndexes = !schema.contains("CREATE INDEX", Qt::CaseInsensitive)
            && !schema.contains("CREATE UNIQUE INDEX", Qt::CaseInsensitive);
        const QRegularExpression column(R"(^\s*`((?:``|[^`])+)`\s+([A-Za-z][A-Za-z0-9_]*))");
        for (const auto& line : metadata.split('\n')) {
            const auto field = column.match(line);
            if (field.hasMatch()) {
                const auto nameBytes = field.captured(1).replace("``","`").toUtf8().toHex();
                schema += "\n-- LINACCESS_SOURCE_TYPE " + QString::fromLatin1(nameBytes) + ' ' + field.captured(2).toUpper();
            }
            if (!line.startsWith("ALTER TABLE ", Qt::CaseInsensitive)) continue;
            if ((needKeys && line.contains(" ADD PRIMARY KEY ", Qt::CaseInsensitive))
                || (needIndexes && (line.contains(" ADD INDEX ", Qt::CaseInsensitive)
                                    || line.contains(" ADD UNIQUE INDEX ", Qt::CaseInsensitive))))
                schema += '\n' + line;
        }
    }
    return schema;
}
}
