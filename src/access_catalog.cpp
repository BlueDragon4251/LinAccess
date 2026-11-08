#include "access_catalog.hpp"
#include "csv.hpp"

#include <QProcess>
#include <QCoreApplication>
#include <QDir>
#include <QFileInfo>
#include <QJsonDocument>
#include <QElapsedTimer>
#include <QProcessEnvironment>
#include <QRegularExpression>
#include <QStandardPaths>
#include <QSet>

#include <stdexcept>

namespace linaccess {
AccessCatalog::AccessCatalog(QString file) : file_(std::move(file)) {}

QByteArray AccessCatalog::tool(const QString& name, const QStringList& args,
                                const QByteArray& input) {
    auto path = QStandardPaths::findExecutable(name);
    if(name=="linaccess-mdb-metadata") {
        path.clear(); const auto root=QCoreApplication::applicationDirPath();
        for(const auto& candidate:{root+"/linaccess-mdb-metadata",QDir(root).absoluteFilePath("../lib/linaccess/linaccess-mdb-metadata")})
            if(QFileInfo(candidate).isFile() && QFileInfo(candidate).isExecutable()) { path=candidate; break; }
    }
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
    QByteArray result,diagnostics;
    const auto collect=[&] {
        const auto output=process.readAllStandardOutput(),errors=process.readAllStandardError();
        if(output.size()>20*1024*1024-result.size() || errors.size()>64*1024-diagnostics.size()) {
            process.kill(); process.waitForFinished(2000);
            throw std::runtime_error("MDBTools-Ausgabe überschreitet die begrenzte Puffergröße");
        }
        result+=output; diagnostics+=errors;
    };
    QElapsedTimer elapsed; elapsed.start();
    while(process.state()!=QProcess::NotRunning) {
        process.waitForReadyRead(static_cast<int>(qBound<qint64>(qint64{1},30000-elapsed.elapsed(),qint64{100}))); collect();
        if(elapsed.elapsed()>=30000 && process.state()!=QProcess::NotRunning) {
            process.kill(); process.waitForFinished(2000);
            throw std::runtime_error("MDBTools: 30-Sekunden-Zeitlimit");
        }
    }
    collect();
    if (process.exitStatus() != QProcess::NormalExit || process.exitCode() != 0)
        throw std::runtime_error("MDBTools: " + diagnostics.toStdString());
    // MDBTools can report damaged pages while returning exit code zero.
    // Never accept a silently truncated schema or row set in that case.
    const QRegularExpression fatalDiagnostic("beyond EOF|error|failed|unable|could not",
        QRegularExpression::CaseInsensitiveOption);
    if (fatalDiagnostic.match(QString::fromUtf8(diagnostics)).hasMatch())
        throw std::runtime_error("MDBTools meldet einen Lesefehler; Import wurde abgebrochen");
    return result;
}

QByteArray AccessCatalog::exportTable(const QString& source, const QString& table,bool binaryHex) {
    QStringList args{"-D","%Y-%m-%d %H:%M:%S","-T","%Y-%m-%d %H:%M:%S","-0",QString(QChar(0x1e))};
    if(binaryHex) args<<"-b"<<"hex";
    args<<source<<table; return tool("mdb-export",args);
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

QMap<QString,QString> AccessCatalog::floatingSourceTypes(const QString& accessSchema) {
    QMap<QString,QString> result;
    // MDBTools' Access backend wraps raw names in brackets without escaping ']'.
    // Match the final bracket before the exact type, retaining the name verbatim.
    const QRegularExpression field(R"(^\s*\[(.+)\]\s+(Currency|Single|Double)(?:\s|,|$))",QRegularExpression::CaseInsensitiveOption);
    for(const auto& line:accessSchema.split('\n')) {
        const auto match=field.match(line);
        if(!match.hasMatch()) continue;
        const auto key=match.captured(1),type=match.captured(2).toUpper();
        if(result.contains(key) && result.value(key)!=type) throw std::runtime_error("Conflicting Access floating-type metadata");
        result.insert(key,type);
    }
    return result;
}

QString AccessCatalog::tableSchema(const QString& name) const {
    auto schema = QString::fromUtf8(tool("mdb-schema", {"--default-values", "-T", name, file_, "sqlite"}));
    // MDBTools 1.0.0's SQLite backend omits all keys and indexes. Its MySQL
    // backend exposes them. Read only structural ALTER lines; never execute them.
    {
        const auto metadata = QString::fromUtf8(tool("mdb-schema", {"-T", name, file_, "mysql"}));
        QSet<QString> knownColumns,floatingColumns;
        const bool needKeys = !schema.contains("PRIMARY KEY", Qt::CaseInsensitive);
        const bool needIndexes = !schema.contains("CREATE INDEX", Qt::CaseInsensitive)
            && !schema.contains("CREATE UNIQUE INDEX", Qt::CaseInsensitive);
        const QRegularExpression column(R"(^\s*`((?:``|[^`])+)`\s+([A-Za-z][A-Za-z0-9_]*)(.*)$)");
        for (const auto& line : metadata.split('\n')) {
            const auto field = column.match(line);
            if (field.hasMatch()) {
                const auto columnName=field.captured(1).replace("``","`");
                knownColumns.insert(columnName);
                if(field.captured(2).compare("FLOAT",Qt::CaseInsensitive)==0) floatingColumns.insert(columnName);
                const auto nameBytes = columnName.toUtf8().toHex();
                schema += "\n-- LINACCESS_SOURCE_TYPE " + QString::fromLatin1(nameBytes) + ' ' + field.captured(2).toUpper();
                schema += "\n-- LINACCESS_SOURCE_ORIGIN " + QString::fromLatin1(nameBytes) + " MDBTOOLS_MYSQL";
                if(field.captured(2).compare("VARCHAR",Qt::CaseInsensitive)==0) {
                    const auto length=QRegularExpression(R"(^\s*\(\s*([0-9]+)\s*\))").match(field.captured(3));
                    bool valid=false; const auto size=length.captured(1).toInt(&valid);
                    if(length.hasMatch() && valid && size>=1 && size<=255)
                        schema += "\n-- LINACCESS_SOURCE_SIZE " + QString::fromLatin1(nameBytes) + ' ' + QString::number(size);
                }
                auto modifiers=field.captured(3);
                modifiers.remove(QRegularExpression(R"sql('(?:''|[^'])*'|"(?:""|[^"])*")sql"));
                if (QRegularExpression("\\bauto_increment\\b",QRegularExpression::CaseInsensitiveOption).match(modifiers).hasMatch())
                    schema += "\n-- LINACCESS_AUTO_NUMBER " + QString::fromLatin1(nameBytes);
            }
            if (!line.startsWith("ALTER TABLE ", Qt::CaseInsensitive)) continue;
            if ((needKeys && line.contains(" ADD PRIMARY KEY ", Qt::CaseInsensitive))
                || (needIndexes && (line.contains(" ADD INDEX ", Qt::CaseInsensitive)
                                    || line.contains(" ADD UNIQUE INDEX ", Qt::CaseInsensitive))))
                schema += '\n' + line;
        }
        const auto exactTypes=floatingSourceTypes(QString::fromUtf8(tool("mdb-schema",{"-T",name,file_,"access"})));
        for(auto it=exactTypes.cbegin();it!=exactTypes.cend();++it) {
            if(!knownColumns.contains(it.key())) throw std::runtime_error("Access floating-type column metadata mismatch");
            schema += "\n-- LINACCESS_SOURCE_TYPE " + QString::fromLatin1(it.key().toUtf8().toHex()) + ' ' + it.value();
            schema += "\n-- LINACCESS_SOURCE_ORIGIN " + QString::fromLatin1(it.key().toUtf8().toHex()) + " MDBTOOLS_ACCESS";
            floatingColumns.remove(it.key());
        }
        if(!floatingColumns.isEmpty()) throw std::runtime_error("Access floating source metadata is unavailable");
    }
    const auto indexes=tool("linaccess-mdb-metadata",{file_,name});
    if(indexes.isEmpty() || indexes.size()>1024*1024 || !QJsonDocument::fromJson(indexes).isObject())
        throw std::runtime_error("Originale Indexmetadaten konnten nicht gelesen werden");
    schema+="\n-- LINACCESS_INDEX_METADATA "+QString::fromLatin1(indexes.toHex())+"\n";
    return schema;
}
}
