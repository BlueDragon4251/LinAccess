#include "access_table_import.hpp"
#include "access_catalog.hpp"
#include "access_index_metadata.hpp"
#include "access_floating_metadata.hpp"
#include "access_autonumber.hpp"
#include "jet_sql_functions.hpp"
#include <QMap>
#include <QFile>
#include <QTemporaryDir>
#include <QCryptographicHash>
#include <QJsonArray>
#include <QJsonObject>
#include <QJsonDocument>
#include <algorithm>
#include <cmath>
#include <optional>
#include <limits>

#include <QRegularExpression>
#include <QUuid>
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
std::optional<QString> nativeDefaultLiteral(const QString& token) {
    if(token.compare("NULL",Qt::CaseInsensitive)==0) return QString("NULL");
    if(token.compare("TRUE",Qt::CaseInsensitive)==0) return QString("-1");
    if(token.compare("FALSE",Qt::CaseInsensitive)==0) return QString("0");
    if(QRegularExpression(R"(^[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?$)").match(token).hasMatch()) {
        bool valid=false; const auto number=token.toDouble(&valid);
        if(valid && std::isfinite(number)) return token;
        return {};
    }
    if(token.size()<2 || (token.front()!=QChar(39) && token.front()!=QChar(34)) || token.back()!=token.front()) return {};
    QString text; const auto quote=token.front();
    for(qsizetype i=1;i<token.size()-1;++i) {
        if(token[i]==quote) {
            if(i+1>=token.size()-1 || token[i+1]!=quote) return {};
            ++i;
        }
        text+=token[i];
    }
    text.replace(QChar(39),QString("''")); return QChar(39)+text+QChar(39);
}
std::runtime_error fail(const QString& error) {
    return std::runtime_error(error.toUtf8().toStdString());
}
}

std::optional<QString> accessLiteralDefaultSql(const QString& token) {auto literal=token.trimmed();if(literal.startsWith('=')) literal=literal.mid(1).trimmed();return nativeDefaultLiteral(literal); }

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
    const auto orderedIndex=[&](QString name,const QString& fields,bool unique)->std::optional<AccessIndex> {
        static const QRegularExpression part(R"(^\s*`((?:``|[^`])+)`(?:\s+(ASC|DESC))?\s*(,|$))",QRegularExpression::CaseInsensitiveOption);
        AccessIndex result; result.name=std::move(name); result.unique=unique;
        qsizetype offset=0;
        while(offset<fields.size()) {
            const auto match=part.match(fields.mid(offset));
            if(!match.hasMatch() || match.capturedLength()==0 || result.columns.size()>=32767) return {};
            result.columns.append(match.captured(1).replace("``","`"));
            result.descending.append(match.captured(2).compare("DESC",Qt::CaseInsensitive)==0);
            offset+=match.capturedLength();
            if(match.captured(3).isEmpty()) break;
            if(offset>=fields.size()) return {};
        }
        if(result.columns.isEmpty() || offset!=fields.size()) return {};
        return result;
    };
    AccessTableSchema result;
    QStringList keys;
    QMap<QString,QString> sourceTypes,sourceOrigins;
    QMap<QString,int> sourceSizes;
    QMap<QString,bool> counters;
    bool inTable = false;
    for (const auto& raw : schema.split('\n')) {
        const auto text = raw.trimmed();
        if(text.startsWith("-- LINACCESS_INDEX_METADATA ")) {
            if(!result.sourceIndexMetadata.isEmpty()) return {};
            const auto encoded=text.mid(QString("-- LINACCESS_INDEX_METADATA ").size()).toLatin1();
            if(encoded.isEmpty() || encoded.size()>2*1024*1024 || encoded.size()%2) return {};
            for(const auto ch:encoded) if(!((ch>='0' && ch<='9') || (ch>='a' && ch<='f') || (ch>='A' && ch<='F'))) return {};
            result.sourceIndexMetadata=QByteArray::fromHex(encoded); continue;
        }
        static const QRegularExpression sourceType(R"(^-- LINACCESS_SOURCE_TYPE ([0-9a-fA-F]+) ([A-Z][A-Z0-9_]*)$)");
        const auto source = sourceType.match(text);
        const auto origin=QRegularExpression(R"(^-- LINACCESS_SOURCE_ORIGIN ([0-9a-fA-F]+) (MDBTOOLS_MYSQL|MDBTOOLS_ACCESS)$)").match(text);
        if(origin.hasMatch()) {
            sourceOrigins.insert(QString::fromUtf8(QByteArray::fromHex(origin.captured(1).toLatin1())),origin.captured(2)); continue;
        }
        if(text.startsWith("-- LINACCESS_SOURCE_SIZE ")) {
            const auto length=QRegularExpression(R"(^-- LINACCESS_SOURCE_SIZE ([0-9a-fA-F]+) ([0-9]+)$)").match(text);
            bool valid=false; const auto size=length.captured(2).toInt(&valid);
            if(!length.hasMatch() || !valid || size<1 || size>255) return {};
            const auto name=QString::fromUtf8(QByteArray::fromHex(length.captured(1).toLatin1()));
            if(sourceSizes.contains(name) && sourceSizes.value(name)!=size) return {};
            sourceSizes.insert(name,size); continue;
        }
        static const QRegularExpression autoNumber(R"(^-- LINACCESS_AUTO_NUMBER ([0-9a-fA-F]+)$)");
        const auto counter=autoNumber.match(text);
        if (counter.hasMatch()) {
            counters.insert(QString::fromUtf8(QByteArray::fromHex(counter.captured(1).toLatin1())),true);
            continue;
        }
        if (source.hasMatch()) {
            sourceTypes.insert(QString::fromUtf8(QByteArray::fromHex(source.captured(1).toLatin1())),source.captured(2));
            continue;
        }
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
            if (idx.hasMatch()) {
                const auto parsed=orderedIndex(idx.captured(2).replace("``","`"),idx.captured(3),!idx.captured(1).isEmpty());
                if(!parsed) return {}; result.indexes.append(*parsed);
            }
            const auto match = index.match(text);
            if (match.hasMatch()) {
                const auto parsed=orderedIndex(match.captured(2).replace("``","`"),match.captured(4),!match.captured(1).isEmpty());
                if(!parsed) return {}; result.indexes.append(*parsed);
            }
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
        const auto rawModifiers=match.captured(3);
        auto modifiers=rawModifiers;
        auto literals=QRegularExpression(R"sql('(?:''|[^'])*'|"(?:""|[^"])*")sql").globalMatch(rawModifiers);
        while(literals.hasNext()) {
            const auto literal=literals.next();
            modifiers.replace(literal.capturedStart(),literal.capturedLength(),QString(literal.capturedLength(),' '));
        }
        const bool generated=QRegularExpression("\\bauto_increment\\b",QRegularExpression::CaseInsensitiveOption).match(modifiers).hasMatch();
        if (generated && type!="INTEGER") return {};
        result.columns.append({match.captured(1).replace("``", "`"), type, native,
            QRegularExpression("\\bNOT\\s+NULL\\b",QRegularExpression::CaseInsensitiveOption).match(modifiers).hasMatch(), 0,generated});
        const auto defaultTag=QRegularExpression("\\bDEFAULT\\b",QRegularExpression::CaseInsensitiveOption).match(modifiers);
        if(defaultTag.hasMatch()) {
            auto token=rawModifiers.mid(defaultTag.capturedEnd()).trimmed();
            if(token.endsWith(',')) { token.chop(1); token=token.trimmed(); }
            const auto translated=nativeDefaultLiteral(token);
            if(!translated) return {};
            result.columns.last().hasDefault=true; result.columns.last().sourceDefault=token; result.columns.last().sqlDefault=*translated;
        }
    }
    for (int i = 0; i < keys.size(); ++i) {
        bool found = false;
        for (auto& col : result.columns) {
            if (col.name == keys[i]) { col.primaryKeyOrder = i + 1; found = true; }
        }
        if (!found) return {};
    }
    for (auto& column : result.columns) {
        if (sourceTypes.contains(column.name)) column.sourceType = sourceTypes.value(column.name);
        column.sourceOrigin=sourceOrigins.take(column.name);
        if(!column.sourceOrigin.isEmpty() && !sourceTypes.contains(column.name)) return {};
        if(sourceSizes.contains(column.name)) {
            if(column.sourceType!="VARCHAR" || column.sourceOrigin!="MDBTOOLS_MYSQL") return {};
            column.sourceSize=sourceSizes.take(column.name);
        }
        if (counters.remove(column.name)) {
            if (column.sqliteType!="INTEGER") return {};
            column.autoNumber=true;
        }
        if (column.autoNumber) column.required=true;
    }
    if (!counters.isEmpty()) return {};
    if(!sourceOrigins.isEmpty() || !sourceSizes.isEmpty()) return {};
    if(!result.sourceIndexMetadata.isEmpty()) {
        try {
            QStringList columnNames,primary; for(const auto& column:result.columns) columnNames.append(column.name);
            for(int order=1;order<=result.columns.size();++order) for(const auto& column:result.columns) if(column.primaryKeyOrder==order) primary.append(column.name);
            const auto indexes=checkedAccessSourceIndexes(result.sourceIndexMetadata,columnNames);
            int primaryCount=0; result.indexes.clear();
            for(const auto& index:indexes) {
                if(index.kind==1) { ++primaryCount; if(index.columns!=primary) return {}; }
                if(index.kind==0) result.indexes.append({index.name,index.columns,(index.flags&1)!=0,index.descending});
            }
            if(primaryCount!=(primary.isEmpty() ? 0 : 1)) return {};
        } catch(...) { return {}; }
    }
    return result;
}

QVector<AccessColumn> parseMdbColumns(const QString& schema) {
    return parseMdbSchema(schema).columns;
}

QByteArray decodeMdbBinaryHex(std::string_view encoded) {
    if(encoded.size()%2 || encoded.size()>32*1024*1024) throw fail("Ungültige oder zu große Binär-Hexlänge");
    for(const unsigned char ch:encoded) if(!((ch>='0' && ch<='9') || (ch>='a' && ch<='f') || (ch>='A' && ch<='F')))
        throw fail("Ungültiger Binär-Hexexport");
    if(encoded.empty()) return QByteArray("",0);
    return QByteArray::fromHex(QByteArray(encoded.data(),static_cast<qsizetype>(encoded.size())));
}

static void importAccessTablePhase(QSqlDatabase db, const QString& sourceFile,
                       const QString& tableName, const AccessTableSchema& schema,
                       const std::vector<std::vector<CsvField>>& rows, bool insertOnly = false,
                       AccessBinaryEncoding binaryEncoding=AccessBinaryEncoding::Unsupported,
                       AccessFloatingSource floatingSource=AccessFloatingSource::Csv,const QString& sourceReadPath={}) {
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
    int primaryCount=0;
    QStringList actualNames,declaredPrimary;
    for (const auto& column:columns) { actualNames.append(column.name); if (column.primaryKeyOrder) ++primaryCount; }
    for(int order=1;order<=columns.size();++order) for(const auto& column:columns) if(column.primaryKeyOrder==order) declaredPrimary.append(column.name);
    const auto sourceIndexes=schema.sourceIndexMetadata.isEmpty() ? QVector<AccessSourceIndex>{} : checkedAccessSourceIndexes(schema.sourceIndexMetadata,actualNames);
    std::optional<AccessSingleValues> exactSingles;
    std::optional<AccessDoubleValues> exactDoubles;
    const auto& floatingReadPath=sourceReadPath.isEmpty() ? sourceFile : sourceReadPath;
    if(floatingSource==AccessFloatingSource::NativeMdbBits && schema.sourceIndexMetadata.isEmpty())
        for(const auto& column:columns) if(column.sourceType=="SINGLE" || column.sourceType=="DOUBLE") throw fail("Originale Gleitkomma-Quellmetadaten fehlen");
    if(floatingSource==AccessFloatingSource::NativeMdbBits && !schema.sourceIndexMetadata.isEmpty()) {
        const auto sourceRoot=QJsonDocument::fromJson(schema.sourceIndexMetadata).object();
        const auto singleNames=sourceRoot.value("single_columns"); QStringList names;
        for(const auto& column:columns) if(column.sourceType=="SINGLE") {
            if(column.sqliteType!="REAL") throw fail("Single-Quelltyp benötigt nachgewiesenen REAL-Speicher");
            names.append(column.name);
        }
        if(!names.isEmpty()) {
            if(!singleNames.isArray() || singleNames.toArray().size()!=names.size()) throw fail("Single-Quelltyp und Bibliothek widersprechen sich");
            for(int i=0;i<names.size();++i) if(singleNames.toArray()[i].toString()!=names[i]) throw fail("Single-Quellspalten widersprechen sich");
            if(insertOnly || rows.size()>1) {
                const auto sourceTable=sourceRoot.value("source_table").toString();
                exactSingles=checkedAccessSingleValues(AccessCatalog(floatingReadPath).singleValues(sourceTable),names,sourceTable);
                if(exactSingles->rowCount!=static_cast<qsizetype>(rows.size()-1)) throw fail("Single-Quelle und CSV-Zeilenzahl widersprechen sich");
                if(accessSourceDigest(floatingReadPath)!=exactSingles->sourceDigest) throw fail("Single-Quelle wurde während des Lesens verändert");
            }
        } else if(singleNames.isArray() && !singleNames.toArray().isEmpty()) throw fail("Single-Quelltypen sind nicht nachgewiesen");
        const auto doubleNames=sourceRoot.value("double_columns"); QStringList doubles;
        for(const auto& column:columns) if(column.sourceType=="DOUBLE") {
            if(column.sqliteType!="REAL") throw fail("Double-Quelltyp benötigt nachgewiesenen REAL-Speicher"); doubles.append(column.name);
        }
        if(!doubles.isEmpty()) {
            if(!doubleNames.isArray() || doubleNames.toArray().size()!=doubles.size()) throw fail("Double-Quelltyp und Bibliothek widersprechen sich");
            for(int i=0;i<doubles.size();++i) if(doubleNames.toArray()[i].toString()!=doubles[i]) throw fail("Double-Quellspalten widersprechen sich");
            if(insertOnly || rows.size()>1) {
                const auto sourceTable=sourceRoot.value("source_table").toString();
                exactDoubles=checkedAccessDoubleValues(AccessCatalog(floatingReadPath).doubleValues(sourceTable),doubles,sourceTable);
                if(exactDoubles->rowCount!=static_cast<qsizetype>(rows.size()-1)) throw fail("Double-Quelle und CSV-Zeilenzahl widersprechen sich");
                if(accessSourceDigest(floatingReadPath)!=exactDoubles->sourceDigest) throw fail("Double-Quelle wurde während des Lesens verändert");
            }
        } else if(doubleNames.isArray() && !doubleNames.toArray().isEmpty()) throw fail("Double-Quelltypen sind nicht nachgewiesen");
    }
    QMap<QString,qint32> sourceCounterStates;
    QByteArray counterDigest;
    if(floatingSource==AccessFloatingSource::NativeMdbBits) {
        QStringList counterNames;
        for(const auto& column:columns) if(column.autoNumber) counterNames.append(column.name);
        if(!counterNames.isEmpty()) {
            const auto root=QJsonDocument::fromJson(schema.sourceIndexMetadata).object();
            const auto states=root.value("counter_columns");
            counterDigest=root.value("source_sha256").toString().toLatin1();
            if(!states.isArray() || states.toArray().size()!=counterNames.size() || counterDigest.size()!=64
                || root.value("source_table").toString().isEmpty()
                || accessSourceDigest(floatingReadPath)!=counterDigest)
                throw fail("Gespeicherter AutoWert-Quellstand fehlt oder wurde verändert");
            for(const auto& state:states.toArray()) {
                if(!state.isObject()) throw fail("Ungültiger AutoWert-Quellstand");
                const auto object=state.toObject(); const auto name=object.value("name").toString();
                const auto value=object.value("last_value"); const auto number=value.toDouble(-1);
                if(!counterNames.contains(name) || sourceCounterStates.contains(name) || !value.isDouble()
                    || number<0 || number>std::numeric_limits<qint32>::max() || std::floor(number)!=number)
                    throw fail("Widersprüchlicher AutoWert-Quellstand");
                sourceCounterStates.insert(name,static_cast<qint32>(number));
            }
        }
    }
    QMap<QString,bool> primaryDirections; int storedPrimary=0,storedOrdinary=0;
    for(const auto& index:sourceIndexes) {
        if(index.kind==1) {
            ++storedPrimary; if(index.columns!=declaredPrimary) throw fail("Primärindex und Tabellenschlüssel widersprechen sich");
            for(int i=0;i<index.columns.size();++i) primaryDirections.insert(index.columns[i],index.descending[i]);
        }
        if(index.kind==0) {
            ++storedOrdinary; int matches=0;
            for(const auto& native:schema.indexes) if(native.name==index.name && native.columns==index.columns && native.unique==((index.flags&1)!=0) && native.descending==index.descending) ++matches;
            if(matches!=1 || (index.flags&2)) throw fail("Indexdefinition oder IgnoreNulls ist nicht nachgewiesen");
            if(index.flags&8) for(const auto& name:index.columns) {
                bool required=false; for(const auto& column:columns) if(column.name==name) required=column.required || column.primaryKeyOrder>0;
                if(!required) throw fail("Index.Required kann nativ nicht nachgewiesen werden");
            }
        }
    }
    if(!schema.sourceIndexMetadata.isEmpty() && (storedPrimary!=(declaredPrimary.isEmpty() ? 0 : 1) || storedOrdinary!=schema.indexes.size()))
        throw fail("Originaler Indexkatalog ist unvollständig");
    const bool descendingSinglePrimary=primaryCount==1 && !declaredPrimary.isEmpty() && primaryDirections.value(declaredPrimary.front());
    const auto nativeCounter=[&](const AccessColumn& column) { return !sourceCounterStates.contains(column.name) && column.autoNumber && column.sqliteType=="INTEGER" && column.primaryKeyOrder==1 && primaryCount==1; };
    for (int i = 0; i < columns.size(); ++i) {
        const auto& column = columns[i];
        if (column.name.isEmpty() || seen.contains(column.name.toCaseFolded())
            || rows[0][i].value != column.name.toUtf8().toStdString()
            || (column.sqliteType != "TEXT" && column.sqliteType != "INTEGER"
                && column.sqliteType != "REAL" && column.sqliteType != "BLOB")
            || (column.autoNumber && column.sqliteType!="INTEGER")
            || column.sourceSize<0 || column.sourceSize>255
            || (column.sourceSize>0 && (column.sourceType!="VARCHAR" || column.sourceOrigin!="MDBTOOLS_MYSQL")))
            throw fail("Spaltenname/-typ stimmt nicht mit MDBTools überein");
        if(column.hasDefault) {
            const auto literal=nativeDefaultLiteral(column.sourceDefault);
            if(!literal || *literal!=column.sqlDefault) throw fail("Nicht unterstützter oder widersprüchlicher Standardwert");
        } else if(!column.sourceDefault.isEmpty() || !column.sqlDefault.isEmpty()) throw fail("Standardwert-Metadaten widersprüchlich");
        if(nativeCounter(column) && primaryDirections.value(column.name)) throw fail("Absteigender AutoNumber-Primärindex ist nicht nachgewiesen");
        seen.insert(column.name.toCaseFolded());
        const bool jetManaged=sourceCounterStates.contains(column.name);
        definitions << sqlQuote(column.name) + ' ' + (jetManaged ? QString("INT") : column.sqliteType)
            + (nativeCounter(column) ? " PRIMARY KEY AUTOINCREMENT" : descendingSinglePrimary && column.primaryKeyOrder ? " PRIMARY KEY DESC" : QString{})
            + (!nativeCounter(column) && column.autoNumber ? " DEFAULT ("+accessCounterDefault(tableName,column.name)+")" : QString{})
            + (!column.autoNumber && column.hasDefault ? " DEFAULT "+column.sqlDefault : QString{})
            + (column.required || column.primaryKeyOrder || column.autoNumber ? " NOT NULL" : "")
            + (column.sourceSize>0 ? " CHECK (linaccess_jet_text_capacity("+sqlQuote(column.name)+","+QString::number(column.sourceSize)+"))" : QString{});
    }

    QStringList primary;
    for (int order = 1; order <= columns.size(); ++order)
        for (const auto& col : columns)
            if (col.primaryKeyOrder == order) primary << sqlQuote(col.name)+(primaryDirections.value(col.name) ? " DESC" : " ASC");
    const bool inlineKey=descendingSinglePrimary || std::any_of(columns.cbegin(),columns.cend(),nativeCounter);
    if (!primary.empty() && !inlineKey) definitions << "PRIMARY KEY (" + primary.join(',') + ")";
    for (const auto& key : schema.foreignKeys) {
        QStringList child, parent;
        for (const auto& name : key.columns) child << sqlQuote(name);
        for (const auto& name : key.parentColumns) parent << sqlQuote(name);
        definitions << "FOREIGN KEY (" + child.join(',') + ") REFERENCES " + sqlQuote(key.parentTable)
            + " (" + parent.join(',') + ") ON UPDATE " + (key.cascadeUpdate ? "CASCADE" : "NO ACTION")
            + " ON DELETE " + (key.cascadeDelete ? "CASCADE" : "NO ACTION")
            + " DEFERRABLE INITIALLY DEFERRED";
    }
    installJetSqlFunctions(db);
    QSqlQuery scope(db);
    if (!scope.exec("SAVEPOINT la_table_import")) throw fail(scope.lastError().text());
    try {
        QSqlQuery query(db);
        if (!insertOnly && !query.exec("CREATE TABLE " + sqlQuote(tableName) + " (" + definitions.join(",") + ")"))
            throw fail("CREATE TABLE: " + query.lastError().text());
        if (!insertOnly) for (const auto& column:columns) if (column.autoNumber && !nativeCounter(column)) createAccessCounter(db,tableName,column.name);
        if(!insertOnly) for(const auto& column:columns) if(sourceCounterStates.contains(column.name)) {
            createJetLongCounter(db,tableName,column.name);
            protectAccessCounterUpdates(db,tableName,column.name);
        }
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
                const auto proof=exactSingles ? exactSingles->columns.constFind(columns[c].name.toCaseFolded()) : QMap<QString,QVector<std::optional<float>>>::const_iterator{};
                const bool hasSingleProof=exactSingles && proof!=exactSingles->columns.cend();
                const auto doubleProof=exactDoubles ? exactDoubles->columns.constFind(columns[c].name.toCaseFolded()) : QMap<QString,QVector<std::optional<double>>>::const_iterator{};
                const bool hasDoubleProof=exactDoubles && doubleProof!=exactDoubles->columns.cend();
                const bool csvNull=!cell.quoted && cell.value=="\x1e";
                if(hasSingleProof && ((*proof)[static_cast<qsizetype>(r-1)].has_value()==csvNull)) throw fail("Single-NULL-Quelle und CSV widersprechen sich");
                if(hasDoubleProof && ((*doubleProof)[static_cast<qsizetype>(r-1)].has_value()==csvNull)) throw fail("Double-NULL-Quelle und CSV widersprechen sich");
                // ASCII 0x1E is supplied by `mdb-export -0`: only unquoted
                // single-character markers mean NULL, not quoted literal text.
                if (!cell.quoted && cell.value == "\x1e") {
                    converted = QVariant();
                } else if (columns[c].sourceType == "BOOLEAN" || columns[c].sourceType == "BIT") {
                    if (value == "1" || value == "-1" || value.compare("TRUE",Qt::CaseInsensitive) == 0) converted = qlonglong(-1);
                    else if (value == "0" || value.compare("FALSE",Qt::CaseInsensitive) == 0) converted = qlonglong(0);
                    else throw fail("Ungültiger Boolean-Exportwert");
                } else if (columns[c].sqliteType == "BLOB") {
                    if(binaryEncoding!=AccessBinaryEncoding::MdbToolsHex)
                        throw fail("Nicht-NULL-Binärdaten benötigen einen ausdrücklich gewählten Hexexport");
                    converted=decodeMdbBinaryHex(cell.value);
                } else if (columns[c].sqliteType == "INTEGER") {
                    bool ok = false;
                    const auto number = value.toLongLong(&ok);
                    if (!ok) throw fail("Ungültiger INTEGER-Wert in Zeile " + QString::number(r));
                    converted = number;
                } else if(hasDoubleProof) {
                    const auto exact=(*doubleProof)[static_cast<qsizetype>(r-1)].value();
                    if(!mdbFloatingExportMatches(value,exact,16)) throw fail("Double-Rohwert und CSV-Snapshot widersprechen sich");
                    converted=exact;
                } else if(hasSingleProof) {
                    const auto exact=(*proof)[static_cast<qsizetype>(r-1)].value();
                    if(!mdbFloatingExportMatches(value,static_cast<double>(exact),8)) throw fail("Single-Rohwert und CSV-Snapshot widersprechen sich");
                    converted=static_cast<double>(exact);
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
        for(const auto& column:columns) if(sourceCounterStates.contains(column.name))
            restoreAccessCounter(db,tableName,column.name,sourceCounterStates.value(column.name),nativeCounter(column));
        if (!insertOnly) {
            for (const auto& index : schema.indexes) {
                QStringList fields;
                if(!index.descending.isEmpty() && index.descending.size()!=index.columns.size()) throw fail("Indexrichtungen unvollständig");
                for (qsizetype ordinal=0;ordinal<index.columns.size();++ordinal) {
                    const auto& name=index.columns[ordinal];
                    if (!seen.contains(name.toCaseFolded())) throw fail("Indexspalte fehlt");
                    fields << sqlQuote(name)+(index.descending.isEmpty() || !index.descending[ordinal] ? " ASC" : " DESC");
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
            QStringList displayColumns;for(const auto& column:columns) displayColumns.append(column.name);const auto sourceDisplay=checkedAccessSourceFieldDisplayProperties(schema.sourceIndexMetadata,displayColumns);
            for (const auto& col : columns) {QJsonObject entry{
                {"name", col.name}, {"source_type", col.sourceType}, {"sqlite_type", col.sqliteType},
                {"required", col.required}, {"primary_key_order", col.primaryKeyOrder},
                {"auto_number", col.autoNumber}, {"auto_number_native", col.autoNumber},
                {"source_origin", col.sourceOrigin}, {"source_size", col.sourceSize},
                {"has_default", col.hasDefault}, {"source_default", col.sourceDefault}, {"default_sql", col.sqlDefault},
                {"default_native", col.hasDefault && !col.autoNumber}};
                if(sourceDisplay) {entry.insert("source_display_properties_proven",true);const auto properties=sourceDisplay->value(col.name.toCaseFolded());if(properties.caption) {entry.insert("source_caption",*properties.caption);entry.insert("source_caption_id",QUuid::createUuid().toString(QUuid::WithoutBraces));}if(properties.description) {entry.insert("source_description",*properties.description);entry.insert("source_description_id",QUuid::createUuid().toString(QUuid::WithoutBraces));}}metadata.append(entry);}

            query.prepare("INSERT INTO _la_table_schema VALUES (?,?)");
            query.addBindValue(tableName);
            query.addBindValue(QString::fromUtf8(QJsonDocument(metadata).toJson(QJsonDocument::Compact)));
            if (!query.exec()) throw fail(query.lastError().text());
            if(!schema.sourceIndexMetadata.isEmpty()) {
                if(!query.exec("CREATE TABLE IF NOT EXISTS main._la_access_indexes(table_name TEXT PRIMARY KEY,definition TEXT NOT NULL)")) throw fail(query.lastError().text());
                query.prepare("INSERT INTO main._la_access_indexes VALUES(?,?)"); query.addBindValue(tableName); query.addBindValue(QString::fromUtf8(schema.sourceIndexMetadata));
                if(!query.exec()) throw fail(query.lastError().text());
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
                             "PK, exported indices, literal defaults and proven text capacities retained; explicit hex preserves raw binary bytes; FK retained by full-database import only; OLE rendering and general Access semantics NOT retained");
            if (!log.exec()) throw fail(log.lastError().text());
        }
        if(!counterDigest.isEmpty() && accessSourceDigest(floatingReadPath)!=counterDigest) throw fail("AutoWert-Quelle wurde während des Imports verändert");
        if(exactSingles && accessSourceDigest(floatingReadPath)!=exactSingles->sourceDigest) throw fail("Single-Quelle wurde während des Imports verändert");
        if(exactDoubles && accessSourceDigest(floatingReadPath)!=exactDoubles->sourceDigest) throw fail("Double-Quelle wurde während des Imports verändert");
        if (!scope.exec("RELEASE la_table_import")) throw fail(scope.lastError().text());
    } catch (...) {
        scope.exec("ROLLBACK TO la_table_import");
        scope.exec("RELEASE la_table_import");
        throw;
    }
}

void importAccessTable(QSqlDatabase db, const QString& sourceFile,
                       const QString& tableName, const AccessTableSchema& schema,
                       const std::vector<std::vector<CsvField>>& rows,AccessBinaryEncoding binaryEncoding,AccessFloatingSource floatingSource) {
    importAccessTablePhase(db, sourceFile, tableName, schema, rows,false,binaryEncoding,floatingSource);
}

void importAccessTable(QSqlDatabase db, const QString& sourceFile,
                       const QString& tableName, const QVector<AccessColumn>& columns,
                       const std::vector<std::vector<CsvField>>& rows,AccessBinaryEncoding binaryEncoding,AccessFloatingSource floatingSource) {
    importAccessTable(db, sourceFile, tableName, AccessTableSchema{columns, {}, {}}, rows,binaryEncoding,floatingSource);
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
            importAccessTablePhase(db, sourceFile, it.key(), it.value(), {header},false,AccessBinaryEncoding::MdbToolsHex,AccessFloatingSource::NativeMdbBits,readPath);
        }
        for (auto it = schemas.cbegin(); it != schemas.cend(); ++it) {
            const auto csv = AccessCatalog::exportTable(readPath,it.key(),true);
            importAccessTablePhase(db, sourceFile, it.key(), it.value(),
                parseCsvDetailed(std::string_view(csv.constData(), csv.size())), true,AccessBinaryEncoding::MdbToolsHex,AccessFloatingSource::NativeMdbBits,readPath);
        }
        if (!db.commit()) throw fail(db.lastError().text());
    } catch (...) { db.rollback(); throw; }
    return schemas.size();
}

}
