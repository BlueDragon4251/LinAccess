#include "access_catalog.hpp"
#include "access_query.hpp"
#include "dao_database.hpp"
#include "access_table_import.hpp"
#include "csv.hpp"
#include "vba_currency.hpp"
#include <QRegularExpression>

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
#include <QJsonObject>
#include <set>

class EnumerationHost final : public linaccess::VbaHost {
public:
    explicit EnumerationHost(QSqlDatabase db) : db_(db) {}
    std::shared_ptr<linaccess::VbaObject> currentDatabase() override { return linaccess::makeDaoDatabase(db_); }
    std::optional<std::string> field(std::string_view) override { throw std::runtime_error("Unexpected control read"); }
    void setField(std::string_view,std::optional<std::string>) override { assert(false); }
    void message(std::string_view) override { assert(false); }
    void saveRecord() override { assert(false); }
private:
    QSqlDatabase db_;
};

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
    int longColumns=0,typedLongValues=0;
    int currencyColumns=0,doubleColumns=0,typedCurrencyValues=0,currencySums=0,unsupportedCurrencySums=0;
    int formattedDates = 0;
    int counters=0,nativeCounters=0,counterProbes=0,managedProbes=0;
    QStringList failures;
    for (const auto& table : tables) {
        if (table.startsWith("MSys", Qt::CaseInsensitive)) continue;
        try {
            const auto csv = linaccess::AccessCatalog::exportTable(source,table);
            const auto rows = linaccess::parseCsvDetailed(std::string_view(csv.constData(), csv.size()));
            const auto schema = linaccess::parseMdbSchema(catalog.tableSchema(table));
            for (const auto& column : schema.columns)
                if (column.sourceType == "BOOLEAN" || column.sourceType == "BIT") ++booleanColumns;
            for (const auto& column:schema.columns) {
                if(column.sourceType=="CURRENCY") ++currencyColumns;
                if(column.sourceType=="DOUBLE") ++doubleColumns;
            }
            linaccess::importAccessTable(db, source, table, schema, rows);
            {
                std::vector<int> indices;
                for(int i=0;i<schema.columns.size();++i) if(schema.columns[i].sourceType=="CURRENCY") indices.push_back(i);
                if(!indices.empty()) {
                    const auto money=linaccess::makeDaoDatabase(db)->member("OpenRecordset",{linaccess::VbaValue::string(table.toStdString())}).object;
                    std::vector<std::shared_ptr<linaccess::VbaObject>> fields;
                    for(const auto i:indices) {
                        const auto field=money->member("Fields",{linaccess::VbaValue::string(schema.columns[i].name.toStdString())}).object;
                        assert(field->member("Type",{}).number==5); fields.push_back(field);
                    }
                    QMap<QByteArray,int> actual,reference;
                    while(!money->member("EOF",{}).truthy()) {
                        QJsonArray values;
                        for(const auto& field:fields) {
                            const auto value=field->member("Value",{});
                            if(value.kind==linaccess::VbaValue::Kind::Null) values.append(QJsonValue::Null);
                            else { assert(value.kind==linaccess::VbaValue::Kind::Currency); values.append(QString::number(value.currencyUnits)); ++typedCurrencyValues; }
                        }
                        ++actual[QJsonDocument(values).toJson(QJsonDocument::Compact)]; money->member("MoveNext",{});
                    }
                    money->member("Close",{});
                    const QRegularExpression decimal("^[+-]?[0-9]+(?:\\.[0-9]{1,4})?$");
                    for(std::size_t row=1;row<rows.size();++row) {
                        QJsonArray values;
                        for(const auto i:indices) {
                            const auto& cell=rows[row][i];
                            if(!cell.quoted && cell.value=="\x1e") values.append(QJsonValue::Null);
                            else {
                                assert(decimal.match(QString::fromStdString(cell.value)).hasMatch());
                                values.append(QString::number(linaccess::convertCurrency(linaccess::VbaValue::string(cell.value)).currencyUnits));
                            }
                        }
                        ++reference[QJsonDocument(values).toJson(QJsonDocument::Compact)];
                    }
                    assert(actual==reference);
                    for(const auto i:indices) {
                        __int128 expected=0; bool any=false;
                        for(std::size_t row=1;row<rows.size();++row) {
                            const auto& cell=rows[row][i];
                            if(!cell.quoted && cell.value=="\x1e") continue;
                            expected+=linaccess::convertCurrency(linaccess::VbaValue::string(cell.value)).currencyUnits; any=true;
                        }
                        const auto quote=[](QString value) { return '['+value.replace(']',"]]")+']'; };
                        const auto field=quote(schema.columns[i].name);
                        const auto plan=linaccess::compileAccessQuery("SELECT Sum(CCur("+field+")) AS total FROM "+quote(table)+" WHERE "+field+" IS NOT NULL");
                        if(expected<=-5497558138880000LL || expected>=5497558138880000LL) {
                            int error=0; try { (void)linaccess::executeAccessQueryPlan(db,plan,{}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
                            assert(error==3251 || error==6); ++unsupportedCurrencySums;
                        } else {
                            const auto result=linaccess::executeAccessQueryPlan(db,plan,{});
                            assert(result.rows.size()==1 && result.types.value("total")=="CURRENCY_REAL");
                            if(!any) assert(result.rows[0][0].isNull());
                            else assert(linaccess::currencyFromRealStorage(result.rows[0][0].toDouble()).currencyUnits==static_cast<std::int64_t>(expected));
                            ++currencySums;
                        }
                    }

                }
            }
            int primaryCount=0;
            for (const auto& column:schema.columns) if (column.primaryKeyOrder) ++primaryCount;
            for (const auto& column:schema.columns) if (column.autoNumber) {
                ++counters;
                const bool native=primaryCount==1 && column.primaryKeyOrder==1;
                if (native) ++nativeCounters;
                QSqlQuery metadata(db); metadata.prepare("SELECT definition FROM _la_table_schema WHERE table_name=?"); metadata.addBindValue(table);
                assert(metadata.exec() && metadata.next());
                const auto stored=QJsonDocument::fromJson(metadata.value(0).toByteArray()).array(); metadata.finish();
                int matched=0;
                for (const auto& value:stored) {
                    const auto field=value.toObject(); if (field["name"].toString()!=column.name) continue;
                    assert(field["auto_number"].toBool() && field["auto_number_native"].toBool()); ++matched;
                }
                assert(matched==1);
                if (native ? counterProbes!=0 : managedProbes!=0) continue;
                const auto quote=[](QString value) { return '"'+value.replace('"',"\"\"")+'"'; };
                QSqlQuery probe(db); assert(probe.exec("SAVEPOINT counter_probe"));
                assert(probe.exec("SELECT COALESCE(MAX("+quote(column.name)+"),0) FROM "+quote(table)) && probe.next());
                const auto previous=probe.value(0).toLongLong(); probe.finish();
                if (probe.exec("INSERT INTO "+quote(table)+" DEFAULT VALUES")) {
                    const auto counterValue=[&](const QVariant& rowId) {
                        QSqlQuery read(db); read.prepare("SELECT "+quote(column.name)+" FROM "+quote(table)+" WHERE rowid=?"); read.addBindValue(rowId);
                        assert(read.exec() && read.next()); return read.value(0).toLongLong();
                    };
                    const auto first=counterValue(probe.lastInsertId()); assert(first>previous);
                    probe.prepare("DELETE FROM "+quote(table)+" WHERE "+quote(column.name)+"=?"); probe.addBindValue(first); assert(probe.exec());
                    assert(probe.exec("INSERT INTO "+quote(table)+" DEFAULT VALUES"));
                    assert(counterValue(probe.lastInsertId())==first+1); if (native) ++counterProbes; else ++managedProbes;
                }
                assert(probe.exec("ROLLBACK TO counter_probe") && probe.exec("RELEASE counter_probe"));
            }
            EnumerationHost enumerationHost(db);
            std::vector<linaccess::VbaValue> arguments{linaccess::VbaValue::string(table.toUtf8().toStdString())};
            const auto enumeration = linaccess::executeVbaProcedureWithArguments(R"VBA(
Function CountFields(ByVal source As String) As Long
Dim rs As DAO.Recordset
Dim fld As DAO.Field
Dim count As Long
Set rs = CurrentDb.OpenRecordset(source, dbOpenSnapshot)
For Each fld In rs.Fields
count = count + 1
Next fld
rs.Close
CountFields = count
End Function
)VBA","CountFields",enumerationHost,arguments);
            assert(enumeration.status == linaccess::VbaExecutionStatus::Completed);
            assert(enumeration.typedReturnValue && enumeration.typedReturnValue->asNumber() == schema.columns.size());

            {
                const auto numericRows=linaccess::makeDaoDatabase(db)->member("OpenRecordset",{linaccess::VbaValue::string(table.toStdString()),linaccess::VbaValue::numeric(4)}).object;
                std::vector<std::shared_ptr<linaccess::VbaObject>> fields;
                QStringList projections; QMap<QByteArray,int> actualIntegers,referenceIntegers;
                const auto quote=[](QString value) { return '"'+value.replace('"',"\"\"")+'"'; };
                for(const auto& column:schema.columns) if(column.sourceType=="INT") {
                    const auto field=numericRows->member("Fields",{linaccess::VbaValue::string(column.name.toStdString())}).object;
                    assert(field->member("Type",{}).number==4 && field->member("Type",{}).numericType==linaccess::VbaValue::NumericType::Integer);
                    fields.push_back(field); projections<<quote(column.name); ++longColumns;
                }
                while(!numericRows->member("EOF",{}).truthy()) {
                    QJsonArray actual;
                    for(const auto& field:fields) {
                        const auto value=field->member("Value",{});
                        if(value.kind==linaccess::VbaValue::Kind::Null) { actual.append(QJsonValue::Null); continue; }
                        assert(value.kind==linaccess::VbaValue::Kind::Number && value.numericType==linaccess::VbaValue::NumericType::Long);
                        assert(value.number>=-2147483648.0 && value.number<=2147483647.0); ++typedLongValues;
                        actual.append(QString::number(static_cast<qlonglong>(value.number)));
                    }
                    ++actualIntegers[QJsonDocument(actual).toJson(QJsonDocument::Compact)];
                    numericRows->member("MoveNext",{});
                }
                QSqlQuery reference(db); assert(reference.exec("SELECT "+(projections.isEmpty() ? QString("1") : projections.join(','))+" FROM "+quote(table)));
                while(reference.next()) {
                    QJsonArray expected;
                    for(int i=0;i<projections.size();++i) {
                        if(reference.value(i).isNull()) expected.append(QJsonValue::Null);
                        else expected.append(QString::number(reference.value(i).toLongLong()));
                    }
                    ++referenceIntegers[QJsonDocument(expected).toJson(QJsonDocument::Compact)];
                }
                reference.finish(); assert(actualIntegers==referenceIntegers);
                numericRows->member("Close",{});
            }
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

            for (const auto& column : schema.columns) {
                if (column.sourceType!="DATETIME" && column.sourceType!="DATE") continue;
                const auto bracket=[](QString name) { return '['+name.replace(']',"]]")+']'; };
                linaccess::AccessQueryModel formatted;
                const auto field=bracket(column.name);
                formatted.execute(db,"SELECT TOP 1 "+field+",Format$("+field+",\"yyyy\"),Format$("+field+",\"mm\"),Format$("+field+",\"dd\") FROM "+bracket(table)+" WHERE "+field+" IS NOT NULL");
                if (formatted.rowCount()==0) continue;
                const auto originalDate=formatted.data(formatted.index(0,0)).toString();
                assert(originalDate.size()>=10);
                assert(formatted.data(formatted.index(0,1)).toString()==originalDate.left(4));
                assert(formatted.data(formatted.index(0,2)).toString()==originalDate.mid(5,2));
                assert(formatted.data(formatted.index(0,3)).toString()==originalDate.mid(8,2));
                ++formattedDates;
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
    assert(formattedDates>0);
    std::cout << "Original date columns formatted through native SQL: " << formattedDates << '\n';
    assert(tables.size() == 29);
    assert(imported == 29 && failures.empty());
    assert(currencyColumns==43 && doubleColumns==57 && typedCurrencyValues>0);
    assert(currencySums+unsupportedCurrencySums==43);
    std::cout<<"Original Currency sums: "<<currencySums<<", outside supported backing: "<<unsupportedCurrencySums<<'\n';
    std::cout<<"Original Currency column metadata: "<<currencyColumns<<", Double columns: "<<doubleColumns<<", typed Currency non-NULL values: "<<typedCurrencyValues<<'\n';
    assert(longColumns==127 && typedLongValues>0);
    std::cout<<"Original Long columns: "<<longColumns<<", typed non-NULL field values: "<<typedLongValues<<'\n';
    assert(counters==21 && nativeCounters==16 && counterProbes==1 && managedProbes==1);
    std::cout<<"Original counter metadata: "<<counters<<", native single-PK counters: "<<nativeCounters<<", isolated insert/delete/no-reuse probes: "<<counterProbes<<'\n';
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
