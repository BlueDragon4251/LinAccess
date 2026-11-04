#include "access_table_import.hpp"
#include "access_catalog.hpp"
#include "access_autonumber.hpp"
#include "jet_sql_functions.hpp"
#include "dao_database.hpp"
#include <QCoreApplication>
#include <QFile>
#include <QTemporaryDir>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QJsonDocument>
#include <QJsonArray>
#include <QJsonObject>
#include <QSqlError>
#include <atomic>
#include <thread>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>

int main(int argc, char** argv) {
    QCoreApplication app(argc, argv);
    using namespace linaccess;
    const auto floating=AccessCatalog::floatingSourceTypes("CREATE TABLE [values]\n(\n[price] Currency,\n[ö] Single,\n[odd] Currency] Single NOT NULL,\n[other] Double,\n-- [comment] Currency\n);\n");
    assert(floating.size()==4 && floating.value("other")=="DOUBLE" && floating.value("price")=="CURRENCY" && floating.value("ö")=="SINGLE" && floating.value("odd] Currency")=="SINGLE");
    assert(AccessCatalog::floatingSourceTypes("[label] SingleText\n[price] CurrencyOther\n[other] DoubleOther").isEmpty());
    bool conflict=false; try { (void)AccessCatalog::floatingSourceTypes("[x] Currency\n[x] Single"); } catch(const std::runtime_error&) { conflict=true; } assert(conflict);
    const auto currencySchema=parseMdbSchema("CREATE TABLE `cash`\n(\n`price` REAL,\n`rate` REAL\n);\n-- LINACCESS_SOURCE_TYPE 7072696365 FLOAT\n-- LINACCESS_SOURCE_TYPE 7072696365 CURRENCY\n-- LINACCESS_SOURCE_TYPE 72617465 SINGLE\n");
    assert(currencySchema.columns.size()==2 && currencySchema.columns[0].sourceType=="CURRENCY" && currencySchema.columns[1].sourceType=="SINGLE");
    assert(currencySchema.columns[0].sqliteType=="REAL" && currencySchema.columns[1].sqliteType=="REAL");
    const auto schema = parseMdbSchema(R"SQL(CREATE TABLE `source`
 (
 `part` INTEGER NOT NULL,
 `line` INTEGER,
 `a``b` varchar,
 `a"b` REAL
 , PRIMARY KEY (`part`, `line`)
 );
CREATE UNIQUE INDEX `pair` ON `source` (`part`, `line`);
CREATE INDEX `label` ON `source` (`a``b`);
)SQL");
    assert(schema.columns.size() == 4 && schema.indexes.size() == 2);
    assert(schema.columns[0].primaryKeyOrder == 1);
    assert(schema.columns[1].primaryKeyOrder == 2);
    assert(schema.columns[2].name == "a`b");
    assert(schema.columns[3].name == "a\"b");
    assert(parseMdbSchema("CREATE TABLE `x`\n(\n`a` UNSUPPORTED\n);").columns.empty());
    auto db = QSqlDatabase::addDatabase("QSQLITE");
    db.setDatabaseName(":memory:");
    assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    const auto rows = parseCsvDetailed("part,line,a`b,a\"\"b\n1,2,\"\x1e\",3.5\n");
    // Build the header explicitly: CSV double quotes are quoted identifiers' data.
    auto data = rows;
    data[0][3].value = "a\"b";
    importAccessTable(db, "not-a-real-file", "copy\"table", schema, data);
    assert(query.exec("SELECT \"a`b\",\"a\"\"b\" FROM \"copy\"\"table\""));
    assert(query.next() && query.value(0).toString() == QString(QChar(0x1e)));
    assert(query.value(1).toDouble() == 3.5);
    assert(!query.exec("INSERT INTO \"copy\"\"table\" (part,line) VALUES (1,2)"));
    assert(!query.exec("INSERT INTO \"copy\"\"table\" (part,line) VALUES (NULL,3)"));
    auto bad = data;
    bad.push_back(data[1]); // duplicate composite primary key must roll back DDL and log
    bool refused = false;
    try { importAccessTable(db, "not-a-real-file", "failed", schema, bad); }
    catch (...) { refused = true; }
    assert(refused && !db.tables().contains("failed"));
    assert(query.exec("SELECT COUNT(*) FROM _la_import_log"));
    assert(query.next() && query.value(0).toInt() == 1);
    query.finish();
    importAccessTable(db,"synthetic-floating-types","floating_types",currencySchema,parseCsvDetailed("price,rate\n1.2345,0.25\n"));
    assert(query.exec("SELECT definition FROM _la_table_schema WHERE table_name='floating_types'") && query.next());
    const auto floatingMetadata=QJsonDocument::fromJson(query.value(0).toByteArray()).array(); query.finish();
    assert(floatingMetadata.size()==2 && floatingMetadata[0].toObject()["source_type"].toString()=="CURRENCY" && floatingMetadata[1].toObject()["source_type"].toString()=="SINGLE");

    const QString textDefinition="CREATE TABLE `source`\n(\n`label` TEXT,\n`note` TEXT\n);\n"
        "-- LINACCESS_SOURCE_TYPE 6c6162656c VARCHAR\n-- LINACCESS_SOURCE_ORIGIN 6c6162656c MDBTOOLS_MYSQL\n-- LINACCESS_SOURCE_SIZE 6c6162656c 20\n"
        "-- LINACCESS_SOURCE_TYPE 6e6f7465 TEXT\n-- LINACCESS_SOURCE_ORIGIN 6e6f7465 MDBTOOLS_MYSQL\n";
    const auto textSchema=parseMdbSchema(textDefinition);
    assert(textSchema.columns.size()==2 && textSchema.columns[0].sourceSize==20 && textSchema.columns[1].sourceSize==0);
    assert(textSchema.columns[0].sourceOrigin=="MDBTOOLS_MYSQL");
    assert(parseMdbSchema(textDefinition+"-- LINACCESS_SOURCE_SIZE 6c6162656c 256\n").columns.isEmpty());
    assert(parseMdbSchema(textDefinition+"-- LINACCESS_SOURCE_SIZE 6c6162656c 19\n").columns.isEmpty());
    importAccessTable(db,"synthetic-source-widths","text_metadata",textSchema,parseCsvDetailed("label,note\nshort,longer text\n"));
    assert(query.exec("SELECT definition FROM _la_table_schema WHERE table_name='text_metadata'") && query.next());
    const auto textMetadata=QJsonDocument::fromJson(query.value(0).toByteArray()).array(); query.finish();
    assert(textMetadata[0].toObject()["source_size"].toInt()==20 && textMetadata[1].toObject()["source_origin"].toString()=="MDBTOOLS_MYSQL");
    const auto textFields=makeDaoDatabase(db)->member("TableDefs",{VbaValue::string("text_metadata")}).object->member("Fields",{}).object;
    const auto label=textFields->member("Item",{VbaValue::string("label")}).object;
    assert(label->member("Type",{}).number==10 && label->member("Size",{}).number==20);
    const auto note=textFields->member("Item",{VbaValue::string("note")}).object;
    assert(note->member("Type",{}).number==12 && note->member("Size",{}).number==0);
    {
        const auto records=makeDaoDatabase(db)->member("OpenRecordset",{VbaValue::string("text_metadata")}).object;
        const auto field=records->member("Fields",{VbaValue::string("label")}).object;
        assert(field->member("Type",{}).number==10 && field->member("Size",{}).number==20);
        assert(field->member("Value",{}).asText()=="short");
        const auto memo=records->member("Fields",{VbaValue::string("note")}).object;
        assert(memo->member("Type",{}).number==12 && memo->member("Size",{}).number==0);
        const auto clone=records->member("Clone",{}).object;
        const auto cloned=clone->member("Fields",{VbaValue::string("label")}).object;
        assert(cloned->member("Type",{}).number==10 && cloned->member("Size",{}).number==20);
        clone->member("Close",{}); records->member("Close",{});
        bool closed=false; try { field->member("Size",{}); } catch(const VbaRuntimeError& error) { closed=error.number==3420; } assert(closed);
    }
    const auto dragon=QString::fromUtf8("🐉");
    query.prepare("INSERT INTO text_metadata(label) VALUES(?)"); query.addBindValue(QString(18,'x')+dragon); assert(query.exec()); query.finish();
    query.prepare("INSERT INTO text_metadata(label) VALUES(?)"); query.addBindValue(QString(19,'x')+dragon); assert(!query.exec());
    assert(jetSqlFunctionErrorNumber(query.lastError().text().toLower().toStdString())==3163); query.finish();
    query.prepare("INSERT INTO text_metadata(label) VALUES(?)"); query.addBindValue(QString(19,'x')+QChar(0)+'x'); assert(!query.exec());
    assert(jetSqlFunctionErrorNumber(query.lastError().text().toLower().toStdString())==3163); query.finish();
    bool overlong=false;
    try { makeDaoDatabase(db)->member("Execute",{VbaValue::string("INSERT INTO text_metadata(label) VALUES('123456789012345678901')")}); }
    catch(const VbaRuntimeError& error) { overlong=error.number==3163; } assert(overlong);
    assert(query.exec("SELECT COUNT(*) FROM text_metadata") && query.next() && query.value(0).toInt()==2); query.finish();


    const auto counter=parseMdbSchema("CREATE TABLE `x`\n(\n`id` int NOT NULL auto_increment,\n`label` varchar\n);\nALTER TABLE `x` ADD PRIMARY KEY (`id`);\n");
    assert(counter.columns.size()==2 && counter.columns[0].autoNumber && !counter.columns[1].autoNumber);
    assert(counter.columns[0].primaryKeyOrder==1);
    const auto sourceCounter=parseMdbSchema("CREATE TABLE `x`\n(\n`id` INTEGER NOT NULL\n);\n-- LINACCESS_AUTO_NUMBER 6964\n");
    assert(sourceCounter.columns.size()==1 && sourceCounter.columns[0].autoNumber && !sourceCounter.columns[0].primaryKeyOrder);
    assert(parseMdbSchema("CREATE TABLE `x`\n(\n`id` TEXT auto_increment\n);\n").columns.empty());
    assert(parseMdbSchema("CREATE TABLE `x`\n(\n`id` INTEGER\n);\n-- LINACCESS_AUTO_NUMBER 616273656e74\n").columns.empty());
    const auto quotedDefault=parseMdbSchema("CREATE TABLE `x`\n(\n`id` INTEGER DEFAULT 'auto_increment'\n);\n");
    assert(quotedDefault.columns.size()==1 && !quotedDefault.columns[0].autoNumber);
    importAccessTable(db,"synthetic","counter",counter,parseCsvDetailed("id,label\n7,first\n41,last\n"));
    assert(query.exec("SELECT id,label FROM counter ORDER BY id") && query.next() && query.value(0).toInt()==7);
    assert(query.next() && query.value(0).toInt()==41 && query.value(1).toString()=="last"); query.finish();
    assert(query.exec("INSERT INTO counter(label) VALUES('generated')") && query.lastInsertId().toLongLong()==42);
    assert(query.exec("DELETE FROM counter WHERE id=42"));
    assert(query.exec("INSERT INTO counter(label) VALUES('no reuse')") && query.lastInsertId().toLongLong()==43);
    assert(query.exec("INSERT INTO counter VALUES(100,'explicit')"));
    assert(query.exec("INSERT INTO counter(label) VALUES('after explicit')") && query.lastInsertId().toLongLong()==101);
    assert(query.exec("SELECT definition FROM _la_table_schema WHERE table_name='counter'") && query.next());
    const auto metadata=QJsonDocument::fromJson(query.value(0).toByteArray()).array(); query.finish();
    assert(metadata.size()==2 && metadata[0].toObject()["auto_number"].toBool() && metadata[0].toObject()["auto_number_native"].toBool());
    assert(!metadata[1].toObject()["auto_number"].toBool());
    importAccessTable(db,"synthetic","counter_without_pk",sourceCounter,parseCsvDetailed("id\n12\n"));
    assert(query.exec("SELECT definition FROM _la_table_schema WHERE table_name='counter_without_pk'") && query.next());
    const auto preserved=QJsonDocument::fromJson(query.value(0).toByteArray()).array(); query.finish();
    assert(preserved[0].toObject()["auto_number"].toBool() && preserved[0].toObject()["auto_number_native"].toBool());
    assert(query.exec("INSERT INTO counter_without_pk DEFAULT VALUES"));
    assert(query.exec("SELECT max(id) FROM counter_without_pk") && query.next() && query.value(0).toInt()==13); query.finish();
    QTemporaryDir persistent; assert(persistent.isValid());
    const auto path=persistent.filePath("counter.linaccess");
    auto file=QSqlDatabase::addDatabase("QSQLITE","persistent-counter"); file.setDatabaseName(path); assert(file.open());
    QSqlQuery write(file); assert(write.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    const auto managed=parseMdbSchema("CREATE TABLE `x`\n(\n`id` INTEGER NOT NULL,\n`label` TEXT\n);\n-- LINACCESS_AUTO_NUMBER 6964\nCREATE UNIQUE INDEX `label` ON `x` (`label`);\n");
    importAccessTable(file,"synthetic","persistent",managed,parseCsvDetailed("id,label\n12,seed\n"));
    const auto idFor=[&](const QString& label) {
        QSqlQuery read(file); read.prepare("SELECT id FROM persistent WHERE label=?"); read.addBindValue(label); assert(read.exec() && read.next()); return read.value(0).toLongLong();
    };
    assert(!write.exec("INSERT INTO persistent(label)VALUES('seed')"));
    assert(write.exec("INSERT INTO persistent(label)VALUES('next')") && idFor("next")==13);
    assert(write.exec("INSERT INTO persistent VALUES(70,'explicit')"));
    assert(write.exec("UPDATE persistent SET id=90 WHERE label='explicit'"));
    assert(write.exec("INSERT INTO persistent(label)VALUES('highest')") && idFor("highest")==91);
    assert(write.exec("DELETE FROM persistent WHERE id=91"));
    assert(write.exec("INSERT INTO persistent(label)VALUES('no reuse')") && idFor("no reuse")==92);
    write.clear(); file.close(); assert(file.open()); installJetSqlFunctions(file); write=QSqlQuery(file);
    assert(write.exec("INSERT INTO persistent(label)VALUES('reopen')") && idFor("reopen")==93);
    auto database=makeDaoDatabase(file); auto recordset=database->member("OpenRecordset",{VbaValue::string("persistent")}).object;
    recordset->member("AddNew",{}); recordset->assignItem("label",VbaValue::string("dao")); recordset->member("Update",{}); assert(recordset->item("id").asNumber()==94);
    recordset->member("AddNew",{}); recordset->assignItem("id",VbaValue{}); recordset->assignItem("label",VbaValue::string("dao null")); recordset->member("Update",{}); assert(recordset->item("id").asNumber()==95);
    recordset->member("AddNew",{}); recordset->assignItem("label",VbaValue::string("seed"));
    int duplicate=0; try { recordset->member("Update",{}); } catch (const VbaRuntimeError& error) { duplicate=error.number; }
    assert(duplicate==3022); recordset->member("Close",{});
    std::atomic_bool good{true};
    const auto worker=[&](int number) {
        const auto connection="counter-worker-"+QString::number(number);
        {
            auto other=QSqlDatabase::addDatabase("QSQLITE",connection); other.setDatabaseName(path);
            if (!other.open()) good=false;
            else {
                try {
                    installJetSqlFunctions(other); QSqlQuery insert(other);
                    for (int i=0;i<20;++i) {
                        insert.prepare("INSERT INTO persistent(label)VALUES(?)"); insert.addBindValue("worker-"+QString::number(number)+'-'+QString::number(i));
                        if (!insert.exec()) { good=false; break; }
                    }
                } catch (...) { good=false; }
            }
        }
        QSqlDatabase::removeDatabase(connection);
    };
    std::thread first(worker,1),second(worker,2); first.join(); second.join(); assert(good);
    assert(write.exec("SELECT count(*),count(DISTINCT id),max(id) FROM persistent WHERE id>95") && write.next());
    assert(write.value(0).toInt()==40 && write.value(1).toInt()==40 && write.value(2).toInt()==135); write.finish();
    assert(!write.exec("SELECT linaccess_next_counter(NULL,'id')"));
    assert(jetSqlFunctionErrorNumber(write.lastError().text().toCaseFolded().toStdString())==94);
    assert(!write.exec("SELECT linaccess_next_counter('missing','id')"));
    assert(jetSqlFunctionErrorNumber(write.lastError().text().toCaseFolded().toStdString())==3251);
    assert(write.exec("UPDATE _la_counter_sequences SET value=9223372036854775807 WHERE table_name='persistent'"));
    assert(!write.exec("INSERT INTO persistent(label)VALUES('overflow')"));
    assert(jetSqlFunctionErrorNumber(write.lastError().text().toCaseFolded().toStdString())==6);
    auto duplicateRows=parseCsvDetailed("id,label\n1,same\n2,same\n"); bool atomic=false;
    try { importAccessTable(file,"synthetic","failed_counter",managed,duplicateRows); } catch (...) { atomic=true; }
    assert(atomic && !file.tables().contains("failed_counter"));
    assert(write.exec("SELECT count(*) FROM _la_counter_sequences WHERE table_name='failed_counter'") && write.next() && write.value(0).toInt()==0); write.finish();
    auto quotedCounter=sourceCounter; quotedCounter.columns[0].name="i'd\" ]";
    const QString quotedTable="counter'\" table";
    const std::vector<std::vector<CsvField>> quotedRows{{{quotedCounter.columns[0].name.toStdString(),true}},{{"12",false}}};
    importAccessTable(file,"synthetic",quotedTable,quotedCounter,quotedRows);
    assert(generatedAccessColumns(file,quotedTable).contains(quotedCounter.columns[0].name.toCaseFolded()));
    assert(write.exec("INSERT INTO \"counter'\"\" table\" DEFAULT VALUES"));
    assert(write.exec("SELECT MAX(\"i'd\"\" ]\") FROM \"counter'\"\" table\"") && write.next() && write.value(0).toInt()==13); write.finish();
    QTemporaryDir tools;
    assert(tools.isValid());
    QFile fake(tools.filePath("mdb-failure"));
    assert(fake.open(QIODevice::WriteOnly));
    assert(fake.write("#!/bin/sh\necho 'offset is beyond EOF' >&2\necho partial\nexit 0\n") > 0);
    fake.close();
    assert(fake.setPermissions(QFile::ReadOwner | QFile::WriteOwner | QFile::ExeOwner));
    bool rejectedPartial = false;
    try { (void)AccessCatalog::tool(fake.fileName(), {}); } catch (...) { rejectedPartial = true; }
    assert(rejectedPartial);
    std::cout << "Schema, composite keys, indices, quoted identifiers and rollback verified\n";
}
