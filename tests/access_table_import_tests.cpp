#include "access_table_import.hpp"
#include "access_catalog.hpp"
#include <QCoreApplication>
#include <QFile>
#include <QTemporaryDir>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QJsonDocument>
#include <QJsonArray>
#include <QJsonObject>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>

int main(int argc, char** argv) {
    QCoreApplication app(argc, argv);
    using namespace linaccess;
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
    assert(preserved[0].toObject()["auto_number"].toBool() && !preserved[0].toObject()["auto_number_native"].toBool());
    assert(!query.exec("INSERT INTO counter_without_pk DEFAULT VALUES")); // Preserved metadata, not yet generated.
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
