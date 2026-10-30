#include "access_table_import.hpp"
#include "access_catalog.hpp"
#include <QCoreApplication>
#include <QFile>
#include <QTemporaryDir>
#include <QSqlDatabase>
#include <QSqlQuery>
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
