#include "access_catalog.hpp"
#include "access_index_metadata.hpp"
#include "access_table_import.hpp"
#include "csv.hpp"
#include "dao_database.hpp"

#include <QCoreApplication>
#include <QCryptographicHash>
#include <QFile>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QTemporaryDir>

#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>

namespace {
QByteArray digest(const QString& path) {
    QFile file(path);
    assert(file.open(QIODevice::ReadOnly));
    return QCryptographicHash::hash(file.readAll(), QCryptographicHash::Sha256);
}
int scalar(QSqlDatabase db, const QString& sql) {
    QSqlQuery query(db);
    assert(query.exec(sql));
    assert(query.next());
    return query.value(0).toInt();
}

void verifyNumericDisplay(const QString& path) {
    using namespace linaccess;
    const auto before=digest(path);
    const AccessCatalog catalog(path);
    const bool ace=catalog.formatVersion()=="ACE14";
    assert(ace || catalog.formatVersion()=="JET4");
    const QString table="decimal_probe";
    const auto schema=parseMdbSchema(catalog.tableSchema(table));
    const QStringList columns{"id","large","small"};
    const auto display=checkedAccessSourceFieldDisplayProperties(schema.sourceIndexMetadata,columns);
    assert(display);
    const auto numeric=display->value("large");
    assert(numeric.caption==QString("Exact decimal"));
    assert(numeric.description==QString("Prüfwert für Dezimalstellen"));
    assert(numeric.format==QString("0.0000;[Red]-0.0000") && numeric.formatProven);
    if(ace) {
        assert(display->size()==1 && !display->contains("id") && !display->contains("small"));
    } else {
        assert(display->size()==3 && display->value("id").caption==QString("Fixture identifier"));
        assert(display->value("small").formatProven && !display->value("small").format);
    }

    auto db=QSqlDatabase::addDatabase("QSQLITE","numeric-display-probe");
    db.setDatabaseName(":memory:");
    assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    assert(importAccessDatabase(db,path)==1);
    assert(scalar(db,"SELECT COUNT(*) FROM decimal_probe")==5);
    const auto fields=makeDaoDatabase(db)->member("TableDefs",{VbaValue::string("decimal_probe")}).object;
    const auto property=[&](const char* column,const char* name) {
        const auto field=fields->member("Fields",{VbaValue::string(column)}).object;
        return field->member("Properties",{VbaValue::string(name)}).object->member("Value",{}).asText();
    };
    assert(property("large","Caption")=="Exact decimal");
    assert(property("large","Description")=="Prüfwert für Dezimalstellen");
    assert(property("large","Format")=="0.0000;[Red]-0.0000");
    if(ace) {
        for(const char* column:{"id","small"}) {
            int error=0;
            try { (void)property(column,"Caption"); }
            catch(const VbaRuntimeError& failure) { error=failure.number; }
            assert(error==3251);
        }

        const auto root=QJsonDocument::fromJson(schema.sourceIndexMetadata).object();
        const auto csv=AccessCatalog::exportTable(path,table);
        const auto rows=parseCsvDetailed(std::string_view(csv.constData(),csv.size()));
        const auto rejects=[&](const QJsonObject& modified) {
            auto candidate=schema;
            candidate.sourceIndexMetadata=QJsonDocument(modified).toJson(QJsonDocument::Compact);
            bool refused=false;
            try { importAccessTable(db,path,"invalid_display",candidate,rows,
                                    AccessBinaryEncoding::Unsupported,AccessFloatingSource::NativeMdbBits); }
            catch(const std::exception&) { refused=true; }
            assert(refused && !db.tables().contains("invalid_display"));
        };
        auto modified=root;
        auto properties=root.value("field_properties").toArray();
        auto field=properties[0].toObject();
        field.insert("caption","Unproven Long caption");
        properties[0]=field;
        modified.insert("field_properties",properties);
        rejects(modified);

        properties=root.value("field_properties").toArray();
        field=properties[1].toObject();
        field.insert("format","0.00");
        properties[1]=field;
        modified.insert("field_properties",properties);
        rejects(modified);

        modified=root;
        modified.insert("source_sha256",QString(64,QChar('0')));
        rejects(modified);
    } else {
        assert(property("id","Caption")=="Fixture identifier");
    }
    assert(digest(path)==before);
    std::cout<<"Numeric display properties verified: "<<(ace ? "ACE 2010" : "Jet 4")<<'\n';
}
}

int main(int argc, char** argv) {
    QCoreApplication app(argc, argv);
    if(argc==3 && app.arguments()[1]=="--numeric-display") {
        verifyNumericDisplay(app.arguments()[2]);
        return 0;
    }
    assert(argc == 2);
    const QString path = QString::fromLocal8Bit(argv[1]);
    const auto before = digest(path);
    linaccess::AccessCatalog source(path);
    assert(source.formatVersion() == "JET4");

    const auto objects = source.objects();
    int tables = 0, forms = 0, reports = 0, queries = 0, relationships = 0;
    for (const auto& object : objects) {
        if (object.internal) continue;
        if (object.kind == "table") ++tables;
        if (object.kind == "form") ++forms;
        if (object.kind == "report") ++reports;
        if (object.kind == "query") ++queries;
        if (object.kind == "relationship") ++relationships;
    }
    assert(tables == 2 && forms == 4 && reports == 1 && queries == 3 && relationships == 1);
    assert(source.querySql("3_qry_Rezepte").contains("2_tbl_Nährwerte"));
    assert(source.missingQuerySources(source.querySql("3_qry_Rezepte"))
                   .contains("2_tbl_Nährwerte"));
    const auto relations = source.relationships();
    assert(relations.size() == 1);
    assert(relations[0].childTable == "33_tbl_ZutatenRezept");
    assert(relations[0].parentTable == "3_tbl_Rezepte");

    QTemporaryDir tmp;
    assert(tmp.isValid());
    auto db = QSqlDatabase::addDatabase("QSQLITE", "real-mdb-test");
    db.setDatabaseName(tmp.filePath("result.linaccess"));
    assert(db.open());
    QSqlQuery init(db);
    assert(init.exec("CREATE TABLE _la_import_log (source_file TEXT,kind TEXT,name TEXT,"
                     "status TEXT,details TEXT)"));

    for (const auto& table : {"3_tbl_Rezepte", "33_tbl_ZutatenRezept"}) {
        const auto schema = linaccess::parseMdbSchema(source.tableSchema(table));
        assert(!schema.columns.isEmpty());
        const auto csv = linaccess::AccessCatalog::exportTable(path,table);
        const auto rows = linaccess::parseCsvDetailed(std::string_view(csv.constData(), csv.size()));
        linaccess::importAccessTable(db, path, table, schema, rows,linaccess::AccessBinaryEncoding::Unsupported,linaccess::AccessFloatingSource::NativeMdbBits);
    }
    assert(scalar(db, "SELECT COUNT(*) FROM \"3_tbl_Rezepte\"") == 4);
    assert(scalar(db, "SELECT COUNT(*) FROM \"33_tbl_ZutatenRezept\"") == 35);
    assert(scalar(db, "SELECT COUNT(*) FROM \"3_tbl_Rezepte\" "
                      "WHERE RezeptNr=4 AND AnzahlPortionen IS NULL "
                      "AND Rezeptbeschreibung IS NULL") == 1);
    assert(scalar(db, "SELECT COUNT(*) FROM \"33_tbl_ZutatenRezept\" "
                      "WHERE typeof(\"Menge (g)\")='real'") > 0);
    assert(scalar(db, "SELECT COUNT(*) FROM \"3_tbl_Rezepte\" "
                      "WHERE typeof(RezeptNr)='integer'") == 4);
    assert(scalar(db, "SELECT COUNT(*) FROM _la_import_log "
                      "WHERE status='partial_typed_nulls'") == 2);
    bool refusedDuplicate = false;
    try {
        const auto csv = linaccess::AccessCatalog::exportTable(path,"3_tbl_Rezepte");
        linaccess::importAccessTable(db, path, "3_tbl_Rezepte",
            linaccess::parseMdbSchema(source.tableSchema("3_tbl_Rezepte")),
            linaccess::parseCsvDetailed(std::string_view(csv.constData(), csv.size())),linaccess::AccessBinaryEncoding::Unsupported,linaccess::AccessFloatingSource::NativeMdbBits);
    } catch (...) { refusedDuplicate = true; }
    assert(refusedDuplicate);
    assert(scalar(db, "SELECT COUNT(*) FROM _la_import_log") == 2);
    auto full = QSqlDatabase::addDatabase("QSQLITE", "full-recipe-copy");
    full.setDatabaseName(tmp.filePath("full.linaccess"));
    assert(full.open());
    QSqlQuery fullQuery(full);
    assert(fullQuery.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    assert(linaccess::importAccessDatabase(full, path) == 2);
    assert(fullQuery.exec("PRAGMA foreign_key_check"));
    assert(!fullQuery.next());
    assert(fullQuery.exec("PRAGMA foreign_key_list(\"33_tbl_ZutatenRezept\")"));
    assert(fullQuery.next());
    assert(fullQuery.value(2).toString() == "3_tbl_Rezepte");
    assert(fullQuery.value(5).toString() == "CASCADE" && fullQuery.value(6).toString() == "CASCADE");
    assert(!fullQuery.exec("UPDATE \"33_tbl_ZutatenRezept\" SET RezeptNr=999999999"));
    assert(fullQuery.exec("DELETE FROM \"3_tbl_Rezepte\""));
    assert(scalar(full, "SELECT COUNT(*) FROM \"33_tbl_ZutatenRezept\"") == 0);
    bool duplicateRefused = false;
    try { linaccess::importAccessDatabase(full, path); } catch (...) { duplicateRefused = true; }
    assert(duplicateRefused);
    assert(digest(path) == before); // Nothing touched the original MDB.
    std::cout << "Real MDB test passed: " << tables << " tables, " << forms
              << " forms, " << reports << " report, " << queries
              << " queries, " << relationships << " relationship; 39 records copied\n";
}
