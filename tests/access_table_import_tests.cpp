#include "access_table_import.hpp"
#include "access_catalog.hpp"
#include "access_index_metadata.hpp"
#include "access_floating_metadata.hpp"
#include <bit>
#include <limits>
#include "access_autonumber.hpp"
#include "jet_sql_functions.hpp"
#include "dao_database.hpp"
#include <QCoreApplication>
#include <QFile>
#include <QFileInfo>
#include <QDir>
#include <QCryptographicHash>
#include <QSet>
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

void verifyJetCounterStateReference(const QString& path) {
    using namespace linaccess;
    const auto digest=[&] { QFile source(path); assert(source.open(QIODevice::ReadOnly)); return QCryptographicHash::hash(source.readAll(),QCryptographicHash::Sha256); };
    const auto before=digest();
    auto db=QSqlDatabase::addDatabase("QSQLITE","jet-counter-state"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db); assert(query.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)")); query.finish();
    assert(importAccessDatabase(db,path)==2);
    for(const auto& table:{QString("native_counter"),QString("managed_counter")}) {
        assert(query.exec("INSERT INTO "+table+" DEFAULT VALUES"));
        assert(query.exec("SELECT MAX(id) FROM "+table) && query.next() && query.value(0).toInt()==4); query.finish();
    }
    const auto reserveEngine=makeDaoEngine(db);
    for(const auto& table:{QString("native_counter"),QString("managed_counter")}) {
        const auto cursor=makeDaoDatabase(db)->member("OpenRecordset",{VbaValue::string(table.toStdString())}).object;
        cursor->member("AddNew",{}); assert(cursor->item("id").asNumber()==5); cursor->member("CancelUpdate",{});
        cursor->member("AddNew",{}); assert(cursor->item("id").asNumber()==6); cursor->member("Update",{});
        reserveEngine->member("BeginTrans",{});
        cursor->member("AddNew",{});assert(cursor->item("id").asNumber()==7);cursor->member("Update",{});
        bool ddlRejected=false;
        try {makeDaoDatabase(db)->member("Execute",{VbaValue::string("CREATE TABLE allocation_ddl_probe (id INTEGER)")});}
        catch(const VbaRuntimeError& error){ddlRejected=error.number==3251;}
        assert(ddlRejected && !db.tables().contains("allocation_ddl_probe"));
        reserveEngine->member("Rollback",{});cursor->member("Requery",{});
        cursor->member("AddNew",{});assert(cursor->item("id").asNumber()==8);cursor->member("CancelUpdate",{});cursor->member("Close",{});
        assert(query.exec("SELECT COUNT(*) FROM "+table+" WHERE id IN (5,7,8)") && query.next() && query.value(0).toInt()==0);query.finish();
    }
    const auto records=makeDaoDatabase(db)->member("OpenRecordset",{VbaValue::string("native_counter")}).object;
    const auto field=records->member("Fields",{VbaValue::string("id")}).object;
    assert(field->member("Attributes",{}).asNumber()==49 && field->member("DataUpdatable",{}).truthy());
    records->member("Edit",{}); bool readonly=false;
    try { field->assign("Value",VbaValue::longInteger(100)); } catch(const VbaRuntimeError& error) { readonly=error.number==3164; }
    assert(readonly); records->member("CancelUpdate",{}); records->member("Close",{});
    assert(!query.exec("UPDATE native_counter SET id=100 WHERE id=1"));
    assert(jetSqlFunctionErrorNumber(query.lastError().text().toLower().toStdString())==3113); query.finish();
    AccessCatalog catalog(path);
    auto schema=parseMdbSchema(catalog.tableSchema("native_counter"));
    auto root=QJsonDocument::fromJson(schema.sourceIndexMetadata).object();
    assert(root.value("counter_columns").toArray().first().toObject().value("last_value").toInt()==3);
    root.insert("counter_columns",QJsonArray{QJsonObject{{"name","id"},{"last_value",1}}});
    schema.sourceIndexMetadata=QJsonDocument(root).toJson(QJsonDocument::Compact);
    const auto rows=parseCsvDetailed(AccessCatalog::exportTable(path,"native_counter",true).toStdString());
    bool rejected=false;
    try { importAccessTable(db,path,"invalid_counter",schema,rows,AccessBinaryEncoding::MdbToolsHex,AccessFloatingSource::NativeMdbBits); }
    catch(const std::exception&) { rejected=true; }
    assert(rejected && !db.tables().contains("invalid_counter"));
    assert(query.exec("SELECT value FROM _la_counter32_sequences WHERE table_name='native_counter'") && query.next() && query.value(0).toInt()==8); query.finish();
    const auto database=makeDaoDatabase(db);
    const auto engine=makeDaoEngine(db);
    engine->member("BeginTrans",{});
    database->member("Execute",{VbaValue::string("DROP TABLE [managed_counter]")});
    assert(!db.tables().contains("managed_counter"));
    assert(query.exec("SELECT COUNT(*) FROM _la_counter_sequences WHERE table_name='managed_counter'") && query.next() && query.value(0).toInt()==0);query.finish();
    engine->member("Rollback",{});
    assert(db.tables().contains("managed_counter"));
    assert(query.exec("SELECT value FROM _la_counter32_sequences WHERE table_name='managed_counter'") && query.next() && query.value(0).toInt()==8);query.finish();
    database->member("Execute",{VbaValue::string("DROP TABLE [managed_counter]")});
    const auto managedSchema=parseMdbSchema(catalog.tableSchema("managed_counter"));
    const auto managedRows=parseCsvDetailed(AccessCatalog::exportTable(path,"managed_counter",true).toStdString());
    importAccessTable(db,path,"managed_counter",managedSchema,managedRows,AccessBinaryEncoding::MdbToolsHex,AccessFloatingSource::NativeMdbBits);
    assert(query.exec("INSERT INTO managed_counter DEFAULT VALUES"));
    assert(query.exec("SELECT id FROM managed_counter") && query.next() && query.value(0).toInt()==4);query.finish();
    bool renameRejected=false;
    try { database->member("Execute",{VbaValue::string("ALTER TABLE [managed_counter] RENAME TO [renamed_counter]")}); }
    catch(const VbaRuntimeError& error) {renameRejected=error.number==3251;}
    assert(renameRejected && db.tables().contains("managed_counter") && !db.tables().contains("renamed_counter"));
    for(const auto& table:{QString("native_counter"),QString("managed_counter")}) {
        assert(query.exec("INSERT INTO "+table+"(id) VALUES(2147483646)"));
        const auto cursor=makeDaoDatabase(db)->member("OpenRecordset",{VbaValue::string(table.toStdString())}).object;
        cursor->member("AddNew",{});assert(cursor->item("id").asNumber()==2147483647);cursor->member("Update",{});
        cursor->member("AddNew",{});assert(cursor->item("id").asNumber()==-2147483648.0);cursor->member("Update",{});
        cursor->member("AddNew",{});assert(cursor->item("id").asNumber()==-2147483647);cursor->member("CancelUpdate",{});
        reserveEngine->member("BeginTrans",{});
        cursor->member("AddNew",{});assert(cursor->item("id").asNumber()==-2147483646);cursor->member("Update",{});
        reserveEngine->member("Rollback",{});cursor->member("Requery",{});
        cursor->member("AddNew",{});assert(cursor->item("id").asNumber()==-2147483645);cursor->member("CancelUpdate",{});cursor->member("Close",{});
        assert(query.exec("SELECT COUNT(*) FROM "+table+" WHERE id IN (-2147483647,-2147483646,-2147483645)") && query.next() && query.value(0).toInt()==0);query.finish();
        assert(query.exec("INSERT INTO "+table+" DEFAULT VALUES"));
        assert(query.exec("SELECT id FROM "+table+" WHERE rowid=last_insert_rowid()") && query.next() && query.value(0).toInt()==-2147483644);query.finish();
        assert(!query.exec("INSERT INTO "+table+"(id) VALUES(NULL)"));
        assert(jetSqlFunctionErrorNumber(query.lastError().text().toLower().toStdString())==3162);query.finish();
    }
    database->member("Execute",{VbaValue::string("CREATE TABLE [user_created_counter] ([number] INTEGER PRIMARY KEY AUTOINCREMENT,[payload] TEXT)")});
    const auto own=database->member("OpenRecordset",{VbaValue::string("USER_CREATED_COUNTER")}).object;
    own->member("AddNew",{});assert(own->item("number").asNumber()==1 && own->item("number").numericType==VbaValue::NumericType::Long);own->member("CancelUpdate",{});
    own->member("AddNew",{});assert(own->item("number").asNumber()==2);own->member("Update",{});own->member("Close",{});
    assert(jetLongAccessColumns(db,"user_created_counter").contains("number"));
    const auto ownQuery=database->member("OpenRecordset",{VbaValue::string("SELECT [number] FROM [user_created_counter]")}).object;
    assert(ownQuery->item("number").numericType==VbaValue::NumericType::Long && ownQuery->item("number").asNumber()==2);ownQuery->member("Close",{});
    assert(digest()==before);
    std::cout<<"Actual Jet stored counters: deleted-highest native/deleted-all managed next4, AddNew timing, signed32 wrap, consumed cancellation/rollback IDs, protected DDL, contradictory state atomic rollback and original hash verified\n";
}

void verifyJetBinaryReference(const QString& path) {
    using namespace linaccess;
    const QDir directory=QFileInfo(path).dir();
    const auto digest=[](const QString& file) { QFile input(file); assert(input.open(QIODevice::ReadOnly)); return QCryptographicHash::hash(input.readAll(),QCryptographicHash::Sha256); };
    const auto before=digest(path);
    auto db=QSqlDatabase::addDatabase("QSQLITE","jet-binary-reference"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db); assert(query.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)")); query.finish();
    assert(importAccessDatabase(db,path)==1);
    assert(query.exec("SELECT id,payload,typeof(payload) FROM binary_probe ORDER BY id"));
    QSet<int> seen;
    while(query.next()) {
        const int id=query.value(0).toInt(); assert(!seen.contains(id)); seen.insert(id);
        if(id==2) { assert(query.value(1).isNull() && query.value(2).toString()=="null"); continue; }
        QFile expected(directory.filePath("payload-"+QString::number(id)+".bin")); assert(expected.open(QIODevice::ReadOnly));
        assert(query.value(2).toString()=="blob" && query.value(1).toByteArray()==expected.readAll());
    }
    query.finish(); assert(seen==QSet<int>({1,2,3,4}));
    const auto rows=makeDaoDatabase(db)->member("OpenRecordset",{VbaValue::string("binary_probe"),VbaValue::integer(4)}).object;
    int count=0;
    while(!rows->member("EOF",{}).truthy()) {
        const int id=static_cast<int>(rows->item("id").asNumber()); const auto field=rows->member("Fields",{VbaValue::string("payload")}).object;
        const auto value=field->member("Value",{});
        if(id==2) assert(value.kind==VbaValue::Kind::Null);
        else {
            QFile expected(directory.filePath("payload-"+QString::number(id)+".bin")); assert(expected.open(QIODevice::ReadOnly)); const auto bytes=expected.readAll();
            assert(value.kind==VbaValue::Kind::ByteArray && QByteArray(reinterpret_cast<const char*>(value.bytes.data()),static_cast<qsizetype>(value.bytes.size()))==bytes);
            assert(field->member("Type",{}).asNumber()==11 && field->member("Size",{}).asNumber()==0);
            assert(field->member("FieldSize",{}).asNumber()==bytes.size());
            const auto tail=field->member("GetChunk",{VbaValue::longInteger(static_cast<std::int32_t>(bytes.size()-2)),VbaValue::integer(20)});
            assert(tail.kind==VbaValue::Kind::ByteArray && tail.bytes.size()==2 && QByteArray(reinterpret_cast<const char*>(tail.bytes.data()),2)==bytes.right(2));
        }
        ++count; rows->member("MoveNext",{});
    }
    rows->member("Close",{}); assert(count==4); db.close(); assert(digest(path)==before);
    const auto oversize=directory.filePath("binary-oversize.mdb"); const auto largeBefore=digest(oversize);
    auto failure=QSqlDatabase::addDatabase("QSQLITE","jet-binary-limit"); failure.setDatabaseName(":memory:"); assert(failure.open());
    QSqlQuery check(failure); assert(check.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    assert(check.exec("CREATE TABLE untouched(value INTEGER)")); assert(check.exec("INSERT INTO untouched VALUES(17)")); check.finish();
    const auto tables=failure.tables(); bool bounded=false;
    try { (void)importAccessDatabase(failure,oversize); } catch(const std::exception& error) { bounded=QString::fromUtf8(error.what()).contains("Puffergröße"); }
    assert(bounded && failure.tables()==tables && check.exec("SELECT value FROM untouched") && check.next() && check.value(0).toInt()==17); check.finish();
    failure.close(); assert(digest(oversize)==largeBefore);
    std::cout<<"Actual technical Jet binary import: exact inline/multi-page bytes, DAO chunks, NULL, bounded oversized rollback and source hashes verified\n";
}

void verifyJetDoublePrecisionReference(const QString& path) {
    using namespace linaccess;
    const auto before=accessSourceDigest(path);
    const auto schema=parseMdbSchema(AccessCatalog(path).tableSchema("double_precision_probe"));
    const auto csv=AccessCatalog::exportTable(path,"double_precision_probe"); const auto exported=parseCsvDetailed(std::string_view(csv.constData(),csv.size()));
    assert(exported.size()==9 && QString::fromStdString(exported[1][1].value).toDouble()==1.0);
    auto db=QSqlDatabase::addDatabase("QSQLITE","jet-double-precision-reference"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db); assert(query.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)")); query.finish();
    assert(importAccessDatabase(db,path)==1);
    const quint64 expected[]{0x3ff0000000000001ULL,std::bit_cast<quint64>(0.1),1ULL,0x7fefffffffffffffULL,std::bit_cast<quint64>(-123.456),0ULL,0x4340000000000000ULL};
    assert(query.exec("SELECT id,sample FROM double_precision_probe ORDER BY id")); int count=0;
    while(query.next()) { const int id=query.value(0).toInt(); assert(id==++count); if(id==8) assert(query.value(1).isNull()); else assert(std::bit_cast<quint64>(query.value(1).toDouble())==expected[id-1]); }
    assert(count==8); query.finish();
    const auto rows=makeDaoDatabase(db)->member("OpenRecordset",{VbaValue::string("double_precision_probe")}).object; count=0;
    while(!rows->member("EOF",{}).truthy()) {
        const auto value=rows->item("sample");
        if(count==7) assert(value.kind==VbaValue::Kind::Null);
        else assert(value.numericType==VbaValue::NumericType::Double && std::bit_cast<quint64>(value.asNumber())==expected[count]);
        const auto field=rows->member("Fields",{VbaValue::string("sample")}).object; assert(field->member("Type",{}).asNumber()==7 && field->member("Size",{}).asNumber()==8);
        ++count; rows->member("MoveNext",{});
    }
    assert(count==8); rows->member("Close",{});
    importAccessTable(db,path,"partial_double_precision",schema,exported,AccessBinaryEncoding::Unsupported,AccessFloatingSource::NativeMdbBits);
    assert(query.exec("SELECT sample FROM partial_double_precision WHERE id=1") && query.next() && std::bit_cast<quint64>(query.value(0).toDouble())==expected[0]); query.finish();
    auto changed=exported; changed[1][1].value="0.25"; bool refused=false;
    try { importAccessTable(db,path,"wrong_double_values",schema,changed,AccessBinaryEncoding::Unsupported,AccessFloatingSource::NativeMdbBits); } catch(const std::exception&) { refused=true; }
    assert(refused && !db.tables().contains("wrong_double_values"));
    changed=exported; changed.back()[1].value="0"; refused=false;
    try { importAccessTable(db,path,"wrong_double_nulls",schema,changed,AccessBinaryEncoding::Unsupported,AccessFloatingSource::NativeMdbBits); } catch(const std::exception&) { refused=true; }
    assert(refused && !db.tables().contains("wrong_double_nulls"));
    auto callerCsv=exported; callerCsv.resize(2); callerCsv[1][1].value="0.25";
    importAccessTable(db,"caller-owned-technical-csv","caller_csv",schema,callerCsv);
    assert(query.exec("SELECT sample FROM caller_csv") && query.next() && query.value(0).toDouble()==0.25); query.finish();
    refused=false;
    try { importAccessTable(db,"missing-native-source","unproven_native",schema,callerCsv,AccessBinaryEncoding::Unsupported,AccessFloatingSource::NativeMdbBits); } catch(const std::exception&) { refused=true; }
    assert(refused && !db.tables().contains("unproven_native")); db.close(); assert(accessSourceDigest(path)==before);
    const auto negative=QFileInfo(path).dir().filePath("double-negative-zero.mdb"); const auto negativeBefore=accessSourceDigest(negative);
    auto rejected=QSqlDatabase::addDatabase("QSQLITE","jet-double-negative-zero"); rejected.setDatabaseName(":memory:"); assert(rejected.open());
    QSqlQuery untouched(rejected); assert(untouched.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    assert(untouched.exec("CREATE TABLE untouched(value INTEGER)")); assert(untouched.exec("INSERT INTO untouched VALUES(17)")); untouched.finish();
    const auto tables=rejected.tables(); refused=false;
    try { (void)importAccessDatabase(rejected,negative); } catch(const std::exception&) { refused=true; }
    assert(refused && rejected.tables()==tables && untouched.exec("SELECT value FROM untouched") && untouched.next() && untouched.value(0).toInt()==17); untouched.finish();
    rejected.close(); assert(accessSourceDigest(negative)==negativeBefore);
    std::cout<<"Actual Jet Double precision: original bits despite16-digit CSV loss/overflow, finite extremes, NULL, full/direct import, DAO and atomic refusal verified\n";
}

void verifyJetSinglePrecisionReference(const QString& path) {
    using namespace linaccess;
    const auto before=accessSourceDigest(path);
    const auto schema=parseMdbSchema(AccessCatalog(path).tableSchema("single_precision_probe"));
    const auto csv=AccessCatalog::exportTable(path,"single_precision_probe");
    const auto exported=parseCsvDetailed(std::string_view(csv.constData(),csv.size()));
    assert(exported.size()==9 && std::bit_cast<quint32>(static_cast<float>(QString::fromStdString(exported[1][1].value).toDouble()))!=509022793u);
    auto db=QSqlDatabase::addDatabase("QSQLITE","jet-single-precision-reference"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db); assert(query.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)")); query.finish();
    assert(importAccessDatabase(db,path)==1);
    const quint32 expected[]{509022793u,std::bit_cast<quint32>(0.1f),1u,0x7f7fffffu,std::bit_cast<quint32>(-123.456f),0u,0x4b800000u};
    assert(query.exec("SELECT id,sample FROM single_precision_probe ORDER BY id")); int count=0;
    while(query.next()) {
        const int id=query.value(0).toInt(); assert(id==++count);
        if(id==8) assert(query.value(1).isNull());
        else assert(query.value(1).toDouble()==static_cast<double>(std::bit_cast<float>(expected[id-1])));
    }
    assert(count==8); query.finish();
    const auto rows=makeDaoDatabase(db)->member("OpenRecordset",{VbaValue::string("single_precision_probe")}).object;
    count=0;
    while(!rows->member("EOF",{}).truthy()) {
        const auto value=rows->item("sample");
        if(count==7) assert(value.kind==VbaValue::Kind::Null);
        else assert(value.numericType==VbaValue::NumericType::Single && std::bit_cast<quint32>(static_cast<float>(value.asNumber()))==expected[count]);
        ++count; rows->member("MoveNext",{});
    }
    assert(count==8); rows->member("Close",{});
    importAccessTable(db,path,"partial_precision",schema,exported,AccessBinaryEncoding::Unsupported,AccessFloatingSource::NativeMdbBits);
    assert(query.exec("SELECT sample FROM partial_precision WHERE id=1") && query.next() && query.value(0).toDouble()==static_cast<double>(std::bit_cast<float>(expected[0]))); query.finish();
    auto mismatched=exported; mismatched[1][1].value="0.25"; bool refused=false;
    try { importAccessTable(db,path,"wrong_values",schema,mismatched,AccessBinaryEncoding::Unsupported,AccessFloatingSource::NativeMdbBits); } catch(const std::exception&) { refused=true; }
    assert(refused && !db.tables().contains("wrong_values"));
    mismatched=exported; mismatched.back()[1].value="0"; refused=false;
    try { importAccessTable(db,path,"wrong_nulls",schema,mismatched,AccessBinaryEncoding::Unsupported,AccessFloatingSource::NativeMdbBits); } catch(const std::exception&) { refused=true; }
    assert(refused && !db.tables().contains("wrong_nulls"));
    mismatched=exported; mismatched.pop_back(); refused=false;
    try { importAccessTable(db,path,"wrong_count",schema,mismatched,AccessBinaryEncoding::Unsupported,AccessFloatingSource::NativeMdbBits); } catch(const std::exception&) { refused=true; }
    assert(refused && !db.tables().contains("wrong_count"));
    auto wrongAffinity=schema; wrongAffinity.columns[1].sqliteType="TEXT"; refused=false;
    try { importAccessTable(db,path,"wrong_affinity",wrongAffinity,exported,AccessBinaryEncoding::Unsupported,AccessFloatingSource::NativeMdbBits); } catch(const std::exception&) { refused=true; }
    assert(refused && !db.tables().contains("wrong_affinity")); db.close(); assert(accessSourceDigest(path)==before);
    const auto negative=QFileInfo(path).dir().filePath("single-negative-zero.mdb");
    const auto negativeBefore=accessSourceDigest(negative);
    auto rejected=QSqlDatabase::addDatabase("QSQLITE","jet-single-negative-zero"); rejected.setDatabaseName(":memory:"); assert(rejected.open());
    QSqlQuery untouched(rejected); assert(untouched.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    assert(untouched.exec("CREATE TABLE untouched(value INTEGER)")); assert(untouched.exec("INSERT INTO untouched VALUES(17)")); untouched.finish();
    const auto tables=rejected.tables(); refused=false;
    try { (void)importAccessDatabase(rejected,negative); } catch(const std::exception&) { refused=true; }
    assert(refused && rejected.tables()==tables && untouched.exec("SELECT value FROM untouched") && untouched.next() && untouched.value(0).toInt()==17); untouched.finish();
    rejected.close(); assert(accessSourceDigest(negative)==negativeBefore);
    std::cout<<"Actual Jet Single precision: original bits despite lossy CSV, finite extremes, NULL, direct/full import, DAO values, mismatched CSV and negative-zero atomic refusal verified\n";
}

void verifyJetIndexCatalogReference(const QString& path) {
    using namespace linaccess;
    QFile input(path); assert(input.open(QIODevice::ReadOnly));
    const auto before=QCryptographicHash::hash(input.readAll(),QCryptographicHash::Sha256); input.close();
    const auto schema=parseMdbSchema(AccessCatalog(path).tableSchema("catalog_probe"));
    assert(schema.columns.size()==3 && schema.indexes.size()==1 && schema.indexes[0].name=="ActualSecondary");
    const QStringList names{"primary_a","primary_b","other_key"};
    const auto source=checkedAccessSourceIndexes(schema.sourceIndexMetadata,names);
    assert(source.size()==2); int primaries=0;
    for(const auto& index:source) if(index.kind==1) { assert(index.name=="DeliberatelyNamedPrimary" && index.columns==QStringList({"primary_a","primary_b"}) && index.descending==QVector<bool>({true,false})); ++primaries; }
    assert(primaries==1);
    auto db=QSqlDatabase::addDatabase("QSQLITE","jet-index-catalog-reference"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db); assert(query.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)")); query.finish();
    assert(importAccessDatabase(db,path)==2);
    assert(query.exec("SELECT definition FROM main._la_access_indexes WHERE table_name='catalog_probe'") && query.next());
    const auto stored=checkedAccessSourceIndexes(query.value(0).toString().toUtf8(),names); assert(stored.size()==2); query.finish();
    assert(query.exec("PRAGMA main.index_list('catalog_probe')")); int indexes=0;
    while(query.next()) {
        const bool primary=query.value(3).toString()=="pk"; assert(primary || query.value(3).toString()=="c");
        assert(query.value(2).toBool()==primary); auto name=query.value(1).toString(); name.replace('\"',"\"\"");
        QSqlQuery fields(db); assert(fields.exec("PRAGMA main.index_xinfo(\""+name+"\")")); int ordinal=0;
        while(fields.next()) if(fields.value(5).toBool()) {
            assert(ordinal<2 && fields.value(2).toString()==(ordinal==0 ? primary ? "primary_a" : "other_key" : "primary_b"));
            assert(fields.value(3).toBool()==(primary ? ordinal==0 : ordinal==1)); ++ordinal;
        }
        assert(ordinal==2); ++indexes;
    }
    assert(indexes==2); query.finish();
    assert(query.exec("PRAGMA main.index_list('single_primary_probe')") && query.next() && query.value(3).toString()=="pk");
    auto singleName=query.value(1).toString(); singleName.replace('\"',"\"\""); query.finish();
    assert(query.exec("PRAGMA main.index_xinfo(\""+singleName+"\")") && query.next() && query.value(2).toString()=="primary_key" && query.value(3).toBool()); query.finish();
    assert(query.exec("SELECT definition FROM main._la_access_indexes WHERE table_name='single_primary_probe'") && query.next());
    const auto singleIndexes=checkedAccessSourceIndexes(query.value(0).toString().toUtf8(),{"primary_key"});
    assert(singleIndexes.size()==1 && singleIndexes[0].name=="NamedDescendingSingle" && singleIndexes[0].descending==QVector<bool>({true})); query.finish();
    auto table=makeDaoDatabase(db)->member("TableDefs",{VbaValue::string("catalog_probe")}).object;
    auto daoIndexes=table->member("Indexes",{}).object; assert(daoIndexes->member("Count",{}).asNumber()==2);
    auto daoPrimary=daoIndexes->member("Item",{VbaValue::string("DeliberatelyNamedPrimary")}).object;
    assert(daoPrimary->member("Primary",{}).truthy() && daoPrimary->member("Unique",{}).truthy());
    assert(daoPrimary->member("Fields",{VbaValue::numeric(0)}).object->member("Attributes",{}).asNumber()==1);
    auto single=makeDaoDatabase(db)->member("TableDefs",{VbaValue::string("single_primary_probe")}).object->member("Indexes",{VbaValue::string("NamedDescendingSingle")}).object;
    assert(single->member("Primary",{}).truthy() && single->member("Fields",{VbaValue::numeric(0)}).object->member("Attributes",{}).asNumber()==1);
    auto malformed=QJsonDocument::fromJson(schema.sourceIndexMetadata).object(); auto entries=malformed["indexes"].toArray(); auto duplicate=entries[0].toObject();
    auto second=entries[1].toObject(); second["name"]=duplicate["name"]; entries[1]=second; malformed["indexes"]=entries;
    auto bad=schema; bad.sourceIndexMetadata=QJsonDocument(malformed).toJson(QJsonDocument::Compact); bool refused=false;
    try { importAccessTable(db,"technical","invalid_catalog",bad,parseCsvDetailed("primary_a,primary_b,other_key\n")); } catch(const std::exception&) { refused=true; }
    assert(refused && !db.tables().contains("invalid_catalog"));
    db.close(); assert(input.open(QIODevice::ReadOnly)); assert(QCryptographicHash::hash(input.readAll(),QCryptographicHash::Sha256)==before);
    std::cout<<"Actual Jet index catalog: original primary/secondary names, flags, mixed primary directions, persistence, refusal and source hash verified\n";
}

void verifyJetSingleReference(const QString& path) {
    using namespace linaccess;
    QFile input(path); assert(input.open(QIODevice::ReadOnly));
    const auto before=QCryptographicHash::hash(input.readAll(),QCryptographicHash::Sha256); input.close();
    auto db=QSqlDatabase::addDatabase("QSQLITE","jet-single-reference"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db); assert(query.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)")); query.finish();
    assert(importAccessDatabase(db,path)==1);
    auto rows=makeDaoDatabase(db)->member("OpenRecordset",{VbaValue::string("single_type_probe")}).object;
    const float expected[]{0.1f,16777216.0f,1.5f}; int count=0;
    while(!rows->member("EOF",{}).truthy()) {
        const auto id=static_cast<int>(rows->item("id").asNumber()); assert(id>=1 && id<=4);
        const auto field=rows->member("Fields",{VbaValue::string("sample")}).object;
        assert(field->member("Type",{}).asNumber()==6 && field->member("Size",{}).asNumber()==4);
        const auto value=field->member("Value",{});
        if(id==4) assert(value.kind==VbaValue::Kind::Null);
        else assert(value.kind==VbaValue::Kind::Number && value.numericType==VbaValue::NumericType::Single && value.asNumber()==static_cast<double>(expected[id-1]));
        ++count; rows->member("MoveNext",{});
    }
    assert(count==4); rows->member("MoveFirst",{}); rows->member("Edit",{}); rows->assignItem("sample",VbaValue::numeric(0.2)); rows->member("Update",{});
    assert(query.exec("SELECT sample FROM single_type_probe WHERE id=1") && query.next() && query.value(0).toDouble()==static_cast<double>(0.2f)); query.finish();
    rows->member("Requery",{}); assert(rows->item("sample").numericType==VbaValue::NumericType::Single && rows->item("sample").asNumber()==static_cast<double>(0.2f));
    rows->member("Close",{}); db.close(); assert(input.open(QIODevice::ReadOnly)); assert(QCryptographicHash::hash(input.readAll(),QCryptographicHash::Sha256)==before);
    std::cout<<"Actual technical Jet Single import: source field identity, Single values, NULL, rounded native write and unchanged hash verified\n";
}

void verifyJetIndexReference(const QString& path) {
    using namespace linaccess;
    QFile input(path); assert(input.open(QIODevice::ReadOnly));
    const auto before=QCryptographicHash::hash(input.readAll(),QCryptographicHash::Sha256); input.close();
    const auto schema=parseMdbSchema(AccessCatalog(path).tableSchema("index_order_probe"));
    assert(schema.columns.size()==2 && schema.indexes.size()==2);
    int mixed=0,ascending=0;
    for(const auto& index:schema.indexes) {
        if(index.columns.size()==2) { assert(index.unique && index.columns==QStringList({"first_key","second_key"}) && index.descending==QVector<bool>({true,false})); ++mixed; }
        else { assert(index.columns==QStringList({"first_key"}) && index.descending==QVector<bool>({false})); ++ascending; }
    }
    assert(mixed==1 && ascending==1);
    auto db=QSqlDatabase::addDatabase("QSQLITE","jet-index-reference"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db); assert(query.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)")); query.finish();
    assert(importAccessDatabase(db,path)==1);
    int indexes=0,keys=0;
    assert(query.exec("PRAGMA main.index_list('index_order_probe')"));
    while(query.next()) {
        assert(query.value(3).toString()=="c" && !query.value(4).toBool());
        const bool unique=query.value(2).toBool();
        auto name=query.value(1).toString(); name.replace('\"',"\"\"");
        QSqlQuery fields(db); assert(fields.exec("PRAGMA main.index_xinfo(\""+name+"\")")); int ordinal=0;
        while(fields.next()) if(fields.value(5).toBool()) {
            assert(fields.value(2).toString()==(ordinal==0 ? "first_key" : "second_key"));
            assert(fields.value(3).toBool()==(unique && ordinal==0));
            assert(fields.value(4).toString()=="BINARY"); ++ordinal; ++keys;
        }
        assert(ordinal==(unique ? 2 : 1)); ++indexes;
    }
    assert(indexes==2 && keys==3); query.finish();
    assert(query.exec("SELECT COUNT(*) FROM index_order_probe") && query.next() && query.value(0).toInt()==0); query.finish();
    db.close(); assert(input.open(QIODevice::ReadOnly)); assert(QCryptographicHash::hash(input.readAll(),QCryptographicHash::Sha256)==before);
    std::cout<<"Actual technical Jet index import: mixed directions, key order, uniqueness and source hash verified\n";
}

int main(int argc, char** argv) {
    QCoreApplication app(argc, argv);
    using namespace linaccess;
    QJsonObject rawSingles{{"version",1},{"source_table","technical"},{"source_sha256",QString(64,'0')},{"columns",QJsonArray{"sample"}},{"rows",QJsonArray{QJsonArray{"3dcccccd"},QJsonArray{QJsonValue::Null}}}};
    const auto decoded=checkedAccessSingleValues(QJsonDocument(rawSingles).toJson(QJsonDocument::Compact),{"sample"},"technical");
    assert(decoded.rowCount==2 && decoded.columns.value("sample")[0]==0.1f && !decoded.columns.value("sample")[1]);
    for(const auto* bits:{"80000000","7f800000","ff800000","7fc00000","GG000000","0000000"}) {
        auto invalid=rawSingles; invalid["rows"]=QJsonArray{QJsonArray{bits}}; bool refused=false;
        try { (void)checkedAccessSingleValues(QJsonDocument(invalid).toJson(QJsonDocument::Compact),{"sample"},"technical"); } catch(const std::exception&) { refused=true; }
        assert(refused);
    }
    assert(mdbFloatingExportMatches("1",std::bit_cast<double>(quint64{0x3ff0000000000001ULL}),16));
    assert(mdbFloatingExportMatches("1.797693134862316e+308",std::numeric_limits<double>::max(),16));
    assert(mdbFloatingExportMatches("4.940656458412465e-324",std::numeric_limits<double>::denorm_min(),16));
    assert(mdbFloatingExportMatches("1.0039062",1.00390625,8) && !mdbFloatingExportMatches("1.0039063",1.00390625,8));
    assert(!mdbFloatingExportMatches("1e999999",1.0,16) && !mdbFloatingExportMatches("1,0",1.0,16));
    auto rawDoubles=rawSingles; rawDoubles["rows"]=QJsonArray{QJsonArray{"3ff0000000000001"}};
    const auto decodedDouble=checkedAccessDoubleValues(QJsonDocument(rawDoubles).toJson(QJsonDocument::Compact),{"sample"},"technical");
    assert(std::bit_cast<quint64>(decodedDouble.columns.value("sample")[0].value())==0x3ff0000000000001ULL);
    for(const auto* bits:{"8000000000000000","7ff0000000000000","fff0000000000000","7ff8000000000000","GG00000000000000","000000000000000"}) {
        auto invalid=rawDoubles; invalid["rows"]=QJsonArray{QJsonArray{bits}}; bool refused=false;
        try { (void)checkedAccessDoubleValues(QJsonDocument(invalid).toJson(QJsonDocument::Compact),{"sample"},"technical"); } catch(const std::exception&) { refused=true; }
        assert(refused);
    }
    const auto floating=AccessCatalog::floatingSourceTypes("CREATE TABLE [values]\n(\n[price] Currency,\n[ö] Single,\n[odd] Currency] Single NOT NULL,\n[other] Double,\n-- [comment] Currency\n);\n");
    assert(floating.size()==4 && floating.value("other")=="DOUBLE" && floating.value("price")=="CURRENCY" && floating.value("ö")=="SINGLE" && floating.value("odd] Currency")=="SINGLE");
    assert(AccessCatalog::floatingSourceTypes("[label] SingleText\n[price] CurrencyOther\n[other] DoubleOther").isEmpty());
    bool conflict=false; try { (void)AccessCatalog::floatingSourceTypes("[x] Currency\n[x] Single"); } catch(const std::runtime_error&) { conflict=true; } assert(conflict);
    const auto currencySchema=parseMdbSchema("CREATE TABLE `cash`\n(\n`price` REAL,\n`rate` REAL\n);\n-- LINACCESS_SOURCE_TYPE 7072696365 FLOAT\n-- LINACCESS_SOURCE_TYPE 7072696365 CURRENCY\n-- LINACCESS_SOURCE_TYPE 72617465 SINGLE\n");
    assert(currencySchema.columns.size()==2 && currencySchema.columns[0].sourceType=="CURRENCY" && currencySchema.columns[1].sourceType=="SINGLE");
    assert(currencySchema.columns[0].sqliteType=="REAL" && currencySchema.columns[1].sqliteType=="REAL");
    const auto literalDefaults=parseMdbSchema("CREATE TABLE `source`\n(\n`qty` INTEGER DEFAULT 7,\n`flag` INTEGER DEFAULT FALSE,\n`label` TEXT DEFAULT 'native''s NOT NULL',\n`optional` TEXT DEFAULT NULL,\n`amount` REAL DEFAULT 1.25\n);\n");
    assert(literalDefaults.columns.size()==5 && literalDefaults.columns[0].hasDefault && literalDefaults.columns[1].sqlDefault=="0");
    assert(!literalDefaults.columns[2].required && literalDefaults.columns[2].sqlDefault=="'native''s NOT NULL'");
    assert(parseMdbSchema("CREATE TABLE `source`\n(\n`stamp` TEXT DEFAULT Now()\n);\n").columns.isEmpty());
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
    const QString orderedPrefix="CREATE TABLE `x`\n(\n`a``b` INTEGER,\n`second` INTEGER\n);\n";
    const auto ordered=parseMdbSchema(orderedPrefix+"CREATE UNIQUE INDEX `ordered` ON `x` (`a``b` DESC, `second` ASC);\n");
    assert(ordered.indexes.size()==1 && ordered.indexes[0].columns==QStringList({"a`b","second"}) && ordered.indexes[0].descending==QVector<bool>({true,false}));
    const auto altered=parseMdbSchema(orderedPrefix+"ALTER TABLE `x` ADD INDEX `ordered` (`a``b` desc, `second`);\n");
    assert(altered.indexes.size()==1 && altered.indexes[0].descending==QVector<bool>({true,false}));
    for(const auto* fields:{"`a``b` SIDEWAYS","`a``b`,","`a``b` DESC extra","`a``b` DESC `second`","`a``b`, lower(`second`)"})
        assert(parseMdbSchema(orderedPrefix+"CREATE INDEX `bad` ON `x` ("+fields+");\n").columns.isEmpty());
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
    auto incompleteDirections=schema; incompleteDirections.indexes[0].descending=QVector<bool>{true};
    bool invalidDirections=false;
    try { importAccessTable(db,"technical","invalid_directions",incompleteDirections,data); } catch(const std::exception&) { invalidDirections=true; }
    assert(invalidDirections && !db.tables().contains("invalid_directions"));
    assert(query.exec("SELECT COUNT(*) FROM _la_import_log"));
    assert(query.next() && query.value(0).toInt() == 1);
    query.finish();
    importAccessTable(db,"synthetic-defaults","literal_defaults",literalDefaults,parseCsvDetailed("qty,flag,label,optional,amount\n"));
    {
        const auto records=makeDaoDatabase(db)->member("OpenRecordset",{VbaValue::string("literal_defaults")}).object;
        const auto quantity=records->member("Fields",{VbaValue::string("qty")}).object;
        const auto caption=records->member("Fields",{VbaValue::string("label")}).object;
        const auto optional=records->member("Fields",{VbaValue::string("optional")}).object;
        assert(quantity->member("DefaultValue",{}).kind==VbaValue::Kind::Text && quantity->member("DefaultValue",{}).asText()=="7");
        assert(caption->member("DefaultValue",{}).asText()=="\"native's NOT NULL\"");
        assert(optional->member("DefaultValue",{}).asText()=="Null");
        const auto schemaField=makeDaoDatabase(db)->member("TableDefs",{VbaValue::string("literal_defaults")}).object->member("Fields",{VbaValue::string("qty")}).object;
        assert(schemaField->member("DefaultValue",{}).asText()=="7");
        bool readOnly=false; try { quantity->assign("DefaultValue",VbaValue::string("8")); } catch(const VbaRuntimeError& error) { readOnly=error.number==383; } assert(readOnly);
        records->member("AddNew",{});
        assert(records->member("Fields",{VbaValue::string("qty")}).object->member("Value",{}).number==7);
        assert(records->member("Fields",{VbaValue::string("label")}).object->member("Value",{}).asText()=="native's NOT NULL");
        assert(records->member("Fields",{VbaValue::string("optional")}).object->member("Value",{}).kind==VbaValue::Kind::Null);
        records->member("CancelUpdate",{});
        assert(query.exec("SELECT COUNT(*) FROM literal_defaults") && query.next() && query.value(0).toInt()==0); query.finish();
        records->member("AddNew",{}); records->member("Update",{});
        assert(records->member("Fields",{VbaValue::string("qty")}).object->member("Value",{}).number==7);
        records->member("Close",{});
        assert(query.exec("SELECT qty,flag,label,optional,amount FROM literal_defaults") && query.next());
        assert(query.value(0).toInt()==7 && query.value(1).toInt()==0 && query.value(2).toString()=="native's NOT NULL" && query.value(3).isNull() && query.value(4).toDouble()==1.25); query.finish();
    }
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
    assert(write.value(0).toInt()==40 && write.value(1).toInt()==40 && write.value(2).toInt()==136); write.finish();
    assert(!write.exec("SELECT linaccess_next_counter(NULL,'id')"));
    assert(jetSqlFunctionErrorNumber(write.lastError().text().toCaseFolded().toStdString())==94);
    assert(!write.exec("SELECT linaccess_next_counter('missing','id')"));
    assert(jetSqlFunctionErrorNumber(write.lastError().text().toCaseFolded().toStdString())==3251);
    makeDaoDatabase(file)->member("Execute",{VbaValue::string("CREATE TABLE durable_jet_counter(id INTEGER PRIMARY KEY AUTOINCREMENT)")});
    auto durable=makeDaoDatabase(file)->member("OpenRecordset",{VbaValue::string("durable_jet_counter")}).object;
    durable->member("AddNew",{});assert(durable->item("id").asNumber()==1);durable->member("CancelUpdate",{});durable->member("Close",{});durable.reset();
    write.clear();file.close();assert(file.open());installJetSqlFunctions(file);write=QSqlQuery(file);
    durable=makeDaoDatabase(file)->member("OpenRecordset",{VbaValue::string("durable_jet_counter")}).object;
    durable->member("AddNew",{});assert(durable->item("id").asNumber()==2);durable->member("Update",{});durable->member("Close",{});
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
    const auto binary=parseMdbSchema("CREATE TABLE `binary`\n(\n`id` INTEGER,\n`payload` BLOB\n);\n");
    const auto binaryRows=parseCsvDetailed("id,payload\n1,00ff\n2,\x1e\n3,\n");
    bool codecRequired=false;
    try { importAccessTable(db,"technical","binary_no_codec",binary,binaryRows); } catch(const std::exception&) { codecRequired=true; }
    assert(codecRequired && !db.tables().contains("binary_no_codec"));
    importAccessTable(db,"technical","binary_hex",binary,binaryRows,AccessBinaryEncoding::MdbToolsHex);
    QSqlQuery binaryCheck(db); assert(binaryCheck.exec("SELECT hex(payload),typeof(payload),length(payload) FROM binary_hex ORDER BY id"));
    assert(binaryCheck.next() && binaryCheck.value(0).toString()=="00FF" && binaryCheck.value(1).toString()=="blob" && binaryCheck.value(2).toInt()==2);
    assert(binaryCheck.next() && binaryCheck.value(1).toString()=="null");
    assert(binaryCheck.next() && binaryCheck.value(1).toString()=="blob" && binaryCheck.value(2).toInt()==0); binaryCheck.finish();
    for(const auto* malformed:{"0","0x00","GG","00 1","00\n00"}) {
        bool rejected=false; try { (void)decodeMdbBinaryHex(malformed); } catch(const std::exception&) { rejected=true; } assert(rejected);
    }
    const auto counterReference=qEnvironmentVariable("LINACCESS_JET_COUNTER_STATE_REFERENCE"); if(!counterReference.isEmpty()) verifyJetCounterStateReference(counterReference);
    const auto doubleReference=qEnvironmentVariable("LINACCESS_JET_DOUBLE_PRECISION_REFERENCE"); if(!doubleReference.isEmpty()) verifyJetDoublePrecisionReference(doubleReference);
    const auto precisionReference=qEnvironmentVariable("LINACCESS_JET_SINGLE_PRECISION_REFERENCE"); if(!precisionReference.isEmpty()) verifyJetSinglePrecisionReference(precisionReference);
    const auto catalogReference=qEnvironmentVariable("LINACCESS_JET_INDEX_CATALOG_REFERENCE"); if(!catalogReference.isEmpty()) verifyJetIndexCatalogReference(catalogReference);
    const auto singleReference=qEnvironmentVariable("LINACCESS_JET_SINGLE_REFERENCE"); if(!singleReference.isEmpty()) verifyJetSingleReference(singleReference);
    const auto indexReference=qEnvironmentVariable("LINACCESS_JET_INDEX_REFERENCE"); if(!indexReference.isEmpty()) verifyJetIndexReference(indexReference);
    const auto reference=qEnvironmentVariable("LINACCESS_JET_BINARY_REFERENCE"); if(!reference.isEmpty()) verifyJetBinaryReference(reference);
    std::cout << "Schema, composite keys, indices, quoted identifiers and rollback verified\n";
}
