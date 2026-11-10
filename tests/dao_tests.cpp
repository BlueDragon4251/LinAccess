#include "native_field_defaults.hpp"
#include "access_table_import.hpp"
#include <QDateTime>
#include <QFile>
#include "dao_database.hpp"
#include "native_schema_editor.hpp"
#include "vba_array.hpp"
#include "vba_decimal.hpp"
#include <QFileInfo>
#include <bit>
#include "vba_qt_value.hpp"
#include "access_query.hpp"
#include "sqlite_execution.hpp"
#include "jet_sql_functions.hpp"
#include "vba_module_context.hpp"
#include "vba_runtime.hpp"
#include <QCoreApplication>
#include <QSqlQuery>
#include <QSqlRecord>
#include <QSqlError>
#include <QJsonDocument>
#include <QJsonArray>
#include <QJsonObject>
#include <QDir>
#include <QTemporaryDir>
#include <QTemporaryFile>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>
#include <map>
#include <limits>
#include <cmath>
#include <thread>

class Host final : public linaccess::VbaHost {
public:
    explicit Host(QSqlDatabase db) : db_(db) {}
    std::shared_ptr<linaccess::VbaObject> databaseEngine() override { return linaccess::makeDaoEngine(db_); }
    std::shared_ptr<linaccess::VbaObject> currentDatabase() override { return linaccess::makeDaoDatabase(db_); }
    std::map<std::string,std::optional<std::string>> fields;
    std::optional<std::string> field(std::string_view name) override { return fields[std::string(name)]; }
    void setField(std::string_view name, std::optional<std::string> value) override { fields[std::string(name)] = std::move(value); }
    void message(std::string_view) override { assert(false); }
    void saveRecord() override { assert(false); }
private:
    QSqlDatabase db_;
};
void testDaoCursors();
void testDaoTransactions();
void testDaoIndexes() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","dao-indexes"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE keyed(primary_a INTEGER NOT NULL,primary_b INTEGER NOT NULL,other_key INTEGER,PRIMARY KEY(primary_a DESC,primary_b ASC))"));
    assert(sql.exec("CREATE INDEX keyed_ActualSecondary ON keyed(other_key ASC,primary_b DESC)"));
    assert(sql.exec("CREATE TABLE _la_access_indexes(table_name TEXT PRIMARY KEY,definition TEXT NOT NULL)"));
    QJsonObject definition{{"version",1},{"source_table","keyed"},{"columns",QJsonArray{"primary_a","primary_b","other_key"}},
        {"indexes",QJsonArray{
            QJsonObject{{"name","NamedPrimary"},{"kind",1},{"flags",137},{"fields",QJsonArray{QJsonObject{{"name","primary_a"},{"descending",true}},QJsonObject{{"name","primary_b"},{"descending",false}}}}},
            QJsonObject{{"name","ActualSecondary"},{"kind",0},{"flags",128},{"fields",QJsonArray{QJsonObject{{"name","other_key"},{"descending",false}},QJsonObject{{"name","primary_b"},{"descending",true}}}}}}}};
    const auto store=[&] { sql.prepare("INSERT OR REPLACE INTO main._la_access_indexes VALUES('keyed',?)"); sql.addBindValue(QJsonDocument(definition).toJson(QJsonDocument::Compact)); assert(sql.exec()); sql.finish(); }; store();
    auto database=linaccess::makeDaoDatabase(db); auto table=database->member("TableDefs",{VbaValue::string("keyed")}).object;
    auto indexes=table->member("Indexes",{}).object; auto primary=indexes->member("Item",{VbaValue::string("namedprimary")}).object;
    auto field=primary->member("Fields",{VbaValue::string("PRIMARY_A")}).object;
    assert(indexes->member("Count",{}).asNumber()==2 && indexes->member("Count",{}).numericType==VbaValue::NumericType::Integer);
    assert(primary==table->member("Indexes",{VbaValue::numeric(0)}).object && primary==indexes->enumerate()[0].object);
    assert(field->member("Attributes",{}).asNumber()==1 && field->member("Attributes",{}).numericType==VbaValue::NumericType::Long);
    const auto error=[](const std::function<void()>& operation,int expected) { int number=0; try { operation(); } catch(const linaccess::VbaRuntimeError& failure) { number=failure.number; } assert(number==expected); };
    error([&] { indexes->member("Item",{VbaValue{}}); },94);
    error([&] { indexes->member("Item",{VbaValue::numeric(2)}); },3265);
    error([&] { indexes->member("Refresh",{VbaValue::numeric(1)}); },450);
    error([&] { primary->assign("Name",VbaValue::string("renamed")); },383);
    error([&] { field->assign("Attributes",VbaValue::numeric(0)); },383);
    error([&] { primary->member("Foreign",{}); },3251);
    error([&] { field->member("Type",{}); },3251);
    Host host(db); const auto result=linaccess::executeVbaProcedure(R"VBA(Sub InspectIndexes()
Dim table As DAO.TableDef, indexes As DAO.Indexes, index As DAO.Index, fields As DAO.Fields, field As DAO.Field
Dim count As Long
Set table=CurrentDb.TableDefs("keyed")
Set indexes=table.Indexes
If Not TypeOf indexes Is DAO.Indexes Then Err.Raise 5
If indexes.Count<>2 Or VarType(indexes.Count)<>vbInteger Then Err.Raise 5
Set index=table.Indexes(Index:="NamedPrimary")
If Not TypeOf index Is DAO.Index Then Err.Raise 5
If Not (index Is indexes.Item(Index:=0)) Then Err.Raise 5
If Not index.Primary Or Not index.Unique Or Not index.Required Or index.IgnoreNulls Then Err.Raise 5
If VarType(index.Primary)<>vbBoolean Then Err.Raise 5
Set fields=index.Fields
Set field=index.Fields(Index:="primary_a")
If field.Name<>"primary_a" Or field.Attributes<>dbDescending Or VarType(field.Attributes)<>vbLong Then Err.Raise 5
If index.Fields(1).Attributes<>0 Then Err.Raise 5
For Each index In indexes
For Each field In index.Fields
count=count+1
Next
Next
If count<>4 Then Err.Raise 5
indexes.Refresh
fields.Refresh
Me!IndexesChecked="yes"
End Sub)VBA","InspectIndexes",host);
    if(result.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Index VBA error="<<result.errorNumber<<", line="<<result.line<<'\n';
    assert(result.status==linaccess::VbaExecutionStatus::Completed && host.fields["IndexesChecked"]=="yes");
    assert(sql.exec("CREATE TEMP TABLE _la_access_indexes(bad)")); assert(sql.exec("CREATE TEMP TABLE keyed(bad)"));
    assert(indexes->member("Count",{}).asNumber()==2 && field->member("Name",{}).asText()=="primary_a");
    auto entries=definition["indexes"].toArray(); auto entry=entries[1].toObject(); entry["flags"]=0; entries[1]=entry; definition["indexes"]=entries; store();
    error([&] { field->member("Name",{}); },3420); primary=indexes->member("Item",{VbaValue::numeric(0)}).object;
    auto engine=linaccess::makeDaoEngine(db); engine->member("BeginTrans",{}); engine->member("Rollback",{});
    error([&] { primary->member("Name",{}); },3420); indexes=table->member("Indexes",{}).object;
    assert(sql.exec("CREATE INDEX main.unexpected_native ON keyed(other_key)")); error([&] { indexes->member("Count",{}); },3420);
    indexes=table->member("Indexes",{}).object; error([&] { indexes->member("Count",{}); },3251);
    assert(sql.exec("DROP INDEX main.unexpected_native")); indexes=table->member("Indexes",{}).object; assert(indexes->member("Count",{}).asNumber()==2);
    assert(sql.exec("CREATE TABLE missing_index_identity(id INTEGER)"));
    auto missing=database->member("TableDefs",{VbaValue::string("missing_index_identity")}).object->member("Indexes",{}).object;
    error([&] { missing->member("Count",{}); },3251);
    db.close(); error([&] { indexes->member("Count",{}); },3420);
}

void testDaoSingleFields() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","dao-single-fields"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE samples(id INTEGER4,sample REAL DEFAULT 0.1)"));
    assert(sql.exec("INSERT INTO samples VALUES(1,1.5),(2,NULL)"));
    assert(sql.exec("CREATE TABLE _la_table_schema(table_name TEXT PRIMARY KEY,definition TEXT NOT NULL)"));
    const QJsonArray metadata{QJsonObject{{"name","id"},{"source_type","INT"}},QJsonObject{{"name","sample"},{"source_type","SINGLE"}}};
    sql.prepare("INSERT INTO _la_table_schema VALUES('samples',?)"); sql.addBindValue(QJsonDocument(metadata).toJson(QJsonDocument::Compact)); assert(sql.exec()); sql.finish();
    auto rows=linaccess::makeDaoDatabase(db)->member("OpenRecordset",{VbaValue::string("samples")}).object;
    const auto field=rows->member("Fields",{VbaValue::string("sample")}).object;
    assert(field->member("Type",{}).asNumber()==6 && field->member("Value",{}).numericType==VbaValue::NumericType::Single);
    const auto error=[](const std::function<void()>& operation,int expected) { int number=0; try { operation(); } catch(const linaccess::VbaRuntimeError& failure) { number=failure.number; } assert(number==expected); };
    rows->member("Edit",{}); rows->assignItem("sample",VbaValue::numeric(16777217.0));
    assert(rows->item("sample").asNumber()==16777216.0 && rows->item("sample").numericType==VbaValue::NumericType::Single);
    error([&] { rows->assignItem("sample",VbaValue::numeric(1e40)); },6);
    error([&] { rows->assignItem("sample",VbaValue::string("invalid")); },13);
    error([&] { rows->assignItem("sample",VbaValue::numeric(1e-50)); },3251);
    assert(rows->item("sample").asNumber()==16777216.0);
    rows->member("Update",{});
    assert(sql.exec("SELECT sample FROM samples WHERE id=1") && sql.next() && sql.value(0).toDouble()==16777216.0); sql.finish();
    auto clone=rows->member("Clone",{}).object; clone->member("MoveFirst",{}); assert(clone->item("sample").numericType==VbaValue::NumericType::Single); clone->member("Close",{});
    rows->member("MoveNext",{}); assert(rows->item("sample").kind==VbaValue::Kind::Null);
    rows->member("Edit",{}); rows->assignItem("sample",VbaValue::empty()); assert(rows->item("sample").asNumber()==0 && rows->item("sample").numericType==VbaValue::NumericType::Single);
    rows->assignItem("sample",VbaValue{}); assert(rows->item("sample").kind==VbaValue::Kind::Null); rows->member("CancelUpdate",{});
    rows->member("AddNew",{}); assert(rows->item("sample").numericType==VbaValue::NumericType::Single && rows->item("sample").asNumber()==static_cast<double>(0.1f)); rows->member("CancelUpdate",{});
    rows->member("Close",{});
    Host host(db); const auto result=linaccess::executeVbaProcedure(R"VBA(Sub InspectSingles()
Dim rows As DAO.Recordset
Set rows=CurrentDb.OpenRecordset("samples")
If VarType(rows!sample)<>vbSingle Then Err.Raise 5
rows.Edit
rows!sample=0.1
If VarType(rows!sample)<>vbSingle Or rows!sample<>CSng(0.1) Then Err.Raise 5
rows.Update
rows.Requery
If VarType(rows!sample)<>vbSingle Or rows!sample<>CSng(0.1) Then Err.Raise 5
rows.MoveNext
If Not IsNull(rows!sample) Then Err.Raise 5
rows.Close
Me!SingleChecked="yes"
End Sub)VBA","InspectSingles",host);
    if(result.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Single VBA error="<<result.errorNumber<<", line="<<result.line<<'\n';
    assert(result.status==linaccess::VbaExecutionStatus::Completed && host.fields["SingleChecked"]=="yes");
    assert(sql.exec("CREATE TABLE ordinary_double(value REAL)")); assert(sql.exec("INSERT INTO ordinary_double VALUES(0.1)"));
    auto ordinary=linaccess::makeDaoDatabase(db)->member("OpenRecordset",{VbaValue::string("ordinary_double")}).object;
    assert(ordinary->item("value").numericType==VbaValue::NumericType::Double); ordinary->member("Close",{});
    assert(sql.exec("UPDATE samples SET sample=1e40 WHERE id=1"));
    auto invalid=linaccess::makeDaoDatabase(db)->member("OpenRecordset",{VbaValue::string("samples")}).object;
    error([&] { invalid->item("sample"); },6); invalid->member("Close",{}); db.close();
}

void testDaoRelations() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","dao-relations"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE primary_rows(primary_a INTEGER,primary_b INTEGER,PRIMARY KEY(primary_a,primary_b))"));
    assert(sql.exec("CREATE TABLE foreign_rows(foreign_a INTEGER,foreign_b INTEGER,FOREIGN KEY(foreign_a,foreign_b) REFERENCES primary_rows(primary_a,primary_b) ON UPDATE CASCADE)"));
    assert(sql.exec("CREATE TABLE _la_access_relations(source_file TEXT,name TEXT,flags INTEGER,ordinal INTEGER,child_table TEXT,child_column TEXT,parent_table TEXT,parent_column TEXT)"));
    assert(sql.exec("INSERT INTO _la_access_relations VALUES('technical','CompositeReference',256,0,'foreign_rows','foreign_a','primary_rows','primary_a'),('technical','CompositeReference',256,1,'foreign_rows','foreign_b','primary_rows','primary_b')"));
    const auto error=[](const std::function<void()>& operation,int expected) { int number=0; try { operation(); } catch(const linaccess::VbaRuntimeError& failure) { number=failure.number; } assert(number==expected); };
    auto database=linaccess::makeDaoDatabase(db);
    auto relations=database->member("Relations",{}).object;
    auto relation=relations->member("Item",{VbaValue::string("compositereference")}).object;
    auto fields=relation->member("Fields",{}).object;
    auto field=fields->member("Item",{VbaValue::string("PRIMARY_A")}).object;
    assert(relation==database->member("Relations",{VbaValue::numeric(0)}).object);
    assert(relations->member("Count",{}).asNumber()==1 && relations->member("Count",{}).numericType==VbaValue::NumericType::Integer);
    assert(fields->member("Count",{}).asNumber()==2 && field->member("Name",{}).asText()=="primary_a" && field->member("ForeignName",{}).asText()=="foreign_a");
    assert(field==fields->enumerate().front().object && relation==relations->enumerate().front().object);
    error([&] { relations->member("Item",{VbaValue{}}); },94);
    error([&] { relations->member("Item",{VbaValue::numeric(1)}); },3265);
    error([&] { relations->member("Refresh",{VbaValue::numeric(1)}); },450);
    error([&] { relation->assign("Name",VbaValue::string("renamed")); },383);
    error([&] { field->assign("ForeignName",VbaValue::string("other")); },383);
    error([&] { field->member("Type",{}); },3251);
    Host host(db);
    const auto result=linaccess::executeVbaProcedure(R"VBA(Sub InspectRelations()
Dim relations As DAO.Relations, relation As DAO.Relation, fields As DAO.Fields, field As DAO.Field
Dim database As DAO.Database, total As Long
Set database=CurrentDb
Set relations=database.Relations
If Not TypeOf relations Is DAO.Relations Then Err.Raise 5
If VarType(relations.Count)<>vbInteger Or relations.Count<>1 Then Err.Raise 5
Set relation=database.Relations(Index:="CompositeReference")
If Not TypeOf relation Is DAO.Relation Then Err.Raise 5
If Not (relation Is relations.Item(Index:=0)) Then Err.Raise 5
If relation.Table<>"primary_rows" Or relation.ForeignTable<>"foreign_rows" Then Err.Raise 5
If VarType(relation.Attributes)<>vbLong Or relation.Attributes<>dbRelationUpdateCascade Then Err.Raise 5
Set fields=relation.Fields
Set field=relation.Fields(Index:="primary_b")
If field.Name<>"primary_b" Or field.ForeignName<>"foreign_b" Then Err.Raise 5
For Each relation In relations
For Each field In relation.Fields
total=total+1
Next
Next
If total<>2 Then Err.Raise 5
relations.Refresh
fields.Refresh
Me!RelationsChecked="yes"
End Sub)VBA","InspectRelations",host);
    if(result.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Relation VBA error="<<result.errorNumber<<", line="<<result.line<<'\n';
    assert(result.status==linaccess::VbaExecutionStatus::Completed && host.fields["RelationsChecked"]=="yes");
    assert(sql.exec("CREATE TEMP TABLE _la_access_relations(bad)"));
    assert(relations->member("Count",{}).asNumber()==1 && field->member("Name",{}).asText()=="primary_a");
    assert(sql.exec("UPDATE main._la_access_relations SET source_file='changed'"));
    error([&] { field->member("Name",{}); },3420);
    relation=relations->member("Item",{VbaValue::numeric(0)}).object;
    auto engine=linaccess::makeDaoEngine(db); engine->member("BeginTrans",{});
    assert(sql.exec("UPDATE main._la_access_relations SET source_file='transient'"));
    auto transient=relations->member("Item",{VbaValue::numeric(0)}).object;
    engine->member("Rollback",{});
    error([&] { transient->member("Name",{}); },3420);
    error([&] { relation->member("Name",{}); },3420);
    relation=relations->member("Item",{VbaValue::numeric(0)}).object;
    assert(sql.exec("CREATE TABLE unknown_foreign(x INTEGER,y INTEGER,FOREIGN KEY(x,y) REFERENCES primary_rows(primary_a,primary_b))"));
    error([&] { relation->member("Name",{}); },3420);
    error([&] { relations->member("Count",{}); },3251);
    assert(sql.exec("DROP TABLE unknown_foreign"));
    assert(relations->member("Count",{}).asNumber()==1);
    assert(sql.exec("UPDATE main._la_access_relations SET flags=2"));
    error([&] { relations->member("Count",{}); },3251); // Enforced native key cannot silently become unenforced.
    assert(sql.exec("DROP TABLE foreign_rows"));
    assert(sql.exec("CREATE TABLE foreign_rows(foreign_a INTEGER,foreign_b INTEGER)"));
    assert(relations->member("Count",{}).asNumber()==1); // Genuine imported don't-enforce metadata remains visible.
    assert(sql.exec("UPDATE main._la_access_relations SET ordinal=3 WHERE ordinal=1"));
    error([&] { relations->member("Count",{}); },3251);
    db.close(); error([&] { relations->member("Count",{}); },3420);
}

void testTableDefSchemaFields() {
    auto db=QSqlDatabase::addDatabase("QSQLITE","tabledef-schema-fields"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE schema_rows(id INTEGER4,amount CURRENCY_REAL,payload BLOB)"));
    assert(sql.exec("CREATE TABLE source_rows(label TEXT,amount REAL,memo TEXT,enabled INTEGER)"));
    assert(sql.exec("CREATE TABLE _la_table_schema(table_name TEXT PRIMARY KEY,definition TEXT NOT NULL)"));
    const QJsonArray sourceFields{
        QJsonObject{{"name","label"},{"source_type","VARCHAR"}},QJsonObject{{"name","amount"},{"source_type","DOUBLE"}},
        QJsonObject{{"name","memo"},{"source_type","MEMO"}},QJsonObject{{"name","enabled"},{"source_type","BOOLEAN"}}};
    sql.prepare("INSERT INTO _la_table_schema VALUES('source_rows',?)"); sql.addBindValue(QJsonDocument(sourceFields).toJson(QJsonDocument::Compact)); assert(sql.exec()); sql.finish();
    Host host(db);
    const auto result=linaccess::executeVbaProcedure(R"VBA(Sub InspectSchema()
Dim table As DAO.TableDef, fields As DAO.Fields, field As DAO.Field, count As Long, positions As Long
Set table=CurrentDb.TableDefs("schema_rows")
Set fields=table.Fields
If Not TypeOf fields Is DAO.Fields Then Err.Raise 5
If fields.Count<>3 Then Err.Raise 5
Set field=fields.Item(Index:="amount")
If Not TypeOf field Is DAO.Field Then Err.Raise 5
If field.Type<>5 Or field.OrdinalPosition<>1 Then Err.Raise 5
If field.Size<>8 Or VarType(field.Size)<>vbLong Then Err.Raise 5
If table.Fields("payload").Size<>0 Then Err.Raise 5
If table.Fields("id").Size<>4 Then Err.Raise 5
If table.Fields(0).Name<>"id" Then Err.Raise 5
For Each field In fields
count=count+1
positions=positions+field.OrdinalPosition
Next
If count<>3 Or positions<>3 Then Err.Raise 5
Dim database As DAO.Database, definitions As DAO.TableDefs, query As DAO.QueryDef
Set database=CurrentDb
Set definitions=database.TableDefs
Set table=database.TableDefs(Index:="schema_rows")
If Not (table Is definitions.Item(Index:="schema_rows")) Then Err.Raise 5
If table.RecordCount<>0 Or VarType(table.RecordCount)<>vbLong Then Err.Raise 5
definitions.Refresh
If Not (table Is definitions.Item(Index:="schema_rows")) Then Err.Raise 5
Set query=database.CreateQueryDef("","SELECT id,amount,payload FROM schema_rows")
Set fields=query.Fields
If Not TypeOf fields Is DAO.Fields Then Err.Raise 5
Set field=query.Fields(Index:="amount")
If Not TypeOf field Is DAO.Field Then Err.Raise 5
If field.Type<>dbCurrency Or field.Size<>8 Then Err.Raise 5
If fields.Item(Index:=0).Name<>"id" Then Err.Raise 5
count=0
For Each field In fields
count=count+1
Next
If count<>3 Then Err.Raise 5
fields.Refresh
query.Close
Set fields=table.Fields
fields.Refresh
Me!SchemaFields="checked"
End Sub
)VBA","InspectSchema",host);
    assert(result.status==linaccess::VbaExecutionStatus::Completed && host.fields["SchemaFields"]=="checked");
    const auto origins=host.currentDatabase()->member("TableDefs",{linaccess::VbaValue::string("source_rows")}).object->member("Fields",{}).object;
    const int expectedTypes[]{10,7,12,1};
    for(int i=0;i<4;++i) assert(origins->member("Item",{linaccess::VbaValue::longInteger(i)}).object->member("Type",{}).number==expectedTypes[i]);
    assert(origins->member("Item",{linaccess::VbaValue::longInteger(1)}).object->member("Size",{}).number==8);
    assert(origins->member("Item",{linaccess::VbaValue::longInteger(2)}).object->member("Size",{}).number==0);
    bool unknownSize=false; try { origins->member("Item",{linaccess::VbaValue::longInteger(0)}).object->member("Size",{}); } catch(const linaccess::VbaRuntimeError& error) { unknownSize=error.number==3251; } assert(unknownSize);
    const auto emptyRows=host.currentDatabase()->member("OpenRecordset",{linaccess::VbaValue::string("schema_rows")}).object;
    assert(emptyRows->member("Fields",{linaccess::VbaValue::string("payload")}).object->member("Size",{}).number==0);
    assert(emptyRows->member("Fields",{linaccess::VbaValue::string("id")}).object->member("Size",{}).number==4); emptyRows->member("Close",{});
    auto table=host.currentDatabase()->member("TableDefs",{linaccess::VbaValue::string("schema_rows")}).object;
    auto fields=table->member("Fields",{}).object;
    auto field=fields->member("Item",{linaccess::VbaValue::string("id")}).object;
    assert(field->member("Type",{}).number==4);
    assert(sql.exec("ALTER TABLE schema_rows ADD COLUMN note TEXT"));
    bool stale=false;
    try { fields->member("Count",{}); } catch(const linaccess::VbaRuntimeError& error) { stale=error.number==3420; }
    assert(stale);
    stale=false; try { field->member("Name",{}); } catch(const linaccess::VbaRuntimeError& error) { stale=error.number==3420; } assert(stale);
    auto fresh=table->member("Fields",{}).object; assert(fresh->member("Count",{}).number==4);
    assert(sql.exec("CREATE TEMP VIEW _la_table_schema AS SELECT 'schema_rows' AS table_name,'[]' AS definition"));
    bool shadowed=false; try { fresh->member("Count",{}); } catch(const linaccess::VbaRuntimeError& error) { shadowed=error.number==3251; } assert(shadowed);
    assert(sql.exec("DROP VIEW temp._la_table_schema")); assert(fresh->member("Count",{}).number==4);

    db.close(); stale=false; try { fresh->member("Count",{}); } catch(const linaccess::VbaRuntimeError& error) { stale=error.number==3420; } assert(stale);
}
void testDaoTableGenerations() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","dao-table-generations"); db.setDatabaseName(":memory:"); assert(db.open());
    const auto database=linaccess::makeDaoDatabase(db);
    const auto execute=[&](const char* source) { database->member("Execute",{VbaValue::string(source)}); };
    const auto error=[&](const std::function<void()>& operation,int expected) { int number=0; try { operation(); } catch(const linaccess::VbaRuntimeError& failure) { number=failure.number; } assert(number==expected); };
    execute("CREATE TABLE epoch_rows(id INTEGER PRIMARY KEY,payload INTEGER4)"); execute("INSERT INTO epoch_rows VALUES(1,11)");
    const auto definitions=database->member("TableDefs",{}).object;
    const auto previous=definitions->item("epoch_rows").object;
    const auto schema=previous->member("Fields",{}).object;
    const auto rows=previous->member("OpenRecordset",{}).object;
    const auto field=rows->member("Fields",{VbaValue::string("payload")}).object;
    const auto clone=rows->member("Clone",{}).object;
    rows->member("Edit",{}); field->assign("Value",VbaValue::integer(777));
    QSqlQuery sql(db); assert(sql.exec("SELECT sql,rootpage FROM main.sqlite_schema WHERE name='epoch_rows'") && sql.next());
    const auto definition=sql.value(0).toString(); const auto root=sql.value(1).toLongLong(); sql.finish();
    execute("CREATE TABLE IF NOT EXISTS epoch_rows(id INTEGER PRIMARY KEY,payload INTEGER4)");
    error([&] { execute("CREATE TABLE epoch_rows(id INTEGER PRIMARY KEY,payload INTEGER4)"); },3010);
    assert(previous==definitions->item("epoch_rows").object && previous->member("Name",{}).asText()=="epoch_rows");
    execute("DROP\nTABLE [epoch_rows]"); execute("CREATE TABLE epoch_rows(id INTEGER PRIMARY KEY,payload INTEGER4)");
    execute("INSERT INTO epoch_rows VALUES(1,99)");
    assert(sql.exec("SELECT sql,rootpage FROM main.sqlite_schema WHERE name='epoch_rows'") && sql.next());
    assert(sql.value(0).toString()==definition && sql.value(1).toLongLong()==root); sql.finish();
    error([&] { (void)previous->member("Name",{}); },3420);
    error([&] { (void)schema->member("Count",{}); },3420);
    error([&] { rows->member("Update",{}); },3420);
    error([&] { rows->member("Requery",{}); },3420);
    error([&] { clone->member("Edit",{}); },3420);
    error([&] { (void)field->member("Name",{}); },3420);
    clone->member("Close",{}); rows->member("Close",{});
    assert(sql.exec("SELECT payload FROM epoch_rows") && sql.next() && sql.value(0).toInt()==99); sql.finish();
    const auto current=definitions->item("epoch_rows").object; assert(current!=previous);
    const auto fresh=current->member("OpenRecordset",{}).object;
    execute("ALTER TABLE epoch_rows ADD COLUMN extra INTEGER");
    assert(current==definitions->item("epoch_rows").object); fresh->member("Requery",{});
    assert(fresh->member("Fields",{}).object->member("Count",{}).asNumber()==3);
    const auto workspace=linaccess::makeDaoEngine(db)->defaultMember({VbaValue::integer(0)},true).object;
    const auto drop=database->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string("DROP TABLE epoch_rows")}).object;
    workspace->member("BeginTrans",{}); drop->member("Execute",{}); workspace->member("Rollback",{});
    error([&] { (void)current->member("Name",{}); },3420); error([&] { fresh->member("Requery",{}); },3420); fresh->member("Close",{});
    const auto restored=definitions->item("epoch_rows").object;
    assert(restored!=current && restored->member("RecordCount",{}).asNumber()==1);
    assert(sql.exec("SELECT payload FROM epoch_rows") && sql.next() && sql.value(0).toInt()==99); sql.finish(); db.close();
}
void testSqlRgbDao() {
    auto db=QSqlDatabase::addDatabase("QSQLITE","sql-rgb-dao"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE color_rows(id INTEGER PRIMARY KEY,r REAL,g REAL,b REAL,packed INTEGER)"));
    assert(sql.exec("INSERT INTO color_rows VALUES(1,255,0,0,0),(2,0,255,0,0)"));
    assert(sql.exec("CREATE TABLE color_values(value INTEGER)"));
    Host host(db);
    {
        assert(sql.exec("CREATE TABLE depreciation_rows(period INTEGER,cost REAL,salvage REAL,life REAL)"));
        assert(sql.exec("INSERT INTO depreciation_rows VALUES(1,10000,1000,5),(2,10000,1000,5),(3,10000,1000,5),(4,10000,1000,5),(5,10000,1000,5)"));
        const auto definition=host.currentDatabase()->member("CreateQueryDef",{
            linaccess::VbaValue::string(""),linaccess::VbaValue::string("SELECT SLN(cost,salvage,life) AS straight,SYD(cost,salvage,life,period) AS accelerated,DDB(cost,salvage,life,period) AS declining FROM depreciation_rows ORDER BY period")}).object;
        const auto rows=definition->member("OpenRecordset",{}).object;
        const auto fields=rows->member("Fields",{}).object;
        double total=0,decliningBalance=10000;
        for(int i=0;i<5;++i) {
            assert(fields->member("Item",{linaccess::VbaValue::string("straight")}).object->member("Value",{}).asNumber()==1800);
            const auto depreciation=fields->member("Item",{linaccess::VbaValue::string("accelerated")}).object->member("Value",{}).asNumber();
            assert(depreciation==3000-i*600); total+=depreciation;
            const auto expectedDecline=std::min(decliningBalance*0.4,decliningBalance-1000);
            const auto decline=fields->member("Item",{linaccess::VbaValue::string("declining")}).object->member("Value",{}).asNumber();
            assert(std::abs(decline-expectedDecline)<1e-9); decliningBalance-=expectedDecline;
            rows->member("MoveNext",{});
        }
        assert(total==9000 && decliningBalance==1000 && rows->member("EOF",{}).truthy()); rows->member("Close",{});
    }
    {
        assert(sql.exec("CREATE TABLE financial_rows(id INTEGER PRIMARY KEY,rate REAL,periods INTEGER,payment REAL,principal REAL)"));
        assert(sql.exec("INSERT INTO financial_rows VALUES(1,0.01,12,-100,1000),(2,0,12,-100,1200)"));
        const auto definition=host.currentDatabase()->member("CreateQueryDef",{
            linaccess::VbaValue::string(""),linaccess::VbaValue::string("SELECT FV(rate,periods,payment) AS future,Pmt(rate,periods,principal) AS due,PV(rate,periods,payment) AS present,NPer(rate,Pmt(rate,periods,principal),principal) AS duration,Rate(periods,Pmt(rate,periods,principal),principal) AS effective_rate,IPmt(rate,1,periods,principal) AS interest,PPmt(rate,1,periods,principal) AS principal_paid FROM financial_rows ORDER BY id")}).object;
        const auto rows=definition->member("OpenRecordset",{}).object;
        const auto fields=rows->member("Fields",{}).object;
        const auto value=[&](const char* name) { return fields->member("Item",{linaccess::VbaValue::string(name)}).object->member("Value",{}).asNumber(); };
        double expected=0; for(int i=0;i<12;++i) expected=expected*1.01+100;
        assert(std::abs(value("future")-expected)<1e-9);
        double discounted=0,discountFactor=1; for(int i=0;i<12;++i) { discountFactor*=1.01; discounted+=100/discountFactor; }
        assert(std::abs(value("present")-discounted)<1e-9);
        double loan=1000; const auto due=value("due"); for(int i=0;i<12;++i) loan=loan*1.01+due;
        assert(std::abs(loan)<1e-9 && std::abs(value("duration")-12)<1e-9 && std::abs(value("effective_rate")-0.01)<1e-11);
        assert(value("interest")==-10 && std::abs(value("interest")+value("principal_paid")-due)<1e-9);
        rows->member("MoveNext",{});
        assert(value("future")==1200 && value("due")==-100);
        assert(value("present")==1200 && value("duration")==12 && std::abs(value("effective_rate"))<1e-11);
        assert(value("interest")==0 && value("principal_paid")==-100);
        rows->member("Close",{});
    }
    {
        assert(sql.exec("CREATE TABLE partition_rows(value INTEGER)"));
        assert(sql.exec("INSERT INTO partition_rows VALUES(-1),(0),(4),(5),(99),(100),(NULL)"));
        const auto definition=host.currentDatabase()->member("CreateQueryDef",{
            linaccess::VbaValue::string(""),linaccess::VbaValue::string("SELECT Partition(value,0,99,5) AS bucket,Count(*) AS n FROM partition_rows GROUP BY Partition(value,0,99,5) ORDER BY bucket")}).object;
        const auto rows=definition->member("OpenRecordset",{}).object;
        const std::vector<std::string> expected{"","   : -1","  0:  4","  5:  9"," 95: 99","100:   "};
        for (std::size_t i=0;i<expected.size();++i) {
            assert(!rows->member("EOF",{}).truthy());
            const auto fields=rows->member("Fields",{}).object;
            const auto bucket=fields->member("Item",{linaccess::VbaValue::string("bucket")}).object->member("Value",{});
            assert(i==0 ? bucket.kind==linaccess::VbaValue::Kind::Null : bucket.asText()==expected[i]);
            assert(fields->member("Item",{linaccess::VbaValue::string("n")}).object->member("Value",{}).asNumber()==(i==2 ? 2 : 1));
            rows->member("MoveNext",{});
        }
        assert(rows->member("EOF",{}).truthy());
        rows->member("Close",{});
    }
    const std::string retainedDao=R"VBA(Private database As DAO.Database
Public rows As DAO.Recordset
Sub ReadFieldArgument(ByVal value As DAO.Field)
Me!DaoObjectArgument=CStr(value.Value)
End Sub
Function ReadFieldFunction(ByVal value As DAO.Field) As Long
ReadFieldFunction=value.Value
End Function
Sub InitializeDao()
Set database=CurrentDb
Set rows=database.OpenRecordset("SELECT id FROM color_rows ORDER BY id")
If Not TypeOf database Is DAO.Database Then Err.Raise 5
If Not TypeOf rows Is DAO.Recordset Then Err.Raise 5
If Not TypeOf rows.Fields Is DAO.Fields Then Err.Raise 5
If Not TypeOf rows!id Is DAO.Field Then Err.Raise 5
ReadFieldArgument rows!id
Me!DaoObjectFunction=ReadFieldFunction(rows!id)
If TypeOf rows!id Is DAO.Recordset Then Err.Raise 5
If Not TypeOf rows.Clone Is DAO.Recordset Then Err.Raise 5
DBEngine.BeginTrans
DBEngine.CommitTrans
Me!DaoTypes=TypeName(database) & ":" & TypeName(rows) & ":" & TypeName(rows.Fields)
End Sub
Sub AdvanceDao()
rows.MoveNext
Me!DaoRetained=CStr(rows!id)
End Sub
Function ReturnDatabase() As DAO.Database
Set ReturnDatabase=database
End Function
Function OptionalDatabase(Optional ByVal value As DAO.Database=Nothing) As Boolean
OptionalDatabase=value Is Nothing
End Function
Sub CheckDao()
Dim returned As DAO.Database
Set returned=ReturnDatabase()
Me!DaoReturned=CStr(returned Is database)
Me!DaoOptional=CStr(OptionalDatabase()) & ":" & CStr(OptionalDatabase(database))
On Error Resume Next
Set rows=database
Me!DaoWrongType=CStr(Err.Number)
Me!DaoUnchanged=TypeName(rows) & ":" & CStr(rows!id)
End Sub
Sub ReleaseDao()
rows.Close
Set rows=Nothing
Set database=Nothing
End Sub
)VBA";
    const auto initializedDao=linaccess::executeVbaProcedure(retainedDao,"InitializeDao",host);
    if(initializedDao.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"TypeOf DAO error="<<initializedDao.errorNumber<<" line="<<initializedDao.line<<" explanation="<<initializedDao.explanation<<"\n";
    assert(initializedDao.status==linaccess::VbaExecutionStatus::Completed);
    assert(host.fields["DaoObjectArgument"]=="1" && host.fields["DaoObjectFunction"]=="1");
    assert(host.fields["DaoTypes"]=="Database:Recordset:Fields");
    assert(linaccess::executeVbaProcedure(retainedDao,"AdvanceDao",host).status==linaccess::VbaExecutionStatus::Completed && host.fields["DaoRetained"]=="2");
    const auto checkedDao=linaccess::executeVbaProcedure(retainedDao,"CheckDao",host);
    if(checkedDao.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Typed DAO error="<<checkedDao.errorNumber<<" line="<<checkedDao.line<<"\n";
    assert(checkedDao.status==linaccess::VbaExecutionStatus::Completed);
    assert(host.fields["DaoReturned"]=="True" && host.fields["DaoOptional"]=="True:False");
    assert(host.fields["DaoWrongType"]=="13" && host.fields["DaoUnchanged"]=="Recordset:2");
    assert(linaccess::executeVbaProcedure(retainedDao,"ReleaseDao",host).status==linaccess::VbaExecutionStatus::Completed);

    assert(sql.exec("CREATE TABLE native_indexed_fields(id INTEGER PRIMARY KEY,value TEXT)"));
    assert(sql.exec("INSERT INTO native_indexed_fields VALUES(1,'before')"));
    const auto indexedFields=linaccess::executeVbaProcedure(R"VBA(Sub IndexedFields()
Dim rows As DAO.Recordset, check As DAO.Recordset
Dim fields As Object, genericField As Object
Dim field As DAO.Field
Set rows=CurrentDb.OpenRecordset("native_indexed_fields")
rows.Edit
rows("value")="indexed"
Set fields=rows.Fields
If Not (fields Is rows.Fields) Then Err.Raise 5
Set genericField=fields.Item(Index:="VALUE")
If genericField.Name<>"value" Then Err.Raise 5
If Not (genericField Is rows.Fields(Index:="value")) Then Err.Raise 5
fields("value")="fields"
rows.Fields("value")="indexed-member"
Set field=rows.Fields("value")
Let field="typed-default"
Set genericField=field
genericField="final"
Me!DefaultFieldIdentity=CStr(field Is genericField)
Me!IndexedPending=rows("value")
Set check=CurrentDb.OpenRecordset("SELECT value FROM native_indexed_fields")
Me!IndexedBeforeUpdate=check!value
check.Close
rows.Update
Me!IndexedCommitted=rows("value")
rows.Close
End Sub
)VBA","IndexedFields",host);
    if(indexedFields.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Indexed DAO error="<<indexedFields.errorNumber<<" line="<<indexedFields.line<<"\n";
    assert(host.fields["DefaultFieldIdentity"]=="True");
    assert(indexedFields.status==linaccess::VbaExecutionStatus::Completed && host.fields["IndexedPending"]=="final" && host.fields["IndexedBeforeUpdate"]=="before" && host.fields["IndexedCommitted"]=="final");
    assert(sql.exec("SELECT value FROM native_indexed_fields WHERE id=1") && sql.next() && sql.value(0).toString()=="final"); sql.finish();

    assert(sql.exec("CREATE TABLE native_dictionary_values(value TEXT)"));
    const auto dictionarySql=linaccess::executeVbaProcedure(R"VBA(Sub DictionarySql()
Dim d As Object
Dim database As DAO.Database
Dim rows As DAO.Recordset
Dim batch As Variant
Set d=CreateObject("Scripting.Dictionary")
Set d("database")=CurrentDb
Set database=d("database")
database.Execute "INSERT INTO native_dictionary_values VALUES('native')"
Set rows=database.OpenRecordset("SELECT value FROM native_dictionary_values")
d("matrix")=rows.GetRows(1)
batch=d("matrix")
Me!DictionarySqlValue=batch(0,0)
rows.Close
End Sub
)VBA","DictionarySql",host);
    if(dictionarySql.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Dictionary SQL error="<<dictionarySql.errorNumber<<" line="<<dictionarySql.line<<"\n";
    assert(dictionarySql.status==linaccess::VbaExecutionStatus::Completed && host.fields["DictionarySqlValue"]=="native");
    const auto resumedWith=linaccess::executeVbaProcedure(R"VBA(Sub ResumedWith()
Dim source As DAO.Recordset,columnName As String,handled As Long
Set source=CurrentDb.OpenRecordset("native_indexed_fields")
columnName="missing"
On Error GoTo caught
With source.Clone
.MoveFirst
Me!ResumedCloneValue=.Fields(columnName).Value
On Error Resume Next
DBEngine.BeginTrans
Me!ResumedCloneLive=CStr(Err.Number)
On Error GoTo caught
On 1 GoTo releasedClone
End With
releasedClone:
DBEngine.BeginTrans
DBEngine.Rollback
Me!ResumedCloneReleased="yes"
Me!ResumedCloneHandled=CStr(handled)
source.Close
Exit Sub
caught:
handled=handled+1
If handled>1 Then Err.Raise Err.Number
columnName="value"
Resume
End Sub
)VBA","ResumedWith",host);
    if(resumedWith.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Resumed clone error="<<resumedWith.errorNumber<<" line="<<resumedWith.line<<"\n";
    assert(resumedWith.status==linaccess::VbaExecutionStatus::Completed && host.fields["ResumedCloneValue"]=="final" && host.fields["ResumedCloneHandled"]=="1");
    assert(host.fields["ResumedCloneLive"]=="3251" && host.fields["ResumedCloneReleased"]=="yes");

    const auto arrayReferenceLifetime=linaccess::executeVbaProcedure(R"VBA(Private holder As Variant
Sub ArrayReferenceLifetime()
Dim rows As DAO.Recordset,copyRows As DAO.Recordset
Set rows=CurrentDb.OpenRecordset("native_indexed_fields")
Set copyRows=rows.Clone
holder=Array(copyRows)
Set copyRows=Nothing
ForwardArrayClone holder(0)
rows.Close
End Sub
Sub ForwardArrayClone(ByRef value As Object)
ReleaseArrayClone value,value
End Sub
Sub ReleaseArrayClone(ByVal copied As Object,ByRef value As Object)
Set copied=Nothing
Set value=Nothing
DBEngine.BeginTrans
DBEngine.Rollback
Me!ArrayCloneReleased="yes"
End Sub
)VBA","ArrayReferenceLifetime",host);
    if(arrayReferenceLifetime.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Array clone error="<<arrayReferenceLifetime.errorNumber<<" line="<<arrayReferenceLifetime.line<<"\n";
    assert(arrayReferenceLifetime.status==linaccess::VbaExecutionStatus::Completed && host.fields["ArrayCloneReleased"]=="yes");
    const auto dictionaryLifetime=linaccess::executeVbaProcedure(R"VBA(Sub DictionaryLifetime()
Dim d As Object, key As Variant
Dim rows As DAO.Recordset, copyRows As DAO.Recordset
Set rows=CurrentDb.OpenRecordset("native_indexed_fields")
Set copyRows=rows.Clone
Set d=CreateObject("Scripting.Dictionary")
Set d("clone")=copyRows
Set copyRows=Nothing
For Each key In d
Set d=Nothing
On Error Resume Next
DBEngine.BeginTrans
Me!DictionaryHeldClone=CStr(Err.Number)
On Error GoTo 0
Exit For
Next
DBEngine.BeginTrans
DBEngine.Rollback
rows.Close
End Sub
)VBA","DictionaryLifetime",host);
    if(dictionaryLifetime.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Dictionary lifetime error="<<dictionaryLifetime.errorNumber<<" line="<<dictionaryLifetime.line<<"\n";
    assert(dictionaryLifetime.status==linaccess::VbaExecutionStatus::Completed && host.fields["DictionaryHeldClone"]=="3251");


    {
        QTemporaryFile file; assert(file.open() && file.write("abc",3)==3 && file.flush());
        auto path=file.fileName(); path.replace(QChar(39),"''");
        const auto source="SELECT FileLen('"+path+"') AS bytes,FileDateTime('"+path+"') AS stamp";
        const auto rows=host.currentDatabase()->member("OpenRecordset",{linaccess::VbaValue::string(source.toStdString())}).object;
        const auto bytes=rows->member("Fields",{linaccess::VbaValue::string("bytes")}).object;
        const auto stamp=rows->member("Fields",{linaccess::VbaValue::string("stamp")}).object;
        assert(bytes->member("Type",{}).number==4 && bytes->member("Value",{}).number==3 && bytes->member("Value",{}).numericType==linaccess::VbaValue::NumericType::Long);
        assert(stamp->member("Type",{}).number==8 && stamp->member("Value",{}).kind==linaccess::VbaValue::Kind::Date);
        rows->member("Close",{});
    }

    {
        const auto savedDirectory=QDir::currentPath();
        QTemporaryDir directory; assert(directory.isValid());
        const auto plan=linaccess::compileAccessQuery("SELECT CurDir() AS cwd,CurDir('') AS optional_drive");
        QSqlQuery current(db); assert(current.prepare(plan.sql));
        assert(current.exec() && current.next() && current.value(0).toString()==savedDirectory);
        current.finish();
        assert(QDir::setCurrent(directory.path()));
        assert(current.exec() && current.next());
        assert(current.value(0).toString()==QDir::currentPath() && current.value(1).toString()==QDir::currentPath());
        current.finish();
        const auto rows=host.currentDatabase()->member("OpenRecordset",{linaccess::VbaValue::string("SELECT CurDir() AS cwd")}).object;
        assert(rows->member("Fields",{linaccess::VbaValue::string("cwd")}).object->member("Value",{}).asText()==QDir::currentPath().toStdString());
        rows->member("Close",{});
        assert(QDir::setCurrent(savedDirectory));
    }

    {
        using linaccess::VbaValue;
        const auto definition=host.currentDatabase()->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string("SELECT 1 AS value")}).object;
        const std::vector<std::pair<std::string,int>> kinds{
            {"PARAMETERS p Long; SELECT p AS value",0},
            {"SELECT 'UNION UPDATE' AS marker",0},
            {"SELECT 1 AS value UNION ALL SELECT 2 AS value",128},
            {"SELECT * FROM (SELECT 1 AS value UNION ALL SELECT 2 AS value) AS nested",0},
            {"SELECT TOP 1 id FROM items ORDER BY id",0},
            {"UPDATE items AS a INNER JOIN items AS b ON a.id=b.id SET a.qty=b.qty",48}
        };
        for(const auto& [source,expected]:kinds) {
            definition->assign("SQL",VbaValue::string(source));
            const auto type=definition->member("Type",{});
            assert(type.asNumber()==expected && type.numericType==VbaValue::NumericType::Integer);
        }
        const auto action=[&](const std::string& source,int expected) {
            definition->assign("SQL",VbaValue::string(source)); assert(definition->member("Type",{}).asNumber()==expected);
            definition->member("Execute",{});
        };
        action("CREATE TABLE native_query_type_rows(id INTEGER4,value INTEGER4)",96);
        action("INSERT INTO native_query_type_rows VALUES(1,4)",64);
        action("UPDATE native_query_type_rows SET value=9 WHERE id=1",48);
        definition->assign("SQL",VbaValue::string("SELECT value FROM native_query_type_rows"));
        assert(definition->member("Type",{}).asNumber()==0);
        const auto rows=definition->member("OpenRecordset",{}).object;
        assert(rows->member("Fields",{VbaValue::string("value")}).object->member("Value",{}).asNumber()==9); rows->member("Close",{});
        action("SELECT id,value INTO native_query_type_copy FROM native_query_type_rows",80);
        action("DELETE FROM native_query_type_rows",32);
        action("ALTER TABLE native_query_type_copy ADD COLUMN extra INTEGER4",96);
        action("DROP TABLE native_query_type_copy",96); action("DROP TABLE native_query_type_rows",96);
        int error=0; try { definition->assign("Type",VbaValue::integer(0)); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
        assert(error==383); definition->member("Close",{});
    }

    const auto painted=linaccess::executeVbaProcedure(R"VBA(Sub PaintRows()
Dim database As DAO.Database
Dim rows As DAO.Recordset
Set database=CurrentDb
Set rows=database.OpenRecordset(Type:=4, Name:="SELECT id,RGB(r,g,b) AS shade FROM color_rows ORDER BY id")
Let rows.AbsolutePosition=0: Me!ShadeOrdinal=rows.Fields("shade").OrdinalPosition
Me!FirstShade=rows!shade
rows.FindFirst Criteria:="id=2"
Me!SecondShade=rows!shade
Me!Ordinal=rows.AbsolutePosition
rows.Close
database.Execute Options:=128, Query:="UPDATE color_rows SET packed=RGB(r,g,b)"
Me!ColorAffected=database.RecordsAffected
End Sub
)VBA","PaintRows",host);
    assert(painted.status==linaccess::VbaExecutionStatus::Completed);
    assert(host.fields["FirstShade"]=="255" && host.fields["SecondShade"]=="65280" && host.fields["ColorAffected"]=="2" && host.fields["Ordinal"]=="1" && host.fields["ShadeOrdinal"]=="1");
    assert(sql.exec("SELECT packed FROM color_rows ORDER BY id") && sql.next() && sql.value(0).toInt()==255 && sql.next() && sql.value(0).toInt()==65280); sql.finish();
    const auto collected=linaccess::executeVbaProcedure(R"VBA(Sub CollectColors()
Dim values As Collection
Dim references As Collection
Dim database As DAO.Database
Dim rows As DAO.Recordset
Dim value As Variant
Set values=New Collection
On Error Resume Next
Set values=CurrentDb
Me!CollectionClassError=Err.Number
On Error GoTo 0
Set references=New Collection
references.Add CurrentDb,"db"
Set database=references("DB")
Set rows=CurrentDb.OpenRecordset("SELECT packed FROM color_rows ORDER BY id",4)
Do While Not rows.EOF
values.Add rows!packed
rows.MoveNext
Loop
rows.Close
For Each value In values
database.Execute "INSERT INTO color_values(value) VALUES(" & CStr(value) & ")"
Next
End Sub
)VBA","CollectColors",host);
    assert(collected.status==linaccess::VbaExecutionStatus::Completed && host.fields["CollectionClassError"]=="13");
    assert(sql.exec("SELECT COUNT(*),SUM(value) FROM color_values") && sql.next() && sql.value(0).toInt()==2 && sql.value(1).toInt()==65535); sql.finish();
    assert(sql.exec("UPDATE color_rows SET packed=99")); assert(sql.exec("INSERT INTO color_rows VALUES(3,NULL,0,0,99)"));
    const auto rejected=linaccess::executeVbaProcedure(R"VBA(Sub InvalidColor()
Dim database As DAO.Database
Set database=CurrentDb
On Error Resume Next
database.Execute "UPDATE color_rows SET packed=RGB(r,g,b)"
Me!ColorError=CStr(Err.Number)
End Sub
)VBA","InvalidColor",host);
    assert(rejected.status==linaccess::VbaExecutionStatus::Completed && host.fields["ColorError"]=="94");
    assert(sql.exec("SELECT COUNT(*),MIN(packed),MAX(packed) FROM color_rows") && sql.next());
    assert(sql.value(0).toInt()==3 && sql.value(1).toInt()==99 && sql.value(2).toInt()==99); sql.finish();
}
void testBooleanQueryParameters() {
    auto db=QSqlDatabase::addDatabase("QSQLITE","boolean-parameters");
    db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE flags(flag INTEGER)"));
    assert(sql.exec("INSERT INTO flags VALUES(7)"));
    Host host(db);
    const auto converted=linaccess::executeVbaProcedure(R"VBA(Sub ConvertBoolean()
Dim definition As DAO.QueryDef
Set definition=CurrentDb.CreateQueryDef(SQLText:="PARAMETERS p BIT; UPDATE flags SET flag=[p]", Name:="")
definition.Parameters("p").Value=CBool("False")
definition.Execute
Me!FalseType=TypeName(CBool(0))
Me!TrueNumber=CStr(CLng(CBool(-2.5)))
On Error Resume Next
definition.Parameters("p").Value=CBool(Null)
Me!NullError=CStr(Err.Number)
definition.Parameters("p").Value=CBool("invalid")
Me!TypeError=CStr(Err.Number)
End Sub
)VBA","ConvertBoolean",host);
    assert(converted.status==linaccess::VbaExecutionStatus::Completed);
    assert(host.fields["FalseType"]=="Boolean" && host.fields["TrueNumber"]=="-1");
    assert(host.fields["NullError"]=="94" && host.fields["TypeError"]=="13");
    assert(sql.exec("SELECT flag FROM flags") && sql.next() && sql.value(0).toInt()==0); sql.finish();
    auto database=linaccess::makeDaoDatabase(db);
    using linaccess::VbaValue;
    for (const auto& alias : {"BIT","BOOLEAN","LOGICAL","LOGICAL1","YESNO"}) {
        const auto source=QString("PARAMETERS p %1; UPDATE flags SET flag=[p]").arg(alias);
        auto definition=database->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string(source.toStdString())}).object;
        auto parameter=definition->member("Parameters",{VbaValue::string("p")}).object;
        assert(parameter->member("Type",{}).asNumber()==1);
        for (const auto& item : {std::pair{VbaValue::string("False"),0},
                std::pair{VbaValue::string("TRUE"),-1},
                std::pair{VbaValue::string("  False  "),0},
                std::pair{VbaValue::numeric(0),0},
                std::pair{VbaValue::numeric(-2.5),-1}}) {
            parameter->assign("Value",item.first);
            definition->member("Execute",{});
            assert(sql.exec("SELECT flag FROM flags") && sql.next() && sql.value(0).toInt()==item.second); sql.finish();
        }
        parameter->assign("Value",VbaValue::string("not a Boolean"));
        bool rejected=false;
        try { definition->member("Execute",{}); }
        catch (const linaccess::VbaRuntimeError& error) { rejected=error.number==13; }
        assert(rejected);
        assert(sql.exec("SELECT flag FROM flags") && sql.next() && sql.value(0).toInt()==-1); sql.finish();
        const auto plan=linaccess::compileAccessQuery(source);
        for (const auto& invalid : {QVariant(QString("")),QVariant(std::numeric_limits<double>::infinity()),
                QVariant(std::numeric_limits<double>::quiet_NaN())}) {
            bool invalidRejected=false;
            try { linaccess::bindAccessQueryParameters(plan,{{"p",invalid}}); }
            catch (const linaccess::VbaRuntimeError& error) { invalidRejected=error.number==13; }
            assert(invalidRejected);
        }
        parameter->assign("Value",VbaValue{});
        definition->member("Execute",{});
        assert(sql.exec("SELECT flag FROM flags") && sql.next() && sql.value(0).isNull()); sql.finish();
    }
}
void testIntegerConversionDao() {
    auto db=QSqlDatabase::addDatabase("QSQLITE","integer-conversion"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE integer_values(narrow INTEGER,wide INTEGER)"));
    Host host(db);
    const auto result=linaccess::executeVbaProcedure(R"VBA(Sub SaveIntegers()
Dim definition As DAO.QueryDef
Dim narrow As Integer
Dim wide As Long
Dim copy As Variant
narrow=CInt(2.5)
wide=CLng(40000.5)
copy=narrow
Me!StoredTypes=TypeName(narrow)&":"&TypeName(wide)&":"&CStr(VarType(copy))
Set definition=CurrentDb.CreateQueryDef("", "PARAMETERS a SHORT,b LONG; INSERT INTO integer_values VALUES([a],[b])")
definition.Parameters("a").Value=narrow
definition.Parameters("b").Value=wide
definition.Execute
On Error Resume Next
narrow=CInt(32767.5)
Me!IntegerOverflow=CStr(Err.Number)&":"&CStr(narrow)&":"&TypeName(narrow)
End Sub
)VBA","SaveIntegers",host);
    assert(result.status==linaccess::VbaExecutionStatus::Completed);
    assert(host.fields["StoredTypes"]=="Integer:Long:2" && host.fields["IntegerOverflow"]=="6:2:Integer");
    assert(sql.exec("SELECT narrow,wide FROM integer_values") && sql.next());
    assert(sql.value(0).toInt()==2 && sql.value(1).toInt()==40000 && !sql.next()); sql.finish();
}
void testDeclaredParameterTypes() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","declared-parameter-types"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE parameter_values(value)")); assert(sql.exec("INSERT INTO parameter_values VALUES(7)"));
    auto database=linaccess::makeDaoDatabase(db);
    const auto definition=[&](const QString& spelling) {
        return database->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string(QString("PARAMETERS p %1; UPDATE parameter_values SET value=[p]").arg(spelling).toStdString())}).object;
    };
    for(const auto* alias:{"INTEGER","LONG","INT","INTEGER4"}) {
        auto query=definition(alias); auto parameter=query->member("Parameters",{VbaValue::string("p")}).object;
        const auto type=parameter->member("Type",{}); assert(type.number==4 && type.numericType==VbaValue::NumericType::Integer);
        for(const auto value:{-2147483648.0,2147483647.0}) {
            parameter->assign("Value",VbaValue::numeric(value)); query->member("Execute",{});
            assert(sql.exec("SELECT value FROM parameter_values") && sql.next() && sql.value(0).toDouble()==value); sql.finish();
        }
        parameter->assign("Value",VbaValue::numeric(2147483648.0));
        int rejected=0; try { query->member("Execute",{}); } catch(const linaccess::VbaRuntimeError& error) { rejected=error.number; } assert(rejected==13);
        assert(sql.exec("SELECT value FROM parameter_values") && sql.next() && sql.value(0).toLongLong()==2147483647); sql.finish();
    }
    for(const auto* alias:{"SHORT","SMALLINT","INTEGER2","BYTE","TINYINT","INTEGER1"}) {
        auto query=definition(alias); auto parameter=query->member("Parameters",{VbaValue::string("p")}).object;
        const bool byte=QString(alias)=="BYTE" || QString(alias)=="TINYINT" || QString(alias)=="INTEGER1";
        assert(parameter->member("Type",{}).number==(byte ? 2 : 3));
        parameter->assign("Value",VbaValue::numeric(byte ? 255 : 32767)); query->member("Execute",{});
        parameter->assign("Value",VbaValue::numeric(byte ? 256 : 32768));
        int rejected=0; try { query->member("Execute",{}); } catch(const linaccess::VbaRuntimeError& error) { rejected=error.number; } assert(rejected==13);
        assert(sql.exec("SELECT value FROM parameter_values") && sql.next() && sql.value(0).toInt()==(byte ? 255 : 32767)); sql.finish();
    }
    for(const auto* alias:{"SINGLE","REAL","FLOAT4","IEEESINGLE","DOUBLE","FLOAT","FLOAT8","IEEEDOUBLE","NUMBER"}) {
        auto query=definition(alias); auto parameter=query->member("Parameters",{VbaValue::string("p")}).object;
        const auto name=QString(alias); const bool single=name=="SINGLE" || name=="REAL" || name=="FLOAT4" || name=="IEEESINGLE";
        assert(parameter->member("Type",{}).number==(single ? 6 : 7));
        parameter->assign("Value",VbaValue::numeric(0.1)); query->member("Execute",{});
        assert(sql.exec("SELECT value FROM parameter_values") && sql.next() && sql.value(0).toDouble()==(single ? static_cast<double>(0.1f) : 0.1)); sql.finish();
        std::vector<double> invalid{std::numeric_limits<double>::infinity(),std::numeric_limits<double>::quiet_NaN()};
        if(single) invalid.push_back(3.5e38);
        for(const auto value:invalid) {
            parameter->assign("Value",VbaValue::numeric(value));
            int rejected=0; try { query->member("Execute",{}); } catch(const linaccess::VbaRuntimeError& error) { rejected=error.number; } assert(rejected==13);
            assert(sql.exec("SELECT value FROM parameter_values") && sql.next() && sql.value(0).toDouble()==(single ? static_cast<double>(0.1f) : 0.1)); sql.finish();
        }
        if(single) {
            parameter->assign("Value",VbaValue::numeric(1e-50));
            int rejected=0; try { query->member("Execute",{}); } catch(const linaccess::VbaRuntimeError& error) { rejected=error.number; } assert(rejected==3251);
        }
        parameter->assign("Value",VbaValue{}); query->member("Execute",{});
        assert(sql.exec("SELECT value FROM parameter_values") && sql.next() && sql.value(0).isNull()); sql.finish();
    }
    for(const auto& item:std::vector<std::pair<const char*,int>>{{"TEXT",10},{"VARCHAR",10},{"CHAR",10},{"DATE",8},{"DATETIME",8}}) {
        auto query=definition(item.first); auto parameter=query->member("Parameters",{VbaValue::string("p")}).object;
        assert(parameter->member("Type",{}).number==item.second);
        int rejected=0; try { parameter->member("Type",{VbaValue::numeric(1)}); } catch(const linaccess::VbaRuntimeError& error) { rejected=error.number; } assert(rejected==450);
    }
    auto unknown=definition("VALUE")->member("Parameters",{VbaValue::string("p")}).object;
    int unknownType=0; try { unknown->member("Type",{}); } catch(const linaccess::VbaRuntimeError& error) { unknownType=error.number; } assert(unknownType==3251);
    auto query=definition("LONG"); auto parameter=query->member("Parameters",{VbaValue::string("p")}).object;
    parameter->assign("Value",VbaValue::numeric(123)); const auto source=query->member("SQL",{}).asText();
    int rejected=0; try { parameter->assign("Type",VbaValue::numeric(3)); } catch(const linaccess::VbaRuntimeError& error) { rejected=error.number; }
    assert(rejected==383 && parameter->member("Type",{}).number==4 && parameter->member("Value",{}).number==123 && query->member("SQL",{}).asText()==source);
    query->assign("SQL",VbaValue::string("PARAMETERS p SHORT; UPDATE parameter_values SET value=[p]"));
    rejected=0; try { parameter->member("Type",{}); } catch(const linaccess::VbaRuntimeError& error) { rejected=error.number; } assert(rejected==3420);
    parameter=query->member("Parameters",{VbaValue::string("p")}).object; assert(parameter->member("Type",{}).number==3);
    query->member("Close",{}); rejected=0; try { parameter->member("Type",{}); } catch(const linaccess::VbaRuntimeError& error) { rejected=error.number; } assert(rejected==3420);
    Host host(db);
    const auto result=linaccess::executeVbaProcedure(R"VBA(Sub InspectType()
Dim definition As DAO.QueryDef
Set definition=CurrentDb.CreateQueryDef("", "PARAMETERS p INTEGER; UPDATE parameter_values SET value=[p]")
On Error Resume Next
definition.Parameters("p").Type=dbInteger
Me!TypeWriteError=CStr(Err.Number)
If definition.Parameters("p").Type=dbLong Then
definition.Parameters("p").Value=40000
definition.Execute
End If
Me!TypeStorage=TypeName(definition.Parameters("p").Type)
End Sub
)VBA","InspectType",host);
    assert(result.status==linaccess::VbaExecutionStatus::Completed && host.fields["TypeStorage"]=="Integer" && host.fields["TypeWriteError"]=="383");
    assert(sql.exec("SELECT value FROM parameter_values") && sql.next() && sql.value(0).toInt()==40000); sql.finish();
}
void testDdlRecordsAffected() {
    using namespace linaccess;
    auto db=QSqlDatabase::addDatabase("QSQLITE","dao-ddl-counts"); db.setDatabaseName(":memory:"); assert(db.open()); QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE seed(value INTEGER)")); assert(sql.exec("INSERT INTO seed VALUES(1),(2)")); sql.finish();
    const auto database=makeDaoDatabase(db);
    const auto execute=[&](const char* source) { database->member("Execute",{VbaValue::string(source)}); const auto count=database->member("RecordsAffected",{}); assert(count.asNumber()==0 && count.numericType==VbaValue::NumericType::Long); };
    execute("CREATE TABLE ddl_probe(value INTEGER)");
    const auto definition=database->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string("INSERT INTO ddl_probe SELECT value FROM seed")}).object;
    definition->member("Execute",{}); auto count=definition->member("RecordsAffected",{}); assert(count.asNumber()==2 && count.numericType==VbaValue::NumericType::Long);
    execute("ALTER TABLE ddl_probe ADD COLUMN label TEXT");
    definition->assign("SQL",VbaValue::string("UPDATE ddl_probe SET label='changed'")); definition->member("Execute",{}); count=definition->member("RecordsAffected",{}); assert(count.asNumber()==2 && count.numericType==VbaValue::NumericType::Long);
    const auto rows=database->member("OpenRecordset",{VbaValue::string("ddl_probe")}).object; count=rows->member("RecordCount",{}); assert(count.asNumber()==2 && count.numericType==VbaValue::NumericType::Long);
    count=rows->member("Fields",{}).object->member("Count",{}); assert(count.asNumber()==2 && count.numericType==VbaValue::NumericType::Integer); rows->member("Close",{});
    execute("DROP TABLE ddl_probe");
    database->member("Execute",{VbaValue::string("SELECT value INTO copied_rows FROM seed")}); count=database->member("RecordsAffected",{}); assert(count.asNumber()==2 && count.numericType==VbaValue::NumericType::Long); execute("DROP TABLE copied_rows");
}

void testTypedSqlTargets() {
    using namespace linaccess;
    auto db=QSqlDatabase::addDatabase("QSQLITE","typed-sql-targets"); db.setDatabaseName(":memory:"); assert(db.open()); QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE typed(id INTEGER PRIMARY KEY,b INTEGER,s INTEGER,l INTEGER,flag INTEGER,stamp TEXT,money REAL)"));
    assert(sql.exec("CREATE TABLE _la_table_schema(table_name TEXT PRIMARY KEY,definition TEXT)"));
    sql.prepare("INSERT INTO _la_table_schema VALUES('typed',?)"); sql.addBindValue(QStringLiteral(R"JSON([{"name":"b","source_type":"TINYINT"},{"name":"s","source_type":"SMALLINT"},{"name":"l","source_type":"INT"},{"name":"flag","source_type":"BOOLEAN"},{"name":"stamp","source_type":"DATETIME"},{"name":"money","source_type":"CURRENCY"}])JSON")); assert(sql.exec());
    assert(sql.exec("INSERT INTO typed VALUES(1,10,20,30,0,'2024-02-29 12:00:00',0.1),(2,11,21,31,-1,NULL,0.2)")); sql.finish();
    const auto database=makeDaoDatabase(db);
    const auto execute=[&](const char* source) { database->member("Execute",{VbaValue::string(source)}); };
    const auto row=[&](int id) { assert(sql.exec("SELECT b,s,l,flag,stamp,money FROM typed WHERE id="+QString::number(id)) && sql.next()); QVariantList values; for(int i=0;i<6;++i) values.append(sql.value(i)); sql.finish(); return values; };
    const auto fail=[&](const char* source,int wanted) { int code=0; try { execute(source); } catch(const VbaRuntimeError& error) { code=error.number; } assert(code==wanted); };
    execute("UPDATE typed SET b=2.5,s=-3.5,l=40000,flag=1,stamp='2024-03-01' WHERE id=1");
    auto values=row(1); assert(values[0].toDouble()==2 && values[1].toDouble()==-4 && values[2].toInt()==40000 && values[3].toInt()==-1 && values[4].toString()=="2024-03-01 00:00:00");
    const auto before=row(1),second=row(2);
    fail("UPDATE typed SET s=IIf(id=1,12,40000)",6); assert(row(1)==before && row(2)==second);
    fail("UPDATE typed SET b=-1 WHERE id=1",6); fail("UPDATE typed SET l=2147483648 WHERE id=1",6);
    fail("UPDATE typed SET stamp='invalid' WHERE id=1",13); fail("UPDATE typed SET stamp='0001-01-01' WHERE id=1",13); assert(row(1)==before);
    execute("UPDATE typed SET flag='false',b=NULL WHERE id=1"); assert(row(1)[0].isNull() && row(1)[3].toInt()==0);
    execute("INSERT INTO typed(id,b,s,l,flag,stamp,money) VALUES(3,2.5,-3.5,40000,1,'2024-03-01','1.23455')"); values=row(3);
    assert(values[0].toInt()==2 && values[1].toInt()==-4 && values[3].toInt()==-1 && values[4].toString()=="2024-03-01 00:00:00" && values[5].toDouble()==1.2346);
    fail("INSERT INTO typed(id,s) SELECT id+10,IIf(id=1,2.5,40000) FROM typed WHERE id IN(1,2)",6);
    assert(sql.exec("SELECT COUNT(*) FROM typed") && sql.next() && sql.value(0).toInt()==3); sql.finish();
    execute("INSERT INTO typed(id,s) SELECT id+10,2.5 FROM typed WHERE id IN(1,2)"); assert(row(11)[1].toInt()==2 && row(12)[1].toInt()==2);
    const auto parameterized=database->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string("PARAMETERS p Double; UPDATE typed SET s=p WHERE id=2")}).object;
    parameterized->member("Parameters",{VbaValue::string("p")}).object->assign("Value",VbaValue::numeric(2.5)); parameterized->member("Execute",{}); assert(row(2)[1].toInt()==2);
    assert(sql.exec("CREATE TABLE ordinary(value INTEGER)")); assert(sql.exec("INSERT INTO ordinary VALUES(1)")); execute("UPDATE ordinary SET value=0.5");
    assert(sql.exec("SELECT value FROM ordinary") && sql.next() && sql.value(0).toDouble()==0.5); sql.finish();
}

void testVariantArraysDao() {
    using namespace linaccess;
    auto db=QSqlDatabase::addDatabase("QSQLITE","variant-array-dao");
    db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE array_rows(id INTEGER PRIMARY KEY,amount INTEGER)"));
    Host host(db);
    const auto result=executeVbaProcedure(R"VBA(Option Base 1
Sub SaveArray()
Dim values As Variant, original As Variant
Dim i As Long
values=Array(12,24,36)
original=values
values(2)=42
For i=LBound(values) To UBound(values)
CurrentDb.Execute "INSERT INTO array_rows VALUES(" & CStr(i) & "," & CStr(values(i)) & ")"
Next
Me!Original=original(2)
End Sub
)VBA","SaveArray",host);
    assert(result.status==VbaExecutionStatus::Completed && host.fields["Original"]=="24");
    assert(sql.exec("SELECT COUNT(*),SUM(amount) FROM array_rows") && sql.next());
    assert(sql.value(0).toInt()==3 && sql.value(1).toInt()==90);
    sql.finish();
    assert(sql.exec("CREATE TABLE text_lists(value TEXT)"));
    const auto textLists=executeVbaProcedure(R"VBA(Function Parts(ParamArray values() As Variant) As String
Parts=Join(Delimiter:="",SourceArray:=values)
End Function
Function StoredLabel(ByVal text As String, Optional ByVal suffix As String="!") As String
StoredLabel=text & suffix
End Function
Sub SaveTextList()
Dim words As Variant, selected As Variant, item As Variant
Dim result As String
words=Split("alpha,beta,alps",",")
words(1)="alt"
selected=Filter(words,"al")
For Each item In selected
If Len(result)>0 Then
result=result & "|"
End If
If item Like "al*" Then
result=result & item
End If
Next
CurrentDb.Execute "INSERT INTO text_lists VALUES('" & StoredLabel(suffix:="!",text:=Parts(result)) & "')"
End Sub
)VBA","SaveTextList",host);
    assert(textLists.status==VbaExecutionStatus::Completed);
    assert(sql.exec("SELECT value FROM text_lists") && sql.next() && sql.value(0).toString()=="alpha|alt|alps!");
    sql.finish(); db.close();
}
void testOwnedQueryDefs() {
    auto db=QSqlDatabase::addDatabase("QSQLITE","owned-querydef-sql"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE records(value INTEGER)"));
    assert(sql.exec("INSERT INTO records VALUES(1),(2),(3)"));
    assert(sql.exec("CREATE TABLE _la_objects(kind TEXT,name TEXT,definition TEXT,PRIMARY KEY(kind,name))"));
    const auto save=[&](const QString& name,const QString& source) {
        sql.prepare("INSERT INTO _la_objects VALUES('query',?,?)"); sql.addBindValue(name);
        sql.addBindValue(QString::fromUtf8(QJsonDocument(QJsonObject{{"sql",source},{"archive","preserve"}}).toJson(QJsonDocument::Compact)));
        assert(sql.exec()); sql.finish();
    };
    save("Owned Base","SELECT value FROM records");
    save("Owned Caller","PARAMETERS p Long; SELECT value FROM [Owned Base] WHERE value>=p");
    Host host(db);
    const auto result=linaccess::executeVbaProcedure(R"VBA(Sub OwnedDefinitions()
Dim query As DAO.QueryDef, other As DAO.QueryDef, base As DAO.QueryDef, parameter As DAO.Parameter
Dim rows As DAO.Recordset, workspace As DAO.Workspace
Set query=CurrentDb.QueryDefs("owned caller")
If query.Type<>dbQSelect Or VarType(query.Type)<>vbInteger Then Err.Raise 5
If query.Name<>"Owned Caller" Then Err.Raise 5
Set parameter=query.Parameters("p")
parameter.Value=2
Set rows=query.OpenRecordset(dbOpenSnapshot)
If rows.RecordCount<>2 Then Err.Raise 5
rows.Close
query.SQL="PARAMETERS p Long; SELECT value*10 AS value FROM [Owned Base] WHERE value>=p"
On Error Resume Next
parameter.Value=1
Me!OwnedParameterError=CStr(Err.Number)
Err.Clear
On Error GoTo 0
Set other=CurrentDb.QueryDefs("Owned Caller")
If other.SQL<>query.SQL Then Err.Raise 5
other.Parameters("p").Value=3
Set rows=other.OpenRecordset(dbOpenSnapshot)
If rows!value<>30 Then Err.Raise 5
rows.Close
other.SQL="PARAMETERS q Long; SELECT value FROM [Owned Base] WHERE value>=q"
If query.SQL<>other.SQL Then Err.Raise 5
If query.Parameters.Count<>1 Then Err.Raise 5
Set base=CurrentDb.QueryDefs("Owned Base")
base.SQL="SELECT value FROM records WHERE value=3"
query.Parameters("q").Value=2
Set rows=query.OpenRecordset(dbOpenSnapshot)
If rows.RecordCount<>1 Or rows!value<>3 Then Err.Raise 5
rows.Close
Set workspace=DBEngine.Workspaces(0)
workspace.BeginTrans
query.SQL="SELECT 99 AS value"
workspace.Rollback
If InStr(query.SQL,"PARAMETERS q")<>1 Then Err.Raise 5
query.Close
On Error Resume Next
Me!OwnedClosedSQL=query.SQL
Me!OwnedClosedError=CStr(Err.Number)
End Sub
)VBA","OwnedDefinitions",host);
    if(result.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Synthetic owned QueryDef: "<<result.errorNumber<<" line "<<result.line<<" "<<result.explanation<<'\n';
    assert(result.status==linaccess::VbaExecutionStatus::Completed);
    assert(host.fields["OwnedParameterError"]=="3420" && host.fields["OwnedClosedError"]=="3420");
    const auto catalog=linaccess::executeVbaProcedure(R"VBA(Sub QueryCatalog()
Dim database As DAO.Database, definitions As DAO.QueryDefs, query As DAO.QueryDef, held As DAO.QueryDef
Dim temporary As DAO.QueryDef, rows As DAO.Recordset, workspace As DAO.Workspace
Dim item As Variant, count As Long
Set database=CurrentDb
Set definitions=database.QueryDefs
If Not TypeOf definitions Is DAO.QueryDefs Then Err.Raise 5
If definitions.Count<>2 Or VarType(definitions.Count)<>vbInteger Then Err.Raise 5
Set query=database.CreateQueryDef(Name:="Created From VBA",SQLText:="SELECT value FROM records")
If query.Name<>"Created From VBA" Then Err.Raise 5
If definitions.Count<>3 Then Err.Raise 5
Set held=definitions.Item("Created From VBA")
If Not (held Is definitions.Item("created from vba")) Then Err.Raise 5
Set rows=held.OpenRecordset(dbOpenSnapshot)
If rows.RecordCount<>3 Then Err.Raise 5
rows.Close
Set temporary=database.CreateQueryDef("","SELECT 1 AS value")
If temporary.Name<>"" Or definitions.Count<>3 Then Err.Raise 5
temporary.Close
For Each item In definitions
If Len(item.Name)=0 Then Err.Raise 5
count=count+1
Next
If count<>3 Then Err.Raise 5
definitions.Refresh
If definitions.Item(0).Name<>"Created From VBA" Then Err.Raise 5
On Error Resume Next
Set query=database.CreateQueryDef("created from vba","SELECT 2 AS value")
Me!CreatedDuplicateError=CStr(Err.Number)
Err.Clear
On Error GoTo 0
If definitions.Count<>3 Then Err.Raise 5
Set workspace=DBEngine.Workspaces(0)
workspace.BeginTrans
Set query=database.CreateQueryDef("Rolled Back Query","SELECT 99 AS value")
If definitions.Count<>4 Then Err.Raise 5
workspace.Rollback
If definitions.Count<>3 Then Err.Raise 5
On Error Resume Next
Me!RolledQueryName=query.Name
Me!RolledQueryError=CStr(Err.Number)
Err.Clear
On Error GoTo 0
Set query=database.CreateQueryDef("Empty Created")
If query.SQL<>"" Then Err.Raise 5
query.SQL="SELECT 42 AS value"
Set rows=query.OpenRecordset(dbOpenSnapshot)
If rows!value<>42 Then Err.Raise 5
rows.Close
definitions.Delete "Empty Created"
Set temporary=database.CreateQueryDef("Empty Created","SELECT 7 AS value")
On Error Resume Next
Me!DeletedOldQuery=query.Name
Me!DeletedOldError=CStr(Err.Number)
Err.Clear
On Error GoTo 0
definitions.Delete "Empty Created"
If definitions.Count<>3 Then Err.Raise 5
workspace.BeginTrans
definitions.Delete "Created From VBA"
If definitions.Count<>2 Then Err.Raise 5
Set rows=definitions.Item("Owned Base").OpenRecordset(dbOpenSnapshot)
rows.Close
workspace.Rollback
If definitions.Count<>3 Then Err.Raise 5
On Error Resume Next
Me!DeletedHeldName=held.Name
Me!DeletedHeldError=CStr(Err.Number)
Err.Clear
On Error GoTo 0
Set held=definitions.Item("Created From VBA")
If held.Name<>"Created From VBA" Then Err.Raise 5
held.SQL="SELECT value+10 AS value FROM records ORDER BY value"
Set temporary=database.CreateQueryDef("","SELECT value FROM [Created From VBA] ORDER BY value")
Set rows=temporary.OpenRecordset(dbOpenSnapshot)
If rows!value<>11 Then Err.Raise 5
rows.Close
End Sub
)VBA","QueryCatalog",host);
    if(catalog.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Synthetic QueryDefs catalog: "<<catalog.errorNumber<<" line "<<catalog.line<<" "<<catalog.explanation<<'\n';
    assert(catalog.status==linaccess::VbaExecutionStatus::Completed);
    assert(host.fields["CreatedDuplicateError"]=="3012" && host.fields["RolledQueryError"]=="3420");
    assert(host.fields["DeletedOldError"]=="3420" && host.fields["DeletedHeldError"]=="3420");
    assert(sql.exec("SELECT COUNT(*) FROM _la_objects WHERE kind='query'") && sql.next() && sql.value(0).toInt()==3); sql.finish();

    assert(sql.exec("SELECT definition FROM _la_objects WHERE name='Owned Caller'") && sql.next());
    const auto definition=QJsonDocument::fromJson(sql.value(0).toByteArray()).object(); sql.finish();
    assert(definition.value("archive").toString()=="preserve" && definition.value("sql").toString().startsWith("PARAMETERS q"));
    assert(sql.exec("SELECT COUNT(*) FROM records") && sql.next() && sql.value(0).toInt()==3); sql.finish();
}

void testAccessEvalNumericOracle() {
    using namespace linaccess;QFile file(QFileInfo(QString::fromUtf8(__FILE__)).absoluteDir().filePath("access2003_reference/expression-observations.json"));assert(file.open(QIODevice::ReadOnly));const auto root=QJsonDocument::fromJson(file.readAll()).object();assert(root.value("schema_version").toInt()==1 && root.value("provenance").toObject().value("access_version").toString()=="11.0");
    auto db=QSqlDatabase::addDatabase("QSQLITE","access-eval-numeric-oracle");db.setDatabaseName(":memory:");assert(db.open());Host host(db);int checked=0;
    for(const auto& entry:root.value("observations").toArray()) {
        const auto row=entry.toObject();if(row.value("verification_group").toString()!="numeric") continue;const auto expression=row.value("expression").toString();VbaValue actual;
        try {actual=host.eval(expression.toStdString());}catch(const VbaRuntimeError& failure) {std::cerr<<"Access Eval numeric oracle "<<row.value("case").toString().toStdString()<<": error="<<failure.number<<"\n";assert(false);}
        const auto expectedType=row.value("vartype").toInt();const auto actualType=evaluateVbaExpression("VarType(value)",host,{{"value",actual}}).asNumber();if(actualType!=expectedType) std::cerr<<"Access Eval oracle "<<row.value("case").toString().toStdString()<<": type "<<actualType<<" expected "<<expectedType<<"\n";assert(actualType==expectedType);
        const auto expected=row.value("result").toObject();
        if(expectedType==0 || expectedType==1) {assert(actual.kind==(expectedType==0 ? VbaValue::Kind::Empty : VbaValue::Kind::Null));}
        else if(expectedType==6 || expectedType==14) {assert(actual.asText()==expected.value("invariant").toString().toStdString());}
        else if(expectedType==8) {assert(actual.asText()==expected.value("value").toString().toStdString());}
        else if(expectedType==5) {const auto bytes=QByteArray::fromHex(expected.value("bits").toString().toLatin1());assert(bytes.size()==8);std::uint64_t bits=0;for(int i=0;i<8;++i) bits|=static_cast<std::uint64_t>(static_cast<unsigned char>(bytes[i]))<<(8*i);const auto value=std::bit_cast<double>(bits);assert(std::abs(actual.asNumber()-value)<=2e-12*std::max(1.0,std::abs(value)));}
        else {assert(actual.asNumber()==expected.value("value").toDouble());}
        ++checked;
    }
    assert(checked==56);
    int checkedErrors=0;for(const auto& entry:root.value("errors").toArray()) {const auto row=entry.toObject();int error=0;try {(void)host.eval(row.value("expression").toString().toStdString());}catch(const VbaRuntimeError& failure) {error=failure.number;}if(error!=row.value("error_number").toInt()) std::cerr<<"Access Eval error oracle "<<row.value("case").toString().toStdString()<<": "<<error<<" expected "<<row.value("error_number").toInt()<<"\n";assert(error==row.value("error_number").toInt());++checkedErrors;}assert(checkedErrors==8);
}

void testOwnedQueryRename() {
    using namespace linaccess;auto db=QSqlDatabase::addDatabase("QSQLITE","owned-query-rename");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);assert(sql.exec("CREATE TABLE values_table(value INTEGER)"));assert(sql.exec("INSERT INTO values_table VALUES(7),(9)"));
    auto database=makeDaoDatabase(db);auto query=database->member("CreateQueryDef",{VbaValue::string("Original Name"),VbaValue::string("PARAMETERS p LONG; SELECT value FROM values_table WHERE value>=p")}).object;
    assert(sql.exec("SELECT definition FROM _la_objects WHERE kind='query' AND name='Original Name'") && sql.next());const auto originalDefinition=sql.value(0).toByteArray();sql.finish();
    auto catalog=database->member("QueryDefs",{}).object;auto other=catalog->member("Item",{VbaValue::string("Original Name")}).object;auto parameter=query->member("Parameters",{VbaValue::string("p")}).object;parameter->assign("Value",VbaValue::longInteger(8));
    const auto error=[](const std::function<void()>& operation,int expected) {int number=0;try {operation();}catch(const VbaRuntimeError& failure) {number=failure.number;}assert(number==expected);};
    error([&] {query->assign("Name",VbaValue{});},94);error([&] {query->assign("Name",VbaValue::string(" "));},5);error([&] {query->assign("Name",VbaValue::string("values_table"));},3012);
    auto engine=makeDaoEngine(db);engine->member("BeginTrans",{});error([&] {query->assign("Name",VbaValue::string("Blocked"));},3251);engine->member("Rollback",{});assert(query->member("Name",{}).asText()=="Original Name");
    assert(sql.exec("PRAGMA query_only=ON"));error([&] {query->assign("Name",VbaValue::string("Blocked"));},3027);assert(sql.exec("PRAGMA query_only=OFF"));sql.finish();
    assert(sql.exec("CREATE TABLE _la_datasheet_page_setup(kind TEXT NOT NULL,name TEXT NOT NULL,layout TEXT NOT NULL,PRIMARY KEY(kind,name))"));assert(sql.exec("INSERT INTO _la_datasheet_page_setup VALUES('query','original name','unchanged page bytes'),('query','page conflict','orphan')"));sql.finish();error([&] {query->assign("Name",VbaValue::string("Page Conflict"));},3251);assert(query->member("Name",{}).asText()=="Original Name" && other->member("Name",{}).asText()=="Original Name");assert(sql.exec("SELECT COUNT(*) FROM _la_objects WHERE name='Original Name'") && sql.next() && sql.value(0).toInt()==1);sql.finish();
    query->assign("Name",VbaValue::string("Renamed Ä"));assert(query->member("Name",{}).asText()=="Renamed Ä");assert(sql.exec("SELECT definition FROM _la_objects WHERE kind='query' AND name='Renamed Ä'") && sql.next() && sql.value(0).toByteArray()==originalDefinition);sql.finish();error([&] {(void)other->member("Name",{});},3420);error([&] {(void)catalog->member("Item",{VbaValue::string("Original Name")});},3265);
    assert(parameter->member("Value",{}).asNumber()==8);auto rows=query->member("OpenRecordset",{}).object;assert(rows->member("RecordCount",{}).asNumber()==1 && rows->item("value").asNumber()==9);rows->member("Close",{});
    auto current=catalog->member("Item",{VbaValue::string("Renamed Ä")}).object;current->assign("Name",VbaValue::string("RENAMED Ä"));assert(catalog->member("Item",{VbaValue::string("RENAMED Ä")}).object==current);error([&] {(void)query->member("Name",{});},3420);
    auto restored=database->member("CreateQueryDef",{VbaValue::string("Original Name"),VbaValue::string("SELECT 17 AS value")}).object;assert(catalog->member("Item",{VbaValue::string("Original Name")}).object->member("Name",{}).asText()=="Original Name");
    current->assign("SQL",VbaValue::string("SELECT 9 AS value"));auto all=database->member("OpenRecordset",{VbaValue::string("SELECT value FROM [RENAMED Ä]")}).object;assert(all->item("value").asNumber()==9);all->member("Close",{});
    current->assign("Name",VbaValue::string("Final Name"));assert(sql.exec("SELECT layout FROM _la_datasheet_page_setup WHERE kind='query' AND name='final name'") && sql.next() && sql.value(0).toString()=="unchanged page bytes");sql.finish();assert(!db.tables(QSql::Views).contains("RENAMED Ä",Qt::CaseInsensitive));assert(db.tables(QSql::Views).contains("Final Name",Qt::CaseInsensitive));
    assert(sql.exec("SELECT definition FROM _la_objects WHERE kind='query' AND name='Final Name'") && sql.next());assert(QJsonDocument::fromJson(sql.value(0).toByteArray()).object().value("sql").toString()=="SELECT 9 AS value");sql.finish();assert(sql.exec("SELECT COUNT(*) FROM values_table") && sql.next() && sql.value(0).toInt()==2);sql.finish();
    assert(sql.exec("ALTER TABLE _la_datasheet_page_setup ADD COLUMN unexpected TEXT"));error([&] {catalog->member("Delete",{VbaValue::string("Final Name")});},3251);assert(current->member("Name",{}).asText()=="Final Name");assert(sql.exec("ALTER TABLE _la_datasheet_page_setup DROP COLUMN unexpected"));sql.finish();
    catalog->member("Delete",{VbaValue::string("Final Name")});error([&] {(void)current->member("Name",{});},3420);assert(!db.tables(QSql::Views).contains("Final Name",Qt::CaseInsensitive));assert(sql.exec("SELECT COUNT(*) FROM _la_datasheet_page_setup WHERE kind='query' AND name='final name'") && sql.next() && sql.value(0).toInt()==0);sql.finish();
    auto replacement=database->member("CreateQueryDef",{VbaValue::string("Final Name"),VbaValue::string("SELECT 33 AS value")}).object;auto fresh=replacement->member("OpenRecordset",{}).object;assert(fresh->item("value").asNumber()==33);fresh->member("Close",{});
    assert(sql.exec("SELECT COUNT(*) FROM values_table") && sql.next() && sql.value(0).toInt()==2);sql.finish();
    assert(sql.exec("INSERT INTO _la_datasheet_page_setup VALUES('table','values_table','old table page')"));sql.finish();database->member("TableDefs",{}).object->member("Delete",{VbaValue::string("values_table")});assert(!db.tables().contains("values_table"));assert(sql.exec("SELECT COUNT(*) FROM _la_datasheet_page_setup WHERE kind='table' AND name='values_table'") && sql.next() && sql.value(0).toInt()==0);sql.finish();
}

void testDaoGetRows() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","dao-getrows"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE batch_records(id INTEGER PRIMARY KEY,label TEXT,price CURRENCY_REAL,stamp DATETIME,enabled BOOLEAN,raw BLOB)"));
    assert(sql.exec("INSERT INTO batch_records VALUES(1,'first',0.1001,'2024-02-29 12:34:56',-1,X'00FF42'),(2,NULL,-0.0001,NULL,0,X''),(3,'last',NULL,NULL,NULL,NULL)"));
    Host host(db);
    const auto result=linaccess::executeVbaProcedure(R"VBA(Option Base 1
Sub ReadBatches()
Dim rows As DAO.Recordset
Dim batch As Variant
Set rows=CurrentDb.OpenRecordset("batch_records",4)
batch=rows.GetRows(NumRows:=2)
Me!BatchBounds=CStr(LBound(batch,1)) & ":" & CStr(UBound(batch,1)) & ":" & CStr(LBound(batch,2)) & ":" & CStr(UBound(batch,2))
Me!BatchTypes=CStr(VarType(batch(2,0))) & ":" & CStr(VarType(batch(3,0))) & ":" & CStr(VarType(batch(4,0))) & ":" & CStr(VarType(batch(5,0)))
Me!BatchMoney=CStr(batch(2,0)) & ":" & CStr(batch(2,1))
Me!BatchNull=CStr(IsNull(batch(1,1)))
Me!BatchBinary=CStr(batch(5,0)(0)) & ":" & CStr(batch(5,0)(1)) & ":" & CStr(batch(5,0)(2))
Me!BatchEmptyBinary=CStr(UBound(batch(5,1)))
Me!BatchPosition=CStr(rows.AbsolutePosition) & ":" & CStr(rows.EOF)
batch(0,0)=77
batch=rows.GetRows(99)
Me!BatchTail=CStr(UBound(batch,2)) & ":" & CStr(batch(0,0)) & ":" & CStr(rows.EOF)
On Error Resume Next
batch=rows.GetRows(1)
Me!BatchEofError=CStr(Err.Number)
Err.Clear
rows.MoveFirst
batch=rows.GetRows(0)
Me!BatchZeroError=CStr(Err.Number)
Err.Clear
batch=rows.GetRows(Null)
Me!BatchNullError=CStr(Err.Number)
Err.Clear
batch=rows.GetRows()
Me!BatchOmittedError=CStr(Err.Number)
Me!BatchErrorPosition=CStr(rows.AbsolutePosition)
rows.Close
Err.Clear
batch=rows.GetRows(1)
Me!BatchClosedError=CStr(Err.Number)
End Sub
)VBA","ReadBatches",host);
    if(result.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Synthetic GetRows runtime error="<<result.errorNumber<<" line="<<result.line<<"\n";
    assert(result.status==linaccess::VbaExecutionStatus::Completed);
    assert(host.fields["BatchBounds"]=="0:5:0:1" && host.fields["BatchTypes"]=="6:7:11:8209");
    assert(host.fields["BatchMoney"]=="0.1001:-0.0001" && host.fields["BatchNull"]=="True");
    assert(host.fields["BatchBinary"]=="0:255:66" && host.fields["BatchEmptyBinary"]=="-1");
    assert(host.fields["BatchPosition"]=="2:False" && host.fields["BatchTail"]=="0:3:True");
    assert(host.fields["BatchEofError"]=="3021" && host.fields["BatchZeroError"]=="3251" && host.fields["BatchNullError"]=="94" && host.fields["BatchOmittedError"]=="3251" && host.fields["BatchClosedError"]=="3420" && host.fields["BatchErrorPosition"]=="0");
    assert(sql.exec("SELECT count(*),sum(id) FROM batch_records") && sql.next() && sql.value(0).toInt()==3 && sql.value(1).toInt()==6); sql.finish();
    const auto references=linaccess::executeVbaProcedure(R"VBA(Sub ObjectArrays()
Dim refs As Variant
Dim saved As Variant
Dim database As DAO.Database
Dim rows As DAO.Recordset
ReDim refs(0 To 1,0 To 1)
Set refs(0,0)=CurrentDb
Set refs(1,0)=Nothing
saved=refs
Set refs(0,0)=Nothing
Set database=saved(0,0)
database.Execute "CREATE TABLE object_array_values(value INTEGER)"
database.Execute "INSERT INTO object_array_values VALUES(7)"
Set refs(0,1)=database
Me!ObjectArrayIdentity=CStr(refs(0,1) Is saved(0,0))
Set rows=database.OpenRecordset("object_array_values",4)
Set refs(1,1)=rows
refs(1,1).MoveNext
Me!ObjectArrayEof=CStr(refs(1,1).EOF)
refs(1,1).Close
On Error Resume Next
Set refs(0,1)=42
Me!ObjectArrayScalarError=CStr(Err.Number)
Err.Clear
Set refs(2,0)=database
Me!ObjectArrayIndexError=CStr(Err.Number)
On Error GoTo 0
Me!ObjectArrayNothing=CStr(refs(0,0) Is Nothing) & ":" & CStr(refs(1,0) Is Nothing)
Erase refs
Me!ObjectArrayErased=TypeName(refs)
ReDim refs(0 To 1) As Object
Me!ObjectTyped=TypeName(refs) & ":" & CStr(VarType(refs)) & ":" & CStr(refs(0) Is Nothing)
Set refs(0)=database
ReDim Preserve refs(0 To 2) As Object
Me!ObjectTypedPreserved=CStr(refs(0) Is saved(0,0)) & ":" & CStr(refs(2) Is Nothing)
On Error Resume Next
refs(0)=7
Me!ObjectTypedValueError=CStr(Err.Number)
On Error GoTo 0
Erase refs
Set database=saved(0,0)
database.Execute "UPDATE object_array_values SET value=9"
End Sub
)VBA","ObjectArrays",host);
    if(references.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Synthetic object array error="<<references.errorNumber<<" line="<<references.line<<"\n";
    assert(references.status==linaccess::VbaExecutionStatus::Completed);
    assert(host.fields["ObjectArrayIdentity"]=="True" && host.fields["ObjectArrayEof"]=="True" && host.fields["ObjectArrayNothing"]=="True:True");
    assert(host.fields["ObjectArrayScalarError"]=="13" && host.fields["ObjectArrayIndexError"]=="9" && host.fields["ObjectArrayErased"]=="Empty");
    assert(host.fields["ObjectTyped"]=="Object():8201:True" && host.fields["ObjectTypedPreserved"]=="True:True" && host.fields["ObjectTypedValueError"]=="3251");
    assert(sql.exec("SELECT value FROM object_array_values") && sql.next() && sql.value(0).toInt()==9); sql.finish();
    const auto binaryFields=linaccess::executeVbaProcedure(R"VBA(Sub BinaryFields()
Dim rows As DAO.Recordset, held As DAO.Field, payload() As Byte, emptyBytes() As Byte
Dim prefix() As Byte, middle() As Byte, suffix() As Byte
Set rows=CurrentDb.OpenRecordset("batch_records")
Set held=rows.Fields("raw")
If held.Type<>11 Then Err.Raise 5
If held.FieldSize<>3 Or VarType(held.FieldSize)<>3 Then Err.Raise 5
payload=held.GetChunk(Bytes:=1,Offset:=1)
If VarType(payload)<>8209 Or UBound(payload)<>0 Or payload(0)<>255 Then Err.Raise 5
payload(0)=8
payload=held.GetChunk(0,99)
If UBound(payload)<>2 Or payload(0)<>0 Or payload(1)<>255 Or payload(2)<>66 Then Err.Raise 5
payload=held.GetChunk(2,99)
If UBound(payload)<>0 Or payload(0)<>66 Then Err.Raise 5
On Error Resume Next
held.FieldSize=1
Me!BinarySizeWriteError=CStr(Err.Number)
Err.Clear
payload=held.GetChunk(3,1)
Me!BinaryChunkBoundary=CStr(Err.Number)
Err.Clear
payload=held.GetChunk(Null,1)
Me!BinaryChunkNullOffset=CStr(Err.Number)
Err.Clear
payload=held.GetChunk(0)
Me!BinaryChunkArity=CStr(Err.Number)
On Error GoTo 0
payload=held.Value
If VarType(payload)<>8209 Or UBound(payload)<>2 Then Err.Raise 5
If payload(0)<>0 Or payload(1)<>255 Or payload(2)<>66 Then Err.Raise 5
prefix=held.GetChunk(0,1)
middle=held.GetChunk(1,1)
suffix=held.GetChunk(2,1)
middle(0)=7
rows.Edit
held.AppendChunk Val:=prefix
payload=held.Value
If UBound(payload)<>0 Or payload(0)<>0 Then Err.Raise 5
held.AppendChunk middle
held.AppendChunk suffix
payload=held.Value
If UBound(payload)<>2 Or payload(1)<>7 Then Err.Raise 5
On Error Resume Next
held.Value=payload
Me!BinaryMixedChunkError=CStr(Err.Number)
On Error GoTo 0
rows.Update
rows.Edit
held.Value=payload
On Error Resume Next
held.AppendChunk prefix
Me!BinaryMixedValueError=CStr(Err.Number)
On Error GoTo 0
rows.Update
rows.Edit
held.AppendChunk suffix
payload=held.Value
If UBound(payload)<>0 Or payload(0)<>66 Then Err.Raise 5
rows.CancelUpdate
payload=held.Value
If UBound(payload)<>2 Or payload(1)<>7 Then Err.Raise 5
rows.MoveNext
emptyBytes=held.Value
If held.FieldSize<>0 Then Err.Raise 5
If UBound(emptyBytes)<>-1 Then Err.Raise 5
rows.Edit
held.AppendChunk emptyBytes
rows.Update
rows.MoveNext
If Not IsNull(held.Value) Then Err.Raise 5
rows.Edit
held.AppendChunk payload
rows.CancelUpdate
If Not IsNull(held.Value) Then Err.Raise 5
rows.AddNew
held.AppendChunk payload
If UBound(held.Value)<>2 Then Err.Raise 5
rows.CancelUpdate
If rows.RecordCount<>3 Then Err.Raise 5
rows.MoveLast
rows.MoveNext
If Not rows.EOF Or held.Name<>"raw" Then Err.Raise 5
rows.Close
On Error Resume Next
payload=held.Value
Me!BinaryClosedError=CStr(Err.Number)
Err.Clear
payload=held.GetChunk(0,1)
Me!BinaryChunkClosedError=CStr(Err.Number)
Err.Clear
Me!ClosedBinaryFieldName=held.Name
Me!BinaryNameClosedError=CStr(Err.Number)
End Sub
)VBA","BinaryFields",host);
    if(binaryFields.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Synthetic binary fields: "<<binaryFields.errorNumber<<" line "<<binaryFields.line<<" "<<binaryFields.explanation<<'\n';
    assert(binaryFields.status==linaccess::VbaExecutionStatus::Completed && host.fields["BinaryClosedError"]=="3420" && host.fields["BinarySizeWriteError"]=="383" && host.fields["BinaryChunkClosedError"]=="3420" && host.fields["BinaryNameClosedError"]=="3420");
    assert(host.fields["BinaryChunkBoundary"]=="3251" && host.fields["BinaryChunkNullOffset"]=="94" && host.fields["BinaryChunkArity"]=="450" && host.fields["BinaryMixedChunkError"]=="3251" && host.fields["BinaryMixedValueError"]=="3251");
    assert(sql.exec("SELECT hex(raw),typeof(raw),length(raw) FROM batch_records ORDER BY id") && sql.next());
    assert(sql.value(0).toString()=="000742" && sql.value(1).toString()=="blob" && sql.value(2).toInt()==3);
    assert(sql.next() && sql.value(1).toString()=="blob" && sql.value(2).toInt()==0);
    assert(sql.next() && sql.value(1).toString()=="null"); sql.finish();
    assert(sql.exec("CREATE TABLE batch_large(value INTEGER)")); assert(sql.exec("INSERT INTO batch_large VALUES(1),(9007199254740993)"));
    auto rows=linaccess::makeDaoDatabase(db)->member("OpenRecordset",{VbaValue::string("batch_large"),VbaValue::numeric(4)}).object;
    int error=0; try { rows->member("GetRows",{VbaValue::numeric(2)}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
    assert(error==6 && rows->member("AbsolutePosition",{}).number==0 && rows->item("value").number==1); rows->member("Close",{});
    db.close();
}
void testSqlVbaFunctions() {
    auto db=QSqlDatabase::addDatabase("QSQLITE","native-vba-sql-functions"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE inputs(value INTEGER4)")); assert(sql.exec("INSERT INTO inputs VALUES(1),(2),(3)"));
    QByteArray module=R"VBA(Public calls As Long
Public held As Object
Public Function NativeScale(ByVal value As Long) As Long
calls=calls+1
NativeScale=SQLDoubleHelper(value)
End Function
Private Function SQLDoubleHelper(ByVal value As Long) As Long
SQLDoubleHelper=value*2
End Function
Public Function NativeCalls() As Long
NativeCalls=calls
End Function
Public Function NativeSequence() As Long
Static counter As Long
counter=counter+1
NativeSequence=counter
End Function
Public Function NativeMaybe(ByVal value As Variant) As Variant
If IsNull(value) Then
NativeMaybe=Null
Else
NativeMaybe=CStr(value)
End If
End Function
Public Function NativeFlag(ByVal value As Long) As Boolean
NativeFlag=value>1
End Function
Public Function NativeNeedsUI() As String
MsgBox "must not display"
NativeNeedsUI="unreachable"
End Function
Public Function NativeMoney() As Currency
calls=calls+1
NativeMoney=CCur("0.1")
End Function
Public Function NativeDate() As Date
NativeDate=DateSerial(2024,2,29)+TimeSerial(12,34,56)
End Function
Public Function NativeVariantMoney() As Variant
NativeVariantMoney=CCur(1)
End Function
Public Function Abs() As Long
Abs=999
End Function
Public Function FV(ByVal rate As Double, ByVal periods As Double, ByVal payment As Double) As Currency
FV=CCur(999)
End Function
)VBA";
    const auto store=[&](const QString& name,const QByteArray& source) {
        sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('vba',?,?)"); sql.addBindValue(name); sql.addBindValue(source);
        assert(sql.exec()); sql.finish();
    };
    store("NativeMath",module);
    store("NativeOther",R"VBA(Public Function NativeNested(ByVal value As Long) As Long
NativeNested=NativeScale(value)+1
End Function
)VBA");
    store("Form_PrivateContext","VERSION 1.0 CLASS\nPublic Function ClassOnly() As Long\nClassOnly=42\nEnd Function\n");
    store("ConditionalMath",R"VBA(#If Not VBA7 Then
Public Function ConditionalScale(ByVal value As Long) As Long
ConditionalScale=value+5
End Function
#Else
Public Function ConditionalScale(ByVal value As Long) As Currency
Err.Raise 7091
End Function
#End If
#If Win32 Then
Public Function WindowsOnlySQL() As Long
Err.Raise 7092
End Function
#End If
)VBA");
    linaccess::installJetSqlFunctions(db);
    const auto conditionalTypes=linaccess::nativeVbaSqlFunctionResultTypes(db);
    assert(conditionalTypes.value("conditionalscale")=="long" && !conditionalTypes.contains("windowsonlysql"));
    const auto conditionalResult=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT ConditionalScale(2) AS result"),{});
    assert(conditionalResult.rows.size()==1 && conditionalResult.rows[0][0].toInt()==7 && conditionalResult.columnTypes.value(0)=="INTEGER4");
    assert(!linaccess::nativeVbaSqlFunctionResultTypes(db).contains("fv"));
    const auto builtinFuture=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT FV(0,2,-10) AS future"),{});
    assert(builtinFuture.rows.size()==1 && builtinFuture.rows[0][0].toDouble()==20 && builtinFuture.columnTypes.value(0)!="CURRENCY_REAL");
    QSqlQuery rows(db); assert(rows.exec("SELECT NativeScale(value) FROM inputs ORDER BY value"));
    for(int expected:{2,4,6}) assert(rows.next() && rows.value(0).toInt()==expected); assert(!rows.next()); rows.finish();
    assert(rows.exec("SELECT NativeCalls(),NativeNested(3),NativeFlag(2),abs(-2)")); assert(rows.next());
    assert(rows.value(0).toInt()==3 && rows.value(1).toInt()==7 && rows.value(2).toInt()==-1 && rows.value(3).toInt()==2); rows.finish();
    assert(rows.exec("SELECT NativeSequence() FROM inputs ORDER BY value"));
    for(int expected:{1,2,3}) assert(rows.next() && rows.value(0).toInt()==expected); assert(!rows.next()); rows.finish();
    assert(rows.exec("SELECT NativeSequence()") && rows.next() && rows.value(0).toInt()==4); rows.finish();
    assert(sql.exec("UPDATE inputs SET value=NativeScale(value)")); sql.finish();
    assert(rows.exec("SELECT value FROM inputs ORDER BY value"));
    for(int expected:{2,4,6}) assert(rows.next() && rows.value(0).toInt()==expected); assert(!rows.next()); rows.finish();
    assert(rows.exec("SELECT NativeMaybe(NULL),NativeMaybe('A'||char(0)||'B')") && rows.next());
    assert(rows.value(0).isNull() && rows.value(1).toString().toUtf8()==QByteArray("A\0B",3)); rows.finish();
    assert(!rows.exec("SELECT SQLDoubleHelper(2)")); rows.finish(); assert(!rows.exec("SELECT ClassOnly()")); rows.finish();
    assert(!rows.exec("SELECT NativeNeedsUI()"));
    assert(linaccess::jetSqlFunctionErrorNumber(rows.lastError().text().toCaseFolded().toStdString())==3251); rows.finish();
    assert(!rows.exec("SELECT NativeVariantMoney()"));
    assert(linaccess::jetSqlFunctionErrorNumber(rows.lastError().text().toCaseFolded().toStdString())==3251); rows.finish();
    const auto typedSource=QStringLiteral("SELECT NativeMoney() AS amount,NativeDate() AS stamp,NativeScale(2) AS integer_value,NativeFlag(2) AS enabled");
    const auto metadataFrame=linaccess::standardVbaModuleState(db,module.toStdString());
    const auto beforeMetadata=metadataFrame->values.at("calls").asNumber();
    QVector<QString> columnTypes;
    const auto columns=linaccess::accessQueryColumns(db,linaccess::compileAccessQuery(typedSource),&columnTypes);
    assert(columns.size()==4 && columnTypes.value(0)=="CURRENCY_REAL" && columnTypes.value(1)=="DATETIME");
    const auto unnamed=linaccess::accessQueryColumns(db,linaccess::compileAccessQuery("SELECT NativeMoney()"));
    assert(unnamed.size()==1 && unnamed.front().contains("NativeMoney",Qt::CaseInsensitive) && !unnamed.front().contains("linaccess_"));
    assert(metadataFrame->values.at("calls").asNumber()==beforeMetadata);
    const auto database=linaccess::makeDaoDatabase(db);
    {
        using linaccess::VbaValue;
        const auto definition=database->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string(typedSource.toStdString())}).object;
        const auto fields=definition->member("Fields",{}).object;
        assert(fields->member("Count",{}).asNumber()==4 && fields->isNativeType("dao.fields"));
        const auto amountField=fields->member("Item",{VbaValue::string("amount")}).object;
        assert(amountField->member("Name",{}).asText()=="amount" && amountField->member("Type",{}).asNumber()==5 && amountField->member("Size",{}).asNumber()==8);
        assert(fields->member("Item",{VbaValue::integer(1)}).object->member("Type",{}).asNumber()==8);
        assert(fields->member("Item",{VbaValue::string("integer_value")}).object->member("Type",{}).asNumber()==4);
        const auto flagField=fields->member("Item",{VbaValue::string("enabled")}).object;
        assert(flagField->member("Type",{}).asNumber()==1 && flagField->member("Size",{}).asNumber()==1);
        const auto enumerated=fields->enumerate(); assert(enumerated.size()==4);
        assert(enumerated.back().object->member("Name",{}).asText()=="enabled" && enumerated.back().object->member("OrdinalPosition",{}).asNumber()==3);
        assert(definition->member("Fields",{}).object==fields && fields->member("Item",{VbaValue::integer(0)}).object==amountField);
        int failure=0; try { (void)amountField->member("DefaultValue",{}); } catch(const linaccess::VbaRuntimeError& error) { failure=error.number; } assert(failure==3251);
        assert(metadataFrame->values.at("calls").asNumber()==beforeMetadata);
        definition->assign("SQL",VbaValue::string("SELECT NativeDate() AS changed"));
        failure=0; try { (void)fields->member("Count",{}); } catch(const linaccess::VbaRuntimeError& error) { failure=error.number; } assert(failure==3420);
        failure=0; try { (void)amountField->member("Name",{}); } catch(const linaccess::VbaRuntimeError& error) { failure=error.number; } assert(failure==3420);
        const auto current=definition->member("Fields",{}).object; assert(current->member("Count",{}).asNumber()==1);
        assert(sql.exec("CREATE TABLE query_fields_schema_probe(id INTEGER)")); sql.finish();
        failure=0; try { (void)current->member("Count",{}); } catch(const linaccess::VbaRuntimeError& error) { failure=error.number; } assert(failure==3420);
        const auto refreshed=definition->member("Fields",{}).object; assert(refreshed->member("Count",{}).asNumber()==1);
        const auto workspace=linaccess::makeDaoEngine(db)->defaultMember({VbaValue::integer(0)},true).object;
        workspace->member("BeginTrans",{});
        workspace->member("Rollback",{});
        failure=0; try { (void)refreshed->member("Count",{}); } catch(const linaccess::VbaRuntimeError& error) { failure=error.number; } assert(failure==3420);
        const auto afterRollback=definition->member("Fields",{}).object; assert(afterRollback->member("Count",{}).asNumber()==1);
        definition->member("Close",{});
        failure=0; try { (void)afterRollback->member("Count",{}); } catch(const linaccess::VbaRuntimeError& error) { failure=error.number; } assert(failure==3420);
        assert(sql.exec("DROP TABLE query_fields_schema_probe")); sql.finish();
        auto duplicate=database->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string("SELECT 1 AS same,2 AS same")}).object;
        failure=0; try { (void)duplicate->member("Fields",{}); } catch(const linaccess::VbaRuntimeError& error) { failure=error.number; } assert(failure==3251);
        auto action=database->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string("UPDATE inputs SET value=0")}).object;
        failure=0; try { (void)action->member("Fields",{}); } catch(const linaccess::VbaRuntimeError& error) { failure=error.number; } assert(failure==3251);
        assert(sql.exec("SELECT SUM(value) FROM inputs") && sql.next() && sql.value(0).toInt()==12); sql.finish();
        assert(metadataFrame->values.at("calls").asNumber()==beforeMetadata);
    }
    const auto typedRows=database->member("OpenRecordset",{linaccess::VbaValue::string(typedSource.toStdString())}).object;
    const auto enabled=typedRows->member("Fields",{linaccess::VbaValue::string("enabled")}).object;
    assert(enabled->member("Type",{}).asNumber()==1 && enabled->member("Size",{}).asNumber()==1);
    assert(enabled->member("Value",{}).kind==linaccess::VbaValue::Kind::Boolean && enabled->member("Value",{}).boolean);
    const auto falseRows=database->member("OpenRecordset",{linaccess::VbaValue::string("SELECT NativeFlag(0) AS enabled")}).object;
    const auto falseField=falseRows->member("Fields",{linaccess::VbaValue::string("enabled")}).object;
    assert(falseField->member("Type",{}).asNumber()==1 && falseField->member("Value",{}).kind==linaccess::VbaValue::Kind::Boolean && !falseField->member("Value",{}).boolean);
    falseRows->member("Close",{});

    const auto amount=typedRows->member("Fields",{linaccess::VbaValue::string("amount")}).object;
    const auto stamp=typedRows->member("Fields",{linaccess::VbaValue::string("stamp")}).object;
    const auto integer=typedRows->member("Fields",{linaccess::VbaValue::string("integer_value")}).object;
    assert(amount->member("Type",{}).asNumber()==5 && amount->member("Value",{}).kind==linaccess::VbaValue::Kind::Currency && amount->member("Value",{}).currencyUnits==1000);
    assert(stamp->member("Type",{}).asNumber()==8 && stamp->member("Value",{}).kind==linaccess::VbaValue::Kind::Date && stamp->member("Value",{}).asText()=="2024-02-29 12:34:56");
    assert(integer->member("Type",{}).asNumber()==4 && integer->member("Value",{}).numericType==linaccess::VbaValue::NumericType::Long);
    assert(typedRows->member("Fields",{linaccess::VbaValue::string("enabled")}).object->member("Value",{}).kind==linaccess::VbaValue::Kind::Boolean); typedRows->member("Close",{});
    const auto moneyResult=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT NativeMoney()+CCur('0.2') AS amount"),{});
    assert(moneyResult.types.value("amount")=="CURRENCY_REAL" && linaccess::qtFieldValue(moneyResult.rows[0][0],"CURRENCY_REAL").currencyUnits==3000);
    assert(sql.exec("CREATE TABLE udf_money_target(amount CURRENCY_REAL,stamp DATETIME)")); sql.finish();
    database->member("Execute",{linaccess::VbaValue::string("INSERT INTO udf_money_target(amount,stamp) VALUES(NativeMoney()+CCur('0.2'),NativeDate())")});
    assert(rows.exec("SELECT amount,stamp FROM udf_money_target") && rows.next() && std::abs(rows.value(0).toDouble()-0.3)<1e-12 && rows.value(1).toString()=="2024-02-29 12:34:56"); rows.finish();
    const auto guardedPlan=linaccess::filterAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT NativeMoney() AS amount"),{});
    QSqlQuery guarded(db); assert(guarded.prepare(guardedPlan.sql)); assert(guarded.exec() && guarded.next()); guarded.finish();
    auto changedSignature=module; changedSignature.replace("Public Function NativeMoney() As Currency","Public Function NativeMoney() As Double");
    sql.prepare("UPDATE _la_access_artifacts SET source_bytes=? WHERE name='NativeMath'"); sql.addBindValue(changedSignature); assert(sql.exec()); sql.finish();
    assert(!guarded.exec()); assert(linaccess::jetSqlFunctionErrorNumber(guarded.lastError().text().toCaseFolded().toStdString())==3251); guarded.finish();
    sql.prepare("UPDATE _la_access_artifacts SET source_bytes=? WHERE name='NativeMath'"); sql.addBindValue(module); assert(sql.exec()); sql.finish();
    assert(guarded.exec() && guarded.next() && std::abs(guarded.value(0).toDouble()-0.1)<1e-12); guarded.clear();
    database->member("Execute",{linaccess::VbaValue::string("SELECT NativeMoney() AS amount,NativeDate() AS stamp INTO typed_udf_copy FROM inputs")});
    assert(database->member("RecordsAffected",{}).asNumber()==3);
    assert(rows.exec("PRAGMA table_info(typed_udf_copy)") && rows.next() && rows.value(2).toString()=="CURRENCY_REAL");
    assert(rows.next() && rows.value(2).toString()=="DATETIME"); rows.finish();
    const auto copied=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT * FROM typed_udf_copy"),{});
    assert(copied.rows.size()==3 && copied.columnTypes[0]=="CURRENCY_REAL" && copied.columnTypes[1]=="DATETIME");
    assert(linaccess::qtFieldValue(copied.rows[0][0],"CURRENCY_REAL").currencyUnits==1000);
    assert(copied.rows[0][1].toString()=="2024-02-29 12:34:56");
    int makeTableError=0;
    try { database->member("Execute",{linaccess::VbaValue::string("SELECT NativeMoney() AS amount,NativeNeedsUI() AS blocked INTO failed_udf_copy FROM inputs")}); }
    catch(const linaccess::VbaRuntimeError& failure) { makeTableError=failure.number; }
    assert(makeTableError==3251 && !db.tables().contains("failed_udf_copy"));
    assert(sql.exec("CREATE TEMP VIEW _la_access_artifacts AS SELECT * FROM main._la_access_artifacts")); sql.finish();
    assert(!rows.exec("SELECT NativeScale(3)"));
    assert(linaccess::jetSqlFunctionErrorNumber(rows.lastError().text().toCaseFolded().toStdString())==3251); rows.finish();
    assert(sql.exec("DROP VIEW temp._la_access_artifacts")); sql.finish();
    const auto sharedFrame=linaccess::standardVbaModuleState(db,module.toStdString());
    const auto beforeBlockedCall=sharedFrame->values.at("calls").asNumber();
    sharedFrame->values["held"]=linaccess::VbaValue::reference(linaccess::makeDaoDatabase(db));
    assert(!rows.exec("SELECT NativeScale(3)"));
    assert(linaccess::jetSqlFunctionErrorNumber(rows.lastError().text().toCaseFolded().toStdString())==3251);
    assert(sharedFrame->values.at("calls").asNumber()==beforeBlockedCall); rows.finish();
    sharedFrame->values["held"]=linaccess::VbaValue::reference(nullptr);
    QSqlQuery prepared(db); assert(prepared.prepare("SELECT NativeScale(?)")); prepared.addBindValue(3);
    assert(prepared.exec() && prepared.next() && prepared.value(0).toInt()==6); prepared.finish();
    module.replace("SQLDoubleHelper=value*2","SQLDoubleHelper=value*3");
    sql.prepare("UPDATE _la_access_artifacts SET source_bytes=? WHERE name='NativeMath'"); sql.addBindValue(module); assert(sql.exec()); sql.finish();
    assert(prepared.exec() && prepared.next() && prepared.value(0).toInt()==9); prepared.finish();
    assert(sql.exec("SAVEPOINT native_source_edit")); sql.finish();
    auto transientModule=module; transientModule.replace("SQLDoubleHelper=value*3","SQLDoubleHelper=value*4");
    sql.prepare("UPDATE _la_access_artifacts SET source_bytes=? WHERE name='NativeMath'"); sql.addBindValue(transientModule); assert(sql.exec()); sql.finish();
    assert(prepared.exec() && prepared.next() && prepared.value(0).toInt()==12); prepared.finish();
    assert(sql.exec("ROLLBACK TO native_source_edit")); assert(sql.exec("RELEASE native_source_edit")); sql.finish();
    assert(prepared.exec() && prepared.next() && prepared.value(0).toInt()==9); prepared.finish();
    store("AmbiguousMath","Public Function NativeScale(ByVal value As Long) As Long\nNativeScale=999\nEnd Function\n");
    linaccess::installJetSqlFunctions(db);
    assert(!rows.exec("SELECT NativeScale(3)")); assert(linaccess::jetSqlFunctionErrorNumber(rows.lastError().text().toCaseFolded().toStdString())==3251); rows.clear(); prepared.clear(); sql.clear();
    db.close(); assert(db.open()); sql=QSqlQuery(db); rows=QSqlQuery(db);
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    store("NativeMath",module); linaccess::installJetSqlFunctions(db);
    assert(rows.exec("SELECT NativeSequence()") && rows.next() && rows.value(0).toInt()==1); rows.finish();
    assert(sql.exec("DELETE FROM _la_access_artifacts")); sql.finish();
    assert(!rows.exec("SELECT NativeScale(3)")); assert(linaccess::jetSqlFunctionErrorNumber(rows.lastError().text().toCaseFolded().toStdString())==3251);
}

void testDetachedTableCreation() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","dao-detached-schema");db.setDatabaseName(":memory:");assert(db.open());
    Host host(db);
    const auto result=linaccess::executeVbaProcedure(R"VBA(Sub BuildSchema()
Dim database As DAO.Database
Dim table As DAO.TableDef
Dim field As DAO.Field
Dim index As DAO.Index
Set database=CurrentDb
Set table=database.CreateTableDef("OwnTable")
Set field=table.CreateField("ID",dbLong)
field.Attributes=dbAutoIncrField
table.Fields.Append field
Set field=table.CreateField("Label",dbText,12)
field.Required=True
field.DefaultValue="'new'"
table.Fields.Append field
Set field=table.CreateField("Enabled",dbBoolean)
table.Fields.Append field
Set index=table.CreateIndex("PrimaryKey")
index.Primary=True
index.Fields.Append index.CreateField("ID")
table.Indexes.Append index
Set index=table.CreateIndex("LabelDescending")
index.Unique=True
Set field=index.CreateField("Label")
field.Attributes=dbDescending
index.Fields.Append field
table.Indexes.Append index
database.TableDefs.Append table
database.Execute "INSERT INTO OwnTable (Enabled) VALUES (True)"
End Sub)VBA","BuildSchema",host);
    if(result.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Schema builder error="<<result.errorNumber<<", line="<<result.line<<", reason="<<result.explanation<<"\n";
    assert(result.status==linaccess::VbaExecutionStatus::Completed);
    QSqlQuery query(db);assert(query.exec("SELECT ID,Label,Enabled FROM OwnTable") && query.next());
    assert(query.value(0).toInt()==1 && query.value(1).toString()=="new" && query.value(2).toInt()==-1);query.finish();
    auto database=linaccess::makeDaoDatabase(db),tables=database->member("TableDefs",{}).object;
    auto table=tables->member("Item",{VbaValue::string("OwnTable")}).object;
    auto field=table->member("Fields",{VbaValue::string("Label")}).object;
    assert(field->member("Size",{}).asNumber()==12);
    assert(field->member("Required",{}).truthy() && field->member("Attributes",{}).asNumber()==2);
    assert(!field->member("DataUpdatable",{}).truthy());
    assert(table->member("Fields",{VbaValue::string("Enabled")}).object->member("Attributes",{}).asNumber()==1);
    assert(!table->member("Fields",{VbaValue::string("Enabled")}).object->member("Required",{}).truthy());
    assert(!table->member("Fields",{VbaValue::string("ID")}).object->member("Required",{}).truthy());
    auto indexes=table->member("Indexes",{}).object;
    assert(indexes->member("Count",{}).asNumber()==2);
    auto primary=indexes->member("Item",{VbaValue::string("PrimaryKey")}).object;
    assert(primary->member("Primary",{}).truthy() && primary->member("Unique",{}).truthy());
    auto descending=indexes->member("Item",{VbaValue::string("LabelDescending")}).object;
    assert(descending->member("Unique",{}).truthy());
    assert(descending->member("Fields",{VbaValue::string("Label")}).object->member("Attributes",{}).asNumber()==1);
    const auto error=[&](const std::function<void()>& operation,int expected) {int number=0;try {operation();} catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;}assert(number==expected);};
    error([&] {database->member("Execute",{VbaValue::string("INSERT INTO OwnTable (Label) VALUES ('1234567890123')")});},3163);
    error([&] {database->member("Execute",{VbaValue::string("INSERT INTO OwnTable (Label) VALUES ('new')")});},3022);
    // Existing-table index mutation preserves data and uses the same checked catalog.
    auto extra=table->member("CreateIndex",{VbaValue::string("EnabledThenID")}).object;
    auto extraFields=extra->member("Fields",{}).object;
    extraFields->member("Append",{extra->member("CreateField",{VbaValue::string("Enabled")})});
    auto descendingID=extra->member("CreateField",{VbaValue::string("ID")}).object;descendingID->assign("Attributes",VbaValue::longInteger(1));
    extraFields->member("Append",{VbaValue::reference(descendingID)});
    table->member("Indexes",{}).object->member("Append",{VbaValue::reference(extra)});
    assert(extra->member("Name",{}).asText()=="EnabledThenID");
    error([&] {primary->member("Name",{});},3420);
    auto liveIndexes=table->member("Indexes",{}).object;assert(liveIndexes->member("Count",{}).asNumber()==3);
    database->member("Execute",{VbaValue::string("INSERT INTO OwnTable (Label,Enabled) VALUES ('two',True)")});
    auto rejected=table->member("CreateIndex",{VbaValue::string("DuplicateEnabled")}).object;
    rejected->assign("Unique",VbaValue::logical(true));
    rejected->member("Fields",{}).object->member("Append",{rejected->member("CreateField",{VbaValue::string("Enabled")})});
    error([&] {liveIndexes->member("Append",{VbaValue::reference(rejected)});},3022);
    assert(table->member("Indexes",{}).object->member("Count",{}).asNumber()==3);
    table->member("Indexes",{}).object->member("Delete",{VbaValue::string("EnabledThenID")});
    error([&] {extra->member("Name",{});},3420);
    assert(table->member("Indexes",{}).object->member("Count",{}).asNumber()==2);
    table->member("Indexes",{}).object->member("Delete",{VbaValue::string("PrimaryKey")});assert(table->member("Indexes",{}).object->member("Count",{}).asNumber()==1);
    auto restoredPrimary=table->member("CreateIndex",{VbaValue::string("PrimaryKey")}).object;restoredPrimary->assign("Primary",VbaValue::logical(true));restoredPrimary->member("Fields",{}).object->member("Append",{restoredPrimary->member("CreateField",{VbaValue::string("ID")})});table->member("Indexes",{}).object->member("Append",{VbaValue::reference(restoredPrimary)});assert(restoredPrimary->member("Primary",{}).truthy());
    auto rollbackIndex=table->member("CreateIndex",{VbaValue::string("RolledBack")}).object;
    rollbackIndex->member("Fields",{}).object->member("Append",{rollbackIndex->member("CreateField",{VbaValue::string("ID")})});
    auto workspace=linaccess::makeDaoEngine(db)->defaultMember({VbaValue::integer(0)},true).object;
    workspace->member("BeginTrans",{});
    table->member("Indexes",{}).object->member("Append",{VbaValue::reference(rollbackIndex)});
    workspace->member("Rollback",{});
    error([&] {rollbackIndex->member("Name",{});},3420);
    assert(table->member("Indexes",{}).object->member("Count",{}).asNumber()==2);
    // The earlier field wrapper predates index DDL; reacquire it for later deletion checks.
    field=table->member("Fields",{VbaValue::string("Label")}).object;
    auto pending=database->member("CreateTableDef",{VbaValue::string("Pending")}).object;
    auto fields=pending->member("Fields",{}).object;
    auto first=pending->member("CreateField",{VbaValue::string("A"),VbaValue::integer(4)}).object;
    auto second=pending->member("CreateField",{VbaValue::string("B"),VbaValue::integer(4)}).object;
    fields->member("Append",{VbaValue::reference(first)});fields->member("Append",{VbaValue::reference(second)});
    second->assign("Name",VbaValue::string("a"));
    error([&] {tables->member("Append",{VbaValue::reference(pending)});},3191);
    assert(!db.tables().contains("Pending"));second->assign("Name",VbaValue::string("B"));
    tables->member("Append",{VbaValue::reference(pending)});
    assert(fields->enumerate().size()==2);
    tables->member("Delete",{VbaValue::string("Pending")});
    error([&] {fields->enumerate();},3420);
    tables->member("Delete",{VbaValue::string("OwnTable")});
    error([&] {field->member("Name",{});},3420);
    // Force a physical index namespace collision after counter/table creation.
    // The enclosing schema transaction must remove every partial object.
    assert(query.exec("CREATE TABLE Occupied(value INTEGER)"));
    assert(query.exec("CREATE INDEX Collision_Key ON Occupied(value)"));
    auto collision=database->member("CreateTableDef",{VbaValue::string("Collision")}).object;
    auto identity=collision->member("CreateField",{VbaValue::string("ID"),VbaValue::integer(4),VbaValue::integer(99)}).object;
    assert(identity->member("Size",{}).asNumber()==4);
    identity->assign("Attributes",VbaValue::longInteger(16));
    collision->member("Fields",{}).object->member("Append",{VbaValue::reference(identity)});
    auto index=collision->member("CreateIndex",{VbaValue::string("Key")}).object;
    index->member("Fields",{}).object->member("Append",{index->member("CreateField",{VbaValue::string("ID")})});
    collision->member("Indexes",{}).object->member("Append",{VbaValue::reference(index)});
    error([&] {tables->member("Append",{VbaValue::reference(collision)});},3010);
    assert(!db.tables().contains("Collision"));
    assert(query.exec("SELECT COUNT(*) FROM _la_counter32_sequences WHERE table_name='Collision'") && query.next() && query.value(0).toInt()==0);query.finish();
    assert(query.exec("DROP INDEX Collision_Key"));
    tables->member("Append",{VbaValue::reference(collision)});
    assert(collision->member("Indexes",{}).object->member("Count",{}).asNumber()==1);
    database->member("Execute",{VbaValue::string("INSERT INTO Collision DEFAULT VALUES")});
    assert(query.exec("SELECT ID FROM Collision") && query.next() && query.value(0).toInt()==1);query.finish();
    auto ranges=database->member("CreateTableDef",{VbaValue::string("Ranges")}).object;
    auto rangeFields=ranges->member("Fields",{}).object;
    for(const auto& definition:std::vector<std::pair<std::string,int>>{{"SmallByte",2},{"SmallInteger",3},{"LargeInteger",4},{"Magnitude",7},{"Moment",8},{"Notes",12},{"Payload",11}}) rangeFields->member("Append",{ranges->member("CreateField",{VbaValue::string(definition.first),VbaValue::integer(static_cast<std::int16_t>(definition.second))})});
    tables->member("Append",{VbaValue::reference(ranges)});
    // INSERT SELECT bypasses the bounded VALUES-conversion path: real column
    // constraints must still prevent unrepresentable primitive storage.
    error([&] {database->member("Execute",{VbaValue::string("INSERT INTO Ranges(SmallByte) SELECT 256")});},6);
    error([&] {database->member("Execute",{VbaValue::string("INSERT INTO Ranges(SmallInteger) SELECT 32768")});},6);
    error([&] {database->member("Execute",{VbaValue::string("INSERT INTO Ranges(LargeInteger) SELECT 2147483648")});},6);
    database->member("Execute",{VbaValue::string("INSERT INTO Ranges(SmallByte,SmallInteger,LargeInteger) SELECT 255,-32768,2147483647")});
    assert(query.exec("SELECT SmallByte,SmallInteger,LargeInteger FROM Ranges") && query.next());assert(query.value(0).toInt()==255 && query.value(1).toInt()==-32768 && query.value(2).toInt()==2147483647);query.finish();
    query.prepare("INSERT INTO Ranges(Magnitude,Moment,Notes,Payload) VALUES(?,?,?,?)");
    query.addBindValue(std::numeric_limits<double>::max());query.addBindValue(QString("2026-11-08 00:00:00"));query.addBindValue(QString::fromUtf8("Technische Notiz\nÄÖÜ"));query.addBindValue(QByteArray::fromHex("0001ff"));assert(query.exec());
    assert(query.exec("SELECT Magnitude,Moment,Notes,Payload FROM Ranges WHERE Payload IS NOT NULL") && query.next());
    assert(query.value(0).toDouble()==std::numeric_limits<double>::max() && query.value(1).toString()=="2026-11-08 00:00:00" && query.value(2).toString()==QString::fromUtf8("Technische Notiz\nÄÖÜ") && query.value(3).toByteArray()==QByteArray::fromHex("0001ff"));query.finish();
    assert(!query.exec("INSERT INTO Ranges(Magnitude) VALUES('invalid')"));
    assert(!query.exec("INSERT INTO Ranges(Moment) VALUES('invalid')"));
    assert(!query.exec("INSERT INTO Ranges(Payload) VALUES('invalid')"));
    assert(!query.exec("INSERT INTO Ranges(Notes) VALUES(X'ff')"));
    assert(query.exec("SELECT COUNT(*) FROM Ranges") && query.next() && query.value(0).toInt()==2);query.finish();
    const auto altered=linaccess::executeVbaProcedure(R"VBA(Sub AlterSchema()
Dim table As DAO.TableDef
Dim field As DAO.Field
Set table=CurrentDb.CreateTableDef("Alterable")
table.Fields.Append table.CreateField("Original",dbLong)
CurrentDb.TableDefs.Append table
CurrentDb.Execute "INSERT INTO Alterable VALUES (41)"
Set table=CurrentDb.TableDefs("Alterable")
Set field=table.CreateField(Name:="Added",Type:=dbText,Size:=8)
field.DefaultValue="'initial'"
table.Fields.Append field
End Sub)VBA","AlterSchema",host);
    if(altered.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Alter builder error="<<altered.errorNumber<<", line="<<altered.line<<", reason="<<altered.explanation<<"\n";
    assert(altered.status==linaccess::VbaExecutionStatus::Completed);
    assert(query.exec("SELECT Original,Added FROM Alterable") && query.next() && query.value(0).toInt()==41 && query.value(1).toString()=="initial");query.finish();
    auto alterable=tables->member("Item",{VbaValue::string("Alterable")}).object;
    assert(alterable->member("Fields",{VbaValue::string("Added")}).object->member("Size",{}).asNumber()==8);
    assert(alterable->member("Indexes",{}).object->member("Count",{}).asNumber()==0);
    auto oldRows=database->member("OpenRecordset",{VbaValue::string("Alterable")}).object;
    auto oldValue=oldRows->member("Fields",{VbaValue::string("Added")}).object;
    auto oldClone=oldRows->member("Clone",{}).object;oldClone->member("MoveFirst",{});oldClone->member("Edit",{});oldClone->assignItem("Added",VbaValue::string("pending"));
    auto extraField=alterable->member("CreateField",{VbaValue::string("Extra"),VbaValue::integer(4)}).object;
    alterable->member("Fields",{}).object->member("Append",{VbaValue::reference(extraField)});
    error([&] {oldRows->member("EOF",{});},3420);error([&] {oldValue->member("Value",{});},3420);error([&] {oldClone->member("Update",{});},3420);
    oldRows->member("Requery",{});assert(oldRows->member("Fields",{}).object->member("Count",{}).asNumber()==3 && oldClone->member("EditMode",{}).asNumber()==0);
    assert(oldRows->item("Added").asText()=="initial");oldClone->member("Close",{});oldRows->member("Close",{});
    alterable->member("Fields",{}).object->member("Delete",{VbaValue::string("Extra")});
    auto dependency=alterable->member("CreateIndex",{VbaValue::string("AddedIndex")}).object;
    dependency->member("Fields",{}).object->member("Append",{dependency->member("CreateField",{VbaValue::string("Added")})});
    alterable->member("Indexes",{}).object->member("Append",{VbaValue::reference(dependency)});
    error([&] {alterable->member("Fields",{}).object->member("Delete",{VbaValue::string("Added")});},3251);
    alterable->member("Indexes",{}).object->member("Delete",{VbaValue::string("AddedIndex")});
    auto heldAdded=alterable->member("Fields",{VbaValue::string("Added")}).object;
    alterable->member("Fields",{}).object->member("Delete",{VbaValue::string("Added")});
    error([&] {heldAdded->member("Name",{});},3420);
    assert(alterable->member("Fields",{}).object->member("Count",{}).asNumber()==1 && alterable->member("Indexes",{}).object->member("Count",{}).asNumber()==0);
    auto requiredField=alterable->member("CreateField",{VbaValue::string("RequiredNoDefault"),VbaValue::integer(4)}).object;requiredField->assign("Required",VbaValue::logical(true));
    error([&] {alterable->member("Fields",{}).object->member("Append",{VbaValue::reference(requiredField)});},3314);
    auto counterField=alterable->member("CreateField",{VbaValue::string("Counter"),VbaValue::integer(4)}).object;counterField->assign("Attributes",VbaValue::longInteger(16));
    error([&] {alterable->member("Fields",{}).object->member("Append",{VbaValue::reference(counterField)});},3251);
    auto reversible=alterable->member("CreateField",{VbaValue::string("Reverted"),VbaValue::integer(2)}).object;
    workspace->member("BeginTrans",{});alterable->member("Fields",{}).object->member("Append",{VbaValue::reference(reversible)});workspace->member("Rollback",{});
    error([&] {reversible->member("Size",{});},3420);
    assert(alterable->member("Fields",{}).object->member("Count",{}).asNumber()==1);
    error([&] {alterable->member("Fields",{}).object->member("Delete",{VbaValue::string("Original")});},3251);
    assert(query.exec("SELECT Original FROM Alterable") && query.next() && query.value(0).toInt()==41);query.finish();
    auto heldFields=alterable->member("Fields",{}).object;
    auto heldOriginal=heldFields->member("Item",{VbaValue::string("Original")}).object;
    auto heldIndexes=alterable->member("Indexes",{}).object;
    auto refreshField=alterable->member("CreateField",{VbaValue::string("RefreshField"),VbaValue::integer(4)}).object;
    heldFields->member("Append",{VbaValue::reference(refreshField)});
    error([&] {heldFields->member("Count",{});},3420);heldFields->member("Refresh",{});
    assert(heldFields->member("Count",{}).asNumber()==2);error([&] {heldOriginal->member("Name",{});},3420);
    heldIndexes->member("Refresh",{});assert(heldIndexes->member("Count",{}).asNumber()==0);
    auto refreshIndex=alterable->member("CreateIndex",{VbaValue::string("RefreshIndex")}).object;
    refreshIndex->member("Fields",{}).object->member("Append",{refreshIndex->member("CreateField",{VbaValue::string("Original")})});
    heldIndexes->member("Append",{VbaValue::reference(refreshIndex)});
    error([&] {heldIndexes->member("Count",{});},3420);heldIndexes->member("Refresh",{});
    auto oldIndex=heldIndexes->member("Item",{VbaValue::string("RefreshIndex")}).object;
    assert(heldIndexes->member("Count",{}).asNumber()==1);
    heldIndexes->member("Delete",{VbaValue::string("RefreshIndex")});heldIndexes->member("Refresh",{});
    assert(heldIndexes->member("Count",{}).asNumber()==0);error([&] {oldIndex->member("Name",{});},3420);
    heldFields->member("Refresh",{});
    auto temporaryField=alterable->member("CreateField",{VbaValue::string("Temporary"),VbaValue::integer(2)}).object;
    workspace->member("BeginTrans",{});heldFields->member("Append",{VbaValue::reference(temporaryField)});heldFields->member("Refresh",{});
    auto oldTemporary=heldFields->member("Item",{VbaValue::string("Temporary")}).object;assert(heldFields->member("Count",{}).asNumber()==3);
    workspace->member("Rollback",{});heldFields->member("Refresh",{});assert(heldFields->member("Count",{}).asNumber()==2);error([&] {oldTemporary->member("Name",{});},3420);
    tables->member("Delete",{VbaValue::string("Alterable")});
    auto replacement=database->member("CreateTableDef",{VbaValue::string("Alterable")}).object;
    replacement->member("Fields",{}).object->member("Append",{replacement->member("CreateField",{VbaValue::string("Original"),VbaValue::integer(4)})});tables->member("Append",{VbaValue::reference(replacement)});
    error([&] {heldFields->member("Refresh",{});},3420);error([&] {heldIndexes->member("Refresh",{});},3420);
}

void testDetachedRelationCreation() {
    using linaccess::VbaValue;
    QTemporaryDir directory;assert(directory.isValid());
    auto db=QSqlDatabase::addDatabase("QSQLITE","dao-created-relations");db.setDatabaseName(directory.filePath("relations.linaccess"));assert(db.open());
    Host host(db);
    const auto result=linaccess::executeVbaProcedure(R"VBA(Sub BuildRelations()
Dim table As DAO.TableDef, relation As DAO.Relation, field As DAO.Field
Set table=CurrentDb.CreateTableDef("RelParent")
table.Fields.Append table.CreateField("ID",dbLong)
CurrentDb.TableDefs.Append table
Set table=CurrentDb.CreateTableDef("RelChild")
table.Fields.Append table.CreateField("ID",dbLong)
table.Fields.Append table.CreateField("ParentID",dbLong)
CurrentDb.TableDefs.Append table
Set relation=CurrentDb.CreateRelation(Name:="ParentChild",Table:="RelParent",ForeignTable:="RelChild",Attributes:=dbRelationDontEnforce)
Set field=relation.CreateField("ID")
field.ForeignName="ParentID"
relation.Fields.Append field
CurrentDb.Relations.Append relation
If relation.Attributes<>dbRelationDontEnforce Or relation.Fields(0).ForeignName<>"ParentID" Then Err.Raise 5
CurrentDb.Execute "INSERT INTO RelParent VALUES (1)"
CurrentDb.Execute "INSERT INTO RelChild VALUES (1,999)"
End Sub)VBA","BuildRelations",host);
    if(result.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Relation builder error="<<result.errorNumber<<", line="<<result.line<<", reason="<<result.explanation<<"\n";
    assert(result.status==linaccess::VbaExecutionStatus::Completed);
    auto database=linaccess::makeDaoDatabase(db),relations=database->member("Relations",{}).object;
    assert(relations->member("Count",{}).asNumber()==1);
    auto original=relations->member("Item",{VbaValue::string("ParentChild")}).object;
    auto oldField=original->member("Fields",{VbaValue::string("ID")}).object;
    assert(oldField->member("ForeignName",{}).asText()=="ParentID");
    QSqlQuery sql(db);assert(sql.exec("PRAGMA foreign_key_list(RelChild)") && !sql.next());sql.finish();
    assert(sql.exec("SELECT ParentID FROM RelChild") && sql.next() && sql.value(0).toInt()==999);sql.finish();
    const auto error=[](const std::function<void()>& operation,int expected) {int number=0;try {operation();}catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;}assert(number==expected);};
    error([&] {database->member("Execute",{VbaValue::string("DROP TABLE RelParent")});},3251);
    auto child=database->member("TableDefs",{VbaValue::string("RelChild")}).object;
    error([&] {child->member("Fields",{}).object->member("Delete",{VbaValue::string("ParentID")});},3251);
    const auto build=[&](const std::string& name,const std::string& foreign,int attributes) {
        auto relation=database->member("CreateRelation",{VbaValue::string(name),VbaValue::string("RelParent"),VbaValue::string("RelChild"),VbaValue::longInteger(attributes)}).object;
        auto field=relation->member("CreateField",{VbaValue::string("ID")}).object;field->assign("ForeignName",VbaValue::string(foreign));relation->member("Fields",{}).object->member("Append",{VbaValue::reference(field)});return relation;
    };
    auto unimplemented=build("Enforced","ParentID",0);error([&] {relations->member("Append",{VbaValue::reference(unimplemented)});},3251);
    auto missing=build("MissingField","Absent",2);error([&] {relations->member("Append",{VbaValue::reference(missing)});},3265);
    assert(relations->member("Count",{}).asNumber()==1);
    relations->member("Delete",{VbaValue::string("ParentChild")});error([&] {original->member("Name",{});},3420);error([&] {oldField->member("Name",{});},3420);
    auto replacement=build("ParentChild","ParentID",2);relations->member("Append",{VbaValue::reference(replacement)});
    auto fresh=relations->member("Item",{VbaValue::string("ParentChild")}).object;assert(fresh!=original);
    error([&] {original->member("Name",{});},3420);error([&] {oldField->member("Name",{});},3420);
    auto workspace=linaccess::makeDaoEngine(db)->defaultMember({VbaValue::integer(0)},true).object;
    workspace->member("BeginTrans",{});relations->member("Delete",{VbaValue::string("ParentChild")});workspace->member("Rollback",{});
    assert(relations->member("Count",{}).asNumber()==1);error([&] {fresh->member("Name",{});},3420);
    auto independent=QSqlDatabase::addDatabase("QSQLITE","dao-created-relations-reader");independent.setDatabaseName(db.databaseName());assert(independent.open());
    QSqlQuery observed(independent);assert(observed.exec("SELECT name,flags,parent_table,child_table,parent_column,child_column FROM _la_access_relations") && observed.next());
    assert(observed.value(0).toString()=="ParentChild" && observed.value(1).toInt()==2 && observed.value(2).toString()=="RelParent" && observed.value(3).toString()=="RelChild" && observed.value(4).toString()=="ID" && observed.value(5).toString()=="ParentID");observed.finish();independent.close();
    relations->member("Delete",{VbaValue::string("ParentChild")});
    child=database->member("TableDefs",{VbaValue::string("RelChild")}).object;child->member("Fields",{}).object->member("Delete",{VbaValue::string("ParentID")});
    database->member("TableDefs",{}).object->member("Delete",{VbaValue::string("RelParent")});
    assert(sql.exec("SELECT ID FROM RelChild") && sql.next() && sql.value(0).toInt()==1);sql.finish();
}

void testEnforcedNativeRelations() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","dao-native-foreign-key-migration");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);assert(sql.exec("PRAGMA foreign_keys=ON"));
    auto database=linaccess::makeDaoDatabase(db),tables=database->member("TableDefs",{}).object;
    const auto field=[&](const std::shared_ptr<linaccess::VbaObject>& table,const char* name,int type,int size=0) {return table->member("CreateField",{VbaValue::string(name),VbaValue::integer(static_cast<std::int16_t>(type)),VbaValue::longInteger(size)}).object;};
    const auto primary=[&](const std::shared_ptr<linaccess::VbaObject>& table) {auto key=table->member("CreateIndex",{VbaValue::string("PrimaryKey")}).object;key->assign("Primary",VbaValue::logical(true));key->member("Fields",{}).object->member("Append",{key->member("CreateField",{VbaValue::string("ID")})});table->member("Indexes",{}).object->member("Append",{VbaValue::reference(key)});};
    auto parent=database->member("CreateTableDef",{VbaValue::string("RIParent")}).object;parent->member("Fields",{}).object->member("Append",{VbaValue::reference(field(parent,"ID",4))});primary(parent);tables->member("Append",{VbaValue::reference(parent)});
    auto child=database->member("CreateTableDef",{VbaValue::string("RIChild")}).object;
    auto id=field(child,"ID",4);id->assign("Attributes",VbaValue::longInteger(16));child->member("Fields",{}).object->member("Append",{VbaValue::reference(id)});
    for(const auto& definition:std::vector<std::pair<const char*,int>>{{"ParentID",4},{"Payload",11},{"Note",10},{"Magnitude",7},{"Moment",8}}) {
        auto item=field(child,definition.first,definition.second,definition.second==10 ? 20 : 0);
        if(definition.second==7) item->assign("DefaultValue",VbaValue::string("1.25"));if(definition.second==8) item->assign("DefaultValue",VbaValue::string("'2026-11-08 00:00:00'"));
        child->member("Fields",{}).object->member("Append",{VbaValue::reference(item)});
    }
    primary(child);tables->member("Append",{VbaValue::reference(child)});
    auto lookup=child->member("CreateIndex",{VbaValue::string("ParentLookup")}).object;
    auto parentKey=lookup->member("CreateField",{VbaValue::string("ParentID")}).object;parentKey->assign("Attributes",VbaValue::longInteger(1));lookup->member("Fields",{}).object->member("Append",{VbaValue::reference(parentKey)});lookup->member("Fields",{}).object->member("Append",{lookup->member("CreateField",{VbaValue::string("ID")})});child->member("Indexes",{}).object->member("Append",{VbaValue::reference(lookup)});
    database->member("Execute",{VbaValue::string("ALTER TABLE RIChild ALTER COLUMN ID COUNTER(100,7)")});
    assert(sql.exec("CREATE TABLE migration_audit(value INTEGER)"));assert(sql.exec("CREATE TRIGGER migration_child_audit AFTER INSERT ON RIChild BEGIN INSERT INTO migration_audit VALUES(NEW.ID);END"));
    assert(sql.exec("CREATE VIEW migration_child_view AS SELECT rowid AS source_rowid,ID,ParentID,Payload,Note,Magnitude,Moment FROM RIChild"));
    database->member("Execute",{VbaValue::string("INSERT INTO RIParent VALUES (7)")});
    sql.prepare("INSERT INTO RIChild(ParentID,Payload,Note,Magnitude,Moment) VALUES(?,?,?,?,?)");
    for(int i=0;i<3;++i) {sql.bindValue(0,i==1 ? 999 : 7);sql.bindValue(1,QByteArray::fromHex("0001ff"));sql.bindValue(2,QString::fromUtf8("Prüfwert %1").arg(i));sql.bindValue(3,i==1 ? std::numeric_limits<double>::max() : 3.141592653589793);sql.bindValue(4,QString("2026-11-08 00:00:00"));assert(sql.exec());}sql.finish();
    database->member("Execute",{VbaValue::string("DELETE FROM RIChild WHERE ID=114")});
    auto oldRows=database->member("OpenRecordset",{VbaValue::string("RIChild")}).object;
    auto relations=database->member("Relations",{}).object;Host host(db);
    const auto cascadeFlags=linaccess::evaluateVbaExpression("dbRelationUpdateCascade Or dbRelationDeleteCascade",host,{});assert(cascadeFlags.asNumber()==4352);
    auto relation=database->member("CreateRelation",{VbaValue::string("RealForeignKey"),VbaValue::string("RIParent"),VbaValue::string("RIChild"),cascadeFlags}).object;
    auto link=relation->member("CreateField",{VbaValue::string("ID")}).object;link->assign("ForeignName",VbaValue::string("ParentID"));relation->member("Fields",{}).object->member("Append",{VbaValue::reference(link)});
    const auto error=[](const std::function<void()>& operation,int expected) {int number=0;try {operation();}catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;if(number!=expected) std::cerr<<"Unexpected migration error="<<number<<", reason="<<failure.what()<<"\n";}assert(number==expected);};
    error([&] {relations->member("Append",{VbaValue::reference(relation)});},3251);
    assert(oldRows->member("RecordCount",{}).asNumber()==2 && relations->member("Count",{}).asNumber()==0);
    assert(sql.exec("PRAGMA foreign_keys") && sql.next() && sql.value(0).toInt()==1);sql.finish();
    assert(sql.exec("PRAGMA legacy_alter_table") && sql.next() && sql.value(0).toInt()==0);sql.finish();
    assert(sql.exec("SELECT COUNT(*) FROM migration_audit") && sql.next() && sql.value(0).toInt()==3);sql.finish();
    database->member("Execute",{VbaValue::string("UPDATE RIChild SET ParentID=7 WHERE ID=107")});
    try {relations->member("Append",{VbaValue::reference(relation)});}catch(const linaccess::VbaRuntimeError& failure) {std::cerr<<"Native FK migration failed="<<failure.number<<", reason="<<failure.what()<<"\n";throw;}
    error([&] {oldRows->member("RecordCount",{});},3420);oldRows->member("Close",{});
    assert(sql.exec("PRAGMA foreign_key_list(RIChild)") && sql.next());assert(sql.value(2).toString()=="RIParent" && sql.value(5).toString()=="CASCADE" && sql.value(6).toString()=="CASCADE");sql.finish();
    assert(sql.exec("SELECT source_rowid,ID,Payload,Note,Magnitude FROM migration_child_view ORDER BY ID") && sql.next());
    assert(sql.value(0).toInt()==1 && sql.value(1).toInt()==100 && sql.value(2).toByteArray()==QByteArray::fromHex("0001ff") && sql.value(3).toString()==QString::fromUtf8("Prüfwert 0") && sql.value(4).toDouble()==3.141592653589793);
    assert(sql.next() && sql.value(0).toInt()==2 && sql.value(1).toInt()==107 && sql.value(4).toDouble()==std::numeric_limits<double>::max());sql.finish();
    assert(sql.exec("SELECT value FROM _la_counter32_sequences WHERE table_name='RIChild'") && sql.next() && sql.value(0).toInt()==114);sql.finish();
    assert(sql.exec("SELECT COUNT(*) FROM migration_audit") && sql.next() && sql.value(0).toInt()==3);sql.finish();
    child=tables->member("Item",{VbaValue::string("RIChild")}).object;
    assert(child->member("Indexes",{VbaValue::string("ParentLookup")}).object->member("Fields",{VbaValue::string("ParentID")}).object->member("Attributes",{}).asNumber()==1);
    database->member("Execute",{VbaValue::string("UPDATE RIParent SET ID=8 WHERE ID=7")});assert(sql.exec("SELECT COUNT(*) FROM RIChild WHERE ParentID=8") && sql.next() && sql.value(0).toInt()==2);sql.finish();
    database->member("Execute",{VbaValue::string("DELETE FROM RIParent WHERE ID=8")});assert(sql.exec("SELECT COUNT(*) FROM RIChild") && sql.next() && sql.value(0).toInt()==0);sql.finish();
    database->member("Execute",{VbaValue::string("INSERT INTO RIParent VALUES(9)")});
    error([&] {database->member("Execute",{VbaValue::string("INSERT INTO RIChild(ParentID) VALUES(999)")});},3146);
    database->member("Execute",{VbaValue::string("INSERT INTO RIChild(ParentID) VALUES(9)")});assert(sql.exec("SELECT ID FROM RIChild") && sql.next() && sql.value(0).toInt()==128);sql.finish();
    auto grandchild=database->member("CreateTableDef",{VbaValue::string("RIGrandChild")}).object;
    grandchild->member("Fields",{}).object->member("Append",{VbaValue::reference(field(grandchild,"ID",4))});grandchild->member("Fields",{}).object->member("Append",{VbaValue::reference(field(grandchild,"ChildID",4))});primary(grandchild);tables->member("Append",{VbaValue::reference(grandchild)});
    database->member("Execute",{VbaValue::string("INSERT INTO RIGrandChild VALUES(1,128)")});
    auto descendant=database->member("CreateRelation",{VbaValue::string("ChildGrandChild"),VbaValue::string("RIChild"),VbaValue::string("RIGrandChild"),VbaValue::longInteger(4096)}).object;
    auto descendantField=descendant->member("CreateField",{VbaValue::string("ID")}).object;descendantField->assign("ForeignName",VbaValue::string("ChildID"));descendant->member("Fields",{}).object->member("Append",{VbaValue::reference(descendantField)});relations->member("Append",{VbaValue::reference(descendant)});
    assert(relations->member("Count",{}).asNumber()==2);
    relations->member("Delete",{VbaValue::string("RealForeignKey")});
    assert(sql.exec("SELECT ChildID FROM RIGrandChild") && sql.next() && sql.value(0).toInt()==128);sql.finish();
    assert(sql.exec("PRAGMA foreign_key_list(RIChild)") && !sql.next());sql.finish();
    database->member("Execute",{VbaValue::string("INSERT INTO RIChild(ParentID) VALUES(999)")});assert(sql.exec("SELECT ID FROM RIChild WHERE ParentID=999") && sql.next() && sql.value(0).toInt()==135);sql.finish();
    assert(sql.exec("PRAGMA foreign_keys") && sql.next() && sql.value(0).toInt()==1);sql.finish();
    auto workspace=linaccess::makeDaoEngine(db)->defaultMember({VbaValue::integer(0)},true).object;
    workspace->member("BeginTrans",{});error([&] {relations->member("Delete",{VbaValue::string("ChildGrandChild")});},3251);workspace->member("Rollback",{});
    assert(relations->member("Count",{}).asNumber()==1);
    // Model a native table customized outside the supported schema writer.
    // Its extra validation must survive refusal rather than disappear in rebuild.
    assert(sql.exec("CREATE TABLE CustomChild(ParentID INTEGER4 CHECK(ParentID>0))"));assert(sql.exec("INSERT INTO CustomChild VALUES(9)"));
    const QJsonArray customFields{QJsonObject{{"name","ParentID"},{"source_type","INT"},{"source_origin","DAO_NATIVE"},{"source_size",0},{"native_attributes",1},{"native_required",false}}};
    sql.prepare("INSERT INTO _la_table_schema VALUES('CustomChild',?)");sql.addBindValue(QJsonDocument(customFields).toJson(QJsonDocument::Compact));assert(sql.exec());
    const QJsonObject customIndexes{{"version",1},{"source_table","CustomChild"},{"columns",QJsonArray{"ParentID"}},{"indexes",QJsonArray{}}};
    sql.prepare("INSERT INTO _la_access_indexes VALUES('CustomChild',?)");sql.addBindValue(QString::fromUtf8(QJsonDocument(customIndexes).toJson(QJsonDocument::Compact)));assert(sql.exec());sql.finish();
    auto custom=database->member("CreateRelation",{VbaValue::string("CustomGuard"),VbaValue::string("RIParent"),VbaValue::string("CustomChild")}).object;
    auto customField=custom->member("CreateField",{VbaValue::string("ID")}).object;customField->assign("ForeignName",VbaValue::string("ParentID"));custom->member("Fields",{}).object->member("Append",{VbaValue::reference(customField)});
    error([&] {relations->member("Append",{VbaValue::reference(custom)});},3251);
    assert(!sql.exec("INSERT INTO CustomChild VALUES(-1)"));assert(sql.exec("SELECT ParentID FROM CustomChild") && sql.next() && sql.value(0).toInt()==9);sql.finish();
    assert(sql.exec("PRAGMA foreign_keys") && sql.next() && sql.value(0).toInt()==1);sql.finish();
}

void testOneToOneNativeRelations() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","dao-native-one-to-one");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);assert(sql.exec("PRAGMA foreign_keys=ON"));
    auto database=linaccess::makeDaoDatabase(db),tables=database->member("TableDefs",{}).object;
    for(const auto& name:{"UniqueParent","UniqueChild"}) {
        auto table=database->member("CreateTableDef",{VbaValue::string(name)}).object;
        table->member("Fields",{}).object->member("Append",{table->member("CreateField",{VbaValue::string("ID"),VbaValue::integer(4)})});
        if(std::string(name)=="UniqueChild") table->member("Fields",{}).object->member("Append",{table->member("CreateField",{VbaValue::string("ParentID"),VbaValue::integer(4)})});
        auto key=table->member("CreateIndex",{VbaValue::string("PrimaryKey")}).object;key->assign("Primary",VbaValue::logical(true));key->member("Fields",{}).object->member("Append",{key->member("CreateField",{VbaValue::string("ID")})});table->member("Indexes",{}).object->member("Append",{VbaValue::reference(key)});tables->member("Append",{VbaValue::reference(table)});
    }
    auto relations=database->member("Relations",{}).object;
    const auto build=[&](const std::string& name,int flags) {auto relation=database->member("CreateRelation",{VbaValue::string(name),VbaValue::string("UniqueParent"),VbaValue::string("UniqueChild"),VbaValue::longInteger(flags)}).object;auto field=relation->member("CreateField",{VbaValue::string("ID")}).object;field->assign("ForeignName",VbaValue::string("ParentID"));relation->member("Fields",{}).object->member("Append",{VbaValue::reference(field)});return relation;};
    const auto error=[](const std::function<void()>& operation,int expected) {int number=0;try {operation();}catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;}assert(number==expected);};
    auto relation=build("UniqueMapping",1);error([&] {relations->member("Append",{VbaValue::reference(relation)});},3251);
    assert(relations->member("Count",{}).asNumber()==0);
    auto child=tables->member("Item",{VbaValue::string("UniqueChild")}).object;
    auto unique=child->member("CreateIndex",{VbaValue::string("ParentUnique")}).object;unique->assign("Unique",VbaValue::logical(true));unique->member("Fields",{}).object->member("Append",{unique->member("CreateField",{VbaValue::string("ParentID")})});child->member("Indexes",{}).object->member("Append",{VbaValue::reference(unique)});
    relations->member("Append",{VbaValue::reference(relation)});assert(relation->member("Attributes",{}).asNumber()==1);
    error([&] {tables->member("Item",{VbaValue::string("UniqueParent")}).object->member("Indexes",{}).object->member("Delete",{VbaValue::string("PrimaryKey")});},3251);
    database->member("Execute",{VbaValue::string("INSERT INTO UniqueParent VALUES(10)")});database->member("Execute",{VbaValue::string("INSERT INTO UniqueChild VALUES(1,10)")});
    error([&] {database->member("Execute",{VbaValue::string("INSERT INTO UniqueChild VALUES(2,10)")});},3022);
    error([&] {database->member("Execute",{VbaValue::string("INSERT INTO UniqueChild VALUES(2,999)")});},3146);
    child=tables->member("Item",{VbaValue::string("UniqueChild")}).object;error([&] {child->member("Indexes",{}).object->member("Delete",{VbaValue::string("ParentUnique")});},3251);
    auto parentKey=tables->member("Item",{VbaValue::string("UniqueParent")}).object->member("Fields",{VbaValue::string("ID")}).object;
    parentKey->assign("Name",VbaValue::string("Parent Key"));assert(parentKey->member("Name",{}).asText()=="Parent Key");
    child->member("Fields",{VbaValue::string("ParentID")}).object->assign("Name",VbaValue::string("Foreign Key"));
    assert(sql.exec("PRAGMA foreign_key_list(UniqueChild)") && sql.next() && sql.value(3).toString()=="Foreign Key" && sql.value(4).toString()=="Parent Key");sql.finish();
    auto renamedRelation=relations->member("Item",{VbaValue::string("UniqueMapping")}).object;
    auto renamedRelationField=renamedRelation->member("Fields",{VbaValue::string("Parent Key")}).object;assert(renamedRelationField->member("ForeignName",{}).asText()=="Foreign Key");
    error([&] {database->member("Execute",{VbaValue::string("INSERT INTO UniqueChild VALUES(2,999)")});},3146);
    parentKey=tables->member("Item",{VbaValue::string("UniqueParent")}).object->member("Fields",{VbaValue::string("Parent Key")}).object;parentKey->assign("Name",VbaValue::string("ID"));child->member("Fields",{VbaValue::string("Foreign Key")}).object->assign("Name",VbaValue::string("ParentID"));
    relations->member("Delete",{VbaValue::string("UniqueMapping")});database->member("Execute",{VbaValue::string("INSERT INTO UniqueChild VALUES(2,999)")});
    auto logical=build("LogicalUnique",3);relations->member("Append",{VbaValue::reference(logical)});assert(logical->member("Attributes",{}).asNumber()==3);
    child=tables->member("Item",{VbaValue::string("UniqueChild")}).object;error([&] {child->member("Indexes",{}).object->member("Delete",{VbaValue::string("ParentUnique")});},3251);
    relations->member("Delete",{VbaValue::string("LogicalUnique")});child->member("Indexes",{}).object->member("Delete",{VbaValue::string("ParentUnique")});
    assert(sql.exec("SELECT COUNT(*) FROM UniqueChild") && sql.next() && sql.value(0).toInt()==2);sql.finish();
}

void testNativeFieldPropertyChanges() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","dao-native-field-properties");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);
    auto database=linaccess::makeDaoDatabase(db),tables=database->member("TableDefs",{}).object;
    auto table=database->member("CreateTableDef",{VbaValue::string("PropertyRows")}).object;
    for(const auto& definition:std::vector<std::pair<const char*,int>>{{"ID",4},{"Label",10},{"Optional",4},{"Payload",11},{"Magnitude",7},{"Moment",8},{"Enabled",1}}) {
        auto field=table->member("CreateField",{VbaValue::string(definition.first),VbaValue::integer(static_cast<std::int16_t>(definition.second)),VbaValue::longInteger(definition.second==10 ? 12 : 0)}).object;
        if(std::string(definition.first)=="ID") field->assign("Attributes",VbaValue::longInteger(16));
        if(definition.second==10) field->assign("DefaultValue",VbaValue::string("'old'"));if(definition.second==7) field->assign("DefaultValue",VbaValue::string("1.25"));if(definition.second==8) field->assign("DefaultValue",VbaValue::string("'2026-11-08 00:00:00'"));
        table->member("Fields",{}).object->member("Append",{VbaValue::reference(field)});
    }
    auto primary=table->member("CreateIndex",{VbaValue::string("PrimaryKey")}).object;primary->assign("Primary",VbaValue::logical(true));primary->member("Fields",{}).object->member("Append",{primary->member("CreateField",{VbaValue::string("ID")})});table->member("Indexes",{}).object->member("Append",{VbaValue::reference(primary)});tables->member("Append",{VbaValue::reference(table)});
    sql.prepare("INSERT INTO PropertyRows(Label,Payload,Magnitude) VALUES(?,?,?)");for(const auto& value:{"keep","long value"}) {sql.bindValue(0,QString(value));sql.bindValue(1,QByteArray::fromHex("0001ff"));sql.bindValue(2,std::numeric_limits<double>::max());assert(sql.exec());}sql.finish();
    const auto error=[](const std::function<void()>& operation,int expected) {int number=0;try {operation();}catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;if(number!=expected) std::cerr<<"Field edit error="<<number<<", reason="<<failure.what()<<"\n";}assert(number==expected);};
    auto definition=tables->member("Item",{VbaValue::string("PropertyRows")}).object;
    auto label=definition->member("Fields",{VbaValue::string("Label")}).object;
    error([&] {label->assign("Size",VbaValue::longInteger(4));},3163);assert(label->member("Size",{}).asNumber()==12);
    linaccess::updateNativeFieldProperties(db,"PropertyRows","Optional",{{"Required",VbaValue::logical(false)}});
    error([&] {linaccess::updateNativeFieldProperties(db,"PropertyRows","Optional",{{"Required",VbaValue::logical(true)},{"DefaultValue",VbaValue::string("7")}});},3314);
    auto rows=database->member("OpenRecordset",{VbaValue::string("PropertyRows")}).object;auto clone=rows->member("Clone",{}).object;clone->member("MoveFirst",{});clone->member("Edit",{});clone->assignItem("Label",VbaValue::string("pending"));
    Host host(db);const auto changed=linaccess::executeVbaProcedure(R"VBA(Sub EditProperties()
Dim field As DAO.Field
Set field=CurrentDb.TableDefs("PropertyRows").Fields("Label")
field.Size=20
If field.Size<>20 Then Err.Raise 5
field.DefaultValue="'new'"
field.Required=True
If Not field.Required Then Err.Raise 5
End Sub)VBA","EditProperties",host);
    if(changed.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"VBA field properties="<<changed.errorNumber<<", line="<<changed.line<<", reason="<<changed.explanation<<"\n";
    assert(changed.status==linaccess::VbaExecutionStatus::Completed);
    error([&] {rows->member("RecordCount",{});},3420);error([&] {clone->member("Update",{});},3420);rows->member("Requery",{});assert(clone->member("EditMode",{}).asNumber()==0 && rows->item("Label").asText()=="keep");clone->member("Close",{});rows->member("Close",{});
    assert(definition->member("Fields",{VbaValue::string("Label")}).object->member("Size",{}).asNumber()==20);
    assert(sql.exec("SELECT Label,Payload,Magnitude FROM PropertyRows ORDER BY ID") && sql.next());assert(sql.value(0).toString()=="keep" && sql.value(1).toByteArray()==QByteArray::fromHex("0001ff") && sql.value(2).toDouble()==std::numeric_limits<double>::max());assert(sql.next() && sql.value(0).toString()=="long value");sql.finish();
    database->member("Execute",{VbaValue::string("INSERT INTO PropertyRows DEFAULT VALUES")});assert(sql.exec("SELECT ID,Label FROM PropertyRows WHERE ID=3") && sql.next() && sql.value(1).toString()=="new");sql.finish();
    linaccess::updateNativeFieldProperties(db,"PropertyRows","Label",{{"Size",VbaValue::longInteger(10)},{"DefaultValue",VbaValue::string("'future'")},{"Required",VbaValue::logical(true)}});
    error([&] {linaccess::updateNativeFieldProperties(db,"PropertyRows","Label",{{"Size",VbaValue::longInteger(4)},{"DefaultValue",VbaValue::string("'x'")}});},3163);
    assert(definition->member("Fields",{VbaValue::string("Label")}).object->member("Size",{}).asNumber()==10);
    database->member("Execute",{VbaValue::string("INSERT INTO PropertyRows DEFAULT VALUES")});assert(sql.exec("SELECT Label FROM PropertyRows WHERE ID=4") && sql.next() && sql.value(0).toString()=="future");sql.finish();
    assert(sql.exec("SELECT COUNT(*) FROM PropertyRows WHERE Optional IS NULL") && sql.next() && sql.value(0).toInt()==4);sql.finish();
    auto optional=definition->member("Fields",{VbaValue::string("Optional")}).object;assert(!optional->member("Required",{}).truthy() && optional->member("DefaultValue",{}).asText().empty());error([&] {optional->assign("DefaultValue",VbaValue::string("2147483648"));},6);
    definition->member("Fields",{VbaValue::string("Enabled")}).object->assign("DefaultValue",VbaValue::string("1"));definition->member("Fields",{VbaValue::string("Moment")}).object->assign("DefaultValue",VbaValue::string("'2027-01-02'"));
    database->member("Execute",{VbaValue::string("INSERT INTO PropertyRows DEFAULT VALUES")});assert(sql.exec("SELECT ID,Enabled,Moment FROM PropertyRows WHERE ID=5") && sql.next() && sql.value(1).toInt()==-1 && sql.value(2).toString()=="2027-01-02 00:00:00");sql.finish();
    auto workspace=linaccess::makeDaoEngine(db)->defaultMember({VbaValue::integer(0)},true).object;workspace->member("BeginTrans",{});error([&] {definition->member("Fields",{VbaValue::string("Label")}).object->assign("Size",VbaValue::longInteger(11));},3251);workspace->member("Rollback",{});
    auto keyParent=database->member("CreateTableDef",{VbaValue::string("PropertyParent")}).object;keyParent->member("Fields",{}).object->member("Append",{keyParent->member("CreateField",{VbaValue::string("Label"),VbaValue::integer(10),VbaValue::longInteger(10)})});tables->member("Append",{VbaValue::reference(keyParent)});
    auto relation=database->member("CreateRelation",{VbaValue::string("PropertyMapping"),VbaValue::string("PropertyParent"),VbaValue::string("PropertyRows"),VbaValue::longInteger(2)}).object;auto link=relation->member("CreateField",{VbaValue::string("Label")}).object;link->assign("ForeignName",VbaValue::string("Label"));relation->member("Fields",{}).object->member("Append",{VbaValue::reference(link)});database->member("Relations",{}).object->member("Append",{VbaValue::reference(relation)});
    label=definition->member("Fields",{VbaValue::string("Label")}).object;error([&] {label->assign("Size",VbaValue::longInteger(11));},3251);error([&] {label->assign("Required",VbaValue::logical(false));},3251);
    database->member("Relations",{}).object->member("Delete",{VbaValue::string("PropertyMapping")});label->assign("Required",VbaValue::logical(false));assert(!label->member("Required",{}).truthy());
    assert(sql.exec("PRAGMA foreign_keys") && sql.next() && sql.value(0).toInt()==0);sql.finish();
    assert(sql.exec("SELECT value FROM _la_counter32_sequences WHERE table_name='PropertyRows'") && sql.next() && sql.value(0).toInt()==5);sql.finish();
    auto labelIndex=definition->member("CreateIndex",{VbaValue::string("LabelDescending")}).object;auto labelIndexField=labelIndex->member("CreateField",{VbaValue::string("Label")}).object;labelIndexField->assign("Attributes",VbaValue::longInteger(1));labelIndex->member("Fields",{}).object->member("Append",{VbaValue::reference(labelIndexField)});definition->member("Indexes",{}).object->member("Append",{VbaValue::reference(labelIndex)});
    assert(sql.exec("CREATE VIEW property_label_view AS SELECT Label AS observed_label FROM PropertyRows"));
    assert(sql.exec("CREATE TABLE rename_audit(value TEXT)"));assert(sql.exec("CREATE TRIGGER property_label_audit AFTER UPDATE OF Label ON PropertyRows BEGIN INSERT INTO rename_audit VALUES(new.Label); END"));
    auto renameCursor=database->member("OpenRecordset",{VbaValue::string("PropertyRows")}).object;
    auto heldFields=definition->member("Fields",{}).object;label=heldFields->member("Item",{VbaValue::string("Label")}).object;
    error([&] {label->assign("Name",VbaValue::string("Optional"));},3191);error([&] {label->assign("Name",VbaValue::string("rowid"));},3251);
    label->assign("Name",VbaValue::string("Beschriftung ä"));assert(label->member("Name",{}).asText()=="Beschriftung ä");assert(label->member("Size",{}).asNumber()==10);
    error([&] {renameCursor->member("RecordCount",{});},3420);renameCursor->member("Requery",{});assert(renameCursor->item("Beschriftung ä").asText()=="keep");
    error([&] {heldFields->member("Count",{});},3420);heldFields->member("Refresh",{});assert(heldFields->member("Item",{VbaValue::string("Beschriftung ä")}).object->member("Size",{}).asNumber()==10);
    assert(sql.exec("SELECT observed_label FROM property_label_view ORDER BY observed_label") && sql.next());sql.finish();
    assert(sql.exec("UPDATE PropertyRows SET [Beschriftung ä]='updated' WHERE ID=1"));assert(sql.exec("SELECT value FROM rename_audit") && sql.next() && sql.value(0).toString()=="updated");sql.finish();
    auto renamedIndex=definition->member("Indexes",{VbaValue::string("LabelDescending")}).object;assert(renamedIndex->member("Fields",{VbaValue::integer(0)}).object->member("Name",{}).asText()=="Beschriftung ä");
    assert(sql.exec("SELECT ID,Payload,Magnitude FROM PropertyRows WHERE ID=1") && sql.next() && sql.value(1).toByteArray()==QByteArray::fromHex("0001ff") && sql.value(2).toDouble()==std::numeric_limits<double>::max());sql.finish();
    database->member("Execute",{VbaValue::string("INSERT INTO PropertyRows DEFAULT VALUES")});assert(sql.exec("SELECT ID,[Beschriftung ä] FROM PropertyRows WHERE ID=6") && sql.next() && sql.value(1).toString()=="future");sql.finish();
    definition->member("Fields",{VbaValue::string("Beschriftung ä")}).object->assign("Name",VbaValue::string("beschriftung ä"));assert(definition->member("Fields",{VbaValue::string("beschriftung ä")}).object->member("Name",{}).asText()=="beschriftung ä");
    auto nonempty=definition->member("Fields",{VbaValue::string("beschriftung ä")}).object;assert(nonempty->member("AllowZeroLength",{}).truthy());nonempty->assign("AllowZeroLength",VbaValue::logical(false));assert(!nonempty->member("AllowZeroLength",{}).truthy());
    error([&] {database->member("Execute",{VbaValue::string("INSERT INTO PropertyRows([beschriftung ä]) VALUES('')")});},3315);
    database->member("Execute",{VbaValue::string("INSERT INTO PropertyRows([beschriftung ä]) VALUES(' ')")});database->member("Execute",{VbaValue::string("INSERT INTO PropertyRows([beschriftung ä]) VALUES(NULL)")});
    error([&] {database->member("Execute",{VbaValue::string("UPDATE PropertyRows SET [beschriftung ä]=''")});},3315);
    assert(sql.exec("SELECT COUNT(*) FROM PropertyRows WHERE [beschriftung ä]=''") && sql.next() && sql.value(0).toInt()==0);sql.finish();
    error([&] {definition->member("Fields",{VbaValue::string("Optional")}).object->assign("AllowZeroLength",VbaValue::logical(false));},3251);
    linaccess::updateNativeFieldProperties(db,"PropertyRows","beschriftung ä",{{"AllowZeroLength",VbaValue::logical(true)},{"DefaultValue",VbaValue::string("''")}});database->member("Execute",{VbaValue::string("INSERT INTO PropertyRows DEFAULT VALUES")});
    nonempty=definition->member("Fields",{VbaValue::string("beschriftung ä")}).object;error([&] {nonempty->assign("AllowZeroLength",VbaValue::logical(false));},3315);assert(nonempty->member("AllowZeroLength",{}).truthy());
    auto memo=definition->member("CreateField",{VbaValue::string("Memo"),VbaValue::integer(12)}).object;memo->assign("AllowZeroLength",VbaValue::logical(false));memo->assign("DefaultValue",VbaValue::string("'memo'"));definition->member("Fields",{}).object->member("Append",{VbaValue::reference(memo)});
    error([&] {database->member("Execute",{VbaValue::string("INSERT INTO PropertyRows(Memo) VALUES('')")});},3315);memo->assign("AllowZeroLength",VbaValue::logical(true));database->member("Execute",{VbaValue::string("INSERT INTO PropertyRows(Memo) VALUES('')")});
    auto zeroTable=database->member("CreateTableDef",{VbaValue::string("ZeroCreation")}).object;auto zeroField=zeroTable->member("CreateField",{VbaValue::string("TextValue"),VbaValue::integer(10),VbaValue::longInteger(8)}).object;zeroField->assign("Required",VbaValue::logical(true));zeroField->assign("AllowZeroLength",VbaValue::logical(false));zeroTable->member("Fields",{}).object->member("Append",{VbaValue::reference(zeroField)});tables->member("Append",{VbaValue::reference(zeroTable)});
    error([&] {database->member("Execute",{VbaValue::string("INSERT INTO ZeroCreation VALUES('')")});},3315);error([&] {database->member("Execute",{VbaValue::string("INSERT INTO ZeroCreation VALUES(NULL)")});},3314);
    auto zeroCursor=database->member("OpenRecordset",{VbaValue::string("ZeroCreation")}).object;auto zeroMetadata=zeroCursor->member("Fields",{VbaValue::string("TextValue")}).object;assert(zeroMetadata->member("Required",{}).truthy() && !zeroMetadata->member("AllowZeroLength",{}).truthy());error([&] {zeroMetadata->assign("AllowZeroLength",VbaValue::logical(true));},383);auto zeroClone=zeroCursor->member("Clone",{}).object;assert(!zeroClone->member("Fields",{VbaValue::string("TextValue")}).object->member("AllowZeroLength",{}).truthy());zeroCursor->member("AddNew",{});zeroCursor->assign("TextValue",VbaValue::string(""));error([&] {zeroCursor->member("Update",{});},3315);zeroCursor->assign("TextValue",VbaValue::string("kept"));zeroCursor->member("Update",{});assert(sql.exec("SELECT TextValue FROM ZeroCreation") && sql.next() && sql.value(0).toString()=="kept");sql.finish();
    const auto buildReentry=[&] {auto object=database->member("CreateTableDef",{VbaValue::string("FieldReentry")}).object;object->member("Fields",{}).object->member("Append",{object->member("CreateField",{VbaValue::string("Label"),VbaValue::integer(10),VbaValue::longInteger(8)})});tables->member("Append",{VbaValue::reference(object)});};buildReentry();
    auto held=tables->member("Item",{VbaValue::string("FieldReentry")}).object->member("Fields",{VbaValue::string("Label")}).object;
    class ReplacingValue final:public linaccess::VbaObject {public:std::function<void()> callback;bool hasScalarDefault() const override {return true;}VbaValue scalarValue() override {callback();return VbaValue::longInteger(16);}VbaValue member(std::string_view,const std::vector<VbaValue>&) override {throw linaccess::VbaRuntimeError(438,"No member");}};
    auto replacing=std::make_shared<ReplacingValue>();replacing->callback=[&] {tables->member("Delete",{VbaValue::string("FieldReentry")});buildReentry();};error([&] {held->assign("Size",VbaValue::reference(replacing));},3420);
    assert(tables->member("Item",{VbaValue::string("FieldReentry")}).object->member("Fields",{VbaValue::string("Label")}).object->member("Size",{}).asNumber()==8);
}

void testExistingNativePrimaryIndexes() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","dao-existing-native-primary");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);
    auto database=linaccess::makeDaoDatabase(db),tables=database->member("TableDefs",{}).object;
    auto table=database->member("CreateTableDef",{VbaValue::string("PrimaryRows")}).object;
    for(const auto& item:std::vector<std::pair<const char*,int>>{{"ID",4},{"Code",10},{"GroupID",4},{"Payload",11}}) {
        auto field=table->member("CreateField",{VbaValue::string(item.first),VbaValue::integer(static_cast<std::int16_t>(item.second)),VbaValue::longInteger(item.second==10 ? 12 : 0)}).object;
        if(std::string(item.first)=="ID") field->assign("Attributes",VbaValue::longInteger(16));if(item.second==10) field->assign("AllowZeroLength",VbaValue::logical(false));table->member("Fields",{}).object->member("Append",{VbaValue::reference(field)});
    }
    tables->member("Append",{VbaValue::reference(table)});
    auto lookup=table->member("CreateIndex",{VbaValue::string("GroupDescending")}).object;auto lookupField=lookup->member("CreateField",{VbaValue::string("GroupID")}).object;lookupField->assign("Attributes",VbaValue::longInteger(1));lookup->member("Fields",{}).object->member("Append",{VbaValue::reference(lookupField)});table->member("Indexes",{}).object->member("Append",{VbaValue::reference(lookup)});
    assert(sql.exec("CREATE TABLE primary_audit(value INTEGER)"));assert(sql.exec("CREATE TRIGGER primary_rows_audit AFTER INSERT ON PrimaryRows BEGIN INSERT INTO primary_audit VALUES(new.ID); END"));assert(sql.exec("CREATE VIEW primary_rows_view AS SELECT Code AS observed FROM PrimaryRows"));
    sql.prepare("INSERT INTO PrimaryRows(Code,GroupID,Payload) VALUES(?,?,?)");for(int row=0;row<2;++row) {sql.bindValue(0,"a");sql.bindValue(1,7+row);sql.bindValue(2,QByteArray::fromHex(row==0 ? "0001ff" : "dead"));assert(sql.exec());}sql.finish();
    const auto error=[](const std::function<void()>& operation,int expected) {int number=0;try {operation();}catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;if(number!=expected) std::cerr<<"Primary edit error="<<number<<", reason="<<failure.what()<<"\n";}assert(number==expected);};
    const auto build=[&](const char* name) {auto index=table->member("CreateIndex",{VbaValue::string(name)}).object;index->assign("Primary",VbaValue::logical(true));auto field=index->member("CreateField",{VbaValue::string("Code")}).object;field->assign("Attributes",VbaValue::longInteger(1));index->member("Fields",{}).object->member("Append",{VbaValue::reference(field)});return index;};
    auto primary=build("CodePrimary");error([&] {table->member("Indexes",{}).object->member("Append",{VbaValue::reference(primary)});},3022);assert(table->member("Indexes",{}).object->member("Count",{}).asNumber()==1);
    database->member("Execute",{VbaValue::string("UPDATE PrimaryRows SET Code=NULL WHERE ID=2")});error([&] {table->member("Indexes",{}).object->member("Append",{VbaValue::reference(primary)});},3314);database->member("Execute",{VbaValue::string("UPDATE PrimaryRows SET Code='b' WHERE ID=2")});
    auto cursor=database->member("OpenRecordset",{VbaValue::string("PrimaryRows")}).object,clone=cursor->member("Clone",{}).object;
    table->member("Indexes",{}).object->member("Append",{VbaValue::reference(primary)});assert(primary->member("Primary",{}).truthy() && primary->member("Unique",{}).truthy());
    error([&] {cursor->member("RecordCount",{});},3420);error([&] {clone->member("RecordCount",{});},3420);cursor->member("Requery",{});assert(clone->member("RecordCount",{}).asNumber()==2);
    assert(sql.exec("PRAGMA table_info(PrimaryRows)"));bool codeKey=false;while(sql.next()) if(sql.value(1).toString()=="Code") codeKey=sql.value(3).toBool() && sql.value(5).toInt()==1;assert(codeKey);sql.finish();
    assert(primary->member("Fields",{VbaValue::integer(0)}).object->member("Attributes",{}).asNumber()==1);
    error([&] {database->member("Execute",{VbaValue::string("UPDATE PrimaryRows SET Code=NULL WHERE ID=2")});},3314);auto second=build("SecondPrimary");error([&] {table->member("Indexes",{}).object->member("Append",{VbaValue::reference(second)});},3283);
    assert(sql.exec("SELECT COUNT(*) FROM primary_audit") && sql.next() && sql.value(0).toInt()==2);sql.finish();assert(sql.exec("SELECT observed FROM primary_rows_view ORDER BY observed") && sql.next() && sql.value(0).toString()=="a");sql.finish();
    assert(sql.exec("SELECT value FROM _la_counter32_sequences WHERE table_name='PrimaryRows'") && sql.next() && sql.value(0).toInt()==2);sql.finish();
    table->member("Indexes",{}).object->member("Delete",{VbaValue::string("CodePrimary")});assert(table->member("Indexes",{}).object->member("Count",{}).asNumber()==1);error([&] {primary->member("Name",{});},3420);
    database->member("Execute",{VbaValue::string("INSERT INTO PrimaryRows(Code,GroupID) VALUES(NULL,9)")});assert(sql.exec("SELECT ID,Code FROM PrimaryRows WHERE GroupID=9") && sql.next() && sql.value(0).toInt()==3 && sql.value(1).isNull());sql.finish();
    error([&] {database->member("Execute",{VbaValue::string("UPDATE PrimaryRows SET Code='' WHERE ID=1")});},3315);
    assert(sql.exec("SELECT hex(Payload) FROM PrimaryRows WHERE ID=1") && sql.next() && sql.value(0).toString()=="0001FF");sql.finish();assert(sql.exec("SELECT COUNT(*) FROM primary_audit") && sql.next() && sql.value(0).toInt()==3);sql.finish();
    auto mandatory=table->member("CreateIndex",{VbaValue::string("CodeRequired")}).object;mandatory->assign("Required",VbaValue::logical(true));mandatory->member("Fields",{}).object->member("Append",{mandatory->member("CreateField",{VbaValue::string("Code")})});
    error([&] {table->member("Indexes",{}).object->member("Append",{VbaValue::reference(mandatory)});},3314);assert(table->member("Indexes",{}).object->member("Count",{}).asNumber()==1);
    database->member("Execute",{VbaValue::string("UPDATE PrimaryRows SET Code='c' WHERE ID=3")});table->member("Indexes",{}).object->member("Append",{VbaValue::reference(mandatory)});assert(mandatory->member("Required",{}).truthy() && !mandatory->member("Unique",{}).truthy());
    assert(!table->member("Fields",{VbaValue::string("Code")}).object->member("Required",{}).truthy());error([&] {database->member("Execute",{VbaValue::string("UPDATE PrimaryRows SET Code=NULL WHERE ID=3")});},3314);
    database->member("Execute",{VbaValue::string("UPDATE PrimaryRows SET Code='a' WHERE ID=2")});table->member("Indexes",{}).object->member("Delete",{VbaValue::string("CodeRequired")});assert(table->member("Indexes",{}).object->member("Count",{}).asNumber()==1);database->member("Execute",{VbaValue::string("UPDATE PrimaryRows SET Code=NULL WHERE ID=3")});
    error([&] {database->member("Execute",{VbaValue::string("UPDATE PrimaryRows SET Code='' WHERE ID=1")});},3315);assert(sql.exec("SELECT value FROM _la_counter32_sequences WHERE table_name='PrimaryRows'") && sql.next() && sql.value(0).toInt()==3);sql.finish();
    clone.reset();cursor.reset();
    auto workspace=linaccess::makeDaoEngine(db)->defaultMember({VbaValue::integer(0)},true).object;workspace->member("BeginTrans",{});error([&] {table->member("Indexes",{}).object->member("Append",{VbaValue::reference(build("InTransaction"))});},3251);workspace->member("Rollback",{});
    assert(sql.exec("PRAGMA foreign_keys") && sql.next() && sql.value(0).toInt()==0);sql.finish();
}

void testShadowedDaoRowIdentity() {
    const char* phase="schema";try {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","dao-shadowed-row-identity");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);
    auto database=linaccess::makeDaoDatabase(db),tables=database->member("TableDefs",{}).object;
    auto table=database->member("CreateTableDef",{VbaValue::string("Aliases")}).object;
    for(const auto& item:std::vector<std::pair<const char*,int>>{{"ID",4},{"rowid",10},{"_ROWID_",10},{"Value",4}}) {auto field=table->member("CreateField",{VbaValue::string(item.first),VbaValue::integer(static_cast<std::int16_t>(item.second)),VbaValue::longInteger(item.second==10 ? 8 : 0)}).object;if(std::string(item.first)=="ID") field->assign("Attributes",VbaValue::longInteger(16));if(item.second==10) field->assign("DefaultValue",VbaValue::string("'same'"));table->member("Fields",{}).object->member("Append",{VbaValue::reference(field)});}tables->member("Append",{VbaValue::reference(table)});
    phase="insert";
    database->member("Execute",{VbaValue::string("INSERT INTO Aliases(Value) VALUES(10)")});database->member("Execute",{VbaValue::string("INSERT INTO Aliases(Value) VALUES(11)")});
    phase="cursor update";
    auto cursor=database->member("OpenRecordset",{VbaValue::string("Aliases")}).object;cursor->member("MoveLast",{});assert(cursor->item("ID").asNumber()==2 && cursor->item("rowid").asText()=="same");cursor->member("Edit",{});cursor->assign("Value",VbaValue::longInteger(20));cursor->member("Update",{});
    assert(sql.exec("SELECT Value FROM Aliases ORDER BY ID") && sql.next() && sql.value(0).toInt()==10 && sql.next() && sql.value(0).toInt()==20);sql.finish();cursor->member("Delete",{});assert(sql.exec("SELECT COUNT(*) FROM Aliases") && sql.next() && sql.value(0).toInt()==1);sql.finish();
    phase="append second alias";
    const auto append=[&](const char* name) {auto field=table->member("CreateField",{VbaValue::string(name),VbaValue::integer(10),VbaValue::longInteger(8)}).object;field->assign("DefaultValue",VbaValue::string("'also'"));table->member("Fields",{}).object->member("Append",{VbaValue::reference(field)});};append("Extra");cursor->member("Requery",{});auto clone=cursor->member("Clone",{}).object;
    phase="add new";
    cursor->member("AddNew",{});cursor->assign("Value",VbaValue::longInteger(30));cursor->member("Update",{});assert(cursor->item("ID").asNumber()==3 && cursor->item("rowid").asText()=="same" && cursor->item("_rowid_").asText()=="same");
    phase="edit user alias";
    cursor->member("Edit",{});cursor->assign("rowid",VbaValue::string("renamed"));cursor->member("Update",{});assert(cursor->item("ID").asNumber()==3 && cursor->item("rowid").asText()=="renamed");assert(clone->member("RecordCount",{}).asNumber()==2);cursor->member("Delete",{});
    assert(sql.exec("SELECT ID,Value,rowid FROM Aliases") && sql.next() && sql.value(0).toInt()==1 && sql.value(1).toInt()==10 && sql.value(2).toString()=="same" && !sql.next());sql.finish();
    phase="all shadowed";
    append("More");const auto error=[](const std::function<void()>& operation,int expected) {int number=0;try {operation();}catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;}assert(number==expected);};error([&] {cursor->member("Edit",{});},3420);cursor->member("Requery",{});assert(cursor->member("RecordCount",{}).asNumber()==1);
    auto blocked=database->member("CreateTableDef",{VbaValue::string("AllAliases")}).object;for(const auto* name:{"rowid","_rowid_","oid"}) blocked->member("Fields",{}).object->member("Append",{blocked->member("CreateField",{VbaValue::string(name),VbaValue::integer(10),VbaValue::longInteger(8)})});tables->member("Append",{VbaValue::reference(blocked)});error([&] {database->member("OpenRecordset",{VbaValue::string("AllAliases")});},3251);
    error([&] {append("oid");},3251);
    assert(sql.exec("SELECT COUNT(*) FROM Aliases WHERE ID=1 AND Value=10") && sql.next() && sql.value(0).toInt()==1);sql.finish();
    assert(sql.exec("CREATE TABLE SingleAlias(rowid TEXT,Value INTEGER)"));assert(sql.exec("INSERT INTO SingleAlias VALUES('same',1),('same',2)"));auto single=database->member("OpenRecordset",{VbaValue::string("SingleAlias")}).object;single->member("MoveLast",{});single->member("Edit",{});single->assign("Value",VbaValue::longInteger(7));single->member("Update",{});single->member("Delete",{});assert(sql.exec("SELECT Value FROM SingleAlias") && sql.next() && sql.value(0).toInt()==1 && !sql.next());sql.finish();
    assert(sql.exec("CREATE TABLE NoRowIdentity(ID INTEGER PRIMARY KEY) WITHOUT ROWID"));error([&] {database->member("OpenRecordset",{VbaValue::string("NoRowIdentity")});},3251);
    }catch(const std::exception& failure) {std::cerr<<"Row identity phase: "<<phase<<"; "<<failure.what()<<"\n";throw;}
}

void testExistingNativeTableRename() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","dao-native-table-rename");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);assert(sql.exec("PRAGMA foreign_keys=ON"));
    auto database=linaccess::makeDaoDatabase(db),tables=database->member("TableDefs",{}).object;
    const auto build=[&](const char* name,bool parent) {auto table=database->member("CreateTableDef",{VbaValue::string(name)}).object;auto id=table->member("CreateField",{VbaValue::string("ID"),VbaValue::integer(4)}).object;id->assign("Attributes",VbaValue::longInteger(16));table->member("Fields",{}).object->member("Append",{VbaValue::reference(id)});
        if(parent) {for(const auto& item:std::vector<std::pair<const char*,int>>{{"Label",10},{"Payload",11},{"Magnitude",7}}) {auto field=table->member("CreateField",{VbaValue::string(item.first),VbaValue::integer(static_cast<std::int16_t>(item.second)),VbaValue::longInteger(item.second==10 ? 12 : 0)}).object;if(item.second==10) {field->assign("DefaultValue",VbaValue::string("'Prüfwert'"));field->assign("AllowZeroLength",VbaValue::logical(false));}table->member("Fields",{}).object->member("Append",{VbaValue::reference(field)});}}
        else table->member("Fields",{}).object->member("Append",{table->member("CreateField",{VbaValue::string("ParentID"),VbaValue::integer(4)})});
        auto key=table->member("CreateIndex",{VbaValue::string("PrimaryKey")}).object;key->assign("Primary",VbaValue::logical(true));key->member("Fields",{}).object->member("Append",{key->member("CreateField",{VbaValue::string("ID")})});table->member("Indexes",{}).object->member("Append",{VbaValue::reference(key)});tables->member("Append",{VbaValue::reference(table)});return table;};
    auto parent=build("RenameParent",true);auto child=build("RenameChild",false);
    auto lookup=parent->member("CreateIndex",{VbaValue::string("LabelDescending")}).object;auto ordered=lookup->member("CreateField",{VbaValue::string("Label")}).object;ordered->assign("Attributes",VbaValue::longInteger(1));lookup->member("Fields",{}).object->member("Append",{VbaValue::reference(ordered)});parent->member("Indexes",{}).object->member("Append",{VbaValue::reference(lookup)});
    database->member("Execute",{VbaValue::string("ALTER TABLE RenameParent ALTER COLUMN ID COUNTER(100,7)")});
    assert(sql.exec("CREATE VIEW rename_parent_view AS SELECT ID,Label,Payload,Magnitude FROM RenameParent"));assert(sql.exec("CREATE VIEW rename_literal_view AS SELECT 'INDEXED BY' AS observed FROM RenameParent"));assert(sql.exec("CREATE TABLE rename_audit(value INTEGER)"));assert(sql.exec("CREATE TRIGGER rename_parent_audit AFTER INSERT ON RenameParent BEGIN INSERT INTO rename_audit VALUES(new.ID);END"));
    sql.prepare("INSERT INTO RenameParent(Payload,Magnitude) VALUES(?,?)");for(int row=0;row<2;++row) {sql.bindValue(0,QByteArray::fromHex("0001ff"));sql.bindValue(1,std::numeric_limits<double>::max());assert(sql.exec());}sql.finish();database->member("Execute",{VbaValue::string("DELETE FROM RenameParent WHERE ID=107")});database->member("Execute",{VbaValue::string("INSERT INTO RenameChild(ParentID) VALUES(100)")});
    auto relation=database->member("CreateRelation",{VbaValue::string("RenameMapping"),VbaValue::string("RenameParent"),VbaValue::string("RenameChild"),VbaValue::longInteger(4096)}).object;auto pair=relation->member("CreateField",{VbaValue::string("ID")}).object;pair->assign("ForeignName",VbaValue::string("ParentID"));relation->member("Fields",{}).object->member("Append",{VbaValue::reference(pair)});database->member("Relations",{}).object->member("Append",{VbaValue::reference(relation)});
    const auto error=[](const std::function<void()>& operation,int expected) {int number=0;try {operation();}catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;if(number!=expected) std::cerr<<"Table rename error="<<number<<", reason="<<failure.what()<<"\n";}assert(number==expected);};
    auto oldCursor=database->member("OpenRecordset",{VbaValue::string("RenameParent")}).object;
    assert(sql.exec("CREATE TABLE RenameOccupied(value INTEGER)"));assert(sql.exec("CREATE INDEX [Renamed parent_LabelDescending] ON RenameOccupied(value)"));error([&] {parent->assign("Name",VbaValue::string("Renamed parent"));},3010);assert(parent->member("Name",{}).asText()=="RenameParent" && oldCursor->member("RecordCount",{}).asNumber()==1);assert(sql.exec("DROP INDEX [Renamed parent_LabelDescending]"));
    assert(sql.exec("CREATE VIEW rename_index_hint AS SELECT Label FROM RenameParent INDEXED /* technical */ BY RenameParent_LabelDescending"));error([&] {parent->assign("Name",VbaValue::string("Renamed parent"));},3251);assert(sql.exec("DROP VIEW rename_index_hint"));
    auto oldFields=parent->member("Fields",{}).object;parent->assign("Name",VbaValue::string("Renamed parent"));assert(parent->member("Name",{}).asText()=="Renamed parent");error([&] {oldCursor->member("Requery",{});},3420);error([&] {oldFields->member("Count",{});},3420);error([&] {tables->member("Item",{VbaValue::string("RenameParent")});},3265);
    auto renamed=tables->member("Item",{VbaValue::string("Renamed parent")}).object;assert(renamed->member("Indexes",{}).object->member("Count",{}).asNumber()==2);assert(renamed->member("Indexes",{VbaValue::string("LabelDescending")}).object->member("Fields",{VbaValue::integer(0)}).object->member("Attributes",{}).asNumber()==1);
    assert(sql.exec("SELECT ID,Label,Payload,Magnitude FROM rename_parent_view") && sql.next() && sql.value(0).toInt()==100 && sql.value(1).toString()=="Prüfwert" && sql.value(2).toByteArray()==QByteArray::fromHex("0001ff") && sql.value(3).toDouble()==std::numeric_limits<double>::max());sql.finish();
    auto renamedRelation=database->member("Relations",{VbaValue::string("RenameMapping")}).object;assert(renamedRelation->member("Table",{}).asText()=="Renamed parent");assert(sql.exec("PRAGMA foreign_key_list(RenameChild)") && sql.next() && sql.value(2).toString()=="Renamed parent");sql.finish();
    database->member("Execute",{VbaValue::string("INSERT INTO [Renamed parent] DEFAULT VALUES")});assert(sql.exec("SELECT MAX(ID) FROM [Renamed parent]") && sql.next() && sql.value(0).toInt()==114);sql.finish();assert(sql.exec("SELECT COUNT(*) FROM rename_audit") && sql.next() && sql.value(0).toInt()==3);sql.finish();
    assert(sql.exec("SELECT COUNT(*) FROM _la_counter32_sequences WHERE table_name='RenameParent'") && sql.next() && sql.value(0).toInt()==0);sql.finish();database->member("Execute",{VbaValue::string("DELETE FROM [Renamed parent] WHERE ID=100")});assert(sql.exec("SELECT COUNT(*) FROM RenameChild") && sql.next() && sql.value(0).toInt()==0);sql.finish();assert(sql.exec("PRAGMA foreign_keys") && sql.next() && sql.value(0).toInt()==1);sql.finish();
}

void testNativeDaoWriteCapability() {
    using linaccess::VbaValue;QTemporaryDir directory;assert(directory.isValid());const auto path=directory.filePath("readonly.linaccess");
    auto db=QSqlDatabase::addDatabase("QSQLITE","dao-native-write-capability");db.setDatabaseName(path);assert(db.open());QSqlQuery sql(db);auto database=linaccess::makeDaoDatabase(db),tables=database->member("TableDefs",{}).object;
    const auto error=[](const std::function<void()>& operation,int expected) {int number=0;try {operation();}catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;if(number!=expected) std::cerr<<"Writability error="<<number<<", reason="<<failure.what()<<"\n";}assert(number==expected);};
    auto table=database->member("CreateTableDef",{VbaValue::string("LocalRows")}).object;assert(table->member("Attributes",{}).asNumber()==0 && table->member("Connect",{}).asText().empty() && table->member("SourceTableName",{}).asText().empty());table->assign("Attributes",VbaValue::longInteger(0));table->assign("Connect",VbaValue::string(""));error([&] {table->assign("Connect",VbaValue::string("unsupported-provider"));},3251);
    auto id=table->member("CreateField",{VbaValue::string("ID"),VbaValue::integer(4)}).object;id->assign("Attributes",VbaValue::longInteger(16));table->member("Fields",{}).object->member("Append",{VbaValue::reference(id)});
    auto label=table->member("CreateField",{VbaValue::string("Label"),VbaValue::integer(10),VbaValue::longInteger(12)}).object;label->assign("DefaultValue",VbaValue::string("'kept'"));table->member("Fields",{}).object->member("Append",{VbaValue::reference(label)});tables->member("Append",{VbaValue::reference(table)});database->member("Execute",{VbaValue::string("INSERT INTO LocalRows DEFAULT VALUES")});
    assert(database->member("Updatable",{}).truthy() && table->member("Updatable",{}).truthy() && table->member("Attributes",{}).asNumber()==0);table->assign("SourceTableName",VbaValue::string(""));error([&] {table->assign("Attributes",VbaValue::longInteger(1));},3251);error([&] {database->assign("Updatable",VbaValue::logical(false));},383);
    assert(sql.exec("CREATE TABLE UnknownLocal(value INTEGER)"));error([&] {tables->member("Item",{VbaValue::string("UnknownLocal")}).object->member("Connect",{});},3251);
    auto cursor=database->member("OpenRecordset",{VbaValue::string("LocalRows")}).object;cursor->member("Edit",{});cursor->assign("Label",VbaValue::string("pending"));assert(sql.exec("PRAGMA query_only=ON"));assert(!database->member("Updatable",{}).truthy() && !table->member("Updatable",{}).truthy() && !cursor->member("Updatable",{}).truthy());
    error([&] {cursor->member("Update",{});},3027);cursor->member("CancelUpdate",{});assert(cursor->item("Label").asText()=="kept");error([&] {database->member("Execute",{VbaValue::string("DELETE FROM LocalRows")});},3027);error([&] {table->assign("Name",VbaValue::string("Changed"));},3027);assert(!db.tables().contains("Changed"));
    assert(cursor->member("Fields",{VbaValue::string("ID")}).object->member("Attributes",{}).asNumber()==17 && !cursor->member("Fields",{VbaValue::string("ID")}).object->member("DataUpdatable",{}).truthy());
    auto readonlyAtOpen=database->member("OpenRecordset",{VbaValue::string("LocalRows")}).object;assert(!readonlyAtOpen->member("Updatable",{}).truthy());assert(sql.exec("PRAGMA query_only=OFF"));assert(database->member("Updatable",{}).truthy() && cursor->member("Updatable",{}).truthy() && !readonlyAtOpen->member("Updatable",{}).truthy());cursor->member("Close",{});readonlyAtOpen->member("Close",{});
    sql.clear();db.close();QFile original(path);assert(original.open(QIODevice::ReadOnly));const auto before=original.readAll();original.close();db.setConnectOptions("QSQLITE_OPEN_READONLY");assert(db.open());database=linaccess::makeDaoDatabase(db);tables=database->member("TableDefs",{}).object;table=tables->member("Item",{VbaValue::string("LocalRows")}).object;cursor=database->member("OpenRecordset",{VbaValue::string("LocalRows")}).object;
    assert(!database->member("Updatable",{}).truthy() && !table->member("Updatable",{}).truthy() && !cursor->member("Updatable",{}).truthy() && cursor->item("ID").asNumber()==1);assert(table->member("Connect",{}).asText().empty());
    error([&] {cursor->member("AddNew",{});},3027);error([&] {cursor->member("Edit",{});},3027);error([&] {cursor->assign("Label",VbaValue::string("lost"));},3027);error([&] {database->member("Execute",{VbaValue::string("INSERT INTO LocalRows DEFAULT VALUES")});},3027);error([&] {table->member("Fields",{VbaValue::string("Label")}).object->assign("Required",VbaValue::logical(true));},3027);error([&] {database->member("CreateQueryDef",{VbaValue::string("ReadonlyQuery"),VbaValue::string("SELECT * FROM LocalRows")});},3027);
    sql=QSqlQuery(db);assert(sql.exec("SELECT COUNT(*),MAX(ID),MIN(Label) FROM LocalRows") && sql.next() && sql.value(0).toInt()==1 && sql.value(1).toInt()==1 && sql.value(2).toString()=="kept");sql.finish();assert(sql.exec("SELECT value FROM _la_counter32_sequences WHERE table_name='LocalRows'") && sql.next() && sql.value(0).toInt()==1);sql.clear();db.close();assert(original.open(QIODevice::ReadOnly) && original.readAll()==before);original.close();
}

void testNativeSingleSchemaAndWrites() {
    assert(linaccess::qtFieldValue(QVariant(true),"SINGLE_REAL").asNumber()==-1);
    using linaccess::VbaValue;auto db=QSqlDatabase::addDatabase("QSQLITE","dao-native-single-writes");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);auto database=linaccess::makeDaoDatabase(db);
    auto table=database->member("CreateTableDef",{VbaValue::string("SingleRows")}).object;auto id=table->member("CreateField",{VbaValue::string("ID"),VbaValue::integer(4)}).object;id->assign("Attributes",VbaValue::longInteger(16));table->member("Fields",{}).object->member("Append",{VbaValue::reference(id)});auto sample=table->member("CreateField",{VbaValue::string("Sample"),VbaValue::integer(6)}).object;sample->assign("DefaultValue",VbaValue::string("0.1"));table->member("Fields",{}).object->member("Append",{VbaValue::reference(sample)});database->member("TableDefs",{}).object->member("Append",{VbaValue::reference(table)});
    assert(sample->member("Type",{}).asNumber()==6 && sample->member("Size",{}).asNumber()==4);database->member("Execute",{VbaValue::string("INSERT INTO SingleRows DEFAULT VALUES")});assert(sql.exec("SELECT Sample FROM SingleRows") && sql.next() && sql.value(0).toDouble()==static_cast<double>(0.1f));sql.finish();
    const auto error=[](const std::function<void()>& operation,int expected) {int number=0;try {operation();}catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;if(number!=expected) std::cerr<<"Single write error="<<number<<", reason="<<failure.what()<<"\n";}assert(number==expected);};
    auto rows=database->member("OpenRecordset",{VbaValue::string("SingleRows")}).object;rows->member("Edit",{});rows->assign("Sample",VbaValue::numeric(16777217));assert(rows->item("Sample").asNumber()==16777216 && rows->item("Sample").numericType==VbaValue::NumericType::Single);error([&] {rows->assign("Sample",VbaValue::numeric(1e40));},6);error([&] {rows->assign("Sample",VbaValue::numeric(1e-50));},3251);error([&] {rows->assign("Sample",VbaValue::numeric(-0.0));},3251);error([&] {rows->assign("Sample",VbaValue::string("invalid"));},13);rows->member("Update",{});
    database->member("Execute",{VbaValue::string("INSERT INTO SingleRows(Sample) VALUES(0.2)")});assert(sql.exec("SELECT Sample FROM SingleRows WHERE ID=2") && sql.next() && sql.value(0).toDouble()==static_cast<double>(0.2f));sql.finish();database->member("Execute",{VbaValue::string("UPDATE SingleRows SET Sample=16777217")});error([&] {database->member("Execute",{VbaValue::string("UPDATE SingleRows SET Sample=1e40")});},6);assert(sql.exec("SELECT COUNT(*) FROM SingleRows WHERE Sample=16777216") && sql.next() && sql.value(0).toInt()==2);sql.finish();
    assert(!sql.exec("INSERT INTO SingleRows(Sample) VALUES(0.1)"));assert(sql.exec("CREATE TABLE FloatInput(value REAL)"));assert(sql.exec("INSERT INTO FloatInput VALUES(0.3),(0.4)"));database->member("Execute",{VbaValue::string("INSERT INTO SingleRows(Sample) SELECT value FROM FloatInput")});assert(sql.exec("SELECT Sample FROM SingleRows WHERE Sample<1 ORDER BY Sample") && sql.next() && sql.value(0).toDouble()==static_cast<double>(0.3f) && sql.next() && sql.value(0).toDouble()==static_cast<double>(0.4f));sql.finish();
    error([&] {database->member("Execute",{VbaValue::string("INSERT INTO SingleRows(Sample) VALUES(-0.0)")});},3251);database->member("Execute",{VbaValue::string("INSERT INTO SingleRows(Sample) VALUES(NULL)")});
    sample=table->member("Fields",{VbaValue::string("Sample")}).object;sample->assign("DefaultValue",VbaValue::string("0.25"));database->member("Execute",{VbaValue::string("INSERT INTO SingleRows DEFAULT VALUES")});assert(sql.exec("SELECT COUNT(*) FROM SingleRows WHERE Sample=0.25") && sql.next() && sql.value(0).toInt()==1);sql.finish();
    rows->member("Requery",{});rows->member("AddNew",{});assert(rows->item("Sample").numericType==VbaValue::NumericType::Single && rows->item("Sample").asNumber()==0.25);rows->member("CancelUpdate",{});rows->member("Close",{});
    auto projected=database->member("OpenRecordset",{VbaValue::string("SELECT Sample FROM SingleRows")}).object;assert(projected->member("Fields",{VbaValue::integer(0)}).object->member("Type",{}).asNumber()==6 && projected->item("Sample").numericType==VbaValue::NumericType::Single);
    assert(sql.exec("CREATE TABLE LegacySingles(ID INTEGER,Sample REAL DEFAULT 0.1)"));const QJsonArray metadata{QJsonObject{{"name","ID"},{"source_type","INT"}},QJsonObject{{"name","Sample"},{"source_type","SINGLE"}}};sql.prepare("INSERT INTO _la_table_schema VALUES('LegacySingles',?)");sql.addBindValue(QJsonDocument(metadata).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();database->member("Execute",{VbaValue::string("INSERT INTO LegacySingles(ID) VALUES(1)")});database->member("Execute",{VbaValue::string("INSERT INTO LegacySingles DEFAULT VALUES")});assert(sql.exec("SELECT Sample FROM LegacySingles") && sql.next() && sql.value(0).toDouble()==static_cast<double>(0.1f) && sql.next() && sql.value(0).toDouble()==static_cast<double>(0.1f));sql.finish();
}

void testNativeSingleRelations() {
    using linaccess::VbaValue;auto db=QSqlDatabase::addDatabase("QSQLITE","dao-native-single-relations");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);assert(sql.exec("PRAGMA foreign_keys=ON"));auto database=linaccess::makeDaoDatabase(db);
    const auto build=[&](const char* name,bool primary) {auto table=database->member("CreateTableDef",{VbaValue::string(name)}).object;table->member("Fields",{}).object->member("Append",{table->member("CreateField",{VbaValue::string("KeyValue"),VbaValue::integer(6)})});table->member("Fields",{}).object->member("Append",{table->member("CreateField",{VbaValue::string("Payload"),VbaValue::integer(11)})});if(primary) {auto key=table->member("CreateIndex",{VbaValue::string("PrimaryKey")}).object;key->assign("Primary",VbaValue::logical(true));key->member("Fields",{}).object->member("Append",{key->member("CreateField",{VbaValue::string("KeyValue")})});table->member("Indexes",{}).object->member("Append",{VbaValue::reference(key)});}database->member("TableDefs",{}).object->member("Append",{VbaValue::reference(table)});};build("SingleParent",true);build("SingleChild",false);
    database->member("Execute",{VbaValue::string("INSERT INTO SingleParent(KeyValue) VALUES(0.1)")});sql.prepare("INSERT INTO SingleChild(KeyValue,Payload) VALUES(?,?)");sql.addBindValue(static_cast<double>(0.1f));sql.addBindValue(QByteArray::fromHex("0001ff"));assert(sql.exec());sql.finish();database->member("Execute",{VbaValue::string("INSERT INTO SingleChild(KeyValue) VALUES(NULL)")});
    const auto relation=[&](const char* name,const char* parent) {auto result=database->member("CreateRelation",{VbaValue::string(name),VbaValue::string(parent),VbaValue::string("SingleChild"),VbaValue::longInteger(256|4096)}).object;auto pair=result->member("CreateField",{VbaValue::string("KeyValue")}).object;pair->assign("ForeignName",VbaValue::string("KeyValue"));result->member("Fields",{}).object->member("Append",{VbaValue::reference(pair)});return result;};
    auto mapping=relation("SingleMapping","SingleParent");database->member("Relations",{}).object->member("Append",{VbaValue::reference(mapping)});database->member("TableDefs",{VbaValue::string("SingleChild")}).object->member("Fields",{VbaValue::string("Payload")}).object->assign("OrdinalPosition",VbaValue::integer(0));assert(sql.exec("PRAGMA foreign_key_list(SingleChild)") && sql.next() && sql.value(2).toString()=="SingleParent");sql.finish();
    const auto error=[](const std::function<void()>& operation,int expected) {int number=0;try {operation();}catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;}assert(number==expected);};error([&] {database->member("Execute",{VbaValue::string("INSERT INTO SingleChild(KeyValue) VALUES(0.3)")});},3146);error([&] {database->member("Execute",{VbaValue::string("INSERT INTO SingleParent(KeyValue) VALUES(0.1000000001)")});},3022);
    database->member("Execute",{VbaValue::string("UPDATE SingleParent SET KeyValue=0.2")});assert(sql.exec("SELECT KeyValue,Payload FROM SingleChild WHERE KeyValue IS NOT NULL") && sql.next() && sql.value(0).toDouble()==static_cast<double>(0.2f) && sql.value(1).toByteArray()==QByteArray::fromHex("0001ff"));sql.finish();database->member("Execute",{VbaValue::string("DELETE FROM SingleParent")});assert(sql.exec("SELECT COUNT(*) FROM SingleChild WHERE KeyValue IS NULL") && sql.next() && sql.value(0).toInt()==1);sql.finish();database->member("Relations",{}).object->member("Delete",{VbaValue::string("SingleMapping")});
    assert(sql.exec("CREATE TABLE LegacySingleParent(KeyValue REAL NOT NULL UNIQUE)"));const QJsonArray metadata{QJsonObject{{"name","KeyValue"},{"source_type","SINGLE"}}};sql.prepare("INSERT INTO _la_table_schema VALUES('LegacySingleParent',?)");sql.addBindValue(QJsonDocument(metadata).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();auto unsupported=relation("LegacySingleMapping","LegacySingleParent");error([&] {database->member("Relations",{}).object->member("Append",{VbaValue::reference(unsupported)});},3251);assert(database->member("Relations",{}).object->member("Count",{}).asNumber()==0);
}

void testNativeIntrinsicFieldProperties() {
    using linaccess::VbaValue;auto db=QSqlDatabase::addDatabase("QSQLITE","dao-native-intrinsic-properties");db.setDatabaseName(":memory:");assert(db.open());auto database=linaccess::makeDaoDatabase(db);auto table=database->member("CreateTableDef",{VbaValue::string("IntrinsicRows")}).object;auto id=table->member("CreateField",{VbaValue::string("ID"),VbaValue::integer(4)}).object;id->assign("Attributes",VbaValue::longInteger(16));table->member("Fields",{}).object->member("Append",{VbaValue::reference(id)});auto field=table->member("CreateField",{VbaValue::string("Label"),VbaValue::integer(10),VbaValue::longInteger(12)}).object;auto properties=field->member("Properties",{}).object;properties->member("Item",{VbaValue::string("DefaultValue")}).object->assign("Value",VbaValue::string("\"x\""));assert(properties->member("Item",{VbaValue::string("Required")}).object->member("Type",{}).asNumber()==1);table->member("Fields",{}).object->member("Append",{VbaValue::reference(field)});database->member("TableDefs",{}).object->member("Append",{VbaValue::reference(table)});database->member("Execute",{VbaValue::string("INSERT INTO IntrinsicRows DEFAULT VALUES")});
    field=table->member("Fields",{VbaValue::string("Label")}).object;properties=field->member("Properties",{}).object;auto size=properties->member("Item",{VbaValue::string("Size")}).object;assert(size->member("Type",{}).asNumber()==4 && size->member("Value",{}).numericType==VbaValue::NumericType::Long && size->member("Value",{}).asNumber()==12);size->assign("Value",VbaValue::longInteger(20));assert(size->member("Value",{}).asNumber()==20);auto mandatory=properties->member("Item",{VbaValue::string("Required")}).object;mandatory->assign("Value",VbaValue::logical(true));assert(mandatory->member("Value",{}).truthy());auto name=properties->member("Item",{VbaValue::string("Name")}).object;name->assign("Value",VbaValue::string("RenamedLabel"));assert(name->member("Value",{}).asText()=="RenamedLabel" && size->member("Value",{}).asNumber()==20);
    const auto error=[](const std::function<void()>& operation,int expected) {int number=0;try {operation();}catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;}assert(number==expected);};error([&] {properties->member("Delete",{VbaValue::string("Name")});},383);error([&] {properties->member("Item",{VbaValue::string("Type")}).object->assign("Value",VbaValue::integer(12));},383);auto rows=database->member("OpenRecordset",{VbaValue::string("IntrinsicRows")}).object;auto recordProperties=rows->member("Fields",{VbaValue::string("ID")}).object->member("Properties",{}).object;assert(recordProperties->member("Item",{VbaValue::string("Attributes")}).object->member("Value",{}).asNumber()==49);auto tableProperties=table->member("Fields",{VbaValue::string("ID")}).object->member("Properties",{}).object;assert(tableProperties->member("Item",{VbaValue::string("Attributes")}).object->member("Value",{}).asNumber()==17);auto required=rows->member("Fields",{VbaValue::string("RenamedLabel")}).object->member("Properties",{VbaValue::string("Required")}).object;assert(required->member("Value",{}).truthy());error([&] {required->assign("Value",VbaValue::logical(false));},383);rows->member("Close",{});
    Host host(db);const auto result=linaccess::executeVbaProcedure("Sub Run()\nDim f As DAO.Field\nDim p As DAO.Property\nSet f=CurrentDb.TableDefs(\"IntrinsicRows\").Fields(\"RenamedLabel\")\nSet p=f.Properties(\"DefaultValue\")\nIf p.Type<>dbText Then Err.Raise 5\np.Value=\"\"\"next\"\"\"\nIf CStr(p)<>\"\"\"next\"\"\" Then Err.Raise 5\nf.Properties(\"AllowZeroLength\").Value=False\nIf f.AllowZeroLength Then Err.Raise 5\nEnd Sub","Run",host);if(result.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Intrinsic property: "<<result.explanation<<"\n";assert(result.status==linaccess::VbaExecutionStatus::Completed);
}

void testNativeFieldDisplayProperties() {
    using linaccess::VbaValue;auto db=QSqlDatabase::addDatabase("QSQLITE","dao-native-field-display");db.setDatabaseName(":memory:");assert(db.open());auto database=linaccess::makeDaoDatabase(db);auto table=database->member("CreateTableDef",{VbaValue::string("DisplayRows")}).object;auto field=table->member("CreateField",{VbaValue::string("StoredName"),VbaValue::integer(10),VbaValue::longInteger(20)}).object;auto properties=field->member("Properties",{}).object;
    auto format=database->member("CreateProperty",{VbaValue::string("Format"),VbaValue::integer(10),VbaValue::string("0.00")}).object;
    properties->member("Append",{VbaValue::reference(format)});
    assert(format->member("Name",{}).asText()=="Format" && format->member("Type",{}).asNumber()==10);
    auto caption=database->member("CreateProperty",{VbaValue::string("Caption"),VbaValue::integer(10),VbaValue::string("Anzeige <b>Ä</b>")}).object;properties->member("Append",{VbaValue::reference(caption)});assert(caption->member("Type",{}).asNumber()==10 && caption->member("Value",{}).asText()=="Anzeige <b>Ä</b>");table->member("Fields",{}).object->member("Append",{VbaValue::reference(field)});database->member("TableDefs",{}).object->member("Append",{VbaValue::reference(table)});caption->assign("Value",VbaValue::string("Gespeicherter Kopf"));assert(field->member("Properties",{VbaValue::string("caption")}).object->member("Value",{}).asText()=="Gespeicherter Kopf");
    const auto error=[](const std::function<void()>& operation,int expected) {int number=0;try {operation();}catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;}assert(number==expected);};database->member("Execute",{VbaValue::string("INSERT INTO DisplayRows VALUES('x')")});auto rows=database->member("OpenRecordset",{VbaValue::string("DisplayRows")}).object;auto readonly=rows->member("Fields",{VbaValue::string("StoredName")}).object->member("Properties",{VbaValue::string("Caption")}).object;assert(readonly->member("Value",{}).asText()=="Gespeicherter Kopf");error([&] {readonly->assign("Value",VbaValue::string("blocked"));},383);field->assignIndexedMember("Properties",{VbaValue::string("Caption")},VbaValue::string("Neuer Kopf"));assert(readonly->member("Value",{}).asText()=="Neuer Kopf" && rows->item("StoredName").asText()=="x");
    properties=field->member("Properties",{}).object;error([&] {(void)properties->member("Count",{});},3251);error([&] {(void)properties->member("Item",{VbaValue::string("Description")});},3270);auto description=database->member("CreateProperty",{VbaValue::string("Description"),VbaValue::integer(10),VbaValue::string("Technischer Hinweis")}).object;properties->member("Append",{VbaValue::reference(description)});auto old=properties->member("Item",{VbaValue::string("Caption")}).object;properties->member("Delete",{VbaValue::string("Caption")});error([&] {(void)old->member("Value",{});},3420);auto replacement=database->member("CreateProperty",{VbaValue::string("Caption"),VbaValue::integer(10),VbaValue::string("Ersetzter Kopf")}).object;properties->member("Append",{VbaValue::reference(replacement)});error([&] {(void)old->member("Value",{});},3420);error([&] {replacement->assign("Value",VbaValue::string(std::string(256,'x')));},3251);assert(linaccess::nativeTableFieldDisplayProperties(db,"DisplayRows").value("storedname").description==QString("Technischer Hinweis"));rows->member("Close",{});error([&] {(void)readonly->member("Value",{});},3420);
    assert(format->member("Value",{}).asText()=="0.00");
    format->assign("Value",VbaValue::string("0.000"));
    assert(linaccess::nativeTableFieldDisplayProperties(db,"DisplayRows").value("storedname").format==QString("0.000"));
    error([&] {format->assign("Value",{});},94);
    auto replacementFormat=database->member("CreateProperty",{VbaValue::string("Format"),VbaValue::integer(10),VbaValue::string("0.0")}).object;
    error([&] {properties->member("Append",{VbaValue::reference(replacementFormat)});},3367);
    properties->member("Delete",{VbaValue::string("Format")});
    error([&] {(void)format->member("Value",{});},3420);
    properties->member("Append",{VbaValue::reference(replacementFormat)});
    error([&] {(void)format->member("Value",{});},3420);
    Host host(db);const auto result=linaccess::executeVbaProcedure("Sub Run()\nDim f As DAO.Field\nDim p As DAO.Property\nDim props As DAO.Properties\nSet f=CurrentDb.TableDefs(\"DisplayRows\").Fields(\"StoredName\")\nSet props=f.Properties\nSet p=props.Item(Index:=\"Caption\")\nIf Not TypeOf p Is DAO.Property Then Err.Raise 5\nf.Properties(\"Caption\")=\"VBA Kopf\"\nIf CStr(p)<>\"VBA Kopf\" Then Err.Raise 5\nprops.Delete Name:=\"Description\"\nSet p=CurrentDb.CreateProperty(Value:=\"VBA Hinweis\",Type:=dbText,Name:=\"Description\")\nprops.Append Object:=p\nEnd Sub","Run",host);if(result.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Native display property: "<<result.explanation<<"\n";assert(result.status==linaccess::VbaExecutionStatus::Completed);
    field=table->member("Fields",{VbaValue::string("StoredName")}).object;field->assign("Size",VbaValue::longInteger(30));field->assign("Name",VbaValue::string("RenamedField"));table->assign("Name",VbaValue::string("RenamedDisplayRows"));const auto display=linaccess::nativeTableFieldDisplayProperties(db,"RenamedDisplayRows");assert(display.value("renamedfield").caption==QString("VBA Kopf") && display.value("renamedfield").description==QString("VBA Hinweis") && display.value("renamedfield").format==QString("0.0"));
}

void testNativeDateIndexSeek() {
    using linaccess::VbaValue;auto db=QSqlDatabase::addDatabase("QSQLITE","dao-native-date-index-seek");db.setDatabaseName(":memory:");assert(db.open());auto database=linaccess::makeDaoDatabase(db);auto table=database->member("CreateTableDef",{VbaValue::string("SeekDates")}).object;for(const auto* name:{"Stamp","OrderValue"}) {auto field=table->member("CreateField",{VbaValue::string(name),VbaValue::integer(QString(name)=="Stamp" ? 8 : 4)}).object;field->assign("Required",VbaValue::logical(true));table->member("Fields",{}).object->member("Append",{VbaValue::reference(field)});}auto index=table->member("CreateIndex",{VbaValue::string("ByDate")}).object;for(const auto* name:{"Stamp","OrderValue"}) index->member("Fields",{}).object->member("Append",{index->member("CreateField",{VbaValue::string(name)})});table->member("Indexes",{}).object->member("Append",{VbaValue::reference(index)});database->member("TableDefs",{}).object->member("Append",{VbaValue::reference(table)});
    for(const auto* values:{"'2026-11-09 12:00:00',2","'2026-11-08',1","'2026-11-09 12:00:00',1"}) database->member("Execute",{VbaValue::string((QString("INSERT INTO SeekDates VALUES(")+values+")").toStdString())});auto rows=database->member("OpenRecordset",{VbaValue::string("SeekDates")}).object;rows->assign("Index",VbaValue::string("ByDate"));const auto key=linaccess::qtFieldValue("2026-11-09 12:00:00","DATETIME");rows->member("Seek",{VbaValue::string("="),key,VbaValue::integer(2)});assert(rows->item("OrderValue").asNumber()==2 && rows->item("Stamp").kind==VbaValue::Kind::Date);rows->member("Seek",{VbaValue::string("<"),key,VbaValue::integer(2)});assert(rows->item("OrderValue").asNumber()==1 && linaccess::sqlFieldValue(rows->item("Stamp"),"DATETIME").toString()=="2026-11-09 12:00:00");rows->member("Seek",{VbaValue::string(">="),linaccess::qtFieldValue("2026-11-08 23:00:00","DATETIME"),VbaValue::integer(0)});assert(rows->item("OrderValue").asNumber()==1 && linaccess::sqlFieldValue(rows->item("Stamp"),"DATETIME").toString()=="2026-11-09 12:00:00");rows->member("Seek",{VbaValue::string(">="),key});assert(rows->item("OrderValue").asNumber()==1);rows->member("Seek",{VbaValue::string("="),key});assert(rows->member("NoMatch",{}).truthy());int problem=0;try {rows->member("Seek",{VbaValue::string("="),VbaValue::string("2026-11-09"),VbaValue::integer(1)});}catch(const linaccess::VbaRuntimeError& error) {problem=error.number;}assert(problem==13);rows->member("Close",{});
    Host host(db);const auto result=linaccess::executeVbaProcedure("Sub Run()\nDim rows As DAO.Recordset\nSet rows=CurrentDb.OpenRecordset(\"SeekDates\")\nrows.Index=\"ByDate\"\nrows.Seek \"=\",#2026-11-09 12:00:00#,2\nIf rows.NoMatch Or rows!OrderValue<>2 Then Err.Raise 5\nrows.Close\nEnd Sub","Run",host);assert(result.status==linaccess::VbaExecutionStatus::Completed);
}

void testNativeTableIndexSeek() {
    using linaccess::VbaValue;auto db=QSqlDatabase::addDatabase("QSQLITE","dao-native-index-seek");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);auto database=linaccess::makeDaoDatabase(db);auto table=database->member("CreateTableDef",{VbaValue::string("SeekRows")}).object;
    for(const auto* name:{"ID","A","B","N","Label"}) {const bool label=QString(name)=="Label";auto field=table->member("CreateField",{VbaValue::string(name),VbaValue::integer(label ? 10 : 4),VbaValue::longInteger(label ? 8 : 0)}).object;field->assign("Required",VbaValue::logical(QString(name)!="N"));table->member("Fields",{}).object->member("Append",{VbaValue::reference(field)});}const auto addIndex=[&](const char* name,QStringList fields,bool primary=false,bool descending=false) {auto index=table->member("CreateIndex",{VbaValue::string(name)}).object;if(primary) index->assign("Primary",VbaValue::logical(true));for(const auto& fieldName:fields) {auto field=index->member("CreateField",{VbaValue::string(fieldName.toStdString())}).object;if(descending) field->assign("Attributes",VbaValue::longInteger(1));index->member("Fields",{}).object->member("Append",{VbaValue::reference(field)});}table->member("Indexes",{}).object->member("Append",{VbaValue::reference(index)});};addIndex("PrimaryKey",{"ID"},true);addIndex("ByA",{"A"});addIndex("ByAB",{"A","B"});addIndex("DescA",{"A"},false,true);addIndex("Nullable",{"N"});addIndex("TextKey",{"Label"});database->member("TableDefs",{}).object->member("Append",{VbaValue::reference(table)});
    for(const auto* values:{"30,3,2,NULL,'c'","10,1,2,NULL,'a'","20,3,1,NULL,'b'","40,5,0,NULL,'d'"}) database->member("Execute",{VbaValue::string((QString("INSERT INTO SeekRows VALUES(")+values+")").toStdString())});
    const auto error=[](const std::function<void()>& operation,int expected) {int number=0;try {operation();}catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;}assert(number==expected);};auto rows=database->member("OpenRecordset",{VbaValue::string("SeekRows")}).object;assert(rows->member("Type",{}).asNumber()==1 && rows->member("Index",{}).asText().empty());error([&] {rows->member("Seek",{VbaValue::string("="),VbaValue::integer(3)});},3251);rows->assign("Index",VbaValue::string("bya"));assert(rows->member("Index",{}).asText()=="ByA");rows->member("MoveFirst",{});assert(rows->item("ID").asNumber()==10);rows->member("MoveLast",{});assert(rows->item("ID").asNumber()==40);
    const auto seek=[&](const char* comparison,int value,int id) {rows->member("Seek",{VbaValue::string(comparison),VbaValue::longInteger(value)});assert(!rows->member("NoMatch",{}).truthy() && rows->item("ID").asNumber()==id);};seek("=",3,30);seek(">=",2,30);seek(">",3,40);seek("<",3,10);seek("<=",3,20);const auto bookmark=rows->member("Bookmark",{});rows->member("Seek",{VbaValue::string("="),VbaValue::integer(99)});assert(rows->member("NoMatch",{}).truthy());error([&] {(void)rows->item("ID");},3021);rows->assign("Bookmark",bookmark);assert(rows->item("ID").asNumber()==20 && !rows->member("NoMatch",{}).truthy());error([&] {rows->member("Seek",{VbaValue::string("="),VbaValue::string("3")});},13);
    rows->assign("Index",VbaValue::string("ByAB"));rows->member("Seek",{VbaValue::string("="),VbaValue::integer(3),VbaValue::integer(1)});assert(rows->item("ID").asNumber()==20);const auto compositeBookmark=rows->member("Bookmark",{});rows->member("Seek",{VbaValue::string("="),VbaValue::integer(3)});assert(rows->member("NoMatch",{}).truthy());rows->assign("Bookmark",compositeBookmark);rows->member("Seek",{VbaValue::string(">="),VbaValue::integer(3)});assert(rows->item("ID").asNumber()==20);rows->member("Seek",{VbaValue::string("<="),VbaValue::integer(3)});assert(rows->item("ID").asNumber()==10);rows->member("Seek",{VbaValue::string(">="),VbaValue::integer(3),VbaValue::missing()});assert(rows->item("ID").asNumber()==20);error([&] {rows->member("Seek",{VbaValue::string("="),VbaValue::integer(3),VbaValue{}});},3251);auto clone=rows->member("Clone",{}).object;assert(clone->member("Index",{}).asText()=="ByAB");clone->member("MoveFirst",{});assert(clone->item("ID").asNumber()==10 && rows->item("ID").asNumber()==20);error([&] {rows->assign("Index",VbaValue::string("ByA"));},3251);clone->member("Close",{});rows->assign("Index",VbaValue::string("DescA"));rows->member("MoveFirst",{});assert(rows->item("ID").asNumber()==40);seek("=",3,30);error([&] {rows->member("Seek",{VbaValue::string(">"),VbaValue::integer(1)});},3251);error([&] {rows->assign("Index",VbaValue::string("Nullable"));},3251);error([&] {rows->assign("Index",VbaValue::string("TextKey"));},3251);error([&] {rows->assign("Index",VbaValue::string("Missing"));},3265);assert(rows->member("Index",{}).asText()=="DescA");rows->member("Close",{});
    auto tableRows=table->member("OpenRecordset",{}).object;assert(tableRows->member("Type",{}).asNumber()==1);tableRows->member("Close",{});auto dynaset=database->member("OpenRecordset",{VbaValue::string("SeekRows"),VbaValue::integer(2)}).object;error([&] {dynaset->assign("Index",VbaValue::string("ByA"));},3251);dynaset->member("Close",{});error([&] {(void)database->member("OpenRecordset",{VbaValue::string("SELECT * FROM SeekRows"),VbaValue::integer(1)});},3251);
    Host host(db);const auto result=linaccess::executeVbaProcedure("Sub Run()\nDim rs As DAO.Recordset\nSet rs=CurrentDb.OpenRecordset(\"SeekRows\",dbOpenTable)\nrs.Index=\"ByAB\"\nrs.Seek Key2:=1,Comparison:=\"=\",Key1:=3\nIf rs.NoMatch Or rs!ID<>20 Then Err.Raise 5\nrs.Close\nEnd Sub","Run",host);assert(result.status==linaccess::VbaExecutionStatus::Completed);
}

void testNativeNumericTypeChanges() {
    using linaccess::VbaValue;auto db=QSqlDatabase::addDatabase("QSQLITE","dao-native-numeric-type-changes");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);auto database=linaccess::makeDaoDatabase(db);auto table=database->member("CreateTableDef",{VbaValue::string("NumberTypes")}).object;auto id=table->member("CreateField",{VbaValue::string("ID"),VbaValue::integer(4)}).object;id->assign("Attributes",VbaValue::longInteger(16));table->member("Fields",{}).object->member("Append",{VbaValue::reference(id)});auto field=table->member("CreateField",{VbaValue::string("Measured"),VbaValue::integer(4)}).object;field->assign("DefaultValue",VbaValue::string("7"));field->assign("ValidationRule",VbaValue::string(">=0"));field->assign("ValidationText",VbaValue::string("Nicht negativ"));table->member("Fields",{}).object->member("Append",{VbaValue::reference(field)});auto index=table->member("CreateIndex",{VbaValue::string("MeasuredIndex")}).object;index->member("Fields",{}).object->member("Append",{index->member("CreateField",{VbaValue::string("Measured")})});table->member("Indexes",{}).object->member("Append",{VbaValue::reference(index)});database->member("TableDefs",{}).object->member("Append",{VbaValue::reference(table)});database->member("Execute",{VbaValue::string("INSERT INTO NumberTypes DEFAULT VALUES")});database->member("Execute",{VbaValue::string("INSERT INTO NumberTypes(Measured) VALUES(NULL)")});
    const auto error=[](const std::function<void()>& operation,int expected) {int number=0;try {operation();}catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;}assert(number==expected);};error([&] {field->assign("Type",VbaValue::integer(7));},383);error([&] {linaccess::changeNativeNumericFieldType(db,"NumberTypes","ID",7);},3251);auto rows=database->member("OpenRecordset",{VbaValue::string("NumberTypes")}).object;linaccess::changeNativeNumericFieldType(db,"NumberTypes","Measured",7);error([&] {(void)rows->item("Measured");},3420);rows->member("Requery",{});assert(rows->member("Fields",{VbaValue::string("Measured")}).object->member("Type",{}).asNumber()==7 && rows->item("Measured").asNumber()==7);rows->member("Close",{});assert(sql.exec("SELECT typeof(Measured),COUNT(*) FROM NumberTypes WHERE Measured=7") && sql.next() && sql.value(0).toString()=="real" && sql.value(1).toInt()==1);sql.finish();assert(sql.exec("SELECT COUNT(*) FROM NumberTypes WHERE Measured IS NULL") && sql.next() && sql.value(0).toInt()==1);sql.finish();
    linaccess::changeNativeNumericFieldType(db,"NumberTypes","Measured",2);assert(table->member("Fields",{VbaValue::string("Measured")}).object->member("Type",{}).asNumber()==2);linaccess::changeNativeNumericFieldType(db,"NumberTypes","Measured",3);linaccess::changeNativeNumericFieldType(db,"NumberTypes","Measured",4);linaccess::changeNativeNumericFieldType(db,"NumberTypes","Measured",6);linaccess::changeNativeNumericFieldType(db,"NumberTypes","Measured",7);database->member("Execute",{VbaValue::string("UPDATE NumberTypes SET Measured=0.1 WHERE ID=1")});error([&] {linaccess::changeNativeNumericFieldType(db,"NumberTypes","Measured",6);},3251);error([&] {linaccess::changeNativeNumericFieldType(db,"NumberTypes","Measured",4);},3251);assert(table->member("Fields",{VbaValue::string("Measured")}).object->member("Type",{}).asNumber()==7);assert(sql.exec("SELECT Measured FROM NumberTypes WHERE ID=1") && sql.next() && sql.value(0).toDouble()==0.1);sql.finish();
    database->member("Execute",{VbaValue::string("UPDATE NumberTypes SET Measured=300 WHERE ID=1")});error([&] {linaccess::changeNativeNumericFieldType(db,"NumberTypes","Measured",2);},6);database->member("Execute",{VbaValue::string("UPDATE NumberTypes SET Measured=7 WHERE ID=1")});field=table->member("Fields",{VbaValue::string("Measured")}).object;field->assign("DefaultValue",VbaValue::string("0.1"));error([&] {linaccess::changeNativeNumericFieldType(db,"NumberTypes","Measured",6);},3251);field->assign("DefaultValue",VbaValue::string("7"));error([&] {database->member("Execute",{VbaValue::string("UPDATE NumberTypes SET Measured=-1")});},3317);assert(table->member("Indexes",{}).object->member("Count",{}).asNumber()==1);assert(sql.exec("SELECT value FROM _la_counter32_sequences WHERE table_name='NumberTypes'") && sql.next() && sql.value(0).toInt()==2);sql.finish();database->member("Execute",{VbaValue::string("INSERT INTO NumberTypes DEFAULT VALUES")});assert(sql.exec("SELECT ID,Measured FROM NumberTypes WHERE ID=3") && sql.next() && sql.value(1).toDouble()==7);sql.finish();
}

void testNativeDateValidation() {
    using linaccess::VbaValue;auto db=QSqlDatabase::addDatabase("QSQLITE","dao-native-date-validation");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);auto database=linaccess::makeDaoDatabase(db);auto table=database->member("CreateTableDef",{VbaValue::string("CheckedDates")}).object;
    auto field=table->member("CreateField",{VbaValue::string("Stamp"),VbaValue::integer(8)}).object;field->assign("DefaultValue",VbaValue::string("Date()"));field->assign("ValidationRule",VbaValue::string("Between #2000-01-01# And #2099-12-31 23:59:59#"));field->assign("ValidationText",VbaValue::string("Datum außerhalb des Bereichs"));table->member("Fields",{}).object->member("Append",{VbaValue::reference(field)});database->member("TableDefs",{}).object->member("Append",{VbaValue::reference(table)});database->member("Execute",{VbaValue::string("INSERT INTO CheckedDates DEFAULT VALUES")});database->member("Execute",{VbaValue::string("INSERT INTO CheckedDates VALUES(NULL)")});
    const auto error=[](const std::function<void()>& operation,int expected) {int number=0;try {operation();}catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;}assert(number==expected);};error([&] {database->member("Execute",{VbaValue::string("INSERT INTO CheckedDates VALUES('1999-12-31')")});},3317);error([&] {field->assign("ValidationRule",VbaValue::string(">=#2026-02-30#"));},3251);error([&] {field->assign("ValidationRule",VbaValue::string(">=#11/08/2026#"));},3251);error([&] {field->assign("ValidationRule",VbaValue::string(">=0"));},3251);
    field->assign("ValidationRule",VbaValue::string(">=#2000-01-01# And <=Date()"));database->member("Execute",{VbaValue::string("INSERT INTO CheckedDates DEFAULT VALUES")});error([&] {database->member("Execute",{VbaValue::string("INSERT INTO CheckedDates VALUES('9999-12-31')")});},3317);field->assign("ValidationRule",VbaValue::string("<=Now()"));database->member("Execute",{VbaValue::string("INSERT INTO CheckedDates DEFAULT VALUES")});field->assign("DefaultValue",VbaValue::string("Time()"));field->assign("ValidationRule",VbaValue::string(">=#1899-12-30# And <=Now()"));database->member("Execute",{VbaValue::string("INSERT INTO CheckedDates DEFAULT VALUES")});assert(sql.exec("SELECT COUNT(*) FROM CheckedDates WHERE Stamp LIKE '1899-12-30 %'") && sql.next() && sql.value(0).toInt()==1);sql.finish();field->assign("Name",VbaValue::string("WhenValue"));table->assign("Name",VbaValue::string("RenamedCheckedDates"));error([&] {database->member("Execute",{VbaValue::string("INSERT INTO RenamedCheckedDates VALUES('9999-12-31')")});},3317);
}

void testNativeFieldOrder() {
    using linaccess::VbaValue;auto db=QSqlDatabase::addDatabase("QSQLITE","dao-native-field-order");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);auto database=linaccess::makeDaoDatabase(db);auto table=database->member("CreateTableDef",{VbaValue::string("OrderRows")}).object;
    auto id=table->member("CreateField",{VbaValue::string("ID"),VbaValue::integer(4)}).object;id->assign("Attributes",VbaValue::longInteger(16));table->member("Fields",{}).object->member("Append",{VbaValue::reference(id)});auto value=table->member("CreateField",{VbaValue::string("Value"),VbaValue::integer(6)}).object;value->assign("DefaultValue",VbaValue::string("0.1"));value->assign("ValidationRule",VbaValue::string(">=0"));table->member("Fields",{}).object->member("Append",{VbaValue::reference(value)});auto stamp=table->member("CreateField",{VbaValue::string("Stamp"),VbaValue::integer(8)}).object;stamp->assign("DefaultValue",VbaValue::string("Now()"));table->member("Fields",{}).object->member("Append",{VbaValue::reference(stamp)});table->member("Fields",{}).object->member("Append",{table->member("CreateField",{VbaValue::string("Payload"),VbaValue::integer(11)})});table->member("Fields",{}).object->member("Append",{table->member("CreateField",{VbaValue::string("Label"),VbaValue::integer(10),VbaValue::longInteger(8)})});
    for(const char* field:{"ID","Value"}) {auto index=table->member("CreateIndex",{VbaValue::string(field)}).object;if(QString(field)=="ID") index->assign("Primary",VbaValue::logical(true));index->member("Fields",{}).object->member("Append",{index->member("CreateField",{VbaValue::string(field)})});table->member("Indexes",{}).object->member("Append",{VbaValue::reference(index)});}database->member("TableDefs",{}).object->member("Append",{VbaValue::reference(table)});
    database->member("Execute",{VbaValue::string("INSERT INTO OrderRows DEFAULT VALUES")});database->member("Execute",{VbaValue::string("INSERT INTO OrderRows DEFAULT VALUES")});sql.prepare("UPDATE OrderRows SET Payload=? WHERE ID=1");sql.addBindValue(QByteArray::fromHex("0001ff"));assert(sql.exec());sql.finish();assert(sql.exec("CREATE VIEW OrderView AS SELECT * FROM OrderRows"));assert(sql.exec("CREATE TABLE AuditOrder(value INTEGER)"));assert(sql.exec("CREATE TRIGGER AuditOrderUpdate AFTER UPDATE ON OrderRows BEGIN INSERT INTO AuditOrder VALUES(1); END"));
    const auto error=[](const std::function<void()>& operation,int expected) {int number=0;try {operation();}catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;}assert(number==expected);};auto payload=table->member("Fields",{VbaValue::string("Payload")}).object;payload->assign("OrdinalPosition",VbaValue::integer(3));error([&] {payload->assign("OrdinalPosition",VbaValue::longInteger(-1));},3251);error([&] {payload->assign("OrdinalPosition",VbaValue::longInteger(5));},3251);
    auto rows=database->member("OpenRecordset",{VbaValue::string("OrderRows")}).object;auto clone=rows->member("Clone",{}).object;rows->member("Edit",{});rows->assign("Value",VbaValue::single(0.25f));payload->assign("OrdinalPosition",VbaValue::integer(0));assert(payload->member("OrdinalPosition",{}).asNumber()==0);error([&] {(void)rows->item("Value");},3420);error([&] {(void)clone->member("Fields",{VbaValue::string("Payload")});},3420);rows->member("Requery",{});assert(rows->member("Fields",{VbaValue::integer(0)}).object->member("Name",{}).asText()=="Payload" && rows->item("Value").asNumber()==static_cast<double>(0.1f));rows->member("Close",{});clone->member("Close",{});
    auto counter=table->member("Fields",{VbaValue::string("ID")}).object;counter->assign("OrdinalPosition",VbaValue::integer(4));assert(counter->member("OrdinalPosition",{}).asNumber()==4);assert(sql.exec("SELECT * FROM OrderView WHERE ID=1") && sql.record().fieldName(0)=="Payload" && sql.record().fieldName(4)=="ID" && sql.next() && sql.value(0).toByteArray()==QByteArray::fromHex("0001ff"));sql.finish();assert(sql.exec("SELECT value FROM _la_counter32_sequences WHERE table_name='OrderRows'") && sql.next() && sql.value(0).toInt()==2);sql.finish();database->member("Execute",{VbaValue::string("INSERT INTO OrderRows DEFAULT VALUES")});database->member("Execute",{VbaValue::string("UPDATE OrderRows SET Value=0.3 WHERE ID=3")});assert(sql.exec("SELECT COUNT(*) FROM AuditOrder") && sql.next() && sql.value(0).toInt()==1);sql.finish();assert(sql.exec("SELECT Value,Stamp FROM OrderRows WHERE ID=3") && sql.next() && sql.value(0).toDouble()==static_cast<double>(0.3f) && QDateTime::fromString(sql.value(1).toString(),"yyyy-MM-dd HH:mm:ss").isValid());sql.finish();
}

void testNativeNumericValidation() {
    using linaccess::VbaValue;auto db=QSqlDatabase::addDatabase("QSQLITE","dao-native-validation");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);auto database=linaccess::makeDaoDatabase(db);auto table=database->member("CreateTableDef",{VbaValue::string("ValidationRows")}).object;
    auto id=table->member("CreateField",{VbaValue::string("ID"),VbaValue::integer(4)}).object;id->assign("Attributes",VbaValue::longInteger(16));table->member("Fields",{}).object->member("Append",{VbaValue::reference(id)});auto amount=table->member("CreateField",{VbaValue::string("Amount"),VbaValue::integer(4)}).object;amount->assign("ValidationRule",VbaValue::string(">=0 And <=10"));amount->assign("ValidationText",VbaValue::string("Wert außerhalb des Bereichs 'ÄÖ'"));amount->assign("DefaultValue",VbaValue::string("5"));table->member("Fields",{}).object->member("Append",{VbaValue::reference(amount)});database->member("TableDefs",{}).object->member("Append",{VbaValue::reference(table)});assert(amount->member("ValidationRule",{}).asText()==">=0 And <=10");
    const auto error=[](const std::function<void()>& operation,int expected) {int number=0;try {operation();}catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;}assert(number==expected);};database->member("Execute",{VbaValue::string("INSERT INTO ValidationRows DEFAULT VALUES")});database->member("Execute",{VbaValue::string("INSERT INTO ValidationRows(Amount) VALUES(10)")});database->member("Execute",{VbaValue::string("INSERT INTO ValidationRows(Amount) VALUES(NULL)")});error([&] {database->member("Execute",{VbaValue::string("UPDATE ValidationRows SET Amount=Amount+1")});},3317);assert(sql.exec("SELECT SUM(Amount),COUNT(*) FROM ValidationRows") && sql.next() && sql.value(0).toInt()==15 && sql.value(1).toInt()==3);sql.finish();
    auto rows=database->member("OpenRecordset",{VbaValue::string("ValidationRows")}).object;auto metadata=rows->member("Fields",{VbaValue::string("Amount")}).object;assert(metadata->member("ValidationRule",{}).asText()==">=0 And <=10" && metadata->member("ValidationText",{}).asText()=="Wert außerhalb des Bereichs 'ÄÖ'");error([&] {metadata->assign("ValidationText",VbaValue::string("changed"));},383);rows->member("MoveFirst",{});rows->member("Edit",{});rows->assign("Amount",VbaValue::longInteger(11));{int number=0;std::string description;try {rows->member("Update",{});}catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;description=failure.what();}assert(number==3317 && description=="Wert außerhalb des Bereichs 'ÄÖ'");}assert(rows->item("Amount").asNumber()==11);rows->assign("Amount",VbaValue::longInteger(6));rows->member("Update",{});rows->member("Close",{});
    auto field=table->member("Fields",{VbaValue::string("Amount")}).object;error([&] {field->assign("ValidationRule",VbaValue::string("Between 0 And 5"));},3317);assert(field->member("ValidationRule",{}).asText()==">=0 And <=10");error([&] {field->assign("ValidationRule",VbaValue::string(">=0; DELETE FROM ValidationRows"));},3251);error([&] {field->assign("ValidationRule",VbaValue::string(">=1e999"));},3251);error([&] {field->assign("ValidationRule",VbaValue::string(">=0 Or"));},3251);error([&] {field->assign("ValidationText",VbaValue::string(std::string(256,'x')));},3251);field->assign("ValidationText",VbaValue::string("Neuer Bereich"));field->assign("ValidationRule",VbaValue::string("Between 0 And 12"));assert(field->member("ValidationRule",{}).asText()=="Between 0 And 12");database->member("Execute",{VbaValue::string("UPDATE ValidationRows SET Amount=12 WHERE Amount=10")});field->assign("Name",VbaValue::string("Measured"));table->assign("Name",VbaValue::string("RenamedValidation"));error([&] {database->member("Execute",{VbaValue::string("INSERT INTO RenamedValidation(Measured) VALUES(13)")});},3317);
    field=table->member("Fields",{VbaValue::string("Measured")}).object;field->assign("ValidationRule",VbaValue::string("<0 Or >=5 And <=12"));database->member("Execute",{VbaValue::string("INSERT INTO RenamedValidation(Measured) VALUES(-2)")});error([&] {database->member("Execute",{VbaValue::string("INSERT INTO RenamedValidation(Measured) VALUES(3)")});},3317);field->assign("ValidationText",VbaValue::string(""));field->assign("ValidationRule",VbaValue::string(""));database->member("Execute",{VbaValue::string("INSERT INTO RenamedValidation(Measured) VALUES(13)")});assert(field->member("ValidationRule",{}).asText().empty());
    for(const int type:{1,2,3,6,7}) {auto checked=table->member("CreateField",{VbaValue::string(("N"+QString::number(type)).toStdString()),VbaValue::integer(type)}).object;checked->assign("ValidationRule",VbaValue::string(">=-1e0 And <=1"));table->member("Fields",{}).object->member("Append",{VbaValue::reference(checked)});}database->member("Execute",{VbaValue::string("INSERT INTO RenamedValidation(N1,N2,N3,N6,N7) VALUES(True,1,-1,0.1,0.1)")});error([&] {database->member("Execute",{VbaValue::string("UPDATE RenamedValidation SET N7=2")});},3317);
}

void testNativeDateRelations() {
    using linaccess::VbaValue;auto db=QSqlDatabase::addDatabase("QSQLITE","dao-native-date-relations");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);assert(sql.exec("PRAGMA foreign_keys=ON"));auto database=linaccess::makeDaoDatabase(db);
    for(const auto* name:{"DateParent","DateChild"}) {auto table=database->member("CreateTableDef",{VbaValue::string(name)}).object;table->member("Fields",{}).object->member("Append",{table->member("CreateField",{VbaValue::string("KeyValue"),VbaValue::integer(8)})});if(QString(name)=="DateParent") {auto key=table->member("CreateIndex",{VbaValue::string("PrimaryKey")}).object;key->assign("Primary",VbaValue::logical(true));key->member("Fields",{}).object->member("Append",{key->member("CreateField",{VbaValue::string("KeyValue")})});table->member("Indexes",{}).object->member("Append",{VbaValue::reference(key)});}else {auto clock=table->member("CreateField",{VbaValue::string("Moment"),VbaValue::integer(8)}).object;clock->assign("DefaultValue",VbaValue::string("Now()"));table->member("Fields",{}).object->member("Append",{VbaValue::reference(clock)});table->member("Fields",{}).object->member("Append",{table->member("CreateField",{VbaValue::string("Payload"),VbaValue::integer(11)})});}database->member("TableDefs",{}).object->member("Append",{VbaValue::reference(table)});assert(linaccess::nativeFieldHasCanonicalDateStorage(db,name,"KeyValue"));}
    database->member("Execute",{VbaValue::string("INSERT INTO DateParent VALUES('2026-11-08')")});sql.prepare("INSERT INTO DateChild(KeyValue,Payload) VALUES(?,?)");sql.addBindValue("2026-11-08 00:00:00");sql.addBindValue(QByteArray::fromHex("0001ff"));assert(sql.exec());sql.finish();database->member("Execute",{VbaValue::string("INSERT INTO DateChild(KeyValue) VALUES(NULL)")});assert(sql.exec("SELECT Moment FROM DateChild WHERE KeyValue IS NOT NULL") && sql.next());const auto frozen=sql.value(0).toString();sql.finish();
    const auto relation=[&](const char* name,const char* parent) {auto result=database->member("CreateRelation",{VbaValue::string(name),VbaValue::string(parent),VbaValue::string("DateChild"),VbaValue::longInteger(256|4096)}).object;auto pair=result->member("CreateField",{VbaValue::string("KeyValue")}).object;pair->assign("ForeignName",VbaValue::string("KeyValue"));result->member("Fields",{}).object->member("Append",{VbaValue::reference(pair)});return result;};auto mapping=relation("DateMapping","DateParent");database->member("Relations",{}).object->member("Append",{VbaValue::reference(mapping)});
    const auto error=[](const std::function<void()>& operation,int expected) {int number=0;try {operation();}catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;}assert(number==expected);};error([&] {database->member("Execute",{VbaValue::string("INSERT INTO DateChild(KeyValue) VALUES('2026-11-10')")});},3146);error([&] {database->member("Execute",{VbaValue::string("INSERT INTO DateParent VALUES('2026-11-08 00:00:00')")});},3022);
    database->member("Execute",{VbaValue::string("UPDATE DateParent SET KeyValue='2026-11-09'")});assert(sql.exec("SELECT KeyValue,Moment,Payload FROM DateChild WHERE KeyValue IS NOT NULL") && sql.next() && sql.value(0).toString()=="2026-11-09 00:00:00" && sql.value(1).toString()==frozen && sql.value(2).toByteArray()==QByteArray::fromHex("0001ff"));sql.finish();database->member("Execute",{VbaValue::string("DELETE FROM DateParent")});assert(sql.exec("SELECT COUNT(*) FROM DateChild WHERE KeyValue IS NULL") && sql.next() && sql.value(0).toInt()==1);sql.finish();database->member("Relations",{}).object->member("Delete",{VbaValue::string("DateMapping")});
    assert(sql.exec("CREATE TABLE LegacyDateParent(KeyValue DATETIME NOT NULL UNIQUE)"));const QJsonArray metadata{QJsonObject{{"name","KeyValue"},{"source_type","DATETIME"},{"source_origin","DAO_NATIVE"},{"native_attributes",1},{"native_required",true}}};sql.prepare("INSERT INTO _la_table_schema VALUES('LegacyDateParent',?)");sql.addBindValue(QJsonDocument(metadata).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();assert(!linaccess::nativeFieldHasCanonicalDateStorage(db,"LegacyDateParent","KeyValue"));auto unsupported=relation("LegacyDateMapping","LegacyDateParent");error([&] {database->member("Relations",{}).object->member("Append",{VbaValue::reference(unsupported)});},3251);assert(database->member("Relations",{}).object->member("Count",{}).asNumber()==0);
}

void testNativeDateDefaults() {
    using linaccess::VbaValue;auto db=QSqlDatabase::addDatabase("QSQLITE","dao-native-date-defaults");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);auto database=linaccess::makeDaoDatabase(db);auto table=database->member("CreateTableDef",{VbaValue::string("DateRows")}).object;
    auto id=table->member("CreateField",{VbaValue::string("ID"),VbaValue::integer(4)}).object;id->assign("Attributes",VbaValue::longInteger(16));table->member("Fields",{}).object->member("Append",{VbaValue::reference(id)});
    for(const auto& item:std::vector<std::pair<const char*,const char*>>{{"Today","Date()"},{"Moment","= Now ( )"},{"Clock","Time()"},{"Fixed","= #2026-11-08#"}}) {auto field=table->member("CreateField",{VbaValue::string(item.first),VbaValue::integer(8)}).object;field->assign("DefaultValue",VbaValue::string(item.second));table->member("Fields",{}).object->member("Append",{VbaValue::reference(field)});}database->member("TableDefs",{}).object->member("Append",{VbaValue::reference(table)});
    assert(table->member("Fields",{VbaValue::string("Moment")}).object->member("DefaultValue",{}).asText()=="Now()");const auto before=QDateTime::currentDateTime();database->member("Execute",{VbaValue::string("INSERT INTO DateRows DEFAULT VALUES")});const auto after=QDateTime::currentDateTime();
    assert(sql.exec("SELECT Today,Moment,Clock,Fixed FROM DateRows") && sql.next());const auto today=sql.value(0).toString();const auto moment=QDateTime::fromString(sql.value(1).toString(),"yyyy-MM-dd HH:mm:ss");const auto clock=sql.value(2).toString();assert(today==before.date().toString(Qt::ISODate)+" 00:00:00" || today==after.date().toString(Qt::ISODate)+" 00:00:00");assert(moment.isValid() && moment.toSecsSinceEpoch()>=before.toSecsSinceEpoch() && moment.toSecsSinceEpoch()<=after.toSecsSinceEpoch());assert(clock.startsWith("1899-12-30 ") && QTime::fromString(clock.mid(11),"HH:mm:ss").isValid());assert(sql.value(3).toString()=="2026-11-08 00:00:00");sql.finish();
    auto rows=database->member("OpenRecordset",{VbaValue::string("DateRows")}).object;rows->member("AddNew",{});assert(rows->item("Today").kind==VbaValue::Kind::Date && rows->item("Moment").kind==VbaValue::Kind::Date && rows->item("Clock").kind==VbaValue::Kind::Date);const auto frozen=linaccess::sqlFieldValue(rows->item("Moment"),"DATETIME").toString();rows->member("Update",{});assert(sql.exec("SELECT Moment FROM DateRows WHERE ID=2") && sql.next() && sql.value(0).toString()==frozen);sql.finish();rows->member("AddNew",{});rows->assign("Moment",VbaValue{});rows->member("Update",{});assert(sql.exec("SELECT Moment FROM DateRows WHERE ID=3") && sql.next() && sql.value(0).isNull());sql.finish();
    const auto literals=linaccess::nativeTableLiteralDefaults(db,"DateRows");assert(!literals.contains("moment") && literals.value("fixed").toString()=="2026-11-08 00:00:00");const auto dates=linaccess::nativeTableDateDefaults(db,"DateRows");assert(dates.size()==4 && !dates.contains("id"));assert(sql.exec("SELECT value FROM _la_counter32_sequences WHERE table_name='DateRows'") && sql.next() && sql.value(0).toInt()==3);sql.finish();
    auto fixed=table->member("Fields",{VbaValue::string("Fixed")}).object;int invalid=0;try {fixed->assign("DefaultValue",VbaValue::string("#2026-02-30#"));}catch(const linaccess::VbaRuntimeError& failure) {invalid=failure.number;}assert(invalid==3251);fixed->assign("DefaultValue",VbaValue::string("= #2026-11-09 12:30#"));
    auto field=table->member("Fields",{VbaValue::string("Moment")}).object;field->assign("DefaultValue",VbaValue::string("Date()"));assert(field->member("DefaultValue",{}).asText()=="Date()");int number=0;try {field->assign("DefaultValue",VbaValue::string("UnsupportedClock()"));}catch(const linaccess::VbaRuntimeError& failure) {number=failure.number;}assert(number==3251 && field->member("DefaultValue",{}).asText()=="Date()");
    rows->member("Requery",{});rows->member("AddNew",{});assert(linaccess::sqlFieldValue(rows->item("Moment"),"DATETIME").toString().endsWith(" 00:00:00"));rows->member("CancelUpdate",{});table->assign("Name",VbaValue::string("RenamedDates"));database->member("Execute",{VbaValue::string("INSERT INTO RenamedDates DEFAULT VALUES")});assert(sql.exec("SELECT Moment FROM RenamedDates WHERE ID=5") && sql.next() && sql.value(0).toString().endsWith(" 00:00:00"));sql.finish();
    assert(sql.exec("SELECT Fixed FROM RenamedDates WHERE ID=5") && sql.next() && sql.value(0).toString()=="2026-11-09 12:30:00");sql.finish();assert(linaccess::accessLiteralDefaultSql(" = 7 ")==QString("7") && linaccess::accessLiteralDefaultSql("=True")==QString("-1") && linaccess::accessLiteralDefaultSql("=\"a=b\"")==QString("'a=b'"));
    assert(!sql.exec("INSERT INTO RenamedDates(Fixed) VALUES('2026-11-08')"));assert(sql.exec("SELECT Moment FROM RenamedDates WHERE ID=2") && sql.next() && sql.value(0).toString()==frozen);sql.finish();
}

void testProjectErrorContexts() {
    using namespace linaccess;
    const QString firstName="native-error-first",secondName="native-error-second";
    {
        auto first=QSqlDatabase::addDatabase("QSQLITE",firstName),second=QSqlDatabase::addDatabase("QSQLITE",secondName);
        first.setDatabaseName(":memory:");second.setDatabaseName(":memory:");assert(first.open() && second.open());
        auto one=projectVbaErrorObject(first),same=projectVbaErrorObject(first),other=projectVbaErrorObject(second);
        assert(one==same && one!=other);one->assign("Number",VbaValue::longInteger(901));assert(same->member("Number",{}).number==901 && other->member("Number",{}).number==0);
        resetVbaProjectState(first);int code=0;try{(void)one->member("Number",{});}catch(const VbaRuntimeError& error){code=error.number;}assert(code==3420);
        auto replacement=projectVbaErrorObject(first);assert(replacement!=one && replacement->member("Number",{}).number==0);
        first.close();code=0;try{replacement->assign("Number",VbaValue::longInteger(902));}catch(const VbaRuntimeError& error){code=error.number;}assert(code==3420);assert(other->member("Number",{}).number==0);
        second.close();
    }
    QSqlDatabase::removeDatabase(firstName);QSqlDatabase::removeDatabase(secondName);
}
void testNativeDaoArrays() {
    auto db=QSqlDatabase::addDatabase("QSQLITE","native-dao-arrays");db.setDatabaseName(":memory:");assert(db.open());QSqlQuery sql(db);assert(sql.exec("CREATE TABLE native_array_rows(id INTEGER PRIMARY KEY,value INTEGER)"));assert(sql.exec("INSERT INTO native_array_rows VALUES(1,17),(2,23)"));Host host(db);
    const std::string source=R"VBA(Option Explicit
Private Sub GrowDatabases(ByRef items() As DAO.Database)
ReDim Preserve items(1 To 2)
Set items(2)=CurrentDb
End Sub
Private Function DatabaseArray As DAO.Database()
Dim items(1 To 1) As DAO.Database
Set items(1)=CurrentDb
DatabaseArray=items
End Function
Function NativeDaoArrayProbe As Long
Dim dbs() As DAO.Database, sets(0 To 0) As DAO.Recordset, fields(0 To 0) As DAO.Field, current As DAO.Database, item As Variant
ReDim dbs(1 To 1)
Set current=CurrentDb
Set dbs(1)=current
If Not dbs(1) Is current Then Err.Raise 5
Set sets(0)=dbs(1).OpenRecordset("SELECT id,value FROM native_array_rows ORDER BY id",dbOpenSnapshot)
Set fields(0)=sets(0).Fields("value")
NativeDaoArrayProbe=fields(0).Value
sets(0).MoveNext
NativeDaoArrayProbe=NativeDaoArrayProbe+fields(0).Value
GrowDatabases dbs
If TypeName(dbs)<>"Database()" Or TypeName(sets)<>"Recordset()" Or TypeName(fields)<>"Field()" Then Err.Raise 5
If VarType(dbs)<>8201 Then Err.Raise 5
NativeDaoArrayProbe=NativeDaoArrayProbe+UBound(dbs)+sets(0).Type
For Each item In DatabaseArray()
If Not TypeOf item Is DAO.Database Then Err.Raise 5
NativeDaoArrayProbe=NativeDaoArrayProbe+1
Next
On Error Resume Next
Set sets(0)=CurrentDb
NativeDaoArrayProbe=NativeDaoArrayProbe+Err.Number
On Error GoTo 0
sets(0).Close
End Function
Function ClosedDaoArrayProbe As Long
Dim sets(0 To 0) As DAO.Recordset, fields(0 To 0) As DAO.Field
Set sets(0)=CurrentDb.OpenRecordset("native_array_rows")
Set fields(0)=sets(0).Fields("value")
sets(0).Close
On Error Resume Next
ClosedDaoArrayProbe=fields(0).Value
ClosedDaoArrayProbe=Err.Number
End Function
)VBA";
    for(const auto& probe:std::vector<std::pair<const char*,int>>{{"NativeDaoArrayProbe",60},{"ClosedDaoArrayProbe",3420}}) {const auto result=linaccess::executeVbaProcedure(source,probe.first,host);if(result.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"DAO array "<<probe.first<<": "<<result.explanation<<" line "<<result.line<<" error "<<result.errorNumber<<"\n";assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->number==probe.second);}
}
int main(int argc, char** argv) {
    QCoreApplication app(argc, argv);
    testProjectErrorContexts();
    testNativeDaoArrays();
    testNativeIntrinsicFieldProperties();
    testNativeFieldDisplayProperties();
    testNativeDateIndexSeek();
    testNativeTableIndexSeek();
    testNativeNumericTypeChanges();
    testNativeDateValidation();
    testNativeFieldOrder();
    testNativeNumericValidation();
    testNativeDateRelations();
    testNativeDateDefaults();
    testNativeSingleRelations();
    testNativeSingleSchemaAndWrites();
    testNativeDaoWriteCapability();
    testExistingNativeTableRename();
    testShadowedDaoRowIdentity();
    testExistingNativePrimaryIndexes();
    testNativeFieldPropertyChanges();
    testOneToOneNativeRelations();
    testEnforcedNativeRelations();
    testDetachedRelationCreation();
    testDetachedTableCreation();
    testVariantArraysDao();
    testAccessEvalNumericOracle();
    testOwnedQueryRename();
    testDaoGetRows();
    testOwnedQueryDefs();
    testSqlVbaFunctions();
    testDdlRecordsAffected();
    testTypedSqlTargets();
    testDeclaredParameterTypes();
    testIntegerConversionDao();
    testBooleanQueryParameters();
    testDaoCursors();
    testDaoTransactions();
    testTableDefSchemaFields();
    testDaoRelations();
    testDaoSingleFields();
    testDaoIndexes();
    testDaoTableGenerations();
    testSqlRgbDao();
    auto db = QSqlDatabase::addDatabase("QSQLITE"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery q(db);
    assert(q.exec("CREATE TABLE items (id INTEGER PRIMARY KEY, qty REAL, label TEXT DEFAULT 'default')"));
    assert(q.exec("INSERT INTO items VALUES (1,2,NULL),(2,3,'two')"));
    assert(q.exec("CREATE TABLE unicode_values(id INTEGER PRIMARY KEY,[Änderung] INTEGER)"));
    assert(q.exec("INSERT INTO unicode_values VALUES(1,3)"));
    Host host(db);
    const auto unicode=linaccess::executeVbaProcedure(R"VBA(Sub Übertragen()
Dim größe As Double
Dim sätze As DAO.Recordset
größe=4
Set sätze=CurrentDb.OpenRecordset("unicode_values")
sätze.Edit
sätze!Änderung=sätze!Änderung+größe
sätze.Update
sätze.Close
End Sub)VBA","Übertragen",host);
    assert(unicode.status==linaccess::VbaExecutionStatus::Completed);
    assert(q.exec("SELECT [Änderung] FROM unicode_values WHERE id=1") && q.next() && q.value(0).toInt()==7); q.finish();
    {
        auto database=linaccess::makeDaoDatabase(db);
        const auto definitions=database->member("TableDefs",{}).object;
        assert(definitions==database->member("TableDefs",{}).object);
        int wrongThreadMember=0,wrongThreadEnumeration=0;
        std::thread foreign([&] {
            try { (void)definitions->member("Count",{}); } catch(const linaccess::VbaRuntimeError& error) { wrongThreadMember=error.number; }
            try { (void)definitions->enumerate(); } catch(const linaccess::VbaRuntimeError& error) { wrongThreadEnumeration=error.number; }
        });
        foreign.join(); assert(wrongThreadMember==3251 && wrongThreadEnumeration==3251);
        assert(definitions->member("Count",{}).asNumber()==2 && definitions->member("Count",{}).numericType==linaccess::VbaValue::NumericType::Integer);
        assert(q.exec("CREATE TABLE _la_private_metadata(secret TEXT)"));
        assert(q.exec("CREATE VIEW visible_query AS SELECT * FROM items"));
        assert(definitions->member("Count",{}).asNumber()==2);
        const auto table=database->member("TableDefs",{linaccess::VbaValue::string("ITEMS")}).object;
        assert(table==definitions->member("Item",{linaccess::VbaValue::numeric(0)}).object);
        assert(table==definitions->item("items").object);
        assert(table->member("Name",{}).asText()=="items");
        const auto recordCount=table->member("RecordCount",{}); assert(recordCount.asNumber()==2 && recordCount.numericType==linaccess::VbaValue::NumericType::Long);
        try { table->assign("RecordCount",linaccess::VbaValue::integer(8)); assert(false); }
        catch(const linaccess::VbaRuntimeError& error) { assert(error.number==383); }
        try { (void)table->member("RecordCount",{linaccess::VbaValue::integer(0)}); assert(false); }
        catch(const linaccess::VbaRuntimeError& error) { assert(error.number==450); }
        assert(q.exec("SAVEPOINT record_count_probe"));
        assert(q.exec("INSERT INTO items VALUES(3,9,'probe')")); assert(table->member("RecordCount",{}).asNumber()==3);
        assert(q.exec("ROLLBACK TO record_count_probe")); assert(q.exec("RELEASE record_count_probe"));
        assert(table->member("RecordCount",{}).asNumber()==2);
        assert(q.exec("CREATE TEMP TABLE items(id INTEGER)")); assert(q.exec("INSERT INTO temp.items VALUES(1)"));
        assert(table->member("RecordCount",{}).asNumber()==2); assert(q.exec("DROP TABLE temp.items"));
        assert(definitions->enumerate().size()==2);
        for (const auto& [arguments,number]:std::vector<std::pair<std::vector<linaccess::VbaValue>,int>>{
            {{},450},{{linaccess::VbaValue{}},94},{{linaccess::VbaValue::numeric(-1)},3265},{{linaccess::VbaValue::string("missing")},3265}}) {
            try { (void)definitions->member("Item",arguments); assert(false); }
            catch(const linaccess::VbaRuntimeError& error) { assert(error.number==number); }
        }
        auto rows=table->member("OpenRecordset",{}).object;
        const auto recordFields=rows->member("Fields",{}).object;
        assert(recordFields==rows->member("Fields",{linaccess::VbaValue::missing()}).object);
        const auto firstField=rows->member("Fields",{linaccess::VbaValue::empty()}).object;
        assert(firstField->member("Name",{}).asText()=="id" && firstField==recordFields->member("Item",{linaccess::VbaValue::numeric(0.5)}).object);
        const auto namedField=recordFields->member("Item",{linaccess::VbaValue::string("QTY")}).object;
        assert(namedField->member("Name",{}).asText()=="qty" && namedField==recordFields->enumerate()[1].object);
        for(const auto& [index,expected]:std::vector<std::pair<linaccess::VbaValue,int>>{{{},94},{linaccess::VbaValue::integer(-1),3265},{linaccess::VbaValue::string("absent"),3265}}) {
            try { (void)rows->member("Fields",{index}); assert(false); }
            catch(const linaccess::VbaRuntimeError& error) { assert(error.number==expected); }
        }
        int duplicateField=0;
        try {
            const auto duplicateRows=database->member("OpenRecordset",{linaccess::VbaValue::string("SELECT id AS same,qty AS same FROM items")}).object;
            assert(duplicateRows->member("RecordCount",{}).asNumber()==2);
            (void)duplicateRows->member("Fields",{linaccess::VbaValue::integer(0)});
        } catch(const linaccess::VbaRuntimeError& error) { duplicateField=error.number; }
        assert(duplicateField==3251);
        rows->member("Edit",{}); rows->assignItem("qty",linaccess::VbaValue::numeric(2.5)); rows->member("Update",{}); rows->member("Close",{});
        assert(q.exec("SELECT qty FROM items WHERE id=1") && q.next() && q.value(0).toDouble()==2.5); q.finish();
        assert(q.exec("UPDATE items SET qty=2 WHERE id=1"));
        assert(q.exec("CREATE TABLE transient_table(value INTEGER)"));
        assert(definitions->member("Count",{}).asNumber()==3);
        auto transient=definitions->item("transient_table").object;
        assert(q.exec("DROP TABLE transient_table"));
        assert(definitions->member("Count",{}).asNumber()==2);
        definitions->member("Refresh",{});
        try { (void)definitions->member("Refresh",{linaccess::VbaValue::integer(1)}); assert(false); }
        catch(const linaccess::VbaRuntimeError& error) { assert(error.number==450); }
        try { (void)transient->member("Name",{}); assert(false); }
        catch(const linaccess::VbaRuntimeError& error) { assert(error.number==3420); }
        assert(q.exec("CREATE TABLE transient_table(value INTEGER)"));
        assert(definitions->item("transient_table").object!=transient);
        assert(q.exec("DROP TABLE transient_table"));
    }
    {
        auto closedDb=QSqlDatabase::addDatabase("QSQLITE","closed-tabledefs"); closedDb.setDatabaseName(":memory:"); assert(closedDb.open());
        QSqlQuery schema(closedDb); assert(schema.exec("CREATE TABLE one_table(value INTEGER)")); schema.finish();
        auto database=linaccess::makeDaoDatabase(closedDb);
        auto definitions=database->member("TableDefs",{}).object;
        auto definition=definitions->item("one_table").object; closedDb.close();
        for (const auto& object:{definitions,definition}) {
            try { (void)object->member(object==definitions ? "Count" : "Name",{}); assert(false); }
            catch(const linaccess::VbaRuntimeError& error) { assert(error.number==3420); }
        }
    }
    {
        const auto database=linaccess::makeDaoDatabase(db);
        auto definition=database->member("CreateQueryDef",{linaccess::VbaValue::string(""),linaccess::VbaValue::string("PARAMETERS [minimum] Long; SELECT id,qty,label INTO [made table] FROM items WHERE id>=[minimum]")}).object;
        definition->member("Parameters",{linaccess::VbaValue::string("minimum")}).object->assign("Value",linaccess::VbaValue::numeric(1));
        assert(db.transaction());
        definition->member("Execute",{});
        assert(definition->member("RecordsAffected",{}).asNumber()==2);
        assert(q.exec("SELECT id,qty,label FROM [made table] ORDER BY id") && q.next());
        assert(q.value(0).toInt()==1 && q.value(1).toDouble()==2 && q.value(2).isNull());
        assert(q.next() && q.value(0).toInt()==2 && q.value(1).toDouble()==3 && q.value(2).toString()=="two"); assert(!q.next()); q.finish();
        bool exists=false;
        try { definition->member("Execute",{}); } catch(const linaccess::VbaRuntimeError& error) { exists=error.number==3010; }
        assert(exists && db.rollback() && !db.tables().contains("made table"));
        database->member("Execute",{linaccess::VbaValue::string("SELECT id,label INTO [empty table] FROM items WHERE id<0")});
        assert(database->member("RecordsAffected",{}).asNumber()==0);
        assert(q.exec("SELECT COUNT(*) FROM [empty table]") && q.next() && q.value(0).toInt()==0); q.finish();
        database->member("Execute",{linaccess::VbaValue::string("DROP TABLE [empty table]")});
        bool external=false;
        try { (void)linaccess::compileAccessQuery("SELECT id INTO [copy] IN 'external.mdb' FROM items"); }
        catch(const linaccess::VbaRuntimeError& error) { external=error.number==3251; }
        assert(external && !db.tables().contains("copy"));
        assert(db.transaction());
        bool failedCreation=false;
        try { database->member("Execute",{linaccess::VbaValue::string("SELECT absent_function(id) AS value INTO [failed table] FROM items")}); }
        catch(const linaccess::VbaRuntimeError& error) { failedCreation=error.number==3085; }
        assert(failedCreation && !db.tables().contains("failed table"));
        assert(q.exec("INSERT INTO items VALUES(3,9,'outer transaction survives')")); q.finish();
        assert(db.rollback());
        assert(q.exec("SELECT COUNT(*) FROM items") && q.next() && q.value(0).toInt()==2); q.finish();
        bool selected=false;
        try { database->member("Execute",{linaccess::VbaValue::string("WITH values_source AS (SELECT 1) SELECT * FROM values_source")}); }
        catch(const linaccess::VbaRuntimeError& error) { selected=error.number==3065; }
        assert(selected);
    }
    for (const auto& mode:std::vector<std::pair<std::string,bool>>{{"Binary",false},{"Text",true},{"Database",true},{"Binary",false}}) {
        const auto result=linaccess::executeVbaProcedure("Option Compare "+mode.first+"\nFunction CompareMode() As Boolean\nWith CurrentDb.TableDefs(\"items\")\nCompareMode=(.Name=\"ITEMS\")\nEnd With\nEnd Function\n","CompareMode",host);
        assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue && result.typedReturnValue->boolean==mode.second);
    }
    const auto withResult=linaccess::executeVbaProcedure(R"VBA(Sub WithRecords()
With OpenOnce()
.Edit
.Fields("qty").Value=2.75
.Update
With .Fields("qty")
Me!NestedValue=.Value
End With
.Close
End With
End Sub
Function OpenOnce() As Object
Me!Opened=CLng(Nz(Me!Opened,0))+1
Set OpenOnce=CurrentDb.OpenRecordset("items")
End Function)VBA","WithRecords",host);
    assert(withResult.status==linaccess::VbaExecutionStatus::Completed);
    assert(host.fields["Opened"]=="1" && host.fields["NestedValue"]=="2.75");
    assert(q.exec("SELECT qty FROM items WHERE id=1") && q.next() && q.value(0).toDouble()==2.75); q.finish();
    const auto withResume=linaccess::executeVbaProcedure(R"VBA(Sub WithResume()
On Error GoTo Repair
With CurrentDb.OpenRecordset("items")
.Edit
.Fields("qty").Value=1/0
.Fields("qty").Value=2.5
.Update
.Close
End With
Exit Sub
Repair:
Resume Next
End Sub)VBA","WithResume",host);
    assert(withResume.status==linaccess::VbaExecutionStatus::Completed);
    assert(q.exec("SELECT qty FROM items WHERE id=1") && q.next() && q.value(0).toDouble()==2.5); q.finish();
    const auto withExit=linaccess::executeVbaProcedure(R"VBA(Sub WithExit()
Dim i As Long
With CurrentDb.OpenRecordset("items")
For i=1 To 3
With .Fields("qty")
If i=1 Then Exit For
End With
Next
.Close
End With
End Sub)VBA","WithExit",host);
    assert(withExit.status==linaccess::VbaExecutionStatus::Completed);
    assert(linaccess::validateVbaProcedure("Sub IllegalWith()\n.Value=1\nEnd Sub","IllegalWith").status!=linaccess::VbaExecutionStatus::Completed);
    assert(q.exec("UPDATE items SET qty=2 WHERE id=1")); q.finish();
    const auto result = linaccess::executeVbaProcedure(R"VBA(
Sub Run()
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim total As Double
Set db = CurrentDb()
Set rs = db.OpenRecordset("items", dbOpenDynaset)
total = 0
Do While Not rs.EOF
    total = total + rs!qty
    rs.Edit
    rs!qty = rs!qty * 2
    rs.Update
    rs.MoveNext
Loop
rs.AddNew
rs!qty = total
rs.Fields("label").Value = "new"
rs.Update
rs.Close
Set rs = Nothing
End Sub
)VBA", "Run", host);
    if (result.status != linaccess::VbaExecutionStatus::Completed) std::cerr << result.explanation << " line " << result.line << '\n';
    assert(result.status == linaccess::VbaExecutionStatus::Completed);
    assert(q.exec("SELECT qty FROM items ORDER BY id"));
    assert(q.next() && q.value(0).toDouble() == 4);
    assert(q.next() && q.value(0).toDouble() == 6);
    assert(q.next() && q.value(0).toDouble() == 5);
    assert(!q.next());
    auto run = [&](const char* body) {
        return linaccess::executeVbaProcedure(std::string("Sub Test()\nDim rs As DAO.Recordset\nSet rs = CurrentDb.OpenRecordset(\"items\")\n") + body + "\nEnd Sub", "Test", host);
    };
    assert(run("Me!DefaultQty=rs(\"qty\")\nMe!DefaultIndex=rs(0)\nMe!DefaultNull=IsNull(rs(\"label\"))\nMe!ImplicitQty=rs.Fields(\"qty\")+1\nrs.Edit\nrs(\"qty\").Value=4.25\nrs.Update\nrs.MoveNext\nMe!DefaultString=(rs.Fields(\"label\")=\"two\")\nrs.Close").status==linaccess::VbaExecutionStatus::Completed);
    assert(host.fields["DefaultQty"]=="4" && host.fields["DefaultIndex"]=="1" && host.fields["DefaultNull"]=="True"
        && host.fields["ImplicitQty"]=="5" && host.fields["DefaultString"]=="True");
    assert(q.exec("SELECT qty FROM items WHERE id=1") && q.next() && q.value(0).toDouble()==4.25); q.finish();
    assert(q.exec("UPDATE items SET qty=4 WHERE id=1")); q.finish();
    assert(run("rs.Edit\nrs!qty = 99\nrs.CancelUpdate\nrs.Requery").status == linaccess::VbaExecutionStatus::Completed);
    assert(run("rs!qty = 99").status == linaccess::VbaExecutionStatus::Error);
    assert(run("rs.Close\nrs.MoveNext").status == linaccess::VbaExecutionStatus::Error);
    assert(run("rs.Delete").status == linaccess::VbaExecutionStatus::Completed);
    assert(q.exec("SELECT COUNT(*) FROM items")); assert(q.next() && q.value(0).toInt() == 2);
    assert(run("rs.AddNew\nrs.Update").status == linaccess::VbaExecutionStatus::Completed);
    assert(q.exec("SELECT label FROM items ORDER BY id DESC LIMIT 1")); assert(q.next() && q.value(0).toString() == "default");
    const auto readonly = linaccess::executeVbaProcedure("Sub Test()\nDim rs As DAO.Recordset\nSet rs = CurrentDb.OpenRecordset(\"SELECT * FROM items\", dbOpenSnapshot)\nrs.Edit\nEnd Sub", "Test", host);
    assert(readonly.status == linaccess::VbaExecutionStatus::Error);
    assert(q.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    q.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES ('query',?,?)");
    q.addBindValue("minimum");
    q.addBindValue(QByteArray("PARAMETERS [Minimum] Long, [Pattern] Text(255); SELECT id,qty,label FROM items WHERE qty >= [Minimum] AND label LIKE [Pattern]"));
    assert(q.exec());
    {
        const auto imported=host.currentDatabase()->member("QueryDefs",{linaccess::VbaValue::string("minimum")}).object;
        int typeError=0; try { (void)imported->member("Type",{}); }
        catch(const linaccess::VbaRuntimeError& failure) { typeError=failure.number; }
        assert(typeError==3251);
    }
    const auto parameterized = linaccess::executeVbaProcedure(R"VBA(
Sub QueryParameters()
Dim definition As DAO.QueryDef
Dim rows As DAO.Recordset
Set definition = CurrentDb.QueryDefs("minimum")
definition.Parameters("minimum").Value = 3
definition.Parameters("pattern").Value = "t*"
Set rows = definition.OpenRecordset(dbOpenSnapshot)
Me!Result = CStr(rows.RecordCount)
rows.Close
Set definition = CurrentDb.CreateQueryDef("", "PARAMETERS [Identifier] Long; UPDATE items SET label = 'changed' WHERE id=[Identifier]")
definition.Parameters(0).Value = 2
definition.Execute(128)
Me!Affected = CStr(definition.RecordsAffected)
End Sub
)VBA", "QueryParameters", host);
    if (parameterized.status != linaccess::VbaExecutionStatus::Completed)
        std::cerr << "QueryDef regression failed at line " << parameterized.line << ": " << parameterized.explanation << '\n';
    assert(parameterized.status == linaccess::VbaExecutionStatus::Completed);
    assert(host.fields["Result"] == "1" && host.fields["Affected"] == "1");
    assert(q.exec("SELECT label FROM items WHERE id=2")); assert(q.next() && q.value(0).toString() == "changed");
    const auto parameterCollection=linaccess::executeVbaProcedure(R"VBA(
Sub CollectionExecution()
Dim definition As DAO.QueryDef
Dim parameters As Object
Dim parameter As DAO.Parameter
Dim first As DAO.Parameter
Dim rows As DAO.Recordset
Dim names As String
Set definition=CurrentDb.CreateQueryDef("", "PARAMETERS [first] Long, [second] Long; SELECT [first]+[second] AS result")
Set parameters=definition.Parameters
Set first=parameters.Item(Index:=0)
If Not (first Is definition.Parameters(Index:="FIRST")) Then Err.Raise 5
Me!ParameterIdentity=CStr(first Is definition.Parameters("FIRST"))
Me!ParameterCount=CStr(parameters.Count)
If VarType(parameters.Count)<>vbInteger Then Err.Raise 5
names=""
For Each parameter In parameters
Let parameter=3
names=names & parameter.Name & ":"
Next
Me!ParameterNames=names
parameters.Refresh
If Not (first Is parameters.Item(Index:="first")) Then Err.Raise 5
If first.Value<>3 Or parameters.Item(Index:="second").Value<>3 Then Err.Raise 5
Set rows=definition.OpenRecordset(dbOpenSnapshot)
Me!CollectionValue=CStr(rows.Fields(0).Value)
rows.Close
definition.SQL="PARAMETERS [replacement] Long; SELECT [replacement] AS result"
parameters.Refresh
Me!ReplacementCount=CStr(parameters.Count)
On Error Resume Next
Me!InvalidName=first.Name
Me!StaleParameter=CStr(Err.Number)
On Error GoTo 0
definition.Close
Err.Clear
On Error Resume Next
Me!InvalidCount=parameters.Count
Me!ClosedParameters=CStr(Err.Number)
Err.Clear
parameters.Refresh
Me!ClosedParameterRefresh=CStr(Err.Number)
End Sub
)VBA","CollectionExecution",host);
    assert(parameterCollection.status==linaccess::VbaExecutionStatus::Completed);
    assert(host.fields["ParameterIdentity"]=="True" && host.fields["ParameterCount"]=="2");
    assert(host.fields["ParameterNames"]=="first:second:" && host.fields["CollectionValue"]=="6");
    assert(host.fields["ReplacementCount"]=="1" && host.fields["StaleParameter"]=="3420" && host.fields["ClosedParameters"]=="3420" && host.fields["ClosedParameterRefresh"]=="3420");
    auto emptyDefinition=host.currentDatabase()->member("CreateQueryDef",{linaccess::VbaValue::string(""),linaccess::VbaValue::string("SELECT 1")}).object;
    auto emptyParameters=emptyDefinition->member("Parameters",{}).object;
    assert(emptyParameters->member("Count",{}).asNumber()==0 && emptyParameters->enumerate().empty());
    emptyParameters->member("Refresh",{});
    int refreshArguments=0; try { emptyParameters->member("Refresh",{linaccess::VbaValue::integer(1)}); }
    catch(const linaccess::VbaRuntimeError& error) { refreshArguments=error.number; } assert(refreshArguments==450);
    bool absentIndex=false;
    try { (void)emptyParameters->member("Item",{linaccess::VbaValue::numeric(0)}); }
    catch(const linaccess::VbaRuntimeError& error) { absentIndex=error.number==3265; }
    assert(absentIndex);
    auto replaced=host.currentDatabase()->member("CreateQueryDef",{linaccess::VbaValue::string(""),linaccess::VbaValue::string("PARAMETERS [same] Long; SELECT [same]")}).object;
    auto collection=replaced->member("Parameters",{}).object;
    auto oldParameter=collection->member("Item",{linaccess::VbaValue::numeric(0)}).object;
    replaced->assign("SQL",linaccess::VbaValue::string("PARAMETERS [same] Text; SELECT [same]"));
    auto newParameter=collection->item("SAME").object;
    assert(newParameter!=oldParameter && collection->member("Count",{}).asNumber()==1);
    bool staleReference=false;
    try { (void)oldParameter->member("Value",{}); }
    catch(const linaccess::VbaRuntimeError& error) { staleReference=error.number==3420; }
    assert(staleReference);
    newParameter->assign("Value",linaccess::VbaValue::string("new value"));
    auto replacementRows=replaced->member("OpenRecordset",{}).object;
    assert(replacementRows->member("Fields",{linaccess::VbaValue::numeric(0)}).object->member("Value",{}).asText()=="new value");
    replacementRows->member("Close",{});
    replaced->member("Close",{});
    bool closedName=false;
    try { (void)newParameter->member("Name",{}); }
    catch(const linaccess::VbaRuntimeError& error) { closedName=error.number==3420; }
    assert(closedName);
    const auto missingParameter = linaccess::executeVbaProcedure(
        "Sub Test()\nDim definition As DAO.QueryDef\nDim rows As DAO.Recordset\nSet definition=CurrentDb.QueryDefs(\"minimum\")\nSet rows=definition.OpenRecordset()\nEnd Sub", "Test", host);
    assert(missingParameter.status == linaccess::VbaExecutionStatus::Error && missingParameter.errorNumber == 3061);
    const auto storeQuery = [&](const QString& name, const QByteArray& sql) {
        q.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES ('query',?,?)");
        q.addBindValue(name); q.addBindValue(sql); assert(q.exec());
    };
    storeQuery("base", "SELECT * FROM items");
    storeQuery("child", "SELECT * FROM base WHERE qty >= 3");
    assert(q.exec("CREATE TEMP VIEW user_view AS SELECT 42 AS answer"));
    storeQuery("user_view", "SELECT 99 AS answer");
    assert(linaccess::installAccessQueryViews(db) >= 2);
    assert(q.exec("SELECT COUNT(*) FROM child")); assert(q.next() && q.value(0).toInt() == 2);
    q.finish();
    storeQuery("BASE", "SELECT * FROM items WHERE id=2");
    assert(linaccess::installAccessQueryViews(db) >= 2);
    assert(q.exec("SELECT COUNT(*) FROM child")); assert(q.next() && q.value(0).toInt() == 1);
    assert(q.exec("SELECT answer FROM user_view")); assert(q.next() && q.value(0).toInt() == 42);
    q.finish();
    assert(q.exec("CREATE TABLE ranking(id INTEGER PRIMARY KEY, category TEXT, score INTEGER)"));
    assert(q.exec("INSERT INTO ranking VALUES (1,'A',10),(2,'A',10),(3,'B',5),(4,'B',5)"));
    const auto native = linaccess::makeDaoDatabase(db);
    const auto countRows = [&](const char* sql) {
        auto rows = native->member("OpenRecordset",{linaccess::VbaValue::string(sql)});
        return rows.object->member("RecordCount",{}).asNumber();
    };
    assert(countRows("SELECT DISTINCTROW category FROM ranking") == 4);
    assert(countRows("SELECT DISTINCT category FROM ranking") == 2);
    assert(countRows("SELECT TOP 1 category FROM ranking ORDER BY score DESC") == 2);
    assert(countRows("SELECT TOP 3 id FROM ranking ORDER BY score DESC") == 4);
    assert(countRows("SELECT TOP 1 id FROM ranking ORDER BY score DESC,id ASC") == 1);
    assert(countRows("SELECT TOP 2 id FROM ranking") == 2);
    auto aliasRows = native->member("OpenRecordset",{linaccess::VbaValue::string("SELECT TOP 1 -score AS score FROM ranking ORDER BY score DESC")});
    assert(aliasRows.object->member("RecordCount",{}).asNumber() == 2);
    assert(aliasRows.object->member("Fields",{linaccess::VbaValue::numeric(0)}).object->member("Value",{}).asNumber() == -5);
    bool rejectedOrderedDistinct = false;
    try { (void)linaccess::compileAccessQuery("SELECT DISTINCT TOP 2 category FROM ranking ORDER BY category"); }
    catch (...) { rejectedOrderedDistinct = true; }
    assert(rejectedOrderedDistinct);
    bool rejectedOrdinal = false;
    try { (void)linaccess::compileAccessQuery("SELECT TOP 1 score FROM ranking ORDER BY 1"); }
    catch (...) { rejectedOrdinal = true; }
    assert(rejectedOrdinal);
    QString expensive = "SELECT 1 AS n UNION ALL SELECT count(*) FROM ";
    for (int i = 0; i < 12; ++i) expensive += (i ? ",ranking r" : "ranking r") + QString::number(i);
    bool rejectedPartialResult = false;
    try { (void)native->member("OpenRecordset",{linaccess::VbaValue::string(expensive.toStdString())}); }
    catch (const linaccess::VbaRuntimeError& error) { rejectedPartialResult = error.number == 3035; }
    assert(rejectedPartialResult);
    bool rejectedExecuteSelect = false;
    try { (void)native->member("Execute",{linaccess::VbaValue::string("SELECT 1")}); }
    catch (const linaccess::VbaRuntimeError& error) { rejectedExecuteSelect = error.number == 3065; }
    assert(rejectedExecuteSelect);
    bool rejectedDistinctRowJoin = false;
    try { (void)linaccess::compileAccessQuery("SELECT DISTINCTROW a.id FROM ranking a JOIN ranking b ON a.id=b.id"); }
    catch (...) { rejectedDistinctRowJoin = true; }
    assert(rejectedDistinctRowJoin);
    const auto mixed = linaccess::compileAccessQuery("PARAMETERS p Text; SELECT p, 'p', \"p\" WHERE 'a%' LIKE p");
    const auto bindings = linaccess::bindAccessQueryParameters(mixed,{{"p","a%"}});
    assert(bindings.value(":la_parameter_0").toString() == "a%");
    assert(bindings.value(":la_parameter_0_like").toString() == "a\\%");
    const auto datePlan = linaccess::compileAccessQuery("SELECT #12/31/2003# = #2003-12-31#, Date(), Now()");
    assert(q.exec(datePlan.sql)); assert(q.next() && q.value(0).toInt() == 1);
    assert(q.value(1).toString().endsWith("00:00:00"));
    assert(q.value(2).toString().size() == 19);
    const auto dateParameters = linaccess::compileAccessQuery("PARAMETERS d DateTime; SELECT d");
    const auto dateValues = linaccess::bindAccessQueryParameters(dateParameters,{{"d","2003-12-31"}});
    assert(dateValues.value(":la_parameter_0").toString() == "2003-12-31 00:00:00");
    bool ambiguousDate = false;
    try { (void)linaccess::compileAccessQuery("SELECT #1/2/03#"); } catch (...) { ambiguousDate = true; }
    assert(ambiguousDate);
    {
        linaccess::SqliteExecutionBudget budget(db,100000);
        assert(!q.exec("WITH RECURSIVE forever(x) AS (VALUES(1) UNION ALL SELECT x+1 FROM forever) SELECT sum(x) FROM forever"));
    }
    assert(q.exec("SELECT 1")); assert(q.next() && q.value(0).toInt() == 1);
    const auto boolValue = linaccess::qtFieldValue(QVariant(-1),"BOOLEAN");
    assert(boolValue.kind == linaccess::VbaValue::Kind::Boolean && boolValue.boolean);
    const auto textValue = linaccess::sqlFieldValue(linaccess::VbaValue::numeric(3),"TEXT");
    assert(textValue.metaType().id() == QMetaType::QString && textValue.toString() == "3");
    assert(linaccess::sqlFieldValue(linaccess::VbaValue::string("2003-12-31"),"DATETIME").toString() == "2003-12-31 00:00:00");
    bool refusedLoss = false;
    try { (void)linaccess::qtFieldValue(QVariant(qlonglong(9007199254740993LL)),"INTEGER"); }
    catch (const linaccess::VbaRuntimeError& error) { refusedLoss = error.number == 6; }
    assert(refusedLoss);

    assert(q.exec("CREATE TABLE enumeration_values(a INTEGER,b INTEGER)"));
    assert(q.exec("INSERT INTO enumeration_values VALUES(2,3)"));
    const auto enumerated = linaccess::executeVbaProcedure(R"VBA(
Sub Enumerate()
Dim rs As DAO.Recordset
Dim fld As DAO.Field
Dim i As Long
Dim visits As Long
Set rs = CurrentDb.OpenRecordset("enumeration_values")
Me!Count = rs.Fields.Count
Me!First = rs.Fields.Item(0).Name
For i = 1 To 2
For Each fld In rs.Fields
visits = visits + 1
Next fld
Next i
Me!Visits = visits
rs.Edit
For Each fld In rs.Fields
fld.Value = fld.Value * 2
Next
rs.Update
rs.Bookmark=rs.LastModified
visits = 0
For Each fld In rs.Fields
If fld.Name = "a" Then
visits = visits + 1
Exit For
End If
visits = 100
Next
Me!Exited = visits
rs.Close
End Sub
)VBA","Enumerate",host);
    if (enumerated.status != linaccess::VbaExecutionStatus::Completed) std::cerr << enumerated.explanation << " at " << enumerated.line << "\n";
    assert(enumerated.status == linaccess::VbaExecutionStatus::Completed);
    assert(host.fields["Count"] == "2" && host.fields["First"] == "a");
    assert(host.fields["Visits"] == "4" && host.fields["Exited"] == "1");
    assert(q.exec("SELECT a,b FROM enumeration_values") && q.next());
    assert(q.value(0).toInt()==4 && q.value(1).toInt()==6); q.finish();
    const auto enumerationError = linaccess::executeVbaProcedure(R"VBA(
Sub Recover()
Dim rs As DAO.Recordset
Dim fld As DAO.Field
Dim count As Long
On Error GoTo Failed
Set rs = CurrentDb.OpenRecordset("enumeration_values")
For Each fld In rs.Fields
Err.Raise 11
count = count + 1
Next fld
Me!Recovered = count
Exit Sub
Failed:
Resume Next
End Sub
)VBA","Recover",host);
    assert(enumerationError.status == linaccess::VbaExecutionStatus::Completed);
    assert(host.fields["Recovered"] == "2");
    const auto wrongCollection = linaccess::executeVbaProcedure(
        "Sub Bad()\nDim x\nFor Each x In CurrentDb\nNext\nEnd Sub","Bad",host);
    assert(wrongCollection.status == linaccess::VbaExecutionStatus::Error && wrongCollection.errorNumber == 438);
    assert(linaccess::validateVbaProcedure("Sub Bad()\nExit For\nEnd Sub","Bad").status == linaccess::VbaExecutionStatus::Error);

    assert(q.exec("CREATE TABLE property_collision(EOF INTEGER,Fields TEXT,Type INTEGER)"));
    assert(q.exec("INSERT INTO property_collision VALUES(9,'field',7)"));
    const auto collision = linaccess::executeVbaProcedure(R"VBA(
Sub Collision()
Dim rs As DAO.Recordset
Dim fld As DAO.Field
Set rs = CurrentDb.OpenRecordset("property_collision")
Me!EofField = rs!EOF
Me!EofProperty = rs.EOF
Me!FieldsField = rs!Fields
Me!FieldsCount = rs.Fields.Count
Me!CursorType = rs.Type
Me!CursorTypeKind = VarType(rs.Type)
Me!TypeField = rs!Type
rs.Edit
Set fld = rs!EOF
fld.Value = 10
Me!FieldReference = rs!EOF.Value
rs!Fields.Value = "changed"
rs!Type.Value = 11
rs.Update
rs.Close
End Sub
)VBA","Collision",host);
    assert(collision.status == linaccess::VbaExecutionStatus::Completed);
    assert(host.fields["EofField"]=="9" && host.fields["EofProperty"]=="False");
    assert(host.fields["FieldsField"]=="field" && host.fields["FieldsCount"]=="3");
    assert(host.fields["FieldReference"]=="10");
    assert(host.fields["CursorType"]=="1" && host.fields["CursorTypeKind"]=="2" && host.fields["TypeField"]=="7");
    assert(q.exec("SELECT Type FROM property_collision") && q.next() && q.value(0).toInt()==11); q.finish();
    {
        using linaccess::VbaValue;
        auto database=host.currentDatabase();
        auto dynaset=database->member("OpenRecordset",{VbaValue::string("property_collision"),VbaValue::integer(2)}).object;
        auto clone=dynaset->member("Clone",{}).object;
        assert(dynaset->member("Type",{}).asNumber()==2 && clone->member("Type",{}).asNumber()==2);
        auto snapshot=database->member("OpenRecordset",{VbaValue::string("SELECT * FROM property_collision"),VbaValue::integer(4)}).object;
        assert(snapshot->member("Type",{}).asNumber()==4 && !snapshot->member("Updatable",{}).boolean);
        auto readonlyDynaset=database->member("OpenRecordset",{VbaValue::string("SELECT * FROM property_collision"),VbaValue::integer(2)}).object;
        assert(readonlyDynaset->member("Type",{}).asNumber()==2 && !readonlyDynaset->member("Updatable",{}).boolean);
        auto definition=database->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string("SELECT * FROM property_collision")}).object;
        auto queryDynaset=definition->member("OpenRecordset",{VbaValue::integer(2)}).object;
        auto querySnapshot=definition->member("OpenRecordset",{VbaValue::integer(4)}).object;
        assert(queryDynaset->member("Type",{}).asNumber()==2 && querySnapshot->member("Type",{}).asNumber()==4);
        for(const auto& kind:std::vector<VbaValue>{VbaValue::numeric(2.5),VbaValue::string("2"),VbaValue::missing()}) {
            auto converted=database->member("OpenRecordset",{VbaValue::string("property_collision"),kind}).object;
            assert(converted->member("Type",{}).asNumber()==(kind.kind==VbaValue::Kind::Missing ? 1 : 2)); converted->member("Close",{});
            auto parameterized=definition->member("OpenRecordset",{kind}).object;
            assert(parameterized->member("Type",{}).asNumber()==2); parameterized->member("Close",{});
        }
        auto rounded=definition->member("OpenRecordset",{VbaValue::numeric(3.5)}).object;
        assert(rounded->member("Type",{}).asNumber()==4); rounded->member("Close",{});
        for(const auto& invalid:std::vector<std::pair<VbaValue,int>>{{VbaValue::numeric(2.6),3251},{VbaValue::empty(),3251},{VbaValue{},94},{VbaValue::string("invalid"),13}}) {
            int problem=0;
            try { (void)database->member("OpenRecordset",{VbaValue::string("property_collision"),invalid.first}); }
            catch(const linaccess::VbaRuntimeError& failure) { problem=failure.number; }
            assert(problem==invalid.second);
        }


        int error=0; try { dynaset->assign("Type",VbaValue::integer(4)); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
        assert(error==383 && dynaset->member("Type",{}).asNumber()==2);
        error=0; try { (void)dynaset->member("Type",{VbaValue::integer(0)}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==450);
        clone->member("Close",{}); dynaset->member("Close",{});
        error=0; try { (void)dynaset->member("Type",{}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==3420);
        auto frozen=linaccess::makeDaoSnapshot(db,{"id"},{},{{QVariant(1)}});
        error=0; try { (void)frozen->member("Type",{}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==3251);
    }

    assert(q.exec("SELECT EOF FROM property_collision") && q.next() && q.value(0).toInt()==10); q.finish();
    assert(q.exec("SELECT Fields FROM property_collision") && q.next() && q.value(0).toString()=="changed"); q.finish();
    std::cout << "DAO recordset iteration, edit, insert, fields, cancel, delete and read-only checks passed\n";
}

void testDaoTransactions() {
    using linaccess::VbaValue;
    QTemporaryDir directory; assert(directory.isValid());
    auto db=QSqlDatabase::addDatabase("QSQLITE","dao-workspace-writer"); db.setDatabaseName(directory.filePath("project.linaccess")); assert(db.open());
    QSqlQuery setup(db); assert(setup.exec("CREATE TABLE balances(id INTEGER PRIMARY KEY,amount INTEGER NOT NULL)"));
    assert(setup.exec("INSERT INTO balances VALUES(1,100),(2,200)")); setup.finish();
    auto reader=QSqlDatabase::addDatabase("QSQLITE","dao-workspace-reader"); reader.setDatabaseName(db.databaseName()); assert(reader.open());
    const auto amount=[&](QSqlDatabase source,int id=1) { QSqlQuery q(source); q.prepare("SELECT amount FROM balances WHERE id=?"); q.addBindValue(id); assert(q.exec() && q.next()); return q.value(0).toInt(); };
    const auto error=[&](const std::function<void()>& operation,int expected) { int number=0; try { operation(); } catch(const linaccess::VbaRuntimeError& failure) { number=failure.number; } assert(number==expected); };
    auto engine=linaccess::makeDaoEngine(db),database=linaccess::makeDaoDatabase(db);
    auto spaces=engine->member("Workspaces",{}).object;
    auto workspace=engine->defaultMember({VbaValue::numeric(0)},true).object;
    assert(engine==linaccess::makeDaoEngine(db));
    assert(workspace==spaces->member("Item",{VbaValue::string("#Default Workspace#")}).object);
    assert(spaces->member("Count",{}).asNumber()==1 && spaces->enumerate().size()==1 && spaces->member("Count",{}).numericType==VbaValue::NumericType::Integer);
    auto databases=workspace->member("Databases",{}).object;
    assert(databases->member("Count",{}).asNumber()==1 && databases->enumerate().size()==1 && databases->member("Count",{}).numericType==VbaValue::NumericType::Integer);
    const auto defaultDatabase=workspace->defaultMember({VbaValue::numeric(0)},true).object;
    assert(defaultDatabase->member("Transactions",{}).truthy());
    assert(defaultDatabase==workspace->defaultMember({VbaValue::numeric(0)},true).object);
    error([&] { engine->defaultMember({VbaValue::numeric(1)},true); },3265);
    error([&] { workspace->member("CommitTrans",{}); },3034);
    error([&] { workspace->member("Rollback",{}); },3034);
    workspace->member("BeginTrans",{});
    database->member("Execute",{VbaValue::string("UPDATE balances SET amount=90 WHERE id=1")});
    database->member("Execute",{VbaValue::string("UPDATE balances SET amount=210 WHERE id=2")});
    assert(amount(db)==90 && amount(reader)==100);
    error([&] { workspace->member("CommitTrans",{VbaValue{}}); },94);
    error([&] { workspace->member("CommitTrans",{VbaValue::numeric(2)}); },5);
    error([&] { workspace->member("CommitTrans",{VbaValue::numeric(1)}); },3251);
    assert(amount(reader)==100);
    workspace->member("CommitTrans",{VbaValue::numeric(0)}); assert(amount(reader)==90 && amount(reader,2)==210);
    auto rows=database->member("OpenRecordset",{VbaValue::string("balances")}).object;
    const auto bookmark=rows->member("Bookmark",{});
    auto field=rows->member("Fields",{VbaValue::string("amount")}).object;
    auto snapshot=linaccess::makeDaoSnapshot(db,{"id","amount"},{{"id","INTEGER"},{"amount","INTEGER"}},{{1,90},{2,210}},{1,2});
    workspace->member("BeginTrans",{}); rows->member("Edit",{}); rows->assignItem("amount",VbaValue::numeric(7)); rows->member("Update",{});
    assert(amount(db)==7 && amount(reader)==90);
    workspace->member("Rollback",{}); assert(amount(db)==90 && amount(reader)==90);
    error([&] { field->member("Value",{}); },3420); error([&] { rows->member("EOF",{}); },3420);
    error([&] { linaccess::refreshDaoSnapshot(snapshot,{"bad"},{},{{1}},{1}); },3251);
    error([&] { snapshot->member("EOF",{}); },3420);
    linaccess::refreshDaoSnapshot(snapshot,{"id","amount"},{{"id","INTEGER"},{"amount","INTEGER"}},{{1,90},{2,210}},{1,2});
    assert(snapshot->item("amount").asNumber()==90);
    rows->member("Requery",{}); assert(rows->item("amount").asNumber()==90);
    error([&] { rows->assign("Bookmark",bookmark); },3159);
    workspace->member("BeginTrans",{}); database->member("Execute",{VbaValue::string("INSERT INTO balances VALUES(3,300)")});
    workspace->member("BeginTrans",{}); database->member("Execute",{VbaValue::string("UPDATE balances SET amount=8 WHERE id=1")});
    error([&] { workspace->member("Rollback",{}); },3251); assert(amount(db)==8 && amount(reader)==90);
    workspace->member("CommitTrans",{}); workspace->member("Rollback",{});
    QSqlQuery count(db); assert(count.exec("SELECT COUNT(*) FROM balances") && count.next() && count.value(0).toInt()==2); count.finish();
    assert(amount(db)==90);
    for(int depth=0;depth<5;++depth) workspace->member("BeginTrans",{});
    error([&] { workspace->member("BeginTrans",{}); },3003);
    for(int depth=0;depth<4;++depth) workspace->member("CommitTrans",{});
    workspace->member("Rollback",{});
    rows->member("Requery",{}); auto clone=rows->member("Clone",{}).object;
    error([&] { workspace->member("BeginTrans",{}); },3251); clone->member("Close",{});
    workspace->member("BeginTrans",{}); error([&] { rows->member("Clone",{}); },3251); workspace->member("Rollback",{});
    Host host(db);
    const auto result=linaccess::executeVbaProcedure(R"VBA(Sub Restore(ByVal session As DAO.Workspace)
session.BeginTrans
CurrentDb.Execute "UPDATE balances SET amount=999 WHERE id=1",dbFailOnError
session.Rollback
End Sub
Sub Transfer()
Dim eng As DAO.DBEngine
Dim spaces As DAO.Workspaces
Dim session As DAO.Workspace
Set eng=DBEngine
Set spaces=eng.Workspaces
Set session=spaces(0)
session.BeginTrans
CurrentDb.Execute "UPDATE balances SET amount=amount-10 WHERE id=1",dbFailOnError
CurrentDb.Execute "UPDATE balances SET amount=amount+10 WHERE id=2",dbFailOnError
session.CommitTrans
Call Restore(DBEngine(0))
End Sub)VBA","Transfer",host);
    if (result.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Workspace VBA status="<<static_cast<int>(result.status)<<", error="<<result.errorNumber<<", line="<<result.line<<"\n";
    assert(result.status==linaccess::VbaExecutionStatus::Completed && amount(reader)==80 && amount(reader,2)==220);
    engine->member("BeginTrans",{}); database->member("Execute",{VbaValue::string("UPDATE balances SET amount=5 WHERE id=1")});
    db.close(); assert(amount(reader)==80);
    error([&] { workspace->member("Name",{}); },3420); error([&] { database->member("Name",{}); },3420); error([&] { spaces->member("Count",{}); },3420);
    assert(db.open()); auto reopened=linaccess::makeDaoEngine(db); assert(reopened!=engine);
    reopened->member("BeginTrans",{}); linaccess::makeDaoDatabase(db)->member("Execute",{VbaValue::string("UPDATE balances SET amount=70 WHERE id=1")});
    reopened->member("CommitTrans",{}); assert(amount(reader)==70);
    reader.close(); db.close();
}
