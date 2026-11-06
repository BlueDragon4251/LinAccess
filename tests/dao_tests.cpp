#include "dao_database.hpp"
#include "vba_array.hpp"
#include "vba_qt_value.hpp"
#include "access_query.hpp"
#include "sqlite_execution.hpp"
#include "jet_sql_functions.hpp"
#include "vba_module_context.hpp"
#include "vba_runtime.hpp"
#include <QCoreApplication>
#include <QSqlQuery>
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
If definitions.Count<>2 Then Err.Raise 5
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
rows.Close
On Error Resume Next
payload=held.Value
Me!BinaryClosedError=CStr(Err.Number)
Err.Clear
payload=held.GetChunk(0,1)
Me!BinaryChunkClosedError=CStr(Err.Number)
End Sub
)VBA","BinaryFields",host);
    if(binaryFields.status!=linaccess::VbaExecutionStatus::Completed) std::cerr<<"Synthetic binary fields: "<<binaryFields.errorNumber<<" line "<<binaryFields.line<<" "<<binaryFields.explanation<<'\n';
    assert(binaryFields.status==linaccess::VbaExecutionStatus::Completed && host.fields["BinaryClosedError"]=="3420" && host.fields["BinarySizeWriteError"]=="383" && host.fields["BinaryChunkClosedError"]=="3420");
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
    linaccess::installJetSqlFunctions(db);
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

int main(int argc, char** argv) {
    QCoreApplication app(argc, argv);
    testVariantArraysDao();
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
        assert(definitions->member("Count",{}).asNumber()==2);
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
    assert(host.fields["CursorType"]=="2" && host.fields["CursorTypeKind"]=="2" && host.fields["TypeField"]=="7");
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
            assert(converted->member("Type",{}).asNumber()==2); converted->member("Close",{});
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
    assert(spaces->member("Count",{}).asNumber()==1 && spaces->enumerate().size()==1);
    auto databases=workspace->member("Databases",{}).object;
    assert(databases->member("Count",{}).asNumber()==1 && databases->enumerate().size()==1);
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
