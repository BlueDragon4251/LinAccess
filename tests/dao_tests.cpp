#include "dao_database.hpp"
#include "vba_qt_value.hpp"
#include "access_query.hpp"
#include "sqlite_execution.hpp"
#include "vba_runtime.hpp"
#include <QCoreApplication>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>
#include <map>

class Host final : public linaccess::VbaHost {
public:
    explicit Host(QSqlDatabase db) : db_(db) {}
    std::shared_ptr<linaccess::VbaObject> currentDatabase() override { return linaccess::makeDaoDatabase(db_); }
    std::map<std::string,std::optional<std::string>> fields;
    std::optional<std::string> field(std::string_view name) override { return fields[std::string(name)]; }
    void setField(std::string_view name, std::optional<std::string> value) override { fields[std::string(name)] = std::move(value); }
    void message(std::string_view) override { assert(false); }
    void saveRecord() override { assert(false); }
private:
    QSqlDatabase db_;
};
int main(int argc, char** argv) {
    QCoreApplication app(argc, argv);
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
        assert(definitions->member("Count",{}).asNumber()==2);
        assert(q.exec("CREATE TABLE _la_private_metadata(secret TEXT)"));
        assert(q.exec("CREATE VIEW visible_query AS SELECT * FROM items"));
        assert(definitions->member("Count",{}).asNumber()==2);
        const auto table=database->member("TableDefs",{linaccess::VbaValue::string("ITEMS")}).object;
        assert(table==definitions->member("Item",{linaccess::VbaValue::numeric(0)}).object);
        assert(table==definitions->item("items").object);
        assert(table->member("Name",{}).asText()=="items");
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
Set first=parameters.Item(0)
Me!ParameterIdentity=CStr(first Is definition.Parameters("FIRST"))
Me!ParameterCount=CStr(parameters.Count)
names=""
For Each parameter In parameters
parameter.Value=3
names=names & parameter.Name & ":"
Next
Me!ParameterNames=names
Set rows=definition.OpenRecordset(dbOpenSnapshot)
Me!CollectionValue=CStr(rows.Fields(0).Value)
rows.Close
definition.SQL="PARAMETERS [replacement] Long; SELECT [replacement] AS result"
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
End Sub
)VBA","CollectionExecution",host);
    assert(parameterCollection.status==linaccess::VbaExecutionStatus::Completed);
    assert(host.fields["ParameterIdentity"]=="True" && host.fields["ParameterCount"]=="2");
    assert(host.fields["ParameterNames"]=="first:second:" && host.fields["CollectionValue"]=="6");
    assert(host.fields["ReplacementCount"]=="1" && host.fields["StaleParameter"]=="3420" && host.fields["ClosedParameters"]=="3420");
    auto emptyDefinition=host.currentDatabase()->member("CreateQueryDef",{linaccess::VbaValue::string(""),linaccess::VbaValue::string("SELECT 1")}).object;
    auto emptyParameters=emptyDefinition->member("Parameters",{}).object;
    assert(emptyParameters->member("Count",{}).asNumber()==0 && emptyParameters->enumerate().empty());
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

    assert(q.exec("CREATE TABLE property_collision(EOF INTEGER,Fields TEXT)"));
    assert(q.exec("INSERT INTO property_collision VALUES(9,'field')"));
    const auto collision = linaccess::executeVbaProcedure(R"VBA(
Sub Collision()
Dim rs As DAO.Recordset
Dim fld As DAO.Field
Set rs = CurrentDb.OpenRecordset("property_collision")
Me!EofField = rs!EOF
Me!EofProperty = rs.EOF
Me!FieldsField = rs!Fields
Me!FieldsCount = rs.Fields.Count
rs.Edit
Set fld = rs!EOF
fld.Value = 10
Me!FieldReference = rs!EOF.Value
rs!Fields.Value = "changed"
rs.Update
rs.Close
End Sub
)VBA","Collision",host);
    assert(collision.status == linaccess::VbaExecutionStatus::Completed);
    assert(host.fields["EofField"]=="9" && host.fields["EofProperty"]=="False");
    assert(host.fields["FieldsField"]=="field" && host.fields["FieldsCount"]=="2");
    assert(host.fields["FieldReference"]=="10");
    assert(q.exec("SELECT EOF FROM property_collision") && q.next() && q.value(0).toInt()==10); q.finish();
    assert(q.exec("SELECT Fields FROM property_collision") && q.next() && q.value(0).toString()=="changed"); q.finish();
    std::cout << "DAO recordset iteration, edit, insert, fields, cancel, delete and read-only checks passed\n";
}
