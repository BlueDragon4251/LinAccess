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
    Host host(db);
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
Set rs = CurrentDb.OpenRecordset("property_collision")
Me!EofField = rs!EOF
Me!EofProperty = rs.EOF
Me!FieldsField = rs!Fields
Me!FieldsCount = rs.Fields.Count
rs.Edit
rs!EOF = 10
rs.Update
rs.Close
End Sub
)VBA","Collision",host);
    assert(collision.status == linaccess::VbaExecutionStatus::Completed);
    assert(host.fields["EofField"]=="9" && host.fields["EofProperty"]=="False");
    assert(host.fields["FieldsField"]=="field" && host.fields["FieldsCount"]=="2");
    assert(q.exec("SELECT EOF FROM property_collision") && q.next() && q.value(0).toInt()==10); q.finish();
    std::cout << "DAO recordset iteration, edit, insert, fields, cancel, delete and read-only checks passed\n";
}
