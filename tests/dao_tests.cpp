#include "dao_database.hpp"
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
    std::cout << "DAO recordset iteration, edit, insert, fields, cancel, delete and read-only checks passed\n";
}
