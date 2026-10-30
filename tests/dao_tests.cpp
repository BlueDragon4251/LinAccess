#include "dao_database.hpp"
#include "vba_runtime.hpp"
#include <QCoreApplication>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>

class Host final : public linaccess::VbaHost {
public:
    explicit Host(QSqlDatabase db) : db_(db) {}
    std::shared_ptr<linaccess::VbaObject> currentDatabase() override { return linaccess::makeDaoDatabase(db_); }
    std::optional<std::string> field(std::string_view) override { return {}; }
    void setField(std::string_view, std::optional<std::string>) override { assert(false); }
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
    std::cout << "DAO recordset iteration, edit, insert, fields, cancel, delete and read-only checks passed\n";
}
