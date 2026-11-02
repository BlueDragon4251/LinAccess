#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_table_import.hpp"
#include "dao_database.hpp"
#include "vba_currency.hpp"
#include <QApplication>
#include <QLineEdit>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

int main(int argc,char** argv) {
    QApplication app(argc,argv);
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","currency-updates"); db.setDatabaseName(":memory:"); assert(db.open()); QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)")); assert(sql.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    const auto schema=linaccess::parseMdbSchema("CREATE TABLE `cash`\n(\n`id` INTEGER,\n`a` REAL,\n`b` REAL\n);\n-- LINACCESS_SOURCE_TYPE 6964 INT\n-- LINACCESS_SOURCE_TYPE 61 CURRENCY\n-- LINACCESS_SOURCE_TYPE 62 CURRENCY\n");
    linaccess::importAccessTable(db,"synthetic-updates","cash",schema,linaccess::parseCsvDetailed("id,a,b\n1,0.1,0.2\n2,1.2345,0.0001\n"));
    auto database=linaccess::makeDaoDatabase(db);
    const auto query=[&](const std::string& source) { return database->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string(source)}).object; };
    const auto execute=[&](const std::string& source) { auto definition=query(source); definition->member("Execute",{}); return definition->member("RecordsAffected",{}).asNumber(); };
    const auto stored=[&](int id,const QString& column="a") { assert(sql.exec("SELECT "+column+" FROM cash WHERE id="+QString::number(id)) && sql.next()); const auto value=sql.value(0); sql.finish(); return value; };
    assert(execute("UPDATE cash SET a=a+b WHERE id=1")==1);
    assert(stored(1).toDouble()==0.3);
    execute("UPDATE cash SET a='1.23455' WHERE id=1"); assert(linaccess::currencyFromRealStorage(stored(1).toDouble()).currencyUnits==12346);
    execute("UPDATE cash SET a=NULL WHERE id=1"); assert(stored(1).isNull());
    auto parameterized=query("PARAMETERS p Text(255); UPDATE cash SET a=p WHERE id=1"); parameterized->member("Parameters",{VbaValue::string("p")}).object->assign("Value",VbaValue::string("0.00005")); parameterized->member("Execute",{}); assert(stored(1).toDouble()==0);
    auto money=query("PARAMETERS p Currency; UPDATE cash SET a=a+p WHERE id=1"); money->member("Parameters",{VbaValue::string("p")}).object->assign("Value",VbaValue::currency(2000)); money->member("Execute",{}); assert(stored(1).toDouble()==0.2);
    execute("UPDATE cash SET a='0.1' WHERE id=1"); execute("UPDATE cash AS c SET a=c.a+c.b WHERE id=1"); assert(stored(1).toDouble()==0.3);
    execute("UPDATE cash SET a=b,b=a WHERE id=1"); assert(stored(1).toDouble()==0.2 && stored(1,"b").toDouble()==0.3);
    for(const auto& [bad,wanted]:std::vector<std::pair<std::string,int>>{{"invalid",13},{"549755813888",3251}}) {
        int error=0; try { execute("UPDATE cash SET a=IIf(id=1,'2.34565','"+bad+"') WHERE id IN(1,2)"); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
        assert(error==wanted && stored(1).toDouble()==0.2 && stored(2).toDouble()==1.2345);
    }
    auto ordered=query("PARAMETERS first Text(255), second Long; UPDATE cash SET a=first WHERE id=second");
    ordered->member("Parameters",{VbaValue::string("first")}).object->assign("Value",VbaValue::string("3.45675"));
    ordered->member("Parameters",{VbaValue::string("second")}).object->assign("Value",VbaValue::longInteger(2)); ordered->member("Execute",{});
    assert(stored(1).toDouble()==0.2 && linaccess::currencyFromRealStorage(stored(2).toDouble()).currencyUnits==34568);
    execute("UPDATE [cash] AS [c] SET [c].[a]=[c].[a]+[c].[b], [id]=[id]+10 WHERE [c].[id]=2;");
    assert(linaccess::currencyFromRealStorage(stored(12).toDouble()).currencyUnits==34569);
    execute("UPDATE cash SET id=2,a='1.2345' WHERE id=12");
    execute("UPDATE cash SET a=a+0.00006 WHERE id=1");
    assert(linaccess::currencyFromRealStorage(stored(1).toDouble()).currencyUnits==2001);
    assert(sql.exec("CREATE TABLE ordinary(a REAL,b REAL)")); assert(sql.exec("INSERT INTO ordinary VALUES(0.1,0.2)")); execute("UPDATE ordinary SET a=a+b");
    assert(sql.exec("SELECT a FROM ordinary") && sql.next() && sql.value(0).toDouble()==0.1+0.2); sql.finish();
    const QByteArray layout="Begin Form\nRecordSource =\"cash\"\nBegin Section\nBegin TextBox\nName =\"a\"\nControlSource =\"a\"\nWidth =5000\nEnd\nEnd\nEnd\n";
    const QByteArray code=R"VBA(Function UpdateCurrency(ByVal text As String)
Dim q As DAO.QueryDef
Set q=CurrentDb.CreateQueryDef("","PARAMETERS p Text(255); UPDATE cash SET a=p WHERE id=1")
q.Parameters("p").Value=text
q.Execute
Me.Requery
UpdateCurrency=VarType(Me!a)
End Function
)VBA";
    AccessFormPreview form("currency_updates",layout,code,db); form.show(); QApplication::processEvents(); auto model=makePreviewFormObject(form);
    std::vector<VbaValue> args{VbaValue::string("1.23455")}; const auto updated=form.executeProcedure("UpdateCurrency",args);
    assert(updated.status==linaccess::VbaExecutionStatus::Completed && updated.typedReturnValue->number==6 && model->item("a").currencyUnits==12346 && !form.hasUnsavedChanges());
    args[0]=VbaValue::string("549755813887.9999"); assert(form.executeProcedure("UpdateCurrency",args).status==linaccess::VbaExecutionStatus::Completed);
    assert(form.findChild<QLineEdit*>("access-control-a")->text()=="549755813887.9999" && model->item("a").currencyUnits==5497558138879999LL); assert(form.close());
}
