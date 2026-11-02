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
    QApplication app(argc,argv); using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","currency-inserts"); db.setDatabaseName(":memory:"); assert(db.open()); QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)")); assert(sql.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    const auto schema=linaccess::parseMdbSchema("CREATE TABLE `cash`\n(\n`id` INTEGER PRIMARY KEY,\n`a` REAL,\n`b` REAL\n);\n-- LINACCESS_SOURCE_TYPE 6964 INT\n-- LINACCESS_SOURCE_TYPE 61 CURRENCY\n-- LINACCESS_SOURCE_TYPE 62 CURRENCY\n");
    linaccess::importAccessTable(db,"synthetic-inserts","cash",schema,linaccess::parseCsvDetailed("id,a,b\n"));
    auto database=linaccess::makeDaoDatabase(db);
    const auto query=[&](const std::string& source) { return database->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string(source)}).object; };
    const auto execute=[&](const std::string& source) { auto definition=query(source); definition->member("Execute",{}); return definition->member("RecordsAffected",{}).asNumber(); };
    const auto stored=[&](int id,const QString& column="a") { assert(sql.exec("SELECT "+column+" FROM cash WHERE id="+QString::number(id)) && sql.next()); const auto value=sql.value(0); sql.finish(); return value; };
    assert(execute("INSERT INTO cash(id,a,b) VALUES(1,'1.23455',NULL)")==1);
    assert(linaccess::currencyFromRealStorage(stored(1).toDouble()).currencyUnits==12346 && stored(1,"b").isNull());
    execute("INSERT INTO cash VALUES(2,'0.00005','0.00015')"); assert(stored(2).toDouble()==0 && linaccess::currencyFromRealStorage(stored(2,"b").toDouble()).currencyUnits==2);
    execute("INSERT INTO [cash]([b],[id],[a]) VALUES('0.2',3,CCur('0.1')+CCur('0.2'));"); assert(stored(3).toDouble()==0.3);
    auto parameterized=query("PARAMETERS amount Text(255), rowid Long; INSERT INTO cash(id,a) VALUES(rowid,amount)");
    parameterized->member("Parameters",{VbaValue::string("amount")}).object->assign("Value",VbaValue::string("549755813887.9999"));
    parameterized->member("Parameters",{VbaValue::string("rowid")}).object->assign("Value",VbaValue::longInteger(4)); parameterized->member("Execute",{});
    assert(linaccess::currencyFromRealStorage(stored(4).toDouble()).currencyUnits==5497558138879999LL && stored(4,"b").isNull());
    for(const auto& [bad,wanted]:std::vector<std::pair<std::string,int>>{{"invalid",13},{"549755813888",3251}}) {
        int error=0; try { execute("INSERT INTO cash VALUES(5,'0.1','"+bad+"')"); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
        assert(error==wanted); assert(sql.exec("SELECT COUNT(*) FROM cash") && sql.next() && sql.value(0).toInt()==4); sql.finish();
    }
    int syntax=0; try { execute("INSERT INTO cash VALUES(5,'0.1','0.2',"); } catch(const linaccess::VbaRuntimeError& failure) { syntax=failure.number; }
    assert(syntax==3075); assert(sql.exec("SELECT COUNT(*) FROM cash") && sql.next() && sql.value(0).toInt()==4); sql.finish();
    assert(sql.exec("CREATE TABLE ordinary(a REAL)")); execute("INSERT INTO ordinary VALUES('1.23455')"); assert(sql.exec("SELECT a FROM ordinary") && sql.next() && sql.value(0).toDouble()==1.23455); sql.finish();
    const QByteArray layout="Begin Form\nRecordSource =\"cash\"\nBegin Section\nBegin TextBox\nName =\"a\"\nControlSource =\"a\"\nWidth =5000\nEnd\nEnd\nEnd\n";
    const QByteArray code=R"VBA(Function InsertCurrency()
Dim q As DAO.QueryDef
Set q=CurrentDb.CreateQueryDef("","PARAMETERS p Currency; INSERT INTO cash(id,a) VALUES(6,p+CCur(""0.2""))")
q.Parameters("p").Value=CCur("0.1")
q.Execute
Me.Requery
DoCmd.GoToRecord , , acLast
InsertCurrency=VarType(Me!a)
End Function
)VBA";
    AccessFormPreview form("currency_inserts",layout,code,db); form.show(); QApplication::processEvents(); auto model=makePreviewFormObject(form); std::vector<VbaValue> args;
    const auto result=form.executeProcedure("InsertCurrency",args);
    assert(result.status==linaccess::VbaExecutionStatus::Completed && result.typedReturnValue->number==6 && model->item("a").currencyUnits==3000 && model->item("id").asNumber()==6 && !form.hasUnsavedChanges());
    assert(form.findChild<QLineEdit*>("access-control-a")->text()=="0.3" && form.close());
    assert(sql.exec("CREATE TABLE append_source(id INTEGER,value TEXT)")); assert(sql.exec("INSERT INTO append_source VALUES(7,'2.34565'),(8,'invalid')"));
    int appendError=0; try { execute("INSERT INTO cash(id,a) SELECT id,value FROM append_source ORDER BY id"); } catch(const linaccess::VbaRuntimeError& failure) { appendError=failure.number; }
    assert(appendError==13); assert(sql.exec("SELECT COUNT(*) FROM cash") && sql.next() && sql.value(0).toInt()==5); sql.finish();
    assert(sql.exec("UPDATE append_source SET value='0.00015' WHERE id=8"));
    assert(execute("INSERT INTO cash(id,a) SELECT id,value FROM append_source ORDER BY id")==2);
    assert(linaccess::currencyFromRealStorage(stored(7).toDouble()).currencyUnits==23456 && linaccess::currencyFromRealStorage(stored(8).toDouble()).currencyUnits==2);
    assert(execute("INSERT INTO cash(id,a) SELECT id+100,a+a FROM cash WHERE id=3")==1 && stored(103).toDouble()==0.6);
    auto appendParameter=query("PARAMETERS amount Text(255), rowid Long; INSERT INTO cash(id,a) SELECT rowid,amount FROM append_source WHERE id=7");
    appendParameter->member("Parameters",{VbaValue::string("amount")}).object->assign("Value",VbaValue::string("1.23455"));
    appendParameter->member("Parameters",{VbaValue::string("rowid")}).object->assign("Value",VbaValue::longInteger(9)); appendParameter->member("Execute",{});
    assert(linaccess::currencyFromRealStorage(stored(9).toDouble()).currencyUnits==12346);

}
