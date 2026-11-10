#include <cmath>
#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_query.hpp"
#include "access_table_import.hpp"
#include "dao_database.hpp"
#include "vba_currency.hpp"
#include <QApplication>
#include <QLineEdit>
#include <QSqlQuery>
#include <QTest>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <optional>

int main(int argc,char** argv) {
    QApplication app(argc,argv); assert(argc==2);
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","currency-query-types"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)")); assert(sql.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    const auto schema=linaccess::parseMdbSchema("CREATE TABLE `cash`\n(\n`amount` REAL\n);\n-- LINACCESS_SOURCE_TYPE 616d6f756e74 CURRENCY\n");
    linaccess::importAccessTable(db,"synthetic-query-types","cash",schema,linaccess::parseCsvDetailed("amount\n0.1\n"));
    if(QString::fromLatin1(argv[1])=="objects") {
        assert(sql.exec("CREATE TABLE object_audit(amount REAL)"));
        const QByteArray layout="Begin Form\nRecordSource =\"cash\"\nBegin Section\nBegin TextBox\nName =\"amount\"\nControlSource =\"amount\"\nWidth =4000\nEnd\nBegin TextBox\nName =\"free\"\nTop =500\nWidth =4000\nEnd\nEnd\nEnd\n";
        const QByteArray code=R"VBA(Sub CaptureReferences()
CurrentDb.Execute "INSERT INTO object_audit SELECT Forms!currency_types!amount"
CurrentDb.Execute "INSERT INTO object_audit SELECT Formulare!currency_types!amount"
CurrentDb.Execute "INSERT INTO object_audit SELECT Forms![currency_types]![amount].Value"
End Sub
Sub CaptureReports()
CurrentDb.Execute "INSERT INTO object_audit SELECT Reports!currency_report!amount"
CurrentDb.Execute "INSERT INTO object_audit SELECT Berichte!currency_report!amount"
End Sub
Sub CaptureFree()
CurrentDb.Execute "INSERT INTO object_audit SELECT Forms!currency_types!free"
End Sub
Function NumericReference()
Dim r As DAO.Recordset
Set r=CurrentDb.OpenRecordset("SELECT Forms!currency_types!amount<1 AS small")
NumericReference=r.Fields("small").Value
r.Close
End Function
)VBA";
        AccessFormPreview form("currency_types",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
        auto model=makePreviewFormObject(form); std::vector<VbaValue> none;
        assert(form.executeProcedure("CaptureReferences",none).status==linaccess::VbaExecutionStatus::Completed);
        const auto numeric=form.executeProcedure("NumericReference",none); assert(numeric.status==linaccess::VbaExecutionStatus::Completed && numeric.typedReturnValue->number==1);
        auto* editor=form.findChild<QLineEdit*>("access-control-amount"); assert(editor); editor->setFocus(); editor->selectAll(); QTest::keyClicks(editor,"0.6");
        assert(form.executeProcedure("CaptureReferences",none).status==linaccess::VbaExecutionStatus::Completed && editor->text()=="0.6" && model->item("amount").currencyUnits==1000 && form.hasUnsavedChanges());
        model->assignItem("free",VbaValue::currency(5497558138879999LL)); assert(form.executeProcedure("CaptureFree",none).status==linaccess::VbaExecutionStatus::Completed);
        model->assignItem("free",VbaValue::currency(5497558138880000LL)); const auto rejected=form.executeProcedure("CaptureFree",none); assert(rejected.status!=linaccess::VbaExecutionStatus::Completed && rejected.errorNumber==3251);
        const QByteArray reportLayout="Begin Report\nRecordSource =\"SELECT CCur('0.25') AS amount\"\nBegin Section\nBegin TextBox\nName =\"amount\"\nControlSource =\"amount\"\nEnd\nEnd\nEnd\n";
        AccessFormPreview report("currency_report",reportLayout,{},db); report.show(); QApplication::processEvents();
        assert(makePreviewFormObject(report)->item("amount").kind==VbaValue::Kind::Currency);
        assert(form.executeProcedure("CaptureReports",none).status==linaccess::VbaExecutionStatus::Completed);
        assert(sql.exec("SELECT amount FROM object_audit ORDER BY rowid"));
        for(int i=0;i<9;++i) assert(sql.next() && linaccess::currencyFromRealStorage(sql.value(0).toDouble()).currencyUnits==(i<6 ? 1000 : i==6 ? 5497558138879999LL : 2500));
        assert(!sql.next()); sql.finish(); assert(sql.exec("SELECT amount FROM cash") && sql.next() && sql.value(0).toDouble()==0.1); sql.finish();
        model->member("Undo",{}); assert(!form.hasUnsavedChanges() && report.close() && form.close()); return 0;
    }
    if(QString::fromLatin1(argv[1])=="schema") {
        assert(sql.exec("INSERT INTO cash VALUES(0.2),(-0.3),(NULL)"));
        auto database=linaccess::makeDaoDatabase(db);
        const auto open=[&](const std::string& source) { return database->member("OpenRecordset",{VbaValue::string(source)}).object; };
    {
        assert(sql.exec("CREATE TABLE cross_probe(row_key INTEGER,category TEXT,cell INTEGER)"));
        assert(sql.exec("INSERT INTO cross_probe VALUES(1,'A',1),(1,'A',NULL),(1,'B',2),(2,'B',NULL),(2,'Z',3),(3,NULL,NULL)"));
        const QString source="TRANSFORM Count([cell]) AS n SELECT [row_key] AS row_id FROM cross_probe GROUP BY row_key ORDER BY row_key PIVOT category IN ('B','A','missing');";
        const auto plan=linaccess::compileAccessQuery(source);assert(plan.daoQueryType==16 && plan.source==source);
        const auto result=linaccess::executeAccessQueryPlan(db,plan,{});
        assert(result.columns==QStringList({"row_id","B","A","missing"}) && result.rows.size()==3 && result.types.value("b")=="INTEGER4");
        assert(result.rows[0][1].toInt()==1 && result.rows[0][2].toInt()==1 && result.rows[0][3].toInt()==0 && result.rows[1][1].toInt()==0 && result.rows[2][1].toInt()==0);
        auto rows=open(source.toStdString());assert(rows->member("Fields",{VbaValue::string("B")}).object->member("Type",{}).number==4);rows->member("Close",{});
        auto definition=database->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string(("PARAMETERS wanted LONG; "+QString(source).replace("FROM cross_probe","FROM cross_probe WHERE row_key=[wanted]")).toStdString())}).object;
        assert(definition->member("Type",{}).number==16);definition->member("Parameters",{VbaValue::string("wanted")}).object->assign("Value",VbaValue::numeric(1));
        rows=definition->member("OpenRecordset",{}).object;assert(rows->member("RecordCount",{}).number==1);rows->member("Close",{});
        const auto stars=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery(QString(source).replace("Count([cell])","Count(*)")),{});
        assert(stars.rows[0][2].toInt()==2 && stars.rows[1][1].toInt()==1);
        assert(sql.exec("CREATE TABLE cross_dates(row_key INTEGER,stamp DATETIME)"));assert(sql.exec("INSERT INTO cross_dates VALUES(1,'2003-12-31 18:05:06'),(1,'2004-02-29 00:00:00'),(2,NULL)"));
        const auto calendar=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("TRANSFORM Count(*) SELECT row_key FROM cross_dates GROUP BY row_key ORDER BY row_key PIVOT Month(stamp) IN (12,+2,-1);"),{});
        assert(calendar.columns==QStringList({"row_key","12","2","-1"}) && calendar.rows.size()==2 && calendar.rows[0][1].toInt()==1 && calendar.rows[0][2].toInt()==1 && calendar.rows[1][1].toInt()==0);
        const auto years=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("TRANSFORM Count(stamp) SELECT row_key FROM cross_dates GROUP BY row_key PIVOT Year(stamp) IN (2003,2004);"),{});assert(years.rows[0][1].toInt()==1 && years.rows[0][2].toInt()==1);
        auto saved=database->member("CreateQueryDef",{VbaValue::string("fixed_cross_probe"),VbaValue::string(source.toStdString())}).object;
        rows=open("fixed_cross_probe");assert(rows->member("RecordCount",{}).number==3);rows->member("Close",{});
        assert(sql.exec("CREATE TABLE cross_extrema(row_key INTEGER4,category TEXT,amount CURRENCY_REAL,quantity INTEGER2,stamp DATETIME)"));
        assert(sql.exec("INSERT INTO cross_extrema VALUES(1,'A',0.2,7,'2003-12-31 18:05:06'),(1,'A',-0.3,-2,'2003-01-01 00:00:00'),(1,'B',NULL,NULL,NULL),(2,'B',0.1,4,'2004-02-29 00:00:00')"));
        const QString extrema="TRANSFORM Min(c.amount) SELECT c.row_key FROM cross_extrema AS c GROUP BY c.row_key ORDER BY c.row_key PIVOT c.category IN ('A','B','missing')";
        for(const auto* function:{"Min","Max"}) {
            const auto query=QString(extrema).replace("Min(",QString(function)+'(');
            const auto minimum=QString(function)=="Min";
            const auto values=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery(query),{});
            assert(values.rows.size()==2 && values.types.value("a")=="CURRENCY_REAL" && values.rows[0][1].toDouble()==(minimum ? -0.3 : 0.2));
            assert(values.rows[0][2].isNull() && values.rows[0][3].isNull() && values.rows[1][1].isNull() && values.rows[1][2].toDouble()==0.1);
            rows=open(query.toStdString());auto amount=rows->member("Fields",{VbaValue::string("A")}).object;
            assert(amount->member("Type",{}).number==5 && amount->member("Value",{}).kind==VbaValue::Kind::Currency && amount->member("Value",{}).currencyUnits==(minimum ? -3000 : 2000));rows->member("Close",{});
            const auto integers=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery(QString(query).replace("c.amount","c.quantity")),{});
            assert(integers.types.value("a")=="INTEGER2" && integers.rows[0][1].toInt()==(minimum ? -2 : 7));
            const auto dates=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery(QString(query).replace("c.amount","c.stamp")),{});
            assert(dates.types.value("a")=="DATETIME" && dates.rows[0][1].toString()==(minimum ? "2003-01-01 00:00:00" : "2003-12-31 18:05:06"));
        }
        assert(sql.exec("CREATE TABLE cross_statistics(row_key INTEGER,category TEXT,value REAL)"));
        assert(sql.exec("INSERT INTO cross_statistics VALUES(1,'A',10),(1,'A',20),(1,'A',30),(1,'B',7),(1,'B',NULL),(2,'A',NULL),(2,'Z',5)"));sql.finish();
        const QString statisticalSource="TRANSFORM Avg(s.value) SELECT s.row_key FROM cross_statistics AS s GROUP BY s.row_key ORDER BY s.row_key PIVOT s.category IN ('A','B','absent')";
        for(const auto& function:QStringList{"Avg","Var","StDev","VarP","StDevP"}) {
            const auto source=QString(statisticalSource).replace("Avg(",function+'(');
            const auto values=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery(source),{});
            assert(values.rows.size()==2 && values.types.value("a")=="DOUBLE" && values.types.value("b")=="DOUBLE");
            const double expected=function=="Avg" ? 20 : function=="Var" ? 100 : function=="StDev" ? 10 : function=="VarP" ? 200.0/3 : std::sqrt(200.0/3);
            assert(qAbs(values.rows[0][1].toDouble()-expected)<1e-10 && values.rows[0][3].isNull() && values.rows[1][1].isNull() && values.rows[1][2].isNull());
            if(function=="Var" || function=="StDev") assert(values.rows[0][2].isNull());
            else assert(!values.rows[0][2].isNull() && values.rows[0][2].toDouble()==(function=="Avg" ? 7 : 0));
            auto cursor=open(source.toStdString());const auto field=cursor->member("Fields",{VbaValue::string("A")}).object;
            assert(field->member("Type",{}).number==7 && field->member("Value",{}).kind==VbaValue::Kind::Number);
            cursor->member("Close",{});
        }
        auto statisticDefinition=database->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string(("PARAMETERS wanted LONG; "+QString(statisticalSource).replace("GROUP BY","WHERE s.row_key=[wanted] GROUP BY")).toStdString())}).object;
        assert(statisticDefinition->member("Type",{}).number==16);
        statisticDefinition->member("Parameters",{VbaValue::string("wanted")}).object->assign("Value",VbaValue::longInteger(1));
        auto statisticCursor=statisticDefinition->member("OpenRecordset",{}).object;assert(statisticCursor->member("RecordCount",{}).number==1);statisticCursor->member("Close",{});
        for(const auto& bad:QStringList{QString(statisticalSource).replace("s.value","s.category"),QString(extrema).replace("Min(","Avg(")}) {
            int error=0;try{(void)linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery(bad),{});}catch(const linaccess::VbaRuntimeError& failure){error=failure.number;}assert(error==3251);
        }
        assert(sql.exec("CREATE TEMP TABLE cross_statistics(row_key INTEGER,category TEXT,value REAL)"));
        int shadowError=0;try{(void)linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery(statisticalSource),{});}catch(const linaccess::VbaRuntimeError& failure){shadowError=failure.number;}assert(shadowError==3251);
        assert(sql.exec("DROP TABLE temp.cross_statistics"));sql.finish();
        auto extremumDefinition=database->member("CreateQueryDef",{VbaValue::string("fixed_extrema"),VbaValue::string(extrema.toStdString())}).object;
        assert(extremumDefinition->member("Type",{}).number==16);
        rows=open("fixed_extrema");assert(rows->member("Fields",{VbaValue::string("A")}).object->member("Type",{}).number==5);rows->member("Close",{});
        const auto filteredExtrema="PARAMETERS wanted LONG; "+QString(extrema).replace("GROUP BY","WHERE c.row_key=[wanted] GROUP BY");
        auto parameterExtrema=database->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string(filteredExtrema.toStdString())}).object;
        parameterExtrema->member("Parameters",{VbaValue::string("wanted")}).object->assign("Value",VbaValue::numeric(2));
        rows=parameterExtrema->member("OpenRecordset",{}).object;assert(rows->member("RecordCount",{}).number==1);rows->member("Close",{});
        for(const auto& unsupported:QStringList{QString(extrema).replace("c.amount","c.category"),QString(extrema).replace("c.amount","c.absent"),QString(extrema).replace("cross_extrema AS c","fixed_cross_probe AS c")}) {
            int refused=0;try {(void)linaccess::accessQueryColumns(db,linaccess::compileAccessQuery(unsupported));}catch(const linaccess::VbaRuntimeError& error){refused=error.number;}assert(refused==3251);
        }
        const auto emptyExtrema=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery(QString(extrema).replace("GROUP BY","WHERE c.row_key=99 GROUP BY")),{});
        assert(emptyExtrema.rows.isEmpty() && emptyExtrema.types.value("a")=="CURRENCY_REAL");
        assert(linaccess::installAccessQueryViews(db)>=0);
        assert(sql.exec("SELECT [A] FROM fixed_extrema ORDER BY row_key") && sql.next() && sql.value(0).toDouble()==-0.3);sql.finish();
        assert(sql.exec("CREATE TEMP TABLE cross_extrema(row_key INTEGER4,category TEXT,amount CURRENCY_REAL)"));
        int shadowRefused=0;try {(void)linaccess::accessQueryColumns(db,linaccess::compileAccessQuery(extrema));}catch(const linaccess::VbaRuntimeError& error){shadowRefused=error.number;}assert(shadowRefused==3251);
        assert(linaccess::installAccessQueryViews(db)>=0);
        assert(sql.exec("SELECT COUNT(*) FROM temp.sqlite_master WHERE type='view' AND name='fixed_extrema'") && sql.next() && sql.value(0).toInt()==0);sql.finish();
        assert(sql.exec("DROP TABLE temp.cross_extrema"));
        assert(linaccess::installAccessQueryViews(db)>=0);
        assert(sql.exec("SELECT [A] FROM fixed_extrema ORDER BY row_key") && sql.next() && sql.value(0).toDouble()==-0.3);sql.finish();
        for(const auto* invalid:{"TRANSFORM Sum(cell) SELECT row_key FROM cross_probe GROUP BY row_key PIVOT category IN ('A')", "TRANSFORM Count(*) SELECT row_key FROM cross_probe GROUP BY row_key PIVOT category", "TRANSFORM Count(*) SELECT row_key FROM cross_probe GROUP BY row_key PIVOT category IN ('A','a')", "TRANSFORM Count(*) SELECT row_key FROM cross_probe GROUP BY row_key PIVOT category IN ('row_key')", "TRANSFORM Count(*) SELECT row_key FROM cross_probe GROUP BY row_key PIVOT category IN ('A',)", "TRANSFORM Count(*) SELECT row_key FROM cross_probe GROUP BY row_key PIVOT category IN (1,01)", "TRANSFORM Count(*) SELECT row_key FROM cross_probe GROUP BY row_key PIVOT category IN (1.5)", "TRANSFORM Count(*) SELECT row_key FROM cross_probe GROUP BY row_key PIVOT category IN (2147483648)", "TRANSFORM Count(*) SELECT row_key FROM cross_dates GROUP BY row_key PIVOT CDate(stamp) IN (1)"}) {
            int refused=0;try {(void)linaccess::compileAccessQuery(invalid);}catch(const linaccess::VbaRuntimeError& error) {refused=error.number;}assert(refused==3251);
        }
    }

        auto rows=open("SELECT amount FROM cash ORDER BY rowid"); auto field=rows->member("Fields",{VbaValue::string("amount")}).object;
        assert(field->member("Type",{}).number==5);
        for(const auto expected:std::vector<std::optional<std::int64_t>>{1000,2000,-3000,{}}) {
            const auto value=field->member("Value",{}); assert(expected ? value.kind==VbaValue::Kind::Currency && value.currencyUnits==*expected : value.kind==VbaValue::Kind::Null); rows->member("MoveNext",{});
        }
        rows->member("Close",{});
        for(const auto* source:{"SELECT cash.amount AS money FROM cash WHERE rowid=1","SELECT c.amount AS money FROM cash AS c WHERE rowid=1","SELECT Sum(c.amount) AS money FROM cash AS c"}) {
            rows=open(source); field=rows->member("Fields",{VbaValue::string("money")}).object;
            assert(field->member("Type",{}).number==5 && field->member("Value",{}).kind==VbaValue::Kind::Currency); rows->member("Close",{});
        }
        const auto base=linaccess::compileAccessQuery("SELECT Sum(amount) AS total FROM cash");
        const auto filtered=linaccess::filterAccessQueryPlan(db,base,"total=0");
        auto result=linaccess::executeAccessQueryPlan(db,filtered,{}); assert(result.rows.size()==1 && result.rows[0][0].toDouble()==0 && result.types.value("total")=="CURRENCY_REAL");
        auto modified=linaccess::compileAccessQuery("SELECT amount FROM cash"); modified.sql="SELECT 2 AS amount";
        result=linaccess::executeAccessQueryPlan(db,modified,{}); assert(result.rows[0][0].toInt()==2 && result.types.isEmpty());
        modified=linaccess::compileAccessQuery("SELECT CCur('0.1') AS amount"); modified.sql="SELECT 2 AS amount";
        result=linaccess::executeAccessQueryPlan(db,modified,{}); assert(result.rows[0][0].toInt()==2 && result.types.isEmpty());
        auto resolved=linaccess::filterAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT amount FROM cash"),{});
        auto changedResolved=resolved; changedResolved.sql="SELECT 2 AS amount";
        result=linaccess::executeAccessQueryPlan(db,changedResolved,{}); assert(result.rows[0][0].toInt()==2 && result.types.isEmpty());
        assert(sql.exec("UPDATE _la_table_schema SET definition=replace(definition,'CURRENCY','FLOAT') WHERE table_name='cash'"));
        result=linaccess::executeAccessQueryPlan(db,resolved,{}); assert(result.types.isEmpty());
        assert(sql.exec("UPDATE _la_table_schema SET definition=replace(definition,'FLOAT','CURRENCY') WHERE table_name='cash'"));
        result=linaccess::executeAccessQueryPlan(db,resolved,{}); assert(result.types.value("amount")=="CURRENCY_REAL");
        assert(sql.exec("CREATE TABLE ordinary(amount REAL)")); assert(sql.exec("INSERT INTO ordinary VALUES(0.1),(0.2),(-0.3)"));
        for(const auto* source:{"SELECT Sum(amount) AS total FROM ordinary","SELECT Sum(amount) AS total FROM cash HAVING 1","SELECT cash.amount FROM cash JOIN ordinary ON 1"}) {
            const auto plan=linaccess::compileAccessQuery(source); assert(linaccess::executeAccessQueryPlan(db,plan,{}).types.isEmpty());
        }
        const QByteArray formLayout="Begin Form\nRecordSource =\"SELECT amount FROM cash WHERE rowid=1\"\nBegin Section\nBegin TextBox\nName =\"amount\"\nControlSource =\"amount\"\nEnd\nEnd\nEnd\n";
        AccessFormPreview form("implicit_currency",formLayout,{},db); form.show(); QApplication::processEvents(); auto model=makePreviewFormObject(form);
        assert(model->item("amount").kind==VbaValue::Kind::Currency && model->item("amount").currencyUnits==1000);
        assert(form.findChild<QLineEdit*>("access-control-amount")->isReadOnly()); assert(form.close());
        const QByteArray reportLayout="Begin Report\nRecordSource =\"SELECT Sum(amount) AS total FROM cash\"\nBegin Section\nBegin TextBox\nName =\"total\"\nControlSource =\"total\"\nEnd\nEnd\nEnd\n";
        AccessFormPreview report("implicit_sum",reportLayout,{},db); report.show(); QApplication::processEvents();
        assert(makePreviewFormObject(report)->item("total").kind==VbaValue::Kind::Currency && makePreviewFormObject(report)->item("total").currencyUnits==0); assert(report.close());
        assert(db.transaction()); assert(sql.exec("UPDATE cash SET amount=1.234567 WHERE rowid=1"));
        rows=open("SELECT amount FROM cash WHERE rowid=1"); int error=0;
        try { (void)rows->member("Fields",{VbaValue::string("amount")}).object->member("Value",{}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==13);
        rows->member("Close",{}); assert(db.rollback()); return 0;
    }
    assert(QString::fromLatin1(argv[1])=="results");
    auto database=linaccess::makeDaoDatabase(db);
    const auto open=[&](const std::string& source) { return database->member("OpenRecordset",{VbaValue::string(source)}).object; };
    for(const auto* expression:{"CCur('549755813887.9999')","(CCur('549755813887.9999'))","+CCur('549755813887.9999')","-CCur('-549755813887.9999')"}) {
        auto rows=open(std::string("SELECT ")+expression+" AS amount"); auto field=rows->member("Fields",{VbaValue::string("amount")}).object;
        assert(field->member("Type",{}).number==5 && field->member("Value",{}).kind==VbaValue::Kind::Currency && field->member("Value",{}).currencyUnits==5497558138879999LL);
        rows->member("MoveNext",{}); assert(rows->member("EOF",{}).truthy() && field->member("Type",{}).number==5); rows->member("Close",{});
    }
    for(const auto* source:{"SELECT TOP 1 CCur('1.2345') AS amount FROM cash ORDER BY amount","SELECT DISTINCT CCur('1.2345') AS amount FROM cash","SELECT CCur('1.2345') AS amount WHERE 0"}) {
        auto rows=open(source); auto field=rows->member("Fields",{VbaValue::string("amount")}).object;
        assert(field->member("Type",{}).number==5);
        if(!rows->member("EOF",{}).truthy()) assert(field->member("Value",{}).currencyUnits==12345);
        rows->member("Close",{});
    }
    sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('query','money_projection',?)"); sql.addBindValue(QByteArray("SELECT CCur('1.2345') AS amount")); assert(sql.exec()); sql.finish();
    auto saved=open("money_projection"); assert(saved->member("Fields",{VbaValue::string("amount")}).object->member("Value",{}).currencyUnits==12345); saved->member("Close",{});
    auto definition=database->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string("PARAMETERS p Money; SELECT [p] AS amount")}).object;
    auto parameter=definition->member("Parameters",{VbaValue::string("p")}).object; parameter->assign("Value",VbaValue::currency(12345));
    auto parameterRows=definition->member("OpenRecordset",{}).object; auto parameterField=parameterRows->member("Fields",{VbaValue::string("amount")}).object;
    assert(parameterField->member("Type",{}).number==5 && parameterField->member("Value",{}).currencyUnits==12345); parameterRows->member("Close",{});
    parameter->assign("Value",{}); parameterRows=definition->member("OpenRecordset",{}).object;
    assert(parameterRows->member("Fields",{VbaValue::string("amount")}).object->member("Value",{}).kind==VbaValue::Kind::Null); parameterRows->member("Close",{});
    const auto base=linaccess::compileAccessQuery("SELECT CCur('1.2345') AS amount,2 AS plain");
    const auto filtered=linaccess::filterAccessQueryPlan(db,base,"amount>1");
    const auto result=linaccess::executeAccessQueryPlan(db,filtered,{});
    assert(result.rows.size()==1 && result.types.value("amount")=="CURRENCY_REAL" && !result.types.contains("plain"));
    const auto empty=linaccess::executeAccessQueryPlan(db,linaccess::filterAccessQueryPlan(db,base,"amount<0"),{}); assert(empty.rows.empty() && empty.types.value("amount")=="CURRENCY_REAL");
    for(const auto* source:{"SELECT CCur('1.2') AS amount UNION ALL SELECT 2 AS amount","SELECT *,CCur('1.2') AS extra FROM cash","SELECT CCur('1.2') AS amount,2 AS amount"}) {
        const auto plan=linaccess::compileAccessQuery(source); const auto unknown=linaccess::executeAccessQueryPlan(db,plan,{}); assert(unknown.types.isEmpty());
    }
    linaccess::AccessQueryModel table; table.execute(db,"SELECT CCur('549755813887.9999') AS amount");
    assert(table.data(table.index(0,0),Qt::DisplayRole).toString()=="549755813887.9999" && linaccess::currencyFromRealStorage(table.data(table.index(0,0),Qt::EditRole).toDouble()).currencyUnits==5497558138879999LL);
    const QByteArray layout="Begin Form\nRecordSource =\"SELECT CCur('549755813887.9999') AS amount\"\nBegin Section\nBegin TextBox\nName =\"amount\"\nControlSource =\"amount\"\nWidth =5000\nEnd\nEnd\nEnd\n";
    AccessFormPreview form("currency_projection",layout,{},db); form.show(); QApplication::processEvents(); auto model=makePreviewFormObject(form);
    auto* editor=form.findChild<QLineEdit*>("access-control-amount"); assert(editor && editor->isReadOnly() && editor->text()=="549755813887.9999" && model->item("amount").kind==VbaValue::Kind::Currency && model->item("amount").currencyUnits==5497558138879999LL);
    model->assign("Filter",VbaValue::string("amount>1")); model->assign("FilterOn",VbaValue::logical(true)); assert(model->item("amount").currencyUnits==5497558138879999LL && !form.hasUnsavedChanges()); assert(form.close());
}
