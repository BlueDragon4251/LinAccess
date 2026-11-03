#include "dao_database.hpp"
#include "vba_qt_value.hpp"
#include "access_query.hpp"
#include "sqlite_execution.hpp"
#include "vba_runtime.hpp"
#include <QCoreApplication>
#include <QSqlQuery>
#include <QTemporaryDir>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>
#include <map>
#include <limits>

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
void testSqlRgbDao() {
    auto db=QSqlDatabase::addDatabase("QSQLITE","sql-rgb-dao"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE color_rows(id INTEGER PRIMARY KEY,r REAL,g REAL,b REAL,packed INTEGER)"));
    assert(sql.exec("INSERT INTO color_rows VALUES(1,255,0,0,0),(2,0,255,0,0)"));
    Host host(db);
    const auto painted=linaccess::executeVbaProcedure(R"VBA(Sub PaintRows()
Dim database As DAO.Database
Dim rows As DAO.Recordset
Set database=CurrentDb
Set rows=database.OpenRecordset(Type:=4, Name:="SELECT id,RGB(r,g,b) AS shade FROM color_rows ORDER BY id")
rows.AbsolutePosition=0
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
    assert(host.fields["FirstShade"]=="255" && host.fields["SecondShade"]=="65280" && host.fields["ColorAffected"]=="2" && host.fields["Ordinal"]=="1");
    assert(sql.exec("SELECT packed FROM color_rows ORDER BY id") && sql.next() && sql.value(0).toInt()==255 && sql.next() && sql.value(0).toInt()==65280); sql.finish();
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
int main(int argc, char** argv) {
    QCoreApplication app(argc, argv);
    testVariantArraysDao();
    testDdlRecordsAffected();
    testTypedSqlTargets();
    testDeclaredParameterTypes();
    testIntegerConversionDao();
    testBooleanQueryParameters();
    testDaoCursors();
    testDaoTransactions();
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
