#include <cmath>
#include "access_query.hpp"
#include "vba_qt_value.hpp"
#include "vba_runtime.hpp"
#include "sqlite_execution.hpp"
#include "jet_sql_functions.hpp"
#include <QCoreApplication>
#include <QSqlQuery>
#include <QSqlError>
#include <QTemporaryFile>
#include <QDateTime>
#include <QTimeZone>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>
#include <vector>
#include <utility>
int main(int argc, char** argv) {
    QCoreApplication app(argc,argv);
    auto db = QSqlDatabase::addDatabase("QSQLITE");
    db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE samples(id INTEGER, score INTEGER, label TEXT)"));
    assert(query.exec("INSERT INTO samples VALUES(1,10,'one'),(2,20,'two'),(3,20,NULL),(4,30,'four')"));
    {
        linaccess::AccessQueryModel labels;
        labels.execute(db,"SELECT FV(0,2,-10),IPmt(0,1,12,1200) AS paid FROM samples");
        assert(labels.rowCount()==4 && labels.index(0,0).data().toDouble()==20);
        assert(labels.headerData(0,Qt::Horizontal).toString().contains("FV") && !labels.headerData(0,Qt::Horizontal).toString().contains("linaccess_"));
        assert(labels.headerData(1,Qt::Horizontal).toString()=="paid");
        labels.execute(db,"SELECT *,IPmt(0,1,12,1200) FROM samples");
        assert(labels.columnCount()==4 && labels.headerData(3,Qt::Horizontal).toString().contains("IPmt") && !labels.headerData(3,Qt::Horizontal).toString().contains("linaccess_"));
        labels.execute(db,"SELECT FV(0,2,-10) UNION ALL SELECT FV(0,3,-10)");
        assert(labels.rowCount()==2 && labels.index(1,0).data().toDouble()==30 && !labels.headerData(0,Qt::Horizontal).toString().contains("linaccess_"));
        labels.execute(db,"SELECT FV(0,2,-10) future"); assert(labels.headerData(0,Qt::Horizontal).toString()=="future");
        labels.execute(db,"SELECT FV(0,2,-10) 'future'"); assert(labels.headerData(0,Qt::Horizontal).toString()=="future");
        labels.execute(db,"SELECT [label] & '!' FROM samples");
        assert(labels.index(0,0).data().toString()=="one!" && !labels.headerData(0,Qt::Horizontal).toString().contains("linaccess_"));
    }
    {
        QTemporaryFile file("/tmp/linaccess-sql-file-ä-XXXXXX"); assert(file.open() && file.write("abc",3)==3 && file.flush());
        const auto modified=QDateTime(QDate(2024,2,29),QTime(12,34,56),QTimeZone(0));
        assert(file.setFileTime(modified,QFileDevice::FileModificationTime));
        const QString source="PARAMETERS path Text; SELECT FileLen([path]) AS bytes,FileDateTime([path]) AS stamp";
        const QMap<QString,QVariant> parameters{{"path",file.fileName()}};
        linaccess::AccessQueryModel metadata; metadata.execute(db,source,parameters);
        assert(metadata.rowCount()==1 && metadata.index(0,0).data().toLongLong()==3);
        assert(metadata.fieldSqlType(0)=="INTEGER4" && metadata.fieldSqlType(1)=="DATETIME");
        assert(metadata.index(0,1).data().toString()==modified.toLocalTime().toString("yyyy-MM-dd HH:mm:ss"));
        assert(file.resize(7) && file.flush()); metadata.execute(db,source,parameters);
        assert(metadata.index(0,0).data().toLongLong()==7);
        assert(file.remove()); const auto plan=linaccess::compileAccessQuery(source); assert(!plan.sql.isEmpty());
        int failure=0; try { metadata.execute(db,source,parameters); } catch(const linaccess::VbaRuntimeError& error) { failure=error.number; }
        assert(failure==53 && metadata.rowCount()==0);
        failure=0; try { metadata.execute(db,"SELECT FileDateTime(Null)"); } catch(const linaccess::VbaRuntimeError& error) { failure=error.number; }
        assert(failure==94);
    }
    {
        assert(query.exec("CREATE TABLE origin_left(id INTEGER,key INTEGER)"));
        assert(query.exec("CREATE TABLE origin_right(id INTEGER,key INTEGER)"));
        assert(query.exec("INSERT INTO origin_left VALUES(11,7)")); assert(query.exec("INSERT INTO origin_right VALUES(22,7)"));
        assert(query.exec("CREATE VIEW origin_saved AS SELECT origin_left.id AS id,origin_right.id AS id FROM origin_left INNER JOIN origin_right ON origin_left.key=origin_right.key"));
        linaccess::AccessQueryModel qualified;
        qualified.execute(db,"SELECT origin_saved.origin_right.id AS chosen FROM origin_saved");
        assert(qualified.rowCount()==1 && qualified.columnCount()==1 && qualified.data(qualified.index(0,0)).toInt()==22);
        qualified.execute(db,"SELECT s.origin_left.id,s.origin_right.id FROM origin_saved AS s WHERE s.origin_right.id=22 ORDER BY s.origin_left.id");
        assert(qualified.rowCount()==1 && qualified.columnCount()==2 && qualified.data(qualified.index(0,0)).toInt()==11 && qualified.data(qualified.index(0,1)).toInt()==22);
        qualified.execute(db,"SELECT origin_saved.origin_left.id FROM origin_saved WHERE origin_saved.origin_right.id=23");
        assert(qualified.rowCount()==0 && qualified.columnCount()==1);
        assert(query.exec("CREATE TABLE origin_shared(id INTEGER,key INTEGER)")); assert(query.exec("INSERT INTO origin_shared VALUES(11,7),(22,7)"));
        assert(query.exec("CREATE VIEW origin_before AS SELECT id,key FROM origin_shared WHERE id=11"));
        assert(query.exec("CREATE VIEW origin_after AS SELECT id,key FROM origin_shared WHERE id=22"));
        assert(query.exec("CREATE VIEW origin_rank_raw AS SELECT origin_before.id AS id,origin_after.id AS id,RANK() OVER(ORDER BY origin_before.id DESC,origin_after.id DESC) AS rank_value FROM origin_before INNER JOIN origin_after ON origin_before.key=origin_after.key"));
        assert(query.exec("CREATE VIEW origin_ranked AS SELECT [id],[id:1] FROM origin_rank_raw WHERE rank_value<=1"));
        qualified.execute(db,"SELECT origin_ranked.origin_before.id,origin_ranked.origin_after.id FROM origin_ranked WHERE origin_ranked.origin_after.id=22");
        assert(qualified.rowCount()==1 && qualified.columnCount()==2 && qualified.data(qualified.index(0,0)).toInt()==11 && qualified.data(qualified.index(0,1)).toInt()==22);
        assert(query.exec("CREATE VIEW origin_ambiguous AS SELECT id AS first_copy,id AS second_copy FROM origin_left"));
        assert(query.exec("CREATE VIEW origin_union AS SELECT id FROM origin_left UNION ALL SELECT id FROM origin_right"));
        for(const auto& source:{"SELECT origin_union.origin_left.id FROM origin_union","SELECT origin_ambiguous.origin_left.id FROM origin_ambiguous","SELECT origin_saved.absent_table.id FROM origin_saved","SELECT main.origin_saved.origin_right.id FROM origin_saved"}) {
            int error=0; try { qualified.execute(db,source); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
            assert(error!=0 && qualified.rowCount()==0 && qualified.columnCount()==0);
        }
    }
    {
        assert(query.exec("CREATE TABLE editable_source(id INTEGER PRIMARY KEY,value TEXT)"));
        assert(linaccess::editableAccessQueryTable(db,"SELECT * FROM [editable_source]")=="editable_source");
        QString editableFilter;
        assert(linaccess::editableAccessQueryTable(db,"SELECT * FROM editable_source WHERE id=1 AND (value='a;b' OR value IS NULL)",&editableFilter)=="editable_source" && !editableFilter.isEmpty());
        assert(query.exec("INSERT INTO editable_source VALUES(1,'a;b'),(2,'other')"));
        assert(query.exec("SELECT * FROM editable_source WHERE "+editableFilter) && query.next() && query.value(0).toInt()==1 && !query.next()); query.finish();
        assert(query.exec("CREATE TABLE editable_nullable(id TEXT PRIMARY KEY,value TEXT)"));
        assert(query.exec("CREATE TABLE editable_desc(id INTEGER PRIMARY KEY DESC,value TEXT)"));
        assert(query.exec("CREATE TABLE editable_nonnull(id TEXT NOT NULL PRIMARY KEY,value TEXT)"));
        assert(linaccess::editableAccessQueryTable(db,"SELECT * FROM editable_nonnull")=="editable_nonnull");
        assert(query.exec("CREATE TABLE editable_generated(id INTEGER PRIMARY KEY,value INTEGER,doubled INTEGER GENERATED ALWAYS AS(value*2))"));
        for(const auto& source:{"SELECT * FROM editable_nullable","SELECT * FROM editable_desc","SELECT id FROM editable_source","SELECT * FROM editable_source ORDER BY id","SELECT * FROM editable_source WHERE id IN(SELECT id FROM editable_source)","SELECT * FROM samples","PARAMETERS p Long; SELECT * FROM editable_source","SELECT * FROM origin_saved","SELECT * FROM editable_generated"}) {
            int error=0; try { (void)linaccess::editableAccessQueryTable(db,source); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
            assert(error==3251);
        }
        assert(query.exec("CREATE TEMP VIEW editable_source AS SELECT * FROM samples"));
        int error=0; try { (void)linaccess::editableAccessQueryTable(db,"SELECT * FROM editable_source"); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
        assert(error==3251); assert(query.exec("DROP VIEW temp.editable_source"));
    }
    {
        linaccess::AccessQueryModel reversed;
        reversed.execute(db,"SELECT InStrRev('abcabc','bc') AS pos,StrReverse('abö') AS backwards, InStrRev('A😀x😀','😀') AS units");
        assert(reversed.rowCount()==1 && reversed.data(reversed.index(0,0)).toInt()==5);
        assert(reversed.data(reversed.index(0,1)).toString()=="öba" && reversed.data(reversed.index(0,2)).toInt()==5);
        reversed.execute(db,"SELECT InStrRev(label,'o') AS pos,pos+1 AS following FROM samples WHERE id=1");
        assert(reversed.data(reversed.index(0,0)).toInt()==1 && reversed.data(reversed.index(0,1)).toInt()==2);
        assert(!query.exec("SELECT linaccess_jet_strreverse(Null)"));
        assert(linaccess::jetSqlFunctionErrorNumber(query.lastError().text().toCaseFolded().toStdString())==94); query.finish();
        reversed.execute(db,"SELECT InStrRev(Null,'x') AS missing,InStrRev('x',Null) AS pos");
        assert(reversed.data(reversed.index(0,0)).isNull() && reversed.data(reversed.index(0,1)).isNull());
        reversed.execute(db,"SELECT Count(InStrRev(label,'o')) FROM samples");
        assert(reversed.data(reversed.index(0,0)).toInt()==3);
        reversed.execute(db,"SELECT Sqr(score),Exp(Log(score)),Sin(Atn(1)),Cos(0),Tan(0) FROM samples WHERE id=1");
        assert(std::abs(reversed.data(reversed.index(0,0)).toDouble()-std::sqrt(10.0))<1e-12);
        assert(std::abs(reversed.data(reversed.index(0,1)).toDouble()-10.0)<1e-12);
        assert(std::abs(reversed.data(reversed.index(0,2)).toDouble()-std::sqrt(0.5))<1e-12);
        assert(reversed.data(reversed.index(0,3)).toDouble()==1 && reversed.data(reversed.index(0,4)).toDouble()==0);
        assert(!query.exec("SELECT linaccess_jet_log(0)"));
        assert(linaccess::jetSqlFunctionErrorNumber(query.lastError().text().toCaseFolded().toStdString())==5); query.finish();
        assert(query.exec("CREATE TABLE native_dates(id INTEGER,stamp DATETIME)"));
        assert(query.exec("INSERT INTO native_dates VALUES(1,'2024-02-28 12:00:00'),(2,NULL)"));
        const auto arithmetic=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT stamp+1 AS following,stamp-CDate('2024-02-27 12:00:00') AS elapsed FROM native_dates ORDER BY id"),{});
        assert(arithmetic.types.value("following")=="DATETIME" && arithmetic.types.value("elapsed")=="DOUBLE");
        assert(arithmetic.rows[0][0].toString()=="2024-02-29 12:00:00" && arithmetic.rows[0][1].toDouble()==1);
        assert(arithmetic.rows[1][0].isNull() && arithmetic.rows[1][1].isNull());
        const auto parts=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT DateValue('2024-02-29 12:34') AS d,TimeValue('14:24') AS t,TimeValue(Null) AS missing"),{});
        assert(parts.types.value("d")=="DATETIME" && parts.rows[0][0].toString()=="2024-02-29 00:00:00");
        assert(parts.rows[0][1].toString()=="1899-12-30 14:24:00" && parts.rows[0][2].isNull());
        reversed.execute(db,"SELECT DateDiff('w','2024-01-01','2024-01-09') AS weeks");
        assert(reversed.data(reversed.index(0,0)).toInt()==1);
        const auto summed=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT SUM(stamp-CDate('2024-02-27 12:00:00')) AS elapsed FROM native_dates"),{});
        assert(summed.rows[0][0].toDouble()==1 && summed.types.value("elapsed")=="DOUBLE");
        reversed.execute(db,"SELECT DateAdd('m',1,'2024-01-31') AS shifted,DateAdd('d',1,Null) AS missing");
        assert(reversed.data(reversed.index(0,0)).toString()=="2024-02-29 00:00:00" && reversed.data(reversed.index(0,1)).isNull());
        assert(!query.exec("SELECT linaccess_jet_date_sum(stamp) FROM native_dates WHERE id<0"));
        assert(linaccess::jetSqlFunctionErrorNumber(query.lastError().text().toCaseFolded().toStdString())==3251); query.finish();
        reversed.execute(db,"SELECT Abs((stamp-CDate('2024-02-27 12:00:00'))+1),Year(stamp+1) FROM native_dates WHERE id=1");
        assert(reversed.data(reversed.index(0,0)).toDouble()==2 && reversed.data(reversed.index(0,1)).toInt()==2024);
        const auto clocks=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT Date() AS d,Time() AS t,Now() AS n"),{});
        assert(clocks.types.value("d")=="DATETIME" && clocks.types.value("t")=="DATETIME" && clocks.types.value("n")=="DATETIME");
        reversed.execute(db,"SELECT 1+DateSerial(2024,2,28) AS following");
        assert(reversed.data(reversed.index(0,0)).toString()=="2024-02-29 00:00:00");
        const auto datePlan=linaccess::compileAccessQuery("SELECT DateSerial(2024,3,0) AS d,TimeSerial(6,-15,0) AS t,CDate(2.5) AS converted");
        const auto dates=linaccess::executeAccessQueryPlan(db,datePlan,{});
        assert(dates.rows.size()==1 && dates.types.value("d")=="DATETIME" && dates.types.value("t")=="DATETIME");
        assert(linaccess::qtFieldValue(dates.rows[0][0],dates.types.value("d")).asText()=="2024-02-29");
        assert(linaccess::qtFieldValue(dates.rows[0][1],dates.types.value("t")).asText()=="05:45:00");
        assert(linaccess::qtFieldValue(dates.rows[0][2],dates.types.value("converted")).asNumber()==2.5);
        reversed.execute(db,"SELECT Choose(2,label,'chosen'),Switch(score=10,'ten',True,'other'),Switch(False,1) FROM samples WHERE id=1");
        assert(reversed.data(reversed.index(0,0)).toString()=="chosen" && reversed.data(reversed.index(0,1)).toString()=="ten");
        assert(reversed.data(reversed.index(0,2)).isNull());
        reversed.execute(db,"SELECT DatePart('y','2024-02-29'),DatePart('q','2024-12-31'),DatePart('m',Null)");
        assert(reversed.data(reversed.index(0,0)).toInt()==60 && reversed.data(reversed.index(0,1)).toInt()==4);
        assert(reversed.data(reversed.index(0,2)).isNull());
        reversed.execute(db,"SELECT UCase(label),LCase(UCase(label)),UCase(Null) FROM samples WHERE id=1");
        assert(reversed.data(reversed.index(0,0)).toString()=="ONE" && reversed.data(reversed.index(0,1)).toString()=="one");
        assert(reversed.data(reversed.index(0,2)).isNull());
        reversed.execute(db,"SELECT Val(label),Val(' 1 2.5 tail'),Val('&HFFFF'),Val('1,5') FROM samples WHERE id=1");
        assert(reversed.data(reversed.index(0,0)).toDouble()==0 && reversed.data(reversed.index(0,1)).toDouble()==12.5);
        assert(reversed.data(reversed.index(0,2)).toInt()==-1 && reversed.data(reversed.index(0,3)).toDouble()==1);
        reversed.execute(db,"SELECT Hex(score),Oct(score),Hex(Null) FROM samples WHERE id=1");
        assert(reversed.data(reversed.index(0,0)).toString()=="A" && reversed.data(reversed.index(0,1)).toString()=="12");
        assert(reversed.data(reversed.index(0,2)).isNull());
        reversed.execute(db,"SELECT Space(2),String(3,label),String(2,321),String(2,Null) FROM samples WHERE id=1");
        assert(reversed.data(reversed.index(0,0)).toString()=="  " && reversed.data(reversed.index(0,1)).toString()=="ooo");
        assert(reversed.data(reversed.index(0,2)).toString()=="AA" && reversed.data(reversed.index(0,3)).isNull());
        reversed.execute(db,"SELECT AscW('Ω') AS code,ChrW(937) AS glyph,Chr(34) AS quotation,Chr(Asc(label)) AS first FROM samples WHERE id=1");
        assert(reversed.data(reversed.index(0,0)).toInt()==937 && reversed.data(reversed.index(0,1)).toString()=="Ω");
        assert(reversed.data(reversed.index(0,2)).toString()=="\"" && reversed.data(reversed.index(0,3)).toString()=="o");
        assert(!query.exec("SELECT linaccess_jet_chrw(55296)"));
        assert(linaccess::jetSqlFunctionErrorNumber(query.lastError().text().toCaseFolded().toStdString())==3251); query.finish();
        assert(!query.exec("SELECT linaccess_jet_strreverse(CAST(x'ff' AS TEXT))"));
        assert(linaccess::jetSqlFunctionErrorNumber(query.lastError().text().toCaseFolded().toStdString())==13); query.finish();
    }
    {
        linaccess::AccessQueryModel dates; dates.execute(db,"SELECT DateDiff('d',#2/28/2024#,#3/1/2024#) AS elapsed, Year('2024-02-29') AS yr, Month('2024-02-29') AS mon, Day('2024-02-29') AS dy, Hour('23:59:58') AS hr, Minute('23:59:58') AS mn, Second('23:59:58') AS sec, Weekday('2024-03-03',2) AS wd, Year(Null) AS missing");
        const int expected[]{2,2024,2,29,23,59,58,7}; assert(dates.rowCount()==1);
        for(int i=0;i<8;++i) assert(dates.data(dates.index(0,i)).toInt()==expected[i]); assert(dates.data(dates.index(0,8)).isNull());
        dates.execute(db,"SELECT Year('2024-02-29') AS yr, yr+1 AS following FROM samples WHERE id=1"); assert(dates.data(dates.index(0,1)).toInt()==2025);
    }
    {
        linaccess::AccessQueryModel lengths; lengths.execute(db,"SELECT Len('Grüße😀') AS units, Len(Null) AS missing");
        assert(lengths.rowCount()==1 && lengths.data(lengths.index(0,0)).toInt()==7 && lengths.data(lengths.index(0,1)).isNull());
        lengths.execute(db,"SELECT Len('😀') AS units, units+1 AS next_units FROM samples WHERE id=1");
        assert(lengths.data(lengths.index(0,0)).toInt()==2 && lengths.data(lengths.index(0,1)).toInt()==3);
        assert(query.exec("CREATE TABLE unicode_lengths(value TEXT)"));
        query.prepare("INSERT INTO unicode_lengths VALUES(?)"); query.addBindValue(QString::fromUtf8("A\0😀",6)); assert(query.exec()); query.finish();
        lengths.execute(db,"SELECT Len(value) AS units FROM unicode_lengths");
        assert(lengths.data(lengths.index(0,0)).toInt()==4);
        lengths.execute(db,"SELECT LenB(value) AS bytes,LenB(Null) AS missing FROM unicode_lengths");
        assert(lengths.data(lengths.index(0,0)).toInt()==8 && lengths.data(lengths.index(0,1)).isNull());
        lengths.execute(db,"SELECT LenB('Grüße😀') AS bytes,bytes+1 AS following FROM samples WHERE id=1");
        assert(lengths.data(lengths.index(0,0)).toInt()==14 && lengths.data(lengths.index(0,1)).toInt()==15);

        assert(query.exec("SELECT linaccess_jet_len(char(65279)||'A')") && query.next() && query.value(0).toInt()==2); query.finish();
        assert(!query.exec("SELECT linaccess_jet_len(CAST(x'ff' AS TEXT))"));
        assert(linaccess::jetSqlFunctionErrorNumber(query.lastError().text().toCaseFolded().toStdString())==13); query.finish();

    }

    {
        assert(query.exec("CREATE TABLE color_channels(id INTEGER PRIMARY KEY,r REAL,g REAL,b REAL,[RGB] TEXT)"));
        assert(query.exec("INSERT INTO color_channels VALUES(1,255,0,0,'first'),(2,0,255,0,'second'),(3,0,0,255,'third'),(4,300,127.5,0.5,'rounded')"));
        linaccess::AccessQueryModel colors;
        colors.execute(db,"SELECT [RGB],RGB(r,g,b) AS shade,shade+1 AS next_shade FROM color_channels ORDER BY id");
        assert(colors.rowCount()==4 && colors.columnCount()==3);
        const int expected[]={255,65280,16711680,33023};
        for(int row=0;row<4;++row) {
            assert(colors.data(colors.index(row,1)).toInt()==expected[row]);
            assert(colors.data(colors.index(row,2)).toInt()==expected[row]+1);
        }
        assert(colors.data(colors.index(0,0)).toString()=="first");
        colors.execute(db,"PARAMETERS [red] Double,[green] Double,[blue] Double; SELECT rGb([red],[green],[blue]) AS shade",{{"red",1.0},{"green",2.0},{"blue",3.0}});
        assert(colors.rowCount()==1 && colors.data(colors.index(0,0)).toInt()==197121);
        colors.execute(db,"SELECT RGB(1e100,0,0) AS shade,RGB(False,0,0) AS black");
        assert(colors.data(colors.index(0,0)).toInt()==255 && colors.data(colors.index(0,1)).toInt()==0);
        const auto rejectColor=[&](const QString& sql,int expectedError) {
            int number=0; try { colors.execute(db,sql); } catch(const linaccess::VbaRuntimeError& failure) { number=failure.number; }
            if(number!=expectedError || colors.rowCount()!=0 || colors.columnCount()!=0) std::cerr<<"Synthetic RGB error: expected="<<expectedError<<", actual="<<number<<", rows="<<colors.rowCount()<<", columns="<<colors.columnCount()<<"\n";
            assert(number==expectedError && colors.rowCount()==0 && colors.columnCount()==0);
        };
        rejectColor("SELECT RGB(Null,0,0)",94);
        rejectColor("SELECT RGB('bad',0,0)",13);
        rejectColor("SELECT RGB(CAST(1 AS BLOB),0,0)",13);
        rejectColor("SELECT RGB(True,0,0)",5);
        rejectColor("SELECT RGB(1,2)",450);
        rejectColor("SELECT RGB()",450);
        assert(query.exec("SELECT linaccess_jet_rgb(255,255,255)") && query.next() && query.value(0).toInt()==16777215); query.finish();
        assert(!query.exec("SELECT linaccess_jet_rgb(Null,0,0)"));
        assert(linaccess::jetSqlFunctionErrorNumber(query.lastError().text().toCaseFolded().toStdString())==94); query.finish();
    }

    {
        assert(query.exec("CREATE TABLE legacy_palette(idx INTEGER,[QBColor] TEXT)"));
        for(int index=0;index<16;++index) { query.prepare("INSERT INTO legacy_palette VALUES(?,'entry')"); query.addBindValue(index); assert(query.exec()); query.finish(); }
        linaccess::AccessQueryModel colors;
        colors.execute(db,"SELECT [QBColor],QBColor(idx) AS shade,shade+1 AS next_shade FROM legacy_palette ORDER BY idx");
        assert(colors.rowCount()==16 && colors.data(colors.index(0,0)).toString()=="entry");
        assert(colors.data(colors.index(1,1)).toInt()==0x800000 && colors.data(colors.index(7,1)).toInt()==0xc0c0c0);
        assert(colors.data(colors.index(12,1)).toInt()==255 && colors.data(colors.index(15,1)).toInt()==0xffffff);
        for(int row=0;row<16;++row) assert(colors.data(colors.index(row,2)).toInt()==colors.data(colors.index(row,1)).toInt()+1);
        colors.execute(db,"PARAMETERS [palette] Double; SELECT qBcOlOr([palette]) AS shade",{{"palette",7.5}});
        assert(colors.data(colors.index(0,0)).toInt()==0x808080);
        for(const auto& item:std::vector<std::pair<QString,int>>{{"SELECT QBColor(Null)",94},{"SELECT QBColor('bad')",13},{"SELECT QBColor(CAST(1 AS BLOB))",13},{"SELECT QBColor(-1)",5},{"SELECT QBColor(16)",5},{"SELECT QBColor()",450},{"SELECT QBColor(1,2)",450}}) {
            int number=0; try { colors.execute(db,item.first); } catch(const linaccess::VbaRuntimeError& failure) { number=failure.number; }
            assert(number==item.second && colors.rowCount()==0 && colors.columnCount()==0);
        }
    }

    const auto expensive=QString("WITH RECURSIVE numbers(x) AS (SELECT 1 UNION ALL SELECT x+1 FROM numbers WHERE x<100000) SELECT sum(x) FROM numbers");
    {
        linaccess::SqliteExecutionBudget outer(db,1000,5000);
        { linaccess::SqliteExecutionBudget inner(db,10000000,5000); assert(query.exec("SELECT 1")); query.finish(); }
        assert(!query.exec(expensive) && query.lastError().databaseText().contains("interrupted",Qt::CaseInsensitive)); query.finish();
    }
    {
        linaccess::SqliteExecutionBudget outer(db,1000,5000);
        linaccess::SqliteExecutionBudget inner(db,10000000,5000);
        assert(!query.exec(expensive) && query.lastError().databaseText().contains("interrupted",Qt::CaseInsensitive)); query.finish();
    }
    linaccess::AccessQueryModel model;
    model.execute(db,"SELECT TOP 2 id,score FROM samples ORDER BY score");
    assert(model.rowCount()==3 && model.columnCount()==2);
    assert(model.headerData(1,Qt::Horizontal).toString()=="score");
    assert(model.data(model.index(2,1)).toInt()==20);
    assert(!(model.flags(model.index(0,0)) & Qt::ItemIsEditable));
    model.execute(db,"PARAMETERS [minimum] Long; SELECT id,label FROM samples WHERE score >= [minimum] ORDER BY id",{{"minimum",20}});
    assert(model.rowCount()==3 && model.data(model.index(0,0)).toInt()==2);
    assert(model.data(model.index(1,1)).isNull());
    bool rejected=false;
    try { model.execute(db,"DELETE FROM samples"); } catch(const linaccess::VbaRuntimeError&) { rejected=true; }
    assert(rejected && model.rowCount()==0 && model.columnCount()==0);
    assert(query.exec("SELECT COUNT(*) FROM samples") && query.next() && query.value(0).toInt()==4);
    rejected=false;
    try { model.execute(db,"SELECT * FROM missing_dependency"); } catch(const linaccess::VbaRuntimeError&) { rejected=true; }
    assert(rejected && model.rowCount()==0);
    // Metadata expansion must reject missing transitive dependencies before
    // SQLite prepares nested outer joins (a crash on Ubuntu's SQLite 3.45).
    assert(query.exec("CREATE TEMP VIEW broken_outer AS SELECT a.id, sum(b.score) AS total FROM (samples AS a RIGHT JOIN samples AS b ON a.id=b.id) LEFT JOIN absent_outer AS c ON b.id=c.id GROUP BY a.id"));
    assert(query.exec("CREATE TEMP VIEW broken_wrapper AS SELECT * FROM broken_outer"));
    for (const auto& sql : {"SELECT * FROM broken_outer", "SELECT * FROM broken_wrapper", "SELECT x.id FROM samples AS x LEFT JOIN broken_wrapper AS y ON x.id=y.id"}) {
        int number=0;
        try { model.execute(db,sql); } catch(const linaccess::VbaRuntimeError& error) { number=error.number; }
        assert(number==3078 && model.rowCount()==0);
    }
    assert(query.exec("CREATE TABLE absent_outer(id INTEGER)"));
    model.execute(db,"SELECT * FROM broken_wrapper");
    assert(model.rowCount()==4);
    model.execute(db,"SELECT id FROM samples WHERE id=4");
    assert(model.rowCount()==1 && model.data(model.index(0,0)).toInt()==4);
    model.execute(db,"SELECT IIf(True, 7, 9) AS chosen, IIf(Null,1,2) AS null_condition, IIf(False,1,Null) AS null_result");
    assert(model.data(model.index(0,0)).toInt()==7 && model.data(model.index(0,1)).toInt()==2);
    assert(model.data(model.index(0,2)).isNull());
    model.execute(db,"SELECT IIf(True, 'kept', 9), IIf(False, 'discarded', 12)");
    assert(model.data(model.index(0,0)).metaType().id()==QMetaType::QString);
    assert(model.data(model.index(0,1)).toInt()==12);
    // Access evaluates both IIf arguments. SQLite's built-in IIf short-circuits.
    assert(query.exec("SELECT IIf(1,7,abs(-9223372036854775808))"));
    assert(query.next() && query.value(0).toInt()==7); query.finish();
    rejected = false;
    try { model.execute(db,"SELECT IIf(True,7,abs(-9223372036854775808))"); }
    catch(const linaccess::VbaRuntimeError&) { rejected = true; }
    assert(rejected && model.rowCount()==0);
    model.execute(db,"SELECT IIf('True',1,2), IIf('False',1,2)");
    assert(model.data(model.index(0,0)).toInt()==1 && model.data(model.index(0,1)).toInt()==2);
    rejected = false;
    try { model.execute(db,"SELECT IIf('invalid',1,2)"); }
    catch(const linaccess::VbaRuntimeError&) { rejected = true; }
    assert(rejected);

    assert(query.exec("CREATE TABLE statistical_values(group_id INTEGER,value REAL)"));
    assert(query.exec("INSERT INTO statistical_values VALUES(1,10),(1,20),(1,NULL),(1,30),(2,1000000000001),(2,1000000000002),(2,1000000000003),(3,NULL),(4,7)"));query.finish();
    model.execute(db,"SELECT group_id,Var(value),StDev(value) FROM statistical_values GROUP BY group_id ORDER BY group_id");
    assert(model.rowCount()==4 && model.columnCount()==3);
    assert(model.index(0,1).data().toDouble()==100 && model.index(0,2).data().toDouble()==10);
    assert(model.index(1,1).data().toDouble()==1 && model.index(1,2).data().toDouble()==1);
    for(int row=2;row<4;++row) assert(model.index(row,1).data().isNull() && model.index(row,2).data().isNull());
    model.execute(db,"SELECT Var(value),StDev(value) FROM statistical_values WHERE 1=0");
    assert(model.rowCount()==1 && model.index(0,0).data().isNull() && model.index(0,1).data().isNull());
    model.execute(db,"SELECT Var(DISTINCT value),StDev(DISTINCT value) FROM statistical_values WHERE group_id=1");
    assert(model.index(0,0).data().toDouble()==100 && model.index(0,1).data().toDouble()==10);
    model.execute(db,"PARAMETERS minimum LONG; SELECT Var(value) FROM statistical_values WHERE group_id=1 AND value>[minimum]",{{"minimum",10}});
    assert(model.index(0,0).data().toDouble()==50);
    for(const auto& invalid:QStringList{"SELECT Var(label) FROM samples","SELECT StDev(value) FROM aggregate_blobs","SELECT Var(value) FROM (SELECT -1e308 AS value UNION ALL SELECT 1e308)"}) {
        int error=0;try{model.execute(db,invalid);}catch(const linaccess::VbaRuntimeError& failure){error=failure.number;}
        // A BLOB fixture is added below; the missing table remains a dependency error.
        assert(error==(invalid.contains("aggregate_blobs") ? 3078 : invalid.contains("label") ? 13 : 6));
        assert(model.rowCount()==0);
    }

    assert(query.exec("CREATE TABLE aggregate_values(group_name TEXT,sequence INTEGER,value TEXT)"));
    assert(query.exec("INSERT INTO aggregate_values VALUES('a',1,NULL),('a',2,'last'),('b',1,'first'),('b',2,NULL)"));
    query.finish();
    model.execute(db,"SELECT group_name,First(value),Last(value) FROM (SELECT * FROM aggregate_values ORDER BY sequence) GROUP BY group_name ORDER BY group_name");
    assert(model.rowCount()==2 && model.columnCount()==3);
    assert(model.data(model.index(0,1)).isNull() && model.data(model.index(0,2)).toString()=="last");
    assert(model.data(model.index(1,1)).toString()=="first" && model.data(model.index(1,2)).isNull());
    model.execute(db,"SELECT First(value),Last(value) FROM aggregate_values WHERE 1=0");
    assert(model.rowCount()==1 && model.data(model.index(0,0)).isNull() && model.data(model.index(0,1)).isNull());
    model.execute(db,"SELECT First(9007199254740993),Last(7.5)");
    assert(model.data(model.index(0,0)).toLongLong()==9007199254740993LL);
    assert(model.data(model.index(0,1)).toDouble()==7.5);
    assert(query.exec("CREATE TABLE aggregate_blobs(value BLOB)"));
    query.prepare("INSERT INTO aggregate_blobs VALUES(?)");
    const auto binary = QByteArray::fromHex("0001ff00"); query.addBindValue(binary); assert(query.exec()); query.finish();
    int binaryStatError=0;try{model.execute(db,"SELECT StDev(value) FROM aggregate_blobs");}catch(const linaccess::VbaRuntimeError& failure){binaryStatError=failure.number;}assert(binaryStatError==13 && model.rowCount()==0);
    model.execute(db,"SELECT First(value),Last(value) FROM aggregate_blobs");
    assert(model.data(model.index(0,0)).toByteArray()==binary && model.data(model.index(0,1)).toByteArray()==binary);
    query.clear(); db.close(); assert(db.open());
    model.execute(db,"SELECT IIf(True,5,6)");
    assert(model.data(model.index(0,0)).toInt()==5);
    model.execute(db,"SELECT First(12),Last('restored')");
    assert(model.data(model.index(0,0)).toInt()==12 && model.data(model.index(0,1)).toString()=="restored");

}
