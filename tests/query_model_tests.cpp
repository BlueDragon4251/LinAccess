#include "access_query.hpp"
#include "vba_runtime.hpp"
#include "sqlite_execution.hpp"
#include "jet_sql_functions.hpp"
#include <QCoreApplication>
#include <QSqlQuery>
#include <QSqlError>
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
        linaccess::AccessQueryModel reversed;
        reversed.execute(db,"SELECT InStrRev('abcabc','bc') AS pos,StrReverse('abö') AS backwards, InStrRev('A😀x😀','😀') AS units");
        assert(reversed.rowCount()==1 && reversed.data(reversed.index(0,0)).toInt()==5);
        assert(reversed.data(reversed.index(0,1)).toString()=="öba" && reversed.data(reversed.index(0,2)).toInt()==5);
        reversed.execute(db,"SELECT InStrRev(label,'o') AS pos,pos+1 AS following FROM samples WHERE id=1");
        assert(reversed.data(reversed.index(0,0)).toInt()==1 && reversed.data(reversed.index(0,1)).toInt()==2);
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
    model.execute(db,"SELECT First(value),Last(value) FROM aggregate_blobs");
    assert(model.data(model.index(0,0)).toByteArray()==binary && model.data(model.index(0,1)).toByteArray()==binary);
    query.clear(); db.close(); assert(db.open());
    model.execute(db,"SELECT IIf(True,5,6)");
    assert(model.data(model.index(0,0)).toInt()==5);
    model.execute(db,"SELECT First(12),Last('restored')");
    assert(model.data(model.index(0,0)).toInt()==12 && model.data(model.index(0,1)).toString()=="restored");

}
