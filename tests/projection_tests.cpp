#include "access_query.hpp"
#include "jet_sql_functions.hpp"
#include "sqlite_execution.hpp"
#include <QCoreApplication>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
int main(int argc,char** argv) {
    QCoreApplication app(argc,argv);
    auto db=QSqlDatabase::addDatabase("QSQLITE"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE a(id INTEGER,flag INTEGER)"));
    assert(query.exec("CREATE TABLE b(id INTEGER,flag INTEGER)"));
    assert(query.exec("CREATE TABLE c(id INTEGER)"));
    assert(query.exec("INSERT INTO a VALUES(1,10)"));
    assert(query.exec("INSERT INTO b VALUES(1,20)"));
    assert(query.exec("INSERT INTO c VALUES(1)"));
    assert(query.exec("CREATE TABLE distinct_parent(k INTEGER)"));
    assert(query.exec("CREATE TABLE distinct_child(k INTEGER)"));
    assert(query.exec("INSERT INTO distinct_parent VALUES(1),(1),(2),(2)"));
    assert(query.exec("INSERT INTO distinct_child VALUES(1),(1)")); query.finish();
    const auto distinctAll=linaccess::compileAccessQuery("SELECT DISTINCTROW * FROM distinct_parent INNER JOIN distinct_child ON distinct_parent.k=distinct_child.k");
    const auto duplicates=linaccess::executeAccessQueryPlan(db,distinctAll,{});
    assert(duplicates.rows.size()==4 && duplicates.columns.size()==2);
    for (const auto& row : duplicates.rows) assert(row[0].toInt()==1 && row[1].toInt()==1);
    const auto outer=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT DISTINCTROW * FROM distinct_parent LEFT JOIN distinct_child ON distinct_parent.k=distinct_child.k WHERE distinct_parent.k=2"),{});
    assert(outer.rows.size()==2);
    for (const auto& row : outer.rows) assert(row[0].toInt()==2 && row[1].isNull());
    const auto derived=linaccess::compileAccessQuery("PARAMETERS [wanted] Value; SELECT DISTINCTROW * FROM [SELECT distinct_parent.k AS k FROM distinct_parent INNER JOIN distinct_child ON distinct_parent.k=distinct_child.k]. AS picked WHERE picked.k=[wanted]");
    assert(derived.parameters.size()==1);
    const auto picked=linaccess::executeAccessQueryPlan(db,derived,linaccess::bindAccessQueryParameters(derived,{{"wanted",1}}));
    assert(picked.rows.size()==4 && picked.columns.size()==1);
    const auto none=linaccess::executeAccessQueryPlan(db,derived,linaccess::bindAccessQueryParameters(derived,{{"wanted",2}}));
    assert(none.rows.empty());
    const auto nested=linaccess::compileAccessQuery("SELECT * FROM [SELECT * FROM [SELECT 'bracket ] and [ text' AS label]. AS inside_query]. AS outside_query");
    const auto nestedRows=linaccess::executeAccessQueryPlan(db,nested,{});
    assert(nestedRows.rows.size()==1 && nestedRows.rows.front().front().toString()=="bracket ] and [ text");
    const auto ordinary=linaccess::executeAccessQueryPlan(db,linaccess::compileAccessQuery("SELECT 1 AS [SELECT ordinary identifier]"),{});
    assert(ordinary.columns.front()=="SELECT ordinary identifier");
    bool identityMappingRequired=false;
    try { (void)linaccess::compileAccessQuery("SELECT DISTINCTROW distinct_parent.k FROM distinct_parent INNER JOIN distinct_child ON distinct_parent.k=distinct_child.k"); }
    catch(const linaccess::VbaRuntimeError& error) { identityMappingRequired=error.number==3085; }
    assert(identityMappingRequired);
    assert(query.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    const auto artifact=[&](const char* name,const char* sql) {
        query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('query',?,?)");
        query.addBindValue(name); query.addBindValue(QByteArray(sql)); assert(query.exec()); query.finish();
    };
    artifact("joined","SELECT b.id, b.flag FROM c RIGHT JOIN (a LEFT JOIN b ON a.id=b.id) ON c.id=a.id");
    artifact("consumer","SELECT joined.id, joined.flag FROM joined");
    assert(linaccess::installAccessQueryViews(db)==2);
    linaccess::AccessQueryModel model;
    model.execute(db,"SELECT consumer.id,consumer.flag FROM consumer");
    assert(model.rowCount()==1 && model.columnCount()==2);
    assert(model.headerData(0,Qt::Horizontal,Qt::DisplayRole).toString()=="id");
    assert(model.headerData(1,Qt::Horizontal,Qt::DisplayRole).toString()=="flag");
    assert(model.data(model.index(0,0)).toInt()==1 && model.data(model.index(0,1)).toInt()==20);
    model.execute(db,"SELECT b.id AS explicit_label,b.flag FROM b");
    assert(model.headerData(0,Qt::Horizontal,Qt::DisplayRole).toString()=="explicit_label");
    model.execute(db,"SELECT DISTINCT b.id,b.flag FROM b");
    assert(model.headerData(0,Qt::Horizontal,Qt::DisplayRole).toString()=="id");
    model.execute(db,"SELECT TOP 1 b.id,b.flag FROM b ORDER BY b.id");
    assert(model.columnCount()==2 && model.rowCount()==1 && model.data(model.index(0,1)).toInt()==20);
    assert(query.exec("INSERT INTO b VALUES(2,20),(3,20),(4,30)")); query.finish();
    artifact("best","SELECT TOP 2 b.id,b.flag FROM b ORDER BY b.flag");
    artifact("best_consumer","SELECT best.id,best.flag FROM best ORDER BY id");
    artifact("aaa_latest","SELECT TOP 1 id,flag FROM best ORDER BY id DESC");
    assert(linaccess::installAccessQueryViews(db)==5);
    assert(linaccess::installAccessQueryViews(db)==5);
    model.execute(db,"SELECT * FROM best_consumer");
    assert(model.rowCount()==3 && model.columnCount()==2);
    assert(model.data(model.index(0,0)).toInt()==1 && model.data(model.index(2,0)).toInt()==3);
    model.execute(db,"SELECT * FROM aaa_latest");
    assert(model.rowCount()==1 && model.columnCount()==2 && model.data(model.index(0,0)).toInt()==3);
    artifact("rank collision","SELECT TOP 1 id+100 AS __la_top_rank,flag FROM b ORDER BY id");
    assert(linaccess::installAccessQueryViews(db)==6);
    model.execute(db,"SELECT __la_top_rank FROM [rank collision]");
    assert(model.rowCount()==1 && model.columnCount()==1 && model.data(model.index(0,0)).toInt()==101);
    artifact("delayed","SELECT TOP 1 id FROM later ORDER BY id");
    assert(linaccess::installAccessQueryViews(db)==6);
    assert(query.exec("CREATE TABLE later(id INTEGER)")); assert(query.exec("INSERT INTO later VALUES(7)")); query.finish();
    assert(linaccess::installAccessQueryViews(db)==7);
    model.execute(db,"SELECT id FROM delayed"); assert(model.rowCount()==1 && model.data(model.index(0,0)).toInt()==7);
    int calls=0;
    linaccess::setJetSqlObjectResolver(db,[&](auto,auto,auto) { ++calls; return linaccess::VbaValue::numeric(2003); });
    artifact("no_evaluation","SELECT TOP 1 Forms!filter!year AS year FROM b ORDER BY id");
    assert(linaccess::installAccessQueryViews(db)==8 && calls==0);
    model.execute(db,"SELECT year FROM no_evaluation");
    assert(model.rowCount()==1 && model.data(model.index(0,0)).toInt()==2003 && calls>0);
    assert(query.exec("CREATE TABLE ordering(id INTEGER PRIMARY KEY,score INTEGER,enabled INTEGER)"));
    assert(query.exec("INSERT INTO ordering VALUES(1,NULL,0),(2,NULL,0),(3,7,1),(4,7,1),(5,9,0),(6,8,1)")); query.finish();
    model.execute(db,"SELECT TOP 1 id,score FROM ordering ORDER BY score");
    assert(model.rowCount()==2 && model.data(model.index(0,1)).isNull() && model.data(model.index(1,1)).isNull());
    model.execute(db,"SELECT TOP 1 id,score FROM ordering ORDER BY score DESC");
    assert(model.rowCount()==1 && model.data(model.index(0,0)).toInt()==5);
    model.execute(db,"SELECT TOP 1 id,score FROM ordering WHERE enabled=1 OR id=5 ORDER BY score");
    assert(model.rowCount()==2);
    model.execute(db,"PARAMETERS p Long; SELECT TOP 1 id,score AS ranked FROM ordering WHERE enabled=[p] ORDER BY ranked",{{"p",1}});
    assert(model.rowCount()==2 && model.columnCount()==2 && model.data(model.index(0,1)).toInt()==7);
    assert(query.exec("CREATE TABLE big(id INTEGER PRIMARY KEY,flag INTEGER)"));
    assert(query.exec("WITH RECURSIVE n(x) AS (SELECT 1 UNION ALL SELECT x+1 FROM n WHERE x<20000) INSERT INTO big SELECT x,0 FROM n"));
    assert(query.exec("CREATE VIEW big_view AS SELECT id,flag FROM big")); query.finish();
    {
        linaccess::SqliteExecutionBudget outer(db,10000,5000);
        model.execute(db,"SELECT TOP 1 id,flag FROM big_view WHERE (((flag)=0)) ORDER BY id");
        assert(model.rowCount()==1 && model.data(model.index(0,0)).toInt()==1);
    }
    calls=0;
    linaccess::setJetSqlObjectResolver(db,[&](auto,auto,auto) { return linaccess::VbaValue::numeric(++calls); });
    assert(query.exec("CREATE VIEW unstable AS SELECT id,linaccess_jet_object_value('forms','filter','year') AS token FROM b")); query.finish();
    model.execute(db,"SELECT TOP 1 id,token FROM unstable ORDER BY token");
    assert(model.rowCount()==1 && calls>0);
    assert(query.exec("CREATE TABLE left_side(k INTEGER,label TEXT)"));
    assert(query.exec("CREATE TABLE right_side(k INTEGER,name TEXT)"));
    assert(query.exec("INSERT INTO left_side VALUES(1,'a'),(1,'b'),(NULL,'left null')"));
    assert(query.exec("INSERT INTO right_side VALUES(1,'matched'),(2,'unmatched'),(NULL,'right null')")); query.finish();
    const auto joined=QString("SELECT left_side.label,right_side.name,IIf(left_side.k IS NULL,0,1) AS matched FROM left_side RIGHT JOIN right_side ON left_side.k=right_side.k ORDER BY right_side.k,left_side.label");
    const auto joinedPlan=linaccess::compileAccessQuery(joined);
    assert(!joinedPlan.rightJoinSql.isEmpty() && joinedPlan.rightJoinSources.size()==2);
    model.execute(db,joined); assert(model.rowCount()==4 && model.columnCount()==3);
    assert(query.exec("SELECT left_side.label,right_side.name,CASE WHEN left_side.k IS NULL THEN 0 ELSE 1 END FROM left_side RIGHT JOIN right_side ON left_side.k=right_side.k ORDER BY right_side.k,left_side.label"));
    int row=0;
    while (query.next()) {
        for (int column=0;column<3;++column) {
            const auto actual=model.data(model.index(row,column)),expected=query.value(column);
            assert(actual.isNull()==expected.isNull());
            if (!actual.isNull()) assert(actual.toString()==expected.toString());
        }
        ++row;
    }
    assert(row==4); query.finish();
    assert(linaccess::compileAccessQuery("SELECT * FROM left_side RIGHT JOIN right_side ON left_side.k=right_side.k").rightJoinSql.isEmpty());
    artifact("lineage_source","SELECT a.id AS original_key,a.flag AS amount FROM a");
    artifact("lineage_consumer","SELECT a.id,a.flag FROM lineage_source");
    artifact("pure_aliases","SELECT a.flag AS first_value,first_value+1 AS second_value,second_value & '!' AS text_value FROM lineage_source");
    artifact("repeated_origins","SELECT a.id AS one,a.id AS two FROM a");
    assert(linaccess::installAccessQueryViews(db)>0);
    model.execute(db,"SELECT * FROM lineage_consumer");
    assert(model.rowCount()==1 && model.data(model.index(0,0)).toInt()==1 && model.data(model.index(0,1)).toInt()==10);
    model.execute(db,"SELECT * FROM pure_aliases");
    assert(model.rowCount()==1 && model.data(model.index(0,1)).toInt()==11 && model.data(model.index(0,2)).toString()=="11!");
    model.execute(db,"SELECT a.flag FROM lineage_source AS saved INNER JOIN c ON a.id=c.id");
    assert(model.rowCount()==1 && model.data(model.index(0,0)).toInt()==10);
    bool ambiguousOrigin=false;
    try { model.execute(db,"SELECT a.id FROM repeated_origins"); }
    catch(const linaccess::VbaRuntimeError& error) { ambiguousOrigin=error.number==3146; }
    assert(ambiguousOrigin);
    calls=0;
    bool volatileAlias=false;
    try { model.execute(db,"SELECT Forms!Filter!Year AS x,x+1 FROM a"); }
    catch(const linaccess::VbaRuntimeError& error) { volatileAlias=error.number==3146; }
    assert(volatileAlias && calls==0);

}
