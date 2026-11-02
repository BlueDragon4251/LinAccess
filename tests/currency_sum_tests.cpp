#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_query.hpp"
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
    auto db=QSqlDatabase::addDatabase("QSQLITE","currency-sum"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db); assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    assert(sql.exec("CREATE TABLE sum_inputs(groupid INTEGER,amount REAL)")); assert(sql.exec("INSERT INTO sum_inputs VALUES(1,0.1),(1,0.2),(1,-0.3),(2,NULL)"));
    const auto run=[&](const QString& source) { const auto plan=linaccess::compileAccessQuery(source); return linaccess::executeAccessQueryPlan(db,plan,linaccess::bindAccessQueryParameters(plan,{})); };
    const auto exact=run("SELECT Sum(CCur(amount)) AS total FROM sum_inputs WHERE amount IS NOT NULL");
    assert(exact.rows.size()==1 && exact.rows[0][0].toDouble()==0 && exact.types.value("total")=="CURRENCY_REAL");
    const auto grouped=run("SELECT groupid,Sum(CCur(amount)) AS total FROM sum_inputs WHERE amount IS NOT NULL GROUP BY groupid");
    assert(grouped.rows.size()==1 && grouped.rows[0][1].toDouble()==0 && grouped.types.value("total")=="CURRENCY_REAL");
    const auto empty=run("SELECT Sum(CCur(amount)) AS total FROM sum_inputs WHERE 0"); assert(empty.rows.size()==1 && empty.rows[0][0].isNull() && empty.types.value("total")=="CURRENCY_REAL");
    auto definition=linaccess::makeDaoDatabase(db)->member("CreateQueryDef",{VbaValue::string(""),VbaValue::string("PARAMETERS p Money; SELECT Sum(p) AS total FROM sum_inputs")}).object;
    auto parameter=definition->member("Parameters",{VbaValue::string("p")}).object; parameter->assign("Value",VbaValue::currency(1000));
    auto rows=definition->member("OpenRecordset",{}).object; auto field=rows->member("Fields",{VbaValue::string("total")}).object;
    assert(field->member("Type",{}).number==5 && field->member("Value",{}).kind==VbaValue::Kind::Currency && field->member("Value",{}).currencyUnits==4000); rows->member("Close",{});
    parameter->assign("Value",{}); rows=definition->member("OpenRecordset",{}).object; assert(rows->member("Fields",{VbaValue::string("total")}).object->member("Value",{}).kind==VbaValue::Kind::Null); rows->member("Close",{});
    assert(sql.exec("CREATE TABLE sum_large(amount REAL)"));
    sql.prepare("INSERT INTO sum_large VALUES(?)");
    for(const auto units:{5497558138879999LL,5497558138879999LL,-5497558138879999LL}) { sql.bindValue(0,linaccess::currencyToRealStorage(VbaValue::currency(units))); assert(sql.exec()); } sql.finish();
    const auto cancellation=run("SELECT Sum(CCur(amount)) AS total FROM sum_large"); assert(linaccess::currencyFromRealStorage(cancellation.rows[0][0].toDouble()).currencyUnits==5497558138879999LL);
    int error=0; try { (void)run("SELECT Sum(CCur(amount)) FROM sum_large WHERE rowid<3"); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==3251);
    assert(sql.exec("CREATE TABLE sum_overflow(amount REAL)")); assert(db.transaction()); sql.prepare("INSERT INTO sum_overflow VALUES(?)"); sql.addBindValue(linaccess::currencyToRealStorage(VbaValue::currency(5497558138879999LL)));
    for(int i=0;i<2000;++i) assert(sql.exec()); sql.finish(); assert(db.commit());
    error=0; try { (void)run("SELECT Sum(CCur(amount)) FROM sum_overflow"); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==6);
    const QByteArray layout="Begin Report\nRecordSource =\"SELECT Sum(CCur(amount)) AS total FROM sum_inputs WHERE amount IS NOT NULL\"\nBegin Section\nBegin TextBox\nName =\"total\"\nControlSource =\"total\"\nWidth =4000\nEnd\nEnd\nEnd\n";
    AccessFormPreview report("currency_sum",layout,{},db); report.show(); QApplication::processEvents();
    auto model=makePreviewFormObject(report); auto* editor=report.findChild<QLineEdit*>("access-control-total"); assert(editor && editor->isReadOnly() && editor->text()=="0" && model->item("total").kind==VbaValue::Kind::Currency && model->item("total").currencyUnits==0); assert(report.close());
}
