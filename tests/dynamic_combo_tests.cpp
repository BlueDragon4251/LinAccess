#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include <QComboBox>
#include <QTreeView>
#include <QSqlQuery>
#include <QTest>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

void testDynamicCombos() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","dynamic-combos"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE choices(id INTEGER,label TEXT,extra REAL)"));
    assert(query.exec("INSERT INTO choices VALUES(10,'first',NULL),(20,'second',2.5)"));
    assert(query.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,choice INTEGER)"));
    assert(query.exec("INSERT INTO records VALUES(1,10)"));
    const QByteArray layout="Begin Form\nRecordSource =\"records\"\nBegin Section\nBegin ComboBox\nName =\"pick\"\nControlSource =\"choice\"\nWidth =3000\nHeight =300\nEnd\nEnd\nEnd\n";
    const QByteArray code=R"VBA(Sub Configure()
Me.pick.ColumnCount=3
Me.pick.BoundColumn=1
Me.pick.RowSourceType="Table/Query"
Me.pick.ColumnWidths="0;2.54 cm;720"
Me.pick.RowSource="SELECT id,label,extra FROM choices ORDER BY id"
End Sub
)VBA";
    AccessFormPreview form("dynamic",layout,code,db); form.show(); QApplication::processEvents();
    std::vector<VbaValue> none; assert(form.executeProcedure("Configure",none).status==linaccess::VbaExecutionStatus::Completed);
    auto root=makePreviewFormObject(form); auto control=root->objectMember("pick",{}).object;
    auto* combo=form.findChild<QComboBox*>("access-control-pick"); assert(combo);
    auto* view=qobject_cast<QTreeView*>(combo->view()); assert(view);
    assert(combo->count()==2 && combo->model()->columnCount()==3 && combo->modelColumn()==1);
    assert(view->isColumnHidden(0) && !view->isColumnHidden(1) && view->columnWidth(1)==96);
    assert(combo->currentText()=="first" && control->member("Value",{}).asNumber()==10);
    assert(control->member("Column",{VbaValue::numeric(2),VbaValue::numeric(0)}).kind==VbaValue::Kind::Null);
    assert(control->member("Column",{VbaValue::numeric(2),VbaValue::numeric(1)}).asNumber()==2.5);
    assert(control->member("ColumnCount",{}).asNumber()==3 && control->member("BoundColumn",{}).asNumber()==1);
    assert(control->member("RowSourceType",{}).asText()=="Table/Query");
    assert(control->member("ColumnWidths",{}).asText()=="0;2.54 cm;720");
    control->member("Dropdown",{}); QApplication::processEvents();
    const auto second=combo->model()->index(1,1);
    QTest::mouseClick(view->viewport(),Qt::LeftButton,Qt::NoModifier,view->visualRect(second).center()); QApplication::processEvents();
    assert(control->member("Value",{}).asNumber()==20 && combo->currentText()=="second");
    root->assign("Dirty",VbaValue::logical(false));
    assert(query.exec("SELECT choice FROM records WHERE id=1") && query.next() && query.value(0).toInt()==20); query.finish();
    const auto originalSource=control->member("RowSource",{}).asText();
    for (const auto& [property,value]:std::vector<std::pair<std::string,VbaValue>>{
        {"ColumnCount",VbaValue::numeric(0)},{"ColumnCount",VbaValue::numeric(256)},
        {"BoundColumn",VbaValue::numeric(4)},{"RowSource",VbaValue::string("SELECT absent FROM choices")},
        {"RowSourceType",VbaValue::string("unsupported callback")},{"ColumnWidths",VbaValue::string("-1")},
        {"ColumnWidths",VbaValue{}},{"BoundColumn",VbaValue{}}}) {
        bool rejected=false; try { control->assign(property,value); } catch(const linaccess::VbaRuntimeError&) { rejected=true; } assert(rejected);
        assert(combo->count()==2 && combo->modelColumn()==1 && combo->currentText()=="second");
        assert(control->member("Value",{}).asNumber()==20 && control->member("RowSource",{}).asText()==originalSource);
    }
    control->assign("RowSource",VbaValue::string("SELECT id,label,extra FROM choices WHERE id=20"));
    assert(combo->count()==1 && combo->currentText()=="second" && control->member("Value",{}).asNumber()==20);
    control->assign("BoundColumn",VbaValue::numeric(0)); control->assign("Value",VbaValue::numeric(0));
    assert(combo->currentText()=="second" && control->member("Value",{}).asNumber()==0);
    control->assign("ColumnWidths",VbaValue::string("1440;0;0"));
    assert(combo->modelColumn()==0 && combo->currentText()=="20" && view->isColumnHidden(1));
    assert(form.close());
}
