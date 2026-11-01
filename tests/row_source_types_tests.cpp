#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_value_list.hpp"
#include <QApplication>
#include <QComboBox>
#include <QPointer>
#include <QSqlQuery>
#include <QTest>
#include <QTreeView>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

void testRowSourceTypes() {
    using namespace linaccess;
    const auto parsed=parseAccessValueList(QString::fromUtf8(R"LIST(1;"alpha;beta";2;"say ""hi""";3;"Grüße")LIST"),2);
    assert(parsed.rows.size()==3 && parsed.rows[0][0].toString()=="1" && parsed.rows[0][1].toString()=="alpha;beta");
    assert(parsed.rows[1][1].toString()=="say \"hi\"" && parsed.rows[2][1].toString()==QString::fromUtf8("Grüße"));
    const auto blanks=parseAccessValueList(";\"\";last",2);
    assert(blanks.rows.size()==2 && blanks.rows[0][0].isNull() && !blanks.rows[0][1].isNull()
        && blanks.rows[0][1].toString().isEmpty() && blanks.rows[1][1].isNull());
    assert(parseAccessValueList("",2).rows.isEmpty());
    auto db=QSqlDatabase::addDatabase("QSQLITE","row-source-types"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db); assert(query.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,choice INTEGER)"));
    assert(query.exec("INSERT INTO records VALUES(1,1),(2,2)"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="records"
Begin Section
Begin ComboBox
Name ="pick"
ControlSource ="choice"
RowSourceType ="Wertliste"
RowSource ="1;first;2;second"
ColumnCount =2
ColumnWidths ="0;1440"
Width =3000
Height =300
End
End
End
)FORM";
    AccessFormPreview form("value list",layout,{},db); form.show(); form.activateWindow(); QApplication::processEvents();
    auto root=makePreviewFormObject(form); auto control=root->member("Controls",{VbaValue::string("pick")}).object;
    auto* combo=form.findChild<QComboBox*>("access-control-pick"); assert(combo); combo->setFocus();
    QTest::keyClick(combo,Qt::Key_Down,Qt::AltModifier); QApplication::processEvents();
    auto* view=qobject_cast<QTreeView*>(combo->view()); assert(view && view->isVisible() && combo->count()==2 && combo->currentText()=="first");
    QTest::qWait(QApplication::doubleClickInterval()+20);
    const auto point=view->visualRect(combo->model()->index(1,1)).center();
    QTest::mouseMove(view->viewport(),point); QTest::mouseClick(view->viewport(),Qt::LeftButton,Qt::NoModifier,point); QApplication::processEvents();
    assert(control->member("Value",{}).asNumber()==2 && combo->currentText()=="second");
    root->assign("Dirty",VbaValue::logical(false));
    assert(query.exec("SELECT choice FROM records WHERE id=1") && query.next() && query.value(0).toInt()==2); query.finish();
    control->assign("RowSourceType",VbaValue::string("Value List"));
    control->assign("RowSource",VbaValue::string(R"LIST(1;"alpha;beta";2;"say ""hi""")LIST"));
    assert(combo->count()==2 && combo->currentText()=="say \"hi\"" && control->member("Value",{}).asNumber()==2);
    assert(control->member("Column",{VbaValue::numeric(1),VbaValue::numeric(0)}).asText()=="alpha;beta");
    QPointer<QAbstractItemModel> before=combo->model(); const auto source=control->member("RowSource",{}).asText();
    int number=0; try { control->assign("RowSource",VbaValue::string("\"unterminated")); }
    catch (const VbaRuntimeError& failure) { number=failure.number; }
    assert(number==5 && combo->model()==before && combo->currentText()=="say \"hi\"" && control->member("RowSource",{}).asText()==source);
    control->assign("ColumnCount",VbaValue::numeric(1));
    number=0; try { control->assign("RowSource",VbaValue::string(QStringList(10001,"entry").join(';').toStdString())); }
    catch (const VbaRuntimeError& failure) { number=failure.number; }
    assert(number==3035 && combo->count()==4 && control->member("RowSource",{}).asText()==source);
    control->assign("ColumnCount",VbaValue::numeric(2)); assert(combo->count()==2 && form.close());

    const auto dangerous=compileAccessQuery("SELECT IIf(True,7,abs(-9223372036854775808)) AS [danger],id AS [identifier] FROM records");
    assert(accessQueryColumns(db,dangerous)==QStringList({"danger","identifier"}));
    number=0; try { (void)executeAccessQueryPlan(db,dangerous,{}); }
    catch (const VbaRuntimeError& failure) { number=failure.number; }
    assert(number!=0); // Describing the query above must not evaluate its expressions.
    const auto parameters=compileAccessQuery("PARAMETERS [p] Long; SELECT [p] AS [argument],id AS [identifier] FROM records");
    assert(accessQueryColumns(db,parameters)==QStringList({"argument","identifier"}));
    assert(accessQueryColumns(db,compileAccessQuery("SELECT TOP 2 id+0 AS [computed] FROM records ORDER BY id"))==QStringList({"computed"}));
    number=0; try { (void)accessQueryColumns(db,compileAccessQuery("SELECT [absent_field] FROM records")); }
    catch (const VbaRuntimeError& failure) { number=failure.number; }
    assert(number!=0);
    auto fieldsLayout=layout; fieldsLayout.replace("RecordSource =\"records\"",""); fieldsLayout.replace("ControlSource =\"choice\"","");
    fieldsLayout.replace("RowSourceType =\"Wertliste\"","RowSourceType =\"Field List\"");
    fieldsLayout.replace("RowSource =\"1;first;2;second\"","RowSource =\"records\"");
    fieldsLayout.replace("ColumnCount =2","ColumnCount =1"); fieldsLayout.replace("ColumnWidths =\"0;1440\"","ColumnWidths =\"1440\"");
    AccessFormPreview fields("field list",fieldsLayout,{},db); fields.show(); fields.activateWindow(); QApplication::processEvents();
    auto fieldControl=makePreviewFormObject(fields)->member("Controls",{VbaValue::string("pick")}).object;
    fieldControl->member("Dropdown",{}); QApplication::processEvents();
    auto* picker=fields.findChild<QComboBox*>("access-control-pick"); assert(picker && picker->count()==2);
    assert(picker->model()->index(0,0).data().toString()=="id" && picker->model()->index(1,0).data().toString()=="choice"); picker->hidePopup();
    fieldControl->assign("RowSource",VbaValue::string("SELECT IIf(True,7,abs(-9223372036854775808)) AS [danger],id AS [identifier] FROM records"));
    assert(picker->count()==2 && picker->model()->index(0,0).data().toString()=="danger" && picker->model()->index(1,0).data().toString()=="identifier");
    fieldControl->assign("RowSourceType",VbaValue::string("Feldliste"));
    assert(picker->count()==2 && fields.close());
}
