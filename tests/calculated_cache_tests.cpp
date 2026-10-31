#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "vba_expression.hpp"
#include "vba_currency.hpp"
#include <QApplication>
#include <QComboBox>
#include <QLineEdit>
#include <QSqlQuery>
#include <QTest>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

void testCalculatedCaching() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","calculated-cache"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db); assert(query.exec("CREATE TABLE items(id INTEGER PRIMARY KEY,label TEXT)"));
    assert(query.exec("INSERT INTO items VALUES(1,'one')"));
    const QByteArray definition=R"FORM(Begin Form
Begin Section
Begin TextBox
Name ="counter"
ControlSource ="=DCount(\"*\",\"items\")"
End
Begin TextBox
Name ="amount"
Top =400
End
Begin TextBox
Name ="doubled"
ControlSource ="=[amount]*2"
Top =800
End
Begin TextBox
Name ="cash"
ControlSource ="=CCur([amount])"
Top =1200
End
End
End
)FORM";
    const QByteArray code="Sub RefreshCount()\nMe.counter.Requery\nEnd Sub\n";
    AccessFormPreview form("counter cache",definition,code,db); auto object=makePreviewFormObject(form);
    assert(object->item("counter").asNumber()==1);
    assert(query.exec("INSERT INTO items VALUES(2,'two')"));
    assert(object->item("counter").asNumber()==1);
    std::vector<VbaValue> none; assert(form.executeProcedure("RefreshCount",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(object->item("counter").asNumber()==2);
    object->assignItem("amount",linaccess::convertCurrency(VbaValue::string("1.2345")));
    assert(object->item("cash").kind==VbaValue::Kind::Currency && object->item("cash").currencyUnits==12345);
    form.show(); QApplication::setActiveWindow(&form); auto* amount=form.findChild<QLineEdit*>("access-control-amount"); assert(amount);
    amount->setFocus(); amount->selectAll(); QTest::keyClicks(amount,"2.5");
    assert(object->item("doubled").asNumber()==5 && object->item("cash").currencyUnits==25000);
    assert(form.close());

    assert(query.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    const QByteArray child=R"FORM(Begin Form
RecordSource ="SELECT id,label FROM items WHERE Forms!parent!selected=-1 OR id=Forms!parent!selected"
Begin Section
Begin TextBox
Name ="label"
ControlSource ="label"
End
End
End
)FORM";
    query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('form','detailform',?)"); query.addBindValue(child); assert(query.exec()); query.finish();
    const QByteArray parent=R"FORM(Begin Form
Begin Section
Begin TextBox
Name ="selected"
End
Begin Subform
Name ="detail"
SourceObject ="Form.detailform"
Top =500
End
Begin TextBox
Name ="copied"
ControlSource ="=[detail].[Form]![label]"
End
End
End
)FORM";
    AccessFormPreview owner("parent",parent,"Sub RefreshCopy()\nMe.copied.Requery\nEnd Sub\n",db);
    auto model=makePreviewFormObject(owner);
    auto detail=model->member("Controls",{VbaValue::string("detail")}).object->member("Form",{}).object;
    auto* embedded=dynamic_cast<AccessFormPreview*>(owner.findChild<QWidget*>("access-control-detail")); assert(embedded);
    auto* selector=embedded->findChild<QComboBox*>("access-record-selector",Qt::FindDirectChildrenOnly); assert(selector);
    model->assignItem("selected",VbaValue::numeric(1)); detail->member("Requery",{});
    assert(selector->count()==2); selector->setCurrentIndex(1);
    assert(owner.executeProcedure("RefreshCopy",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(model->item("copied").asText()=="one");
    model->assignItem("selected",VbaValue::numeric(0)); detail->member("Requery",{});
    assert(model->item("copied").asText()=="one");
    assert(owner.executeProcedure("RefreshCopy",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(model->item("copied").kind==VbaValue::Kind::Null);
    model->assignItem("selected",VbaValue::numeric(-1)); detail->member("Requery",{});
    assert(selector->count()==3);
    selector->setCurrentIndex(2); assert(model->item("copied").asText()=="two");
    assert(owner.close());
}
