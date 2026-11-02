#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_list_box.hpp"
#include <QApplication>
#include <QComboBox>
#include <QLineEdit>
#include <QTreeView>
#include <QSqlQuery>
#include <QTest>
#include <QTimer>
#include <QStyle>
#include <QStyleOptionComboBox>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>

void testComboAutoPopup();
void testCommandRequery();
void testComboKeyboardRefresh();
void testComboListProperties();
void testNativeListBox();

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
    QPointer<QAbstractItemModel> previousModel=combo->model();
    control->assign("ColumnWidths",VbaValue::string("1440;0;0"));
    assert(!previousModel);
    assert(combo->modelColumn()==0 && combo->currentText()=="20" && view->isColumnHidden(1));
    assert(form.close());
    testCommandRequery();
    testComboAutoPopup();
    testComboKeyboardRefresh();
    testComboListProperties();
    testNativeListBox();
}

void testNativeListBox() {
    using linaccess::VbaValue;
    const auto metadata=linaccess::parseAccessFormLayout(QStringLiteral("Begin Form\nBegin ListBox\nMultiSelect =2\nEnd\nBegin Section\nBegin ListBox\nName =\"inherited\"\nEnd\nBegin ListBox\nName =\"single\"\nMultiSelect =0\nEnd\nEnd\nEnd"));
    assert(metadata.controls.size()==2 && metadata.controls[0].multiSelect==2 && metadata.controls[1].multiSelect==0);
    auto db=QSqlDatabase::addDatabase("QSQLITE","native-list-box"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE choices(id INTEGER,label TEXT,extra REAL)"));
    assert(sql.exec("INSERT INTO choices VALUES(10,'first',NULL),(20,'second',2.5)"));
    assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,choice INTEGER)")); assert(sql.exec("INSERT INTO records VALUES(1,10)"));
    assert(sql.exec("CREATE TABLE events(kind TEXT)"));
    const QByteArray layout="Begin Form\nRecordSource =\"records\"\nBegin Section\nBegin ListBox\nName =\"pick\"\nControlSource =\"choice\"\nRowSource =\"SELECT id,label,extra FROM choices ORDER BY id\"\nColumnCount =3\nColumnWidths =\"0;1440;720\"\nWidth =4000\nHeight =1800\nBeforeUpdate =\"[Event Procedure]\"\nAfterUpdate =\"[Event Procedure]\"\nOnClick =\"[Event Procedure]\"\nEnd\nEnd\nEnd";
    const QByteArray code="Sub pick_BeforeUpdate(Cancel As Integer)\nIf Me.Tag=\"block\" Then Cancel=True\nEnd Sub\nSub pick_AfterUpdate()\nCurrentDb.Execute \"INSERT INTO events VALUES('updated')\"\nEnd Sub\nSub pick_Click()\nCurrentDb.Execute \"INSERT INTO events VALUES('clicked')\"\nEnd Sub";
    AccessFormPreview form("native list",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
    auto* list=dynamic_cast<AccessListBox*>(form.findChild<QTreeView*>("access-control-pick"));
    assert(list && list->isVisible() && !list->isWindow() && !form.findChild<QComboBox*>("access-control-pick"));
    assert(list->rowCount()==2 && list->model()->columnCount()==3 && list->currentIndex().row()==0);
    assert(list->isColumnHidden(0) && list->columnWidth(1)==96 && !QApplication::activePopupWidget());
    auto root=makePreviewFormObject(form),pick=root->objectMember("pick",{}).object;
    assert(pick->member("ListCount",{}).asNumber()==2 && pick->member("Column",{VbaValue::numeric(2),VbaValue::numeric(0)}).kind==VbaValue::Kind::Null);
    const auto click=[&](int row) { QTest::mouseClick(list->viewport(),Qt::LeftButton,Qt::NoModifier,list->visualRect(list->model()->index(row,1)).center()); QApplication::processEvents(); };
    const auto count=[&](const char* kind) { QSqlQuery query(db); query.prepare("SELECT COUNT(*) FROM events WHERE kind=?"); query.addBindValue(kind); assert(query.exec() && query.next()); return query.value(0).toInt(); };
    click(1); assert(pick->member("Value",{}).asNumber()==20 && form.hasUnsavedChanges() && count("updated")==1 && count("clicked")==1);
    assert(sql.exec("SELECT choice FROM records WHERE id=1") && sql.next() && sql.value(0).toInt()==10); sql.finish();
    root->assign("Tag",VbaValue::string("block")); click(0);
    assert(count("updated")==1 && count("clicked")==1);
    int error=0; try { root->assign("Dirty",VbaValue::logical(false)); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==2501);
    root->member("Undo",{}); root->assign("Tag",VbaValue::string("")); click(1); root->assign("Dirty",VbaValue::logical(false));
    assert(sql.exec("SELECT choice FROM records WHERE id=1") && sql.next() && sql.value(0).toInt()==20); sql.finish();
    list->setFocus(); QTest::keyClick(list,Qt::Key_Up); assert(pick->member("Value",{}).asNumber()==10 && form.hasUnsavedChanges()); root->member("Undo",{});
    pick->assign("Locked",VbaValue::logical(true)); assert(pick->member("Locked",{}).truthy()); click(0); QTest::keyClick(list,Qt::Key_Up);
    assert(pick->member("Value",{}).asNumber()==20 && !form.hasUnsavedChanges()); pick->assign("Locked",VbaValue::logical(false));
    QPointer<QAbstractItemModel> previous=list->model(); assert(sql.exec("INSERT INTO choices VALUES(30,'third',3.5)"));
    list->setFocus(); QTest::keyClick(list,Qt::Key_F9);
    assert(!previous && list->rowCount()==3 && pick->member("Value",{}).asNumber()==20 && !form.hasUnsavedChanges());
    previous=list->model(); error=0; try { pick->assign("RowSource",VbaValue::string("SELECT absent FROM choices")); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
    assert(error && list->model()==previous && list->rowCount()==3);
    pick->assign("Value",VbaValue{}); assert(!list->currentIndex().isValid() && list->selectionModel()->selectedRows().isEmpty()); root->member("Undo",{});
    assert(list->currentIndex().row()==1 && !form.hasUnsavedChanges());
    assert(form.close());
    auto unbound=layout; unbound.replace("RecordSource =\"records\"\n",""); unbound.replace("ControlSource =\"choice\"\n","");
    auto failedSource=unbound; failedSource.replace("FROM choices ORDER BY id","FROM deferred_choices ORDER BY id");
    AccessFormPreview recovery("list recovery",failedSource,{},db); recovery.show(); QApplication::setActiveWindow(&recovery); QApplication::processEvents();
    auto* recovering=recovery.findChild<QTreeView*>("access-control-pick"); assert(recovering && recovering->property("la_rows_error").toInt()>0);
    assert(sql.exec("CREATE TABLE deferred_choices(id INTEGER,label TEXT,extra REAL)")); assert(sql.exec("INSERT INTO deferred_choices VALUES(10,'recovered',1)"));
    recovering->setFocus(); QTest::keyClick(recovering,Qt::Key_F9);
    assert(recovering->model()->rowCount()==1 && recovering->property("la_rows_error").toInt()==0 && recovering->toolTip().isEmpty()); assert(recovery.close());
    auto values=unbound; values.replace("SELECT id,label,extra FROM choices ORDER BY id","10;first;1;20;second;2");
    values.replace("Name =\"pick\"","Name =\"pick\"\nRowSourceType =\"Value List\"\nDefaultValue =\"10\"");
    AccessFormPreview valueForm("value list",values,{},db); valueForm.show(); QApplication::setActiveWindow(&valueForm); QApplication::processEvents();
    auto valuePick=makePreviewFormObject(valueForm)->objectMember("pick",{}).object;
    auto* valueList=valueForm.findChild<QTreeView*>("access-control-pick"); assert(valueList && valueList->model()->rowCount()==2 && valueList->currentIndex().row()==0);
    assert(valuePick->member("Column",{VbaValue::numeric(2),VbaValue::numeric(1)}).asText()=="2");
    QTest::mouseClick(valueList->viewport(),Qt::LeftButton,Qt::NoModifier,valueList->visualRect(valueList->model()->index(1,1)).center());
    assert(valuePick->member("Value",{}).asText()=="20" && !valueForm.hasUnsavedChanges()); assert(valueForm.close());
    auto fields=unbound; fields.replace("SELECT id,label,extra FROM choices ORDER BY id","choices");
    fields.replace("ColumnCount =3","RowSourceType =\"Field List\"\nColumnCount =1"); fields.replace("0;1440;720","1440");
    AccessFormPreview fieldForm("field list",fields,{},db); auto fieldPick=makePreviewFormObject(fieldForm)->objectMember("pick",{}).object;
    auto* fieldList=fieldForm.findChild<QTreeView*>("access-control-pick"); assert(fieldList && fieldList->model()->rowCount()==3);
    fieldPick->assign("Value",VbaValue::string("label")); assert(fieldList->currentIndex().row()==1 && fieldPick->member("Value",{}).asText()=="label"); assert(fieldForm.close());
    auto reportLayout=layout; reportLayout.replace("Begin Form","Begin Report");
    AccessFormPreview report("list report",reportLayout,{},db); report.show(); QApplication::setActiveWindow(&report); QApplication::processEvents();
    auto reportPick=makePreviewFormObject(report)->objectMember("pick",{}).object; auto* reportList=report.findChild<QTreeView*>("access-control-pick"); assert(reportList && reportList->currentIndex().row()==1);
    QTest::mouseClick(reportList->viewport(),Qt::LeftButton,Qt::NoModifier,reportList->visualRect(reportList->model()->index(0,1)).center());
    QTest::keyClick(reportList,Qt::Key_Up); assert(reportPick->member("Value",{}).asNumber()==20 && !report.hasUnsavedChanges()); assert(report.close());
    auto multiple=layout; multiple.replace("Begin ListBox\n","Begin ListBox\nMultiSelect =2\n");
    AccessFormPreview unsupported("multiple list",multiple,{},db);
    assert(!unsupported.findChild<QTreeView*>("access-control-pick"));
    auto rejected=makePreviewFormObject(unsupported)->objectMember("pick",{}).object; error=0;
    try { rejected->member("Value",{}); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; } assert(error==3251); assert(unsupported.close());
}

void testComboListProperties() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","combo-list-properties"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE choices(id INTEGER,label TEXT)"));
    assert(sql.exec("INSERT INTO choices VALUES(10,'first'),(20,'second'),(NULL,'none'),(30,'')"));
    assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,choice TEXT)")); assert(sql.exec("INSERT INTO records VALUES(1,10)"));
    const QByteArray layout="Begin Form\nRecordSource =\"records\"\nBegin Section\nBegin ComboBox\nName =\"pick\"\nControlSource =\"choice\"\nRowSource =\"SELECT id,label FROM choices ORDER BY rowid\"\nColumnCount =2\nColumnWidths =\"0;1440\"\nEnd\nBegin TextBox\nName =\"plain\"\nEnd\nEnd\nEnd";
    const QByteArray code="Sub Inspect()\nMe.Tag=CStr(Me.pick.ListIndex)&\":\"&CStr(Me.pick.ItemData(1))\nEnd Sub\nSub RejectWrite()\nMe.pick.ListIndex=1\nEnd Sub";
    AccessFormPreview form("list properties",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
    auto root=makePreviewFormObject(form),pick=root->objectMember("pick",{}).object;
    auto* combo=form.findChild<QComboBox*>("access-control-pick"); assert(combo && combo->count()==0);
    assert(pick->member("ListIndex",{}).asNumber()==0 && combo->count()==4 && !form.hasUnsavedChanges());
    std::vector<VbaValue> none; assert(form.executeProcedure("Inspect",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(root->member("Tag",{}).asText()=="0:20");
    assert(pick->member("ItemData",{VbaValue::numeric(2)}).kind==VbaValue::Kind::Null);
    assert(pick->member("ItemData",{VbaValue::empty()}).asNumber()==10);
    assert(pick->member("ItemData",{VbaValue::numeric(0.5)}).asNumber()==10);
    assert(pick->member("ItemData",{VbaValue::numeric(1.5)}).kind==VbaValue::Kind::Null);
    assert(pick->member("ItemData",{VbaValue::string("1")}).asNumber()==20);
    class Index final:public linaccess::VbaObject {
    public:
        int reads{};
        VbaValue scalarValue() override { ++reads; return VbaValue::numeric(1); }
        VbaValue member(std::string_view,const std::vector<VbaValue>&) override { throw linaccess::VbaRuntimeError(438,"No members"); }
    };
    auto index=std::make_shared<Index>(); assert(pick->member("ItemData",{VbaValue::reference(index)}).asNumber()==20 && index->reads==1);
    const auto error=[&](const char* name,const std::vector<VbaValue>& args,int expected) {
        int actual=0; try { pick->member(name,args); } catch(const linaccess::VbaRuntimeError& failure) { actual=failure.number; } assert(actual==expected);
    };
    error("ListIndex",{VbaValue::numeric(0)},450); error("ItemData",{},450);
    error("ItemData",{VbaValue::numeric(0),VbaValue::numeric(1)},450);
    error("ItemData",{VbaValue{}},94); error("ItemData",{VbaValue::string("bad")},13);
    error("ItemData",{VbaValue::numeric(-1)},5); error("ItemData",{VbaValue::numeric(4)},5);
    error("ItemData",{VbaValue::numeric(2147483648.0)},6);
    for(const auto* name:{"ListIndex","ItemData"}) {
        int actual=0; try { pick->assign(name,VbaValue::numeric(1)); } catch(const linaccess::VbaRuntimeError& failure) { actual=failure.number; } assert(actual==383);
        actual=0; try { root->objectMember("plain",{}).object->member(name,{}); } catch(const linaccess::VbaRuntimeError& failure) { actual=failure.number; } assert(actual==438);
    }
    const auto rejected=form.executeProcedure("RejectWrite",none); assert(rejected.status==linaccess::VbaExecutionStatus::Error && rejected.errorNumber==383);
    pick->assign("Value",VbaValue::numeric(20)); assert(pick->member("ListIndex",{}).asNumber()==1);
    pick->assign("Value",VbaValue{}); assert(pick->member("ListIndex",{}).asNumber()==-1 && combo->currentText().isEmpty());
    pick->assign("Value",VbaValue::numeric(999)); assert(pick->member("ListIndex",{}).asNumber()==-1);
    pick->assign("Value",VbaValue::numeric(10)); root->member("Undo",{});
    combo->lineEdit()->setFocus(); combo->lineEdit()->selectAll(); QTest::keyClicks(combo->lineEdit(),"pending");
    error("ListIndex",{},3251); assert(pick->member("ItemData",{VbaValue::numeric(1)}).asNumber()==20 && combo->currentText()=="pending");
    QTest::keyClick(combo->lineEdit(),Qt::Key_Escape); assert(!form.hasUnsavedChanges());
    pick->assign("BoundColumn",VbaValue::numeric(2)); pick->assign("Value",VbaValue::string("second"));
    assert(pick->member("ListIndex",{}).asNumber()==1 && pick->member("ItemData",{VbaValue::numeric(1)}).asText()=="second");
    const auto empty=pick->member("ItemData",{VbaValue::numeric(3)}); assert(empty.kind==VbaValue::Kind::Text && empty.asText().empty());
    pick->assign("Value",VbaValue::string("")); assert(pick->member("ListIndex",{}).asNumber()==3 && combo->currentText().isEmpty());
    pick->assign("Value",VbaValue{}); assert(pick->member("ListIndex",{}).asNumber()==-1 && combo->currentText().isEmpty());
    pick->assign("BoundColumn",VbaValue::numeric(0)); pick->assign("Value",VbaValue::numeric(1));
    assert(pick->member("ListIndex",{}).asNumber()==1 && pick->member("ItemData",{VbaValue::numeric(1)}).asNumber()==1);
    pick->assign("BoundColumn",VbaValue::numeric(1)); root->member("Undo",{});
    assert(pick->member("ListIndex",{}).asNumber()==0 && !form.hasUnsavedChanges());
    assert(sql.exec("DELETE FROM choices WHERE id=10")); pick->member("Requery",{});
    assert(pick->member("ListIndex",{}).asNumber()==-1 && pick->member("ItemData",{VbaValue::numeric(0)}).asNumber()==20);
    assert(sql.exec("SELECT choice FROM records WHERE id=1") && sql.next() && sql.value(0).toInt()==10); sql.finish();
    assert(form.close());
    auto reportLayout=layout; reportLayout.replace("Begin Form","Begin Report");
    AccessFormPreview report("list report",reportLayout,{},db);
    auto reportPick=makePreviewFormObject(report)->objectMember("pick",{}).object;
    int actual=0; try { reportPick->member("ListIndex",{}); } catch(const linaccess::VbaRuntimeError& failure) { actual=failure.number; } assert(actual==3251);
    assert(reportPick->member("ItemData",{VbaValue::numeric(0)}).asNumber()==20); assert(report.close());
}

void testComboKeyboardRefresh() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","combo-keyboard-refresh"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE choices(id INTEGER,label TEXT)"));
    assert(sql.exec("INSERT INTO choices VALUES(10,'first'),(20,'second')"));
    assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,choice INTEGER)"));
    assert(sql.exec("INSERT INTO records VALUES(1,10)"));
    const QByteArray layout="Begin Form\nRecordSource =\"records\"\nKeyPreview =NotDefault\nOnKeyDown =\"[Event Procedure]\"\nBegin Section\nBegin ComboBox\nName =\"pick\"\nControlSource =\"choice\"\nRowSource =\"SELECT id,label FROM choices ORDER BY id\"\nColumnCount =2\nColumnWidths =\"0;1440\"\nEnd\nEnd\nEnd";
    const QByteArray code="Sub Form_KeyDown(KeyCode As Integer,Shift As Integer)\nIf Me.Tag=\"block\" And (KeyCode=vbKeyF9 Or KeyCode=vbKeyF4) Then KeyCode=0\nEnd Sub";
    AccessFormPreview form("keyboard refresh",layout,code,db); form.show(); QApplication::setActiveWindow(&form); QApplication::processEvents();
    auto root=makePreviewFormObject(form); auto control=root->objectMember("pick",{}).object;
    assert(control->member("Column",{VbaValue::numeric(1),VbaValue::numeric(0)}).asText()=="first");
    auto* combo=form.findChild<QComboBox*>("access-control-pick"); assert(combo);
    auto* editor=combo->lineEdit(); editor->setFocus();
    root->assign("Tag",VbaValue::string("block"));
    QTest::keyClick(editor,Qt::Key_F4); assert(!combo->view()->isVisible());
    assert(sql.exec("UPDATE choices SET label='updated' WHERE id=10"));
    assert(sql.exec("INSERT INTO choices VALUES(30,'third')"));
    QPointer<QAbstractItemModel> previous=combo->model();
    QTest::keyClick(editor,Qt::Key_F9); assert(combo->model()==previous && combo->count()==2 && combo->currentText()=="first");
    root->assign("Tag",VbaValue::string(""));
    QTest::keyClick(editor,Qt::Key_F9);
    assert(!previous && combo->count()==3 && combo->currentText()=="updated" && editor->hasFocus());
    assert(!form.hasUnsavedChanges() && control->member("Value",{}).asNumber()==10);
    QTest::keyClick(editor,Qt::Key_F4); QApplication::processEvents(); assert(combo->view()->isVisible()); combo->hidePopup(); editor->setFocus();
    editor->selectAll(); QTest::keyClicks(editor,"pen");
    assert(editor->isModified() && form.hasUnsavedChanges());
    QTest::keyClick(editor,Qt::Key_F9);
    assert(combo->currentText()=="pen" && editor->isModified() && editor->hasFocus() && form.hasUnsavedChanges());
    previous=combo->model(); assert(sql.exec("DROP TABLE choices"));
    QTest::keyClick(editor,Qt::Key_F9);
    assert(combo->model()==previous && combo->count()==3 && combo->currentText()=="pen" && editor->isModified());
    assert(sql.exec("CREATE TABLE choices(id INTEGER,label TEXT)")); assert(sql.exec("INSERT INTO choices VALUES(10,'recovered')"));
    QTest::keyClick(editor,Qt::Key_F9);
    assert(!previous && combo->count()==1 && combo->currentText()=="pen" && editor->isModified());
    assert(sql.exec("SELECT choice FROM records WHERE id=1") && sql.next() && sql.value(0).toInt()==10); sql.finish();
    QTest::keyClick(editor,Qt::Key_Escape);
    assert(control->member("Value",{}).asNumber()==10 && combo->currentText()=="recovered" && !form.hasUnsavedChanges());
    assert(form.close());
}

void testComboAutoPopup() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","auto-combo-popup"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db); assert(query.exec("CREATE TABLE choices(id INTEGER,label TEXT)"));
    assert(query.exec("INSERT INTO choices VALUES(1,'first'),(2,'second')"));
    assert(query.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,choice INTEGER)"));
    assert(query.exec("INSERT INTO records VALUES(1,1)"));
    const QByteArray layout="Begin Form\nRecordSource =\"records\"\nBegin Section\nBegin ComboBox\nName =\"pick\"\nControlSource =\"choice\"\nRowSource =\"SELECT id,label FROM choices ORDER BY id\"\nColumnCount =2\nColumnWidths =\"0;1440\"\nWidth =3000\nHeight =300\nEnd\nEnd\nEnd\n";
    const auto clickArrow=[](QComboBox* combo) {
        QStyleOptionComboBox option; option.initFrom(combo); option.editable=combo->isEditable();
        const auto arrow=combo->style()->subControlRect(QStyle::CC_ComboBox,&option,QStyle::SC_ComboBoxArrow,combo);
        QTest::mouseClick(combo,Qt::LeftButton,Qt::NoModifier,arrow.center()); QApplication::processEvents();
    };
    AccessFormPreview mouse("mouse popup",layout,{},db); mouse.show(); QApplication::setActiveWindow(&mouse); QApplication::processEvents();
    auto* combo=mouse.findChild<QComboBox*>("access-control-pick"); assert(combo && combo->count()==0);
    clickArrow(combo); assert(combo->count()==2 && combo->view()->isVisible() && combo->currentText()=="first");
    assert(!mouse.hasUnsavedChanges());
    auto* view=qobject_cast<QTreeView*>(combo->view()); assert(view);
    bool timerDelivered=false; QTimer popupTimer; popupTimer.setSingleShot(true);
    QObject::connect(&popupTimer,&QTimer::timeout,[&] { timerDelivered=true; }); popupTimer.start(20);
    QTest::qWait(QApplication::doubleClickInterval()+20); // Qt blocks the first release during its double-click interval.
    assert(timerDelivered);
    QTest::mouseMove(view->viewport(),view->visualRect(combo->model()->index(1,1)).center()); QApplication::processEvents();
    QTest::mouseClick(view->viewport(),Qt::LeftButton,Qt::NoModifier,view->visualRect(combo->model()->index(1,1)).center()); QApplication::processEvents();
    if (makePreviewFormObject(mouse)->item("pick").asNumber()!=2) {
        const auto rect=view->visualRect(combo->model()->index(1,1));
        std::cerr<<"Popup selection diagnostic: row="<<combo->currentIndex()<<", view row="<<view->currentIndex().row()
            <<", viewport="<<view->viewport()->width()<<"x"<<view->viewport()->height()<<", target="<<rect.x()<<","<<rect.y()<<","<<rect.width()<<","<<rect.height()
            <<", visible="<<view->isVisible()<<'\n';
    }
    assert(makePreviewFormObject(mouse)->item("pick").asNumber()==2);
    makePreviewFormObject(mouse)->assign("Dirty",VbaValue::logical(false));
    assert(query.exec("SELECT choice FROM records WHERE id=1") && query.next() && query.value(0).toInt()==2); query.finish();
    assert(mouse.close());
    AccessFormPreview keyboard("keyboard popup",layout,{},db); keyboard.show(); QApplication::setActiveWindow(&keyboard); QApplication::processEvents();
    auto* keys=keyboard.findChild<QComboBox*>("access-control-pick"); assert(keys && keys->count()==0); keys->setFocus();
    QTest::keyClick(keys,Qt::Key_Down,Qt::AltModifier); QApplication::processEvents();
    assert(keys->count()==2 && keys->view()->isVisible() && keys->currentText()=="second"); keys->hidePopup(); assert(keyboard.close());
    AccessFormPreview draft("typed popup",layout,{},db); draft.show(); QApplication::setActiveWindow(&draft); QApplication::processEvents();
    auto* typed=draft.findChild<QComboBox*>("access-control-pick"); assert(typed && typed->count()==0);
    typed->setFocus(); typed->lineEdit()->selectAll(); QTest::keyClicks(typed->lineEdit(),"sec");
    QTest::keyClick(typed,Qt::Key_Down,Qt::AltModifier); QApplication::processEvents();
    assert(typed->count()==2 && typed->view()->isVisible() && typed->currentText()=="sec");
    auto* typedView=qobject_cast<QTreeView*>(typed->view()); assert(typedView);
    QTest::mouseClick(typedView->viewport(),Qt::LeftButton,Qt::NoModifier,typedView->visualRect(typed->model()->index(0,1)).center()); QApplication::processEvents();
    assert(makePreviewFormObject(draft)->item("pick").asNumber()==1 && draft.close());
    assert(query.exec("SELECT choice FROM records WHERE id=1") && query.next() && query.value(0).toInt()==1); query.finish();
    auto failedLayout=layout; failedLayout.replace("RecordSource =\"records\"",""); failedLayout.replace("ControlSource =\"choice\"","");
    failedLayout.replace("SELECT id,label FROM choices ORDER BY id","SELECT id,label FROM pending ORDER BY id");
    AccessFormPreview recovery("popup recovery",failedLayout,{},db); recovery.show(); QApplication::setActiveWindow(&recovery); QApplication::processEvents();
    auto* pending=recovery.findChild<QComboBox*>("access-control-pick"); assert(pending);
    auto model=makePreviewFormObject(recovery); model->assignItem("pick",VbaValue::numeric(42));
    clickArrow(pending);
    assert(pending->count()==0 && !pending->view()->isVisible() && pending->property("la_rows_error").toInt()>0);
    assert(model->item("pick").asNumber()==42);
    assert(query.exec("CREATE TABLE pending(id INTEGER,label TEXT)")); assert(query.exec("INSERT INTO pending VALUES(42,'recovered')"));
    QTest::keyClick(pending,Qt::Key_Down,Qt::AltModifier); QApplication::processEvents();
    assert(pending->count()==1 && pending->view()->isVisible() && pending->currentText()=="recovered");
    assert(model->item("pick").asNumber()==42 && pending->property("la_rows_error").toInt()==0);
    pending->hidePopup(); assert(recovery.close());
}

void testCommandRequery() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","command-requery"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,name TEXT)"));
    assert(sql.exec("INSERT INTO records VALUES(1,'old')"));
    assert(sql.exec("CREATE TABLE choices(id INTEGER,label TEXT)"));
    assert(sql.exec("INSERT INTO choices VALUES(1,'first')"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="records"
BeforeUpdate ="[Event Procedure]"
Begin Section
Begin TextBox
Name ="name"
ControlSource ="name"
End
Begin ComboBox
Name ="pick"
RowSource ="SELECT id,label FROM choices ORDER BY id"
ColumnCount =2
ColumnWidths ="0;1440"
End
End
End
)FORM";
    const QByteArray targetCode=R"VBA(Sub Form_BeforeUpdate(Cancel As Integer)
If Me.Tag="block" Then Cancel=1
End Sub
)VBA";
    auto callerLayout=layout; callerLayout.replace("RecordSource =\"records\"",""); callerLayout.replace("ControlSource =\"name\"","");
    const QByteArray callerCode=R"VBA(Sub NamedControl()
DoCmd.Requery "PiCk"
Me.Tag=Screen.ActiveForm.Name
End Sub
Sub BoundControl()
Application.DoCmd.Requery "name"
End Sub
Sub ExplicitControl()
Me.pick.Requery
End Sub
)VBA";
    AccessFormPreview target("requery target",layout,targetCode,db);
    AccessFormPreview caller("requery caller",callerLayout,callerCode,db);
    auto targetRoot=makePreviewFormObject(target),callerRoot=makePreviewFormObject(caller);
    auto targetPick=targetRoot->objectMember("pick",{}).object,callerPick=callerRoot->objectMember("pick",{}).object;
    const auto label=[](const auto& pick) { return pick->member("Column",{VbaValue::numeric(1),VbaValue::numeric(0)}).asText(); };
    assert(label(targetPick)=="first" && label(callerPick)=="first");
    auto* targetCombo=target.findChild<QComboBox*>("access-control-pick");
    auto* callerCombo=caller.findChild<QComboBox*>("access-control-pick"); assert(targetCombo && callerCombo);
    assert(sql.exec("INSERT INTO choices VALUES(2,'second')"));
    assert(sql.exec("UPDATE choices SET label='updated' WHERE id=1"));
    assert(sql.exec("UPDATE records SET name='external' WHERE id=1"));
    assert(targetRoot->item("name").asText()=="old");
    caller.show(); target.show(); QApplication::setActiveWindow(&target); QApplication::processEvents();
    std::vector<VbaValue> none;
    assert(caller.executeProcedure("NamedControl",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(targetCombo->count()==2 && label(targetPick)=="updated");
    assert(callerCombo->count()==1 && label(callerPick)=="first");
    assert(callerRoot->member("Tag",{}).asText()=="requery target");
    assert(caller.executeProcedure("BoundControl",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(targetRoot->item("name").asText()=="external");
    targetRoot->assign("Tag",VbaValue::string("block"));
    targetRoot->assignItem("name",VbaValue::string("pending"));
    const auto cancelled=caller.executeProcedure("BoundControl",none);
    assert(cancelled.errorNumber==2501 && target.hasUnsavedChanges() && targetRoot->item("name").asText()=="pending");
    assert(sql.exec("SELECT name FROM records WHERE id=1") && sql.next() && sql.value(0).toString()=="external"); sql.finish();
    targetRoot->member("Undo",{});
    assert(caller.executeProcedure("ExplicitControl",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(callerCombo->count()==2 && label(callerPick)=="updated");
    assert(caller.close() && target.close());
}
