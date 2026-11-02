#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
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
