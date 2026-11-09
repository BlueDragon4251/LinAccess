#include "native_project_gui_extensions.hpp"
#include "mainwindow.hpp"
#include "native_standard_module_host.hpp"
#include "native_macro_storage.hpp"
#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_table_view.hpp"
#include "dao_database.hpp"
#include "access_query.hpp"
#include <QComboBox>
#include <QInputDialog>
#include <QPointer>
#include <QAction>
#include <QApplication>
#include <QDialog>
#include <QEvent>
#include <QLineEdit>
#include <QPlainTextEdit>
#include <QPrintPreviewWidget>
#include <QPushButton>
#include <QSqlQuery>
#include <QSqlRecord>
#include <QTabWidget>
#include <QTableWidget>
#include <QTest>
#include <QTimer>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
namespace {
QAction* action(MainWindow& window,const QString& text) {
    for(auto* candidate:window.findChildren<QAction*>()) if(candidate->text()==text) return candidate;
    assert(false);return nullptr;
}
}
void testNativeRepaintObject(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;class PaintProbe final : public QObject {public:int paints{};bool eventFilter(QObject*,QEvent* event) override {if(event->type()==QEvent::Paint) ++paints;return false;}};QSqlQuery sql(db);assert(sql.exec("CREATE TABLE repaint_rows(id INTEGER PRIMARY KEY,n INTEGER)"));assert(sql.exec("INSERT INTO repaint_rows VALUES(1,10),(2,20)"));assert(sql.exec("CREATE TABLE repaint_lookup(id INTEGER PRIMARY KEY,n INTEGER)"));assert(sql.exec("INSERT INTO repaint_lookup VALUES(1,5)"));sql.finish();linaccess::NativeStandardModuleHost host(&window,db,[] {});auto application=host.applicationObject();auto commands=application->member("DoCmd",{}).object;commands->member("OpenTable",{VbaValue::string("repaint_rows")});auto* tabs=window.findChild<QTabWidget*>();auto* view=dynamic_cast<linaccess::AccessTableView*>(tabs->currentWidget()->findChild<QTableView*>());assert(view);auto* model=qobject_cast<linaccess::AccessTableModel*>(view->model());assert(model);const int row=model->index(0,0).data().toInt()==1 ? 0 : 1,column=model->fieldIndex("n");view->setCurrentIndex(model->index(row,column));assert(model->setData(model->index(row,column),77));view->edit(model->index(row,column));QApplication::processEvents();auto* editor=QApplication::focusWidget();assert(editor);auto* editorText=qobject_cast<QLineEdit*>(editor);if(!editorText) editorText=editor->findChild<QLineEdit*>();assert(editorText);editorText->setFocus();QTest::keyClick(editorText,Qt::Key_A,Qt::ControlModifier);QTest::keyClicks(editorText,"78");PaintProbe tablePaint;view->viewport()->installEventFilter(&tablePaint);const auto current=view->currentIndex();const auto beforeDraft=editorText->text();commands->member("RepaintObject",{VbaValue::integer(0),VbaValue::string("repaint_rows")});if(!(tablePaint.paints>0 && view->currentIndex()==current && model->isDirty() && model->index(row,column).data(Qt::EditRole).toInt()==77 && editorText->text()=="78")) qFatal("Repaint table: paint=%d index=%d dirty=%d model=%d draft=%d visible=%d updates=%d before=%d type=%s",tablePaint.paints,view->currentIndex()==current,model->isDirty(),model->index(row,column).data(Qt::EditRole).toInt(),editorText->text().toInt(),view->viewport()->isVisible(),view->viewport()->updatesEnabled(),beforeDraft.toInt(),editor->metaObject()->className());assert(sql.exec("SELECT n FROM repaint_rows WHERE id=1") && sql.next() && sql.value(0).toInt()==10);sql.finish();QTest::keyClick(editor,Qt::Key_Escape);model->revertAll();commands->member("Close",{VbaValue::integer(0),VbaValue::string("repaint_rows"),VbaValue::integer(2)});
    auto database=linaccess::makeDaoDatabase(db);database->member("CreateQueryDef",{VbaValue::string("RepaintQuery"),VbaValue::string("SELECT id,n FROM repaint_rows")});commands->member("OpenQuery",{VbaValue::string("RepaintQuery"),VbaValue::integer(0),VbaValue::integer(2)});auto* query=tabs->currentWidget()->findChild<QTableView*>("native-query-view");assert(query);PaintProbe queryPaint;query->viewport()->installEventFilter(&queryPaint);linaccess::storeNativeMacro(db,"RepaintDefault",{{"version",1},{"actions",QJsonArray{QJsonObject{{"action","RepaintObject"},{"arguments",QJsonArray{}}}}}});commands->member("RunMacro",{VbaValue::string("RepaintDefault")});assert(queryPaint.paints>0);commands->member("Close",{VbaValue::integer(1),VbaValue::string("RepaintQuery")});
    const QByteArray source=R"FORM(Begin Form
RecordSource ="repaint_rows"
Width =2400
Begin Section
Height =1000
Begin TextBox
Name ="Value"
ControlSource ="n"
Width =1800
Height =300
End
Begin TextBox
Name ="Calc"
ControlSource ="=DLookup(""n"",""repaint_lookup"",""id=1"")"
Top =400
Width =1800
Height =300
End
End
End)FORM";AccessFormPreview form("RepaintForm",source,{},db,&window);form.show();QApplication::processEvents();auto root=makePreviewFormObject(form);auto* calculated=form.findChild<QLineEdit*>("access-control-Calc");assert(calculated && calculated->text()=="5");root->member("Controls",{VbaValue::string("Value")}).object->assign("Value",VbaValue::integer(44));assert(sql.exec("UPDATE repaint_lookup SET n=30 WHERE id=1"));assert(calculated->text()=="5");commands->member("RepaintObject",{VbaValue::integer(2),VbaValue::string("RepaintForm")});assert(calculated->text()=="30" && root->member("Dirty",{}).boolean);assert(sql.exec("SELECT n FROM repaint_rows WHERE id=1") && sql.next() && sql.value(0).toInt()==10);sql.finish();assert(sql.exec("UPDATE repaint_lookup SET n=31 WHERE id=1"));root->member("Repaint",{});assert(calculated->text()=="31");assert(sql.exec("UPDATE repaint_lookup SET n=32 WHERE id=1"));makePreviewApplicationObject(form)->member("DoCmd",{}).object->member("RunMacro",{VbaValue::string("RepaintDefault")});assert(calculated->text()=="32" && root->member("Dirty",{}).boolean);root->assign("Dirty",VbaValue::logical(false));assert(form.close());
    sql.prepare("INSERT INTO _la_objects VALUES('report','RepaintReport',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"sql","SELECT id,n FROM repaint_rows"},{"title","Native repaint reference"}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();commands->member("OpenReport",{VbaValue::string("RepaintReport"),VbaValue::integer(2)});AccessFormPreview* report=nullptr;for(const auto& document:AccessFormPreview::projectDocuments(db)) if(document && document->documentName()=="RepaintReport") report=document;assert(report);auto* pages=report->findChild<QPrintPreviewWidget*>("access-report-pages");assert(pages);int renders=0;const auto repaintConnection=QObject::connect(pages,&QPrintPreviewWidget::paintRequested,pages,[&](QPrinter*) {++renders;});commands->member("RepaintObject",{VbaValue::integer(3),VbaValue::string("RepaintReport")});assert(renders==1 && report->findChild<QWidget*>("access-report-page-view")->property("la_report_pages_valid").toBool());QObject::disconnect(repaintConnection);commands->member("Close",{VbaValue::integer(3),VbaValue::string("RepaintReport"),VbaValue::integer(2)});
    QTimer::singleShot(0,&window,[&] {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-module-editor");dialog->findChild<QLineEdit*>("native-module-name")->setText("RepaintEditorModule");auto* text=dialog->findChild<QPlainTextEdit*>("native-module-source");text->setPlainText("Public Sub RepaintReference()\nEnd Sub");QApplication::setActiveWindow(dialog);dialog->activateWindow();text->setFocus();QApplication::processEvents();QTest::keyClick(text,Qt::Key_S,Qt::ControlModifier);QApplication::processEvents();assert(dialog->isVisible() && !dialog->isWindowModified());text->appendPlainText("' pending repaint");const auto draft=text->toPlainText();assert(dialog->isWindowModified());PaintProbe editorPaint;text->viewport()->installEventFilter(&editorPaint);commands->member("RepaintObject",{VbaValue::integer(5),VbaValue::string("RepaintEditorModule")});commands->member("RepaintObject",{});assert(editorPaint.paints>0 && dialog->isWindowModified() && text->toPlainText()==draft);dialog->reject();});action(window,"+ Modul")->trigger();int error=0;try {commands->member("RepaintObject",{VbaValue::integer(2),VbaValue::string("ClosedRepaintTarget")});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==2450);
}


void testNativeSetValueDoubleEvaluation(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;
    QSqlQuery sql(db);assert(sql.exec("CREATE TABLE double_eval_rows(id INTEGER PRIMARY KEY,n INTEGER)"));assert(sql.exec("INSERT INTO double_eval_rows VALUES(1,10)"));assert(sql.exec("CREATE TABLE double_eval_audit(n INTEGER)"));
    const auto source=QByteArray("Begin Form\nRecordSource =\"double_eval_rows\"\nWidth =2400\nBegin Section\nHeight =1000\nBegin TextBox\nName =\"Value\"\nControlSource =\"n\"\nWidth =1800\nHeight =300\nEnd\nEnd\nEnd");
    AccessFormPreview form("DoubleEvaluationForm",source,{},db,&window);form.show();QApplication::processEvents();auto root=makePreviewFormObject(form);auto control=root->member("Controls",{VbaValue::string("Value")}).object;
    sql.prepare("INSERT INTO _la_objects VALUES('module','DoubleEvaluationCode',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba","Public Function GetSetValueExpression() As String\nCurrentDb.Execute \"INSERT INTO double_eval_audit VALUES(1)\"\nGetSetValueExpression=\"Forms!DoubleEvaluationForm!Value+2\"\nEnd Function"}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();
    const auto macro=[&](const QString& name,const QString& expression) {linaccess::storeNativeMacro(db,name,{{"version",1},{"actions",QJsonArray{QJsonObject{{"action","SetValue"},{"arguments",QJsonArray{"Forms!DoubleEvaluationForm!Value",expression}}}}}});};
    macro("DoubleEvaluationNumber","=42");macro("DoubleEvaluationReference","=\"Forms!DoubleEvaluationForm!Value+1\"");macro("DoubleEvaluationFunction","=GetSetValueExpression()");macro("DoubleEvaluationNull","=Null");macro("DoubleEvaluationError","=CVErr(2015)");macro("DoubleEvaluationDivision","=\"1/0\"");macro("DoubleEvaluationSyntax","=\"1+\"");macro("DoubleEvaluationPrefix","=\"=1\"");macro("DoubleEvaluationLength","=String(4097,65)");
    linaccess::NativeStandardModuleHost host(&window,db,[] {});const auto native=host.applicationObject()->member("DoCmd",{}).object;const auto bound=makePreviewApplicationObject(form)->member("DoCmd",{}).object;
    for(const auto& commands:{native,bound}) {
        commands->member("RunMacro",{VbaValue::string("DoubleEvaluationNumber")});assert(control->scalarValue().number==42);commands->member("RunMacro",{VbaValue::string("DoubleEvaluationReference")});assert(control->scalarValue().number==43);commands->member("RunMacro",{VbaValue::string("DoubleEvaluationFunction")});assert(control->scalarValue().number==45);assert(root->member("Dirty",{}).boolean);root->assign("Dirty",VbaValue::logical(false));
        for(const auto& failure:std::vector<std::pair<const char*,int>>{{"DoubleEvaluationNull",94},{"DoubleEvaluationError",2015},{"DoubleEvaluationDivision",11},{"DoubleEvaluationSyntax",3251},{"DoubleEvaluationPrefix",3251},{"DoubleEvaluationLength",5}}) {int error=0;try {commands->member("RunMacro",{VbaValue::string(failure.first)});}catch(const linaccess::VbaRuntimeError& result) {error=result.number;}if(error!=failure.second) qFatal("SetValue double evaluation %s expected=%d actual=%d",failure.first,failure.second,error);assert(control->scalarValue().number==45 && !root->member("Dirty",{}).boolean);}
    }
    assert(sql.exec("SELECT COUNT(*) FROM double_eval_audit") && sql.next() && sql.value(0).toInt()==2);sql.finish();assert(sql.exec("SELECT n FROM double_eval_rows WHERE id=1") && sql.next() && sql.value(0).toInt()==45);sql.finish();assert(form.close());assert(sql.exec("DELETE FROM _la_objects WHERE kind='module' AND name='DoubleEvaluationCode'"));
}

void testNativeFormRecordFields(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;
    QSqlQuery sql(db);assert(sql.exec("CREATE TABLE record_field_rows(id INTEGER PRIMARY KEY AUTOINCREMENT,visible INTEGER,hidden INTEGER DEFAULT 7 CHECK(hidden>=0),note TEXT)"));assert(sql.exec("INSERT INTO record_field_rows(visible,hidden,note) VALUES(10,20,'initial'),(30,40,NULL)"));
    const QByteArray source=R"FORM(Begin Form
RecordSource ="record_field_rows"
Width =2400
Begin Section
Height =1000
Begin TextBox
Name ="VisibleControl"
ControlSource ="visible"
Width =1800
Height =300
End
Begin TextBox
Name ="CalculatedControl"
ControlSource ="=[hidden]+1"
Top =400
Width =1800
Height =300
End
End
End)FORM";
    AccessFormPreview form("RecordFieldsForm",source,{},db,&window);form.show();QApplication::processEvents();auto root=makePreviewFormObject(form);auto field=root->objectItem("hidden").object;auto note=root->objectItem("note").object;auto visible=root->objectItem("visible").object;auto* calculated=form.findChild<QLineEdit*>("access-control-CalculatedControl");assert(field && field->isNativeType("access.formfield") && field->member("Name",{}).asText()=="hidden" && field->scalarValue().number==20 && calculated->text()=="21");assert(visible->isNativeType("access.control") && root->member("Controls",{}).object->member("Count",{}).number==2);
    linaccess::NativeStandardModuleHost host(&window,db,[] {});const auto commands=host.applicationObject()->member("DoCmd",{}).object;
    const auto run=linaccess::executeVbaProcedure(R"VBA(Public Sub ChangeUnboundRecordField()
Forms!RecordFieldsForm!hidden=Forms!RecordFieldsForm!hidden+5
Forms!RecordFieldsForm!note="Änderung"
Forms!RecordFieldsForm!visible=11
End Sub)VBA","ChangeUnboundRecordField",host);if(run.status!=linaccess::VbaExecutionStatus::Completed) qFatal("record field assignment error=%d",run.errorNumber);assert(field->scalarValue().number==25 && note->scalarValue().asText()=="Änderung" && visible->scalarValue().number==11 && calculated->text()=="26" && root->member("Dirty",{}).boolean);assert(sql.exec("SELECT hidden,note,visible FROM record_field_rows WHERE id=1") && sql.next() && sql.value(0).toInt()==20 && sql.value(1).toString()=="initial" && sql.value(2).toInt()==10);sql.finish();root->assign("Dirty",VbaValue::logical(false));assert(sql.exec("SELECT hidden,note,visible FROM record_field_rows WHERE id=1") && sql.next() && sql.value(0).toInt()==25 && sql.value(1).toString()==QString::fromUtf8("Änderung") && sql.value(2).toInt()==11);sql.finish();
    linaccess::storeNativeMacro(db,"RecordFieldMacro",{{"version",1},{"actions",QJsonArray{QJsonObject{{"action","SetValue"},{"arguments",QJsonArray{"Forms!RecordFieldsForm!hidden","Forms!RecordFieldsForm!hidden+1"}}},QJsonObject{{"action","SetValue"},{"arguments",QJsonArray{"Forms!RecordFieldsForm!note","Null"}}}}}});for(const auto& cmd:{commands,makePreviewApplicationObject(form)->member("DoCmd",{}).object}) {cmd->member("RunMacro",{VbaValue::string("RecordFieldMacro")});assert(field->scalarValue().number==26 && note->scalarValue().kind==VbaValue::Kind::Null);root->member("Undo",{});assert(field->scalarValue().number==25 && note->scalarValue().asText()=="Änderung" && !root->member("Dirty",{}).boolean);}
    field->assign("Value",VbaValue::integer(-1));int error=0;try {root->assign("Dirty",VbaValue::logical(false));}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==2501 && root->member("Dirty",{}).boolean && field->scalarValue().number==-1);assert(sql.exec("SELECT hidden FROM record_field_rows WHERE id=1") && sql.next() && sql.value(0).toInt()==25);sql.finish();field->assign("Value",VbaValue::integer(27));root->assign("Dirty",VbaValue::logical(false));
    field->assign("Value",VbaValue::integer(28));auto* visibleEditor=form.findChild<QLineEdit*>("access-control-VisibleControl");assert(visibleEditor);visibleEditor->setFocus();QApplication::processEvents();QTest::keyClick(visibleEditor,Qt::Key_A,Qt::ControlModifier);QTest::keyClicks(visibleEditor,"99");assert(visibleEditor->text()=="99");visible->member("Undo",{});assert(root->member("Dirty",{}).boolean && field->scalarValue().number==28 && visible->scalarValue().number==11);root->member("Undo",{});
    error=0;try {root->objectItem("id").object->assign("Value",VbaValue::integer(99));}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==2448 && !root->member("Dirty",{}).boolean);
    commands->member("GoToRecord",{VbaValue::integer(2),VbaValue::string("RecordFieldsForm"),VbaValue::integer(5)});assert(field->scalarValue().number==7);field->assign("Value",VbaValue::integer(8));note->assign("Value",VbaValue::string("new field row"));root->assign("Dirty",VbaValue::logical(false));assert(sql.exec("SELECT hidden,note FROM record_field_rows WHERE id=3") && sql.next() && sql.value(0).toInt()==8 && sql.value(1).toString()=="new field row");sql.finish();commands->member("GoToRecord",{VbaValue::integer(2),VbaValue::string("RecordFieldsForm"),VbaValue::integer(2)});assert(field->scalarValue().number==27 && note->scalarValue().asText()=="Änderung");
    assert(form.close());error=0;try {field->assign("Value",VbaValue::integer(100));}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==2450);
    assert(sql.exec("CREATE VIEW record_field_readonly AS SELECT hidden,note FROM record_field_rows"));auto readonlySource=source;readonlySource.replace("record_field_rows","record_field_readonly");AccessFormPreview readonly("RecordFieldsReadonly",readonlySource,{},db,&window);auto readonlyRoot=makePreviewFormObject(readonly);assert(readonlyRoot->objectItem("hidden").object->scalarValue().number==27);error=0;try {readonlyRoot->assignItem("hidden",VbaValue::integer(100));}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==2448 && !readonlyRoot->member("Dirty",{}).boolean);assert(readonly.close());
}

void testNativeProjectFilters(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;QSqlQuery sql(db);assert(sql.exec("CREATE TABLE filter_rows(id INTEGER PRIMARY KEY,n INTEGER,label TEXT,stamp DATETIME)"));assert(sql.exec("INSERT INTO filter_rows VALUES(1,10,'Alpha','2003-01-01 00:00:00'),(2,20,'Beta','2004-02-29 00:00:00'),(3,30,NULL,NULL)"));
    const auto saveQuery=[&](const QString& name,const QString& source) {sql.prepare("INSERT INTO _la_objects VALUES('query',?,?)");sql.addBindValue(name);sql.addBindValue(QJsonDocument(QJsonObject{{"sql",source}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();};saveQuery("FilterEditable","SELECT * FROM filter_rows WHERE n>=20");saveQuery("FilterReadonly","SELECT id,n,label FROM filter_rows ORDER BY id");saveQuery("FilterParameterized","PARAMETERS wanted LONG; SELECT id,n FROM filter_rows WHERE n>=[wanted] ORDER BY id");
    linaccess::storeNativeMacro(db,"FilterMacro",{{"version",1},{"actions",QJsonArray{QJsonObject{{"action","ApplyFilter"},{"arguments",QJsonArray{QJsonValue::Null,"n>=20"}}}}}});linaccess::storeNativeMacro(db,"FilterClearMacro",{{"version",1},{"actions",QJsonArray{QJsonObject{{"action","ShowAllRecords"},{"arguments",QJsonArray{}}}}}});
    linaccess::NativeStandardModuleHost host(&window,db,[] {});auto commands=host.applicationObject()->member("DoCmd",{}).object;commands->member("OpenTable",{VbaValue::string("filter_rows")});auto* tabs=window.findChild<QTabWidget*>();assert(tabs);auto* page=tabs->currentWidget();auto* view=page->findChild<QTableView*>();auto* model=qobject_cast<linaccess::AccessTableModel*>(view->model());assert(model && model->rowCount()==3);model->sort(model->fieldIndex("id"),Qt::AscendingOrder);
    const auto apply=[&](const QString& filter) {commands->member("ApplyFilter",{VbaValue::missing(),VbaValue::string(filter.toStdString())});};const auto failure=[&](const QString& filter,int expected) {int error=0;try {apply(filter);}catch(const linaccess::VbaRuntimeError& result) {error=result.number;}if(error!=expected) qFatal("ApplyFilter expected=%d actual=%d",expected,error);};
    apply("label Like 'A*'");assert(model->rowCount()==1 && model->record(0).value("id").toInt()==1);commands->member("ShowAllRecords",{});assert(model->rowCount()==3);commands->member("RunMacro",{VbaValue::string("FilterMacro")});assert(model->rowCount()==2 && model->record(0).value("id").toInt()==2);failure("n=",3075);assert(model->rowCount()==2 && page->property("la_datasheet_filter").toString()=="n>=20");failure("n=20; DELETE FROM filter_rows",3251);failure("n IN (SELECT n FROM filter_rows)",3251);failure(QString(32769,'1'),5);assert(model->rowCount()==2);
    assert(model->setData(model->index(0,model->fieldIndex("n")),21));failure("id=3",2501);assert(model->isDirty() && model->record(0).value("n").toInt()==21);model->revertAll();commands->member("RunMacro",{VbaValue::string("FilterClearMacro")});assert(model->rowCount()==3);
    const auto filterDialog=[&](const QString& value,bool clear,bool cancel) {QTimer::singleShot(0,&window,[&,value,clear,cancel] {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="native-datasheet-filter-dialog");auto* input=dialog->findChild<QLineEdit*>("native-datasheet-filter-expression");assert(input);input->setText(value);if(cancel) dialog->reject();else dialog->findChild<QPushButton*>(clear ? "native-datasheet-filter-clear" : "native-datasheet-filter-apply")->click();});page->findChild<QPushButton*>("native-datasheet-filter-open")->click();};filterDialog("id=3",false,false);assert(model->rowCount()==1 && model->record(0).value("id").toInt()==3);filterDialog("id=1",false,true);assert(model->rowCount()==1 && model->record(0).value("id").toInt()==3);filterDialog({},true,false);assert(model->rowCount()==3);apply("stamp>=#1/1/2004#");assert(model->rowCount()==1 && model->record(0).value("id").toInt()==2);commands->member("Close",{VbaValue::integer(0),VbaValue::string("filter_rows")});
    commands->member("OpenQuery",{VbaValue::string("FilterEditable"),VbaValue::integer(0),VbaValue::integer(1)});page=tabs->currentWidget();view=page->findChild<QTableView*>();model=qobject_cast<linaccess::AccessTableModel*>(view->model());assert(model && model->rowCount()==2);model->sort(model->fieldIndex("id"),Qt::AscendingOrder);apply("id=3");assert(model->rowCount()==1 && model->record(0).value("id").toInt()==3);commands->member("ShowAllRecords",{});assert(model->rowCount()==2 && model->record(0).value("id").toInt()==2);commands->member("Close",{VbaValue::integer(1),VbaValue::string("FilterEditable")});
    commands->member("OpenQuery",{VbaValue::string("FilterReadonly"),VbaValue::integer(0),VbaValue::integer(2)});page=tabs->currentWidget();view=page->findChild<QTableView*>("native-query-view");assert(view && view->model()->rowCount()==3);apply("n>=20");assert(view->model()->rowCount()==2 && view->model()->index(0,0).data().toInt()==2);const auto* previous=view->model();failure("n=",3075);assert(view->model()==previous && view->model()->rowCount()==2);commands->member("Requery",{});assert(view->model()->rowCount()==2 && page->property("la_datasheet_filter").toString()=="n>=20");commands->member("ShowAllRecords",{});assert(view->model()->rowCount()==3);commands->member("Close",{VbaValue::integer(1),VbaValue::string("FilterReadonly")});
    QTimer::singleShot(0,&window,[&] {auto* dialog=qobject_cast<QInputDialog*>(QApplication::activeModalWidget());assert(dialog);dialog->setTextValue("20");dialog->accept();});commands->member("OpenQuery",{VbaValue::string("FilterParameterized"),VbaValue::integer(0),VbaValue::integer(2)});page=tabs->currentWidget();view=page->findChild<QTableView*>("native-query-view");assert(view && view->model()->rowCount()==2);apply("id=3");assert(view->model()->rowCount()==1 && view->model()->index(0,0).data().toInt()==3);commands->member("ShowAllRecords",{});assert(view->model()->rowCount()==2);commands->member("Close",{VbaValue::integer(1),VbaValue::string("FilterParameterized")});
    const QByteArray formSource=R"FORM(Begin Form
RecordSource ="filter_rows"
Width =2400
Begin Section
Height =500
Begin TextBox
Name ="Value"
ControlSource ="n"
Width =1800
Height =300
End
End
End)FORM";
    AccessFormPreview form("FilterForm",formSource,{},db,&window);form.show();QApplication::processEvents();auto root=makePreviewFormObject(form);auto bound=makePreviewApplicationObject(form)->member("DoCmd",{}).object;auto* records=form.findChild<QComboBox*>("access-record-selector");assert(records && records->count()==4);root->objectItem("Value").object->assign("Value",VbaValue::integer(12));bound->member("RunMacro",{VbaValue::string("FilterMacro")});assert(root->member("FilterOn",{}).boolean && records->count()==3 && root->objectItem("Value").object->scalarValue().number==20);assert(sql.exec("SELECT n FROM filter_rows WHERE id=1") && sql.next() && sql.value(0).toInt()==12);sql.finish();bound->member("ShowAllRecords",{});assert(!root->member("FilterOn",{}).boolean && records->count()==4);commands->member("ApplyFilter",{VbaValue::missing(),VbaValue::string("id=3")});assert(records->count()==2 && root->objectItem("Value").object->scalarValue().number==30);assert(form.close());
    auto reportSource=formSource;reportSource.replace("Begin Form","Begin Report");reportSource.replace("RecordSource =\"filter_rows\"","RecordSource =\"filter_rows\"\nOnOpen =\"[Event Procedure]\"");AccessFormPreview report("FilterReport",reportSource,"Private Sub Report_Open(Cancel As Integer)\nDoCmd.ApplyFilter WhereCondition:=\"id=2\"\nEnd Sub",db,&window);assert(report.eventResult("OnOpen") && report.eventResult("OnOpen")->status==linaccess::VbaExecutionStatus::Completed);auto reportRoot=makePreviewFormObject(report);assert(reportRoot->member("FilterOn",{}).boolean && reportRoot->objectItem("Value").object->scalarValue().number==20);report.show();QApplication::processEvents();auto reportCommands=makePreviewApplicationObject(report)->member("DoCmd",{}).object;int reportError=0;try {reportCommands->member("ApplyFilter",{VbaValue::missing(),VbaValue::string("id=1")});}catch(const linaccess::VbaRuntimeError& result) {reportError=result.number;}assert(reportError==2046 && reportRoot->objectItem("Value").object->scalarValue().number==20);assert(report.close());
    window.focusDatabaseWindow();int error=0;try {commands->member("ShowAllRecords",{});}catch(const linaccess::VbaRuntimeError& result) {error=result.number;}assert(error==2046);
}

void testNativeWindowStateCommands(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;QSqlQuery sql(db);assert(sql.exec("CREATE TABLE native_window_rows(id INTEGER PRIMARY KEY,n INTEGER)"));assert(sql.exec("INSERT INTO native_window_rows VALUES(1,10)"));
    const auto macro=[&](const QString& name,const QStringList& commands) {QJsonArray actions;for(const auto& command:commands) actions.append(QJsonObject{{"action",command},{"arguments",QJsonArray{}}});linaccess::storeNativeMacro(db,name,{{"version",1},{"actions",actions}});};macro("NativeWindowSequence",{"Maximize","Minimize","Restore"});macro("NativeWindowMaxRestore",{"Maximize","Restore"});
    class StateProbe final:public QObject {public:QVector<Qt::WindowStates> states;bool saw(Qt::WindowState flag) const {for(const auto flags:states) if(flag==Qt::WindowNoState ? flags==Qt::WindowNoState : flags.testFlag(flag)) return true;return false;}bool eventFilter(QObject* watched,QEvent* event) override {if(event->type()==QEvent::WindowStateChange) if(auto* widget=qobject_cast<QWidget*>(watched)) states.append(widget->windowState() & (Qt::WindowMaximized|Qt::WindowMinimized));return false;}};
    linaccess::NativeStandardModuleHost host(&window,db,[] {});auto commands=host.applicationObject()->member("DoCmd",{}).object;commands->member("OpenTable",{VbaValue::string("native_window_rows")});window.showNormal();QApplication::processEvents();auto* page=window.findChild<QTabWidget*>()->currentWidget();auto* view=page->findChild<QTableView*>();auto* model=qobject_cast<linaccess::AccessTableModel*>(view->model());assert(model && model->setData(model->index(0,model->fieldIndex("n")),77));
    const auto execution=linaccess::executeVbaProcedure("Public Sub NativeMaximize()\nDoCmd.Maximize\nEnd Sub","NativeMaximize",host);assert(execution.status==linaccess::VbaExecutionStatus::Completed && window.isMaximized());commands->member("Minimize",{});assert(window.isMinimized());commands->member("Restore",{});assert(!window.isMaximized() && !window.isMinimized());StateProbe rootStates;window.installEventFilter(&rootStates);commands->member("RunMacro",{VbaValue::string("NativeWindowSequence")});window.removeEventFilter(&rootStates);assert(rootStates.saw(Qt::WindowMaximized) && rootStates.saw(Qt::WindowMinimized) && rootStates.saw(Qt::WindowNoState));assert(model->isDirty() && model->record(0).value("n").toInt()==77);assert(sql.exec("SELECT n FROM native_window_rows") && sql.next() && sql.value(0).toInt()==10);sql.finish();int error=0;try {commands->member("Maximize",{VbaValue::integer(1)});}catch(const linaccess::VbaRuntimeError& result) {error=result.number;}assert(error==450);model->revertAll();commands->member("Close",{VbaValue::integer(0),VbaValue::string("native_window_rows")});
    const QByteArray source=R"FORM(Begin Form
RecordSource ="native_window_rows"
Width =2400
Begin Section
Height =500
Begin TextBox
Name ="Value"
ControlSource ="n"
Width =1800
Height =300
End
End
End)FORM";
    AccessFormPreview form("NativeWindowForm",source,{},db,&window);form.show();QApplication::processEvents();auto root=makePreviewFormObject(form);root->objectItem("Value").object->assign("Value",VbaValue::integer(55));commands->member("Maximize",{});assert(form.isMaximized() && !window.isMaximized());commands->member("Minimize",{});assert(form.isMinimized());commands->member("SelectObject",{VbaValue::integer(2),VbaValue::string("NativeWindowForm"),VbaValue::logical(false)});commands->member("Restore",{});assert(!form.isMinimized() && !form.isMaximized());StateProbe formStates;form.installEventFilter(&formStates);makePreviewApplicationObject(form)->member("DoCmd",{}).object->member("RunMacro",{VbaValue::string("NativeWindowMaxRestore")});form.removeEventFilter(&formStates);assert(formStates.saw(Qt::WindowMaximized) && formStates.saw(Qt::WindowNoState));assert(root->member("Dirty",{}).boolean && root->objectItem("Value").object->scalarValue().number==55);assert(sql.exec("SELECT n FROM native_window_rows") && sql.next() && sql.value(0).toInt()==10);sql.finish();root->member("Undo",{});assert(form.close());
    QTimer::singleShot(0,&window,[&] {auto* editor=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(editor && editor->objectName()=="native-module-editor");auto* text=editor->findChild<QPlainTextEdit*>("native-module-source");text->setPlainText("Public Sub NativeWindowDraft()\nEnd Sub");const auto draft=text->toPlainText();int failure=0;try {commands->member("Maximize",{});}catch(const linaccess::VbaRuntimeError& result) {failure=result.number;}assert(failure==3251 && editor->isWindowModified() && text->toPlainText()==draft);editor->reject();});action(window,"+ Modul")->trigger();window.showNormal();
}

void testNativeMacroContinuations(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;QSqlQuery sql(db);assert(sql.exec("CREATE TABLE macro_continuation_gate(id INTEGER PRIMARY KEY,n INTEGER)"));assert(sql.exec("INSERT INTO macro_continuation_gate VALUES(1,0)"));assert(sql.exec("CREATE TABLE macro_continuation_audit(n INTEGER)"));
    sql.prepare("INSERT INTO _la_objects VALUES('module','MacroContinuationCode',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba","Public Function MacroContinuationGate() As Boolean\nCurrentDb.Execute \"UPDATE macro_continuation_gate SET n=n+1 WHERE id=1\"\nMacroContinuationGate=(DLookup(\"n\",\"macro_continuation_gate\",\"id=1\") Mod 2)=1\nEnd Function"}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();
    QTimer::singleShot(0,&window,[&] {auto* editor=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(editor && editor->objectName()=="native-macro-designer");editor->findChild<QLineEdit*>("native-macro-name")->setText("ContinuationEditor");auto* rows=editor->findChild<QTableWidget*>("native-macro-actions");assert(rows);rows->item(0,1)->setText("INSERT INTO macro_continuation_audit VALUES(11)");rows->item(0,4)->setText("MacroContinuationGate()");editor->findChild<QPushButton*>("native-macro-add")->click();rows->item(1,1)->setText("INSERT INTO macro_continuation_audit VALUES(22)");rows->item(1,4)->setText("...");editor->findChild<QPushButton*>("native-macro-add")->click();rows->item(2,1)->setText("INSERT INTO macro_continuation_audit VALUES(33)");editor->findChild<QPushButton*>("native-macro-save")->click();assert(!editor->isWindowModified());});action(window,"+ Makro")->trigger();
    assert(sql.exec("SELECT definition FROM _la_objects WHERE kind='macro' AND name='ContinuationEditor'") && sql.next());assert(QJsonDocument::fromJson(sql.value(0).toByteArray()).object().value("actions").toArray()[1].toObject().value("condition").toString()=="...");sql.finish();
    const auto reset=[&](int value=0) {assert(sql.exec("DELETE FROM macro_continuation_audit"));assert(sql.exec("UPDATE macro_continuation_gate SET n="+QString::number(value)));};const auto count=[&](int value) {QSqlQuery query(db);assert(query.exec("SELECT COUNT(*) FROM macro_continuation_audit WHERE n="+QString::number(value)) && query.next());return query.value(0).toInt();};const auto gate=[&] {QSqlQuery query(db);assert(query.exec("SELECT n FROM macro_continuation_gate WHERE id=1") && query.next());return query.value(0).toInt();};
    linaccess::NativeStandardModuleHost host(&window,db,[] {});auto commands=host.applicationObject()->member("DoCmd",{}).object;commands->member("RunMacro",{VbaValue::string("ContinuationEditor"),VbaValue::integer(2)});assert(gate()==2 && count(11)==1 && count(22)==1 && count(33)==2);
    const QByteArray source=R"FORM(Begin Form
RecordSource ="macro_continuation_gate"
Width =1800
Begin Section
Height =500
Begin TextBox
Name ="Value"
ControlSource ="n"
Width =1500
Height =300
End
End
End)FORM";
    reset();AccessFormPreview form("MacroContinuationForm",source,{},db,&window);form.show();QApplication::processEvents();auto bound=makePreviewApplicationObject(form)->member("DoCmd",{}).object;bound->member("RunMacro",{VbaValue::string("ContinuationEditor"),VbaValue::integer(2)});assert(gate()==2 && count(11)==1 && count(22)==1 && count(33)==2);assert(form.close());
    const auto actionDefinition=[](const QString& name,const QString& argument,const QString& condition=QString{}) {QJsonObject result{{"action",name},{"arguments",QJsonArray{argument}}};if(!condition.isEmpty()) result.insert("condition",condition);return result;};
    linaccess::storeNativeMacro(db,"ContinuationChild",{{"version",1},{"actions",QJsonArray{actionDefinition("RunSQL","INSERT INTO macro_continuation_audit VALUES(99)","False"),actionDefinition("RunSQL","INSERT INTO macro_continuation_audit VALUES(44)")}}});linaccess::storeNativeMacro(db,"ContinuationParent",{{"version",1},{"actions",QJsonArray{actionDefinition("RunMacro","ContinuationChild","MacroContinuationGate()"),actionDefinition("RunSQL","INSERT INTO macro_continuation_audit VALUES(55)","..."),actionDefinition("RunSQL","INSERT INTO macro_continuation_audit VALUES(66)")}}});reset();commands->member("RunMacro",{VbaValue::string("ContinuationParent")});assert(gate()==1 && count(44)==1 && count(55)==1 && count(66)==1 && count(99)==0);reset(1);commands->member("RunMacro",{VbaValue::string("ContinuationParent")});assert(gate()==2 && count(44)==0 && count(55)==0 && count(66)==1);
    const auto raw=[&](const QString& name,const QJsonObject& definition) {sql.prepare("INSERT INTO _la_objects VALUES('macro',?,?)");sql.addBindValue(name);sql.addBindValue(QJsonDocument(definition).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();};raw("ContinuationOrphan",{{"version",1},{"actions",QJsonArray{actionDefinition("RunSQL","INSERT INTO macro_continuation_audit VALUES(77)"),actionDefinition("RunSQL","INSERT INTO macro_continuation_audit VALUES(88)","...")}}});raw("ContinuationOtherGroup",{{"version",2},{"groups",QJsonArray{QJsonObject{{"name","First"},{"actions",QJsonArray{actionDefinition("RunSQL","INSERT INTO macro_continuation_audit VALUES(77)")}}},QJsonObject{{"name","Broken"},{"actions",QJsonArray{actionDefinition("RunSQL","INSERT INTO macro_continuation_audit VALUES(88)","...")}}}}}});reset();for(const auto* name:{"ContinuationOrphan","ContinuationOtherGroup"}) {int error=0;try {commands->member("RunMacro",{VbaValue::string(name)});}catch(const linaccess::VbaRuntimeError& failure) {error=failure.number;}assert(error==3251 && gate()==0 && count(77)==0 && count(88)==0);}commands->member("RunMacro",{VbaValue::string("ContinuationEditor")});assert(gate()==1 && count(11)==1 && count(22)==1 && count(33)==1);assert(sql.exec("DELETE FROM _la_objects WHERE kind='module' AND name='MacroContinuationCode'"));
}

void testNativeProjectGuiExtensions(MainWindow& window,QSqlDatabase db) {
    testNativeRepaintObject(window,db);
    testNativeSetValueDoubleEvaluation(window,db);
    testNativeFormRecordFields(window,db);
    testNativeProjectFilters(window,db);
    testNativeWindowStateCommands(window,db);
    testNativeMacroContinuations(window,db);
}
