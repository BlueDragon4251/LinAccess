#include "execution_capabilities.hpp"
#include "native_project_gui_extensions.hpp"
#include "mainwindow.hpp"
#include "native_standard_module_host.hpp"
#include "native_macro_editor_draft.hpp"
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
#include <QTreeWidget>
#include <QTextBlock>
#include <QTextBrowser>
#include <QLabel>
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

void testNativeMoveSize(MainWindow& window,QSqlDatabase database) {
    using linaccess::VbaValue;
    const auto definition=[](const QJsonArray& arguments) {
        return QJsonObject{{"version",1},{"actions",QJsonArray{
            QJsonObject{{"action","MoveSize"},{"arguments",arguments}}}}};
    };
    const auto errorFrom=[](const auto& operation) {
        try {operation();} catch(const linaccess::VbaRuntimeError& error) {return error.number;}
        return 0;
    };
    const QJsonValue omitted(QJsonValue::Null);
    assert(errorFrom([&] {linaccess::parseNativeMacro(definition({}));})==450);
    assert(errorFrom([&] {linaccess::parseNativeMacro(definition({omitted,omitted,omitted,omitted}));})==450);
    assert(errorFrom([&] {linaccess::parseNativeMacro(definition({15,15,15,15,15}));})==450);
    assert(errorFrom([&] {linaccess::parseNativeMacro(definition({15.5}));})==13);
    assert(errorFrom([&] {linaccess::parseNativeMacro(definition({"15"}));})==13);
    assert(errorFrom([&] {linaccess::parseNativeMacro(definition({32768}));})==6);
    assert(errorFrom([&] {linaccess::parseNativeMacro(definition({omitted,omitted,0}));})==5);

    const auto full=definition({210,390,7500,5250});
    const auto height=definition({omitted,omitted,omitted,6000});
    const auto heightActions=height.value("actions").toArray();
    assert(linaccess::nativeMacroDraftActions(linaccess::nativeMacroDraftRows(heightActions))==heightActions);
    linaccess::storeNativeMacro(database,"ResizeProjectForm",full);
    linaccess::storeNativeMacro(database,"ResizeFormHeight",height);
    linaccess::storeNativeMacro(database,"ResizeFormTooSmall",definition({omitted,omitted,30}));
    const auto persisted=linaccess::loadNativeMacro(database,"ResizeFormHeight");
    assert(persisted.size()==1 && persisted.front().arguments.size()==4);
    assert(persisted.front().arguments[0].kind==VbaValue::Kind::Missing && persisted.front().arguments[3].number==6000);

    linaccess::NativeStandardModuleHost host(&window,database,[] {});
    auto commands=host.applicationObject()->member("DoCmd",{}).object;
    const auto run=[&](const auto& commandHost,const std::string& name) {
        commandHost->member("RunMacro",{VbaValue::string(name)});
    };
    window.focusDatabaseWindow();
    const auto projectGeometry=window.geometry();
    assert(errorFrom([&] {run(commands,"ResizeProjectForm");})==2046);
    assert(window.geometry()==projectGeometry);

    QSqlQuery sql(database);
    assert(sql.exec("CREATE TABLE native_move_size_rows(id INTEGER PRIMARY KEY,n INTEGER)"));
    assert(sql.exec("INSERT INTO native_move_size_rows VALUES(1,10)"));
    const QByteArray source=R"FORM(Begin Form
RecordSource ="native_move_size_rows"
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
    AccessFormPreview form("NativeMoveSizeForm",source,{},database,&window);
    form.show();
    QApplication::processEvents();
    QApplication::setActiveWindow(&form);
    auto formObject=makePreviewFormObject(form);
    formObject->objectItem("Value").object->assign("Value",VbaValue::integer(71));
    const auto origin=window.mapToGlobal(QPoint{});
    run(commands,"ResizeProjectForm");
    assert(form.geometry()==QRect(origin+QPoint(14,26),QSize(500,350)));
    auto formCommands=makePreviewApplicationObject(form)->member("DoCmd",{}).object;
    run(formCommands,"ResizeFormHeight");
    const QRect expected(origin+QPoint(14,26),QSize(500,400));
    assert(form.geometry()==expected && window.geometry()==projectGeometry);
    form.setMinimumWidth(90);
    assert(errorFrom([&] {run(formCommands,"ResizeFormTooSmall");})==5 && form.geometry()==expected);
    assert(formObject->member("Dirty",{}).boolean);
    assert(formObject->objectItem("Value").object->scalarValue().number==71);
    assert(sql.exec("SELECT n FROM native_move_size_rows") && sql.next() && sql.value(0).toInt()==10);
    sql.finish();

    form.showMaximized();
    QApplication::processEvents();
    const auto maximized=form.geometry();
    assert(errorFrom([&] {run(formCommands,"ResizeFormHeight");})==2046);
    assert(form.isMaximized() && form.geometry()==maximized);
    form.showNormal();
    formObject->member("Undo",{});
    assert(form.close());
    window.focusDatabaseWindow();
    assert(linaccess::loadNativeMacro(database,"ResizeFormHeight").front().arguments[3].number==6000);

    // An Access popup uses the desktop frame even when the project origin is offset.
    window.move(180,135);
    auto popupSource=source;
    popupSource.replace("Begin Form\n","Begin Form\nPopUp =NotDefault\n");
    AccessFormPreview popup("NativeMoveSizePopup",popupSource,{},database,&window);
    assert(popup.isPopupDocument());
    popup.show();
    QApplication::processEvents();
    QApplication::setActiveWindow(&popup);
    const auto projectOrigin=window.mapToGlobal(QPoint{});
    assert(!projectOrigin.isNull());
    run(commands,"ResizeProjectForm");
    const QRect popupFrame(14,26,500,350);
    assert(popup.frameGeometry()==popupFrame);
    auto popupCommands=makePreviewApplicationObject(popup)->member("DoCmd",{}).object;
    run(popupCommands,"ResizeFormHeight");
    assert(popup.frameGeometry()==QRect(14,26,500,400));
    assert(window.mapToGlobal(QPoint{})==projectOrigin);
    assert(popup.close());
    window.focusDatabaseWindow();
    window.setGeometry(projectGeometry);
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

void testNativeSqlObjectProperties(MainWindow& window,QSqlDatabase db) {
    using linaccess::VbaValue;QSqlQuery sql(db);assert(sql.exec("CREATE TABLE sql_property_rows(id INTEGER PRIMARY KEY,n INTEGER,hidden TEXT)"));assert(sql.exec("INSERT INTO sql_property_rows VALUES(1,10,'initial'),(2,20,NULL),(3,30,NULL)"));
    const QByteArray source=R"FORM(Begin Form
Caption ="Property header"
RecordSource ="sql_property_rows"
Width =3600
Begin Section
Height =1400
Begin TextBox
Name ="Value"
ControlSource ="n"
Width =1500
Height =300
End
Begin TextBox
Name ="dot.name"
Top =350
Width =1500
Height =300
End
Begin Label
Name ="Label Caption"
Caption ="Label's caption"
Top =700
Width =3000
Height =300
End
Begin ComboBox
Name ="Choice"
RowSource ="SELECT id,n FROM sql_property_rows ORDER BY id"
ColumnCount =2
BoundColumn =1
Top =1050
Width =1800
Height =300
End
End
End)FORM";
    AccessFormPreview form("SQL Property Form",source,{},db,&window);form.show();QApplication::processEvents();auto root=makePreviewFormObject(form);root->assign("Tag",VbaValue::string("form's tag"));auto control=root->objectItem("Value").object;control->assign("Tag",VbaValue::string("control tag"));root->objectItem("dot.name").object->assign("Tag",VbaValue::string("dotted tag"));linaccess::AccessQueryModel model;
    model.execute(db,"SELECT Forms![SQL Property Form].Caption AS title, Forms![SQL Property Form].Visible AS shown, Forms![SQL Property Form].Tag AS tag, Forms![SQL Property Form]!Value.Tag AS control_tag, Formulare![SQL Property Form]![dot.name].Tag AS dotted, Forms![SQL Property Form]![Label Caption].Caption AS label, Forms![SQL Property Form]!Choice.ListCount AS choices, Forms![SQL Property Form].OpenArgs AS args");assert(model.rowCount()==1 && model.columnCount()==8 && model.data(model.index(0,0)).toString()=="Property header" && model.data(model.index(0,1)).toInt()==-1 && model.data(model.index(0,2)).toString()=="form's tag" && model.data(model.index(0,3)).toString()=="control tag" && model.data(model.index(0,4)).toString()=="dotted tag" && model.data(model.index(0,5)).toString()=="Label's caption" && model.data(model.index(0,6)).toInt()==3 && model.data(model.index(0,7)).isNull());assert(!root->member("Dirty",{}).boolean);
    control->assign("Value",VbaValue::integer(42));root->objectItem("hidden").object->assign("Value",VbaValue::string("draft hidden"));model.execute(db,"SELECT Forms![SQL Property Form].Dirty AS dirty, Forms![SQL Property Form].CurrentRecord AS current_row, Forms![SQL Property Form]!Value.OldValue AS previous, Forms![SQL Property Form]!Value.Value AS current_value, Forms![SQL Property Form]!hidden AS hidden_value, Forms![SQL Property Form]!hidden.Name AS field_name");assert(model.data(model.index(0,0)).toInt()==-1 && model.data(model.index(0,1)).toInt()==1 && model.data(model.index(0,2)).toInt()==10 && model.data(model.index(0,3)).toInt()==42 && model.data(model.index(0,4)).toString()=="draft hidden" && model.data(model.index(0,5)).toString()=="hidden");assert(sql.exec("SELECT n,hidden FROM sql_property_rows WHERE id=1") && sql.next() && sql.value(0).toInt()==10 && sql.value(1).toString()=="initial");sql.finish();
    const auto reject=[&](const QString& statement,int expected) {int number=0;try {model.execute(db,statement);}catch(const linaccess::VbaRuntimeError& error) {number=error.number;}if(number!=expected) qFatal("SQL property expected=%d actual=%d",expected,number);};reject("SELECT Forms![SQL Property Form].Undo",438);reject("SELECT Forms![SQL Property Form]!Value.Undo",438);reject("SELECT Forms![SQL Property Form]!Value.Parent",438);assert(root->member("Dirty",{}).boolean && control->scalarValue().number==42);
    assert(sql.exec("CREATE VIEW live_sql_property_view AS "+linaccess::compileAccessQuery("SELECT Forms![SQL Property Form].Caption AS title, Forms![SQL Property Form]!Value.Tag AS tag").sql));root->assign("Caption",VbaValue::string("Changed property header"));control->assign("Tag",VbaValue::string("changed tag"));model.execute(db,"SELECT title,tag FROM live_sql_property_view");assert(model.data(model.index(0,0)).toString()=="Changed property header" && model.data(model.index(0,1)).toString()=="changed tag");
    {auto other=QSqlDatabase::addDatabase("QSQLITE","foreign-sql-properties");other.setDatabaseName(":memory:");assert(other.open());QSqlQuery otherSql(other);assert(otherSql.exec("CREATE TABLE sql_property_rows(id INTEGER PRIMARY KEY,n INTEGER,hidden TEXT)"));assert(otherSql.exec("INSERT INTO sql_property_rows VALUES(1,99,NULL)"));auto foreignSource=source;foreignSource.replace("Property header","Foreign property header");AccessFormPreview otherForm("SQL Property Form",foreignSource,{},other);linaccess::AccessQueryModel otherModel;otherModel.execute(other,"SELECT Forms![SQL Property Form].Caption AS title");assert(otherModel.data(otherModel.index(0,0)).toString()=="Foreign property header");model.execute(db,"SELECT Forms![SQL Property Form].Caption AS title");assert(model.data(model.index(0,0)).toString()=="Changed property header");assert(otherForm.close());}QSqlDatabase::removeDatabase("foreign-sql-properties");
    {AccessFormPreview duplicate("SQL Property Form",source,{},db,&window);reject("SELECT Forms![SQL Property Form].Caption",5);reject("SELECT Forms![SQL Property Form]!Value",5);assert(duplicate.close());}model.execute(db,"SELECT Forms![SQL Property Form]!Value AS current_value");assert(model.data(model.index(0,0)).toInt()==42);
    const QByteArray reportSource=R"FORM(Begin Report
Caption ="Report property header"
RecordSource ="sql_property_rows"
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
End)FORM";AccessFormPreview report("SQL Property Report",reportSource,{},db,&window);model.execute(db,"SELECT Berichte![SQL Property Report].Caption AS title, Reports![SQL Property Report]!Value.Name AS control_name");assert(model.data(model.index(0,0)).toString()=="Report property header" && model.data(model.index(0,1)).toString()=="Value");assert(report.close());reject("SELECT Reports![SQL Property Report].Caption",2451);root->member("Undo",{});assert(form.close());reject("SELECT Forms![SQL Property Form].Caption",2450);
}

void testNativeModuleDebugger(MainWindow& window,QSqlDatabase db) {
    QSqlQuery sql(db);assert(sql.exec("CREATE TABLE debugger_audit(n INTEGER)"));
    const QString code=R"VBA(Public Sub DebugEntry()
Dim n As Long
Dim a(0 To 299) As Long
Dim text As String
n=1
a(299)=9
text=String(600,"x")
DebugChild n
CurrentDb.Execute "INSERT INTO debugger_audit VALUES(" & n & ")"
n=1/0
CurrentDb.Execute "INSERT INTO debugger_audit VALUES(99)"
End Sub
Private Sub DebugChild(ByRef n As Long)
n=n+2
End Sub)VBA";
    const auto audit=[&] {QSqlQuery query(db);assert(query.exec("SELECT n FROM debugger_audit ORDER BY rowid"));QList<int> values;while(query.next()) values.append(query.value(0).toInt());return values;};
    QTimer::singleShot(0,&window,[&] {auto* editor=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(editor && editor->objectName()=="native-module-editor");editor->findChild<QLineEdit*>("native-module-name")->setText("NativeDebuggerProbe");auto* source=editor->findChild<QPlainTextEdit*>("native-module-source");source->setPlainText(code);QApplication::setActiveWindow(editor);source->setFocus();QApplication::processEvents();QTest::keyClick(source,Qt::Key_S,Qt::ControlModifier);QApplication::processEvents();assert(!editor->isWindowModified());
        const auto mark=[&] {source->setTextCursor(QTextCursor(source->document()->findBlockByNumber(7)));source->setFocus();QApplication::processEvents();QTest::keyClick(source,Qt::Key_F9);QApplication::processEvents();};mark();assert(editor->findChild<QLabel*>("native-module-breakpoint-status")->text().contains("gesetzt") && !editor->isWindowModified());
        const auto local=[&](QDialog* debugger,const QString& name) {auto* values=debugger->findChild<QTreeWidget*>("native-debugger-locals");assert(values);const auto found=values->findItems(name,Qt::MatchExactly,0);assert(found.size()==1);return found.front();};const auto openValue=[&](QDialog* debugger,QTreeWidgetItem* item) {auto* values=debugger->findChild<QTreeWidget*>("native-debugger-locals");values->doItemsLayout();values->scrollToItem(item);values->setCurrentItem(item);const auto rectangle=values->visualItemRect(item);assert(rectangle.isValid());QTest::mouseClick(values->viewport(),Qt::LeftButton,Qt::NoModifier,rectangle.center());QTest::mouseDClick(values->viewport(),Qt::LeftButton,Qt::NoModifier,rectangle.center());};
        int stage=0;bool inspecting=false;QTimer probe;probe.setInterval(1);QObject::connect(&probe,&QTimer::timeout,&window,[&] {if(inspecting) return;auto* debugger=qobject_cast<QDialog*>(QApplication::activeModalWidget());if(!debugger || debugger->objectName()!="native-module-debugger") return;auto* into=debugger->findChild<QPushButton*>("native-debugger-step-into");auto* skip=debugger->findChild<QPushButton*>("native-debugger-skip-error");if(!into->isEnabled() && !skip->isEnabled()) return;auto* stack=debugger->findChild<QTreeWidget*>("native-debugger-stack");assert(stack);
            if(stage==0) {assert(stack->topLevelItemCount()==1 && stack->topLevelItem(0)->text(1)=="8" && local(debugger,"n")->text(3)=="1" && audit().isEmpty() && !source->isEnabled());assert(!window.close() && db.isOpen());inspecting=true;bool textRead=false;QTimer textProbe;textProbe.setInterval(1);QObject::connect(&textProbe,&QTimer::timeout,&window,[&] {auto* inspector=qobject_cast<QDialog*>(QApplication::activeModalWidget());if(!inspector || inspector->objectName()!="native-debugger-value-inspector") return;assert(inspector->findChild<QPlainTextEdit*>("native-debugger-inspector-text")->toPlainText()==QString(600,'x'));textRead=true;inspector->accept();});textProbe.start();openValue(debugger,local(debugger,"text"));textProbe.stop();assert(textRead);bool arrayRead=false;QTimer arrayProbe;arrayProbe.setInterval(1);QObject::connect(&arrayProbe,&QTimer::timeout,&window,[&] {auto* inspector=qobject_cast<QDialog*>(QApplication::activeModalWidget());if(!inspector || inspector->objectName()!="native-debugger-value-inspector") return;auto* next=inspector->findChild<QPushButton*>("native-debugger-array-next");next->click();next->click();auto* items=inspector->findChild<QTreeWidget*>("native-debugger-array-items");assert(items->topLevelItemCount()==44 && items->topLevelItem(43)->text(0)=="#299" && items->topLevelItem(43)->text(2)=="9");arrayRead=true;inspector->accept();});arrayProbe.start();openValue(debugger,local(debugger,"a"));arrayProbe.stop();assert(arrayRead);inspecting=false;stage=1;into->click();}
            else if(stage==1) {assert(stack->topLevelItemCount()==2 && stack->topLevelItem(0)->text(0).endsWith(".DebugChild") && local(debugger,"n")->text(3)=="1");stage=2;debugger->findChild<QPushButton*>("native-debugger-step-out")->click();}
            else if(stage==2) {assert(stack->topLevelItemCount()==1 && stack->topLevelItem(0)->text(1)=="9" && local(debugger,"n")->text(3)=="3");stage=3;debugger->findChild<QPushButton*>("native-debugger-step-over")->click();}
            else if(stage==3) {assert(stack->topLevelItem(0)->text(1)=="10" && audit()==QList<int>({3}));stage=4;debugger->findChild<QPushButton*>("native-debugger-continue")->click();}
            else if(stage==4) {assert(skip->isEnabled() && debugger->findChild<QLabel*>("native-debugger-status")->text().contains("11") && local(debugger,"n")->text(3)=="3");stage=5;skip->click();}
            else {assert(stage==5 && stack->topLevelItem(0)->text(1)=="11" && local(debugger,"err")->child(0)->text(3)=="11" && audit()==QList<int>({3}));stage=6;debugger->findChild<QPushButton*>("native-debugger-continue")->click();}
        });probe.start();editor->findChild<QPushButton*>("native-module-run")->click();probe.stop();assert(stage==6 && audit()==QList<int>({3,99}) && !editor->isWindowModified() && source->isEnabled());auto* result=editor->findChild<QTextBrowser*>("native-module-result");assert(result->toPlainText().startsWith("Prozedur ausgeführt."));
        assert(sql.exec("DELETE FROM debugger_audit"));bool aborted=false;QTimer abortProbe;abortProbe.setInterval(1);QObject::connect(&abortProbe,&QTimer::timeout,&window,[&] {auto* debugger=qobject_cast<QDialog*>(QApplication::activeModalWidget());if(debugger && debugger->objectName()=="native-module-debugger" && debugger->findChild<QPushButton*>("native-debugger-continue")->isEnabled()) {aborted=true;debugger->findChild<QPushButton*>("native-debugger-stop")->click();}});abortProbe.start();editor->findChild<QPushButton*>("native-module-run")->click();abortProbe.stop();assert(aborted && audit().isEmpty() && result->toPlainText().contains("im Debugger beendet"));mark();if(!editor->findChild<QLabel*>("native-module-breakpoint-status")->text().contains("entfernt")) qFatal("Debugger F9 after stop: status=%s focus=%d active=%d enabled=%d",qPrintable(editor->findChild<QLabel*>("native-module-breakpoint-status")->text()),source->hasFocus(),editor->isActiveWindow(),source->isEnabled());
        int errors=0;QTimer errorProbe;errorProbe.setInterval(1);QObject::connect(&errorProbe,&QTimer::timeout,&window,[&] {auto* debugger=editor->findChild<QDialog*>("native-module-debugger");if(!debugger || !debugger->isVisible()) return;auto* skip=debugger->findChild<QPushButton*>("native-debugger-skip-error");auto* resume=debugger->findChild<QPushButton*>("native-debugger-continue");if(skip->isEnabled()) {++errors;resume->click();}else if(resume->isEnabled()) resume->click();else if(debugger->findChild<QPushButton*>("native-debugger-close")->isEnabled()) {assert(!debugger->findChild<QTreeWidget*>("native-debugger-stack")->isEnabled() && !debugger->findChild<QTreeWidget*>("native-debugger-locals")->isEnabled());debugger->findChild<QPushButton*>("native-debugger-close")->click();}});errorProbe.start();editor->findChild<QPushButton*>("native-module-run")->click();errorProbe.stop();assert(errors==1 && audit()==QList<int>({3}) && result->toPlainText().contains("Fehler 11"));
        auto fixed=code;fixed.replace("n=1/0","n=4");source->setPlainText(fixed);source->setFocus();QApplication::processEvents();QTest::keyClick(source,Qt::Key_S,Qt::ControlModifier);QApplication::processEvents();assert(!editor->isWindowModified());assert(sql.exec("DELETE FROM debugger_audit"));editor->findChild<QPushButton*>("native-module-run")->click();assert(audit()==QList<int>({3,99}) && result->toPlainText().startsWith("Prozedur ausgeführt."));editor->reject();
    });action(window,"+ Modul")->trigger();
}

void testNativeCapabilityReport(MainWindow& window,QSqlDatabase db) {
    using namespace linaccess;
    clearExecutionObservations(db);QSqlQuery query(db);assert(query.exec("CREATE TABLE capability_audit(n INTEGER)"));
    const QByteArray source="Public Sub CapabilityRoot()\nCapabilityChild\nEnd Sub\nPrivate Sub CapabilityChild()\nCurrentDb.Execute \"INSERT INTO capability_audit VALUES(7)\"\nEnd Sub\nPublic Sub CapabilityFailure()\nDim n As Long\nn=1/0\nEnd Sub\nPublic Sub CapabilityNeverRun()\nCurrentDb.Execute \"INSERT INTO capability_audit VALUES(99)\"\nEnd Sub";
    const auto save=[&](QByteArray bytes) {query.prepare("INSERT OR REPLACE INTO _la_objects VALUES('module','CapabilityProbe',?)");query.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba",QString::fromUtf8(bytes)}}).toJson(QJsonDocument::Compact));assert(query.exec());query.finish();};save(source);
    const auto find=[&](const QJsonObject& report) {for(const auto& value:report["artifacts"].toArray()) {const auto item=value.toObject();if(item["origin"]=="native" && item["name"]=="CapabilityProbe") return item;}assert(false);return QJsonObject{};};
    auto item=find(projectCapabilityReport(db));assert(item["procedures"].toArray().size()==4 && executionObservations(db).isEmpty());for(const auto& value:item["procedures"].toArray()) assert(value.toObject()["execution"]=="not_observed");
    assert(query.exec("SELECT COUNT(*) FROM capability_audit") && query.next() && query.value(0).toInt()==0);query.finish();
    NativeStandardModuleHost host(&window,db,[] {});assert(executeVbaProcedure(source.toStdString(),"CapabilityRoot",host).status==VbaExecutionStatus::Completed);assert(executeVbaProcedure(source.toStdString(),"CapabilityFailure",host).errorNumber==11);
    item=find(projectCapabilityReport(db));bool root=false,child=false,failure=false,never=false;for(const auto& value:item["procedures"].toArray()) {const auto entry=value.toObject();const auto name=entry["name"].toString();if(name=="CapabilityRoot") root=entry["last_source_result"]=="completed";if(name=="CapabilityChild") child=entry["last_source_result"]=="completed";if(name=="CapabilityFailure") failure=entry["last_source_result"]=="error" && entry["observations"].toArray().last().toObject()["error"].toInt()==11;if(name=="CapabilityNeverRun") never=entry["execution"]=="not_observed";}assert(root && child && failure && never);
    assert(query.exec("SELECT n FROM capability_audit") && query.next() && query.value(0).toInt()==7 && !query.next());query.finish();
    save(source+"\n'changed revision");item=find(projectCapabilityReport(db));for(const auto& value:item["procedures"].toArray()) assert(value.toObject()["execution"]=="not_observed");
    const auto encoded=QJsonDocument(projectCapabilityReport(db)).toJson();assert(!encoded.contains("INSERT INTO capability_audit") && !encoded.contains("n=1/0"));
    QTimer::singleShot(0,&window,[&] {auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(dialog && dialog->objectName()=="project-capability-report" && dialog->findChild<QTreeWidget*>("project-capability-items")->topLevelItemCount()>0);dialog->reject();});window.findChild<QAction*>("project-capabilities")->trigger();
    assert(query.exec("INSERT INTO _la_objects VALUES('module','CapabilityBroken','{}')"));const auto unavailableTypes=projectCapabilityReport(db);assert(unavailableTypes["enum_type_environment"]=="unavailable" && unavailableTypes["enum_type_environment_error"]==3251);bool invalid=false;for(const auto& value:unavailableTypes["artifacts"].toArray()) {const auto entry=value.toObject();if(entry["name"]=="CapabilityBroken") invalid=entry["capability"]=="invalid_definition";}assert(invalid);assert(query.exec("DELETE FROM _la_objects WHERE kind='module' AND name='CapabilityBroken'"));query.finish();
    query.prepare("INSERT INTO _la_objects VALUES('module','CapabilityBadEnum',?)");query.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba","Public Enum Double\nFirst=1\nEnd Enum\nSub BadEnum()\nMsgBox \"unreached\"\nEnd Sub"}}).toJson(QJsonDocument::Compact));assert(query.exec());query.finish();
    const auto badEnumReport=projectCapabilityReport(db);assert(badEnumReport["enum_type_environment_error"]==5);bool rejectedEnum=false;
    for(const auto& value:badEnumReport["artifacts"].toArray()) {const auto item=value.toObject();if(item["name"]!="CapabilityBadEnum") continue;for(const auto& value:item["procedures"].toArray()) if(value.toObject()["name"]=="BadEnum") rejectedEnum=value.toObject()["validation_error"]==5;}
    assert(rejectedEnum && query.exec("DELETE FROM _la_objects WHERE kind='module' AND name='CapabilityBadEnum'"));query.finish();
    QSqlQuery rollback(db);assert(rollback.exec("SAVEPOINT capability_observations"));assert(executeVbaProcedure(source.toStdString(),"CapabilityRoot",host).status==VbaExecutionStatus::Completed);assert(rollback.exec("ROLLBACK TO capability_observations") && rollback.exec("RELEASE capability_observations"));assert(!executionObservations(db).isEmpty());
    const QString select="SELECT n FROM capability_audit";query.prepare("INSERT INTO _la_objects VALUES('query','CapabilityQuery',?)");query.addBindValue(QJsonDocument(QJsonObject{{"sql",select}}).toJson(QJsonDocument::Compact));assert(query.exec());query.finish();assert(executeAccessQueryPlan(db,compileAccessQuery(select),{}).rows.size()==1);
    const QString missing="SELECT missing_column FROM capability_audit";query.prepare("INSERT INTO _la_objects VALUES('query','CapabilityMissingQuery',?)");query.addBindValue(QJsonDocument(QJsonObject{{"sql",missing}}).toJson(QJsonDocument::Compact));assert(query.exec());query.finish();bool failedQuery=false;try {(void)executeAccessQueryPlan(db,compileAccessQuery(missing),{});}catch(const VbaRuntimeError&) {failedQuery=true;}assert(failedQuery);
    storeNativeMacro(db,"CapabilityMacro",{{"version",1},{"actions",QJsonArray{QJsonObject{{"action","RunSQL"},{"arguments",QJsonArray{"INSERT INTO capability_audit VALUES(8)"}}},QJsonObject{{"action","StopMacro"},{"arguments",QJsonArray{}}}}}});host.applicationObject()->member("DoCmd",{}).object->member("RunMacro",{VbaValue::string("CapabilityMacro")});
    bool observedSelect=false,observedFailure=false,observedMacro=false;for(const auto& value:projectCapabilityReport(db)["artifacts"].toArray()) {const auto entry=value.toObject();if(entry["origin"]!="native") continue;if(entry["name"]=="CapabilityQuery") observedSelect=entry["last_source_result"]=="completed";if(entry["name"]=="CapabilityMissingQuery") observedFailure=entry["last_source_result"]=="error";if(entry["name"]=="CapabilityMacro") observedMacro=entry["last_source_result"]=="stopped" && entry["capability"]=="macro_syntax_accepted_runtime_unproven";}assert(observedSelect && observedFailure && observedMacro);
    clearExecutionObservations(db);assert(executionObservations(db).isEmpty());
}

void testNativeConditionCoercions(MainWindow& window,QSqlDatabase db) {
    using namespace linaccess;QSqlQuery query(db);assert(query.exec("CREATE TABLE condition_audit(n INTEGER)"));
    const auto action=[](QString condition,int value) {return QJsonObject{{"action","RunSQL"},{"arguments",QJsonArray{QString("INSERT INTO condition_audit VALUES(%1)").arg(value)}},{"condition",condition}};};
    storeNativeMacro(db,"ConditionCoercion",{{"version",1},{"actions",QJsonArray{action("\"True\"",1),action("\"False\"",99),action("Null",99),action("\"#TRUE#\"",2),action("\"#FALSE#\"",99)}}});NativeStandardModuleHost host(&window,db,[] {});auto commands=host.applicationObject()->member("DoCmd",{}).object;
    commands->member("RunMacro",{VbaValue::string("ConditionCoercion"),VbaValue::longInteger(2),VbaValue::string("\"True\"")});assert(query.exec("SELECT n FROM condition_audit ORDER BY rowid"));QList<int> values;while(query.next()) values.append(query.value(0).toInt());assert(values==QList<int>({1,2,1,2}));query.finish();
    commands->member("RunMacro",{VbaValue::string("ConditionCoercion"),VbaValue::longInteger(2),VbaValue::string("\"False\"")});assert(query.exec("SELECT COUNT(*) FROM condition_audit") && query.next() && query.value(0).toInt()==4);query.finish();
}

void testNativeProjectGuiExtensions(MainWindow& window,QSqlDatabase db) {
    testNativeRepaintObject(window,db);
    testNativeSetValueDoubleEvaluation(window,db);
    testNativeFormRecordFields(window,db);
    testNativeProjectFilters(window,db);
    testNativeWindowStateCommands(window,db);
    testNativeMoveSize(window,db);
    testNativeMacroContinuations(window,db);
    testNativeSqlObjectProperties(window,db);
    testNativeModuleDebugger(window,db);
    testNativeCapabilityReport(window,db);
    testNativeConditionCoercions(window,db);
}
