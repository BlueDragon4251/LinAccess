#include "native_project_gui_extensions.hpp"
#include "mainwindow.hpp"
#include "native_standard_module_host.hpp"
#include "native_macro_storage.hpp"
#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_table_view.hpp"
#include "dao_database.hpp"
#include <QAction>
#include <QApplication>
#include <QDialog>
#include <QEvent>
#include <QLineEdit>
#include <QPlainTextEdit>
#include <QPrintPreviewWidget>
#include <QPushButton>
#include <QSqlQuery>
#include <QTabWidget>
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

