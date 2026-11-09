#include <QMessageBox>
#include <QTest>
#include "native_class_catalog.hpp"
#include "native_class_definition.hpp"
#include "native_standard_module_host.hpp"
#include "native_project_context.hpp"
#include "vba_class_fixture.hpp"
#include "vba_module_context.hpp"
#include "jet_sql_functions.hpp"
#include "access_form_preview.hpp"
#include "mainwindow.hpp"
#include <QApplication>
#include <QTimer>
#include <QDialog>
#include <QLineEdit>
#include <QPlainTextEdit>
#include <QPushButton>
#include <QAction>
#include <QTreeWidget>
#include <QJsonDocument>
#include <QSqlQuery>
#include <QSqlError>
#include <iostream>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
void exerciseNativeClasses(MainWindow& window,QSqlDatabase db) {using namespace linaccess;
    QSqlQuery sql(db);assert(sql.exec("CREATE TABLE native_class_audit(value INTEGER)"));
    auto project=nativeProjectContext(db,&window,false);auto catalog=project->member("AllModules",{}).object;assert(catalog);
    bool saved=false;QTimer::singleShot(0,&window,[&] {
        auto* editor=qobject_cast<QDialog*>(QApplication::activeModalWidget());assert(editor && editor->objectName()=="native-class-editor");auto* name=editor->findChild<QLineEdit*>("native-class-name");auto* source=editor->findChild<QPlainTextEdit*>("native-class-source");auto* save=editor->findChild<QPushButton*>("native-class-save");assert(name && source && save);name->setText("NativeCounter");const auto body=nativeClassFixture.substr(nativeClassFixture.find("Option Explicit"));source->setPlainText(QString::fromStdString(body));QTimer errors;errors.setInterval(1);QObject::connect(&errors,&QTimer::timeout,editor,[]{for(auto* widget:QApplication::allWidgets()) if(auto* box=qobject_cast<QMessageBox*>(widget);box && box->isVisible() && box->windowTitle()=="LinAccess — Fehler") {std::cerr<<"Technical class editor error: "<<box->text().toStdString()<<"\n";box->accept();}});errors.start();QApplication::setActiveWindow(editor);source->setFocus();QApplication::processEvents();QTest::keyClick(source,Qt::Key_S,Qt::ControlModifier);errors.stop();assert(!editor->isWindowModified() && name->isReadOnly() && source->toPlainText().startsWith("VERSION 1.0 CLASS\nAttribute VB_Name = \"NativeCounter\"\n"));auto descriptor=catalog->member("Item",{VbaValue::string("NativeCounter")}).object;assert(descriptor && descriptor->member("IsLoaded",{}).boolean);saved=true;editor->reject();
    });
    QAction* create=nullptr;for(auto* item:window.findChildren<QAction*>()) if(item->text()=="+ Klasse") create=item;assert(create);create->trigger();assert(saved);
    auto descriptor=catalog->member("Item",{VbaValue::string("NativeCounter")}).object;assert(!descriptor->member("IsLoaded",{}).boolean);
    bool found=false;for(auto* tree:window.findChildren<QTreeWidget*>()) for(auto* item:tree->findItems("NativeCounter",Qt::MatchExactly|Qt::MatchRecursive)) if(item->data(0,Qt::UserRole).toString()=="class") found=true;assert(found);
    assert(nativeProjectClassAvailable(db,"nativecounter"));
    std::shared_ptr<VbaObject> retained;
    {QDialog editor(&window);NativeStandardModuleHost host(&editor,db,[]{});retained=evaluateVbaObjectExpression("New NativeCounter",host,{}).object;assert(retained && retained->member("Value",{}).number==1);}
    assert(retained->member("Value",{}).number==1); // Project lifetime, not the caller editor's lifetime.
    assert(retained->member("SaveAudit",{VbaValue::longInteger(17)}).number==17);assert(sql.exec("SELECT value FROM native_class_audit") && sql.next() && sql.value(0).toInt()==17 && !sql.next());sql.finish();
    installJetSqlFunctions(db);assert(!sql.exec("SELECT SaveAudit(77)"));sql.finish();assert(sql.exec("SELECT COUNT(*) FROM native_class_audit") && sql.next() && sql.value(0).toInt()==1);sql.finish();
    const QByteArray layout="Begin Form\nBegin Section\nBegin TextBox\nName =\"New Copy\"\nEnd\nEnd\nEnd\n";AccessFormPreview form("Native class caller",layout,QByteArray::fromStdString(nativeClassConsumer+"Public Sub WriteControl\nMe![New Copy]=7\nEnd Sub\nPublic Function RetainedClass As Object\nSet RetainedClass=New NativeCounter\nEnd Function\n"),db,&window);form.show();QApplication::processEvents();std::vector<VbaValue> arguments;
    for(const auto& probe:std::vector<std::pair<const char*,int>>{{"ClassProbe",911},{"ArrayProbe",25},{"ValueGuard",325101},{"RefGuard",325101}}){auto result=form.executeProcedure(probe.first,arguments);if(result.status!=VbaExecutionStatus::Completed)std::cerr<<"Form class "<<probe.first<<": "<<result.explanation<<" line "<<result.line<<" error "<<result.errorNumber<<"\n";assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==probe.second);}
    auto control=form.executeProcedure("WriteControl",arguments);assert(control.status==VbaExecutionStatus::Completed && form.findChild<QLineEdit*>("access-control-New Copy")->text()=="7");
    auto result=form.executeProcedure("MethodProbe",arguments);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="x18:20");result=form.executeProcedure("IdentityProbe",arguments);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->boolean);
    result=form.executeProcedure("RetainedClass",arguments);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->object);auto afterForm=result.typedReturnValue->object;assert(form.close());assert(afterForm->member("Calls",{}).number==1);
    assert(sql.exec("SELECT definition FROM main._la_objects WHERE kind='class' AND name='NativeCounter'") && sql.next());const auto original=sql.value(0).toByteArray();sql.finish();auto altered=QJsonDocument::fromJson(original).object();altered.insert("vba",altered.value("vba").toString()+"\n");sql.prepare("UPDATE main._la_objects SET definition=? WHERE kind='class' AND name='NativeCounter'");sql.addBindValue(QString::fromUtf8(QJsonDocument(altered).toJson(QJsonDocument::Compact)));assert(sql.exec());sql.finish();
    int error=0;try{retained->member("Value",{});}catch(const VbaRuntimeError& e){error=e.number;}assert(error==3420);
    sql.prepare("UPDATE main._la_objects SET definition=? WHERE kind='class' AND name='NativeCounter'");sql.addBindValue(QString::fromUtf8(original));assert(sql.exec());sql.finish();error=0;try{retained->member("Value",{});}catch(const VbaRuntimeError& e){error=e.number;}assert(error==3420);
    auto fresh=constructNativeProjectClass(db,&window,"NativeCounter");assert(fresh->member("Value",{}).number==1);
    assert(sql.exec("CREATE TEMP TABLE _la_objects(kind TEXT,name TEXT,definition TEXT)"));error=0;try{(void)nativeProjectClassAvailable(db,"NativeCounter");}catch(const VbaRuntimeError& e){error=e.number;}assert(error==3251);assert(sql.exec("DROP TABLE temp._la_objects"));
    assert(sql.exec("DELETE FROM main._la_objects WHERE kind='class' AND name='NativeCounter'"));error=0;try{fresh->member("Value",{});}catch(const VbaRuntimeError& e){error=e.number;}assert(error==3420);
    retained.reset();afterForm.reset();fresh.reset();installJetSqlFunctions(db);
}
