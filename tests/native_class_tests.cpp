#include "access_form_objects.hpp"
#include "vba_typed_class_fixture.hpp"
#include <QCryptographicHash>
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
    {NativeStandardModuleHost typedHost(&window,db,[]{});
        for(const auto& probe:std::vector<std::pair<const char*,int>>{{"TypedProbe",33},{"TypedMismatch",13},{"TypedRefMismatch",13}}) {
            const auto result=executeVbaProcedure(typedClassConsumer,probe.first,typedHost);
            if(result.status!=VbaExecutionStatus::Completed) std::cerr<<"Project typed class "<<probe.first<<": "<<result.explanation<<" error "<<result.errorNumber<<"\n";
            assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue && result.typedReturnValue->number==probe.second);
        }
    }
    std::shared_ptr<VbaObject> retained;
    {QDialog editor(&window);NativeStandardModuleHost host(&editor,db,[]{});retained=evaluateVbaObjectExpression("New NativeCounter",host,{}).object;assert(retained && retained->member("Value",{}).number==1);}
    assert(retained->member("Value",{}).number==1); // Project lifetime, not the caller editor's lifetime.
    assert(retained->member("SaveAudit",{VbaValue::longInteger(17)}).number==17);assert(sql.exec("SELECT value FROM native_class_audit") && sql.next() && sql.value(0).toInt()==17 && !sql.next());sql.finish();
    installJetSqlFunctions(db);assert(!sql.exec("SELECT SaveAudit(77)"));sql.finish();assert(sql.exec("SELECT COUNT(*) FROM native_class_audit") && sql.next() && sql.value(0).toInt()==1);sql.finish();
    sql.prepare("INSERT INTO main._la_objects(kind,name,definition) VALUES('module','TypedSqlBoundary',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba","Public typedHits As Long\nPublic Function TypedSqlResult As NativeCounter\ntypedHits=typedHits+1\nEnd Function\nPublic Function TypedSqlHits As Long\nTypedSqlHits=typedHits\nEnd Function\n"}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();
    installJetSqlFunctions(db);assert(!sql.exec("SELECT TypedSqlResult()") && sql.lastError().text().contains("3251"));sql.finish();assert(sql.exec("SELECT TypedSqlHits()") && sql.next() && sql.value(0).toInt()==0);sql.finish();assert(sql.exec("DELETE FROM main._la_objects WHERE kind='module' AND name='TypedSqlBoundary'"));
    {AccessFormPreview invalid("Missing class field",QByteArray("Begin Form\nEnd\n"),QByteArray("Public held As NoSuchClass\n"),db,&window);int failure=0;try{(void)makePreviewFormObject(invalid)->member("held",{});}catch(const VbaRuntimeError& e){failure=e.number;}assert(failure==3251);}
    const QByteArray layout="Begin Form\nBegin Section\nBegin TextBox\nName =\"New Copy\"\nEnd\nEnd\nEnd\n";AccessFormPreview form("Native class caller",layout,QByteArray::fromStdString(typedClassConsumer+nativeClassConsumer.substr(nativeClassConsumer.find("Public Function"))+"Public Sub WriteControl\nMe![New Copy]=7\nEnd Sub\nPublic Function RetainedClass As Object\nSet RetainedClass=New NativeCounter\nEnd Function\n"),db,&window);form.show();QApplication::processEvents();std::vector<VbaValue> arguments;
    for(const auto& probe:std::vector<std::pair<const char*,int>>{{"ClassProbe",911},{"ArrayProbe",25},{"ValueGuard",325101},{"RefGuard",325101}}){auto result=form.executeProcedure(probe.first,arguments);if(result.status!=VbaExecutionStatus::Completed)std::cerr<<"Form class "<<probe.first<<": "<<result.explanation<<" line "<<result.line<<" error "<<result.errorNumber<<"\n";assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==probe.second);}
    auto typedResult=form.executeProcedure("TypedProbe",arguments);if(typedResult.status!=VbaExecutionStatus::Completed) std::cerr<<"Typed form error "<<typedResult.errorNumber<<" line "<<typedResult.line<<": "<<typedResult.explanation<<"\n";assert(typedResult.status==VbaExecutionStatus::Completed && typedResult.typedReturnValue->number==33);auto facade=makePreviewFormObject(form);facade->assignObject("held",VbaValue::reference(retained));assert(facade->member("held",{}).object==retained);facade->assignObject("held",VbaValue::reference(nullptr));
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
    const auto archiveHash=[&] {QSqlQuery rows(db);assert(rows.exec("SELECT id,name,source_bytes FROM main._la_access_artifacts ORDER BY id"));QCryptographicHash hash(QCryptographicHash::Sha256);while(rows.next()) {const auto name=rows.value(1).toString().toUtf8(),raw=rows.value(2).toByteArray();hash.addData(QByteArray::number(rows.value(0).toLongLong())+":"+QByteArray::number(name.size())+":"+name+":"+QByteArray::number(raw.size())+":"+raw);}return hash.result();};
    const auto beforeArchive=archiveHash();QVector<qint64> introduced;
    const auto archiveClass=[&](const QString& kind,const QString& name,const QByteArray& raw) {QSqlQuery add(db);add.prepare("INSERT INTO main._la_access_artifacts(source_sha256,source_filename,kind,name,status,format,detail,source_bytes) VALUES(?,'technical.class',?,?,'source_only','technical-class-scope','',?)");add.addBindValue(QString::fromLatin1(QCryptographicHash::hash(raw,QCryptographicHash::Sha256).toHex()));add.addBindValue(kind);add.addBindValue(name);add.addBindValue(raw);assert(add.exec());const auto id=add.lastInsertId().toLongLong();introduced.push_back(id);return id;};
    const auto readBytes=[&](qint64 id){QSqlQuery read(db);read.prepare("SELECT source_bytes FROM main._la_access_artifacts WHERE id=?");read.addBindValue(id);assert(read.exec() && read.next());return read.value(0).toByteArray();};
    auto importedText=QString::fromStdString(nativeClassFixture);importedText.replace("NativeCounter","ImportedCounter");importedText+="Public Function Caption As String\nCaption=\"ä\"\nEnd Function\n";const auto ansi=importedText.toLatin1();const auto ansiId=archiveClass("vba","ImportedCounter",ansi);assert(ansi.contains(char(0xe4)) && nativeProjectClassAvailable(db,"ImportedCounter"));
    auto imported=constructNativeProjectClass(db,&window,"ImportedCounter");assert(imported->member("Caption",{}).text=="ä" && imported->member("Value",{}).number==1);(void)imported->invokeMember("Add",{VbaValue::longInteger(2)},false);assert(imported->member("Value",{}).number==3 && readBytes(ansiId)==ansi);
    assert(catalog->member("Item",{VbaValue::string("ImportedCounter")}).object->member("Name",{}).text=="ImportedCounter");
    assert(imported->member("SaveAudit",{VbaValue::longInteger(19)}).number==19);assert(sql.exec("SELECT COUNT(*) FROM native_class_audit") && sql.next() && sql.value(0).toInt()==2);sql.finish();
    QByteArray utf16("\xff\xfe",2);for(const auto ch:importedText) {const auto unit=ch.unicode();utf16.append(static_cast<char>(unit&255));utf16.append(static_cast<char>(unit>>8));}const auto utf16Id=archiveClass("vba","ImportedCounter",utf16);error=0;try{imported->member("Value",{});}catch(const VbaRuntimeError& e){error=e.number;}assert(error==3420);
    auto importedFresh=constructNativeProjectClass(db,&window,"ImportedCounter");assert(importedFresh->member("Value",{}).number==1 && importedFresh->member("Caption",{}).text=="ä" && readBytes(ansiId)==ansi && readBytes(utf16Id)==utf16);
    const std::string importedConsumer="Function ReadImported As String\nDim value As Object\nSet value=New ImportedCounter\nReadImported=value.Caption & CStr(value.Value)\nEnd Function";QDialog importedOwner(&window);NativeStandardModuleHost importedHost(&importedOwner,db,[]{});result=executeVbaProcedure(importedConsumer,"ReadImported",importedHost);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="ä1");
    auto documentText=QString::fromStdString(nativeClassFixture);documentText.replace("NativeCounter","Form_ImportedDocument");archiveClass("form","ImportedDocument","Begin Form\nEnd\n");archiveClass("vba","Form_ImportedDocument",documentText.toUtf8());assert(!nativeProjectClassAvailable(db,"Form_ImportedDocument"));error=0;try{catalog->member("Item",{VbaValue::string("Form_ImportedDocument")});}catch(const VbaRuntimeError& e){error=e.number;}assert(error==3265);
    auto prefixed=QString::fromStdString(nativeClassFixture);prefixed.replace("NativeCounter","Form_IndependentClass");archiveClass("vba","Form_IndependentClass",prefixed.toUtf8());auto prefixObject=constructNativeProjectClass(db,&window,"Form_IndependentClass");assert(prefixObject->member("Value",{}).number==1 && catalog->member("Item",{VbaValue::string("Form_IndependentClass")}).object);
    auto based=QString::fromStdString(nativeClassFixture);based.replace("NativeCounter","DocumentBasedClass");based.replace("Option Explicit","Attribute VB_Base = \"document-base\"\nOption Explicit");archiveClass("vba","DocumentBasedClass",based.toUtf8());assert(!nativeProjectClassAvailable(db,"DocumentBasedClass"));
    auto mismatch=QString::fromStdString(nativeClassFixture);archiveClass("vba","MismatchedClass",mismatch.toUtf8());assert(!nativeProjectClassAvailable(db,"MismatchedClass"));
    auto loss=QString::fromStdString(nativeClassFixture);loss.replace("NativeCounter","LossyClass");QByteArray lossBytes=loss.toUtf8();lossBytes+="'";lossBytes+=char(0x81);lossBytes+='\n';archiveClass("vba","LossyClass",lossBytes);error=0;try{(void)nativeProjectClassAvailable(db,"LossyClass");}catch(const VbaRuntimeError& e){error=e.number;}assert(error==3251);
    assert(sql.exec("CREATE TEMP TABLE _la_access_artifacts(id INTEGER,kind TEXT,name TEXT,source_bytes BLOB)"));error=0;try{(void)nativeProjectClassAvailable(db,"ImportedCounter");}catch(const VbaRuntimeError& e){error=e.number;}assert(error==3251);assert(sql.exec("DROP TABLE temp._la_access_artifacts"));
    sql.prepare("INSERT INTO main._la_objects(kind,name,definition) VALUES('class','ImportedCounter',?)");sql.addBindValue(QString::fromUtf8(QJsonDocument(QJsonObject{{"version",1},{"vba",importedText}}).toJson(QJsonDocument::Compact)));assert(sql.exec());sql.finish();error=0;try{(void)nativeProjectClassAvailable(db,"ImportedCounter");}catch(const VbaRuntimeError& e){error=e.number;}assert(error==5);assert(sql.exec("DELETE FROM main._la_objects WHERE kind='class' AND name='ImportedCounter'"));
    for(const auto id:introduced) {sql.prepare("DELETE FROM main._la_access_artifacts WHERE id=?");sql.addBindValue(id);assert(sql.exec() && sql.numRowsAffected()==1);sql.finish();}assert(archiveHash()==beforeArchive);error=0;try{importedFresh->member("Value",{});}catch(const VbaRuntimeError& e){error=e.number;}assert(error==3420);installJetSqlFunctions(db);

}
