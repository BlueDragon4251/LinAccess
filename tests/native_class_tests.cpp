#include "access_format.hpp"
#include "vba_set_timing_fixture.hpp"
#include "vba_named_setter_fixture.hpp"
#include "vba_paramarray_alias_fixture.hpp"
#include "vba_array_lock_fixture.hpp"
#include "vba_error_context_fixture.hpp"
#include "vba_error_clear_fixture.hpp"
#include "vba_redim_declaration_fixture.hpp"
#include "vba_native_object_array_fixture.hpp"
#include "vba_enum_type_fixture.hpp"
#include "vba_automatic_class_fixture.hpp"
#include "vba_default_class_fixture.hpp"
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
        for(const auto& probe:std::vector<std::pair<const char*,int>>{{"TypedProbe",33},{"ClassArrayProbe",28}}) {
            const auto result=executeVbaProcedure(typedClassConsumer,probe.first,typedHost);
            if(result.status!=VbaExecutionStatus::Completed) std::cerr<<"Project typed class "<<probe.first<<": "<<result.explanation<<" error "<<result.errorNumber<<"\n";
            assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue && result.typedReturnValue->number==probe.second);
        }
    }
    sql.prepare("INSERT INTO main._la_objects(kind,name,definition) VALUES('class','NativeDefault',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba",QString::fromStdString(nativeDefaultClassFixture)}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();
    {NativeStandardModuleHost defaultHost(&window,db,[]{});const auto probe=executeVbaProcedure(nativeDefaultClassConsumer,"DefaultClassProbe",defaultHost);if(probe.status!=VbaExecutionStatus::Completed) std::cerr<<"Project default class: "<<probe.explanation<<" line "<<probe.line<<" error "<<probe.errorNumber<<"\n";assert(probe.status==VbaExecutionStatus::Completed && probe.typedReturnValue->number==1058);}
    const auto storeAutomatic=[&](const QString& kind,const QString& name,const std::string& source) {sql.prepare("INSERT INTO main._la_objects(kind,name,definition) VALUES(?,?,?)");sql.addBindValue(kind);sql.addBindValue(name);sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba",QString::fromStdString(source)}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();};
    storeAutomatic("class","AutoTracked",automaticClassFixture);storeAutomatic("module","AutomaticState",automaticClassState);
    {NativeStandardModuleHost automaticHost(&window,db,[]{});const auto invoke=[&](const char* name){const auto outcome=executeVbaProcedure(automaticClassConsumer,name,automaticHost);if(outcome.status!=VbaExecutionStatus::Completed) std::cerr<<"Project automatic "<<name<<": "<<outcome.explanation<<" line "<<outcome.line<<" error "<<outcome.errorNumber<<"\n";assert(outcome.status==VbaExecutionStatus::Completed);return outcome.typedReturnValue ? *outcome.typedReturnValue : VbaValue::empty();};assert(invoke("AutoLocalProbe").text=="0:0:1:1:2:3");assert(invoke("AutoStaticProbe").number==2 && invoke("AutoStaticProbe").number==3);assert(invoke("AutoModuleProbe").number==5 && invoke("AutoModuleProbe").number==5);invoke("ResetAutomaticModule");assert(invoke("AutoModuleProbe").number==6 && invoke("AutoDefaultProbe").number==11 && invoke("AutoOwnedProbe").boolean);
        AccessFormPreview automaticForm("Automatic field",QByteArray("Begin Form\nEnd\n"),QByteArray("Public held As New AutoTracked\n"),db,&window);auto facade=makePreviewFormObject(automaticForm);const auto one=facade->member("held",{}).object;assert(one && facade->member("held",{}).object==one);facade->assignObject("held",VbaValue::reference(nullptr));const auto two=facade->member("held",{}).object;assert(two && two!=one && two->member("Serial",{}).number==one->member("Serial",{}).number+1);
    }
    assert(sql.exec("DELETE FROM main._la_objects WHERE (kind='class' AND name='AutoTracked') OR (kind='module' AND name='AutomaticState')"));
    assert(sql.exec("DELETE FROM main._la_objects WHERE kind='class' AND name='NativeDefault'"));
    sql.prepare("INSERT INTO main._la_objects(kind,name,definition) VALUES('module','EnumTypeScope',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba",QString::fromStdString(enumTypeFixture)}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();
    {assert(sql.exec("DELETE FROM main._la_objects WHERE kind='module' AND name='EnumTypeScope'"));sql.prepare("INSERT INTO main._la_objects(kind,name,definition) VALUES('module','EnumTypeScope',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba",QString::fromStdString(qualifiedEnumSqlFixture)}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();installJetSqlFunctions(db);assert(nativeVbaSqlFunctionResultTypes(db).value("enumsqlidentity")=="long");assert(sql.exec("SELECT EnumSqlIdentity(1234)") && sql.next() && sql.value(0).toInt()==1234);sql.finish();}
    assert(sql.exec("DELETE FROM main._la_objects WHERE kind='module' AND name='EnumTypeScope'"));
    {const QByteArray layout="Begin Form\nBegin Section\nBegin TextBox\nName =\"ArrayFirst\"\nEnd\nBegin TextBox\nName =\"ArraySecond\"\nEnd\nEnd\nEnd\n";AccessFormPreview arraysForm("Native array controls",layout,QByteArray::fromStdString(nativeAccessArrayFixture),db,&window);arraysForm.show();QApplication::processEvents();std::vector<VbaValue> inputs;const auto probe=arraysForm.executeProcedure("NativeAccessArrayProbe",inputs);if(probe.status!=VbaExecutionStatus::Completed) std::cerr<<"Access native array: "<<probe.explanation<<" line "<<probe.line<<" error "<<probe.errorNumber<<"\n";assert(probe.status==VbaExecutionStatus::Completed && probe.typedReturnValue->number==39);assert(arraysForm.close());}
    {NativeStandardModuleHost declarationHost(&window,db,[]{});for(const auto& probe:std::vector<std::pair<const char*,int>>{{"ReDimDeclarationProbe",30},{"ReDimExistingProbe",17},{"ReDimStaticProbe",7},{"ReDimStaticProbe",14}}) {const auto outcome=executeVbaProcedure(nativeReDimDeclarationFixture,probe.first,declarationHost);if(outcome.status!=VbaExecutionStatus::Completed) std::cerr<<"Project ReDim declaration "<<probe.first<<": "<<outcome.explanation<<" line "<<outcome.line<<" error "<<outcome.errorNumber<<"\n";assert(outcome.status==VbaExecutionStatus::Completed && outcome.typedReturnValue->number==probe.second);}}
    {NativeStandardModuleHost errorHost(&window,db,[]{});for(const auto& probe:std::vector<std::pair<const char*,int>>{{"ErrorClearProbe",704},{"ErrorExitProbe",17}}) {const auto outcome=executeVbaProcedure(nativeErrorClearFixture,probe.first,errorHost);if(outcome.status!=VbaExecutionStatus::Completed) std::cerr<<"Project Err reset "<<probe.first<<": "<<outcome.explanation<<" line "<<outcome.line<<" error "<<outcome.errorNumber<<"\n";assert(outcome.status==VbaExecutionStatus::Completed && outcome.typedReturnValue->number==probe.second);}}
    {NativeStandardModuleHost errorHost(&window,db,[]{});const auto probe=executeVbaProcedure(nativeErrorContextFixture,"EnvironmentErrorProbe",errorHost);if(probe.status!=VbaExecutionStatus::Completed) std::cerr<<"Project shared error context: "<<probe.explanation<<" line "<<probe.line<<" error "<<probe.errorNumber<<"\n";assert(probe.status==VbaExecutionStatus::Completed && probe.typedReturnValue->number==17);const auto converted=executeVbaProcedure(nativeErrorContextFixture,"ErrorVariantConversionProbe",errorHost);assert(converted.status==VbaExecutionStatus::Completed && converted.typedReturnValue->number==196);const auto textProbe=executeVbaProcedure(nativeErrorContextFixture,"ErrorTextProbe",errorHost);if(textProbe.status!=VbaExecutionStatus::Completed) std::cerr<<"Project error text: "<<textProbe.explanation<<" line "<<textProbe.line<<" error "<<textProbe.errorNumber<<"\n";assert(textProbe.status==VbaExecutionStatus::Completed && textProbe.typedReturnValue->number==204);for(const auto& shadow:std::vector<std::pair<const char*,int>>{{"ErrorProcedureShadow",17},{"ErrorArrayShadow",19}}) {const auto outcome=executeVbaProcedure(nativeErrorShadowFixture,shadow.first,errorHost);assert(outcome.status==VbaExecutionStatus::Completed && outcome.typedReturnValue->number==shadow.second);}sql.prepare("INSERT INTO main._la_objects(kind,name,definition) VALUES('module','NativeErrorScalarScope',?)");sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba","Public Function NativeErrorNumber As Long\nNativeErrorNumber=Err.Number\nEnd Function\nPublic Function NativeCVErrText As String\nNativeCVErrText=CStr(CVErr(CVErr(2015)))\nEnd Function\nPublic Function NativeErrorText As String\nNativeErrorText=Error()\nEnd Function\n"}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();installJetSqlFunctions(db);errorHost.errorObject()->assign("Number",VbaValue::longInteger(907));assert(sql.exec("SELECT NativeErrorNumber(),NativeCVErrText(),NativeErrorText()") && sql.next() && sql.value(0).toInt()==907 && sql.value(1).toString()=="Error 2015" && sql.value(2).toString()=="Object variable or With block variable not set");sql.finish();assert(errorHost.errorObject()->member("Number",{}).number==907);errorHost.errorObject()->member("Clear",{});assert(sql.exec("DELETE FROM main._la_objects WHERE kind='module' AND name='NativeErrorScalarScope'"));NativeStandardModuleHost anotherHost(&window,db,[]{});const auto identity=executeVbaProcedure(nativeErrorContextFixture,"EnvironmentIdentityProbe",anotherHost);assert(identity.status==VbaExecutionStatus::Completed && identity.typedReturnValue->boolean);AccessFormPreview form("Native error environment",QByteArray("Begin Form\nEnd\n"),QByteArray::fromStdString(nativeFormErrorContextFixture),db,&window);std::vector<VbaValue> inputs{VbaValue::reference(errorHost.errorObject())};const auto formIdentity=form.executeProcedure("FormEnvironmentError",inputs);assert(formIdentity.status==VbaExecutionStatus::Completed && formIdentity.typedReturnValue->boolean);assert(form.close());}
    {NativeStandardModuleHost calendarHost(&window,db,[]{});
        const auto names=executeVbaProcedure(R"VBA(Function CalendarNames As Boolean
CalendarNames=MonthName(Month:=3,Abbreviate:=True)=VBA.MonthName(3,True) And Len(WeekdayName(Weekday:=1,FirstDayOfWeek:=vbMonday))>0
End Function)VBA","CalendarNames",calendarHost);
        assert(names.status==VbaExecutionStatus::Completed && names.typedReturnValue->boolean);
        sql.prepare("INSERT INTO main._la_objects(kind,name,definition) VALUES('module','NativeCalendarScope',?)");
        sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba","Public Function NativeCalendarMonth(ByVal value As Long) As String\nNativeCalendarMonth=MonthName(value)\nEnd Function\nPublic Function NativeCalendarDay(ByVal value As Long) As String\nNativeCalendarDay=WeekdayName(value,,vbSunday)\nEnd Function\n"}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();installJetSqlFunctions(db);
        assert(sql.exec("SELECT NativeCalendarMonth(3),NativeCalendarDay(1)") && sql.next());
        assert(sql.value(0).toString()==QString::fromStdString(accessCalendarName(true,3,false,0).asText()) && sql.value(1).toString()==QString::fromStdString(accessCalendarName(false,1,false,1).asText()));sql.finish();
        assert(sql.exec("DELETE FROM main._la_objects WHERE kind='module' AND name='NativeCalendarScope'"));

        AccessFormPreview calendarForm("Calendar names",QByteArray("Begin Form\nEnd\n"),QByteArray("Function CalendarNames As Boolean\nCalendarNames=Len(MonthName(3))>0 And Len(WeekdayName(1,,vbSunday))>0\nEnd Function"),db,&window);
        std::vector<VbaValue> inputs;const auto formNames=calendarForm.executeProcedure("CalendarNames",inputs);
        assert(formNames.status==VbaExecutionStatus::Completed && formNames.typedReturnValue->boolean);assert(calendarForm.close());
    }
    {NativeStandardModuleHost lockHost(&window,db,[]{});const auto probe=executeVbaProcedure(nativeArrayLockFixture,"ArrayLockProbe",lockHost);if(probe.status!=VbaExecutionStatus::Completed) std::cerr<<"Project array lock: "<<probe.explanation<<" line "<<probe.line<<" error "<<probe.errorNumber<<"\n";assert(probe.status==VbaExecutionStatus::Completed && probe.typedReturnValue->number==41);}
    {NativeStandardModuleHost packHost(&window,db,[]{});for(const auto& probe:std::vector<std::pair<const char*,int>>{{"ParamArrayAliasProbe",3300},{"ParamArrayElementProbe",19},{"ParamArrayEmptyProbe",9}}) {const auto outcome=executeVbaProcedure(nativeParamArrayAliasFixture,probe.first,packHost);if(outcome.status!=VbaExecutionStatus::Completed) std::cerr<<"Project ParamArray alias "<<probe.first<<": "<<outcome.explanation<<" line "<<outcome.line<<" error "<<outcome.errorNumber<<"\n";assert(outcome.status==VbaExecutionStatus::Completed && outcome.typedReturnValue->number==probe.second);}std::vector<VbaValue> values{VbaValue::longInteger(7)};const auto outcome=executeVbaProcedureWithArguments(nativeParamArrayAliasFixture,"UnboundPack",packHost,values);assert(outcome.status==VbaExecutionStatus::Error && outcome.errorNumber==3251 && values[0].number==7);}
    {NativeStandardModuleHost namedHost(&window,db,[]{});
        // Full named-index and mismatch matrices run in core; retain real
        // project/class/default/form facade integration below.
        sql.prepare("INSERT INTO main._la_objects(kind,name,definition) VALUES('class','NamedWriter',?)");
        sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba",QString::fromStdString(nativeNamedSetterClassFixture)}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();
        const auto classProbe=executeVbaProcedure(nativeNamedSetterClassConsumer,"NamedClassSetterProbe",namedHost);
        if(classProbe.status!=VbaExecutionStatus::Completed) std::cerr<<"Project named class setter: "<<classProbe.explanation<<" line "<<classProbe.line<<" error "<<classProbe.errorNumber<<"\n";
        assert(classProbe.status==VbaExecutionStatus::Completed && classProbe.typedReturnValue->number==461);
        AccessFormPreview namedForm("Named setter source",QByteArray("Begin Form\nEnd\n"),QByteArray::fromStdString(nativeNamedSetterFixture),db,&window);
        std::vector<VbaValue> inputs;const auto formProbe=namedForm.executeProcedure("NamedFormSetterProbe",inputs);
        if(formProbe.status!=VbaExecutionStatus::Completed) std::cerr<<"Form named setter: "<<formProbe.explanation<<" line "<<formProbe.line<<" error "<<formProbe.errorNumber<<"\n";
        assert(formProbe.status==VbaExecutionStatus::Completed && formProbe.typedReturnValue->number==3993);assert(namedForm.close());
        assert(sql.exec("DELETE FROM main._la_objects WHERE kind='class' AND name='NamedWriter'") && sql.numRowsAffected()==1);
    }
    {
        sql.prepare("INSERT INTO main._la_objects(kind,name,definition) VALUES('class','NativeSetBox',?)");
        sql.addBindValue(QJsonDocument(QJsonObject{{"version",1},{"vba",QString::fromStdString(nativeSetTimingClass)}}).toJson(QJsonDocument::Compact));assert(sql.exec());sql.finish();
        NativeStandardModuleHost timingHost(&window,db,[]{});
        const auto defaultProbe=executeVbaProcedure(nativeSetTimingFixture,"DefaultSetTiming",timingHost);
        if(defaultProbe.status!=VbaExecutionStatus::Completed) std::cerr<<"Project Set timing: "<<defaultProbe.explanation<<" line "<<defaultProbe.line<<" error "<<defaultProbe.errorNumber<<"\n";
        assert(defaultProbe.status==VbaExecutionStatus::Completed && defaultProbe.typedReturnValue->number==111);
        AccessFormPreview timingForm("Set evaluation source",QByteArray("Begin Form\nEnd\n"),QByteArray::fromStdString(nativeSetTimingFixture),db,&window);
        std::vector<VbaValue> inputs;const auto formProbe=timingForm.executeProcedure("FormSetTiming",inputs);
        if(formProbe.status!=VbaExecutionStatus::Completed) std::cerr<<"Form Set timing: "<<formProbe.explanation<<" line "<<formProbe.line<<" error "<<formProbe.errorNumber<<"\n";
        assert(formProbe.status==VbaExecutionStatus::Completed && formProbe.typedReturnValue->number==111);assert(timingForm.close());
        assert(executeVbaProcedure(nativeSetTimingFixture,"ClearTimingReferences",timingHost).status==VbaExecutionStatus::Completed);
        assert(sql.exec("DELETE FROM main._la_objects WHERE kind='class' AND name='NativeSetBox'") && sql.numRowsAffected()==1);
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
    {auto source=QString::fromStdString(nativeDefaultClassFixture);source.replace("NativeDefault","ImportedDefault");archiveClass("vba","ImportedDefault",source.toUtf8());auto object=constructNativeProjectClass(db,&window,"ImportedDefault");assert(object->defaultMember({VbaValue::longInteger(3)},false).number==7 && object->member("Item",{VbaValue::longInteger(3)}).number==1002);}
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
