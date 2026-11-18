#include "vba_def_type_fixture.hpp"
#include <iostream>
#include "access_form_preview.hpp"
#include "mainwindow.hpp"
#include "native_macro_storage.hpp"
#include "native_standard_module_host.hpp"
#include "vba_static_fixture.hpp"
#include <QApplication>
#include <QLineEdit>
#include <QPlainTextEdit>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

namespace {
void exerciseMacroOpenModule(QSqlDatabase db) {
    using namespace linaccess;
    MainWindow* window=nullptr;
    for(auto* widget:QApplication::topLevelWidgets()) {
        if(auto* candidate=qobject_cast<MainWindow*>(widget);candidate && candidate->isVisible()) {
            assert(!window);
            window=candidate;
        }
    }
    assert(window);

    const auto rawObject=[&](const QString& kind,const QString& name,const QJsonObject& definition) {
        QSqlQuery query(db);
        query.prepare("INSERT INTO _la_objects(kind,name,definition) VALUES(?,?,?)");
        query.addBindValue(kind);
        query.addBindValue(name);
        query.addBindValue(QJsonDocument(definition).toJson(QJsonDocument::Compact));
        assert(query.exec());
    };
    const auto open=[](const QJsonArray& arguments) {
        return QJsonObject{{"action","OpenModule"},{"arguments",arguments}};
    };
    const auto macro=[](const QJsonArray& actions) {
        return QJsonObject{{"version",1},{"actions",actions}};
    };
    rawObject("module","MacroModuleProbe",{{"version",1},{"vba","Option Explicit\n\nPublic Sub ProbeMacroModuleProcedure()\nEnd Sub"}});
    storeNativeMacro(db,"MacroModuleByName",macro({open({"macromoduleprobe"})}));
    storeNativeMacro(db,"MacroModuleByProcedure",macro({open({QJsonValue::Null,"probemacromoduleprocedure"})}));
    storeNativeMacro(db,"MacroModuleMissing",macro({open({"MacroModuleProbe","MissingProcedure"})}));
    rawObject("macro","MacroModuleInvalid",macro({open({"MacroModuleProbe"}),open({QJsonValue::Null,""})}));

    const auto editorFor=[&] {
        QDialog* result=nullptr;
        for(auto* editor:window->findChildren<QDialog*>("native-module-editor")) {
            if(editor->isVisible() && editor->property("la_project_object_name").toString()=="MacroModuleProbe") {
                assert(!result);
                result=editor;
            }
        }
        return result;
    };
    const auto verifyHost=[&](const std::shared_ptr<VbaObject>& commands) {
        const auto run=[&](const char* name) {commands->member("RunMacro",{VbaValue::string(name)});};
        try {
            run("MacroModuleInvalid");
            assert(false);
        } catch(const VbaRuntimeError& failure) {
            assert(failure.number==5);
        }
        assert(!editorFor());
        try {
            run("MacroModuleMissing");
            assert(false);
        } catch(const VbaRuntimeError& failure) {
            assert(failure.number==3251 && std::string(failure.what()).find("Aktion 1 (OpenModule)")!=std::string::npos);
        }
        assert(!editorFor());

        run("MacroModuleByName");
        auto* editor=editorFor();
        assert(editor && !editor->isModal());
        auto* source=editor->findChild<QPlainTextEdit*>("native-module-source");
        assert(source && source->textCursor().blockNumber()==0);
        run("MacroModuleByProcedure");
        assert(editorFor()==editor && source->textCursor().blockNumber()==2);
        editor->reject();
        assert(!editorFor());
    };
    NativeStandardModuleHost host(window,db,[] {});
    verifyHost(host.applicationObject()->member("DoCmd",{}).object);

    AccessFormPreview form("Macro module caller","Begin Form\nEnd\n",{},db,window);
    form.show();
    QApplication::processEvents();
    verifyHost(makePreviewApplicationObject(form)->member("DoCmd",{}).object);
    assert(form.close());

    QSqlQuery cleanup(db);
    assert(cleanup.exec("DELETE FROM _la_objects WHERE kind='module' AND name='MacroModuleProbe'"));
    assert(cleanup.exec("DELETE FROM _la_objects WHERE kind='macro' AND name IN ('MacroModuleByName','MacroModuleByProcedure','MacroModuleMissing','MacroModuleInvalid')"));
}
}

void exerciseNativeStaticProcedures(QSqlDatabase db){using namespace linaccess;
    const QByteArray layout="Begin Form\nBegin Section\nEnd\nEnd\n";const auto source=QByteArray::fromStdString(nativeStaticFixture);
    AccessFormPreview first("Static instance one",layout,source,db),second("Static instance two",layout,source,db);first.show();second.show();QApplication::processEvents();std::vector<VbaValue> args;
    auto result=first.executeProcedure("Retained",args);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="1:2:3:1:A1");result=first.executeProcedure("Retained",args);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="2:4:6:2:A2");result=second.executeProcedure("Retained",args);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="1:2:3:1:A1");
    result=first.executeProcedure("Implicit",args);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="1:1:1:3:1");result=first.executeProcedure("Implicit",args);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="2:2:2:6:2");assert(first.close() && second.close());
    QDialog moduleEditor;
    for(int call=1;call<=2;++call){NativeStandardModuleHost host(&moduleEditor,db,[]{});result=executeVbaProcedure(nativeStaticFixture,"Trailing",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==call);}
    AccessFormPreview reopened("Static new instance",layout,source,db);reopened.show();QApplication::processEvents();result=reopened.executeProcedure("Retained",args);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="1:2:3:1:A1");assert(reopened.close());

    const QByteArray implicitLayout="Begin Form\nBegin Section\nBegin TextBox\nName =\"Actual\"\nEnd\nEnd\nEnd\n";
    const QByteArray implicitCode="Public Function ImplicitProbe As Long\ncount=count+1\nActual=count\nImplicitProbe=Actual+count\nEnd Function";
    AccessFormPreview ordinary("Transient locals",implicitLayout,implicitCode,db);ordinary.show();QApplication::processEvents();auto* actual=ordinary.findChild<QLineEdit*>("access-control-Actual");assert(actual);
    for(int call=0;call<2;++call){result=ordinary.executeProcedure("ImplicitProbe",args);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==2 && actual->text()=="1");}assert(ordinary.close());
    AccessFormPreview typed("Def typed instance",layout,QByteArray::fromStdString(nativeDefTypeFixture),db);typed.show();QApplication::processEvents();
    for(const auto& probe:std::vector<std::pair<const char*,int>>{{"LValue",7},{"ArrayProbe",819533},{"ImplicitArrayProbe",819509},{"MismatchProbe",1300},{"OverrideProbe",805},{"OverflowProbe",602},{"LStatic",1},{"LStatic",2},{"LOrdinary",1},{"LOrdinary",1}}){args.clear();result=typed.executeProcedure(probe.first,args);if(result.status!=VbaExecutionStatus::Completed)std::cerr<<"Form Def "<<probe.first<<": "<<result.explanation<<" line "<<result.line<<" error "<<result.errorNumber<<"\n";assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==probe.second);}
    args.clear();result=typed.executeProcedure("ParameterProbe",args);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="value:0:2");args.clear();result=typed.executeProcedure("PropertyProbe",args);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="cell");assert(typed.close());
    exerciseMacroOpenModule(db);
}
