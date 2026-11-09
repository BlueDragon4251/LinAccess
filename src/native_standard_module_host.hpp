#pragma once
#include "vba_runtime.hpp"
#include "vba_expression.hpp"
#include "vba_module_context.hpp"
#include "vba_source_codec.hpp"
#include "vba_analyzer.hpp"
#include "access_format.hpp"
#include "dao_database.hpp"
#include "native_input_box.hpp"
#include "native_public_procedure_run.hpp"
#include "native_domain_lookup.hpp"
#include "native_macro_runner.hpp"
#include "native_stored_form.hpp"
#include "native_stored_report.hpp"
#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QApplication>
#include <QSqlQuery>
#include <QSqlError>
#include <QMainWindow>
#include <QVariant>
#include <QThread>
#include <QPointer>
#include <QWidget>
#include <QSqlDatabase>
#include <QRegularExpression>
#include <functional>
int nativeModuleMessageBox(QWidget*,std::string_view,int,const std::optional<std::string>&);
namespace linaccess {
// Real database/editor context. There is deliberately no fabricated Me/Form.
class NativeStandardModuleHost final : public VbaHost {
public:
    NativeStandardModuleHost(QWidget* owner,QSqlDatabase db,std::function<void()> guard)
        : owner_(owner),db_(std::move(db)),guard_(std::move(guard)) {}
    void check() const {if(!owner_ || !db_.isOpen()) throw VbaRuntimeError(3420,"Module execution context is closed");if(owner_->thread()!=QThread::currentThread()) throw VbaRuntimeError(3251,"Module application context requires its GUI thread");guard_();}
    void enterModule(std::string_view) override {check();}
    std::optional<std::string> field(std::string_view) override {throw VbaRuntimeError(438,"Me requires a real form instance");}
    void setField(std::string_view,std::optional<std::string>) override {throw VbaRuntimeError(438,"Me requires a real form instance");}
    void saveRecord() override {throw VbaRuntimeError(3251,"Module editor has no current record");}
    int compareStrings(std::string_view a,std::string_view b,VbaCompareMode mode) override {
        if(mode==VbaCompareMode::Binary) return VbaHost::compareStrings(a,b,mode);
        return QString::compare(QString::fromUtf8(a.data(),a.size()),QString::fromUtf8(b.data(),b.size()),Qt::CaseInsensitive);
    }
    bool isNumericText(std::string_view text) override {return isAccessNumericText(text);}
    VbaValue formatValue(const VbaValue& value,std::string_view mask,bool stringResult) override {return formatAccessValue(value,mask,stringResult);}
    VbaValue formatNumericValue(const VbaValue& value,bool currency,const VbaNumericFormatOptions& options) override {return formatAccessNumericValue(value,currency,options);}
    void message(std::string_view value) override {(void)messageBox(value,0,std::nullopt);}
    int messageBox(std::string_view prompt,int buttons,const std::optional<std::string>& title) override {check();const auto result=nativeModuleMessageBox(owner_,prompt,buttons,title);check();return result;}
    std::string inputBox(const VbaInputBoxArguments& arguments) override {check();auto result=nativeInputBox(owner_,arguments);check();return result;}
    std::shared_ptr<VbaObject> currentDatabase() override {check();return makeDaoDatabase(db_);}
    std::shared_ptr<VbaObject> databaseEngine() override {check();return makeDaoEngine(db_);}
    std::shared_ptr<VbaObject> applicationObject() override;
    VbaValue domainAggregate(std::string_view function,const std::vector<VbaValue>& args) override {check();return nativeDomainValue(db_,args,function);}
    void executeSql(std::string_view sql) override {currentDatabase()->member("Execute",{VbaValue::string(std::string(sql))});}
    std::shared_ptr<VbaModuleState> moduleState(std::string_view source) override {check();return isVbaClassModule(source) ? VbaHost::moduleState(source) : standardVbaModuleState(db_,source);}
    std::vector<std::shared_ptr<VbaModuleState>> globalModuleStates() override {check();return publicVbaModuleStates(db_);}
    std::optional<std::string> moduleForProcedure(std::string_view procedure) override {
        check();std::optional<std::string> found;const auto target=QString::fromUtf8(procedure.data(),procedure.size());
        for(const auto& bytes:standardVbaModuleSources(db_)) {
            const auto source=decodeVbaSource(bytes);const auto code=source.toUtf8().toStdString();if(isVbaClassModule(code)) continue;
            const auto analysis=analyzeVba(code);const auto lines=source.split('\n');
            for(const auto& item:analysis.procedures) {
                if(QString::fromStdString(item.name).compare(target,Qt::CaseInsensitive)!=0) continue;
                if(item.line>0 && item.line<=lines.size() && QRegularExpression("^Private\\b",QRegularExpression::CaseInsensitiveOption).match(lines[item.line-1].trimmed()).hasMatch()) continue;
                if(found) throw VbaRuntimeError(5,"Ambiguous public procedure");found=code;
            }
        }
        return found;
    }
private:
    QPointer<QWidget> owner_;QSqlDatabase db_;std::function<void()> guard_;
};
class NativeModuleDoCmd final : public VbaObject {
public:
    NativeModuleDoCmd(QWidget* owner,QSqlDatabase db,std::function<void()> guard):owner_(owner),db_(std::move(db)),guard_(std::move(guard)) {}
    std::string_view className() const override {return "DoCmd";}
    bool isNativeType(std::string_view type) const override {return type=="access.docmd";}
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        const auto key=QString::fromUtf8(name.data(),name.size()).toCaseFolded();
        if(key=="deleteobject") return VbaArgumentSignature{{"ObjectType","ObjectName"},0};
        if(key=="runmacro") return VbaArgumentSignature{{"MacroName","RepeatCount","RepeatExpression"},1};
        if(key=="runsql") return VbaArgumentSignature{{"SQLStatement","UseTransaction"},1};
        if(key=="opentable") return VbaArgumentSignature{{"TableName","View","DataMode"},1};
        if(key=="openquery") return VbaArgumentSignature{{"QueryName","View","DataMode"},1};
        if(key=="openform") return VbaArgumentSignature{{"FormName","View"},1};
        if(key=="openreport") return VbaArgumentSignature{{"ReportName","View"},1};
        if(key=="runcommand") return VbaArgumentSignature{{"Command"},1};
        if(key=="requery") return VbaArgumentSignature{{"ControlName"},0};return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        NativeStandardModuleHost host(owner_,db_,guard_);host.check();const auto key=QString::fromUtf8(name.data(),name.size()).toCaseFolded();
        if(key=="runmacro") {
            runNativeMacro(db_,args,[&] {host.check();},[&](const QString& expression) {return host.eval(expression.toUtf8().toStdString());},[&](const NativeMacroAction& action) {
                if(action.name=="RunCode") {
                    const auto expression=action.arguments.front().asText();const auto target=nativeMacroFunctionName(QString::fromUtf8(expression));
                    QMap<QString,QString> names;(void)standardVbaModuleSources(db_,&names);if(names.contains(target.toCaseFolded())) throw VbaRuntimeError(3251,"RunCode Function/module name conflict");
                    const auto source=host.moduleForProcedure(target.toStdString());if(!source) throw VbaRuntimeError(2517,"RunCode public Function was not found");
                    const auto signature=nativeVbaProcedureSignature(*source,target.toStdString());if(!signature || !signature->function) throw VbaRuntimeError(3251,"RunCode requires a Function");
                    const auto result=executeVbaEventExpression({},"="+expression,host);if(result.status!=VbaExecutionStatus::Completed) throw VbaRuntimeError(result.errorNumber ? result.errorNumber : 3251,result.explanation,result.errorSource,result.errorHelpFile,result.errorHelpContext);
                }else (void)member(action.name.toStdString(),action.arguments);
            });return VbaValue::empty();
        }
        std::vector<VbaValue> values;for(const auto& argument:args) {values.push_back(scalarVbaValue(argument));host.check();}
        const auto number=[&](std::size_t index,int fallback) {return index>=values.size() || values[index].kind==VbaValue::Kind::Missing ? fallback : static_cast<int>(coerceVbaModuleValue(values[index],"long").number);};
        const auto target=[&] {if(values.empty() || values[0].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(450,"Object name is required");return QString::fromStdString(coerceVbaModuleValue(values[0],"string").asText());};
        if(key=="deleteobject") {
            if(values.size()!=2 || values[0].kind==VbaValue::Kind::Missing || values[1].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(3251,"DeleteObject requires explicit type and name");
            const int type=number(0,0);const auto object=QString::fromStdString(coerceVbaModuleValue(values[1],"string").asText());QVariantMap result;
            if(!QMetaObject::invokeMethod(owner_,"deleteDatabaseObject",Qt::DirectConnection,Q_RETURN_ARG(QVariantMap,result),Q_ARG(int,type),Q_ARG(QString,object))) throw VbaRuntimeError(3251,"Native object deletion is unavailable");resultError(result);host.check();return VbaValue::empty();
        }
        if(key=="runsql") {
            if(values.empty() || values.size()>2) throw VbaRuntimeError(450,"RunSQL requires SQL and optional UseTransaction");
            const auto sql=target().toStdString();const bool transaction=values.size()<2 || values[1].kind==VbaValue::Kind::Missing || coerceVbaModuleValue(values[1],"boolean").boolean;
            const auto execute=[&] {try {host.executeSql(sql);}catch(const VbaRuntimeError& error) {if(error.number==3065) throw VbaRuntimeError(2342,"RunSQL requires action SQL");throw;}};
            if(!transaction) {execute();return VbaValue::empty();}
            static thread_local quint64 sequence=0;const auto savepoint=QString("la_module_runsql_%1").arg(++sequence);QSqlQuery query(db_);
            const auto fail=[&] {const int code=query.lastError().nativeErrorCode().toInt()&255;throw VbaRuntimeError(code==5 || code==6 ? 3260 : 3146,"RunSQL savepoint operation failed");};
            if(!query.exec("SAVEPOINT "+savepoint)) fail();
            try {execute();if(!query.exec("RELEASE "+savepoint)) fail();}
            catch(...) {const bool rollback=query.exec("ROLLBACK TO "+savepoint),release=query.exec("RELEASE "+savepoint);if(!rollback || !release) throw VbaRuntimeError(3146,"RunSQL savepoint cleanup failed");throw;}
            return VbaValue::empty();
        }
        if(key=="opentable" || key=="openquery") {
            if(values.empty() || values.size()>3) throw VbaRuntimeError(450,"OpenTable/OpenQuery argument count invalid");
            const auto object=target();const int view=number(1,0),mode=number(2,1);
            if(key=="opentable") {int error=0;if(!QMetaObject::invokeMethod(owner_,"openTableDocument",Qt::DirectConnection,Q_RETURN_ARG(int,error),Q_ARG(QString,object),Q_ARG(int,view),Q_ARG(int,mode))) throw VbaRuntimeError(3251,"Table opening is unavailable");if(error) throw VbaRuntimeError(error,"Table opening failed");}
            else {QVariantMap result;if(!QMetaObject::invokeMethod(owner_,"openQueryDocument",Qt::DirectConnection,Q_RETURN_ARG(QVariantMap,result),Q_ARG(QString,object),Q_ARG(int,view),Q_ARG(int,mode))) throw VbaRuntimeError(3251,"Query opening is unavailable");resultError(result);}
            host.check();return VbaValue::empty();
        }
        if(key=="openform" || key=="openreport") {
            if(values.empty() || values.size()>2) throw VbaRuntimeError(450,"Stored document opening supports name and view");
            if(QApplication::activeModalWidget()) throw VbaRuntimeError(2046,"Document opening is blocked by a modal editor");
            const bool report=key=="openreport";const int view=number(1,0);if(view!=(report ? 2 : 0)) throw VbaRuntimeError(3251,"Only normal native forms and report preview are implemented");
            const auto requested=target();const auto kind=report ? AccessDocumentKind::Report : AccessDocumentKind::Form;
            const auto form=report ? std::optional<NativeStoredForm>{} : loadNativeStoredForm(db_,requested);
            const auto storedReport=report ? loadNativeStoredReport(db_,requested) : std::optional<NativeStoredReport>{};
            if(!form && !storedReport) throw VbaRuntimeError(report ? 2103 : 2102,"Native stored document was not found");
            const auto canonical=report ? storedReport->name : form->name;const auto layout=report ? storedReport->layout : form->layout;const auto code=report ? storedReport->code : form->code;
            bool existed=false;for(const auto& document:AccessFormPreview::projectDocuments(db_)) if(document && document->documentKind()==kind && document->documentName().compare(canonical,Qt::CaseInsensitive)==0) existed=true;
            QPointer<AccessFormPreview> document=AccessFormPreview::showDocument(kind,canonical,layout,code,db_,owner_,report ? std::optional<QString>(storedReport->recordSource) : std::nullopt);
            if(!document) throw VbaRuntimeError(2501,"Document opening was cancelled");
            try {
                for(const auto* event:{"OnOpen","OnLoad","OnActivate","OnCurrent","OnNoData"}) if(const auto outcome=document->eventResult(event);outcome && outcome->status!=VbaExecutionStatus::Completed && outcome->status!=VbaExecutionStatus::ProcedureMissing) throw VbaRuntimeError(outcome->errorNumber ? outcome->errorNumber : 3251,outcome->explanation,outcome->errorSource,outcome->errorHelpFile,outcome->errorHelpContext);
                if((report ? !storedReport->recordSource.isEmpty() : !parseAccessFormLayout(QString::fromUtf8(layout)).recordSource.isEmpty()) && !document->hasDataSource()) throw VbaRuntimeError(3078,"Document record source is unavailable");
                host.check();
            }catch(...) {if(!existed && document) delete document.data();throw;}
            return VbaValue::empty();
        }
        if(key=="runcommand" || key=="requery") {
            if(auto* active=activePreviewDocument()) for(const auto& document:AccessFormPreview::projectDocuments(db_)) if(document==active) return makePreviewApplicationObject(*active)->member("DoCmd",{}).object->member(name,args);
            QVariantMap result;
            if(key=="runcommand") {if(values.size()!=1) throw VbaRuntimeError(450,"RunCommand requires one command");const int command=number(0,0);if(!QMetaObject::invokeMethod(owner_,"dispatchTableCommand",Qt::DirectConnection,Q_RETURN_ARG(QVariantMap,result),Q_ARG(int,command))) throw VbaRuntimeError(3251,"Table command is unavailable");}
            else {if(values.size()>1) throw VbaRuntimeError(450,"Requery accepts one control name");if(!values.empty() && values[0].kind!=VbaValue::Kind::Missing && !coerceVbaModuleValue(values[0],"string").asText().empty()) throw VbaRuntimeError(3251,"Named control Requery requires a form");if(!QMetaObject::invokeMethod(owner_,"requeryTableDocument",Qt::DirectConnection,Q_RETURN_ARG(QVariantMap,result))) throw VbaRuntimeError(3251,"Requery is unavailable");}
            if(!result.value("handled").toBool()) throw VbaRuntimeError(2046,"No active document supports this command");resultError(result);return VbaValue::empty();
        }
        if(key=="cancelevent") throw VbaRuntimeError(3251,"CancelEvent requires an event context");
        throw VbaRuntimeError(438,"DoCmd member is unavailable in the module application host");
    }
private:
    static void resultError(const QVariantMap& result) {if(const int error=result.value("error").toInt()) throw VbaRuntimeError(error,result.value("description").toString().toStdString());if(vbaExecutionStopRequested()) throw VbaExecutionStopped{};}
    QPointer<QWidget> owner_;QSqlDatabase db_;std::function<void()> guard_;
};
class NativeModuleApplication final : public VbaObject {
public:
    NativeModuleApplication(QWidget* owner,QSqlDatabase db,std::function<void()> guard):owner_(owner),db_(std::move(db)),guard_(std::move(guard)) {}
    std::string_view className() const override {return "Application";}
    bool isNativeType(std::string_view type) const override {return type=="access.application";}
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        if(QString::fromUtf8(name.data(),name.size()).compare("Run",Qt::CaseInsensitive)==0) return nativeApplicationRunSignature();return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        NativeStandardModuleHost host(owner_,db_,guard_);host.check();const auto key=QString::fromUtf8(name.data(),name.size()).toCaseFolded();
        if(key=="docmd") {if(!args.empty()) throw VbaRuntimeError(450,"DoCmd requires no arguments");if(auto command=commands_.lock()) return VbaValue::reference(command);auto command=std::make_shared<NativeModuleDoCmd>(owner_,db_,guard_);commands_=command;return VbaValue::reference(command);}
        if(key=="run") return runPublicStandardProcedure(host,args);
        if(key=="eval") {
            if(args.size()!=1) throw VbaRuntimeError(450,"Eval requires one expression");
            const auto expression=coerceVbaModuleValue(scalarVbaValue(args.front()),"string").asText();host.check();return host.eval(expression);
        }
        if(key=="dbengine") {if(!args.empty()) throw VbaRuntimeError(450,"DBEngine requires no arguments");return VbaValue::reference(host.databaseEngine());}
        throw VbaRuntimeError(438,"Application member is unavailable in the module execution host");
    }
private:
    QPointer<QWidget> owner_;QSqlDatabase db_;std::function<void()> guard_;std::weak_ptr<VbaObject> commands_;
};
inline std::shared_ptr<VbaObject> NativeStandardModuleHost::applicationObject() {
    check();QWidget* project=owner_;
    for(auto* ancestor=owner_->parentWidget();ancestor;ancestor=ancestor->parentWidget()) if(qobject_cast<QMainWindow*>(ancestor)) {project=ancestor;break;}
    const auto connection=db_.connectionName();
    const auto identity=QString::number(connection.size())+":"+connection+db_.databaseName()+":"+project->property("la_project_generation").toString();
    auto contexts=project->property("la_native_module_applications").toMap();
    auto object=contexts.value(identity).value<std::weak_ptr<VbaObject>>().lock();
    if(!object) {
        for(auto iterator=contexts.begin();iterator!=contexts.end();) if(iterator->value<std::weak_ptr<VbaObject>>().expired()) iterator=contexts.erase(iterator);else ++iterator;
        object=std::make_shared<NativeModuleApplication>(project,db_,guard_);contexts.insert(identity,QVariant::fromValue(std::weak_ptr<VbaObject>(object)));project->setProperty("la_native_module_applications",contexts);
    }
    return object;
}

}
