#include "native_class_catalog.hpp"
#include "vba_conditional.hpp"
#include "execution_capabilities.hpp"
#include "native_standard_module_host.hpp"
#include "vba_runtime.hpp"
#include "vba_expression.hpp"
#include "vba_module_context.hpp"
#include "vba_source_codec.hpp"
#include "vba_analyzer.hpp"
#include "access_format.hpp"
#include "dao_database.hpp"
#include "native_input_box.hpp"
#include "native_access_file_dialog.hpp"
#include "native_access_echo.hpp"
#include "native_access_cursor.hpp"
#include <QScopeGuard>
#include "native_public_procedure_run.hpp"
#include "native_application_run.hpp"
#include "native_project_context.hpp"
#include "native_project_documents.hpp"
#include "native_document_open.hpp"
#include "native_domain_lookup.hpp"
#include "native_macro_runner.hpp"
#include "native_macro_message_box.hpp"
#include "native_macro_set_value.hpp"
#include "native_find_record_arguments.hpp"
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
NativeStandardModuleHost::~NativeStandardModuleHost()=default;
std::string NativeStandardModuleHost::projectClassName(std::string_view name){check();return nativeProjectClassName(db_,name);}
bool NativeStandardModuleHost::knowsClassType(std::string_view name){check();return nativeProjectClassTypeKnown(db_,name);}
bool NativeStandardModuleHost::canConstructClass(std::string_view name){check();return nativeProjectClassAvailable(db_,name);}
std::shared_ptr<VbaObject> NativeStandardModuleHost::constructClass(std::string_view name){check();return constructNativeProjectClass(db_,owner_,name,guard_);}
NativeStandardModuleHost::NativeStandardModuleHost(QWidget* owner,QSqlDatabase db,std::function<void()> guard)
        : owner_(owner),db_(std::move(db)),guard_(std::move(guard)){}
void NativeStandardModuleHost::check() const{if(!owner_ || !db_.isOpen()) throw VbaRuntimeError(3420,"Module execution context is closed");if(owner_->thread()!=QThread::currentThread()) throw VbaRuntimeError(3251,"Module application context requires its GUI thread");guard_();}
void NativeStandardModuleHost::enterModule(std::string_view){check();enterNativeUiExecution();}
void NativeStandardModuleHost::leaveModule() noexcept {leaveNativeUiExecution();}
std::optional<std::string> NativeStandardModuleHost::field(std::string_view){throw VbaRuntimeError(438,"Me requires a real form instance");}
void NativeStandardModuleHost::setField(std::string_view,std::optional<std::string>){throw VbaRuntimeError(438,"Me requires a real form instance");}
void NativeStandardModuleHost::saveRecord(){throw VbaRuntimeError(3251,"Module editor has no current record");}
void NativeStandardModuleHost::maximizeForm(){check();(void)nativeProjectWindowState(db_,owner_,"maximize");check();}
int NativeStandardModuleHost::compareStrings(std::string_view a,std::string_view b,VbaCompareMode mode){
        if(mode==VbaCompareMode::Binary) return VbaHost::compareStrings(a,b,mode);
        return QString::compare(QString::fromUtf8(a.data(),a.size()),QString::fromUtf8(b.data(),b.size()),Qt::CaseInsensitive);
    }
bool NativeStandardModuleHost::isNumericText(std::string_view text){return isAccessNumericText(text);}
VbaValue NativeStandardModuleHost::formatValue(const VbaValue& value,std::string_view mask,bool stringResult){return formatAccessValue(value,mask,stringResult);}
VbaValue NativeStandardModuleHost::formatNumericValue(const VbaValue& value,bool currency,const VbaNumericFormatOptions& options){return formatAccessNumericValue(value,currency,options);}
void NativeStandardModuleHost::message(std::string_view value){(void)messageBox(value,0,std::nullopt);}
int NativeStandardModuleHost::messageBox(std::string_view prompt,int buttons,const std::optional<std::string>& title){check();const auto result=nativeModuleMessageBox(owner_,prompt,buttons,title);check();return result;}
std::string NativeStandardModuleHost::inputBox(const VbaInputBoxArguments& arguments){check();auto result=nativeInputBox(owner_,arguments);check();return result;}
std::shared_ptr<VbaObject> NativeStandardModuleHost::formsCollection(){check();return applicationObject()->member("Forms",{}).object;}
std::shared_ptr<VbaObject> NativeStandardModuleHost::reportsCollection(){check();return applicationObject()->member("Reports",{}).object;}
std::shared_ptr<VbaObject> NativeStandardModuleHost::currentDatabase(){check();return makeDaoDatabase(db_);}
std::shared_ptr<VbaObject> NativeStandardModuleHost::databaseEngine(){check();return makeDaoEngine(db_);}
VbaValue NativeStandardModuleHost::domainAggregate(std::string_view function,const std::vector<VbaValue>& args){check();return nativeDomainValue(db_,args,function);}
void NativeStandardModuleHost::executeSql(std::string_view sql){currentDatabase()->member("Execute",{VbaValue::string(std::string(sql))});}
std::shared_ptr<VbaModuleState> NativeStandardModuleHost::moduleState(std::string_view source) {check();return isVbaClassModule(source) ? VbaHost::moduleState(source) : standardVbaModuleState(db_,source);}
std::vector<std::shared_ptr<VbaModuleState>> NativeStandardModuleHost::globalModuleStates(){check();return publicVbaModuleStates(db_);}
std::optional<std::string> NativeStandardModuleHost::moduleForProcedure(std::string_view procedure){
        check();std::optional<std::string> found;const auto target=QString::fromUtf8(procedure.data(),procedure.size());
        for(const auto& bytes:standardVbaModuleSources(db_)) {
            const auto source=decodeVbaSource(bytes);const auto code=source.toUtf8().toStdString();if(isVbaClassModule(code)) continue;
            const auto analysis=analyzeVba(preprocessVbaConditional(code));const auto lines=source.split('\n');
            for(const auto& item:analysis.procedures) {
                if(QString::fromStdString(item.name).compare(target,Qt::CaseInsensitive)!=0) continue;
                if(item.line>0 && item.line<=lines.size() && QRegularExpression("^Private\\b",QRegularExpression::CaseInsensitiveOption).match(lines[item.line-1].trimmed()).hasMatch()) continue;
                if(found) throw VbaRuntimeError(5,"Ambiguous public procedure");found=code;
            }
        }
        return found;
    }
class NativeModuleDoCmd final : public VbaObject {
public:
    NativeModuleDoCmd(QWidget* owner,QSqlDatabase db,std::function<void()> guard):owner_(owner),db_(std::move(db)),guard_(std::move(guard)) {}
    std::string_view className() const override {return "DoCmd";}
    bool isNativeType(std::string_view type) const override {return type=="access.docmd";}
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        const auto key=QString::fromUtf8(name.data(),name.size()).toCaseFolded();
        if(key=="gotocontrol") return VbaArgumentSignature{{"ControlName"},1};
        if(key=="hourglass") return VbaArgumentSignature{{"HourglassOn"},1};
        if(key=="echo") return VbaArgumentSignature{{"EchoOn","StatusBarText"},1};
        if(key=="selectobject") return VbaArgumentSignature{{"ObjectType","ObjectName","InDatabaseWindow"},1};
        if(key=="close") return VbaArgumentSignature{{"ObjectType","ObjectName","Save"},0};
        if(key=="rename") return VbaArgumentSignature{{"NewName","ObjectType","OldName"},1};
        if(key=="deleteobject") return VbaArgumentSignature{{"ObjectType","ObjectName"},0};
        if(key=="runmacro") return VbaArgumentSignature{{"MacroName","RepeatCount","RepeatExpression"},1};
        if(key=="runsql") return VbaArgumentSignature{{"SQLStatement","UseTransaction"},1};
        if(key=="opentable") return VbaArgumentSignature{{"TableName","View","DataMode"},1};
        if(key=="openquery") return VbaArgumentSignature{{"QueryName","View","DataMode"},1};
        if(key=="findrecord") return VbaArgumentSignature{{"FindWhat","Match","MatchCase","Search","SearchAsFormatted","OnlyCurrentField","FindFirst"},1};
        if(key=="repaintobject") return VbaArgumentSignature{{"ObjectType","ObjectName"},0};
        if(key=="applyfilter") return VbaArgumentSignature{{"FilterName","WhereCondition"},0};
        if(key=="showallrecords" || key=="maximize" || key=="minimize" || key=="restore") return VbaArgumentSignature{{},0};
        if(key=="gotorecord") return VbaArgumentSignature{{"ObjectType","ObjectName","Record","Offset"},0};
        if(key=="openform") return VbaArgumentSignature{{"FormName","View","FilterName","WhereCondition","DataMode","WindowMode","OpenArgs"},1};
        if(key=="openreport") return VbaArgumentSignature{{"ReportName","View","FilterName","WhereCondition","WindowMode","OpenArgs"},1};
        if(key=="runcommand") return VbaArgumentSignature{{"Command"},1};
        if(key=="requery") return VbaArgumentSignature{{"ControlName"},0};return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        NativeStandardModuleHost host(owner_,db_,guard_);host.check();const auto key=QString::fromUtf8(name.data(),name.size()).toCaseFolded();
        if(key=="runmacro") {
            enterNativeUiExecution();const auto leave=qScopeGuard([] {leaveNativeUiExecution();});
            runNativeMacro(db_,args,[&] {host.check();},[&](const QString& expression) {return host.eval(expression.toUtf8().toStdString());},[&](const NativeMacroAction& action) {
                if(action.name=="MessageBox") nativeMacroMessageBox(owner_,action.arguments,[&](const QString& expression) {return host.eval(expression.toUtf8().toStdString());},[&] {host.check();});
                else if(action.name=="SetValue") {QPointer<AccessFormPreview> active;for(const auto& document:nativeLoadedProjectDocuments(db_,owner_)) if(document && document==activePreviewDocument()) active=document;if(active) runNativeMacroSetValue(*active,action.arguments);else nativeMacroSetValue(host,action.arguments,[&] {host.check();});}
                else if(action.name=="RunCode") {
                    const auto expression=action.arguments.front().asText();const auto target=nativeMacroFunctionName(QString::fromUtf8(expression));
                    QMap<QString,QString> names;(void)standardVbaModuleSources(db_,&names);if(names.contains(target.toCaseFolded())) throw VbaRuntimeError(3251,"RunCode Function/module name conflict");
                    const auto source=host.moduleForProcedure(target.toStdString());if(!source) throw VbaRuntimeError(2517,"RunCode public Function was not found");
                    const auto signature=nativeVbaProcedureSignature(*source,target.toStdString());if(!signature || !signature->function) throw VbaRuntimeError(3251,"RunCode requires a Function");
                    const auto result=executeVbaEventExpression({},"="+expression,host);if(result.status!=VbaExecutionStatus::Completed) throw VbaRuntimeError(result.errorNumber ? result.errorNumber : 3251,result.explanation,result.errorSource,result.errorHelpFile,result.errorHelpContext);
                }else (void)member(action.name.toStdString(),action.arguments);
            },[] {restoreNativeEcho();});return VbaValue::empty();
        }
        if(key=="echo") {auto result=nativeEcho(*owner_,args);host.check();return result;}
        std::vector<VbaValue> values;for(const auto& argument:args) {values.push_back(scalarVbaValue(argument));host.check();}
        const auto number=[&](std::size_t index,int fallback) {return index>=values.size() || values[index].kind==VbaValue::Kind::Missing ? fallback : static_cast<int>(coerceVbaModuleValue(values[index],"long").number);};
        const auto target=[&] {if(values.empty() || values[0].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(450,"Object name is required");return QString::fromStdString(coerceVbaModuleValue(values[0],"string").asText());};
        if(key=="maximize" || key=="minimize" || key=="restore") {if(!values.empty()) throw VbaRuntimeError(450,"Window command takes no arguments");auto result=nativeProjectWindowState(db_,owner_,key);host.check();return result;}
        if(key=="applyfilter" || key=="showallrecords") {auto result=nativeProjectApplyFilter(db_,owner_,values,key=="showallrecords");host.check();return result;}
        if(key=="repaintobject") {auto result=nativeProjectRepaintObject(db_,owner_,values);host.check();return result;}
        if(key=="hourglass") {
            if(values.size()!=1 || values[0].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(450,"Hourglass requires one Boolean argument");
            nativeHourglass(coerceVbaModuleValue(values[0],"boolean").boolean);host.check();return VbaValue::empty();
        }
        if(key=="gotocontrol") {
            if(values.size()!=1 || values[0].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(450,"GoToControl requires one control name");
            if(values[0].kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null control name");
            if(auto* active=activePreviewDocument()) for(const auto& document:AccessFormPreview::projectDocuments(db_)) if(document==active && document->documentKind()==AccessDocumentKind::Form) {auto result=invokeNativePreviewCommand(*active,"GoToControl",values);host.check();return result;}
            throw VbaRuntimeError(2046,"GoToControl requires an active project form");
        }
        if(key=="findrecord") {
            const auto request=nativeFindRecordArguments(values);
            if(auto* active=activePreviewDocument()) for(const auto& document:AccessFormPreview::projectDocuments(db_)) if(document==active) {auto result=invokeNativePreviewCommand(*active,"FindRecord",values);host.check();return result;}
            const auto payload=request.payload();QVariantMap result;
            if(!QMetaObject::invokeMethod(owner_,"findTableRecord",Qt::DirectConnection,Q_RETURN_ARG(QVariantMap,result),Q_ARG(QVariantMap,payload))) throw VbaRuntimeError(3251,"FindRecord data dispatch is unavailable");
            if(!result.value("handled").toBool()) throw VbaRuntimeError(2046,"No active data document supports FindRecord");resultError(result);host.check();return VbaValue::empty();
        }
        if(key=="gotorecord") {
            if(values.size()>4) throw VbaRuntimeError(450,"GoToRecord accepts at most four arguments");
            const auto omitted=[&](std::size_t index) {return index>=values.size() || values[index].kind==VbaValue::Kind::Missing || values[index].kind==VbaValue::Kind::Empty;};
            const int type=omitted(0) ? -1 : number(0,-1),action=omitted(2) ? 1 : number(2,1),offset=omitted(3) ? 1 : number(3,1);
            const auto object=omitted(1) ? QString{} : QString::fromStdString(coerceVbaModuleValue(values[1],"string").asText());
            if((type!=-1 && type!=0 && type!=1 && type!=2) || (!object.isEmpty() && type==-1)) throw VbaRuntimeError(3251,"GoToRecord object type unsupported");
            if(action<0 || action>5) throw VbaRuntimeError(5,"Invalid GoToRecord action");if((action==0 || action==1 || action==4) && offset<1) throw VbaRuntimeError(2105,"Invalid record offset");
            if(type==2 || type==-1) {
                QPointer<AccessFormPreview> selected;
                for(const auto& document:nativeLoadedProjectDocuments(db_,owner_)) if(document && document->documentKind()==AccessDocumentKind::Form && (object.isEmpty() ? document==activePreviewDocument() : document->documentName().compare(object,Qt::CaseInsensitive)==0)) {if(selected) throw VbaRuntimeError(5,"GoToRecord form target is ambiguous");selected=document;}
                if(selected) {auto forwarded=values;for(auto& value:forwarded) if(value.kind==VbaValue::Kind::Missing) value=VbaValue::empty();auto result=invokeNativePreviewCommand(*selected,"GoToRecord",forwarded);host.check();return result;}
                if(type==2) throw VbaRuntimeError(object.isEmpty() ? 2046 : 2450,"GoToRecord form is not open");
            }
            QVariantMap result;
            if(type==1) {if(!QMetaObject::invokeMethod(owner_,"navigateQueryDocument",Qt::DirectConnection,Q_RETURN_ARG(QVariantMap,result),Q_ARG(QString,object),Q_ARG(int,action),Q_ARG(int,offset))) throw VbaRuntimeError(3251,"Query navigation is unavailable");}
            else {const bool explicitTable=type==0;if(!QMetaObject::invokeMethod(owner_,"navigateTableDocument",Qt::DirectConnection,Q_RETURN_ARG(QVariantMap,result),Q_ARG(QString,object),Q_ARG(int,action),Q_ARG(int,offset),Q_ARG(bool,explicitTable))) throw VbaRuntimeError(3251,"Data navigation is unavailable");}
            if(!result.value("handled").toBool()) throw VbaRuntimeError(2046,"No active data document supports GoToRecord");resultError(result);host.check();return VbaValue::empty();
        }
        if(key=="selectobject") {
            if(values.empty() || values.size()>3 || values[0].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(450,"SelectObject requires a type and optional name/window mode");
            const int type=number(0,-1);const auto object=values.size()>1 && values[1].kind!=VbaValue::Kind::Missing ? QString::fromStdString(coerceVbaModuleValue(values[1],"string").asText()) : QString{};
            const bool database=values.size()>2 && values[2].kind!=VbaValue::Kind::Missing ? coerceVbaModuleValue(values[2],"boolean").boolean : false;
            if(type<0 || type>5 || (!database && type>3)) throw VbaRuntimeError(3251,"SelectObject type/window context is not implemented");
            if(database) {int error=0;const bool activate=true;if(!QMetaObject::invokeMethod(owner_,"selectDatabaseObject",Qt::DirectConnection,Q_RETURN_ARG(int,error),Q_ARG(int,type),Q_ARG(QString,object),Q_ARG(bool,activate))) throw VbaRuntimeError(3251,"Database selection is unavailable");if(error) throw VbaRuntimeError(error,"Database selection failed");host.check();return VbaValue::empty();}
            if(object.isEmpty()) throw VbaRuntimeError(450,"SelectObject requires an open object name");
            if(type==0 || type==1) {int error=0;const auto* method=type==0 ? "activateOpenTable" : "activateOpenQuery";if(!QMetaObject::invokeMethod(owner_,method,Qt::DirectConnection,Q_RETURN_ARG(int,error),Q_ARG(QString,object))) throw VbaRuntimeError(3251,"Open data document selection is unavailable");if(error) throw VbaRuntimeError(error,"Data document selection failed");host.check();return VbaValue::empty();}
            QPointer<AccessFormPreview> selected;for(const auto& document:AccessFormPreview::projectDocuments(db_)) if(document && type==(document->documentKind()==AccessDocumentKind::Form ? 2 : 3) && document->documentName().compare(object,Qt::CaseInsensitive)==0) {if(selected) throw VbaRuntimeError(5,"SelectObject target is ambiguous");selected=document;}
            if(!selected) throw VbaRuntimeError(type==3 ? 2451 : 2450,"SelectObject target is not open");auto forwarded=values;for(auto& value:forwarded) if(value.kind==VbaValue::Kind::Missing) value=VbaValue::empty();return invokeNativePreviewCommand(*selected,"SelectObject",forwarded);
        }
        if(key=="rename") {
            if(values.size()!=3 || values[0].kind==VbaValue::Kind::Missing || values[1].kind==VbaValue::Kind::Missing || values[2].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(3251,"Rename requires explicit new name, type and old name");
            const auto newName=target(),oldName=QString::fromStdString(coerceVbaModuleValue(values[2],"string").asText());const int type=number(1,0);QVariantMap result;
            if(!QMetaObject::invokeMethod(owner_,"renameDatabaseObject",Qt::DirectConnection,Q_RETURN_ARG(QVariantMap,result),Q_ARG(QString,newName),Q_ARG(int,type),Q_ARG(QString,oldName))) throw VbaRuntimeError(3251,"Native object rename is unavailable");resultError(result);host.check();return VbaValue::empty();
        }
        if(key=="close") {
            if(values.size()>3) throw VbaRuntimeError(450,"Close accepts at most three arguments");
            const int type=number(0,-1),save=number(2,0);if(type!=-1 && type!=0 && type!=1 && type!=2 && type!=3) throw VbaRuntimeError(3251,"Closing this object type is not implemented");if(save<0 || save>2) throw VbaRuntimeError(5,"Invalid Close save option");
            const auto object=values.size()>1 && values[1].kind!=VbaValue::Kind::Missing ? QString::fromStdString(coerceVbaModuleValue(values[1],"string").asText()) : QString{};
            if(type==0 || type==1 || type==-1) for(const auto* method:{"closeQueryDocument","closeTableDocument"}) {QVariantMap result;if(!QMetaObject::invokeMethod(owner_,method,Qt::DirectConnection,Q_RETURN_ARG(QVariantMap,result),Q_ARG(int,type),Q_ARG(QString,object),Q_ARG(int,save))) throw VbaRuntimeError(3251,"Data document close is unavailable");if(result.value("handled").toBool()) {resultError(result);host.check();return VbaValue::empty();}}
            QPointer<AccessFormPreview> selected;
            for(const auto& document:nativeLoadedProjectDocuments(db_,owner_)) if(document && (type==-1 || type==(document->documentKind()==AccessDocumentKind::Form ? 2 : 3))) {
                if(object.isEmpty()) {if(document==activePreviewDocument() && !document->isHidden()) selected=document;}
                else if(document->documentName().compare(object,Qt::CaseInsensitive)==0) {if(selected) throw VbaRuntimeError(5,"Close target is ambiguous");selected=document;}
            }
            if(!selected) throw VbaRuntimeError(type==3 ? 2451 : type==2 ? 2450 : 2046,"Close target is unavailable");
            auto forwarded=values;for(auto& value:forwarded) if(value.kind==VbaValue::Kind::Missing) value=VbaValue::empty();return invokeNativePreviewCommand(*selected,"Close",forwarded);
        }
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
            if(QApplication::activeModalWidget()) throw VbaRuntimeError(2046,"Document opening is blocked by a modal editor");
            if(key=="openform") {const auto request=nativeFormOpenArguments(values);openNativeProjectDocument(db_,owner_,AccessDocumentKind::Form,QString::fromStdString(request.name),QString::fromStdString(request.where),request.openArgs ? std::optional<QString>(QString::fromStdString(*request.openArgs)) : std::nullopt,request.windowMode,request.dataMode);}
            else {const auto request=nativeReportOpenArguments(values);openNativeProjectDocument(db_,owner_,AccessDocumentKind::Report,request.name,request.where,request.openArgs,request.windowMode,-1);}
            host.check();return VbaValue::empty();
        }
        if(key=="runcommand" || key=="requery") {
            if(auto* active=activePreviewDocument()) for(const auto& document:AccessFormPreview::projectDocuments(db_)) if(document==active) return invokeNativePreviewCommand(*active,name,args);
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
        const auto key=QString::fromUtf8(name.data(),name.size()).toCaseFolded();if(key=="forms" || key=="reports") return VbaArgumentSignature{{"Index"},0};if(key=="filedialog") return VbaArgumentSignature{{"DialogType"},1};if(key=="syscmd") return VbaArgumentSignature{{"Action","Argument2","Argument3"},1};if(QStringList{"dlookup","dmin","dmax","davg","dsum","dcount","dfirst","dlast","dvar","dstdev"}.contains(key)) return VbaArgumentSignature{{"Expr","Domain","Criteria"},2};
        if(QString::fromUtf8(name.data(),name.size()).compare("Echo",Qt::CaseInsensitive)==0) return VbaArgumentSignature{{"EchoOn","bstrStatusBarText"},1};
        if(QString::fromUtf8(name.data(),name.size()).compare("Run",Qt::CaseInsensitive)==0) return nativeApplicationRunSignature();return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        NativeStandardModuleHost host(owner_,db_,guard_);host.check();const auto key=QString::fromUtf8(name.data(),name.size()).toCaseFolded();
        if(key=="syscmd") return nativeProjectSysCmd(owner_,db_,args,[&] {host.check();});
        if(key=="currentdb") {if(!args.empty()) throw VbaRuntimeError(450,"CurrentDb requires no arguments");return VbaValue::reference(host.currentDatabase());}
        if(QStringList{"dlookup","dmin","dmax","davg","dsum","dcount","dfirst","dlast","dvar","dstdev"}.contains(key)) {
            if(args.size()<2 || args.size()>3) throw VbaRuntimeError(450,"Domain function requires two or three arguments");std::vector<VbaValue> values;for(const auto& argument:args) {values.push_back(scalarVbaValue(argument));host.check();}if(values[0].kind==VbaValue::Kind::Missing || values[1].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(450,"Domain expression and object are required");while(values.size()>2 && values.back().kind==VbaValue::Kind::Missing) values.pop_back();return nativeDomainValue(db_,values,key.toStdString());
        }
        if(key=="echo") {auto result=nativeEcho(*owner_,args);host.check();return result;}
        if(key=="docmd") {if(!args.empty()) throw VbaRuntimeError(450,"DoCmd requires no arguments");if(auto command=commands_.lock()) return VbaValue::reference(command);auto command=std::make_shared<NativeModuleDoCmd>(owner_,db_,guard_);commands_=command;return VbaValue::reference(command);}
        if(key=="filedialog") {
            if(args.size()!=1 || args[0].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(450,"FileDialog requires one type");
            const auto type=scalarVbaValue(args[0]);host.check();if(type.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null FileDialog type");const int value=native_picker::integer(type);host.check();
            const auto owner=owner_;const auto database=db_;const auto lease=guard_;auto guard=[owner,database,lease] {NativeStandardModuleHost context(owner,database,lease);context.check();};
            return VbaValue::reference(native_picker::create(owner_,value,guard));
        }
        if(key=="forms" || key=="reports") {
            if(args.size()>1) throw VbaRuntimeError(450,"Document collection accepts one optional index");auto& cache=key=="forms" ? forms_ : reports_;auto documents=cache.lock();if(!documents) {documents=nativeProjectDocuments(db_,owner_,key=="reports");cache=documents;}return args.empty() ? VbaValue::reference(documents) : documents->member("Item",args);
        }
        if(key=="screen") {if(!args.empty()) throw VbaRuntimeError(450,"Screen requires no arguments");auto screen=screen_.lock();if(!screen) {screen=nativeProjectScreen(db_,owner_);screen_=screen;}return VbaValue::reference(screen);}
        if(key=="run") return runPublicStandardProcedure(host,args);
        if(key=="eval") {
            if(args.size()!=1) throw VbaRuntimeError(450,"Eval requires one expression");
            const auto expression=coerceVbaModuleValue(scalarVbaValue(args.front()),"string").asText();host.check();return host.eval(expression);
        }
        if(key=="currentproject" || key=="currentdata") {if(!args.empty()) throw VbaRuntimeError(450,"Project/data context requires no arguments");return VbaValue::reference(nativeProjectContext(db_,owner_,key=="currentdata"));}
        if(key=="dbengine") {if(!args.empty()) throw VbaRuntimeError(450,"DBEngine requires no arguments");return VbaValue::reference(host.databaseEngine());}
        throw VbaRuntimeError(438,"Application member is unavailable in the module execution host");
    }
private:
    QPointer<QWidget> owner_;QSqlDatabase db_;std::function<void()> guard_;std::weak_ptr<VbaObject> commands_,forms_,reports_,screen_;
};
std::shared_ptr<VbaObject> NativeStandardModuleHost::applicationObject() {
    check();QWidget* project=owner_;
    for(auto* ancestor=owner_->parentWidget();ancestor;ancestor=ancestor->parentWidget()) if(qobject_cast<QMainWindow*>(ancestor)) {project=ancestor;break;}
    const auto connection=db_.connectionName();
    const auto identity=QString::number(connection.size())+":"+connection+db_.databaseName()+":"+project->property("la_project_generation").toString();
    auto contexts=project->property("la_native_module_applications").toMap();
    auto object=contexts.value(identity).value<std::weak_ptr<VbaObject>>().lock();
    if(!object) {
        for(auto iterator=contexts.begin();iterator!=contexts.end();) if(iterator->value<std::weak_ptr<VbaObject>>().expired()) iterator=contexts.erase(iterator);else ++iterator;
        const QPointer<QWidget> owner=project;const auto database=db_;const auto databaseName=db_.databaseName();const auto generation=project->property("la_project_generation");const auto outer=guard_;
        auto lease=[owner,database,databaseName,generation,outer] {if(!owner || !database.isOpen() || database.databaseName()!=databaseName || owner->property("la_project_generation")!=generation) throw VbaRuntimeError(3420,"Application belongs to an earlier project generation");outer();};
        object=std::make_shared<NativeModuleApplication>(project,db_,std::move(lease));contexts.insert(identity,QVariant::fromValue(std::weak_ptr<VbaObject>(object)));project->setProperty("la_native_module_applications",contexts);
    }
    return object;
}

}

void linaccess::NativeStandardModuleHost::executionObserved(std::string_view source,std::string_view procedure,const VbaExecutionResult& result,bool stopped) { check();recordVbaObservation(db_,source,procedure,result,stopped,"standard_module"); }
