#include "vba_condition.hpp"
#include "artifact_execution_observer.hpp"
#include "native_macro_runner.hpp"
#include "native_access_warnings.hpp"
namespace linaccess {
QVector<NativeMacroFrame>& nativeMacroFrames() {static thread_local QVector<NativeMacroFrame> frames;return frames;}
bool nativeProjectMacrosActive(QSqlDatabase db) {
    for(const auto& frame:nativeMacroFrames()) if(frame.connection==db.connectionName() && frame.database==db.databaseName()) return true;return false;
}
bool nativeMacroObjectActive(QSqlDatabase db,const QString& name) {
    for(const auto& frame:nativeMacroFrames()) if(frame.connection==db.connectionName() && frame.database==db.databaseName() && frame.name.compare(name,Qt::CaseInsensitive)==0) return true;return false;
}
bool& nativeMacroStopAllPending() {static thread_local bool pending=false;return pending;}
void runNativeMacro(QSqlDatabase db,const std::vector<VbaValue>& args,
    const std::function<void()>& check,const std::function<VbaValue(const QString&)>& evaluate,
    const std::function<void(const NativeMacroAction&)>& dispatch,const std::function<void()>& restoreUi) {
    VbaExecutionScope executionScope;
    if(!nativeMacroFrames().isEmpty() && nativeMacroStopAllPending()) return;
    if(args.empty() || args.size()>3) throw VbaRuntimeError(450,"RunMacro requires a name and optional repeat count");
    const auto scalar=[&](const VbaValue& value) {auto result=scalarVbaValue(value);check();return result;};
    const auto macro=coerceVbaModuleValue(scalar(args[0]),"string");
    const auto macroName=QString::fromStdString(macro.asText());QByteArray revision;
    QString observedStatus="completed";int observedError=0;const auto observation=qScopeGuard([&] {try {recordArtifactExecution(db,"macro",macroName,revision,observedStatus,observedError);}catch(...) {}});
    try {
    const bool counted=args.size()>1 && args[1].kind!=VbaValue::Kind::Missing;
    const auto count=counted ? coerceVbaModuleValue(scalar(args[1]),"long").number : 1;
    const auto expression=args.size()>2 && args[2].kind!=VbaValue::Kind::Missing ? nativeMacroRepeatExpression(QString::fromStdString(coerceVbaModuleValue(scalar(args[2]),"string").asText())) : std::optional<QString>{};
    if(count<1 || count>1000) throw VbaRuntimeError(3251,"Native macro repeat count requires 1 to 1000");
    static thread_local int depth=0,visited=0;
    if(depth>=16) throw VbaRuntimeError(3251,"Native macro call depth exceeded");
    if(!depth) {visited=0;nativeMacroStopAllPending()=false;}++depth;const auto budget=qScopeGuard([] {if(--depth==0) {visited=0;nativeMacroStopAllPending()=false;}});
    check();try {QSqlQuery query(db);query.prepare("SELECT definition FROM main._la_objects WHERE kind='macro' AND name=? COLLATE NOCASE LIMIT 2");query.addBindValue(macroName.section('.',0,0));if(query.exec() && query.next()) {revision=query.value(0).toByteArray();if(query.next()) revision.clear();}}catch(...) {}
    const auto actions=loadNativeMacro(db,QString::fromStdString(macro.asText()));
    bool warningsTouched=false;const auto warnings=qScopeGuard([&] {if(warningsTouched) restoreNativeWarnings();});
    nativeMacroFrames().push_back({db.connectionName(),db.databaseName(),QString::fromStdString(macro.asText()).section('.',0,0)});
    const auto frame=qScopeGuard([] {nativeMacroFrames().removeLast();});
    if(!expression && count*actions.size()>10000) throw VbaRuntimeError(3251,"Native macro action budget exceeded");
    for(int repeat=0;;++repeat) {
        if(nativeMacroStopAllPending()) {observedStatus="stopped";return;}
        if((counted || !expression) && repeat>=count) return;
        check();if(vbaExecutionStopRequested()) throw VbaExecutionStopped{};
        if(expression) {
            try {const auto condition=scalarVbaValue(evaluate(*expression));check();if(nativeMacroStopAllPending()) {observedStatus="stopped";return;}if(!vbaBooleanCondition(condition)) return;}
            catch(const VbaRuntimeError& error) {throw VbaRuntimeError(error.number,"Makro "+macro.asText()+", Wiederholung "+std::to_string(repeat+1)+": "+error.what(),error.source,error.helpFile,error.helpContext);}
        }
        if(repeat>=1000) throw VbaRuntimeError(3251,"RunMacro repeat expression iteration budget exceeded");
        std::optional<bool> blockCondition;
        for(int index=0;index<actions.size();++index) {
        if(nativeMacroStopAllPending()) {observedStatus="stopped";return;}
        check();if(vbaExecutionStopRequested()) throw VbaExecutionStopped{};
        if(++visited>10000) throw VbaRuntimeError(3251,"Nested native macro action budget exceeded");
        const auto& action=actions[index];
        try {
            if(action.condition) {
                if(*action.condition=="...") {if(!blockCondition) throw VbaRuntimeError(3251,"Macro condition continuation has no active block");if(!*blockCondition) continue;}
                else {const auto condition=scalarVbaValue(evaluate(*action.condition));check();if(nativeMacroStopAllPending()) {observedStatus="stopped";return;}blockCondition=vbaBooleanCondition(condition);if(!*blockCondition) continue;}
            } else blockCondition.reset();
            if(action.name=="StopAllMacros") {observedStatus="stopped";nativeMacroStopAllPending()=true;if(restoreUi) restoreUi();return;}
            if(action.name=="StopMacro") {observedStatus="stopped";return;}
            dispatch(action);if(action.name=="SetWarnings") warningsTouched=true;
        }catch(const VbaRuntimeError& error) {throw VbaRuntimeError(error.number,"Makro "+macro.asText()+", Aktion "+std::to_string(index+1)+" ("+action.name.toStdString()+"): "+error.what(),error.source,error.helpFile,error.helpContext);}
        }
    }
    }catch(const VbaExecutionStopped&) {observedStatus="stopped";throw;}
    catch(const VbaRuntimeError& failure) {observedStatus="error";observedError=failure.number;throw;}
    catch(...) {observedStatus="error";throw;}

}
}
