#pragma once
#include "native_macro.hpp"
#include <QScopeGuard>
#include <functional>
namespace linaccess {
struct NativeMacroFrame {QString connection,database,name;};
inline QVector<NativeMacroFrame>& nativeMacroFrames() {static thread_local QVector<NativeMacroFrame> frames;return frames;}
inline bool nativeProjectMacrosActive(QSqlDatabase db) {
    for(const auto& frame:nativeMacroFrames()) if(frame.connection==db.connectionName() && frame.database==db.databaseName()) return true;return false;
}
inline bool nativeMacroObjectActive(QSqlDatabase db,const QString& name) {
    for(const auto& frame:nativeMacroFrames()) if(frame.connection==db.connectionName() && frame.database==db.databaseName() && frame.name.compare(name,Qt::CaseInsensitive)==0) return true;return false;
}
inline void runNativeMacro(QSqlDatabase db,const std::vector<VbaValue>& args,
    const std::function<void()>& check,const std::function<VbaValue(const QString&)>& evaluate,
    const std::function<void(const NativeMacroAction&)>& dispatch) {
    VbaExecutionScope executionScope;
    if(args.empty() || args.size()>3) throw VbaRuntimeError(450,"RunMacro requires a name and optional repeat count");
    const auto scalar=[&](const VbaValue& value) {auto result=scalarVbaValue(value);check();return result;};
    const auto macro=coerceVbaModuleValue(scalar(args[0]),"string");
    const bool counted=args.size()>1 && args[1].kind!=VbaValue::Kind::Missing;
    const auto count=counted ? coerceVbaModuleValue(scalar(args[1]),"long").number : 1;
    const auto expression=args.size()>2 && args[2].kind!=VbaValue::Kind::Missing ? nativeMacroRepeatExpression(QString::fromStdString(coerceVbaModuleValue(scalar(args[2]),"string").asText())) : std::optional<QString>{};
    if(count<1 || count>1000) throw VbaRuntimeError(3251,"Native macro repeat count requires 1 to 1000");
    static thread_local int depth=0,visited=0;
    if(depth>=16) throw VbaRuntimeError(3251,"Native macro call depth exceeded");
    if(!depth) visited=0;++depth;const auto budget=qScopeGuard([] {if(--depth==0) visited=0;});
    check();const auto actions=loadNativeMacro(db,QString::fromStdString(macro.asText()));
    nativeMacroFrames().push_back({db.connectionName(),db.databaseName(),QString::fromStdString(macro.asText()).section('.',0,0)});
    const auto frame=qScopeGuard([] {nativeMacroFrames().removeLast();});
    if(!expression && count*actions.size()>10000) throw VbaRuntimeError(3251,"Native macro action budget exceeded");
    for(int repeat=0;;++repeat) {
        if((counted || !expression) && repeat>=count) return;
        check();if(vbaExecutionStopRequested()) throw VbaExecutionStopped{};
        if(expression) {
            try {const auto condition=scalarVbaValue(evaluate(*expression));check();if(!condition.truthy()) return;}
            catch(const VbaRuntimeError& error) {throw VbaRuntimeError(error.number,"Makro "+macro.asText()+", Wiederholung "+std::to_string(repeat+1)+": "+error.what(),error.source,error.helpFile,error.helpContext);}
        }
        if(repeat>=1000) throw VbaRuntimeError(3251,"RunMacro repeat expression iteration budget exceeded");
        for(int index=0;index<actions.size();++index) {
        check();if(vbaExecutionStopRequested()) throw VbaExecutionStopped{};
        if(++visited>10000) throw VbaRuntimeError(3251,"Nested native macro action budget exceeded");
        const auto& action=actions[index];
        try {
            if(action.condition) {const auto condition=scalarVbaValue(evaluate(*action.condition));check();if(!condition.truthy()) continue;}
            if(action.name=="StopMacro") return;
            dispatch(action);
        }catch(const VbaRuntimeError& error) {throw VbaRuntimeError(error.number,"Makro "+macro.asText()+", Aktion "+std::to_string(index+1)+" ("+action.name.toStdString()+"): "+error.what(),error.source,error.helpFile,error.helpContext);}
        }
    }
}
}
