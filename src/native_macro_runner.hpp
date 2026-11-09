#pragma once
#include "native_macro.hpp"
#include <QScopeGuard>
#include <functional>
namespace linaccess {
inline void runNativeMacro(QSqlDatabase db,const std::vector<VbaValue>& args,
    const std::function<void()>& check,const std::function<VbaValue(const QString&)>& evaluate,
    const std::function<void(const NativeMacroAction&)>& dispatch) {
    VbaExecutionScope executionScope;
    if(args.empty() || args.size()>3) throw VbaRuntimeError(450,"RunMacro requires a name and optional repeat count");
    const auto scalar=[&](const VbaValue& value) {auto result=scalarVbaValue(value);check();return result;};
    const auto macro=coerceVbaModuleValue(scalar(args[0]),"string");
    const auto count=args.size()>1 && args[1].kind!=VbaValue::Kind::Missing ? coerceVbaModuleValue(scalar(args[1]),"long").number : 1;
    if(args.size()>2 && args[2].kind!=VbaValue::Kind::Missing) throw VbaRuntimeError(3251,"RunMacro RepeatExpression is not implemented");
    if(count<1 || count>1000) throw VbaRuntimeError(3251,"Native macro repeat count requires 1 to 1000");
    static thread_local int depth=0,visited=0;
    if(depth>=16) throw VbaRuntimeError(3251,"Native macro call depth exceeded");
    if(!depth) visited=0;++depth;const auto budget=qScopeGuard([] {if(--depth==0) visited=0;});
    check();const auto actions=loadNativeMacro(db,QString::fromStdString(macro.asText()));
    if(count*actions.size()>10000) throw VbaRuntimeError(3251,"Native macro action budget exceeded");
    for(int repeat=0;repeat<count;++repeat) for(int index=0;index<actions.size();++index) {
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
