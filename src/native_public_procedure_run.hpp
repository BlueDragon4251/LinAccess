#pragma once
#include "vba_expression.hpp"
#include "vba_runtime.hpp"
namespace linaccess {
inline VbaArgumentSignature nativeApplicationRunSignature() {
    VbaArgumentSignature signature{{"Procedure"},1};
    for(int index=1;index<=30;++index) signature.names.push_back("Arg"+std::to_string(index));
    signature.objectArguments.assign(31,true);signature.objectArguments[0]=false;
    return signature;
}
inline VbaValue runPublicStandardProcedure(VbaHost& host,const std::vector<VbaValue>& arguments) {
    if(arguments.empty() || arguments.size()>31) throw VbaRuntimeError(450,"Run requires a procedure and at most 30 arguments");
    const auto name=coerceVbaModuleValue(scalarVbaValue(arguments.front()),"string").asText();
    if(name.empty() || name.size()>255 || name.find('\0')!=std::string::npos) throw VbaRuntimeError(5,"Invalid Run procedure name");
    if(name.find('.')!=std::string::npos) throw VbaRuntimeError(3251,"Run across referenced projects is not implemented");
    static thread_local unsigned depth=0;
    if(depth>=32) throw VbaRuntimeError(28,"Run call depth limit exceeded");
    struct Guard {unsigned& depth;explicit Guard(unsigned& value):depth(value) {++depth;}~Guard() {--depth;}} guard(depth);
    const auto module=host.moduleForProcedure(name);
    if(!module) throw VbaRuntimeError(2517,"Run public procedure was not found");
    auto actuals=std::vector<VbaValue>(arguments.begin()+1,arguments.end());
    const auto result=executeVbaProcedureWithArguments(*module,name,host,actuals);
    if(result.status!=VbaExecutionStatus::Completed)
        throw VbaRuntimeError(result.errorNumber ? result.errorNumber : 3251,result.explanation,result.errorSource,result.errorHelpFile,result.errorHelpContext);
    return result.typedReturnValue ? *result.typedReturnValue : VbaValue::empty();
}
}
