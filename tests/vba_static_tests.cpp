#include "vba_expression.hpp"
#include "vba_static_fixture.hpp"
#include "vba_analyzer.hpp"
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>
#include <algorithm>
#include <cctype>
namespace {
struct StaticHost final:linaccess::VbaHost {
    std::optional<std::string> field(std::string_view)override{return {};}
    void setField(std::string_view,std::optional<std::string>)override{assert(false);}
    void saveRecord()override{assert(false);}
    void message(std::string_view)override{assert(false);}
    std::optional<std::string> moduleForProcedure(std::string_view name)override {
        auto key=std::string(name);std::transform(key.begin(),key.end(),key.begin(),[](unsigned char ch){return static_cast<char>(std::tolower(ch));});
        if(key=="externalzero") return "Public Function ExternalZero As Long\nExternalZero=19\nEnd Function";
        if(key=="externalsub") return "Public Sub ExternalSub\nMsgBox \"unreached\"\nEnd Sub";
        return {};
    }
};
}
void testVbaStaticProcedures(){using namespace linaccess;StaticHost host,other;
    for(int call=1;call<=2;++call){
        auto result=executeVbaProcedure(nativeStaticFixture,"Retained",host);if(result.status!=VbaExecutionStatus::Completed)std::cerr<<"Static declared: "<<result.explanation<<" line "<<result.line<<" error "<<result.errorNumber<<"\n";assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text==std::to_string(call)+":"+std::to_string(call*2)+":"+std::to_string(call*3)+":"+std::to_string(call)+":A"+std::to_string(call));
        result=executeVbaProcedure(nativeStaticFixture,"Implicit",host);if(result.status!=VbaExecutionStatus::Completed)std::cerr<<"Static implicit: "<<result.explanation<<" line "<<result.line<<" error "<<result.errorNumber<<"\n";assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text==std::to_string(call)+":"+std::to_string(call)+":"+std::to_string(call)+":"+std::to_string(call*3)+":"+std::to_string(call));
        result=executeVbaProcedure(nativeStaticFixture,"Trailing",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==call);
    }
    auto result=executeVbaProcedure(nativeStaticFixture,"OtherImplicit",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==100);
    result=executeVbaProcedure(nativeStaticFixture,"Retained",other);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="1:2:3:1:A1");
    result=executeVbaProcedure(nativeStaticFixture+"\n' distinct source revision\n","Retained",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="1:2:3:1:A1");
    std::vector<VbaValue> arguments{VbaValue::longInteger(2)};result=executeVbaProcedureWithArguments(nativeStaticFixture,"Recursive",host,arguments);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==3);
    arguments={VbaValue::logical(true)};result=executeVbaProcedureWithArguments(nativeStaticFixture,"FreshResult",host,arguments);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==5);
    arguments={VbaValue::logical(false)};result=executeVbaProcedureWithArguments(nativeStaticFixture,"FreshResult",host,arguments);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==0);
    for(int call=1;call<=2;++call){arguments={VbaValue::longInteger(call*10)};result=executeVbaProcedureWithArguments(nativeStaticFixture,"FreshParameter",host,arguments);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==call*11+1 && arguments[0].number==call*10);}
    for(int call=1;call<=2;++call){arguments={VbaValue::longInteger(7)};result=executeVbaProperty(nativeStaticFixture,"Accumulated",VbaPropertyAccess::Let,host,arguments);assert(result.status==VbaExecutionStatus::Completed);arguments.clear();result=executeVbaProperty(nativeStaticFixture,"Accumulated",VbaPropertyAccess::Get,host,arguments);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==7+call);}
    const auto state=host.moduleState(nativeStaticFixture);assert(state->procedureStatics.at("property let accumulated").at("writes").number==2);
    for(int call=1;call<=2;++call){result=executeVbaProcedure(nativeStaticFixture,"Failure",host);assert(result.status==VbaExecutionStatus::Error && result.errorNumber==11 && state->procedureStatics.at("failure").at("count").number==call);}
    result=executeVbaProcedure("Option Explicit\nPublic Static Function Declared() As Long\nDim count As Long\ncount=count+1\nDeclared=count\nEnd Function","Declared",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==1);
    result=executeVbaProcedure("Option Explicit\nStatic Function Guarded() As Long\nGuarded=missing\nEnd Function","Guarded",host);assert(result.status==VbaExecutionStatus::Error && result.errorNumber==5);
    result=executeVbaProcedure("Option Explicit\nFunction Ordinary() As Long\nOrdinary=missing\nEnd Function","Ordinary",host);assert(result.status==VbaExecutionStatus::Error && result.errorNumber==5);
    result=executeVbaProcedure("DefInt A-Z\nStatic Sub GuardedType()\nMsgBox \"unreached\"\nEnd Sub","GuardedType",host);assert(result.status==VbaExecutionStatus::Unsupported && result.errorNumber==3251);
    result=executeVbaProcedure("Static Function Suffix() As Long\nSuffix=VarType(undeclared$)\nEnd Function","Suffix",host);assert(result.status==VbaExecutionStatus::Error && result.errorNumber==3251);
    const auto inventory=analyzeVba(nativeStaticFixture);assert(!inventory.procedures.empty() && inventory.procedures.front().name=="Retained");

    auto transient=nativeStaticFixture;const auto modifier=transient.find("Static Function Implicit()");assert(modifier!=std::string::npos);transient.erase(modifier,7);
    for(int call=0;call<2;++call){result=executeVbaProcedure(transient,"Implicit",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="1:1:1:3:1");}
    assert(!host.moduleState(transient)->procedureStatics.contains("implicit"));
    result=executeVbaProcedure("Function EmptyProbe As Long\nEmptyProbe=VarType(uninitialized)\nEnd Function","EmptyProbe",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==0);
    result=executeVbaProcedure("Option Explicit\nSub StrictWrite\nuninitialized=1\nEnd Sub","StrictWrite",host);assert(result.status==VbaExecutionStatus::Error && result.errorNumber==5);
    result=executeVbaProcedure("Function ExternalProbe As Long\nExternalProbe=ExternalZero\nEnd Function","ExternalProbe",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==19);
    result=executeVbaProcedure("Function ExternalGuard As Long\nExternalGuard=ExternalSub\nEnd Function","ExternalGuard",host);assert(result.status==VbaExecutionStatus::Error && result.errorNumber==3251);
    result=executeVbaProcedure("Sub ExternalWrite\nExternalZero=1\nEnd Sub","ExternalWrite",host);assert(result.status==VbaExecutionStatus::Error && result.errorNumber==3251);
    result=executeVbaProcedure(R"VBA(DefInt A-Z
Sub GuardedOrdinary
MsgBox "unreached"
End Sub)VBA","GuardedOrdinary",host);assert(result.status==VbaExecutionStatus::Unsupported && result.errorNumber==3251);
    const std::string moduleField="Private known As Long\nFunction KnownProbe As Long\nknown=known+1\nKnownProbe=known\nEnd Function";
    for(int call=1;call<=2;++call){result=executeVbaProcedure(moduleField,"KnownProbe",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==call);}
}
