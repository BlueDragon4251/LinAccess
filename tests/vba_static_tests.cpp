#include "vba_def_type_fixture.hpp"
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
    result=executeVbaProcedure("DefLngLng A-Z\nStatic Sub GuardedType()\nMsgBox \"unreached\"\nEnd Sub","GuardedType",host);assert(result.status==VbaExecutionStatus::Unsupported && result.errorNumber==3251);
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
    result=executeVbaProcedure(R"VBA(DefLngLng A-Z
Sub GuardedOrdinary
MsgBox "unreached"
End Sub)VBA","GuardedOrdinary",host);assert(result.status==VbaExecutionStatus::Unsupported && result.errorNumber==3251);
    const std::string moduleField="Private known As Long\nFunction KnownProbe As Long\nknown=known+1\nKnownProbe=known\nEnd Function";
    for(int call=1;call<=2;++call){result=executeVbaProcedure(moduleField,"KnownProbe",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==call);}
    const auto defRun=[&](const char* name){auto outcome=executeVbaProcedure(nativeDefTypeFixture,name,host);if(outcome.status!=VbaExecutionStatus::Completed)std::cerr<<"Def type "<<name<<": "<<outcome.explanation<<" line "<<outcome.line<<" error "<<outcome.errorNumber<<"\n";assert(outcome.status==VbaExecutionStatus::Completed && outcome.typedReturnValue);return *outcome.typedReturnValue;};
    assert(defRun("TypeProbe").text=="2:3:8:11:17:6:7:5:4:9:0");
    assert(defRun("PropertyProbe").text=="cell");
    assert(defRun("XTypes").text=="17:2:8:7:0");assert(defRun("SValue").text=="typed" && defRun("LValue").numericType==VbaValue::NumericType::Long);
    const auto firstObject=defRun("OValue");assert(firstObject.object && defRun("OValue").object==firstObject.object);
    assert(defRun("ParameterProbe").text=="value:0:2");assert(defRun("ArrayProbe").number==819533);assert(defRun("MismatchProbe").number==1300);assert(defRun("OverrideProbe").number==805);assert(defRun("OverflowProbe").number==602);
    for(int call=1;call<=2;++call){assert(defRun("LStatic").number==call && defRun("LOrdinary").number==1);assert(defRun("ModuleProbe").text==std::to_string(call)+":"+std::string(call,'x')+":3:8");}
    assert(defRun("CurrencyLoop").currencyUnits==6 && defRun("ImplicitArrayProbe").number==819509);
    result=executeVbaProcedure("DefInt I\nSub InvalidLoop\nMsgBox \"unreached\"\nFor Each item In Array(1,2)\nNext\nEnd Sub","InvalidLoop",host);assert(result.status==VbaExecutionStatus::Unsupported && result.errorNumber==3251);
    const auto parameterSignature=nativeVbaProcedureSignature(nativeDefTypeFixture,"Bump");assert(parameterSignature && parameterSignature->names[0]=="index");
    const auto returnSignature=nativeVbaProcedureSignature(nativeDefTypeFixture,"LValue");assert(returnSignature && returnSignature->resultType=="long");
    for(const auto& directive:std::vector<std::pair<std::string,int>>{{"DefInt A-C\nDefStr B",5},{"DefInt A-Z\nDefVar Q",5},{"DefInt A,A",5},{"DefInt AB",3251},{"DefInt",3251},{"DefInt 1",3251},{"DefLngPtr A",3251}}){result=executeVbaProcedure(directive.first+"\nSub InvalidDefs\nMsgBox \"unreached\"\nEnd Sub","InvalidDefs",host);assert(result.status==VbaExecutionStatus::Unsupported && result.errorNumber==directive.second);}
    result=executeVbaProcedure("DefLng Z-X\nFunction Descending As Long\nDescending=VarType(xValue)*100+VarType(zValue)\nEnd Function","Descending",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==303);
    result=executeVbaProcedure("DefLng A-Z\nFunction Universal As Long\nDim äValue\nUniversal=VarType(äValue)\nEnd Function","Universal",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==3);
    result=executeVbaProcedure("Option Explicit\nDefLng A-Z\nFunction Strict As Long\nStrict=missing\nEnd Function","Strict",host);assert(result.status==VbaExecutionStatus::Error && result.errorNumber==5);
    result=executeVbaProcedure("DefVar A-Z\nFunction MissingProbe(Optional value) As Boolean\nMissingProbe=IsMissing(value)\nEnd Function","MissingProbe",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->truthy());
    result=executeVbaProcedure("#If False Then\nDefLngLng A-Z\n#End If\nDefStr S\nFunction Conditional As Long\nConditional=VarType(sValue)\nEnd Function","Conditional",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==8);

}
