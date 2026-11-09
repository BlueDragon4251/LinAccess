#include "vba_debugger.hpp"
#include "vba_condition.hpp"
#include "vba_expression.hpp"
#include "vba_decimal.hpp"
#include <vector>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
namespace {
struct Host final:linaccess::VbaHost {
    std::optional<std::string> field(std::string_view) override {return {};}
    linaccess::VbaValue expressionValue(std::string_view name,bool) override {if(name=="Error") return linaccess::VbaValue::longInteger(17);throw linaccess::VbaRuntimeError(438,"Unknown technical expression field");}
    void setField(std::string_view,std::optional<std::string>) override {assert(false);}
    void saveRecord() override {assert(false);}
    void message(std::string_view) override {assert(false);}
};
struct DefaultValue final:linaccess::VbaObject {
    linaccess::VbaValue value;int reads{};
    bool hasScalarDefault() const override {return true;}
    linaccess::VbaValue scalarValue() override {++reads;return value;}
    linaccess::VbaValue member(std::string_view,const std::vector<linaccess::VbaValue>&) override {assert(false);return {};}
};
}
void testVbaConditions() {
    using namespace linaccess;Host host;
    const auto run=[&](const std::string& body,const VbaValue& value) {std::vector<VbaValue> args{value};return executeVbaProcedureWithArguments("Function Probe(v As Variant) As Long\nDim n As Long\n"+body+"\nEnd Function","Probe",host,args);};
    const auto number=[&](const std::string& body,const VbaValue& value,int expected) {const auto result=run(body,value);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue && result.typedReturnValue->number==expected);};
    for(const auto& [value,truth]:std::vector<std::pair<VbaValue,bool>>{{VbaValue{},false},{VbaValue::empty(),false},{VbaValue::logical(false),false},{VbaValue::logical(true),true},{VbaValue::integer(0),false},{VbaValue::numeric(-0.5),true},{VbaValue::currency(1),true},{decimalLiteral("0.0000000000000000000000000001"),true},{decimalLiteral("0"),false},{VbaValue::string("True"),true},{VbaValue::string("fAlSe"),false},{VbaValue::string("#TRUE#"),true},{VbaValue::string("#FALSE#"),false},{VbaValue::string(" 0 "),false},{VbaValue::string(" 2.5 "),true}}) {
        number("Probe=IIf(v,1,2)",value,truth?1:2);
        number("If(v)Then\nProbe=1\nElse\nProbe=2\nEndIf",value,truth?1:2);
        number("If False Then\nProbe=9\nElseIf(v)Then\nProbe=1\nElse\nProbe=2\nEnd\t If",value,truth?1:2);
        number("If(v)Then Probe=1 Else Probe=2",value,truth?1:2);
        number("Do While(v)\nProbe=1\nExit\tDo\nLoop",value,truth?1:0);
        number("Do Until(v)\nProbe=1\nExit  Do\nLoop",value,truth?0:1);
        number("While(v)\nProbe=1\nExit Function\nWend",value,truth?1:0);
        number("Do\nn=n+1\nIf n=2 Then Exit Do\nLoop While(v)\nProbe=n",value,truth?2:1);
        number("Do\nn=n+1\nIf n=2 Then Exit Do\nLoop Until(v)\nProbe=n",value,truth?1:2);
        number("If v Then Else Probe=3",value,truth?0:3);
        number("If v Then Probe=3 Else",value,truth?3:0);
        number("If v Then Else",value,0);
        auto object=std::make_shared<DefaultValue>();object->value=value;number("If(v)Then Probe=1 Else Probe=2",VbaValue::reference(object),truth?1:2);assert(object->reads==1);
    }
    for(const auto& value:std::vector<VbaValue>{VbaValue::string("invalid"),VbaValue::string("#true#"),VbaValue::string("#false#"),VbaValue::errorValue(2015),VbaValue::byteArray({1})}) {
        const auto result=run("If v Then Probe=1",value);assert(result.status==VbaExecutionStatus::Error && result.errorNumber==13);
    }
    assert(run("If v Then Probe=1",VbaValue::reference(nullptr)).errorNumber==91);
    auto object=std::make_shared<DefaultValue>();object->value=VbaValue::errorValue(2015);
    number("If False Then\nIf v Then Probe=1\nEnd If\nProbe=3",VbaValue::reference(object),3);assert(object->reads==0);
    number("If True Then Probe=4 Else If v Then Probe=8",VbaValue::reference(object),4);assert(object->reads==0);
    number("If False Then Probe=4 Else If v Then Else Probe=8",VbaValue{},8);
    number("On Error GoTo recovered\nIf v Then Probe=7\nExit Function\nrecovered:\nv=\"True\"\nResume",VbaValue::string("invalid"),7);
    auto reference=std::make_shared<DefaultValue>();reference->value=VbaValue::string("False");number("With(v)\nProbe=5\nEnd\t With",VbaValue::reference(reference),5);assert(reference->reads==0);
    number("Select Case(v)\nCase(1)\nProbe=4\nCase Else\nProbe=2\nEnd Select",VbaValue::integer(1),4);
    const auto invalid=run("Probe=5\nSelect Case v\nCase\nProbe=9\nEnd Select",VbaValue::integer(1));assert(invalid.status!=VbaExecutionStatus::Completed);
    number("On Error Resume Next\nIf v Then Probe=90 Else Probe=99\nProbe=Err.Number",VbaValue::string("invalid"),13);
    number("On Error Resume Next\nIf True Then\nIf v Then Probe=90 Else Probe=99\nProbe=Probe+7\nEnd If\nProbe=Probe+Err.Number",VbaValue::string("invalid"),20);
    number("On Error Resume Next\nIf False Then\nProbe=90\nElseIf v Then\nProbe=91\nElseIf True Then\nProbe=92\nElse\nProbe=99\nEnd If\nProbe=Err.Number",VbaValue::string("invalid"),13);
    number("On Error GoTo recovered\nIf False Then\nProbe=90\nElseIf v Then\nProbe=7\nEnd If\nExit Function\nrecovered:\nv=\"True\"\nResume",VbaValue::string("invalid"),7);
    number("On Error GoTo recovered\nIf v Then Probe=90 Else Probe=99\nProbe=3\nExit Function\nrecovered:\nResume Next",VbaValue::string("invalid"),3);
    {int errors=0;VbaDebuggerSession debugger([&](const VbaDebugPause& pause) {if(pause.reason!=VbaDebugReason::Error) return VbaDebugAction::Continue;assert(pause.error->number==13 && pause.recoverable);++errors;return VbaDebugAction::SkipError;});number("If v Then Probe=90 Else Probe=99\nProbe=Err.Number",VbaValue::string("invalid"),13);assert(errors==1);}
    for(const auto& opening:std::vector<std::string>{"While v","Do While v","Do Until v"}) {const auto closing=opening=="While v" ? "Wend" : "Loop";number("On Error Resume Next\n"+opening+"\nProbe=99\n"+closing+"\nProbe=Probe+Err.Number",VbaValue::string("invalid"),13);}
    number("On Error Resume Next\nFor n=v To 2\nProbe=99\nNext n\nProbe=Probe+Err.Number",VbaValue::string("invalid"),13);
    number("On Error Resume Next\nDim item As Variant\nFor Each item In v\nProbe=99\nNext\nProbe=Probe+Err.Number",VbaValue::string("invalid"),424);
    number("On Error Resume Next\nWith v\nProbe=99\nEnd With\nProbe=Probe+Err.Number",VbaValue::integer(1),424);
    number("On Error Resume Next\nIf True Then\nWith v\nProbe=99\nEnd With\nProbe=Probe+7\nEnd If\nProbe=Probe+Err.Number",VbaValue::integer(1),431);
    number("On Error GoTo recovered\nWith v\nProbe=7\nEnd With\nExit Function\nrecovered:\nSet v=New Collection\nResume",VbaValue::integer(1),7);
    number("On Error GoTo recovered\nWith v\nProbe=99\nEnd With\nProbe=3\nExit Function\nrecovered:\nResume Next",VbaValue::integer(1),3);
    number("On Error GoTo recovered\nDo While v\nProbe=99\nLoop\nProbe=3\nExit Function\nrecovered:\nResume Next",VbaValue::string("invalid"),3);
    number("On Error Resume Next\nDim i As Long\nFor i=1 To 2\nFor n=1 To v\nProbe=99\nNext n\nProbe=Probe+1\nNext i\nProbe=Probe+Err.Number",VbaValue::string("invalid"),15);
    number("On Error GoTo recovered\nFor n=1 To v\nProbe=Probe+1\nNext n\nExit Function\nrecovered:\nv=2\nResume",VbaValue::string("invalid"),2);
    {int errors=0;VbaDebuggerSession debugger([&](const VbaDebugPause& pause) {if(pause.reason!=VbaDebugReason::Error) return VbaDebugAction::Continue;assert(pause.error->number==13 && pause.recoverable);++errors;return VbaDebugAction::SkipError;});number("Do While v\nProbe=99\nLoop\nProbe=Probe+Err.Number",VbaValue::string("invalid"),13);assert(errors==1);}
    assert(evaluateAccessExpression("Error",host).number==17 && evaluateAccessExpression("Error(0)",host).text.empty());
    const auto expression=[&](const char* text) {return evaluateVbaExpression(text,host,{});};assert(expression("CBool(\"#TRUE#\")").boolean && !expression("CBool(\"#FALSE#\")").boolean);
    for(const auto& ending:std::vector<std::string>{"EndIf","End   If","End\tIf","EndIf: Probe=9"}) number("If v Then\nProbe=9\n"+ending,VbaValue::logical(true),9);
}
