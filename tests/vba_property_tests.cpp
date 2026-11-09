#include "vba_expression.hpp"
#include "vba_analyzer.hpp"
#include "vba_property_fixture.hpp"
#include "vba_parameterless_fixture.hpp"
#include "vba_default_property_fixture.hpp"
#include "vba_byref_property_fixture.hpp"
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>
namespace {
struct PropertyHost final:linaccess::VbaHost {
    int effects{};
    std::optional<std::string> field(std::string_view)override{return {};}
    void setField(std::string_view,std::optional<std::string>)override{assert(false);}
    void saveRecord()override{assert(false);}
    void message(std::string_view)override{++effects;}
};
}
void testVbaProperties(){using namespace linaccess;PropertyHost host,other;
    const auto inventory=analyzeVba(nativePropertyFixture);assert(vbaProcedureExecutionName(inventory.procedures.front())=="property get Stored" && vbaProcedureExecutionName(inventory.procedures[1])=="property let Stored");
    std::vector<VbaValue> arguments{VbaValue::longInteger(7)};
    auto result=executeVbaProperty(nativePropertyFixture,"Stored",VbaPropertyAccess::Let,host,arguments);assert(result.status==VbaExecutionStatus::Completed && arguments[0].number==7);
    arguments.clear();result=executeVbaProperty(nativePropertyFixture,"ReadCount",VbaPropertyAccess::Get,host,arguments);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==0);
    result=executeVbaProperty(nativePropertyFixture,"Stored",VbaPropertyAccess::Get,host,arguments);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==7);
    result=executeVbaProperty(nativePropertyFixture,"Stored",VbaPropertyAccess::Get,other,arguments);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==0);
    assert(!nativeVbaPropertySignature(nativePropertyFixture,"Secret",VbaPropertyAccess::Get,true));assert(nativeVbaPropertySignature(nativePropertyFixture,"Secret",VbaPropertyAccess::Get,false));
    result=executeVbaProcedure(nativePropertyFixture,"LocalProbe",host);if(result.status!=VbaExecutionStatus::Completed)std::cerr<<"Property probe: "<<result.explanation<<" line "<<result.line<<" error "<<result.errorNumber<<"\n";assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==24);
    arguments.clear();result=executeVbaProperty(nativePropertyFixture,"ObjectValue",VbaPropertyAccess::Get,host,arguments);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->object);
    const auto object=result.typedReturnValue->object;assert(object->member("Item",{VbaValue::string("entry")}).number==7);
    result=executeVbaProperty(nativePropertyFixture,"ObjectValue",VbaPropertyAccess::Get,host,arguments);assert(result.typedReturnValue->object==object);
    result=executeVbaProperty(nativePropertyFixture,"Failure",VbaPropertyAccess::Get,host,arguments);assert(result.status==VbaExecutionStatus::Error && result.errorNumber==11);
    result=executeVbaProcedure(nativePropertyFixture+"\nFunction ReadOnlyProbe() As Long\nOn Error Resume Next\nReadCount=4\nReadOnlyProbe=Err.Number\nEnd Function","ReadOnlyProbe",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==383);
    const std::string early="Public Property Get Early() As Long\nEarly=3\nIf True Then Exit Property\nEarly=99\nEnd Property\n";
    result=executeVbaProperty(early,"Early",VbaPropertyAccess::Get,host,arguments);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==3);
    bool rejected=false;try{(void)nativeVbaPropertySignature("Property Get Bad(ByRef index As Long) As Long\nBad=index\nEnd Property","Bad",VbaPropertyAccess::Get);}catch(const VbaRuntimeError& e){rejected=e.number==3251;}assert(rejected);
    rejected=false;try{(void)nativeVbaPropertySignature(early+early,"Early",VbaPropertyAccess::Get);}catch(const VbaRuntimeError& e){rejected=e.number==5;}assert(rejected && host.effects==0);

    result=executeVbaProcedure(nativeParameterlessFixture,"Answer",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==1);
    result=executeVbaProcedure(nativeParameterlessFixture,"CallBare",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==2);
    arguments.clear();result=executeVbaProperty(nativeParameterlessFixture,"Value",VbaPropertyAccess::Get,host,arguments);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==2);
    result=executeVbaProperty(nativeParameterlessFixture,"EmptyValue",VbaPropertyAccess::Get,host,arguments);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->kind==VbaValue::Kind::Empty);
    for(int call=1;call<=2;++call){for(const auto* name:{"Counter","TailCounter"}) {result=executeVbaProcedure(nativeParameterlessFixture,name,host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==call);}}
    const auto bareInventory=analyzeVba(nativeParameterlessFixture);bool valueListed=false,tailListed=false;for(const auto& declaration:bareInventory.procedures){valueListed|=vbaProcedureExecutionName(declaration)=="property get Value";tailListed|=declaration.name=="TailCounter";}assert(valueListed && tailListed);
    result=executeVbaProcedure("Function Bad value As Long\nMsgBox \"unreached\"\nEnd Function","Bad",host);assert(result.status==VbaExecutionStatus::Unsupported && host.effects==0);
    result=executeVbaProcedure("Sub Bad As Long\nMsgBox \"unreached\"\nEnd Sub","Bad",host);assert(result.status==VbaExecutionStatus::Unsupported && host.effects==0);
    const auto invalidSetter=validateVbaProcedure("Property Let Broken\nMsgBox \"unreached\"\nEnd Property","property let Broken");assert(invalidSetter.status==VbaExecutionStatus::Unsupported && invalidSetter.errorNumber==3251);

    PropertyHost defaultHost;arguments.clear();result=executeVbaProcedure(nativeDefaultPropertyFixture,"DefaultProbe",defaultHost);if(result.status!=VbaExecutionStatus::Completed)std::cerr<<"Default property: "<<result.explanation<<" line "<<result.line<<" error "<<result.errorNumber<<"\n";assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==26);
    result=executeVbaProperty(nativeDefaultPropertyFixture,"ReadCount",VbaPropertyAccess::Get,defaultHost,arguments);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==4);
    result=executeVbaProperty(nativeDefaultPropertyFixture,"FunctionCalls",VbaPropertyAccess::Get,defaultHost,arguments);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==1);
    result=executeVbaProcedure(nativeDefaultPropertyFixture,"RejectScalar",defaultHost);assert(result.status==VbaExecutionStatus::Error && result.errorNumber==450);
    result=executeVbaProperty(nativeDefaultPropertyFixture,"ReadCount",VbaPropertyAccess::Get,defaultHost,arguments);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==4);
    result=executeVbaProcedure(nativeDefaultPropertyFixture,"MissingProbe",defaultHost);assert(result.status==VbaExecutionStatus::Error && result.errorNumber==91);
    result=executeVbaProcedure(nativeDefaultPropertyFixture,"FailedProbe",defaultHost);assert(result.status==VbaExecutionStatus::Error && result.errorNumber==11);

    PropertyHost byrefHost;
    for(const auto& probe:std::vector<std::pair<std::string,int>>{{"LetProbe",277},{"GetProbe",22},{"TemporaryProbe",21},{"ArrayProbe",330},{"ObjectProbe",319},{"ErrorProbe",211},{"MismatchProbe",13}}){result=executeVbaProcedure(nativeByRefPropertyFixture,probe.first,byrefHost);if(result.status!=VbaExecutionStatus::Completed)std::cerr<<"ByRef property "<<probe.first<<": "<<result.explanation<<" line "<<result.line<<" error "<<result.errorNumber<<"\n";assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==probe.second);}
    arguments={VbaValue::longInteger(1)};result=executeVbaProperty(nativeByRefPropertyFixture,"Counter",VbaPropertyAccess::Get,byrefHost,arguments,false);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==20 && arguments[0].number==2);
    bool facadeRejected=false;try{(void)nativeVbaPropertySignature(nativeByRefPropertyFixture,"Counter",VbaPropertyAccess::Get,true);}catch(const VbaRuntimeError& e){facadeRejected=e.number==3251;}assert(facadeRejected);
}
