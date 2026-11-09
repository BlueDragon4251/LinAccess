#include "vba_expression.hpp"
#include "vba_analyzer.hpp"
#include "vba_property_fixture.hpp"
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
}
