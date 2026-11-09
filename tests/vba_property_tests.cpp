#include "vba_array_return_fixture.hpp"
#include "vba_array.hpp"
#include "vba_expression.hpp"
#include "vba_analyzer.hpp"
#include "vba_property_fixture.hpp"
#include "vba_parameterless_fixture.hpp"
#include "vba_default_property_fixture.hpp"
#include "vba_byref_property_fixture.hpp"
#include "vba_optional_property_fixture.hpp"
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

    PropertyHost optionalHost;
    for(const auto& probe:std::vector<std::pair<std::string,int>>{{"DefaultsProbe",487},{"HoleProbe",543},{"EmptyCallProbe",457},{"OwnedProbe",3067},{"ManyProbe",673},{"EmptyManyProbe",7},{"OptionalObjectProbe",439},{"ManyObjectProbe",249},{"GuardedValuesProbe",10251}}){result=executeVbaProcedure(nativeOptionalPropertyFixture,probe.first,optionalHost);if(result.status!=VbaExecutionStatus::Completed)std::cerr<<"Optional property "<<probe.first<<": "<<result.explanation<<" line "<<result.line<<" error "<<result.errorNumber<<"\n";assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==probe.second);}
    arguments={VbaValue::longInteger(7)};result=executeVbaProperty(nativeOptionalPropertyFixture,"Cell",VbaPropertyAccess::Let,optionalHost,arguments);assert(result.status==VbaExecutionStatus::Completed && arguments.size()==1 && arguments[0].number==7);
    arguments={VbaValue::longInteger(1),VbaValue::longInteger(7)};result=executeVbaProperty(nativeOptionalPropertyFixture,"Owned",VbaPropertyAccess::Let,optionalHost,arguments,false);assert(result.status==VbaExecutionStatus::Completed && arguments.size()==2 && arguments[0].number==2 && arguments[1].number==7);
    arguments.clear();result=executeVbaProperty(nativeOptionalPropertyFixture,"Written",VbaPropertyAccess::Get,optionalHost,arguments);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==2707);
    result=executeVbaProcedure(R"VBA(Sub ArrayHole
Dim values(1 To 2,1 To 2) As Long
MsgBox "unreached"
values(,1)=7
End Sub)VBA","ArrayHole",optionalHost);assert(result.status==VbaExecutionStatus::Unsupported && result.errorNumber==3251 && optionalHost.effects==0);
    const auto mixed=validateVbaProcedure("Property Get Bad(Optional ByVal first As Long,ParamArray rest() As Variant) As Long\nEnd Property","property get Bad");assert(mixed.status==VbaExecutionStatus::Unsupported);
    PropertyHost arrayHost;
    const auto arrayRun=[&](const char* name){auto value=executeVbaProcedure(nativeArrayReturnFixture,name,arrayHost);if(value.status!=VbaExecutionStatus::Completed)std::cerr<<"Array return "<<name<<": "<<value.explanation<<" line "<<value.line<<" error "<<value.errorNumber<<"\n";assert(value.status==VbaExecutionStatus::Completed && value.typedReturnValue);return *value.typedReturnValue;};
    for(const auto& probe:std::vector<std::pair<const char*,int>>{{"LinArrayReadProbe",297},{"LinArrayCopyProbe",1299},{"LinArrayQueueProbe",14},{ "LinArrayEachProbe",3901},{"LinArrayEachObjects",19},{"LinArrayEachCopy",3999},{"LinArrayEachNested",9}}) assert(arrayRun(probe.first).number==probe.second);
    assert(arrayRun("LinArrayEachMatrix").text=="21:31:20:30:");
    const auto longs=arrayRun("LinArrayReturnLongs");assert(isVbaArray(longs) && vbaArrayElementType(longs)=="long" && (vbaArrayBounds(longs)==std::vector<VbaArrayDimension>{{2,4}}) && vbaArrayElement(longs,{VbaValue::longInteger(4)}).number==14);
    const auto bytes=arrayRun("LinArrayBytes");assert(isVbaArray(bytes) && vbaArrayElementType(bytes)=="byte" && vbaArrayElement(bytes,{VbaValue::longInteger(2)}).number==255);
    const auto strings=arrayRun("LinArrayStrings");assert(vbaArrayElementType(strings)=="string" && vbaArrayElement(strings,{VbaValue::longInteger(2)}).text=="ä");
    const auto objects=arrayRun("LinArrayObjects");const auto child=vbaArrayElement(objects,{VbaValue::longInteger(1)});assert(vbaArrayElementType(objects)=="object" && child.object && child.object->member("Item",{VbaValue::string("value")}).number==19);
    arguments.clear();result=executeVbaProperty(nativeArrayReturnFixture,"LinArrayReturnMatrix",VbaPropertyAccess::Get,arrayHost,arguments);assert(result.status==VbaExecutionStatus::Completed && (vbaArrayBounds(*result.typedReturnValue)==std::vector<VbaArrayDimension>{{2,3},{-1,0}}));
    const auto beforeBounds=arrayRun("LinArrayCallCount").number;result=executeVbaProcedure(nativeArrayReturnFixture,"LinArrayBoundsProbe",arrayHost);assert(result.status==VbaExecutionStatus::Error && result.errorNumber==9 && arrayRun("LinArrayCallCount").number==beforeBounds+1);
    const auto unallocated=arrayRun("LinArrayUnallocated");assert(isVbaArray(unallocated) && vbaArrayElementType(unallocated)=="long" && vbaArraySize(unallocated)==0);
    result=executeVbaProcedure(nativeArrayReturnFixture,"LinArrayWrong",arrayHost);assert(result.status==VbaExecutionStatus::Error && result.errorNumber==13);
    result=executeVbaProcedure("Function Bad() As Scripting.Dictionary()\nMsgBox \"unreached\"\nEnd Function","Bad",arrayHost);assert(result.status==VbaExecutionStatus::Unsupported && arrayHost.effects==0);

}
