#include "vba_error_clear_fixture.hpp"
#include "vba_redim_declaration_fixture.hpp"
#include "vba_native_object_array_fixture.hpp"
#include "vba_enum_type_fixture.hpp"
#include "vba_automatic_class_fixture.hpp"
#include "vba_default_class_fixture.hpp"
#include "vba_typed_class_fixture.hpp"
#include "vba_class.hpp"
#include "vba_class_fixture.hpp"
#include "vba_array.hpp"
#include <cctype>
#include <iostream>
#ifdef NDEBUG
#undef NDEBUG
#endif
// Assertions must remain active in the existing Release core target.
#include <cassert>
namespace {
using namespace linaccess;
struct ClassState {std::shared_ptr<VbaModuleState> automaticState;bool alive{true};int messages{};std::function<void()> callback;std::unordered_map<std::string,std::shared_ptr<VbaModuleState>> modules;std::unordered_map<std::string,std::string> classes{{"nativecounter",nativeClassFixture}};};
class ClassServices final:public VbaHost {
public:
    explicit ClassServices(std::shared_ptr<ClassState> state):state(std::move(state)){}
    std::shared_ptr<ClassState> state;
    std::optional<std::string> field(std::string_view) override {throw VbaRuntimeError(438,"No form in class services");}
    void setField(std::string_view,std::optional<std::string>) override {assert(false);}
    void saveRecord() override {assert(false);}
    std::shared_ptr<VbaObject> currentForm() override {throw VbaRuntimeError(438,"No hidden Me in standard service modules");}
    void message(std::string_view) override {++state->messages;if(state->callback) state->callback();}
    std::shared_ptr<VbaModuleState> moduleState(std::string_view source) override {auto& value=state->modules[std::string(source)];if(!value) value=createVbaModuleState(source);return value;}
    std::vector<std::shared_ptr<VbaModuleState>> globalModuleStates() override {return state->automaticState ? std::vector<std::shared_ptr<VbaModuleState>>{state->automaticState} : std::vector<std::shared_ptr<VbaModuleState>>{};}
    std::optional<std::string> moduleForProcedure(std::string_view name) override {auto key=std::string(name);for(auto& ch:key) ch=static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));if(key=="projectvalue") return "Public Function ProjectValue As Long\nProjectValue=19\nEnd Function";if(key=="projectme") return "Public Function ProjectMe As Long\nProjectMe=VarType(Me)\nEnd Function";return {};}
    std::string projectClassName(std::string_view name) override {if(!knowsClassType(name)) throw VbaRuntimeError(3251,"Unknown technical class type");return name=="nativecounter" ? "NativeCounter" : std::string(name);}
    bool knowsClassType(std::string_view name) override {auto key=std::string(name);for(auto& ch:key) ch=static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));return state->classes.contains(key);}
    bool canConstructClass(std::string_view name) override {const auto found=state->classes.find(std::string(name));if(found==state->classes.end()) return false;(void)compileNativeVbaClass(name=="nativecounter" ? "NativeCounter" : std::string(name),found->second);return true;}
    std::shared_ptr<VbaObject> constructClass(std::string_view name) override {auto key=std::string(name);for(auto& ch:key) ch=static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));const auto found=state->classes.find(key);if(found==state->classes.end()) throw VbaRuntimeError(3251,"Unknown technical class");VbaClassContext context;context.check=[state=state]{if(!state->alive) throw VbaRuntimeError(3420,"Class technical context closed");};context.host=[state=state]{return std::make_unique<ClassServices>(state);};return constructNativeVbaClass(compileNativeVbaClass(key=="nativecounter" ? "NativeCounter" : key,found->second),std::move(context));}
};
}
void testVbaClasses() {using namespace linaccess;auto state=std::make_shared<ClassState>();ClassServices host(state);state->classes["nativedefault"]=nativeDefaultClassFixture;
    const auto run=[&](const char* name){auto result=executeVbaProcedure(nativeClassConsumer,name,host);if(result.status!=VbaExecutionStatus::Completed)std::cerr<<"Class "<<name<<": "<<result.explanation<<" line "<<result.line<<" error "<<result.errorNumber<<"\n";assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue);return *result.typedReturnValue;};
    assert(run("ClassProbe").number==911);assert(run("ArrayProbe").number==25);assert(run("MethodProbe").text=="x18:20");assert(run("IdentityProbe").boolean);assert(run("ValueGuard").number==325101);assert(run("RefGuard").number==325101);
    auto first=host.constructClass("NativeCounter"),second=host.constructClass("NativeCounter");assert(first->className()=="NativeCounter" && first->member("Calls",{}).number==1 && first->member("ReadCount",{}).number==0);
    first->assign("Value",VbaValue::longInteger(7));assert(first->member("ReadCount",{}).number==0 && first->member("Value",{}).number==7 && second->member("Value",{}).number==1);
    first->assignIndexedMember("Bag",{VbaValue::string("value")},VbaValue::longInteger(11));assert(first->member("Bag",{VbaValue::string("value")}).number==11 && second->member("Bag",{VbaValue::string("value")}).number==5);
    first->assignObject("Child",VbaValue::reference(second));assert(first->member("Child",{}).object==second && first->member("Identity",{}).object==first);
    assert(first->member("OwnedRef",{}).number==3 && first->member("ExternalValue",{}).number==19 && first->member("ExternalMe",{}).number==438);
    int error=0;try{first->member("Hidden",{VbaValue::longInteger(1)});}catch(const VbaRuntimeError& e){error=e.number;}assert(error==438);
    error=0;try{first->assign("ConstantValue",VbaValue::longInteger(9));}catch(const VbaRuntimeError& e){error=e.number;}assert(error==383 && first->member("ConstantValue",{}).number==4);
    auto result=executeVbaProcedure("Sub MissingClass\nMsgBox \"unreached\"\nDim value As Object\nSet value=New Missing\nEnd Sub","MissingClass",host);assert(result.status==VbaExecutionStatus::Error && result.errorNumber==3251 && state->messages==0);
    result=executeVbaProcedure("Function Literal As String\nLiteral=\"New Missing\"\nEnd Function","Literal",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->text=="New Missing");
    error=0;try{(void)compileNativeVbaClass("NativeCounter",nativeClassFixture+"Private Sub Class_Terminate()\nEnd Sub\n");}catch(const VbaRuntimeError& e){error=e.number;}assert(error==3251);
    state->callback=[&] {
        ClassServices other(state);VbaVariables variables{{"held",VbaValue::reference(first)}};int privateError=0;
        try{(void)evaluateVbaExpression("held.Hidden(1)",other,variables);}catch(const VbaRuntimeError& e){privateError=e.number;}assert(privateError==438);
        const std::string source="Public carried As Object\nFunction Reentrant As Long\nOn Error Resume Next\nReentrant=carried.Hidden(1)\nReentrant=Err.Number\nEnd Function";
        other.moduleState(source)->values["carried"]=VbaValue::reference(first);const auto nested=executeVbaProcedure(source,"Reentrant",other);assert(nested.status==VbaExecutionStatus::Completed && nested.typedReturnValue->number==438);other.moduleState(source)->values["carried"]=VbaValue::reference(nullptr);
    };
    const auto beforeCallback=first->member("Value",{}).number;(void)first->invokeMember("Trigger",{},false);state->callback={};assert(state->messages==1 && first->member("Value",{}).number==beforeCallback+1);
    for(const auto& probe:std::vector<std::pair<const char*,int>>{{"TypedProbe",33},{"TypedMismatch",13},{"TypedRefMismatch",13},{"ClassArrayProbe",28},{"ClassArrayMismatch",13},{"ClassArrayRefMismatch",13}}) {
        const auto result=executeVbaProcedure(typedClassConsumer,probe.first,host);
        if(result.status!=VbaExecutionStatus::Completed) std::cerr<<"Typed class "<<probe.first<<": "<<result.explanation<<" error "<<result.errorNumber<<"\n";
        assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue && result.typedReturnValue->number==probe.second);
    }
    result=executeVbaProcedure(typedClassConsumer,"TypedReturnMismatch",host);assert(result.status==VbaExecutionStatus::Error && result.errorNumber==13);
    for(const auto& declaration:std::vector<std::string>{"Dim item As Missing", ""}) {
        const auto source=declaration.empty() ? "Function UnknownTyped As Missing\nMsgBox \"unreached\"\nEnd Function" : "Sub UnknownTyped\nMsgBox \"unreached\"\n"+declaration+"\nEnd Sub";
        const auto before=state->messages;result=executeVbaProcedure(source,"UnknownTyped",host);assert(result.status==VbaExecutionStatus::Error && result.errorNumber==3251 && state->messages==before);
    }
    state->classes["selflink"]="VERSION 1.0 CLASS\nAttribute VB_Name = \"SelfLink\"\nPublic NextLink As SelfLink\nPublic Function Identity As SelfLink\nSet Identity=Me\nEnd Function\n";
    auto self=host.constructClass("SelfLink");assert(self->member("Identity",{}).object==self && !self->member("NextLink",{}).object);
    result=executeVbaProcedure(nativeDefaultClassConsumer,"DefaultClassProbe",host);if(result.status!=VbaExecutionStatus::Completed) std::cerr<<"Default class: "<<result.explanation<<" line "<<result.line<<" error "<<result.errorNumber<<"\n";assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==1058);
    auto defaultValue=host.constructClass("NativeDefault");assert(defaultValue->hasScalarDefault() && scalarVbaValue(VbaValue::reference(defaultValue)).number==4 && defaultValue->member("Reads",{}).number==1);
    defaultValue->assignDefaultMember({VbaValue::longInteger(2)},VbaValue::longInteger(17));assert(defaultValue->defaultMember({},false).number==15);
    error=0;try{(void)first->defaultMember({},true);}catch(const VbaRuntimeError& e){error=e.number;}assert(error==438);
    for(const auto& invalid:std::vector<std::string>{nativeDefaultClassFixture+"Attribute Item.VB_UserMemId = 0\n",nativeDefaultClassFixture+"Attribute Item.VB_UserMemId = -4\n"}) {
        error=0;try{(void)compileNativeVbaClass("NativeDefault",invalid);}catch(const VbaRuntimeError& e){error=e.number;}assert(error==5 || error==3251);
    }
    state->classes["autotracked"]=automaticClassFixture;state->automaticState=createVbaModuleState(automaticClassState);
    const auto automatic=[&](const char* name){const auto outcome=executeVbaProcedure(automaticClassConsumer,name,host);if(outcome.status!=VbaExecutionStatus::Completed) std::cerr<<"Automatic "<<name<<": "<<outcome.explanation<<" line "<<outcome.line<<" error "<<outcome.errorNumber<<"\n";assert(outcome.status==VbaExecutionStatus::Completed);return outcome.typedReturnValue ? *outcome.typedReturnValue : VbaValue::empty();};
    assert(automatic("AutoLocalProbe").text=="0:0:1:1:2:3");assert(automatic("AutoStaticProbe").number==2 && automatic("AutoStaticProbe").number==3);assert(automatic("AutoModuleProbe").number==5 && automatic("AutoModuleProbe").number==5);automatic("ResetAutomaticModule");assert(automatic("AutoModuleProbe").number==6);assert(automatic("AutoDefaultProbe").number==11 && automatic("AutoOwnedProbe").boolean);
    auto parent=host.constructClass("AutoTracked");auto child=parent->member("Child",{}).object;assert(child && parent->member("Child",{}).object==child);const auto childSerial=child->member("Serial",{}).number;parent->assignObject("Child",VbaValue::reference(nullptr));auto replacement=parent->member("Child",{}).object;assert(replacement!=child && replacement->member("Serial",{}).number==childSerial+1);
    const auto beforeAutomaticFailure=state->messages;result=executeVbaProcedure("Sub AutoMissing\nMsgBox \"unreached\"\nDim item As New Missing\nEnd Sub","AutoMissing",host);assert(result.status==VbaExecutionStatus::Error && result.errorNumber==3251 && state->messages==beforeAutomaticFailure);
    for(const auto& probe:std::vector<std::pair<const char*,int>>{{"EnumTypeProbe",64},{"EnumTypeOverflow",6}}) {const auto outcome=executeVbaProcedure(enumTypeFixture,probe.first,host);if(outcome.status!=VbaExecutionStatus::Completed) std::cerr<<"Enum type "<<probe.first<<": "<<outcome.explanation<<" line "<<outcome.line<<" error "<<outcome.errorNumber<<"\n";assert(outcome.status==VbaExecutionStatus::Completed && outcome.typedReturnValue->number==probe.second);}
    const auto beforeEnumGuard=state->messages;result=executeVbaProcedure(enumTypeFixture+"Sub RejectEnumClass\nMsgBox \"unreached\"\nDim item As Object\nIf TypeOf item Is NativeState_2 Then Exit Sub\nEnd Sub\n","RejectEnumClass",host);assert(result.status==VbaExecutionStatus::Error && result.errorNumber==3251 && state->messages==beforeEnumGuard);
    result=executeVbaProcedure(nativeObjectArrayFixture,"NativeObjectArrayProbe",host);if(result.status!=VbaExecutionStatus::Completed) std::cerr<<"Native object array: "<<result.explanation<<" line "<<result.line<<" error "<<result.errorNumber<<"\n";assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==25);
    for(const auto& probe:std::vector<std::pair<const char*,int>>{{"ReDimDeclarationProbe",30},{"ReDimExistingProbe",17},{"ReDimStaticProbe",7},{"ReDimStaticProbe",14}}) {const auto outcome=executeVbaProcedure(nativeReDimDeclarationFixture,probe.first,host);if(outcome.status!=VbaExecutionStatus::Completed) std::cerr<<"ReDim declaration "<<probe.first<<": "<<outcome.explanation<<" line "<<outcome.line<<" error "<<outcome.errorNumber<<"\n";assert(outcome.status==VbaExecutionStatus::Completed && outcome.typedReturnValue->number==probe.second);}
    for(const auto& probe:std::vector<std::pair<const char*,int>>{{"ErrorClearProbe",704},{"ErrorExitProbe",17}}) {const auto outcome=executeVbaProcedure(nativeErrorClearFixture,probe.first,host);if(outcome.status!=VbaExecutionStatus::Completed) std::cerr<<"Err reset "<<probe.first<<": "<<outcome.explanation<<" line "<<outcome.line<<" error "<<outcome.errorNumber<<"\n";assert(outcome.status==VbaExecutionStatus::Completed && outcome.typedReturnValue->number==probe.second);}
    state->classes["database"]="VERSION 1.0 CLASS\nAttribute VB_Name = \"Database\"\nPublic Value As Long\n";
    result=executeVbaProcedure("Function RejectNominalDao As Long\nDim items(0 To 0) As DAO.Database\nOn Error Resume Next\nSet items(0)=New Database\nRejectNominalDao=Err.Number\nEnd Function","RejectNominalDao",host);assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue->number==13);
    state->alive=false;error=0;try{first->member("Calls",{});}catch(const VbaRuntimeError& e){error=e.number;}assert(error==3420);
}
