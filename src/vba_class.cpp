#include "vba_project_types.hpp"
#include "vba_class.hpp"
#include "vba_analyzer.hpp"
#include "vba_conditional.hpp"
#include "vba_array.hpp"
#include "native_object_types.hpp"
#include <algorithm>
#include <cctype>
#include <regex>
#include <sstream>

namespace linaccess {
namespace {
std::string folded(std::string_view input) {std::string value(input);for(auto& ch:value) ch=static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));return value;}
void failed(const VbaExecutionResult& result) {if(result.status!=VbaExecutionStatus::Completed) throw VbaRuntimeError(result.errorNumber ? result.errorNumber : 3251,result.explanation,result.errorSource,result.errorHelpFile,result.errorHelpContext);}
}
struct NativeVbaClassDefinition {
    std::string name,source,defaultName;
    VbaEnumAliases enums;
    std::shared_ptr<VbaModuleState> initial;
    std::unordered_map<std::string,bool> visible;
    bool initialize{};
    std::unordered_set<std::string> classTypes;
};
std::shared_ptr<const NativeVbaClassDefinition> compileNativeVbaClass(std::string name,std::string source,const VbaEnumAliases* aliases) {
    static const std::regex identifier(R"(^[A-Za-z_][A-Za-z0-9_]{0,63}$)");
    if(!std::regex_match(name,identifier) || !isProjectClassIdentifier(folded(name)) || source.size()>1024*1024) throw VbaRuntimeError(3251,"Native class name/source is outside its supported bounds");
    const auto active=preprocessVbaConditional(source);std::vector<std::string> lines;std::istringstream stream(active);std::string line;
    static const std::regex version(R"(^\s*VERSION\s+[^\r\n]*\bCLASS\s*$)",std::regex::icase);
    static const std::regex attribute(R"VBA(^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"\s*$)VBA",std::regex::icase);
    static const std::regex document(R"(^\s*Attribute\s+VB_Base\b)",std::regex::icase);
    static const std::regex privateScope(R"(^\s*(?:Private|Friend)\b)",std::regex::icase);
    static const std::regex memberAttribute(R"(^\s*Attribute\s+([A-Za-z_][A-Za-z0-9_]*)\.VB_UserMemId\s*=\s*(-?[0-9]+)\s*(?:\'.*)?$)",std::regex::icase);
    std::string defaultName;
    int versions=0,names=0;
    while(std::getline(stream,line)) {
        lines.push_back(line);std::smatch match;
        versions+=std::regex_match(line,version);
        if(std::regex_match(line,match,memberAttribute)) {
            if(match[2].str()!="0") throw VbaRuntimeError(3251,"Native nondefault member IDs/enumerators are not implemented");
            const auto candidate=folded(match[1].str());if(!defaultName.empty() && defaultName!=candidate) throw VbaRuntimeError(5,"Ambiguous native class default member");defaultName=candidate;
        }
        if(std::regex_search(line,document)) throw VbaRuntimeError(3251,"Document CodeBehind is not a standalone native class");
        if(std::regex_match(line,match,attribute)) {++names;if(folded(match[1].str())!=folded(name)) throw VbaRuntimeError(5,"Native class source name does not match its catalog name");}
    }
    if(versions!=1 || names!=1) throw VbaRuntimeError(3251,"Native class requires one CLASS version and matching VB_Name attribute");
    auto definition=std::make_shared<NativeVbaClassDefinition>();definition->name=std::move(name);definition->source=std::move(source);definition->defaultName=std::move(defaultName);
    if(aliases) definition->enums=*aliases;
    definition->initial=createVbaModuleState(definition->source,aliases);
    const auto analysis=analyzeVba(active);
    for(const auto& procedure:analysis.procedures) {
        const auto key=folded(vbaProcedureExecutionName(procedure));
        if(procedure.line<1 || static_cast<std::size_t>(procedure.line)>lines.size()) throw VbaRuntimeError(3251,"Native class procedure source mapping unavailable");
        const bool visible=!std::regex_search(lines[procedure.line-1],privateScope);
        if(!definition->visible.emplace(key,visible).second) throw VbaRuntimeError(5,"Ambiguous native class member");
        failed(validateVbaProcedure(definition->source,key,aliases));
        for(const auto& type:nativeVbaProcedureClassTypes(definition->source,key,aliases)) definition->classTypes.insert(type);
        if(key=="class_terminate") throw VbaRuntimeError(3251,"Native class termination events are not implemented");
        if(key=="class_initialize") {
            const auto signature=nativeVbaProcedureSignature(definition->source,key,aliases);
            if(visible || !signature || signature->function || !signature->names.empty()) throw VbaRuntimeError(3251,"Class_Initialize requires a private parameterless Sub");
            definition->initialize=true;
        }
    }
    if(!definition->defaultName.empty()) {
        const auto& name=definition->defaultName;auto shape=nativeVbaPropertySignature(definition->source,name,VbaPropertyAccess::Get,true,aliases);
        const auto found=definition->visible.find(name);
        if(!shape && found!=definition->visible.end() && found->second) shape=nativeVbaProcedureSignature(definition->source,name,aliases);
        if(!shape || !shape->function) throw VbaRuntimeError(3251,"Default member requires a public Function or Property Get");
    }
    for(const auto& [name,type]:definition->initial->types) if(isProjectClassType(type)) definition->classTypes.insert(std::string(projectClassTypeName(type)));
    for(const auto& [name,value]:definition->initial->values)
        if(definition->initial->procedures.contains(name)) throw VbaRuntimeError(5,"Native class field/procedure names collide");
    return definition;
}
std::vector<std::string> nativeVbaClassDeclaredTypes(const NativeVbaClassDefinition& definition) {return {definition.classTypes.begin(),definition.classTypes.end()};}
namespace {
class NativeClass;
class ClassHost final:public VbaHost {
public:
    ClassHost(NativeClass& owner,std::unique_ptr<VbaHost> services):owner_(owner),services_(std::move(services)){if(!services_) throw VbaRuntimeError(3251,"Native class service context is unavailable");}
    void enterModule(std::string_view source) override {services_->enterModule(source);}
    void leaveModule() noexcept override {services_->leaveModule();}
    std::shared_ptr<VbaObject> errorObject() override {return services_->errorObject();}
    std::shared_ptr<VbaRandomState> randomState() override {return services_->randomState();}
    bool hasClassInstance() const override;
    bool hasInstanceIdentifier(std::string_view) override;
    VbaEnumAliases projectEnumAliases() override {return services_->projectEnumAliases();}
    std::shared_ptr<VbaModuleState> moduleState(std::string_view) override;
    std::shared_ptr<VbaObject> currentForm() override;
    std::optional<std::string> field(std::string_view) override;
    void setField(std::string_view,std::optional<std::string>) override;
    VbaValue fieldValue(std::string_view) override;
    void setFieldValue(std::string_view,const VbaValue&) override;
    void setControlValue(std::string_view name,const VbaValue& value) override {if(hasClassInstance()) throw VbaRuntimeError(3251,"Class bang/default writes are not implemented");services_->setControlValue(name,value);}
    void executionObserved(std::string_view source,std::string_view name,const VbaExecutionResult& result,bool event) override {services_->executionObserved(source,name,result,event);}
    std::vector<std::shared_ptr<VbaModuleState>> globalModuleStates() override {return services_->globalModuleStates();}
    std::optional<std::string> moduleForProcedure(std::string_view name) override {return services_->moduleForProcedure(name);}
    std::string projectClassName(std::string_view name) override {return services_->projectClassName(name);}
    bool knowsClassType(std::string_view name) override {return services_->knowsClassType(name);}
    bool canConstructClass(std::string_view name) override {return services_->canConstructClass(name);}
    std::shared_ptr<VbaObject> constructClass(std::string_view name) override {return services_->constructClass(name);}
    int compareStrings(std::string_view a,std::string_view b,VbaCompareMode mode) override {return services_->compareStrings(a,b,mode);}
    VbaValue calendarName(bool month,int number,bool abbreviated,int firstDay) override {return services_->calendarName(month,number,abbreviated,firstDay);}
    VbaValue formatValue(const VbaValue& value,std::string_view mask,bool string) override {return services_->formatValue(value,mask,string);}
    VbaValue formatNumericValue(const VbaValue& value,bool currency,const VbaNumericFormatOptions& options) override {return services_->formatNumericValue(value,currency,options);}
    bool isNumericText(std::string_view text) override {return services_->isNumericText(text);}
    void message(std::string_view text) override {services_->message(text);}
    int messageBox(std::string_view prompt,int buttons,const std::optional<std::string>& title) override {return services_->messageBox(prompt,buttons,title);}
    std::string inputBox(const VbaInputBoxArguments& arguments) override {return services_->inputBox(arguments);}
    void saveRecord() override {services_->saveRecord();}
    int doEvents() override {return services_->doEvents();}
    std::shared_ptr<VbaObject> applicationObject() override {return services_->applicationObject();}
    std::shared_ptr<VbaObject> formsCollection() override {return services_->formsCollection();}
    std::shared_ptr<VbaObject> reportsCollection() override {return services_->reportsCollection();}
    std::shared_ptr<VbaObject> currentDatabase() override {return services_->currentDatabase();}
    std::shared_ptr<VbaObject> databaseEngine() override {return services_->databaseEngine();}
    VbaValue domainAggregate(std::string_view name,const std::vector<VbaValue>& arguments) override {return services_->domainAggregate(name,arguments);}
    void executeSql(std::string_view sql) override {services_->executeSql(sql);}
private:NativeClass& owner_;std::unique_ptr<VbaHost> services_;
};
class NativeClass final:public VbaObject,public std::enable_shared_from_this<NativeClass> {
public:
    NativeClass(std::shared_ptr<const NativeVbaClassDefinition> definition,VbaClassContext context):definition_(std::move(definition)),context_(std::move(context)),state_(std::make_shared<VbaModuleState>(*definition_->initial)) {}
    std::string_view className() const override {return definition_->name;}
    bool isNativeType(std::string_view name) const override {return folded(name)==folded(definition_->name);}
    void check() const {if(context_.check) context_.check();}
    bool ownScope() const {return activeVbaModuleSource()==definition_->source;}
    bool declared(std::string_view name) const {const auto key=folded(name);return state_->values.contains(key) || state_->procedures.contains(key);}
    std::optional<VbaArgumentSignature> signature(std::string_view name,VbaPropertyAccess access) const {
        auto result=nativeVbaPropertySignature(definition_->source,name,access,!ownScope(),&definition_->enums);
        if(result) facade(*result);return result;
    }
    static void facade(const VbaArgumentSignature& signature) {
        if(std::any_of(signature.byReference.begin(),signature.byReference.end(),[](bool value){return value;})) throw VbaRuntimeError(3251,"Native class object calls currently require explicit ByVal parameters");
    }
    std::optional<VbaArgumentSignature> method(std::string_view name) const {
        const auto key=folded(name);const auto item=definition_->visible.find(key);
        if(key=="class_initialize" || key=="class_terminate" || item==definition_->visible.end() || (!ownScope() && !item->second)) return {};
        auto result=nativeVbaProcedureSignature(definition_->source,name,&definition_->enums);if(result) facade(*result);return result;
    }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        check();if(auto result=signature(name,VbaPropertyAccess::Get)) return result;return method(name);
    }
    std::optional<VbaArgumentSignature> assignmentSignature(std::string_view name,bool reference) const override {
        check();return signature(name,reference ? VbaPropertyAccess::Set : VbaPropertyAccess::Let);
    }
    std::optional<VbaArgumentSignature> defaultAssignmentSignature(bool reference) const override {
        check();if(definition_->defaultName.empty()) return {};return assignmentSignature(definition_->defaultName,reference);
    }
    VbaValue run(std::string_view name,std::vector<VbaValue> arguments,std::optional<VbaPropertyAccess> access={}) {
        check();ClassHost host(*this,context_.host());
        const auto result=access ? executeVbaProperty(definition_->source,name,*access,host,arguments,!ownScope()) : executeVbaProcedureWithArguments(definition_->source,name,host,arguments);
        failed(result);check();return result.typedReturnValue ? *result.typedReturnValue : VbaValue::empty();
    }
    bool fieldVisible(const std::string& key) const {return state_->values.contains(key) && (ownScope() || state_->publicVariables.contains(key));}
    VbaValue member(std::string_view name,const std::vector<VbaValue>& arguments) override {return invokeMember(name,arguments,true);}
    VbaValue invokeMember(std::string_view name,const std::vector<VbaValue>& arguments,bool needsValue) override {
        check();const auto key=folded(name);
        if(fieldVisible(key)) {if(state_->automaticObjects.contains(key)) {auto host=context_.host();(void)materializeAutomaticVbaObject(state_->values.at(key),state_->types.at(key),*host);check();}return indexed(state_->values.at(key),arguments);}
        auto shape=signature(name,VbaPropertyAccess::Get);const bool property=shape.has_value();if(!shape) shape=method(name);
        if(!shape) throw VbaRuntimeError(438,"Native class member is unavailable");
        if(needsValue && !shape->function) throw VbaRuntimeError(3251,"A class Sub cannot be used as a value");
        if(shape->function && shape->names.empty() && !arguments.empty() && (shape->resultType.empty() || shape->resultType=="variant" || shape->resultType.ends_with("()") || isDeclaredObjectType(shape->resultType)))
            return indexed(run(name,{},property ? std::optional<VbaPropertyAccess>(VbaPropertyAccess::Get) : std::nullopt),arguments);
        return run(name,arguments,property ? std::optional<VbaPropertyAccess>(VbaPropertyAccess::Get) : std::nullopt);
    }
    std::optional<VbaArgumentSignature> defaultArgumentSignature() const override {
        check();if(definition_->defaultName.empty()) return {};
        if(auto shape=signature(definition_->defaultName,VbaPropertyAccess::Get)) return shape;
        return method(definition_->defaultName);
    }
    bool hasScalarDefault() const override {const auto shape=defaultArgumentSignature();return shape && shape->required==0;}
    VbaValue scalarValue() override {return defaultMember({},false);}
    VbaValue defaultMember(const std::vector<VbaValue>& arguments,bool) override {
        check();if(definition_->defaultName.empty()) throw VbaRuntimeError(438,"Native class has no declared default member");
        return invokeMember(definition_->defaultName,arguments,true);
    }
    void assignDefaultMember(const std::vector<VbaValue>& indices,const VbaValue& value) override {
        check();if(definition_->defaultName.empty()) throw VbaRuntimeError(438,"Native class has no declared default member");
        write(definition_->defaultName,indices,value,value.kind==VbaValue::Kind::Object);
    }
    VbaValue indexed(const VbaValue& value,const std::vector<VbaValue>& arguments) {
        if(arguments.empty()) return value;
        if(isVbaArray(value)) return vbaArrayElement(value,arguments);
        if(value.kind!=VbaValue::Kind::Object) throw VbaRuntimeError(13,"Class member has no indexed value");
        if(!value.object) throw VbaRuntimeError(91,"Class object member is Nothing");
        const auto result=value.object->defaultMember(arguments,true);check();return result;
    }
    void write(std::string_view name,const std::vector<VbaValue>& indices,const VbaValue& input,bool object) {
        check();const auto key=folded(name);
        if(fieldVisible(key)) {
            if(state_->constants.contains(key)) throw VbaRuntimeError(383,"Class constant is read-only");
            auto& target=state_->values.at(key);const auto& type=state_->types.at(key);
            if(state_->automaticObjects.contains(key) && (!object || !indices.empty())) {auto host=context_.host();(void)materializeAutomaticVbaObject(target,type,*host);check();}
            if(!indices.empty()) {
                if(isVbaArray(target)) {target=withVbaArrayElement(target,indices,input);return;}
                if(target.kind!=VbaValue::Kind::Object) throw VbaRuntimeError(13,"Class member has no indexed target");
                if(!target.object) throw VbaRuntimeError(91,"Class object member is Nothing");target.object->assignDefaultMember(indices,input);check();return;
            }
            if(object) {
                if(input.kind!=VbaValue::Kind::Object || (!type.empty() && type!="variant" && !isDeclaredObjectType(type))) throw VbaRuntimeError(13,"Set requires an object or Variant field");
            } else if(isDeclaredObjectType(type)) {
                if(!target.object) throw VbaRuntimeError(91,"Class object member is Nothing");
                if(!target.object->hasScalarDefault()) throw VbaRuntimeError(3251,"Class object field has no supported scalar default");target.object->assignDefaultMember({},scalarVbaValue(input));check();return;
            }
            if(state_->fixedArrays.contains(key)) throw VbaRuntimeError(3251,"Whole fixed-array assignment is not implemented");
            target=coerceVbaModuleValue(input,type);return;
        }
        const auto access=object ? VbaPropertyAccess::Set : VbaPropertyAccess::Let;
        if(signature(name,access)) {auto arguments=indices;arguments.push_back(input);(void)run(name,std::move(arguments),access);return;}
        if(const auto getter=signature(name,VbaPropertyAccess::Get)) {
            const bool forwarded=getter->names.empty() && !indices.empty() && (getter->resultType.empty() || getter->resultType=="variant" || isDeclaredObjectType(getter->resultType));
            if(forwarded || (!object && isDeclaredObjectType(getter->resultType))) {
                const auto target=run(name,forwarded ? std::vector<VbaValue>{} : indices,VbaPropertyAccess::Get);
                if(target.kind!=VbaValue::Kind::Object) throw VbaRuntimeError(13,"Class default assignment requires an object result");
                if(!target.object) throw VbaRuntimeError(91,"Class getter result is Nothing");
                target.object->assignDefaultMember(forwarded ? indices : std::vector<VbaValue>{},input);check();return;
            }
            throw VbaRuntimeError(383,"Class property is read-only");
        }
        throw VbaRuntimeError(438,"Native class assignment member is unavailable");
    }
    void assign(std::string_view name,const VbaValue& value) override {write(name,{},value,false);}
    void assignObject(std::string_view name,const VbaValue& value) override {write(name,{},value,true);}
    void assignIndexedMember(std::string_view name,const std::vector<VbaValue>& indices,const VbaValue& value) override {write(name,indices,value,false);}
    void assignIndexedObject(std::string_view name,const std::vector<VbaValue>& indices,const VbaValue& value) override {write(name,indices,value,true);}
    VbaValue item(std::string_view) override {throw VbaRuntimeError(3251,"Native class default/bang members require unsupported attribute binding");}
    std::shared_ptr<const NativeVbaClassDefinition> definition_;VbaClassContext context_;std::shared_ptr<VbaModuleState> state_;
};
bool ClassHost::hasClassInstance() const {return owner_.ownScope();}
bool ClassHost::hasInstanceIdentifier(std::string_view name) {return owner_.ownScope() ? owner_.declared(name) : services_->hasInstanceIdentifier(name);}
std::shared_ptr<VbaModuleState> ClassHost::moduleState(std::string_view source) {return source==owner_.definition_->source ? owner_.state_ : services_->moduleState(source);}
std::shared_ptr<VbaObject> ClassHost::currentForm() {return owner_.ownScope() ? owner_.shared_from_this() : services_->currentForm();}
std::optional<std::string> ClassHost::field(std::string_view name) {return owner_.ownScope() ? owner_.member(name,{}).asField() : services_->field(name);}
void ClassHost::setField(std::string_view name,std::optional<std::string> value) {if(owner_.ownScope()) owner_.assign(name,value ? VbaValue::string(*value) : VbaValue{});else services_->setField(name,std::move(value));}
VbaValue ClassHost::fieldValue(std::string_view name) {return owner_.ownScope() ? owner_.member(name,{}) : services_->fieldValue(name);}
void ClassHost::setFieldValue(std::string_view name,const VbaValue& value) {if(owner_.ownScope()) owner_.assign(name,value);else services_->setFieldValue(name,value);}
}
std::shared_ptr<VbaObject> constructNativeVbaClass(std::shared_ptr<const NativeVbaClassDefinition> definition,VbaClassContext context) {
    if(!definition || !context.host) throw VbaRuntimeError(3251,"Native class definition/context is unavailable");
    struct Retained {std::weak_ptr<VbaObject> object;std::size_t bytes;};static thread_local std::vector<Retained> live;
    live.erase(std::remove_if(live.begin(),live.end(),[](const auto& item){return item.object.expired();}),live.end());
    std::size_t bytes=definition->source.size();for(const auto& item:live) bytes+=item.bytes;
    if(live.size()>=4096 || bytes>16*1024*1024) throw VbaRuntimeError(3035,"Native class live-instance/source budget exceeded");
    auto services=context.host();if(!services) throw VbaRuntimeError(3251,"Native class type context is unavailable");
    for(const auto& type:definition->classTypes) if(!services->knowsClassType(type)) throw VbaRuntimeError(3251,"Class member project type is unavailable");
    auto object=std::make_shared<NativeClass>(std::move(definition),std::move(context));object->check();live.push_back({object,object->definition_->source.size()});
    if(object->definition_->initialize) (void)object->run("class_initialize",{});return object;
}
}
