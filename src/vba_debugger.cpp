#include "vba_project_types.hpp"
#include "vba_debugger.hpp"
#include "vba_runtime.hpp"
#include "vba_array.hpp"
#include <algorithm>
#include <map>
#include <set>
#include <limits>
namespace linaccess {
namespace {
thread_local std::weak_ptr<VbaDebuggerState> currentDebugger;
std::string shorten(std::string_view value,bool* shortened) {
    auto end=std::min<std::size_t>(value.size(),512);while(end<value.size() && end && (static_cast<unsigned char>(value[end])&0xc0)==0x80) --end;
    *shortened=end<value.size();std::string result;for(char ch:value.substr(0,end)) {if(ch=='\n') result+="\\n";else if(ch=='\r') result+="\\r";else if(ch=='\t') result+="\\t";else if(ch=='\0') result+="\\0";else result+=ch;}if(*shortened) result+="…";return result;
}
VbaDebugVariable describe(const std::string& name,const std::string& declared,const VbaValue& value) {
    VbaDebugVariable item;item.name=name;item.declaredType=isProjectClassType(declared) ? std::string(declared.substr(projectClassPrefix.size())) : declared;
    try {
        switch(value.kind) {
        case VbaValue::Kind::Null:item.valueType="Null";item.value="Null";break;
        case VbaValue::Kind::Empty:item.valueType="Empty";item.value="Empty";break;
        case VbaValue::Kind::Missing:item.valueType="Missing";item.value="<Missing>";break;
        case VbaValue::Kind::Text:item.valueType="String";item.textBytes=value.text.size();item.value='"'+shorten(value.text,&item.truncated)+'"';break;
        case VbaValue::Kind::Boolean:item.valueType="Boolean";item.value=value.boolean ? "True" : "False";break;
        case VbaValue::Kind::Number:item.valueType=value.numericType==VbaValue::NumericType::Integer ? "Integer" : value.numericType==VbaValue::NumericType::Long ? "Long" : value.numericType==VbaValue::NumericType::Single ? "Single" : value.numericType==VbaValue::NumericType::Byte ? "Byte" : "Double";item.value=value.asText();break;
        case VbaValue::Kind::Currency:item.valueType="Currency";item.value=value.asText();break;
        case VbaValue::Kind::Decimal:item.valueType="Decimal";item.value=value.asText();break;
        case VbaValue::Kind::Date:item.valueType="Date";item.value=value.asText();break;
        case VbaValue::Kind::Error:item.valueType="Error";item.value="Error "+std::to_string(value.errorNumber);break;
        case VbaValue::Kind::ByteArray:item.valueType="Byte()";item.array=true;item.elements=value.bytes.size();if(value.bytes.size()<=static_cast<std::size_t>(std::numeric_limits<int>::max())) item.bounds.push_back({0,static_cast<int>(value.bytes.size())-1});item.value="Byte("+std::to_string(item.elements)+")";break;
        case VbaValue::Kind::Array:
            item.array=true;item.valueType=vbaArrayTypeName(value);try {item.elements=vbaArraySize(value);for(const auto& bound:vbaArrayBounds(value)) item.bounds.push_back({bound.lower,bound.upper});item.value="Array("+std::to_string(item.elements)+")";}catch(const VbaRuntimeError&) {item.value="(nicht dimensioniert)";}break;
        case VbaValue::Kind::Object:item.valueType=value.object && !value.object->className().empty() ? std::string(value.object->className()) : "Object";item.value=value.object ? "<"+item.valueType+">" : "Nothing";break;
        }
    }catch(...) {item.valueType="unverfügbar";item.value="<Wert nicht darstellbar>";}
    return item;
}
}
struct VbaDebuggerState {
    struct Frame {VbaDebugFrameInfo info;const VbaVariables* variables{};std::function<std::string(std::string_view)> types;std::function<bool()> handlesError;std::function<VbaDebugError()> errorState;const std::set<int>* breakpoints{};int lastLine{};std::size_t lastInstruction=std::numeric_limits<std::size_t>::max();};
    std::vector<Frame> frames;std::map<std::string,std::set<int>> breakpoints;
    VbaDebuggerSession::Handler handler;VbaDebuggerSession::SourceName names;std::function<void()> poll;
    VbaDebugAction mode{VbaDebugAction::Continue};std::size_t stepDepth{},ticks{};std::uint64_t nextId{};
    bool active=true,paused{},polling{},aborted{},pauseRequested{},propagating{};std::string failure;
    Frame& frame(std::uint64_t id) {for(auto& item:frames) if(item.info.id==id) return item;throw VbaRuntimeError(3251,"Debugger frame is no longer active");}
    void requirePaused() const {if(!active || !paused) throw VbaRuntimeError(3251,"Debugger inspection requires a live pause");}
    void abort() {aborted=true;requestVbaExecutionStop();}
    VbaDebugAction pause(VbaDebugReason reason,const VbaDebugError* error=nullptr,bool recoverable=false) {
        VbaDebugPause snapshot;snapshot.reason=reason;snapshot.recoverable=recoverable;if(error) snapshot.error=*error;for(const auto& item:frames) {auto info=item.info;info.variables=item.variables->size();if(item.errorState) info.errorState=item.errorState();snapshot.frames.push_back(std::move(info));}
        paused=true;VbaDebugAction action=VbaDebugAction::Stop;
        try {action=handler(snapshot);}catch(const std::exception& error) {failure=error.what();}catch(...) {failure="Debugger callback failed";}
        paused=false;if(!active || !failure.empty() || action==VbaDebugAction::Stop || (action==VbaDebugAction::SkipError && (!error || !recoverable))) {if(failure.empty() && action==VbaDebugAction::SkipError) failure="Ungültige Fehlerfortsetzung";abort();throw VbaExecutionStopped{};}
        pauseRequested=false;stepDepth=frames.size();mode=action==VbaDebugAction::SkipError ? VbaDebugAction::StepInto : action;
        if(action==VbaDebugAction::StepOut && stepDepth) --stepDepth;
        if(error && action!=VbaDebugAction::SkipError) propagating=true;
        return action;
    }
    const VbaValue& variable(std::uint64_t id,std::string_view name) {requirePaused();auto& selected=frame(id);const auto found=selected.variables->find(std::string(name));if(found==selected.variables->end()) throw VbaRuntimeError(5,"Unknown debugger variable");return found->second;}
};
bool vbaDebuggerBusy() noexcept {const auto state=currentDebugger.lock();return state && state->active && (state->paused || state->polling);}
bool vbaDebuggerActive() noexcept {const auto state=currentDebugger.lock();return state && state->active;}
VbaDebuggerSession::VbaDebuggerSession(Handler handler,SourceName names,std::function<void()> poll,bool stepFirst):state_(std::make_shared<VbaDebuggerState>()) {
    if(const auto current=currentDebugger.lock();current && current->active) throw VbaRuntimeError(3251,"Another debugger session is active");if(!handler) throw VbaRuntimeError(5,"Debugger requires a pause handler");state_->handler=std::move(handler);state_->names=std::move(names);state_->poll=std::move(poll);state_->mode=stepFirst ? VbaDebugAction::StepInto : VbaDebugAction::Continue;currentDebugger=state_;
}
VbaDebuggerSession::~VbaDebuggerSession() {if(!state_->frames.empty()) state_->abort();state_->active=false;if(currentDebugger.lock()==state_) currentDebugger.reset();}
void VbaDebuggerSession::setBreakpoint(std::string_view source,int line,bool enabled) {
    if(source.empty() || source.size()>1024*1024 || line<1 || line>1000000) throw VbaRuntimeError(5,"Invalid debugger breakpoint");auto& lines=state_->breakpoints[std::string(source)];if(enabled) lines.insert(line);else lines.erase(line);for(auto& frame:state_->frames) if(*frame.info.source==source) frame.breakpoints=&lines;
}
void VbaDebuggerSession::requestPause(){state_->pauseRequested=true;}
void VbaDebuggerSession::stop(){state_->abort();}
bool VbaDebuggerSession::paused() const{return state_->paused;}
bool VbaDebuggerSession::aborted() const{return state_->aborted;}
std::string VbaDebuggerSession::failure() const{return state_->failure;}
std::vector<VbaDebugVariable> VbaDebuggerSession::variables(std::uint64_t id,std::size_t offset,std::size_t count) const {
    state_->requirePaused();auto& frame=state_->frame(id);if(count>256) throw VbaRuntimeError(5,"Debugger page exceeds 256 variables");std::vector<std::string> names;for(const auto& [name,value]:*frame.variables) names.push_back(name);std::sort(names.begin(),names.end());std::vector<VbaDebugVariable> result;for(auto index=std::min(offset,names.size());index<names.size() && result.size()<count;++index) {const auto& name=names[index];result.push_back(describe(name,frame.types ? frame.types(name) : std::string{},frame.variables->at(name)));}return result;
}
namespace {
VbaValue pathValue(const VbaValue& root,const std::vector<std::size_t>& path) {if(path.size()>32) throw VbaRuntimeError(5,"Debugger array path exceeds 32 levels");auto value=root;for(const auto index:path) {if(value.kind==VbaValue::Kind::ByteArray) {if(index>=value.bytes.size()) throw VbaRuntimeError(9,"Debugger byte index unavailable");value=VbaValue::byte(value.bytes[index]);}else {if(index>=vbaArraySize(value)) throw VbaRuntimeError(9,"Debugger array index unavailable");value=vbaArrayLinearElement(value,index);}}return value;}
}
std::vector<VbaDebugVariable> VbaDebuggerSession::elements(std::uint64_t id,std::string_view name,const std::vector<std::size_t>& path,std::size_t offset,std::size_t count) const {
    if(count>256) throw VbaRuntimeError(5,"Debugger page exceeds 256 array elements");const auto value=pathValue(state_->variable(id,name),path);const auto size=value.kind==VbaValue::Kind::ByteArray ? value.bytes.size() : vbaArraySize(value);std::vector<VbaDebugVariable> result;for(auto index=std::min(offset,size);index<size && result.size()<count;++index) {auto item=describe("#"+std::to_string(index),{},value.kind==VbaValue::Kind::ByteArray ? VbaValue::byte(value.bytes[index]) : vbaArrayLinearElement(value,index));item.index=index;result.push_back(std::move(item));}return result;
}
std::string VbaDebuggerSession::text(std::uint64_t id,std::string_view name,const std::vector<std::size_t>& path) const {const auto value=pathValue(state_->variable(id,name),path);if(value.kind!=VbaValue::Kind::Text) throw VbaRuntimeError(13,"Debugger text inspection requires a String");if(value.text.size()>16*1024*1024) throw VbaRuntimeError(3035,"Debugger text exceeds 16 MiB");return value.text;}
VbaDebugFrameScope::VbaDebugFrameScope(std::string_view source,std::string_view procedure,const VbaVariables& variables,std::function<std::string(std::string_view)> types,std::function<bool()> handlesError,std::function<VbaDebugError()> errorState) {
    state_=currentDebugger.lock();if(!state_ || !state_->active) {state_.reset();return;}if(state_->paused || state_->polling) throw VbaRuntimeError(3251,"Reentrant execution while debugger is paused is unavailable");VbaDebuggerState::Frame frame;frame.info.id=id_=++state_->nextId;frame.info.procedure=procedure;frame.variables=&variables;frame.types=std::move(types);frame.handlesError=std::move(handlesError);frame.errorState=std::move(errorState);for(const auto& active:state_->frames) if(*active.info.source==source) {frame.info.source=active.info.source;break;}if(!frame.info.source) frame.info.source=std::make_shared<const std::string>(source);frame.info.moduleName=state_->names ? state_->names(source) : "Modul";if(const auto found=state_->breakpoints.find(std::string(source));found!=state_->breakpoints.end()) frame.breakpoints=&found->second;state_->frames.push_back(std::move(frame));
}
VbaDebugFrameScope::~VbaDebugFrameScope(){if(state_) {const auto found=std::find_if(state_->frames.begin(),state_->frames.end(),[&](const auto& frame) {return frame.info.id==id_;});if(found!=state_->frames.end()) state_->frames.erase(found);}}
void VbaDebugFrameScope::checkpoint(int line,std::size_t instruction,bool executable) {
    if(!state_ || !state_->active) return;if(state_->aborted) throw VbaExecutionStopped{};auto& frame=state_->frame(id_);frame.info.line=line;frame.info.instruction=instruction;
    if(state_->poll && ++state_->ticks%16==0) {state_->polling=true;try {state_->poll();}catch(const std::exception& error) {state_->failure=error.what();state_->abort();}catch(...) {state_->failure="Debugger event polling failed";state_->abort();}state_->polling=false;if(state_->aborted || vbaExecutionStopRequested()) throw VbaExecutionStopped{};}
    if(!executable) return;const bool duplicate=line==frame.lastLine && instruction>frame.lastInstruction;frame.lastLine=line;frame.lastInstruction=instruction;state_->propagating=false;
    const bool breakpoint=frame.breakpoints && frame.breakpoints->contains(line) && !duplicate;
    const bool step=state_->pauseRequested || state_->mode==VbaDebugAction::StepInto || ((state_->mode==VbaDebugAction::StepOver || state_->mode==VbaDebugAction::StepOut) && state_->frames.size()<=state_->stepDepth);
    if(breakpoint || step) (void)state_->pause(breakpoint ? VbaDebugReason::Breakpoint : VbaDebugReason::Step);
}
void VbaDebugFrameScope::halt(VbaDebugReason reason) {
    if(reason!=VbaDebugReason::Stop && reason!=VbaDebugReason::Assertion) throw VbaRuntimeError(5,"Invalid explicit debugger halt");
    if(!state_ || !state_->active) throw VbaRuntimeError(3251,"Explicit halt requires a native debugger session");
    if(state_->aborted) throw VbaExecutionStopped{};
    (void)state_->pause(reason);
}
bool VbaDebugFrameScope::unhandled(const VbaDebugError& error,bool recoverable,int line) {
    if(!state_ || !state_->active || state_->propagating) return false;if(line>0) state_->frame(id_).info.line=line;if(state_->aborted) throw VbaExecutionStopped{};for(const auto& frame:state_->frames) {if(frame.info.id==id_) break;if(frame.handlesError && frame.handlesError()) return false;}return state_->pause(VbaDebugReason::Error,&error,recoverable)==VbaDebugAction::SkipError;
}
}
