#include "vba_debugger.hpp"
#include "vba_runtime.hpp"
#include <algorithm>
#include <sstream>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
namespace {
struct ObservedObject final:linaccess::VbaObject {
    int reads{};std::string_view className() const override {return "ObservedObject";}
    bool hasScalarDefault() const override {return true;}
    linaccess::VbaValue scalarValue() override {++reads;return linaccess::VbaValue::integer(99);}
    linaccess::VbaValue member(std::string_view,const std::vector<linaccess::VbaValue>&) override {++reads;throw linaccess::VbaRuntimeError(438,"Debugger must not call object getters");}
};
struct Host final:linaccess::VbaHost {
    std::vector<std::string> messages;std::shared_ptr<ObservedObject> object=std::make_shared<ObservedObject>();
    std::optional<std::string> field(std::string_view) override {return {};}
    void setField(std::string_view,std::optional<std::string>) override {assert(false);}
    void saveRecord() override {assert(false);}
    void message(std::string_view value) override {messages.emplace_back(value);}
    std::optional<linaccess::VbaValue> resolveIdentifier(std::string_view name,bool) override {if(name=="Observed" || name=="observed") return linaccess::VbaValue::reference(object);return {};}
};
int lineOf(const std::string& code,const std::string& statement) {std::istringstream input(code);std::string line;int number=0;while(std::getline(input,line)) {++number;if(line==statement) return number;}assert(false);return 0;}
const linaccess::VbaDebugVariable& variable(const std::vector<linaccess::VbaDebugVariable>& values,const char* name) {const auto found=std::find_if(values.begin(),values.end(),[&](const auto& value) {return value.name==name;});assert(found!=values.end());return *found;}
}
void testVbaDebugger() {
    using namespace linaccess;
    const std::string code=R"VBA(Sub Root()
Dim n As Long
Dim a(1 To 3) As Long
Dim big(0 To 299) As Long
Dim missing() As Long
Dim nested As Variant
Dim s As String
Dim obj As Object
n=1
s=String(600,"x")
a(1)=7
big(299)=9
nested=Array(Array("inner"))
Set obj=Observed
Child n
n=n+10
MsgBox CStr(n)
End Sub
Sub Child(ByRef item As Long)
item=item+2
End Sub)VBA";
    Host host;int pauses=0;VbaDebuggerSession* active=nullptr;
    {VbaDebuggerSession debugger([&](const VbaDebugPause& pause) {assert(active && active->paused() && vbaDebuggerBusy());++pauses;const auto& frame=pause.frames.back();const auto values=active->variables(frame.id);
        if(pauses==1) {assert(pause.reason==VbaDebugReason::Step && frame.line==lineOf(code,"n=1") && variable(values,"n").value=="0");return VbaDebugAction::StepOver;}
        if(pauses==2) {assert(frame.line==lineOf(code,"s=String(600,\"x\")") && variable(values,"n").value=="1");return VbaDebugAction::Continue;}
        if(pauses==3) {assert(pause.reason==VbaDebugReason::Breakpoint && frame.line==lineOf(code,"Child n") && pause.frames.size()==1);assert(variable(values,"s").truncated && active->text(frame.id,"s")==std::string(600,'x'));assert((variable(values,"a").bounds==std::vector<std::pair<int,int>>({{1,3}})));assert(active->elements(frame.id,"a")[0].value=="7");const auto tail=active->elements(frame.id,"big",{},256);assert(tail.size()==44 && tail.back().index==299 && tail.back().value=="9");assert(variable(values,"missing").array && variable(values,"missing").elements==0);assert(active->elements(frame.id,"nested")[0].array && active->text(frame.id,"nested",{0,0})=="inner");assert(variable(values,"obj").value=="<ObservedObject>" && host.object->reads==0);const auto rejected=executeVbaProcedure("Sub Reenter()\nMsgBox \"wrong\"\nEnd Sub","Reenter",host);assert(rejected.status==VbaExecutionStatus::Error && rejected.errorNumber==3251 && host.messages.empty());return VbaDebugAction::StepInto;}
        if(pauses==4) {assert(pause.frames.size()==2 && frame.procedure=="Child" && frame.line==lineOf(code,"item=item+2") && variable(values,"item").value=="1");return VbaDebugAction::StepOut;}
        assert(pauses==5 && pause.frames.size()==1 && frame.line==lineOf(code,"n=n+10") && variable(values,"n").value=="3");return VbaDebugAction::Continue;
    },[](auto) {return "ActualModule";},{},true);active=&debugger;debugger.setBreakpoint(code,lineOf(code,"Child n"));const auto outcome=executeVbaProcedure(code,"Root",host);assert(outcome.status==VbaExecutionStatus::Completed && pauses==5 && host.messages==std::vector<std::string>({"13"}) && !debugger.aborted() && host.object->reads==0);bool refused=false;try {(void)debugger.variables(1);}catch(const VbaRuntimeError& error) {refused=error.number==3251;}assert(refused);}
    host.messages.clear();pauses=0;{VbaDebuggerSession debugger([&](const VbaDebugPause& pause) {++pauses;assert(pause.frames.size()==1);return pauses==1 ? VbaDebugAction::StepOver : VbaDebugAction::Continue;});debugger.setBreakpoint(code,lineOf(code,"Child n"));assert(executeVbaProcedure(code,"Root",host).status==VbaExecutionStatus::Completed && pauses==2 && host.messages==std::vector<std::string>({"13"}));}
    const std::string branch="Sub Branch()\nIf False Then\nMsgBox \"wrong\"\nEnd If\nMsgBox \"right\"\nEnd Sub";host.messages.clear();{VbaDebuggerSession debugger([](const auto&) {assert(false);return VbaDebugAction::Stop;});debugger.setBreakpoint(branch,3);assert(executeVbaProcedure(branch,"Branch",host).status==VbaExecutionStatus::Completed && host.messages==std::vector<std::string>({"right"}));}
    const std::string colon="Sub Colon()\nDim n As Long: n=1\nMsgBox CStr(n)\nEnd Sub";pauses=0;{VbaDebuggerSession debugger([&](const VbaDebugPause& point) {++pauses;assert(point.frames.back().line==2);return VbaDebugAction::Continue;});debugger.setBreakpoint(colon,2);assert(executeVbaProcedure(colon,"Colon",host).status==VbaExecutionStatus::Completed && pauses==1);}
    const std::string recover="Sub Recover()\nDim n As Long\nn=7\nn=1/0\nMsgBox CStr(n)\nEnd Sub";pauses=0;host.messages.clear();{VbaDebuggerSession debugger([&](const VbaDebugPause& pause) {++pauses;if(pauses==1) {assert(pause.reason==VbaDebugReason::Error && pause.error && pause.error->number==11 && pause.recoverable && pause.frames.back().line==4);return VbaDebugAction::SkipError;}assert(pauses==2 && pause.reason==VbaDebugReason::Step && pause.frames.back().errorState.number==11);return VbaDebugAction::Continue;});assert(executeVbaProcedure(recover,"Recover",host).status==VbaExecutionStatus::Completed && pauses==2 && host.messages==std::vector<std::string>({"7"}));}
    const std::string errors="Sub Caller()\nFailing\nMsgBox \"wrong\"\nEnd Sub\nSub Failing()\nDim n As Long\nn=1/0\nEnd Sub";pauses=0;{VbaDebuggerSession debugger([&](const VbaDebugPause& pause) {++pauses;assert(pause.frames.size()==2 && pause.frames.back().procedure=="Failing" && pause.frames.back().line==7);return VbaDebugAction::Continue;});const auto failure=executeVbaProcedure(errors,"Caller",host);assert(failure.status==VbaExecutionStatus::Error && failure.errorNumber==11 && pauses==1);}
    const std::string handled="Sub Caller()\nOn Error Resume Next\nFailing\nMsgBox CStr(Err.Number)\nEnd Sub\nSub Failing()\nDim n As Long\nn=1/0\nEnd Sub";host.messages.clear();{VbaDebuggerSession debugger([](const auto&) {assert(false);return VbaDebugAction::Stop;});assert(executeVbaProcedure(handled,"Caller",host).status==VbaExecutionStatus::Completed && host.messages==std::vector<std::string>({"11"}));}
    host.messages.clear();active=nullptr;int polls=0;{VbaDebuggerSession debugger([](const VbaDebugPause& pause) {assert(pause.reason==VbaDebugReason::Step);return VbaDebugAction::Stop;},{},[&] {if(++polls==1) active->requestPause();});active=&debugger;const auto outcome=executeVbaProcedure("Sub LongRun()\nDim n As Long\nFor n=1 To 30\nNext\nMsgBox \"wrong\"\nEnd Sub","LongRun",host);assert(outcome.status==VbaExecutionStatus::Completed && outcome.exitSub && debugger.aborted() && host.messages.empty() && polls>=1);}
    {VbaDebuggerSession debugger([](const auto&) -> VbaDebugAction {throw std::runtime_error("debug callback failure");},{},{},true);assert(executeVbaProcedure(recover,"Recover",host).status==VbaExecutionStatus::Completed && debugger.aborted() && debugger.failure()=="debug callback failure");}
    assert(!vbaDebuggerBusy() && !vbaDebuggerActive());assert(executeVbaProcedure("Sub Fresh()\nMsgBox \"fresh\"\nEnd Sub","Fresh",host).status==VbaExecutionStatus::Completed && host.messages==std::vector<std::string>({"fresh"}));
}
