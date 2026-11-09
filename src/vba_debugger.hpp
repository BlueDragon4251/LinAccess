#pragma once
#include "vba_expression.hpp"
#include <cstdint>
#include <functional>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>
namespace linaccess {
enum class VbaDebugAction { Continue, StepInto, StepOver, StepOut, Stop, SkipError };
enum class VbaDebugReason { Breakpoint, Step, Error, Stop, Assertion };
struct VbaDebugError {int number{};std::string description,source,helpFile;int helpContext{};};
struct VbaDebugVariable {
    std::string name,declaredType,valueType,value;
    bool array{},truncated{};std::size_t elements{},textBytes{},index{};
    std::vector<std::pair<int,int>> bounds;
};
struct VbaDebugFrameInfo {
    std::uint64_t id{};std::shared_ptr<const std::string> source;
    std::string moduleName,procedure;int line{};std::size_t instruction{},variables{};VbaDebugError errorState;
};
struct VbaDebugPause {
    VbaDebugReason reason{VbaDebugReason::Step};std::vector<VbaDebugFrameInfo> frames;
    std::optional<VbaDebugError> error;bool recoverable{};
};
struct VbaDebuggerState;
class VbaDebuggerSession {
public:
    using Handler=std::function<VbaDebugAction(const VbaDebugPause&)>;
    using SourceName=std::function<std::string(std::string_view)>;
    VbaDebuggerSession(Handler handler,SourceName names={},std::function<void()> poll={},bool stepFirst=false);
    ~VbaDebuggerSession();
    VbaDebuggerSession(const VbaDebuggerSession&)=delete;
    VbaDebuggerSession& operator=(const VbaDebuggerSession&)=delete;
    void setBreakpoint(std::string_view source,int line,bool enabled=true);
    void requestPause();void stop();
    bool paused() const;bool aborted() const;std::string failure() const;
    std::vector<VbaDebugVariable> variables(std::uint64_t frame,std::size_t offset=0,std::size_t count=128) const;
    std::vector<VbaDebugVariable> elements(std::uint64_t frame,std::string_view name,const std::vector<std::size_t>& path={},std::size_t offset=0,std::size_t count=128) const;
    std::string text(std::uint64_t frame,std::string_view name,const std::vector<std::size_t>& path={}) const;
private:
    std::shared_ptr<VbaDebuggerState> state_;
};
// Runtime-only live frame registration. Inspection never evaluates defaults or methods.
class VbaDebugFrameScope {
public:
    VbaDebugFrameScope(std::string_view source,std::string_view procedure,const VbaVariables& variables,
        std::function<std::string(std::string_view)> types,std::function<bool()> handlesError,std::function<VbaDebugError()> errorState={});
    ~VbaDebugFrameScope();
    VbaDebugFrameScope(const VbaDebugFrameScope&)=delete;
    VbaDebugFrameScope& operator=(const VbaDebugFrameScope&)=delete;
    void checkpoint(int line,std::size_t instruction,bool executable);
    void halt(VbaDebugReason reason);
    bool unhandled(const VbaDebugError& error,bool recoverable=true,int line=0);
private:
    std::shared_ptr<VbaDebuggerState> state_;std::uint64_t id_{};
};
bool vbaDebuggerBusy() noexcept;
bool vbaDebuggerActive() noexcept;
std::vector<int> vbaDebugExecutableLines(std::string_view source,std::string_view procedure);
}
