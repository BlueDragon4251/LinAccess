#pragma once

#include <optional>
#include <cstdint>
#include <exception>
#include <memory>
#include <stdexcept>
#include <string>
#include <string_view>
#include <vector>
#include <unordered_map>

namespace linaccess {

class VbaRuntimeError : public std::runtime_error {
public:
    VbaRuntimeError(int code, const std::string& message, const std::string& origin = {},
                    const std::string& help = {}, int context = 0)
        : std::runtime_error(message), number(code), source(origin), helpFile(help), helpContext(context) {}
    int number;
    std::string source, helpFile;
    int helpContext{};
};
// Successful native application Quit stops existing GUI-thread VBA frames;
// later independent engine executions get a fresh epoch.
class VbaExecutionStopped final : public std::exception {
public:
    const char* what() const noexcept override { return "VBA execution stopped"; }
};
class VbaExecutionScope final {
public:
    VbaExecutionScope();
    ~VbaExecutionScope();
    VbaExecutionScope(const VbaExecutionScope&)=delete;
    VbaExecutionScope& operator=(const VbaExecutionScope&)=delete;
};
std::uint64_t vbaExecutionEpoch() noexcept;
bool vbaExecutionStopRequested() noexcept;
void requestVbaExecutionStop() noexcept;
enum class VbaCompareMode { Binary, Text, Database };
class VbaObject;
struct VbaValue;
struct VbaModuleState;
std::shared_ptr<VbaModuleState> createVbaModuleState(std::string_view module);
struct VbaOpenFormArguments { std::string name,where; std::optional<std::string> openArgs; int windowMode{}; int dataMode{-1}; };
struct VbaInputBoxArguments {
    std::string prompt;
    std::optional<std::string> title;
    std::string defaultValue;
    std::optional<int> xTwips,yTwips;
};
VbaOpenFormArguments parseVbaOpenFormArguments(const std::vector<VbaValue>& arguments);

struct VbaNumericFormatOptions {
    std::optional<int> digits;
    std::optional<bool> leading,parens,group;
};
VbaNumericFormatOptions parseVbaNumericFormatOptions(const std::vector<VbaValue>& arguments);

struct VbaExecutionResult;
class VbaHost {
public:
    virtual ~VbaHost() = default;
    // Observation only: exceptions from observers must never change execution.
    virtual void executionObserved(std::string_view, std::string_view, const VbaExecutionResult&, bool) {}
    virtual std::optional<std::string> field(std::string_view name) = 0;
    virtual void setField(std::string_view name, std::optional<std::string> value) = 0;
    virtual int compareStrings(std::string_view left,std::string_view right,VbaCompareMode mode);
    virtual VbaValue formatValue(const VbaValue& value,std::string_view mask,bool stringResult);
    virtual VbaValue formatNumericValue(const VbaValue& value,bool currency,const VbaNumericFormatOptions& options);
    virtual bool isNumericText(std::string_view text);
    virtual void enterModule(std::string_view) {}
    virtual void leaveModule() noexcept {}
    virtual bool hasInstanceIdentifier(std::string_view) { return false; }
    virtual std::shared_ptr<VbaModuleState> moduleState(std::string_view module);
    virtual std::vector<std::shared_ptr<VbaModuleState>> globalModuleStates() { return {}; }
    virtual std::optional<VbaValue> resolveIdentifier(std::string_view name,bool objectReference);
    virtual bool assignIdentifier(std::string_view name,const VbaValue& value);
    virtual VbaValue expressionValue(std::string_view name,bool objectReference);
    virtual VbaValue bracketValue(std::string_view name,bool objectReference);
    virtual VbaValue domainAggregate(std::string_view function,const std::vector<VbaValue>& arguments);
    virtual VbaValue expressionAggregate(std::string_view function,std::string_view field,bool allRecords);
    virtual VbaValue expressionAggregateExpression(std::string_view function,std::string_view expression);
    virtual VbaValue fieldValue(std::string_view name);
    virtual VbaValue controlValue(std::string_view name);
    virtual void setControlValue(std::string_view name,const VbaValue& value);
    virtual void setFieldValue(std::string_view name,const VbaValue& value);
    virtual void message(std::string_view text) = 0;
    virtual int messageBox(std::string_view prompt,int buttons,const std::optional<std::string>& title);
    virtual std::string inputBox(const VbaInputBoxArguments& arguments);
    virtual VbaValue eval(std::string_view expression);
    virtual int doEvents() { throw VbaRuntimeError(3251,"DoEvents requires a native event-loop host"); }
    virtual void saveRecord() = 0;
    virtual void quitApplication() { throw std::runtime_error("DoCmd.Quit: kein Anwendungskontext"); }
    virtual void maximizeForm() { throw std::runtime_error("DoCmd.Maximize: kein Formularkontext"); }
    virtual void gotoLastRecord() { throw std::runtime_error("DoCmd.GoToRecord: keine Datensatznavigation"); }
    virtual void gotoNextRecord() { throw std::runtime_error("DoCmd.GoToRecord acNext: keine Datensatznavigation"); }
    virtual void gotoNewRecord() { throw std::runtime_error("DoCmd.GoToRecord acNewRec: keine Datensatznavigation"); }
    virtual void requery() { throw std::runtime_error("DoCmd.Requery: kein Formularkontext"); }
    virtual std::string memberControlName(std::string_view name) { return std::string(name); }
    virtual void requeryControl(std::string_view) {
        throw std::runtime_error("Me.Control.Requery: kein Formularkontext");
    }
    virtual void focusControl(std::string_view) {
        throw std::runtime_error("Me.Control.SetFocus: kein Formularkontext");
    }
    virtual void setControlVisible(std::string_view, bool) {
        throw std::runtime_error("Me.Control.Visible: kein Formularkontext");
    }
    virtual std::optional<std::string> moduleForProcedure(std::string_view) { return std::nullopt; }
    virtual std::shared_ptr<VbaObject> applicationObject() { throw VbaRuntimeError(438,"Application object unavailable"); }
    virtual std::shared_ptr<VbaObject> currentForm() {
        throw VbaRuntimeError(438,"Me: form object unavailable");
    }
    virtual std::shared_ptr<VbaObject> reportsCollection() { throw VbaRuntimeError(438,"Reports collection unavailable"); }
    virtual std::shared_ptr<VbaObject> formsCollection() {
        throw VbaRuntimeError(438,"Forms collection unavailable");
    }
    virtual std::shared_ptr<VbaObject> databaseEngine() { throw VbaRuntimeError(3251,"DBEngine requires a native DAO host"); }
    virtual std::shared_ptr<VbaObject> currentDatabase() {
        throw std::runtime_error("CurrentDb: keine verbundene Datenbank");
    }
    virtual void executeSql(std::string_view) {
        throw std::runtime_error("DAO.Execute/DoCmd.RunSQL: keine verbundene Datenbank");
    }
    virtual void sendKeys(std::string_view,bool) { throw VbaRuntimeError(3251,"SendKeys requires a native keyboard host"); }
    virtual void openFormArguments(const std::vector<VbaValue>& arguments);
    virtual void openForm(std::string_view) {
        throw std::runtime_error("DoCmd.OpenForm: kein Access-Formularkatalog");
    }
private:
    std::unordered_map<std::string,std::shared_ptr<VbaModuleState>> moduleStates_;
};

enum class VbaExecutionStatus { Completed, ProcedureMissing, Unsupported, Error };

struct VbaExecutionResult {
    VbaExecutionStatus status{VbaExecutionStatus::ProcedureMissing};
    std::string explanation;
    int line{};
    bool exitSub{};
    std::optional<std::string> returnValue;
    bool cancelled{};
    std::shared_ptr<VbaValue> typedReturnValue;
    int errorNumber{};
    std::string errorSource, errorHelpFile;
    int errorHelpContext{};
};

// Syntax inspection only; Completed here does not mean a procedure was executed.
VbaExecutionResult validateVbaProcedure(std::string_view module, std::string_view procedure);

VbaExecutionResult executeVbaProcedureWithArguments(std::string_view module,
    std::string_view procedure, VbaHost& host, std::vector<VbaValue>& arguments);

VbaExecutionResult executeVbaEventExpression(std::string_view module,std::string_view expression,VbaHost& host);

// Strict, explicitly limited interpreter. NOT the Microsoft VBA runtime.
// The whole procedure is syntax-validated before any host-side effect.
VbaExecutionResult executeVbaProcedure(std::string_view module,
                                        std::string_view procedure,
                                        VbaHost& host);

}  // namespace linaccess
