#pragma once

#include <optional>
#include <memory>
#include <stdexcept>
#include <string>
#include <string_view>
#include <vector>

namespace linaccess {

class VbaRuntimeError : public std::runtime_error {
public:
    VbaRuntimeError(int code, const std::string& message) : std::runtime_error(message), number(code) {}
    int number;
};
class VbaObject;
struct VbaValue;

class VbaHost {
public:
    virtual ~VbaHost() = default;
    virtual std::optional<std::string> field(std::string_view name) = 0;
    virtual void setField(std::string_view name, std::optional<std::string> value) = 0;
    virtual VbaValue formatValue(const VbaValue& value,std::string_view mask,bool stringResult);
    virtual VbaValue expressionValue(std::string_view name,bool objectReference);
    virtual VbaValue expressionAggregate(std::string_view function,std::string_view field,bool allRecords);
    virtual VbaValue fieldValue(std::string_view name);
    virtual VbaValue controlValue(std::string_view name);
    virtual void setControlValue(std::string_view name,const VbaValue& value);
    virtual void setFieldValue(std::string_view name,const VbaValue& value);
    virtual void message(std::string_view text) = 0;
    virtual void saveRecord() = 0;
    virtual void quitApplication() { throw std::runtime_error("DoCmd.Quit: kein Anwendungskontext"); }
    virtual void maximizeForm() { throw std::runtime_error("DoCmd.Maximize: kein Formularkontext"); }
    virtual void gotoLastRecord() { throw std::runtime_error("DoCmd.GoToRecord: keine Datensatznavigation"); }
    virtual void gotoNextRecord() { throw std::runtime_error("DoCmd.GoToRecord acNext: keine Datensatznavigation"); }
    virtual void gotoNewRecord() { throw std::runtime_error("DoCmd.GoToRecord acNewRec: keine Datensatznavigation"); }
    virtual void requery() { throw std::runtime_error("DoCmd.Requery: kein Formularkontext"); }
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
    virtual std::shared_ptr<VbaObject> currentDatabase() {
        throw std::runtime_error("CurrentDb: keine verbundene Datenbank");
    }
    virtual void executeSql(std::string_view) {
        throw std::runtime_error("DAO.Execute/DoCmd.RunSQL: keine verbundene Datenbank");
    }
    virtual void openForm(std::string_view) {
        throw std::runtime_error("DoCmd.OpenForm: kein Access-Formularkatalog");
    }
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
};

// Syntax inspection only; Completed here does not mean a procedure was executed.
VbaExecutionResult validateVbaProcedure(std::string_view module, std::string_view procedure);

VbaExecutionResult executeVbaProcedureWithArguments(std::string_view module,
    std::string_view procedure, VbaHost& host, std::vector<VbaValue>& arguments);

// Strict, explicitly limited interpreter. NOT the Microsoft VBA runtime.
// The whole procedure is syntax-validated before any host-side effect.
VbaExecutionResult executeVbaProcedure(std::string_view module,
                                        std::string_view procedure,
                                        VbaHost& host);

}  // namespace linaccess
