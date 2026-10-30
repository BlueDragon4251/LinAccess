#pragma once

#include <optional>
#include <stdexcept>
#include <string>
#include <string_view>

namespace linaccess {

class VbaHost {
public:
    virtual ~VbaHost() = default;
    virtual std::optional<std::string> field(std::string_view name) = 0;
    virtual void setField(std::string_view name, std::optional<std::string> value) = 0;
    virtual void message(std::string_view text) = 0;
    virtual void saveRecord() = 0;
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
};

// Strict, explicitly limited interpreter. NOT the Microsoft VBA runtime.
// The whole procedure is syntax-validated before any host-side effect.
VbaExecutionResult executeVbaProcedure(std::string_view module,
                                        std::string_view procedure,
                                        VbaHost& host);

}  // namespace linaccess
