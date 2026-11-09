#pragma once
#include "vba_runtime.hpp"
#include "vba_expression.hpp"
#include <QPointer>
#include <QSqlDatabase>
#include <functional>
class QWidget;
namespace linaccess {
// A real project/database execution host without a fabricated Form/Me.
class NativeStandardModuleHost final : public VbaHost {
public:
    NativeStandardModuleHost(QWidget* owner,QSqlDatabase db,std::function<void()> guard);
    ~NativeStandardModuleHost() override;
    void check() const;
    void enterModule(std::string_view) override;
    void leaveModule() noexcept override;
    std::optional<std::string> field(std::string_view) override;
    void setField(std::string_view,std::optional<std::string>) override;
    void saveRecord() override;
    void maximizeForm() override;
    int compareStrings(std::string_view,std::string_view,VbaCompareMode) override;
    bool isNumericText(std::string_view) override;
    VbaValue formatValue(const VbaValue&,std::string_view,bool) override;
    VbaValue formatNumericValue(const VbaValue&,bool,const VbaNumericFormatOptions&) override;
    void message(std::string_view) override;
    int messageBox(std::string_view,int,const std::optional<std::string>&) override;
    std::string inputBox(const VbaInputBoxArguments&) override;
    std::shared_ptr<VbaObject> formsCollection() override;
    std::shared_ptr<VbaObject> reportsCollection() override;
    std::shared_ptr<VbaObject> currentDatabase() override;
    std::shared_ptr<VbaObject> databaseEngine() override;
    std::shared_ptr<VbaObject> applicationObject() override;
    VbaValue domainAggregate(std::string_view,const std::vector<VbaValue>&) override;
    void executeSql(std::string_view) override;
    std::shared_ptr<VbaModuleState> moduleState(std::string_view) override;
    std::vector<std::shared_ptr<VbaModuleState>> globalModuleStates() override;
    std::optional<std::string> moduleForProcedure(std::string_view) override;
private:
    QPointer<QWidget> owner_;QSqlDatabase db_;std::function<void()> guard_;
};
}
