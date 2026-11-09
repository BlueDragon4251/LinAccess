#pragma once
#include "vba_runtime.hpp"
#include "vba_expression.hpp"
#include "vba_module_context.hpp"
#include "vba_source_codec.hpp"
#include "vba_analyzer.hpp"
#include "access_format.hpp"
#include "dao_database.hpp"
#include "native_input_box.hpp"
#include <QPointer>
#include <QWidget>
#include <QSqlDatabase>
#include <QRegularExpression>
#include <functional>
int nativeModuleMessageBox(QWidget*,std::string_view,int,const std::optional<std::string>&);
namespace linaccess {
// Real database/editor context. There is deliberately no fabricated Me/Form.
class NativeStandardModuleHost final : public VbaHost {
public:
    NativeStandardModuleHost(QWidget* owner,QSqlDatabase db,std::function<void()> guard)
        : owner_(owner),db_(std::move(db)),guard_(std::move(guard)) {}
    void check() const {if(!owner_ || !db_.isOpen()) throw VbaRuntimeError(3420,"Module execution context is closed");guard_();}
    void enterModule(std::string_view) override {check();}
    std::optional<std::string> field(std::string_view) override {throw VbaRuntimeError(438,"Me requires a real form instance");}
    void setField(std::string_view,std::optional<std::string>) override {throw VbaRuntimeError(438,"Me requires a real form instance");}
    void saveRecord() override {throw VbaRuntimeError(3251,"Module editor has no current record");}
    int compareStrings(std::string_view a,std::string_view b,VbaCompareMode mode) override {
        if(mode==VbaCompareMode::Binary) return VbaHost::compareStrings(a,b,mode);
        return QString::compare(QString::fromUtf8(a.data(),a.size()),QString::fromUtf8(b.data(),b.size()),Qt::CaseInsensitive);
    }
    bool isNumericText(std::string_view text) override {return isAccessNumericText(text);}
    VbaValue formatValue(const VbaValue& value,std::string_view mask,bool stringResult) override {return formatAccessValue(value,mask,stringResult);}
    VbaValue formatNumericValue(const VbaValue& value,bool currency,const VbaNumericFormatOptions& options) override {return formatAccessNumericValue(value,currency,options);}
    void message(std::string_view value) override {(void)messageBox(value,0,std::nullopt);}
    int messageBox(std::string_view prompt,int buttons,const std::optional<std::string>& title) override {check();const auto result=nativeModuleMessageBox(owner_,prompt,buttons,title);check();return result;}
    std::string inputBox(const VbaInputBoxArguments& arguments) override {check();auto result=nativeInputBox(owner_,arguments);check();return result;}
    std::shared_ptr<VbaObject> currentDatabase() override {check();return makeDaoDatabase(db_);}
    std::shared_ptr<VbaObject> databaseEngine() override {check();return makeDaoEngine(db_);}
    void executeSql(std::string_view sql) override {currentDatabase()->member("Execute",{VbaValue::string(std::string(sql))});}
    std::shared_ptr<VbaModuleState> moduleState(std::string_view source) override {check();return isVbaClassModule(source) ? VbaHost::moduleState(source) : standardVbaModuleState(db_,source);}
    std::vector<std::shared_ptr<VbaModuleState>> globalModuleStates() override {check();return publicVbaModuleStates(db_);}
    std::optional<std::string> moduleForProcedure(std::string_view procedure) override {
        check();std::optional<std::string> found;const auto target=QString::fromUtf8(procedure.data(),procedure.size());
        for(const auto& bytes:standardVbaModuleSources(db_)) {
            const auto source=decodeVbaSource(bytes);const auto code=source.toUtf8().toStdString();if(isVbaClassModule(code)) continue;
            const auto analysis=analyzeVba(code);const auto lines=source.split('\n');
            for(const auto& item:analysis.procedures) {
                if(QString::fromStdString(item.name).compare(target,Qt::CaseInsensitive)!=0) continue;
                if(item.line>0 && item.line<=lines.size() && QRegularExpression("^Private\\b",QRegularExpression::CaseInsensitiveOption).match(lines[item.line-1].trimmed()).hasMatch()) continue;
                if(found) throw VbaRuntimeError(5,"Ambiguous public procedure");found=code;
            }
        }
        return found;
    }
private:
    QPointer<QWidget> owner_;QSqlDatabase db_;std::function<void()> guard_;
};
}
