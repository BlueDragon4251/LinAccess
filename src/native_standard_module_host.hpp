#pragma once
#include "vba_runtime.hpp"
#include "vba_expression.hpp"
#include "vba_module_context.hpp"
#include "vba_source_codec.hpp"
#include "vba_analyzer.hpp"
#include "access_format.hpp"
#include "dao_database.hpp"
#include "native_input_box.hpp"
#include "native_public_procedure_run.hpp"
#include "native_domain_lookup.hpp"
#include <QMainWindow>
#include <QVariant>
#include <QThread>
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
    void check() const {if(!owner_ || !db_.isOpen()) throw VbaRuntimeError(3420,"Module execution context is closed");if(owner_->thread()!=QThread::currentThread()) throw VbaRuntimeError(3251,"Module application context requires its GUI thread");guard_();}
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
    std::shared_ptr<VbaObject> applicationObject() override;
    VbaValue domainAggregate(std::string_view function,const std::vector<VbaValue>& args) override {check();return nativeDomainValue(db_,args,function);}
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
class NativeModuleApplication final : public VbaObject {
public:
    NativeModuleApplication(QWidget* owner,QSqlDatabase db,std::function<void()> guard):owner_(owner),db_(std::move(db)),guard_(std::move(guard)) {}
    std::string_view className() const override {return "Application";}
    bool isNativeType(std::string_view type) const override {return type=="access.application";}
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        if(QString::fromUtf8(name.data(),name.size()).compare("Run",Qt::CaseInsensitive)==0) return nativeApplicationRunSignature();return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        NativeStandardModuleHost host(owner_,db_,guard_);host.check();const auto key=QString::fromUtf8(name.data(),name.size()).toCaseFolded();
        if(key=="run") return runPublicStandardProcedure(host,args);
        if(key=="eval") {
            if(args.size()!=1) throw VbaRuntimeError(450,"Eval requires one expression");
            const auto expression=coerceVbaModuleValue(scalarVbaValue(args.front()),"string").asText();host.check();return host.eval(expression);
        }
        if(key=="dbengine") {if(!args.empty()) throw VbaRuntimeError(450,"DBEngine requires no arguments");return VbaValue::reference(host.databaseEngine());}
        throw VbaRuntimeError(438,"Application member is unavailable in the module execution host");
    }
private:
    QPointer<QWidget> owner_;QSqlDatabase db_;std::function<void()> guard_;
};
inline std::shared_ptr<VbaObject> NativeStandardModuleHost::applicationObject() {
    check();QWidget* project=owner_;
    for(auto* ancestor=owner_->parentWidget();ancestor;ancestor=ancestor->parentWidget()) if(qobject_cast<QMainWindow*>(ancestor)) {project=ancestor;break;}
    const auto connection=db_.connectionName();
    const auto identity=QString::number(connection.size())+":"+connection+db_.databaseName()+":"+project->property("la_project_generation").toString();
    auto contexts=project->property("la_native_module_applications").toMap();
    auto object=contexts.value(identity).value<std::weak_ptr<VbaObject>>().lock();
    if(!object) {
        for(auto iterator=contexts.begin();iterator!=contexts.end();) if(iterator->value<std::weak_ptr<VbaObject>>().expired()) iterator=contexts.erase(iterator);else ++iterator;
        object=std::make_shared<NativeModuleApplication>(project,db_,guard_);contexts.insert(identity,QVariant::fromValue(std::weak_ptr<VbaObject>(object)));project->setProperty("la_native_module_applications",contexts);
    }
    return object;
}

}
