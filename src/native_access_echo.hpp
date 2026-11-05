#pragma once
#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "native_access_status.hpp"
#include <QApplication>
#include <QEvent>
#include <QPointer>
#include <QVector>

namespace linaccess {
class NativeEchoState final : public QObject {
public:
    explicit NativeEchoState(QObject* parent) : QObject(parent) { setObjectName("access-native-echo-state"); }
    ~NativeEchoState() override { restore(); }
    void enable(bool enabled) {
        if(suppressed_==!enabled) return;
        suppressed_=!enabled;
        if(suppressed_) {
            qApp->installEventFilter(this);
            for(auto* widget:QApplication::allWidgets()) observe(widget);
        } else {
            qApp->removeEventFilter(this); restore();
        }
    }
protected:
    bool eventFilter(QObject* watched,QEvent* event) override {
        if(suppressed_ && event->type()==QEvent::Show)
            if(auto* widget=qobject_cast<QWidget*>(watched)) observe(widget);
        return false;
    }
private:
    struct Surface { QPointer<QWidget> widget; bool updatesEnabled{}; };
    bool suppressed_{};
    QVector<Surface> surfaces_;
    static bool ordinaryForm(AccessFormPreview* form) {
        if(!form || !form->isWindow() || form->isModal()) return false;
        try {
            const auto object=makePreviewFormObject(*form);
            return !object->member("PopUp",{}).truthy() && !object->member("Modal",{}).truthy();
        } catch(const std::exception&) { return false; }
    }
    void suspend(QWidget* widget) {
        if(!widget) return;
        for(const auto& surface:surfaces_) if(surface.widget==widget) return;
        surfaces_.append({widget,widget->updatesEnabled()}); widget->setUpdatesEnabled(false);
    }
    void observe(QWidget* widget) {
        if(auto* window=qobject_cast<QMainWindow*>(widget)) { suspend(window->centralWidget()); return; }
        if(auto* form=dynamic_cast<AccessFormPreview*>(widget);ordinaryForm(form)) {
            for(auto* child:form->findChildren<QWidget*>(QString{},Qt::FindDirectChildrenOnly))
                if(!child->isWindow() && !qobject_cast<QStatusBar*>(child)) suspend(child);
        }
        if(auto* form=dynamic_cast<AccessFormPreview*>(widget->parentWidget());ordinaryForm(form))
            if(!widget->isWindow() && !qobject_cast<QStatusBar*>(widget)) suspend(widget);
    }
    void restore() {
        auto surfaces=std::move(surfaces_); surfaces_.clear();
        for(const auto& surface:surfaces) if(surface.widget) surface.widget->setUpdatesEnabled(surface.updatesEnabled);
    }
};
inline VbaValue nativeEcho(QWidget& document,const std::vector<VbaValue>& arguments) {
    auto count=arguments.size();
    while(count && arguments[count-1].kind==VbaValue::Kind::Missing) --count;
    if(count<1 || count>2 || arguments[0].kind==VbaValue::Kind::Missing)
        throw VbaRuntimeError(450,"Echo requires EchoOn and optional status text");
    const bool enabled=coerceVbaModuleValue(scalarVbaValue(arguments[0]),"boolean").boolean;
    QString text;
    if(count==2) {
        text=QString::fromStdString(coerceVbaModuleValue(scalarVbaValue(arguments[1]),"string").asText());
        if(text.contains(QChar(0)) || text.contains('\n') || text.contains('\r')) throw VbaRuntimeError(5,"Invalid Echo status text");
    }
    auto* object=qApp->findChild<QObject*>("access-native-echo-state",Qt::FindDirectChildrenOnly);
    auto* state=dynamic_cast<NativeEchoState*>(object);
    if(object && !state) throw VbaRuntimeError(3251,"Native Echo state is unavailable");
    if(!state && !enabled) state=new NativeEchoState(qApp);
    if(state) state->enable(enabled);
    if(count==2) nativeStatusSurface(document)->showMessage(text);
    return {};
}
} // namespace linaccess
