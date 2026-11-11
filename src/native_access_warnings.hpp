#pragma once
#include "vba_expression.hpp"
#include <QCoreApplication>
#include <QThread>
#include <QVariant>
#include <functional>

namespace linaccess {
inline QCoreApplication* nativeWarningsApplication() {
    auto* app=QCoreApplication::instance();
    if(!app) throw VbaRuntimeError(3420,"Application is unavailable");
    if(app->thread()!=QThread::currentThread()) throw VbaRuntimeError(3251,"Cross-thread warning state is unsupported");
    return app;
}
inline bool nativeWarningsEnabled() {
    const auto state=nativeWarningsApplication()->property("la_access_warnings_enabled");
    if(!state.isValid()) return true;
    if(state.metaType().id()!=QMetaType::Bool) throw VbaRuntimeError(3251,"Application warning state is unavailable");
    return state.toBool();
}
inline void setNativeWarningsEnabled(bool enabled) {
    nativeWarningsApplication()->setProperty("la_access_warnings_enabled",enabled);
}
inline void restoreNativeWarnings() noexcept {
    auto* app=QCoreApplication::instance();
    if(app && app->thread()==QThread::currentThread()) app->setProperty("la_access_warnings_enabled",true);
}
inline VbaValue nativeSetWarnings(const std::vector<VbaValue>& args,const std::function<void()>& check) {
    if(args.size()!=1 || args[0].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(450,"SetWarnings requires WarningsOn");
    const bool enabled=coerceVbaModuleValue(scalarVbaValue(args[0]),"boolean").boolean;
    check();setNativeWarningsEnabled(enabled);return VbaValue::empty();
}
}
