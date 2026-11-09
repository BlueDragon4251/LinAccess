#pragma once
#include "native_macro.hpp"
#include <QScopeGuard>
#include <functional>
namespace linaccess {
struct NativeMacroFrame {QString connection,database,name;};
QVector<NativeMacroFrame>& nativeMacroFrames();
bool nativeProjectMacrosActive(QSqlDatabase db);
bool nativeMacroObjectActive(QSqlDatabase db,const QString& name);
bool& nativeMacroStopAllPending();
void runNativeMacro(QSqlDatabase db,const std::vector<VbaValue>& args,
    const std::function<void()>& check,const std::function<VbaValue(const QString&)>& evaluate,
    const std::function<void(const NativeMacroAction&)>& dispatch,const std::function<void()>& restoreUi={});
}
