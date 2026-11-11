#pragma once
#include "vba_runtime.hpp"
#include <QByteArray>
#include <QMap>
#include <QSet>
#include <QString>
#include <functional>
class QDialog;class QPlainTextEdit;class QSqlDatabase;
struct NativeModuleBreakpoints {QMap<QByteArray,QSet<int>> lines;};
void installNativeModuleBreakpoints(QDialog&,QPlainTextEdit&,const std::shared_ptr<NativeModuleBreakpoints>&,QSqlDatabase);
struct NativeModuleDebugResult {linaccess::VbaExecutionResult execution;bool stopped{};QString debuggerError;};
NativeModuleDebugResult runNativeModuleDebugger(QDialog& editor,QSqlDatabase database,const QString& moduleName,const QByteArray& code,const QString& procedure,NativeModuleBreakpoints& breakpoints,bool stepFirst,const std::function<void()>& checkProject);
