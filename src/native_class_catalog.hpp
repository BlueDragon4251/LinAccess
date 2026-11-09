#pragma once
#include "vba_expression.hpp"
#include <QSqlDatabase>
#include <functional>
class QWidget;
namespace linaccess {
bool nativeProjectClassTypeKnown(QSqlDatabase database,std::string_view name);
void validateNativeProjectClassFields(QSqlDatabase database,const VbaModuleState& state);
bool nativeProjectClassAvailable(QSqlDatabase database,std::string_view name);
std::shared_ptr<VbaObject> constructNativeProjectClass(QSqlDatabase database,QWidget* caller,std::string_view name,std::function<void()> standaloneGuard={});
}
