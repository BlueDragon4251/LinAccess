#pragma once
#include "vba_runtime.hpp"
class QSqlDatabase;
namespace linaccess {
bool isVbaClassModule(std::string_view source);
std::shared_ptr<VbaModuleState> standardVbaModuleState(QSqlDatabase db,std::string_view source);
std::vector<std::shared_ptr<VbaModuleState>> publicVbaModuleStates(QSqlDatabase db);
void resetVbaProjectState(QSqlDatabase db);
}
