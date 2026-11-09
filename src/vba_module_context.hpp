#pragma once
#include "vba_runtime.hpp"
#include <QMap>
#include <QByteArray>
#include <QString>
class QSqlDatabase;
namespace linaccess {
bool isVbaClassModule(std::string_view source);
QMap<QString,QByteArray> standardVbaModuleSources(QSqlDatabase db,QMap<QString,QString>* canonicalNames=nullptr);
std::shared_ptr<VbaModuleState> standardVbaModuleState(QSqlDatabase db,std::string_view source);
std::vector<std::shared_ptr<VbaModuleState>> publicVbaModuleStates(QSqlDatabase db);
void resetVbaProjectState(QSqlDatabase db);
}
