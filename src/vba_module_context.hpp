#pragma once
#include "vba_runtime.hpp"
#include <QMap>
#include <QByteArray>
#include <QString>
#include <QSet>
class QSqlDatabase;
namespace linaccess {
bool isVbaClassModule(std::string_view source);
QMap<QString,QByteArray> standardVbaModuleSources(QSqlDatabase db,QMap<QString,QString>* canonicalNames=nullptr,QSet<QString>* standaloneClasses=nullptr);
VbaEnumAliases projectVbaEnumAliases(QSqlDatabase db);
std::shared_ptr<VbaModuleState> standardVbaModuleState(QSqlDatabase db,std::string_view source,const VbaEnumAliases* aliases=nullptr);
std::vector<std::shared_ptr<VbaModuleState>> publicVbaModuleStates(QSqlDatabase db,const VbaEnumAliases* aliases=nullptr);
std::shared_ptr<VbaObject> projectVbaErrorObject(QSqlDatabase db);
std::shared_ptr<VbaRandomState> projectVbaRandomState(QSqlDatabase db);
void resetVbaProjectState(QSqlDatabase db);
}
