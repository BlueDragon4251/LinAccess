#pragma once
#include <QSqlDatabase>
#include <QMap>
#include "vba_expression.hpp"
#include <optional>
#include <functional>
#include <string_view>
#include <string>
namespace linaccess {
// Registers internal scalar and aggregate adapters on the verified SQLite connection.
std::optional<int> jetSqlFunctionErrorNumber(std::string_view foldedMessage);
std::string jetSqlFunctionErrorExplanation(std::string_view originalDatabaseMessage);
using JetSqlObjectResolver=std::function<VbaValue(std::string_view collection,std::string_view object,std::string_view field)>;
void setJetSqlObjectResolver(QSqlDatabase db,JetSqlObjectResolver resolver);
void installJetSqlFunctions(QSqlDatabase db);
QMap<QString,QString> nativeVbaSqlFunctionResultTypes(QSqlDatabase db);
}
