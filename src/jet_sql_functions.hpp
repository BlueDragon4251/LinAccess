#pragma once
#include <QSqlDatabase>
#include <optional>
#include <string_view>
namespace linaccess {
// Registers internal scalar and aggregate adapters on the distro SQLite connection.
std::optional<int> jetSqlFunctionErrorNumber(std::string_view foldedMessage);
void installJetSqlFunctions(QSqlDatabase db);
}
