#pragma once
#include <QSqlDatabase>
namespace linaccess {
// Registers internal scalar and aggregate adapters on the distro SQLite connection.
void installJetSqlFunctions(QSqlDatabase db);
}
