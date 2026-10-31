#pragma once
#include <QSqlDatabase>
namespace linaccess {
// Registers internal scalar adapters on the distro SQLite connection.
void installJetSqlFunctions(QSqlDatabase db);
}
