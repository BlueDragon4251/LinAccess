#pragma once
#include "vba_expression.hpp"
#include <QSqlDatabase>

namespace linaccess {
// SQLite-backed subset of DAO. Unsupported members throw; original MDBs are
// never opened for writing. Database must refer to the LinAccess project copy.
std::shared_ptr<VbaObject> makeDaoDatabase(QSqlDatabase database);
}
