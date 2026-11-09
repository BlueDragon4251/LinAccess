#pragma once
#include "vba_expression.hpp"
#include <QSqlDatabase>
class QObject;
// Both form and standalone applications share the actual project's catalog identity.
std::shared_ptr<linaccess::VbaObject> nativeProjectContext(QSqlDatabase,QObject*,bool data);
