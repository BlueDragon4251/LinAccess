#pragma once
#include "vba_expression.hpp"
#include <QSqlDatabase>
class QObject;
class QWidget;
// Both form and standalone applications share the actual project's catalog identity.
std::shared_ptr<linaccess::VbaObject> nativeProjectContext(QSqlDatabase,QObject*,bool data);

linaccess::VbaValue nativeProjectSysCmd(QWidget*,QSqlDatabase,const std::vector<linaccess::VbaValue>&,const std::function<void()>&);
