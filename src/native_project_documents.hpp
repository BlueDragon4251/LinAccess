#pragma once
#include "vba_expression.hpp"
#include <QSqlDatabase>
class QObject;
std::shared_ptr<linaccess::VbaObject> nativeProjectDocuments(QSqlDatabase,QObject*,bool reports);
std::shared_ptr<linaccess::VbaObject> nativeProjectScreen(QSqlDatabase,QObject*);
