#pragma once
#include "vba_expression.hpp"
#include <QSqlDatabase>
#include <QPointer>
#include <QVector>
class QObject;
class AccessFormPreview;
std::shared_ptr<linaccess::VbaObject> nativeProjectDocuments(QSqlDatabase,QObject*,bool reports);
std::shared_ptr<linaccess::VbaObject> nativeProjectScreen(QSqlDatabase,QObject*);

QVector<QPointer<AccessFormPreview>> nativeLoadedProjectDocuments(QSqlDatabase,QObject*);
