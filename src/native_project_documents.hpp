#pragma once
#include "vba_expression.hpp"
#include <QSqlDatabase>
#include <QPointer>
#include <QVector>
class QObject;
class QWidget;
class AccessFormPreview;
namespace linaccess {enum class AccessDocumentKind;}
std::shared_ptr<linaccess::VbaObject> nativeProjectDocuments(QSqlDatabase,QObject*,bool reports);
std::shared_ptr<linaccess::VbaObject> nativeProjectScreen(QSqlDatabase,QObject*);

QVector<QPointer<AccessFormPreview>> nativeLoadedProjectDocuments(QSqlDatabase,QObject*);

void openNativeProjectDocument(QSqlDatabase,QWidget*,linaccess::AccessDocumentKind,const QString&,const QString&,std::optional<QString>,int windowMode,int dataMode);

linaccess::VbaValue nativeProjectRepaintObject(QSqlDatabase,QWidget*,const std::vector<linaccess::VbaValue>&,AccessFormPreview* caller=nullptr);

linaccess::VbaValue nativeProjectApplyFilter(QSqlDatabase,QWidget*,const std::vector<linaccess::VbaValue>&,bool clear,AccessFormPreview* caller=nullptr);

linaccess::VbaValue nativeProjectWindowState(QSqlDatabase,QWidget*,const QString& command);
void nativeProjectMoveSize(QSqlDatabase,QWidget*,const std::vector<linaccess::VbaValue>&);
