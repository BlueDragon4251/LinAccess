#pragma once
#include <QSqlDatabase>
#include <QString>
class QWidget;
namespace linaccess {
void editNativeRelations(QWidget* parent,QSqlDatabase database,const QString& suggestedParent={});
}
