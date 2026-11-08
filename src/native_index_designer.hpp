#pragma once
#include <QSqlDatabase>
#include <QString>
class QWidget;
namespace linaccess {
void editNativeIndexes(QWidget* parent,QSqlDatabase database,const QString& table);
}
