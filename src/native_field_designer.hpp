#pragma once
#include <QSqlDatabase>
#include <QString>
class QWidget;
namespace linaccess {
void editNativeFields(QWidget* parent,QSqlDatabase database,const QString& table);
}
