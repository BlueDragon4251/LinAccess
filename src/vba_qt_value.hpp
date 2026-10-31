#pragma once
#include "vba_expression.hpp"
#include <QMap>
#include <QSqlDatabase>
#include <QVariant>
namespace linaccess {
QMap<QString,QString> sqliteColumnTypes(QSqlDatabase db,const QString& table);
VbaValue qtFieldValue(const QVariant& value,const QString& sqlType = {});
QVariant sqlFieldValue(const VbaValue& value,const QString& sqlType = {});
}
