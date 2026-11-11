#pragma once
#include "vba_expression.hpp"
#include <QMap>
#include <QSqlDatabase>
#include <QVariant>
namespace linaccess {
QMap<QString,QString> sqliteColumnTypes(QSqlDatabase db,const QString& table);
QString qtFieldText(const QVariant& value,const QString& type);
VbaValue qtFieldValue(const QVariant& value,const QString& sqlType = {});
QVariant sqlDecimalFieldValue(const VbaValue& value,int precision,int scale);
QVariant sqlFieldValue(const VbaValue& value,const QString& sqlType = {});
}
