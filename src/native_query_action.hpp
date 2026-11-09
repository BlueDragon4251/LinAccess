#pragma once
#include <QJsonObject>
#include <QSqlDatabase>
#include <QStringList>
#include <functional>
#include <optional>
class QWidget;
namespace linaccess {
// Only SELECT-derived actions; execution remains with the existing native DAO path.
QString nativeQueryActionSql(QSqlDatabase database,const QJsonObject& action,
    const QString& header,const QString& selection,const QString& intoSelection,
    const QString& physicalSource,const QString& predicate);
std::optional<QJsonObject> editNativeQueryAction(QWidget* parent,QSqlDatabase database,
    const QStringList& outputs,const QString& physicalSource,const QJsonObject& initial,const std::function<void()>& checkProject);
}
