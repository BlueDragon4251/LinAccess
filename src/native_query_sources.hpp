#pragma once
#include <QJsonArray>
#include <QMap>
#include <QSqlDatabase>
#include <QStringList>
#include <functional>
#include <optional>
class QWidget;
namespace linaccess {
struct NativeQuerySources {
    QString sql;
    QStringList order;
    QMap<QString,QStringList> fields;
    QMap<QString,QString> origins;
};
NativeQuerySources nativeQuerySources(QSqlDatabase database,const QString& main,const QJsonArray& joins);
std::optional<QJsonArray> editNativeQueryJoins(QWidget* parent,QSqlDatabase database,
    const QString& main,const QJsonArray& initial,const std::function<void()>& checkProject);
}
