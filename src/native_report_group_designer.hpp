#pragma once
#include <QJsonArray>
#include <QSqlDatabase>
#include <QString>
#include <functional>
#include <optional>
class QWidget;
namespace linaccess {
std::optional<QJsonArray> editNativeReportGroups(QWidget* owner,QSqlDatabase database,const QString& source,
    const QJsonArray& groups,const std::function<void()>& guard);
}
