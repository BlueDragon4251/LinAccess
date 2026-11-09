#pragma once
#include <QJsonObject>
#include <QSqlDatabase>
#include <QString>
#include <functional>
#include <optional>
class QWidget;
namespace linaccess {
struct NativeQueryDesign {QString sql;QJsonObject state;};
std::optional<NativeQueryDesign> editNativeQueryDesign(QWidget* parent,QSqlDatabase database,
    const QJsonObject& initial,const std::function<void()>& checkProject,bool allowParameters=true);
}
