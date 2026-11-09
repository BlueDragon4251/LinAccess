#pragma once
#include <QJsonArray>
#include <QJsonValue>
#include <QStringList>
#include <QSet>
class QWidget;
namespace linaccess {
struct NativeFormControls {QStringList lines,names;int width{},bottom{};};
NativeFormControls nativeFormControls(const QJsonValue& definition,const QStringList& fields,const QStringList& schema,const QSet<QString>& procedures);
bool editNativeFormControls(QWidget* parent,QJsonArray& controls,const QStringList& fields,const QStringList& schema);
}
