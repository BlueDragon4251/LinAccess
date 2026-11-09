#pragma once
#include "vba_runtime.hpp"
#include "vba_expression.hpp"
#include "native_macro_function.hpp"
#include "native_macro_set_value_contract.hpp"
#include <QJsonArray>
#include <QElapsedTimer>
#include <QJsonDocument>
#include <QJsonObject>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlError>
#include <QStringList>
#include <QVector>
#include <cmath>
#include <limits>
namespace linaccess {
struct NativeMacroSignature {QString name;QStringList types;int required;};
struct NativeMacroAction {QString name;std::vector<VbaValue> arguments;std::optional<QString> condition;};
const QVector<NativeMacroSignature>& nativeMacroSignatures();
bool nativeMacroNameValid(const QString& name);
bool nativeMacroCallNameValid(const QString& name);
std::optional<QString> nativeMacroRepeatExpression(const QString& value);
QVector<NativeMacroAction> parseNativeMacro(const QJsonObject& definition,const QString& requestedGroup={});
QString nativeMacroCatalogName(QSqlDatabase database,const QString& requested);
QVector<NativeMacroAction> loadNativeMacro(QSqlDatabase database,const QString& name);
}
