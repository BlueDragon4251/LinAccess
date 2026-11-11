#pragma once
#include "vba_expression.hpp"
#include <QSqlDatabase>
#include <QMap>
#include <QString>
#include <optional>
namespace linaccess {
struct NativeFieldDisplayProperties {std::optional<QString> caption,description;bool nativeOwned{};std::optional<QString> format;};
// Only declared exact Decimal columns with the matching physical constraint.
QMap<QString,QPair<int,int>> nativeTableDecimalTypes(QSqlDatabase database,const QString& table);
QMap<QString,NativeFieldDisplayProperties> nativeTableFieldDisplayProperties(QSqlDatabase database,const QString& table);
// True only for an owned native Currency declaration with exact supported storage.
bool nativeFieldHasExactCurrencyStorage(QSqlDatabase database,const QString& table,const QString& field);
// True only for an owned native Date declaration with canonical storage.
bool nativeFieldHasCanonicalDateStorage(QSqlDatabase database,const QString& table,const QString& field);
// Atomic same-type native field edit: Required, DefaultValue, AllowZeroLength,
// numeric ValidationRule/ValidationText, Text Size, or a separate Name rename.
// Existing values are never truncated or filled.
// Designer operation: numeric non-counter fields, exact values/defaults only.
void changeNativeNumericFieldType(QSqlDatabase database,const QString& table,const QString& field,int type);
void updateNativeFieldProperties(QSqlDatabase database,const QString& table,const QString& field,const QMap<QString,VbaValue>& properties);
}
