#pragma once
#include "vba_expression.hpp"
#include <QSqlDatabase>
#include <QMap>
#include <QString>
namespace linaccess {
// True only for an owned native Date declaration with canonical storage.
bool nativeFieldHasCanonicalDateStorage(QSqlDatabase database,const QString& table,const QString& field);
// Atomic same-type native field edit: Required, DefaultValue, AllowZeroLength,
// numeric ValidationRule/ValidationText, Text Size, or a separate Name rename.
// Existing values are never truncated or filled.
void updateNativeFieldProperties(QSqlDatabase database,const QString& table,const QString& field,const QMap<QString,VbaValue>& properties);
}
