#pragma once
#include "vba_expression.hpp"
#include <QSqlDatabase>
#include <QMap>
#include <QString>
namespace linaccess {
// One atomic same-type native field edit; supported keys: Required,
// DefaultValue and Text Size, or a separate Name rename. Existing values are
// never truncated or filled.
void updateNativeFieldProperties(QSqlDatabase database,const QString& table,const QString& field,const QMap<QString,VbaValue>& properties);
}
