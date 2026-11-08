#pragma once
#include <QSqlDatabase>
#include <QSet>
#include <QString>
struct sqlite3_context;
struct sqlite3_value;
namespace linaccess {
QString accessCounterDefault(const QString& table,const QString& column);
void createAccessCounter(QSqlDatabase db,const QString& table,const QString& column);
void protectAccessCounterUpdates(QSqlDatabase db,const QString& table,const QString& column);
void restoreAccessCounter(QSqlDatabase db,const QString& table,const QString& column,qint32 last,bool native);
QSet<QString> generatedAccessColumns(QSqlDatabase db,const QString& table);
// Registered only after installJetSqlFunctions verifies the SQLite backend.
void nextAccessCounter(sqlite3_context* context,int count,sqlite3_value** values) noexcept;
}
