#pragma once
#include <QSqlDatabase>
#include <QSet>
#include <QVector>
#include <QString>
struct sqlite3_context;
struct sqlite3_value;
namespace linaccess {
struct AccessCounterReservation {
    QString table,column;
    qlonglong value{};
    bool native{},jetLong{};
};
AccessCounterReservation reserveAccessCounter(QSqlDatabase db,const QString& table,const QString& column);
AccessCounterReservation currentAccessCounterReservation(QSqlDatabase db,AccessCounterReservation reservation);
void configureNewAccessCounterStep(QSqlDatabase db,const QString& table,const QString& column,qint32 seed,qint32 increment);
void configureNewAccessCounters(QSqlDatabase db,const QString& table);
QSet<QString> jetLongAccessColumns(QSqlDatabase db,const QString& table);
void createJetLongCounter(QSqlDatabase db,const QString& table,const QString& column);
void retainAccessCounterReservation(QSqlDatabase db,const AccessCounterReservation& reservation);
QString accessCounterDefault(const QString& table,const QString& column);
void createAccessCounter(QSqlDatabase db,const QString& table,const QString& column);
void protectAccessCounterUpdates(QSqlDatabase db,const QString& table,const QString& column);
void restoreAccessCounter(QSqlDatabase db,const QString& table,const QString& column,qint32 last,bool native);
QSet<QString> generatedAccessColumns(QSqlDatabase db,const QString& table);
// Registered only after installJetSqlFunctions verifies the SQLite backend.
void nextAccessCounter(sqlite3_context* context,int count,sqlite3_value** values) noexcept;
}
