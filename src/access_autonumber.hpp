#pragma once
#include <QSqlDatabase>
#include <QSet>
#include <QVector>
#include <QString>
#include <memory>
struct sqlite3_context;
struct sqlite3_value;
namespace linaccess {
struct AccessCounterReservation {
    QString table,column;
    qlonglong value{};
    bool native{},jetLong{};
};
// Captures generator calls made by a single application SQL operation. Failed
// statements restore consumed frontiers without overwriting a competing writer.
class AccessCounterSqlScope {
public:
    explicit AccessCounterSqlScope(QSqlDatabase db);
    ~AccessCounterSqlScope();
    AccessCounterSqlScope(const AccessCounterSqlScope&)=delete;
    AccessCounterSqlScope& operator=(const AccessCounterSqlScope&)=delete;
    void complete(bool succeeded);
    struct State;
private:
    std::unique_ptr<State> state_;
};
void trackAccessCounterReservations(QSqlDatabase db,const QVector<AccessCounterReservation>& reservations);
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
