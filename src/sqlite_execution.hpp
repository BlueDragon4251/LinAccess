#pragma once
#include <QSqlDatabase>
#include <memory>
namespace linaccess {
// Applies only to this connection and this synchronous execution scope. The
// distro Qt SQLite plugin exposes its native handle; other drivers are rejected.
class SqliteExecutionBudget {
public:
    explicit SqliteExecutionBudget(QSqlDatabase db, qint64 operations = 20000000,
                                   int milliseconds = 5000);
    ~SqliteExecutionBudget();
    SqliteExecutionBudget(const SqliteExecutionBudget&) = delete;
    SqliteExecutionBudget& operator=(const SqliteExecutionBudget&) = delete;
private:
    struct Impl;
    std::unique_ptr<Impl> impl_;
};
}
