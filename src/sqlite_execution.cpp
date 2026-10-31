#include "sqlite_execution.hpp"
#include "jet_sql_functions.hpp"
#include "vba_runtime.hpp"
#include <QElapsedTimer>
#include <QLibrary>
#include <QSqlDriver>
#include <algorithm>
struct sqlite3;
namespace linaccess {
struct SqliteExecutionBudget::Impl {
    using Progress = void (*)(sqlite3*,int,int (*)(void*),void*);
    sqlite3* database{};
    Progress progress{};
    qint64 remaining{};
    int interval{};
    int milliseconds{};
    QElapsedTimer clock;
    static int check(void* pointer) {
        auto& self = *static_cast<Impl*>(pointer);
        self.remaining -= self.interval;
        return self.remaining <= 0 || self.clock.hasExpired(self.milliseconds);
    }
};
SqliteExecutionBudget::SqliteExecutionBudget(QSqlDatabase db, qint64 operations, int milliseconds)
    : impl_(std::make_unique<Impl>()) {
    static QLibrary library("sqlite3",0);
    library.setLoadHints(QLibrary::PreventUnloadHint);
    const auto function = reinterpret_cast<Impl::Progress>(library.resolve("sqlite3_progress_handler"));
    auto handle = db.driver()->handle();
    if (!function || !handle.isValid() || QByteArray(handle.typeName()) != "sqlite3*")
        throw VbaRuntimeError(3251,"SQLite query execution budget unavailable");
    installJetSqlFunctions(db);
    impl_->database = *static_cast<sqlite3**>(handle.data());
    if (!impl_->database || operations <= 0 || milliseconds <= 0)
        throw VbaRuntimeError(5,"Invalid SQLite execution budget");
    impl_->progress = function;
    impl_->remaining = operations;
    impl_->interval = static_cast<int>(std::min<qint64>(1000,operations));
    impl_->milliseconds = milliseconds;
    impl_->clock.start();
    function(impl_->database,impl_->interval,&Impl::check,impl_.get());
}
SqliteExecutionBudget::~SqliteExecutionBudget() {
    if (impl_ && impl_->progress) impl_->progress(impl_->database,0,nullptr,nullptr);
}
}
