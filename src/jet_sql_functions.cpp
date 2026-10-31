#include "jet_sql_functions.hpp"
#include "vba_runtime.hpp"
#include <QLibrary>
#include <QSqlDriver>
#include <QSqlQuery>
struct sqlite3;
struct sqlite3_context;
struct sqlite3_value;
namespace linaccess {
namespace {
struct Api {
    QLibrary library{"sqlite3",0};
    using Callback = void (*)(sqlite3_context*,int,sqlite3_value**);
    using Create = int (*)(sqlite3*,const char*,int,int,void*,Callback,Callback,void (*)(sqlite3_context*),void (*)(void*));
    Create create{};
    int (*type)(sqlite3_value*){};
    double (*number)(sqlite3_value*){};
    const unsigned char* (*text)(sqlite3_value*){};
    int (*bytes)(sqlite3_value*){};
    void (*result)(sqlite3_context*,sqlite3_value*){};
    void (*error)(sqlite3_context*,const char*,int){};
    void* (*aggregate)(sqlite3_context*,int){};
    sqlite3_value* (*duplicate)(const sqlite3_value*){};
    void (*freeValue)(sqlite3_value*){};
    void (*nullResult)(sqlite3_context*){};
    Api() {
        library.setLoadHints(QLibrary::PreventUnloadHint);
        create = reinterpret_cast<Create>(library.resolve("sqlite3_create_function_v2"));
        type = reinterpret_cast<decltype(type)>(library.resolve("sqlite3_value_type"));
        number = reinterpret_cast<decltype(number)>(library.resolve("sqlite3_value_double"));
        text = reinterpret_cast<decltype(text)>(library.resolve("sqlite3_value_text"));
        bytes = reinterpret_cast<decltype(bytes)>(library.resolve("sqlite3_value_bytes"));
        result = reinterpret_cast<decltype(result)>(library.resolve("sqlite3_result_value"));
        error = reinterpret_cast<decltype(error)>(library.resolve("sqlite3_result_error"));
        aggregate = reinterpret_cast<decltype(aggregate)>(library.resolve("sqlite3_aggregate_context"));
        duplicate = reinterpret_cast<decltype(duplicate)>(library.resolve("sqlite3_value_dup"));
        freeValue = reinterpret_cast<decltype(freeValue)>(library.resolve("sqlite3_value_free"));
        nullResult = reinterpret_cast<decltype(nullResult)>(library.resolve("sqlite3_result_null"));
    }
    bool valid() const { return create && type && number && text && bytes && result && error && aggregate && duplicate && freeValue && nullResult; }
};
Api& api() { static Api value; return value; }
void iif(sqlite3_context* context,int,sqlite3_value** args) noexcept {
    auto& native = api();
    try {
        bool condition = false;
        const auto type = native.type(args[0]);
        if (type == 1 || type == 2) condition = native.number(args[0]) != 0;
        else if (type == 3) {
            const auto value = QString::fromUtf8(reinterpret_cast<const char*>(native.text(args[0])),native.bytes(args[0])).trimmed();
            if (value.compare("True",Qt::CaseInsensitive)==0) condition = true;
            else if (value.compare("False",Qt::CaseInsensitive)!=0) {
                bool valid = false;
                condition = value.toDouble(&valid) != 0;
                if (!valid) { native.error(context,"Jet IIf condition type mismatch",-1); return; }
            }
        } else if (type != 5) { native.error(context,"Jet IIf condition type mismatch",-1); return; }
        native.result(context,args[condition ? 1 : 2]);
    } catch (...) { native.error(context,"Jet IIf evaluation failed",-1); }
}
// aggregate_context supplies zeroed storage. Each group retains one copied value;
// no pointers into the temporary SQLite argument storage outlive a step callback.
struct AggregateState { sqlite3_value* value; bool failed; };
void aggregateStep(sqlite3_context* context,sqlite3_value* value,bool first) noexcept {
    auto& native = api();
    auto* state = static_cast<AggregateState*>(native.aggregate(context,sizeof(AggregateState)));
    if (!state) { native.error(context,"Jet aggregate allocation failed",-1); return; }
    if (state->failed || (first && state->value)) return;
    auto* copied = native.duplicate(value);
    if (!copied) { state->failed = true; native.error(context,"Jet aggregate allocation failed",-1); return; }
    if (state->value) native.freeValue(state->value);
    state->value = copied;
}
void firstStep(sqlite3_context* context,int,sqlite3_value** args) noexcept { aggregateStep(context,args[0],true); }
void lastStep(sqlite3_context* context,int,sqlite3_value** args) noexcept { aggregateStep(context,args[0],false); }
void aggregateFinal(sqlite3_context* context) noexcept {
    auto& native = api();
    auto* state = static_cast<AggregateState*>(native.aggregate(context,0));
    if (!state) { native.nullResult(context); return; }
    if (state->failed) native.error(context,"Jet aggregate allocation failed",-1);
    else if (state->value) native.result(context,state->value);
    else native.nullResult(context);
    if (state->value) native.freeValue(state->value);
    state->value = nullptr;
}

}
void installJetSqlFunctions(QSqlDatabase db) {
    auto& native = api();
    auto handle = db.driver()->handle();
    if (!native.valid() || !handle.isValid() || QByteArray(handle.typeName()) != "sqlite3*")
        throw VbaRuntimeError(3251,"Jet SQL adapters unavailable");
    auto* connection = *static_cast<sqlite3**>(handle.data());
    if (!connection) throw VbaRuntimeError(3251,"Jet SQL connection unavailable");
    // Probe the live connection, rather than caching a native pointer that could
    // be reused after close/open. Preparing this call does not execute arguments.
    const auto ensure = [&](const char* name,int count,Api::Callback scalar,Api::Callback step,void (*final)(sqlite3_context*)) {
        QSqlQuery probe(db);
        QStringList args;
        for (int i = 0; i < count; ++i) args.append("NULL");
        if (probe.prepare("SELECT " + QString::fromLatin1(name) + "(" + args.join(',') + ")")) return;
        probe.clear();
        if (native.create(connection,name,count,1,nullptr,scalar,step,final,nullptr) != 0)
            throw VbaRuntimeError(3251,"Jet SQL function registration failed");
    };
    ensure("linaccess_jet_iif",3,&iif,nullptr,nullptr);
    ensure("linaccess_jet_first",1,nullptr,&firstStep,&aggregateFinal);
    ensure("linaccess_jet_last",1,nullptr,&lastStep,&aggregateFinal);
}
}
