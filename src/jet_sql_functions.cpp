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
    Api() {
        library.setLoadHints(QLibrary::PreventUnloadHint);
        create = reinterpret_cast<Create>(library.resolve("sqlite3_create_function_v2"));
        type = reinterpret_cast<decltype(type)>(library.resolve("sqlite3_value_type"));
        number = reinterpret_cast<decltype(number)>(library.resolve("sqlite3_value_double"));
        text = reinterpret_cast<decltype(text)>(library.resolve("sqlite3_value_text"));
        bytes = reinterpret_cast<decltype(bytes)>(library.resolve("sqlite3_value_bytes"));
        result = reinterpret_cast<decltype(result)>(library.resolve("sqlite3_result_value"));
        error = reinterpret_cast<decltype(error)>(library.resolve("sqlite3_result_error"));
    }
    bool valid() const { return create && type && number && text && bytes && result && error; }
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
}
void installJetSqlFunctions(QSqlDatabase db) {
    auto& native = api();
    auto handle = db.driver()->handle();
    if (!native.valid() || !handle.isValid() || QByteArray(handle.typeName()) != "sqlite3*")
        throw VbaRuntimeError(3251,"Jet scalar adapters unavailable");
    auto* connection = *static_cast<sqlite3**>(handle.data());
    if (!connection) throw VbaRuntimeError(3251,"Jet scalar connection unavailable");
    // Probe the live connection, rather than caching a native pointer that could
    // be reused after close/open. Preparing this call does not execute arguments.
    QSqlQuery probe(db);
    if (probe.prepare("SELECT linaccess_jet_iif(0,NULL,NULL)")) return;
    probe.clear();
    if (native.create(connection,"linaccess_jet_iif",3,1,nullptr,&iif,nullptr,nullptr,nullptr) != 0)
        throw VbaRuntimeError(3251,"Jet scalar adapter registration failed");
}
}
