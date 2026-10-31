#include "jet_sql_functions.hpp"
#include "access_format.hpp"
#include "access_query.hpp"
#include <QLocale>
#include <QObject>
#include <charconv>
#include <cmath>
#include <cstring>
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
    qlonglong (*integer)(sqlite3_value*){};
    void* (*userData)(sqlite3_context*){};
    void (*integerResult)(sqlite3_context*,qlonglong){};
    void (*numberResult)(sqlite3_context*,double){};
    const unsigned char* (*text)(sqlite3_value*){};
    int (*bytes)(sqlite3_value*){};
    void (*result)(sqlite3_context*,sqlite3_value*){};
    void (*error)(sqlite3_context*,const char*,int){};
    void* (*aggregate)(sqlite3_context*,int){};
    sqlite3_value* (*duplicate)(const sqlite3_value*){};
    void (*freeValue)(sqlite3_value*){};
    void (*nullResult)(sqlite3_context*){};
    void (*textResult)(sqlite3_context*,const char*,int,void (*)(void*)){};
    Api() {
        library.setLoadHints(QLibrary::PreventUnloadHint);
        create = reinterpret_cast<Create>(library.resolve("sqlite3_create_function_v2"));
        type = reinterpret_cast<decltype(type)>(library.resolve("sqlite3_value_type"));
        number = reinterpret_cast<decltype(number)>(library.resolve("sqlite3_value_double"));
        integer = reinterpret_cast<decltype(integer)>(library.resolve("sqlite3_value_int64"));
        userData = reinterpret_cast<decltype(userData)>(library.resolve("sqlite3_user_data"));
        integerResult = reinterpret_cast<decltype(integerResult)>(library.resolve("sqlite3_result_int64"));
        numberResult = reinterpret_cast<decltype(numberResult)>(library.resolve("sqlite3_result_double"));
        text = reinterpret_cast<decltype(text)>(library.resolve("sqlite3_value_text"));
        bytes = reinterpret_cast<decltype(bytes)>(library.resolve("sqlite3_value_bytes"));
        result = reinterpret_cast<decltype(result)>(library.resolve("sqlite3_result_value"));
        error = reinterpret_cast<decltype(error)>(library.resolve("sqlite3_result_error"));
        aggregate = reinterpret_cast<decltype(aggregate)>(library.resolve("sqlite3_aggregate_context"));
        duplicate = reinterpret_cast<decltype(duplicate)>(library.resolve("sqlite3_value_dup"));
        freeValue = reinterpret_cast<decltype(freeValue)>(library.resolve("sqlite3_value_free"));
        nullResult = reinterpret_cast<decltype(nullResult)>(library.resolve("sqlite3_result_null"));
        textResult = reinterpret_cast<decltype(textResult)>(library.resolve("sqlite3_result_text"));
    }
    bool valid() const { return create && type && number && integer && userData && integerResult && numberResult && text && bytes && result && error && aggregate && duplicate && freeValue && nullResult && textResult; }
};
Api& api() { static Api value; return value; }
class ObjectContext final : public QObject {
public:
    explicit ObjectContext(QObject* owner) : QObject(owner) { setObjectName("linaccess-sql-objects"); }
    JetSqlObjectResolver resolver;
};
ObjectContext* objectContext(QSqlDatabase db) {
    auto* driver=db.driver();
    for (auto* child : driver->children()) if (auto* context=dynamic_cast<ObjectContext*>(child)) return context;
    return new ObjectContext(driver);
}
void functionError(sqlite3_context* context,const VbaRuntimeError& error) noexcept {
    constexpr char prefix[]="LinAccess SQL function error ";
    char buffer[80];
    std::memcpy(buffer,prefix,sizeof(prefix)-1);
    const auto result=std::to_chars(buffer+sizeof(prefix)-1,buffer+sizeof(buffer),error.number);
    api().error(context,buffer,static_cast<int>(result.ptr-buffer));
}
QString scalarText(sqlite3_value* value) {
    auto& native=api();
    const auto type=native.type(value);
    if (type==5) return {};
    if (type==3) {
        const auto* text=native.text(value);
        if (!text) throw VbaRuntimeError(7,"SQL string unavailable");
        return QString::fromUtf8(reinterpret_cast<const char*>(text),native.bytes(value));
    }
    QLocale locale;
    locale.setNumberOptions(locale.numberOptions()|QLocale::OmitGroupSeparator);
    if (type==1) return locale.toString(native.integer(value));
    if (type==2) {
        const auto number=native.number(value);
        if (!std::isfinite(number)) throw VbaRuntimeError(6,"SQL numeric string overflow");
        return locale.toString(number,'g',15);
    }
    throw VbaRuntimeError(13,"SQL value cannot be concatenated");
}
void returnText(sqlite3_context* context,const QString& value) {
    const auto bytes=value.toUtf8();
    api().textResult(context,bytes.constData(),bytes.size(),reinterpret_cast<void (*)(void*)>(-1));
}
void concatenate(sqlite3_context* context,int,sqlite3_value** args) noexcept {
    auto& native=api();
    try {
        if (native.type(args[0])==5 && native.type(args[1])==5) { native.nullResult(context); return; }
        returnText(context,scalarText(args[0])+scalarText(args[1]));
    } catch (const VbaRuntimeError& error) { functionError(context,error); }
      catch (...) { native.error(context,"Jet concatenation failed",-1); }
}
void likePattern(sqlite3_context* context,int,sqlite3_value** args) noexcept {
    auto& native=api();
    try {
        if (native.type(args[0])==5) { native.nullResult(context); return; }
        returnText(context,accessLikePattern(scalarText(args[0])));
    } catch (const VbaRuntimeError& error) { functionError(context,error); }
      catch (...) { native.error(context,"Jet LIKE pattern failed",-1); }
}
void objectValue(sqlite3_context* context,int,sqlite3_value** args) noexcept {
    auto& native=api();
    try {
        auto* objects=static_cast<ObjectContext*>(native.userData(context));
        if (!objects || !objects->resolver) throw VbaRuntimeError(3061,"SQL object context unavailable");
        const auto collection=scalarText(args[0]).toUtf8();
        const auto object=scalarText(args[1]).toUtf8();
        const auto field=scalarText(args[2]).toUtf8();
        const auto result=objects->resolver(
            std::string_view(collection.constData(),collection.size()),std::string_view(object.constData(),object.size()),std::string_view(field.constData(),field.size()));
        switch (result.kind) {
        case VbaValue::Kind::Null: native.nullResult(context); break;
        case VbaValue::Kind::Empty: native.integerResult(context,0); break;
        case VbaValue::Kind::Boolean: native.integerResult(context,result.boolean ? -1 : 0); break;
        case VbaValue::Kind::Number:
            if (!std::isfinite(result.number)) throw VbaRuntimeError(6,"SQL object number overflow");
            native.numberResult(context,result.number); break;
        case VbaValue::Kind::Text:
            native.textResult(context,result.text.data(),static_cast<int>(result.text.size()),reinterpret_cast<void (*)(void*)>(-1)); break;
        default: throw VbaRuntimeError(13,"SQL object value must be scalar");
        }
    } catch (const VbaRuntimeError& error) { functionError(context,error); }
      catch (...) { native.error(context,"SQL object resolution failed",-1); }
}
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
void formatDate(sqlite3_context* context,sqlite3_value** args,bool stringResult) noexcept {
    auto& native=api();
    try {
        if (native.type(args[1])==5) throw VbaRuntimeError(94,"Format mask is Null");
        if (native.type(args[1])!=3) throw VbaRuntimeError(13,"Format mask requires text");
        const auto readText=[&](sqlite3_value* value) {
            const auto* data=native.text(value);
            if (!data) throw VbaRuntimeError(7,"Format string unavailable");
            return std::string(reinterpret_cast<const char*>(data),native.bytes(value));
        };
        VbaValue input;
        const auto type=native.type(args[0]);
        if (type==1 || type==2) input=VbaValue::numeric(native.number(args[0]));
        else if (type==3) input=VbaValue::string(readText(args[0]));
        else if (type!=5) throw VbaRuntimeError(13,"Format input type unsupported");
        const auto result=formatAccessValue(input,readText(args[1]),stringResult);
        if (result.kind==VbaValue::Kind::Null) native.nullResult(context);
        else native.textResult(context,result.text.data(),static_cast<int>(result.text.size()),reinterpret_cast<void (*)(void*)>(-1));
    } catch (const VbaRuntimeError& error) { functionError(context,error); }
      catch (...) { native.error(context,"Jet date format failed",-1); }
}
void formatString(sqlite3_context* context,int,sqlite3_value** args) noexcept { formatDate(context,args,true); }
void formatVariant(sqlite3_context* context,int,sqlite3_value** args) noexcept { formatDate(context,args,false); }
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
std::optional<int> jetSqlFunctionErrorNumber(std::string_view foldedMessage) {
    constexpr std::string_view prefix="linaccess sql function error ";
    const auto position=foldedMessage.find(prefix);
    if (position==std::string_view::npos) return {};
    const auto code=foldedMessage.substr(position+prefix.size());
    int number=0;
    const auto parsed=std::from_chars(code.data(),code.data()+code.size(),number);
    if (parsed.ec!=std::errc{}) return {};
    if (number==5 || number==6 || number==7 || number==13 || number==94 || number==3061 || number==2450 || number==2451 || number==2465 || number==3251 || number==3085) return number;
    return {};
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
    auto* objects=objectContext(db);
    const auto ensure = [&](const char* name,int count,Api::Callback scalar,Api::Callback step,void (*final)(sqlite3_context*)) {
        QSqlQuery probe(db);
        QStringList args;
        for (int i = 0; i < count; ++i) args.append("NULL");
        if (probe.prepare("SELECT " + QString::fromLatin1(name) + "(" + args.join(',') + ")")) return;
        probe.clear();
        if (native.create(connection,name,count,1,objects,scalar,step,final,nullptr) != 0)
            throw VbaRuntimeError(3251,"Jet SQL function registration failed");
    };
    ensure("linaccess_jet_iif",3,&iif,nullptr,nullptr);
    ensure("linaccess_jet_first",1,nullptr,&firstStep,&aggregateFinal);
    ensure("linaccess_jet_last",1,nullptr,&lastStep,&aggregateFinal);
    ensure("linaccess_jet_format_string",2,&formatString,nullptr,nullptr);
    ensure("linaccess_jet_format",2,&formatVariant,nullptr,nullptr);
    ensure("linaccess_jet_concat",2,&concatenate,nullptr,nullptr);
    ensure("linaccess_jet_like_pattern",1,&likePattern,nullptr,nullptr);
    ensure("linaccess_jet_object_value",3,&objectValue,nullptr,nullptr);
}
void setJetSqlObjectResolver(QSqlDatabase db,JetSqlObjectResolver resolver) {
    installJetSqlFunctions(db);
    objectContext(db)->resolver=std::move(resolver);
}

}
