#include "jet_sql_functions.hpp"
#include "access_autonumber.hpp"
#include "access_format.hpp"
#include "access_query.hpp"
#include <QLocale>
#include <QObject>
#include <charconv>
#include <cmath>
#include <cstring>
#include "vba_runtime.hpp"
#include "vba_expression.hpp"
#include "vba_color.hpp"
#include "vba_round.hpp"
#include "vba_calendar.hpp"
#include "vba_qt_value.hpp"
#include "vba_currency.hpp"
#include "vba_qt_value.hpp"
#include <QSet>
#include <QLibrary>
#include <QSqlDriver>
#include <QSqlQuery>
#include <QVariant>
#include <sqlite3.h>
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
        case VbaValue::Kind::Date: {
            const auto text=sqlFieldValue(result,"DATETIME").toString().toUtf8();
            native.textResult(context,text.data(),static_cast<int>(text.size()),reinterpret_cast<void (*)(void*)>(-1)); break;
        }
        case VbaValue::Kind::Currency: native.numberResult(context,currencyToRealStorage(result)); break;
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
void colorFunction(sqlite3_context* context,int count,sqlite3_value** args,bool indexed) noexcept {
    auto& native=api();
    try {
        if (count!=(indexed ? 1 : 3)) throw VbaRuntimeError(450,"Color function argument count invalid");
        std::vector<VbaValue> values;
        for(int index=0;index<count;++index) {
            const auto type=native.type(args[index]);
            if(type==5) values.emplace_back();
            else if(type==1) values.push_back(VbaValue::numeric(static_cast<double>(native.integer(args[index]))));
            else if(type==2) values.push_back(VbaValue::numeric(native.number(args[index])));
            else if(type==3) {
                const auto* bytes=native.text(args[index]);
                if(!bytes) throw VbaRuntimeError(7,"SQL color string unavailable");
                values.push_back(VbaValue::string(std::string(reinterpret_cast<const char*>(bytes),native.bytes(args[index]))));
            } else throw VbaRuntimeError(13,"SQL color function requires scalar arguments");
        }
        native.integerResult(context,static_cast<qlonglong>((indexed ? vbaQbColor(values) : vbaRgb(values)).asNumber()));
    } catch(const VbaRuntimeError& failure) { functionError(context,failure); }
      catch(...) { native.error(context,"SQL color calculation failed",-1); }
}
void rgbColor(sqlite3_context* context,int count,sqlite3_value** args) noexcept { colorFunction(context,count,args,false); }
void qbColor(sqlite3_context* context,int count,sqlite3_value** args) noexcept { colorFunction(context,count,args,true); }

// Scalar SQL adapter: functions run in the native VBA expression engine without
// a form, database-object or UI context. SQL data are bound values, never code.
class ScalarStringHost final : public VbaHost {
public:
    std::optional<std::string> field(std::string_view) override { throw VbaRuntimeError(3251,"SQL scalar has no form context"); }
    void setField(std::string_view,std::optional<std::string>) override { throw VbaRuntimeError(3251,"SQL scalar cannot assign fields"); }
    void message(std::string_view) override { throw VbaRuntimeError(3251,"SQL scalar has no UI context"); }
    void saveRecord() override { throw VbaRuntimeError(3251,"SQL scalar cannot save forms"); }
};
void scalarBuiltinFunction(sqlite3_context* context,int count,sqlite3_value** args,std::string_view name,int minimum,int maximum,unsigned dateArguments=0) noexcept {
    auto& native=api();
    try {
        if(count<minimum || count>maximum) throw VbaRuntimeError(450,"Invalid scalar SQL arity");
        VbaVariables values;
        std::string expression=dateArguments ? std::string{} : std::string(name)+"(";
        for(int index=0;index<count;++index) {
            const auto name="_la_arg"+std::to_string(index);
            VbaValue value;
            const auto type=native.type(args[index]);
            if(type==1) value=VbaValue::numeric(static_cast<double>(native.integer(args[index])));
            else if(type==2) value=VbaValue::numeric(native.number(args[index]));
            else if(type==3) {
                const auto* bytes=native.text(args[index]); const auto size=native.bytes(args[index]);
                if(!bytes && size) throw VbaRuntimeError(13,"SQL text conversion failed");
                value=VbaValue::string(bytes ? std::string(reinterpret_cast<const char*>(bytes),size) : std::string{});
            } else if(type!=5) throw VbaRuntimeError(13,"Binary scalar string input is unsupported");
            if(dateArguments && (dateArguments & (1U<<index)) && value.kind!=VbaValue::Kind::Null) value=vbaDateValue(value);
            values.emplace(name,std::move(value));
            if(index) expression+=',';
            expression+=name;
        }
        if(dateArguments) expression="_la_arg0"+std::string(name)+"_la_arg1";
        else expression+=')';
        ScalarStringHost host;
        const auto result=evaluateVbaExpression(expression,host,values);
        if(dateArguments && (name=="+" || dateArguments!=3) && result.kind==VbaValue::Kind::Number)
            throw VbaRuntimeError(6,"Date SQL projection exceeds its declared Date range");
        if(result.kind==VbaValue::Kind::Null) native.nullResult(context);
        else if(result.kind==VbaValue::Kind::Empty) native.integerResult(context,0);
        else if(result.kind==VbaValue::Kind::Boolean) native.integerResult(context,result.boolean ? -1 : 0);
        else if(result.kind==VbaValue::Kind::Date) {
            const auto text=sqlFieldValue(result,"DATETIME").toString().toUtf8();
            native.textResult(context,text.data(),static_cast<int>(text.size()),reinterpret_cast<void (*)(void*)>(-1));
        }
        else if(result.kind==VbaValue::Kind::Currency) native.numberResult(context,currencyToRealStorage(result));
        else if(result.kind==VbaValue::Kind::Text) native.textResult(context,result.text.data(),static_cast<int>(result.text.size()),reinterpret_cast<void (*)(void*)>(-1));
        else if(result.kind==VbaValue::Kind::Number) {
            if(result.numericType==VbaValue::NumericType::Double || result.numericType==VbaValue::NumericType::Single)
                native.numberResult(context,result.number);
            else native.integerResult(context,static_cast<qlonglong>(result.number));
        }
        else throw VbaRuntimeError(13,"Unsupported scalar SQL result");
    } catch(const VbaRuntimeError& failure) { functionError(context,failure); }
      catch(...) { native.error(context,"SQL scalar evaluation failed",-1); }
}
void stringReverse(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"StrReverse",1,1); }
void stringInStrRev(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"InStrRev",2,4); }
void mathSqr(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Sqr",1,1); }
void mathExp(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Exp",1,1); }
void mathLog(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Log",1,1); }
void mathSin(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Sin",1,1); }
void mathCos(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Cos",1,1); }
void mathTan(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Tan",1,1); }
void mathAtn(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Atn",1,1); }
void stringLCase(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"LCase",1,1); }
void stringUCase(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"UCase",1,1); }
template<unsigned mask,bool subtract>
void calendarDateArithmetic(sqlite3_context* context,int count,sqlite3_value** args) noexcept {
    scalarBuiltinFunction(context,count,args,subtract ? "-" : "+",2,2,mask);
}
void calendarDateValue(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"DateValue",1,1); }
void calendarTimeValue(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"TimeValue",1,1); }
void calendarDateAdd(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"DateAdd",3,3); }
void unsupportedDateSumStep(sqlite3_context*,int,sqlite3_value**) noexcept {}
void unsupportedDateSumFinal(sqlite3_context* context) noexcept { functionError(context,VbaRuntimeError(3251,"Date-valued SUM is not implemented")); }
void calendarDateSerial(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"DateSerial",3,3); }
void calendarTimeSerial(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"TimeSerial",3,3); }
void calendarCDate(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"CDate",1,1); }
void calendarDatePart(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"DatePart",2,4); }
void scalarChoose(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Choose",2,1000); }
void scalarSwitch(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Switch",2,1000); }
void stringVal(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Val",1,1); }
void stringHex(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Hex",1,1); }
void stringOct(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Oct",1,1); }
void stringSpace(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Space",1,1); }
void stringRepeat(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"String",2,2); }
void stringAsc(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Asc",1,1); }
void stringAscW(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"AscW",1,1); }
void stringChr(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Chr",1,1); }
void stringChrW(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"ChrW",1,1); }
void calendarFunction(sqlite3_context* context,int count,sqlite3_value** args,std::string_view name) noexcept {
    auto& native=api();
    try {
        std::vector<VbaValue> values;
        for(int i=0;i<count;++i) {
            const auto type=native.type(args[i]);
            if(type==SQLITE_NULL) values.emplace_back();
            else if(type==SQLITE_INTEGER || type==SQLITE_FLOAT) values.push_back(VbaValue::numeric(native.number(args[i])));
            else if(type==SQLITE_TEXT) {
                const auto* text=native.text(args[i]); if(!text) throw VbaRuntimeError(7,"SQL date text unavailable");
                values.push_back(VbaValue::string(std::string(reinterpret_cast<const char*>(text),native.bytes(args[i]))));
            } else throw VbaRuntimeError(13,"SQL calendar argument must be scalar");
        }
        const auto result=vbaCalendarFunction(name,values);
        if(result.kind==VbaValue::Kind::Null) native.nullResult(context);
        else native.integerResult(context,static_cast<qlonglong>(result.number));
    } catch(const VbaRuntimeError& failure) { functionError(context,failure); }
      catch(...) { native.error(context,"SQL calendar calculation failed",-1); }
}
void calendarDatediff(sqlite3_context* c,int n,sqlite3_value** a) noexcept { calendarFunction(c,n,a,"datediff"); }
void calendarYear(sqlite3_context* c,int n,sqlite3_value** a) noexcept { calendarFunction(c,n,a,"year"); }
void calendarMonth(sqlite3_context* c,int n,sqlite3_value** a) noexcept { calendarFunction(c,n,a,"month"); }
void calendarDay(sqlite3_context* c,int n,sqlite3_value** a) noexcept { calendarFunction(c,n,a,"day"); }
void calendarHour(sqlite3_context* c,int n,sqlite3_value** a) noexcept { calendarFunction(c,n,a,"hour"); }
void calendarMinute(sqlite3_context* c,int n,sqlite3_value** a) noexcept { calendarFunction(c,n,a,"minute"); }
void calendarSecond(sqlite3_context* c,int n,sqlite3_value** a) noexcept { calendarFunction(c,n,a,"second"); }
void calendarWeekday(sqlite3_context* c,int n,sqlite3_value** a) noexcept { calendarFunction(c,n,a,"weekday"); }

void currencyNumber(sqlite3_context* context,int count,sqlite3_value** args) noexcept {
    auto& native=api();
    try {
        if(count!=1) throw VbaRuntimeError(450,"CCur requires one argument");
        VbaValue value;
        const auto type=native.type(args[0]);
        if(type==1) {
            const auto integer=native.integer(args[0]);
            if(integer < -9007199254740992LL || integer>9007199254740992LL)
                throw VbaRuntimeError(3251,"Exact SQL integer Currency input is unavailable");
            value=VbaValue::numeric(static_cast<double>(integer));
        } else if(type==2) value=VbaValue::numeric(native.number(args[0]));
        else if(type==3) {
            const auto* bytes=native.text(args[0]);
            if(!bytes) throw VbaRuntimeError(7,"SQL Currency string unavailable");
            value=VbaValue::string(std::string(reinterpret_cast<const char*>(bytes),native.bytes(args[0])));
        } else if(type!=5) throw VbaRuntimeError(13,"CCur requires a scalar argument");
        native.numberResult(context,currencyToRealStorage(value));
    } catch(const VbaRuntimeError& failure) { functionError(context,failure); }
      catch(...) { native.error(context,"SQL currency conversion failed",-1); }
}

void currencyStore(sqlite3_context* context,int count,sqlite3_value** args) noexcept {
    if(count==1 && api().type(args[0])==SQLITE_NULL) { api().nullResult(context); return; }
    currencyNumber(context,count,args);
}

void fieldStore(sqlite3_context* context,int count,sqlite3_value** args) noexcept {
    auto& native=api();
    try {
        if(count!=2 || native.type(args[0])!=SQLITE_TEXT) throw VbaRuntimeError(450,"Field conversion requires type and value");
        const auto type=scalarText(args[0]);
        static const QSet<QString> supported{"INTEGER1","INTEGER2","INTEGER4","BOOL","BOOLEAN","DATE","DATETIME"};
        if(!supported.contains(type)) throw VbaRuntimeError(3251,"SQL field target type unsupported");
        if(native.type(args[1])==SQLITE_NULL) { native.nullResult(context); return; }
        VbaValue input;
        if(native.type(args[1])==SQLITE_INTEGER) input=VbaValue::numeric(static_cast<double>(native.integer(args[1])));
        else if(native.type(args[1])==SQLITE_FLOAT) input=VbaValue::numeric(native.number(args[1]));
        else if(native.type(args[1])==SQLITE_TEXT) {
            const auto* bytes=native.text(args[1]); if(!bytes) throw VbaRuntimeError(7,"SQL field text unavailable");
            input=VbaValue::string(std::string(reinterpret_cast<const char*>(bytes),native.bytes(args[1])));
        } else throw VbaRuntimeError(13,"SQL field input must be scalar");
        const auto result=sqlFieldValue(input,type);
        if(result.isNull()) native.nullResult(context);
        else if(result.metaType().id()==QMetaType::QString) returnText(context,result.toString());
        else native.integerResult(context,result.toLongLong());
    } catch(const VbaRuntimeError& failure) { functionError(context,failure); }
      catch(...) { native.error(context,"SQL field conversion failed",-1); }
}

void currencyArithmetic(sqlite3_context* context,sqlite3_value** args,std::string_view operation) noexcept {
    auto& native=api();
    try {
        if(native.type(args[0])==SQLITE_NULL || native.type(args[1])==SQLITE_NULL) { native.nullResult(context); return; }
        for(int i=0;i<2;++i) if(native.type(args[i])!=SQLITE_INTEGER && native.type(args[i])!=SQLITE_FLOAT) throw VbaRuntimeError(13,"Currency arithmetic requires numeric encoding");
        const auto left=currencyFromRealStorage(native.number(args[0])),right=currencyFromRealStorage(native.number(args[1]));
        native.numberResult(context,currencyToRealStorage(currencyBinary(operation,left,right)));
    } catch(const VbaRuntimeError& failure) { functionError(context,failure); }
      catch(...) { native.error(context,"SQL Currency arithmetic failed",-1); }
}
void currencyAdd(sqlite3_context* context,int,sqlite3_value** args) noexcept { currencyArithmetic(context,args,"+"); }
void currencySubtract(sqlite3_context* context,int,sqlite3_value** args) noexcept { currencyArithmetic(context,args,"-"); }
void currencyMultiply(sqlite3_context* context,int,sqlite3_value** args) noexcept { currencyArithmetic(context,args,"*"); }

struct CurrencySumState { std::int64_t units; int error; bool seen; };
void currencySumStep(sqlite3_context* context,int,sqlite3_value** args) noexcept {
    auto& native=api(); auto* state=static_cast<CurrencySumState*>(native.aggregate(context,sizeof(CurrencySumState)));
    if(!state) { native.error(context,"LinAccess SQL function error 7",-1); return; }
    if(state->error || native.type(args[0])==SQLITE_NULL) return;
    try {
        const auto type=native.type(args[0]);
        if(type!=SQLITE_INTEGER && type!=SQLITE_FLOAT) throw VbaRuntimeError(13,"Currency sum requires numeric encoding");
        const auto value=currencyFromRealStorage(native.number(args[0]));
        state->units=currencyBinary("+",VbaValue::currency(state->units),value).currencyUnits;
        state->seen=true;
    } catch(const VbaRuntimeError& failure) { state->error=failure.number; functionError(context,failure); }
      catch(...) { state->error=7; native.error(context,"LinAccess SQL function error 7",-1); }
}
void currencySumFinal(sqlite3_context* context) noexcept {
    auto& native=api(); const auto* state=static_cast<CurrencySumState*>(native.aggregate(context,0));
    try {
        if(state && state->error) throw VbaRuntimeError(state->error,"Currency sum failed");
        if(!state || !state->seen) { native.nullResult(context); return; }
        native.numberResult(context,currencyToRealStorage(VbaValue::currency(state->units)));
    } catch(const VbaRuntimeError& failure) { functionError(context,failure); }
      catch(...) { native.error(context,"SQL Currency sum failed",-1); }
}

void roundNumber(sqlite3_context* context,int count,sqlite3_value** args) noexcept {
    auto& native=api();
    try {
        if(count<1 || count>2) throw VbaRuntimeError(450,"Round requires one or two arguments");
        std::vector<VbaValue> values;
        for(int index=0;index<count;++index) {
            const auto type=native.type(args[index]);
            if(type==5) values.emplace_back();
            else if(type==1) {
                const auto integer=native.integer(args[index]);
                if(index==0 && (integer < -9007199254740992LL || integer>9007199254740992LL))
                    throw VbaRuntimeError(3251,"Exact SQL integer Round input is unavailable");
                values.push_back(VbaValue::numeric(static_cast<double>(integer)));
            } else if(type==2) values.push_back(VbaValue::numeric(native.number(args[index])));
            else if(type==3) {
                const auto* bytes=native.text(args[index]);
                if(!bytes) throw VbaRuntimeError(7,"SQL Round string unavailable");
                values.push_back(VbaValue::string(std::string(reinterpret_cast<const char*>(bytes),native.bytes(args[index]))));
            } else throw VbaRuntimeError(13,"Round requires scalar arguments");
        }
        const auto result=vbaRound(values);
        if(result.kind==VbaValue::Kind::Null) native.nullResult(context);
        else if(result.kind==VbaValue::Kind::Number) native.numberResult(context,result.number);
        else throw VbaRuntimeError(13,"SQL Round result type unavailable");
    } catch(const VbaRuntimeError& failure) { functionError(context,failure); }
      catch(...) { native.error(context,"SQL rounding failed",-1); }
}

void textLength(sqlite3_context* context,int,sqlite3_value** args) noexcept {
    auto& native=api();
    try {
        if (native.type(args[0])==SQLITE_NULL) { native.nullResult(context); return; }
        if (native.type(args[0])==SQLITE_TEXT) {
            const auto* bytes=native.text(args[0]); const auto size=native.bytes(args[0]);
            if (!bytes) throw VbaRuntimeError(7,"SQL string unavailable");
            native.integerResult(context,vbaStringCodeUnits(std::string_view(reinterpret_cast<const char*>(bytes),size)));
        } else {
            const auto text=scalarText(args[0]).toUtf8();
            native.integerResult(context,vbaStringCodeUnits(std::string_view(text.constData(),text.size())));
        }
    } catch(const VbaRuntimeError& failure) { functionError(context,failure); }
      catch(...) { native.error(context,"SQL Len evaluation failed",-1); }
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
struct SingleState { sqlite3_value* value; int error; bool seen; };
bool identicalUpdateValue(sqlite3_value* left,sqlite3_value* right) noexcept {
    auto& native=api(); const auto type=native.type(left);
    if (type!=native.type(right)) return false;
    if (type==SQLITE_NULL) return true;
    if (type==SQLITE_INTEGER) return native.integer(left)==native.integer(right);
    if (type==SQLITE_FLOAT) return native.number(left)==native.number(right);
    if (type==SQLITE_TEXT || type==SQLITE_BLOB) {
        const auto* a=type==SQLITE_TEXT ? static_cast<const void*>(native.text(left)) : sqlite3_value_blob(left);
        const auto* b=type==SQLITE_TEXT ? static_cast<const void*>(native.text(right)) : sqlite3_value_blob(right);
        const auto count=native.bytes(left);
        return count==native.bytes(right) && (!count || (a && b && std::memcmp(a,b,count)==0));
    }
    return false;
}
void singleStep(sqlite3_context* context,int,sqlite3_value** args) noexcept {
    auto& native=api(); auto* state=static_cast<SingleState*>(native.aggregate(context,sizeof(SingleState)));
    if (!state) { native.error(context,"LinAccess SQL function error 7",-1); return; }
    if (state->error) return;
    if (state->seen) { if (!identicalUpdateValue(state->value,args[0])) state->error=3251; }
    else { state->seen=true; state->value=native.duplicate(args[0]); if (!state->value) state->error=7; }
    if (state->error) native.error(context,state->error==7 ? "LinAccess SQL function error 7" : "LinAccess SQL function error 3251",-1);
}
void singleFinal(sqlite3_context* context) noexcept {
    auto& native=api(); auto* state=static_cast<SingleState*>(native.aggregate(context,0));
    if (!state) { native.nullResult(context); return; }
    if (state->error) native.error(context,state->error==7 ? "LinAccess SQL function error 7" : "LinAccess SQL function error 3251",-1);
    else if (state->value) native.result(context,state->value);
    else native.nullResult(context);
    if (state->value) native.freeValue(state->value);
    state->value=nullptr;
}
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
    if (number==5 || number==6 || number==7 || number==13 || number==94 || number==438 || number==450 || number==3035 || number==3146 || number==3061 || number==2450 || number==2451 || number==2455 || number==2465 || number==3251 || number==3085) return number;
    return {};
}
void installJetSqlFunctions(QSqlDatabase db) {
    // Keep QSQLITE and all native handle calls on the exact pinned backend.
    // Probe through Qt before passing its opaque handle to a native API.
    QSqlQuery backend(db);
    if (!backend.exec("SELECT sqlite_source_id()") || !backend.next()
        || backend.value(0).toString()!=QString::fromLatin1(SQLITE_SOURCE_ID)
        || QByteArray(sqlite3_sourceid())!=QByteArray(SQLITE_SOURCE_ID))
        throw VbaRuntimeError(3251,"SQLite backend identity mismatch");
    backend.finish();
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
    ensure("linaccess_next_counter",2,&nextAccessCounter,nullptr,nullptr);
    ensure("linaccess_jet_iif",3,&iif,nullptr,nullptr);
    ensure("linaccess_jet_rgb",-1,&rgbColor,nullptr,nullptr);
    ensure("linaccess_jet_qbcolor",-1,&qbColor,nullptr,nullptr);
    ensure("linaccess_jet_round",-1,&roundNumber,nullptr,nullptr);
    ensure("linaccess_jet_asc",-1,&stringAsc,nullptr,nullptr);
    ensure("linaccess_jet_ascw",-1,&stringAscW,nullptr,nullptr);
    ensure("linaccess_jet_chr",-1,&stringChr,nullptr,nullptr);
    ensure("linaccess_jet_sqr",-1,&mathSqr,nullptr,nullptr);
    ensure("linaccess_jet_exp",-1,&mathExp,nullptr,nullptr);
    ensure("linaccess_jet_log",-1,&mathLog,nullptr,nullptr);
    ensure("linaccess_jet_sin",-1,&mathSin,nullptr,nullptr);
    ensure("linaccess_jet_cos",-1,&mathCos,nullptr,nullptr);
    ensure("linaccess_jet_tan",-1,&mathTan,nullptr,nullptr);
    ensure("linaccess_jet_atn",-1,&mathAtn,nullptr,nullptr);
    ensure("linaccess_jet_lcase",-1,&stringLCase,nullptr,nullptr);
    ensure("linaccess_jet_ucase",-1,&stringUCase,nullptr,nullptr);
    ensure("linaccess_jet_date_add_1",-1,&calendarDateArithmetic<1,false>,nullptr,nullptr);
    ensure("linaccess_jet_date_add_2",-1,&calendarDateArithmetic<2,false>,nullptr,nullptr);
    ensure("linaccess_jet_date_add_3",-1,&calendarDateArithmetic<3,false>,nullptr,nullptr);
    ensure("linaccess_jet_date_subtract_1",-1,&calendarDateArithmetic<1,true>,nullptr,nullptr);
    ensure("linaccess_jet_date_subtract_2",-1,&calendarDateArithmetic<2,true>,nullptr,nullptr);
    ensure("linaccess_jet_date_subtract_3",-1,&calendarDateArithmetic<3,true>,nullptr,nullptr);
    ensure("linaccess_jet_datevalue",-1,&calendarDateValue,nullptr,nullptr);
    ensure("linaccess_jet_timevalue",-1,&calendarTimeValue,nullptr,nullptr);
    ensure("linaccess_jet_dateadd",-1,&calendarDateAdd,nullptr,nullptr);
    ensure("linaccess_jet_date_sum",1,nullptr,&unsupportedDateSumStep,&unsupportedDateSumFinal);
    ensure("linaccess_jet_dateserial",-1,&calendarDateSerial,nullptr,nullptr);
    ensure("linaccess_jet_timeserial",-1,&calendarTimeSerial,nullptr,nullptr);
    ensure("linaccess_jet_cdate",-1,&calendarCDate,nullptr,nullptr);
    ensure("linaccess_jet_datepart",-1,&calendarDatePart,nullptr,nullptr);
    ensure("linaccess_jet_choose",-1,&scalarChoose,nullptr,nullptr);
    ensure("linaccess_jet_switch",-1,&scalarSwitch,nullptr,nullptr);
    ensure("linaccess_jet_val",-1,&stringVal,nullptr,nullptr);
    ensure("linaccess_jet_hex",-1,&stringHex,nullptr,nullptr);
    ensure("linaccess_jet_oct",-1,&stringOct,nullptr,nullptr);
    ensure("linaccess_jet_space",-1,&stringSpace,nullptr,nullptr);
    ensure("linaccess_jet_string",-1,&stringRepeat,nullptr,nullptr);
    ensure("linaccess_jet_chrw",-1,&stringChrW,nullptr,nullptr);
    ensure("linaccess_jet_strreverse",-1,&stringReverse,nullptr,nullptr);
    ensure("linaccess_jet_instrrev",-1,&stringInStrRev,nullptr,nullptr);
    ensure("linaccess_jet_datediff",-1,&calendarDatediff,nullptr,nullptr);
    ensure("linaccess_jet_year",-1,&calendarYear,nullptr,nullptr);
    ensure("linaccess_jet_month",-1,&calendarMonth,nullptr,nullptr);
    ensure("linaccess_jet_day",-1,&calendarDay,nullptr,nullptr);
    ensure("linaccess_jet_hour",-1,&calendarHour,nullptr,nullptr);
    ensure("linaccess_jet_minute",-1,&calendarMinute,nullptr,nullptr);
    ensure("linaccess_jet_second",-1,&calendarSecond,nullptr,nullptr);
    ensure("linaccess_jet_weekday",-1,&calendarWeekday,nullptr,nullptr);
    ensure("linaccess_jet_ccur",-1,&currencyNumber,nullptr,nullptr);
    ensure("linaccess_jet_currency_store",1,&currencyStore,nullptr,nullptr);
    ensure("linaccess_jet_field_store",2,&fieldStore,nullptr,nullptr);
    ensure("linaccess_jet_currency_sum",1,nullptr,&currencySumStep,&currencySumFinal);
    ensure("linaccess_jet_currency_add",2,&currencyAdd,nullptr,nullptr);
    ensure("linaccess_jet_currency_subtract",2,&currencySubtract,nullptr,nullptr);
    ensure("linaccess_jet_currency_multiply",2,&currencyMultiply,nullptr,nullptr);
    ensure("linaccess_jet_first",1,nullptr,&firstStep,&aggregateFinal);
    ensure("linaccess_jet_last",1,nullptr,&lastStep,&aggregateFinal);
    ensure("linaccess_jet_update_value",1,nullptr,&singleStep,&singleFinal);
    ensure("linaccess_jet_format_string",2,&formatString,nullptr,nullptr);
    ensure("linaccess_jet_format",2,&formatVariant,nullptr,nullptr);
    ensure("linaccess_jet_concat",2,&concatenate,nullptr,nullptr);
    ensure("linaccess_jet_len",1,&textLength,nullptr,nullptr);
    ensure("linaccess_jet_like_pattern",1,&likePattern,nullptr,nullptr);
    ensure("linaccess_jet_object_value",3,&objectValue,nullptr,nullptr);
}
void setJetSqlObjectResolver(QSqlDatabase db,JetSqlObjectResolver resolver) {
    installJetSqlFunctions(db);
    objectContext(db)->resolver=std::move(resolver);
}

}
