#include <QDateTime>
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
#include "native_vba_sql.hpp"
#include "native_report_aggregate.hpp"
#include <memory>
#include "vba_currency.hpp"
#include "vba_decimal.hpp"
#include "vba_qt_value.hpp"
#include <QSet>
#include <QStringList>
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
    JetSqlObjectPropertyResolver propertyResolver;
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
void objectResult(sqlite3_context* context,const VbaValue& result) {
    auto& native=api();
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
        objectResult(context,result);
    } catch (const VbaRuntimeError& error) { functionError(context,error); }
      catch (...) { native.error(context,"SQL object resolution failed",-1); }
}
void objectProperty(sqlite3_context* context,int,sqlite3_value** args) noexcept {
    auto& native=api();try {
        auto* objects=static_cast<ObjectContext*>(native.userData(context));if(!objects || !objects->propertyResolver) throw VbaRuntimeError(3061,"SQL object property context unavailable");
        const auto collection=scalarText(args[0]).toUtf8(),object=scalarText(args[1]).toUtf8(),field=scalarText(args[2]).toUtf8(),property=scalarText(args[3]).toUtf8();
        const auto view=[](const QByteArray& value) {return std::string_view(value.constData(),value.size());};const std::optional<std::string_view> target=native.type(args[2])==5 ? std::nullopt : std::optional<std::string_view>(view(field));
        objectResult(context,objects->propertyResolver(view(collection),view(object),target,view(property)));
    }catch(const VbaRuntimeError& error) {functionError(context,error);}catch(...) {native.error(context,"SQL object property resolution failed",-1);}
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
    bool isNumericText(std::string_view text) override {return isAccessNumericText(text);}
    VbaValue formatValue(const VbaValue& value,std::string_view mask,bool stringResult) override {return formatAccessValue(value,mask,stringResult);}
    VbaValue calendarName(bool month,int number,bool abbreviated,int firstDay) override {return accessCalendarName(month,number,abbreviated,firstDay);}
    VbaValue formatNumericValue(const VbaValue& value,bool currency,const VbaNumericFormatOptions& options) override {return formatAccessNumericValue(value,currency,options);}
    std::optional<std::string> field(std::string_view) override { throw VbaRuntimeError(3251,"SQL scalar has no form context"); }
    void setField(std::string_view,std::optional<std::string>) override { throw VbaRuntimeError(3251,"SQL scalar cannot assign fields"); }
    void message(std::string_view) override { throw VbaRuntimeError(3251,"SQL scalar has no UI context"); }
    void saveRecord() override { throw VbaRuntimeError(3251,"SQL scalar cannot save forms"); }
};
void scalarBuiltinFunction(sqlite3_context* context,int count,sqlite3_value** args,std::string_view name,int minimum,int maximum,unsigned dateArguments=0,bool infix=false) noexcept {
    auto& native=api();
    try {
        if(count<minimum || count>maximum) throw VbaRuntimeError(450,"Invalid scalar SQL arity");
        const bool numericFormat=name=="FormatNumber" || name=="FormatCurrency" || name=="FormatPercent";
        VbaVariables values;
        std::string expression=(dateArguments || infix) ? std::string{} : std::string(name)+"(";
        for(int index=0;index<count;++index) {
            const auto name="_la_arg"+std::to_string(index);
            VbaValue value;
            const auto type=native.type(args[index]);
            if(type==1) value=numericFormat ? decimalLiteral(std::to_string(native.integer(args[index]))) : VbaValue::numeric(static_cast<double>(native.integer(args[index])));
            else if(type==2) value=VbaValue::numeric(native.number(args[index]));
            else if(type==3) {
                const auto* bytes=native.text(args[index]); const auto size=native.bytes(args[index]);
                if(!bytes && size) throw VbaRuntimeError(13,"SQL text conversion failed");
                value=VbaValue::string(bytes ? std::string(reinterpret_cast<const char*>(bytes),size) : std::string{});
            } else if(type!=5) throw VbaRuntimeError(13,"Binary scalar string input is unsupported");
            if(dateArguments && (dateArguments & (1U<<index)) && value.kind!=VbaValue::Kind::Null) value=vbaDateValue(value);
            if(infix && value.kind==VbaValue::Kind::Number && !std::isfinite(value.number)) throw VbaRuntimeError(6,"Nonfinite arithmetic operand");
            values.emplace(name,std::move(value));
            if(index) expression+=',';
            expression+=name;
        }
        if(dateArguments || infix) expression="_la_arg0"+std::string(name)+"_la_arg1";
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
void double_add(sqlite3_context* context,int count,sqlite3_value** args) noexcept {scalarBuiltinFunction(context,count,args,"+",2,2,0,true);}
void double_subtract(sqlite3_context* context,int count,sqlite3_value** args) noexcept {scalarBuiltinFunction(context,count,args,"-",2,2,0,true);}
void double_multiply(sqlite3_context* context,int count,sqlite3_value** args) noexcept {scalarBuiltinFunction(context,count,args,"*",2,2,0,true);}
void double_divide(sqlite3_context* context,int count,sqlite3_value** args) noexcept {scalarBuiltinFunction(context,count,args,"/",2,2,0,true);}
void numericPower(sqlite3_context* context,int count,sqlite3_value** args) noexcept {scalarBuiltinFunction(context,count,args,"^",2,2,0,true);}
void boundScalarSqlFunction(sqlite3_context* context,int count,sqlite3_value** args,std::string_view builtin,int minimum,int maximum,bool typed,bool dateComparison=false) noexcept {
    auto& native=api();
    try {
        const int offset=typed ? 1 : 0,arguments=count-offset;
        if(arguments<minimum || arguments>maximum) throw VbaRuntimeError(450,"Invalid bound scalar SQL arity");
        QStringList declared;
        if(typed) {
            if(native.type(args[0])!=SQLITE_TEXT) throw VbaRuntimeError(450,"Typed scalar SQL requires argument types");
            declared=scalarText(args[0]).split(',');
            if(declared.size()!=arguments) throw VbaRuntimeError(450,"Typed scalar SQL argument types mismatch");
            static const QSet<QString> supported{"","CURRENCY_REAL","DATETIME","SINGLE_REAL","INTEGER1","INTEGER2","INTEGER4","BOOL","BOOLEAN","DOUBLE"};
            for(const auto& type:declared) if(!supported.contains(type)) throw VbaRuntimeError(3251,"Scalar SQL source type is unsupported");
        }
        if(dateComparison && declared!=QStringList{"DATETIME","DATETIME"}) throw VbaRuntimeError(450,"Date comparison requires two proven Date operands");
        if(builtin=="Format" || builtin=="Format$") {
            if(native.type(args[offset+1])==SQLITE_NULL) throw VbaRuntimeError(94,"Format mask is Null");
            if(native.type(args[offset+1])!=SQLITE_TEXT) throw VbaRuntimeError(13,"Format mask requires text");
        }
        VbaVariables values;std::string expression=dateComparison ? std::string{} : std::string(builtin)+"(";
        for(int index=0;index<arguments;++index) {
            auto* argument=args[index+offset];QVariant storage;
            switch(native.type(argument)) {
            case SQLITE_NULL: break;
            case SQLITE_INTEGER: storage=QVariant::fromValue<qlonglong>(native.integer(argument));break;
            case SQLITE_FLOAT: storage=native.number(argument);break;
            case SQLITE_TEXT: {
                const auto* bytes=native.text(argument);const auto size=native.bytes(argument);
                if(!bytes && size) throw VbaRuntimeError(7,"Scalar SQL text unavailable");
                const QByteArray input(bytes ? reinterpret_cast<const char*>(bytes) : "",size);
                if(!input.isValidUtf8()) throw VbaRuntimeError(13,"Scalar SQL text is not valid UTF-8");
                storage=QString::fromUtf8(input);break;
            }
            default: throw VbaRuntimeError(13,"Scalar SQL binary input is unsupported");
            }
            const auto name="_la_arg"+std::to_string(index);
            // Formatting already preserves unproven SQLite int64 input exactly.
            // A proven source declaration still goes through checked field decoding.
            const bool exactFormatInteger=index==0 && builtin.starts_with("Format") && declared.value(index).isEmpty() && native.type(argument)==SQLITE_INTEGER;
            values.emplace(name,exactFormatInteger ? decimalLiteral(std::to_string(native.integer(argument))) : qtFieldValue(storage,declared.value(index)));
            if(index) expression+=dateComparison ? std::string(builtin) : ",";expression+=name;
        }
        if(!dateComparison) expression+=')';ScalarStringHost host;
        const auto result=evaluateVbaExpression(expression,host,values);
        if(result.kind==VbaValue::Kind::Null) native.nullResult(context);
        else if(result.kind==VbaValue::Kind::Boolean) native.integerResult(context,result.boolean ? -1 : 0);
        else if(result.kind==VbaValue::Kind::Number && (result.numericType==VbaValue::NumericType::Double || result.numericType==VbaValue::NumericType::Single)) {
            if(!std::isfinite(result.number)) throw VbaRuntimeError(6,"Nonfinite bound scalar SQL result");
            native.numberResult(context,result.number);
        }
        else if(result.kind==VbaValue::Kind::Currency) native.numberResult(context,currencyToRealStorage(result));
        else if(result.kind==VbaValue::Kind::Date) returnText(context,sqlFieldValue(result,"DATETIME").toString());
        else if(result.kind==VbaValue::Kind::Text) returnText(context,QString::fromUtf8(result.text.data(),static_cast<qsizetype>(result.text.size())));
        else if(result.kind==VbaValue::Kind::Number && (result.numericType==VbaValue::NumericType::Integer || result.numericType==VbaValue::NumericType::Long || result.numericType==VbaValue::NumericType::Byte)) native.integerResult(context,static_cast<qlonglong>(result.number));
        else throw VbaRuntimeError(13,"Scalar SQL result type mismatch");
    } catch(const VbaRuntimeError& failure) { functionError(context,failure); }
      catch(...) { native.error(context,"SQL bound scalar evaluation failed",-1); }
}
void bound_rgb(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"RGB",3,3,true);}
void bound_qbcolor(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"QBColor",1,1,true);}
void bound_partition(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Partition",4,4,true);}
void normalized_abs_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Abs",1,1,true);}
void normalized_abs(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Abs",1,1,false);}
void normalized_int_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Int",1,1,true);}
void normalized_int(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Int",1,1,false);}
void normalized_fix_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Fix",1,1,true);}
void normalized_fix(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Fix",1,1,false);}
void normalized_sgn_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Sgn",1,1,true);}
void normalized_sgn(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Sgn",1,1,false);}
void normalized_round_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Round",1,2,true);}
void normalized_ccur_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"CCur",1,1,true);}
void inspect_isnumeric(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"IsNumeric",1,1,false);}
void inspect_isnumeric_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"IsNumeric",1,1,true);}
void inspect_vartype(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"VarType",1,1,false);}
void inspect_vartype_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"VarType",1,1,true);}
void inspect_typename(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"TypeName",1,1,false);}
void inspect_typename_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"TypeName",1,1,true);}
void stringCStr(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"CStr",1,1,false);}
void stringCStrTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"CStr",1,1,true);}
void stringStr(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Str",1,1,false);}
void stringStrTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Str",1,1,true);}
void stringLeft(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Left",2,2,false);}
void stringLeftTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Left",2,2,true);}
void stringLeftString(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Left$",2,2,false);}
void stringLeftStringTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Left$",2,2,true);}
void stringRight(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Right",2,2,false);}
void stringRightTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Right",2,2,true);}
void stringRightString(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Right$",2,2,false);}
void stringRightStringTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Right$",2,2,true);}
void stringMid(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Mid",2,3,false);}
void stringMidTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Mid",2,3,true);}
void stringMidString(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Mid$",2,3,false);}
void stringMidStringTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Mid$",2,3,true);}
void stringReplace(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Replace",3,6,false);}
void stringReplaceTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Replace",3,6,true);}
void stringInStr(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"InStr",2,4,false);}
void stringInStrTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"InStr",2,4,true);}
void stringStrComp(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"StrComp",2,3,false);}
void stringStrCompTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"StrComp",2,3,true);}
void stringTrim(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Trim",1,1,false);}
void stringTrimTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Trim",1,1,true);}
void stringTrimString(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Trim$",1,1,false);}
void stringTrimStringTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Trim$",1,1,true);}
void stringLTrim(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"LTrim",1,1,false);}
void stringLTrimTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"LTrim",1,1,true);}
void stringLTrimString(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"LTrim$",1,1,false);}
void stringLTrimStringTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"LTrim$",1,1,true);}
void stringRTrim(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"RTrim",1,1,false);}
void stringRTrimTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"RTrim",1,1,true);}
void stringRTrimString(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"RTrim$",1,1,false);}
void stringRTrimStringTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"RTrim$",1,1,true);}
void stringLCaseString(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"LCase$",1,1,false);}
void stringLCaseStringTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"LCase$",1,1,true);}
void stringUCaseString(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"UCase$",1,1,false);}
void stringUCaseStringTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"UCase$",1,1,true);}
void convert_cbool_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"CBool",1,1,true);}
void convert_cbool(sqlite3_context* context,int count,sqlite3_value** args) noexcept {scalarBuiltinFunction(context,count,args,"CBool",1,1);}
void convert_cbyte_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"CByte",1,1,true);}
void convert_cbyte(sqlite3_context* context,int count,sqlite3_value** args) noexcept {scalarBuiltinFunction(context,count,args,"CByte",1,1);}
void convert_cint_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"CInt",1,1,true);}
void convert_cint(sqlite3_context* context,int count,sqlite3_value** args) noexcept {scalarBuiltinFunction(context,count,args,"CInt",1,1);}
void convert_clng_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"CLng",1,1,true);}
void convert_clng(sqlite3_context* context,int count,sqlite3_value** args) noexcept {scalarBuiltinFunction(context,count,args,"CLng",1,1);}
void convert_csng_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"CSng",1,1,true);}
void convert_csng(sqlite3_context* context,int count,sqlite3_value** args) noexcept {scalarBuiltinFunction(context,count,args,"CSng",1,1);}
void convert_cdbl_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"CDbl",1,1,true);}
void convert_cdbl(sqlite3_context* context,int count,sqlite3_value** args) noexcept {scalarBuiltinFunction(context,count,args,"CDbl",1,1);}
void financialDdbTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"DDB",4,5,true);}
void financialDdb(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"DDB",4,5); }
void financialIpmtTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"IPmt",4,6,true);}
void financialIpmt(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"IPmt",4,6); }
void financialPpmtTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"PPmt",4,6,true);}
void financialPpmt(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"PPmt",4,6); }
void financialRateTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Rate",3,6,true);}
void financialRate(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Rate",3,6); }
void financialNperTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"NPer",3,5,true);}
void financialNper(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"NPer",3,5); }
void financialPvTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"PV",3,5,true);}
void financialPv(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"PV",3,5); }
void financialSlnTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"SLN",3,3,true);}
void financialSln(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"SLN",3,3); }
void financialSydTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"SYD",4,4,true);}
void financialSyd(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"SYD",4,4); }
void financialFvTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"FV",3,5,true);}
void financialFv(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"FV",3,5); }
void financialPmtTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Pmt",3,5,true);}
void financialPmt(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Pmt",3,5); }
void partitionRange(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Partition",4,4); }
void textLengthTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Len",1,1,true);}
void stringLenBTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"LenB",1,1,true);}
void stringLenB(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"LenB",1,1); }
void stringReverse(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"StrReverse",1,1,false);}
void stringReverseTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"StrReverse",1,1,true);}
void stringInStrRev(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"InStrRev",2,4,false);}
void stringInStrRevTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"InStrRev",2,4,true);}
void mathSqrTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Sqr",1,1,true);}
void mathSqr(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Sqr",1,1); }
void mathExpTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Exp",1,1,true);}
void mathExp(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Exp",1,1); }
void mathLogTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Log",1,1,true);}
void mathLog(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Log",1,1); }
void mathSinTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Sin",1,1,true);}
void mathSin(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Sin",1,1); }
void mathCosTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Cos",1,1,true);}
void mathCos(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Cos",1,1); }
void mathTanTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Tan",1,1,true);}
void mathTan(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Tan",1,1); }
void mathAtnTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Atn",1,1,true);}
void mathAtn(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Atn",1,1); }
void stringLCase(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"LCase",1,1,false);}
void stringLCaseTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"LCase",1,1,true);}
void stringUCase(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"UCase",1,1,false);}
void stringUCaseTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"UCase",1,1,true);}
void calendarDateValueTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"DateValue",1,1,true);}
void calendarTimeValueTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"TimeValue",1,1,true);}
void calendarDateAddTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"DateAdd",3,3,true);}
void calendarDateSerialTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"DateSerial",3,3,true);}
void calendarTimeSerialTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"TimeSerial",3,3,true);}
void calendarCDateTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"CDate",1,1,true);}
void calendarDatePartTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"DatePart",2,4,true);}
void calendarDatediffTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"DateDiff",3,5,true);}
void calendarYearTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Year",1,1,true);}
void calendarMonthTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Month",1,1,true);}
void calendarDayTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Day",1,1,true);}
void calendarHourTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Hour",1,1,true);}
void calendarMinuteTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Minute",1,1,true);}
void calendarSecondTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Second",1,1,true);}
void calendarWeekdayTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Weekday",1,2,true);}
void monthNameTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"MonthName",1,2,true);}
void weekdayNameTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"WeekdayName",1,3,true);}
void calendarCompare_equal(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"=",2,2,true,true);}
void calendarCompare_not_equal(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"<>",2,2,true,true);}
void calendarCompare_less(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"<",2,2,true,true);}
void calendarCompare_less_equal(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"<=",2,2,true,true);}
void calendarCompare_greater(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,">",2,2,true,true);}
void calendarCompare_greater_equal(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,">=",2,2,true,true);}
void currentDirectory(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"CurDir",0,1); }
void fileLength(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"FileLen",1,1); }
void fileModified(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"FileDateTime",1,1); }
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
void stringHexTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Hex",1,1,true);}
void alias_hex_string(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Hex$",1,1,false);}
void alias_hex_string_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Hex$",1,1,true);}
void alias_oct_string(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Oct$",1,1,false);}
void alias_oct_string_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Oct$",1,1,true);}
void alias_space_string(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Space$",1,1,false);}
void alias_space_string_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Space$",1,1,true);}
void alias_string_string(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"String$",2,2,false);}
void alias_string_string_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"String$",2,2,true);}
void alias_chr_string(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Chr$",1,1,false);}
void alias_chr_string_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Chr$",1,1,true);}
void alias_chrw_string(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"ChrW$",1,1,false);}
void alias_chrw_string_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"ChrW$",1,1,true);}
void stringHex(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Hex",1,1); }
void stringOctTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Oct",1,1,true);}
void stringOct(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Oct",1,1); }
void stringSpaceTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Space",1,1,true);}
void stringSpace(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Space",1,1); }
void stringRepeatTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"String",2,2,true);}
void stringRepeat(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"String",2,2); }
void stringAscTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Asc",1,1,true);}
void stringAsc(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Asc",1,1); }
void stringAscWTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"AscW",1,1,true);}
void stringAscW(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"AscW",1,1); }
void stringChrTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Chr",1,1,true);}
void stringChr(sqlite3_context* context,int count,sqlite3_value** args) noexcept { scalarBuiltinFunction(context,count,args,"Chr",1,1); }
void stringChrWTyped(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"ChrW",1,1,true);}
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
void monthName(sqlite3_context* c,int n,sqlite3_value** a) noexcept {scalarBuiltinFunction(c,n,a,"MonthName",1,2);}
void weekdayName(sqlite3_context* c,int n,sqlite3_value** a) noexcept {scalarBuiltinFunction(c,n,a,"WeekdayName",1,3);}
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
        static const QSet<QString> supported{"REAL","DOUBLE","SINGLE_REAL","INTEGER1","INTEGER2","INTEGER4","BOOL","BOOLEAN","DATE","DATETIME"};
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
        else if(result.metaType().id()==QMetaType::Double) native.numberResult(context,result.toDouble());
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

// These callbacks consume private integer ticks only; numeric SQL amounts and
// physical REAL-backed Currency continue to use their existing checked adapter.
void currencyTicksArithmetic(sqlite3_context* context,sqlite3_value** args,std::string_view operation) noexcept {
    auto& native=api();try{if(native.type(args[0])==SQLITE_NULL || native.type(args[1])==SQLITE_NULL){native.nullResult(context);return;}
        if(native.type(args[0])!=SQLITE_INTEGER || native.type(args[1])!=SQLITE_INTEGER)throw VbaRuntimeError(13,"Exact Currency arithmetic requires integer ticks");
        const auto value=currencyBinary(operation,VbaValue::currency(native.integer(args[0])),VbaValue::currency(native.integer(args[1])));native.integerResult(context,value.currencyUnits);
    }catch(const VbaRuntimeError& failure){functionError(context,failure);}catch(...){native.error(context,"Exact SQL Currency arithmetic failed",-1);}
}
void currencyTicksAdd(sqlite3_context* c,int,sqlite3_value** a) noexcept {currencyTicksArithmetic(c,a,"+");}
void currencyTicksSubtract(sqlite3_context* c,int,sqlite3_value** a) noexcept {currencyTicksArithmetic(c,a,"-");}
void currencyTicksMultiply(sqlite3_context* c,int,sqlite3_value** a) noexcept {currencyTicksArithmetic(c,a,"*");}

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

template<int Mode> void dateFieldDefault(sqlite3_context* context,int,sqlite3_value**) noexcept {
    try {const auto now=QDateTime::currentDateTime();if(!now.isValid()) throw VbaRuntimeError(3251,"Native clock unavailable");const auto date=Mode==2 ? QString("1899-12-30") : now.date().toString(Qt::ISODate);const auto time=Mode==0 ? QString("00:00:00") : now.time().toString("HH:mm:ss");returnText(context,date+" "+time);}catch(const VbaRuntimeError& failure) {functionError(context,failure);}catch(...) {api().error(context,"Native clock unavailable",-1);}
}

void nonemptyText(sqlite3_context* context,int,sqlite3_value** args) noexcept {
    auto& native=api();try {
        if(native.type(args[0])==SQLITE_NULL) {native.integerResult(context,1);return;}
        if(native.type(args[0])!=SQLITE_TEXT) throw VbaRuntimeError(13,"Text field requires text");
        if(!native.text(args[0])) throw VbaRuntimeError(7,"SQL text unavailable");
        if(native.bytes(args[0])==0) throw VbaRuntimeError(3315,"Field cannot contain a zero-length string");
        native.integerResult(context,1);
    }catch(const VbaRuntimeError& failure) {functionError(context,failure);}catch(...) {native.error(context,"Text validation failed",-1);}
}

void fieldValidation(sqlite3_context* context,int count,sqlite3_value** args) noexcept {
    auto& native=api();try {
        if(native.type(args[0])==SQLITE_NULL || native.integer(args[0])!=0) {native.integerResult(context,1);return;}
        if(count==2) {if(native.type(args[1])!=SQLITE_TEXT || native.bytes(args[1])>1020) throw VbaRuntimeError(3251,"Invalid ValidationText");const auto text=scalarText(args[1]);if(text.size()>255 || text.contains(QChar(0)) || !text.isValidUtf16()) throw VbaRuntimeError(3251,"Invalid ValidationText");if(!text.isEmpty()) {const auto message=QByteArray("LinAccess SQL function error 3317: ")+text.toUtf8();native.error(context,message.constData(),message.size());return;}}
        throw VbaRuntimeError(3317,"Field value violates its ValidationRule");
    }catch(const VbaRuntimeError& failure) {functionError(context,failure);}catch(...) {native.error(context,"Field validation failed",-1);}
}

void textCapacity(sqlite3_context* context,int,sqlite3_value** args) noexcept {
    auto& native=api();
    try {
        if(native.type(args[1])!=SQLITE_INTEGER) throw VbaRuntimeError(5,"Invalid source text capacity");
        const auto maximum=native.integer(args[1]);
        if(maximum<1 || maximum>255) throw VbaRuntimeError(5,"Invalid source text capacity");
        if(native.type(args[0])==SQLITE_NULL) { native.integerResult(context,1); return; }
        if(native.type(args[0])!=SQLITE_TEXT) throw VbaRuntimeError(13,"Source text field requires text");
        const auto* bytes=native.text(args[0]); const auto size=native.bytes(args[0]);
        if(!bytes) throw VbaRuntimeError(7,"SQL string unavailable");
        if(vbaStringCodeUnits(std::string_view(reinterpret_cast<const char*>(bytes),size))>maximum)
            throw VbaRuntimeError(3163,"Text exceeds the preserved source field capacity");
        native.integerResult(context,1);
    } catch(const VbaRuntimeError& failure) { functionError(context,failure); }
      catch(...) { native.error(context,"Source text capacity validation failed",-1); }
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
        if(type==1) input=decimalLiteral(std::to_string(native.integer(args[0])));
        else if(type==2) input=VbaValue::numeric(native.number(args[0]));
        else if (type==3) input=VbaValue::string(readText(args[0]));
        else if (type!=5) throw VbaRuntimeError(13,"Format input type unsupported");
        const auto result=formatAccessValue(input,readText(args[1]),stringResult);
        if (result.kind==VbaValue::Kind::Null) native.nullResult(context);
        else native.textResult(context,result.text.data(),static_cast<int>(result.text.size()),reinterpret_cast<void (*)(void*)>(-1));
    } catch (const VbaRuntimeError& error) { functionError(context,error); }
      catch (...) { native.error(context,"Jet date format failed",-1); }
}
void format_format_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Format",2,2,true);}
void format_format_string_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"Format$",2,2,true);}
void format_formatnumber_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"FormatNumber",1,5,true);}
void format_formatcurrency_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"FormatCurrency",1,5,true);}
void format_formatpercent_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"FormatPercent",1,5,true);}
void format_formatdatetime_typed(sqlite3_context* context,int count,sqlite3_value** args) noexcept {boundScalarSqlFunction(context,count,args,"FormatDateTime",1,2,true);}
void formatNumber(sqlite3_context* context,int count,sqlite3_value** args) noexcept {scalarBuiltinFunction(context,count,args,"FormatNumber",1,5);}
void formatCurrency(sqlite3_context* context,int count,sqlite3_value** args) noexcept {scalarBuiltinFunction(context,count,args,"FormatCurrency",1,5);}
void formatPercent(sqlite3_context* context,int count,sqlite3_value** args) noexcept {scalarBuiltinFunction(context,count,args,"FormatPercent",1,5);}
void formatDateTime(sqlite3_context* context,int count,sqlite3_value** args) noexcept {scalarBuiltinFunction(context,count,args,"FormatDateTime",1,2);}
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
// SQLite owns only the small pointer/error slot; the C++ accumulator has its
// normal alignment/lifetime and is destroyed even when a group fails.
struct StatisticalState { NativeReportAggregate* aggregate; int error; };
template<bool Deviation,bool Population=false> void statisticalStep(sqlite3_context* context,int,sqlite3_value** args) noexcept {
    auto& native=api();auto* state=static_cast<StatisticalState*>(native.aggregate(context,sizeof(StatisticalState)));
    if(!state) {native.error(context,"LinAccess SQL function error 7",-1);return;}
    if(state->error || native.type(args[0])==SQLITE_NULL) return;
    try {
        if(native.type(args[0])!=SQLITE_INTEGER && native.type(args[0])!=SQLITE_FLOAT)
            throw VbaRuntimeError(13,"Statistical SQL aggregate requires numeric values");
        if(!state->aggregate) state->aggregate=new NativeReportAggregate(Population ? (Deviation ? "stdevp" : "varp") : (Deviation ? "stdev" : "var"));
        state->aggregate->add(VbaValue::numeric(native.number(args[0])));
    } catch(const VbaRuntimeError& failure) {state->error=failure.number;functionError(context,failure);}
      catch(...) {state->error=7;native.error(context,"LinAccess SQL function error 7",-1);}
}
void statisticalFinal(sqlite3_context* context) noexcept {
    auto& native=api();auto* state=static_cast<StatisticalState*>(native.aggregate(context,0));
    if(!state) {native.nullResult(context);return;}
    std::unique_ptr<NativeReportAggregate> aggregate(state->aggregate);state->aggregate=nullptr;
    try {
        if(state->error) throw VbaRuntimeError(state->error,"Statistical SQL aggregate failed");
        const auto value=aggregate ? aggregate->result() : VbaValue{};
        if(value.kind==VbaValue::Kind::Null) native.nullResult(context);else native.numberResult(context,value.number);
    } catch(const VbaRuntimeError& failure) {functionError(context,failure);}
      catch(...) {native.error(context,"LinAccess SQL function error 7",-1);}
}
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
    if (state->error) native.error(context,state->error==7 ? "LinAccess SQL function error 7" : "LinAccess SQL function error 3251: ambiguous joined update",-1);
}
void singleFinal(sqlite3_context* context) noexcept {
    auto& native=api(); auto* state=static_cast<SingleState*>(native.aggregate(context,0));
    if (!state) { native.nullResult(context); return; }
    if (state->error) native.error(context,state->error==7 ? "LinAccess SQL function error 7" : "LinAccess SQL function error 3251: ambiguous joined update",-1);
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
    if (number==5 || number==6 || number==7 || number==11 || number==13 || number==52 || number==53 || number==70 || number==75 || number==76 || number==94 || number==438 || number==450 || number==3035 || number==3113 || number==3146 || number==3162 || number==3163 || number==3315 || number==3317 || number==3061 || number==2450 || number==2451 || number==2455 || number==2465 || number==3251 || number==3420 || number==3085) return number;
    return {};
}
std::string jetSqlFunctionErrorExplanation(std::string_view message) {
    const auto original=QString::fromUtf8(message.data(),message.size());const QString marker="LinAccess SQL function error 3317: ";if(original.startsWith(marker)) {const auto text=original.mid(marker.size());if(!text.isEmpty() && text.size()<=255 && !text.contains(QChar(0)) && text.isValidUtf16()) return text.toStdString();}
    const auto storage=original.toCaseFolded().toStdString();const std::string_view foldedMessage(storage);
    if(foldedMessage.find("linaccess sql function error 3317")!=std::string_view::npos) return "Field value violates its ValidationRule";
    if(foldedMessage.find("linaccess sql function error 3315:")!=std::string_view::npos) return "Field cannot contain a zero-length string";
    if(foldedMessage.find("linaccess sql function error 3163:")!=std::string_view::npos) return "Text is too long for this field";
    if(foldedMessage.find("linaccess sql function error 3251: ambiguous joined update")!=std::string_view::npos)
        return "Joined UPDATE has conflicting source values for one target";
    return "Jet SQL function argument invalid";
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
    const auto ensure = [&](const char* name,int count,Api::Callback scalar,Api::Callback step,void (*final)(sqlite3_context*),int flags=SQLITE_UTF8) {
        QSqlQuery probe(db);
        QStringList args;
        for (int i = 0; i < count; ++i) args.append("NULL");
        if (probe.prepare("SELECT " + QString::fromLatin1(name) + "(" + args.join(',') + ")")) return;
        probe.clear();
        if (native.create(connection,name,count,flags,objects,scalar,step,final,nullptr) != 0)
            throw VbaRuntimeError(3251,"Jet SQL function registration failed");
    };
    ensure("linaccess_next_counter",2,&nextAccessCounter,nullptr,nullptr);
    ensure("linaccess_jet_iif",3,&iif,nullptr,nullptr);
    ensure("linaccess_jet_rgb",-1,&rgbColor,nullptr,nullptr);
    ensure("linaccess_jet_isnumeric",-1,&inspect_isnumeric,nullptr,nullptr);
    ensure("linaccess_jet_isnumeric_typed",-1,&inspect_isnumeric_typed,nullptr,nullptr);
    ensure("linaccess_jet_vartype",-1,&inspect_vartype,nullptr,nullptr);
    ensure("linaccess_jet_vartype_typed",-1,&inspect_vartype_typed,nullptr,nullptr);
    ensure("linaccess_jet_typename",-1,&inspect_typename,nullptr,nullptr);
    ensure("linaccess_jet_typename_typed",-1,&inspect_typename_typed,nullptr,nullptr);

    ensure("linaccess_jet_rgb_typed",-1,&bound_rgb,nullptr,nullptr);
    ensure("linaccess_jet_qbcolor_typed",-1,&bound_qbcolor,nullptr,nullptr);
    ensure("linaccess_jet_partition_typed",-1,&bound_partition,nullptr,nullptr);

    ensure("linaccess_jet_qbcolor",-1,&qbColor,nullptr,nullptr);
    ensure("linaccess_jet_partition",-1,&partitionRange,nullptr,nullptr);
    ensure("linaccess_jet_fv",-1,&financialFv,nullptr,nullptr);
    ensure("linaccess_jet_fv_typed",-1,&financialFvTyped,nullptr,nullptr);
    ensure("linaccess_jet_pv",-1,&financialPv,nullptr,nullptr);
    ensure("linaccess_jet_pv_typed",-1,&financialPvTyped,nullptr,nullptr);
    ensure("linaccess_jet_nper",-1,&financialNper,nullptr,nullptr);
    ensure("linaccess_jet_nper_typed",-1,&financialNperTyped,nullptr,nullptr);
    ensure("linaccess_jet_rate",-1,&financialRate,nullptr,nullptr);
    ensure("linaccess_jet_rate_typed",-1,&financialRateTyped,nullptr,nullptr);
    ensure("linaccess_jet_ddb",-1,&financialDdb,nullptr,nullptr);
    ensure("linaccess_jet_ddb_typed",-1,&financialDdbTyped,nullptr,nullptr);
    ensure("linaccess_jet_sln",-1,&financialSln,nullptr,nullptr);
    ensure("linaccess_jet_sln_typed",-1,&financialSlnTyped,nullptr,nullptr);
    ensure("linaccess_jet_syd",-1,&financialSyd,nullptr,nullptr);
    ensure("linaccess_jet_syd_typed",-1,&financialSydTyped,nullptr,nullptr);
    ensure("linaccess_jet_pmt",-1,&financialPmt,nullptr,nullptr);
    ensure("linaccess_jet_pmt_typed",-1,&financialPmtTyped,nullptr,nullptr);
    ensure("linaccess_jet_ipmt",-1,&financialIpmt,nullptr,nullptr);
    ensure("linaccess_jet_ipmt_typed",-1,&financialIpmtTyped,nullptr,nullptr);
    ensure("linaccess_jet_ppmt",-1,&financialPpmt,nullptr,nullptr);
    ensure("linaccess_jet_ppmt_typed",-1,&financialPpmtTyped,nullptr,nullptr);
    ensure("linaccess_jet_round",-1,&roundNumber,nullptr,nullptr);
    ensure("linaccess_jet_abs_typed",-1,&normalized_abs_typed,nullptr,nullptr);
    ensure("linaccess_jet_abs",-1,&normalized_abs,nullptr,nullptr);
    ensure("linaccess_jet_int_typed",-1,&normalized_int_typed,nullptr,nullptr);
    ensure("linaccess_jet_int",-1,&normalized_int,nullptr,nullptr);
    ensure("linaccess_jet_fix_typed",-1,&normalized_fix_typed,nullptr,nullptr);
    ensure("linaccess_jet_fix",-1,&normalized_fix,nullptr,nullptr);
    ensure("linaccess_jet_sgn_typed",-1,&normalized_sgn_typed,nullptr,nullptr);
    ensure("linaccess_jet_sgn",-1,&normalized_sgn,nullptr,nullptr);
    ensure("linaccess_jet_round_typed",-1,&normalized_round_typed,nullptr,nullptr);
    ensure("linaccess_jet_ccur_typed",-1,&normalized_ccur_typed,nullptr,nullptr);

    ensure("linaccess_jet_asc",-1,&stringAsc,nullptr,nullptr);
    ensure("linaccess_jet_asc_typed",-1,&stringAscTyped,nullptr,nullptr);
    ensure("linaccess_jet_ascw",-1,&stringAscW,nullptr,nullptr);
    ensure("linaccess_jet_ascw_typed",-1,&stringAscWTyped,nullptr,nullptr);
    ensure("linaccess_jet_chr",-1,&stringChr,nullptr,nullptr);
    ensure("linaccess_jet_chr_typed",-1,&stringChrTyped,nullptr,nullptr);
    ensure("linaccess_jet_sqr",-1,&mathSqr,nullptr,nullptr);
    ensure("linaccess_jet_sqr_typed",-1,&mathSqrTyped,nullptr,nullptr);
    ensure("linaccess_jet_exp",-1,&mathExp,nullptr,nullptr);
    ensure("linaccess_jet_exp_typed",-1,&mathExpTyped,nullptr,nullptr);
    ensure("linaccess_jet_log",-1,&mathLog,nullptr,nullptr);
    ensure("linaccess_jet_log_typed",-1,&mathLogTyped,nullptr,nullptr);
    ensure("linaccess_jet_sin",-1,&mathSin,nullptr,nullptr);
    ensure("linaccess_jet_sin_typed",-1,&mathSinTyped,nullptr,nullptr);
    ensure("linaccess_jet_cos",-1,&mathCos,nullptr,nullptr);
    ensure("linaccess_jet_cos_typed",-1,&mathCosTyped,nullptr,nullptr);
    ensure("linaccess_jet_power",2,&numericPower,nullptr,nullptr);
    ensure("linaccess_jet_double_add",2,&double_add,nullptr,nullptr);
    ensure("linaccess_jet_double_subtract",2,&double_subtract,nullptr,nullptr);
    ensure("linaccess_jet_double_multiply",2,&double_multiply,nullptr,nullptr);
    ensure("linaccess_jet_double_divide",2,&double_divide,nullptr,nullptr);

    ensure("linaccess_jet_cstr",-1,&stringCStr,nullptr,nullptr);
    ensure("linaccess_jet_cstr_typed",-1,&stringCStrTyped,nullptr,nullptr);
    ensure("linaccess_jet_str",-1,&stringStr,nullptr,nullptr);
    ensure("linaccess_jet_str_typed",-1,&stringStrTyped,nullptr,nullptr);
    ensure("linaccess_jet_left",-1,&stringLeft,nullptr,nullptr);
    ensure("linaccess_jet_left_typed",-1,&stringLeftTyped,nullptr,nullptr);
    ensure("linaccess_jet_left_string",-1,&stringLeftString,nullptr,nullptr);
    ensure("linaccess_jet_left_string_typed",-1,&stringLeftStringTyped,nullptr,nullptr);
    ensure("linaccess_jet_right",-1,&stringRight,nullptr,nullptr);
    ensure("linaccess_jet_right_typed",-1,&stringRightTyped,nullptr,nullptr);
    ensure("linaccess_jet_right_string",-1,&stringRightString,nullptr,nullptr);
    ensure("linaccess_jet_right_string_typed",-1,&stringRightStringTyped,nullptr,nullptr);
    ensure("linaccess_jet_mid",-1,&stringMid,nullptr,nullptr);
    ensure("linaccess_jet_mid_typed",-1,&stringMidTyped,nullptr,nullptr);
    ensure("linaccess_jet_mid_string",-1,&stringMidString,nullptr,nullptr);
    ensure("linaccess_jet_mid_string_typed",-1,&stringMidStringTyped,nullptr,nullptr);
    ensure("linaccess_jet_replace",-1,&stringReplace,nullptr,nullptr);
    ensure("linaccess_jet_replace_typed",-1,&stringReplaceTyped,nullptr,nullptr);
    ensure("linaccess_jet_instr",-1,&stringInStr,nullptr,nullptr);
    ensure("linaccess_jet_instr_typed",-1,&stringInStrTyped,nullptr,nullptr);
    ensure("linaccess_jet_strcomp",-1,&stringStrComp,nullptr,nullptr);
    ensure("linaccess_jet_strcomp_typed",-1,&stringStrCompTyped,nullptr,nullptr);
    ensure("linaccess_jet_trim",-1,&stringTrim,nullptr,nullptr);
    ensure("linaccess_jet_trim_typed",-1,&stringTrimTyped,nullptr,nullptr);
    ensure("linaccess_jet_trim_string",-1,&stringTrimString,nullptr,nullptr);
    ensure("linaccess_jet_trim_string_typed",-1,&stringTrimStringTyped,nullptr,nullptr);
    ensure("linaccess_jet_ltrim",-1,&stringLTrim,nullptr,nullptr);
    ensure("linaccess_jet_ltrim_typed",-1,&stringLTrimTyped,nullptr,nullptr);
    ensure("linaccess_jet_ltrim_string",-1,&stringLTrimString,nullptr,nullptr);
    ensure("linaccess_jet_ltrim_string_typed",-1,&stringLTrimStringTyped,nullptr,nullptr);
    ensure("linaccess_jet_rtrim",-1,&stringRTrim,nullptr,nullptr);
    ensure("linaccess_jet_rtrim_typed",-1,&stringRTrimTyped,nullptr,nullptr);
    ensure("linaccess_jet_rtrim_string",-1,&stringRTrimString,nullptr,nullptr);
    ensure("linaccess_jet_rtrim_string_typed",-1,&stringRTrimStringTyped,nullptr,nullptr);
    ensure("linaccess_jet_lcase_typed",-1,&stringLCaseTyped,nullptr,nullptr);
    ensure("linaccess_jet_ucase_typed",-1,&stringUCaseTyped,nullptr,nullptr);
    ensure("linaccess_jet_lcase_string",-1,&stringLCaseString,nullptr,nullptr);
    ensure("linaccess_jet_lcase_string_typed",-1,&stringLCaseStringTyped,nullptr,nullptr);
    ensure("linaccess_jet_ucase_string",-1,&stringUCaseString,nullptr,nullptr);
    ensure("linaccess_jet_ucase_string_typed",-1,&stringUCaseStringTyped,nullptr,nullptr);
    ensure("linaccess_jet_strreverse_typed",-1,&stringReverseTyped,nullptr,nullptr);
    ensure("linaccess_jet_instrrev_typed",-1,&stringInStrRevTyped,nullptr,nullptr);
    ensure("linaccess_jet_cbool",-1,&convert_cbool,nullptr,nullptr);
    ensure("linaccess_jet_cbool_typed",-1,&convert_cbool_typed,nullptr,nullptr);
    ensure("linaccess_jet_cbyte",-1,&convert_cbyte,nullptr,nullptr);
    ensure("linaccess_jet_cbyte_typed",-1,&convert_cbyte_typed,nullptr,nullptr);
    ensure("linaccess_jet_cint",-1,&convert_cint,nullptr,nullptr);
    ensure("linaccess_jet_cint_typed",-1,&convert_cint_typed,nullptr,nullptr);
    ensure("linaccess_jet_clng",-1,&convert_clng,nullptr,nullptr);
    ensure("linaccess_jet_clng_typed",-1,&convert_clng_typed,nullptr,nullptr);
    ensure("linaccess_jet_csng",-1,&convert_csng,nullptr,nullptr);
    ensure("linaccess_jet_csng_typed",-1,&convert_csng_typed,nullptr,nullptr);
    ensure("linaccess_jet_cdbl",-1,&convert_cdbl,nullptr,nullptr);
    ensure("linaccess_jet_cdbl_typed",-1,&convert_cdbl_typed,nullptr,nullptr);

    ensure("linaccess_jet_tan",-1,&mathTan,nullptr,nullptr);
    ensure("linaccess_jet_tan_typed",-1,&mathTanTyped,nullptr,nullptr);
    ensure("linaccess_jet_atn",-1,&mathAtn,nullptr,nullptr);
    ensure("linaccess_jet_atn_typed",-1,&mathAtnTyped,nullptr,nullptr);
    ensure("linaccess_jet_lcase",-1,&stringLCase,nullptr,nullptr);
    ensure("linaccess_jet_ucase",-1,&stringUCase,nullptr,nullptr);
    ensure("linaccess_jet_filelen",-1,&fileLength,nullptr,nullptr);
    ensure("linaccess_jet_curdir",-1,&currentDirectory,nullptr,nullptr);
    ensure("linaccess_jet_filedatetime",-1,&fileModified,nullptr,nullptr);
    ensure("linaccess_jet_date_add_1",-1,&calendarDateArithmetic<1,false>,nullptr,nullptr);
    ensure("linaccess_jet_date_add_2",-1,&calendarDateArithmetic<2,false>,nullptr,nullptr);
    ensure("linaccess_jet_date_add_3",-1,&calendarDateArithmetic<3,false>,nullptr,nullptr);
    ensure("linaccess_jet_date_subtract_1",-1,&calendarDateArithmetic<1,true>,nullptr,nullptr);
    ensure("linaccess_jet_date_subtract_2",-1,&calendarDateArithmetic<2,true>,nullptr,nullptr);
    ensure("linaccess_jet_date_subtract_3",-1,&calendarDateArithmetic<3,true>,nullptr,nullptr);
    ensure("linaccess_jet_date_equal",-1,&calendarCompare_equal,nullptr,nullptr);
    ensure("linaccess_jet_date_not_equal",-1,&calendarCompare_not_equal,nullptr,nullptr);
    ensure("linaccess_jet_date_less",-1,&calendarCompare_less,nullptr,nullptr);
    ensure("linaccess_jet_date_less_equal",-1,&calendarCompare_less_equal,nullptr,nullptr);
    ensure("linaccess_jet_date_greater",-1,&calendarCompare_greater,nullptr,nullptr);
    ensure("linaccess_jet_date_greater_equal",-1,&calendarCompare_greater_equal,nullptr,nullptr);
    ensure("linaccess_jet_datevalue",-1,&calendarDateValue,nullptr,nullptr);
    ensure("linaccess_jet_datevalue_typed",-1,&calendarDateValueTyped,nullptr,nullptr);
    ensure("linaccess_jet_timevalue",-1,&calendarTimeValue,nullptr,nullptr);
    ensure("linaccess_jet_timevalue_typed",-1,&calendarTimeValueTyped,nullptr,nullptr);
    ensure("linaccess_jet_dateadd",-1,&calendarDateAdd,nullptr,nullptr);
    ensure("linaccess_jet_dateadd_typed",-1,&calendarDateAddTyped,nullptr,nullptr);
    ensure("linaccess_jet_date_sum",1,nullptr,&unsupportedDateSumStep,&unsupportedDateSumFinal);
    ensure("linaccess_jet_dateserial",-1,&calendarDateSerial,nullptr,nullptr);
    ensure("linaccess_jet_dateserial_typed",-1,&calendarDateSerialTyped,nullptr,nullptr);
    ensure("linaccess_jet_timeserial",-1,&calendarTimeSerial,nullptr,nullptr);
    ensure("linaccess_jet_timeserial_typed",-1,&calendarTimeSerialTyped,nullptr,nullptr);
    ensure("linaccess_jet_cdate",-1,&calendarCDate,nullptr,nullptr);
    ensure("linaccess_jet_cdate_typed",-1,&calendarCDateTyped,nullptr,nullptr);
    ensure("linaccess_jet_datepart",-1,&calendarDatePart,nullptr,nullptr);
    ensure("linaccess_jet_datepart_typed",-1,&calendarDatePartTyped,nullptr,nullptr);
    ensure("linaccess_jet_choose",-1,&scalarChoose,nullptr,nullptr);
    ensure("linaccess_jet_switch",-1,&scalarSwitch,nullptr,nullptr);
    ensure("linaccess_jet_val",-1,&stringVal,nullptr,nullptr);
    ensure("linaccess_jet_hex_string",-1,&alias_hex_string,nullptr,nullptr);
    ensure("linaccess_jet_hex_string_typed",-1,&alias_hex_string_typed,nullptr,nullptr);
    ensure("linaccess_jet_oct_string",-1,&alias_oct_string,nullptr,nullptr);
    ensure("linaccess_jet_oct_string_typed",-1,&alias_oct_string_typed,nullptr,nullptr);
    ensure("linaccess_jet_space_string",-1,&alias_space_string,nullptr,nullptr);
    ensure("linaccess_jet_space_string_typed",-1,&alias_space_string_typed,nullptr,nullptr);
    ensure("linaccess_jet_string_string",-1,&alias_string_string,nullptr,nullptr);
    ensure("linaccess_jet_string_string_typed",-1,&alias_string_string_typed,nullptr,nullptr);
    ensure("linaccess_jet_chr_string",-1,&alias_chr_string,nullptr,nullptr);
    ensure("linaccess_jet_chr_string_typed",-1,&alias_chr_string_typed,nullptr,nullptr);
    ensure("linaccess_jet_chrw_string",-1,&alias_chrw_string,nullptr,nullptr);
    ensure("linaccess_jet_chrw_string_typed",-1,&alias_chrw_string_typed,nullptr,nullptr);
    ensure("linaccess_jet_hex",-1,&stringHex,nullptr,nullptr);
    ensure("linaccess_jet_hex_typed",-1,&stringHexTyped,nullptr,nullptr);
    ensure("linaccess_jet_oct",-1,&stringOct,nullptr,nullptr);
    ensure("linaccess_jet_oct_typed",-1,&stringOctTyped,nullptr,nullptr);
    ensure("linaccess_jet_space",-1,&stringSpace,nullptr,nullptr);
    ensure("linaccess_jet_space_typed",-1,&stringSpaceTyped,nullptr,nullptr);
    ensure("linaccess_jet_string",-1,&stringRepeat,nullptr,nullptr);
    ensure("linaccess_jet_string_typed",-1,&stringRepeatTyped,nullptr,nullptr);
    ensure("linaccess_jet_chrw",-1,&stringChrW,nullptr,nullptr);
    ensure("linaccess_jet_chrw_typed",-1,&stringChrWTyped,nullptr,nullptr);
    ensure("linaccess_jet_strreverse",-1,&stringReverse,nullptr,nullptr);
    ensure("linaccess_jet_instrrev",-1,&stringInStrRev,nullptr,nullptr);
    ensure("linaccess_jet_datediff",-1,&calendarDatediff,nullptr,nullptr);
    ensure("linaccess_jet_datediff_typed",-1,&calendarDatediffTyped,nullptr,nullptr);
    ensure("linaccess_jet_year",-1,&calendarYear,nullptr,nullptr);
    ensure("linaccess_jet_year_typed",-1,&calendarYearTyped,nullptr,nullptr);
    ensure("linaccess_jet_month",-1,&calendarMonth,nullptr,nullptr);
    ensure("linaccess_jet_month_typed",-1,&calendarMonthTyped,nullptr,nullptr);
    ensure("linaccess_jet_day",-1,&calendarDay,nullptr,nullptr);
    ensure("linaccess_jet_day_typed",-1,&calendarDayTyped,nullptr,nullptr);
    ensure("linaccess_jet_hour",-1,&calendarHour,nullptr,nullptr);
    ensure("linaccess_jet_hour_typed",-1,&calendarHourTyped,nullptr,nullptr);
    ensure("linaccess_jet_minute",-1,&calendarMinute,nullptr,nullptr);
    ensure("linaccess_jet_minute_typed",-1,&calendarMinuteTyped,nullptr,nullptr);
    ensure("linaccess_jet_second",-1,&calendarSecond,nullptr,nullptr);
    ensure("linaccess_jet_second_typed",-1,&calendarSecondTyped,nullptr,nullptr);
    ensure("linaccess_jet_weekday",-1,&calendarWeekday,nullptr,nullptr);
    ensure("linaccess_jet_weekday_typed",-1,&calendarWeekdayTyped,nullptr,nullptr);
    ensure("linaccess_jet_monthname",-1,&monthName,nullptr,nullptr,SQLITE_UTF8|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_monthname_typed",-1,&monthNameTyped,nullptr,nullptr,SQLITE_UTF8|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_weekdayname",-1,&weekdayName,nullptr,nullptr,SQLITE_UTF8|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_weekdayname_typed",-1,&weekdayNameTyped,nullptr,nullptr,SQLITE_UTF8|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_ccur",-1,&currencyNumber,nullptr,nullptr);
    ensure("linaccess_jet_currency_store",1,&currencyStore,nullptr,nullptr);
    ensure("linaccess_jet_field_store",2,&fieldStore,nullptr,nullptr);
    ensure("linaccess_jet_currency_sum",1,nullptr,&currencySumStep,&currencySumFinal);
    ensure("linaccess_jet_currency_add",2,&currencyAdd,nullptr,nullptr);
    ensure("linaccess_jet_currency_subtract",2,&currencySubtract,nullptr,nullptr);
    ensure("linaccess_jet_currency_multiply",2,&currencyMultiply,nullptr,nullptr);
    ensure("linaccess_jet_currency_ticks_add",2,&currencyTicksAdd,nullptr,nullptr);
    ensure("linaccess_jet_currency_ticks_subtract",2,&currencyTicksSubtract,nullptr,nullptr);
    ensure("linaccess_jet_currency_ticks_multiply",2,&currencyTicksMultiply,nullptr,nullptr);
    ensure("linaccess_jet_var",1,nullptr,&statisticalStep<false>,&statisticalFinal);
    ensure("linaccess_jet_stdev",1,nullptr,&statisticalStep<true>,&statisticalFinal);
    ensure("linaccess_jet_varp",1,nullptr,&statisticalStep<false,true>,&statisticalFinal);
    ensure("linaccess_jet_stdevp",1,nullptr,&statisticalStep<true,true>,&statisticalFinal);
    ensure("linaccess_jet_first",1,nullptr,&firstStep,&aggregateFinal);
    ensure("linaccess_jet_last",1,nullptr,&lastStep,&aggregateFinal);
    ensure("linaccess_jet_update_value",1,nullptr,&singleStep,&singleFinal);
    ensure("linaccess_jet_format_typed",-1,&format_format_typed,nullptr,nullptr,SQLITE_UTF8|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_format_string_typed",-1,&format_format_string_typed,nullptr,nullptr,SQLITE_UTF8|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_formatnumber_typed",-1,&format_formatnumber_typed,nullptr,nullptr,SQLITE_UTF8|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_formatcurrency_typed",-1,&format_formatcurrency_typed,nullptr,nullptr,SQLITE_UTF8|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_formatpercent_typed",-1,&format_formatpercent_typed,nullptr,nullptr,SQLITE_UTF8|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_formatdatetime_typed",-1,&format_formatdatetime_typed,nullptr,nullptr,SQLITE_UTF8|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_format_string",2,&formatString,nullptr,nullptr);
    ensure("linaccess_jet_format",2,&formatVariant,nullptr,nullptr);
    ensure("linaccess_jet_formatnumber",-1,&formatNumber,nullptr,nullptr,SQLITE_UTF8|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_formatcurrency",-1,&formatCurrency,nullptr,nullptr,SQLITE_UTF8|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_formatpercent",-1,&formatPercent,nullptr,nullptr,SQLITE_UTF8|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_formatdatetime",-1,&formatDateTime,nullptr,nullptr,SQLITE_UTF8|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_concat",2,&concatenate,nullptr,nullptr);
    ensure("linaccess_jet_len",1,&textLength,nullptr,nullptr);
    ensure("linaccess_jet_len_typed",-1,&textLengthTyped,nullptr,nullptr);
    ensure("linaccess_jet_lenb_typed",-1,&stringLenBTyped,nullptr,nullptr);
    ensure("linaccess_jet_default_date",0,&dateFieldDefault<0>,nullptr,nullptr,SQLITE_UTF8|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_default_now",0,&dateFieldDefault<1>,nullptr,nullptr,SQLITE_UTF8|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_default_time",0,&dateFieldDefault<2>,nullptr,nullptr,SQLITE_UTF8|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_nonempty_text",1,&nonemptyText,nullptr,nullptr,SQLITE_UTF8|SQLITE_DETERMINISTIC|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_validation",2,&fieldValidation,nullptr,nullptr,SQLITE_UTF8|SQLITE_DETERMINISTIC|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_validation",1,&fieldValidation,nullptr,nullptr,SQLITE_UTF8|SQLITE_DETERMINISTIC|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_text_capacity",2,&textCapacity,nullptr,nullptr,SQLITE_UTF8|SQLITE_DETERMINISTIC|SQLITE_INNOCUOUS);
    ensure("linaccess_jet_lenb",-1,&stringLenB,nullptr,nullptr);
    ensure("linaccess_jet_like_pattern",1,&likePattern,nullptr,nullptr);
    ensure("linaccess_jet_object_value",3,&objectValue,nullptr,nullptr);
    ensure("linaccess_jet_object_property",4,&objectProperty,nullptr,nullptr);
    vba_sql::install(db,connection);
}
QMap<QString,QString> nativeVbaSqlFunctionResultTypes(QSqlDatabase db) {
    installJetSqlFunctions(db); return vba_sql::declaredTypes(db);
}
void setJetSqlObjectResolver(QSqlDatabase db,JetSqlObjectResolver resolver) {
    installJetSqlFunctions(db);
    objectContext(db)->resolver=std::move(resolver);
}
void setJetSqlObjectPropertyResolver(QSqlDatabase db,JetSqlObjectPropertyResolver resolver) {
    installJetSqlFunctions(db);objectContext(db)->propertyResolver=std::move(resolver);
}

}
