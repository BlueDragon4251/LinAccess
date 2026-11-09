#pragma once
#include "native_find_value.hpp"
namespace linaccess {
struct NativeFindRecordArguments {
    VbaValue sought;
    int match{1},direction{2};
    bool sensitive{},formatted{},all{},first{true};
    QVariantMap payload() const {
        auto request=nativeFindValuePayload(sought);
        request.insert("match",match);request.insert("sensitive",sensitive);
        request.insert("first",first);request.insert("direction",direction);
        request.insert("formatted",formatted);request.insert("allfields",all);
        return request;
    }
};
// Inputs have already been scalarized once by the actual document/module host.
inline NativeFindRecordArguments nativeFindRecordArguments(const std::vector<VbaValue>& args) {
    if(args.empty() || args.size()>7 || args[0].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(450,"FindRecord argument count invalid");
    for(const auto& argument:args) if(argument.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null FindRecord argument");
    const auto number=[](const VbaValue& value) {
        const auto integer=sqlFieldValue(value,"INTEGER").toLongLong();
        if(integer<std::numeric_limits<std::int32_t>::min() || integer>std::numeric_limits<std::int32_t>::max()) throw VbaRuntimeError(6,"FindRecord argument overflow");
        return static_cast<int>(integer);
    };
    const auto optionalNumber=[&](std::size_t index,int fallback) {return index>=args.size() || args[index].kind==VbaValue::Kind::Missing ? fallback : number(args[index]);};
    const auto optionalBoolean=[&](std::size_t index,bool fallback) {return index>=args.size() || args[index].kind==VbaValue::Kind::Missing ? fallback : coerceVbaModuleValue(args[index],"boolean").boolean;};
    NativeFindRecordArguments result;result.match=optionalNumber(1,1);result.direction=optionalNumber(3,2);const auto field=optionalNumber(5,-1);result.all=field==0;
    result.sensitive=optionalBoolean(2,false);result.formatted=optionalBoolean(4,false);result.first=optionalBoolean(6,true);
    if(result.match<0 || result.match>2 || result.direction<0 || result.direction>2 || (field!=0 && field!=-1)) throw VbaRuntimeError(5,"Invalid FindRecord option");
    if(result.formatted && args[0].kind!=VbaValue::Kind::Text) throw VbaRuntimeError(3251,"Formatted FindRecord requires literal String input");
    result.sought=normalizeNativeFindValue(args[0]);const auto text=nativeFindValueText(result.sought);const bool isText=result.sought.kind==VbaValue::Kind::Text;
    if(isText && text.isEmpty()) throw VbaRuntimeError(3251,"Empty FindRecord is not implemented");
    if(isText && nativeFindHasWildcards(text)) (void)nativeFindTextMatches({},text,result.match,result.sensitive,true);
    if(!isText && result.match!=1) throw VbaRuntimeError(3251,"Typed FindRecord currently requires whole-field matching");
    return result;
}
}
