#pragma once
#include "vba_qt_value.hpp"
#include <QString>
namespace linaccess {
inline VbaOpenFormArguments nativeFormOpenArguments(std::vector<VbaValue> values) {
    if(values.empty() || values[0].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(450,"OpenForm requires a name");
    for(std::size_t index=1;index<values.size();++index) if(values[index].kind==VbaValue::Kind::Missing) values[index]=VbaValue::empty();
    return parseVbaOpenFormArguments(values);
}
struct NativeReportOpenArguments {QString name,where;std::optional<QString> openArgs;int windowMode{};};
// The actual caller scalarizes once and rechecks its project/document context.
inline NativeReportOpenArguments nativeReportOpenArguments(const std::vector<VbaValue>& values) {
    if(values.empty() || values.size()>6 || values[0].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(450,"OpenReport requires a name and up to five optional arguments");
    for(std::size_t index=0;index<values.size();++index) if(index!=5 && values[index].kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null OpenReport argument");
    const auto omitted=[&](std::size_t index) {return index>=values.size() || values[index].kind==VbaValue::Kind::Missing || values[index].kind==VbaValue::Kind::Empty;};
    const auto number=[&](std::size_t index,int fallback) {return omitted(index) ? fallback : static_cast<int>(coerceVbaModuleValue(values[index],"long").number);};
    NativeReportOpenArguments result;result.name=QString::fromStdString(values[0].asText());if(result.name.isEmpty()) throw VbaRuntimeError(2103,"Report name unavailable");
    if(number(1,0)!=2) throw VbaRuntimeError(3251,"OpenReport supports native preview only");
    if(!omitted(2) && !values[2].asText().empty()) throw VbaRuntimeError(3251,"Named report filters unsupported");
    if(!omitted(3)) result.where=QString::fromStdString(values[3].asText());if(result.where.size()>32768) throw VbaRuntimeError(5,"Report WhereCondition too long");
    result.windowMode=number(4,0);if(result.windowMode<0 || result.windowMode>3) throw VbaRuntimeError(3251,"Report window mode unsupported");
    if(!omitted(5) && values[5].kind!=VbaValue::Kind::Null) {if(values[5].kind!=VbaValue::Kind::Text) throw VbaRuntimeError(3251,"Report OpenArgs currently requires text or Null");(void)vbaStringCodeUnits(values[5].text);result.openArgs=QString::fromStdString(values[5].text);}
    return result;
}
}
