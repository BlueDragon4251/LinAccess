#pragma once
#include "native_object_types.hpp"
#include <string_view>
namespace linaccess {
inline constexpr std::string_view projectClassPrefix="projectclass:";
inline bool isProjectClassType(std::string_view type) {return type.starts_with(projectClassPrefix);}
inline std::string_view projectClassTypeName(std::string_view type) {return isProjectClassType(type) ? type.substr(projectClassPrefix.size()) : std::string_view{};}
inline bool isDeclaredObjectType(std::string_view type) {return isProjectClassType(type) || !nativeObjectType(type).empty();}
}
namespace linaccess {
inline bool isProjectClassIdentifier(std::string_view name) {
    const auto letter=[](unsigned char ch){return (ch>='a' && ch<='z') || ch=='_';};
    if(name.empty() || name.size()>64 || !letter(name.front())) return false;
    for(const unsigned char ch:name) if(!letter(ch) && !(ch>='0' && ch<='9')) return false;
    if(!nativeObjectType(name).empty()) return false;
    constexpr std::string_view reserved[]={"variant","boolean","byte","integer","long","single","double","currency","date","string","decimal","longlong","longptr","any","new","as","type","enum","vba","scripting","dao","access"};
    for(const auto word:reserved) if(word==name) return false;
    return true;
}
}
