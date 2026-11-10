#pragma once
#include <string>
#include <string_view>
namespace linaccess {
inline std::string_view vbaIdentifierType(std::string_view name) {
    if(name.empty()) return {};
    switch(name.back()) {
    case '%': return "integer"; case '&': return "long";
    case '!': return "single"; case '#': return "double";
    case '@': return "currency"; case '$': return "string";
    default: return {};
    }
}
inline std::string vbaIdentifierName(std::string_view name) {
    return std::string(vbaIdentifierType(name).empty() ? name : name.substr(0,name.size()-1));
}
}
