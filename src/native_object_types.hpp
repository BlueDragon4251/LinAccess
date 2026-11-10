#pragma once
#include <string_view>

namespace linaccess {
// Closed native class vocabulary; actual reference checks still use the binder.
// Keep vocabulary data out of this header so adding a type rebuilds one source.
std::string_view nativeAccessObjectType(std::string_view type);
std::string_view nativeScriptingObjectType(std::string_view type);
std::string_view nativeObjectType(std::string_view type);
}
