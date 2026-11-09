#pragma once
#include "vba_runtime.hpp"
#include <string>
#include <string_view>
namespace linaccess {
class VbaConditionalError final:public VbaRuntimeError {
public:
    VbaConditionalError(int code,const std::string& message,int sourceLine):VbaRuntimeError(code,message),line(sourceLine) {}
    int line;
};
// Pure, bounded native preprocessing. Source line numbers are preserved; the
// caller retains original bytes for source identity, storage and debugging.
std::string preprocessVbaConditional(std::string_view source);
}
