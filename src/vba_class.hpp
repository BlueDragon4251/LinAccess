#pragma once
#include "vba_expression.hpp"
#include <functional>
namespace linaccess {
struct NativeVbaClassDefinition;
struct VbaClassContext {
    std::function<std::unique_ptr<VbaHost>()> host;
    std::function<void()> check;
};
// Compile actual class source without running initialization or user code.
std::shared_ptr<const NativeVbaClassDefinition> compileNativeVbaClass(std::string name,std::string source);
std::vector<std::string> nativeVbaClassDeclaredTypes(const NativeVbaClassDefinition& definition);
std::shared_ptr<VbaObject> constructNativeVbaClass(std::shared_ptr<const NativeVbaClassDefinition> definition,VbaClassContext context);
}
