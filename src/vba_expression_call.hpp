#pragma once
#include "vba_expression.hpp"
#include <algorithm>
#include <cctype>
namespace linaccess {
// Expression callbacks have temporary actuals, unlike a procedure frame's
// addressable ByRef bindings. Reorder named actuals without inventing aliases.
inline std::vector<VbaValue> orderVbaExpressionActuals(const VbaArgumentSignature& signature,
    const std::vector<VbaValue>& arguments,const std::vector<std::string>& references) {
    const bool named=std::any_of(references.begin(),references.end(),[](const auto& reference){return reference.starts_with("@named:");});
    if(!named) return arguments;
    if(signature.variadic) throw VbaRuntimeError(3251,"Named ParamArray expression calls are not implemented");
    if(references.size()!=arguments.size()) throw VbaRuntimeError(450,"Expression argument bindings are inconsistent");
    std::vector<VbaValue> ordered(signature.names.size(),VbaValue::missing());
    std::vector<bool> assigned(signature.names.size());bool namedSeen=false;
    const auto key=[](std::string value) {for(auto& ch:value) ch=static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));return value;};
    for(std::size_t index=0;index<arguments.size();++index) {
        auto target=index;const auto& reference=references[index];
        if(reference.starts_with("@named:")) {
            namedSeen=true;const auto separator=reference.find(':',7);
            if(separator==std::string::npos) throw VbaRuntimeError(450,"Invalid expression named binding");
            const auto parameter=key(reference.substr(7,separator-7));target=signature.names.size();
            for(std::size_t candidate=0;candidate<signature.names.size();++candidate)
                if(key(signature.names[candidate])==parameter) {target=candidate;break;}
            if(target==signature.names.size()) throw VbaRuntimeError(448,"Named argument not found");
        } else if(namedSeen) throw VbaRuntimeError(450,"Positional arguments cannot follow named arguments");
        if(target>=ordered.size() || assigned[target]) throw VbaRuntimeError(450,"Expression argument supplied more than once or too many arguments");
        ordered[target]=arguments[index];assigned[target]=true;
    }
    return ordered;
}
}
