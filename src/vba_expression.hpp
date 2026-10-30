#pragma once

#include "vba_runtime.hpp"

#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>

namespace linaccess {

struct VbaValue {
    enum class Kind { Null, Number, Text, Boolean } kind{Kind::Null};
    double number{};
    std::string text;
    bool boolean{};

    static VbaValue numeric(double value);
    static VbaValue string(std::string value);
    static VbaValue logical(bool value);
    std::string asText() const;
    double asNumber() const;
    bool truthy() const;
    std::optional<std::string> asField() const;
};

using VbaVariables = std::unordered_map<std::string, VbaValue>;

bool validVbaExpression(std::string_view expression);
VbaValue evaluateVbaExpression(std::string_view expression, VbaHost& host,
                               const VbaVariables& variables);
}
