#pragma once

#include "vba_runtime.hpp"

#include <optional>
#include <functional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

namespace linaccess {

struct VbaValue {
    enum class Kind { Null, Empty, Number, Text, Boolean, Object } kind{Kind::Null};
    double number{};
    std::string text;
    bool boolean{};
    std::shared_ptr<VbaObject> object;

    static VbaValue empty();
    static VbaValue numeric(double value);
    static VbaValue string(std::string value);
    static VbaValue logical(bool value);
    static VbaValue reference(std::shared_ptr<VbaObject> value);
    std::string asText() const;
    double asNumber() const;
    bool truthy() const;
    std::optional<std::string> asField() const;
};

// Late-bound Access/DAO members. Unknown properties and methods must throw.
class VbaObject {
public:
    virtual ~VbaObject() = default;
    virtual VbaValue member(std::string_view name, const std::vector<VbaValue>& arguments) = 0;
    virtual void assign(std::string_view, const VbaValue&) {
        throw std::runtime_error("Objekteigenschaft ist nicht schreibbar");
    }
};

using VbaVariables = std::unordered_map<std::string, VbaValue>;

using VbaFunctionResolver = std::function<VbaValue(std::string_view, std::vector<VbaValue>&,
                                                 const std::vector<std::string>&)>;
bool validVbaExpression(std::string_view expression);
VbaValue evaluateVbaExpression(std::string_view expression, VbaHost& host,
                               const VbaVariables& variables, const VbaFunctionResolver& resolver = {});
}
