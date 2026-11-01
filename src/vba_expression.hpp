#pragma once

#include "vba_runtime.hpp"

#include <optional>
#include <cstdint>
#include <functional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <vector>

namespace linaccess {

// UTF-16 code-unit count for valid UTF-8; rejects malformed text.
std::size_t vbaStringCodeUnits(std::string_view text);

struct VbaValue {
    enum class Kind { Null, Empty, Number, Text, Boolean, Object, Error, Currency, ByteArray } kind{Kind::Null};
    double number{};
    std::int64_t currencyUnits{};
    std::string text;
    bool boolean{};
    int errorNumber{};
    std::shared_ptr<VbaObject> object;
    std::vector<unsigned char> bytes;

    static VbaValue empty();
    static VbaValue numeric(double value);
    static VbaValue currency(std::int64_t scaledUnits);
    static VbaValue string(std::string value);
    static VbaValue byteArray(std::vector<unsigned char> value);
    static VbaValue logical(bool value);
    static VbaValue reference(std::shared_ptr<VbaObject> value);
    static VbaValue errorValue(int number);
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
    virtual VbaValue scalarValue() { throw VbaRuntimeError(13,"Object has no scalar default value"); }
    virtual VbaValue defaultMember(const std::vector<VbaValue>& arguments,bool reference) { return reference ? objectMember("Item",arguments) : member("Item",arguments); }
    virtual VbaValue item(std::string_view name) { return member(name,{}); }
    virtual VbaValue objectMember(std::string_view name,const std::vector<VbaValue>& arguments) { return member(name,arguments); }
    virtual VbaValue objectItem(std::string_view name) { return item(name); }
    virtual void assignItem(std::string_view name,const VbaValue& value) { assign(name,value); }
    virtual std::vector<VbaValue> enumerate() {
        throw VbaRuntimeError(438,"Object does not support enumeration");
    }
    virtual void assign(std::string_view, const VbaValue&) {
        throw std::runtime_error("Objekteigenschaft ist nicht schreibbar");
    }
};

// Shared relational semantics for expression operators and Select Case.
VbaValue compareVbaValues(std::string_view relation,const VbaValue& left,const VbaValue& right,
                         VbaHost& host,VbaCompareMode mode);

// Resolve scalar defaults without erasing Null/Empty/Error subtype information.
VbaValue scalarVbaValue(VbaValue value);

using VbaVariables = std::unordered_map<std::string, VbaValue>;
struct VbaModuleState {
    VbaVariables values;
    std::unordered_map<std::string,std::string> types;
    std::unordered_set<std::string> publicVariables;
};
VbaValue coerceVbaModuleValue(const VbaValue& value,std::string_view type);

using VbaFunctionResolver = std::function<VbaValue(std::string_view, std::vector<VbaValue>&,
                                                 const std::vector<std::string>&)>;
bool validVbaExpression(std::string_view expression,bool withContext=false);
VbaValue evaluateVbaExpression(std::string_view expression, VbaHost& host,
                               const VbaVariables& variables, const VbaFunctionResolver& resolver = {},const VbaValue* withContext=nullptr,VbaCompareMode compareMode=VbaCompareMode::Binary);
VbaValue evaluateVbaObjectExpression(std::string_view expression,VbaHost& host,
                                    const VbaVariables& variables,const VbaFunctionResolver& resolver = {},const VbaValue* withContext=nullptr,VbaCompareMode compareMode=VbaCompareMode::Binary);
VbaValue evaluateAccessExpression(std::string_view expression,VbaHost& host);
}
