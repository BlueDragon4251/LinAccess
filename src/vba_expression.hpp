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
    enum class Kind { Null, Empty, Number, Text, Boolean, Object, Error, Currency, ByteArray, Array, Missing, Date } kind{Kind::Null};
    enum class NumericType { Double, Integer, Long, Single, Byte } numericType{NumericType::Double};
    double number{};
    std::int64_t currencyUnits{};
    std::string text;
    bool boolean{};
    int errorNumber{};
    std::shared_ptr<VbaObject> object;
    std::vector<unsigned char> bytes;

    static VbaValue empty();
    static VbaValue missing();
    static VbaValue numeric(double value);
    static VbaValue dateSerial(double value);
    static VbaValue single(float value);
    static VbaValue byte(unsigned char value);
    static VbaValue integer(std::int16_t value);
    static VbaValue longInteger(std::int32_t value);
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

struct VbaArgumentSignature { std::vector<std::string> names; std::size_t required{}; };

// Late-bound Access/DAO members. Unknown properties and methods must throw.
class VbaObject {
public:
    virtual ~VbaObject() = default;
    virtual std::string_view className() const { return {}; }
    virtual std::optional<VbaArgumentSignature> argumentSignature(std::string_view) const { return {}; }
    virtual VbaValue member(std::string_view name, const std::vector<VbaValue>& arguments) = 0;
    // Advertise a zero-argument scalar default explicitly; getter errors propagate.
    virtual bool hasScalarDefault() const { return false; }
    virtual VbaValue scalarValue() { throw VbaRuntimeError(13,"Object has no scalar default value"); }
    virtual VbaValue defaultMember(const std::vector<VbaValue>& arguments,bool reference) { return reference ? objectMember("Item",arguments) : member("Item",arguments); }
    virtual VbaValue item(std::string_view name) { return member(name,{}); }
    virtual VbaValue objectMember(std::string_view name,const std::vector<VbaValue>& arguments) { return member(name,arguments); }
    virtual VbaValue objectItem(std::string_view name) { return item(name); }
    virtual void assignItem(std::string_view name,const VbaValue& value) { assign(name,value); }
    virtual void assignIndexedMember(std::string_view,const std::vector<VbaValue>&,const VbaValue&) {
        throw VbaRuntimeError(438,"Indexed property assignment is not supported by this object");
    }
    virtual void assignDefaultMember(const std::vector<VbaValue>& arguments,const VbaValue& value) {
        assignIndexedMember("Item",arguments,value);
    }
    virtual std::shared_ptr<void> enumerationGuard() { return {}; }
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
    std::unordered_map<std::string,VbaVariables> procedureStatics;
    std::unordered_map<std::string,std::string> types;
    std::unordered_set<std::string> publicVariables;
    std::unordered_set<std::string> fixedArrays;
    std::unordered_set<std::string> constants;
};
VbaValue coerceVbaModuleValue(const VbaValue& value,std::string_view type);

struct VbaArrayElementReference {
    std::string root,elementType;
    std::size_t offset{};
    std::shared_ptr<const void> allocation;
};
struct VbaEvaluatedArgument {
    VbaValue value;
    std::optional<VbaArrayElementReference> arrayElement;
};
// Metadata belongs to this synchronous invocation; legacy callbacks can still
// read the string references through the implicit vector view.
struct VbaArgumentReferences {
    std::vector<std::string> names;
    std::vector<std::optional<VbaArrayElementReference>> arrayElements;
    VbaArgumentReferences() = default;
    VbaArgumentReferences(std::vector<std::string> values) : names(std::move(values)) {}
    VbaArgumentReferences(std::vector<std::string> values,std::vector<std::optional<VbaArrayElementReference>> elements)
        : names(std::move(values)),arrayElements(std::move(elements)) {}
    operator const std::vector<std::string>&() const { return names; }
    auto begin() const { return names.begin(); }
    auto end() const { return names.end(); }
    auto size() const { return names.size(); }
    bool empty() const { return names.empty(); }
    const auto& operator[](std::size_t index) const { return names[index]; }
    const auto& at(std::size_t index) const { return names.at(index); }
    const auto& front() const { return names.front(); }
    const auto& back() const { return names.back(); }
};
// Native procedure calls encode named bindings as @named:parameter:reference.
using VbaFunctionResolver = std::function<VbaValue(std::string_view, std::vector<VbaValue>&,
                                                 const VbaArgumentReferences&)>;
VbaEvaluatedArgument evaluateVbaArgumentExpression(std::string_view expression,VbaHost& host,
    const VbaVariables& variables,const VbaFunctionResolver& resolver={},const VbaValue* withContext=nullptr,
    VbaCompareMode compareMode=VbaCompareMode::Binary,int arrayBase=0);
bool validVbaExpression(std::string_view expression,bool withContext=false);
VbaValue evaluateVbaExpression(std::string_view expression, VbaHost& host,
                               const VbaVariables& variables, const VbaFunctionResolver& resolver = {},const VbaValue* withContext=nullptr,VbaCompareMode compareMode=VbaCompareMode::Binary,int arrayBase=0);
VbaValue evaluateVbaObjectExpression(std::string_view expression,VbaHost& host,
                                    const VbaVariables& variables,const VbaFunctionResolver& resolver = {},const VbaValue* withContext=nullptr,VbaCompareMode compareMode=VbaCompareMode::Binary,int arrayBase=0);
VbaValue evaluateAccessExpression(std::string_view expression,VbaHost& host);
}
