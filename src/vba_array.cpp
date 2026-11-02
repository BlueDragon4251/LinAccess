#include "vba_array.hpp"
#include <limits>
namespace linaccess {
namespace {
// Immutable storage gives Variant assignment value semantics, including nested arrays.
class ArrayStorage final : public VbaObject {
public:
    const std::vector<VbaValue> values;
    const int lower;
    ArrayStorage(std::vector<VbaValue> elements,int bound):values(std::move(elements)),lower(bound) {}
    VbaValue member(std::string_view,const std::vector<VbaValue>&) override { throw VbaRuntimeError(13,"Array is not an object"); }
};
const ArrayStorage& storage(const VbaValue& value) {
    const auto* result=dynamic_cast<const ArrayStorage*>(value.object.get());
    if(value.kind!=VbaValue::Kind::Array || !result) throw VbaRuntimeError(13,"Array required");
    return *result;
}
std::size_t offset(const VbaValue& value,const std::vector<VbaValue>& indices) {
    if(!isVbaArray(value)) throw VbaRuntimeError(13,"Array required");
    if(indices.size()!=1) throw VbaRuntimeError(9,"Array requires one subscript");
    const auto index=static_cast<std::int64_t>(coerceVbaModuleValue(scalarVbaValue(indices[0]),"long").number);
    const auto lower=value.kind==VbaValue::Kind::ByteArray ? 0 : storage(value).lower;
    const auto size=value.kind==VbaValue::Kind::ByteArray ? value.bytes.size() : storage(value).values.size();
    if(index<lower || static_cast<std::uint64_t>(index-lower)>=size) throw VbaRuntimeError(9,"Subscript out of range");
    return static_cast<std::size_t>(index-lower);
}
}
bool isVbaArray(const VbaValue& value) { return value.kind==VbaValue::Kind::Array || value.kind==VbaValue::Kind::ByteArray; }
VbaValue makeVbaArray(std::vector<VbaValue> values,int lowerBound) {
    if(lowerBound<0 || lowerBound>1 || values.size()>static_cast<std::size_t>(std::numeric_limits<std::int32_t>::max())) throw VbaRuntimeError(6,"Array size overflow");
    VbaValue value; value.kind=VbaValue::Kind::Array;
    value.object=std::make_shared<ArrayStorage>(std::move(values),lowerBound); return value;
}
VbaValue vbaArrayBound(const VbaValue& value,bool upper) {
    if(!isVbaArray(value)) throw VbaRuntimeError(13,"Array required");
    const auto lower=value.kind==VbaValue::Kind::ByteArray ? 0 : storage(value).lower;
    const auto size=value.kind==VbaValue::Kind::ByteArray ? value.bytes.size() : storage(value).values.size();
    const auto result=upper ? static_cast<std::int64_t>(lower)+static_cast<std::int64_t>(size)-1 : lower;
    if(result>std::numeric_limits<std::int32_t>::max()) throw VbaRuntimeError(6,"Array bound overflow");
    return VbaValue::longInteger(static_cast<std::int32_t>(result));
}
VbaValue vbaArrayElement(const VbaValue& value,const std::vector<VbaValue>& indices) {
    const auto index=offset(value,indices);
    return value.kind==VbaValue::Kind::ByteArray ? VbaValue::byte(value.bytes[index]) : storage(value).values[index];
}
VbaValue withVbaArrayElement(const VbaValue& value,const std::vector<VbaValue>& indices,const VbaValue& element) {
    const auto index=offset(value,indices);
    if(value.kind==VbaValue::Kind::ByteArray) {
        auto bytes=value.bytes; bytes[index]=static_cast<unsigned char>(coerceVbaModuleValue(element,"byte").number);
        return VbaValue::byteArray(std::move(bytes));
    }
    auto values=storage(value).values; values[index]=element;
    return makeVbaArray(std::move(values),storage(value).lower);
}
}
