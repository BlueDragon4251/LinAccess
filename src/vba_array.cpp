#include "vba_array.hpp"
#include <limits>
#include <algorithm>
namespace linaccess {
namespace {
// Immutable storage gives Variant assignment value semantics, including nested arrays.
class ArrayStorage final : public VbaObject {
public:
    const std::vector<VbaValue> values;
    const std::vector<VbaArrayDimension> bounds;
    const bool strings;
    const std::shared_ptr<const char> allocation;
    ArrayStorage(std::vector<VbaValue> elements,std::vector<VbaArrayDimension> dimensions,bool stringElements=false,std::shared_ptr<const char> identity={})
        :values(std::move(elements)),bounds(std::move(dimensions)),strings(stringElements),allocation(identity ? std::move(identity) : std::make_shared<const char>(0)) {}
    VbaValue member(std::string_view,const std::vector<VbaValue>&) override { throw VbaRuntimeError(13,"Array is not an object"); }
};
const ArrayStorage& storage(const VbaValue& value) {
    const auto* result=dynamic_cast<const ArrayStorage*>(value.object.get());
    if(value.kind!=VbaValue::Kind::Array || !result) throw VbaRuntimeError(13,"Array required");
    return *result;
}
std::vector<VbaArrayDimension> singleDimension(int lower,std::size_t count) {
    if(count>static_cast<std::size_t>(std::numeric_limits<std::int32_t>::max())) throw VbaRuntimeError(6,"Array size overflow");
    const auto upper=static_cast<std::int64_t>(lower)+static_cast<std::int64_t>(count)-1;
    if(upper<std::numeric_limits<std::int32_t>::min() || upper>std::numeric_limits<std::int32_t>::max()) throw VbaRuntimeError(6,"Array bound overflow");
    return {{lower,static_cast<int>(upper)}};
}
VbaValue array(std::vector<VbaValue> values,std::vector<VbaArrayDimension> bounds,bool strings=false,std::shared_ptr<const char> allocation={}) {
    VbaValue value; value.kind=VbaValue::Kind::Array;
    value.object=std::make_shared<ArrayStorage>(std::move(values),std::move(bounds),strings,std::move(allocation)); return value;
}
std::size_t boundedShapeSize(const std::vector<VbaArrayDimension>& bounds) {
    if(bounds.empty() || bounds.size()>60) throw VbaRuntimeError(9,"Array dimension count out of range");
    std::uint64_t count=1;
    for(const auto& bound:bounds) {
        const auto length=static_cast<std::int64_t>(bound.upper)-bound.lower+1;
        if(length<0) throw VbaRuntimeError(9,"Invalid array bounds");
        if(length && count>1000000/static_cast<std::uint64_t>(length)) throw VbaRuntimeError(3251,"Array allocation exceeds the supported runtime limit");
        count*=static_cast<std::uint64_t>(length);
    }
    return static_cast<std::size_t>(count);
}
std::size_t offset(const VbaValue& value,const std::vector<VbaValue>& indices) {
    if(!isVbaArray(value)) throw VbaRuntimeError(13,"Array required");
    if(value.kind==VbaValue::Kind::ByteArray) {
        if(indices.size()!=1) throw VbaRuntimeError(9,"Byte array requires one subscript");
        const auto index=static_cast<std::int64_t>(coerceVbaModuleValue(scalarVbaValue(indices[0]),"long").number);
        if(index<0 || static_cast<std::uint64_t>(index)>=value.bytes.size()) throw VbaRuntimeError(9,"Subscript out of range");
        return static_cast<std::size_t>(index);
    }
    const auto& original=storage(value);
    if(indices.size()!=original.bounds.size()) throw VbaRuntimeError(9,"Array subscript count mismatch");
    std::size_t result=0,stride=1;
    // VBA/SAFEARRAY traversal varies the first dimension fastest.
    for(std::size_t dimension=0;dimension<indices.size();++dimension) {
        const auto index=static_cast<std::int64_t>(coerceVbaModuleValue(scalarVbaValue(indices[dimension]),"long").number);
        const auto bound=original.bounds[dimension];
        if(index<bound.lower || index>bound.upper) throw VbaRuntimeError(9,"Subscript out of range");
        result+=static_cast<std::size_t>(index-bound.lower)*stride;
        stride*=static_cast<std::size_t>(static_cast<std::int64_t>(bound.upper)-bound.lower+1);
    }
    return result;
}
}
bool isVbaArray(const VbaValue& value) { return value.kind==VbaValue::Kind::Array || value.kind==VbaValue::Kind::ByteArray; }
VbaValue makeVbaArray(std::vector<VbaValue> values,int lowerBound) {
    auto bounds=singleDimension(lowerBound,values.size()); return array(std::move(values),std::move(bounds));
}
VbaValue makeVbaArrayWithBounds(std::vector<VbaValue> values,std::vector<VbaArrayDimension> bounds) {
    const auto count=boundedShapeSize(bounds);
    if(count!=values.size()) throw VbaRuntimeError(9,"Array shape and storage differ");
    return array(std::move(values),std::move(bounds));
}
VbaValue makeVbaStringArray(std::vector<std::string> values) {
    std::vector<VbaValue> elements; elements.reserve(values.size());
    for(auto& value:values) elements.push_back(VbaValue::string(std::move(value)));
    auto bounds=singleDimension(0,elements.size()); return array(std::move(elements),std::move(bounds),true);
}
bool sameVbaArrayAllocation(const VbaValue& left,const VbaValue& right) {
    return left.kind==VbaValue::Kind::Array && right.kind==VbaValue::Kind::Array && storage(left).allocation==storage(right).allocation;
}
int vbaArrayType(const VbaValue& value) {
    if(value.kind==VbaValue::Kind::ByteArray) return 8209;
    return storage(value).strings ? 8200 : 8204;
}
std::size_t vbaArrayRank(const VbaValue& value) { return value.kind==VbaValue::Kind::ByteArray ? 1 : storage(value).bounds.size(); }
std::size_t vbaArraySize(const VbaValue& value) { return value.kind==VbaValue::Kind::ByteArray ? value.bytes.size() : storage(value).values.size(); }
std::vector<VbaValue> vbaArrayValues(const VbaValue& value) { return storage(value).values; }
VbaValue vbaArrayLinearElement(const VbaValue& value,std::size_t index) {
    if(index>=vbaArraySize(value)) throw VbaRuntimeError(9,"Subscript out of range");
    return value.kind==VbaValue::Kind::ByteArray ? VbaValue::byte(value.bytes[index]) : storage(value).values[index];
}
VbaValue resizeVbaArray(const VbaValue& value,int lowerBound,int upperBound,bool preserve) {
    const bool bytes=value.kind==VbaValue::Kind::ByteArray;
    const bool strings=value.kind==VbaValue::Kind::Array && storage(value).strings;
    if((bytes || strings) && !preserve) throw VbaRuntimeError(3251,"Non-Preserve typed array resize is not implemented");
    if(bytes && lowerBound!=0) throw VbaRuntimeError(9,"Preserve cannot change the Byte array lower bound");
    const auto count=static_cast<std::int64_t>(upperBound)-lowerBound+1;
    if(count<=0) throw VbaRuntimeError(9,"Invalid array bounds");
    if(count>1000000) throw VbaRuntimeError(3251,"Array allocation exceeds the supported runtime limit");
    if(bytes) {
        auto result=value.bytes; result.resize(static_cast<std::size_t>(count),0);
        return VbaValue::byteArray(std::move(result));
    }
    if(preserve && value.kind!=VbaValue::Kind::Empty && value.kind!=VbaValue::Kind::Array) throw VbaRuntimeError(13,"Preserve requires an array or uninitialized Variant");
    if(preserve && value.kind==VbaValue::Kind::Array && (storage(value).bounds.size()!=1 || storage(value).bounds[0].lower!=lowerBound))
        throw VbaRuntimeError(9,"Preserve cannot change rank or lower bound");
    std::vector<VbaValue> values(static_cast<std::size_t>(count),strings ? VbaValue::string("") : VbaValue::empty());
    if(preserve && value.kind==VbaValue::Kind::Array) {
        const auto& previous=storage(value).values;
        std::copy_n(previous.begin(),std::min(previous.size(),values.size()),values.begin());
    }
    return array(std::move(values),{{lowerBound,upperBound}},strings);
}
VbaValue resizeVbaArrayWithBounds(const VbaValue& value,std::vector<VbaArrayDimension> bounds,bool preserve) {
    if(bounds.size()==1) return resizeVbaArray(value,bounds[0].lower,bounds[0].upper,preserve);
    const auto count=boundedShapeSize(bounds);
    for(const auto& bound:bounds) if(bound.upper<bound.lower) throw VbaRuntimeError(9,"Invalid ReDim bounds");
    if(value.kind==VbaValue::Kind::ByteArray || (value.kind==VbaValue::Kind::Array && storage(value).strings))
        throw VbaRuntimeError(3251,"Multidimensional typed array resize is not implemented");
    if(preserve && value.kind!=VbaValue::Kind::Empty && value.kind!=VbaValue::Kind::Array) throw VbaRuntimeError(13,"Preserve requires an array or uninitialized Variant");
    if(preserve && value.kind==VbaValue::Kind::Array) {
        const auto& original=storage(value);
        if(original.bounds.size()!=bounds.size()) throw VbaRuntimeError(9,"Preserve cannot change array rank");
        for(std::size_t dimension=0;dimension<bounds.size();++dimension)
            if(bounds[dimension].lower!=original.bounds[dimension].lower || (dimension+1<bounds.size() && bounds[dimension].upper!=original.bounds[dimension].upper))
                throw VbaRuntimeError(9,"Preserve can resize only the last upper bound");
    }
    std::vector<VbaValue> values(count,VbaValue::empty());
    if(preserve && value.kind==VbaValue::Kind::Array) {
        const auto& previous=storage(value).values;
        std::copy_n(previous.begin(),std::min(previous.size(),values.size()),values.begin());
    }
    return array(std::move(values),std::move(bounds));
}
VbaValue vbaArrayBound(const VbaValue& value,bool upper,int dimension) {
    if(!isVbaArray(value)) throw VbaRuntimeError(13,"Array required");
    if(dimension<1 || static_cast<std::size_t>(dimension)>vbaArrayRank(value)) throw VbaRuntimeError(9,"Array dimension out of range");
    if(value.kind==VbaValue::Kind::ByteArray) {
        const auto bound=singleDimension(0,value.bytes.size())[0];
        return VbaValue::longInteger(upper ? bound.upper : bound.lower);
    }
    const auto bound=storage(value).bounds[static_cast<std::size_t>(dimension-1)];
    return VbaValue::longInteger(upper ? bound.upper : bound.lower);
}
VbaValue vbaArrayElement(const VbaValue& value,const std::vector<VbaValue>& indices) { return vbaArrayLinearElement(value,offset(value,indices)); }
VbaValue withVbaArrayElement(const VbaValue& value,const std::vector<VbaValue>& indices,const VbaValue& element) {
    const auto index=offset(value,indices);
    if(value.kind==VbaValue::Kind::ByteArray) {
        auto bytes=value.bytes; bytes[index]=static_cast<unsigned char>(coerceVbaModuleValue(element,"byte").number);
        return VbaValue::byteArray(std::move(bytes));
    }
    const auto& original=storage(value); auto values=original.values;
    values[index]=original.strings ? coerceVbaModuleValue(scalarVbaValue(element),"string") : element;
    return array(std::move(values),original.bounds,original.strings,original.allocation);
}
}
