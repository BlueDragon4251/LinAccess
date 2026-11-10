#include "vba_project_types.hpp"
#include "vba_array.hpp"
#include "vba_fixed_string.hpp"
#include <limits>
#include <algorithm>
namespace linaccess {
namespace {
// Immutable storage gives Variant assignment value semantics, including nested arrays.
class ArrayStorage final : public VbaObject {
public:
    const std::vector<VbaValue> values;
    const std::vector<VbaArrayDimension> bounds;
    const std::string elementType;
    const bool detached;
    const std::shared_ptr<const char> allocation;
    ArrayStorage(std::vector<VbaValue> elements,std::vector<VbaArrayDimension> dimensions,std::string type="variant",std::shared_ptr<const char> identity={},bool snapshot=false)
        :values(std::move(elements)),bounds(std::move(dimensions)),elementType(std::move(type)),detached(snapshot),allocation(identity ? std::move(identity) : std::make_shared<const char>(0)) {}
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
VbaValue array(std::vector<VbaValue> values,std::vector<VbaArrayDimension> bounds,std::string elementType="variant",std::shared_ptr<const char> allocation={},bool detached=false) {
    if(elementType=="byte" && bounds.size()==1 && bounds[0].lower==0) {
        std::vector<unsigned char> bytes; bytes.reserve(values.size());
        for(const auto& value:values) bytes.push_back(static_cast<unsigned char>(value.number));
        return VbaValue::byteArray(std::move(bytes));
    }
    VbaValue value; value.kind=VbaValue::Kind::Array;
    value.object=std::make_shared<ArrayStorage>(std::move(values),std::move(bounds),std::move(elementType),std::move(allocation),detached); return value;
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
VbaValue makeUnallocatedVbaArray(std::string elementType) { (void)vbaFixedStringWidth(elementType);return array({}, {}, std::move(elementType)); }
VbaValue makeVbaArray(std::vector<VbaValue> values,int lowerBound) {
    auto bounds=singleDimension(lowerBound,values.size()); return array(std::move(values),std::move(bounds));
}
VbaValue makeDetachedVbaArray(std::vector<VbaValue> values,int lowerBound) {
    auto bounds=singleDimension(lowerBound,values.size()); return array(std::move(values),std::move(bounds),"variant",{},true);
}
bool isDetachedVbaArray(const VbaValue& value) { return value.kind==VbaValue::Kind::Array && storage(value).detached; }
VbaValue makeVbaArrayWithBounds(std::vector<VbaValue> values,std::vector<VbaArrayDimension> bounds) {
    const auto count=boundedShapeSize(bounds);
    if(count!=values.size()) throw VbaRuntimeError(9,"Array shape and storage differ");
    return array(std::move(values),std::move(bounds));
}
VbaValue makeVbaStringArray(std::vector<std::string> values) {
    std::vector<VbaValue> elements; elements.reserve(values.size());
    for(auto& value:values) elements.push_back(VbaValue::string(std::move(value)));
    auto bounds=singleDimension(0,elements.size()); return array(std::move(elements),std::move(bounds),"string");
}
bool sameVbaArrayAllocation(const VbaValue& left,const VbaValue& right) {
    if(left.kind==VbaValue::Kind::ByteArray && right.kind==VbaValue::Kind::ByteArray)
        return left.object && left.object==right.object;
    return left.kind==VbaValue::Kind::Array && right.kind==VbaValue::Kind::Array && storage(left).allocation==storage(right).allocation;
}
std::shared_ptr<const void> vbaArrayAllocation(const VbaValue& value) {
    if(value.kind==VbaValue::Kind::ByteArray) return value.object;
    return storage(value).allocation;
}
std::size_t vbaArrayElementOffset(const VbaValue& value,const std::vector<VbaValue>& indices) { return offset(value,indices); }
int vbaArrayType(const VbaValue& value) {
    if(value.kind==VbaValue::Kind::ByteArray) return 8209;
    const auto& type=storage(value).elementType;
    if(type=="byte") return 8209;
    if(type=="integer") return 8194;
    if(type=="long") return 8195;
    if(type=="single") return 8196;
    if(type=="double") return 8197;
    if(type=="currency") return 8198;
    if(type=="date") return 8199;
    if(type=="string" || type.starts_with("string*")) return 8200;
    if(isDeclaredObjectType(type)) return 8201;
    if(type=="boolean") return 8203;
    return 8204;
}
std::string vbaArrayElementType(const VbaValue& value) {return value.kind==VbaValue::Kind::ByteArray ? "byte" : storage(value).elementType;}
std::string vbaArrayTypeName(const VbaValue& value) {
    if(value.kind==VbaValue::Kind::ByteArray) return "Byte()";
    auto type=storage(value).elementType;if(isProjectClassType(type)) return std::string(projectClassTypeName(type))+"()";
    if(const auto native=nativeObjectType(type);!native.empty()) {
        const auto dot=native.rfind('.');const auto shortName=dot==std::string_view::npos ? native : native.substr(dot+1);
        static constexpr std::pair<std::string_view,std::string_view> names[]{{"errobject","ErrObject"},{"dbengine","DBEngine"},{"querydef","QueryDef"},{"querydefs","QueryDefs"},{"tabledef","TableDef"},{"tabledefs","TableDefs"},{"filesystemobject","FileSystemObject"},{"textstream","TextStream"},{"textbox","TextBox"},{"combobox","ComboBox"},{"listbox","ListBox"},{"checkbox","CheckBox"},{"optionbutton","OptionButton"},{"togglebutton","ToggleButton"},{"commandbutton","CommandButton"},{"optiongroup","OptionGroup"},{"tabcontrol","TabControl"},{"filedialog","FileDialog"},{"filedialogfilters","FileDialogFilters"},{"filedialogfilter","FileDialogFilter"},{"filedialogselecteditems","FileDialogSelectedItems"}};
        for(const auto& [key,name]:names) if(shortName==key) return std::string(name)+"()";
        type=std::string(shortName);
    }
    if(type.starts_with("string*")) type="string";
    if(!type.empty()) type[0]=static_cast<char>(type[0]-'a'+'A');
    return type+"()";
}
std::vector<VbaArrayDimension> vbaArrayBounds(const VbaValue& value) { return value.kind==VbaValue::Kind::ByteArray ? singleDimension(0,value.bytes.size()) : storage(value).bounds; }
std::size_t vbaArrayRank(const VbaValue& value) { return value.kind==VbaValue::Kind::ByteArray ? 1 : storage(value).bounds.size(); }
std::size_t vbaArraySize(const VbaValue& value) { return value.kind==VbaValue::Kind::ByteArray ? value.bytes.size() : storage(value).values.size(); }
std::vector<VbaValue> vbaArrayValues(const VbaValue& value) { return storage(value).values; }
VbaValue vbaArrayLinearElement(const VbaValue& value,std::size_t index) {
    if(index>=vbaArraySize(value)) throw VbaRuntimeError(9,"Subscript out of range");
    auto element=value.kind==VbaValue::Kind::ByteArray ? VbaValue::byte(value.bytes[index]) : storage(value).values[index];
    const bool variant=value.kind!=VbaValue::Kind::ByteArray && storage(value).elementType=="variant";
    return bindArithmeticValue(std::move(element),variant ? VbaValue::ArithmeticBinding::Variant : VbaValue::ArithmeticBinding::Static);
}
VbaValue resizeVbaArray(const VbaValue& value,int lowerBound,int upperBound,bool preserve) {
    return resizeVbaArrayWithBounds(value,{{lowerBound,upperBound}},preserve);
}
VbaValue resizeVbaArrayWithBounds(const VbaValue& value,std::vector<VbaArrayDimension> bounds,bool preserve,std::string elementType) {
    const auto count=boundedShapeSize(bounds);
    for(const auto& bound:bounds) if(bound.upper<bound.lower) throw VbaRuntimeError(9,"Invalid ReDim bounds");
    const bool originalArray=isVbaArray(value);
    const auto previousType=value.kind==VbaValue::Kind::ByteArray ? std::string("byte")
        : value.kind==VbaValue::Kind::Array ? storage(value).elementType : std::string("variant");
    if(elementType.empty()) elementType=previousType;
    if(elementType!="variant" && elementType!="byte" && elementType!="integer" && elementType!="long" && elementType!="single"
        && elementType!="double" && elementType!="currency" && elementType!="date" && elementType!="boolean" && elementType!="string" && !elementType.starts_with("string*") && elementType!="object" && !isDeclaredObjectType(elementType))
        throw VbaRuntimeError(3251,"Array element type is not implemented");
    const auto fixedWidth=vbaFixedStringWidth(elementType);
    // At most three UTF-8 bytes per UTF-16 unit, including future element writes.
    if(fixedWidth && count>(16*1024*1024)/(3*(*fixedWidth))) throw VbaRuntimeError(3251,"Fixed String array text storage exceeds16MiB");
    if(preserve && value.kind!=VbaValue::Kind::Empty && !originalArray) throw VbaRuntimeError(13,"Preserve requires an array or uninitialized Variant");
    if(preserve && originalArray) {
        if(elementType!=previousType) throw VbaRuntimeError(13,"Preserve cannot change array element type");
        const auto previousBounds=value.kind==VbaValue::Kind::ByteArray ? singleDimension(0,value.bytes.size()) : storage(value).bounds;
        if(!previousBounds.empty() && previousBounds.size()!=bounds.size()) throw VbaRuntimeError(9,"Preserve cannot change array rank");
        for(std::size_t dimension=0;dimension<previousBounds.size();++dimension)
            if(bounds[dimension].lower!=previousBounds[dimension].lower || (dimension+1<bounds.size() && bounds[dimension].upper!=previousBounds[dimension].upper))
                throw VbaRuntimeError(9,"Preserve can resize only the last upper bound");
    }
    if(elementType=="byte" && bounds.size()==1 && bounds[0].lower==0) {
        std::vector<unsigned char> bytes(count,0);
        if(preserve && value.kind==VbaValue::Kind::ByteArray) std::copy_n(value.bytes.begin(),std::min(count,value.bytes.size()),bytes.begin());
        return VbaValue::byteArray(std::move(bytes));
    }
    const auto initial=fixedWidth ? VbaValue::string(std::string(*fixedWidth,'\0')) : elementType=="variant" ? VbaValue::empty() : isDeclaredObjectType(elementType) ? VbaValue::reference(nullptr) : coerceVbaModuleValue(VbaValue::empty(),elementType);
    std::vector<VbaValue> values(count,initial);
    if(preserve && originalArray) for(std::size_t i=0;i<std::min(count,vbaArraySize(value));++i) values[i]=vbaArrayLinearElement(value,i);
    return array(std::move(values),std::move(bounds),std::move(elementType));
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
    return withVbaArrayLinearElement(value,offset(value,indices),element);
}
VbaValue withVbaArrayLinearElement(const VbaValue& value,std::size_t index,const VbaValue& element) {
    if(index>=vbaArraySize(value)) throw VbaRuntimeError(9,"Subscript out of range");
    if(value.kind==VbaValue::Kind::ByteArray) {
        auto result=value; result.bytes[index]=static_cast<unsigned char>(coerceVbaModuleValue(element,"byte").number);
        return result;
    }
    const auto& original=storage(value); auto values=original.values;
    if(original.elementType=="object" && element.kind!=VbaValue::Kind::Object) throw VbaRuntimeError(13,"Object array requires object references");
    if(isDeclaredObjectType(original.elementType)) values[index]=coerceVbaModuleValue(element,original.elementType);
    else values[index]=(original.elementType=="variant" || original.elementType=="object") ? element : coerceVbaModuleValue(scalarVbaValue(element),original.elementType);
    return array(std::move(values),original.bounds,original.elementType,original.allocation,original.detached);
}
}
