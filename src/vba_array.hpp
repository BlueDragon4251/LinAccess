#pragma once
#include "vba_expression.hpp"
namespace linaccess {
struct VbaArrayDimension { int lower,upper; };
VbaValue resizeVbaArray(const VbaValue& value,int lowerBound,int upperBound,bool preserve);
VbaValue resizeVbaArrayWithBounds(const VbaValue& value,std::vector<VbaArrayDimension> bounds,bool preserve,std::string elementType={});
VbaValue makeVbaStringArray(std::vector<std::string> values);
bool sameVbaArrayAllocation(const VbaValue& left,const VbaValue& right);
int vbaArrayType(const VbaValue& value);
std::string vbaArrayTypeName(const VbaValue& value);
std::vector<VbaValue> vbaArrayValues(const VbaValue& value);
bool isVbaArray(const VbaValue& value);
VbaValue makeVbaArray(std::vector<VbaValue> values,int lowerBound=0);
VbaValue makeVbaArrayWithBounds(std::vector<VbaValue> values,std::vector<VbaArrayDimension> bounds);
std::size_t vbaArrayRank(const VbaValue& value);
std::size_t vbaArraySize(const VbaValue& value);
VbaValue vbaArrayLinearElement(const VbaValue& value,std::size_t index);
VbaValue vbaArrayBound(const VbaValue& value,bool upper,int dimension=1);
VbaValue vbaArrayElement(const VbaValue& value,const std::vector<VbaValue>& indices);
VbaValue withVbaArrayElement(const VbaValue& value,const std::vector<VbaValue>& indices,const VbaValue& element);
}
