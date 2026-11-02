#pragma once
#include "vba_expression.hpp"
namespace linaccess {
bool isVbaArray(const VbaValue& value);
VbaValue makeVbaArray(std::vector<VbaValue> values,int lowerBound=0);
VbaValue vbaArrayBound(const VbaValue& value,bool upper);
VbaValue vbaArrayElement(const VbaValue& value,const std::vector<VbaValue>& indices);
VbaValue withVbaArrayElement(const VbaValue& value,const std::vector<VbaValue>& indices,const VbaValue& element);
}
