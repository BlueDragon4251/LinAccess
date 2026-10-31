#pragma once
#include "vba_expression.hpp"
#include <cstdint>
namespace linaccess {
VbaValue convertCurrency(const VbaValue& value);
// Checked fixed-point operations; division deliberately remains Double.
VbaValue currencyBinary(std::string_view op,const VbaValue& left,const VbaValue& right);
VbaValue negateCurrency(const VbaValue& value);
VbaValue floorCurrency(const VbaValue& value);
}
