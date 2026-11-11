#pragma once
#include "vba_expression.hpp"
#include <cstdint>
namespace linaccess {
// Decimal Variant: unsigned 96-bit coefficient, sign and scale0..28.
// Its invariant text is authoritative; never round-trip through Double internally.
VbaValue decimalLiteral(std::string_view text);
VbaValue convertDecimal(const VbaValue& value);
VbaValue decimalBinary(std::string_view op,const VbaValue& left,const VbaValue& right);
VbaValue decimalUnary(std::string_view op,const VbaValue& value);
VbaValue decimalRound(const VbaValue& value,int places);
std::vector<unsigned char> decimalSqlKey(const VbaValue& value);
VbaValue decimalFromSqlKey(const std::vector<unsigned char>& bytes);
int decimalCompare(const VbaValue& left,const VbaValue& right);
std::int64_t decimalInteger(const VbaValue& value,std::int64_t minimum,std::int64_t maximum);
}
