#pragma once
#include "vba_expression.hpp"
#include <string_view>
namespace linaccess {
// Bounded date/numeric formatting shared by SQL and Qt VBA hosts.
// ISO date/timestamp text and numeric OLE Automation days are accepted.
VbaValue formatAccessValue(const VbaValue& value,std::string_view mask,bool stringResult);
VbaValue formatAccessNumericValue(const VbaValue& value,bool currency,const VbaNumericFormatOptions& options={});
bool isAccessNumericText(std::string_view text);
}
