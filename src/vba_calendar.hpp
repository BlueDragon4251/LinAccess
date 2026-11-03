#pragma once
#include "vba_expression.hpp"
namespace linaccess {
// Gregorian ISO conversion and OLE Automation Date serials; locale parsing is not inferred.
VbaValue vbaDateValue(const VbaValue& value);
std::string vbaDateText(double serial);
VbaValue vbaCalendarFunction(std::string_view name,const std::vector<VbaValue>& arguments);
}
