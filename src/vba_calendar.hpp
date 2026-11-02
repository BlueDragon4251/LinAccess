#pragma once
#include "vba_expression.hpp"
namespace linaccess {
// Gregorian, unambiguous ISO text input; no locale or Date Variant inference.
VbaValue vbaCalendarFunction(std::string_view name,const std::vector<VbaValue>& arguments);
}
