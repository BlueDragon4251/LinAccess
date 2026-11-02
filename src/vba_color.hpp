#pragma once

#include "vba_expression.hpp"

namespace linaccess {
// Shared native VBA/SQL RGB channel conversion and packed OLE color.
VbaValue vbaRgb(const std::vector<VbaValue>& arguments);
}
