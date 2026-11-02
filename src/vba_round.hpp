#pragma once
#include "vba_expression.hpp"
namespace linaccess {
// Bounded native bankers rounding, shared by VBA and the Jet SQL adapter.
VbaValue vbaRound(const std::vector<VbaValue>& arguments);
}
