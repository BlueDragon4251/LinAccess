#pragma once
#include <string>
#include <string_view>
#include <vector>
namespace linaccess {
struct VbaIdentifierUse {std::string name,type;bool call{},singleWordOperator{};};
// Native lexical root uses only; members, strings, bracket fields and intrinsic
// names are excluded. This does not execute a host or resolve project symbols.
std::vector<VbaIdentifierUse> vbaIdentifierUses(std::string_view expression);
}
