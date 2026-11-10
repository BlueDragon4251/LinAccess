#pragma once
#include <cstdint>
namespace linaccess {
class VbaHost;
struct VbaValue;
// Actual mutable state belongs to one host/project, never to a parsed expression.
struct VbaRandomState {std::uint32_t seed{0x50000};};
VbaValue nextVbaRandom(VbaHost& host,const VbaValue& number);
void randomizeVbaGenerator(VbaHost& host,const VbaValue& number);
}
