#pragma once
namespace linaccess {
struct VbaValue;
// MS-VBAL5.6.16.3: resolve a scalar default once, Null=False, otherwise Let-Boolean.
bool vbaBooleanCondition(const VbaValue&);
}
