#pragma once

#include <string>
#include <string_view>
#include <vector>

namespace linaccess {
struct VbaProcedure {
    std::string kind;
    std::string name;
    int line{};
    bool eventCandidate{};
};

struct VbaAnalysis {
    std::vector<VbaProcedure> procedures;
    std::vector<std::string> compatibilityFlags;
};

// Static cataloging only. This is deliberately NOT a VBA interpreter/executor.
VbaAnalysis analyzeVba(std::string_view source);
// Accessors sharing a source name need distinct invocation and observation identities.
std::string vbaProcedureExecutionName(const VbaProcedure& procedure);
}
