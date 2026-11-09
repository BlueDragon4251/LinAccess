#include "vba_analyzer.hpp"

#include <algorithm>
#include <cctype>
#include <regex>
#include <sstream>

namespace linaccess {
std::string vbaProcedureExecutionName(const VbaProcedure& procedure) {
    std::istringstream kind(procedure.kind);std::string first,second;kind>>first>>second;
    for(auto* word:{&first,&second}) for(auto& ch:*word) ch=static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
    if(first=="property" && (second=="get" || second=="let" || second=="set")) return "property "+second+" "+procedure.name;
    return procedure.name;
}
VbaAnalysis analyzeVba(std::string_view source) {
    VbaAnalysis analysis;
    std::istringstream input{std::string(source)};
    std::string line;
    const std::regex declaration(
        // UTF-8 non-ASCII bytes also occur in legal VBA identifiers, e.g.
        // German form event handlers such as Mengenänderung_Click.
        R"(^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*(?:\(|As\b|Static\b|$))",
        std::regex_constants::icase);
    const std::vector<std::pair<std::string, std::regex>> dependencies = {
        {"Windows-API/Declare", std::regex(R"(\bDeclare\s+(?:PtrSafe\s+)?(?:Function|Sub)\b)", std::regex_constants::icase)},
        {"COM-Automation", std::regex(R"(\b(?:CreateObject|GetObject)\s*\()", std::regex_constants::icase)},
        {"Access DoCmd", std::regex(R"(\bDoCmd\s*\.)", std::regex_constants::icase)},
        {"Access CurrentDb/DAO", std::regex(R"(\b(?:CurrentDb\b|DAO\.))", std::regex_constants::icase)},
        {"Access Forms/Reports", std::regex(R"(\b(?:Forms|Reports)\s*[!\(])", std::regex_constants::icase)},
        {"ADO", std::regex(R"(\bADODB\.)", std::regex_constants::icase)},
    };
    int number = 0;
    while (std::getline(input, line)) {
        ++number;
        // Single-quote comments are common in VBA. Discard them outside string literals.
        std::string uncommented;
        bool inString = false;
        for (std::size_t i = 0; i < line.size(); ++i) {
            const char c = line[i];
            if (c == '"') {
                if (inString && i + 1 < line.size() && line[i + 1] == '"') {
                    uncommented += "\"\"";
                    ++i;
                    continue;
                }
                inString = !inString;
            }
            if (c == '\'' && !inString)
                break;
            uncommented.push_back(c);
        }
        std::smatch match;
        if (std::regex_search(uncommented, match, declaration)) {
            const auto name = match[2].str();
            analysis.procedures.push_back({match[1].str(), name, number,
                                           name.find('_') != std::string::npos});
        }
        // Ignore string literals when looking for potentially unsupported APIs.
        // Seeing "DoCmd.OpenForm" inside MsgBox is not a real dependency.
        std::string outsideStrings = uncommented;
        inString = false;
        for (std::size_t i = 0; i < outsideStrings.size(); ++i) {
            if (outsideStrings[i] == '"') {
                if (inString && i + 1 < outsideStrings.size() && outsideStrings[i + 1] == '"') {
                    outsideStrings[i++] = ' ';
                    outsideStrings[i] = ' ';
                    continue;
                }
                inString = !inString;
            }
            if (inString || outsideStrings[i] == '"') outsideStrings[i] = ' ';
        }
        for (const auto& [label, pattern] : dependencies) {
            if (std::regex_search(outsideStrings, pattern) &&
                std::find(analysis.compatibilityFlags.begin(), analysis.compatibilityFlags.end(), label)
                    == analysis.compatibilityFlags.end()) {
                analysis.compatibilityFlags.push_back(label);
            }
        }
    }
    return analysis;
}
}
