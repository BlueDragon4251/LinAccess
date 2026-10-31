#include "vba_runtime.hpp"
#include "vba_expression.hpp"

#include <algorithm>
#include <cctype>
#include <cmath>
#include <regex>
#include <sstream>
#include <stdexcept>
#include <unordered_map>
#include <vector>

namespace linaccess {
namespace {
// This interpreter deliberately reports unsupported VBA instead of silently
// treating Access/DAO/COM calls as successful no-ops.
struct Statement {
    enum class Type { If, Else, EndIf, For, Next, Message, Save, Maximize, GoLast,
                      GoNext, GoNew, Requery,
                      ExecuteSql, OpenForm, ControlRequery, ControlFocus, ControlVisible,
                      Call, SetField, SetVar, Exit, Dim, ObjectSet, ObjectAssign, ObjectCall, While, EndWhile, Quit, ErrorResumeNext, ErrorDisable, Label, Jump, ErrorHandler, Resume, ResumeNext, ResumeLabel } type;
    std::string argument;
    std::string expression;
    int line{};
    std::string endExpression;
    std::string stepExpression;
    int partner{-1};
};

std::string trim(std::string text) {
    const auto first = text.find_first_not_of(" \t\r\n");
    if (first == std::string::npos) return {};
    const auto last = text.find_last_not_of(" \t\r\n");
    return text.substr(first, last - first + 1);
}
std::string lower(std::string text) {
    std::transform(text.begin(), text.end(), text.begin(), [](unsigned char c) {
        return static_cast<char>(std::tolower(c));
    });
    return text;
}
std::string withoutComment(const std::string& line) {
    std::string result;
    bool inString = false;
    for (std::size_t i = 0; i < line.size(); ++i) {
        const char ch = line[i];
        if (ch == '"') {
            if (inString && i + 1 < line.size() && line[i + 1] == '"') {
                result += "\"\"";
                ++i;
                continue;
            }
            inString = !inString;
        }
        if (!inString && ch == '\'') break;
        result.push_back(ch);
    }
    return trim(result);
}

class ErrorObject final : public VbaObject {
public:
    int number{};
    std::string description;
    VbaValue member(std::string_view name, const std::vector<VbaValue>& args) override {
        const auto key = lower(std::string(name));
        if (key == "raise") {
            if (args.size() != 1) throw VbaRuntimeError(5, "Err.Raise benötigt eine Fehlernummer");
            const auto code = args[0].asNumber();
            if (code <= 0 || code > 65535 || std::floor(code) != code)
                throw VbaRuntimeError(5, "Ungültige Fehlernummer");
            throw VbaRuntimeError(static_cast<int>(code), "VBA: explizit ausgelöster Fehler");
        }
        if (!args.empty()) throw VbaRuntimeError(5, "Unerwartete Err-Argumente");
        if (key == "number") return VbaValue::numeric(number);
        if (key == "description") return VbaValue::string(description);
        if (key == "clear") { number = 0; description.clear(); return {}; }
        throw VbaRuntimeError(438, "Err-Mitglied nicht unterstützt");
    }
};

struct Parameter { std::string name; bool byReference{true}; };
std::vector<std::string> splitArguments(const std::string& input) {
    std::vector<std::string> args;
    if (trim(input).empty()) return args;
    bool quoted = false;
    int parentheses = 0;
    std::size_t start = 0;
    for (std::size_t i = 0; i < input.size(); ++i) {
        const char ch = input[i];
        if (ch == '"') {
            if (quoted && i + 1 < input.size() && input[i + 1] == '"') { ++i; continue; }
            quoted = !quoted;
        }
        if (quoted) continue;
        if (ch == '(') ++parentheses;
        if (ch == ')') --parentheses;
        if (ch == ',' && parentheses == 0) { args.push_back(trim(input.substr(start, i - start))); start = i + 1; }
    }
    args.push_back(trim(input.substr(start)));
    return args;
}

VbaExecutionResult parseProcedure(std::string_view module, std::string_view name,
                                  std::vector<Statement>& statements, bool& isFunction, bool& hasCancel,
                                  std::vector<Parameter>* parameters = nullptr) {
    const std::regex begin(
        R"(^\s*(?:(?:Public|Private|Friend|Static)\s+)?(Sub|Function)\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*\(\s*(.*?)\s*\)\s*(?:As\s+[A-Za-z_][A-Za-z0-9_]*)?\s*$)",
        std::regex::icase);
    const std::regex anyBegin(
        R"(^\s*(?:(?:Public|Private|Friend|Static)\s+)?(?:Sub|Function)\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)(?:\s*\(.*\))?.*$)",
        std::regex::icase);
    const std::regex end(R"(^\s*End\s+(Sub|Function)\s*$)", std::regex::icase);
    const std::regex ifThen(R"(^If\s+(.+)\s+Then\s*$)", std::regex::icase);
    const std::regex assignment(R"(^Me\s*[!.]\s*([A-Za-z_][A-Za-z0-9_]*|\[[^\]]+\])\s*=\s*(.+)$)", std::regex::icase);
    const std::regex variableAssignment(R"(^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.+)$)", std::regex::icase);
    const std::regex message(R"(^MsgBox\s+(.+)$)", std::regex::icase);
    const std::regex save(R"(^DoCmd\s*\.\s*RunCommand\s+acCmdSaveRecord$)", std::regex::icase);
    const std::regex maximize(R"(^DoCmd\s*\.\s*Maximize\s*$)", std::regex::icase);
    const std::regex goLast(R"(^DoCmd\s*\.\s*GoToRecord\s*,\s*,\s*acLast\s*$)", std::regex::icase);
    const std::regex goNext(R"(^DoCmd\s*\.\s*GoToRecord\s*,\s*,\s*acNext\s*$)", std::regex::icase);
    const std::regex goNew(R"(^DoCmd\s*\.\s*GoToRecord\s*,\s*,\s*acNewRec\s*$)", std::regex::icase);
    const std::regex requery(R"(^DoCmd\s*\.\s*Requery\s*$)", std::regex::icase);
    const std::regex runSql(R"(^DoCmd\s*\.\s*RunSQL\s+(.+)$)", std::regex::icase);
    const std::regex daoExecute(R"(^CurrentDb(?:\s*\(\s*\))?\s*\.\s*Execute\s+(.+)$)", std::regex::icase);
    const std::regex openForm(R"(^DoCmd\s*\.\s*OpenForm\s+(.+)$)", std::regex::icase);
    const std::regex controlRequery(R"(^Me\s*[!.]\s*(\[[^\]]+\]|[A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*\.\s*Requery$)", std::regex::icase);
    const std::regex controlFocus(R"(^Me\s*[!.]\s*(\[[^\]]+\]|[A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*\.\s*SetFocus$)", std::regex::icase);
    const std::regex controlVisible(R"(^Me\s*[!.]\s*(\[[^\]]+\]|[A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*\.\s*Visible\s*=\s*(.+)$)", std::regex::icase);
    const std::regex call(R"(^Call\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)(?:\s*\((.*)\))?$)", std::regex::icase);
    const std::regex implicitCall(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)(?:\s+(.+)|\((.*)\))?$)", std::regex::icase);
    const std::regex parameter(R"(^(?:(ByRef|ByVal)\s+)?([A-Za-z_][A-Za-z0-9_]*)(?:\s+As\s+(?:Integer|Long|String|Boolean|Double|Single|Variant|Object|DAO\.Recordset|DAO\.Database))?$)", std::regex::icase);
    const std::regex cancelParam(R"(^Cancel\s+As\s+(?:Integer|Boolean)$)", std::regex::icase);
    const std::regex objectSet(R"(^Set\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.+)$)", std::regex::icase);
    const std::regex objectAssign(R"(^([A-Za-z_][A-Za-z0-9_]*(?:\([^)]*\))?(?:[.!][A-Za-z_][A-Za-z0-9_]*(?:\([^)]*\))?)*[.!](?:[A-Za-z_][A-Za-z0-9_]*|\[[^\]]+\]))\s*=\s*(.+)$)", std::regex::icase);
    const std::regex label(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*):$)");
    const std::regex jump(R"(^GoTo\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)$)", std::regex::icase);
    const std::regex errorHandler(R"(^On\s+Error\s+GoTo\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)$)", std::regex::icase);
    const std::regex resumeLabel(R"(^Resume\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)$)", std::regex::icase);
    const std::regex whileStart(R"(^Do\s+While\s+(.+)$)", std::regex::icase);
    const std::regex dim(R"(^Dim\s+(.+)$)", std::regex::icase);
    const std::regex declaration(R"(^([A-Za-z_][A-Za-z0-9_]*)(?:\s+As\s+([A-Za-z_][A-Za-z0-9_.]*))?$)", std::regex::icase);
    const std::regex forStart(R"(^For\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.+?)\s+To\s+(.+?)(?:\s+Step\s+(.+))?$)", std::regex::icase);
    const std::regex next(R"(^Next(?:\s+([A-Za-z_][A-Za-z0-9_]*))?$)", std::regex::icase);

    std::istringstream stream{std::string(module)};
    std::string raw;
    int line = 0;
    int nesting = 0;
    std::vector<bool> elseSeen;
    std::vector<Statement::Type> blocks;
    std::vector<int> forStarts;
    std::vector<int> whileStarts;
    bool within = false;
    std::unordered_map<std::string,int> labels;
    while (std::getline(stream, raw)) {
        ++line;
        // VBA line continuations join a statement before it is parsed.
        // This is essential for long SQL strings in Access 2003 modules.
        while (true) {
            const auto clean = withoutComment(raw);
            if (clean.size() < 2 || clean.back() != '_'
                || !std::isspace(static_cast<unsigned char>(clean[clean.size() - 2]))) break;
            raw = trim(clean.substr(0, clean.size() - 1));
            std::string continuation;
            if (!std::getline(stream, continuation))
                return {VbaExecutionStatus::Error, "VBA-Fortsetzungszeile fehlt", line};
            ++line;
            raw += ' ' + trim(continuation);
        }
        const auto source = withoutComment(raw);
        if (source.empty()) continue;
        std::smatch match;
        if (!within) {
            if (std::regex_match(source, match, begin) &&
                lower(match[2].str()) == lower(std::string(name))) {
                const auto params = trim(match[3].str());
                hasCancel = std::regex_match(params, cancelParam);
                for (const auto& declaration : splitArguments(params)) {
                    std::smatch declared;
                    if (!std::regex_match(declaration, declared, parameter))
                        return {VbaExecutionStatus::Unsupported, "VBA-Parameterdeklaration nicht unterstützt", line};
                    if (parameters) parameters->push_back({declared[2].str(), lower(declared[1].str()) != "byval"});
                }
                within = true;
                isFunction = lower(match[1].str()) == "function";
            } else if (std::regex_match(source, match, anyBegin)
                       && lower(match[1].str()) == lower(std::string(name))) {
                return {VbaExecutionStatus::Unsupported,
                    "VBA-Prozedursignatur/Parameter derzeit nicht unterstützt: " + source, line};
            }
            continue;
        }
        if (std::regex_match(source, match, end)) {
            if ((lower(match[1].str()) == "function") != isFunction)
                return {VbaExecutionStatus::Error, "End Sub/Function passt nicht", line};
            if (!blocks.empty()) return {VbaExecutionStatus::Error, "End If oder Next fehlt", line};
            for (auto& instruction : statements) {
                if (instruction.type != Statement::Type::Jump && instruction.type != Statement::Type::ErrorHandler
                    && instruction.type != Statement::Type::ResumeLabel) continue;
                const auto target = labels.find(lower(instruction.argument));
                if (target == labels.end()) return {VbaExecutionStatus::Error, "Sprungmarke fehlt", instruction.line};
                instruction.partner = target->second;
            }
            return {VbaExecutionStatus::Completed, {}, line};
        }
        if (std::regex_match(source, match, label)) {
            if (!blocks.empty()) return {VbaExecutionStatus::Unsupported, "Sprungmarke innerhalb eines Blocks", line};
            if (!labels.emplace(lower(match[1].str()), static_cast<int>(statements.size())).second)
                return {VbaExecutionStatus::Error, "Doppelte Sprungmarke", line};
            statements.push_back({Statement::Type::Label, {}, {}, line});
        } else if (std::regex_match(source, match, jump)) {
            statements.push_back({Statement::Type::Jump, match[1].str(), {}, line});
        } else if (std::regex_match(source, match, errorHandler)) {
            statements.push_back({Statement::Type::ErrorHandler, match[1].str(), {}, line});
        } else if (lower(source) == "resume next") {
            statements.push_back({Statement::Type::ResumeNext, {}, {}, line});
        } else if (lower(source) == "resume") {
            statements.push_back({Statement::Type::Resume, {}, {}, line});
        } else if (std::regex_match(source, match, resumeLabel)) {
            statements.push_back({Statement::Type::ResumeLabel, match[1].str(), {}, line});
        } else if (std::regex_match(source, match, ifThen)) {
            const auto expr = trim(match[1].str());
            if (!validVbaExpression(expr)) return {VbaExecutionStatus::Unsupported, "If-Bedingung: " + expr, line};
            statements.push_back({Statement::Type::If, expr, {}, line});
            ++nesting;
            elseSeen.push_back(false);
            blocks.push_back(Statement::Type::If);
        } else if (lower(source) == "else") {
            if (blocks.empty() || blocks.back() != Statement::Type::If)
                return {VbaExecutionStatus::Error, "Else ohne If", line};
            if (elseSeen.back()) return {VbaExecutionStatus::Error, "Doppeltes Else", line};
            elseSeen.back() = true;
            statements.push_back({Statement::Type::Else, {}, {}, line});
        } else if (lower(source) == "end if") {
            if (blocks.empty() || blocks.back() != Statement::Type::If)
                return {VbaExecutionStatus::Error, "End If ohne If", line};
            --nesting;
            elseSeen.pop_back();
            blocks.pop_back();
            statements.push_back({Statement::Type::EndIf, {}, {}, line});
        } else if (std::regex_match(source, match, forStart)) {
            const auto start = trim(match[2].str());
            const auto endValue = trim(match[3].str());
            const auto step = match[4].matched ? trim(match[4].str()) : "1";
            if (!validVbaExpression(start) || !validVbaExpression(endValue)
                || !validVbaExpression(step))
                return {VbaExecutionStatus::Unsupported, "For-Ausdruck: " + source, line};
            forStarts.push_back(static_cast<int>(statements.size()));
            blocks.push_back(Statement::Type::For);
            statements.push_back({Statement::Type::For, match[1].str(), start,
                                  line, endValue, step});
        } else if (std::regex_match(source, match, next)) {
            if (blocks.empty() || blocks.back() != Statement::Type::For)
                return {VbaExecutionStatus::Error, "Next ohne For", line};
            const int startIndex = forStarts.back();
            if (match[1].matched && lower(match[1].str()) != lower(statements[startIndex].argument))
                return {VbaExecutionStatus::Error, "Next-Variable passt nicht zum For", line};
            forStarts.pop_back();
            blocks.pop_back();
            const int nextIndex = static_cast<int>(statements.size());
            statements[startIndex].partner = nextIndex;
            statements.push_back({Statement::Type::Next, {}, {}, line});
            statements.back().partner = startIndex;
        } else if (std::regex_match(source, match, whileStart)) {
            const auto expr = trim(match[1].str());
            if (!validVbaExpression(expr)) return {VbaExecutionStatus::Unsupported, "Do While-Ausdruck", line};
            whileStarts.push_back(static_cast<int>(statements.size()));
            blocks.push_back(Statement::Type::While);
            statements.push_back({Statement::Type::While, {}, expr, line});
        } else if (lower(source) == "loop") {
            if (blocks.empty() || blocks.back() != Statement::Type::While)
                return {VbaExecutionStatus::Error, "Loop ohne Do While", line};
            const int start = whileStarts.back();
            whileStarts.pop_back(); blocks.pop_back();
            statements[start].partner = static_cast<int>(statements.size());
            statements.push_back({Statement::Type::EndWhile, {}, {}, line});
            statements.back().partner = start;
        } else if (lower(source) == "on error resume next") {
            statements.push_back({Statement::Type::ErrorResumeNext, {}, {}, line});
        } else if (lower(source) == "on error goto 0") {
            statements.push_back({Statement::Type::ErrorDisable, {}, {}, line});
        } else if (std::regex_match(source, match, objectSet)) {
            const auto expr = trim(match[2].str());
            if (!validVbaExpression(expr)) return {VbaExecutionStatus::Unsupported, "Set-Ausdruck", line};
            statements.push_back({Statement::Type::ObjectSet, match[1].str(), expr, line});
        } else if (std::regex_match(source, match, message)) {
            const auto expr = trim(match[1].str());
            if (!validVbaExpression(expr)) return {VbaExecutionStatus::Unsupported, "MsgBox-Ausdruck: " + expr, line};
            statements.push_back({Statement::Type::Message, {}, expr, line});
        } else if (lower(source) == "docmd.quit") {
            statements.push_back({Statement::Type::Quit, {}, {}, line});
        } else if (std::regex_match(source, save)) {
            statements.push_back({Statement::Type::Save, {}, {}, line});
        } else if (std::regex_match(source, maximize)) {
            statements.push_back({Statement::Type::Maximize, {}, {}, line});
        } else if (std::regex_match(source, goLast)) {
            statements.push_back({Statement::Type::GoLast, {}, {}, line});
        } else if (std::regex_match(source, goNext)) {
            statements.push_back({Statement::Type::GoNext, {}, {}, line});
        } else if (std::regex_match(source, goNew)) {
            statements.push_back({Statement::Type::GoNew, {}, {}, line});
        } else if (std::regex_match(source, requery)) {
            statements.push_back({Statement::Type::Requery, {}, {}, line});
        } else if (std::regex_match(source, match, runSql)
                   || std::regex_match(source, match, daoExecute)) {
            const auto expr = trim(match[1].str());
            if (!validVbaExpression(expr))
                return {VbaExecutionStatus::Unsupported, "SQL-Ausdruck: " + expr, line};
            statements.push_back({Statement::Type::ExecuteSql, {}, expr, line});
        } else if (std::regex_match(source, match, openForm)) {
            const auto expr = trim(match[1].str());
            if (!validVbaExpression(expr))
                return {VbaExecutionStatus::Unsupported, "Formularname: " + expr, line};
            statements.push_back({Statement::Type::OpenForm, {}, expr, line});
        } else if (std::regex_match(source, match, controlRequery)) {
            auto name = match[1].str();
            if (name.starts_with('[') && name.ends_with(']')) name = name.substr(1, name.size() - 2);
            statements.push_back({Statement::Type::ControlRequery, name, {}, line});
        } else if (std::regex_match(source, match, controlFocus)) {
            auto name = match[1].str();
            if (name.starts_with('[') && name.ends_with(']')) name = name.substr(1, name.size() - 2);
            statements.push_back({Statement::Type::ControlFocus, name, {}, line});
        } else if (std::regex_match(source, match, controlVisible)) {
            auto name = match[1].str();
            if (name.starts_with('[') && name.ends_with(']')) name = name.substr(1, name.size() - 2);
            const auto expr = trim(match[2].str());
            if (!validVbaExpression(expr))
                return {VbaExecutionStatus::Unsupported, "Visible-Ausdruck: " + expr, line};
            statements.push_back({Statement::Type::ControlVisible, name, expr, line});
        } else if (std::regex_match(source, match, call)) {
            const auto args = match[2].str();
            for (const auto& arg : splitArguments(args))
                if (!validVbaExpression(arg)) return {VbaExecutionStatus::Unsupported, "Aufrufargument", line};
            statements.push_back({Statement::Type::Call, match[1].str(), args, line});
        } else if (std::regex_match(source, match, assignment)) {
            const auto expr = trim(match[2].str());
            if (!validVbaExpression(expr)) return {VbaExecutionStatus::Unsupported, "Zuweisung: " + expr, line};
            auto field = match[1].str();
            if (field.starts_with('[') && field.ends_with(']')) field = field.substr(1, field.size() - 2);
            statements.push_back({Statement::Type::SetField, field, expr, line});
        } else if (std::regex_match(source, match, objectAssign)) {
            const auto target = match[1].str();
            const auto expr = trim(match[2].str());
            if (!validVbaExpression(target) || !validVbaExpression(expr))
                return {VbaExecutionStatus::Unsupported, "Objektzuweisung", line};
            statements.push_back({Statement::Type::ObjectAssign, target, expr, line});
        } else if (std::regex_match(source, match, dim)) {
            std::istringstream declarations(match[1].str());
            std::string entry;
            while (std::getline(declarations, entry, ',')) {
                std::smatch declared;
                entry = trim(entry);
                if (!std::regex_match(entry, declared, declaration))
                    return {VbaExecutionStatus::Unsupported, "Variablendeklaration", line};
                const auto type = lower(declared[2].str());
                if (!type.empty() && type != "variant" && type != "object" && type != "dao.database"
                    && type != "dao.recordset" && type != "string" && type != "boolean"
                    && type != "integer" && type != "long" && type != "single" && type != "double"
                    && type != "byte") return {VbaExecutionStatus::Unsupported, "Variablentyp: " + type, line};
                statements.push_back({Statement::Type::Dim, declared[1].str(), type, line});
            }
        } else if (std::regex_match(source, match, variableAssignment)) {
            const auto expr = trim(match[2].str());
            if (!validVbaExpression(expr)) return {VbaExecutionStatus::Unsupported, "Zuweisung: " + expr, line};
            statements.push_back({Statement::Type::SetVar, match[1].str(), expr, line});
        } else if (lower(source) == "exit sub" || lower(source) == "exit function") {
            statements.push_back({Statement::Type::Exit, {}, {}, line});
        } else if (source.find('.') != std::string::npos && validVbaExpression(source)) {
            statements.push_back({Statement::Type::ObjectCall, {}, source, line});
        } else if (std::regex_match(source, match, implicitCall)) {
            const auto args = match[2].matched ? match[2].str() : match[3].str();
            for (const auto& arg : splitArguments(args))
                if (!validVbaExpression(arg)) return {VbaExecutionStatus::Unsupported, "Aufrufargument", line};
            statements.push_back({Statement::Type::Call, match[1].str(), args, line});
        } else {
            return {VbaExecutionStatus::Unsupported, "VBA-Anweisung noch nicht unterstützt: " + source, line};
        }
    }
    return within ? VbaExecutionResult{VbaExecutionStatus::Error, "End Sub fehlt", line}
                  : VbaExecutionResult{VbaExecutionStatus::ProcedureMissing, {}, 0};
}
}

VbaExecutionResult validateVbaProcedure(std::string_view module, std::string_view procedure) {
    if (module.size() > 1024 * 1024)
        return {VbaExecutionStatus::Unsupported, "VBA-Modul größer als 1 MiB", 0};
    std::vector<Statement> instructions;
    bool function = false, cancel = false;
    return parseProcedure(module, procedure, instructions, function, cancel);
}

static VbaExecutionResult executeProcedure(std::string_view module,
                                       std::string_view procedure,
                                       VbaHost& host, std::vector<VbaValue>& arguments,
                                       const std::vector<VbaValue*>* references = nullptr) {
    // Independent of host type, recursive VBA calls must not overflow the
    // native C++ stack. Access' own runtime also terminates stack overflow.
    static thread_local int callDepth = 0;
    if (callDepth >= 48)
        return {VbaExecutionStatus::Error, "VBA: Aufrufstapel erschöpft", 0};
    struct DepthGuard {
        int& depth;
        explicit DepthGuard(int& ref) : depth(ref) { ++depth; }
        ~DepthGuard() { --depth; }
    } guard(callDepth);
    if (module.size() > 1024 * 1024)
        return {VbaExecutionStatus::Unsupported, "VBA-Modul größer als 1 MiB", 0};
    std::vector<Statement> instructions;
    bool isFunction = false;
    bool hasCancel = false;
    std::vector<Parameter> parameters;
    auto result = parseProcedure(module, procedure, instructions, isFunction, hasCancel, &parameters);
    if (result.status != VbaExecutionStatus::Completed) return result;
    // Resolve direct Sub calls before any effects in the calling procedure.
    // Syntax acceptance alone must not turn unknown APIs into successful calls.
    try {
        for (const auto& instruction : instructions) {
            if (instruction.type != Statement::Type::Call) continue;
            auto declaration = validateVbaProcedure(module, instruction.argument);
            if (declaration.status == VbaExecutionStatus::ProcedureMissing) {
                const auto external = host.moduleForProcedure(instruction.argument);
                if (external) declaration = validateVbaProcedure(*external, instruction.argument);
            }
            if (declaration.status != VbaExecutionStatus::Completed)
                return {VbaExecutionStatus::Unsupported, "Aufgerufene Prozedur nicht verfügbar oder nicht unterstützt", instruction.line};
        }
    } catch (const std::exception& exception) {
        return {VbaExecutionStatus::Error, exception.what(), 0};
    }
    if (instructions.size() > 20000)
        return {VbaExecutionStatus::Unsupported, "Prozedur überschreitet 20.000 Anweisungen", 0};
    struct Frame { bool parent; bool condition; bool inElse; };
    struct LoopFrame { int forIndex; double end; double step; };
    std::vector<Frame> frames;
    std::vector<LoopFrame> loops;
    VbaVariables locals;
    auto errorState = std::make_shared<ErrorObject>();
    locals["err"] = VbaValue::reference(errorState);
    bool resumeNext = false;
    int handler = -1;
    bool handlingError = false;
    std::size_t faultInstruction = 0;
    std::vector<Frame> faultFrames;
    std::vector<LoopFrame> faultLoops;
    bool faultActive = true;
    if (isFunction) locals[lower(std::string(procedure))] = VbaValue{};
    const bool implicitCancel = hasCancel && arguments.empty();
    if (!implicitCancel && parameters.size() != arguments.size())
        return {VbaExecutionStatus::Error, "Falsche Argumentanzahl", 0};
    for (std::size_t i = 0; i < parameters.size(); ++i)
        locals[lower(parameters[i].name)] = implicitCancel ? VbaValue::numeric(0) : arguments[i];
    struct Alias { std::string name; VbaValue* source; VbaValue before; };
    std::vector<Alias> aliases;
    if (references && !implicitCancel)
        for (std::size_t i = 0; i < parameters.size(); ++i)
            if (parameters[i].byReference && (*references)[i])
                aliases.push_back({lower(parameters[i].name), (*references)[i], *(*references)[i]});
    const auto equalValue = [](const VbaValue& a, const VbaValue& b) {
        return a.kind == b.kind && a.number == b.number && a.text == b.text
            && a.boolean == b.boolean && a.object == b.object;
    };
    const auto refreshAliases = [&] {
        for (auto& alias : aliases) { locals[alias.name] = *alias.source; alias.before = *alias.source; }
    };
    const auto storeAliases = [&] {
        for (auto& alias : aliases)
            if (!equalValue(locals[alias.name], alias.before)) *alias.source = locals[alias.name];
        refreshAliases();
    };
    const auto copyArguments = [&] {
        if (!implicitCancel)
            for (std::size_t i = 0; i < parameters.size(); ++i)
                if (parameters[i].byReference) arguments[i] = locals[lower(parameters[i].name)];
    };
    auto completed = [&](int line, bool exited) {
        VbaExecutionResult outcome{VbaExecutionStatus::Completed, {}, line, exited};
        if (isFunction) {
            const auto value = locals[lower(std::string(procedure))];
            outcome.typedReturnValue = std::make_shared<VbaValue>(value);
            if (value.kind != VbaValue::Kind::Object) outcome.returnValue = value.asField();
        }
        storeAliases();
        copyArguments();
        if (hasCancel) outcome.cancelled = locals["cancel"].truthy();
        return outcome;
    };
    const auto invoke = [&](std::string_view name, std::vector<VbaValue>& actuals,
                            const std::vector<std::string>& names, bool needsValue) -> VbaValue {
        std::vector<VbaValue*> bindings;
        for (const auto& expression : names) {
            auto found = locals.find(lower(expression));
            VbaValue* binding = found == locals.end() ? nullptr : &found->second;
            for (const auto& alias : aliases) if (alias.name == lower(expression)) binding = alias.source;
            bindings.push_back(binding);
        }
        auto nested = executeProcedure(module, name, host, actuals, &bindings);
        if (nested.status == VbaExecutionStatus::ProcedureMissing) {
            const auto external = host.moduleForProcedure(name);
            if (external) nested = executeProcedure(*external, name, host, actuals, &bindings);
        }
        for (std::size_t i = 0; i < names.size(); ++i) {
            auto found = locals.find(lower(names[i]));
            if (found != locals.end()) found->second = actuals[i];
        }
        if (nested.status != VbaExecutionStatus::Completed)
            throw VbaRuntimeError(nested.errorNumber ? nested.errorNumber : 5,
                nested.status == VbaExecutionStatus::ProcedureMissing ? "Prozedur nicht gefunden" : nested.explanation);
        if (needsValue && !nested.typedReturnValue) throw VbaRuntimeError(5, "Sub hat keinen Rückgabewert");
        return nested.typedReturnValue ? *nested.typedReturnValue : VbaValue::empty();
    };
    const VbaFunctionResolver resolver = [&](std::string_view name, std::vector<VbaValue>& args,
                                            const std::vector<std::string>& refs) {
        return invoke(name, args, refs, true);
    };
    const auto evaluate = [&](std::string_view expression) {
        return evaluateVbaExpression(expression, host, locals, resolver);
    };
    bool active = true;
    int currentLine = 0;
    int steps = 0;
    int effects = 0;
    try {
        for (std::size_t ip = 0; ip < instructions.size(); ++ip) {
            const auto& instruction = instructions[ip];
            currentLine = instruction.line;
            if (++steps > 10000)
                return {VbaExecutionStatus::Error, "VBA-Ausführungsbudget (10.000 Schritte) erreicht", currentLine};
            refreshAliases();
            try {
            switch (instruction.type) {
            case Statement::Type::ErrorResumeNext:
                if (active) { resumeNext = true; handler = -1; }
                break;
            case Statement::Type::ErrorDisable:
                if (active) { resumeNext = false; handler = -1; }
                break;
            case Statement::Type::Label:
                break;
            case Statement::Type::Jump:
                if (active) { frames.clear(); loops.clear(); active = true; ip = static_cast<std::size_t>(instruction.partner); }
                break;
            case Statement::Type::ErrorHandler:
                if (active) { handler = instruction.partner; resumeNext = false; }
                break;
            case Statement::Type::Resume:
            case Statement::Type::ResumeNext:
            case Statement::Type::ResumeLabel:
                if (active) {
                    if (!handlingError) throw VbaRuntimeError(20, "Resume ohne aktiven Fehler");
                    handlingError = false;
                    errorState->number = 0; errorState->description.clear();
                    if (instruction.type == Statement::Type::ResumeLabel) {
                        frames.clear(); loops.clear(); active = true;
                        ip = static_cast<std::size_t>(instruction.partner);
                    } else {
                        frames = faultFrames; loops = faultLoops; active = faultActive;
                        ip = instruction.type == Statement::Type::Resume ? faultInstruction - 1 : faultInstruction;
                    }
                }
                break;
            case Statement::Type::If: {
                const bool matched = active
                    ? evaluate(instruction.argument).truthy() : false;
                frames.push_back({active, matched, false});
                active = active && matched;
                break;
            }
            case Statement::Type::Else:
                if (frames.empty() || frames.back().inElse)
                    return {VbaExecutionStatus::Error, "Else doppelt oder ohne If", instruction.line};
                frames.back().inElse = true;
                active = frames.back().parent && !frames.back().condition;
                break;
            case Statement::Type::EndIf:
                if (frames.empty())
                    return {VbaExecutionStatus::Error, "End If ohne If", instruction.line};
                active = frames.back().parent;
                frames.pop_back();
                break;
            case Statement::Type::Message:
                if (active) {
                    if (++effects > 1000) throw std::runtime_error("VBA-Effektbudget erreicht");
                    host.message(evaluate(instruction.expression).asText());
                }
                break;
            case Statement::Type::Quit:
                if (active) { host.quitApplication(); return completed(instruction.line, true); }
                break;
            case Statement::Type::Save:
                if (active) {
                    if (++effects > 1000) throw std::runtime_error("VBA-Effektbudget erreicht");
                    host.saveRecord();
                }
                break;
            case Statement::Type::Maximize:
                if (active) {
                    if (++effects > 1000) throw std::runtime_error("VBA-Effektbudget erreicht");
                    host.maximizeForm();
                }
                break;
            case Statement::Type::GoLast:
                if (active) {
                    if (++effects > 1000) throw std::runtime_error("VBA-Effektbudget erreicht");
                    host.gotoLastRecord();
                }
                break;
            case Statement::Type::GoNext:
                if (active) host.gotoNextRecord();
                break;
            case Statement::Type::GoNew:
                if (active) host.gotoNewRecord();
                break;
            case Statement::Type::Requery:
                if (active) host.requery();
                break;
            case Statement::Type::ExecuteSql:
                if (active) {
                    if (++effects > 1000) throw std::runtime_error("VBA-Effektbudget erreicht");
                    host.executeSql(evaluate(instruction.expression).asText());
                }
                break;
            case Statement::Type::OpenForm:
                if (active) {
                    if (++effects > 1000) throw std::runtime_error("VBA-Effektbudget erreicht");
                    host.openForm(evaluate(instruction.expression).asText());
                }
                break;
            case Statement::Type::ControlRequery:
                if (active) host.requeryControl(instruction.argument);
                break;
            case Statement::Type::ControlFocus:
                if (active) host.focusControl(instruction.argument);
                break;
            case Statement::Type::ControlVisible:
                if (active) host.setControlVisible(instruction.argument,
                    evaluate(instruction.expression).truthy());
                break;
            case Statement::Type::Call:
                if (active) {
                    const auto expressions = splitArguments(instruction.expression);
                    std::vector<VbaValue> actuals;
                    for (const auto& expression : expressions) actuals.push_back(evaluate(expression));
                    (void)invoke(instruction.argument, actuals, expressions, false);
                }
                break;
            case Statement::Type::SetField:
                if (active) {
                    if (++effects > 1000) throw std::runtime_error("VBA-Effektbudget erreicht");
                    host.setField(instruction.argument,
                        evaluate(instruction.expression).asField());
                }
                break;
            case Statement::Type::ObjectSet:
                if (active) {
                    auto value = evaluate(instruction.expression);
                    if (value.kind != VbaValue::Kind::Object) throw std::runtime_error("Set benötigt ein Objekt");
                    locals[lower(instruction.argument)] = std::move(value);
                }
                break;
            case Statement::Type::ObjectAssign:
                if (active) {
                    const auto pos = instruction.argument.find_last_of(".!");
                    auto name = instruction.argument.substr(pos + 1);
                    if (name.starts_with('[') && name.ends_with(']')) name = name.substr(1, name.size() - 2);
                    const auto object = evaluate(instruction.argument.substr(0, pos));
                    if (!object.object) throw std::runtime_error("Objektvariable nicht gesetzt");
                    object.object->assign(name, evaluate(instruction.expression));
                }
                break;
            case Statement::Type::ObjectCall:
                if (active) (void)evaluate(instruction.expression);
                break;
            case Statement::Type::While:
                if (!active || !evaluate(instruction.expression).truthy())
                    ip = static_cast<std::size_t>(instruction.partner);
                break;
            case Statement::Type::EndWhile:
                if (active) ip = static_cast<std::size_t>(instruction.partner) - 1;
                break;
            case Statement::Type::SetVar:
                if (active) locals[lower(instruction.argument)] =
                    evaluate(instruction.expression);
                break;
            case Statement::Type::For: {
                if (!active) {
                    ip = static_cast<std::size_t>(instruction.partner);
                    break;
                }
                const auto initial = evaluate(instruction.expression).asNumber();
                const auto end = evaluate(instruction.endExpression).asNumber();
                const auto step = evaluate(instruction.stepExpression).asNumber();
                if (!std::isfinite(initial) || !std::isfinite(end)
                    || !std::isfinite(step) || step == 0)
                    throw std::runtime_error("Ungültige For-Schrittweite/Grenze");
                locals[lower(instruction.argument)] = VbaValue::numeric(initial);
                if ((step > 0 && initial > end) || (step < 0 && initial < end)) {
                    ip = static_cast<std::size_t>(instruction.partner);
                } else loops.push_back({static_cast<int>(ip), end, step});
                break;
            }
            case Statement::Type::Next: {
                if (!active) break;
                if (loops.empty() || loops.back().forIndex != instruction.partner)
                    throw std::runtime_error("VBA-For/Next-Laufzeitstapel ungültig");
                const auto& start = instructions[loops.back().forIndex];
                const auto name = lower(start.argument);
                const auto nextValue = locals[name].asNumber() + loops.back().step;
                if (!std::isfinite(nextValue)) throw std::runtime_error("For-Zählerüberlauf");
                locals[name] = VbaValue::numeric(nextValue);
                if ((loops.back().step > 0 && nextValue <= loops.back().end)
                    || (loops.back().step < 0 && nextValue >= loops.back().end))
                    ip = static_cast<std::size_t>(loops.back().forIndex);
                else loops.pop_back();
                break;
            }
            case Statement::Type::Exit:
                if (active) return completed(instruction.line, true);
                break;
            case Statement::Type::Dim:
                if (active) {
                    const auto& type = instruction.expression;
                    auto initial = VbaValue::empty();
                    if (type == "object" || type.starts_with("dao.")) initial = VbaValue::reference(nullptr);
                    else if (type == "string") initial = VbaValue::string("");
                    else if (type == "boolean") initial = VbaValue::logical(false);
                    else if (!type.empty() && type != "variant") initial = VbaValue::numeric(0);
                    locals.try_emplace(lower(instruction.argument), std::move(initial));
                }
                break;
            }
            storeAliases();
            } catch (const std::exception& exception) {
                storeAliases();
                if (!resumeNext && (handler < 0 || handlingError)) throw;
                const auto* typed = dynamic_cast<const VbaRuntimeError*>(&exception);
                errorState->number = typed ? typed->number : 5;
                errorState->description = exception.what();
                if (!resumeNext) {
                    handlingError = true; faultInstruction = ip;
                    faultFrames = frames; faultLoops = loops; faultActive = active;
                    frames.clear(); loops.clear(); active = true;
                    ip = static_cast<std::size_t>(handler);
                }
            }
        }
    } catch (const std::exception& exc) {
        storeAliases();
        copyArguments();
        VbaExecutionResult failure{VbaExecutionStatus::Error, exc.what(), currentLine};
        const auto* typed = dynamic_cast<const VbaRuntimeError*>(&exc);
        failure.errorNumber = typed ? typed->number : 5;
        return failure;
    }
    return completed(0, false);
}
VbaExecutionResult executeVbaProcedureWithArguments(std::string_view module, std::string_view procedure,
                                                    VbaHost& host, std::vector<VbaValue>& arguments) {
    return executeProcedure(module, procedure, host, arguments);
}
VbaExecutionResult executeVbaProcedure(std::string_view module, std::string_view procedure, VbaHost& host) {
    std::vector<VbaValue> arguments;
    return executeVbaProcedureWithArguments(module, procedure, host, arguments);
}
}  // namespace linaccess
