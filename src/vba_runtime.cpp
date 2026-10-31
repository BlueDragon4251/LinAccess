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
    enum class Type { If, Else, EndIf, For, ForEach, ExitFor, Next, Message, Save, Maximize, GoLast,
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
std::size_t memberSeparator(const std::string& expression) {
    std::size_t result = std::string::npos;
    bool quoted = false, bracket = false;
    int depth = 0;
    for (std::size_t i = 0; i < expression.size(); ++i) {
        const char ch = expression[i];
        if (bracket) { if (ch == ']') bracket = false; continue; }
        if (ch == '"') {
            if (quoted && i + 1 < expression.size() && expression[i+1] == '"') { ++i; continue; }
            quoted = !quoted; continue;
        }
        if (quoted) continue;
        if (ch == '[') bracket = true;
        else if (ch == '(') ++depth;
        else if (ch == ')') --depth;
        else if (depth == 0 && (ch == '.' || ch == '!')) result = i;
    }
    return result;
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
    VbaValue scalarValue() override { return VbaValue::numeric(number); }
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

std::size_t outsideKeyword(const std::string& source,std::string_view keyword,std::size_t start=0) {
    bool quoted=false,bracket=false; int depth=0;
    for (std::size_t i=0;i<source.size();++i) {
        const char ch=source[i];
        if (quoted) { if (ch=='"') { if (i+1<source.size() && source[i+1]=='"') ++i; else quoted=false; } continue; }
        if (bracket) { if (ch==']') { if (i+1<source.size() && source[i+1]==']') ++i; else bracket=false; } continue; }
        if (ch=='"') { quoted=true; continue; }
        if (ch=='[') { bracket=true; continue; }
        if (ch=='(') ++depth;
        else if (ch==')') --depth;
        if (i<start || depth!=0 || i+keyword.size()>source.size()) continue;
        const auto part=lower(source.substr(i,keyword.size()));
        const auto identifier=[](char value) { return std::isalnum(static_cast<unsigned char>(value)) || static_cast<unsigned char>(value)>=0x80 || value=='_'; };
        if (part==keyword && (i==0 || !identifier(source[i-1])) && (i+keyword.size()==source.size() || !identifier(source[i+keyword.size()]))) return i;
    }
    return std::string::npos;
}
std::size_t inlineElse(const std::string& source,std::size_t start) {
    int nested=0;
    while (start<source.size()) {
        const auto condition=outsideKeyword(source,"if",start),otherwise=outsideKeyword(source,"else",start);
        if (otherwise==std::string::npos) return otherwise;
        if (condition!=std::string::npos && condition<otherwise) { ++nested; start=condition+2; }
        else if (nested==0) return otherwise;
        else { --nested; start=otherwise+4; }
    }
    return std::string::npos;
}
std::string inlineStatementLines(const std::string& source) {
    std::string result=source; bool quoted=false,bracket=false; int depth=0;
    for (std::size_t i=0;i<result.size();++i) {
        const char ch=result[i];
        if (quoted) { if (ch=='"') { if (i+1<result.size() && result[i+1]=='"') ++i; else quoted=false; } continue; }
        if (bracket) { if (ch==']') { if (i+1<result.size() && result[i+1]==']') ++i; else bracket=false; } continue; }
        if (ch=='"') quoted=true; else if (ch=='[') bracket=true;
        else if (ch=='(') ++depth; else if (ch==')') --depth;
        else if (ch==':' && depth==0) result[i]='\n';
    }
    return result;
}

VbaExecutionResult parseProcedure(std::string_view module, std::string_view name,
                                  std::vector<Statement>& statements, bool& isFunction, bool& hasCancel,
                                  std::vector<Parameter>* parameters = nullptr,bool resolveLabels=true,std::string* resultType=nullptr) {
    const std::regex begin(
        R"(^\s*(?:(?:Public|Private|Friend|Static)\s+)?(Sub|Function)\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*\(\s*(.*?)\s*\)\s*(?:As\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*))?\s*$)",
        std::regex::icase);
    const std::regex anyBegin(
        R"(^\s*(?:(?:Public|Private|Friend|Static)\s+)?(?:Sub|Function)\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)(?:\s*\(.*\))?.*$)",
        std::regex::icase);
    const std::regex end(R"(^\s*End\s+(Sub|Function)\s*$)", std::regex::icase);
    const std::regex ifThen(R"(^If\s+(.+)\s+Then\s*$)", std::regex::icase);
    const std::regex assignment(R"(^Me\s*([!.])\s*([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*|\[[^\]]+\])\s*=\s*(.+)$)", std::regex::icase);
    const std::regex variableAssignment(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*=\s*(.+)$)", std::regex::icase);
    const std::regex message(R"(^MsgBox(?:\s+(.+)|\((.*)\))$)", std::regex::icase);
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
    const std::regex memberCall(R"(^((?:[A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*(?:\([^)]*\))?)(?:[.!][A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*(?:\([^)]*\))?)+)\s+(.+)$)",std::regex::icase);
    const std::regex parameter(R"(^(?:(ByRef|ByVal)\s+)?([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)(?:\s+As\s+(?:Integer|Long|String|Boolean|Double|Single|Variant|Object|DAO\.Recordset|DAO\.Database|DAO\.TableDef|DAO\.TableDefs))?$)", std::regex::icase);
    const std::regex cancelParam(R"(^Cancel\s+As\s+(?:Integer|Boolean)$)", std::regex::icase);
    const std::regex objectSet(R"(^Set\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*=\s*(.+)$)", std::regex::icase);
    const std::regex objectAssign(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*(?:\([^)]*\))*(?:[.!](?:[A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*|\[[^\]]+\])(?:\([^)]*\))*)*[.!](?:[A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*|\[[^\]]+\]))\s*=\s*(.+)$)", std::regex::icase);
    const std::regex label(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*):$)");
    const std::regex jump(R"(^GoTo\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)$)", std::regex::icase);
    const std::regex errorHandler(R"(^On\s+Error\s+GoTo\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)$)", std::regex::icase);
    const std::regex resumeLabel(R"(^Resume\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)$)", std::regex::icase);
    const std::regex whileStart(R"(^Do\s+While\s+(.+)$)", std::regex::icase);
    const std::regex dim(R"(^Dim\s+(.+)$)", std::regex::icase);
    const std::regex declaration(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)(?:\s+As\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_.\x80-\xFF]*))?$)", std::regex::icase);
    const std::regex forStart(R"(^For\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*=\s*(.+?)\s+To\s+(.+?)(?:\s+Step\s+(.+))?$)", std::regex::icase);
    const std::regex forEach(R"(^For\s+Each\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s+In\s+(.+)$)", std::regex::icase);
    const std::regex next(R"(^Next(?:\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*))?$)", std::regex::icase);

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
                const auto declaredResult=lower(match[4].str());
                if (isFunction && !declaredResult.empty() && declaredResult!="variant" && declaredResult!="boolean" && declaredResult!="string" && declaredResult!="integer" && declaredResult!="long" && declaredResult!="single" && declaredResult!="double" && declaredResult!="byte" && declaredResult!="object")
                    return {VbaExecutionStatus::Unsupported,"Function result type unsupported",line};
                if (resultType) *resultType=declaredResult;
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
            // Single-line branches are parsed in a temporary wrapper. An
            // Exit For there binds to the enclosing real procedure's loop.
            if (resolveLabels) {
                int forDepth=0;
                for (const auto& instruction:statements) {
                    if (instruction.type==Statement::Type::For || instruction.type==Statement::Type::ForEach) ++forDepth;
                    else if (instruction.type==Statement::Type::Next) --forDepth;
                    else if (instruction.type==Statement::Type::ExitFor && forDepth==0)
                        return {VbaExecutionStatus::Error,"Exit For without For",instruction.line};
                }
            }
            if (resolveLabels) for (auto& instruction : statements) {
                if (instruction.type != Statement::Type::Jump && instruction.type != Statement::Type::ErrorHandler
                    && instruction.type != Statement::Type::ResumeLabel) continue;
                const auto target = labels.find(lower(instruction.argument));
                if (target == labels.end()) return {VbaExecutionStatus::Error, "Sprungmarke fehlt", instruction.line};
                instruction.partner = target->second;
            }
            return {VbaExecutionStatus::Completed, {}, line};
        }
        if (lower(source).starts_with("if ") || lower(source).starts_with("if\t")) {
            const auto then=outsideKeyword(source,"then",2);
            if (then!=std::string::npos && !trim(source.substr(then+4)).empty()) {
                static thread_local int inlineDepth=0;
                if (inlineDepth>=16) return {VbaExecutionStatus::Unsupported,"Single-line If nesting limit",line};
                struct Guard { int& depth; Guard(int& value):depth(value) { ++depth; } ~Guard() { --depth; } } guard(inlineDepth);
                const auto otherwise=inlineElse(source,then+4);
                const auto condition=trim(source.substr(2,then-2));
                const auto yes=trim(source.substr(then+4,otherwise==std::string::npos ? std::string::npos : otherwise-then-4));
                const auto no=otherwise==std::string::npos ? std::string{} : trim(source.substr(otherwise+4));
                if (yes.empty() || (otherwise!=std::string::npos && no.empty())) return {VbaExecutionStatus::Unsupported,"Single-line If branch missing",line};
                const auto wrapper="Sub __la_inline()\nIf "+condition+" Then\n"+inlineStatementLines(yes)
                    +(otherwise==std::string::npos ? "" : "\nElse\n"+inlineStatementLines(no))+"\nEnd If\nEnd Sub\n";
                std::vector<Statement> body; bool function=false,cancel=false;
                auto parsed=parseProcedure(wrapper,"__la_inline",body,function,cancel,nullptr,false);
                if (parsed.status!=VbaExecutionStatus::Completed) { parsed.line=line; return parsed; }
                const auto offset=static_cast<int>(statements.size());
                for (auto& statement:body) { if (statement.partner>=0) statement.partner+=offset; statement.line=line; statements.push_back(std::move(statement)); }
                continue;
            }
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
        } else if (std::regex_match(source, match, forEach)) {
            const auto expression = trim(match[2].str());
            if (!validVbaExpression(expression))
                return {VbaExecutionStatus::Unsupported,"For Each collection expression",line};
            forStarts.push_back(static_cast<int>(statements.size()));
            blocks.push_back(Statement::Type::For);
            statements.push_back({Statement::Type::ForEach,match[1].str(),expression,line});
        } else if (lower(source) == "exit for") {
            if (forStarts.empty() && resolveLabels) return {VbaExecutionStatus::Error,"Exit For without For",line};
            statements.push_back({Statement::Type::ExitFor,{},{},line});
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
            const auto expr = "MsgBox("+trim(match[1].matched ? match[1].str() : match[2].str())+")";
            if (!validVbaExpression(expr)) return {VbaExecutionStatus::Unsupported, "MsgBox argument syntax unsupported", line};
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
            const auto args=splitArguments(match[1].str());
            if (args.empty() || args.front().empty() || args.size()>7) return {VbaExecutionStatus::Unsupported,"OpenForm arguments unsupported",line};
            for (const auto& arg:args) if (!arg.empty() && !validVbaExpression(arg)) return {VbaExecutionStatus::Unsupported,"OpenForm argument expression unsupported",line};
            statements.push_back({Statement::Type::OpenForm,{},match[1].str(),line});
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
            if (lower(match[1].str())=="msgbox") {
                const auto expression="MsgBox("+args+")";
                if (!validVbaExpression(expression)) return {VbaExecutionStatus::Unsupported,"MsgBox argument syntax unsupported",line};
                statements.push_back({Statement::Type::Message,{},expression,line}); continue;
            }
            for (const auto& arg : splitArguments(args))
                if (!validVbaExpression(arg)) return {VbaExecutionStatus::Unsupported, "Aufrufargument", line};
            statements.push_back({Statement::Type::Call, match[1].str(), args, line});
        } else if (std::regex_match(source, match, assignment)) {
            const auto expr = trim(match[3].str());
            if (!validVbaExpression(expr)) return {VbaExecutionStatus::Unsupported, "Zuweisung: " + expr, line};
            auto field = match[2].str();
            if (field.starts_with('[') && field.ends_with(']')) field = field.substr(1, field.size() - 2);
            statements.push_back({Statement::Type::SetField,field,expr,line,match[1].str()});
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
                    && type != "control" && type != "form" && type != "access.control" && type != "access.form"
                    && type != "dao.tabledef" && type != "dao.tabledefs" && type != "dao.recordset" && type != "dao.field" && type != "dao.fields" && type != "dao.querydef" && type != "dao.parameter" && type != "string" && type != "boolean"
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
        } else if (std::regex_match(source,match,memberCall)) {
            const auto expression = match[1].str() + "(" + match[2].str() + ")";
            if (!validVbaExpression(expression)) return {VbaExecutionStatus::Unsupported,"Object method arguments",line};
            statements.push_back({Statement::Type::ObjectCall,{},expression,line});
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

std::shared_ptr<VbaModuleState> createVbaModuleState(std::string_view module) {
    auto state=std::make_shared<VbaModuleState>();
    const std::regex procedure(R"(^\s*(?:(?:Public|Private|Friend|Static)\s+)?(?:Sub|Function|Property)\s+)",std::regex::icase);
    const std::regex declaration(R"(^\s*(Dim|Private|Public|Global)\s+(.+)$)",std::regex::icase);
    const std::regex variable(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)(?:\s+As\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*))?$)",std::regex::icase);
    std::istringstream stream{std::string(module)}; std::string line;
    while (std::getline(stream,line)) {
        line=withoutComment(line); if (std::regex_search(line,procedure)) break;
        std::smatch match; if (!std::regex_match(line,match,declaration)) continue;
        const auto visibility=lower(match[1].str());
        const auto declarations=match[2].str();
        const auto start=lower(trim(declarations));
        if (start.starts_with("declare ") || start.starts_with("type ") || start.starts_with("enum ")) continue;
        for (const auto& entry:splitArguments(declarations)) {
            std::smatch field;
            if (!std::regex_match(entry,field,variable)) throw VbaRuntimeError(3251,"Module declaration unsupported");
            const auto name=lower(field[1].str()),type=lower(field[2].str());
            if (!type.empty() && type!="variant" && type!="boolean" && type!="string")
                throw VbaRuntimeError(3251,"Module field type unsupported");
            if (state->values.contains(name)) throw VbaRuntimeError(5,"Duplicate module field");
            state->values.emplace(name,type=="boolean" ? VbaValue::logical(false) : type=="string" ? VbaValue::string("") : VbaValue::empty());
            state->types.emplace(name,type);
            if (visibility=="public" || visibility=="global") state->publicVariables.insert(name);
        }
    }
    return state;
}
std::shared_ptr<VbaModuleState> VbaHost::moduleState(std::string_view module) {
    const auto key=std::string(module);
    if (const auto found=moduleStates_.find(key);found!=moduleStates_.end()) return found->second;
    auto state=createVbaModuleState(module); moduleStates_.emplace(key,state); return state;
}
VbaValue coerceVbaModuleValue(const VbaValue& value,std::string_view type) {
    if (type.empty() || type=="variant") return value;
    if (value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null module field assignment");
    if (value.kind==VbaValue::Kind::Error || value.kind==VbaValue::Kind::Object) throw VbaRuntimeError(13,"Module field type mismatch");
    if (type=="string") {
        if (value.kind==VbaValue::Kind::Empty) return VbaValue::string("");
        if (value.kind==VbaValue::Kind::Boolean) return VbaValue::string(value.boolean ? "True" : "False");
        return VbaValue::string(value.asText());
    }
    if (type=="boolean") {
        if (value.kind==VbaValue::Kind::Text) {
            const auto text=lower(trim(value.text));
            if (text=="true" || text=="false") return VbaValue::logical(text=="true");
        }
        try { return VbaValue::logical(value.asNumber()!=0); }
        catch (...) { throw VbaRuntimeError(13,"Module Boolean field type mismatch"); }
    }
    throw VbaRuntimeError(3251,"Module field coercion unsupported");
}
VbaOpenFormArguments parseVbaOpenFormArguments(const std::vector<VbaValue>& args) {
    if (args.empty()) throw VbaRuntimeError(450,"OpenForm requires a form name");
    if (args.size()>4) throw VbaRuntimeError(3251,"OpenForm data/window modes and OpenArgs are unsupported");
    if (args[0].kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null form name");
    VbaOpenFormArguments result{args[0].asText(),{}};
    if (result.name.empty()) throw VbaRuntimeError(2102,"Form name is unavailable");
    if (args.size()>1 && args[1].kind!=VbaValue::Kind::Empty) {
        if (args[1].kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null form view");
        const auto number=args[1].asNumber(); const auto low=std::floor(number); const auto fraction=number-low;
        const auto view=fraction<0.5 ? low : fraction>0.5 ? low+1 : std::fmod(std::abs(low),2)==0 ? low : low+1;
        if (!std::isfinite(view) || view < -2147483648.0 || view > 2147483647.0) throw VbaRuntimeError(6,"OpenForm view overflow");
        if (view!=0) throw VbaRuntimeError(3251,"OpenForm view is unsupported");
    }
    if (args.size()>2) {
        if (args[2].kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null filter name");
        if (!args[2].asText().empty()) throw VbaRuntimeError(3251,"OpenForm named filter is unsupported");
    }
    if (args.size()>3) {
        if (args[3].kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null WhereCondition");
        result.where=args[3].asText();
    }
    return result;
}
void VbaHost::openFormArguments(const std::vector<VbaValue>& args) {
    const auto request=parseVbaOpenFormArguments(args);
    if (!request.where.empty()) throw VbaRuntimeError(3251,"Host does not support a form WhereCondition");
    openForm(request.name);
}
namespace {
VbaValue initialLocalValue(const std::string& type) {
    if (type=="object" || type=="control" || type=="form" || type.starts_with("access.") || type.starts_with("dao.")) return VbaValue::reference(nullptr);
    if (type=="string") return VbaValue::string("");
    if (type=="boolean") return VbaValue::logical(false);
    if (!type.empty() && type!="variant") return VbaValue::numeric(0);
    return VbaValue::empty();
}
struct ReferenceBinding { VbaValue* source{}; std::string type; };
}

static VbaExecutionResult executeProcedure(std::string_view module,
                                       std::string_view procedure,
                                       VbaHost& host, std::vector<VbaValue>& arguments,
                                       const std::vector<ReferenceBinding>* references = nullptr) {
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
    std::string resultType;
    auto result = parseProcedure(module, procedure, instructions, isFunction, hasCancel, &parameters,true,&resultType);
    if (result.status != VbaExecutionStatus::Completed) return result;
    host.enterModule(module);
    struct ModuleScope { VbaHost& host; ~ModuleScope() { host.leaveModule(); } } moduleScope{host};
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
    } catch (const VbaRuntimeError& exception) {
        VbaExecutionResult failure{VbaExecutionStatus::Error,"Procedure resolution failed",0}; failure.errorNumber=exception.number; return failure;
    } catch (const std::exception& exception) {
        return {VbaExecutionStatus::Error, exception.what(), 0};
    }
    if (instructions.size() > 20000)
        return {VbaExecutionStatus::Unsupported, "Prozedur überschreitet 20.000 Anweisungen", 0};
    struct Frame { bool parent; bool condition; bool inElse; };
    struct LoopFrame {
        int forIndex;
        double end{};
        double step{};
        std::vector<VbaValue> items;
        std::size_t itemIndex{};
        std::size_t frameDepth{};
    };
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
    if (isFunction) locals[lower(std::string(procedure))] = initialLocalValue(resultType);
    const bool implicitCancel = hasCancel && arguments.empty();
    if (!implicitCancel && parameters.size() != arguments.size())
        return {VbaExecutionStatus::Error, "Falsche Argumentanzahl", 0};
    for (std::size_t i = 0; i < parameters.size(); ++i)
        locals[lower(parameters[i].name)] = implicitCancel ? VbaValue::numeric(0) : arguments[i];
    // Declarations bind for the whole procedure, including skipped branches.
    for (const auto& instruction:instructions) if (instruction.type==Statement::Type::Dim)
        locals.try_emplace(lower(instruction.argument),initialLocalValue(instruction.expression));
    struct Alias { std::string name; VbaValue* source; VbaValue before; std::string type; };
    std::vector<Alias> aliases;
    if (references && !implicitCancel)
        for (std::size_t i = 0; i < parameters.size(); ++i)
            if (parameters[i].byReference && (*references)[i].source)
                aliases.push_back({lower(parameters[i].name), (*references)[i].source, *(*references)[i].source,(*references)[i].type});
    std::vector<std::shared_ptr<VbaModuleState>> stateOwners;
    try {
        auto state=host.moduleState(module); stateOwners.push_back(state);
        for (auto& [name,value]:state->values) if (!locals.contains(name)) {
            locals.emplace(name,value); aliases.push_back({name,&value,value,state->types.at(name)});
        }
        const auto reserved=locals;
        std::unordered_set<std::string> globals;
        auto globalStates=host.globalModuleStates();
        for (auto& global:globalStates) {
            stateOwners.push_back(global);
            if (global==state) continue;
            for (const auto& name:global->publicVariables) {
                if (reserved.contains(name) || host.hasInstanceIdentifier(name)) continue;
                if (!globals.insert(name).second) throw VbaRuntimeError(5,"Ambiguous public module field");
                auto& value=global->values.at(name);
                locals.emplace(name,value); aliases.push_back({name,&value,value,global->types.at(name)});
            }
        }
    } catch (const VbaRuntimeError& error) {
        VbaExecutionResult failure{VbaExecutionStatus::Error,"Module state initialization failed",0}; failure.errorNumber=error.number; return failure;
    }
    const auto equalValue = [](const VbaValue& a, const VbaValue& b) {
        return a.kind == b.kind && a.number == b.number && a.currencyUnits==b.currencyUnits && a.text == b.text
            && a.boolean == b.boolean && a.object == b.object && a.errorNumber==b.errorNumber;
    };
    const auto refreshAliases = [&] {
        for (auto& alias : aliases) { locals[alias.name] = *alias.source; alias.before = *alias.source; }
    };
    const auto storeAliases = [&] {
        for (auto& alias : aliases)
            if (!equalValue(locals[alias.name], alias.before)) {
                try { *alias.source=coerceVbaModuleValue(locals[alias.name],alias.type); }
                catch (...) { locals[alias.name]=*alias.source; alias.before=*alias.source; throw; }
            }
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
            if (value.kind != VbaValue::Kind::Object && value.kind!=VbaValue::Kind::Error) outcome.returnValue = value.asField();
        }
        storeAliases();
        copyArguments();
        if (hasCancel) outcome.cancelled = locals["cancel"].truthy();
        return outcome;
    };
    const auto invoke = [&](std::string_view name, std::vector<VbaValue>& actuals,
                            const std::vector<std::string>& names, bool needsValue) -> VbaValue {
        std::vector<ReferenceBinding> bindings;
        for (const auto& expression : names) {
            auto found = locals.find(lower(expression));
            ReferenceBinding binding{found==locals.end() ? nullptr : &found->second,{}};
            for (const auto& alias:aliases) if (alias.name==lower(expression)) binding={alias.source,alias.type};
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
        storeAliases();
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
    const auto evaluateObject = [&](std::string_view expression) {
        return evaluateVbaObjectExpression(expression,host,locals,resolver);
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
                    (void)evaluate(instruction.expression);
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
                    std::vector<VbaValue> args;
                    for (const auto& expression:splitArguments(instruction.expression)) args.push_back(expression.empty() ? VbaValue::empty() : evaluate(expression));
                    host.openFormArguments(args);
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
                    const auto value = evaluate(instruction.expression);
                    if (instruction.endExpression == "!") host.setControlValue(instruction.argument,value);
                    else host.setFieldValue(instruction.argument,value);
                }
                break;
            case Statement::Type::ObjectSet:
                if (active) {
                    auto value = evaluateObject(instruction.expression);
                    if (value.kind != VbaValue::Kind::Object) throw VbaRuntimeError(13,"Set benötigt ein Objekt");
                    locals[lower(instruction.argument)] = std::move(value);
                }
                break;
            case Statement::Type::ObjectAssign:
                if (active) {
                    const auto pos = memberSeparator(instruction.argument);
                    auto name = instruction.argument.substr(pos + 1);
                    if (name.starts_with('[') && name.ends_with(']')) name = name.substr(1, name.size() - 2);
                    const auto object = evaluateObject(instruction.argument.substr(0, pos));
                    if (!object.object) throw VbaRuntimeError(91,"Objektvariable nicht gesetzt");
                    const auto value = evaluate(instruction.expression);
                    if (instruction.argument[pos] == '!') object.object->assignItem(name,value);
                    else object.object->assign(name,value);
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
                if (active) {
                    const auto value=evaluate(instruction.expression);
                    const auto name=lower(instruction.argument);
                    if (locals.contains(name) || !host.assignIdentifier(instruction.argument,value)) locals[name]=value;
                }
                break;
            case Statement::Type::ForEach: {
                if (!active) { ip = static_cast<std::size_t>(instruction.partner); break; }
                const auto collection = evaluate(instruction.expression);
                if (collection.kind != VbaValue::Kind::Object)
                    throw VbaRuntimeError(424,"For Each requires an object collection");
                if (!collection.object) throw VbaRuntimeError(91,"For Each object is not set");
                LoopFrame loop{static_cast<int>(ip)};
                loop.items = collection.object->enumerate();
                loop.frameDepth = frames.size();
                if (loop.items.empty()) ip = static_cast<std::size_t>(instruction.partner);
                else {
                    locals[lower(instruction.argument)] = loop.items.front();
                    loops.push_back(std::move(loop));
                }
                break;
            }
            case Statement::Type::ExitFor:
                if (active) {
                    if (loops.empty()) throw VbaRuntimeError(5,"Exit For stack invalid");
                    const auto loop = std::move(loops.back()); loops.pop_back();
                    frames.resize(loop.frameDepth); active = true;
                    ip = static_cast<std::size_t>(instructions[loop.forIndex].partner);
                }
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
                } else {
                    LoopFrame loop{static_cast<int>(ip),end,step};
                    loop.frameDepth = frames.size();
                    loops.push_back(std::move(loop));
                }
                break;
            }
            case Statement::Type::Next: {
                if (!active) break;
                if (loops.empty() || loops.back().forIndex != instruction.partner)
                    throw std::runtime_error("VBA-For/Next-Laufzeitstapel ungültig");
                const auto& start = instructions[loops.back().forIndex];
                const auto name = lower(start.argument);
                if (start.type == Statement::Type::ForEach) {
                    auto& loop = loops.back();
                    if (++loop.itemIndex < loop.items.size()) {
                        locals[name] = loop.items[loop.itemIndex];
                        ip = static_cast<std::size_t>(loop.forIndex);
                    } else loops.pop_back();
                    break;
                }
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
                break; // Already initialized in the procedure declaration scope.
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
VbaExecutionResult executeVbaEventExpression(std::string_view module,std::string_view expression,VbaHost& host) {
    auto source=trim(std::string(expression));
    if (source.empty() || source.front()!='=') {
        VbaExecutionResult result{VbaExecutionStatus::Unsupported,"Event binding requires a function expression",0}; result.errorNumber=3251; return result;
    }
    source=trim(source.substr(1));
    if (!validVbaExpression(source)) return {VbaExecutionStatus::Unsupported,"Event expression syntax unsupported",0};
    int line=0;
    try {
        const VbaFunctionResolver resolver=[&](std::string_view name,std::vector<VbaValue>& args,const std::vector<std::string>&) {
            auto outcome=executeVbaProcedureWithArguments(module,name,host,args);
            if (outcome.status==VbaExecutionStatus::ProcedureMissing) {
                if (const auto external=host.moduleForProcedure(name)) outcome=executeVbaProcedureWithArguments(*external,name,host,args);
            }
            line=outcome.line;
            if (outcome.status!=VbaExecutionStatus::Completed || !outcome.typedReturnValue)
                throw VbaRuntimeError(outcome.errorNumber ? outcome.errorNumber : 5,outcome.explanation.empty() ? "Event function execution failed" : outcome.explanation);
            return *outcome.typedReturnValue;
        };
        const auto value=evaluateVbaExpression(source,host,{},resolver);
        VbaExecutionResult result{VbaExecutionStatus::Completed,{},line};
        result.typedReturnValue=std::make_shared<VbaValue>(value);
        if (value.kind!=VbaValue::Kind::Object && value.kind!=VbaValue::Kind::Error) result.returnValue=value.asField();
        return result;
    } catch (const VbaRuntimeError& error) {
        VbaExecutionResult result{VbaExecutionStatus::Error,error.what(),line}; result.errorNumber=error.number; return result;
    } catch (const std::exception&) {
        VbaExecutionResult result{VbaExecutionStatus::Error,"Event expression execution failed",line}; result.errorNumber=5; return result;
    }
}
}  // namespace linaccess
