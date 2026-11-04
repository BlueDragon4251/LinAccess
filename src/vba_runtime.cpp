#include "vba_array.hpp"
#include "vba_runtime.hpp"
#include "vba_expression.hpp"
#include "vba_currency.hpp"
#include "vba_calendar.hpp"

#include <algorithm>
#include <deque>
#include <list>
#include <new>
#include <cctype>
#include <cmath>
#include <regex>
#include <limits>
#include <sstream>
#include <stdexcept>
#include <unordered_map>
#include <vector>

namespace linaccess {
namespace {
thread_local std::uint64_t stopEpoch=0;
thread_local std::vector<std::uint64_t> activeExecutionEpochs;
// This interpreter deliberately reports unsupported VBA instead of silently
// treating Access/DAO/COM calls as successful no-ops.
struct Statement {
    enum class Type { With, EndWith, Select, Case, EndSelect, SendKeys, If, ElseIf, Else, EndIf, For, ForEach, ExitFor, Next, Evaluate, Save, Maximize, GoLast,
                      GoNext, GoNew, Requery,
                      ExecuteSql, OpenForm, ControlRequery, ControlFocus, ControlVisible,
                      Call, SetField, IndexedAssign, ResizeArray, EraseArray, SetVar, Exit, Dim, ObjectSet, ObjectAssign, ObjectCall, While, EndWhile, ExitDo, Quit, ErrorResumeNext, ErrorDisable, Label, Jump, ErrorHandler, Resume, ResumeNext, ResumeLabel } type;
    std::string argument;
    std::string expression;
    int line{};
    std::string endExpression;
    std::string stepExpression;
    int partner{-1};
    int conditionalUnwind{}, withUnwind{}, forUnwind{};
    std::shared_ptr<const VbaValue> constantValue;
    std::vector<VbaArrayDimension> fixedBounds;
    bool persistent{};
};

std::string trim(std::string text) {
    const auto first = text.find_first_not_of(" \t\r\n");
    if (first == std::string::npos) return {};
    const auto last = text.find_last_not_of(" \t\r\n");
    return text.substr(first, last - first + 1);
}
bool nativeDaoType(std::string_view type) {
    return type=="dao.database" || type=="dao.databases" || type=="dao.workspace" || type=="dao.workspaces"
        || type=="dao.dbengine" || type=="dao.recordset" || type=="dao.field" || type=="dao.fields"
        || type=="dao.querydef" || type=="dao.parameter" || type=="dao.parameters"
        || type=="dao.tabledef" || type=="dao.tabledefs";
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

struct MemberTarget { std::string name; std::optional<std::string> arguments; };
std::optional<MemberTarget> memberTarget(const std::string& tail) {
    std::size_t end=0;
    if(tail.starts_with('[')) {
        const auto closing=tail.find(']'); if(closing==std::string::npos || closing==1) return {};
        end=closing+1;
    } else {
        const auto first=tail.empty() ? 0 : static_cast<unsigned char>(tail[0]);
        if(!std::isalpha(first) && first!='_' && first<128) return {};
        while(end<tail.size()) {
            const auto ch=static_cast<unsigned char>(tail[end]);
            if(!std::isalnum(ch) && ch!='_' && ch<128) break;
            ++end;
        }
    }
    MemberTarget target{tail.substr(0,end),{}};
    if(target.name.starts_with('[')) target.name=target.name.substr(1,target.name.size()-2);
    const auto suffix=trim(tail.substr(end));
    if(suffix.empty()) return target;
    if(!suffix.starts_with('(') || !suffix.ends_with(')')) return {};
    bool quoted=false,bracket=false; int depth=0;
    for(std::size_t index=0;index<suffix.size();++index) {
        const char ch=suffix[index];
        if(bracket) { if(ch==']') bracket=false; continue; }
        if(ch=='"') {
            if(quoted && index+1<suffix.size() && suffix[index+1]=='"') { ++index; continue; }
            quoted=!quoted; continue;
        }
        if(quoted) continue;
        if(ch=='[') bracket=true;
        else if(ch=='(') ++depth;
        else if(ch==')' && --depth==0 && index+1<suffix.size()) return {};
    }
    if(depth || quoted || bracket) return {};
    target.arguments=suffix.substr(1,suffix.size()-2); return target;
}

bool objectPath(const std::string& path) {
    std::size_t position=0;
    const auto spaces=[&] { while(position<path.size() && std::isspace(static_cast<unsigned char>(path[position]))) ++position; };
    while(position<path.size()) {
        spaces(); if(position==path.size()) return false;
        if(path[position]=='[') {
            const auto closing=path.find(']',position+1); if(closing==std::string::npos || closing==position+1) return false;
            position=closing+1;
        } else {
            const auto first=static_cast<unsigned char>(path[position]);
            if(!std::isalpha(first) && first!='_' && first<128) return false;
            while(position<path.size()) {
                const auto ch=static_cast<unsigned char>(path[position]);
                if(!std::isalnum(ch) && ch!='_' && ch<128) break;
                ++position;
            }
        }
        spaces();
        while(position<path.size() && path[position]=='(') {
            bool quoted=false,bracket=false; int depth=0;
            do {
                const char ch=path[position++];
                if(bracket) { if(ch==']') bracket=false; continue; }
                if(ch=='"') {
                    if(quoted && position<path.size() && path[position]=='"') { ++position; continue; }
                    quoted=!quoted; continue;
                }
                if(quoted) continue;
                if(ch=='[') bracket=true;
                else if(ch=='(') ++depth;
                else if(ch==')') --depth;
            } while(position<path.size() && depth>0);
            if(depth || quoted || bracket) return false;
            spaces();
        }
        if(position==path.size()) return true;
        if(path[position]!='.' && path[position]!='!') return false;
        ++position;
    }
    return false;
}

std::optional<std::pair<std::string,std::string>> topLevelAssignment(const std::string& source) {
    bool quoted=false,bracket=false; int depth=0;
    for(std::size_t index=0;index<source.size();++index) {
        const char ch=source[index];
        if(bracket) { if(ch==']') bracket=false; continue; }
        if(ch=='"') {
            if(quoted && index+1<source.size() && source[index+1]=='"') { ++index; continue; }
            quoted=!quoted; continue;
        }
        if(quoted) continue;
        if(ch=='[') bracket=true;
        else if(ch=='(') ++depth;
        else if(ch==')') --depth;
        else if(ch=='=' && depth==0) {
            auto target=trim(source.substr(0,index));
            return std::pair{std::move(target),trim(source.substr(index+1))};
        }
    }
    return {};
}
std::optional<std::pair<std::string,std::string>> objectAssignment(const std::string& source) {
    auto assignment=topLevelAssignment(source); if(!assignment) return {};
    const auto separator=memberSeparator(assignment->first);
    if(separator==std::string::npos || !objectPath(trim(assignment->first.substr(0,separator)))
        || !memberTarget(trim(assignment->first.substr(separator+1)))) return {};
    return assignment;
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
    bool hasScalarDefault() const override { return true; }
    VbaValue scalarValue() override { return VbaValue::longInteger(number); }
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
        if (key == "number") return VbaValue::longInteger(number);
        if (key == "description") return VbaValue::string(description);
        if (key == "clear") { number = 0; description.clear(); return {}; }
        throw VbaRuntimeError(438, "Err-Mitglied nicht unterstützt");
    }
};

struct Parameter { std::string name; bool byReference{true}; std::string type; bool optional{}; std::string defaultExpression; bool paramArray{}; std::shared_ptr<const VbaValue> defaultValue; };
std::vector<std::string> splitArguments(const std::string& input) {
    std::vector<std::string> args;
    if (trim(input).empty()) return args;
    bool quoted = false, bracket = false;
    int parentheses = 0;
    std::size_t start = 0;
    for (std::size_t i = 0; i < input.size(); ++i) {
        const char ch = input[i];
        if (bracket) {
            if (ch==']') { if (i+1<input.size() && input[i+1]==']') ++i; else bracket=false; }
            continue;
        }
        if (!quoted && ch=='[') { bracket=true; continue; }
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

struct CallArgument { std::string name,expression; };
CallArgument callArgument(const std::string& source) {
    static const std::regex named(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*:=\s*(.*)$)");
    std::smatch match;
    if(std::regex_match(source,match,named)) return {lower(match[1].str()),trim(match[2].str())};
    return {{},source};
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
    std::string result=source; bool quoted=false,bracket=false,dated=false,beginning=true; int depth=0;
    for (std::size_t i=0;i<result.size();++i) {
        const char ch=result[i];
        if(dated) { if(ch=='#') dated=false; continue; }
        if (quoted) { if (ch=='"') { if (i+1<result.size() && result[i+1]=='"') ++i; else quoted=false; } continue; }
        if (bracket) { if (ch==']') { if (i+1<result.size() && result[i+1]==']') ++i; else bracket=false; } continue; }
        if(beginning && lower(result.substr(i,3))=="rem" && (i+3==result.size() || std::isspace(static_cast<unsigned char>(result[i+3])) || result[i+3]==':')) { result.resize(i); break; }
        if(!std::isspace(static_cast<unsigned char>(ch))) beginning=false;
        if(ch=='#') dated=true;
        else if (ch=='"') quoted=true; else if (ch=='[') bracket=true;
        else if (ch=='(') ++depth; else if (ch==')') --depth;
        else if (ch==':' && depth==0 && (i+1>=result.size() || result[i+1]!='=')) { result[i]='\n'; beginning=true; }
    }
    return result;
}

int moduleArrayBase(std::string_view module) {
    std::istringstream source{std::string(module)}; std::string line;
    const std::regex option(R"(^\s*Option\s+Base\s+([01])\s*(?:'.*)?$)",std::regex::icase);
    int result=0; while(std::getline(source,line)) { std::smatch match; if(std::regex_match(line,match,option)) result=match[1].str()=="1" ? 1 : 0; }
    return result;
}
class ConstantExpressionHost final : public VbaHost {
public:
    std::optional<std::string> field(std::string_view) override { throw VbaRuntimeError(3251,"A constant cannot read runtime variables"); }
    void setField(std::string_view,std::optional<std::string>) override { throw VbaRuntimeError(3251,"A constant cannot write fields"); }
    void message(std::string_view) override { throw VbaRuntimeError(3251,"A constant cannot execute messages"); }
    void saveRecord() override { throw VbaRuntimeError(3251,"A constant cannot save records"); }
};
VbaCompareMode moduleCompareMode(std::string_view module) {
    std::istringstream stream{std::string(module)}; std::string line;
    static const std::regex option(R"(^\s*Option\s+Compare\s+(Binary|Text|Database)\s*(?:'.*)?$)",std::regex::icase);
    auto mode=VbaCompareMode::Binary;
    while(std::getline(stream,line)) { std::smatch match; if(std::regex_match(line,match,option)) mode=lower(match[1].str())=="text" ? VbaCompareMode::Text : lower(match[1].str())=="database" ? VbaCompareMode::Database : VbaCompareMode::Binary; }
    return mode;
}
VbaValue constantExpression(const std::string& expression,const VbaVariables& values,VbaCompareMode mode) {
    static const std::regex number(R"((?:&[Hh][0-9A-Fa-f]+|&[Oo][0-7]+|(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?)[%&!#@]?)");
    const auto word=[](unsigned char ch) { return std::isalnum(ch) || ch=='_' || ch>=128; };
    const std::unordered_set<std::string> operators{"not","and","or","xor","eqv","imp","mod","like"};
    for(std::size_t i=0;i<expression.size();) {
        const auto ch=static_cast<unsigned char>(expression[i]);
        if(std::isspace(ch)) { ++i; continue; }
        if(ch=='"' || ch=='#') {
            const auto quote=ch; bool closed=false; ++i;
            while(i<expression.size()) {
                if(static_cast<unsigned char>(expression[i++])!=quote) continue;
                if(quote=='"' && i<expression.size() && expression[i]=='"') { ++i; continue; }
                closed=true; break;
            }
            if(!closed) throw VbaRuntimeError(3251,"Unclosed constant literal");
            continue;
        }
        if(std::isdigit(ch) || (ch=='.' && i+1<expression.size() && std::isdigit(static_cast<unsigned char>(expression[i+1])))
            || (ch=='&' && i+1<expression.size() && (expression[i+1]=='H' || expression[i+1]=='h' || expression[i+1]=='O' || expression[i+1]=='o'))) {
            std::smatch match; const auto remainder=expression.substr(i);
            if(!std::regex_search(remainder,match,number,std::regex_constants::match_continuous)) throw VbaRuntimeError(3251,"Invalid constant number");
            i+=match.length(); continue;
        }
        if(std::isalpha(ch) || ch=='_' || ch>=128) {
            const auto start=i++; while(i<expression.size() && word(static_cast<unsigned char>(expression[i]))) ++i;
            const auto name=lower(expression.substr(start,i-start));
            if(name=="date" || name=="time" || name=="now" || name=="timer") throw VbaRuntimeError(3251,"Clock reads are not constant expressions");
            auto next=i; while(next<expression.size() && std::isspace(static_cast<unsigned char>(expression[next]))) ++next;
            if((next<expression.size() && expression[next]=='(' && !operators.contains(name))
                || (!values.contains(name) && name!="true" && name!="false" && !operators.contains(name)
                    && !name.starts_with("vb") && !name.starts_with("ac") && !name.starts_with("db")))
                throw VbaRuntimeError(3251,"Constant expressions require literals, known constants and operators");
            continue;
        }
        if(std::string("+-*/\\^&()=<>").find(static_cast<char>(ch))==std::string::npos) throw VbaRuntimeError(3251,"Unsupported constant expression token");
        ++i;
    }
    ConstantExpressionHost host;
    auto value=evaluateVbaExpression(expression,host,values,{},nullptr,mode);
    if(value.kind==VbaValue::Kind::Null || value.kind==VbaValue::Kind::Empty || value.kind==VbaValue::Kind::Error
        || value.kind==VbaValue::Kind::Object || value.kind==VbaValue::Kind::Missing || isVbaArray(value))
        throw VbaRuntimeError(3251,"Only scalar non-Null constants are implemented");
    if(value.kind==VbaValue::Kind::Number && !std::isfinite(value.number)) throw VbaRuntimeError(6,"Constant overflow");
    return value;
}
struct ConstantDeclaration { std::string name,type; VbaValue value; };
std::vector<ConstantDeclaration> constantDeclarations(const std::string& text,VbaVariables& known,VbaCompareMode mode) {
    static const std::regex declaration(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)(?:\s+As\s+([A-Za-z]+))?\s*=\s*(.+)$)",std::regex::icase);
    std::vector<ConstantDeclaration> result; std::unordered_set<std::string> names;
    for(const auto& entry:splitArguments(text)) {
        std::smatch match; if(!std::regex_match(entry,match,declaration)) throw VbaRuntimeError(3251,"Unsupported Const declaration");
        const auto name=lower(match[1].str()); auto type=lower(match[2].str());
        if(!names.insert(name).second) throw VbaRuntimeError(5,"Duplicate constant");
        if(!type.empty() && type!="variant" && type!="byte" && type!="integer" && type!="long" && type!="single" && type!="double"
            && type!="currency" && type!="date" && type!="boolean" && type!="string") throw VbaRuntimeError(3251,"Unsupported constant type");
        auto value=constantExpression(match[3].str(),known,mode);
        if(!type.empty()) value=coerceVbaModuleValue(value,type);
        else if(value.kind==VbaValue::Kind::Text) type="string";
        else if(value.kind==VbaValue::Kind::Boolean) type="boolean";
        else if(value.kind==VbaValue::Kind::Currency) type="currency";
        else if(value.kind==VbaValue::Kind::Date) type="date";
        else type=value.numericType==VbaValue::NumericType::Byte ? "byte" : value.numericType==VbaValue::NumericType::Integer ? "integer" : value.numericType==VbaValue::NumericType::Long ? "long" : value.numericType==VbaValue::NumericType::Single ? "single" : "double";
        known[name]=value; result.push_back({name,type,std::move(value)});
    }
    if(result.empty()) throw VbaRuntimeError(3251,"Empty Const declaration");
    return result;
}
VbaVariables moduleConstantValues(std::string_view module) {
    static const std::regex procedure(R"(^\s*(?:(?:Public|Private|Friend|Static)\s+)?(?:Sub|Function|Property)\s+)",std::regex::icase);
    static const std::regex declaration(R"(^\s*(?:(?:Public|Private)\s+)?Const\s+(.+)$)",std::regex::icase);
    VbaVariables values; std::unordered_set<std::string> names; const auto mode=moduleCompareMode(module);
    std::istringstream source{std::string(module)}; std::string line;
    while(std::getline(source,line)) {
        line=withoutComment(line); if(std::regex_search(line,procedure)) break;
        std::smatch match; if(!std::regex_match(line,match,declaration)) continue;
        for(const auto& constant:constantDeclarations(match[1].str(),values,mode)) if(!names.insert(constant.name).second) throw VbaRuntimeError(5,"Duplicate module constant");
    }
    return values;
}
std::optional<std::vector<VbaArrayDimension>> literalArrayBounds(const std::string& text,int base,const VbaVariables& values={},VbaCompareMode mode=VbaCompareMode::Binary) {
    const auto dimensions=splitArguments(text); if(dimensions.empty() || dimensions.size()>60) return {};
    const auto bound=[&](const std::string& expression) {
        const auto value=constantExpression(trim(expression),values,mode);
        if(value.kind!=VbaValue::Kind::Number || std::trunc(value.number)!=value.number) throw VbaRuntimeError(3251,"Array bounds require integral constants");
        return static_cast<int>(coerceVbaModuleValue(value,"long").number);
    };
    std::vector<VbaArrayDimension> result;
    try { for(const auto& entry:dimensions) {
        const auto range=outsideKeyword(entry,"to");
        const auto lower=range==std::string::npos ? base : bound(entry.substr(0,range));
        const auto upper=bound(range==std::string::npos ? entry : entry.substr(range+2));
        if(lower>upper) return {}; result.push_back({lower,upper});
    } } catch(const std::exception&) { return {}; }
    return result;
}
VbaExecutionResult parseProcedure(std::string_view module, std::string_view name,
                                  std::vector<Statement>& statements, bool& isFunction, bool& hasCancel,
                                  std::vector<Parameter>* parameters = nullptr,bool resolveLabels=true,std::string* resultType=nullptr,bool inheritedWith=false);
VbaExecutionResult parseProcedureUncached(std::string_view module, std::string_view name,
                                  std::vector<Statement>& statements, bool& isFunction, bool& hasCancel,
                                  std::vector<Parameter>* parameters = nullptr,bool resolveLabels=true,std::string* resultType=nullptr,bool inheritedWith=false) {
    VbaVariables declarationConstants; const auto constantMode=moduleCompareMode(module);
    try { declarationConstants=moduleConstantValues(module); }
    catch(const VbaRuntimeError& error) { VbaExecutionResult result{VbaExecutionStatus::Unsupported,"Module constant declaration unsupported",0}; result.errorNumber=error.number; return result; }
    std::unordered_set<std::string> localConstants;
    const std::regex begin(
        R"(^\s*(?:(?:Public|Private|Friend|Static)\s+)?(Sub|Function)\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*\(\s*(.*?)\s*\)\s*(?:As\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*(?:\.[A-Za-z_][A-Za-z0-9_]*)?))?\s*$)",
        std::regex::icase);
    const std::regex anyBegin(
        R"(^\s*(?:(?:Public|Private|Friend|Static)\s+)?(?:Sub|Function)\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)(?:\s*\(.*\))?.*$)",
        std::regex::icase);
    const std::regex end(R"(^\s*End\s+(Sub|Function)\s*$)", std::regex::icase);
    const std::regex selectStart(R"(^Select\s+Case\s+(.+)$)",std::regex::icase);
    const std::regex selectEnd(R"(^End\s+Select$)",std::regex::icase);
    const std::regex elseIfThen(R"(^ElseIf\s+(.+)\s+Then\s*$)",std::regex::icase);
    const std::regex ifThen(R"(^If\s+(.+)\s+Then\s*$)", std::regex::icase);
    const std::regex assignment(R"(^Me\s*([!.])\s*([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*|\[[^\]]+\])\s*=\s*(.+)$)", std::regex::icase);
    const std::regex variableAssignment(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*=\s*(.+)$)", std::regex::icase);
    const std::regex message(R"(^MsgBox(?:\s+(.+)|\((.*)\))$)", std::regex::icase);
    const std::regex evaluation(R"(^Eval(?:\s+(.+)|\((.*)\))$)",std::regex::icase);
    const std::regex save(R"(^DoCmd\s*\.\s*RunCommand\s+acCmdSaveRecord$)", std::regex::icase);
    const std::regex maximize(R"(^DoCmd\s*\.\s*Maximize\s*$)", std::regex::icase);
    const std::regex goLast(R"(^DoCmd\s*\.\s*GoToRecord\s*,\s*,\s*acLast\s*$)", std::regex::icase);
    const std::regex goNext(R"(^DoCmd\s*\.\s*GoToRecord\s*,\s*,\s*acNext\s*$)", std::regex::icase);
    const std::regex goNew(R"(^DoCmd\s*\.\s*GoToRecord\s*,\s*,\s*acNewRec\s*$)", std::regex::icase);
    const std::regex requery(R"(^DoCmd\s*\.\s*Requery\s*$)", std::regex::icase);
    const std::regex runSql(R"(^DoCmd\s*\.\s*RunSQL\s+(.+)$)", std::regex::icase);
    const std::regex daoExecute(R"(^CurrentDb(?:\s*\(\s*\))?\s*\.\s*Execute\s+(.+)$)", std::regex::icase);
    const std::regex openForm(R"(^DoCmd\s*\.\s*OpenForm\s+(.+)$)", std::regex::icase);
    const std::regex controlRequery(R"(^Me\s*([!.])\s*(\[[^\]]+\]|[A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*\.\s*Requery$)", std::regex::icase);
    const std::regex controlFocus(R"(^Me\s*([!.])\s*(\[[^\]]+\]|[A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*\.\s*SetFocus$)", std::regex::icase);
    const std::regex controlVisible(R"(^Me\s*([!.])\s*(\[[^\]]+\]|[A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*\.\s*Visible\s*=\s*(.+)$)", std::regex::icase);
    const std::regex call(R"(^Call\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)(?:\s*\((.*)\))?$)", std::regex::icase);
    const std::regex implicitCall(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)(?:\s+(.+)|\((.*)\))?$)", std::regex::icase);
    const std::regex memberCall(R"(^((?:[A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*(?:\([^)]*\))?)(?:[.!][A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*(?:\([^)]*\))?)+)\s+(.+)$)",std::regex::icase);
    const std::regex parameter(R"(^(?:(Optional)\s+)?(?:(ByRef|ByVal)\s+)?([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)(?:\s+As\s+(Byte|Integer|Long|String|Boolean|Double|Single|Currency|Date|Variant|Object|Collection|Scripting\.Dictionary|DAO\.Recordset|DAO\.Database|DAO\.Databases|DAO\.TableDef|DAO\.TableDefs|DAO\.Workspace|DAO\.Workspaces|DAO\.DBEngine|DAO\.Field|DAO\.Fields|DAO\.QueryDef|DAO\.Parameter|DAO\.Parameters))?(?:\s*=\s*(.+))?$)", std::regex::icase);
    const std::regex cancelParam(R"(^Cancel\s+As\s+(?:Integer|Boolean)$)", std::regex::icase);
    const std::regex objectSet(R"(^Set\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*=\s*(.+)$)", std::regex::icase);
    const std::regex label(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*):$)");
    const std::regex jump(R"(^GoTo\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)$)", std::regex::icase);
    const std::regex errorHandler(R"(^On\s+Error\s+GoTo\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)$)", std::regex::icase);
    const std::regex resumeLabel(R"(^Resume\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)$)", std::regex::icase);
    const std::regex whileStart(R"(^(Do\s+(While|Until)|While)\s+(.+)$)", std::regex::icase);
    const std::regex loopEnd(R"(^(Loop|Wend)(?:\s+(While|Until)\s+(.+))?$)",std::regex::icase);
    const std::regex dim(R"(^(?:Dim|Static)\s+(.+)$)", std::regex::icase);
    const std::regex declaration(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)(?:\s+As\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_.\x80-\xFF]*))?$)", std::regex::icase);
    const std::regex forStart(R"(^For\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*=\s*(.+?)\s+To\s+(.+?)(?:\s+Step\s+(.+))?$)", std::regex::icase);
    const std::regex forEach(R"(^For\s+Each\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s+In\s+(.+)$)", std::regex::icase);
    const std::regex next(R"(^Next(?:\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*))?$)", std::regex::icase);

    std::istringstream stream{std::string(module)};
    std::string raw;
    int line = 0;
    int nesting = 0;
    std::vector<bool> elseSeen;
    struct SelectBlock { bool hasCase{}; bool hasElse{}; };
    std::vector<SelectBlock> selects;
    std::vector<Statement::Type> blocks;
    const auto withActive=[&] { return inheritedWith || std::find(blocks.begin(),blocks.end(),Statement::Type::With)!=blocks.end(); };
    const auto validExpression=[&](std::string_view expression) { return validVbaExpression(expression,withActive()); };
    std::vector<int> forStarts;
    std::vector<int> whileStarts;
    bool within = false;
    std::unordered_map<std::string,int> labels;
    std::deque<std::string> pendingStatements;
    while (!pendingStatements.empty() || std::getline(stream, raw)) {
        if(pendingStatements.empty()) ++line;
        else { raw=std::move(pendingStatements.front()); pendingStatements.pop_front(); }
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
        auto source = withoutComment(raw);
        if (source.empty()) continue;
        if(within) {
            const auto expanded=inlineStatementLines(source);
            if(expanded.size()<source.size()) source=trim(source.substr(0,expanded.size()));
            if(source.empty()) continue;
            const auto folded=lower(source); const auto then=outsideKeyword(source,"then");
            if((folded.starts_with("if ") || folded.starts_with("elseif ")) && then!=std::string::npos) {
                const auto tail=trim(source.substr(then+4));
                if(!tail.empty() && inlineStatementLines(tail).empty()) source=trim(source.substr(0,then+4));
            }
            const bool inlineBranch=(folded.starts_with("if ") || folded.starts_with("if(") || folded.starts_with("elseif "))
                && then!=std::string::npos && !trim(source.substr(then+4)).empty();
            const auto separator=inlineStatementLines(source).find('\n');
            if(!inlineBranch && separator!=std::string::npos) {
                const auto following=trim(source.substr(separator+1));
                if(!following.empty()) pendingStatements.push_front(following);
                source=trim(source.substr(0,separator));
                if(source.empty()) continue;
                static const std::regex identifier(R"(^[A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*$)");
                static const std::unordered_set<std::string> statements{"else","do","loop","wend","next","stop","return","end","exit"};
                if(std::regex_match(source,identifier) && !statements.contains(lower(source))) source+=':';
            }
        }
        std::smatch match;
        if (!within) {
            if (std::regex_match(source, match, begin) &&
                lower(match[2].str()) == lower(std::string(name))) {
                const auto params = trim(match[3].str());
                hasCancel = std::regex_match(params, cancelParam);
                bool optionalSeen=false,paramArraySeen=false;
                const std::regex paramArray(R"(^ParamArray\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*\(\s*\)\s*(?:As\s+Variant)?$)",std::regex::icase);
                const std::regex arrayParameter(R"(^(?:ByRef\s+)?([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*\(\s*\)\s*(?:As\s+(Byte|Integer|Long|Single|Double|Currency|Date|Boolean|String|Object|Variant))?$)",std::regex::icase);
                for (const auto& declaration : splitArguments(params)) {
                    std::smatch declared;
                    if(paramArraySeen) return {VbaExecutionStatus::Unsupported,"ParamArray must be the last parameter",line};
                    if(std::regex_match(declaration,declared,paramArray)) {
                        if(optionalSeen) return {VbaExecutionStatus::Unsupported,"Combined Optional and ParamArray declarations are not implemented",line};
                        paramArraySeen=true;
                        if(parameters) parameters->push_back({declared[1].str(),false,"paramarray",false,{},true});
                        continue;
                    }
                    if(std::regex_match(declaration,declared,arrayParameter)) {
                        if(optionalSeen) return {VbaExecutionStatus::Unsupported,"Required array parameter after Optional is not implemented",line};
                        auto type=lower(declared[2].str()); if(type.empty()) type="variant";
                        if(parameters) parameters->push_back({declared[1].str(),true,type+"()",false,{}});
                        continue;
                    }
                    if (!std::regex_match(declaration, declared, parameter))
                        return {VbaExecutionStatus::Unsupported, "VBA-Parameterdeklaration nicht unterstützt", line};
                    const bool optional=declared[1].matched;
                    const auto defaultExpression=trim(declared[5].str());
                    if((optionalSeen && !optional) || (!optional && !defaultExpression.empty()))
                        return {VbaExecutionStatus::Unsupported,"Optional parameters must follow required parameters",line};
                    const auto type=lower(declared[4].str());
                    const bool objectDefault=(type=="collection" || type=="object" || type=="scripting.dictionary" || nativeDaoType(type)) && lower(defaultExpression)=="nothing";
                    if(optional && (type=="collection" || type=="object" || type=="scripting.dictionary" || nativeDaoType(type)) && !defaultExpression.empty() && !objectDefault)
                        return {VbaExecutionStatus::Unsupported,"Optional native object default must be Nothing",line};
                    if(optional && lower(defaultExpression)=="null" && !type.empty() && type!="variant")
                        return {VbaExecutionStatus::Unsupported,"Null defaults require a Variant parameter",line};
                    if(optional && type.starts_with("dao.") && !nativeDaoType(type))
                        return {VbaExecutionStatus::Unsupported,"Optional object parameters are not implemented",line};
                    optionalSeen=optionalSeen || optional;
                    std::shared_ptr<const VbaValue> defaultValue;
                    if(!defaultExpression.empty()) {
                        try {
                            const auto folded=lower(defaultExpression)=="null" ? VbaValue{} : lower(defaultExpression)=="empty" ? VbaValue::empty()
                                : objectDefault ? VbaValue::reference(nullptr) : constantExpression(defaultExpression,declarationConstants,constantMode);
                            defaultValue=std::make_shared<const VbaValue>(folded);
                        } catch(const VbaRuntimeError& error) { VbaExecutionResult result{VbaExecutionStatus::Unsupported,"Optional constant default unsupported",line}; result.errorNumber=error.number; return result; }
                    }
                    if (parameters) parameters->push_back({declared[3].str(), lower(declared[2].str()) != "byval",type,optional,defaultExpression,false,std::move(defaultValue)});
                }
                if(lower(trim(source)).starts_with("static ") || lower(trim(source)).starts_with("static\t"))
                    return {VbaExecutionStatus::Unsupported,"Static procedures require full implicit-local lifetime support; use explicit Static declarations",line};
                within = true;
                isFunction = lower(match[1].str()) == "function";
                const auto declaredResult=lower(match[4].str());
                if (isFunction && !declaredResult.empty() && declaredResult!="variant" && declaredResult!="boolean" && declaredResult!="string" && declaredResult!="integer" && declaredResult!="long" && declaredResult!="single" && declaredResult!="double" && declaredResult!="currency" && declaredResult!="date" && declaredResult!="byte" && declaredResult!="object" && declaredResult!="collection" && declaredResult!="scripting.dictionary" && !nativeDaoType(declaredResult))
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
                // Bind inline Exit Do only after merging its temporary If
                // wrapper into the real procedure. All loop partners exist now.
                std::vector<int> lexical;
                for (int i=0;i<static_cast<int>(statements.size());++i) {
                    auto& instruction=statements[i]; const auto type=instruction.type;
                    if (type==Statement::Type::If || type==Statement::Type::Select || type==Statement::Type::With
                        || type==Statement::Type::For || type==Statement::Type::ForEach || type==Statement::Type::While)
                        lexical.push_back(i);
                    else if (type==Statement::Type::EndIf || type==Statement::Type::EndSelect || type==Statement::Type::EndWith
                             || type==Statement::Type::Next || type==Statement::Type::EndWhile) lexical.pop_back();
                    else if (type==Statement::Type::ExitDo) {
                        auto target=lexical.rbegin();
                        while (target!=lexical.rend() && (statements[*target].type!=Statement::Type::While || statements[*target].argument=="wend")) ++target;
                        if (target==lexical.rend()) return {VbaExecutionStatus::Error,"Exit Do without Do",instruction.line};
                        instruction.partner=statements[*target].partner;
                        for (auto inner=lexical.rbegin();inner!=target;++inner) {
                            const auto nested=statements[*inner].type;
                            if (nested==Statement::Type::If || nested==Statement::Type::Select) ++instruction.conditionalUnwind;
                            else if (nested==Statement::Type::With) ++instruction.withUnwind;
                            else if (nested==Statement::Type::For || nested==Statement::Type::ForEach) ++instruction.forUnwind;
                        }
                    }
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
        const auto normalized=lower(source);
        const bool caseLine=normalized.starts_with("case ") || normalized.starts_with("case\t");
        const bool endSelectLine=std::regex_match(source,selectEnd);
        if (!blocks.empty() && blocks.back()==Statement::Type::Select && !selects.back().hasCase
            && !caseLine && !endSelectLine)
            return {VbaExecutionStatus::Error,"Statement before first Case",line};
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
                auto parsed=parseProcedure(wrapper,"__la_inline",body,function,cancel,nullptr,false,nullptr,withActive());
                if (parsed.status!=VbaExecutionStatus::Completed) { parsed.line=line; return parsed; }
                const auto offset=static_cast<int>(statements.size());
                for (auto& statement:body) { if (statement.partner>=0) statement.partner+=offset; statement.line=line; statements.push_back(std::move(statement)); }
                continue;
            }
        }
        const bool explicitLet=source.size()>=3 && lower(source.substr(0,3))=="let"
            && (source.size()==3 || std::isspace(static_cast<unsigned char>(source[3])));
        if(explicitLet) {
            source=trim(source.substr(3));
            if(source.empty()) return {VbaExecutionStatus::Error,"Let requires a value assignment",line};
        }
        const auto statementBegin=statements.size();
        const auto relativeSource=(source.starts_with('.') || source.starts_with('!')) ? "__la_with"+source : std::string{};
        if (std::regex_match(source,match,selectStart)) {
            const auto expression=trim(match[1].str());
            if (!validExpression(expression)) return {VbaExecutionStatus::Unsupported,"Select expression unsupported",line};
            statements.push_back({Statement::Type::Select,{},expression,line});
            blocks.push_back(Statement::Type::Select); selects.push_back({});
        } else if (caseLine) {
            if (blocks.empty() || blocks.back()!=Statement::Type::Select)
                return {VbaExecutionStatus::Error,"Case without Select",line};
            auto& block=selects.back();
            if (block.hasElse) return {VbaExecutionStatus::Error,"Case after Case Else",line};
            const auto expression=trim(source.substr(4));
            block.hasCase=true;
            if (lower(expression)=="else") block.hasElse=true;
            else {
                for (const auto& clause:splitArguments(expression)) {
                    const auto range=outsideKeyword(clause,"to");
                    if (outsideKeyword(clause,"is")==0 || (!clause.empty() && std::string("<>=").find(clause.front())!=std::string::npos)) {
                        const auto comparison=trim(outsideKeyword(clause,"is")==0 ? clause.substr(2) : clause);
                        const auto op=comparison.find_first_not_of("<>=");
                        const auto relation=comparison.substr(0,op);
                        if ((relation!="=" && relation!="<>" && relation!="<" && relation!="<=" && relation!=">" && relation!=">=")
                            || op==std::string::npos || !validExpression(trim(comparison.substr(op))))
                            return {VbaExecutionStatus::Unsupported,"Case relation unsupported",line};
                    } else if (range!=std::string::npos) {
                        if (!validExpression(trim(clause.substr(0,range))) || !validExpression(trim(clause.substr(range+2))))
                            return {VbaExecutionStatus::Unsupported,"Case range unsupported",line};
                    } else if (!validExpression(clause)) return {VbaExecutionStatus::Unsupported,"Case expression unsupported",line};
                }
            }
            statements.push_back({Statement::Type::Case,{},expression,line});
        } else if (endSelectLine) {
            if (blocks.empty() || blocks.back()!=Statement::Type::Select)
                return {VbaExecutionStatus::Error,"End Select without Select",line};
            blocks.pop_back(); selects.pop_back();
            statements.push_back({Statement::Type::EndSelect,{},{},line});
        } else if (lower(source).starts_with("with ") || lower(source).starts_with("with\t")) {
            const auto expression=trim(source.substr(4));
            if (!validExpression(expression)) return {VbaExecutionStatus::Unsupported,"With expression unsupported",line};
            statements.push_back({Statement::Type::With,{},expression,line}); blocks.push_back(Statement::Type::With);
        } else if (lower(source)=="end with") {
            if (blocks.empty() || blocks.back()!=Statement::Type::With) return {VbaExecutionStatus::Error,"End With without With",line};
            blocks.pop_back(); statements.push_back({Statement::Type::EndWith,{},{},line});
        } else if (lower(source).starts_with("sendkeys ") || lower(source).starts_with("sendkeys\t")) {
            const auto args=splitArguments(trim(source.substr(8)));
            if (args.empty() || args.size()>2 || args[0].empty()) return {VbaExecutionStatus::Unsupported,"SendKeys arguments unsupported",line};
            for (const auto& arg:args) if (!arg.empty() && !validExpression(arg)) return {VbaExecutionStatus::Unsupported,"SendKeys expression unsupported",line};
            statements.push_back({Statement::Type::SendKeys,{},trim(source.substr(8)),line});
        } else if (std::regex_match(source, match, label)) {
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
            if (!validExpression(expr)) return {VbaExecutionStatus::Unsupported, "If-Bedingung: " + expr, line};
            statements.push_back({Statement::Type::If, expr, {}, line});
            ++nesting;
            elseSeen.push_back(false);
            blocks.push_back(Statement::Type::If);
        } else if (std::regex_match(source,match,elseIfThen)) {
            if (blocks.empty() || blocks.back()!=Statement::Type::If || elseSeen.back())
                return {VbaExecutionStatus::Error,"ElseIf without If or after Else",line};
            const auto expression=trim(match[1].str());
            if (!validExpression(expression)) return {VbaExecutionStatus::Unsupported,"ElseIf condition unsupported",line};
            statements.push_back({Statement::Type::ElseIf,{},expression,line});
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
            if (!validExpression(expression))
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
            if (!validExpression(start) || !validExpression(endValue)
                || !validExpression(step))
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
        } else if (normalized=="do" || std::regex_match(source,match,whileStart)) {
            const auto expression=normalized=="do" ? std::string{} : trim(match[3].str());
            const auto mode=normalized=="do" ? "while" : match[2].matched ? lower(match[2].str()) : "wend";
            if (!expression.empty() && !validExpression(expression)) return {VbaExecutionStatus::Unsupported,"Loop condition unsupported",line};
            whileStarts.push_back(static_cast<int>(statements.size())); blocks.push_back(Statement::Type::While);
            statements.push_back({Statement::Type::While,mode,expression,line});
        } else if (normalized=="exit do") {
            statements.push_back({Statement::Type::ExitDo,{},{},line});
        } else if (std::regex_match(source,match,loopEnd)) {
            if (blocks.empty() || blocks.back()!=Statement::Type::While)
                return {VbaExecutionStatus::Error,"Loop/Wend without matching start",line};
            const int start=whileStarts.back();
            const bool wend=lower(match[1].str())=="wend";
            const auto expression=match[3].matched ? trim(match[3].str()) : std::string{};
            if (wend!=(statements[start].argument=="wend") || (wend && match[2].matched))
                return {VbaExecutionStatus::Error,"Loop/Wend start mismatch",line};
            if (!expression.empty() && !statements[start].expression.empty())
                return {VbaExecutionStatus::Error,"Do has both entry and exit conditions",line};
            if (!expression.empty() && !validExpression(expression)) return {VbaExecutionStatus::Unsupported,"Loop exit condition unsupported",line};
            whileStarts.pop_back(); blocks.pop_back();
            statements[start].partner=static_cast<int>(statements.size());
            statements.push_back({Statement::Type::EndWhile,match[2].matched ? lower(match[2].str()) : "while",expression,line});
            statements.back().partner=start;
        } else if (lower(source) == "on error resume next") {
            statements.push_back({Statement::Type::ErrorResumeNext, {}, {}, line});
        } else if (lower(source) == "on error goto 0") {
            statements.push_back({Statement::Type::ErrorDisable, {}, {}, line});
        } else if (const auto objectArrayAssignment=topLevelAssignment((lower(source).starts_with("set ") || lower(source).starts_with("set\t")) ? trim(source.substr(3)) : std::string{});
            objectArrayAssignment && memberSeparator(objectArrayAssignment->first)==std::string::npos && memberTarget(objectArrayAssignment->first) && memberTarget(objectArrayAssignment->first)->arguments) {
            const auto target=*memberTarget(objectArrayAssignment->first); const auto indices=splitArguments(*target.arguments);
            if(indices.empty() || indices.size()>60 || !std::all_of(indices.begin(),indices.end(),[&](const auto& index) { return !trim(index).empty() && validExpression(index); }) || !validExpression(objectArrayAssignment->second))
                return {VbaExecutionStatus::Unsupported,"Object array assignment syntax unsupported",line};
            statements.push_back({Statement::Type::IndexedAssign,target.name,objectArrayAssignment->second,line,*target.arguments,"object"});
        } else if (std::regex_match(source, match, objectSet)) {
            const auto expr = trim(match[2].str());
            if (!validExpression(expr)) return {VbaExecutionStatus::Unsupported, "Set-Ausdruck", line};
            statements.push_back({Statement::Type::ObjectSet, match[1].str(), expr, line});
        } else if (std::regex_match(source, match, message)) {
            const auto expr = "MsgBox("+trim(match[1].matched ? match[1].str() : match[2].str())+")";
            if (!validExpression(expr)) return {VbaExecutionStatus::Unsupported, "MsgBox argument syntax unsupported", line};
            statements.push_back({Statement::Type::Evaluate, {}, expr, line});
        } else if (std::regex_match(source,match,evaluation)) {
            const auto expr="Eval("+trim(match[1].matched ? match[1].str() : match[2].str())+")";
            if (!validExpression(expr)) return {VbaExecutionStatus::Unsupported,"Eval argument syntax unsupported",line};
            statements.push_back({Statement::Type::Evaluate,{},expr,line});
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
        } else if (std::regex_match(source,match,daoExecute)) {
            const auto expression="CurrentDb().Execute("+trim(match[1].str())+")";
            if (!validExpression(expression)) return {VbaExecutionStatus::Unsupported,"DAO Execute argument syntax unsupported",line};
            statements.push_back({Statement::Type::Evaluate,{},expression,line});
        } else if (std::regex_match(source, match, runSql)) {
            const auto expr = trim(match[1].str());
            const auto arguments=splitArguments(expr);
            if(arguments.size()>1 || std::any_of(arguments.begin(),arguments.end(),[](const std::string& argument){ return !callArgument(argument).name.empty(); })) {
                const auto expression="DoCmd.RunSQL("+expr+")";
                if(!validExpression(expression)) return {VbaExecutionStatus::Unsupported,"Named RunSQL syntax unsupported",line};
                statements.push_back({Statement::Type::ObjectCall,{},expression,line}); continue;
            }
            if (!validExpression(expr))
                return {VbaExecutionStatus::Unsupported, "SQL-Ausdruck: " + expr, line};
            statements.push_back({Statement::Type::ExecuteSql, {}, expr, line});
        } else if (std::regex_match(source, match, openForm)) {
            const auto args=splitArguments(match[1].str());
            if(std::any_of(args.begin(),args.end(),[](const std::string& argument){ return !callArgument(argument).name.empty(); })) {
                const auto expression="DoCmd.OpenForm("+match[1].str()+")";
                if(!validExpression(expression)) return {VbaExecutionStatus::Unsupported,"Named OpenForm syntax unsupported",line};
                statements.push_back({Statement::Type::ObjectCall,{},expression,line}); continue;
            }
            if (args.empty() || args.front().empty() || args.size()>7) return {VbaExecutionStatus::Unsupported,"OpenForm arguments unsupported",line};
            for (const auto& arg:args) if (!arg.empty() && !validExpression(arg)) return {VbaExecutionStatus::Unsupported,"OpenForm argument expression unsupported",line};
            statements.push_back({Statement::Type::OpenForm,{},match[1].str(),line});
        } else if (std::regex_match(source, match, controlRequery)) {
            auto name = match[2].str();
            if (name.starts_with('[') && name.ends_with(']')) name = name.substr(1, name.size() - 2);
            statements.push_back({Statement::Type::ControlRequery, name, {}, line,match[1].str()});
        } else if (std::regex_match(source, match, controlFocus)) {
            auto name = match[2].str();
            if (name.starts_with('[') && name.ends_with(']')) name = name.substr(1, name.size() - 2);
            statements.push_back({Statement::Type::ControlFocus, name, {}, line,match[1].str()});
        } else if (std::regex_match(source, match, controlVisible)) {
            auto name = match[2].str();
            if (name.starts_with('[') && name.ends_with(']')) name = name.substr(1, name.size() - 2);
            const auto expr = trim(match[3].str());
            if (!validExpression(expr))
                return {VbaExecutionStatus::Unsupported, "Visible-Ausdruck: " + expr, line};
            statements.push_back({Statement::Type::ControlVisible, name, expr, line,match[1].str()});
        } else if (std::regex_match(source, match, call)) {
            const auto args = match[2].str();
            if (lower(match[1].str())=="msgbox" || lower(match[1].str())=="eval") {
                const auto expression=match[1].str()+"("+args+")";
                if (!validExpression(expression)) return {VbaExecutionStatus::Unsupported,"Builtin argument syntax unsupported",line};
                statements.push_back({Statement::Type::Evaluate,{},expression,line}); continue;
            }
            for (const auto& arg : splitArguments(args))
                if(const auto parsed=callArgument(arg); (!parsed.expression.empty() && !validExpression(parsed.expression)) || (!parsed.name.empty() && parsed.expression.empty()))
                    return {VbaExecutionStatus::Unsupported,"Aufrufargument",line};
            statements.push_back({Statement::Type::Call, match[1].str(), args, line});
        } else if (std::regex_match(source, match, assignment)) {
            const auto expr = trim(match[3].str());
            if (!validExpression(expr)) return {VbaExecutionStatus::Unsupported, "Zuweisung: " + expr, line};
            auto field = match[2].str();
            if (field.starts_with('[') && field.ends_with(']')) field = field.substr(1, field.size() - 2);
            statements.push_back({Statement::Type::SetField,field,expr,line,match[1].str()});
        } else if (const auto assignment=objectAssignment(relativeSource); !relativeSource.empty() && assignment) {
            const auto target=assignment->first.substr(9),expr=assignment->second;
            if (!validExpression(target) || !validExpression(expr)) return {VbaExecutionStatus::Unsupported,"With assignment unsupported",line};
            statements.push_back({Statement::Type::ObjectAssign,target,expr,line});
        } else if (const auto assignment=objectAssignment(source); assignment) {
            const auto& target=assignment->first;
            const auto& expr=assignment->second;
            if (!validExpression(target) || !validExpression(expr))
                return {VbaExecutionStatus::Unsupported, "Objektzuweisung", line};
            statements.push_back({Statement::Type::ObjectAssign, target, expr, line});
        } else if (const auto assignment=topLevelAssignment(source); assignment && memberSeparator(assignment->first)==std::string::npos && memberTarget(assignment->first) && memberTarget(assignment->first)->arguments) {
            const auto target=*memberTarget(assignment->first);
            const auto indices=splitArguments(*target.arguments);
            if(indices.empty() || indices.size()>60 || !std::all_of(indices.begin(),indices.end(),[&](const auto& index) { return !trim(index).empty() && validExpression(index); }) || !validExpression(assignment->second))
                return {VbaExecutionStatus::Unsupported,"Array assignment syntax unsupported",line};
            statements.push_back({Statement::Type::IndexedAssign,target.name,assignment->second,line,*target.arguments});
        } else if (lower(source).starts_with("redim ") || lower(source).starts_with("redim\t")) {
            auto declaration=trim(source.substr(5));
            bool preserve=false;
            if(lower(declaration).starts_with("preserve ") || lower(declaration).starts_with("preserve\t")) {
                preserve=true; declaration=trim(declaration.substr(8));
            }
            std::string elementType;
            if(const auto as=outsideKeyword(declaration,"as");as!=std::string::npos) {
                elementType=lower(trim(declaration.substr(as+2))); declaration=trim(declaration.substr(0,as));
                if(elementType!="variant" && elementType!="byte" && elementType!="integer" && elementType!="long" && elementType!="single"
                    && elementType!="double" && elementType!="currency" && elementType!="date" && elementType!="boolean" && elementType!="string" && elementType!="object")
                    return {VbaExecutionStatus::Unsupported,"ReDim element type unsupported",line};
            }
            const auto target=memberTarget(declaration);
            if(!target || !target->arguments) return {VbaExecutionStatus::Unsupported,"Invalid Variant ReDim target",line};
            const auto dimensions=splitArguments(*target->arguments);
            if(dimensions.empty() || dimensions.size()>60) return {VbaExecutionStatus::Unsupported,"Invalid ReDim dimension count",line};
            for(const auto& dimension:dimensions) {
                const auto range=outsideKeyword(dimension,"to");
                const auto lowerBound=range==std::string::npos ? std::string{} : trim(dimension.substr(0,range));
                const auto upperBound=range==std::string::npos ? dimension : trim(dimension.substr(range+2));
                if((!lowerBound.empty() && !validExpression(lowerBound)) || !validExpression(upperBound))
                    return {VbaExecutionStatus::Unsupported,"Invalid ReDim bounds",line};
            }
            statements.push_back({Statement::Type::ResizeArray,target->name,*target->arguments,line,elementType,preserve ? "preserve" : ""});
        } else if (lower(source).starts_with("erase ") || lower(source).starts_with("erase\t")) {
            const auto target=memberTarget(trim(source.substr(5)));
            if(!target || target->arguments) return {VbaExecutionStatus::Unsupported,"Only a single Variant array Erase is implemented",line};
            statements.push_back({Statement::Type::EraseArray,target->name,{},line});
        } else if(lower(source).starts_with("const ") || lower(source).starts_with("const\t")) {
            try {
                for(auto& constant:constantDeclarations(trim(source.substr(5)),declarationConstants,constantMode)) {
                    if(!localConstants.insert(constant.name).second) return {VbaExecutionStatus::Unsupported,"Duplicate local constant",line};
                    Statement declaration{Statement::Type::Dim,constant.name,constant.type,line};
                    declaration.constantValue=std::make_shared<const VbaValue>(std::move(constant.value)); statements.push_back(std::move(declaration));
                }
            } catch(const VbaRuntimeError& error) { VbaExecutionResult result{VbaExecutionStatus::Unsupported,"Const declaration unsupported",line}; result.errorNumber=error.number; return result; }
        } else if (std::regex_match(source, match, dim)) {
            const bool persistent=lower(source.substr(0,6))=="static";
            const std::regex dynamicArray(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*\(\s*(.*)\s*\)\s*(?:As\s+([A-Za-z]+))?$)",std::regex::icase);
            for (auto entry:splitArguments(match[1].str())) {
                std::smatch declared;
                entry = trim(entry);
                if(std::regex_match(entry,declared,dynamicArray)) {
                    auto type=lower(declared[3].str()); if(type.empty()) type="variant";
                    if(type!="variant" && type!="byte" && type!="integer" && type!="long" && type!="single" && type!="double"
                        && type!="currency" && type!="date" && type!="boolean" && type!="string" && type!="object")
                        return {VbaExecutionStatus::Unsupported,"Dynamic array element type unsupported",line};
                    const auto bounds=trim(declared[2].str());
                    const auto dimensions=bounds.empty() ? std::optional<std::vector<VbaArrayDimension>>{} : literalArrayBounds(bounds,moduleArrayBase(module),declarationConstants,constantMode);
                    if(!bounds.empty() && !dimensions) return {VbaExecutionStatus::Unsupported,"Fixed array bounds require supported constant expressions",line};
                    Statement declaration{Statement::Type::Dim,declared[1].str(),type+"()",line,bounds,bounds.empty() ? "" : "fixed"};
                    declaration.persistent=persistent;
                    if(dimensions) declaration.fixedBounds=*dimensions; statements.push_back(std::move(declaration));
                    continue;
                }
                if (!std::regex_match(entry, declared, declaration))
                    return {VbaExecutionStatus::Unsupported, "Variablendeklaration", line};
                const auto type = lower(declared[2].str());
                if (!type.empty() && type != "variant" && type != "object" && type != "collection" && type != "scripting.dictionary" && !nativeDaoType(type)
                    && type != "control" && type != "form" && type != "access.control" && type != "access.form"
                    && type != "string" && type != "boolean"
                    && type != "integer" && type != "long" && type != "single" && type != "double" && type != "currency" && type != "date"
                    && type != "byte") return {VbaExecutionStatus::Unsupported, "Variablentyp: " + type, line};
                Statement local{Statement::Type::Dim, declared[1].str(), type, line};
                local.persistent=persistent; statements.push_back(std::move(local));
            }
        } else if (std::regex_match(source, match, variableAssignment)) {
            const auto expr = trim(match[2].str());
            if (!validExpression(expr)) return {VbaExecutionStatus::Unsupported, "Zuweisung: " + expr, line};
            statements.push_back({Statement::Type::SetVar, match[1].str(), expr, line});
        } else if (lower(source) == "exit sub" || lower(source) == "exit function") {
            statements.push_back({Statement::Type::Exit, {}, {}, line});
        } else if (!relativeSource.empty() && std::regex_match(relativeSource,match,memberCall)) {
            const auto expression=match[1].str().substr(9)+"("+match[2].str()+")";
            if (!validExpression(expression)) return {VbaExecutionStatus::Unsupported,"With method arguments unsupported",line};
            statements.push_back({Statement::Type::ObjectCall,{},expression,line});
        } else if (std::regex_match(source,match,memberCall)) {
            const auto expression = match[1].str() + "(" + match[2].str() + ")";
            if (!validExpression(expression)) return {VbaExecutionStatus::Unsupported,"Object method arguments",line};
            statements.push_back({Statement::Type::ObjectCall,{},expression,line});
        } else if (const auto assignment=topLevelAssignment(source); assignment
            && memberSeparator(assignment->first)!=std::string::npos
            && objectPath(trim(assignment->first.substr(0,memberSeparator(assignment->first))))) {
            return {VbaExecutionStatus::Unsupported,"Invalid property assignment target",line};
        } else if (source.find('.') != std::string::npos && validExpression(source)) {
            statements.push_back({Statement::Type::ObjectCall, {}, source, line});
        } else if (std::regex_match(source, match, implicitCall)) {
            const auto args = match[2].matched ? match[2].str() : match[3].str();
            for (const auto& arg : splitArguments(args))
                if(const auto parsed=callArgument(arg); (!parsed.expression.empty() && !validExpression(parsed.expression)) || (!parsed.name.empty() && parsed.expression.empty()))
                    return {VbaExecutionStatus::Unsupported,"Aufrufargument",line};
            statements.push_back({Statement::Type::Call, match[1].str(), args, line});
        } else {
            return {VbaExecutionStatus::Unsupported, "VBA-Anweisung noch nicht unterstützt: " + source, line};
        }
        if(explicitLet) {
            if(statements.size()!=statementBegin+1) return {VbaExecutionStatus::Error,"Let requires one value assignment",line};
            const auto type=statements.back().type;
            if(type!=Statement::Type::SetField && type!=Statement::Type::SetVar
                && type!=Statement::Type::IndexedAssign && type!=Statement::Type::ObjectAssign)
                return {VbaExecutionStatus::Error,"Let requires a value assignment",line};
        }
    }
    return within ? VbaExecutionResult{VbaExecutionStatus::Error, "End Sub fehlt", line}
                  : VbaExecutionResult{VbaExecutionStatus::ProcedureMissing, {}, 0};
}
// Cache source-only translation plans, never host/module/DAO/document state.
struct ProcedurePlan {
    std::string module,name,resultType;
    std::size_t moduleHash{},cost{};
    bool resolveLabels{},inheritedWith{},function{},cancel{};
    std::vector<Statement> statements;
    std::vector<Parameter> parameters;
    VbaExecutionResult result;
};
VbaExecutionResult parseProcedure(std::string_view module,std::string_view name,
                                 std::vector<Statement>& statements,bool& function,bool& cancel,
                                 std::vector<Parameter>* parameters,bool resolveLabels,std::string* resultType,bool inheritedWith) {
    if(!statements.empty()) return parseProcedureUncached(module,name,statements,function,cancel,parameters,resolveLabels,resultType,inheritedWith);
    constexpr std::size_t budget=16*1024*1024,entryLimit=128;
    static thread_local std::list<ProcedurePlan> plans;
    static thread_local std::size_t accounted=0;
    const auto normalized=lower(std::string(name)); const auto hash=std::hash<std::string_view>{}(module);
    const auto copy=[&](const ProcedurePlan& plan) {
        statements=plan.statements; function=plan.function; cancel=plan.cancel;
        if(parameters) *parameters=plan.parameters;
        if(resultType) *resultType=plan.resultType;
    };
    const auto found=std::find_if(plans.begin(),plans.end(),[&](const ProcedurePlan& plan) {
        return plan.moduleHash==hash && plan.name==normalized && plan.resolveLabels==resolveLabels
            && plan.inheritedWith==inheritedWith && plan.module==module;
    });
    if(found!=plans.end()) {
        copy(*found); const auto result=found->result; plans.splice(plans.begin(),plans,found); return result;
    }
    std::vector<Parameter> inferredParameters; std::string inferredType;
    auto* parsedParameters=parameters ? parameters : &inferredParameters;
    auto* parsedType=resultType ? resultType : &inferredType;
    const auto result=parseProcedureUncached(module,name,statements,function,cancel,parsedParameters,resolveLabels,parsedType,inheritedWith);
    if(result.status!=VbaExecutionStatus::Completed || result.typedReturnValue || result.returnValue) return result;
    ProcedurePlan plan; plan.result=result; plan.function=function; plan.cancel=cancel;
    plan.moduleHash=hash; plan.resolveLabels=resolveLabels; plan.inheritedWith=inheritedWith;
    const auto add=[&](std::size_t amount) {
        if(plan.cost>budget || amount>budget-plan.cost) plan.cost=budget+1; else plan.cost+=amount;
    };
    const auto scalar=[&](const std::shared_ptr<const VbaValue>& value) {
        if(!value) return;
        if(value->object || isVbaArray(*value)) { plan.cost=budget+1; return; }
        add(sizeof(VbaValue)+64+value->text.capacity()+value->bytes.capacity());
    };
    add(sizeof(ProcedurePlan)+2*sizeof(void*)+module.size()+normalized.size()+parsedType->size()+result.explanation.capacity());
    add(statements.capacity()*sizeof(Statement)); add(parsedParameters->capacity()*sizeof(Parameter));
    for(const auto& instruction:statements) {
        add(instruction.argument.capacity()+instruction.expression.capacity()+instruction.endExpression.capacity()+instruction.stepExpression.capacity());
        add(instruction.fixedBounds.capacity()*sizeof(VbaArrayDimension)); scalar(instruction.constantValue);
    }
    for(const auto& parameter:*parsedParameters) {
        add(parameter.name.capacity()+parameter.type.capacity()+parameter.defaultExpression.capacity()); scalar(parameter.defaultValue);
    }
    if(plan.cost>budget) return result;
    try {
        plan.module=std::string(module); plan.name=normalized; plan.resultType=*parsedType;
        plan.statements=statements; plan.parameters=*parsedParameters;
        while(!plans.empty() && (plans.size()>=entryLimit || accounted>budget-plan.cost)) { accounted-=plans.back().cost; plans.pop_back(); }
        plans.push_front(std::move(plan)); accounted+=plans.front().cost;
    } catch(const std::bad_alloc&) { /* The optional cache must not reject a translated procedure. */ }
    return result;
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
    const std::regex variable(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)(?:\s+As\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*(?:\.[A-Za-z_][A-Za-z0-9_]*)?))?$)",std::regex::icase);
    const std::regex dynamicArray(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*\(\s*(.*)\s*\)\s*(?:As\s+([A-Za-z]+))?$)",std::regex::icase);
    const std::regex constant(R"(^\s*(?:(Public|Private)\s+)?Const\s+(.+)$)",std::regex::icase);
    VbaVariables knownConstants; const auto constantMode=moduleCompareMode(module);
    std::istringstream stream{std::string(module)}; std::string line;
    while (std::getline(stream,line)) {
        line=withoutComment(line); if (std::regex_search(line,procedure)) break;
        std::smatch match;
        if(std::regex_match(line,match,constant)) {
            const bool publicConstant=lower(match[1].str())=="public";
            for(auto& item:constantDeclarations(match[2].str(),knownConstants,constantMode)) {
                if(state->values.contains(item.name)) throw VbaRuntimeError(5,"Duplicate module constant");
                state->values.emplace(item.name,std::move(item.value)); state->types.emplace(item.name,item.type); state->constants.insert(item.name);
                if(publicConstant) state->publicVariables.insert(item.name);
            }
            continue;
        }
        if (!std::regex_match(line,match,declaration)) continue;
        const auto visibility=lower(match[1].str());
        const auto declarations=match[2].str();
        const auto start=lower(trim(declarations));
        if (start.starts_with("declare ") || start.starts_with("type ") || start.starts_with("enum ")) continue;
        for (const auto& entry:splitArguments(declarations)) {
            std::smatch field;
            if(std::regex_match(entry,field,dynamicArray)) {
                const auto name=lower(field[1].str()); auto type=lower(field[3].str()); if(type.empty()) type="variant";
                if(type!="variant" && type!="byte" && type!="integer" && type!="long" && type!="single" && type!="double"
                    && type!="currency" && type!="date" && type!="boolean" && type!="string" && type!="object")
                    throw VbaRuntimeError(3251,"Module dynamic array element type unsupported");
                if(state->values.contains(name)) throw VbaRuntimeError(5,"Duplicate module field");
                const auto bounds=trim(field[2].str()); auto value=makeUnallocatedVbaArray(type);
                if(!bounds.empty()) {
                    const auto dimensions=literalArrayBounds(bounds,moduleArrayBase(module),knownConstants,constantMode);
                    if(!dimensions) throw VbaRuntimeError(3251,"Fixed module array bounds require supported decimal literals");
                    value=resizeVbaArrayWithBounds(value,*dimensions,false,type); state->fixedArrays.insert(name);
                }
                state->values.emplace(name,std::move(value)); state->types.emplace(name,type+"()");
                if(visibility=="public" || visibility=="global") state->publicVariables.insert(name);
                continue;
            }
            if (!std::regex_match(entry,field,variable)) throw VbaRuntimeError(3251,"Module declaration unsupported");
            const auto name=lower(field[1].str()),type=lower(field[2].str());
            if (!type.empty() && type!="variant" && type!="boolean" && type!="string" && type!="integer" && type!="long" && type!="single" && type!="double" && type!="currency" && type!="date" && type!="byte" && type!="collection" && type!="object" && type!="scripting.dictionary" && !nativeDaoType(type))
                throw VbaRuntimeError(3251,"Module field type unsupported");
            if (state->values.contains(name)) throw VbaRuntimeError(5,"Duplicate module field");
            state->values.emplace(name,(type=="collection" || type=="object" || type=="scripting.dictionary" || nativeDaoType(type)) ? VbaValue::reference(nullptr) : type=="boolean" ? VbaValue::logical(false) : type=="string" ? VbaValue::string("") : type=="integer" ? VbaValue::integer(0) : type=="long" ? VbaValue::longInteger(0) : type=="single" ? VbaValue::single(0) : type=="double" ? VbaValue::numeric(0) : type=="currency" ? VbaValue::currency(0) : type=="date" ? VbaValue::dateSerial(0) : type=="byte" ? VbaValue::byte(0) : VbaValue::empty());
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
    if(type.ends_with("()")) {
        if(!isVbaArray(value) || lower(vbaArrayTypeName(value))!=type) throw VbaRuntimeError(13,"Dynamic array element type mismatch");
        return value;
    }
    if (type.empty() || type=="variant") return value;
    if(nativeDaoType(type)) {
        if(value.kind!=VbaValue::Kind::Object || (value.object && lower(std::string(value.object->className()))!=type.substr(4)))
            throw VbaRuntimeError(13,"DAO binding requires the declared native object type");
        return value;
    }
    if(type=="scripting.dictionary") {
        if(value.kind!=VbaValue::Kind::Object || (value.object && value.object->className()!="Dictionary"))
            throw VbaRuntimeError(13,"Dictionary binding requires a native Dictionary object");
        return value;
    }
    if(type=="object") {
        if(value.kind!=VbaValue::Kind::Object) throw VbaRuntimeError(13,"Object binding requires an object reference");
        return value;
    }
    if(type=="collection") {
        if(value.kind!=VbaValue::Kind::Object || (value.object && value.object->className()!="Collection"))
            throw VbaRuntimeError(13,"Collection binding requires a Collection object");
        return value;
    }
    if (value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null module field assignment");
    if (value.kind==VbaValue::Kind::Error || value.kind==VbaValue::Kind::Object) throw VbaRuntimeError(13,"Module field type mismatch");
    if (type=="string") {
        if (value.kind==VbaValue::Kind::Empty) return VbaValue::string("");
        if (value.kind==VbaValue::Kind::Boolean) return VbaValue::string(value.boolean ? "True" : "False");
        return VbaValue::string(value.asText());
    }
    if (type=="currency") return convertCurrency(value);
    if (type=="date") return vbaDateValue(value);
    if (type=="boolean") {
        if (value.kind==VbaValue::Kind::Text) {
            const auto text=lower(trim(value.text));
            if (text=="true" || text=="false") return VbaValue::logical(text=="true");
            return VbaValue::logical(VbaValue::string(trim(value.text)).asNumber()!=0);
        }
        try { return VbaValue::logical(value.asNumber()!=0); }
        catch (...) { throw VbaRuntimeError(13,"Module Boolean field type mismatch"); }
    }
    if(type=="single" || type=="double") {
        const auto number=value.kind==VbaValue::Kind::Text ? VbaValue::string(trim(value.text)).asNumber() : value.asNumber();
        if(!std::isfinite(number)) throw VbaRuntimeError(6,"Floating field overflow");
        if(type=="double") return VbaValue::numeric(number);
        if(std::abs(number)>std::numeric_limits<float>::max()) throw VbaRuntimeError(6,"Single field overflow");
        if(number!=0 && std::abs(number)<std::numeric_limits<float>::denorm_min()) throw VbaRuntimeError(3251,"Single underflow conversion is not validated");
        return VbaValue::single(static_cast<float>(number));
    }
    if(type=="byte" && value.kind==VbaValue::Kind::Boolean) return VbaValue::byte(value.boolean ? 255 : 0);
    if(type=="byte" || type=="integer" || type=="long") {
        const auto number=value.kind==VbaValue::Kind::Text ? VbaValue::string(trim(value.text)).asNumber() : value.asNumber();
        if(!std::isfinite(number)) throw VbaRuntimeError(6,"Module numeric field overflow");
        const auto lower=std::floor(static_cast<long double>(number));
        const auto fraction=static_cast<long double>(number)-lower;
        const auto rounded=fraction<0.5L ? lower : fraction>0.5L ? lower+1 : std::fmod(std::abs(lower),2.0L)==0 ? lower : lower+1;
        const auto minimum=type=="byte" ? 0.0L : type=="integer" ? -32768.0L : -2147483648.0L;
        const auto maximum=type=="byte" ? 255.0L : type=="integer" ? 32767.0L : 2147483647.0L;
        if(rounded<minimum || rounded>maximum) throw VbaRuntimeError(6,"Module numeric field overflow");
        return type=="byte" ? VbaValue::byte(static_cast<unsigned char>(rounded)) : type=="integer" ? VbaValue::integer(static_cast<std::int16_t>(rounded)) : VbaValue::longInteger(static_cast<std::int32_t>(rounded));
    }
    throw VbaRuntimeError(3251,"Module field coercion unsupported");
}
VbaOpenFormArguments parseVbaOpenFormArguments(const std::vector<VbaValue>& args) {
    if (args.empty()) throw VbaRuntimeError(450,"OpenForm requires a form name");
    if (args.size()>7) throw VbaRuntimeError(450,"Too many OpenForm arguments");
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
    for(std::size_t index=4;index<std::min<std::size_t>(6,args.size());++index) {
        if(args[index].kind==VbaValue::Kind::Empty) continue;
        const auto mode=coerceVbaModuleValue(scalarVbaValue(args[index]),"long").number;
        if(index==4 && (mode==0 || mode==1 || mode==2)) result.dataMode=static_cast<int>(mode);
        else if(index==5 && (mode==1 || mode==2 || mode==3)) result.windowMode=static_cast<int>(mode);
        else if(mode!=(index==4 ? -1 : 0)) throw VbaRuntimeError(3251,"OpenForm data/window mode is unsupported");
    }
    if(args.size()>6) {
        const auto value=scalarVbaValue(args[6]);
        if(value.kind!=VbaValue::Kind::Null && value.kind!=VbaValue::Kind::Empty) {
            if(value.kind!=VbaValue::Kind::Text) throw VbaRuntimeError(3251,"OpenArgs currently requires text or Null");
            (void)vbaStringCodeUnits(value.text); result.openArgs=value.text;
        }
    }
    return result;
}
void VbaHost::openFormArguments(const std::vector<VbaValue>& args) {
    const auto request=parseVbaOpenFormArguments(args);
    if(request.dataMode!=-1) throw VbaRuntimeError(3251,"Host does not support OpenForm data modes");
    if(request.windowMode!=0) throw VbaRuntimeError(3251,"Host does not support OpenForm window modes");
    if(request.openArgs) throw VbaRuntimeError(3251,"Host does not support OpenArgs");
    if (!request.where.empty()) throw VbaRuntimeError(3251,"Host does not support a form WhereCondition");
    openForm(request.name);
}
namespace {
VbaValue initialLocalValue(const std::string& type) {
    if(type.ends_with("()")) return makeUnallocatedVbaArray(type.substr(0,type.size()-2));
    if (type=="object" || type=="collection" || type=="scripting.dictionary" || type=="control" || type=="form" || type.starts_with("access.") || type.starts_with("dao.")) return VbaValue::reference(nullptr);
    if (type=="string") return VbaValue::string("");
    if (type=="boolean") return VbaValue::logical(false);
    if (type=="byte") return VbaValue::byte(0);
    if (type=="integer") return VbaValue::integer(0);
    if (type=="long") return VbaValue::longInteger(0);
    if (type=="single") return VbaValue::single(0);
    if (type=="currency") return VbaValue::currency(0);
    if (type=="date") return VbaValue::dateSerial(0);
    if (!type.empty() && type!="variant") return VbaValue::numeric(0);
    return VbaValue::empty();
}
struct ReferenceBinding { VbaValue* source{}; std::string type; bool indexedArray{}; bool paramArray{}; bool fixedArray{}; };
}

VbaExecutionScope::VbaExecutionScope() { activeExecutionEpochs.push_back(activeExecutionEpochs.empty() ? stopEpoch : activeExecutionEpochs.back()); }
VbaExecutionScope::~VbaExecutionScope() { activeExecutionEpochs.pop_back(); }
std::uint64_t vbaExecutionEpoch() noexcept { return stopEpoch; }
bool vbaExecutionStopRequested() noexcept { return !activeExecutionEpochs.empty() && activeExecutionEpochs.back()!=stopEpoch; }
void requestVbaExecutionStop() noexcept { ++stopEpoch; }

static VbaExecutionResult executeProcedure(std::string_view module,
                                       std::string_view procedure,
                                       VbaHost& host, std::vector<VbaValue>& arguments,
                                       const std::vector<ReferenceBinding>* references = nullptr) {
    VbaExecutionScope executionFrame;
    if(vbaExecutionStopRequested()) return {VbaExecutionStatus::Completed,{},0,true};
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
            if(declaration.status==VbaExecutionStatus::ProcedureMissing && lower(instruction.argument)=="doevents") continue;
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
    struct Frame { bool parent; bool condition; bool inElse; VbaValue selector; bool select{}; bool pending{}; std::size_t instruction{}; };
    struct LoopFrame {
        int forIndex;
        double end{};
        double step{};
        std::vector<VbaValue> items;
        std::size_t itemIndex{};
        std::size_t frameDepth{};
        std::size_t withDepth{};
        VbaValue array;
        std::string arrayRoot;
        std::shared_ptr<void> enumerationLease;
        std::shared_ptr<VbaObject> enumerationSource;
    };
    std::vector<Frame> frames;
    std::vector<LoopFrame> loops;
    struct WithFrame { std::size_t instruction; VbaValue value; bool pending{}; };
    std::vector<WithFrame> withFrames,faultWithFrames;
    VbaVariables locals;
    std::unordered_map<std::string,std::string> localTypes;
    std::unordered_set<std::string> fixedArrays;
    std::unordered_set<std::string> constantNames;
    const int arrayBase=moduleArrayBase(module);
    const auto integralType=[](std::string_view type) { return type=="integer" || type=="long"; };
    const auto checkedType=[&](std::string_view type) { return type.ends_with("()") || integralType(type) || type=="boolean" || type=="string" || type=="single" || type=="double" || type=="currency" || type=="date" || type=="byte" || type=="collection" || type=="object" || type=="scripting.dictionary" || nativeDaoType(type); };
    const auto scalarType=[&](std::string_view type) { return checkedType(type) && type!="collection" && type!="object" && type!="scripting.dictionary" && !nativeDaoType(type) && !type.ends_with("()"); };
    const auto objectType=[](std::string_view type) { return type=="object" || type=="collection" || type=="scripting.dictionary" || type=="control" || type=="form" || type.starts_with("dao.") || type.starts_with("access."); };
    auto errorState = std::make_shared<ErrorObject>();
    locals["err"] = VbaValue::reference(errorState);
    bool resumeNext = false;
    int handler = -1;
    bool handlingError = false;
    std::size_t faultInstruction = 0;
    std::vector<Frame> faultFrames;
    std::vector<LoopFrame> faultLoops;
    bool faultActive = true;
    if (isFunction) {
        const auto name=lower(std::string(procedure));
        locals[name]=initialLocalValue(resultType); localTypes[name]=resultType;
    }
    const bool implicitCancel = hasCancel && !isFunction && arguments.empty();
    const auto required=static_cast<std::size_t>(std::count_if(parameters.begin(),parameters.end(),[](const Parameter& parameter){ return !parameter.optional && !parameter.paramArray; }));
    const bool variadic=!parameters.empty() && parameters.back().paramArray;
    if (!implicitCancel && (arguments.size()<required || (!variadic && arguments.size()>parameters.size()))) {
        VbaExecutionResult failure{VbaExecutionStatus::Error,"Falsche Argumentanzahl",0};
        failure.errorNumber=450; return failure;
    }
    std::vector<bool> omittedParameters(parameters.size(),false);
    try {
        for(std::size_t i=0;i<parameters.size();++i) {
            const auto& parameter=parameters[i]; const auto name=lower(parameter.name);
            if(parameter.paramArray) {
                for(auto position=i;position<arguments.size();++position) if(arguments[position].kind==VbaValue::Kind::Missing)
                    throw VbaRuntimeError(450,"ParamArray argument holes are not implemented");
                locals[name]=makeVbaArray(std::vector<VbaValue>(arguments.begin()+static_cast<std::ptrdiff_t>(i),arguments.end()),0);
                localTypes[name]="paramarray";
                continue;
            }
            const bool omitted=!implicitCancel && (i>=arguments.size() || arguments[i].kind==VbaValue::Kind::Missing);
            omittedParameters[i]=omitted;
            if(omitted && !parameter.optional) throw VbaRuntimeError(450,"Required argument omitted");
            if(!omitted && references && i<references->size() && parameter.byReference && (*references)[i].paramArray)
                throw VbaRuntimeError(3251,"ByRef ParamArray forwarding is not implemented");
            if(!omitted && references && i<references->size() && !implicitCancel && parameter.byReference && (*references)[i].indexedArray)
                throw VbaRuntimeError(3251,"ByRef array element binding is not implemented");
            const auto value=implicitCancel ? VbaValue::numeric(0) : !omitted ? arguments[i]
                : parameter.defaultValue ? *parameter.defaultValue
                : parameter.type.empty() || parameter.type=="variant" ? VbaValue::missing() : initialLocalValue(parameter.type);
            if(!omitted && references && i<references->size() && !parameter.byReference && (*references)[i].type.ends_with("()"))
                throw VbaRuntimeError(3251,"Declared array arguments require ByRef or an explicit Variant/value copy");
            const auto converted=checkedType(parameter.type) ? coerceVbaModuleValue(value,parameter.type) : value;
            if(!omitted && references && i<references->size() && !implicitCancel && parameter.byReference && (*references)[i].source) {
                const auto& sourceType=(*references)[i].type;
                if(sourceType.ends_with("()") && !parameter.type.ends_with("()")) throw VbaRuntimeError(3251,"Declared arrays require a matching array parameter or explicit Variant copy");
                if((checkedType(parameter.type) || checkedType(sourceType)) && !sourceType.empty() && sourceType!="variant"
                    && !parameter.type.empty() && parameter.type!="variant" && sourceType!=parameter.type)
                    throw VbaRuntimeError(13,"Known ByRef argument types do not match");
                if(checkedType(parameter.type) && (converted.kind!=value.kind
                    || (converted.kind==VbaValue::Kind::Number && (converted.number!=value.number || converted.numericType!=value.numericType))
                    || (converted.kind==VbaValue::Kind::Text && converted.text!=value.text)
                    || (converted.kind==VbaValue::Kind::Boolean && converted.boolean!=value.boolean)))
                    throw VbaRuntimeError(13,"ByRef argument requires an unchanged value representation");
            }
            locals[name]=converted; localTypes[name]=parameter.type;
            if(!omitted && references && i<references->size() && parameter.byReference && (*references)[i].fixedArray) fixedArrays.insert(name);
        }
    } catch(const VbaRuntimeError& error) {
        VbaExecutionResult failure{VbaExecutionStatus::Error,"Parameter conversion failed",0}; failure.errorNumber=error.number; return failure;
    }
    // Declarations bind for the whole procedure, including skipped branches.
    for (const auto& instruction:instructions) if (instruction.type==Statement::Type::Dim) {
        const auto name=lower(instruction.argument);
        if(locals.contains(name)) {
            if(instruction.persistent || instruction.constantValue || constantNames.contains(name)) { VbaExecutionResult result{VbaExecutionStatus::Unsupported,"Duplicate persistent or constant binding",instruction.line}; result.errorNumber=3251; return result; }
            continue;
        }
        try {
            auto value=instruction.constantValue ? *instruction.constantValue : initialLocalValue(instruction.expression);
            if(instruction.constantValue) constantNames.insert(name);
            if(instruction.stepExpression=="fixed") {
                if(instruction.fixedBounds.empty()) throw VbaRuntimeError(3251,"Fixed array bounds unsupported");
                value=resizeVbaArrayWithBounds(value,instruction.fixedBounds,false,instruction.expression.substr(0,instruction.expression.size()-2));
                fixedArrays.insert(name);
            }
            locals.emplace(name,std::move(value)); localTypes[name]=instruction.expression;
        } catch(const VbaRuntimeError& error) {
            VbaExecutionResult failure{VbaExecutionStatus::Error,"Array declaration failed",instruction.line}; failure.errorNumber=error.number; return failure;
        }
    }
    struct Alias { std::string name; VbaValue* source; VbaValue before; std::string type; };
    std::vector<Alias> aliases;
    if (references && !implicitCancel)
        for (std::size_t i = 0; i < parameters.size() && i<references->size(); ++i)
            if (!omittedParameters[i] && parameters[i].byReference && (*references)[i].source)
                aliases.push_back({lower(parameters[i].name), (*references)[i].source, *(*references)[i].source,(*references)[i].type});
    std::vector<std::shared_ptr<VbaModuleState>> stateOwners;
    try {
        auto state=host.moduleState(module); stateOwners.push_back(state);
        for(const auto& instruction:instructions) if(instruction.type==Statement::Type::Dim && instruction.persistent) {
            const auto name=lower(instruction.argument);
            auto& saved=state->procedureStatics[lower(std::string(procedure))];
            auto [slot,inserted]=saved.try_emplace(name,locals.at(name));
            locals[name]=slot->second;
            aliases.push_back({name,&slot->second,slot->second,instruction.expression});
        }
        for (auto& [name,value]:state->values) if (!locals.contains(name)) {
            locals.emplace(name,value); aliases.push_back({name,&value,value,state->types.at(name)});
            if(state->fixedArrays.contains(name)) fixedArrays.insert(name);
            if(state->constants.contains(name)) constantNames.insert(name);
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
                if(global->fixedArrays.contains(name)) fixedArrays.insert(name);
                if(global->constants.contains(name)) constantNames.insert(name);
            }
        }
    } catch (const VbaRuntimeError& error) {
        VbaExecutionResult failure{VbaExecutionStatus::Error,"Module state initialization failed",0}; failure.errorNumber=error.number; return failure;
    }
    const auto variableType=[&](const std::string& name) -> std::string {
        if(const auto found=localTypes.find(name); found!=localTypes.end()) return found->second;
        for(const auto& alias:aliases) if(alias.name==name) return alias.type;
        return {};
    };
    for(const auto& instruction:instructions) {
        if(constantNames.contains(lower(instruction.argument)) && (instruction.type==Statement::Type::SetVar || instruction.type==Statement::Type::ObjectSet
            || instruction.type==Statement::Type::IndexedAssign || instruction.type==Statement::Type::ResizeArray || instruction.type==Statement::Type::EraseArray
            || instruction.type==Statement::Type::For || instruction.type==Statement::Type::ForEach)) {
            VbaExecutionResult result{VbaExecutionStatus::Unsupported,"Assignment to a constant is not permitted",instruction.line}; result.errorNumber=3251; return result;
        }
        const auto type=variableType(lower(instruction.argument));
        if(((instruction.type==Statement::Type::ForEach || instruction.type==Statement::Type::ObjectSet) && (scalarType(type) || type=="paramarray"))
            || (instruction.type==Statement::Type::For && (type.ends_with("()") || type=="boolean" || type=="string" || type=="single" || type=="currency" || type=="date" || type=="byte" || objectType(type)))) {
            VbaExecutionResult failure{VbaExecutionStatus::Unsupported,"Declared variable type is unsupported for this loop/object target",instruction.line}; failure.errorNumber=3251; return failure;
        }
    }
    const auto assignVariable=[&](const std::string& name,const VbaValue& value,bool arrayElement=false) {
        if(!arrayElement) for(const auto& loop:loops) if(!loop.arrayRoot.empty() && loop.arrayRoot==name)
            throw VbaRuntimeError(3251,"Array source replacement during For Each is not implemented");
        if(constantNames.contains(name)) throw VbaRuntimeError(3251,"Assignment to a constant is not permitted");
        if(!arrayElement && fixedArrays.contains(name)) throw VbaRuntimeError(3251,"Whole fixed-array assignment is not implemented");
        const auto type=variableType(name);
        if(type=="paramarray") throw VbaRuntimeError(3251,"ParamArray slot replacement is not implemented");
        locals[name]=checkedType(type) ? coerceVbaModuleValue(value,type) : value;
    };
    const auto equalValue = [](const VbaValue& a, const VbaValue& b) {
        return a.kind == b.kind && a.numericType==b.numericType && a.number == b.number && a.currencyUnits==b.currencyUnits && a.text == b.text
            && a.boolean == b.boolean && a.object == b.object && a.bytes == b.bytes && a.errorNumber==b.errorNumber;
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
            for (std::size_t i = 0; i < parameters.size() && i<arguments.size(); ++i)
                if (!omittedParameters[i] && parameters[i].byReference) arguments[i] = locals[lower(parameters[i].name)];
    };
    auto completed = [&](int line, bool exited) {
        VbaExecutionResult outcome{VbaExecutionStatus::Completed, {}, line, exited};
        if (isFunction) {
            const auto value = locals[lower(std::string(procedure))];
            outcome.typedReturnValue = std::make_shared<VbaValue>(value);
            if (value.kind != VbaValue::Kind::Object && value.kind!=VbaValue::Kind::Error && value.kind!=VbaValue::Kind::Missing && !isVbaArray(value)) outcome.returnValue = value.asField();
        }
        storeAliases();
        copyArguments();
        if (hasCancel) outcome.cancelled = locals["cancel"].truthy();
        return outcome;
    };
    const auto invoke = [&](std::string_view name, std::vector<VbaValue>& actuals,
                            const std::vector<std::string>& names, bool needsValue) -> VbaValue {
        std::vector<std::string> argumentNames,argumentRefs;
        for(const auto& reference:names) {
            if(reference.starts_with("@named:")) {
                const auto separator=reference.find(':',7);
                if(separator==std::string::npos) throw VbaRuntimeError(450,"Invalid named argument binding");
                argumentNames.push_back(reference.substr(7,separator-7)); argumentRefs.push_back(reference.substr(separator+1));
            } else { argumentNames.emplace_back(); argumentRefs.push_back(reference); }
        }
        for(auto& expression:argumentRefs) if(const auto target=memberTarget(expression);target && target->arguments && trim(*target->arguments).empty()) {
            const auto variable=locals.find(lower(target->name));
            if(variable!=locals.end() && isVbaArray(variable->second)) expression=target->name;
        }
        std::vector<ReferenceBinding> bindings;
        for (const auto& expression : argumentRefs) {
            auto found = locals.find(lower(expression));
            const auto type=variableType(lower(expression));
            ReferenceBinding binding{found==locals.end() ? nullptr : &found->second,checkedType(type) ? type : std::string{}};
            for (const auto& alias:aliases) if (alias.name==lower(expression)) binding={alias.source,alias.type};
            binding.paramArray=type=="paramarray";
            binding.fixedArray=fixedArrays.contains(lower(expression));
            if(constantNames.contains(lower(expression))) binding.source=nullptr;
            if(expression=="@array_element") binding.indexedArray=true;
            if(const auto target=memberTarget(expression); target && target->arguments) {
                const auto array=locals.find(lower(target->name));
                if(array!=locals.end() && isVbaArray(array->second)) binding.indexedArray=true;
            }
            bindings.push_back(binding);
        }
        VbaExecutionResult nested;
        const bool named=std::any_of(argumentNames.begin(),argumentNames.end(),[](const std::string& value){ return !value.empty(); });
        if(named) {
            std::string targetModule(module);
            std::vector<Statement> compiled; std::vector<Parameter> parameters;
            bool function=false,cancel=false;
            auto declaration=parseProcedure(targetModule,name,compiled,function,cancel,&parameters);
            if(declaration.status==VbaExecutionStatus::ProcedureMissing) {
                const auto external=host.moduleForProcedure(name);
                if(external) {
                    targetModule=*external; compiled.clear(); parameters.clear();
                    declaration=parseProcedure(targetModule,name,compiled,function,cancel,&parameters);
                }
            }
            if(declaration.status!=VbaExecutionStatus::Completed) throw VbaRuntimeError(3251,"Named arguments require a supported VBA procedure declaration");
            if(std::any_of(parameters.begin(),parameters.end(),[](const Parameter& parameter){ return parameter.paramArray; }))
                throw VbaRuntimeError(3251,"Named ParamArray calls are not implemented");
            std::vector<VbaValue> ordered(parameters.size(),VbaValue::missing());
            std::vector<ReferenceBinding> orderedBindings(parameters.size());
            std::vector<bool> assigned(parameters.size(),false);
            std::vector<std::size_t> mapping; bool namedSeen=false;
            for(std::size_t index=0;index<actuals.size();++index) {
                std::size_t target=index;
                if(!argumentNames[index].empty()) {
                    namedSeen=true; target=parameters.size();
                    for(std::size_t parameter=0;parameter<parameters.size();++parameter)
                        if(lower(parameters[parameter].name)==lower(argumentNames[index])) { target=parameter; break; }
                    if(target==parameters.size()) throw VbaRuntimeError(448,"Named argument not found");
                } else if(namedSeen) throw VbaRuntimeError(450,"Positional arguments cannot follow named arguments");
                if(target>=parameters.size() || assigned[target]) throw VbaRuntimeError(450,"Argument supplied more than once or too many arguments");
                assigned[target]=true; ordered[target]=actuals[index]; orderedBindings[target]=bindings[index]; mapping.push_back(target);
            }
            nested=executeProcedure(targetModule,name,host,ordered,&orderedBindings);
            for(std::size_t index=0;index<actuals.size();++index) actuals[index]=ordered[mapping[index]];
        } else {
            nested=executeProcedure(module,name,host,actuals,&bindings);
            if(nested.status==VbaExecutionStatus::ProcedureMissing) {
                const auto external=host.moduleForProcedure(name);
                if(external) nested=executeProcedure(*external,name,host,actuals,&bindings);
            }
        }
        if(nested.status==VbaExecutionStatus::ProcedureMissing && lower(std::string(name))=="doevents") {
            if(!actuals.empty()) throw VbaRuntimeError(450,"DoEvents takes no arguments");
            const auto result=VbaValue::integer(host.doEvents());
            if(vbaExecutionStopRequested()) throw VbaExecutionStopped{};
            return result;
        }
        if(vbaExecutionStopRequested()) throw VbaExecutionStopped{};
        for (std::size_t i = 0; i < names.size(); ++i) {
            auto found = locals.find(lower(argumentRefs[i]));
            if (found != locals.end()) {
                if(constantNames.contains(found->first)) continue;
                if(variableType(found->first)=="paramarray") {
                    if(actuals[i].kind!=VbaValue::Kind::Array || actuals[i].object!=found->second.object)
                        throw VbaRuntimeError(3251,"ParamArray forwarding cannot replace argument slots");
                    continue;
                }
                bool arrayUpdate=false;
                if(fixedArrays.contains(found->first)) {
                    if(!isVbaArray(actuals[i]) || vbaArrayBounds(actuals[i])!=vbaArrayBounds(found->second)) throw VbaRuntimeError(3251,"Fixed array argument bounds changed");
                    arrayUpdate=true;
                }
                for(const auto& loop:loops) if(loop.arrayRoot==found->first && sameVbaArrayAllocation(loop.array,actuals[i])) arrayUpdate=true;
                assignVariable(found->first,actuals[i],arrayUpdate);
            }
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
    auto compareMode=VbaCompareMode::Binary;
    {
        std::istringstream source{std::string(module)}; std::string line;
        const std::regex option(R"(^\s*Option\s+Compare\s+(Binary|Text|Database)\s*(?:'.*)?$)",std::regex::icase);
        while (std::getline(source,line)) {
            std::smatch match;
            if (std::regex_match(line,match,option)) {
                const auto mode=lower(match[1].str());
                compareMode=mode=="database" ? VbaCompareMode::Database : mode=="text" ? VbaCompareMode::Text : VbaCompareMode::Binary;
            }
        }
    }
    const auto evaluate = [&](std::string_view expression) {
        return evaluateVbaExpression(expression, host, locals, resolver,withFrames.empty() ? nullptr : &withFrames.back().value,compareMode,arrayBase);
    };
    const auto evaluateObject = [&](std::string_view expression) {
        return evaluateVbaObjectExpression(expression,host,locals,resolver,withFrames.empty() ? nullptr : &withFrames.back().value,compareMode,arrayBase);
    };
    bool active = true;
    int currentLine = 0;
    int steps = 0;
    int effects = 0;
    try {
        for (std::size_t ip = 0; ip < instructions.size(); ++ip) {
            if(vbaExecutionStopRequested()) throw VbaExecutionStopped{};
            const auto& instruction = instructions[ip];
            currentLine = instruction.line;
            if (++steps > 10000)
                return {VbaExecutionStatus::Error, "VBA-Ausführungsbudget (10.000 Schritte) erreicht", currentLine};
            refreshAliases();
            try {
            switch (instruction.type) {
            case Statement::Type::With:
                if (!withFrames.empty() && withFrames.back().instruction==ip && withFrames.back().pending) withFrames.pop_back();
                if (!active) { withFrames.push_back({ip,VbaValue::reference(nullptr),false}); break; }
                try {
                    const auto value=evaluateObject(instruction.expression);
                    if (value.kind!=VbaValue::Kind::Object) throw VbaRuntimeError(424,"With requires an object");
                    withFrames.push_back({ip,value,false});
                } catch (...) { withFrames.push_back({ip,VbaValue::reference(nullptr),true}); throw; }
                break;
            case Statement::Type::EndWith:
                if (withFrames.empty()) throw VbaRuntimeError(5,"With stack invalid");
                withFrames.pop_back(); break;
            case Statement::Type::SendKeys:
                if (active) {
                    const auto expressions=splitArguments(instruction.expression);
                    const auto keys=scalarVbaValue(evaluate(expressions[0]));
                    if (keys.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null SendKeys input");
                    const bool wait=expressions.size()>1 && !expressions[1].empty() ? coerceVbaModuleValue(scalarVbaValue(evaluate(expressions[1])),"boolean").boolean : false;
                    host.sendKeys(keys.asText(),wait);
                }
                break;
            case Statement::Type::ErrorResumeNext:
                if (active) { resumeNext = true; handler = -1; }
                break;
            case Statement::Type::ErrorDisable:
                if (active) { resumeNext = false; handler = -1; }
                break;
            case Statement::Type::Label:
                break;
            case Statement::Type::Jump:
                if (active) { frames.clear(); loops.clear(); withFrames.clear(); active = true; ip = static_cast<std::size_t>(instruction.partner); }
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
                        frames.clear(); loops.clear(); withFrames.clear(); active = true;
                        ip = static_cast<std::size_t>(instruction.partner);
                    } else {
                        frames = faultFrames; loops = faultLoops; withFrames=faultWithFrames; active = faultActive;
                        ip = instruction.type == Statement::Type::Resume ? faultInstruction - 1 : faultInstruction;
                    }
                }
                break;
            case Statement::Type::Select: {
                // Preserve a failed selector frame for Resume, and skip its
                // cases under Resume Next rather than consuming an outer frame.
                if (!frames.empty() && frames.back().select && frames.back().pending && frames.back().instruction==ip) {
                    active=frames.back().parent; frames.pop_back();
                }
                frames.push_back({active,false,false,{},true,true,ip});
                active=false;
                if (frames.back().parent) frames.back().selector=scalarVbaValue(evaluate(instruction.expression));
                frames.back().pending=false;
                break;
            }
            case Statement::Type::Case: {
                if (frames.empty() || !frames.back().select) throw VbaRuntimeError(5,"Select stack invalid");
                auto& frame=frames.back(); active=false;
                if (!frame.parent || frame.condition || frame.pending) break;
                bool matched=lower(instruction.expression)=="else";
                for (const auto& clause:matched ? std::vector<std::string>{} : splitArguments(instruction.expression)) {
                    const auto relation=[&](std::string_view op,const std::string& expression) {
                        return compareVbaValues(op,frame.selector,scalarVbaValue(evaluate(expression)),host,compareMode).truthy();
                    };
                    if (outsideKeyword(clause,"is")==0 || (!clause.empty() && std::string("<>=").find(clause.front())!=std::string::npos)) {
                        const auto comparison=trim(outsideKeyword(clause,"is")==0 ? clause.substr(2) : clause);
                        const auto end=comparison.find_first_not_of("<>=");
                        matched=relation(comparison.substr(0,end),trim(comparison.substr(end)));
                    } else if (const auto range=outsideKeyword(clause,"to"); range!=std::string::npos) {
                        // Case lists and range bounds stop once a match or a
                        // failed lower bound determines the result.
                        matched=relation(">=",trim(clause.substr(0,range)))
                            && relation("<=",trim(clause.substr(range+2)));
                    } else matched=relation("=",clause);
                    if (matched) break;
                }
                frame.condition=matched; active=matched;
                break;
            }
            case Statement::Type::EndSelect:
                if (frames.empty() || !frames.back().select) throw VbaRuntimeError(5,"End Select stack invalid");
                active=frames.back().parent; frames.pop_back(); break;
            case Statement::Type::If: {
                const bool matched = active
                    ? evaluate(instruction.argument).truthy() : false;
                frames.push_back({active, matched, false});
                active = active && matched;
                break;
            }
            case Statement::Type::ElseIf: {
                if (frames.empty() || frames.back().select || frames.back().inElse)
                    throw VbaRuntimeError(5,"ElseIf stack invalid");
                auto& frame=frames.back(); active=false;
                if (frame.parent && !frame.condition) {
                    const bool matched=evaluate(instruction.expression).truthy();
                    frame.condition=matched; active=matched;
                }
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
            case Statement::Type::Evaluate:
                if (active) {
                    if (++effects > 1000) throw std::runtime_error("VBA-Effektbudget erreicht");
                    (void)evaluate(instruction.expression);
                }
                break;
            case Statement::Type::Quit:
                if (active) { host.quitApplication(); if(vbaExecutionStopRequested()) throw VbaExecutionStopped{}; return completed(instruction.line,true); }
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
                    const auto statement=scalarVbaValue(evaluate(instruction.expression));
                    if(statement.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null RunSQL statement");
                    host.executeSql(statement.asText());
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
                if (active) host.requeryControl(instruction.endExpression=="." ? host.memberControlName(instruction.argument) : instruction.argument);
                break;
            case Statement::Type::ControlFocus:
                if (active) host.focusControl(instruction.endExpression=="." ? host.memberControlName(instruction.argument) : instruction.argument);
                break;
            case Statement::Type::ControlVisible:
                if (active) {
                    const auto value=scalarVbaValue(evaluate(instruction.expression));
                    if (value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null visibility");
                    host.setControlVisible(instruction.endExpression=="." ? host.memberControlName(instruction.argument) : instruction.argument,value.truthy());
                }
                break;
            case Statement::Type::Call:
                if (active) {
                    const auto expressions = splitArguments(instruction.expression);
                    std::vector<VbaValue> actuals;
                    std::vector<std::string> references;
                    for(const auto& expression:expressions) {
                        const auto parsed=callArgument(expression);
                        actuals.push_back(parsed.expression.empty() ? VbaValue::missing() : evaluate(parsed.expression));
                        references.push_back(parsed.name.empty() ? parsed.expression : "@named:"+parsed.name+":"+parsed.expression);
                    }
                    (void)invoke(instruction.argument,actuals,references,false);
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
                    if(variableType(lower(instruction.argument)).ends_with("()")) throw VbaRuntimeError(13,"Set cannot replace an array variable");
                    auto value = evaluateObject(instruction.expression);
                    if (value.kind != VbaValue::Kind::Object) throw VbaRuntimeError(13,"Set benötigt ein Objekt");
                    assignVariable(lower(instruction.argument),value);
                }
                break;
            case Statement::Type::ObjectAssign:
                if (active) {
                    const auto pos = memberSeparator(instruction.argument);
                    const auto target=memberTarget(trim(instruction.argument.substr(pos+1)));
                    if(!target) throw VbaRuntimeError(5,"Invalid property assignment target");
                    const auto& name=target->name;
                    const bool indexed=target->arguments.has_value();
                    const auto object = pos==0 && !withFrames.empty() ? withFrames.back().value : evaluateObject(instruction.argument.substr(0, pos));
                    if (!object.object) throw VbaRuntimeError(91,"Objektvariable nicht gesetzt");
                    std::vector<VbaValue> indices;
                    if(indexed) {
                        const auto inner=trim(*target->arguments);
                        if(!inner.empty()) for(const auto& expression:splitArguments(inner)) indices.push_back(evaluate(expression));
                    }
                    const auto value = evaluate(instruction.expression);
                    if(indexed && value.kind==VbaValue::Kind::Object) throw VbaRuntimeError(3251,"Indexed object property writes require supported Set syntax");
                    if(indexed) object.object->assignIndexedMember(name,indices,value);
                    else if (instruction.argument[pos] == '!') object.object->assignItem(name,value);
                    else object.object->assign(name,value);
                }
                break;
            case Statement::Type::ObjectCall:
                if (active) (void)evaluate(instruction.expression);
                break;
            case Statement::Type::While:
                if (!active || (!instruction.expression.empty()
                    && evaluate(instruction.expression).truthy()==(instruction.argument=="until")))
                    ip=static_cast<std::size_t>(instruction.partner);
                break;
            case Statement::Type::EndWhile:
                if (active && (instruction.expression.empty()
                    || evaluate(instruction.expression).truthy()!=(instruction.argument=="until")))
                    ip=static_cast<std::size_t>(instruction.partner)-1;
                break;
            case Statement::Type::ExitDo:
                if (active) {
                    if (frames.size()<static_cast<std::size_t>(instruction.conditionalUnwind)
                        || withFrames.size()<static_cast<std::size_t>(instruction.withUnwind)
                        || loops.size()<static_cast<std::size_t>(instruction.forUnwind))
                        throw VbaRuntimeError(5,"Exit Do stack invalid");
                    frames.resize(frames.size()-instruction.conditionalUnwind);
                    withFrames.resize(withFrames.size()-instruction.withUnwind);
                    loops.resize(loops.size()-instruction.forUnwind);
                    active=true; ip=static_cast<std::size_t>(instruction.partner);
                }
                break;
            case Statement::Type::ResizeArray:
                if(active) {
                    const auto name=lower(instruction.argument),type=variableType(name);
                    if(fixedArrays.contains(name)) throw VbaRuntimeError(3251,"ReDim of a fixed array is not implemented");
                    if(type=="paramarray") throw VbaRuntimeError(3251,"ParamArray resize is not implemented");
                    if(!instruction.endExpression.empty() && type.empty() && !locals.contains(name))
                        throw VbaRuntimeError(3251,"Typed ReDim requires an existing Variant variable");
                    const bool declaredArray=type.ends_with("()");
                    const auto declaredElement=declaredArray ? type.substr(0,type.size()-2) : std::string{};
                    if(!type.empty() && type!="variant" && !declaredArray) throw VbaRuntimeError(13,"ReDim requires an array or Variant variable");
                    if(declaredArray && !instruction.endExpression.empty() && instruction.endExpression!=declaredElement)
                        throw VbaRuntimeError(13,"ReDim cannot change a declared array element type");
                    std::vector<VbaArrayDimension> bounds;
                    for(const auto& dimension:splitArguments(instruction.expression)) {
                        const auto range=outsideKeyword(dimension,"to");
                        const auto lowerBound=range==std::string::npos ? arrayBase : static_cast<int>(coerceVbaModuleValue(scalarVbaValue(evaluate(dimension.substr(0,range))),"long").number);
                        const auto upperBound=static_cast<int>(coerceVbaModuleValue(scalarVbaValue(evaluate(range==std::string::npos ? dimension : dimension.substr(range+2))),"long").number);
                        bounds.push_back({lowerBound,upperBound});
                    }
                    refreshAliases();
                    const auto found=locals.find(name);
                    const auto original=found==locals.end() ? VbaValue::empty() : found->second;
                    assignVariable(name,resizeVbaArrayWithBounds(original,std::move(bounds),instruction.stepExpression=="preserve",declaredArray ? declaredElement : instruction.endExpression));
                }
                break;
            case Statement::Type::EraseArray:
                if(active) {
                    const auto name=lower(instruction.argument);
                    const auto found=locals.find(name);
                    if(found==locals.end() || !isVbaArray(found->second)) throw VbaRuntimeError(13,"Erase requires an array");
                    const auto type=variableType(name);
                    if(fixedArrays.contains(name)) {
                        for(const auto& loop:loops) if(loop.arrayRoot==name) throw VbaRuntimeError(3251,"Fixed array Erase during For Each is not implemented");
                        assignVariable(name,resizeVbaArrayWithBounds(found->second,vbaArrayBounds(found->second),false,type.substr(0,type.size()-2)),true);
                    } else assignVariable(name,type.ends_with("()") ? makeUnallocatedVbaArray(type.substr(0,type.size()-2)) : VbaValue::empty());
                }
                break;
            case Statement::Type::IndexedAssign:
                if(active) {
                    const auto name=lower(instruction.argument);
                    const auto initial=locals.find(name);
                    const bool objectTarget=initial!=locals.end() && initial->second.kind==VbaValue::Kind::Object;
                    const auto target=objectTarget ? initial->second.object : std::shared_ptr<VbaObject>{};
                    if(objectTarget && !target) throw VbaRuntimeError(91,"Object variable is not set");
                    std::vector<VbaValue> indices;
                    for(const auto& expression:splitArguments(instruction.endExpression)) {
                        const auto value=evaluate(expression);
                        indices.push_back(objectTarget ? value : coerceVbaModuleValue(scalarVbaValue(value),"long"));
                    }
                    const bool object=instruction.stepExpression=="object";
                    const auto value=object ? evaluateObject(instruction.expression) : evaluate(instruction.expression);
                    if(object && value.kind!=VbaValue::Kind::Object) throw VbaRuntimeError(13,"Set requires an object reference");
                    refreshAliases();
                    if(objectTarget) {
                        if(!object && value.kind==VbaValue::Kind::Object) throw VbaRuntimeError(3251,"Object default-member writes require Set or a scalar default");
                        target->assignDefaultMember(indices,value); break;
                    }
                    const auto found=locals.find(name);
                    if(found==locals.end()) throw VbaRuntimeError(13,"Array variable required");
                    if(object && (!isVbaArray(found->second) || (vbaArrayType(found->second)!=8204 && vbaArrayType(found->second)!=8201))) throw VbaRuntimeError(13,"Object references require a Variant or Object array");
                    if(!object && isVbaArray(found->second) && vbaArrayType(found->second)==8201) throw VbaRuntimeError(3251,"Object array default-member assignment is not implemented");
                    assignVariable(name,withVbaArrayElement(found->second,indices,value),true);
                }
                break;
            case Statement::Type::SetVar:
                if (active) {
                    const auto name=lower(instruction.argument);
                    if(objectType(variableType(name))) {
                        const auto binding=locals.find(name);
                        if(binding==locals.end() || binding->second.kind!=VbaValue::Kind::Object) throw VbaRuntimeError(13,"Object binding required");
                        const auto target=binding->second.object;
                        if(!target) throw VbaRuntimeError(91,"Object variable is not set");
                        if(!target->hasScalarDefault()) throw VbaRuntimeError(3251,"Object scalar default assignment is not implemented for this class");
                        target->assignDefaultMember({},scalarVbaValue(evaluate(instruction.expression)));
                    } else {
                        const auto value=evaluate(instruction.expression);
                        if (locals.contains(name) || !host.assignIdentifier(instruction.argument,value)) assignVariable(name,value);
                    }
                }
                break;
            case Statement::Type::ForEach: {
                if (!active) { ip = static_cast<std::size_t>(instruction.partner); break; }
                const auto collection = evaluate(instruction.expression);
                if(isVbaArray(collection)) {
                    if(vbaArrayRank(collection)==0) throw VbaRuntimeError(9,"For Each requires an allocated array");
                    if(collection.kind!=VbaValue::Kind::Array || vbaArrayType(collection)==8209) throw VbaRuntimeError(3251,"For Each Byte arrays is not implemented");
                    const auto name=lower(instruction.argument),type=variableType(name);
                    if(!type.empty() && type!="variant") throw VbaRuntimeError(13,"Array For Each control must be Variant");
                    LoopFrame loop{static_cast<int>(ip)};
                    loop.array=collection;
                    const auto source=lower(trim(instruction.expression));
                    if(const auto variable=locals.find(source); variable!=locals.end()) {
                        if(source==name) throw VbaRuntimeError(3251,"Array source and For Each control must be distinct");
                        loop.arrayRoot=source;
                    } else if(!isDetachedVbaArray(collection) && !std::regex_search(source,std::regex(R"(^(array|vba\s*\.\s*array|split|filter)\s*\()"))) {
                        throw VbaRuntimeError(3251,"Array For Each requires a named variable or Array/VBA.Array/Split/Filter result");
                    }
                    loop.frameDepth=frames.size(); loop.withDepth=withFrames.size();
                    if(vbaArraySize(collection)==0) ip=static_cast<std::size_t>(instruction.partner);
                    else {
                        locals[name]=vbaArrayLinearElement(collection,0);
                        loops.push_back(std::move(loop));
                    }
                    break;
                }
                if (collection.kind != VbaValue::Kind::Object)
                    throw VbaRuntimeError(424,"For Each requires an object collection or supported array");
                if (!collection.object) throw VbaRuntimeError(91,"For Each object is not set");
                LoopFrame loop{static_cast<int>(ip)};
                loop.enumerationSource=collection.object;
                loop.enumerationLease=collection.object->enumerationGuard();
                loop.items = collection.object->enumerate();
                loop.frameDepth = frames.size(); loop.withDepth=withFrames.size();
                if (loop.items.empty()) ip = static_cast<std::size_t>(instruction.partner);
                else {
                    assignVariable(lower(instruction.argument),loop.items.front());
                    loops.push_back(std::move(loop));
                }
                break;
            }
            case Statement::Type::ExitFor:
                if (active) {
                    if (loops.empty()) throw VbaRuntimeError(5,"Exit For stack invalid");
                    const auto loop = std::move(loops.back()); loops.pop_back();
                    frames.resize(loop.frameDepth); withFrames.resize(loop.withDepth); active = true;
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
                const auto counter=lower(instruction.argument),type=variableType(counter);
                if(integralType(type)) {
                    if(std::floor(initial)!=initial || std::floor(end)!=end || std::floor(step)!=step)
                        throw VbaRuntimeError(3251,"Fractional typed For limits and steps are not implemented");
                    (void)coerceVbaModuleValue(VbaValue::numeric(end),type);
                    (void)coerceVbaModuleValue(VbaValue::numeric(step),type);
                }
                assignVariable(counter,VbaValue::numeric(initial));
                if ((step > 0 && initial > end) || (step < 0 && initial < end)) {
                    ip = static_cast<std::size_t>(instruction.partner);
                } else {
                    LoopFrame loop{static_cast<int>(ip),end,step};
                    loop.frameDepth = frames.size(); loop.withDepth=withFrames.size();
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
                    if(loop.array.kind==VbaValue::Kind::Array) {
                        const auto found=locals.find(loop.arrayRoot);
                        const auto& current=loop.arrayRoot.empty() ? loop.array : found==locals.end() ? VbaValue{} : found->second;
                        if(!sameVbaArrayAllocation(loop.array,current)) throw VbaRuntimeError(3251,"Array replacement or resizing during For Each is not implemented");
                        if(++loop.itemIndex<vbaArraySize(loop.array)) {
                            locals[name]=vbaArrayLinearElement(current,loop.itemIndex);
                            ip=static_cast<std::size_t>(loop.forIndex);
                        } else loops.pop_back();
                        break;
                    }
                    if (++loop.itemIndex < loop.items.size()) {
                        assignVariable(name,loop.items[loop.itemIndex]);
                        ip = static_cast<std::size_t>(loop.forIndex);
                    } else loops.pop_back();
                    break;
                }
                const auto nextValue = locals[name].asNumber() + loops.back().step;
                if (!std::isfinite(nextValue)) throw std::runtime_error("For-Zählerüberlauf");
                assignVariable(name,VbaValue::numeric(nextValue));
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
            if(vbaExecutionStopRequested()) throw VbaExecutionStopped{};
            storeAliases();
            } catch (const VbaExecutionStopped&) { throw;
            } catch (const std::exception& exception) {
                if(vbaExecutionStopRequested()) throw VbaExecutionStopped{};
                storeAliases();
                if (!resumeNext && (handler < 0 || handlingError)) throw;
                const auto* typed = dynamic_cast<const VbaRuntimeError*>(&exception);
                errorState->number = typed ? typed->number : 5;
                errorState->description = exception.what();
                if (!resumeNext) {
                    handlingError = true; faultInstruction = ip;
                    faultFrames = frames; faultLoops = loops; faultWithFrames=withFrames; faultActive = active;
                    frames.clear(); loops.clear(); withFrames.clear(); active = true;
                    ip = static_cast<std::size_t>(handler);
                }
            }
        }
    } catch (const VbaExecutionStopped&) {
        return {VbaExecutionStatus::Completed,{},currentLine,true};
    } catch (const std::exception& exc) {
        if(vbaExecutionStopRequested()) return {VbaExecutionStatus::Completed,{},currentLine,true};
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
            const auto epoch=vbaExecutionEpoch();
            auto outcome=executeVbaProcedureWithArguments(module,name,host,args);
            if (outcome.status==VbaExecutionStatus::ProcedureMissing) {
                if (const auto external=host.moduleForProcedure(name)) outcome=executeVbaProcedureWithArguments(*external,name,host,args);
            }
            if(vbaExecutionStopRequested() || epoch!=vbaExecutionEpoch()) throw VbaExecutionStopped{};
            line=outcome.line;
            if (outcome.status!=VbaExecutionStatus::Completed || !outcome.typedReturnValue)
                throw VbaRuntimeError(outcome.errorNumber ? outcome.errorNumber : 5,outcome.explanation.empty() ? "Event function execution failed" : outcome.explanation);
            return *outcome.typedReturnValue;
        };
        const auto value=evaluateVbaExpression(source,host,{},resolver);
        VbaExecutionResult result{VbaExecutionStatus::Completed,{},line};
        result.typedReturnValue=std::make_shared<VbaValue>(value);
        if (value.kind!=VbaValue::Kind::Object && value.kind!=VbaValue::Kind::Error && value.kind!=VbaValue::Kind::Missing && !isVbaArray(value)) result.returnValue=value.asField();
        return result;
    } catch (const VbaExecutionStopped&) {
        return {VbaExecutionStatus::Completed,{},line,true};
    } catch (const VbaRuntimeError& error) {
        VbaExecutionResult result{VbaExecutionStatus::Error,error.what(),line}; result.errorNumber=error.number; return result;
    } catch (const std::exception&) {
        VbaExecutionResult result{VbaExecutionStatus::Error,"Event expression execution failed",line}; result.errorNumber=5; return result;
    }
}
}  // namespace linaccess
