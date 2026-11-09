#include "native_object_types.hpp"
#include "native_file_operations.hpp"
#include "vba_array.hpp"
#include "vba_runtime.hpp"
#include "vba_debugger.hpp"
#include "vba_expression.hpp"
#include "vba_decimal.hpp"
#include "vba_currency.hpp"
#include "vba_calendar.hpp"

#include <algorithm>
#include <deque>
#include <filesystem>
#include <cerrno>
#include <unistd.h>
#if defined(__linux__)
#include <sys/syscall.h>
#include <fcntl.h>
#include <linux/fs.h>
#endif
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
thread_local std::vector<const VbaValue*> activeArrayElementRoots;
struct ArrayElementRootScope {
    std::size_t base=activeArrayElementRoots.size();
    ~ArrayElementRootScope() { activeArrayElementRoots.resize(base); }
};
struct AliasRefreshScope;
thread_local std::vector<AliasRefreshScope*> activeAliasFrames;
struct AliasRefreshScope {
    std::function<void()> refresh;
    explicit AliasRefreshScope(std::function<void()> callback) : refresh(std::move(callback)) { activeAliasFrames.push_back(this); }
    ~AliasRefreshScope() { activeAliasFrames.pop_back(); }
};
// This interpreter deliberately reports unsupported VBA instead of silently
// treating Access/DAO/COM calls as successful no-ops.
struct ArrayResizeTarget { std::string name,bounds,elementType; };
struct BranchDestination {
    std::string label;
    int partner{-1};
    int conditionalUnwind{},withUnwind{},forUnwind{};
};
struct Statement {
    enum class Type { With, EndWith, Select, Case, EndSelect, SendKeys, Directory, FileCopy, FileRename, FileKill, If, ElseIf, Else, EndIf, For, ForEach, ExitFor, Next, Evaluate, Save, Maximize, GoLast,
                      GoNext, GoNew, Requery,
                      ExecuteSql, OpenForm, ControlRequery, ControlFocus, ControlVisible,
                      Call, SetField, IndexedAssign, ResizeArray, EraseArray, SetVar, StringAlign, Exit, Dim, ObjectSet, ObjectAssign, ObjectCall, While, EndWhile, ExitDo, Quit, RaiseError, ErrorResumeNext, ErrorDisable, Label, Jump, GoSub, ComputedJump, ComputedGoSub, ReturnSubroutine, ErrorHandler, Resume, ResumeNext, ResumeLabel } type;
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
    std::vector<ArrayResizeTarget> resizeTargets;
    int sourceStatementEnd{-1};
    std::vector<BranchDestination> destinations;
};

bool debugExecutable(const Statement& statement) {
    switch(statement.type) {
    case Statement::Type::Dim:case Statement::Type::Label:case Statement::Type::EndIf:case Statement::Type::EndSelect:case Statement::Type::EndWith:case Statement::Type::Else:return false;
    default:return true;
    }
}

std::string trim(std::string text) {
    const auto first = text.find_first_not_of(" \t\r\n");
    if (first == std::string::npos) return {};
    const auto last = text.find_last_not_of(" \t\r\n");
    return text.substr(first, last - first + 1);
}
std::string nativeDeclaredType(std::string type) {
    if(type=="vba.collection") return "collection";
    const auto access=nativeAccessObjectType(type);
    return access.empty() ? std::move(type) : std::string(access);
}
bool nativeDaoType(std::string_view type) {
    return type.starts_with("dao.") && nativeObjectType(type)==type;
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
std::optional<std::string> canonicalLineLabel(std::string name) {
    name=lower(trim(std::move(name)));
    if(!name.empty() && std::all_of(name.begin(),name.end(),[](unsigned char ch) { return std::isdigit(ch); })) {
        const auto first=name.find_first_not_of('0');
        name=first==std::string::npos ? "0" : name.substr(first);
        if(name.size()>5 || (name.size()==5 && name>"65529")) return {};
    }
    return name;
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
    int number{}, helpContext{};
    std::string_view className() const override { return "ErrObject"; }
    bool isNativeType(std::string_view type) const override { return type=="vba.errobject"; }
    std::string description, source, helpFile;
    void clear() { number=0; helpContext=0; description.clear(); source.clear(); helpFile.clear(); }
    bool hasScalarDefault() const override { return true; }
    VbaValue scalarValue() override { return VbaValue::longInteger(number); }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        if(lower(std::string(name))=="raise") return VbaArgumentSignature{{"Number","Source","Description","HelpFile","HelpContext"},1};
        return {};
    }
    void assign(std::string_view name,const VbaValue& value) override {
        const auto key=lower(std::string(name));
        if(key=="number" || key=="helpcontext") {
            const auto converted=coerceVbaModuleValue(scalarVbaValue(value),"long");
            (key=="number" ? number : helpContext)=static_cast<int>(converted.asNumber());
        } else if(key=="description" || key=="source" || key=="helpfile") {
            const auto converted=coerceVbaModuleValue(scalarVbaValue(value),"string").asText();
            (key=="description" ? description : key=="source" ? source : helpFile)=converted;
        } else throw VbaRuntimeError(438,"Err property assignment unsupported");
    }
    VbaValue member(std::string_view name, const std::vector<VbaValue>& args) override {
        const auto key = lower(std::string(name));
        if (key == "raise") {
            if(args.empty() || args.size()>5 || args[0].kind==VbaValue::Kind::Missing)
                throw VbaRuntimeError(450,"Err.Raise requires Number and up to four optional arguments");
            const auto code=static_cast<int>(coerceVbaModuleValue(scalarVbaValue(args[0]),"long").asNumber());
            if(code==0) throw VbaRuntimeError(5,"Invalid error number");
            const auto textArgument=[&](std::size_t index,const std::string& fallback) {
                return index<args.size() && args[index].kind!=VbaValue::Kind::Missing
                    ? coerceVbaModuleValue(scalarVbaValue(args[index]),"string").asText() : fallback;
            };
            const auto context=args.size()>4 && args[4].kind!=VbaValue::Kind::Missing
                ? static_cast<int>(coerceVbaModuleValue(scalarVbaValue(args[4]),"long").asNumber()) : helpContext;
            throw VbaRuntimeError(code,textArgument(2,description.empty() ? "VBA: explicit error" : description),
                textArgument(1,source),textArgument(3,helpFile),context);
        }
        if (!args.empty()) throw VbaRuntimeError(450, "Unexpected Err arguments");
        if (key == "number") return VbaValue::longInteger(number);
        if (key == "description") return VbaValue::string(description);
        if (key == "source") return VbaValue::string(source);
        if (key == "helpfile") return VbaValue::string(helpFile);
        if (key == "helpcontext") return VbaValue::longInteger(helpContext);
        if (key == "clear") { clear(); return {}; }
        throw VbaRuntimeError(438, "Err member unsupported");
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
bool assignsBareName(const std::string& source,std::string_view name) {
    const auto assignment=topLevelAssignment(source);
    return assignment && lower(assignment->first)==name;
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
std::string inlineStatementLines(const std::string& source,bool preserveLabels=false) {
    std::vector<std::size_t> labelColons; std::size_t statementStart=0;
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
        else if (ch==':' && depth==0 && (i+1>=result.size() || result[i+1]!='=')) {
            if(preserveLabels) {
                static const std::regex identifier(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*|[0-9]+)$)");
                static const std::unordered_set<std::string> keywords{"else","do","loop","wend","next","stop","return","end","exit","error","gosub"};
                const auto segment=trim(result.substr(statementStart,i-statementStart));
                if(std::regex_match(segment,identifier) && !keywords.contains(lower(segment))) labelColons.push_back(i);
            }
            result[i]='\n'; beginning=true; statementStart=i+1;
        }
    }
    for(auto position=labelColons.rbegin();position!=labelColons.rend();++position) result.insert(*position,1,':');
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
            auto name=lower(expression.substr(start,i-start));
            auto dot=i; while(dot<expression.size() && std::isspace(static_cast<unsigned char>(expression[dot]))) ++dot;
            if(dot<expression.size() && expression[dot]=='.') {
                auto member=dot+1; while(member<expression.size() && std::isspace(static_cast<unsigned char>(expression[member]))) ++member;
                auto end=member; while(end<expression.size() && word(static_cast<unsigned char>(expression[end]))) ++end;
                const auto qualified=name+"."+lower(expression.substr(member,end-member));
                if(values.contains(qualified)) { name=qualified; i=end; }
            }
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
struct ModuleConstantDeclaration { ConstantDeclaration value; bool publicValue{}; int line{}; std::string qualifiedName; };
std::vector<ModuleConstantDeclaration> moduleConstantDeclarations(std::string_view module) {
    static const std::regex procedure(R"(^\s*(?:(?:Public|Private|Friend|Static)\s+)?(?:Sub|Function|Property)\s+)",std::regex::icase);
    static const std::regex declaration(R"(^\s*(?:(Public|Private)\s+)?Const\s+(.+)$)",std::regex::icase);
    static const std::regex enumStart(R"(^\s*(?:(Public|Private)\s+)?Enum\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*$)",std::regex::icase);
    static const std::regex enumEnd(R"(^\s*End\s+Enum\s*$)",std::regex::icase);
    static const std::regex enumMember(R"(^\s*([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)(?:\s*=\s*(.+))?\s*$)");
    VbaVariables values; std::unordered_set<std::string> names,enumTypes;
    std::vector<ModuleConstantDeclaration> result; const auto mode=moduleCompareMode(module);
    bool inEnum=false,enumPublic=false,hasMembers=false; std::int64_t nextValue=0; std::string enumName;
    const auto append=[&](ConstantDeclaration value,bool visible,int line,std::string qualified={}) {
        if(!names.insert(value.name).second) throw VbaRuntimeError(5,"Duplicate module constant or Enum member");
        values[value.name]=value.value;
        if(!qualified.empty()) values[qualified]=value.value;
        result.push_back({std::move(value),visible,line,std::move(qualified)});
    };
    std::istringstream source{std::string(module)}; std::string line; int number=0;
    while(std::getline(source,line)) {
        ++number; line=withoutComment(line); if(trim(line).empty()) continue;
        if(std::regex_search(line,procedure)) {
            if(inEnum) throw VbaRuntimeError(3251,"Enum declaration must end before a procedure");
            break;
        }
        std::smatch match;
        if(inEnum) {
            if(std::regex_match(line,enumEnd)) {
                if(!hasMembers) throw VbaRuntimeError(3251,"Empty Enum declaration unsupported");
                inEnum=false; continue;
            }
            if(!std::regex_match(line,match,enumMember)) throw VbaRuntimeError(3251,"Enum member requires a name and optional integer constant");
            std::int64_t value=nextValue;
            if(match[2].matched) {
                const auto folded=constantExpression(trim(match[2].str()),values,mode);
                if(folded.kind!=VbaValue::Kind::Number || std::trunc(folded.number)!=folded.number)
                    throw VbaRuntimeError(3251,"Enum values require integral numeric constants");
                if(folded.number<std::numeric_limits<std::int32_t>::min() || folded.number>std::numeric_limits<std::int32_t>::max())
                    throw VbaRuntimeError(6,"Enum member overflow");
                value=static_cast<std::int64_t>(folded.number);
            }
            if(value<std::numeric_limits<std::int32_t>::min() || value>std::numeric_limits<std::int32_t>::max())
                throw VbaRuntimeError(6,"Enum member overflow");
            append({lower(match[1].str()),"long",VbaValue::longInteger(static_cast<std::int32_t>(value))},enumPublic,number,enumName+"."+lower(match[1].str()));
            nextValue=value+1; hasMembers=true; continue;
        }
        if(std::regex_match(line,match,enumStart)) {
            const auto name=lower(match[2].str());
            if(!enumTypes.insert(name).second) throw VbaRuntimeError(5,"Duplicate Enum type");
            enumName=name; inEnum=true; enumPublic=lower(match[1].str())!="private"; nextValue=0; hasMembers=false; continue;
        }
        if(std::regex_match(line,enumEnd)) throw VbaRuntimeError(3251,"End Enum without Enum");
        if(!std::regex_match(line,match,declaration)) continue;
        const bool visible=lower(match[1].str())=="public";
        for(auto value:constantDeclarations(match[2].str(),values,mode)) append(std::move(value),visible,number);
    }
    if(inEnum) throw VbaRuntimeError(3251,"Enum declaration is not closed");
    return result;
}
VbaVariables moduleConstantValues(std::string_view module) {
    VbaVariables values;
    for(const auto& item:moduleConstantDeclarations(module)) {
        values.emplace(item.value.name,item.value.value);
        if(!item.qualifiedName.empty()) values.emplace(item.qualifiedName,item.value.value);
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
    const std::regex stringAlignment(R"(^(LSet|RSet)\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*=\s*(.+)$)",std::regex::icase);
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
    const std::regex parameter(R"(^(?:(Optional)\s+)?(?:(ByRef|ByVal)\s+)?([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)(?:\s+As\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_.\x80-\xFF]*))?(?:\s*=\s*(.+))?$)", std::regex::icase);
    const std::regex cancelParam(R"(^Cancel\s+As\s+(?:Integer|Boolean)$)", std::regex::icase);
    const std::regex objectSet(R"(^Set\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*=\s*(.+)$)", std::regex::icase);
    const std::regex label(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*|[0-9]+):$)");
    const std::regex goSub(R"(^GoSub\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*|[0-9]+)$)", std::regex::icase);
    const std::regex jump(R"(^GoTo\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*|[0-9]+)$)", std::regex::icase);
    const std::regex errorHandler(R"(^On\s+Error\s+GoTo\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*|[0-9]+)$)", std::regex::icase);
    const std::regex computedBranch(R"(^On\s+(.+)\s+(GoTo|GoSub)\s+(.+)$)",std::regex::icase);
    const std::regex destinationName(R"(^(?:[A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*|[0-9]+)$)");
    const std::regex resumeZero(R"(^Resume\s+0+$)",std::regex::icase);
    const std::regex resumeLabel(R"(^Resume\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*|[0-9]+)$)", std::regex::icase);
    const std::regex whileStart(R"(^(Do\s+(While|Until)|While)\s+(.+)$)", std::regex::icase);
    const std::regex loopEnd(R"(^(Loop|Wend)(?:\s+(While|Until)\s+(.+))?$)",std::regex::icase);
    const std::regex dim(R"(^(?:Dim|Static)\s+(.+)$)", std::regex::icase);
    const std::regex declaration(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)(?:\s+As\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_.\x80-\xFF]*))?$)", std::regex::icase);
    const std::regex forStart(R"(^For\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*=\s*(.+?)\s+To\s+(.+?)(?:\s+Step\s+(.+))?$)", std::regex::icase);
    const std::regex forEach(R"(^For\s+Each\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s+In\s+(.+)$)", std::regex::icase);
    const std::regex next(R"(^Next(?:\s+(.+))?$)", std::regex::icase);

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
        const bool physicalLine=pendingStatements.empty();
        if(physicalLine) ++line;
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
            if(physicalLine && std::isdigit(static_cast<unsigned char>(source.front()))) {
                std::size_t digits=0; while(digits<source.size() && std::isdigit(static_cast<unsigned char>(source[digits]))) ++digits;
                if(digits==source.size() || source[digits]==':' || std::isspace(static_cast<unsigned char>(source[digits]))) {
                    const auto number=canonicalLineLabel(source.substr(0,digits));
                    if(!number) return {VbaExecutionStatus::Unsupported,"Line number exceeds 65529",line};
                    auto following=trim(source.substr(digits));
                    if(following.starts_with(':')) following=trim(following.substr(1));
                    if(!following.empty()) pendingStatements.push_front(following);
                    source=*number+":";
                }
            }
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
                static const std::regex identifier(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*|[0-9]+)$)");
                static const std::unordered_set<std::string> statements{"else","do","loop","wend","next","stop","return","end","exit","error","gosub"};
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
                    const auto type=nativeDeclaredType(lower(declared[4].str()));
                    if(!type.empty() && type!="variant" && type!="byte" && type!="integer" && type!="long" && type!="string"
                        && type!="boolean" && type!="double" && type!="single" && type!="currency" && type!="date" && type!="object"
                        && type!="collection" && nativeScriptingObjectType(type).empty() && !nativeDaoType(type) && nativeAccessObjectType(type).empty())
                        return {VbaExecutionStatus::Unsupported,"Declared parameter type is not implemented",line};
                    const bool objectDefault=(type=="collection" || type=="object" || !nativeScriptingObjectType(type).empty() || nativeDaoType(type) || !nativeAccessObjectType(type).empty()) && lower(defaultExpression)=="nothing";
                    if(optional && (type=="collection" || type=="object" || !nativeScriptingObjectType(type).empty() || nativeDaoType(type) || !nativeAccessObjectType(type).empty()) && !defaultExpression.empty() && !objectDefault)
                        return {VbaExecutionStatus::Unsupported,"Optional native object default must be Nothing",line};
                    if(optional && lower(defaultExpression)=="null" && !type.empty() && type!="variant")
                        return {VbaExecutionStatus::Unsupported,"Null defaults require a Variant parameter",line};
                    if(type.starts_with("access.") && nativeAccessObjectType(type).empty())
                        return {VbaExecutionStatus::Unsupported,"Native Access parameter type is unsupported",line};
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
                const auto declaredResult=nativeDeclaredType(lower(match[4].str()));
                if (isFunction && !declaredResult.empty() && declaredResult!="variant" && declaredResult!="boolean" && declaredResult!="string" && declaredResult!="integer" && declaredResult!="long" && declaredResult!="single" && declaredResult!="double" && declaredResult!="currency" && declaredResult!="date" && declaredResult!="byte" && declaredResult!="object" && declaredResult!="collection" && nativeScriptingObjectType(declaredResult).empty() && !nativeDaoType(declaredResult) && nativeAccessObjectType(declaredResult).empty())
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
            std::vector<int> statementScopes;
            if (resolveLabels) {
                statementScopes.resize(statements.size(),-1);
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
                    statementScopes[i]=lexical.empty() ? -1 : lexical.back();
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
            if (resolveLabels) {
                // Inline wrappers carry label names into the merged real plan.
                labels.clear();
                for(int index=0;index<static_cast<int>(statements.size());++index) {
                    const auto& item=statements[index];
                    if(item.type==Statement::Type::Label && !labels.emplace(lower(item.argument),index).second)
                        return {VbaExecutionStatus::Error,"Duplicate line label",item.line};
                }
                const auto bindDestination=[&](BranchDestination& destination,bool jump,int index,int line) -> std::optional<VbaExecutionResult> {
                    const auto target=labels.find(lower(destination.label));
                    if(target==labels.end()) return VbaExecutionResult{VbaExecutionStatus::Error,"Sprungmarke fehlt",line};
                    destination.partner=target->second;
                    const int destinationScope=statementScopes[target->second];
                    if(!jump) {
                        if(destinationScope>=0) return VbaExecutionResult{VbaExecutionStatus::Unsupported,"GoSub/error/Resume targets inside blocks are not implemented",line};
                        return {};
                    }
                    int scope=statementScopes[index];
                    while(scope>=0 && scope!=destinationScope) {
                        const auto type=statements[scope].type;
                        if(type==Statement::Type::If || type==Statement::Type::Select) ++destination.conditionalUnwind;
                        else if(type==Statement::Type::With) ++destination.withUnwind;
                        else if(type==Statement::Type::For || type==Statement::Type::ForEach) ++destination.forUnwind;
                        scope=statementScopes[scope];
                    }
                    if(scope!=destinationScope) return VbaExecutionResult{VbaExecutionStatus::Unsupported,"GoTo cannot enter an inactive or sibling block",line};
                    return {};
                };
                for(int index=0;index<static_cast<int>(statements.size());++index) {
                    auto& instruction=statements[index];
                    if(instruction.type==Statement::Type::ComputedJump || instruction.type==Statement::Type::ComputedGoSub) {
                        for(auto& destination:instruction.destinations)
                            if(const auto failure=bindDestination(destination,instruction.type==Statement::Type::ComputedJump,index,instruction.line)) return *failure;
                    } else if(instruction.type==Statement::Type::Jump || instruction.type==Statement::Type::GoSub
                        || instruction.type==Statement::Type::ErrorHandler || instruction.type==Statement::Type::ResumeLabel) {
                        BranchDestination destination{instruction.argument};
                        if(const auto failure=bindDestination(destination,instruction.type==Statement::Type::Jump,index,instruction.line)) return *failure;
                        instruction.partner=destination.partner; instruction.conditionalUnwind=destination.conditionalUnwind;
                        instruction.withUnwind=destination.withUnwind; instruction.forUnwind=destination.forUnwind;
                    }
                }
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
                const auto wrapper="Sub __la_inline()\nIf "+condition+" Then\n"+inlineStatementLines(yes,true)
                    +(otherwise==std::string::npos ? "" : "\nElse\n"+inlineStatementLines(no,true))+"\nEnd If\nEnd Sub\n";
                std::vector<Statement> body; bool function=false,cancel=false;
                auto parsed=parseProcedure(wrapper,"__la_inline",body,function,cancel,nullptr,false,nullptr,withActive());
                if (parsed.status!=VbaExecutionStatus::Completed) { parsed.line=line; return parsed; }
                const auto offset=static_cast<int>(statements.size());
                for (auto& statement:body) {
                    if(statement.partner>=0) statement.partner+=offset;
                    if(statement.sourceStatementEnd>=0) statement.sourceStatementEnd+=offset;
                    statement.line=line; statements.push_back(std::move(statement));
                }
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
        } else if (normalized=="error" || normalized.starts_with("error ") || normalized.starts_with("error\t")) {
            const auto argument=trim(source.substr(5));
            if(!validExpression(argument)) return {VbaExecutionStatus::Unsupported,"Error statement requires one expression",line};
            statements.push_back({Statement::Type::RaiseError,{},argument,line});
        } else if ((normalized.starts_with("name ") || normalized.starts_with("name\t"))
                   && !assignsBareName(source,"name") && outsideKeyword(source,"as",4)!=std::string::npos) {
            const auto separator=outsideKeyword(source,"as",4);
            const auto from=trim(source.substr(4,separator-4)),to=trim(source.substr(separator+2));
            if(from.empty() || to.empty() || !validExpression(from) || !validExpression(to))
                return {VbaExecutionStatus::Unsupported,"Name requires two pathname expressions separated by As",line};
            statements.push_back({Statement::Type::FileRename,from,to,line});
        } else if ((normalized=="kill" || normalized.starts_with("kill ") || normalized.starts_with("kill\t")) && !assignsBareName(source,"kill")) {
            const auto expression=trim(source.substr(4));
            if(expression.empty() || !validExpression(expression))
                return {VbaExecutionStatus::Unsupported,"Kill requires a pathname expression",line};
            statements.push_back({Statement::Type::FileKill,{},expression,line});
        } else if ((normalized=="filecopy" || normalized.starts_with("filecopy ") || normalized.starts_with("filecopy\t")) && !assignsBareName(source,"filecopy")) {
            const auto expression=trim(source.substr(8)); const auto arguments=splitArguments(expression);
            if(arguments.size()!=2 || arguments[0].empty() || arguments[1].empty()
               || !validExpression(arguments[0]) || !validExpression(arguments[1]))
                return {VbaExecutionStatus::Unsupported,"FileCopy requires source and destination expressions",line};
            statements.push_back({Statement::Type::FileCopy,{},expression,line});
        } else if ((normalized.starts_with("mkdir ") || normalized.starts_with("mkdir\t")
                   || normalized.starts_with("rmdir ") || normalized.starts_with("rmdir\t")
                   || normalized.starts_with("chdir ") || normalized.starts_with("chdir\t")) && !assignsBareName(source,normalized.substr(0,5))) {
            const auto expression=trim(source.substr(5));
            if (!validExpression(expression)) return {VbaExecutionStatus::Unsupported,"Directory statement requires one path expression",line};
            statements.push_back({Statement::Type::Directory,normalized.substr(0,5),expression,line});
        } else if (lower(source).starts_with("sendkeys ") || lower(source).starts_with("sendkeys\t")) {
            const auto args=splitArguments(trim(source.substr(8)));
            if (args.empty() || args.size()>2 || args[0].empty()) return {VbaExecutionStatus::Unsupported,"SendKeys arguments unsupported",line};
            for (const auto& arg:args) if (!arg.empty() && !validExpression(arg)) return {VbaExecutionStatus::Unsupported,"SendKeys expression unsupported",line};
            statements.push_back({Statement::Type::SendKeys,{},trim(source.substr(8)),line});
        } else if (std::regex_match(source, match, label)) {
            const auto name=canonicalLineLabel(match[1].str());
            if(!name) return {VbaExecutionStatus::Unsupported,"Line number exceeds 65529",line};
            if (!labels.emplace(*name, static_cast<int>(statements.size())).second)
                return {VbaExecutionStatus::Error, "Doppelte Sprungmarke", line};
            statements.push_back({Statement::Type::Label,*name,{},line});
        } else if (std::regex_match(source, match, goSub)) {
            const auto name=canonicalLineLabel(match[1].str());
            if(!name) return {VbaExecutionStatus::Unsupported,"Line number exceeds 65529",line};
            statements.push_back({Statement::Type::GoSub,*name,{},line});
        } else if (normalized=="gosub") {
            return {VbaExecutionStatus::Unsupported,"GoSub requires a local line label",line};
        } else if (normalized=="return") {
            statements.push_back({Statement::Type::ReturnSubroutine,{},{},line});
        } else if (normalized.starts_with("return ") || normalized.starts_with("return\t")) {
            return {VbaExecutionStatus::Unsupported,"Return does not take an argument",line};
        } else if (std::regex_match(source, match, jump)) {
            const auto name=canonicalLineLabel(match[1].str());
            if(!name) return {VbaExecutionStatus::Unsupported,"Line number exceeds 65529",line};
            statements.push_back({Statement::Type::Jump,*name,{},line});
        } else if (std::regex_match(source, match, errorHandler)) {
            const auto name=canonicalLineLabel(match[1].str());
            if(!name) return {VbaExecutionStatus::Unsupported,"Line number exceeds 65529",line};
            statements.push_back({*name=="0" ? Statement::Type::ErrorDisable : Statement::Type::ErrorHandler,*name,{},line});
        } else if (std::regex_match(source,match,computedBranch)) {
            const auto expression=trim(match[1].str());
            if(!validExpression(expression)) return {VbaExecutionStatus::Unsupported,"Computed branch selector is unsupported",line};
            Statement branch{lower(match[2].str())=="goto" ? Statement::Type::ComputedJump : Statement::Type::ComputedGoSub,{},expression,line};
            for(const auto& value:splitArguments(match[3].str())) {
                if(!std::regex_match(value,destinationName)) return {VbaExecutionStatus::Error,"Computed branch requires local line labels",line};
                const auto name=canonicalLineLabel(value);
                if(!name) return {VbaExecutionStatus::Unsupported,"Line number exceeds 65529",line};
                branch.destinations.push_back({*name});
            }
            if(branch.destinations.empty()) return {VbaExecutionStatus::Error,"Computed branch requires destinations",line};
            statements.push_back(std::move(branch));
        } else if (lower(source) == "resume next") {
            statements.push_back({Statement::Type::ResumeNext, {}, {}, line});
        } else if (lower(source) == "resume" || std::regex_match(source,resumeZero)) {
            statements.push_back({Statement::Type::Resume, {}, {}, line});
        } else if (std::regex_match(source, match, resumeLabel)) {
            const auto name=canonicalLineLabel(match[1].str());
            if(!name) return {VbaExecutionStatus::Unsupported,"Line number exceeds 65529",line};
            statements.push_back({Statement::Type::ResumeLabel,*name,{},line});
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
            const auto names=match[1].matched ? splitArguments(match[1].str()) : std::vector<std::string>{""};
            const auto groupStart=statements.size();
            static const std::regex identifier(R"(^[A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*$)");
            for(const auto& name:names) {
                if(match[1].matched && !std::regex_match(name,identifier)) return {VbaExecutionStatus::Error,"Invalid Next control variable",line};
                if(blocks.empty() || blocks.back()!=Statement::Type::For || forStarts.empty())
                    return {VbaExecutionStatus::Error,"Next ohne For",line};
                const int startIndex=forStarts.back();
                if(!name.empty() && lower(name)!=lower(statements[startIndex].argument))
                    return {VbaExecutionStatus::Error,"Next-Variable passt nicht zum For",line};
                forStarts.pop_back(); blocks.pop_back();
                const int nextIndex=static_cast<int>(statements.size());
                statements[startIndex].partner=nextIndex;
                statements.push_back({Statement::Type::Next,{},{},line}); statements.back().partner=startIndex;
            }
            if(statements.size()>groupStart)
                for(auto index=groupStart;index<statements.size();++index) statements[index].sourceStatementEnd=static_cast<int>(statements.size()-1);
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
            Statement resize{Statement::Type::ResizeArray,{},{},line,{},preserve ? "preserve" : ""};
            for(auto entry:splitArguments(declaration)) {
                std::string elementType;
                if(const auto as=outsideKeyword(entry,"as");as!=std::string::npos) {
                    elementType=lower(trim(entry.substr(as+2))); entry=trim(entry.substr(0,as));
                    if(elementType!="variant" && elementType!="byte" && elementType!="integer" && elementType!="long" && elementType!="single"
                        && elementType!="double" && elementType!="currency" && elementType!="date" && elementType!="boolean" && elementType!="string" && elementType!="object")
                        return {VbaExecutionStatus::Unsupported,"ReDim element type unsupported",line};
                }
                const auto target=memberTarget(entry);
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
                resize.resizeTargets.push_back({target->name,*target->arguments,std::move(elementType)});
            }
            if(resize.resizeTargets.empty()) return {VbaExecutionStatus::Error,"ReDim requires an array",line};
            resize.argument=resize.resizeTargets.front().name;
            statements.push_back(std::move(resize));
        } else if (lower(source).starts_with("erase ") || lower(source).starts_with("erase\t")) {
            std::string names;
            for(const auto& entry:splitArguments(trim(source.substr(5)))) {
                const auto target=memberTarget(entry);
                if(!target || target->arguments) return {VbaExecutionStatus::Unsupported,"Erase requires supported named array variables",line};
                if(!names.empty()) names+=',';
                names+=target->name;
            }
            if(names.empty()) return {VbaExecutionStatus::Error,"Erase requires an array",line};
            statements.push_back({Statement::Type::EraseArray,std::move(names),{},line});
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
                const auto type = nativeDeclaredType(lower(declared[2].str()));
                if (!type.empty() && type != "variant" && type != "object" && type != "collection" && nativeScriptingObjectType(type).empty() && !nativeDaoType(type)
                    && nativeAccessObjectType(type).empty()
                    && type != "string" && type != "boolean"
                    && type != "integer" && type != "long" && type != "single" && type != "double" && type != "currency" && type != "date"
                    && type != "byte") return {VbaExecutionStatus::Unsupported, "Variablentyp: " + type, line};
                Statement local{Statement::Type::Dim, declared[1].str(), type, line};
                local.persistent=persistent; statements.push_back(std::move(local));
            }
        } else if (std::regex_match(source,match,stringAlignment)) {
            const auto expression=trim(match[3].str());
            if(!validExpression(expression)) return {VbaExecutionStatus::Unsupported,"String alignment expression unsupported",line};
            Statement alignment{Statement::Type::StringAlign,match[2].str(),expression,line};
            alignment.stepExpression=lower(match[1].str()); statements.push_back(std::move(alignment));
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
        add(instruction.destinations.capacity()*sizeof(BranchDestination));
        for(const auto& destination:instruction.destinations) add(destination.label.capacity());
        add(instruction.resizeTargets.capacity()*sizeof(ArrayResizeTarget));
        for(const auto& target:instruction.resizeTargets) add(target.name.capacity()+target.bounds.capacity()+target.elementType.capacity());
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
std::unordered_set<std::string> moduleProcedureNames(std::string_view module) {
    static const std::regex declaration(R"(^\s*(?:(?:Public|Private|Friend|Static)\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*))",std::regex::icase);
    std::unordered_set<std::string> names; std::istringstream source{std::string(module)}; std::string line;
    while(std::getline(source,line)) {
        line=withoutComment(line); std::smatch match;
        if(std::regex_search(line,match,declaration)) names.insert(lower(match[1].str()));
    }
    return names;
}
std::optional<VbaArgumentSignature> localProcedureSignature(std::string_view module,std::string_view name) {
    std::vector<Statement> statements; std::vector<Parameter> parameters; bool function=false,cancel=false; std::string resultType;
    if(parseProcedure(module,name,statements,function,cancel,&parameters,true,&resultType).status!=VbaExecutionStatus::Completed) return {};
    VbaArgumentSignature signature; signature.localProcedure=true; signature.function=function; signature.resultType=std::move(resultType);
    for(const auto& parameter:parameters) {
        signature.names.push_back(lower(parameter.name));
        signature.required+=!parameter.optional && !parameter.paramArray;
        const auto& type=parameter.type;
        signature.objectArguments.push_back(!parameter.paramArray && (type=="object" || type=="collection"
            || !nativeScriptingObjectType(type).empty() || nativeDaoType(type) || !nativeAccessObjectType(type).empty()));
    }
    return signature;
}

}

std::optional<VbaArgumentSignature> nativeVbaProcedureSignature(std::string_view module,std::string_view procedure) {
    if(module.size()>1024*1024) return {};
    return localProcedureSignature(module,procedure);
}

std::vector<int> vbaDebugExecutableLines(std::string_view source,std::string_view procedure) {
    if(source.size()>1024*1024) throw VbaRuntimeError(3035,"Debugger source exceeds one MiB");std::vector<Statement> statements;bool function=false,cancel=false;const auto parsed=parseProcedure(source,procedure,statements,function,cancel);if(parsed.status!=VbaExecutionStatus::Completed) return {};std::vector<int> lines;for(const auto& statement:statements) if(debugExecutable(statement)) lines.push_back(statement.line);std::sort(lines.begin(),lines.end());lines.erase(std::unique(lines.begin(),lines.end()),lines.end());return lines;
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
    state->procedures=moduleProcedureNames(module); state->procedureInventoryComplete=true;
    const std::regex procedure(R"(^\s*(?:(?:Public|Private|Friend|Static)\s+)?(?:Sub|Function|Property)\s+)",std::regex::icase);
    const std::regex declaration(R"(^\s*(Dim|Private|Public|Global)\s+(.+)$)",std::regex::icase);
    const std::regex variable(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)(?:\s+As\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*(?:\.[A-Za-z_][A-Za-z0-9_]*)?))?$)",std::regex::icase);
    const std::regex dynamicArray(R"(^([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)\s*\(\s*(.*)\s*\)\s*(?:As\s+([A-Za-z]+))?$)",std::regex::icase);
    const auto scalarConstants=moduleConstantDeclarations(module); std::size_t constantIndex=0; int sourceLine=0;
    VbaVariables knownConstants; const auto constantMode=moduleCompareMode(module);
    std::istringstream stream{std::string(module)}; std::string line;
    while (std::getline(stream,line)) {
        ++sourceLine;
        line=withoutComment(line); if (std::regex_search(line,procedure)) break;
        std::smatch match;
        if(constantIndex<scalarConstants.size() && scalarConstants[constantIndex].line==sourceLine) {
            while(constantIndex<scalarConstants.size() && scalarConstants[constantIndex].line==sourceLine) {
                const auto& declaration=scalarConstants[constantIndex++]; const auto& item=declaration.value;
                if(state->values.contains(item.name)) throw VbaRuntimeError(5,"Duplicate module constant");
                state->values.emplace(item.name,item.value); state->types.emplace(item.name,item.type); state->constants.insert(item.name);
                knownConstants.emplace(item.name,item.value);
                if(declaration.publicValue) state->publicVariables.insert(item.name);
                if(!declaration.qualifiedName.empty()) {
                    const auto& qualified=declaration.qualifiedName;
                    state->values.emplace(qualified,item.value); state->types.emplace(qualified,"long");
                    state->constants.insert(qualified); knownConstants.emplace(qualified,item.value);
                    if(declaration.publicValue) state->publicVariables.insert(qualified);
                }
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
            const auto name=lower(field[1].str()); const auto type=nativeDeclaredType(lower(field[2].str()));
            if (!type.empty() && type!="variant" && type!="boolean" && type!="string" && type!="integer" && type!="long" && type!="single" && type!="double" && type!="currency" && type!="date" && type!="byte" && type!="collection" && type!="object" && nativeScriptingObjectType(type).empty() && !nativeDaoType(type) && nativeAccessObjectType(type).empty())
                throw VbaRuntimeError(3251,"Module field type unsupported");
            if (state->values.contains(name)) throw VbaRuntimeError(5,"Duplicate module field");
            state->values.emplace(name,(type=="collection" || type=="object" || !nativeScriptingObjectType(type).empty() || nativeDaoType(type) || !nativeAccessObjectType(type).empty()) ? VbaValue::reference(nullptr) : type=="boolean" ? VbaValue::logical(false) : type=="string" ? VbaValue::string("") : type=="integer" ? VbaValue::integer(0) : type=="long" ? VbaValue::longInteger(0) : type=="single" ? VbaValue::single(0) : type=="double" ? VbaValue::numeric(0) : type=="currency" ? VbaValue::currency(0) : type=="date" ? VbaValue::dateSerial(0) : type=="byte" ? VbaValue::byte(0) : VbaValue::empty());
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
    if(const auto access=nativeAccessObjectType(type);!access.empty()) {
        if(value.kind!=VbaValue::Kind::Object || (value.object && !value.object->isNativeType(access)))
            throw VbaRuntimeError(13,"Access binding requires the declared native class or interface");
        return value;
    }
    if(!nativeScriptingObjectType(type).empty()) {
        if(value.kind!=VbaValue::Kind::Object || (value.object && !value.object->isNativeType(nativeScriptingObjectType(type))))
            throw VbaRuntimeError(13,"Scripting binding requires the declared native object type");
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
        if(value.kind==VbaValue::Kind::Decimal) {const auto result=decimalInteger(value,type=="byte" ? 0 : type=="integer" ? -32768 : -2147483648LL,type=="byte" ? 255 : type=="integer" ? 32767 : 2147483647LL);return type=="byte" ? VbaValue::byte(static_cast<unsigned char>(result)) : type=="integer" ? VbaValue::integer(static_cast<std::int16_t>(result)) : VbaValue::longInteger(static_cast<std::int32_t>(result));}
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
    if (type=="object" || type=="collection" || !nativeScriptingObjectType(type).empty() || type=="control" || type=="form" || type.starts_with("access.") || type.starts_with("dao.")) return VbaValue::reference(nullptr);
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
struct ReferenceBinding {
    VbaValue* source{}; std::string type; bool indexedArray{}; bool paramArray{}; bool fixedArray{};
    std::function<VbaValue()> read;
    std::function<void(const VbaValue&)> write;
    VbaValue* arrayRoot{};
    bool temporaryCell{};
};
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
    if(vbaDebuggerBusy()) {VbaExecutionResult failure{VbaExecutionStatus::Error,"VBA-Ausführung während eines Debugger-Halts nicht verfügbar",0};failure.errorNumber=3251;return failure;}
    VbaExecutionScope executionFrame;
    ArrayElementRootScope arrayElementScope;
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
            if(declaration.status==VbaExecutionStatus::ProcedureMissing
               && (lower(instruction.argument)=="doevents" || lower(instruction.argument)=="syscmd")) continue;
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
        bool currencyCounter{};
        VbaValue currencyEnd,currencyStep;
    };
    std::vector<Frame> frames;
    std::vector<LoopFrame> loops;
    struct WithFrame { std::size_t instruction; VbaValue value; bool pending{}; };
    std::vector<WithFrame> withFrames,faultWithFrames;
    struct SubroutineFrame { std::size_t instruction,frameDepth,loopDepth,withDepth; bool active; };
    std::vector<SubroutineFrame> subroutines,faultSubroutines;
    VbaVariables locals;
    std::unordered_map<std::string,std::string> localTypes;
    std::unordered_set<std::string> fixedArrays;
    std::unordered_set<std::string> constantNames;
    const int arrayBase=moduleArrayBase(module);
    const auto integralType=[](std::string_view type) { return type=="integer" || type=="long"; };
    const auto checkedType=[&](std::string_view type) { return type.ends_with("()") || integralType(type) || type=="boolean" || type=="string" || type=="single" || type=="double" || type=="currency" || type=="date" || type=="byte" || type=="collection" || type=="object" || !nativeScriptingObjectType(type).empty() || nativeDaoType(type) || !nativeAccessObjectType(type).empty(); };
    const auto scalarType=[&](std::string_view type) { return checkedType(type) && type!="collection" && type!="object" && nativeScriptingObjectType(type).empty() && !nativeDaoType(type) && nativeAccessObjectType(type).empty() && !type.ends_with("()"); };
    const auto objectType=[](std::string_view type) { return type=="object" || type=="collection" || !nativeScriptingObjectType(type).empty() || type.starts_with("dao.") || type.starts_with("access.") || !nativeAccessObjectType(type).empty(); };
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
            if(!omitted && references && i<references->size() && parameter.byReference && (*references)[i].arrayRoot)
                activeArrayElementRoots.push_back((*references)[i].arrayRoot);
            if(!omitted && references && i<references->size() && parameter.byReference && (*references)[i].read)
                *(*references)[i].source=(*references)[i].read();
            const auto value=implicitCancel ? VbaValue::numeric(0) : !omitted
                ? parameter.byReference && references && i<references->size() && (*references)[i].source ? *(*references)[i].source : arguments[i]
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
    struct Alias {
        std::string name; VbaValue* source; VbaValue before; std::string type;
        std::function<VbaValue()> read;
        std::function<void(const VbaValue&)> write;
        VbaValue* arrayRoot{};
    };
    std::vector<Alias> aliases;
    if (references && !implicitCancel)
        for (std::size_t i = 0; i < parameters.size() && i<references->size(); ++i)
            if (!omittedParameters[i] && parameters[i].byReference && (*references)[i].source)
                aliases.push_back({lower(parameters[i].name), (*references)[i].source, *(*references)[i].source,(*references)[i].type,
                    (*references)[i].read,(*references)[i].write,(*references)[i].arrayRoot});
    // Locals/aliases now own the semantic arguments. Internal transport copies
    // must not keep replaced objects alive; public API argument vectors remain.
    if(references && !implicitCancel) {
        std::unordered_set<VbaValue*> liveSources;
        for(std::size_t i=0;i<parameters.size() && i<references->size();++i)
            if(!omittedParameters[i] && parameters[i].byReference && !parameters[i].paramArray && (*references)[i].source)
                liveSources.insert((*references)[i].source);
        for(const auto& binding:*references)
            if(binding.temporaryCell && binding.source && !liveSources.contains(binding.source)) *binding.source=VbaValue::empty();
        for(auto& value:arguments) value=VbaValue::empty();
    }
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
        bool constantTarget=constantNames.contains(lower(instruction.argument));
        if(instruction.type==Statement::Type::IndexedAssign && lower(instruction.argument)=="mid" && !locals.contains("mid") && instruction.stepExpression!="object") {
            const auto arguments=splitArguments(instruction.endExpression);
            if(!arguments.empty()) constantTarget=constantNames.contains(lower(trim(arguments[0])));
        }

        if(instruction.type==Statement::Type::ObjectAssign) {
            constantTarget=false;
            const auto dot=memberSeparator(instruction.argument);
            if(dot!=std::string::npos) {
                const auto root=lower(trim(instruction.argument.substr(0,dot)));
                const auto member=memberTarget(trim(instruction.argument.substr(dot+1)));
                if(member && !member->arguments && !locals.contains(root) && !host.hasInstanceIdentifier(root))
                    constantTarget=constantNames.contains(root+"."+lower(member->name));
            }
        }
        if(instruction.type==Statement::Type::EraseArray)
            for(const auto& target:splitArguments(instruction.argument)) if(constantNames.contains(lower(target))) constantTarget=true;
        if(instruction.type==Statement::Type::ResizeArray)
            for(const auto& target:instruction.resizeTargets) if(constantNames.contains(lower(target.name))) constantTarget=true;
        if(constantTarget && (instruction.type==Statement::Type::ObjectAssign || instruction.type==Statement::Type::SetVar || instruction.type==Statement::Type::StringAlign || instruction.type==Statement::Type::ObjectSet
            || instruction.type==Statement::Type::IndexedAssign || instruction.type==Statement::Type::ResizeArray || instruction.type==Statement::Type::EraseArray
            || instruction.type==Statement::Type::For || instruction.type==Statement::Type::ForEach)) {
            VbaExecutionResult result{VbaExecutionStatus::Unsupported,"Assignment to a constant is not permitted",instruction.line}; result.errorNumber=3251; return result;
        }
        const auto type=variableType(lower(instruction.argument));
        if(((instruction.type==Statement::Type::ForEach || instruction.type==Statement::Type::ObjectSet) && (scalarType(type) || type=="paramarray"))
            || (instruction.type==Statement::Type::For && (type.ends_with("()") || type=="boolean" || type=="string" || type=="date" || objectType(type)))) {
            VbaExecutionResult failure{VbaExecutionStatus::Unsupported,"Declared variable type is unsupported for this loop/object target",instruction.line}; failure.errorNumber=3251; return failure;
        }
    }
    const auto assignVariable=[&](const std::string& name,const VbaValue& value,bool arrayElement=false) {
        if(!arrayElement) {
            const auto local=locals.find(name); const VbaValue* source=local==locals.end() ? nullptr : &local->second;
            for(const auto& alias:aliases) if(alias.name==name) { source=alias.source; break; }
            if(source && std::find(activeArrayElementRoots.begin(),activeArrayElementRoots.end(),source)!=activeArrayElementRoots.end())
                throw VbaRuntimeError(3251,"Array replacement while a ByRef element is active is not implemented");
        }
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
        for (auto& alias : aliases) {
            if(alias.read) *alias.source=alias.read();
            locals[alias.name] = *alias.source; alias.before = *alias.source;
        }
    };
    AliasRefreshScope* aliasFrame=nullptr;
    const auto storeAliases = [&] {
        bool changed=false;
        for (auto& alias : aliases)
            if (!equalValue(locals[alias.name], alias.before)) {
                try {
                    const auto value=coerceVbaModuleValue(locals[alias.name],alias.type);
                    if(alias.write) alias.write(value);
                    *alias.source=value; changed=true;
                }
                catch (...) { locals[alias.name]=*alias.source; alias.before=*alias.source; throw; }
            }
        if(changed) for(auto* frame:activeAliasFrames) if(frame!=aliasFrame) frame->refresh();
        refreshAliases();
    };
    AliasRefreshScope observingAliases(refreshAliases); aliasFrame=&observingAliases;
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
                            const VbaArgumentReferences& referenceInfo, bool needsValue) -> VbaValue {
        storeAliases();
        const auto& names=referenceInfo.names;
        struct ArrayCell { VbaValue* root; std::size_t offset; std::shared_ptr<const void> allocation; VbaValue value; };
        std::vector<std::shared_ptr<ArrayCell>> cells;
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
        for (std::size_t argument=0;argument<argumentRefs.size();++argument) {
            const auto& expression=argumentRefs[argument];
            auto found = locals.find(lower(expression));
            const auto type=variableType(lower(expression));
            ReferenceBinding binding{found==locals.end() ? nullptr : &found->second,checkedType(type) ? type : std::string{}};
            for (const auto& alias:aliases) if (alias.name==lower(expression))
                binding={alias.source,alias.type,false,false,false,alias.read,alias.write,alias.arrayRoot};
            binding.paramArray=type=="paramarray";
            binding.fixedArray=fixedArrays.contains(lower(expression));
            if(constantNames.contains(lower(expression))) binding.source=nullptr;
            if(expression=="@array_element") binding.indexedArray=true;
            if(const auto target=memberTarget(expression); target && target->arguments) {
                const auto array=locals.find(lower(target->name));
                if(array!=locals.end() && isVbaArray(array->second)) binding.indexedArray=true;
            }
            if(argument<referenceInfo.arrayElements.size() && referenceInfo.arrayElements[argument]) {
                const auto& reference=*referenceInfo.arrayElements[argument];
                binding.indexedArray=true;
                const auto root=locals.find(reference.root); VbaValue* source=root==locals.end() ? nullptr : &root->second;
                bool nested=false;
                for(const auto& alias:aliases) if(alias.name==reference.root) { source=alias.source; nested=bool(alias.read) || bool(alias.write); break; }
                if(source && reference.allocation && !nested && variableType(reference.root)!="paramarray" && !constantNames.contains(reference.root)) {
                    std::shared_ptr<ArrayCell> cell;
                    for(const auto& candidate:cells) if(candidate->root==source && candidate->offset==reference.offset && candidate->allocation==reference.allocation) { cell=candidate; break; }
                    if(!cell) { cell=std::make_shared<ArrayCell>(ArrayCell{source,reference.offset,reference.allocation,actuals[argument]}); cells.push_back(cell); }
                    const auto read=[cell]() {
                        if(!isVbaArray(*cell->root) || vbaArrayAllocation(*cell->root)!=cell->allocation)
                            throw VbaRuntimeError(3251,"ByRef array source allocation changed");
                        return vbaArrayLinearElement(*cell->root,cell->offset);
                    };
                    const auto write=[cell,read](const VbaValue& value) {
                        (void)read();
                        *cell->root=withVbaArrayLinearElement(*cell->root,cell->offset,value);
                    };
                    binding={&cell->value,reference.elementType,false,false,false,read,write,source};
                    binding.temporaryCell=true;
                }
            }
            bindings.push_back(std::move(binding));
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
            if(declaration.status==VbaExecutionStatus::ProcedureMissing && lower(std::string(name))=="syscmd") {
                static const std::vector<std::string> names{"action","argument2","argument3"};
                std::vector<VbaValue> ordered(3,VbaValue::missing()); std::vector<bool> assigned(3);
                bool namedSeen=false;
                for(std::size_t index=0;index<actuals.size();++index) {
                    auto target=index;
                    if(!argumentNames[index].empty()) {
                        namedSeen=true; const auto argument=lower(argumentNames[index]);
                        const auto found=std::find(names.begin(),names.end(),argument);
                        if(found==names.end()) throw VbaRuntimeError(448,"Named argument not found");
                        target=static_cast<std::size_t>(found-names.begin());
                    } else if(namedSeen) throw VbaRuntimeError(450,"Positional arguments cannot follow named arguments");
                    if(target>=ordered.size() || assigned[target]) throw VbaRuntimeError(450,"SysCmd argument supplied more than once or too many arguments");
                    assigned[target]=true; ordered[target]=actuals[index];
                }
                return host.applicationObject()->member("SysCmd",ordered);
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
                assigned[target]=true; ordered[target]=std::move(actuals[index]); orderedBindings[target]=bindings[index]; mapping.push_back(target);
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
        if(nested.status==VbaExecutionStatus::ProcedureMissing && lower(std::string(name))=="syscmd")
            return host.applicationObject()->member("SysCmd",actuals);
        if(nested.status==VbaExecutionStatus::ProcedureMissing && lower(std::string(name))=="doevents") {
            if(!actuals.empty()) throw VbaRuntimeError(450,"DoEvents takes no arguments");
            const auto result=VbaValue::integer(host.doEvents());
            if(vbaExecutionStopRequested()) throw VbaExecutionStopped{};
            return result;
        }
        if(vbaExecutionStopRequested()) throw VbaExecutionStopped{};
        for(std::size_t i=0;i<argumentRefs.size();++i) {
            if(!bindings[i].source) continue;
            for(const auto& loop:loops) if(loop.arrayRoot==lower(argumentRefs[i]) && !sameVbaArrayAllocation(loop.array,*bindings[i].source))
                throw VbaRuntimeError(3251,"Array source replacement during For Each is not implemented");
        }
        // ByRef assignments already update their actual source through aliases.
        // ByVal arguments must never copy an old value back over that source.
        refreshAliases();
        if (nested.status != VbaExecutionStatus::Completed)
            throw VbaRuntimeError(nested.errorNumber ? nested.errorNumber : 5,
                nested.status == VbaExecutionStatus::ProcedureMissing ? "Prozedur nicht gefunden" : nested.explanation,
                nested.errorSource,nested.errorHelpFile,nested.errorHelpContext);
        if (needsValue && !nested.typedReturnValue) throw VbaRuntimeError(5, "Sub hat keinen Rückgabewert");
        return nested.typedReturnValue ? *nested.typedReturnValue : VbaValue::empty();
    };
    std::unordered_map<std::string,std::optional<VbaArgumentSignature>> signatureCache;
    VbaFunctionResolver resolver = [&](std::string_view name, std::vector<VbaValue>& args,
                                            const VbaArgumentReferences& refs) {
        return invoke(name, args, refs, true);
    };
    std::optional<std::unordered_set<std::string>> fallbackProcedures;
    resolver.argumentSignature=[&](std::string_view name) -> std::optional<VbaArgumentSignature> {
        const auto key=lower(std::string(name));
        const auto& state=stateOwners.front();
        if(!state->procedureInventoryComplete && !fallbackProcedures) fallbackProcedures=moduleProcedureNames(module);
        const auto& names=state->procedureInventoryComplete ? state->procedures : *fallbackProcedures;
        if(!names.contains(key)) return {};
        const auto found=signatureCache.find(key);
        if(found!=signatureCache.end()) return found->second;
        auto signature=localProcedureSignature(module,name);
        if(!signature) throw VbaRuntimeError(3251,"Declared local procedure binding is unsupported");
        return signatureCache.emplace(key,std::move(signature)).first->second;
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
    const auto selectedBranch=[&](const Statement& instruction) -> const BranchDestination* {
        const auto selector=static_cast<int>(coerceVbaModuleValue(scalarVbaValue(evaluate(instruction.expression)),"integer").number);
        if(selector<0 || selector>255) throw VbaRuntimeError(5,"Computed branch selector is outside 0-255");
        if(selector==0 || static_cast<std::size_t>(selector)>instruction.destinations.size()) return nullptr;
        return &instruction.destinations[static_cast<std::size_t>(selector-1)];
    };
    const auto leaveFailedNextGroup=[&](std::size_t first) {
        const int end=instructions[first].sourceStatementEnd;
        if(end<0) return;
        for(auto index=first;index<=static_cast<std::size_t>(end);++index)
            if(instructions[index].type==Statement::Type::Next && !loops.empty() && loops.back().forIndex==instructions[index].partner)
                loops.pop_back();
    };
    bool active = true;
    int currentLine = 0;
    int steps = 0;
    int effects = 0;
    std::optional<VbaDebugFrameScope> debugger;
    if(vbaDebuggerActive()) debugger.emplace(module,procedure,locals,[&](std::string_view variable) {const auto name=std::string(variable);if(const auto type=localTypes.find(name);type!=localTypes.end()) return type->second;for(const auto& alias:aliases) if(alias.name==name) return alias.type;return name=="err" ? std::string("ErrObject") : std::string("variant");},[&] {return resumeNext || (handler>=0 && !handlingError);},[&] {return VbaDebugError{errorState->number,errorState->description,errorState->source,errorState->helpFile,errorState->helpContext};});
    try {
        for (std::size_t ip = 0; ip < instructions.size(); ++ip) {
            if(vbaExecutionStopRequested()) throw VbaExecutionStopped{};
            const auto& instruction = instructions[ip];
            const auto executingIp=ip;
            currentLine = instruction.line;
            if (++steps > 10000)
                return {VbaExecutionStatus::Error, "VBA-Ausführungsbudget (10.000 Schritte) erreicht", currentLine};
            refreshAliases();
            if(debugger) {
                bool executable=active && debugExecutable(instruction);
                if(instruction.type==Statement::Type::ElseIf || instruction.type==Statement::Type::Case) executable=!frames.empty() && frames.back().parent && !frames.back().condition && !frames.back().pending;
                debugger->checkpoint(currentLine,executingIp,executable);
            }
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
            case Statement::Type::FileKill:
                if(active) nativeFileKill(scalarVbaValue(evaluate(instruction.expression)));
                break;
            case Statement::Type::FileRename:
                if(active) {
                    const auto source=scalarVbaValue(evaluate(instruction.argument));
                    const auto destination=scalarVbaValue(evaluate(instruction.expression));
                    nativeFileRename(source,destination);
                }
                break;
            case Statement::Type::FileCopy:
                if(active) {
                    const auto expressions=splitArguments(instruction.expression);
                    const auto source=scalarVbaValue(evaluate(expressions[0]));
                    const auto destination=scalarVbaValue(evaluate(expressions[1]));
                    nativeFileCopy(source,destination);
                }
                break;
            case Statement::Type::Directory:
                if (active) {
                    const auto value=scalarVbaValue(evaluate(instruction.expression));
                    if (value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null directory path");
                    const auto text=value.asText();
                    if (text.empty() && instruction.argument=="chdir") throw VbaRuntimeError(76,"Directory path not found");
                    if (text.empty() || text.find('\0')!=std::string::npos || text.find('*')!=std::string::npos || text.find('?')!=std::string::npos)
                        throw VbaRuntimeError(52,"Invalid directory path");
                    if (text.find('\\')!=std::string::npos || text.find(':')!=std::string::npos)
                        throw VbaRuntimeError(3251,"Windows directory paths require an unavailable drive adapter");
                    const std::filesystem::path path(text);
                    std::error_code failure;
                    if (instruction.argument=="mkdir") {
                        const bool created=std::filesystem::create_directory(path,failure);
                        if (!created && !failure) throw VbaRuntimeError(75,"Directory already exists");
                    } else if (instruction.argument=="chdir") {
                        std::filesystem::current_path(path,failure);
                    } else {
                        const auto status=std::filesystem::symlink_status(path,failure);
                        if (!failure) {
                            if (std::filesystem::is_symlink(status)) throw VbaRuntimeError(3251,"RmDir symbolic-link behavior is unsupported");
                            if (!std::filesystem::exists(status)) throw VbaRuntimeError(76,"Directory path not found");
                            if (!std::filesystem::is_directory(status)) throw VbaRuntimeError(75,"RmDir requires a directory");
                            const auto current=std::filesystem::current_path(failure);
                            if (!failure && std::filesystem::equivalent(path,current,failure))
                                throw VbaRuntimeError(75,"Cannot remove the current directory");
                            if (!failure && ::rmdir(path.c_str())!=0) failure=std::error_code(errno,std::generic_category());
                        }
                    }
                    if (failure) {
                        if (failure==std::errc::no_such_file_or_directory || failure==std::errc::not_a_directory)
                            throw VbaRuntimeError(76,"Directory path not found");
                        throw VbaRuntimeError(75,"Directory operation failed");
                    }
                }
                break;
            case Statement::Type::SendKeys:
                if (active) {
                    const auto expressions=splitArguments(instruction.expression);
                    const auto keys=scalarVbaValue(evaluate(expressions[0]));
                    if (keys.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null SendKeys input");
                    const bool wait=expressions.size()>1 && !expressions[1].empty() ? coerceVbaModuleValue(scalarVbaValue(evaluate(expressions[1])),"boolean").boolean : false;
                    host.sendKeys(keys.asText(),wait);
                }
                break;
            case Statement::Type::RaiseError:
                if(active) (void)errorState->member("Raise",{evaluate(instruction.expression)});
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
            case Statement::Type::ComputedJump:
                if (active) {
                    const BranchDestination direct{{},instruction.partner,instruction.conditionalUnwind,instruction.withUnwind,instruction.forUnwind};
                    const auto* destination=instruction.type==Statement::Type::ComputedJump ? selectedBranch(instruction) : &direct;
                    if(!destination) break;
                    const auto frameDepth=subroutines.empty() ? 0 : subroutines.back().frameDepth;
                    const auto loopDepth=subroutines.empty() ? 0 : subroutines.back().loopDepth;
                    const auto withDepth=subroutines.empty() ? 0 : subroutines.back().withDepth;
                    const auto unwindFrames=static_cast<std::size_t>(destination->conditionalUnwind);
                    const auto unwindLoops=static_cast<std::size_t>(destination->forUnwind);
                    const auto unwindWith=static_cast<std::size_t>(destination->withUnwind);
                    if(frames.size()<frameDepth+unwindFrames || loops.size()<loopDepth+unwindLoops || withFrames.size()<withDepth+unwindWith)
                        throw VbaRuntimeError(5,"Invalid GoTo scope");
                    frames.resize(frames.size()-unwindFrames); loops.resize(loops.size()-unwindLoops); withFrames.resize(withFrames.size()-unwindWith);
                    active=true; ip=static_cast<std::size_t>(destination->partner);
                }
                break;
            case Statement::Type::GoSub:
            case Statement::Type::ComputedGoSub:
                if(active) {
                    const auto* destination=instruction.type==Statement::Type::ComputedGoSub ? selectedBranch(instruction) : nullptr;
                    if(instruction.type==Statement::Type::ComputedGoSub && !destination) break;
                    if(subroutines.size()>=48) throw VbaRuntimeError(28,"GoSub return stack exhausted");
                    subroutines.push_back({ip,frames.size(),loops.size(),withFrames.size(),active});
                    active=true; ip=static_cast<std::size_t>(destination ? destination->partner : instruction.partner);
                }
                break;
            case Statement::Type::ReturnSubroutine:
                if(active) {
                    if(subroutines.empty()) {
                        if(handlingError && !faultSubroutines.empty()) throw VbaRuntimeError(3251,"Return from an error handler into a suspended GoSub requires Resume");
                        throw VbaRuntimeError(3,"Return without GoSub");
                    }
                    const auto caller=subroutines.back();
                    if(frames.size()<caller.frameDepth || loops.size()<caller.loopDepth || withFrames.size()<caller.withDepth)
                        throw VbaRuntimeError(5,"Invalid GoSub return scope");
                    frames.resize(caller.frameDepth); loops.resize(caller.loopDepth); withFrames.resize(caller.withDepth);
                    subroutines.pop_back(); active=caller.active; ip=caller.instruction;
                }
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
                    errorState->clear();
                    const int groupEnd=instructions[faultInstruction].sourceStatementEnd;
                    if (instruction.type == Statement::Type::ResumeLabel) {
                        frames.clear(); loops.clear(); withFrames.clear(); active = true;
                        faultFrames.clear(); faultLoops.clear(); faultWithFrames.clear();
                        subroutines.clear(); faultSubroutines.clear();
                        ip = static_cast<std::size_t>(instruction.partner);
                    } else {
                        frames = faultFrames; loops = faultLoops; withFrames=faultWithFrames; active = faultActive;
                        subroutines=faultSubroutines;
                        faultFrames.clear(); faultLoops.clear(); faultWithFrames.clear(); faultSubroutines.clear();
                        if(instruction.type==Statement::Type::Resume) {
                            ip=faultInstruction-1;
                        } else if(groupEnd>=0) {
                            leaveFailedNextGroup(faultInstruction);
                            ip=static_cast<std::size_t>(groupEnd);
                        } else ip=faultInstruction;
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
                    VbaArgumentReferences references;
                    const auto signature=resolver.argumentSignature(instruction.argument);
                    for(const auto& expression:expressions) {
                        const auto parsed=callArgument(expression);
                        auto argument=parsed.expression.empty() ? VbaEvaluatedArgument{VbaValue::missing(),{}}
                            : evaluateVbaArgumentExpression(parsed.expression,host,locals,resolver,withFrames.empty() ? nullptr : &withFrames.back().value,compareMode,arrayBase,
                                signature && signature->objectArgument(actuals.size(),parsed.name));
                        actuals.push_back(std::move(argument.value)); references.arrayElements.push_back(std::move(argument.arrayElement));
                        references.names.push_back(parsed.name.empty() ? parsed.expression : "@named:"+parsed.name+":"+parsed.expression);
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
                if(active) for(const auto& target:instruction.resizeTargets) {
                    const auto name=lower(target.name),type=variableType(name);
                    if(fixedArrays.contains(name)) throw VbaRuntimeError(3251,"ReDim of a fixed array is not implemented");
                    if(type=="paramarray") throw VbaRuntimeError(3251,"ParamArray resize is not implemented");
                    if(!target.elementType.empty() && type.empty() && !locals.contains(name))
                        throw VbaRuntimeError(3251,"Typed ReDim requires an existing Variant variable");
                    const bool declaredArray=type.ends_with("()");
                    const auto declaredElement=declaredArray ? type.substr(0,type.size()-2) : std::string{};
                    if(!type.empty() && type!="variant" && !declaredArray) throw VbaRuntimeError(13,"ReDim requires an array or Variant variable");
                    if(declaredArray && !target.elementType.empty() && target.elementType!=declaredElement)
                        throw VbaRuntimeError(13,"ReDim cannot change a declared array element type");
                    std::vector<VbaArrayDimension> bounds;
                    for(const auto& dimension:splitArguments(target.bounds)) {
                        const auto range=outsideKeyword(dimension,"to");
                        const auto lowerBound=range==std::string::npos ? arrayBase : static_cast<int>(coerceVbaModuleValue(scalarVbaValue(evaluate(dimension.substr(0,range))),"long").number);
                        const auto upperBound=static_cast<int>(coerceVbaModuleValue(scalarVbaValue(evaluate(range==std::string::npos ? dimension : dimension.substr(range+2))),"long").number);
                        bounds.push_back({lowerBound,upperBound});
                    }
                    refreshAliases();
                    const auto found=locals.find(name);
                    const auto original=found==locals.end() ? VbaValue::empty() : found->second;
                    assignVariable(name,resizeVbaArrayWithBounds(original,std::move(bounds),instruction.stepExpression=="preserve",declaredArray ? declaredElement : target.elementType));
                    storeAliases();
                }
                break;
            case Statement::Type::EraseArray:
                if(active) for(const auto& target:splitArguments(instruction.argument)) {
                    const auto name=lower(target);
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
                    // Mid statements share indexed syntax; actual variables named Mid keep their normal array/object path.
                    if(name=="mid" && initial==locals.end() && instruction.stepExpression!="object") {
                        const auto parts=splitArguments(instruction.endExpression);
                        if(parts.size()<2 || parts.size()>3) throw VbaRuntimeError(450,"Mid assignment requires a String variable, start and optional length");
                        static const std::regex targetName(R"(^[A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*$)");
                        const auto destination=trim(parts[0]);
                        if(!std::regex_match(destination,targetName)) throw VbaRuntimeError(3251,"Mid assignment requires a simple String variable");
                        const auto start=static_cast<std::int64_t>(coerceVbaModuleValue(scalarVbaValue(evaluate(parts[1])),"long").number);
                        const auto requested=parts.size()==3 ? static_cast<std::int64_t>(coerceVbaModuleValue(scalarVbaValue(evaluate(parts[2])),"long").number) : std::numeric_limits<std::int32_t>::max();
                        const auto input=coerceVbaModuleValue(scalarVbaValue(evaluate(instruction.expression)),"string");
                        const auto target=locals.find(lower(destination));
                        if(target==locals.end() || target->second.kind!=VbaValue::Kind::Text) throw VbaRuntimeError(3251,"Mid assignment requires an existing String variable");
                        const auto width=vbaStringCodeUnits(target->second.text),inputWidth=vbaStringCodeUnits(input.text);
                        if(start<1 || requested<0) throw VbaRuntimeError(5,"Invalid Mid assignment start or length");
                        if(width>std::numeric_limits<std::int32_t>::max() || inputWidth>std::numeric_limits<std::int32_t>::max()) throw VbaRuntimeError(6,"Mid assignment length overflow");
                        if(start>static_cast<std::int64_t>(width)) throw VbaRuntimeError(3251,"Out-of-range Mid assignment requires reference validation");
                        const auto count=std::min({requested,static_cast<std::int64_t>(inputWidth),static_cast<std::int64_t>(width)-start+1});
                        if(count==0) break;
                        if(start+count>std::numeric_limits<std::int32_t>::max()) throw VbaRuntimeError(6,"Mid assignment boundary overflow");
                        const VbaVariables values{{"_mid_target",target->second},{"_mid_input",input},
                            {"_mid_prefix",VbaValue::longInteger(static_cast<std::int32_t>(start-1))},
                            {"_mid_count",VbaValue::longInteger(static_cast<std::int32_t>(count))},
                            {"_mid_tail",VbaValue::longInteger(static_cast<std::int32_t>(start+count))}};
                        const auto result=evaluateVbaExpression("Left$(_mid_target,_mid_prefix) & Left$(_mid_input,_mid_count) & Mid$(_mid_target,_mid_tail)",host,values);
                        assignVariable(lower(destination),result); break;
                    }
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
            case Statement::Type::StringAlign:
                if(active) {
                    const auto input=coerceVbaModuleValue(scalarVbaValue(evaluate(instruction.expression)),"string");
                    const auto name=lower(instruction.argument); const auto target=locals.find(name);
                    if(target==locals.end() || target->second.kind!=VbaValue::Kind::Text)
                        throw VbaRuntimeError(3251,"LSet/RSet requires an existing String variable");
                    const auto width=vbaStringCodeUnits(target->second.text);
                    if(width>static_cast<std::size_t>(std::numeric_limits<std::int32_t>::max())) throw VbaRuntimeError(6,"String alignment length overflow");
                    const VbaVariables values{{"_align_input",input},{"_align_width",VbaValue::longInteger(static_cast<std::int32_t>(width))}};
                    auto result=evaluateVbaExpression("Left$(_align_input,_align_width)",host,values).text;
                    const auto padding=width-vbaStringCodeUnits(result);
                    if(instruction.stepExpression=="rset") result.insert(0,padding,' ');
                    else result.append(padding,' ');
                    assignVariable(name,VbaValue::string(std::move(result)));
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
                const auto counter=lower(instruction.argument),type=variableType(counter);
                if(type=="currency") {
                    const auto initial=convertCurrency(scalarVbaValue(evaluate(instruction.expression)));
                    const auto end=convertCurrency(scalarVbaValue(evaluate(instruction.endExpression)));
                    const auto step=convertCurrency(scalarVbaValue(evaluate(instruction.stepExpression)));
                    assignVariable(counter,initial);
                    if((step.currencyUnits>=0 && initial.currencyUnits>end.currencyUnits)
                        || (step.currencyUnits<0 && initial.currencyUnits<end.currencyUnits)) {
                        ip=static_cast<std::size_t>(instruction.partner);
                    } else {
                        LoopFrame loop{static_cast<int>(ip)};
                        loop.currencyCounter=true;loop.currencyEnd=end;loop.currencyStep=step;
                        loop.frameDepth=frames.size();loop.withDepth=withFrames.size();
                        loops.push_back(std::move(loop));
                    }
                    break;
                }
                const auto initial = evaluate(instruction.expression).asNumber();
                const auto end = evaluate(instruction.endExpression).asNumber();
                const auto step = evaluate(instruction.stepExpression).asNumber();
                if (!std::isfinite(initial) || !std::isfinite(end) || !std::isfinite(step))
                    throw VbaRuntimeError(6,"For bound or step overflows Double");
                // MS-VBAL 5.4.2.3 keeps evaluated limits/steps uncoerced. Only
                // the counter assignment uses its declared numeric type.
                assignVariable(counter,VbaValue::numeric(initial));
                const auto assigned=locals.at(counter).asNumber();
                if ((step >= 0 && assigned > end) || (step < 0 && assigned < end)) {
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
                    if(isVbaArray(loop.array)) {
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
                if(loops.back().currencyCounter) {
                    const auto next=currencyBinary("+",convertCurrency(scalarVbaValue(locals.at(name))),loops.back().currencyStep);
                    assignVariable(name,next);
                    const auto& loop=loops.back();
                    if((loop.currencyStep.currencyUnits>=0 && next.currencyUnits<=loop.currencyEnd.currencyUnits)
                        || (loop.currencyStep.currencyUnits<0 && next.currencyUnits>=loop.currencyEnd.currencyUnits))
                        ip=static_cast<std::size_t>(loop.forIndex);
                    else loops.pop_back();
                    break;
                }
                const auto nextValue = locals[name].asNumber() + loops.back().step;
                if (!std::isfinite(nextValue)) throw VbaRuntimeError(6,"For counter overflow");
                assignVariable(name,VbaValue::numeric(nextValue));
                const auto assigned=locals.at(name).asNumber();
                if ((loops.back().step >= 0 && assigned <= loops.back().end)
                    || (loops.back().step < 0 && assigned >= loops.back().end))
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
                const auto* typed = dynamic_cast<const VbaRuntimeError*>(&exception);
                bool debuggerResume=false;
                if(!resumeNext && (handler<0 || handlingError)) {
                    const VbaDebugError failure{typed ? typed->number : 5,exception.what(),typed ? typed->source : std::string{},typed ? typed->helpFile : std::string{},typed ? typed->helpContext : 0};
                    if(debugger) debuggerResume=debugger->unhandled(failure,true,currentLine);
                    if(!debuggerResume) throw;
                }
                errorState->number = typed ? typed->number : 5;
                errorState->description = exception.what();
                errorState->source=typed ? typed->source : std::string{};
                errorState->helpFile=typed ? typed->helpFile : std::string{};
                errorState->helpContext=typed ? typed->helpContext : 0;
                if (!resumeNext && !debuggerResume) {
                    handlingError = true; faultInstruction = instruction.sourceStatementEnd>=0 ? executingIp : ip;
                    faultFrames = frames; faultLoops = loops; faultWithFrames=withFrames; faultActive = active;
                    faultSubroutines=subroutines;
                    frames.clear(); loops.clear(); withFrames.clear(); subroutines.clear(); active = true;
                    ip = static_cast<std::size_t>(handler);
                } else if(instruction.sourceStatementEnd>=0) {
                    leaveFailedNextGroup(executingIp);
                    ip=static_cast<std::size_t>(instruction.sourceStatementEnd);
                }
            }
        }
    } catch (const VbaExecutionStopped&) {
        return {VbaExecutionStatus::Completed,{},currentLine,true};
    } catch (const std::exception& exc) {
        if(vbaExecutionStopRequested()) return {VbaExecutionStatus::Completed,{},currentLine,true};
        try { storeAliases(); } catch(const std::exception&) { /* Preserve the original execution failure. */ }
        copyArguments();
        VbaExecutionResult failure{VbaExecutionStatus::Error, exc.what(), currentLine};
        const auto* typed = dynamic_cast<const VbaRuntimeError*>(&exc);
        failure.errorNumber = typed ? typed->number : 5;
        if(typed) { failure.errorSource=typed->source; failure.errorHelpFile=typed->helpFile; failure.errorHelpContext=typed->helpContext; }
        if(debugger) try {(void)debugger->unhandled({failure.errorNumber,failure.explanation,failure.errorSource,failure.errorHelpFile,failure.errorHelpContext},false,currentLine);}catch(const VbaExecutionStopped&) {return {VbaExecutionStatus::Completed,{},currentLine,true};}
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
        VbaFunctionResolver resolver=[&](std::string_view name,std::vector<VbaValue>& args,const std::vector<std::string>&) {
            const auto epoch=vbaExecutionEpoch();
            auto outcome=executeVbaProcedureWithArguments(module,name,host,args);
            if (outcome.status==VbaExecutionStatus::ProcedureMissing) {
                if (const auto external=host.moduleForProcedure(name)) outcome=executeVbaProcedureWithArguments(*external,name,host,args);
            }
            if(vbaExecutionStopRequested() || epoch!=vbaExecutionEpoch()) throw VbaExecutionStopped{};
            line=outcome.line;
            if (outcome.status!=VbaExecutionStatus::Completed || !outcome.typedReturnValue)
                throw VbaRuntimeError(outcome.errorNumber ? outcome.errorNumber : 5,outcome.explanation.empty() ? "Event function execution failed" : outcome.explanation,
                    outcome.errorSource,outcome.errorHelpFile,outcome.errorHelpContext);
            return *outcome.typedReturnValue;
        };
        const auto declared=moduleProcedureNames(module);
        resolver.argumentSignature=[&](std::string_view name) -> std::optional<VbaArgumentSignature> {
            if(!declared.contains(lower(std::string(name)))) return {};
            auto signature=localProcedureSignature(module,name);
            if(!signature) throw VbaRuntimeError(3251,"Declared event function binding is unsupported");
            return signature;
        };
        const auto value=evaluateVbaExpression(source,host,{},resolver);
        VbaExecutionResult result{VbaExecutionStatus::Completed,{},line};
        result.typedReturnValue=std::make_shared<VbaValue>(value);
        if (value.kind!=VbaValue::Kind::Object && value.kind!=VbaValue::Kind::Error && value.kind!=VbaValue::Kind::Missing && !isVbaArray(value)) result.returnValue=value.asField();
        return result;
    } catch (const VbaExecutionStopped&) {
        return {VbaExecutionStatus::Completed,{},line,true};
    } catch (const VbaRuntimeError& error) {
        VbaExecutionResult result{VbaExecutionStatus::Error,error.what(),line}; result.errorNumber=error.number; result.errorSource=error.source; result.errorHelpFile=error.helpFile; result.errorHelpContext=error.helpContext; return result;
    } catch (const std::exception&) {
        VbaExecutionResult result{VbaExecutionStatus::Error,"Event expression execution failed",line}; result.errorNumber=5; return result;
    }
}
}  // namespace linaccess
