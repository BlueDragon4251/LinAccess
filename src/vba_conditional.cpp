#include "vba_conditional.hpp"
#include "vba_expression.hpp"
#include "vba_array.hpp"
#include <algorithm>
#include <cctype>
#include <cmath>
#include <functional>
#include <map>
#include <regex>
#include <sstream>
#include <unordered_set>

namespace linaccess {
namespace {
std::string folded(std::string text) {for(auto& c:text) if(c>='A' && c<='Z') c+=('a'-'A');return text;}
std::string stripped(std::string text) {
    const auto first=text.find_first_not_of(" \t\r\n"),last=text.find_last_not_of(" \t\r\n");
    return first==std::string::npos ? std::string{} : text.substr(first,last-first+1);
}
std::string uncommented(std::string_view text) {
    bool quoted=false;
    for(std::size_t i=0;i<text.size();++i) {
        if(text[i]=='"') {if(quoted && i+1<text.size() && text[i+1]=='"') {++i;continue;}quoted=!quoted;}
        else if(text[i]=='\'' && !quoted) return stripped(std::string(text.substr(0,i)));
    }
    return stripped(std::string(text));
}
class ConditionalHost final:public VbaHost {
public:
    std::optional<std::string> field(std::string_view) override {throw VbaRuntimeError(3251,"Conditional expressions cannot read runtime state");}
    void setField(std::string_view,std::optional<std::string>) override {throw VbaRuntimeError(3251,"Conditional expressions cannot write runtime state");}
    void message(std::string_view) override {throw VbaRuntimeError(3251,"Conditional expressions cannot display messages");}
    void saveRecord() override {throw VbaRuntimeError(3251,"Conditional expressions cannot save records");}
};
VbaValue expressionValue(const std::string& expression,const std::function<VbaValue(const std::string&)>& resolve) {
    if(expression.empty() || expression.size()>4096) throw VbaRuntimeError(3251,"Conditional expression length unsupported");
    static const std::regex number(R"((?:&[Hh][0-9A-Fa-f]+|&[Oo][0-7]+|(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?)[%&!#@]?)");
    const std::unordered_set<std::string> functions{"int","fix","abs","sgn","len","lenb","cbool","cbyte","ccur","cdate","cdbl","cint","clng","csng","cstr","cvar"};
    const std::unordered_set<std::string> terminals{"true","false","empty","null","nothing","not","and","or","xor","eqv","imp","mod","like","is"};
    const std::unordered_set<std::string> operators{"not","and","or","xor","eqv","imp","mod","like","is"};
    const std::unordered_set<std::string> runtimeNames{"date","time","now","timer","new","me","currentdb","dbengine","application","forms","reports","screen","doevents","rnd"};
    const auto word=[](unsigned char c){return std::isalnum(c) || c=='_' || c>=128;};
    VbaVariables values;std::size_t textCost=expression.size();
    for(std::size_t i=0;i<expression.size();) {
        const auto c=static_cast<unsigned char>(expression[i]);
        if(std::isspace(c)) {++i;continue;}
        if(c=='"' || c=='#') {
            const auto quote=c;bool closed=false;++i;
            while(i<expression.size()) {if(static_cast<unsigned char>(expression[i++])!=quote) continue;if(quote=='"' && i<expression.size() && expression[i]=='"') {++i;continue;}closed=true;break;}
            if(!closed) throw VbaRuntimeError(3251,"Unclosed conditional literal");continue;
        }
        if(std::isdigit(c) || (c=='.' && i+1<expression.size() && std::isdigit(static_cast<unsigned char>(expression[i+1]))) || (c=='&' && i+1<expression.size() && (expression[i+1]=='H' || expression[i+1]=='h' || expression[i+1]=='O' || expression[i+1]=='o'))) {
            std::smatch match;const auto rest=expression.substr(i);if(!std::regex_search(rest,match,number,std::regex_constants::match_continuous)) throw VbaRuntimeError(3251,"Invalid conditional number");i+=match.length();continue;
        }
        if(std::isalpha(c) || c=='_' || c>=128) {
            const auto start=i++;while(i<expression.size() && word(static_cast<unsigned char>(expression[i]))) ++i;
            const auto name=folded(expression.substr(start,i-start));auto next=i;while(next<expression.size() && std::isspace(static_cast<unsigned char>(expression[next]))) ++next;
            if(next<expression.size() && expression[next]=='(') {if(!functions.contains(name) && !operators.contains(name)) throw VbaRuntimeError(3251,"Conditional function is unsupported");}
            else if(!terminals.contains(name)) {if(runtimeNames.contains(name)) throw VbaRuntimeError(3251,"Runtime names are unavailable in conditional expressions");auto value=resolve(name);
                if(value.text.size()>65536-textCost) throw VbaRuntimeError(3035,"Conditional expression text budget exceeded");
                textCost+=value.text.size();values[name]=std::move(value);}
            continue;
        }
        if(std::string("+-*/\\^&()=<>").find(static_cast<char>(c))==std::string::npos) throw VbaRuntimeError(3251,"Unsupported conditional expression token");++i;
    }
    ConditionalHost host;auto value=evaluateVbaExpression(expression,host,values);
    if(value.kind==VbaValue::Kind::Missing || value.kind==VbaValue::Kind::Error || isVbaArray(value) || (value.kind==VbaValue::Kind::Object && value.object)) throw VbaRuntimeError(3251,"Conditional result must be a constant scalar or Nothing");
    if(value.text.size()>65536) throw VbaRuntimeError(3035,"Conditional constant text limit exceeded");
    if(value.kind==VbaValue::Kind::Number && !std::isfinite(value.number)) throw VbaRuntimeError(6,"Conditional constant overflow");
    return value;
}
struct LogicalLine {std::string raw,text,name,expression;int number{},physical{};enum Kind{Ordinary,Const,If,ElseIf,Else,EndIf} kind{Ordinary};};
}
std::string preprocessVbaConditional(std::string_view source) {
    if(source.size()>1024*1024) throw VbaConditionalError(3035,"Conditional source exceeds one MiB",0);
    if(source.find('#')==std::string_view::npos) return std::string(source);
    static const std::regex definition(R"(^#\s*Const\s+([A-Za-z_\x80-\xFF][A-Za-z0-9_\x80-\xFF]*)[%&!#@$]?\s*=\s*(.+)$)",std::regex::icase);
    static const std::regex branch(R"(^#\s*(If|ElseIf)\s+(.+)\s+Then\s*$)",std::regex::icase);
    static const std::regex otherwise(R"(^#\s*Else\s*$)",std::regex::icase),ending(R"(^#\s*(?:EndIf|End\s+If)\s*$)",std::regex::icase);
    std::vector<LogicalLine> lines;std::map<std::string,std::size_t> definitions;std::istringstream input{std::string(source)};std::string physical;int number=0,directives=0;
    while(std::getline(input,physical)) {
        LogicalLine line;line.number=++number;line.physical=1;line.raw=physical;line.text=uncommented(physical);
        if(!input.eof()) line.raw+='\n';
        if(!line.text.empty() && line.text.front()=='#') {
            while(line.text.ends_with(" _")) {
                if(line.physical>=32 || !std::getline(input,physical)) throw VbaConditionalError(3251,"Conditional continuation is incomplete or too long",line.number);
                ++number;++line.physical;line.text.resize(line.text.size()-2);line.text+=' '+uncommented(physical);line.raw+=physical;if(!input.eof()) line.raw+='\n';
            }
            if(++directives>4096) throw VbaConditionalError(3035,"Conditional directive limit exceeded",line.number);
            std::smatch match;
            if(std::regex_match(line.text,match,definition)) {
                line.kind=LogicalLine::Const;line.name=folded(match[1].str());line.expression=match[2].str();
                if(!definitions.emplace(line.name,lines.size()).second) throw VbaConditionalError(5,"Duplicate conditional constant",line.number);
            } else if(std::regex_match(line.text,match,branch)) {line.kind=folded(match[1].str())=="if" ? LogicalLine::If : LogicalLine::ElseIf;line.expression=match[2].str();}
            else if(std::regex_match(line.text,otherwise)) line.kind=LogicalLine::Else;
            else if(std::regex_match(line.text,ending)) line.kind=LogicalLine::EndIf;
            else throw VbaConditionalError(3251,"Conditional directive syntax unsupported",line.number);
        }
        lines.push_back(std::move(line));
    }
    if(!directives) return std::string(source);
    VbaVariables constants{{"win16",VbaValue::logical(false)},{"win32",VbaValue::logical(false)},{"win64",VbaValue::logical(false)},{"mac",VbaValue::logical(false)},{"vba6",VbaValue::logical(true)},{"vba7",VbaValue::logical(false)}};
    std::unordered_set<std::string> ready,resolving;std::size_t constantText=0;
    std::function<VbaValue(const std::string&)> resolve=[&](const std::string& name)->VbaValue {
        const auto found=definitions.find(name);if(found==definitions.end()) {const auto value=constants.find(name);return value==constants.end() ? VbaValue::empty() : value->second;}
        if(ready.contains(name)) return constants.at(name);
        const auto& line=lines[found->second];
        if(resolving.size()>=64 || !resolving.insert(name).second) throw VbaConditionalError(3251,"Conditional constant dependency cycle or depth limit",line.number);
        try {auto value=expressionValue(line.expression,resolve);if(value.text.size()>1024*1024-constantText) throw VbaRuntimeError(3035,"Conditional module constant text budget exceeded");constantText+=value.text.size();constants[name]=value;ready.insert(name);resolving.erase(name);return value;}
        catch(const VbaConditionalError&) {throw;}
        catch(const VbaRuntimeError& error) {throw VbaConditionalError(error.number,"Conditional constant evaluation failed",line.number);}
    };
    // Even constants in excluded blocks are bindings and must be evaluated.
    for(const auto& [name,index]:definitions) (void)resolve(name);
    struct Frame {bool parent{},taken{},included{},otherwise{};int line{};};std::vector<Frame> frames;std::string output;output.reserve(source.size());
    for(const auto& line:lines) {
        try {
            const bool included=frames.empty() || frames.back().included;
            if(line.kind==LogicalLine::Ordinary) {if(included) output+=line.raw;else for(char c:line.raw) if(c=='\n') output+='\n';continue;}
            for(char c:line.raw) if(c=='\n') output+='\n';
            if(line.kind==LogicalLine::Const) continue;
            if(line.kind==LogicalLine::If || line.kind==LogicalLine::ElseIf) {
                const auto value=expressionValue(line.expression,resolve);
                const bool condition=coerceVbaModuleValue(value,"boolean").boolean;
                if(line.kind==LogicalLine::If) {
                    if(frames.size()>=64) throw VbaRuntimeError(3035,"Conditional nesting limit exceeded");
                    frames.push_back({included,condition,included && condition,false,line.number});
                } else {
                    if(frames.empty() || frames.back().otherwise) throw VbaRuntimeError(3251,"Conditional ElseIf has no matching If");
                    auto& frame=frames.back();frame.included=frame.parent && !frame.taken && condition;frame.taken|=condition;
                }
            } else {
                if(frames.empty()) throw VbaRuntimeError(3251,"Conditional terminator has no matching If");
                if(line.kind==LogicalLine::Else) {auto& frame=frames.back();if(frame.otherwise) throw VbaRuntimeError(3251,"Duplicate conditional Else");frame.otherwise=true;frame.included=frame.parent && !frame.taken;frame.taken=true;}
                else frames.pop_back();
            }
        } catch(const VbaConditionalError&) {throw;}
          catch(const VbaRuntimeError& error) {throw VbaConditionalError(error.number,"Conditional directive evaluation or structure failed",line.number);}
    }
    if(!frames.empty()) throw VbaConditionalError(3251,"Conditional If is unclosed",frames.back().line);
    return output;
}
}
