#include "vba_expression.hpp"
#include "vba_color.hpp"
#include "vba_currency.hpp"
#include "vba_analyzer.hpp"

#include <algorithm>
#include <chrono>
#include <cmath>
#include <cctype>
#include <ctime>
#include <iomanip>
#include <cstdint>
#include <limits>
#include <locale>
#include <sstream>
#include <stdexcept>
#include <vector>

namespace linaccess {
namespace {
std::string lower(std::string s) {
    std::transform(s.begin(), s.end(), s.begin(), [](unsigned char ch) {
        return static_cast<char>(std::tolower(ch));
    });
    return s;
}
VbaRuntimeError error(const std::string& reason, int code = 5) { return VbaRuntimeError(code, reason); }
VbaValue clockValue(const std::string& word) {
    const auto now=std::chrono::system_clock::to_time_t(std::chrono::system_clock::now());
    std::tm local{};
    if (!::localtime_r(&now,&local)) throw error("Systemdatum nicht verfügbar");
    std::ostringstream formatted;
    formatted.imbue(std::locale::classic());
    formatted<<std::put_time(&local,word=="now" ? "%Y-%m-%d %H:%M:%S" : word=="date" ? "%Y-%m-%d" : "%H:%M:%S");
    return VbaValue::string(formatted.str());
}

// Runtime strings use UTF-8 at the API boundary; VBA substring positions use
// UTF-16 code units. Do not silently split a UTF-8 byte sequence.
std::u16string utf16String(std::string_view input) {
    std::u16string output;
    for (std::size_t i=0;i<input.size();) {
        const auto first=static_cast<unsigned char>(input[i++]);
        std::uint32_t point=first,minimum=0; unsigned continuation=0;
        if (first<128) {}
        else if (first>=0xc2 && first<=0xdf) { point=first&31; continuation=1; minimum=0x80; }
        else if (first>=0xe0 && first<=0xef) { point=first&15; continuation=2; minimum=0x800; }
        else if (first>=0xf0 && first<=0xf4) { point=first&7; continuation=3; minimum=0x10000; }
        else throw error("Invalid UTF-8 runtime string",13);
        if (continuation>input.size()-i) throw error("Truncated UTF-8 runtime string",13);
        for (unsigned n=0;n<continuation;++n) {
            const auto byte=static_cast<unsigned char>(input[i++]);
            if ((byte&0xc0)!=0x80) throw error("Invalid UTF-8 continuation",13);
            point=(point<<6)|(byte&63);
        }
        if (point<minimum || point>0x10ffff || (point>=0xd800 && point<=0xdfff)) throw error("Invalid Unicode scalar",13);
        if (point<=0xffff) output.push_back(static_cast<char16_t>(point));
        else { point-=0x10000; output.push_back(static_cast<char16_t>(0xd800+(point>>10))); output.push_back(static_cast<char16_t>(0xdc00+(point&1023))); }
    }
    return output;
}
std::string utf8String(std::u16string_view input) {
    std::string output;
    for (std::size_t i=0;i<input.size();++i) {
        std::uint32_t point=input[i];
        if (point>=0xd800 && point<=0xdbff) {
            if (i+1>=input.size() || input[i+1]<0xdc00 || input[i+1]>0xdfff) throw error("Isolated UTF-16 surrogate unsupported",3251);
            point=0x10000+((point-0xd800)<<10)+(input[++i]-0xdc00);
        } else if (point>=0xdc00 && point<=0xdfff) throw error("Isolated UTF-16 surrogate unsupported",3251);
        if (point<0x80) output.push_back(static_cast<char>(point));
        else if (point<0x800) { output.push_back(static_cast<char>(0xc0|(point>>6))); output.push_back(static_cast<char>(0x80|(point&63))); }
        else if (point<0x10000) { output.push_back(static_cast<char>(0xe0|(point>>12))); output.push_back(static_cast<char>(0x80|((point>>6)&63))); output.push_back(static_cast<char>(0x80|(point&63))); }
        else { output.push_back(static_cast<char>(0xf0|(point>>18))); output.push_back(static_cast<char>(0x80|((point>>12)&63))); output.push_back(static_cast<char>(0x80|((point>>6)&63))); output.push_back(static_cast<char>(0x80|(point&63))); }
    }
    return output;
}

struct Token {
    enum class Type { Word, Number, Text, Bracket, Symbol, End } type{Type::End};
    std::string value;
};

std::vector<Token> tokenize(std::string_view input,bool accessExpression=false) {
    std::vector<Token> tokens;
    for (std::size_t pos = 0; pos < input.size();) {
        const char ch = input[pos];
        if (std::isspace(static_cast<unsigned char>(ch))) { ++pos; continue; }
        if (std::isalpha(static_cast<unsigned char>(ch)) || static_cast<unsigned char>(ch) >= 0x80 || ch == '_') {
            const auto start = pos++;
            while (pos < input.size() &&
                   (std::isalnum(static_cast<unsigned char>(input[pos])) || static_cast<unsigned char>(input[pos]) >= 0x80 || input[pos] == '_')) ++pos;
            if (pos<input.size() && input[pos]=='$') ++pos;
            tokens.push_back({Token::Type::Word, std::string(input.substr(start, pos - start))});
        } else if (std::isdigit(static_cast<unsigned char>(ch)) || (ch=='.' && pos+1<input.size() && std::isdigit(static_cast<unsigned char>(input[pos+1])))) {
            const auto start=pos;
            while (pos<input.size() && std::isdigit(static_cast<unsigned char>(input[pos]))) ++pos;
            if (pos<input.size() && input[pos]=='.') {
                ++pos; while (pos<input.size() && std::isdigit(static_cast<unsigned char>(input[pos]))) ++pos;
            }
            if (pos<input.size() && (input[pos]=='e' || input[pos]=='E')) {
                ++pos; if(pos<input.size() && (input[pos]=='+' || input[pos]=='-')) ++pos;
                const auto exponent=pos;
                while(pos<input.size() && std::isdigit(static_cast<unsigned char>(input[pos]))) ++pos;
                if(pos==exponent) throw error("Numeric exponent requires digits",5);
            }
            tokens.push_back({Token::Type::Number,std::string(input.substr(start,pos-start))});
        } else if (ch == '"' || (accessExpression && ch=='\'')) {
            const auto delimiter=ch;
            ++pos;
            std::string value;
            bool closed = false;
            while (pos < input.size()) {
                if (input[pos] == delimiter) {
                    ++pos;
                    if (pos < input.size() && input[pos] == delimiter) {
                        value.push_back(delimiter); ++pos;
                    } else { closed = true; break; }
                } else value.push_back(input[pos++]);
            }
            if (!closed) throw error("VBA-String nicht geschlossen");
            tokens.push_back({Token::Type::Text, value});
        } else if (ch == '[') {
            const auto start = ++pos;
            while (pos < input.size() && input[pos] != ']') ++pos;
            if (pos == input.size()) throw error("Schließende ] fehlt");
            tokens.push_back({Token::Type::Bracket, std::string(input.substr(start, pos - start))});
            ++pos;
        } else {
            std::string symbol(1, ch);
            if ((ch == '<' || ch == '>') && pos + 1 < input.size() &&
                (input[pos + 1] == '=' || (ch == '<' && input[pos + 1] == '>'))) {
                symbol.push_back(input[++pos]);
            }
            if (std::string("+-*/\\^&=<>!.,()").find(ch) == std::string::npos && !(accessExpression && ch==';'))
                throw error("Nicht unterstütztes VBA-Ausdruckszeichen: " + symbol);
            tokens.push_back({Token::Type::Symbol, symbol});
            ++pos;
        }
        if (tokens.size() > 512) throw error("VBA-Ausdruck zu lang");
    }
    tokens.push_back({});
    return tokens;
}

int precedence(const Token& token) {
    if (token.type!=Token::Type::Symbol && token.type!=Token::Type::Word) return 0;
    const auto text = lower(token.value);
    if (text == "imp") return 1;
    if (text == "eqv") return 2;
    if (text == "xor") return 3;
    if (text == "or") return 4;
    if (text == "and") return 5;
    if (text == "=" || text == "<>" || text == "<" || text == "<=" ||
        text == ">" || text == ">=" || text == "is") return 6;
    if (text == "&") return 7;
    if (text == "+" || text == "-") return 8;
    if (text == "mod") return 9;
    if (text == "\\") return 10;
    if (text == "*" || text == "/") return 11;
    if (text == "^") return 13;
    return 0;
}

double bankersRound(double value) {
    const auto floor = std::floor(value);
    const auto fraction = value - floor;
    return fraction < 0.5 ? floor : fraction > 0.5 ? floor + 1
        : std::fmod(std::abs(floor), 2.0) == 0 ? floor : floor + 1;
}
std::int32_t integer(const VbaValue& value) {
    const auto rounded = bankersRound(value.asNumber());
    if (!std::isfinite(rounded) || rounded < std::numeric_limits<std::int32_t>::min()
        || rounded > std::numeric_limits<std::int32_t>::max()) throw error("VBA: Long-Überlauf", 6);
    return static_cast<std::int32_t>(rounded);
}
VbaValue scalarObjectValue(VbaValue value);
VbaValue typeInspectionValue(VbaValue value);
VbaValue binary(const std::string& op, const VbaValue& left, const VbaValue& right,VbaHost* host,VbaCompareMode mode) {
    if (left.kind==VbaValue::Kind::Error || right.kind==VbaValue::Kind::Error)
        throw error("Error Variant cannot be used as an operand",13);
    if (op == "is") {
        if (left.kind != VbaValue::Kind::Object || right.kind != VbaValue::Kind::Object)
            throw error("VBA: Is benötigt Objekte");
        return VbaValue::logical(left.object == right.object);
    }
    if (left.kind==VbaValue::Kind::Object) return binary(op,scalarObjectValue(left),right,host,mode);
    if (right.kind==VbaValue::Kind::Object) return binary(op,left,scalarObjectValue(right),host,mode);
    if (op == "&") {
        if (left.kind == VbaValue::Kind::Null && right.kind == VbaValue::Kind::Null) return {};
        return VbaValue::string(left.asText() + right.asText());
    }
    if (op == "and" || op == "or") {
        const bool leftNull = left.kind == VbaValue::Kind::Null;
        const bool rightNull = right.kind == VbaValue::Kind::Null;
        if (leftNull || rightNull) {
            if (op == "and" && ((!leftNull && !left.truthy()) || (!rightNull && !right.truthy())))
                return VbaValue::logical(false);
            if (op == "or" && ((!leftNull && left.truthy()) || (!rightNull && right.truthy())))
                return VbaValue::logical(true);
            return {};
        }
        if (left.kind == VbaValue::Kind::Boolean && right.kind == VbaValue::Kind::Boolean)
            return VbaValue::logical(op == "and" ? left.boolean && right.boolean : left.boolean || right.boolean);
        return VbaValue::numeric(op == "and" ? integer(left) & integer(right) : integer(left) | integer(right));
    }
    if (op=="xor" || op=="eqv" || op=="imp") {
        const bool leftNull=left.kind==VbaValue::Kind::Null,rightNull=right.kind==VbaValue::Kind::Null;
        if (leftNull || rightNull) {
            if (op!="imp") return {};
            const auto certainTrue=[](const VbaValue& known) { return known.kind==VbaValue::Kind::Boolean ? VbaValue::logical(true) : VbaValue::longInteger(-1); };
            if (!leftNull && coerceVbaModuleValue(left,"long").number==0) return certainTrue(left);
            if (!rightNull && coerceVbaModuleValue(right,"long").number==-1) return certainTrue(right);
            return {};
        }
        if (left.kind==VbaValue::Kind::Boolean && right.kind==VbaValue::Kind::Boolean)
            return VbaValue::logical(op=="xor" ? left.boolean!=right.boolean : op=="eqv" ? left.boolean==right.boolean : !left.boolean || right.boolean);
        const auto a=static_cast<std::int32_t>(coerceVbaModuleValue(left,"long").number);
        const auto b=static_cast<std::int32_t>(coerceVbaModuleValue(right,"long").number);
        const auto answer=op=="xor" ? a^b : op=="eqv" ? ~(a^b) : (~a)|b;
        const bool narrow=left.kind==VbaValue::Kind::Number && right.kind==VbaValue::Kind::Number
            && left.numericType==VbaValue::NumericType::Integer && right.numericType==VbaValue::NumericType::Integer;
        return narrow ? VbaValue::integer(static_cast<std::int16_t>(answer)) : VbaValue::longInteger(answer);
    }
    if (op=="mod" || op=="\\") {
        if (left.kind==VbaValue::Kind::Null || right.kind==VbaValue::Kind::Null) return {};
        const auto a=static_cast<std::int64_t>(coerceVbaModuleValue(left,"long").number);
        const auto b=static_cast<std::int64_t>(coerceVbaModuleValue(right,"long").number);
        if (b==0) throw error("Integer division by zero",11);
        const auto answer=op=="mod" ? a%b : a/b;
        const bool narrow=left.kind==VbaValue::Kind::Number && right.kind==VbaValue::Kind::Number
            && left.numericType==VbaValue::NumericType::Integer && right.numericType==VbaValue::NumericType::Integer;
        if (answer<(narrow ? std::numeric_limits<std::int16_t>::min() : std::numeric_limits<std::int32_t>::min())
            || answer>(narrow ? std::numeric_limits<std::int16_t>::max() : std::numeric_limits<std::int32_t>::max()))
            throw error("Integer arithmetic overflow",6);
        return narrow ? VbaValue::integer(static_cast<std::int16_t>(answer)) : VbaValue::longInteger(static_cast<std::int32_t>(answer));
    }
    if (op=="^") {
        if (left.kind==VbaValue::Kind::Null || right.kind==VbaValue::Kind::Null) return {};
        const auto base=coerceVbaModuleValue(left,"double").number;
        const auto exponent=coerceVbaModuleValue(right,"double").number;
        if (base<0 && std::floor(exponent)!=exponent) throw error("Negative power base requires an integer exponent",5);
        if (base==0 && exponent<0) throw error("Zero-base negative powers are not validated",3251);
        const auto answer=std::pow(base,exponent);
        if (!std::isfinite(answer)) throw error("Power result overflow",6);
        if (answer==0 && base!=0) throw error("Power underflow is not validated",3251);
        return VbaValue::numeric(answer);
    }
    if (left.kind == VbaValue::Kind::Empty && right.kind == VbaValue::Kind::Text)
        return binary(op, VbaValue::string(""), right,host,mode);
    if (right.kind == VbaValue::Kind::Empty && left.kind == VbaValue::Kind::Text)
        return binary(op, left, VbaValue::string(""),host,mode);
    if (left.kind == VbaValue::Kind::Null || right.kind == VbaValue::Kind::Null)
        return {}; // Simplified three-valued logic: NULL comparisons yield NULL.

    if (op == "=" || op == "<>" || op == "<" || op == "<=" || op == ">" || op == ">=") {
        int cmp = 0;
        if (left.kind == VbaValue::Kind::Text && right.kind == VbaValue::Kind::Text) {
            cmp = host ? host->compareStrings(left.text,right.text,mode) : left.text.compare(right.text);
        } else if (left.kind==VbaValue::Kind::Currency && right.kind==VbaValue::Kind::Currency) {
            cmp=(left.currencyUnits>right.currencyUnits)-(left.currencyUnits<right.currencyUnits);
        } else if (left.kind == VbaValue::Kind::Text || right.kind == VbaValue::Kind::Text) {
            // These values are Variants: a numeric Variant sorts before a
            // String Variant, including numeric-looking and empty strings.
            // Empty was coerced above according to the other operand.
            cmp = left.kind == VbaValue::Kind::Text ? 1 : -1;
        } else {
            const auto a = left.asNumber(), b = right.asNumber();
            cmp = (a > b) - (a < b);
        }
        return VbaValue::logical(op == "=" ? cmp == 0 : op == "<>" ? cmp != 0 :
            op == "<" ? cmp < 0 : op == "<=" ? cmp <= 0 : op == ">" ? cmp > 0 : cmp >= 0);
    }
    if (op == "+" && left.kind == VbaValue::Kind::Text && right.kind == VbaValue::Kind::Text)
        return VbaValue::string(left.text + right.text);
    if (left.kind==VbaValue::Kind::Currency && right.kind==VbaValue::Kind::Currency && (op=="+" || op=="-" || op=="*"))
        return currencyBinary(op,left,right);
    const auto a = left.asNumber(), b = right.asNumber();
    double answer = 0;
    if (op == "+") answer = a + b;
    else if (op == "-") answer = a - b;
    else if (op == "*") answer = a * b;
    else if (op == "/") {
        if (b == 0) throw error("VBA: Division durch Null", 11);
        answer = a / b;
    } else throw error("Nicht unterstützter Operator " + op);
    if (!std::isfinite(answer)) throw error("VBA: Ungültiges numerisches Ergebnis");
    return VbaValue::numeric(answer);
}

class ExpressionParser {
public:
    ExpressionParser(std::string_view source, VbaHost* host, const VbaVariables& locals,
                     const VbaFunctionResolver& resolver = {},bool objectContext=false,bool accessExpression=false,const VbaValue* withContext=nullptr,bool allowWith=false,VbaCompareMode compareMode=VbaCompareMode::Binary)
        : tokens_(tokenize(source,accessExpression)), host_(host), locals_(locals), resolver_(resolver),objectContext_(objectContext),accessExpression_(accessExpression),withContext_(withContext),allowWith_(allowWith || withContext),compareMode_(compareMode) {}

    VbaValue parse() {
        auto value = expression(1,objectContext_);
        if (peek().type != Token::Type::End)
            throw error("Zusätzliche Zeichen im VBA-Ausdruck: " + peek().value);
        return value;
    }
private:
    const Token& peek() const { return tokens_[index_]; }
    Token consume() { return tokens_[index_++]; }
    bool nextSymbol(std::string_view symbol) const { return peek().type==Token::Type::Symbol && peek().value==symbol; }
    bool take(const std::string& wanted) {
        if ((peek().type!=Token::Type::Symbol && peek().type!=Token::Type::Word) || lower(peek().value) != wanted) return false;
        consume();
        return true;
    }
    void require(const std::string& wanted) {
        if (!take(wanted)) throw error("Erwartet: " + wanted);
    }
    VbaValue expression(int minPrecedence,bool objectContext=false) {
        auto lhs = primary(objectContext,minPrecedence);
        while (true) {
            const int prec = precedence(peek());
            if (prec < minPrecedence) break;
            const auto op = lower(consume().value);
            if (accessExpression_ && op=="is" && (lower(peek().value)=="null" || lower(peek().value)=="not")) {
                const bool negate=take("not"); require("null");
                if (host_ && lhs.kind!=VbaValue::Kind::Error) lhs=VbaValue::logical((lhs.kind==VbaValue::Kind::Null)!=negate);
                continue;
            }
            auto rhs = expression(prec + 1);
            if (!host_) lhs={};
            else if (accessExpression_ && (lhs.kind==VbaValue::Kind::Error || rhs.kind==VbaValue::Kind::Error))
                lhs=lhs.kind==VbaValue::Kind::Error ? lhs : rhs;
            else lhs=binary(op,lhs,rhs,host_,compareMode_);
        }
        return lhs;
    }
    VbaValue primary(bool objectContext,int minPrecedence) {
        auto value = [&] {
            if (!nextSymbol(".") && !nextSymbol("!")) return atom(objectContext,minPrecedence);
            if (!host_ && allowWith_) return VbaValue{};
            if (!withContext_) throw error("With object context unavailable",91);
            return *withContext_;
        }();
        while (nextSymbol("(") || nextSymbol(".") || nextSymbol("!")) {
            if (take("(")) {
                std::vector<VbaValue> arguments;
                if (!take(")")) {
                    do { arguments.push_back(expression(1)); } while (take(",") || (accessExpression_ && take(";")));
                    require(")");
                }
                if (host_) {
                    try {
                        if (value.kind!=VbaValue::Kind::Object || !value.object) throw error("Default member requires an object",91);
                        value=value.object->defaultMember(arguments,objectContext || nextSymbol(".") || nextSymbol("!") || nextSymbol("("));
                    } catch(const VbaRuntimeError& failure) { if (!accessExpression_) throw; value=VbaValue::errorValue(failure.number); }
                }
                continue;
            }
            const bool item = consume().value == "!";
            const auto member = consume();
            if (member.type != Token::Type::Word && member.type != Token::Type::Bracket)
                throw error("Objektmitglied erwartet");
            std::vector<VbaValue> args;
            if (take("(")) {
                if (!take(")")) {
                    do {
                        const bool omitted=(lower(member.value)=="openform" || lower(member.value)=="openreport" || lower(member.value)=="gotorecord" || lower(member.value)=="domenuitem" || lower(member.value)=="selectobject") && (nextSymbol(",") || nextSymbol(")"));
                        args.push_back(omitted ? VbaValue::empty() : expression(1));
                    } while (take(",") || (accessExpression_ && take(";")));
                    require(")");
                }
            }
            if (host_) {
                if (accessExpression_ && value.kind==VbaValue::Kind::Error) continue;
                try {
                if (value.kind != VbaValue::Kind::Object || !value.object)
                    throw error("Objektvariable nicht gesetzt", 91);
                if (item && !args.empty()) throw error("Bang lookup arguments unsupported",450);
                const bool reference=objectContext || nextSymbol(".") || nextSymbol("!");
                value = item ? (reference ? value.object->objectItem(member.value) : value.object->item(member.value))
                    : (reference ? value.object->objectMember(member.value,args) : value.object->member(member.value,args));
                } catch(const VbaRuntimeError& failure) { if (!accessExpression_) throw; value=VbaValue::errorValue(failure.number); }
            }
        }
        return value;
    }
    VbaValue atom(bool objectContext,int minPrecedence) {
        if (take("-")) {
            const auto arg = expression(std::max(12,minPrecedence));
            return host_ ? arg.kind==VbaValue::Kind::Currency ? negateCurrency(arg) : VbaValue::numeric(-arg.asNumber()) : VbaValue{};
        }
        if (take("+")) {
            const auto arg = expression(std::max(12,minPrecedence));
            return host_ ? arg.kind==VbaValue::Kind::Currency ? arg : VbaValue::numeric(arg.asNumber()) : VbaValue{};
        }
        if (take("not")) {
            const auto arg = expression(std::max(6,minPrecedence)); // comparisons bind before Not in VBA
            if (!host_) return {};
            const auto value=scalarObjectValue(arg);
            if (value.kind==VbaValue::Kind::Null) return {};
            if (value.kind==VbaValue::Kind::Boolean) return VbaValue::logical(!value.boolean);
            if (value.kind==VbaValue::Kind::Empty) return VbaValue::integer(-1);
            const auto number=static_cast<std::int32_t>(coerceVbaModuleValue(value,"long").number);
            if(value.kind==VbaValue::Kind::Number && value.numericType==VbaValue::NumericType::Byte)
                return VbaValue::byte(static_cast<unsigned char>(255-number));
            if(value.kind==VbaValue::Kind::Number && value.numericType==VbaValue::NumericType::Integer)
                return VbaValue::integer(static_cast<std::int16_t>(~number));
            return VbaValue::longInteger(~number);
        }
        if (take("(")) {
            const auto value = expression(1,objectContext);
            require(")");
            return value;
        }
        const auto tok = consume();
        if (tok.type==Token::Type::Bracket) {
            if (!host_) return {};
            const auto root=lower(tok.value);
            if (nextSymbol(".") || nextSymbol("!")) {
                if (root=="forms" || root=="formulare") return VbaValue::reference(host_->formsCollection());
                if (root=="reports" || root=="berichte") return VbaValue::reference(host_->reportsCollection());
            }
            const bool reference=objectContext || nextSymbol(".") || nextSymbol("!");
            return accessExpression_ ? host_->expressionValue(tok.value,reference) : host_->bracketValue(tok.value,reference);
        }
        if (tok.type == Token::Type::Text) return VbaValue::string(tok.value);
        if (tok.type == Token::Type::Number) {
            std::size_t consumed = 0;
            double number;
            try { number=std::stod(tok.value,&consumed); }
            catch(const std::out_of_range&) { throw error("Numeric literal overflow",6); }
            catch(const std::invalid_argument&) { throw error("Invalid numeric literal",5); }
            if (consumed!=tok.value.size()) throw error("Ungültige VBA-Zahl");
            if (!std::isfinite(number)) throw error("Numeric literal overflow",6);
            return VbaValue::numeric(number);
        }
        if (tok.type != Token::Type::Word)
            throw error("Ungültiger VBA-Ausdruck: " + tok.value);
        auto word = lower(tok.value);
        if (accessExpression_) {
            static const std::unordered_map<std::string,std::string> localized{{"summe","sum"},{"anzahl","count"},{"wenn","iif"},{"istnull","isnull"},{"istfehler","iserror"},{"datum","date"},{"jetzt","now"},{"zeit","time"},{"formulare","forms"},{"berichte","reports"}};
            if (const auto found=localized.find(word);found!=localized.end()) word=found->second;
        }
        if (word == "null") return {};
        if (word == "empty") return VbaValue::empty();
        if (word == "nothing") return VbaValue::reference(nullptr);
        if (word=="dbengine") return host_ ? VbaValue::reference(host_->databaseEngine()) : VbaValue{};
        if (word == "currentdb") {
            if (take("(")) require(")");
            return host_ ? VbaValue::reference(host_->currentDatabase()) : VbaValue{};
        }
        static const std::unordered_map<std::string,int> controlTypes{
            {"aclabel",100},{"acrectangle",101},{"acline",102},{"acoptiongroup",103},{"accommandbutton",104},
            {"acoptionbutton",105},{"accheckbox",106},{"actextbox",109},{"aclistbox",110},
            {"accombobox",111},{"acsubform",112},{"actogglebutton",122}};
        if (const auto type = controlTypes.find(word); type != controlTypes.end()) return VbaValue::numeric(type->second);
        static const std::unordered_map<std::string,int> daoTypes{
            {"dbboolean",1},{"dbbyte",2},{"dbinteger",3},{"dblong",4},{"dbcurrency",5},
            {"dbsingle",6},{"dbdouble",7},{"dbdate",8},{"dbbinary",9},{"dbtext",10},
            {"dblongbinary",11},{"dbmemo",12},{"dbguid",15}};
        if (const auto type=daoTypes.find(word);type!=daoTypes.end()) return VbaValue::longInteger(type->second);
        if (word == "dbopendynaset") return VbaValue::numeric(2);
        if (word == "dbopensnapshot") return VbaValue::numeric(4);
        if (word == "dbfailonerror") return VbaValue::numeric(128);
        if (word == "dbforceosflush") return VbaValue::numeric(1);
        static const std::unordered_map<std::string,int> variantTypes{
            {"vbempty",0},{"vbnull",1},{"vbinteger",2},{"vblong",3},{"vbsingle",4},
            {"vbdouble",5},{"vbcurrency",6},{"vbdate",7},{"vbstring",8},{"vbobject",9},
            {"vberror",10},{"vbboolean",11},{"vbvariant",12},{"vbdataobject",13},
            {"vbdecimal",14},{"vbbyte",17},{"vbarray",8192}};
        if (const auto type=variantTypes.find(word);type!=variantTypes.end()) return VbaValue::numeric(type->second);
        if (word == "true") return VbaValue::logical(true);
        if (word == "false") return VbaValue::logical(false);
        static const std::unordered_map<std::string,int> messageConstants{
            {"vbokonly",0},{"vbokcancel",1},{"vbabortretryignore",2},{"vbyesnocancel",3},{"vbyesno",4},{"vbretrycancel",5},
            {"vbcritical",16},{"vbquestion",32},{"vbexclamation",48},{"vbinformation",64},
            {"vbdefaultbutton1",0},{"vbdefaultbutton2",256},{"vbdefaultbutton3",512},{"vbdefaultbutton4",768},
            {"vbapplicationmodal",0},{"vbsystemmodal",4096},{"vbmsgboxhelpbutton",16384},{"vbmsgboxsetforeground",65536},
            {"vbmsgboxright",524288},{"vbmsgboxrtlreading",1048576},
            {"vbok",1},{"vbcancel",2},{"vbabort",3},{"vbretry",4},{"vbignore",5},{"vbyes",6},{"vbno",7}};
        if (const auto constant=messageConstants.find(word);constant!=messageConstants.end()) return VbaValue::numeric(constant->second);
        if (word == "vbcrlf") return VbaValue::string("\r\n");
        if (word == "vbnewline") return VbaValue::string("\n");
        if (word == "vbtab") return VbaValue::string("\t");
        static const std::unordered_map<std::string,int> keyConstants{{"vbkeyback",8},{"vbkeytab",9},{"vbkeyclear",12},
            {"vbkeyreturn",13},{"vbkeyshift",16},{"vbkeycontrol",17},{"vbkeymenu",18},{"vbkeypause",19},
            {"vbkeycapital",20},{"vbkeyescape",27},{"vbkeyspace",32},{"vbkeypageup",33},{"vbkeypagedown",34},
            {"vbkeyend",35},{"vbkeyhome",36},{"vbkeyleft",37},{"vbkeyup",38},{"vbkeyright",39},{"vbkeydown",40},
            {"vbkeyinsert",45},{"vbkeydelete",46},{"vbkeynumlock",144},{"vbkeyscrolllock",145}};
        if (const auto found=keyConstants.find(word);found!=keyConstants.end()) return VbaValue::numeric(found->second);
        if (word.starts_with("vbkey") && word.size()==6 && std::isalnum(static_cast<unsigned char>(word.back())))
            return VbaValue::numeric(std::toupper(static_cast<unsigned char>(word.back())));
        if (word.starts_with("vbkeyf")) {
            const auto digits=word.substr(6);
            if (!digits.empty() && std::all_of(digits.begin(),digits.end(),[](unsigned char ch) { return std::isdigit(ch); })) {
                const auto number=digits.size()<=2 ? std::stoi(digits) : 0;
                if (number>=1 && number<=24) return VbaValue::numeric(111+number);
            }
        }
        static const std::unordered_map<std::string,int> documentConstants{{"vbusecompareoption",-1},{"vbbinarycompare",0},{"vbtextcompare",1},{"vbdatabasecompare",2},{"accopy",2},{"accmdcopy",190},{"accmdpasteappend",38},{"acformbar",0},{"aceditmenu",1},{"acmenuver70",70},{"acmenuver20",20},{"acmenuver1x",11},{"acselectrecord",8},{"acdelete",6},{"accmdselectrecord",50},{"accmddeleterecord",223},{"acdataerrcontinue",0},{"acdataerrdisplay",1},{"acdeleteok",0},{"acdeletecancel",1},{"acdeleteusercancel",2},{"accmdsaverecord",97},{"accmdfind",30},{"accmdfindnext",341},{"accmdrecordsgotofirst",67},{"accmdrecordsgotoprevious",66},{"accmdrecordsgotonext",65},{"accmdrecordsgotolast",68},{"accmdrecordsgotonew",28},{"acdataform",2},{"acdatatable",0},{"acdataquery",1},{"acdatareport",3},{"acactivedataobject",-1},{"acprevious",0},{"acfirst",2},{"acgoto",4},{"acnext",1},{"aclast",3},{"acnewrec",5},{"acviewnormal",0},{"acviewpreview",2},{"acviewdesign",1},{"acwindownormal",0},{"acnormal",0},{"acdesign",1},{"acpreview",2},{"acformds",3},{"accustomcontrol",119},{"acshiftmask",1},{"acctrlmask",2},{"acaltmask",4},{"acdefault",-1},{"actable",0},{"acquery",1},{"acform",2},{"acreport",3},{"acmacro",4},{"acmodule",5},{"acsaveprompt",0},{"acsaveyes",1},{"acsaveno",2}};
        if (const auto constant=documentConstants.find(word);constant!=documentConstants.end()) return VbaValue::numeric(constant->second);
        if (word=="screen") return host_ ? host_->applicationObject()->member("Screen",{}) : VbaValue{};
        if (word=="docmd") return host_ ? host_->applicationObject()->member("DoCmd",{}) : VbaValue{};
        if (word=="codecontextobject") return host_ ? VbaValue::reference(host_->currentForm()) : VbaValue{};
        if (word == "application") return host_ ? VbaValue::reference(host_->applicationObject()) : VbaValue{};
        if (word=="commandbars") {
            std::vector<VbaValue> arguments;
            if (take("(")) { arguments.push_back(expression(1)); require(")"); }
            return host_ ? host_->applicationObject()->member("CommandBars",arguments) : VbaValue{};
        }
        if (word == "forms" || word == "reports") {
            std::vector<VbaValue> args;
            const bool item = take("(");
            if (item) { args.push_back(expression(1)); require(")"); }
            if (!host_) return {};
            const auto forms = word == "forms" ? host_->formsCollection() : host_->reportsCollection();
            return item ? forms->member("Item",args) : VbaValue::reference(forms);
        }
        if (word == "me" && !nextSymbol("!") && !nextSymbol(".")) {
            if (take("(")) {
                const auto index=expression(1); require(")");
                return host_ ? host_->currentForm()->member("Controls",{index}) : VbaValue{};
            }
            return host_ ? VbaValue::reference(host_->currentForm()) : VbaValue{};
        }
        if (word == "me" && (nextSymbol("!") || nextSymbol("."))) {
            const bool item = consume().value == "!";
            const auto field = consume();
            if (field.type != Token::Type::Word && field.type != Token::Type::Bracket)
                throw error("Me! benötigt einen Feldnamen");
            if (!item && lower(field.value) == "controls") {
                std::vector<VbaValue> args;
                if (take("(")) { args.push_back(expression(1)); require(")"); }
                return host_ ? host_->currentForm()->member("Controls",args) : VbaValue{};
            }
            const auto key = lower(field.value);
            if (!item && (key=="requery" || key=="recalc" || key=="undo" || key=="setfocus")) {
                if (take("(")) require(")");
                return host_ ? host_->currentForm()->member(field.value,{}) : VbaValue{};
            }
            if (!host_) return {};
            const bool formProperty = !item && (key == "name" || key == "dirty" || key == "newrecord" || key == "recordsource" || key == "caption" || key=="tag" || key=="filter" || key=="filteron" || key=="keypreview" || key=="bookmark" || key=="recordset");
            if (!item && !formProperty)
                if (auto value=host_->resolveIdentifier(field.value,objectContext || nextSymbol(".") || nextSymbol("!"))) return *value;
            if (!formProperty && objectContext)
                return item ? host_->currentForm()->objectItem(field.value) : host_->currentForm()->objectMember(field.value,{});
            if (!formProperty && (nextSymbol(".") || nextSymbol("!")))
                return host_->currentForm()->member("Controls",{VbaValue::string(field.value)});
            return item ? host_->controlValue(field.value) : host_->fieldValue(field.value);
        }
        if ((word=="now" || word=="date" || word=="time") && !nextSymbol("("))
            return host_ ? clockValue(word) : VbaValue{};
        if (accessExpression_ && (word=="sum" || word=="count") && take("(")) {
            const auto field=consume();
            if (field.type!=Token::Type::Word && field.type!=Token::Type::Bracket && field.value!="*")
                throw error("Aggregate field reference required");
            require(")");
            return host_ ? host_->expressionAggregate(word,field.value,field.value=="*") : VbaValue{};
        }
        if (nextSymbol("(")) {
            const auto variable=locals_.find(word);
            if (variable!=locals_.end() && variable->second.kind==VbaValue::Kind::Object) return variable->second;
        }
        if (take("(")) {
            std::vector<VbaValue> args;
            std::vector<std::string> references;
            std::vector<bool> omitted;
            if (!take(")")) {
                do {
                    const auto start = index_;
                    const bool missing=(word=="msgbox" || word=="inputbox" || word=="inputbox$" || word=="instr" || word=="strcomp") && (nextSymbol(",") || nextSymbol(")") || (accessExpression_ && nextSymbol(";")));
                    args.push_back(missing ? VbaValue::empty() : expression(1));
                    omitted.push_back(missing);
                    references.push_back(!missing && index_ == start + 1 && tokens_[start].type == Token::Type::Word
                        ? tokens_[start].value : "");
                } while (take(",") || (accessExpression_ && take(";")));
                require(")");
            }
            const auto one = args.size() == 1;
            if (word=="eval") {
                if (!host_) return {};
                if (args.size()!=1) throw error("Eval requires one argument",450);
                if (args[0].kind==VbaValue::Kind::Null) throw error("Invalid Null Eval argument",94);
                return host_->eval(args[0].asText());
            }
            if (word=="inputbox" || word=="inputbox$") {
                if (!host_) return {};
                if (args.empty() || args.size()>7 || omitted.front()) throw error("InputBox requires a prompt and at most seven arguments",450);
                for (std::size_t i=0;i<args.size();++i)
                    if (!omitted[i] && args[i].kind==VbaValue::Kind::Null) throw error("Invalid Null InputBox argument",94);
                for (std::size_t i=5;i<args.size();++i)
                    if (!omitted[i]) throw error("Windows Help arguments are unsupported",3251);
                VbaInputBoxArguments request;
                request.prompt=args[0].asText();
                if (args.size()>1 && !omitted[1]) request.title=args[1].asText();
                if (args.size()>2 && !omitted[2]) request.defaultValue=args[2].asText();
                const auto position=[&](std::size_t index) {
                    try { return integer(args[index]); }
                    catch (const VbaRuntimeError& failure) {
                        if (failure.number==5) throw error("InputBox position type mismatch",13);
                        throw;
                    }
                };
                if (args.size()>3 && !omitted[3]) request.xTwips=position(3);
                if (args.size()>4 && !omitted[4]) request.yTwips=position(4);
                return VbaValue::string(host_->inputBox(request));
            }
            if (word=="msgbox") {
                if (!host_) return {};
                if (args.empty() || args.size()>5 || omitted.front()) throw error("MsgBox requires a prompt and at most five arguments",450);
                if (args[0].kind==VbaValue::Kind::Null || (args.size()>1 && args[1].kind==VbaValue::Kind::Null) || (args.size()>2 && args[2].kind==VbaValue::Kind::Null))
                    throw error("Invalid Null MsgBox argument",94);
                for (std::size_t i=3;i<args.size();++i) if (!omitted[i]) throw error("Windows Help arguments are unsupported",3251);
                const auto buttons=args.size()>1 && !omitted[1] ? integer(args[1]) : 0;
                const auto title=args.size()>2 && !omitted[2] ? std::optional<std::string>(args[2].asText()) : std::nullopt;
                const auto result=host_->messageBox(args[0].asText(),buttons,title);
                if (result<1 || result>7) throw error("Invalid MsgBox result",5);
                return VbaValue::numeric(result);
            }

            if (word == "now" || word == "date" || word == "time") {
                if (!host_) return {};
                if (!args.empty()) throw error("Clock function takes no arguments",450);
                return clockValue(word);
            }
            if (word=="str" || word=="str$") {
                if (!host_) return {};
                if (args.size()!=1) throw error("Str requires one argument",450);
                const auto& value=args[0];
                if (value.kind==VbaValue::Kind::Null) throw error("Invalid Null Str argument",94);
                double number;
                try { number=value.asNumber(); }
                catch (const VbaRuntimeError& failure) { if (failure.number==5) throw error("Str type mismatch",13); throw; }
                if (!std::isfinite(number)) throw error("Str numeric overflow",6);
                auto result=value.kind==VbaValue::Kind::Currency ? value.asText() : VbaValue::numeric(number==0 ? 0 : number).asText();
                if (result.starts_with("0.")) result.erase(0,1);
                else if (result.starts_with("-0.")) result.erase(1,1);
                std::replace(result.begin(),result.end(),'e','E');
                if (number>=0) result.insert(result.begin(),' ');
                return VbaValue::string(std::move(result));
            }
            if (word=="isarray") {
                if (!host_) return {};
                if (args.size()!=1) throw error("IsArray requires one argument",450);
                return VbaValue::logical(args[0].kind==VbaValue::Kind::ByteArray);
            }
            if(word=="asc" || word=="ascw" || word=="chr" || word=="chrw") {
                if(!host_) return {};
                if(args.size()!=1) throw error("Character conversion requires one argument",450);
                const auto value=scalarObjectValue(args.front());
                if(value.kind==VbaValue::Kind::Null) throw error("Invalid Null character argument",94);
                if(word=="asc" || word=="ascw") {
                    const auto input=utf16String(value.asText());
                    if(input.empty()) throw error("Empty character argument",5);
                    const auto unit=static_cast<int>(input.front());
                    if(word=="asc" && unit>127) throw error("ANSI character codepage conversion is not implemented",3251);
                    return VbaValue::integer(static_cast<std::int16_t>(unit>=32768 ? unit-65536 : unit));
                }
                const auto code=static_cast<std::int32_t>(coerceVbaModuleValue(value,"long").number);
                if(code < -32768 || code>65535) throw error("Invalid character code",5);
                if(word=="chr" && (code<0 || code>127)) throw error("ANSI character codepage conversion is not implemented",3251);
                const auto unit=static_cast<char16_t>(code<0 ? code+65536 : code);
                return VbaValue::string(utf8String(std::u16string(1,unit)));
            }
            if (word=="cbyte" || word=="cint" || word=="clng" || word=="csng" || word=="cdbl" || word=="cbool") {
                if (!host_) return {};
                if (args.size()!=1) throw error("Conversion requires one argument",450);
                const auto type=word=="cbyte" ? "byte" : word=="cint" ? "integer" : word=="clng" ? "long" : word=="csng" ? "single" : word=="cdbl" ? "double" : "boolean";
                return coerceVbaModuleValue(scalarObjectValue(args.front()),type);
            }
            if (word=="ccur" || word=="vartype" || word=="typename") {
                if (!host_) return {};
                if (args.size()!=1) throw error("Conversion/type function requires one argument",450);
                const auto value=word=="vartype" ? typeInspectionValue(args.front()) : args.front();
                if (word=="ccur") return convertCurrency(value);
                const int type=value.kind==VbaValue::Kind::Empty ? 0 : value.kind==VbaValue::Kind::Null ? 1
                    : value.kind==VbaValue::Kind::Number ? (value.numericType==VbaValue::NumericType::Byte ? 17 : value.numericType==VbaValue::NumericType::Integer ? 2 : value.numericType==VbaValue::NumericType::Long ? 3 : value.numericType==VbaValue::NumericType::Single ? 4 : 5) : value.kind==VbaValue::Kind::Currency ? 6
                    : value.kind==VbaValue::Kind::Text ? 8 : value.kind==VbaValue::Kind::Object ? 9
                    : value.kind==VbaValue::Kind::Error ? 10 : value.kind==VbaValue::Kind::ByteArray ? 8209 : 11;
                if (word=="vartype") return VbaValue::integer(static_cast<std::int16_t>(type));
                if (type==9) {
                    if (!value.object) return VbaValue::string("Nothing");
                    throw error("Native object class name unavailable",3251);
                }
                return VbaValue::string(type==0 ? "Empty" : type==1 ? "Null" : type==5 ? "Double"
                    : type==17 ? "Byte" : type==2 ? "Integer" : type==3 ? "Long" : type==4 ? "Single" : type==6 ? "Currency" : type==8 ? "String" : type==9 ? "Object" : type==10 ? "Error" : type==8209 ? "Byte()" : "Boolean");
            }
            if (word=="dsum" || word=="dcount") {
                if (!host_) return {};
                if (args.size()<2 || args.size()>3) throw error("Domain aggregate requires two or three arguments",450);
                return host_->domainAggregate(word,args);
            }
            if (word=="iserror" || word=="cverr") {
                if (!host_) return {};
                if (args.size()!=1) throw error("Error function requires one argument",450);
                if (word=="iserror") return VbaValue::logical(args[0].kind==VbaValue::Kind::Error);
                if (args[0].kind==VbaValue::Kind::Null) throw error("Invalid Null CVErr argument",94);
                const auto number=integer(args[0]);
                if (number<0 || number>65535) throw error("CVErr number out of range");
                return VbaValue::errorValue(number);
            }
            if (word=="format" || word=="format$") {
                if (args.size()!=2) throw error("Format arguments unsupported",450);
                if (!host_) return {};
                if (args[1].kind==VbaValue::Kind::Null) throw error("Format mask is Null",94);
                return host_->formatValue(args[0],args[1].asText(),word=="format$");
            }
            if (word=="instr" || word=="strcomp") {
                if (!host_) return {};
                const bool search=word=="instr";
                if (args.size()<2 || args.size()>(search ? 4u : 3u)) throw error("String search/compare argument count invalid",450);
                const std::size_t first=search && args.size()>2 ? 1 : 0;
                if (omitted[first] || omitted[first+1]) throw error("String arguments required",450);
                std::vector<VbaValue> values; for(const auto& arg:args) values.push_back(scalarObjectValue(arg));
                std::int32_t start=1;
                if (first && !omitted[0]) start=integer(values[0]);
                if (start<1) throw error("InStr start must be positive",5);
                auto mode=compareMode_;
                const auto compare=first+2;
                if (values.size()>compare && !omitted[compare]) {
                    if (search && first && omitted[0]) throw error("InStr start required with compare",450);
                    const auto choice=integer(values[compare]);
                    if (choice>=0 && choice<=2) mode=static_cast<VbaCompareMode>(choice);
                    else if (choice!=-1) throw error(choice>2 && search ? "Locale-ID search comparison unsupported" : "Invalid compare setting",choice>2 && search ? 3251 : 5);
                }
                if (values[first].kind==VbaValue::Kind::Null || values[first+1].kind==VbaValue::Kind::Null) return {};
                const auto left=values[first].asText(),right=values[first+1].asText();
                if (!search) {
                    const int result=host_->compareStrings(left,right,mode);
                    return VbaValue::numeric((result>0)-(result<0));
                }
                const auto haystack=utf16String(left),needle=utf16String(right);
                if (haystack.empty()) return VbaValue::numeric(0);
                if (needle.empty()) return VbaValue::numeric(start);
                if (static_cast<std::size_t>(start)>haystack.size() || needle.size()>haystack.size()) return VbaValue::numeric(0);
                if (mode==VbaCompareMode::Binary) {
                    const auto found=haystack.find(needle,static_cast<std::size_t>(start-1));
                    return VbaValue::numeric(found==std::u16string::npos ? 0 : found+1);
                }
                const auto boundary=[&](std::size_t position) {
                    return position==0 || position==haystack.size() || !(haystack[position-1]>=0xd800 && haystack[position-1]<=0xdbff && haystack[position]>=0xdc00 && haystack[position]<=0xdfff);
                };
                for (std::size_t position=start-1;position+needle.size()<=haystack.size();++position) {
                    if (!boundary(position) || !boundary(position+needle.size())) continue;
                    const auto candidate=utf8String(std::u16string_view(haystack).substr(position,needle.size()));
                    if (host_->compareStrings(candidate,right,mode)==0) return VbaValue::numeric(position+1);
                }
                return VbaValue::numeric(0);
            }
            if (word=="trim" || word=="ltrim" || word=="rtrim" || word=="trim$" || word=="ltrim$" || word=="rtrim$") {
                if (!host_) return {};
                if (args.size()!=1) throw error("Trim function requires one argument",450);
                const auto value=scalarObjectValue(args[0]);
                if (value.kind==VbaValue::Kind::Null) {
                    if (word.ends_with('$')) throw error("Invalid Null String result",94);
                    return {};
                }
                const auto text=value.asText(); std::size_t first=0,last=text.size();
                if (word!="rtrim" && word!="rtrim$") while(first<last && text[first]==' ') ++first;
                if (word!="ltrim" && word!="ltrim$") while(last>first && text[last-1]==' ') --last;
                return VbaValue::string(text.substr(first,last-first));
            }
            if (word=="rgb" || word=="qbcolor") {
                if (!host_) return {};
                return word=="rgb" ? vbaRgb(args) : vbaQbColor(args);
            }
            if (word=="abs" || word=="int" || word=="fix" || word=="sgn") {
                if (!host_) return {};
                if (args.size()!=1) throw error("Numeric function requires one argument",450);
                const auto value=scalarObjectValue(args.front());
                if (value.kind==VbaValue::Kind::Null) {
                    if(word=="sgn") throw error("Invalid Null sign argument",94);
                    return {};
                }
                if (word=="sgn") {
                    const auto number=coerceVbaModuleValue(value,"double").number;
                    return VbaValue::integer(static_cast<std::int16_t>((number>0)-(number<0)));
                }
                if (value.kind==VbaValue::Kind::Currency) {
                    if(word=="abs") return value.currencyUnits<0 ? negateCurrency(value) : value;
                    if(word=="int") return floorCurrency(value);
                    return VbaValue::currency((value.currencyUnits/10000)*10000);
                }
                const auto number=coerceVbaModuleValue(value,"double").number;
                const auto answer=word=="abs" ? std::abs(number) : word=="int" ? std::floor(number) : std::trunc(number);
                const auto result=VbaValue::numeric(answer==0 ? 0 : answer);
                if (value.kind==VbaValue::Kind::Boolean || value.kind==VbaValue::Kind::Empty)
                    return coerceVbaModuleValue(result,"integer");
                if (value.kind==VbaValue::Kind::Number) {
                    if(value.numericType==VbaValue::NumericType::Byte) return coerceVbaModuleValue(result,"byte");
                    if(value.numericType==VbaValue::NumericType::Integer) return coerceVbaModuleValue(result,"integer");
                    if(value.numericType==VbaValue::NumericType::Long) return coerceVbaModuleValue(result,"long");
                    if(value.numericType==VbaValue::NumericType::Single) return coerceVbaModuleValue(result,"single");
                }
                return result;
            }
            if ((word == "isnull" || word == "isempty" || word == "len" || word == "cstr"
                || word == "ucase"
                || word == "lcase"
                || word == "isnumeric" || word == "val") && one) {
                if (!host_) return {};
                const auto& value = args[0];
                if (word == "isnull") return VbaValue::logical(value.kind == VbaValue::Kind::Null);
                if (word == "isempty") return VbaValue::logical(value.kind == VbaValue::Kind::Empty);
                if (word == "len") {
                    const auto scalar=scalarObjectValue(value);
                    return scalar.kind==VbaValue::Kind::Null ? VbaValue{} : VbaValue::numeric(vbaStringCodeUnits(scalar.asText()));
                }
                if (word == "cstr") {
                    if (value.kind==VbaValue::Kind::Error) return VbaValue::string("Error "+std::to_string(value.errorNumber));
                    if (value.kind == VbaValue::Kind::Null) throw error("VBA: Ungültige Verwendung von Null", 94);
                    return VbaValue::string(value.asText());
                }
                if (word == "isnumeric") {
                    try { (void)value.asNumber(); return VbaValue::logical(true); }
                    catch (...) { return VbaValue::logical(false); }
                }
                if (word == "val") {
                    const auto str = value.asText();
                    std::size_t index = 0;
                    try { return VbaValue::numeric(std::stod(str, &index)); }
                    catch (...) { return VbaValue::numeric(0); }
                }
                auto result = value.asText();
                std::transform(result.begin(), result.end(), result.begin(), [&](unsigned char ch) {
                    return word == "ucase" ? std::toupper(ch) : std::tolower(ch);
                });
                return VbaValue::string(result);
            }
            if (word == "nz" && (args.size() == 1 || args.size() == 2)) {
                if (!host_) return {};
                return args[0].kind == VbaValue::Kind::Null
                    ? (args.size() == 2 ? args[1] : VbaValue::string("")) : args[0];
            }
            if (word == "iif" && args.size() == 3) {
                if (!host_) return {};
                return args[0].truthy() ? args[1] : args[2];
            }
            if (word=="left" || word=="right" || word=="left$" || word=="right$" || word=="mid" || word=="mid$") {
                if (!host_) return {};
                const bool middle=word=="mid" || word=="mid$";
                if (args.size()<2 || args.size()>(middle ? 3u : 2u)) throw error("Substring argument count invalid",450);
                const auto value=scalarObjectValue(args[0]);
                const auto position=integer(args[1]);
                const auto length=middle && args.size()==3 ? integer(args[2]) : middle ? std::numeric_limits<std::int32_t>::max() : position;
                if (length<0 || (middle && position<1)) throw error("Invalid substring position or length",5);
                if (value.kind==VbaValue::Kind::Null) {
                    if (word.ends_with('$')) throw error("Invalid Null String result",94);
                    return {};
                }
                const auto input=utf16String(value.asText());
                const auto count=std::min<std::size_t>(input.size(),static_cast<std::size_t>(length));
                const auto start=middle ? std::min<std::size_t>(input.size(),static_cast<std::size_t>(position-1))
                    : word=="right" || word=="right$" ? input.size()-count : 0;
                return VbaValue::string(utf8String(std::u16string_view(input).substr(start,count)));
            }
            if (!host_) return {}; // syntactic function call; bind at execution
            if (resolver_) return resolver_(tok.value, args, references);
            throw error("Nicht unterstützte VBA-Funktion/Parameter: " + tok.value);
        }
        if (!host_) return {}; // Validate identifiers syntactically, resolve on execution.
        const auto variable = locals_.find(word);
        if (variable==locals_.end() && accessExpression_)
            return host_->expressionValue(tok.value,objectContext || nextSymbol(".") || nextSymbol("!"));
        if (variable==locals_.end()) {
            if (auto value=host_->resolveIdentifier(tok.value,objectContext || nextSymbol(".") || nextSymbol("!"))) return *value;
            throw error("VBA-Variable nicht deklariert: " + tok.value);
        }
        return variable->second;
    }
    std::vector<Token> tokens_;
    std::size_t index_{};
    VbaHost* host_{};
    const VbaVariables& locals_;
    VbaFunctionResolver resolver_;
    bool objectContext_{};
    bool accessExpression_{};
    const VbaValue* withContext_{};
    bool allowWith_{};
    VbaCompareMode compareMode_;
};
} // namespace

VbaValue VbaHost::formatValue(const VbaValue&,std::string_view,bool) {
    throw VbaRuntimeError(438,"Format host adapter unavailable");
}
int VbaHost::messageBox(std::string_view prompt,int buttons,const std::optional<std::string>& title) {
    if (buttons!=0 || title.has_value()) throw VbaRuntimeError(3251,"Styled MsgBox requires a dialog host");
    message(prompt); return 1;
}
std::string VbaHost::inputBox(const VbaInputBoxArguments&) {
    throw VbaRuntimeError(3251,"InputBox requires a native dialog host");
}
VbaValue VbaHost::eval(std::string_view source) {
    static thread_local int depth=0;
    if (depth>=16) throw VbaRuntimeError(28,"Eval nesting limit exceeded");
    struct Guard { int& depth; explicit Guard(int& value):depth(value) { ++depth; } ~Guard() { --depth; } } guard(depth);
    const VbaFunctionResolver functions=[&](std::string_view name,std::vector<VbaValue>& arguments,const std::vector<std::string>&) {
        const auto module=moduleForProcedure(name);
        if (!module) throw VbaRuntimeError(2425,"Eval function unavailable");
        const auto analysis=analyzeVba(*module);
        bool publicFunction=false;
        for (const auto& procedure:analysis.procedures) {
            if (lower(procedure.name)!=lower(std::string(name)) || lower(procedure.kind)!="function") continue;
            std::istringstream lines(*module); std::string declaration;
            for (int line=0;line<procedure.line && std::getline(lines,declaration);++line) {}
            const auto start=declaration.find_first_not_of(" \t\r");
            const auto header=start==std::string::npos ? std::string{} : lower(declaration.substr(start));
            publicFunction=!header.starts_with("private ") && !header.starts_with("private\t");
        }
        if (!publicFunction) throw VbaRuntimeError(2425,"Eval requires a public function");
        const auto result=executeVbaProcedureWithArguments(*module,name,*this,arguments);
        if (result.status!=VbaExecutionStatus::Completed || !result.typedReturnValue)
            throw VbaRuntimeError(result.errorNumber ? result.errorNumber : 2425,"Eval function execution failed");
        return *result.typedReturnValue;
    };
    const VbaVariables variables;
    auto value=scalarVbaValue(ExpressionParser(source,this,variables,functions,false,true,nullptr,false,VbaCompareMode::Database).parse());
    if (value.kind==VbaValue::Kind::Error) throw VbaRuntimeError(value.errorNumber ? value.errorNumber : 13,"Eval expression failed");
    if (value.kind==VbaValue::Kind::ByteArray) throw VbaRuntimeError(13,"Eval cannot return a byte array");
    return value;
}
std::optional<VbaValue> VbaHost::resolveIdentifier(std::string_view,bool) { return std::nullopt; }
bool VbaHost::assignIdentifier(std::string_view,const VbaValue&) { return false; }
VbaValue VbaHost::bracketValue(std::string_view,bool) { throw VbaRuntimeError(438,"Access bracket evaluation context unavailable"); }
VbaValue VbaHost::expressionValue(std::string_view,bool) { throw VbaRuntimeError(438,"Access expression context unavailable"); }
VbaValue VbaHost::domainAggregate(std::string_view,const std::vector<VbaValue>&) { throw VbaRuntimeError(438,"Access domain context unavailable"); }
VbaValue VbaHost::expressionAggregate(std::string_view,std::string_view,bool) { throw VbaRuntimeError(438,"Access aggregate context unavailable"); }
VbaValue VbaHost::controlValue(std::string_view name) { return fieldValue(name); }
void VbaHost::setControlValue(std::string_view name,const VbaValue& value) { setFieldValue(name,value); }
VbaValue VbaHost::fieldValue(std::string_view name) {
    const auto value = field(name);
    return value ? VbaValue::string(*value) : VbaValue{};
}
void VbaHost::setFieldValue(std::string_view name,const VbaValue& value) { setField(name,value.asField()); }
VbaValue VbaValue::empty() { VbaValue v; v.kind = Kind::Empty; return v; }
VbaValue VbaValue::errorValue(int number) { VbaValue v; v.kind=Kind::Error; v.errorNumber=number; return v; }
VbaValue VbaValue::reference(std::shared_ptr<VbaObject> value) { VbaValue v; v.kind = Kind::Object; v.object = std::move(value); return v; }
VbaValue VbaValue::numeric(double value) { VbaValue v; v.kind = Kind::Number; v.number = value; return v; }
VbaValue VbaValue::byte(unsigned char value) { auto v=numeric(value); v.numericType=NumericType::Byte; return v; }
VbaValue VbaValue::single(float value) { auto v=numeric(value); v.numericType=NumericType::Single; return v; }
VbaValue VbaValue::integer(std::int16_t value) { auto v=numeric(value); v.numericType=NumericType::Integer; return v; }
VbaValue VbaValue::longInteger(std::int32_t value) { auto v=numeric(value); v.numericType=NumericType::Long; return v; }
VbaValue VbaValue::string(std::string value) { VbaValue v; v.kind = Kind::Text; v.text = std::move(value); return v; }
VbaValue VbaValue::byteArray(std::vector<unsigned char> value) { VbaValue v; v.kind=Kind::ByteArray; v.bytes=std::move(value); return v; }
VbaValue VbaValue::currency(std::int64_t units) { VbaValue v; v.kind=Kind::Currency; v.currencyUnits=units; return v; }
VbaValue VbaValue::logical(bool value) { VbaValue v; v.kind = Kind::Boolean; v.boolean = value; return v; }
namespace {
VbaValue typeInspectionValue(VbaValue value) {
    for (int depth=0;;++depth) {
        if(value.kind!=VbaValue::Kind::Object || !value.object || !value.object->hasScalarDefault()) return value;
        if(depth>=16) throw VbaRuntimeError(28,"Recursive object default value");
        value=value.object->scalarValue();
    }
}
VbaValue scalarObjectValue(VbaValue value) {
    for (int depth=0;depth<16;++depth) {
        if (value.kind!=VbaValue::Kind::Object) return value;
        if (!value.object) throw VbaRuntimeError(91,"Object variable is Nothing");
        value=value.object->scalarValue();
    }
    throw VbaRuntimeError(28,"Recursive object default value");
}
}
VbaValue vbaRgb(const std::vector<VbaValue>& arguments) {
    if (arguments.size()!=3) throw error("RGB requires three color channels",450);
    std::uint32_t color=0;
    for (std::size_t index=0;index<3;++index) {
        const auto value=scalarObjectValue(arguments[index]);
        const auto channel=bankersRound(value.asNumber());
        if (!std::isfinite(channel)) throw error("RGB channel overflow",6);
        if (channel<0) throw error("Invalid negative RGB channel",5);
        color|=static_cast<std::uint32_t>(std::min(channel,255.0))<<(index*8);
    }
    return VbaValue::numeric(color);
}

VbaValue vbaQbColor(const std::vector<VbaValue>& arguments) {
    if (arguments.size()!=1) throw error("QBColor requires one palette index",450);
    const auto index=integer(scalarObjectValue(arguments[0]));
    if (index<0 || index>15) throw error("Invalid QBColor palette index",5);
    constexpr std::uint32_t palette[]={0x000000,0x800000,0x008000,0x808000,
        0x000080,0x800080,0x008080,0xc0c0c0,0x808080,0xff0000,0x00ff00,
        0xffff00,0x0000ff,0xff00ff,0x00ffff,0xffffff};
    return VbaValue::numeric(palette[index]);
}

int VbaHost::compareStrings(std::string_view left,std::string_view right,VbaCompareMode mode) {
    if (mode==VbaCompareMode::Binary) return utf16String(left).compare(utf16String(right));
    const auto fold=[](std::string_view text) {
        auto result=std::string(text);
        for (auto& ch:result) if (ch>='A' && ch<='Z') ch=static_cast<char>(ch-'A'+'a');
        return result;
    };
    const auto a=fold(left),b=fold(right);
    if (a==b) return 0;
    const auto nonAscii=[](std::string_view text) { return std::any_of(text.begin(),text.end(),[](unsigned char ch) { return ch>=128; }); };
    if (nonAscii(a) || nonAscii(b)) throw VbaRuntimeError(3251,"Non-ASCII text comparison requires a locale-aware host");
    return a.compare(b);
}
std::size_t vbaStringCodeUnits(std::string_view text) { return utf16String(text).size(); }
VbaValue scalarVbaValue(VbaValue value) { return scalarObjectValue(std::move(value)); }
std::string VbaValue::asText() const {
    if (kind==Kind::ByteArray) throw error("Byte array cannot be implicitly converted to text",13);
    if (kind==Kind::Error) throw error("Error Variant cannot be implicitly converted to text",13);
    if (kind==Kind::Object) return scalarObjectValue(*this).asText();
    if (kind == Kind::Text) return text;
    if (kind==Kind::Currency) {
        const auto magnitude=currencyUnits<0 ? static_cast<std::uint64_t>(-(currencyUnits+1))+1 : static_cast<std::uint64_t>(currencyUnits);
        auto whole=std::to_string(magnitude/10000);
        if (magnitude%10000) {
            auto fraction=std::to_string(10000+magnitude%10000).substr(1);
            while (fraction.back()=='0') fraction.pop_back();
            whole+='.'+fraction;
        }
        return (currencyUnits<0 ? "-" : "")+whole;
    }
    if (kind == Kind::Boolean) return boolean ? "True" : "False";
    if (kind == Kind::Null || kind == Kind::Empty) return {};
    std::ostringstream stream;
    stream << std::setprecision(numericType==NumericType::Single ? 7 : 15) << number;
    return stream.str();
}
double VbaValue::asNumber() const {
    if (kind==Kind::ByteArray) throw error("Byte array cannot be implicitly converted to number",13);
    if (kind==Kind::Error) throw error("Error Variant cannot be implicitly converted to number",13);
    if (kind==Kind::Object) return scalarObjectValue(*this).asNumber();
    if (kind == Kind::Empty) return 0;
    if (kind == Kind::Number) return number;
    if (kind == Kind::Currency) return static_cast<double>(currencyUnits)/10000;
    if (kind == Kind::Boolean) return boolean ? -1.0 : 0.0; // VBA True = -1.
    if (kind == Kind::Null) throw error("VBA: Null kann nicht in Zahl umgewandelt werden",94);
    std::size_t consumed = 0;
    try {
        const auto result = std::stod(text, &consumed);
        if (consumed == text.size() && std::isfinite(result)) return result;
    } catch (...) {}
    throw error("VBA: Zahl erwartet: " + text,13);
}
bool VbaValue::truthy() const {
    if (kind == Kind::Null) return false; // documented reduced semantics
    if (kind == Kind::Boolean) return boolean;
    return asNumber() != 0;
}
std::optional<std::string> VbaValue::asField() const {
    return kind == Kind::Null ? std::nullopt : std::optional<std::string>{asText()};
}
VbaValue compareVbaValues(std::string_view relation,const VbaValue& left,const VbaValue& right,
                         VbaHost& host,VbaCompareMode mode) {
    if (relation!="=" && relation!="<>" && relation!="<" && relation!="<=" && relation!=">" && relation!=">=")
        throw VbaRuntimeError(5,"Invalid relational operator");
    return binary(std::string(relation),left,right,&host,mode);
}
bool validVbaExpression(std::string_view text,bool withContext) {
    const VbaVariables vars;
    try { (void)ExpressionParser(text, nullptr, vars,{},false,false,nullptr,withContext).parse(); return true; }
    catch (...) { return false; }
}
VbaValue evaluateVbaExpression(std::string_view text, VbaHost& host,
                               const VbaVariables& vars, const VbaFunctionResolver& resolver,const VbaValue* withContext,VbaCompareMode compareMode) {
    return ExpressionParser(text, &host, vars, resolver,false,false,withContext,false,compareMode).parse();
}
VbaValue evaluateVbaObjectExpression(std::string_view text,VbaHost& host,
                                    const VbaVariables& vars,const VbaFunctionResolver& resolver,const VbaValue* withContext,VbaCompareMode compareMode) {
    return ExpressionParser(text,&host,vars,resolver,true,false,withContext,false,compareMode).parse();
}
VbaValue evaluateAccessExpression(std::string_view text,VbaHost& host) {
    const VbaVariables variables;
    return ExpressionParser(text,&host,variables,{},false,true).parse();
}
}
