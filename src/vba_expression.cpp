#include "vba_expression.hpp"

#include <algorithm>
#include <chrono>
#include <cmath>
#include <cctype>
#include <ctime>
#include <iomanip>
#include <cstdint>
#include <limits>
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

struct Token {
    enum class Type { Word, Number, Text, Bracket, Symbol, End } type{Type::End};
    std::string value;
};

std::vector<Token> tokenize(std::string_view input) {
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
        } else if (std::isdigit(static_cast<unsigned char>(ch))) {
            const auto start = pos++;
            while (pos < input.size() &&
                   (std::isdigit(static_cast<unsigned char>(input[pos])) || input[pos] == '.')) ++pos;
            tokens.push_back({Token::Type::Number, std::string(input.substr(start, pos - start))});
        } else if (ch == '"') {
            ++pos;
            std::string value;
            bool closed = false;
            while (pos < input.size()) {
                if (input[pos] == '"') {
                    ++pos;
                    if (pos < input.size() && input[pos] == '"') {
                        value.push_back('"'); ++pos;
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
            if (std::string("+-*/&=<>!.,()").find(ch) == std::string::npos)
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
    const auto text = lower(token.value);
    if (text == "or") return 1;
    if (text == "and") return 2;
    if (text == "=" || text == "<>" || text == "<" || text == "<=" ||
        text == ">" || text == ">=" || text == "is") return 3;
    if (text == "&") return 4;
    if (text == "+" || text == "-") return 5;
    if (text == "*" || text == "/" || text == "mod") return 6;
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
VbaValue binary(const std::string& op, const VbaValue& left, const VbaValue& right) {
    if (op == "is") {
        if (left.kind != VbaValue::Kind::Object || right.kind != VbaValue::Kind::Object)
            throw error("VBA: Is benötigt Objekte");
        return VbaValue::logical(left.object == right.object);
    }
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
    if (left.kind == VbaValue::Kind::Empty && right.kind == VbaValue::Kind::Text)
        return binary(op, VbaValue::string(""), right);
    if (right.kind == VbaValue::Kind::Empty && left.kind == VbaValue::Kind::Text)
        return binary(op, left, VbaValue::string(""));
    if (left.kind == VbaValue::Kind::Null || right.kind == VbaValue::Kind::Null)
        return {}; // Simplified three-valued logic: NULL comparisons yield NULL.

    if (op == "=" || op == "<>" || op == "<" || op == "<=" || op == ">" || op == ">=") {
        int cmp = 0;
        if (left.kind == VbaValue::Kind::Text && right.kind == VbaValue::Kind::Text) {
            cmp = left.text.compare(right.text);
        } else {
            const auto a = left.asNumber(), b = right.asNumber();
            cmp = (a > b) - (a < b);
        }
        return VbaValue::logical(op == "=" ? cmp == 0 : op == "<>" ? cmp != 0 :
            op == "<" ? cmp < 0 : op == "<=" ? cmp <= 0 : op == ">" ? cmp > 0 : cmp >= 0);
    }
    if (op == "+" && left.kind == VbaValue::Kind::Text && right.kind == VbaValue::Kind::Text)
        return VbaValue::string(left.text + right.text);
    const auto a = left.asNumber(), b = right.asNumber();
    double answer = 0;
    if (op == "+") answer = a + b;
    else if (op == "-") answer = a - b;
    else if (op == "*") answer = a * b;
    else if (op == "/") {
        if (b == 0) throw error("VBA: Division durch Null", 11);
        answer = a / b;
    } else if (op == "mod") {
        if (b == 0) throw error("VBA: Modulo durch Null", 11);
        answer = std::fmod(a, b);
    } else throw error("Nicht unterstützter Operator " + op);
    if (!std::isfinite(answer)) throw error("VBA: Ungültiges numerisches Ergebnis");
    return VbaValue::numeric(answer);
}

class ExpressionParser {
public:
    ExpressionParser(std::string_view source, VbaHost* host, const VbaVariables& locals,
                     const VbaFunctionResolver& resolver = {})
        : tokens_(tokenize(source)), host_(host), locals_(locals), resolver_(resolver) {}

    VbaValue parse() {
        auto value = expression(1);
        if (peek().type != Token::Type::End)
            throw error("Zusätzliche Zeichen im VBA-Ausdruck: " + peek().value);
        return value;
    }
private:
    const Token& peek() const { return tokens_[index_]; }
    Token consume() { return tokens_[index_++]; }
    bool take(const std::string& wanted) {
        if (lower(peek().value) != wanted) return false;
        consume();
        return true;
    }
    void require(const std::string& wanted) {
        if (!take(wanted)) throw error("Erwartet: " + wanted);
    }
    VbaValue expression(int minPrecedence) {
        auto lhs = primary();
        while (true) {
            const int prec = precedence(peek());
            if (prec < minPrecedence) break;
            const auto op = lower(consume().value);
            auto rhs = expression(prec + 1);
            lhs = host_ ? binary(op, lhs, rhs) : VbaValue{};
        }
        return lhs;
    }
    VbaValue primary() {
        auto value = atom();
        while (peek().value == "." || peek().value == "!") {
            const bool item = consume().value == "!";
            const auto member = consume();
            if (member.type != Token::Type::Word && member.type != Token::Type::Bracket)
                throw error("Objektmitglied erwartet");
            std::vector<VbaValue> args;
            if (take("(")) {
                if (!take(")")) {
                    do { args.push_back(expression(1)); } while (take(","));
                    require(")");
                }
            }
            if (host_) {
                if (value.kind != VbaValue::Kind::Object || !value.object)
                    throw error("Objektvariable nicht gesetzt", 91);
                if (item && !args.empty()) throw error("Bang lookup arguments unsupported",450);
                value = item ? value.object->item(member.value) : value.object->member(member.value,args);
            }
        }
        return value;
    }
    VbaValue atom() {
        if (take("-")) {
            const auto arg = expression(7);
            return host_ ? VbaValue::numeric(-arg.asNumber()) : VbaValue{};
        }
        if (take("+")) {
            const auto arg = expression(7);
            return host_ ? VbaValue::numeric(arg.asNumber()) : VbaValue{};
        }
        if (take("not")) {
            const auto arg = expression(3); // comparisons bind before Not in VBA
            if (!host_ || arg.kind == VbaValue::Kind::Null) return {};
            if (arg.kind == VbaValue::Kind::Boolean) return VbaValue::logical(!arg.boolean);
            return VbaValue::numeric(~integer(arg));
        }
        if (take("(")) {
            const auto value = expression(1);
            require(")");
            return value;
        }
        const auto tok = consume();
        if (tok.type == Token::Type::Text) return VbaValue::string(tok.value);
        if (tok.type == Token::Type::Number) {
            std::size_t consumed = 0;
            const auto number = std::stod(tok.value, &consumed);
            if (consumed != tok.value.size()) throw error("Ungültige VBA-Zahl");
            return VbaValue::numeric(number);
        }
        if (tok.type != Token::Type::Word)
            throw error("Ungültiger VBA-Ausdruck: " + tok.value);
        const auto word = lower(tok.value);
        if (word == "null") return {};
        if (word == "empty") return VbaValue::empty();
        if (word == "nothing") return VbaValue::reference(nullptr);
        if (word == "currentdb") {
            if (take("(")) require(")");
            return host_ ? VbaValue::reference(host_->currentDatabase()) : VbaValue{};
        }
        static const std::unordered_map<std::string,int> controlTypes{
            {"aclabel",100},{"acrectangle",101},{"acline",102},{"accommandbutton",104},
            {"acoptionbutton",105},{"accheckbox",106},{"actextbox",109},{"aclistbox",110},
            {"accombobox",111},{"acsubform",112},{"actogglebutton",122}};
        if (const auto type = controlTypes.find(word); type != controlTypes.end()) return VbaValue::numeric(type->second);
        if (word == "dbopendynaset") return VbaValue::numeric(2);
        if (word == "dbopensnapshot") return VbaValue::numeric(4);
        if (word == "dbfailonerror") return VbaValue::numeric(128);
        if (word == "true") return VbaValue::logical(true);
        if (word == "false") return VbaValue::logical(false);
        if (word == "vbcrlf") return VbaValue::string("\r\n");
        if (word == "vbnewline") return VbaValue::string("\n");
        if (word == "vbtab") return VbaValue::string("\t");
        if (word == "application") return host_ ? VbaValue::reference(host_->applicationObject()) : VbaValue{};
        if (word == "forms" || word == "reports") {
            std::vector<VbaValue> args;
            const bool item = take("(");
            if (item) { args.push_back(expression(1)); require(")"); }
            if (!host_) return {};
            const auto forms = word == "forms" ? host_->formsCollection() : host_->reportsCollection();
            return item ? forms->member("Item",args) : VbaValue::reference(forms);
        }
        if (word == "me" && peek().value != "!" && peek().value != ".")
            return host_ ? VbaValue::reference(host_->currentForm()) : VbaValue{};
        if (word == "me" && (peek().value == "!" || peek().value == ".")) {
            const bool item = consume().value == "!";
            const auto field = consume();
            if (field.type != Token::Type::Word && field.type != Token::Type::Bracket)
                throw error("Me! benötigt einen Feldnamen");
            if (!item && lower(field.value) == "controls") {
                std::vector<VbaValue> args;
                if (take("(")) { args.push_back(expression(1)); require(")"); }
                return host_ ? host_->currentForm()->member("Controls",args) : VbaValue{};
            }
            if (!host_) return {};
            const auto key = lower(field.value);
            const bool formProperty = !item && (key == "name" || key == "dirty" || key == "newrecord" || key == "recordsource" || key == "caption");
            if (!formProperty && (peek().value == "." || peek().value == "!"))
                return host_->currentForm()->member("Controls",{VbaValue::string(field.value)});
            return item ? host_->controlValue(field.value) : host_->fieldValue(field.value);
        }
        if (take("(")) {
            std::vector<VbaValue> args;
            std::vector<std::string> references;
            if (!take(")")) {
                do {
                    const auto start = index_;
                    args.push_back(expression(1));
                    references.push_back(index_ == start + 1 && tokens_[start].type == Token::Type::Word
                        ? tokens_[start].value : "");
                } while (take(","));
                require(")");
            }
            const auto one = args.size() == 1;
            if ((word == "now" || word == "date" || word == "time") && args.empty()) {
                if (!host_) return {};
                const auto now = std::chrono::system_clock::to_time_t(std::chrono::system_clock::now());
                const auto* local = std::localtime(&now);
                if (!local) throw error("Systemdatum nicht verfügbar");
                std::ostringstream formatted;
                formatted << std::put_time(local, word == "now" ? "%Y-%m-%d %H:%M:%S"
                    : word == "date" ? "%Y-%m-%d" : "%H:%M:%S");
                return VbaValue::string(formatted.str());
            }
            if (word=="format" || word=="format$") {
                if (args.size()!=2) throw error("Format arguments unsupported",450);
                if (!host_) return {};
                if (args[1].kind==VbaValue::Kind::Null) throw error("Format mask is Null",94);
                return host_->formatValue(args[0],args[1].asText(),word=="format$");
            }
            if ((word == "isnull" || word == "isempty" || word == "len" || word == "cstr"
                || word == "clng" || word == "cdbl" || word == "ucase"
                || word == "lcase" || word == "abs" || word == "int"
                || word == "trim" || word == "isnumeric" || word == "val") && one) {
                if (!host_) return {};
                const auto& value = args[0];
                if (word == "isnull") return VbaValue::logical(value.kind == VbaValue::Kind::Null);
                if (word == "isempty") return VbaValue::logical(value.kind == VbaValue::Kind::Empty);
                if (word == "len") return value.kind == VbaValue::Kind::Null ? VbaValue{} : VbaValue::numeric(value.asText().size());
                if (word == "cstr") {
                    if (value.kind == VbaValue::Kind::Null) throw error("VBA: Ungültige Verwendung von Null", 94);
                    return VbaValue::string(value.asText());
                }
                if (word == "clng") return VbaValue::numeric(integer(value));
                if (word == "cdbl") return VbaValue::numeric(value.asNumber());
                if (word == "abs") return VbaValue::numeric(std::abs(value.asNumber()));
                if (word == "int") return VbaValue::numeric(std::floor(value.asNumber()));
                if (word == "trim") return VbaValue::string([&] {
                    const auto str = value.asText();
                    const auto first = str.find_first_not_of(" \t\n\r");
                    if (first == std::string::npos) return std::string{};
                    const auto last = str.find_last_not_of(" \t\n\r");
                    return str.substr(first, last - first + 1);
                }());
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
            if ((word == "left" || word == "right") && args.size() == 2) {
                if (!host_) return {};
                const auto input = args[0].asText();
                const auto length = static_cast<long>(args[1].asNumber());
                if (length < 0) throw error("VBA: Negative Teilstringlänge");
                const auto count = std::min<std::size_t>(input.size(), static_cast<std::size_t>(length));
                return VbaValue::string(word == "left" ? input.substr(0, count)
                                                        : input.substr(input.size() - count));
            }
            if (word == "mid" && (args.size() == 2 || args.size() == 3)) {
                if (!host_) return {};
                const auto input = args[0].asText();
                const auto start = static_cast<long>(args[1].asNumber());
                if (start < 1) throw error("VBA: Mid-Position muss >= 1 sein");
                if (args.size() == 2) return VbaValue::string(input.substr(
                    std::min<std::size_t>(input.size(), static_cast<std::size_t>(start - 1))));
                const auto length = static_cast<long>(args[2].asNumber());
                if (length < 0) throw error("VBA: Negative Teilstringlänge");
                return VbaValue::string(input.substr(
                    std::min<std::size_t>(input.size(), static_cast<std::size_t>(start - 1)),
                    static_cast<std::size_t>(length)));
            }
            if (!host_) return {}; // syntactic function call; bind at execution
            if (resolver_) return resolver_(tok.value, args, references);
            throw error("Nicht unterstützte VBA-Funktion/Parameter: " + tok.value);
        }
        if (!host_) return {}; // Validate identifiers syntactically, resolve on execution.
        const auto variable = locals_.find(word);
        if (variable == locals_.end()) throw error("VBA-Variable nicht deklariert: " + tok.value);
        return variable->second;
    }
    std::vector<Token> tokens_;
    std::size_t index_{};
    VbaHost* host_{};
    const VbaVariables& locals_;
    VbaFunctionResolver resolver_;
};
} // namespace

VbaValue VbaHost::formatValue(const VbaValue&,std::string_view,bool) {
    throw VbaRuntimeError(438,"Format host adapter unavailable");
}
VbaValue VbaHost::controlValue(std::string_view name) { return fieldValue(name); }
void VbaHost::setControlValue(std::string_view name,const VbaValue& value) { setFieldValue(name,value); }
VbaValue VbaHost::fieldValue(std::string_view name) {
    const auto value = field(name);
    return value ? VbaValue::string(*value) : VbaValue{};
}
void VbaHost::setFieldValue(std::string_view name,const VbaValue& value) { setField(name,value.asField()); }
VbaValue VbaValue::empty() { VbaValue v; v.kind = Kind::Empty; return v; }
VbaValue VbaValue::reference(std::shared_ptr<VbaObject> value) { VbaValue v; v.kind = Kind::Object; v.object = std::move(value); return v; }
VbaValue VbaValue::numeric(double value) { VbaValue v; v.kind = Kind::Number; v.number = value; return v; }
VbaValue VbaValue::string(std::string value) { VbaValue v; v.kind = Kind::Text; v.text = std::move(value); return v; }
VbaValue VbaValue::logical(bool value) { VbaValue v; v.kind = Kind::Boolean; v.boolean = value; return v; }
std::string VbaValue::asText() const {
    if (kind == Kind::Object) throw error("VBA: Objekt kann nicht in Text umgewandelt werden");
    if (kind == Kind::Text) return text;
    if (kind == Kind::Boolean) return boolean ? "True" : "False";
    if (kind == Kind::Null || kind == Kind::Empty) return {};
    std::ostringstream stream;
    stream << std::setprecision(15) << number;
    return stream.str();
}
double VbaValue::asNumber() const {
    if (kind == Kind::Object) throw error("VBA: Objekt kann nicht in Zahl umgewandelt werden");
    if (kind == Kind::Empty) return 0;
    if (kind == Kind::Number) return number;
    if (kind == Kind::Boolean) return boolean ? -1.0 : 0.0; // VBA True = -1.
    if (kind == Kind::Null) throw error("VBA: Null kann nicht in Zahl umgewandelt werden");
    std::size_t consumed = 0;
    try {
        const auto result = std::stod(text, &consumed);
        if (consumed == text.size() && std::isfinite(result)) return result;
    } catch (...) {}
    throw error("VBA: Zahl erwartet: " + text);
}
bool VbaValue::truthy() const {
    if (kind == Kind::Null) return false; // documented reduced semantics
    if (kind == Kind::Boolean) return boolean;
    return asNumber() != 0;
}
std::optional<std::string> VbaValue::asField() const {
    return kind == Kind::Null ? std::nullopt : std::optional<std::string>{asText()};
}
bool validVbaExpression(std::string_view text) {
    const VbaVariables vars;
    try { (void)ExpressionParser(text, nullptr, vars).parse(); return true; }
    catch (...) { return false; }
}
VbaValue evaluateVbaExpression(std::string_view text, VbaHost& host,
                               const VbaVariables& vars, const VbaFunctionResolver& resolver) {
    return ExpressionParser(text, &host, vars, resolver).parse();
}
}
