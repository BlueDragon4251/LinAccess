#include "native_object_types.hpp"
#include "native_utf.hpp"
#include "native_file_system_object.hpp"
#include "vba_array.hpp"
#include "vba_calendar.hpp"
#include "vba_expression.hpp"
#include "vba_color.hpp"
#include "vba_currency.hpp"
#include "vba_decimal.hpp"
#include "vba_round.hpp"
#include "vba_analyzer.hpp"

#include <algorithm>
#include <bit>
#include <list>
#include <chrono>
#include <charconv>
#include <cmath>
#include <cctype>
#include <ctime>
#include <iomanip>
#include <filesystem>
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
// Opaque allocation identity for packed Byte() values; never an Object Variant.
class ByteArrayAllocation final : public VbaObject {
public:
    VbaValue member(std::string_view,const std::vector<VbaValue>&) override {
        throw VbaRuntimeError(3251,"Internal Byte array metadata is not a callable object");
    }
};
struct FinancialAnnuity { long double growth,factor; };
FinancialAnnuity nativeFinancialAnnuity(long double rate,long double periods,long double timing) {
    const auto exponent=periods*std::log1p(rate);
    const auto growth=std::exp(exponent);
    const auto annuity=rate==0 ? periods : std::expm1(exponent)/rate;
    return {growth,annuity*(1+rate*timing)};
}
struct FinancialBalance { long double residual,derivative,scale; };
template<class Evaluate>
double nativeFinancialRoot(long double guess,Evaluate evaluate) {
    if(guess<=-1) throw error("Financial initial guess outside supported range",3251);
    auto rate=guess;
    for(int attempt=0;attempt<20;++attempt) {
        const auto state=evaluate(rate);
        if(!std::isfinite(state.residual) || !std::isfinite(state.derivative) || !std::isfinite(state.scale))
            throw error("Financial iteration overflow",6);
        if(state.derivative==0) throw error("Financial iteration has no usable derivative",5);
        if(state.scale>0 && std::abs(state.residual)<=state.scale*1e-12L) {
            const auto result=static_cast<double>(rate);
            if(!std::isfinite(result)) throw error("Financial root result overflow",6);
            return result;
        }
        auto next=rate-state.residual/state.derivative;
        if(!std::isfinite(next)) throw error("Financial iteration did not converge",5);
        if(next<=-1) next=(rate-1)/2;
        if(next==rate) throw error("Financial iteration did not converge",5);
        rate=next;
    }
    throw error("Financial root did not converge within twenty iterations",5);
}
double nativePeriodicRate(long double periods,long double payment,long double present,long double future,long double timing,long double guess) {
    if(periods<=0) throw error("Rate requires positive periods",5);
    if(periods>32767 || std::floor(periods)!=periods || (timing!=0 && timing!=1))
        throw error("Rate period or timing outside supported range",3251);
    if(present==0 && payment==0 && future==0) throw error("Rate has no unique solution",5);
    return nativeFinancialRoot(guess,[&](long double rate) {
        const auto multiplier=1+rate;
        long double balance=present,derivative=0,scale=std::abs(present);
        for(int period=0;period<static_cast<int>(periods);++period) {
            const auto opening=balance+(timing==1 ? payment : 0);
            derivative=derivative*multiplier+opening;
            balance=opening*multiplier+(timing==0 ? payment : 0);
            scale=scale*multiplier+std::abs(payment)*(timing==1 ? multiplier : 1);
        }
        return FinancialBalance{balance+future,derivative,scale+std::abs(future)};
    });
}
std::size_t nativeFinancialArraySize(const VbaValue& value) {
    if(value.kind==VbaValue::Kind::Null) throw error("Invalid Null financial array",94);
    if(!isVbaArray(value)) throw error("Financial function requires an array",13);
    constexpr int doubleArrayType=8192|5;
    if(vbaArrayRank(value)!=1 || vbaArrayType(value)!=doubleArrayType)
        throw error("Financial function requires a one-dimensional Double array",3251);
    const auto count=vbaArraySize(value);
    if(count==0 || count>1000000) throw error("Financial array size outside supported range",3251);
    return count;
}
VbaValue stringLongResult(std::size_t value) {
    if(value>static_cast<std::size_t>(std::numeric_limits<std::int32_t>::max())) throw error("String result exceeds VBA Long",6);
    return VbaValue::longInteger(static_cast<std::int32_t>(value));
}
VbaValue nativeFileInformation(const VbaValue& input,bool modifiedTime) {
    const auto text=coerceVbaModuleValue(input,"string").asText();
    if(text.find('\0')!=std::string::npos || text.find_first_of("*?")!=std::string::npos) throw error("Invalid native file pathname",52);
    if(text.find('\\')!=std::string::npos || (text.size()>1 && text[1]==':')) throw error("Windows drive/UNC path mapping is not implemented",3251);
    if(text.empty()) throw error("File not found",53);
    namespace fs=std::filesystem;
    const fs::path path(text);
    const auto failure=[&](const std::error_code& reason) -> VbaRuntimeError {
        if(reason==std::errc::permission_denied) return error("File permission denied",70);
        if(reason==std::errc::filename_too_long || reason==std::errc::invalid_argument) return error("Invalid native file pathname",52);
        if(reason==std::errc::no_such_file_or_directory || reason==std::errc::not_a_directory) {
            std::error_code parentError;
            const bool parentExists=!path.has_parent_path() || fs::is_directory(path.parent_path(),parentError);
            if(parentError==std::errc::permission_denied) return error("File permission denied",70);
            return parentExists ? error("File not found",53) : error("Path not found",76);
        }
        return error("Native file metadata unavailable",75);
    };
    std::error_code problem;
    const auto status=fs::status(path,problem);
    if(problem) throw failure(problem);
    if(!fs::exists(status)) throw failure(std::make_error_code(std::errc::no_such_file_or_directory));
    if(!fs::is_regular_file(status) && !(modifiedTime && fs::is_directory(status))) throw error("Unsupported file metadata target",75);
    if(!modifiedTime) {
        const auto length=fs::file_size(path,problem); if(problem) throw failure(problem);
        if(length>static_cast<std::uintmax_t>(std::numeric_limits<std::int32_t>::max())) throw error("File length exceeds VBA Long",6);
        return VbaValue::longInteger(static_cast<std::int32_t>(length));
    }
    const auto modified=fs::last_write_time(path,problem); if(problem) throw failure(problem);
    return nativeFileDate(modified);
}
class NativeVariantContainer {
public:
    virtual ~NativeVariantContainer() = default;
    virtual void visitOwnedValues(const std::function<void(const VbaValue&)>& visitor) const = 0;
};
void checkNativeContainerItem(const VbaObject* owner,const VbaValue& value) {
    std::unordered_set<const VbaObject*> visited; std::size_t budget=1000000;
    const auto check=[&](auto&& self,const VbaValue& entry,int depth) -> void {
        if(!budget-- || depth>64) throw error("Container item traversal limit",3251);
        if(entry.kind==VbaValue::Kind::Object && entry.object) {
            if(entry.object.get()==owner) throw error("Cyclic native container lifetime is not implemented",3251);
            if(visited.insert(entry.object.get()).second) if(const auto* container=dynamic_cast<const NativeVariantContainer*>(entry.object.get()))
                container->visitOwnedValues([&](const VbaValue& child) { self(self,child,depth+1); });
        } else if(entry.kind==VbaValue::Kind::Array && visited.insert(entry.object.get()).second)
            for(std::size_t i=0;i<vbaArraySize(entry);++i) self(self,vbaArrayLinearElement(entry,i),depth+1);
    };
    check(check,value,0);
}
// VBA Collection owns ordered Variant values and optional case-insensitive keys.
class VbaCollection final : public VbaObject,public NativeVariantContainer {
    struct Entry { VbaValue value; std::optional<std::string> key; };
    std::vector<Entry> entries_;
    std::shared_ptr<int> enumerators_=std::make_shared<int>(0);
    static std::string key(const VbaValue& value) {
        const auto converted=coerceVbaModuleValue(scalarVbaValue(value),"string");
        auto text=converted.asText();
        for(auto& ch:text) {
            if(static_cast<unsigned char>(ch)>=128) throw error("Regional Collection keys are not implemented",3251);
            if(ch>='A' && ch<='Z') ch=static_cast<char>(ch-'A'+'a');
        }
        return text;
    }
    std::size_t index(const VbaValue& value) const {
        if(value.kind==VbaValue::Kind::Text) {
            const auto wanted=key(value);
            for(std::size_t i=0;i<entries_.size();++i) if(entries_[i].key && *entries_[i].key==wanted) return i;
        } else {
            const auto ordinal=coerceVbaModuleValue(scalarVbaValue(value),"long").number;
            if(ordinal>=1 && ordinal<=static_cast<double>(entries_.size())) return static_cast<std::size_t>(ordinal-1);
        }
        throw error("Invalid Collection index or key",5);
    }
    void checkItem(const VbaValue& value) const { checkNativeContainerItem(this,value); }
    void requireMutable() const {
        if(*enumerators_) throw error("Collection mutation during For Each is not implemented",3251);
    }
public:
    std::string_view className() const override { return "Collection"; }
    bool isNativeType(std::string_view type) const override { return type=="vba.collection"; }
    void visitOwnedValues(const std::function<void(const VbaValue&)>& visitor) const override { for(const auto& entry:entries_) visitor(entry.value); }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        const auto member=lower(std::string(name));
        if(member=="add") return VbaArgumentSignature{{"Item","Key","Before","After"},1};
        if(member=="item" || member=="remove") return VbaArgumentSignature{{"Index"},1};
        return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        const auto member=lower(std::string(name));
        if(member=="count") {
            if(!args.empty()) throw error("Collection.Count takes no arguments",450);
            return VbaValue::longInteger(static_cast<std::int32_t>(entries_.size()));
        }
        if(member=="item" || member=="remove") {
            if(args.size()!=1 || args[0].kind==VbaValue::Kind::Missing) throw error("Collection member requires an index",450);
            if(member=="remove") requireMutable();
            const auto position=index(args[0]);
            if(member=="item") return entries_[position].value;
            entries_.erase(entries_.begin()+static_cast<std::ptrdiff_t>(position));
            return VbaValue::empty();
        }
        if(member=="add") {
            if(args.empty() || args.size()>4 || args[0].kind==VbaValue::Kind::Missing) throw error("Collection.Add arguments",450);
            requireMutable();
            // Shared acyclic graphs are safe; reject a path back to this owner.
            checkItem(args[0]);
            const auto supplied=[&](std::size_t i) { return i<args.size() && args[i].kind!=VbaValue::Kind::Missing; };
            std::optional<std::string> itemKey;
            if(supplied(1)) {
                itemKey=key(args[1]);
                for(const auto& entry:entries_) if(entry.key==itemKey) throw error("Duplicate Collection key",457);
            }
            if(supplied(2) && supplied(3)) throw error("Collection.Add cannot use both Before and After",5);
            const auto position=supplied(2) ? index(args[2]) : supplied(3) ? index(args[3])+1 : entries_.size();
            if(entries_.size()>=1000000) throw error("Collection allocation limit",3251);
            entries_.insert(entries_.begin()+static_cast<std::ptrdiff_t>(position),Entry{args[0],std::move(itemKey)});
            return VbaValue::empty();
        }
        throw error("Unknown Collection member",438);
    }
    std::shared_ptr<void> enumerationGuard() override {
        struct Lease {
            std::shared_ptr<int> count;
            explicit Lease(std::shared_ptr<int> value):count(std::move(value)) { ++*count; }
            ~Lease() { --*count; }
        };
        return std::make_shared<Lease>(enumerators_);
    }
    std::vector<VbaValue> enumerate() override {
        std::vector<VbaValue> values; values.reserve(entries_.size());
        for(const auto& entry:entries_) values.push_back(entry.value);
        return values;
    }
    void assign(std::string_view,const VbaValue&) override { throw error("Collection properties are read-only",438); }
};

VbaValue clockValue(const std::string& word) {
    const auto sampled=std::chrono::system_clock::now();
    const auto seconds=std::chrono::floor<std::chrono::seconds>(sampled);
    const auto now=std::chrono::system_clock::to_time_t(seconds);
    std::tm local{};
    if (!::localtime_r(&now,&local)) throw error("Systemdatum nicht verfügbar");
    if(word=="timer") {
        const auto fraction=std::chrono::duration<double>(sampled-seconds).count();
        return VbaValue::single(static_cast<float>(local.tm_hour*3600+local.tm_min*60+local.tm_sec+fraction));
    }
    std::ostringstream formatted;
    formatted.imbue(std::locale::classic());
    formatted<<std::put_time(&local,word=="now" ? "%Y-%m-%d %H:%M:%S" : word=="date" ? "%Y-%m-%d" : "%H:%M:%S");
    return vbaDateValue(VbaValue::string(formatted.str()));
}

// Runtime strings use UTF-8 at the API boundary; VBA substring positions use
// UTF-16 code units. Do not silently split a UTF-8 byte sequence.
std::u16string utf16String(std::string_view input) { return nativeUtf16(input); }
std::string utf8String(std::u16string_view input) { return nativeUtf8(input); }

std::string dateLiteral(std::string text) {
    const auto first=text.find_first_not_of(" \t\r\n"),last=text.find_last_not_of(" \t\r\n");
    if(first==std::string::npos) throw error("Empty Date literal",5);
    text=text.substr(first,last-first+1);
    if(text.size()>=10 && text[4]=='-' && text[7]=='-') {
        (void)vbaDateValue(VbaValue::string(text)); return text;
    }
    if(text.find('/')==std::string::npos && text.find(':')!=std::string::npos) {
        (void)vbaDateValue(VbaValue::string(text)); return text;
    }
    const auto end=text.find_first_of(" \t");
    const auto date=text.substr(0,end);
    const auto slash=date.find('/'),second=slash==std::string::npos ? slash : date.find('/',slash+1);
    if(slash==std::string::npos || second==std::string::npos || date.find('/',second+1)!=std::string::npos)
        throw error("Date literal requires ISO or month/day/four-digit-year syntax",3251);
    const auto month=date.substr(0,slash),day=date.substr(slash+1,second-slash-1),year=date.substr(second+1);
    const auto digits=[](const std::string& part) { return !part.empty() && std::all_of(part.begin(),part.end(),[](unsigned char ch){ return ch>='0' && ch<='9'; }); };
    if(month.size()>2 || day.size()>2 || year.size()!=4 || !digits(month) || !digits(day) || !digits(year))
        throw error("Date literal requires numeric month/day and four-digit year",3251);
    std::ostringstream iso; iso<<year<<'-'<<std::setfill('0')<<std::setw(2)<<std::stoi(month)<<'-'<<std::setw(2)<<std::stoi(day);
    if(end!=std::string::npos) {
        const auto time=text.find_first_not_of(" \t",end);
        if(time==std::string::npos) throw error("Date literal time is missing",5);
        iso<<' '<<text.substr(time);
    }
    const auto normalized=iso.str(); (void)vbaDateValue(VbaValue::string(normalized)); return normalized;
}

struct Token {
    enum class Type { Word, Number, Text, Date, Bracket, Symbol, End } type{Type::End};
    std::string value;
    std::size_t offset{};
};

std::vector<Token> tokenize(std::string_view input,bool accessExpression=false) {
    std::vector<Token> tokens;
    for (std::size_t pos = 0; pos < input.size();) {
        const char ch = input[pos];
        if (std::isspace(static_cast<unsigned char>(ch))) { ++pos; continue; }
        const auto tokenOffset=pos;
        if (std::isalpha(static_cast<unsigned char>(ch)) || static_cast<unsigned char>(ch) >= 0x80 || ch == '_') {
            const auto start = pos++;
            while (pos < input.size() &&
                   (std::isalnum(static_cast<unsigned char>(input[pos])) || static_cast<unsigned char>(input[pos]) >= 0x80 || input[pos] == '_')) ++pos;
            if (pos<input.size() && input[pos]=='$') ++pos;
            tokens.push_back({Token::Type::Word, std::string(input.substr(start, pos - start))});
        } else if(ch=='&' && pos+2<input.size()
            && (((input[pos+1]=='h' || input[pos+1]=='H') && std::isxdigit(static_cast<unsigned char>(input[pos+2])))
                || ((input[pos+1]=='o' || input[pos+1]=='O') && input[pos+2]>='0' && input[pos+2]<='7'))) {
            const auto start=pos; const bool hex=input[pos+1]=='h' || input[pos+1]=='H'; pos+=2;
            while(pos<input.size() && (hex ? std::isxdigit(static_cast<unsigned char>(input[pos]))!=0 : input[pos]>='0' && input[pos]<='7')) ++pos;
            if(pos<input.size() && (input[pos]=='&' || input[pos]=='%')) ++pos;
            tokens.push_back({Token::Type::Number,std::string(input.substr(start,pos-start))});
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
        } else if(ch=='#') {
            const auto end=input.find('#',pos+1);
            if(end==std::string_view::npos) throw error("Unterminated Date literal",5);
            tokens.push_back({Token::Type::Date,dateLiteral(std::string(input.substr(pos+1,end-pos-1)))}); pos=end+1;
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
        } else if(ch==':' && pos+1<input.size() && input[pos+1]=='=') {
            tokens.push_back({Token::Type::Symbol,":="}); pos+=2;
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
        tokens.back().offset=tokenOffset;
        if (tokens.size() > 512) throw error("VBA-Ausdruck zu lang");
    }
    tokens.push_back({Token::Type::End,{},input.size()});
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
        text == ">" || text == ">=" || text == "is" || text == "like") return 6;
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
    if(value.kind==VbaValue::Kind::Decimal) return static_cast<std::int32_t>(decimalInteger(value,std::numeric_limits<std::int32_t>::min(),std::numeric_limits<std::int32_t>::max()));
    const auto rounded = bankersRound(value.asNumber());
    if (!std::isfinite(rounded) || rounded < std::numeric_limits<std::int32_t>::min()
        || rounded > std::numeric_limits<std::int32_t>::max()) throw error("VBA: Long-Überlauf", 6);
    return static_cast<std::int32_t>(rounded);
}
VbaValue scalarObjectValue(VbaValue value);
VbaValue typeInspectionValue(VbaValue value);
bool likeMatch(std::u16string_view input,std::u16string_view pattern,VbaHost& host,VbaCompareMode mode) {
    struct Part {
        enum class Kind { Literal, Star, Any, Digit, List } kind{Kind::Literal};
        char16_t literal{};
        bool negate{};
        std::vector<std::pair<char16_t,char16_t>> ranges;
    };
    const auto compare=[&](char16_t left,char16_t right) {
        if(mode==VbaCompareMode::Binary) return (left>right)-(left<right);
        if((left>=0xd800 && left<=0xdfff) || (right>=0xd800 && right<=0xdfff))
            throw error("Locale-aware Like surrogate comparisons are not implemented",3251);
        return host.compareStrings(utf8String(std::u16string(1,left)),utf8String(std::u16string(1,right)),mode);
    };
    std::vector<Part> parts;
    for(std::size_t index=0;index<pattern.size();++index) {
        const auto unit=pattern[index]; Part part;
        if(unit==u'*') part.kind=Part::Kind::Star;
        else if(unit==u'?') part.kind=Part::Kind::Any;
        else if(unit==u'#') part.kind=Part::Kind::Digit;
        else if(unit==u'[') {
            part.kind=Part::Kind::List;
            const auto closing=pattern.find(u']',index+1);
            if(closing==std::u16string_view::npos) throw error("Invalid Like pattern",93);
            if(closing==index+1) { index=closing; continue; }
            std::size_t cursor=index+1;
            if(pattern[cursor]==u'!') { part.negate=true; ++cursor; }
            if(cursor==closing) throw error("Invalid empty negated Like class",93);
            while(cursor<closing) {
                const auto first=pattern[cursor++]; auto last=first;
                if(cursor+1<closing && pattern[cursor]==u'-') {
                    last=pattern[cursor+1]; cursor+=2;
                    if(mode!=VbaCompareMode::Binary && (first>127 || last>127))
                        throw error("Non-ASCII Like range collation is not implemented",3251);
                    if(compare(first,last)>0) throw error("Descending Like range",93);
                }
                part.ranges.emplace_back(first,last);
            }
            index=closing;
        } else part.literal=unit;
        parts.push_back(std::move(part));
    }
    // Two rows avoid recursive wildcard backtracking and retain UTF-16 positions.
    std::vector<unsigned char> previous(input.size()+1),current(input.size()+1);
    previous[0]=1;
    for(const auto& part:parts) {
        std::fill(current.begin(),current.end(),0);
        if(part.kind==Part::Kind::Star) {
            current[0]=previous[0];
            for(std::size_t index=1;index<=input.size();++index) current[index]=previous[index] || current[index-1];
        } else for(std::size_t index=1;index<=input.size();++index) {
            if(!previous[index-1]) continue;
            const auto unit=input[index-1]; bool match=false;
            switch(part.kind) {
            case Part::Kind::Any: match=true; break;
            case Part::Kind::Digit: match=unit>=u'0' && unit<=u'9'; break;
            case Part::Kind::Literal: match=compare(unit,part.literal)==0; break;
            case Part::Kind::List:
                for(const auto& range:part.ranges) {
                    if(mode!=VbaCompareMode::Binary && range.first!=range.second && unit>127)
                        throw error("Non-ASCII Like range collation is not implemented",3251);
                    if(compare(unit,range.first)>=0 && compare(unit,range.second)<=0) { match=true; break; }
                }
                match=match!=part.negate; break;
            case Part::Kind::Star: break;
            }
            current[index]=match;
        }
        previous.swap(current);
    }
    return previous.back()!=0;
}
VbaValue binary(const std::string& op, const VbaValue& left, const VbaValue& right,VbaHost* host,VbaCompareMode mode) {
    if (left.kind==VbaValue::Kind::Error || right.kind==VbaValue::Kind::Error || left.kind==VbaValue::Kind::Missing || right.kind==VbaValue::Kind::Missing)
        throw error("Error Variant cannot be used as an operand",13);
    if (op == "is") {
        if (left.kind != VbaValue::Kind::Object || right.kind != VbaValue::Kind::Object)
            throw error("VBA: Is benötigt Objekte");
        return VbaValue::logical(left.object == right.object);
    }
    if (left.kind==VbaValue::Kind::Object) return binary(op,scalarObjectValue(left),right,host,mode);
    if (right.kind==VbaValue::Kind::Object) return binary(op,left,scalarObjectValue(right),host,mode);
    if(op=="like") {
        if(left.kind==VbaValue::Kind::Null || right.kind==VbaValue::Kind::Null) return {};
        return VbaValue::logical(likeMatch(utf16String(left.asText()),utf16String(right.asText()),*host,mode));
    }
    if (op == "&") {
        if (left.kind == VbaValue::Kind::Null && right.kind == VbaValue::Kind::Null) return {};
        return VbaValue::string(left.asText() + right.asText());
    }
    if (op == "and" || op == "or") {
        const bool leftNull = left.kind == VbaValue::Kind::Null;
        const bool rightNull = right.kind == VbaValue::Kind::Null;
        if (leftNull || rightNull) {
            if(leftNull && rightNull) return {};
            const auto& known=leftNull ? right : left;const auto bits=static_cast<std::int32_t>(coerceVbaModuleValue(known,"long").number);
            if((op=="and" && bits==0) || (op=="or" && bits==-1)) {
                if(known.kind==VbaValue::Kind::Boolean) return VbaValue::logical(op=="or");
                return known.kind==VbaValue::Kind::Number && known.numericType==VbaValue::NumericType::Integer ? VbaValue::integer(op=="or" ? -1 : 0) : VbaValue::longInteger(op=="or" ? -1 : 0);
            }
            return {};
        }
        if (left.kind == VbaValue::Kind::Boolean && right.kind == VbaValue::Kind::Boolean)
            return VbaValue::logical(op == "and" ? left.boolean && right.boolean : left.boolean || right.boolean);
        const auto a=static_cast<std::int32_t>(coerceVbaModuleValue(left,"long").number);
        const auto b=static_cast<std::int32_t>(coerceVbaModuleValue(right,"long").number);
        const auto answer=op=="and" ? a&b : a|b;
        const bool narrow=left.kind==VbaValue::Kind::Number && right.kind==VbaValue::Kind::Number
            && left.numericType==VbaValue::NumericType::Integer && right.numericType==VbaValue::NumericType::Integer;
        return narrow ? VbaValue::integer(static_cast<std::int16_t>(answer)) : VbaValue::longInteger(answer);
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
        } else if(left.kind==VbaValue::Kind::Decimal || right.kind==VbaValue::Kind::Decimal) cmp=decimalCompare(left,right);
        else {
            const auto a = left.asNumber(), b = right.asNumber();
            cmp = (a > b) - (a < b);
        }
        return VbaValue::logical(op == "=" ? cmp == 0 : op == "<>" ? cmp != 0 :
            op == "<" ? cmp < 0 : op == "<=" ? cmp <= 0 : op == ">" ? cmp > 0 : cmp >= 0);
    }
    if (op == "+" && left.kind == VbaValue::Kind::Text && right.kind == VbaValue::Kind::Text)
        return VbaValue::string(left.text + right.text);
    if(left.kind==VbaValue::Kind::Decimal || right.kind==VbaValue::Kind::Decimal) return decimalBinary(op,left,right);
    if (left.kind==VbaValue::Kind::Currency && right.kind==VbaValue::Kind::Currency && (op=="+" || op=="-" || op=="*"))
        return currencyBinary(op,left,right);
    if((left.kind==VbaValue::Kind::Date || right.kind==VbaValue::Kind::Date) && (op=="+" || op=="-")) {
        if(left.kind==VbaValue::Kind::Text || right.kind==VbaValue::Kind::Text) throw error("Text/Date arithmetic coercion is not implemented",3251);
        const auto a=left.asNumber(),b=right.asNumber();
        if((left.kind==VbaValue::Kind::Date && a<0 && a!=std::trunc(a)) || (right.kind==VbaValue::Kind::Date && b<0 && b!=std::trunc(b)))
            throw error("Negative fractional Date arithmetic is not validated",3251);
        if(op=="-" && left.kind==VbaValue::Kind::Date && right.kind==VbaValue::Kind::Date) return VbaValue::numeric(a-b);
        const auto result=op=="+" ? a+b : a-b;
        if(!std::isfinite(result)) throw error("Date arithmetic overflow",6);
        if(result<=-657435.0 || result>=2958466.0) return VbaValue::numeric(result);
        if(result<0 && result!=std::trunc(result)) throw error("Negative fractional Date arithmetic is not validated",3251);
        return VbaValue::dateSerial(result);
    }
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

class VbaDictionary final : public VbaObject,public NativeVariantContainer {
    struct Key { bool numeric{}; std::uint64_t number{}; std::string text; bool operator==(const Key&) const = default; };
    struct Hash { std::size_t operator()(const Key& key) const { return key.numeric ? std::hash<std::uint64_t>{}(key.number)^0x9e3779b9U : std::hash<std::string>{}(key.text); } };
    struct Entry { Key key; VbaValue original,value; };
    std::list<Entry> entries_;
    std::unordered_map<Key,std::list<Entry>::iterator,Hash> index_;
    std::shared_ptr<int> enumerators_=std::make_shared<int>(0);
    int compare_{};
    Key key(const VbaValue& value) const {
        if(value.kind==VbaValue::Kind::Number) {
            if(!std::isfinite(value.number)) throw error("Non-finite Dictionary keys are not implemented",3251);
            return {true,std::bit_cast<std::uint64_t>(value.number==0 ? 0.0 : value.number),{}};
        }
        if(value.kind!=VbaValue::Kind::Text) throw error("Dictionary keys require native String or numeric values",3251);
        (void)utf16String(value.text); auto text=value.text;
        if(compare_) for(auto& ch:text) {
            if(static_cast<unsigned char>(ch)>127) throw error("Regional Dictionary text keys are not implemented",3251);
            if(ch>='A' && ch<='Z') ch=static_cast<char>(ch-'A'+'a');
        }
        return {false,0,std::move(text)};
    }
    void structural() const { if(*enumerators_) throw error("Dictionary structural mutation during enumeration is not implemented",3251); }
    void add(const VbaValue& original,const VbaValue& value) {
        const auto normalized=key(original); structural();
        if(index_.contains(normalized)) throw error("Duplicate Dictionary key",457);
        if(entries_.size()>=1000000) throw error("Dictionary allocation limit",3251);
        checkNativeContainerItem(this,value);
        entries_.push_back({normalized,original,value});
        try { index_.emplace(normalized,std::prev(entries_.end())); } catch(...) { entries_.pop_back(); throw; }
    }
    static void count(const std::vector<VbaValue>& args,std::size_t expected) {
        if(args.size()!=expected || std::any_of(args.begin(),args.end(),[](const VbaValue& value){ return value.kind==VbaValue::Kind::Missing; })) throw error("Dictionary argument count",450);
    }
public:
    std::string_view className() const override { return "Dictionary"; }
    bool isNativeType(std::string_view type) const override { return type=="scripting.dictionary"; }
    void visitOwnedValues(const std::function<void(const VbaValue&)>& visitor) const override { for(const auto& entry:entries_) visitor(entry.value); }
    std::optional<VbaArgumentSignature> argumentSignature(std::string_view name) const override {
        const auto member=lower(std::string(name));
        if(member=="add") return VbaArgumentSignature{{"Key","Item"},2};
        if(member=="item" || member=="exists" || member=="remove" || member=="key") return VbaArgumentSignature{{"Key"},1};
        return {};
    }
    VbaValue member(std::string_view name,const std::vector<VbaValue>& args) override {
        const auto member=lower(std::string(name));
        if(member=="count" || member=="comparemode") { count(args,0); return VbaValue::longInteger(static_cast<std::int32_t>(member=="count" ? entries_.size() : compare_)); }
        if(member=="add") { count(args,2); add(args[0],args[1]); return VbaValue::empty(); }
        if(member=="item" || member=="exists" || member=="remove") {
            count(args,1); const auto normalized=key(args[0]); const auto found=index_.find(normalized);
            if(member=="exists") return VbaValue::logical(found!=index_.end());
            if(member=="remove") {
                structural(); if(found==index_.end()) throw error("Removing an absent Dictionary key is not implemented",3251);
                entries_.erase(found->second); index_.erase(found); return VbaValue::empty();
            }
            if(found==index_.end()) { add(args[0],VbaValue::empty()); return VbaValue::empty(); }
            return found->second->value;
        }
        if(member=="keys" || member=="items") {
            count(args,0); std::vector<VbaValue> values; values.reserve(entries_.size());
            for(const auto& entry:entries_) values.push_back(member=="keys" ? entry.original : entry.value);
            return makeDetachedVbaArray(std::move(values));
        }
        if(member=="removeall") { count(args,0); structural(); index_.clear(); entries_.clear(); return VbaValue::empty(); }
        throw error("Unknown Dictionary member",438);
    }
    void assign(std::string_view name,const VbaValue& value) override {
        if(lower(std::string(name))!="comparemode") throw error("Dictionary property is read-only or unknown",438);
        structural(); if(!entries_.empty()) throw error("Dictionary CompareMode cannot change after adding data",5);
        const auto mode=static_cast<int>(coerceVbaModuleValue(scalarVbaValue(value),"long").number);
        if(mode!=0 && mode!=1) throw error("Dictionary database/locale comparison is not implemented",3251);
        compare_=mode;
    }
    void assignIndexedMember(std::string_view name,const std::vector<VbaValue>& args,const VbaValue& value) override {
        count(args,1); const auto member=lower(std::string(name));
        const auto normalized=key(args[0]); const auto found=index_.find(normalized);
        if(member=="item") {
            if(found==index_.end()) { add(args[0],value); return; }
            checkNativeContainerItem(this,value); auto staged=value; found->second->value=std::move(staged); return;
        }
        if(member=="key") {
            structural(); if(found==index_.end()) throw error("Renaming an absent Dictionary key is not implemented",3251);
            auto replacement=key(value);
            if(replacement==normalized) throw error("Equivalent Dictionary key renaming is not implemented",3251);
            if(index_.contains(replacement)) throw error("Duplicate Dictionary key",457);
            const auto entry=found->second; auto original=value;
            index_.emplace(replacement,entry); index_.erase(normalized);
            entry->key=std::move(replacement); entry->original=std::move(original); return;
        }
        throw error("Unknown Dictionary indexed property",438);
    }
    std::shared_ptr<void> enumerationGuard() override {
        struct Lease { std::shared_ptr<int> count; explicit Lease(std::shared_ptr<int> value):count(std::move(value)) { ++*count; } ~Lease() { --*count; } };
        return std::make_shared<Lease>(enumerators_);
    }
    std::vector<VbaValue> enumerate() override {
        std::vector<VbaValue> keys; keys.reserve(entries_.size()); for(const auto& entry:entries_) keys.push_back(entry.original); return keys;
    }
};

bool vbaIntrinsicName(const std::string& name) {
    static const std::unordered_set<std::string> intrinsics{
                "array","lbound","ubound","isarray","ismissing","isobject","iserror","cverr","vartype","typename",
                "createobject","doevents","msgbox","inputbox","inputbox$","round","asc","ascw","chr","chr$","chrw","chrw$",
                "cbyte","cint","clng","csng","cdbl","cbool","ccur","cdec","cdate","cstr","str","str$","format","format$","formatnumber","formatcurrency","formatpercent",
                "left","left$","right","right$","mid","mid$","trim","trim$","ltrim","ltrim$","rtrim","rtrim$","len","lenb",
                "ucase","ucase$","lcase","lcase$","instr","strcomp","instrrev","strreverse","replace","split","join","filter",
                "hex","hex$","oct","oct$","space","space$","string","string$","val","abs","int","fix","sgn",
                "sqr","exp","log","sin","cos","tan","atn","date","time","now","timer","dateadd","datediff","datepart",
                "year","month","day","hour","minute","second","weekday","dateserial","timeserial","datevalue","timevalue",
                "filelen","filedatetime","curdir","curdir$","choose","switch","iif","isnumeric","isnull","isempty","rgb","qbcolor","partition","fv","pmt","ipmt","ppmt","pv","nper","rate","npv","irr","mirr","sln","syd","ddb"};
    return intrinsics.contains(name);
}

class ExpressionParser {
public:
    ExpressionParser(std::string_view source, VbaHost* host, const VbaVariables& locals,
                     const VbaFunctionResolver& resolver = {},bool objectContext=false,bool accessExpression=false,const VbaValue* withContext=nullptr,bool allowWith=false,VbaCompareMode compareMode=VbaCompareMode::Binary,int arrayBase=0,bool jetEval=false)
        : source_(source),tokens_(tokenize(source,accessExpression)), host_(host), locals_(locals), resolver_(resolver),objectContext_(objectContext),accessExpression_(accessExpression),jetEval_(jetEval),withContext_(withContext),allowWith_(allowWith || withContext),compareMode_(compareMode),arrayBase_(arrayBase),executionEpoch_(vbaExecutionEpoch()) {}

    VbaValue parse() {
        auto value = expression(1,objectContext_);
        if(host_ && (vbaExecutionStopRequested() || executionEpoch_!=vbaExecutionEpoch())) throw VbaExecutionStopped{};
        if (peek().type != Token::Type::End)
            throw error("Zusätzliche Zeichen im VBA-Ausdruck: " + peek().value);
        return value;
    }
    VbaEvaluatedArgument parseArgument() {
        auto result=argumentExpression(objectContext_);
        if(host_ && (vbaExecutionStopRequested() || executionEpoch_!=vbaExecutionEpoch())) throw VbaExecutionStopped{};
        if(peek().type!=Token::Type::End) throw error("Additional argument expression tokens");
        return result;
    }
private:
    struct ReferenceCapture { std::size_t start{},end{}; std::optional<VbaArrayElementReference> element; };
    std::optional<ReferenceCapture> referenceCapture_;
    VbaEvaluatedArgument argumentExpression(bool objectContext=false) {
        auto previous=std::move(referenceCapture_);
        referenceCapture_=ReferenceCapture{index_,index_,{}};
        try {
            VbaEvaluatedArgument result{expression(1,objectContext),{}};
            if(referenceCapture_ && referenceCapture_->end==index_) result.arrayElement=std::move(referenceCapture_->element);
            referenceCapture_=std::move(previous); return result;
        } catch(...) { referenceCapture_=std::move(previous); throw; }
    }
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
        if(host_ && (vbaExecutionStopRequested() || executionEpoch_!=vbaExecutionEpoch())) throw VbaExecutionStopped{};
        auto lhs = primary(objectContext,minPrecedence);
        while (true) {
            const int prec = precedence(peek());
            if (prec < minPrecedence) break;
            const auto op = lower(consume().value);
            if(referenceCapture_) referenceCapture_->element.reset();
            if (accessExpression_ && op=="is" && (lower(peek().value)=="null" || lower(peek().value)=="not")) {
                const bool negate=take("not"); require("null");
                if (host_ && lhs.kind!=VbaValue::Kind::Error) lhs=VbaValue::logical((lhs.kind==VbaValue::Kind::Null)!=negate);
                continue;
            }
            auto rhs = expression(prec + 1);
            if (!host_) lhs={};
            else if (accessExpression_ && (lhs.kind==VbaValue::Kind::Error || rhs.kind==VbaValue::Kind::Error))
                lhs=lhs.kind==VbaValue::Kind::Error ? lhs : rhs;
            else {
                const auto left=lhs;
                try {lhs=binary(op,left,rhs,host_,compareMode_);}catch(const VbaRuntimeError& failure) {if(!jetEval_) throw;lhs=VbaValue::errorValue(failure.number);}
                if(jetEval_ && lhs.kind==VbaValue::Kind::Boolean && (op=="and" || op=="or") && (left.kind==VbaValue::Kind::Null || rhs.kind==VbaValue::Kind::Null)) {const auto& known=left.kind==VbaValue::Kind::Null ? rhs : left;if(known.kind==VbaValue::Kind::Number && known.numericType==VbaValue::NumericType::Integer) lhs=VbaValue::integer(lhs.boolean ? -1 : 0);}
                const auto integral=[](const VbaValue& value) {return value.kind==VbaValue::Kind::Number && (value.numericType==VbaValue::NumericType::Long || value.numericType==VbaValue::NumericType::Integer || value.numericType==VbaValue::NumericType::Byte);};
                if(jetEval_ && (op=="+" || op=="-" || op=="*") && integral(left) && integral(rhs) && lhs.kind==VbaValue::Kind::Number && lhs.number>=-2147483648.0 && lhs.number<=2147483647.0 && std::trunc(lhs.number)==lhs.number) lhs=VbaValue::longInteger(static_cast<std::int32_t>(lhs.number));
            }
        }
        return lhs;
    }
    VbaValue primary(bool objectContext,int minPrecedence) {
        const auto primaryStart=index_;
        auto value = [&] {
            if (!nextSymbol(".") && !nextSymbol("!")) return atom(objectContext,minPrecedence);
            if (!host_ && allowWith_) return VbaValue{};
            if (!withContext_) throw error("With object context unavailable",91);
            return *withContext_;
        }();
        while (nextSymbol("(") || nextSymbol(".") || nextSymbol("!")) {
            if(host_ && (vbaExecutionStopRequested() || executionEpoch_!=vbaExecutionEpoch())) throw VbaExecutionStopped{};
            std::string compoundArrayRoot;
            if(referenceCapture_ && referenceCapture_->element && referenceCapture_->end==index_) {
                if(nextSymbol("(") && isVbaArray(value)) compoundArrayRoot=referenceCapture_->element->root;
                referenceCapture_->element.reset();
            }
            const auto suffixStart=index_;
            if (take("(")) {
                std::vector<VbaValue> arguments;
                if (!take(")")) {
                    do { arguments.push_back(expression(1)); } while (take(",") || (accessExpression_ && take(";")));
                    require(")");
                }
                if (host_) {
                    try {
                        if(isVbaArray(value)) {
                            if(!arguments.empty() || accessExpression_) {
                                const bool root=!accessExpression_ && suffixStart==primaryStart+1 && tokens_[primaryStart].type==Token::Type::Word;
                                const auto source=root ? locals_.find(lower(tokens_[primaryStart].value)) : locals_.end();
                                if(source!=locals_.end() && isVbaArray(source->second)) {
                                    if(!sameVbaArrayAllocation(value,source->second)) throw error("Array source changed while evaluating subscripts",3251);
                                    value=source->second;
                                }
                                const auto element=vbaArrayElementOffset(value,arguments);
                                if(source!=locals_.end()) {
                                    if(!isVbaArray(source->second) || !sameVbaArrayAllocation(value,source->second)) throw error("Array source changed while coercing subscripts",3251);
                                    value=source->second;
                                }
                                if(referenceCapture_ && referenceCapture_->start==primaryStart && source!=locals_.end()) {
                                    auto type=lower(vbaArrayTypeName(value)); type.resize(type.size()-2);
                                    referenceCapture_->element=VbaArrayElementReference{source->first,std::move(type),element,vbaArrayAllocation(value)};
                                    referenceCapture_->end=index_;
                                }
                                value=vbaArrayLinearElement(value,element);
                            }
                            if(referenceCapture_ && !compoundArrayRoot.empty()) {
                                // A compound address is recognized but not bound as a temporary.
                                referenceCapture_->element=VbaArrayElementReference{std::move(compoundArrayRoot),{},0,{}};
                                referenceCapture_->end=index_;
                            }
                            continue;
                        }
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
            const auto memberKey=lower(member.value);
            const bool missingDefaults=memberKey=="add" || memberKey=="opentable" || memberKey=="openquery" || memberKey=="rename" || memberKey=="deleteobject" || memberKey=="findrecord" || memberKey=="raise";
            std::vector<VbaValue> args;
            std::vector<std::string> argumentNames;
            std::vector<bool> omittedArguments;
            if (take("(")) {
                if (!take(")")) {
                    const auto signature=host_ && value.kind==VbaValue::Kind::Object && value.object ? value.object->argumentSignature(member.value) : std::nullopt;
                    do {
                        std::string argumentName;
                        if(peek().type==Token::Type::Word && index_+1<tokens_.size() && tokens_[index_+1].value==":=") {
                            argumentName=lower(consume().value); consume();
                        }
                        const bool omitted=(missingDefaults || memberKey=="sendkeys" || lower(member.value)=="openform" || lower(member.value)=="openreport" || lower(member.value)=="gotorecord" || lower(member.value)=="domenuitem" || lower(member.value)=="selectobject" || lower(member.value)=="close") && (nextSymbol(",") || nextSymbol(")"));
                        args.push_back(omitted ? missingDefaults ? VbaValue::missing() : VbaValue::empty()
                            : expression(1,signature && signature->objectArgument(args.size(),argumentName)));
                        if(!argumentName.empty() && omitted) throw error("Named method argument requires a value",450);
                        argumentNames.push_back(std::move(argumentName)); omittedArguments.push_back(omitted);
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
                if(std::any_of(argumentNames.begin(),argumentNames.end(),[](const std::string& name){ return !name.empty(); })) {
                    const auto signature=value.object->argumentSignature(member.value);
                    if(!signature) throw error("Named method arguments are not implemented for this member",3251);
                    std::vector<VbaValue> ordered(signature->names.size(),missingDefaults ? VbaValue::missing() : VbaValue::empty());
                    std::vector<bool> assigned(signature->names.size()),provided(signature->names.size());
                    bool namedSeen=false; std::size_t last=0;
                    for(std::size_t index=0;index<args.size();++index) {
                        std::size_t target=index;
                        if(!argumentNames[index].empty()) {
                            namedSeen=true; target=signature->names.size();
                            for(std::size_t parameter=0;parameter<signature->names.size();++parameter)
                                if(lower(signature->names[parameter])==argumentNames[index]) { target=parameter; break; }
                            if(target==signature->names.size()) throw error("Named method argument not found",448);
                        } else if(namedSeen) throw error("Positional argument after named argument",450);
                        if(target>=ordered.size() || assigned[target]) throw error("Method argument supplied more than once or too many arguments",450);
                        ordered[target]=args[index]; assigned[target]=true; provided[target]=!omittedArguments[index]; last=std::max(last,target);
                    }
                    for(std::size_t index=0;index<signature->required;++index) if(!provided[index]) throw error("Required method argument omitted",450);
                    ordered.resize(last+1); args=std::move(ordered);
                }
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
            const auto arg = expression(std::max(jetEval_ ? 14 : 12,minPrecedence));
            if(!host_) return {};
            const auto value=scalarObjectValue(arg);
            if(value.kind==VbaValue::Kind::Null) return {};
            if(value.kind==VbaValue::Kind::Decimal) return decimalUnary("-",value);
            if(jetEval_ && value.kind==VbaValue::Kind::Number && value.numericType==VbaValue::NumericType::Long && value.number!=-2147483648.0) return VbaValue::longInteger(static_cast<std::int32_t>(-value.number));
            return value.kind==VbaValue::Kind::Currency ? negateCurrency(value) : VbaValue::numeric(-value.asNumber());
        }
        if (take("+")) {
            const auto arg = expression(std::max(jetEval_ ? 14 : 12,minPrecedence));
            if(!host_) return {};
            const auto value=scalarObjectValue(arg);
            if(value.kind==VbaValue::Kind::Null) return {};
            if(value.kind==VbaValue::Kind::Decimal || (jetEval_ && value.kind==VbaValue::Kind::Number)) return value;
            return value.kind==VbaValue::Kind::Currency ? value : VbaValue::numeric(value.asNumber());
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
        if (take("typeof")) {
            if(accessExpression_) throw error("TypeOf is unavailable in calculated expressions",3251);
            const auto root=lower(peek().value);
            const bool reference=peek().type==Token::Type::Word || peek().type==Token::Type::Bracket || nextSymbol(".") || nextSymbol("!");
            if(!reference || root=="nothing" || root=="null" || root=="empty" || root=="true" || root=="false" || root=="new" || root=="typeof")
                throw error("TypeOf requires a native reference expression",3251);
            const auto object=primary(true,7);
            require("is");
            auto token=consume();
            if(token.type!=Token::Type::Word) throw error("TypeOf requires a class name",3251);
            auto type=lower(token.value);
            while(take(".")) {
                token=consume(); if(token.type!=Token::Type::Word) throw error("TypeOf class qualifier is incomplete",3251);
                type+="."+lower(token.value);
            }
            const auto nativeType=nativeObjectType(type);
            if(nativeType.empty()) throw error("TypeOf class binding is not implemented",3251);
            if(referenceCapture_) referenceCapture_->element.reset();
            if(!host_) return {};
            if(object.kind!=VbaValue::Kind::Object) return VbaValue::logical(false);
            if(!object.object) throw error("TypeOf object variable is not set",91);
            return VbaValue::logical(nativeType=="object" || object.object->isNativeType(nativeType));
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
        if(tok.type==Token::Type::Date) return host_ ? vbaDateValue(VbaValue::string(tok.value)) : VbaValue{};
        if (tok.type == Token::Type::Text) return VbaValue::string(tok.value);
        if (tok.type == Token::Type::Number) {
            if(tok.value.starts_with('&')) {
                const bool forceLong=tok.value.back()=='&',forceInteger=tok.value.back()=='%';
                const auto digits=tok.value.substr(2,tok.value.size()-2-(forceLong || forceInteger ? 1 : 0));
                std::size_t consumed=0; std::uint64_t magnitude;
                try { magnitude=std::stoull(digits,&consumed,tok.value[1]=='h' || tok.value[1]=='H' ? 16 : 8); }
                catch(const std::out_of_range&) { throw error("Based integer literal overflow",6); }
                catch(const std::invalid_argument&) { throw error("Invalid based integer literal",5); }
                if(consumed!=digits.size()) throw error("Invalid based integer literal",5);
                if(magnitude>0xffffffffULL || (forceInteger && magnitude>0xffff)) throw error("Based integer literal overflow",6);
                if(!forceLong && magnitude<=0xffff) {
                    const auto signedValue=static_cast<std::int64_t>(magnitude)-(magnitude>0x7fff ? 0x10000LL : 0);
                    return VbaValue::integer(static_cast<std::int16_t>(signedValue));
                }
                const auto signedValue=static_cast<std::int64_t>(magnitude)-(magnitude>0x7fffffff ? 0x100000000LL : 0);
                return VbaValue::longInteger(static_cast<std::int32_t>(signedValue));
            }
            if(jetEval_ && tok.value.find_first_of(".eE")!=std::string::npos && tok.value.find_first_of("eE")==std::string::npos) return decimalLiteral(tok.value);
            std::size_t consumed = 0;
            double number;
            try { number=std::stod(tok.value,&consumed); }
            catch(const std::out_of_range&) { throw error("Numeric literal overflow",6); }
            catch(const std::invalid_argument&) { throw error("Invalid numeric literal",5); }
            if (consumed!=tok.value.size()) throw error("Ungültige VBA-Zahl");
            if (!std::isfinite(number)) throw error("Numeric literal overflow",6);
            if(jetEval_ && tok.value.find_first_of(".eE")==std::string::npos && number>=0 && number<=2147483647.0) return VbaValue::longInteger(static_cast<std::int32_t>(number));
            return VbaValue::numeric(number);
        }
        if (tok.type != Token::Type::Word)
            throw error("Ungültiger VBA-Ausdruck: " + tok.value);
        auto word = lower(tok.value);
        if(!accessExpression_ && nextSymbol(".") && index_+1<tokens_.size() && tokens_[index_+1].type==Token::Type::Word
            && !locals_.contains(word) && (!host_ || !host_->hasInstanceIdentifier(word))) {
            const auto qualified=word+"."+lower(tokens_[index_+1].value);
            if(const auto found=locals_.find(qualified);found!=locals_.end()) {
                consume(); consume(); return found->second;
            }
        }
        bool qualifiedArray=false,qualifiedIntrinsic=false;
        if(!accessExpression_ && word=="vba" && nextSymbol(".") && index_+1<tokens_.size() && tokens_[index_+1].type==Token::Type::Word) {
            if(const auto local=locals_.find(word);local!=locals_.end()) return local->second;
            if(host_) if(auto resolved=host_->resolveIdentifier(tok.value,true)) return *resolved;
            consume(); word=lower(consume().value);

            if(!vbaIntrinsicName(word) && !word.starts_with("vb")) throw error("Qualified VBA member is not implemented",3251);
            qualifiedIntrinsic=true; qualifiedArray=word=="array";
        }
        if (accessExpression_) {
            static const std::unordered_map<std::string,std::string> localized{{"summe","sum"},{"anzahl","count"},{"wenn","iif"},{"istnull","isnull"},{"istfehler","iserror"},{"datum","date"},{"jetzt","now"},{"zeit","time"},{"formulare","forms"},{"berichte","reports"}};
            if (const auto found=localized.find(word);found!=localized.end()) word=found->second;
        }
        if(word=="new") {
            if(accessExpression_) throw error("New is unavailable in calculated expressions",3251);
            auto type=consume();
            if(type.type==Token::Type::Word && lower(type.value)=="scripting" && take(".")) {
                type=consume();
                if(type.type!=Token::Type::Word) throw error("Native Scripting class is not implemented",3251);
                const auto name=lower(type.value);
                if(name=="dictionary") return host_ ? VbaValue::reference(std::make_shared<VbaDictionary>()) : VbaValue{};
                if(name=="filesystemobject") return host_ ? VbaValue::reference(std::make_shared<VbaFileSystemObject>()) : VbaValue{};
                throw error("Native Scripting class is not implemented",3251);
            }
            if(type.type==Token::Type::Word && lower(type.value)=="vba" && take(".")) type=consume();
            if(type.type!=Token::Type::Word || lower(type.value)!="collection") throw error("Native class construction is not implemented",3251);
            return host_ ? VbaValue::reference(std::make_shared<VbaCollection>()) : VbaValue{};
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
        static const std::unordered_map<std::string,int> fieldAttributes{
            {"dbfixedfield",1},{"dbvariablefield",2},{"dbautoincrfield",16},
            {"dbupdatablefield",32},{"dbsystemfield",8192},{"dbhyperlinkfield",32768},
            {"dbdescending",1}};
        if(const auto attribute=fieldAttributes.find(word);attribute!=fieldAttributes.end()) return VbaValue::longInteger(attribute->second);
        if(word=="dbrelationunique") return VbaValue::longInteger(1);
        if(word=="dbrelationdontenforce") return VbaValue::longInteger(2);
        if(word=="dbrelationupdatecascade") return VbaValue::longInteger(256);
        if(word=="dbrelationdeletecascade") return VbaValue::longInteger(4096);
        static const std::unordered_map<std::string,int> queryTypes{
            {"dbqselect",0},{"dbqcrosstab",16},{"dbqdelete",32},{"dbqupdate",48},
            {"dbqappend",64},{"dbqmaketable",80},{"dbqddl",96},{"dbqsqlpassthrough",112},
            {"dbqsetoperation",128},{"dbqsptbulk",144},{"dbqcompound",160},{"dbqprocedure",224},{"dbqaction",240}};
        if(const auto type=queryTypes.find(word);type!=queryTypes.end()) return VbaValue::integer(static_cast<std::int16_t>(type->second));
        static const std::unordered_map<std::string,int> fileDialogTypes{
            {"msofiledialogopen",1},{"msofiledialogsaveas",2},{"msofiledialogfilepicker",3},{"msofiledialogfolderpicker",4}};
        if(const auto type=fileDialogTypes.find(word);type!=fileDialogTypes.end()) return VbaValue::longInteger(type->second);
        static const std::unordered_map<std::string,int> statusActions{
            {"acsyscmdinitmeter",1},{"acsyscmdupdatemeter",2},{"acsyscmdremovemeter",3},
            {"acsyscmdsetstatus",4},{"acsyscmdclearstatus",5},{"acsyscmdgetobjectstate",10},
            {"acobjstateopen",1},{"acobjstatedirty",2},{"acobjstatenew",4}};
        if(const auto action=statusActions.find(word);action!=statusActions.end()) return VbaValue::longInteger(action->second);
        if(word=="vbusedefault") return VbaValue::longInteger(-2);
        if(word=="vbtrue") return VbaValue::longInteger(-1);
        if(word=="vbfalse") return VbaValue::longInteger(0);
        if(word=="vbdefault") return VbaValue::integer(0);
        if(word=="vbhourglass") return VbaValue::integer(11);
        static const std::unordered_map<std::string,std::int32_t> relationAttributes{{"dbdescending",1},{"dbrelationunique",1},{"dbrelationdontenforce",2},{"dbrelationinherited",4},
            {"dbrelationupdatecascade",256},{"dbrelationdeletecascade",4096},{"dbrelationleft",16777216},{"dbrelationright",33554432}};
        if(const auto attribute=relationAttributes.find(word);attribute!=relationAttributes.end()) return VbaValue::longInteger(attribute->second);
        if (word == "dbopentable") return VbaValue::longInteger(1);
        if (word == "dbopendynaset") return VbaValue::numeric(2);
        if (word == "dbopensnapshot") return VbaValue::numeric(4);
        if (word == "acquitprompt") return VbaValue::integer(0);
        if (word == "acquitsaveall") return VbaValue::integer(1);
        if (word == "acquitsavenone") return VbaValue::integer(2);
        if (word == "dbeditnone") return VbaValue::integer(0);
        if (word == "dbeditinprogress") return VbaValue::integer(1);
        if (word == "dbeditadd") return VbaValue::integer(2);
        if (word == "dbfailonerror") return VbaValue::numeric(128);
        if (word == "dbforceosflush") return VbaValue::numeric(1);
        static const std::unordered_map<std::string,int> variantTypes{
            {"vbempty",0},{"vbnull",1},{"vbinteger",2},{"vblong",3},{"vbsingle",4},
            {"vbdouble",5},{"vbcurrency",6},{"vbdate",7},{"vbstring",8},{"vbobject",9},
            {"vberror",10},{"vbboolean",11},{"vbvariant",12},{"vbdataobject",13},
            {"vbdecimal",14},{"vbbyte",17},{"vbarray",8192}};
        if (const auto type=variantTypes.find(word);type!=variantTypes.end()) return VbaValue::numeric(type->second);
        static const std::unordered_map<std::string,int> calendarConstants{{"vbusesystem",0},{"vbsunday",1},{"vbmonday",2},{"vbtuesday",3},{"vbwednesday",4},{"vbthursday",5},{"vbfriday",6},{"vbsaturday",7},{"vbfirstjan1",1},{"vbfirstfourdays",2},{"vbfirstfullweek",3}};
        if(const auto constant=calendarConstants.find(word);constant!=calendarConstants.end()) return VbaValue::integer(constant->second);
        if (word == "true") return jetEval_ ? VbaValue::integer(-1) : VbaValue::logical(true);
        if (word == "false") return jetEval_ ? VbaValue::integer(0) : VbaValue::logical(false);
        static const std::unordered_map<std::string,int> messageConstants{
            {"vbokonly",0},{"vbokcancel",1},{"vbabortretryignore",2},{"vbyesnocancel",3},{"vbyesno",4},{"vbretrycancel",5},
            {"vbcritical",16},{"vbquestion",32},{"vbexclamation",48},{"vbinformation",64},
            {"vbdefaultbutton1",0},{"vbdefaultbutton2",256},{"vbdefaultbutton3",512},{"vbdefaultbutton4",768},
            {"vbapplicationmodal",0},{"vbsystemmodal",4096},{"vbmsgboxhelpbutton",16384},{"vbmsgboxsetforeground",65536},
            {"vbmsgboxright",524288},{"vbmsgboxrtlreading",1048576},
            {"vbok",1},{"vbcancel",2},{"vbabort",3},{"vbretry",4},{"vbignore",5},{"vbyes",6},{"vbno",7}};
        if (const auto constant=messageConstants.find(word);constant!=messageConstants.end()) return VbaValue::numeric(constant->second);
        if (word == "vbcrlf") return VbaValue::string("\r\n");
        if (word == "vbnewline") return VbaValue::string("\r\n");
        if (word == "vbtab") return VbaValue::string("\t");
        if (word == "vbcr") return VbaValue::string("\r");
        if (word == "vblf") return VbaValue::string("\n");
        if (word == "vbback") return VbaValue::string("\b");
        if (word == "vbformfeed") return VbaValue::string("\f");
        if (word == "vbverticaltab") return VbaValue::string("\v");
        if (word == "vbnullchar") return VbaValue::string(std::string(1,'\0'));
        if (word == "vbnullstring") return VbaValue::string("");
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
        if(word=="vbobjecterror") return VbaValue::longInteger(-2147221504);
        static const std::unordered_map<std::string,int> documentConstants{{"vbusecompareoption",-1},{"vbbinarycompare",0},{"vbtextcompare",1},{"vbdatabasecompare",2},{"accopy",2},{"accmdcut",189},{"accmdundo",292},{"accmdredo",199},{"accmdcopy",190},{"accmdpaste",191},{"accmdpasteappend",38},{"acformbar",0},{"aceditmenu",1},{"acmenuver70",70},{"acmenuver20",20},{"acmenuver1x",11},{"acselectrecord",8},{"acdelete",6},{"accmdselectrecord",50},{"accmdselectallrecords",109},{"accmddeleterecord",223},{"acdataerrcontinue",0},{"acdataerrdisplay",1},{"acdeleteok",0},{"acdeletecancel",1},{"acdeleteusercancel",2},{"accmdsortascending",163},{"accmdsortdescending",164},{"accmdsaverecord",97},{"accmdzoombox",179},{"accmdreplace",29},{"accmdfind",30},{"accmdfindnext",341},{"accmdpagesetup",32},{"accmdprint",340},{"accmdprintpreview",54},{"acanywhere",0},{"acentire",1},{"acstart",2},{"acsearchall",2},{"acdown",1},{"acup",0},{"acall",0},{"accurrent",-1},{"accmdrecordsgotofirst",67},{"accmdrecordsgotoprevious",66},{"accmdrecordsgotonext",65},{"accmdrecordsgotolast",68},{"accmdrecordsgotonew",28},{"acdataform",2},{"acdatatable",0},{"acdataquery",1},{"acdatareport",3},{"acactivedataobject",-1},{"acprevious",0},{"acfirst",2},{"acgoto",4},{"acnext",1},{"aclast",3},{"acnewrec",5},{"acdetail",0},{"acheader",1},{"acfooter",2},{"acpageheader",3},{"acpagefooter",4},{"acviewnormal",0},{"acviewpreview",2},{"acviewdesign",1},{"acwindownormal",0},{"acdialog",3},{"achidden",1},{"acicon",2},{"acedit",1},{"acreadonly",2},{"acadd",0},{"acformpropertysettings",-1},{"acformadd",0},{"acformedit",1},{"acformreadonly",2},{"acnormal",0},{"acdesign",1},{"acpreview",2},{"acformds",3},{"accustomcontrol",119},{"acshiftmask",1},{"acctrlmask",2},{"acaltmask",4},{"acdefault",-1},{"actable",0},{"acquery",1},{"acform",2},{"acreport",3},{"acmacro",4},{"acmodule",5},{"acsaveprompt",0},{"acsaveyes",1},{"acsaveno",2}};
        if (const auto constant=documentConstants.find(word);constant!=documentConstants.end()) return VbaValue::numeric(constant->second);
        if (word=="screen") return host_ ? host_->applicationObject()->member("Screen",{}) : VbaValue{};
        if (word=="docmd") return host_ ? host_->applicationObject()->member("DoCmd",{}) : VbaValue{};
        if((word=="currentproject" || word=="currentdata") && !nextSymbol("(")) {
            if(const auto local=locals_.find(word);local!=locals_.end()) return local->second;
            if(!host_) return {};
            if(!accessExpression_) if(auto resolved=host_->resolveIdentifier(tok.value,true)) return *resolved;
            return host_->applicationObject()->member(word=="currentproject" ? "CurrentProject" : "CurrentData",{});
        }
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
            if (!item && (lower(field.value)=="controls" || lower(field.value)=="section")) {
                std::vector<VbaValue> args;
                if (take("(")) { args.push_back(expression(1)); require(")"); }
                return host_ ? host_->currentForm()->member(field.value,args) : VbaValue{};
            }
            const auto key = lower(field.value);
            if (!item && (key=="requery" || key=="recalc" || key=="undo" || key=="setfocus")) {
                if (take("(")) require(")");
                return host_ ? host_->currentForm()->member(field.value,{}) : VbaValue{};
            }
            if (!host_) return {};
            const bool formProperty = !item && (key == "name" || key == "dirty" || key == "newrecord" || key == "recordsource" || key == "caption" || key=="visible" || key=="tag" || key=="filter" || key=="filteron" || key=="keypreview" || key=="bookmark" || key=="recordset" || key=="openargs" || key=="modal" || key=="popup");
            if (!item && !formProperty)
                if (auto value=host_->resolveIdentifier(field.value,objectContext || nextSymbol(".") || nextSymbol("!"))) return *value;
            if (!formProperty && objectContext)
                return item ? host_->currentForm()->objectItem(field.value) : host_->currentForm()->objectMember(field.value,{});
            if (!formProperty && (nextSymbol(".") || nextSymbol("!")))
                return host_->currentForm()->member("Controls",{VbaValue::string(field.value)});
            return item ? host_->controlValue(field.value) : host_->fieldValue(field.value);
        }
        const auto currentDirectory=[&]() -> VbaValue {
            if(!host_) return {};
            std::error_code failure;
            const auto path=std::filesystem::current_path(failure);
            if(failure) {
                if(failure==std::errc::permission_denied) throw error("Current directory access denied",70);
                if(failure==std::errc::no_such_file_or_directory) throw error("Current directory not found",76);
                throw error("Current directory unavailable",75);
            }
            const auto text=path.string(); (void)utf16String(text);
            return VbaValue::string(text);
        };
        if((word=="curdir" || word=="curdir$") && !nextSymbol("(") && !accessExpression_) {
            if(!host_) return {};
            if(!qualifiedIntrinsic) {
                if(const auto variable=locals_.find(word); variable!=locals_.end()) return variable->second;
                if(resolver_.argumentSignature) {
                    const auto signature=resolver_.argumentSignature(tok.value);
                    if(signature && signature->localProcedure) {
                        if(!signature->function) throw error("A Sub cannot be used as a function expression",3251);
                        std::vector<VbaValue> arguments;
                        return resolver_(tok.value,arguments,VbaArgumentReferences{});
                    }
                }
                if(auto value=host_->resolveIdentifier(tok.value,objectContext || nextSymbol(".") || nextSymbol("!"))) return *value;
            }
            return currentDirectory();
        }
        if(word=="doevents" && !nextSymbol("(")) {
            if(!host_) return {};
            if(accessExpression_) throw error("DoEvents is unavailable in calculated expressions",3251);
            return VbaValue::integer(host_->doEvents());
        }
        if ((word=="now" || word=="date" || word=="time" || word=="timer") && !nextSymbol("("))
            return host_ ? clockValue(word) : VbaValue{};
        if (accessExpression_ && (word=="sum" || word=="count" || word=="avg" || word=="min" || word=="max" || word=="first" || word=="last" || word=="var" || word=="stdev") && take("(")) {
            const auto start=index_;
            auto end=start; int depth=1;
            for (;end<tokens_.size() && tokens_[end].type!=Token::Type::End;++end) {
                if (tokens_[end].type!=Token::Type::Symbol) continue;
                if (tokens_[end].value=="(") ++depth;
                else if (tokens_[end].value==")" && --depth==0) break;
            }
            if (depth || end==start) throw error("Aggregate argument is missing or unclosed",450);
            index_=end+1;
            const auto& field=tokens_[start];
            const auto fieldWord=lower(field.value);
            if (end==start+1 && ((field.type==Token::Type::Word && fieldWord!="null"
                    && fieldWord!="true" && fieldWord!="false" && fieldWord!="empty")
                || field.type==Token::Type::Bracket || (field.type==Token::Type::Symbol && field.value=="*")))
                return host_ ? host_->expressionAggregate(word,field.value,field.type==Token::Type::Symbol && field.value=="*") : VbaValue{};
            const auto body=source_.substr(field.offset,tokens_[end].offset-field.offset);
            if (!host_) { (void)ExpressionParser(body,nullptr,locals_,{},false,true).parse(); return {}; }
            return host_->expressionAggregateExpression(word,body);
        }
        if (nextSymbol("(")) {
            const auto variable=locals_.find(word);
            if (variable!=locals_.end() && (variable->second.kind==VbaValue::Kind::Object || isVbaArray(variable->second))) return variable->second;
        }
        if (take("(")) {
            std::optional<VbaArgumentSignature> signature;
            if(host_ && resolver_.argumentSignature && !accessExpression_ && !qualifiedIntrinsic) signature=resolver_.argumentSignature(tok.value);
            bool localProcedure=signature && signature->localProcedure;
            if(localProcedure && !signature->function) throw error("A Sub cannot be used as a function expression",3251);
            std::vector<VbaValue> args;
            std::vector<std::string> references;
            std::vector<std::optional<VbaArrayElementReference>> arrayReferences;
            std::vector<bool> omitted;
            if (!take(")")) {
                do {
                    std::string argumentName;
                    if(peek().type==Token::Type::Word && index_+1<tokens_.size() && tokens_[index_+1].value==":=") {
                        argumentName=lower(consume().value); consume();
                    }
                    const auto start = index_;
                    const bool builtinDefaults=!localProcedure && (word=="msgbox" || word=="inputbox" || word=="inputbox$" || word=="instr" || word=="instrrev" || word=="strcomp" || word=="round" || word=="replace" || word=="split" || word=="join" || word=="filter" || word=="datediff" || word=="datepart" || word=="weekday");
                    const bool missing=nextSymbol(",") || nextSymbol(")") || (accessExpression_ && nextSymbol(";"));
                    auto argument=missing ? VbaEvaluatedArgument{builtinDefaults ? VbaValue::empty() : VbaValue::missing(),{}}
                        : argumentExpression((!localProcedure && word=="isobject") || (signature && signature->objectArgument(args.size(),argumentName)));
                    args.push_back(std::move(argument.value)); arrayReferences.push_back(std::move(argument.arrayElement));
                    if(!argumentName.empty() && missing) throw error("Named argument requires a value",450);
                    omitted.push_back(missing);
                    references.push_back(!missing && index_ == start + 1 && tokens_[start].type == Token::Type::Word
                        ? tokens_[start].value : "");
                    if(host_ && index_==start+3 && tokens_[start].type==Token::Type::Word && tokens_[start+1].value=="(" && tokens_[start+2].value==")") {
                        const auto variable=locals_.find(lower(tokens_[start].value));
                        if(variable!=locals_.end() && isVbaArray(variable->second)) references.back()=tokens_[start].value;
                    }
                    if(host_ && index_>start+3 && tokens_[start].type==Token::Type::Word && tokens_[start+1].value=="(" && tokens_[index_-1].value==")") {
                        const auto variable=locals_.find(lower(tokens_[start].value));
                        int depth=0; bool whole=true;
                        for(auto cursor=start+1;cursor<index_;++cursor) {
                            if(tokens_[cursor].value=="(") ++depth;
                            else if(tokens_[cursor].value==")" && --depth==0 && cursor+1<index_) whole=false;
                        }
                        if(whole && variable!=locals_.end() && isVbaArray(variable->second)) references.back()="@array_element";
                    }
                    if(!argumentName.empty()) references.back()="@named:"+argumentName+":"+references.back();
                } while (take(",") || (accessExpression_ && take(";")));
                require(")");
            }
            // MS-VBAL default binding has a narrow Left exception: a scalar
            // zero-argument function does not hide the two-argument library call.
            if(localProcedure && word=="left" && signature->names.empty() && args.size()==2
                && !signature->resultType.empty() && signature->resultType!="variant" && nativeObjectType(signature->resultType).empty()) localProcedure=false;
            if(localProcedure) return resolver_(tok.value,args,VbaArgumentReferences{std::move(references),std::move(arrayReferences)});
            const bool named=std::any_of(references.begin(),references.end(),[](const std::string& reference){ return reference.starts_with("@named:"); });
            if(named) {
                if(!host_) return {};
                if(qualifiedArray) throw error("Named qualified Array arguments are not implemented",3251);
                if(accessExpression_) throw error("Named arguments are not implemented in Access expression context",3251);
                static const std::unordered_map<std::string,VbaArgumentSignature> signatures{
                    {"syscmd",{{"Action","Argument2","Argument3"},1}},
                    {"msgbox",{{"Prompt","Buttons","Title","HelpFile","Context"},1}},
                    {"inputbox",{{"Prompt","Title","Default","XPos","YPos","HelpFile","Context"},1}},
                    {"inputbox$",{{"Prompt","Title","Default","XPos","YPos","HelpFile","Context"},1}},
                    {"sqr",{{"Number"},1}}, {"exp",{{"Number"},1}}, {"log",{{"Number"},1}},
                    {"sin",{{"Number"},1}}, {"cos",{{"Number"},1}}, {"tan",{{"Number"},1}}, {"atn",{{"Number"},1}},
                    {"lcase",{{"String"},1}}, {"lcase$",{{"String"},1}},
                    {"ucase",{{"String"},1}}, {"ucase$",{{"String"},1}},
                    {"val",{{"String"},1}},
                    {"createobject",{{"Class","ServerName"},1}},
                    {"filelen",{{"PathName"},1}}, {"filedatetime",{{"PathName"},1}},
                    {"curdir",{{"Drive"},0}}, {"curdir$",{{"Drive"},0}},
                    {"hex",{{"Number"},1}}, {"hex$",{{"Number"},1}},
                    {"oct",{{"Number"},1}}, {"oct$",{{"Number"},1}},
                    {"space",{{"Number"},1}}, {"space$",{{"Number"},1}},
                    {"string",{{"Number","Character"},2}}, {"string$",{{"Number","Character"},2}},
                    {"asc",{{"String"},1}}, {"ascw",{{"String"},1}},
                    {"chr",{{"CharCode"},1}}, {"chr$",{{"CharCode"},1}},
                    {"chrw",{{"CharCode"},1}}, {"chrw$",{{"CharCode"},1}},
                    {"replace",{{"Expression","Find","Replace","Start","Count","Compare"},3}},
                    {"split",{{"Expression","Delimiter","Limit","Compare"},1}},
                    {"join",{{"SourceArray","Delimiter"},1}},
                    {"filter",{{"SourceArray","Match","Include","Compare"},2}},
                    {"instrrev",{{"StringCheck","StringMatch","Start","Compare"},2}},
                    {"strreverse",{{"Expression"},1}},
                    {"instr",{{"Start","String1","String2","Compare"},0}},
                    {"strcomp",{{"String1","String2","Compare"},2}},
                    {"pv",{{"Rate","NPer","Pmt","Fv","Type"},3}},
                    {"nper",{{"Rate","Pmt","Pv","Fv","Type"},3}},
                    {"rate",{{"NPer","Pmt","Pv","Fv","Type","Guess"},3}},
                    {"npv",{{"Rate","Values"},2}},
                    {"irr",{{"Values","Guess"},1}},
                    {"mirr",{{"Values","Finance_rate","Reinvest_rate"},3}},
                    {"sln",{{"Cost","Salvage","Life"},3}},
                    {"syd",{{"Cost","Salvage","Life","Period"},4}},
                    {"ddb",{{"Cost","Salvage","Life","Period","Factor"},4}},
                    {"fv",{{"Rate","NPer","Pmt","Pv","Type"},3}},
                    {"pmt",{{"Rate","NPer","Pv","Fv","Type"},3}},
                    {"ipmt",{{"Rate","Per","NPer","Pv","Fv","Type"},4}},
                    {"ppmt",{{"Rate","Per","NPer","Pv","Fv","Type"},4}},
                    {"partition",{{"Number","Start","Stop","Interval"},4}},
                    {"rgb",{{"Red","Green","Blue"},3}}, {"qbcolor",{{"Color"},1}},
                    {"cdate",{{"Expression"},1}}, {"cdec",{{"Expression"},1}},
                    {"formatnumber",{{"Expression","NumDigitsAfterDecimal","IncludeLeadingDigit","UseParensForNegativeNumbers","GroupDigits"},1}},
                    {"formatcurrency",{{"Expression","NumDigitsAfterDecimal","IncludeLeadingDigit","UseParensForNegativeNumbers","GroupDigits"},1}},
                    {"formatpercent",{{"Expression","NumDigitsAfterDecimal","IncludeLeadingDigit","UseParensForNegativeNumbers","GroupDigits"},1}},
                    {"datevalue",{{"Date"},1}}, {"timevalue",{{"Time"},1}},
                    {"dateadd",{{"Interval","Number","Date"},3}},
                    {"dateserial",{{"Year","Month","Day"},3}},
                    {"timeserial",{{"Hour","Minute","Second"},3}},
                    {"datepart",{{"Interval","Date","FirstDayOfWeek","FirstWeekOfYear"},2}},
                    {"datediff",{{"Interval","Date1","Date2","FirstDayOfWeek","FirstWeekOfYear"},3}},
                    {"weekday",{{"Date","FirstDayOfWeek"},1}},
                    {"left",{{"String","Length"},2}}, {"left$",{{"String","Length"},2}},
                    {"right",{{"String","Length"},2}}, {"right$",{{"String","Length"},2}},
                    {"mid",{{"String","Start","Length"},2}}, {"mid$",{{"String","Start","Length"},2}},
                    {"nz",{{"Value","ValueIfNull"},1}}, {"ismissing",{{"ArgName"},1}},
                    {"dsum",{{"Expr","Domain","Criteria"},2}}, {"dcount",{{"Expr","Domain","Criteria"},2}}, {"dlookup",{{"Expr","Domain","Criteria"},2}}, {"dmin",{{"Expr","Domain","Criteria"},2}}, {"dmax",{{"Expr","Domain","Criteria"},2}}, {"davg",{{"Expr","Domain","Criteria"},2}}, {"dfirst",{{"Expr","Domain","Criteria"},2}}, {"dlast",{{"Expr","Domain","Criteria"},2}}, {"dvar",{{"Expr","Domain","Criteria"},2}}, {"dstdev",{{"Expr","Domain","Criteria"},2}}
                };
                const auto signature=signatures.find(word);
                if(signature==signatures.end()) {
                    if(qualifiedIntrinsic || !resolver_) throw error("Named builtin/procedure arguments are not implemented",3251);
                    return resolver_(word,args,VbaArgumentReferences{std::move(references),std::move(arrayReferences)});
                }
                std::vector<VbaValue> ordered(signature->second.names.size(),VbaValue::empty());
                std::vector<bool> assigned(ordered.size()),missing(ordered.size(),true);
                std::vector<std::string> orderedRefs(ordered.size());
                bool namedSeen=false; std::size_t last=0;
                for(std::size_t index=0;index<args.size();++index) {
                    std::size_t target=index; auto reference=references[index];
                    if(reference.starts_with("@named:")) {
                        namedSeen=true; const auto separator=reference.find(':',7);
                        const auto argument=reference.substr(7,separator-7); target=ordered.size();
                        for(std::size_t parameter=0;parameter<signature->second.names.size();++parameter)
                            if(lower(signature->second.names[parameter])==argument) { target=parameter; break; }
                        if(target==ordered.size()) throw error("Named builtin argument not found",448);
                        reference=reference.substr(separator+1);
                    } else if(namedSeen) throw error("Positional argument after named builtin argument",450);
                    if(target>=ordered.size() || assigned[target]) throw error("Builtin argument supplied more than once or too many arguments",450);
                    assigned[target]=true; missing[target]=omitted[index]; ordered[target]=args[index]; orderedRefs[target]=std::move(reference); last=std::max(last,target);
                }
                for(std::size_t index=0;index<signature->second.required;++index) if(missing[index]) throw error("Required builtin argument omitted",450);
                // InStr has two positional overloads; retain its full four-slot shape.
                const auto size=word=="instr" ? ordered.size() : last+1;
                ordered.resize(size); missing.resize(size); orderedRefs.resize(size);
                args=std::move(ordered); omitted=std::move(missing); references=std::move(orderedRefs);
            }
            if(word=="syscmd") {
                if(!host_) return {};
                if(accessExpression_) throw error("SysCmd is unavailable in calculated expressions",3251);
                auto values=args;
                for(std::size_t i=0;i<values.size();++i) if(omitted[i]) values[i]=VbaValue::missing();
                return host_->applicationObject()->member("SysCmd",values);
            }
            if(word=="doevents") {
                if(!args.empty()) throw error("DoEvents takes no arguments",450);
                if(!host_) return {};
                if(accessExpression_) throw error("DoEvents is unavailable in calculated expressions",3251);
                return VbaValue::integer(host_->doEvents());
            }
            const auto one = args.size() == 1;
            if(word=="cdate") {
                if(!host_) return {};
                if(args.size()!=1 || omitted[0]) throw error("CDate requires one argument",450);
                return vbaDateValue(args[0]);
            }
            if(word=="datevalue" || word=="timevalue" || word=="dateadd" || word=="dateserial" || word=="timeserial" || word=="datepart" || word=="datediff" || word=="year" || word=="month" || word=="day" || word=="hour" || word=="minute" || word=="second" || word=="weekday") {
                if(!host_) return {};
                for(std::size_t i=0;i<omitted.size();++i) if(omitted[i]) {
                    if(i<(word=="dateadd" || word=="datediff" || word=="dateserial" || word=="timeserial"?3u:word=="datepart"?2u:1u)) throw error("Required calendar argument omitted",450);
                    args[i]=VbaValue::integer(1);
                }
                return vbaCalendarFunction(word,args);
            }
            if (word=="isobject") {
                if (!host_) return {};
                if(args.size()!=1) throw error("IsObject requires one argument",450);
                return VbaValue::logical(args[0].kind==VbaValue::Kind::Object);
            }
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

            if (word == "now" || word == "date" || word == "time" || word == "timer") {
                if (!host_) return {};
                if (!args.empty()) throw error("Clock function takes no arguments",450);
                return clockValue(word);
            }
            if (word=="str" || word=="str$") {
                if (!host_) return {};
                if (args.size()!=1) throw error("Str requires one argument",450);
                const auto value=scalarObjectValue(args[0]);
                if (value.kind==VbaValue::Kind::Null) throw error("Invalid Null Str argument",94);
                double number;
                try { number=value.asNumber(); }
                catch (const VbaRuntimeError& failure) { if (failure.number==5) throw error("Str type mismatch",13); throw; }
                if (!std::isfinite(number)) throw error("Str numeric overflow",6);
                auto result=(value.kind==VbaValue::Kind::Currency || value.kind==VbaValue::Kind::Decimal) ? value.asText() : VbaValue::numeric(number==0 ? 0 : number).asText();
                if (result.starts_with("0.")) result.erase(0,1);
                else if (result.starts_with("-0.")) result.erase(1,1);
                std::replace(result.begin(),result.end(),'e','E');
                if (number>=0) result.insert(result.begin(),' ');
                return VbaValue::string(std::move(result));
            }
            if(word=="createobject") {
                if(!host_) return {};
                if(args.empty() || args.size()>2 || omitted[0]) throw error("CreateObject requires a class",450);
                if(accessExpression_) throw error("CreateObject is unavailable in calculated expressions",3251);
                const auto name=coerceVbaModuleValue(scalarObjectValue(args[0]),"string").asText();
                if(args.size()>1 && !omitted[1] && !coerceVbaModuleValue(scalarObjectValue(args[1]),"string").asText().empty()) throw error("Remote object activation is not implemented",3251);
                const auto classId=lower(name);
                if(classId=="scripting.filesystemobject") return VbaValue::reference(std::make_shared<VbaFileSystemObject>());
                if(classId=="scripting.dictionary") return VbaValue::reference(std::make_shared<VbaDictionary>());
                throw error("Native object class is not implemented",3251);
            }
            if(word=="curdir" || word=="curdir$") {
                if(!host_) return {};
                if(args.size()>1) throw error("CurDir accepts one optional drive",450);
                if(!args.empty() && !omitted[0]) {
                    const auto drive=coerceVbaModuleValue(scalarObjectValue(args[0]),"string").asText();
                    (void)utf16String(drive);
                    if(!drive.empty()) throw error("Windows drive directories require an unavailable drive adapter",3251);
                }
                return currentDirectory();
            }
            if(word=="filelen" || word=="filedatetime") {
                if(!host_) return {};
                if(args.size()!=1 || omitted[0]) throw error("File metadata requires one pathname",450);
                const auto path=scalarObjectValue(args[0]);
                const auto text=coerceVbaModuleValue(path,"string"); (void)utf16String(text.asText());
                return nativeFileInformation(text,word=="filedatetime");
            }
            if(word=="array") {
                if(!host_) return {};
                if(std::find(omitted.begin(),omitted.end(),true)!=omitted.end()) throw error("Array elements cannot be omitted",450);
                return makeVbaArray(std::move(args),qualifiedArray ? 0 : arrayBase_);
            }
            if(word=="lbound" || word=="ubound") {
                if(!host_) return {};
                if(args.empty() || args.size()>2 || omitted[0]) throw error("Array bound requires one or two arguments",450);
                const int dimension=args.size()==2 && !omitted[1] ? static_cast<int>(coerceVbaModuleValue(scalarVbaValue(args[1]),"long").number) : 1;
                return vbaArrayBound(args[0],word=="ubound",dimension);
            }
            if (word=="isarray") {
                if (!host_) return {};
                if (args.size()!=1) throw error("IsArray requires one argument",450);
                return VbaValue::logical(isVbaArray(args[0]));
            }
            if(word=="round") {
                if(!host_) return {};
                if(args.empty() || omitted[0]) throw error("Round requires an expression",450);
                return vbaRound(args);
            }
            if(word=="hex" || word=="hex$" || word=="oct" || word=="oct$") {
                if(!host_) return {};
                if(args.size()!=1 || omitted[0]) throw error("Base conversion requires one argument",450);
                const auto value=scalarObjectValue(args[0]);
                if(value.kind==VbaValue::Kind::Null) {
                    if(word.back()=='$') throw error("Invalid Null base conversion argument",94);
                    return {};
                }
                auto bits=static_cast<std::uint32_t>(static_cast<std::int32_t>(coerceVbaModuleValue(value,"long").number));
                if(value.kind==VbaValue::Kind::Number && value.numericType==VbaValue::NumericType::Integer) bits&=0xffffU;
                const auto base=word.starts_with("hex") ? 16U : 8U;
                std::string result;
                do { result.push_back("0123456789ABCDEF"[bits%base]); bits/=base; } while(bits);
                std::reverse(result.begin(),result.end());
                return VbaValue::string(std::move(result));
            }
            if(word=="space" || word=="space$" || word=="string" || word=="string$") {
                if(!host_) return {};
                const bool spaces=word=="space" || word=="space$";
                if(args.size()!=(spaces ? 1U : 2U) || std::any_of(omitted.begin(),omitted.end(),[](bool value){ return value; }))
                    throw error("Repetition requires all arguments",450);
                const auto number=scalarObjectValue(args[0]);
                const auto character=spaces ? VbaValue::string(" ") : scalarObjectValue(args[1]);
                if(number.kind==VbaValue::Kind::Null || character.kind==VbaValue::Kind::Null) {
                    if(spaces || word.back()=='$') throw error("Invalid Null repetition argument",94);
                    return {};
                }
                const auto count=static_cast<std::int32_t>(coerceVbaModuleValue(number,"long").number);
                if(count<0) throw error("Invalid repetition count",5);
                if(count>1000000) throw error("Native repetition size limit exceeded",3251);
                std::string unit;
                if(character.kind==VbaValue::Kind::Text) {
                    const auto text=utf16String(character.asText());
                    if(text.empty()) throw error("Empty repetition character",5);
                    unit=utf8String(std::u16string(1,text.front()));
                } else {
                    const auto code=static_cast<std::int32_t>(coerceVbaModuleValue(character,"long").number);
                    if(code<0) throw error("Invalid repetition character code",5);
                    const auto byte=code%256;
                    if(byte>127) throw error("ANSI character codepage conversion is not implemented",3251);
                    unit.assign(1,static_cast<char>(byte));
                }
                std::string result; result.reserve(static_cast<std::size_t>(count)*unit.size());
                for(std::int32_t index=0;index<count;++index) result+=unit;
                return VbaValue::string(std::move(result));
            }
            if(word=="asc" || word=="ascw" || word=="chr" || word=="chr$" || word=="chrw" || word=="chrw$") {
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
                if((word=="chr" || word=="chr$") && (code<0 || code>127)) throw error("ANSI character codepage conversion is not implemented",3251);
                const auto unit=static_cast<char16_t>(code<0 ? code+65536 : code);
                return VbaValue::string(utf8String(std::u16string(1,unit)));
            }
            if (word=="cbyte" || word=="cint" || word=="clng" || word=="csng" || word=="cdbl" || word=="cbool") {
                if (!host_) return {};
                if (args.size()!=1) throw error("Conversion requires one argument",450);
                const auto type=word=="cbyte" ? "byte" : word=="cint" ? "integer" : word=="clng" ? "long" : word=="csng" ? "single" : word=="cdbl" ? "double" : "boolean";
                return coerceVbaModuleValue(scalarObjectValue(args.front()),type);
            }
            if (word=="ccur" || word=="cdec" || word=="vartype" || word=="typename") {
                if (!host_) return {};
                if (args.size()!=1) throw error("Conversion/type function requires one argument",450);
                const auto value=word=="vartype" ? typeInspectionValue(args.front()) : args.front();
                if (word=="ccur") return convertCurrency(value);
                if(word=="cdec" && jetEval_) throw error("CDec is unavailable in the observed Access Eval context",2439);
                if (word=="cdec") return convertDecimal(value);
                const int type=value.kind==VbaValue::Kind::Empty ? 0 : value.kind==VbaValue::Kind::Null ? 1
                    : value.kind==VbaValue::Kind::Number ? (value.numericType==VbaValue::NumericType::Byte ? 17 : value.numericType==VbaValue::NumericType::Integer ? 2 : value.numericType==VbaValue::NumericType::Long ? 3 : value.numericType==VbaValue::NumericType::Single ? 4 : 5) : value.kind==VbaValue::Kind::Currency ? 6
                    : value.kind==VbaValue::Kind::Decimal ? 14 : value.kind==VbaValue::Kind::Date ? 7 : value.kind==VbaValue::Kind::Text ? 8 : value.kind==VbaValue::Kind::Object ? 9
                    : (value.kind==VbaValue::Kind::Error || value.kind==VbaValue::Kind::Missing) ? 10 : value.kind==VbaValue::Kind::ByteArray ? 8209 : value.kind==VbaValue::Kind::Array ? vbaArrayType(value) : 11;
                if (word=="vartype") return jetEval_ ? VbaValue::longInteger(type) : VbaValue::integer(static_cast<std::int16_t>(type));
                if(isVbaArray(value)) return VbaValue::string(vbaArrayTypeName(value));
                if (type==9) {
                    if (!value.object) return VbaValue::string("Nothing");
                    if(!value.object->className().empty()) return VbaValue::string(std::string(value.object->className()));
                    throw error("Native object class name unavailable",3251);
                }
                return VbaValue::string(type==0 ? "Empty" : type==1 ? "Null" : type==5 ? "Double"
                    : type==17 ? "Byte" : type==2 ? "Integer" : type==3 ? "Long" : type==4 ? "Single" : type==6 ? "Currency" : type==14 ? "Decimal" : type==7 ? "Date" : type==8 ? "String" : type==9 ? "Object" : type==10 ? "Error" : type==8209 ? "Byte()" : type==8204 ? "Variant()" : type==8200 ? "String()" : "Boolean");
            }
            if (word=="dsum" || word=="dcount" || word=="dlookup" || word=="dmin" || word=="dmax" || word=="davg" || word=="dfirst" || word=="dlast" || word=="dvar" || word=="dstdev") {
                if (!host_) return {};
                if (args.size()<2 || args.size()>3) throw error("Domain aggregate requires two or three arguments",450);
                return host_->domainAggregate(word,args);
            }
            if(word=="ismissing") {
                if(!host_) return {};
                if(args.size()!=1) throw error("IsMissing requires one argument",450);
                return VbaValue::logical(args[0].kind==VbaValue::Kind::Missing);
            }
            if (word=="iserror" || word=="cverr") {
                if (!host_) return {};
                if (args.size()!=1) throw error("Error function requires one argument",450);
                if (word=="iserror") return VbaValue::logical(args[0].kind==VbaValue::Kind::Error || args[0].kind==VbaValue::Kind::Missing);
                if (args[0].kind==VbaValue::Kind::Null) throw error("Invalid Null CVErr argument",94);
                const auto number=integer(args[0]);
                if (number<0 || number>65535) throw error("CVErr number out of range");
                return VbaValue::errorValue(number);
            }
            if(word=="formatpercent") {
                if(!host_) return {};
                auto values=args;
                for(std::size_t i=0;i<values.size();++i) if(omitted[i]) values[i]=VbaValue::missing();
                const auto options=parseVbaNumericFormatOptions(values);
                const bool leading=options.leading.value_or(true),group=options.group.value_or(true);
                std::string mask=group ? (leading ? "#,##0" : "#,###") : (leading ? "0" : "#");
                const int digits=options.digits.value_or(2);
                if(digits) mask+='.'+std::string(digits,'0');
                mask+='%';
                if(options.parens.value_or(false)) mask+=';'+std::string("(")+mask+")";
                return host_->formatValue(values[0],mask,true);
            }
            if(word=="formatnumber" || word=="formatcurrency") {
                if(!host_) return {};auto values=args;for(std::size_t i=0;i<values.size();++i) if(omitted[i]) values[i]=VbaValue::missing();const auto options=parseVbaNumericFormatOptions(values);return host_->formatNumericValue(values[0],word=="formatcurrency",options);
            }
            if (word=="format" || word=="format$") {
                if (args.size()!=2) throw error("Format arguments unsupported",450);
                if (!host_) return {};
                if (args[1].kind==VbaValue::Kind::Null) throw error("Format mask is Null",94);
                if(jetEval_ && word=="format" && args[0].kind==VbaValue::Kind::Null) return VbaValue::string("");
                return host_->formatValue(args[0],args[1].asText(),word=="format$");
            }
            if(word=="filter") {
                if(!host_) return {};
                if(args.size()<2 || args.size()>4 || omitted[0] || omitted[1]) throw error("Filter requires two to four arguments",450);
                if(args[0].kind==VbaValue::Kind::Null) throw error("Invalid Null Filter array",94);
                if(args[0].kind!=VbaValue::Kind::Array || vbaArrayRank(args[0])!=1) throw error("Filter requires a one-dimensional String or Variant array",13);
                if(vbaArrayType(args[0])!=8200 && vbaArrayType(args[0])!=8204) throw error("Filter typed array element type is not implemented",3251);
                const auto match=scalarObjectValue(args[1]);
                if(match.kind==VbaValue::Kind::Null) throw error("Invalid Null Filter match",94);
                const auto needle=utf16String(match.asText());
                const bool include=args.size()>2 && !omitted[2] ? coerceVbaModuleValue(scalarObjectValue(args[2]),"boolean").boolean : true;
                auto mode=VbaCompareMode::Binary;
                if(args.size()>3 && !omitted[3]) {
                    const auto choice=integer(scalarObjectValue(args[3]));
                    if(choice==-1) mode=compareMode_;
                    else if(choice>=0 && choice<=2) mode=static_cast<VbaCompareMode>(choice);
                    else throw error("Invalid Filter comparison",5);
                }
                std::vector<std::string> selected;
                for(const auto& value:vbaArrayValues(args[0])) {
                    if(value.kind==VbaValue::Kind::Null) throw error("Invalid Null Filter element",94);
                    if(value.kind!=VbaValue::Kind::Text) throw error("Filter currently requires text elements",3251);
                    const auto input=utf16String(value.text);
                    bool found=needle.empty();
                    for(std::size_t position=0; !found && position<=input.size() && needle.size()<=input.size()-position; ++position) {
                        const auto boundary=[&](std::size_t index) {
                            return index==0 || index==input.size() || !(input[index-1]>=0xd800 && input[index-1]<=0xdbff && input[index]>=0xdc00 && input[index]<=0xdfff);
                        };
                        const auto candidate=std::u16string_view(input).substr(position,needle.size());
                        found=mode==VbaCompareMode::Binary ? candidate==needle : boundary(position) && boundary(position+needle.size()) && host_->compareStrings(utf8String(candidate),match.asText(),mode)==0;
                    }
                    if(found==include) selected.push_back(value.text);
                }
                return makeVbaStringArray(std::move(selected));
            }
            if(word=="split" || word=="join") {
                if(!host_) return {};
                const bool split=word=="split";
                if(args.empty() || args.size()>(split ? 4u : 2u) || omitted[0]) throw error("Invalid text array function arguments",450);
                std::vector<VbaValue> values; values.reserve(args.size());
                for(std::size_t i=0;i<args.size();++i) {
                    values.push_back(!split && i==0 ? args[i] : scalarObjectValue(args[i]));
                    if(!omitted[i] && values.back().kind==VbaValue::Kind::Null) throw error("Invalid Null text array argument",94);
                }
                const auto delimiter=values.size()>1 && !omitted[1] ? values[1].asText() : " ";
                if(!split) {
                    if(values[0].kind!=VbaValue::Kind::Array || vbaArrayRank(values[0])!=1) throw error("Join requires a one-dimensional Variant or String array",13);
                    if(vbaArrayType(values[0])!=8200 && vbaArrayType(values[0])!=8204) throw error("Join typed array element type is not implemented",3251);
                    (void)utf16String(delimiter);
                    std::string output; bool first=true;
                    for(const auto& item:vbaArrayValues(values[0])) {
                        const auto value=scalarObjectValue(item);
                        if(value.kind==VbaValue::Kind::Null) throw error("Invalid Null Join element",94);
                        if(!first) output+=delimiter;
                        const auto text=value.asText(); (void)utf16String(text);
                        output+=text; first=false;
                    }
                    return VbaValue::string(std::move(output));
                }
                const auto limit=values.size()>2 && !omitted[2] ? integer(values[2]) : -1;
                if(limit< -1) throw error("Invalid Split limit",5);
                auto mode=VbaCompareMode::Binary;
                if(values.size()>3 && !omitted[3]) {
                    const auto choice=integer(values[3]);
                    if(choice==-1) mode=compareMode_;
                    else if(choice>=0 && choice<=2) mode=static_cast<VbaCompareMode>(choice);
                    else throw error("Invalid Split comparison",5);
                }
                const auto input=utf16String(values[0].asText()),needle=utf16String(delimiter);
                std::vector<std::string> parts;
                if(limit==0 || input.empty()) return makeVbaStringArray(std::move(parts));
                std::size_t begin=0,position=0;
                const auto boundary=[&](std::size_t index) {
                    return index==0 || index==input.size() || !(input[index-1]>=0xd800 && input[index-1]<=0xdbff && input[index]>=0xdc00 && input[index]<=0xdfff);
                };
                while(!needle.empty() && needle.size()<=input.size()-position && (limit==-1 || parts.size()+1<static_cast<std::size_t>(limit))) {
                    const auto candidate=std::u16string_view(input).substr(position,needle.size());
                    const bool match=mode==VbaCompareMode::Binary ? candidate==needle
                        : boundary(position) && boundary(position+needle.size()) && host_->compareStrings(utf8String(candidate),delimiter,mode)==0;
                    if(match) {
                        parts.push_back(utf8String(std::u16string_view(input).substr(begin,position-begin)));
                        position+=needle.size(); begin=position;
                    } else ++position;
                }
                parts.push_back(utf8String(std::u16string_view(input).substr(begin)));
                return makeVbaStringArray(std::move(parts));
            }
            if (word=="replace") {
                if (!host_) return {};
                if (args.size()<3 || args.size()>6 || omitted[0] || omitted[1] || omitted[2])
                    throw error("Replace requires three to six arguments",450);
                std::vector<VbaValue> values;
                for(std::size_t i=0;i<args.size();++i) {
                    values.push_back(scalarObjectValue(args[i]));
                    if(!omitted[i] && values.back().kind==VbaValue::Kind::Null) throw error("Invalid Null Replace argument",94);
                }
                const auto start=values.size()>3 && !omitted[3] ? integer(values[3]) : 1;
                const auto count=values.size()>4 && !omitted[4] ? integer(values[4]) : -1;
                if(start<1 || count<-1) throw error("Invalid Replace start/count",5);
                auto mode=VbaCompareMode::Binary;
                if(values.size()>5 && !omitted[5]) {
                    const auto choice=integer(values[5]);
                    if(choice==-1) mode=compareMode_;
                    else if(choice>=0 && choice<=2) mode=static_cast<VbaCompareMode>(choice);
                    else throw error("Invalid Replace comparison",5);
                }
                const auto input=utf16String(values[0].asText()),needle=utf16String(values[1].asText()),replacement=utf16String(values[2].asText());
                if(static_cast<std::size_t>(start)>input.size()) return VbaValue::string("");
                const auto boundary=[&](std::size_t position) {
                    return position==0 || position==input.size() || !(input[position-1]>=0xd800 && input[position-1]<=0xdbff && input[position]>=0xdc00 && input[position]<=0xdfff);
                };
                std::size_t position=start-1;
                if(!boundary(position)) throw error("Replace start splits a UTF-16 surrogate",3251);
                std::u16string output;
                std::size_t replaced=0;
                while(position<input.size()) {
                    bool match=false;
                    if(!needle.empty() && (count==-1 || replaced<static_cast<std::size_t>(count)) && needle.size()<=input.size()-position) {
                        const auto candidate=std::u16string_view(input).substr(position,needle.size());
                        match=mode==VbaCompareMode::Binary ? candidate==needle
                            : boundary(position) && boundary(position+needle.size()) && host_->compareStrings(utf8String(candidate),values[1].asText(),mode)==0;
                    }
                    if(match) { output.append(replacement); position+=needle.size(); ++replaced; }
                    else output.push_back(input[position++]);
                }
                return VbaValue::string(utf8String(output));
            }
            if(word=="strreverse" || word=="instrrev") {
                if(!host_) return {};
                const bool reverse=word=="strreverse";
                if(args.empty() || args.size()>(reverse ? 1u : 4u) || omitted[0] || (!reverse && (args.size()<2 || omitted[1])))
                    throw error("Invalid reverse string function arguments",450);
                std::vector<VbaValue> values;
                for(std::size_t i=0;i<args.size();++i) {
                    values.push_back(scalarObjectValue(args[i]));
                    if(!omitted[i] && values.back().kind==VbaValue::Kind::Null && (reverse || i>=2)) throw error("Invalid Null reverse string argument",94);
                }
                if(!reverse && (values[0].kind==VbaValue::Kind::Null || values[1].kind==VbaValue::Kind::Null)) return {};
                auto input=utf16String(values[0].asText());
                if(reverse) {
                    std::reverse(input.begin(),input.end());
                    return VbaValue::string(utf8String(input));
                }
                const auto needle=utf16String(values[1].asText());
                const auto start=values.size()>2 && !omitted[2] ? integer(values[2]) : -1;
                if(start==0 || start< -1) throw error("Invalid InStrRev start",5);
                auto mode=VbaCompareMode::Binary;
                if(values.size()>3 && !omitted[3]) {
                    const auto choice=integer(values[3]);
                    if(choice==-1) mode=compareMode_;
                    else if(choice>=0 && choice<=2) mode=static_cast<VbaCompareMode>(choice);
                    else throw error("Invalid InStrRev comparison",5);
                }
                if(input.size()>static_cast<std::size_t>(std::numeric_limits<std::int32_t>::max())) throw error("InStrRev string length overflow",6);
                const auto end=start==-1 ? input.size() : static_cast<std::size_t>(start);
                if(input.empty() || end>input.size()) return VbaValue::longInteger(0);
                if(needle.empty()) return VbaValue::longInteger(static_cast<std::int32_t>(end));
                if(needle.size()>end) return VbaValue::longInteger(0);
                const auto boundary=[&](std::size_t index) {
                    return index==0 || index==input.size() || !(input[index-1]>=0xd800 && input[index-1]<=0xdbff && input[index]>=0xdc00 && input[index]<=0xdfff);
                };
                for(auto position=static_cast<std::int64_t>(end-needle.size());position>=0;--position) {
                    const auto offset=static_cast<std::size_t>(position);
                    const auto candidate=std::u16string_view(input).substr(offset,needle.size());
                    const bool found=mode==VbaCompareMode::Binary ? candidate==needle : boundary(offset) && boundary(offset+needle.size()) && host_->compareStrings(utf8String(candidate),values[1].asText(),mode)==0;
                    if(found) return VbaValue::longInteger(static_cast<std::int32_t>(position+1));
                }
                return VbaValue::longInteger(0);
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
                    return VbaValue::integer(static_cast<std::int16_t>((result>0)-(result<0)));
                }
                const auto haystack=utf16String(left),needle=utf16String(right);
                if (haystack.empty()) return stringLongResult(0);
                if (needle.empty()) return stringLongResult(start);
                if (static_cast<std::size_t>(start)>haystack.size() || needle.size()>haystack.size()) return stringLongResult(0);
                if (mode==VbaCompareMode::Binary) {
                    const auto found=haystack.find(needle,static_cast<std::size_t>(start-1));
                    return stringLongResult(found==std::u16string::npos ? 0 : found+1);
                }
                const auto boundary=[&](std::size_t position) {
                    return position==0 || position==haystack.size() || !(haystack[position-1]>=0xd800 && haystack[position-1]<=0xdbff && haystack[position]>=0xdc00 && haystack[position]<=0xdfff);
                };
                for (std::size_t position=start-1;position+needle.size()<=haystack.size();++position) {
                    if (!boundary(position) || !boundary(position+needle.size())) continue;
                    const auto candidate=utf8String(std::u16string_view(haystack).substr(position,needle.size()));
                    if (host_->compareStrings(candidate,right,mode)==0) return stringLongResult(position+1);
                }
                return stringLongResult(0);
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
            if (word=="sln" || word=="syd" || word=="ddb") {
                if (!host_) return {};
                const std::size_t count=word=="sln" ? 3 : 4;
                if(args.size()<count || args.size()>(word=="ddb" ? 5u : count)
                   || std::any_of(omitted.begin(),omitted.begin()+std::min(count,omitted.size()),[](bool value){return value;}))
                    throw error("Depreciation arguments are required",450);
                long double values[5]{0,0,0,0,2};
                for(std::size_t i=0;i<args.size();++i) {
                    if(omitted[i]) continue;
                    const auto value=scalarObjectValue(args[i]);
                    if(value.kind==VbaValue::Kind::Null) throw error("Invalid Null depreciation argument",94);
                    values[i]=coerceVbaModuleValue(value,"double").number;
                    if(!std::isfinite(values[i])) throw error("Depreciation argument overflow",6);
                }
                const auto cost=values[0],salvage=values[1],life=values[2],period=values[3];
                if(cost<0 || salvage<0) throw error("Negative asset values outside supported scope",3251);
                if(life<=0 || (word!="sln" && (period<=0 || period>life))) throw error("Invalid depreciation lifetime or period",5);
                if(word=="ddb") {
                    const auto factor=values[4];
                    if(factor<=0) throw error("DDB requires positive factor",5);
                    if(salvage>cost || life>32767 || std::floor(life)!=life || std::floor(period)!=period)
                        throw error("DDB asset or fractional period range requires reference validation",3251);
                    const auto fraction=factor/life;
                    long double depreciation;
                    if(fraction>=1) depreciation=period==1 ? cost-salvage : 0;
                    else {
                        const auto balance=cost*std::exp((period-1)*std::log1p(-fraction));
                        depreciation=std::min(balance*fraction,std::max(0.0L,balance-salvage));
                    }
                    const auto result=static_cast<double>(depreciation);
                    if(!std::isfinite(result)) throw error("DDB result overflow",6);
                    return VbaValue::numeric(result);
                }
                const auto result=word=="sln" ? (cost-salvage)/life
                    : (cost-salvage)*(life-period+1)*2/(life*(life+1));
                const auto converted=static_cast<double>(result);
                if(!std::isfinite(converted)) throw error("Depreciation result overflow",6);
                return VbaValue::numeric(converted);
            }
            if(word=="irr") {
                if(!host_) return {};
                if(args.empty() || args.size()>2 || omitted[0]) throw error("IRR requires cash-flow array and optional guess",450);
                const auto array=scalarObjectValue(args[0]); const auto count=nativeFinancialArraySize(array);
                long double guess=0.1L;
                if(args.size()==2 && !omitted[1]) {
                    const auto value=scalarObjectValue(args[1]);
                    if(value.kind==VbaValue::Kind::Null) throw error("Invalid Null IRR guess",94);
                    guess=coerceVbaModuleValue(value,"double").number;
                    if(!std::isfinite(guess)) throw error("IRR guess overflow",6);
                }
                std::vector<double> flows; flows.reserve(count); bool positive=false,negative=false;
                for(std::size_t i=0;i<count;++i) {
                    const auto flow=coerceVbaModuleValue(vbaArrayLinearElement(array,i),"double").number;
                    if(!std::isfinite(flow)) throw error("IRR cash-flow overflow",6);
                    positive|=flow>0; negative|=flow<0; flows.push_back(flow);
                }
                if(!positive || !negative) throw error("IRR requires positive and negative flows",5);
                return VbaValue::numeric(nativeFinancialRoot(guess,[&](long double rate) {
                    const auto multiplier=1+rate;
                    long double balance=flows.front(),derivative=0,scale=std::abs(balance);
                    for(std::size_t i=1;i<flows.size();++i) {
                        derivative=derivative*multiplier+balance;
                        balance=balance*multiplier+flows[i];
                        scale=scale*multiplier+std::abs(flows[i]);
                    }
                    return FinancialBalance{balance,derivative,scale};
                }));
            }
            if(word=="mirr") {
                if(!host_) return {};
                if(args.size()!=3 || omitted[0] || omitted[1] || omitted[2])
                    throw error("MIRR requires cash flows, financing rate and reinvestment rate",450);
                const auto array=scalarObjectValue(args[0]); const auto count=nativeFinancialArraySize(array);
                const auto financeValue=scalarObjectValue(args[1]),reinvestValue=scalarObjectValue(args[2]);
                if(financeValue.kind==VbaValue::Kind::Null || reinvestValue.kind==VbaValue::Kind::Null)
                    throw error("Invalid Null MIRR rate",94);
                const long double finance=coerceVbaModuleValue(financeValue,"double").number;
                const long double reinvest=coerceVbaModuleValue(reinvestValue,"double").number;
                if(!std::isfinite(finance) || !std::isfinite(reinvest)) throw error("MIRR rate overflow",6);
                if(finance<=-1 || reinvest<=-1) throw error("MIRR rate outside supported range",3251);
                if(count<2) throw error("MIRR requires positive and negative flows",5);
                long double positiveFuture=0,negativePresent=0;
                for(std::size_t i=0;i<count;++i) {
                    const long double forward=coerceVbaModuleValue(vbaArrayLinearElement(array,i),"double").number;
                    const long double reverse=coerceVbaModuleValue(vbaArrayLinearElement(array,count-i-1),"double").number;
                    if(!std::isfinite(forward) || !std::isfinite(reverse)) throw error("MIRR cash-flow overflow",6);
                    positiveFuture=positiveFuture*(1+reinvest)+std::max(0.0L,forward);
                    negativePresent=negativePresent/(1+finance)+std::min(0.0L,reverse);
                    if(!std::isfinite(positiveFuture) || !std::isfinite(negativePresent)) throw error("MIRR intermediate overflow",6);
                }
                if(positiveFuture<=0 || negativePresent>=0) throw error("MIRR requires positive and negative flows",5);
                const auto result=static_cast<double>(std::expm1((std::log(positiveFuture)-std::log(-negativePresent))/(count-1)));
                if(!std::isfinite(result)) throw error("MIRR result overflow",6);
                return VbaValue::numeric(result);
            }
            if(word=="npv") {
                if(!host_) return {};
                if(args.size()!=2 || omitted[0] || omitted[1]) throw error("NPV requires rate and cash-flow array",450);
                const auto rateValue=scalarObjectValue(args[0]); const auto array=scalarObjectValue(args[1]);
                if(rateValue.kind==VbaValue::Kind::Null || array.kind==VbaValue::Kind::Null) throw error("Invalid Null NPV argument",94);
                const long double rate=coerceVbaModuleValue(rateValue,"double").number;
                if(!std::isfinite(rate)) throw error("NPV rate overflow",6);
                if(rate<=-1) throw error("NPV rate outside supported range",3251);
                const auto count=nativeFinancialArraySize(array);
                long double total=0; bool positive=false,negative=false;
                // Reverse Horner discounts the first flow at period end and
                // avoids an independent power/weight that could overflow on zero flows.
                for(auto i=count;i>0;--i) {
                    const auto flow=coerceVbaModuleValue(vbaArrayLinearElement(array,i-1),"double").number;
                    if(!std::isfinite(flow)) throw error("NPV cash-flow overflow",6);
                    positive|=flow>0; negative|=flow<0;
                    total=(total+static_cast<long double>(flow))/(1+rate);
                    if(!std::isfinite(total)) throw error("NPV intermediate overflow",6);
                }
                if(!positive || !negative) throw error("NPV single-sign array behavior requires reference validation",3251);
                const auto result=static_cast<double>(total);
                if(!std::isfinite(result)) throw error("NPV result overflow",6);
                return VbaValue::numeric(result);
            }
            if(word=="ipmt" || word=="ppmt") {
                if(!host_) return {};
                if(args.size()<4 || args.size()>6 || omitted[0] || omitted[1] || omitted[2] || omitted[3])
                    throw error("Interest/principal payment requires four to six arguments",450);
                long double values[6]{};
                for(std::size_t i=0;i<args.size();++i) {
                    if(omitted[i]) continue;
                    const auto value=scalarObjectValue(args[i]);
                    if(value.kind==VbaValue::Kind::Null) throw error("Invalid Null payment argument",94);
                    values[i]=coerceVbaModuleValue(value,"double").number;
                    if(!std::isfinite(values[i])) throw error("Payment argument overflow",6);
                }
                const auto rate=values[0],period=values[1],periods=values[2],present=values[3],future=values[4],timing=values[5];
                if(periods<=0 || period<1 || period>periods) throw error("Payment period out of range",5);
                if(periods>32767 || std::floor(periods)!=periods || std::floor(period)!=period || rate<=-1 || (timing!=0 && timing!=1))
                    throw error("Payment period, rate or timing outside supported range",3251);
                const auto factors=nativeFinancialAnnuity(rate,periods,timing);
                const auto payment=-(present*factors.growth+future)/factors.factor;
                if(!std::isfinite(static_cast<double>(payment))) throw error("Payment result overflow",6);
                long double balance=present;
                for(int i=1;i<static_cast<int>(period);++i) {
                    balance=timing==1 ? (balance+payment)*(1+rate) : balance*(1+rate)+payment;
                    if(!std::isfinite(balance)) throw error("Payment balance overflow",6);
                }
                const auto interest=timing==1 ? (period==1 ? 0.0L : -balance*rate/(1+rate)) : -balance*rate;
                const auto result=static_cast<double>(word=="ipmt" ? interest : payment-interest);
                if(!std::isfinite(result)) throw error("Payment component overflow",6);
                return VbaValue::numeric(result);
            }
            if (word=="fv" || word=="pmt" || word=="pv" || word=="nper" || word=="rate") {
                if (!host_) return {};
                if (args.size()<3 || args.size()>(word=="rate" ? 6u : 5u) || omitted[0] || omitted[1] || omitted[2])
                    throw error("Invalid financial argument count",450);
                long double values[6]{0,0,0,0,0,0.1L};
                for (std::size_t i=0;i<args.size();++i) {
                    if (omitted[i]) continue;
                    const auto value=scalarObjectValue(args[i]);
                    if (value.kind==VbaValue::Kind::Null) throw error("Invalid Null financial argument",94);
                    values[i]=coerceVbaModuleValue(value,"double").number;
                    if (!std::isfinite(values[i])) throw error("Financial argument overflow",6);
                }
                if(word=="rate") return VbaValue::numeric(nativePeriodicRate(values[0],values[1],values[2],values[3],values[4],values[5]));
                if(word=="nper") {
                    const auto rate=values[0],payment=values[1],present=values[2],future=values[3],timing=values[4];
                    if(rate<=-1 || (timing!=0 && timing!=1))
                        throw error("NPer rate or timing outside supported range",3251);
                    long double periods;
                    if(rate==0) {
                        if(payment==0) throw error("NPer has no unique solution",5);
                        periods=-(present+future)/payment;
                    } else {
                        const auto adjustedPayment=payment*(1+rate*timing);
                        const auto numerator=adjustedPayment-rate*future;
                        const auto denominator=rate*present+adjustedPayment;
                        if(!std::isfinite(numerator) || !std::isfinite(denominator)) throw error("NPer intermediate overflow",6);
                        if(numerator==0 || denominator==0 || std::signbit(numerator)!=std::signbit(denominator))
                            throw error("NPer has no finite real solution",5);
                        const auto delta=-rate*(present+future)/denominator;
                        const auto logarithm=std::abs(delta)<0.5L ? std::log1p(delta)
                            : std::log(std::abs(numerator))-std::log(std::abs(denominator));
                        periods=logarithm/std::log1p(rate);
                    }
                    const auto result=static_cast<double>(periods);
                    if(!std::isfinite(result)) throw error("NPer result overflow",6);
                    return VbaValue::numeric(result);
                }
                const auto rate=values[0],periods=values[1],timing=values[4];
                if (periods<0 || periods>32767 || std::floor(periods)!=periods)
                    throw error("Financial period range requires an unavailable compatibility oracle",3251);
                if (rate<=-1 || (timing!=0 && timing!=1))
                    throw error("Financial rate or payment timing outside supported range",3251);
                if (word=="pmt" && periods==0) throw error("Pmt requires nonzero periods",5);
                const auto factors=nativeFinancialAnnuity(rate,periods,timing);
                const auto growth=factors.growth,factor=factors.factor;
                const auto result=word=="fv" ? -(values[3]*growth+values[2]*factor)
                    : word=="pv" ? -(values[3]+values[2]*factor)/growth
                    : -(values[2]*growth+values[3])/factor;
                const auto converted=static_cast<double>(result);
                if (!std::isfinite(converted)) throw error("Financial result overflow",6);
                return VbaValue::numeric(converted);
            }
            if (word=="partition") {
                if (!host_) return {};
                if (args.size()!=4 || std::any_of(omitted.begin(),omitted.end(),[](bool value){return value;}))
                    throw error("Partition requires four arguments",450);
                std::vector<VbaValue> values;
                for (const auto& argument : args) values.push_back(scalarObjectValue(argument));
                for (const auto& value : values) if (value.kind==VbaValue::Kind::Null) return {};
                std::int64_t numbers[4];
                for (int i=0;i<4;++i) numbers[i]=static_cast<std::int64_t>(coerceVbaModuleValue(values[i],"long").number);
                const auto number=numbers[0],start=numbers[1],stop=numbers[2],interval=numbers[3];
                if (start<0 || stop<=start || interval<1) throw error("Invalid Partition range",5);
                // Keep boundary arithmetic inside the supported VBA Long domain.
                if (stop==2147483647) throw error("Partition boundary exceeds supported Long range",3251);
                const auto width=std::max(std::to_string(stop+1).size(),std::to_string(start-1).size());
                const auto padded=[&](std::int64_t value) {
                    const auto text=std::to_string(value);
                    return std::string(width>text.size() ? width-text.size() : 0,' ')+text;
                };
                if (number<start) return VbaValue::string(std::string(width,' ')+":"+padded(start-1));
                if (number>stop) return VbaValue::string(padded(stop+1)+":"+std::string(width,' '));
                const auto lower=((number-start)/interval)*interval+start;
                return VbaValue::string(padded(lower)+":"+padded(std::min(lower+interval-1,stop)));
            }
            if (word=="rgb" || word=="qbcolor") {
                if (!host_) return {};
                return word=="rgb" ? vbaRgb(args) : vbaQbColor(args);
            }
            if(word=="sqr" || word=="exp" || word=="log" || word=="sin" || word=="cos" || word=="tan" || word=="atn") {
                if(!host_) return {};
                if(args.size()!=1 || omitted[0]) throw error("Math function requires one argument",450);
                const auto value=scalarObjectValue(args[0]);
                if(value.kind==VbaValue::Kind::Null) throw error("Invalid Null math argument",94);
                const auto number=coerceVbaModuleValue(value,"double").number;
                if(!std::isfinite(number)) throw error("Non-finite math argument",6);
                if((word=="sqr" && number<0) || (word=="log" && number<=0)) throw error("Invalid math domain",5);
                const auto result=word=="sqr" ? std::sqrt(number) : word=="exp" ? std::exp(number)
                    : word=="log" ? std::log(number) : word=="sin" ? std::sin(number)
                    : word=="cos" ? std::cos(number) : word=="tan" ? std::tan(number) : std::atan(number);
                if(!std::isfinite(result)) throw error("Math result overflow",6);
                return VbaValue::numeric(result);
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
                if(value.kind==VbaValue::Kind::Decimal) return decimalUnary(word,value);
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
            if(word=="choose" || word=="switch") {
                if(!host_) return {};
                if(args.size()<2 || (word=="switch" && args.size()%2) || std::any_of(omitted.begin(),omitted.end(),[](bool missing){ return missing; }))
                    throw error("Invalid selection function arguments",450);
                if(jetEval_ && word=="switch" && std::any_of(args.begin(),args.end(),[](const VbaValue& value) {return value.kind==VbaValue::Kind::Error;})) throw error("Access Eval Switch cannot convert an Error Variant",6);
                if(word=="choose") {
                    const auto index=coerceVbaModuleValue(scalarObjectValue(args[0]),"long").number;
                    if(index<1 || index>=args.size()) return {};
                    return args[static_cast<std::size_t>(index)];
                }
                for(std::size_t index=0;index<args.size();index+=2) {
                    const auto condition=scalarObjectValue(args[index]);
                    if(condition.kind!=VbaValue::Kind::Null && coerceVbaModuleValue(condition,"boolean").boolean) return args[index+1];
                }
                return {};
            }
            if(word=="lcase" || word=="lcase$" || word=="ucase" || word=="ucase$") {
                if(!host_) return {};
                if(args.size()!=1 || omitted[0]) throw error("Case conversion requires one argument",450);
                const auto value=scalarObjectValue(args[0]);
                if(value.kind==VbaValue::Kind::Null) {
                    if(word.back()=='$') throw error("Invalid Null case argument",94);
                    return {};
                }
                auto text=value.asText(); (void)utf16String(text);
                const bool upper=word.starts_with("ucase");
                for(auto& ch:text) {
                    if(static_cast<unsigned char>(ch)>127) throw error("Non-ASCII case conversion requires a Windows locale adapter",3251);
                    if(upper && ch>='a' && ch<='z') ch-=('a'-'A');
                    else if(!upper && ch>='A' && ch<='Z') ch+=('a'-'A');
                }
                return VbaValue::string(std::move(text));
            }
            if(word=="val") {
                if(!host_) return {};
                if(args.size()!=1 || omitted[0]) throw error("Val requires one argument",450);
                const auto value=scalarObjectValue(args[0]);
                if(value.kind==VbaValue::Kind::Null) throw error("Invalid Null Val argument",94);
                auto text=value.asText();
                std::erase_if(text,[](char ch){ return ch==' ' || ch=='\t' || ch=='\n' || ch=='\r'; });
                if(text.size()>=2 && text[0]=='&' && (text[1]=='H' || text[1]=='h' || text[1]=='O' || text[1]=='o')) {
                    const unsigned base=(text[1]=='H' || text[1]=='h') ? 16 : 8;
                    std::uint32_t bits=0; std::size_t end=2;
                    for(;end<text.size();++end) {
                        const auto ch=static_cast<unsigned char>(text[end]);
                        const unsigned digit=ch>='0' && ch<='9' ? ch-'0' : ch>='a' && ch<='f' ? ch-'a'+10 : ch>='A' && ch<='F' ? ch-'A'+10 : 16;
                        if(digit>=base) break;
                        if(bits>(0xffffffffU-digit)/base) throw error("Val radix overflow",6);
                        bits=bits*base+digit;
                    }
                    if(end==2) return VbaValue::numeric(0);
                    if(end<text.size() && std::string_view("!@#$^").find(text[end])!=std::string_view::npos)
                        throw error("Val radix suffix conversion is not implemented",3251);
                    const bool longValue=end<text.size() && text[end]=='&';
                    const bool integerValue=end<text.size() && text[end]=='%';
                    if(integerValue && bits>0xffffU) throw error("Val Integer overflow",6);
                    if(!longValue && bits<=0xffffU) return VbaValue::integer(static_cast<std::int16_t>(bits>=0x8000U ? static_cast<std::int64_t>(bits)-0x10000 : static_cast<std::int64_t>(bits)));
                    return VbaValue::longInteger(static_cast<std::int32_t>(bits>=0x80000000U ? static_cast<std::int64_t>(bits)-0x100000000LL : bits));
                }
                std::size_t end=0;
                if(end<text.size() && (text[end]=='+' || text[end]=='-')) ++end;
                bool digits=false;
                while(end<text.size() && text[end]>='0' && text[end]<='9') { digits=true; ++end; }
                if(end<text.size() && text[end]=='.') {
                    ++end; while(end<text.size() && text[end]>='0' && text[end]<='9') { digits=true; ++end; }
                }
                if(!digits) return VbaValue::numeric(0);
                if(end<text.size() && (text[end]=='e' || text[end]=='E' || text[end]=='d' || text[end]=='D')) {
                    auto exponent=end+1;
                    if(exponent<text.size() && (text[exponent]=='+' || text[exponent]=='-')) ++exponent;
                    const auto start=exponent;
                    while(exponent<text.size() && text[exponent]>='0' && text[exponent]<='9') ++exponent;
                    if(exponent>start) { text[end]='e'; end=exponent; }
                }
                const auto start=!text.empty() && text[0]=='+' ? 1U : 0U;
                double number=0;
                const auto parsed=std::from_chars(text.data()+start,text.data()+end,number);
                if(parsed.ec==std::errc::result_out_of_range || !std::isfinite(number)) throw error("Val overflow",6);
                if(parsed.ec!=std::errc{}) throw error("Invalid Val number",13);
                const auto result=VbaValue::numeric(number);
                if(end<text.size()) {
                    const auto suffix=text[end];
                    if(suffix=='%' || suffix=='&') {
                        if(std::trunc(number)!=number) throw error("Val integral suffix mismatch",13);
                        return coerceVbaModuleValue(result,suffix=='%' ? "integer" : "long");
                    }
                    if(suffix=='!') return coerceVbaModuleValue(result,"single");
                    if(suffix=='@') return convertCurrency(result);
                    if(suffix=='$') throw error("Val String suffix mismatch",13);
                    if(suffix=='^') throw error("64-bit Val suffix is not implemented",3251);
                }
                return result;
            }
            if(word=="lenb") {
                if(!host_) return {};
                if(args.size()!=1 || omitted[0]) throw error("LenB requires one argument",450);
                const auto scalar=scalarObjectValue(args[0]);
                if(scalar.kind==VbaValue::Kind::Null) return {};
                if(scalar.kind==VbaValue::Kind::Empty) return VbaValue::longInteger(0);
                if(scalar.kind!=VbaValue::Kind::Text) throw error("LenB non-string storage layout is not implemented",3251);
                const auto units=vbaStringCodeUnits(scalar.text);
                if(units>static_cast<std::size_t>(std::numeric_limits<std::int32_t>::max()/2)) throw error("LenB result overflow",6);
                return VbaValue::longInteger(static_cast<std::int32_t>(units*2));
            }
            if ((word == "isnull" || word == "isempty" || word == "len" || word == "cstr"
                || word == "isnumeric") && one) {
                if (!host_) return {};
                const auto& value = args[0];
                if (word == "isnull") return VbaValue::logical(value.kind == VbaValue::Kind::Null);
                if (word == "isempty") return VbaValue::logical(value.kind == VbaValue::Kind::Empty);
                if (word == "len") {
                    const auto scalar=scalarObjectValue(value);
                    if(scalar.kind==VbaValue::Kind::Date) return VbaValue::longInteger(8);
                    return scalar.kind==VbaValue::Kind::Null ? VbaValue{} : stringLongResult(vbaStringCodeUnits(scalar.asText()));
                }
                if (word == "cstr") {
                    if (value.kind==VbaValue::Kind::Error) return VbaValue::string("Error "+std::to_string(value.errorNumber));
                    if (value.kind == VbaValue::Kind::Null) throw error("VBA: Ungültige Verwendung von Null", 94);
                    return VbaValue::string(value.asText());
                }
                if (word == "isnumeric") {
                    try { const auto scalar=scalarObjectValue(value); if(scalar.kind==VbaValue::Kind::Date) return VbaValue::logical(false);if(scalar.kind==VbaValue::Kind::Text) return VbaValue::logical(host_->isNumericText(scalar.text)); (void)scalar.asNumber(); return VbaValue::logical(true); }
                    catch (...) { return VbaValue::logical(false); }
                }
            }
            if (word == "nz" && (args.size() == 1 || args.size() == 2)) {
                if (!host_) return {};
                if(omitted[0]) throw error("Nz requires a value",450);
                return args[0].kind == VbaValue::Kind::Null
                    ? (args.size() == 2 && !omitted[1] ? args[1] : accessExpression_ && !jetEval_ ? VbaValue::string("") : VbaValue::empty()) : args[0];
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
            if(qualifiedIntrinsic) throw error("Invalid qualified VBA intrinsic arguments",450);
            if (resolver_) return resolver_(tok.value,args,VbaArgumentReferences{std::move(references),std::move(arrayReferences)});
            throw error("Nicht unterstützte VBA-Funktion/Parameter: " + tok.value);
        }
        if (!host_) return {}; // Validate identifiers syntactically, resolve on execution.
        if(qualifiedIntrinsic) throw error("Qualified VBA member requires supported arguments",3251);
        const auto variable = locals_.find(word);
        if (variable==locals_.end() && accessExpression_)
            return host_->expressionValue(tok.value,objectContext || nextSymbol(".") || nextSymbol("!"));
        if (variable==locals_.end()) {
            if (auto value=host_->resolveIdentifier(tok.value,objectContext || nextSymbol(".") || nextSymbol("!"))) return *value;
            throw error("VBA-Variable nicht deklariert: " + tok.value);
        }
        return variable->second;
    }
    std::string_view source_;
    std::vector<Token> tokens_;
    std::size_t index_{};
    VbaHost* host_{};
    const VbaVariables& locals_;
    VbaFunctionResolver resolver_;
    bool objectContext_{};
    bool accessExpression_{};
    bool jetEval_{};
    const VbaValue* withContext_{};
    bool allowWith_{};
    VbaCompareMode compareMode_;
    int arrayBase_;
    std::uint64_t executionEpoch_;
};
} // namespace

bool VbaArgumentSignature::objectArgument(std::size_t position,std::string_view name) const {
    if(!name.empty()) {
        const auto requested=lower(std::string(name)); position=names.size();
        for(std::size_t index=0;index<names.size();++index) if(lower(names[index])==requested) { position=index; break; }
    }
    return position<objectArguments.size() && objectArguments[position];
}

VbaNumericFormatOptions parseVbaNumericFormatOptions(const std::vector<VbaValue>& arguments) {
    if(arguments.empty() || arguments.size()>5 || arguments[0].kind==VbaValue::Kind::Missing) throw VbaRuntimeError(450,"Numeric format requires an expression and at most five arguments");VbaNumericFormatOptions result;
    for(std::size_t i=1;i<arguments.size();++i) {if(arguments[i].kind==VbaValue::Kind::Missing) continue;const auto value=scalarVbaValue(arguments[i]);const auto number=coerceVbaModuleValue(value,"long").asNumber();if(i==1) {if(number==-1) continue;if(number<0 || number>28) throw VbaRuntimeError(5,"Numeric format decimal places outside0..28");result.digits=static_cast<int>(number);}else {if(number==-2) continue;if(number!=-1 && number!=0) throw VbaRuntimeError(5,"Numeric format requires a TriState flag");if(i==2) result.leading=number==-1;else if(i==3) result.parens=number==-1;else result.group=number==-1;}}
    return result;
}
VbaValue VbaHost::formatNumericValue(const VbaValue&,bool,const VbaNumericFormatOptions&) {throw VbaRuntimeError(3251,"Numeric formatting requires a locale-aware native host");}
bool VbaHost::isNumericText(std::string_view text) {try {(void)VbaValue::string(std::string(text)).asNumber();return true;}catch(...) {return false;}}
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
        const auto epoch=vbaExecutionEpoch();
        const auto result=executeVbaProcedureWithArguments(*module,name,*this,arguments);
        if(vbaExecutionStopRequested() || epoch!=vbaExecutionEpoch()) throw VbaExecutionStopped{};
        if (result.status!=VbaExecutionStatus::Completed || !result.typedReturnValue)
            throw VbaRuntimeError(result.errorNumber ? result.errorNumber : 2425,"Eval function execution failed");
        return *result.typedReturnValue;
    };
    const VbaVariables variables;
    auto value=scalarVbaValue(ExpressionParser(source,this,variables,functions,false,true,nullptr,false,VbaCompareMode::Database,0,true).parse());
    if (value.kind==VbaValue::Kind::Error) throw VbaRuntimeError(value.errorNumber ? value.errorNumber : 13,"Eval expression failed");
    if (value.kind==VbaValue::Kind::Missing) throw VbaRuntimeError(13,"Eval cannot return a missing argument");
    if (isVbaArray(value)) throw VbaRuntimeError(13,"Eval cannot return a byte array");
    return value;
}
std::optional<VbaValue> VbaHost::resolveIdentifier(std::string_view,bool) { return std::nullopt; }
bool VbaHost::assignIdentifier(std::string_view,const VbaValue&) { return false; }
VbaValue VbaHost::bracketValue(std::string_view,bool) { throw VbaRuntimeError(438,"Access bracket evaluation context unavailable"); }
VbaValue VbaHost::expressionValue(std::string_view,bool) { throw VbaRuntimeError(438,"Access expression context unavailable"); }
VbaValue VbaHost::domainAggregate(std::string_view,const std::vector<VbaValue>&) { throw VbaRuntimeError(438,"Access domain context unavailable"); }
VbaValue VbaHost::expressionAggregate(std::string_view,std::string_view,bool) { throw VbaRuntimeError(438,"Access aggregate context unavailable"); }
VbaValue VbaHost::expressionAggregateExpression(std::string_view,std::string_view) { throw VbaRuntimeError(3251,"Compound aggregate context unavailable"); }
VbaValue VbaHost::controlValue(std::string_view name) { return fieldValue(name); }
void VbaHost::setControlValue(std::string_view name,const VbaValue& value) { setFieldValue(name,value); }
VbaValue VbaHost::fieldValue(std::string_view name) {
    const auto value = field(name);
    return value ? VbaValue::string(*value) : VbaValue{};
}
void VbaHost::setFieldValue(std::string_view name,const VbaValue& value) { setField(name,value.asField()); }
VbaValue VbaValue::missing() { VbaValue v; v.kind=Kind::Missing; return v; }
VbaValue VbaValue::empty() { VbaValue v; v.kind = Kind::Empty; return v; }
VbaValue VbaValue::errorValue(int number) { VbaValue v; v.kind=Kind::Error; v.errorNumber=number; return v; }
VbaValue VbaValue::reference(std::shared_ptr<VbaObject> value) { VbaValue v; v.kind = Kind::Object; v.object = std::move(value); return v; }
VbaValue VbaValue::dateSerial(double value) {
    if(!std::isfinite(value) || value<=-657435.0 || value>=2958466.0) throw error("Date serial out of range",6);
    VbaValue result; result.kind=Kind::Date; result.number=value; return result;
}
VbaValue VbaValue::numeric(double value) { VbaValue v; v.kind = Kind::Number; v.number = value; return v; }
VbaValue VbaValue::byte(unsigned char value) { auto v=numeric(value); v.numericType=NumericType::Byte; return v; }
VbaValue VbaValue::single(float value) { auto v=numeric(value); v.numericType=NumericType::Single; return v; }
VbaValue VbaValue::integer(std::int16_t value) { auto v=numeric(value); v.numericType=NumericType::Integer; return v; }
VbaValue VbaValue::longInteger(std::int32_t value) { auto v=numeric(value); v.numericType=NumericType::Long; return v; }
VbaValue VbaValue::string(std::string value) { VbaValue v; v.kind = Kind::Text; v.text = std::move(value); return v; }
VbaValue VbaValue::byteArray(std::vector<unsigned char> value) { VbaValue v; v.kind=Kind::ByteArray; v.bytes=std::move(value); v.object=std::make_shared<ByteArrayAllocation>(); return v; }
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
    if (kind==Kind::ByteArray || kind==Kind::Array) throw error("Byte array cannot be implicitly converted to text",13);
    if (kind==Kind::Error || kind==Kind::Missing) throw error("Error Variant cannot be implicitly converted to text",13);
    if (kind==Kind::Object) return scalarObjectValue(*this).asText();
    if (kind == Kind::Text || kind==Kind::Decimal) return text;
    if (kind==Kind::Date) return vbaDateText(number);
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
    if (kind==Kind::ByteArray || kind==Kind::Array) throw error("Byte array cannot be implicitly converted to number",13);
    if (kind==Kind::Error || kind==Kind::Missing) throw error("Error Variant cannot be implicitly converted to number",13);
    if (kind==Kind::Object) return scalarObjectValue(*this).asNumber();
    if (kind == Kind::Empty) return 0;
    if (kind == Kind::Number || kind==Kind::Date) return number;
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
                               const VbaVariables& vars, const VbaFunctionResolver& resolver,const VbaValue* withContext,VbaCompareMode compareMode,int arrayBase) {
    return ExpressionParser(text, &host, vars, resolver,false,false,withContext,false,compareMode,arrayBase).parse();
}
VbaEvaluatedArgument evaluateVbaArgumentExpression(std::string_view expression,VbaHost& host,
    const VbaVariables& variables,const VbaFunctionResolver& resolver,const VbaValue* withContext,VbaCompareMode compareMode,int arrayBase,bool objectContext) {
    return ExpressionParser(expression,&host,variables,resolver,objectContext,false,withContext,false,compareMode,arrayBase).parseArgument();
}
VbaValue evaluateVbaObjectExpression(std::string_view text,VbaHost& host,
                                    const VbaVariables& vars,const VbaFunctionResolver& resolver,const VbaValue* withContext,VbaCompareMode compareMode,int arrayBase) {
    return ExpressionParser(text,&host,vars,resolver,true,false,withContext,false,compareMode,arrayBase).parse();
}
VbaValue evaluateAccessExpression(std::string_view text,VbaHost& host) {
    const VbaVariables variables;
    return ExpressionParser(text,&host,variables,{},false,true).parse();
}
}
