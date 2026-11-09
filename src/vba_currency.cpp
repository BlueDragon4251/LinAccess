#include "vba_currency.hpp"
#include <cmath>
#include <limits>
namespace linaccess {
namespace {
VbaValue checked(__int128 units) {
    if (units<std::numeric_limits<std::int64_t>::min() || units>std::numeric_limits<std::int64_t>::max())
        throw VbaRuntimeError(6,"Currency overflow");
    return VbaValue::currency(static_cast<std::int64_t>(units));
}
VbaValue decimalCurrency(std::string_view input) {
    const auto first=input.find_first_not_of(" \t\r\n");
    if (first==std::string_view::npos) throw VbaRuntimeError(13,"Currency type mismatch");
    input=input.substr(first,input.find_last_not_of(" \t\r\n")-first+1);
    bool negative=false;
    if (input.front()=='+' || input.front()=='-') { negative=input.front()=='-'; input.remove_prefix(1); }
    std::string digits; bool point=false; int fraction=0;
    std::size_t pos=0;
    for (;pos<input.size();++pos) {
        const char ch=input[pos];
        if (ch=='.' && !point) { point=true; continue; }
        if (ch<'0' || ch>'9') break;
        digits+=ch; if (point) ++fraction;
    }
    if (digits.empty()) throw VbaRuntimeError(13,"Currency type mismatch");
    int exponent=0;
    if (pos<input.size() && (input[pos]=='e' || input[pos]=='E')) {
        ++pos; bool minus=false;
        if (pos<input.size() && (input[pos]=='+' || input[pos]=='-')) { minus=input[pos]=='-'; ++pos; }
        const auto begin=pos;
        for (;pos<input.size() && input[pos]>='0' && input[pos]<='9';++pos) {
            if (exponent>10000) throw VbaRuntimeError(6,"Currency exponent overflow");
            exponent=exponent*10+(input[pos]-'0');
        }
        if (begin==pos) throw VbaRuntimeError(13,"Currency type mismatch");
        if (minus) exponent=-exponent;
    }
    if (pos!=input.size()) throw VbaRuntimeError(13,"Currency type mismatch");
    // Work in decimal digits so the extreme four-place values never pass
    // through a binary Double. Ties round to an even scaled integer.
    const auto nonzero=digits.find_first_not_of('0');
    if (nonzero==std::string::npos) return VbaValue::currency(0);
    digits.erase(0,nonzero);
    const long long kept=static_cast<long long>(digits.size())+exponent-fraction+4;
    if (kept>19) throw VbaRuntimeError(6,"Currency overflow");
    __int128 units=0;
    for (long long i=0;i<kept;++i) units=units*10+(i<static_cast<long long>(digits.size()) ? digits[i]-'0' : 0);
    if (kept>=0 && kept<static_cast<long long>(digits.size())) {
        const auto next=digits[kept];
        const bool tail=digits.find_first_not_of('0',kept+1)!=std::string::npos;
        if (next>'5' || (next=='5' && (tail || units%2!=0))) ++units;
    }
    return checked(negative ? -units : units);
}
}
VbaValue convertCurrency(const VbaValue& value) {
    if (value.kind==VbaValue::Kind::Currency) return value;
    if (value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null Currency conversion");
    if (value.kind==VbaValue::Kind::Text || value.kind==VbaValue::Kind::Decimal) return decimalCurrency(value.text);
    if (value.kind==VbaValue::Kind::Empty) return VbaValue::currency(0);
    if (value.kind==VbaValue::Kind::Boolean) return VbaValue::currency(value.boolean ? -10000 : 0);
    if (value.kind!=VbaValue::Kind::Number) throw VbaRuntimeError(13,"Currency type mismatch");
    const long double scaled=static_cast<long double>(value.number)*10000;
    if (!std::isfinite(scaled)) throw VbaRuntimeError(6,"Currency overflow");
    const auto floor=std::floor(scaled), fraction=scaled-floor;
    const auto rounded=fraction<0.5L ? floor : fraction>0.5L ? floor+1 : std::fmod(std::abs(floor),2.L)==0 ? floor : floor+1;
    if (rounded<std::numeric_limits<std::int64_t>::min() || rounded>std::numeric_limits<std::int64_t>::max())
        throw VbaRuntimeError(6,"Currency overflow");
    return VbaValue::currency(static_cast<std::int64_t>(rounded));
}
VbaValue currencyFromRealStorage(double value) {
    static_assert(std::numeric_limits<double>::is_iec559 && std::numeric_limits<double>::digits==53);
    if(!std::isfinite(value)) throw VbaRuntimeError(13,"Currency REAL storage is not finite");
    if(std::abs(value)>=549755813888.0) throw VbaRuntimeError(3251,"Currency REAL storage outside the lossless supported range");
    if(value==0) return VbaValue::currency(0);
    int exponent=0;
    const auto fraction=std::frexp(std::abs(value),&exponent);
    // value*10000 = (53-bit mantissa * 625) / 2^(49-exponent).
    // The product fits 63 bits. Below 2^39 every four-decimal tick has
    // a distinct binary64 encoding; use integer rounding, not double scaling.
    const auto product=static_cast<std::uint64_t>(std::ldexp(fraction,53))*625ULL;
    const int shift=49-exponent; // At least 10 after the range check.
    std::uint64_t units=0;
    if(shift<64) {
        units=product>>shift;
        const auto remainder=product&((1ULL<<shift)-1),half=1ULL<<(shift-1);
        if(remainder>half || (remainder==half && (units&1))) ++units;
    }
    const auto signedUnits=std::signbit(value) ? -static_cast<std::int64_t>(units) : static_cast<std::int64_t>(units);
    if(static_cast<double>(signedUnits)/10000.0!=value)
        throw VbaRuntimeError(13,"Currency REAL storage is not a canonical four-decimal value");
    return VbaValue::currency(signedUnits);
}
double currencyToRealStorage(const VbaValue& value) {
    const auto exact=convertCurrency(scalarVbaValue(value));
    constexpr std::int64_t limit=5497558138880000LL;
    if(exact.currencyUnits<=-limit || exact.currencyUnits>=limit)
        throw VbaRuntimeError(3251,"Currency REAL storage outside the lossless supported range");
    const auto encoded=static_cast<double>(exact.currencyUnits)/10000.0;
    if(currencyFromRealStorage(encoded).currencyUnits!=exact.currencyUnits)
        throw VbaRuntimeError(3251,"Currency cannot round trip through REAL storage");
    return encoded;
}
VbaValue currencyBinary(std::string_view op,const VbaValue& left,const VbaValue& right) {
    const __int128 a=left.currencyUnits,b=right.currencyUnits;
    if (op=="+") return checked(a+b);
    if (op=="-") return checked(a-b);
    if (op=="*") {
        const auto product=a*b;
        auto units=product/10000;
        const auto remainder=product%10000;
        const auto absolute=remainder<0 ? -remainder : remainder;
        if (absolute>5000 || (absolute==5000 && units%2!=0)) units+=product<0 ? -1 : 1;
        return checked(units);
    }
    throw VbaRuntimeError(3251,"Currency operation unsupported");
}
VbaValue negateCurrency(const VbaValue& value) { return checked(-static_cast<__int128>(value.currencyUnits)); }
VbaValue floorCurrency(const VbaValue& value) {
    const __int128 units=value.currencyUnits;
    return checked((units/10000-(units%10000<0 ? 1 : 0))*10000);
}
}
