#include "vba_decimal.hpp"
#include <boost/multiprecision/cpp_int.hpp>
#include <cmath>
#include <iomanip>
#include <locale>
#include <sstream>
#include <algorithm>

namespace linaccess {
namespace {
using boost::multiprecision::cpp_int;
struct Parts {cpp_int magnitude;int scale{};bool negative{};};
const cpp_int& maximum() {static const cpp_int result=(cpp_int(1)<<96)-1;return result;}
cpp_int power(int exponent) {if(exponent<0 || exponent>10000) throw VbaRuntimeError(6,"Decimal exponent overflow");cpp_int result=1;for(int i=0;i<exponent;++i) result*=10;return result;}
cpp_int rounded(const cpp_int& numerator,const cpp_int& denominator) {
    cpp_int quotient=numerator/denominator;const cpp_int remainder=numerator%denominator,twice=remainder*2;
    if(twice>denominator || (twice==denominator && (quotient&1)!=0)) ++quotient;return quotient;
}
VbaValue store(Parts parts,bool trim=false) {
    if(parts.scale<0) {parts.magnitude*=power(-parts.scale);parts.scale=0;}
    const auto original=parts.magnitude;const auto scale=parts.scale;
    for(int removed=std::max(0,scale-28);; ++removed) {
        if(removed>scale) throw VbaRuntimeError(6,"Decimal overflow");
        const cpp_int magnitude=removed ? rounded(original,power(removed)) : original;
        if(magnitude>maximum()) continue;parts.magnitude=magnitude;parts.scale=scale-removed;break;
    }
    if(trim) while(parts.scale && parts.magnitude%10==0) {parts.magnitude/=10;--parts.scale;}
    auto digits=parts.magnitude.convert_to<std::string>();
    if(parts.scale) {if(digits.size()<=static_cast<std::size_t>(parts.scale)) digits.insert(0,parts.scale+1-digits.size(),'0');digits.insert(digits.size()-parts.scale,1,'.');}
    if(parts.negative && parts.magnitude!=0) digits.insert(0,1,'-');VbaValue value;value.kind=VbaValue::Kind::Decimal;value.text=std::move(digits);return value;
}
Parts parse(std::string_view text) {
    if(text.size()>512) throw VbaRuntimeError(6,"Decimal literal exceeds its supported bound");
    const auto begin=text.find_first_not_of(" \t\r\n");if(begin==std::string_view::npos) throw VbaRuntimeError(13,"Decimal type mismatch");text=text.substr(begin,text.find_last_not_of(" \t\r\n")-begin+1);Parts result;
    if(text.front()=='-' || text.front()=='+') {result.negative=text.front()=='-';text.remove_prefix(1);}
    bool point=false,any=false;std::size_t position=0;
    for(;position<text.size();++position) {const auto ch=text[position];if(ch=='.' && !point) {point=true;continue;}if(ch<'0' || ch>'9') break;any=true;result.magnitude*=10;result.magnitude+=ch-'0';if(point) ++result.scale;}
    if(!any) throw VbaRuntimeError(13,"Decimal type mismatch");
    int exponent=0;if(position<text.size() && (text[position]=='e' || text[position]=='E')) {++position;bool negative=false;if(position<text.size() && (text[position]=='+' || text[position]=='-')) {negative=text[position]=='-';++position;}const auto first=position;for(;position<text.size() && text[position]>='0' && text[position]<='9';++position) {if(exponent>1000) throw VbaRuntimeError(6,"Decimal exponent overflow");exponent=exponent*10+text[position]-'0';}if(first==position) throw VbaRuntimeError(13,"Decimal type mismatch");if(negative) exponent=-exponent;}
    if(position!=text.size()) throw VbaRuntimeError(13,"Decimal type mismatch");result.scale-=exponent;return result;
}
Parts parts(const VbaValue& value) {return parse(convertDecimal(value).text);}
cpp_int signedMagnitude(const Parts& value,int scale) {cpp_int result=value.magnitude*power(scale-value.scale);return value.negative ? -result : result;}
}
VbaValue decimalLiteral(std::string_view text) {return store(parse(text));}
VbaValue convertDecimal(const VbaValue& input) {
    const auto value=scalarVbaValue(input);if(value.kind==VbaValue::Kind::Decimal) return value;
    if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null Decimal conversion");
    if(value.kind==VbaValue::Kind::Empty) return decimalLiteral("0");
    if(value.kind==VbaValue::Kind::Boolean) return decimalLiteral(value.boolean ? "-1" : "0");
    if(value.kind==VbaValue::Kind::Text || value.kind==VbaValue::Kind::Currency) return decimalLiteral(value.asText());
    if(value.kind!=VbaValue::Kind::Number && value.kind!=VbaValue::Kind::Date) throw VbaRuntimeError(13,"Decimal type mismatch");
    if(!std::isfinite(value.number)) throw VbaRuntimeError(6,"Decimal conversion overflow");
    std::ostringstream stream;stream.imbue(std::locale::classic());stream<<std::setprecision(value.kind==VbaValue::Kind::Number && value.numericType==VbaValue::NumericType::Single ? 7 : 15)<<value.number;return decimalLiteral(stream.str());
}
VbaValue decimalBinary(std::string_view op,const VbaValue& left,const VbaValue& right) {
    if(left.kind==VbaValue::Kind::Null || right.kind==VbaValue::Kind::Null) return {};const auto a=parts(left),b=parts(right);
    if(op=="+" || op=="-") {const int scale=std::max(a.scale,b.scale);cpp_int result=signedMagnitude(a,scale);const auto other=signedMagnitude(b,scale);if(op=="+") result+=other;else result-=other;return store({result<0 ? -result : result,scale,result<0});}
    if(op=="*") return store({a.magnitude*b.magnitude,a.scale+b.scale,a.negative!=b.negative});
    if(op=="/") {
        if(b.magnitude==0) throw VbaRuntimeError(11,"Decimal division by zero");
        for(int scale=28;scale>=0;--scale) {const int exponent=scale+b.scale-a.scale;cpp_int numerator=a.magnitude,denominator=b.magnitude;if(exponent>=0) numerator*=power(exponent);else denominator*=power(-exponent);const auto result=rounded(numerator,denominator);if(result<=maximum()) return store({result,scale,a.negative!=b.negative},true);}
        throw VbaRuntimeError(6,"Decimal division overflow");
    }
    throw VbaRuntimeError(3251,"Decimal operator is not implemented");
}
VbaValue decimalUnary(std::string_view op,const VbaValue& input) {
    auto value=parts(input);if(op=="-") value.negative=!value.negative;else if(op=="abs") value.negative=false;
    else if(op=="fix" || op=="int") {const auto divisor=power(value.scale);const bool remainder=value.magnitude%divisor!=0;value.magnitude/=divisor;if(op=="int" && value.negative && remainder) ++value.magnitude;value.scale=0;}else if(op!="+") throw VbaRuntimeError(3251,"Decimal unary operator is not implemented");return store(value);
}
VbaValue decimalRound(const VbaValue& input,int places) {if(places<0) throw VbaRuntimeError(5,"Decimal Round precision must be nonnegative");if(places>28) throw VbaRuntimeError(3251,"Decimal Round precision exceeds28");auto value=parts(input);if(places<value.scale) {value.magnitude=rounded(value.magnitude,power(value.scale-places));value.scale=places;}return store(value);}
int decimalCompare(const VbaValue& left,const VbaValue& right) {const auto a=parts(left),b=parts(right);const auto scale=std::max(a.scale,b.scale);const auto first=signedMagnitude(a,scale),second=signedMagnitude(b,scale);return (first>second)-(first<second);}
std::int64_t decimalInteger(const VbaValue& input,std::int64_t minimum,std::int64_t maximumValue) {const auto value=parts(decimalRound(input,0));cpp_int integer=value.negative ? -value.magnitude : value.magnitude;if(integer<minimum || integer>maximumValue) throw VbaRuntimeError(6,"Decimal integer conversion overflow");return integer.convert_to<std::int64_t>();}
}
