#include "vba_round.hpp"
#include <cmath>
#include <limits>

namespace linaccess {
VbaValue vbaRound(const std::vector<VbaValue>& arguments) {
    if(arguments.empty() || arguments.size()>2) throw VbaRuntimeError(450,"Round requires one or two arguments");
    const auto value=scalarVbaValue(arguments[0]);
    const auto precision=arguments.size()==2 ? scalarVbaValue(arguments[1]) : VbaValue::integer(0);
    const auto places=static_cast<int>(coerceVbaModuleValue(precision,"long").number);
    if(places<0) throw VbaRuntimeError(5,"Round precision must be nonnegative");
    if(places>15) throw VbaRuntimeError(3251,"Round precision above 15 is not validated");
    if(value.kind==VbaValue::Kind::Null) return {};
    if(value.kind==VbaValue::Kind::Currency) {
        if(places>=4) return value;
        std::int64_t quantum=1;
        for(int i=places;i<4;++i) quantum*=10;
        auto quotient=value.currencyUnits/quantum; const auto remainder=value.currencyUnits%quantum;
        const auto magnitude=remainder<0 ? -remainder : remainder;
        if(magnitude>quantum/2 || (magnitude==quantum/2 && quotient%2!=0)) quotient+=value.currencyUnits<0 ? -1 : 1;
        if(quotient>std::numeric_limits<std::int64_t>::max()/quantum || quotient<std::numeric_limits<std::int64_t>::min()/quantum)
            throw VbaRuntimeError(6,"Rounded Currency overflow");
        return VbaValue::currency(quotient*quantum);
    }
    if(value.kind==VbaValue::Kind::Empty || value.kind==VbaValue::Kind::Boolean)
        return VbaValue::integer(value.kind==VbaValue::Kind::Boolean && value.boolean ? -1 : 0);
    if(value.kind==VbaValue::Kind::Number) {
        if(value.numericType==VbaValue::NumericType::Byte) return coerceVbaModuleValue(value,"byte");
        if(value.numericType==VbaValue::NumericType::Integer) return coerceVbaModuleValue(value,"integer");
        if(value.numericType==VbaValue::NumericType::Long) return coerceVbaModuleValue(value,"long");
        if(value.numericType==VbaValue::NumericType::Single) throw VbaRuntimeError(3251,"Single Round precision semantics are not validated");
    }
    const auto number=coerceVbaModuleValue(value,"double").number;
    static constexpr double factors[]{1,10,100,1000,10000,100000,1000000,10000000,100000000,1000000000,10000000000,100000000000,1000000000000,10000000000000,100000000000000,1000000000000000};
    const auto scale=factors[places],scaled=number*scale;
    if(!std::isfinite(scaled)) throw VbaRuntimeError(3251,"Round scaling overflow is not validated");
    const auto lower=std::floor(scaled),fraction=scaled-lower;
    const auto rounded=fraction<0.5 ? lower : fraction>0.5 ? lower+1 : std::fmod(std::abs(lower),2.0)==0 ? lower : lower+1;
    const auto result=rounded/scale;
    if(!std::isfinite(result)) throw VbaRuntimeError(6,"Round overflow");
    return VbaValue::numeric(result==0 ? 0 : result);
}
}
