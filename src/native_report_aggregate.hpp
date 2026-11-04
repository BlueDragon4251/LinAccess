#pragma once
#include "vba_expression.hpp"
#include <cmath>
#include <optional>
#include <string>

namespace linaccess {
// Numeric report aggregate state only; never retains a host, object or cursor.
class NativeReportAggregate {
public:
    NativeReportAggregate(std::string_view function,bool allowTextSum=false)
        : function_(function),allowTextSum_(allowTextSum) {
        if(function_!="sum" && function_!="count" && function_!="avg" && function_!="min" && function_!="max")
            throw VbaRuntimeError(3251,"Report aggregate unsupported");
    }
    void add(const VbaValue& value) {
        if(value.kind==VbaValue::Kind::Null) return;
        if(function_=="count") { ++count_; return; }
        if(function_!="sum" && value.kind!=VbaValue::Kind::Number)
            throw VbaRuntimeError(3251,"Report aggregate value type unsupported");
        if(function_=="sum" && !allowTextSum_ && value.kind==VbaValue::Kind::Text)
            throw VbaRuntimeError(13,"Sum requires a numeric expression");
        const double number=value.asNumber();
        if(!std::isfinite(number)) throw VbaRuntimeError(6,"Aggregate overflow");
        ++count_;
        if(function_=="min" || function_=="max") {
            if(!extremum_ || (function_=="min" ? number<extremum_->number : number>extremum_->number)) extremum_=value;
        } else if(function_=="sum") {
            total_+=number;
            if(!std::isfinite(total_)) throw VbaRuntimeError(6,"Aggregate overflow");
        } else {
            mean_+=(static_cast<long double>(number)-mean_)/count_;
            if(!std::isfinite(mean_)) throw VbaRuntimeError(6,"Aggregate overflow");
        }
    }
    VbaValue result() const {
        if(function_=="count") return VbaValue::numeric(count_);
        if(!count_) return {};
        if(extremum_) return *extremum_;
        const double value=function_=="avg" ? static_cast<double>(mean_) : total_;
        if(!std::isfinite(value)) throw VbaRuntimeError(6,"Aggregate overflow");
        return VbaValue::numeric(value);
    }
private:
    std::string function_;
    bool allowTextSum_{};
    std::size_t count_{};
    double total_{};
    long double mean_{};
    std::optional<VbaValue> extremum_;
};
}
