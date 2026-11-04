#pragma once
#include "vba_expression.hpp"
#include <cmath>
#include <optional>
#include <limits>
#include <string>

namespace linaccess {
// Scalar report aggregate state only; never retains a host, object or cursor.
class NativeReportAggregate {
public:
    NativeReportAggregate(std::string_view function,bool allowTextSum=false)
        : function_(function),allowTextSum_(allowTextSum) {
        if(function_!="sum" && function_!="count" && function_!="avg" && function_!="min" && function_!="max" && function_!="first" && function_!="last" && function_!="var" && function_!="stdev")
            throw VbaRuntimeError(3251,"Report aggregate unsupported");
    }
    void add(const VbaValue& value) {
        if(function_=="first" || function_=="last") {
            switch(value.kind) {
                case VbaValue::Kind::Null: case VbaValue::Kind::Empty: case VbaValue::Kind::Number:
                case VbaValue::Kind::Text: case VbaValue::Kind::Boolean: case VbaValue::Kind::Currency:
                case VbaValue::Kind::Date: break;
                default: throw VbaRuntimeError(3251,"Report selector value type unsupported");
            }
            if((value.kind==VbaValue::Kind::Number || value.kind==VbaValue::Kind::Date) && !std::isfinite(value.number))
                throw VbaRuntimeError(6,"Aggregate overflow");
            ++count_;
            if(!extremum_ || function_=="last") extremum_=value;
            return;
        }
        if(value.kind==VbaValue::Kind::Null) return;
        if(function_=="count") { ++count_; return; }
        if(function_=="min" || function_=="max") {
            if(value.kind!=VbaValue::Kind::Number && value.kind!=VbaValue::Kind::Date && value.kind!=VbaValue::Kind::Currency)
                throw VbaRuntimeError(3251,"Report extremum value type unsupported");
            if(extremum_ && value.kind!=extremum_->kind)
                throw VbaRuntimeError(3251,"Mixed report extremum types unsupported");
            if(value.kind!=VbaValue::Kind::Currency && !std::isfinite(value.number))
                throw VbaRuntimeError(6,"Aggregate overflow");
            const int order=!extremum_ ? 0 : value.kind==VbaValue::Kind::Currency
                ? (value.currencyUnits>extremum_->currencyUnits)-(value.currencyUnits<extremum_->currencyUnits)
                : (value.number>extremum_->number)-(value.number<extremum_->number);
            if(!extremum_ || (function_=="min" ? order<0 : order>0)) extremum_=value;
            ++count_;
            return;
        }
        if(function_=="sum" && value.kind==VbaValue::Kind::Currency) {
            if(count_ && !currencyTotal_) throw VbaRuntimeError(3251,"Mixed Currency sum types unsupported");
            const auto previous=currencyTotal_.value_or(0),units=value.currencyUnits;
            if((units>0 && previous>std::numeric_limits<std::int64_t>::max()-units)
                || (units<0 && previous<std::numeric_limits<std::int64_t>::min()-units))
                throw VbaRuntimeError(6,"Currency sum overflow");
            currencyTotal_=previous+units;
            ++count_;
            return;
        }
        if(currencyTotal_) throw VbaRuntimeError(3251,"Mixed Currency sum types unsupported");
        if(function_!="sum" && value.kind!=VbaValue::Kind::Number)
            throw VbaRuntimeError(3251,"Report aggregate value type unsupported");
        if(function_=="sum" && !allowTextSum_ && value.kind==VbaValue::Kind::Text)
            throw VbaRuntimeError(13,"Sum requires a numeric expression");
        const double number=value.asNumber();
        if(!std::isfinite(number)) throw VbaRuntimeError(6,"Aggregate overflow");
        ++count_;
        if(function_=="sum") {
            total_+=number;
            if(!std::isfinite(total_)) throw VbaRuntimeError(6,"Aggregate overflow");
        } else {
            const long double delta=static_cast<long double>(number)-mean_;
            mean_+=delta/count_;
            if(!std::isfinite(mean_)) throw VbaRuntimeError(6,"Aggregate overflow");
            if(function_=="var" || function_=="stdev") {
                secondMoment_+=delta*(static_cast<long double>(number)-mean_);
                if(!std::isfinite(secondMoment_)) throw VbaRuntimeError(6,"Aggregate overflow");
            }
        }
    }
    VbaValue result() const {
        if(function_=="count") return VbaValue::numeric(count_);
        if(!count_) return {};
        if(function_=="var" || function_=="stdev") {
            if(count_<2) return {};
            const long double variance=secondMoment_/(count_-1);
            if(variance<0) throw VbaRuntimeError(6,"Variance computation overflow");
            const double value=static_cast<double>(function_=="stdev" ? std::sqrt(variance) : variance);
            if(!std::isfinite(value)) throw VbaRuntimeError(6,"Aggregate overflow");
            return VbaValue::numeric(value);
        }
        if(extremum_) return *extremum_;
        if(currencyTotal_) return VbaValue::currency(*currencyTotal_);
        const double value=function_=="avg" ? static_cast<double>(mean_) : total_;
        if(!std::isfinite(value)) throw VbaRuntimeError(6,"Aggregate overflow");
        return VbaValue::numeric(value);
    }
private:
    std::string function_;
    bool allowTextSum_{};
    std::size_t count_{};
    double total_{};
    long double mean_{},secondMoment_{};
    std::optional<VbaValue> extremum_;
    std::optional<std::int64_t> currencyTotal_;
};
}
