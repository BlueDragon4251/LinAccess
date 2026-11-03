#include "vba_calendar.hpp"
#include <chrono>
#include <limits>
#include <cmath>
#include <iomanip>
#include <sstream>

namespace linaccess {
namespace {
struct Parts { int year=1899,month=12,day=30,hour=0,minute=0,second=0; std::chrono::sys_days date; };
Parts dateParts(const VbaValue& value) {
    if(value.kind==VbaValue::Kind::Date) return dateParts(VbaValue::string(vbaDateText(value.number)));
    Parts result;
    if(value.kind!=VbaValue::Kind::Empty) {
        if(value.kind!=VbaValue::Kind::Text) throw VbaRuntimeError(3251,"Calendar input requires ISO text; numeric Date coercion is unsupported");
        const auto& text=value.text;
        const auto digits=[&](std::size_t begin,std::size_t count) {
            int number=0;
            for(auto i=begin;i<begin+count;++i) {
                if(text[i]<'0' || text[i]>'9') throw VbaRuntimeError(13,"Invalid ISO date digits");
                number=number*10+text[i]-'0';
            }
            return number;
        };
        const bool timeOnly=text.size()==8;
        if(!timeOnly && text.size()!=10 && text.size()!=19) throw VbaRuntimeError(13,"Unambiguous ISO date or time required");
        if(!timeOnly) {
            if(text[4]!='-' || text[7]!='-' || (text.size()==19 && text[10]!=' ' && text[10]!='T')) throw VbaRuntimeError(13,"Invalid ISO date separators");
            result.year=digits(0,4); result.month=digits(5,2); result.day=digits(8,2);
        }
        if(timeOnly || text.size()==19) {
            const auto offset=timeOnly ? 0u : 11u;
            if(text[offset+2]!=':' || text[offset+5]!=':') throw VbaRuntimeError(13,"Invalid ISO time separators");
            result.hour=digits(offset,2); result.minute=digits(offset+3,2); result.second=digits(offset+6,2);
        }
    }
    const auto date=std::chrono::year(result.year)/std::chrono::month(result.month)/std::chrono::day(result.day);
    if(result.year<100 || result.year>9999 || !date.ok() || result.hour>23 || result.minute>59 || result.second>59) throw VbaRuntimeError(13,"Gregorian date or time out of range");
    result.date=std::chrono::sys_days(date); return result;
}
std::int64_t floorWeeks(std::int64_t days) { auto weeks=days/7; if(days%7<0) --weeks; return weeks; }
}
VbaValue vbaDateValue(const VbaValue& input) {
    const auto value=scalarVbaValue(input);
    if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null Date conversion");
    if(value.kind==VbaValue::Kind::Date) return value;
    if(value.kind!=VbaValue::Kind::Text) return VbaValue::dateSerial(value.asNumber());
    const auto parts=dateParts(value);
    const auto epoch=std::chrono::sys_days(std::chrono::year(1899)/12/30);
    const auto days=(parts.date-epoch).count();
    const auto fraction=(parts.hour*3600+parts.minute*60+parts.second)/86400.0;
    return VbaValue::dateSerial(days<0 ? days-fraction : days+fraction);
}
std::string vbaDateText(double serial) {
    (void)VbaValue::dateSerial(serial); // Check before integral conversion.
    const auto days=static_cast<int>(std::trunc(serial));
    auto seconds=static_cast<int>(std::floor(std::abs(serial-days)*86400.0+0.00005));
    if(seconds==86400) seconds=86399; // Keep whole-second display within the represented day.
    const auto date=std::chrono::year_month_day(std::chrono::sys_days(std::chrono::year(1899)/12/30)+std::chrono::days(days));
    std::ostringstream text; text.imbue(std::locale::classic()); text<<std::setfill('0');
    if(days!=0) text<<std::setw(4)<<int(date.year())<<'-'<<std::setw(2)<<unsigned(date.month())<<'-'<<std::setw(2)<<unsigned(date.day());
    if(seconds || days==0) {
        if(days!=0) text<<' ';
        text<<std::setw(2)<<seconds/3600<<':'<<std::setw(2)<<(seconds/60)%60<<':'<<std::setw(2)<<seconds%60;
    }
    return text.str();
}
VbaValue vbaCalendarFunction(std::string_view name,const std::vector<VbaValue>& arguments) {
    if(name=="dateserial" || name=="timeserial") {
        if(arguments.size()!=3) throw VbaRuntimeError(450,"Date construction requires three arguments");
        int values[3];
        for(int index=0;index<3;++index) values[index]=static_cast<int>(coerceVbaModuleValue(scalarVbaValue(arguments[index]),"integer").number);
        if(name=="timeserial") {
            const auto total=static_cast<std::int64_t>(values[0])*3600+values[1]*60+values[2];
            auto days=total/86400,seconds=total%86400;
            if(seconds<0) { --days; seconds+=86400; }
            const auto fraction=seconds/86400.0;
            return VbaValue::dateSerial(days<0 ? days-fraction : days+fraction);
        }
        if(values[0]<0 || values[0]>9999) throw VbaRuntimeError(5,"DateSerial year out of range");
        if(values[0]<100) throw VbaRuntimeError(3251,"Two-digit DateSerial years require a regional year window");
        const int offset=values[1]-1;
        auto years=offset/12; if(offset%12<0) --years;
        const auto month=static_cast<unsigned>(offset-years*12+1);
        const auto date=std::chrono::sys_days(std::chrono::year(values[0]+years)/std::chrono::month(month)/1)+std::chrono::days(values[2]-1);
        const auto year=static_cast<int>(std::chrono::year_month_day(date).year());
        if(year<100 || year>9999) throw VbaRuntimeError(5,"DateSerial result out of range");
        const auto epoch=std::chrono::sys_days(std::chrono::year(1899)/12/30);
        return VbaValue::dateSerial(static_cast<double>((date-epoch).count()));
    }
    if(name=="datepart") {
        if(arguments.size()<2 || arguments.size()>4) throw VbaRuntimeError(450,"DatePart requires two to four arguments");
        std::vector<VbaValue> args; for(const auto& value:arguments) args.push_back(scalarVbaValue(value));
        if(args[0].kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null DatePart interval");
        const auto option=[&](std::size_t index,int maximum) {
            if(index>=args.size()) return 1;
            const auto value=static_cast<int>(coerceVbaModuleValue(args[index],"long").number);
            if(value==0) throw VbaRuntimeError(3251,"DatePart system calendar option requires a regional adapter");
            if(value<1 || value>maximum) throw VbaRuntimeError(5,"Invalid DatePart calendar option");
            return value;
        };
        const auto firstDay=option(2,7),firstWeek=option(3,3);
        if(args[1].kind==VbaValue::Kind::Null) return {};
        auto interval=args[0].asText(); for(auto& ch:interval) if(ch>='A' && ch<='Z') ch+=('a'-'A');
        const auto date=dateParts(args[1]);
        const auto january=std::chrono::sys_days(std::chrono::year(date.year)/std::chrono::January/1);
        const auto dayOfYear=static_cast<int>((date.date-january).count())+1;
        int result;
        if(interval=="yyyy") result=date.year;
        else if(interval=="q") result=(date.month-1)/3+1;
        else if(interval=="m") result=date.month;
        else if(interval=="y") result=dayOfYear;
        else if(interval=="d") result=date.day;
        else if(interval=="w") result=(static_cast<int>(std::chrono::weekday(date.date).c_encoding())+7-(firstDay-1))%7+1;
        else if(interval=="ww") {
            if(firstWeek!=1) throw VbaRuntimeError(3251,"Non-default first-week DatePart is not validated");
            const auto offset=(static_cast<int>(std::chrono::weekday(january).c_encoding())+7-(firstDay-1))%7;
            result=(dayOfYear-1+offset)/7+1;
        } else if(interval=="h") result=date.hour;
        else if(interval=="n") result=date.minute;
        else if(interval=="s") result=date.second;
        else throw VbaRuntimeError(5,"Invalid DatePart interval");
        return VbaValue::integer(static_cast<std::int16_t>(result));
    }
    const bool difference=name=="datediff",weekday=name=="weekday";
    if(arguments.size()<(difference?3u:1u) || arguments.size()>(difference?5u:weekday?2u:1u)) throw VbaRuntimeError(450,"Calendar function argument count invalid");
    std::vector<VbaValue> args; for(const auto& value:arguments) args.push_back(scalarVbaValue(value));
    int firstDay=1;
    const auto option=difference?3u:1u;
    if((difference || weekday) && args.size()>option) firstDay=static_cast<int>(coerceVbaModuleValue(args[option],"long").number);
    if(firstDay==0) throw VbaRuntimeError(3251,"System first weekday requires a regional adapter");
    if(firstDay<1 || firstDay>7) throw VbaRuntimeError(5,"Invalid first weekday");
    int firstWeek=1;
    if(difference && args.size()>4) {
        firstWeek=static_cast<int>(coerceVbaModuleValue(args[4],"long").number);
        if(firstWeek==0) throw VbaRuntimeError(3251,"System first week requires a regional adapter");
        if(firstWeek<1 || firstWeek>3) throw VbaRuntimeError(5,"Invalid first week");
    }
    if(!difference) {
        if(args[0].kind==VbaValue::Kind::Null) return {};
        const auto date=dateParts(args[0]); int result;
        if(name=="year") result=date.year;
        else if(name=="month") result=date.month;
        else if(name=="day") result=date.day;
        else if(name=="hour") result=date.hour;
        else if(name=="minute") result=date.minute;
        else if(name=="second") result=date.second;
        else if(weekday) result=(static_cast<int>(std::chrono::weekday(date.date).c_encoding())+7-(firstDay-1))%7+1;
        else throw VbaRuntimeError(438,"Calendar function unsupported");
        return VbaValue::integer(static_cast<std::int16_t>(result));
    }
    for(const auto& value:args) if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null DateDiff argument");
    auto interval=args[0].asText(); for(auto& c:interval) if(c>='A' && c<='Z') c+=('a'-'A');
    const auto left=dateParts(args[1]),right=dateParts(args[2]);
    const auto bucket=[&](const Parts& date)->std::int64_t {
        const auto days=date.date.time_since_epoch().count();
        if(interval=="yyyy") return date.year;
        if(interval=="q") return date.year*4+(date.month-1)/3;
        if(interval=="m") return date.year*12+date.month-1;
        if(interval=="y" || interval=="d") return days;
        if(interval=="h") return days*24+date.hour;
        if(interval=="n") return days*1440+date.hour*60+date.minute;
        if(interval=="s") return days*86400+date.hour*3600+date.minute*60+date.second;
        if(interval=="ww") {
            if(firstWeek!=1) throw VbaRuntimeError(3251,"Non-default first-week DateDiff is not validated");
            return floorWeeks(days-3-(firstDay-1)); // 1970-01-04 was Sunday.
        }
        if(interval=="w") throw VbaRuntimeError(3251,"Weekday-interval DateDiff is not validated");
        throw VbaRuntimeError(5,"Invalid DateDiff interval");
    };
    const auto result=bucket(right)-bucket(left);
    if(result<std::numeric_limits<std::int32_t>::min() || result>std::numeric_limits<std::int32_t>::max()) throw VbaRuntimeError(6,"DateDiff Long overflow");
    return VbaValue::longInteger(static_cast<std::int32_t>(result));
}
}
