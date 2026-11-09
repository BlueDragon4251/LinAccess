#include "access_format.hpp"
#include "vba_decimal.hpp"
#include <QDate>
#include <QTime>
#include <QLocale>
#include <QRegularExpression>
#include <QStringList>
#include <QVector>
#include <cmath>
#include <algorithm>
#include <optional>
namespace linaccess {
namespace {
QString localDigits(QString value,const QLocale& locale) {
    const auto zero=locale.zeroDigit();if(zero=="0") return value;if(zero.size()!=1) throw VbaRuntimeError(3251,"Locale digit representation unsupported");
    for(auto& ch:value) if(ch>='0' && ch<='9') ch=QChar(zero[0].unicode()+ch.unicode()-'0');return value;
}
VbaValue numericValue(const VbaValue& input,const QLocale& locale) {
    const auto value=scalarVbaValue(input);if(value.kind!=VbaValue::Kind::Text) return convertDecimal(value);
    auto text=QString::fromUtf8(value.text);bool valid=false;const auto number=locale.toDouble(text,&valid);
    if(valid && std::isfinite(number)) {
        text=text.trimmed();text.remove(locale.groupSeparator());text.replace(locale.decimalPoint(),".");text.replace(locale.negativeSign(),"-");text.replace(locale.positiveSign(),"+");
        const auto zero=locale.zeroDigit();if(zero.size()==1) for(auto& ch:text) if(ch.unicode()>=zero[0].unicode() && ch.unicode()<=zero[0].unicode()+9) ch=QChar('0'+ch.unicode()-zero[0].unicode());
        return decimalLiteral(text.toStdString());
    }
    return convertDecimal(value);
}
struct CurrencyStyle {QString before,after;int digits{};};
CurrencyStyle currencyStyle(const QLocale& locale) {
    const auto sample=locale.toCurrencyString(1.0),one=locale.toString(1),zero=locale.toString(0);const int start=sample.indexOf(one);if(start<0 || one.isEmpty() || zero.isEmpty()) throw VbaRuntimeError(3251,"Locale currency pattern unsupported");int end=start+one.size(),digits=0;
    if(sample.mid(end).startsWith(locale.decimalPoint())) {end+=locale.decimalPoint().size();while(sample.mid(end).startsWith(zero)) {end+=zero.size();if(++digits>28) throw VbaRuntimeError(3251,"Locale currency precision unsupported");}}
    const auto spaces=[](QString text) {return text.replace(QChar(0x00a0),' ').replace(QChar(0x202f),' ');};
    return {spaces(sample.left(start)),spaces(sample.mid(end)),digits};
}
QString numberText(const VbaValue& input,int minimumInteger,int minimumFraction,int maximumFraction,bool group,bool leading,bool parens,bool currency,bool percent,const QLocale& locale) {
    if(maximumFraction<0 || maximumFraction>28 || minimumFraction<0 || minimumFraction>maximumFraction || minimumInteger>255) throw VbaRuntimeError(5,"Numeric format precision unsupported");
    auto value=numericValue(input,locale);if(percent) value=decimalBinary("*",value,decimalLiteral("100"));value=decimalRound(value,maximumFraction);auto text=QString::fromStdString(value.asText());const bool negative=text.startsWith('-');if(negative) text.remove(0,1);
    const auto split=text.indexOf('.');auto whole=split<0 ? text : text.left(split);auto fraction=split<0 ? QString{} : text.mid(split+1);fraction=fraction.leftJustified(maximumFraction,'0');while(fraction.size()>minimumFraction && fraction.endsWith('0')) fraction.chop(1);
    if(whole=="0" && (!leading || minimumInteger==0)) whole.clear();else whole=whole.rightJustified(minimumInteger,'0');
    if(group) for(int index=whole.size()-3;index>0;index-=3) whole.insert(index,locale.groupSeparator());
    QString result=localDigits(whole,locale);if(!fraction.isEmpty()) result+=locale.decimalPoint()+localDigits(fraction,locale);
    if(currency) {const auto style=currencyStyle(locale);result=style.before+result+style.after;}if(percent) result+=locale.percent();
    if(negative) result=parens ? "("+result+")" : locale.negativeSign()+result;return result;
}
struct NumericMaskPiece {QChar character;bool literal{};};
std::optional<QString> customNumericText(const VbaValue& input,const QString& pattern,const QLocale& locale) {
    QVector<QVector<NumericMaskPiece>> sections(1);
    bool placeholders=false;
    for(int position=0;position<pattern.size();++position) {
        const auto ch=pattern[position];
        if(ch=='\\') {
            if(++position>=pattern.size()) throw VbaRuntimeError(5,"Incomplete format escape");
            sections.last().append({pattern[position],true});
        } else if(ch=='"') {
            bool closed=false;
            while(++position<pattern.size()) {if(pattern[position]=='"') {closed=true;break;}sections.last().append({pattern[position],true});}
            if(!closed) throw VbaRuntimeError(5,"Unterminated format literal");
        } else if(ch==';') sections.append(QVector<NumericMaskPiece>{});
        else {sections.last().append({ch,false});placeholders|=ch=='0' || ch=='#';}
    }
    if(!placeholders) return {};
    if(sections.size()>3) throw VbaRuntimeError(5,"Numeric Null sections are not validated");
    auto value=numericValue(input,locale);
    const int sign=decimalCompare(value,decimalLiteral("0"));
    int selected=0;
    if(sign<0 && sections.size()>1 && !sections[1].isEmpty()) {selected=1;value=decimalUnary("abs",value);}
    else if(sign==0 && sections.size()>2 && !sections[2].isEmpty()) selected=2;
    const bool automaticNegative=sign<0 && selected==0;
    if(automaticNegative) value=decimalUnary("abs",value);
    const auto& pieces=sections[selected];
    int first=-1,last=-1,percents=0;
    for(int index=0;index<pieces.size();++index) {
        const auto& piece=pieces[index];if(piece.literal) continue;
        if(piece.character=='0' || piece.character=='#') {if(first<0) first=index;last=index;}
        else if(piece.character=='%') ++percents;
        else if(piece.character.isLetter() || piece.character=='[' || piece.character==']' || piece.character=='@' || piece.character=='&' || piece.character=='!' || piece.character=='<' || piece.character=='>') throw VbaRuntimeError(5,"Numeric mask token unsupported");
    }
    if(percents>1) throw VbaRuntimeError(5,"Repeated numeric percent tokens are not validated");
    const auto literal=[&](int begin,int end) {
        QString result;
        for(int index=begin;index<end;++index) {
            const auto& piece=pieces[index];
            if(!piece.literal && (piece.character=='.' || piece.character==',')) throw VbaRuntimeError(5,"Numeric separator outside its digit mask");
            result+=!piece.literal && piece.character=='%' ? locale.percent() : QString(piece.character);
        }
        return result;
    };
    if(first<0) return literal(0,pieces.size());
    if(first>0 && !pieces[first-1].literal && pieces[first-1].character=='.') --first;
    // Trailing commas belong to the scale, not to the literal suffix.
    while(last+1<pieces.size() && !pieces[last+1].literal && pieces[last+1].character==',') ++last;
    QString core;
    for(int index=first;index<=last;++index) {
        const auto& piece=pieces[index];
        if(piece.literal || (piece.character!='0' && piece.character!='#' && piece.character!='.' && piece.character!=',')) throw VbaRuntimeError(5,"Interleaved numeric literals are not validated");
        core+=piece.character;
    }
    if(core.count('.')>1) throw VbaRuntimeError(5,"Multiple numeric decimal points");
    const int point=core.indexOf('.');
    auto integer=point<0 ? core : core.left(point),fraction=point<0 ? QString{} : core.mid(point+1);
    int scale=0;
    while(integer.endsWith(',')) {integer.chop(1);++scale;}
    while(fraction.endsWith(',')) {fraction.chop(1);++scale;}
    if(integer.startsWith(',') || integer.contains(",,") || fraction.contains(',') || (integer.isEmpty() && point<0) || (point>=0 && fraction.isEmpty())) throw VbaRuntimeError(5,"Numeric grouping mask unsupported");
    if(scale>8) throw VbaRuntimeError(3251,"Numeric scaling exceeds the exact supported range");
    for(int index=0;index<scale;++index) {
        const auto divisor=decimalLiteral("1000"),next=decimalBinary("/",value,divisor);
        if(decimalCompare(decimalBinary("*",next,divisor),value)!=0) throw VbaRuntimeError(3251,"Numeric scaling would lose Decimal precision");
        value=next;
    }
    if(percents) value=decimalBinary("*",value,decimalLiteral("100"));
    const bool group=integer.contains(',');integer.remove(',');
    const int required=integer.indexOf('0');
    const auto rendered=numberText(value,required<0 ? 0 : integer.size()-required,fraction.lastIndexOf('0')+1,fraction.size(),group,true,false,false,false,locale);
    return (automaticNegative ? locale.negativeSign() : QString{})+literal(0,first)+rendered+literal(last+1,pieces.size());
}
struct DateParts {QDate date;QTime time;};
DateParts dateParts(const VbaValue& value) {
    DateParts result;const QDate epoch(1899,12,30);
    if(value.kind==VbaValue::Kind::Number || value.kind==VbaValue::Kind::Date || value.kind==VbaValue::Kind::Empty || value.kind==VbaValue::Kind::Decimal || value.kind==VbaValue::Kind::Currency || value.kind==VbaValue::Kind::Boolean) {
        const auto number=value.asNumber(),days=std::trunc(number);if(!std::isfinite(days) || days<epoch.daysTo(QDate(100,1,1)) || days>epoch.daysTo(QDate(9999,12,31))) throw VbaRuntimeError(6,"Access date format overflow");result.date=epoch.addDays(static_cast<qint64>(days));const auto seconds=std::min(86399,static_cast<int>(std::floor(std::abs(number-days)*86400.0+0.00005)));result.time=QTime(0,0).addSecs(seconds);
    }else if(value.kind==VbaValue::Kind::Text) {
        const auto input=QString::fromUtf8(value.text);static const QRegularExpression iso("^\\d{4}-\\d{2}-\\d{2}(?:[ T]\\d{2}:\\d{2}:\\d{2}(?:\\.\\d{1,3})?)?$");
        if(input.size()<=23 && iso.match(input).hasMatch()) {result.date=QDate::fromString(input.left(10),Qt::ISODate);result.time=input.size()>10 ? QTime::fromString(input.mid(11,8),"HH:mm:ss") : QTime(0,0);}
    }
    if(!result.date.isValid() || !result.time.isValid() || result.date.year()<100 || result.date.year()>9999) throw VbaRuntimeError(13,"Access date format requires an ISO date or OA day value");return result;
}
struct DateToken {QString key,raw;};
QString dateText(const VbaValue& value,const QString& pattern,const QLocale& locale) {
    const auto lower=pattern.toCaseFolded();
    if(lower=="long date" || lower=="short date") return locale.toString(dateParts(value).date,lower=="long date" ? QLocale::LongFormat : QLocale::ShortFormat);
    if(lower=="long time" || lower=="short time") return dateParts(value).time.toString(lower=="long time" ? "HH:mm:ss" : "HH:mm");
    static const QStringList names{"am/pm","a/p","dddddd","ddddd","dddd","ddd","dd","d","yyyy","yy","y","mmmm","mmm","mm","m","hh","h","nn","n","ss","s","ww","w","q"};QVector<DateToken> tokens;bool ampm=false,hasDate=false;
    for(int position=0;position<pattern.size();) {
        const auto ch=pattern[position];if(ch=='\\') {if(++position>=pattern.size()) throw VbaRuntimeError(5,"Incomplete format escape");tokens.append({{},pattern.mid(position++,1)});continue;}
        if(ch=='"') {QString literal;++position;bool closed=false;while(position<pattern.size()) {if(pattern[position]=='"') {++position;closed=true;break;}literal+=pattern[position++];}if(!closed) throw VbaRuntimeError(5,"Unterminated format literal");tokens.append({{},literal});continue;}
        bool found=false;for(const auto& name:names) if(lower.mid(position,name.size())==name) {tokens.append({name,pattern.mid(position,name.size())});ampm|=name=="am/pm" || name=="a/p";position+=name.size();found=true;hasDate=true;break;}if(found) continue;
        if(ch.isLetter() || ch=='0' || ch=='#' || ch==';' || ch=='@' || ch=='&' || ch=='!' || ch=='<' || ch=='>') throw VbaRuntimeError(5,"Access format token unsupported");tokens.append({{},pattern.mid(position++,1)});
    }
    if(!hasDate) {QString result;for(const auto& token:tokens) result+=token.raw;return result;}const auto parts=dateParts(value);const auto date=parts.date;const auto time=parts.time;QString result,previous;
    const auto number=[&](int n,int width=1) {return localDigits(QString::number(n).rightJustified(width,'0'),locale);};
    for(const auto& token:tokens) {const auto& key=token.key;if(key.isEmpty()) {result+=token.raw;continue;}
        if(key=="yyyy") result+=number(date.year(),4);else if(key=="yy") result+=number(date.year()%100,2);else if(key=="y") result+=number(date.dayOfYear());
        else if(key=="ddddd" || key=="dddddd") result+=locale.toString(date,key=="ddddd" ? QLocale::ShortFormat : QLocale::LongFormat);
        else if(key=="d" || key=="dd") result+=number(date.day(),key.size());else if(key=="ddd" || key=="dddd") result+=locale.dayName(date.dayOfWeek(),key=="ddd" ? QLocale::ShortFormat : QLocale::LongFormat);
        else if(key=="m" || key=="mm") result+=number(previous=="h" || previous=="hh" ? time.minute() : date.month(),key.size());else if(key=="mmm" || key=="mmmm") result+=locale.monthName(date.month(),key=="mmm" ? QLocale::ShortFormat : QLocale::LongFormat);
        else if(key=="h" || key=="hh") result+=number(ampm ? (time.hour()%12==0 ? 12 : time.hour()%12) : time.hour(),key.size());else if(key=="n" || key=="nn") result+=number(time.minute(),key.size());else if(key=="s" || key=="ss") result+=number(time.second(),key.size());
        else if(key=="q") result+=number((date.month()-1)/3+1);else if(key=="w") result+=number(date.dayOfWeek()%7+1);else if(key=="ww") result+=number((date.dayOfYear()-1+QDate(date.year(),1,1).dayOfWeek()%7)/7+1);
        else {auto marker=time.hour()<12 ? locale.amText() : locale.pmText();if(key=="a/p") marker=marker.left(1);if(token.raw==token.raw.toLower()) marker=marker.toLower();else if(token.raw==token.raw.toUpper()) marker=marker.toUpper();result+=marker;}previous=key;
    }
    return result;
}
}
bool isAccessNumericText(std::string_view text) {if(text.size()>4096) return false;const auto string=QString::fromUtf8(text.data(),text.size());if(!string.isValidUtf16() || string.contains(QChar(0))) return false;bool valid=false;const auto number=QLocale().toDouble(string,&valid);return valid && std::isfinite(number);}
VbaValue formatAccessNumericValue(const VbaValue& input,bool currency,const VbaNumericFormatOptions& options) {
    const auto value=scalarVbaValue(input);if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Numeric format: invalid use of Null");const QLocale locale;const auto digits=options.digits.value_or(currency ? currencyStyle(locale).digits : 2);
    return VbaValue::string(numberText(value,1,digits,digits,options.group.value_or(true),options.leading.value_or(true),options.parens.value_or(false),currency,false,locale).toUtf8().toStdString());
}
VbaValue formatAccessValue(const VbaValue& input,std::string_view mask,bool stringResult) {
    const auto value=scalarVbaValue(input);if(value.kind==VbaValue::Kind::Null) {if(stringResult) throw VbaRuntimeError(94,"Format$: invalid use of Null");return {};}
    const auto pattern=QString::fromUtf8(mask.data(),mask.size()).trimmed(),lower=pattern.toCaseFolded();if(mask.size()>1024 || pattern.isEmpty() || pattern.contains(QChar(0)) || !pattern.isValidUtf16()) throw VbaRuntimeError(5,"Access format mask invalid");const QLocale locale;
    if(lower=="yes/no" || lower=="true/false" || lower=="on/off") {
        const bool on=decimalCompare(numericValue(value,locale),decimalLiteral("0"))!=0;QString result;if(locale.language()==QLocale::German) result=lower=="yes/no" ? (on ? "Ja" : "Nein") : lower=="true/false" ? (on ? "Wahr" : "Falsch") : (on ? "Ein" : "Aus");else if(locale.language()==QLocale::English || locale.language()==QLocale::C) result=lower=="yes/no" ? (on ? "Yes" : "No") : lower=="true/false" ? (on ? "True" : "False") : (on ? "On" : "Off");else throw VbaRuntimeError(3251,"Boolean format language unsupported");return VbaValue::string(result.toStdString());
    }
    if(lower=="currency" || lower=="standard" || lower=="fixed") {VbaNumericFormatOptions options;if(lower!="currency") {options.digits=2;options.group=lower=="standard";}return formatAccessNumericValue(value,lower=="currency",options);}
    if(const auto numeric=customNumericText(value,pattern,locale)) return VbaValue::string(numeric->toUtf8().toStdString());
    return VbaValue::string(dateText(value,pattern,locale).toUtf8().toStdString());
}
}
