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
#include <sstream>
#include <iomanip>
#include <locale>
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
    // Access 2003 prefixes the symbol in the observed Austrian German profile.
    if(locale.language()==QLocale::German && locale.territory()==QLocale::Austria)
        return {locale.currencySymbol()+QLatin1Char(' '),{},2};
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
QString generalNumberText(const VbaValue& input,const QLocale& locale) {
    const auto value=scalarVbaValue(input);
    if(value.kind==VbaValue::Kind::Number || value.kind==VbaValue::Kind::Date) {
        if(!std::isfinite(value.number)) throw VbaRuntimeError(6,"General Number is not finite");
        auto ungrouped=locale;ungrouped.setNumberOptions(ungrouped.numberOptions()|QLocale::OmitGroupSeparator);
        return ungrouped.toString(value.number==0 ? 0.0 : value.number,'g',value.numericType==VbaValue::NumericType::Single ? 7 : 15).replace(ungrouped.exponential(),"E");
    }
    auto text=QString::fromStdString(numericValue(value,locale).asText());
    if(text.contains('.')) {while(text.endsWith('0')) text.chop(1);if(text.endsWith('.')) text.chop(1);}
    const bool negative=text.startsWith('-');if(negative) text.remove(0,1);
    text.replace('.',locale.decimalPoint());
    return (negative ? locale.negativeSign() : QString{})+localDigits(text,locale);
}
struct ScientificParts {VbaValue mantissa;int exponent{};};
ScientificParts scientificParts(const VbaValue& input,const QLocale& locale) {
    const auto value=scalarVbaValue(input);
    if(value.kind==VbaValue::Kind::Number || value.kind==VbaValue::Kind::Date) {
        if(!std::isfinite(value.number)) throw VbaRuntimeError(6,"Scientific value is not finite");
        std::ostringstream stream;stream.imbue(std::locale::classic());
        stream<<std::scientific<<std::setprecision(value.numericType==VbaValue::NumericType::Single ? 6 : 14)<<value.number;
        const auto text=stream.str();const auto marker=text.find('e');
        if(marker==std::string::npos) throw VbaRuntimeError(5,"Scientific conversion failed");
        return {decimalLiteral(text.substr(0,marker)),std::stoi(text.substr(marker+1))};
    }
    auto text=QString::fromStdString(numericValue(value,locale).asText());const bool negative=text.startsWith('-');if(negative) text.remove(0,1);
    int point=text.indexOf('.');if(point<0) point=text.size();text.remove('.');
    int first=0;while(first<text.size() && text[first]=='0') ++first;
    if(first==text.size()) return {decimalLiteral("0"),0};
    const int exponent=point-first-1;const auto digits=text.mid(first);
    const auto normalized=(negative ? QString("-") : QString{})+digits.left(1)+(digits.size()>1 ? "."+digits.mid(1) : QString{});
    return {decimalLiteral(normalized.toStdString()),exponent};
}
QString scientificText(const VbaValue& value,const QString& fraction,QChar marker,QChar sign,int width,const QLocale& locale) {
    if(fraction.size()>28 || width<1 || width>255) throw VbaRuntimeError(5,"Scientific mask precision unsupported");
    auto parts=scientificParts(value,locale);parts.mantissa=decimalRound(parts.mantissa,fraction.size());
    if(decimalCompare(decimalUnary("abs",parts.mantissa),decimalLiteral("10"))>=0) {parts.mantissa=decimalBinary("/",parts.mantissa,decimalLiteral("10"));++parts.exponent;}
    const auto mantissa=numberText(parts.mantissa,1,fraction.lastIndexOf('0')+1,fraction.size(),false,true,false,false,false,locale);
    const auto exponentSign=parts.exponent<0 ? locale.negativeSign() : sign=='+' ? locale.positiveSign() : QString{};
    return mantissa+marker+exponentSign+localDigits(QString::number(std::abs(parts.exponent)).rightJustified(width,'0'),locale);
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
std::optional<QString> customStringText(const VbaValue& value,const QString& pattern,const QLocale& locale) {
    QVector<QVector<NumericMaskPiece>> sections(1);
    bool directive=false,placeholder=false,numeric=false;
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
        else {
            sections.last().append({ch,false});
            placeholder|=ch=='@' || ch=='&';directive|=ch=='@' || ch=='&' || ch=='<' || ch=='>' || ch=='!';numeric|=ch=='0' || ch=='#';
        }
    }
    if(!directive || (numeric && !placeholder)) return {};
    if(sections.size()>2) throw VbaRuntimeError(5,"String masks support at most two sections");
    if(value.kind!=VbaValue::Kind::Text && value.kind!=VbaValue::Kind::Empty) throw VbaRuntimeError(13,"String mask requires a String or Empty value");
    constexpr qsizetype limit=16*1024*1024;
    if(value.text.size()>static_cast<std::size_t>(limit)*4) throw VbaRuntimeError(3251,"String format input exceeds its native bound");
    const auto bytes=QByteArray::fromStdString(value.kind==VbaValue::Kind::Empty ? std::string{} : value.text);
    auto text=QString::fromUtf8(bytes);
    if(!text.isValidUtf16() || text.toUtf8()!=bytes) throw VbaRuntimeError(13,"Invalid Unicode string format input");
    if(text.size()>limit || text.contains(QChar(0))) throw VbaRuntimeError(3251,"String format input is not representable");
    const auto& pieces=sections[text.isEmpty() && sections.size()>1 ? 1 : 0];
    int casing=0,count=0;bool left=false;
    QStringList rendered;rendered.resize(pieces.size());
    for(int index=0;index<pieces.size();++index) {
        const auto& piece=pieces[index];const auto ch=piece.character;
        if(!piece.literal && (ch=='<' || ch=='>')) {
            const int next=ch=='<' ? 1 : 2;if(casing && casing!=next) throw VbaRuntimeError(3251,"Mixed string case directives are not validated");casing=next;
        } else if(!piece.literal && ch=='!') left=true;
        else if(!piece.literal && (ch=='@' || ch=='&')) ++count;
        else rendered[index]=QString(ch);
    }
    const auto converted=casing==1 ? locale.toLower(text) : casing==2 ? locale.toUpper(text) : text;
    if(count && converted.size()!=text.size()) throw VbaRuntimeError(3251,"Length-changing case conversion with string placeholders is not validated");
    text=converted;
    if(!count) {
        const auto literal=rendered.join(QString{});
        if(text.isEmpty()) return literal;
        if(!literal.isEmpty()) throw VbaRuntimeError(3251,"Literal-only case masks with nonempty input are not validated");
        if(text.size()>limit) throw VbaRuntimeError(3251,"String format output exceeds its native bound");
        return text;
    }
    int cursor=left ? 0 : text.size()-1;
    for(int offset=0;offset<pieces.size();++offset) {
        const int index=left ? offset : pieces.size()-offset-1;const auto& piece=pieces[index];
        if(piece.literal || (piece.character!='@' && piece.character!='&')) continue;
        if(cursor>=0 && cursor<text.size()) {rendered[index]=QString(text[cursor]);cursor+=left ? 1 : -1;}
        else if(piece.character=='@') rendered[index]=" ";
    }
    const auto result=(left ? QString{} : text.left(cursor+1))+rendered.join(QString{})+(left ? text.mid(cursor) : QString{});
    if(result.size()>limit || !result.isValidUtf16()) throw VbaRuntimeError(3251,"String mask would split Unicode characters or exceed its output bound");
    return result;
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
QString longTimeText(const QTime& time,const QLocale& locale) {
    const auto pattern=locale.timeFormat(QLocale::LongFormat);
    QString local;
    bool quoted=false;
    for(int position=0;position<pattern.size();++position) {
        const auto ch=pattern[position];
        if(ch==QChar(39)) {
            local+=ch;
            if(position+1<pattern.size() && pattern[position+1]==QChar(39)) {local+=pattern[++position];continue;}
            quoted=!quoted;
        } else if(!quoted && ch=='t') {
            while(position+1<pattern.size() && pattern[position+1]=='t') ++position;
        } else local+=ch;
    }
    return locale.toString(time,local.trimmed().replace(QChar(0x00a0),' ').replace(QChar(0x202f),' '));
}
QString dateText(const VbaValue& value,const QString& pattern,const QLocale& locale) {
    const auto lower=pattern.toCaseFolded();
    if(lower=="general date") {
        const auto parts=dateParts(value);
        if(parts.date==QDate(1899,12,30)) return longTimeText(parts.time,locale);
        const auto date=locale.toString(parts.date,QLocale::ShortFormat);
        return parts.time==QTime(0,0) ? date : date+" "+longTimeText(parts.time,locale);
    }
    if(lower=="long date" || lower=="short date") return locale.toString(dateParts(value).date,lower=="long date" ? QLocale::LongFormat : QLocale::ShortFormat);
    if(lower=="long time" || lower=="short time") {const auto time=dateParts(value).time;return lower=="long time" ? longTimeText(time,locale) : time.toString("HH:mm");}
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
    if(lower=="general number") return VbaValue::string(generalNumberText(value,locale).toUtf8().toStdString());
    if(lower=="scientific") return VbaValue::string(scientificText(value,"00",'E','+',2,locale).toUtf8().toStdString());
    static const QRegularExpression scientific("^0(?:\\.([0#]+))?([Ee])([+-])(0+)$");
    if(const auto match=scientific.match(pattern);match.hasMatch()) return VbaValue::string(scientificText(value,match.captured(1),match.captured(2)[0],match.captured(3)[0],match.captured(4).size(),locale).toUtf8().toStdString());
    if(lower=="percent") return VbaValue::string(numberText(value,1,2,2,false,true,false,false,true,locale).toUtf8().toStdString());
    if(lower=="currency" || lower=="standard" || lower=="fixed") {VbaNumericFormatOptions options;if(lower!="currency") {options.digits=2;options.group=lower=="standard";}return formatAccessNumericValue(value,lower=="currency",options);}
    if(const auto string=customStringText(value,pattern,locale)) return VbaValue::string(string->toUtf8().toStdString());
    if(const auto numeric=customNumericText(value,pattern,locale)) return VbaValue::string(numeric->toUtf8().toStdString());
    return VbaValue::string(dateText(value,pattern,locale).toUtf8().toStdString());
}
VbaValue accessCalendarName(bool month,int number,bool abbreviated,int firstDay) {
    if(number<1 || number>(month ? 12 : 7) || firstDay<0 || firstDay>7) throw VbaRuntimeError(5,"Invalid calendar name settings");
    const QLocale locale;const auto style=abbreviated ? QLocale::ShortFormat : QLocale::LongFormat;
    QString name;
    if(month) name=locale.standaloneMonthName(number,style);
    else {
        const int first=firstDay==0 ? static_cast<int>(locale.firstDayOfWeek())%7 : firstDay-1;
        const int sundayBased=(first+number-1)%7;
        name=locale.standaloneDayName((sundayBased+6)%7+1,style);
    }
    if(name.isEmpty()) throw VbaRuntimeError(3251,"Native calendar name is unavailable");
    return VbaValue::string(name.toUtf8().toStdString());
}

}
