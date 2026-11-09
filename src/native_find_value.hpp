#pragma once
#include "vba_qt_value.hpp"
#include "vba_like.hpp"
#include <QVariantMap>
#include <cmath>
#include <limits>
#include <optional>
namespace linaccess {
inline bool nativeFindHasWildcards(const QString& text) {
    return text.contains('*') || text.contains('?') || text.contains('#') || text.contains('[');
}
inline bool nativeFindTextMatches(const QString& candidate,const QString& text,int mode,bool sensitive,bool wildcards) {
    if(!wildcards) {const auto cs=sensitive ? Qt::CaseSensitive : Qt::CaseInsensitive;return mode==1 ? candidate.compare(text,cs)==0 : mode==2 ? candidate.startsWith(text,cs) : candidate.contains(text,cs);}
    if(mode<0 || mode>2) throw VbaRuntimeError(5,"Invalid wildcard match mode");
    if(candidate.contains(QChar(0)) || text.contains(QChar(0)) || !candidate.isValidUtf16() || !text.isValidUtf16()) throw VbaRuntimeError(3251,"Wildcard search text is not representable");
    auto pattern=text;if(mode!=1) pattern+='*';if(mode==0) pattern.prepend('*');
    if(pattern.size()>16384 || (!pattern.isEmpty() && candidate.size()>1000000/pattern.size())) throw VbaRuntimeError(3251,"Wildcard search exceeds its native pattern bound");
    const auto input=candidate.toStdU16String(),mask=pattern.toStdU16String();
    return matchVbaLikePattern(input,mask,!sensitive,[](char16_t left,char16_t right) {return QString(QChar(left)).compare(QString(QChar(right)),Qt::CaseInsensitive);});
}
inline bool nativeFindTextType(const QString& type) {
    return type.contains("TEXT",Qt::CaseInsensitive) || type.contains("CHAR",Qt::CaseInsensitive) || type.contains("CLOB",Qt::CaseInsensitive);
}
enum class NativeFindFieldKind { Unsupported,Text,Number,Currency,Date,Boolean };
inline NativeFindFieldKind nativeFindFieldKind(const QString& sqlType) {
    const auto type=sqlType.toUpper();
    if(nativeFindTextType(type)) return NativeFindFieldKind::Text;
    if(type.contains("CURRENCY")) return NativeFindFieldKind::Currency;
    if(type.contains("DATE")) return NativeFindFieldKind::Date;
    if(type.contains("BOOL")) return NativeFindFieldKind::Boolean;
    if(type.contains("INT") || type.contains("REAL") || type.contains("FLOA") || type.contains("DOUB") || type.contains("NUM") || type.contains("DEC")) return NativeFindFieldKind::Number;
    return NativeFindFieldKind::Unsupported;
}
inline bool nativeFindSupports(const VbaValue& sought,const QString& sqlType) {
    switch(nativeFindFieldKind(sqlType)) {
        case NativeFindFieldKind::Text:return sought.kind==VbaValue::Kind::Text;
        case NativeFindFieldKind::Number:return sought.kind==VbaValue::Kind::Number;
        case NativeFindFieldKind::Currency:return sought.kind==VbaValue::Kind::Currency;
        case NativeFindFieldKind::Date:return sought.kind==VbaValue::Kind::Date;
        case NativeFindFieldKind::Boolean:return sought.kind==VbaValue::Kind::Boolean;
        default:return false;
    }
}
inline QString nativeFindInputHint(const QString& sqlType) {
    switch(nativeFindFieldKind(sqlType)) {
        case NativeFindFieldKind::Number:return sqlType.contains("INT",Qt::CaseInsensitive) ? QStringLiteral("Ganze Zahlen eingeben.") : QStringLiteral("Zahlen mit Dezimalpunkt eingeben.");
        case NativeFindFieldKind::Currency:return "Währung mit Dezimalpunkt und bis zu vier Nachkommastellen eingeben.";
        case NativeFindFieldKind::Date:return "Datum: JJJJ-MM-TT, optional hh:mm:ss; Uhrzeit: hh:mm:ss.";
        case NativeFindFieldKind::Boolean:return "Ja/Nein: True oder False (auch -1 oder 0).";
        default:return "Text wird wörtlich gesucht.";
    }
}
inline VbaValue normalizeNativeFindValue(const VbaValue& sought) {
    if(sought.kind==VbaValue::Kind::Number) {
        if(!std::isfinite(sought.number)) throw VbaRuntimeError(6,"Nonfinite FindRecord number");
        return sought;
    }
    if(sought.kind==VbaValue::Kind::Date)
        return qtFieldValue(sqlFieldValue(sought,"DATETIME"),"DATETIME");
    if(sought.kind==VbaValue::Kind::Text || sought.kind==VbaValue::Kind::Currency || sought.kind==VbaValue::Kind::Boolean) return sought;
    throw VbaRuntimeError(3251,"FindRecord value subtype is not implemented");
}
inline VbaValue nativeFindInput(const QString& text,const QString& sqlType) {
    const auto value=qtFieldValue(text,sqlType);
    if(value.kind==VbaValue::Kind::Null) throw VbaRuntimeError(13,"Search value is empty for the selected field");
    return normalizeNativeFindValue(value);
}
inline bool nativeFindEqual(const VbaValue& candidate,const VbaValue& sought) {
    if(candidate.kind==VbaValue::Kind::Null) return false;
    if(candidate.kind!=sought.kind) throw VbaRuntimeError(13,"FindRecord field/value subtype mismatch");
    if(sought.kind==VbaValue::Kind::Currency) return candidate.currencyUnits==sought.currencyUnits;
    if(sought.kind==VbaValue::Kind::Boolean) return candidate.boolean==sought.boolean;
    if(sought.kind==VbaValue::Kind::Number || sought.kind==VbaValue::Kind::Date) return candidate.number==sought.number;
    throw VbaRuntimeError(3251,"Typed FindRecord comparison is not implemented");
}
inline bool sameNativeFindValue(const std::optional<VbaValue>& left,const std::optional<VbaValue>& right) {
    if(!left || !right) return !left && !right;
    return left->kind==right->kind && nativeFindEqual(*left,*right);
}
inline QString nativeFindValueText(const VbaValue& value) {
    return value.kind==VbaValue::Kind::Number ? QString::number(value.number,'g',std::numeric_limits<double>::max_digits10)
        : QString::fromStdString(value.asText());
}
inline QVariantMap nativeFindValuePayload(const VbaValue& input) {
    const auto value=normalizeNativeFindValue(input);
    QVariantMap result{{"kind",static_cast<int>(value.kind)}};
    if(value.kind==VbaValue::Kind::Number || value.kind==VbaValue::Kind::Date) result.insert("number",value.number);
    else if(value.kind==VbaValue::Kind::Currency) result.insert("currency",QVariant::fromValue(static_cast<qlonglong>(value.currencyUnits)));
    else if(value.kind==VbaValue::Kind::Boolean) result.insert("boolean",value.boolean);
    else result.insert("text",QString::fromStdString(value.text));
    return result;
}
inline VbaValue nativeFindValueFromPayload(const QVariantMap& request) {
    bool valid=false; const int kind=request.value("kind").toInt(&valid);
    if(!valid) throw VbaRuntimeError(5,"Missing FindRecord value kind");
    if(kind==static_cast<int>(VbaValue::Kind::Text)) {
        if(request.value("text").metaType().id()!=QMetaType::QString) throw VbaRuntimeError(13,"Invalid FindRecord text");
        return VbaValue::string(request.value("text").toString().toUtf8().toStdString());
    }
    if(kind==static_cast<int>(VbaValue::Kind::Currency)) {
        const auto units=request.value("currency").toLongLong(&valid);
        if(!valid) throw VbaRuntimeError(13,"Invalid FindRecord Currency");
        return VbaValue::currency(units);
    }
    if(kind==static_cast<int>(VbaValue::Kind::Boolean)) {
        if(request.value("boolean").metaType().id()!=QMetaType::Bool) throw VbaRuntimeError(13,"Invalid FindRecord Boolean");
        return VbaValue::logical(request.value("boolean").toBool());
    }
    if(kind==static_cast<int>(VbaValue::Kind::Number) || kind==static_cast<int>(VbaValue::Kind::Date)) {
        const double number=request.value("number").toDouble(&valid);
        if(!valid || !std::isfinite(number)) throw VbaRuntimeError(6,"Invalid FindRecord number");
        return normalizeNativeFindValue(kind==static_cast<int>(VbaValue::Kind::Date) ? VbaValue::dateSerial(number) : VbaValue::numeric(number));
    }
    throw VbaRuntimeError(3251,"FindRecord value kind is not implemented");
}
}
