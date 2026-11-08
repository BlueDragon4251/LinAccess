#pragma once
#include "vba_expression.hpp"
#include "native_field_defaults.hpp"
#include "vba_qt_value.hpp"
#include <QRegularExpression>
#include <QString>
#include <QStringList>
#include <cmath>
namespace linaccess {
// Bounded numeric field rules: comparisons joined by And/Or, or Between.
// No identifiers, subqueries, calls, locale-dependent numbers or implicit SQL.
inline QString nativeNumericValidationPredicate(QString rule,const QString& column) {
    if(rule.isEmpty()) return {};
    if(rule.size()>255 || rule.contains(QChar(0))) throw VbaRuntimeError(3251,"Native numeric ValidationRule exceeds the supported shape");
    const QString number=QStringLiteral(R"([+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?)");
    const auto literal=[](const QString& token) {bool valid=false;const auto value=token.toDouble(&valid);if(!valid || !std::isfinite(value)) throw VbaRuntimeError(3251,"ValidationRule numeric literal is not representable");return QString::number(value,'g',17);};
    const QRegularExpression between("^\\s*Between\\s+("+number+")\\s+And\\s+("+number+")\\s*$",QRegularExpression::CaseInsensitiveOption);
    const auto bounds=between.match(rule);if(bounds.hasMatch()) return "("+column+">="+literal(bounds.captured(1))+" AND "+column+"<="+literal(bounds.captured(2))+")";
    const QRegularExpression term("\\G\\s*(>=|<=|<>|=|>|<)\\s*("+number+")\\s*(?:(And|Or)\\s+|$)",QRegularExpression::CaseInsensitiveOption);
    QStringList parts;qsizetype offset=0;int terms=0;bool connector=false;
    while(offset<rule.size()) {const auto match=term.match(rule,offset);if(!match.hasMatch() || match.capturedStart()!=offset || ++terms>16) throw VbaRuntimeError(3251,"Native ValidationRule requires numeric comparisons joined by And/Or");parts.append(column+match.captured(1)+literal(match.captured(2)));connector=!match.captured(3).isEmpty();if(connector) parts.append(match.captured(3).toUpper());offset=match.capturedEnd();}
    if(parts.isEmpty() || connector) throw VbaRuntimeError(3251,"Native ValidationRule comparison is incomplete");return "("+parts.join(' ')+")";
}
inline QString nativeDateValidationPredicate(QString rule,const QString& column) {
    if(rule.isEmpty()) return {};
    if(rule.size()>255 || rule.contains(QChar(0))) throw VbaRuntimeError(3251,"Native Date ValidationRule exceeds the supported shape");
    const QString atom=QStringLiteral(R"((?:#[^#]+#|(?:Date|Now|Time)\s*\(\s*\)))");
    const auto value=[](QString token) {
        if(const auto clock=nativeDateDefaultSql(token)) return *clock;
        if(const auto literal=nativeDateLiteralSql(token)) return *literal;
        throw VbaRuntimeError(3251,"Date ValidationRule requires an ISO Date literal or Date/Now/Time clock");
    };
    const QRegularExpression between("^\\s*Between\\s+("+atom+")\\s+And\\s+("+atom+")\\s*$",QRegularExpression::CaseInsensitiveOption);
    const auto bounds=between.match(rule);if(bounds.hasMatch()) return "("+column+">="+value(bounds.captured(1))+" AND "+column+"<="+value(bounds.captured(2))+")";
    const QRegularExpression term("\\G\\s*(>=|<=|<>|=|>|<)\\s*("+atom+")\\s*(?:(And|Or)\\s+|$)",QRegularExpression::CaseInsensitiveOption);
    QStringList parts;qsizetype offset=0;int terms=0;bool connector=false;
    while(offset<rule.size()) {const auto match=term.match(rule,offset);if(!match.hasMatch() || match.capturedStart()!=offset || ++terms>16) throw VbaRuntimeError(3251,"Native Date ValidationRule requires Date comparisons joined by And/Or");parts.append(column+match.captured(1)+value(match.captured(2)));connector=!match.captured(3).isEmpty();if(connector) parts.append(match.captured(3).toUpper());offset=match.capturedEnd();}
    if(parts.isEmpty() || connector) throw VbaRuntimeError(3251,"Native Date ValidationRule comparison is incomplete");return "("+parts.join(' ')+")";
}
inline QString nativeFieldValidationPredicate(int type,const QString& rule,const QString& column) {
    return type==8 ? nativeDateValidationPredicate(rule,column) : nativeNumericValidationPredicate(rule,column);
}

}
