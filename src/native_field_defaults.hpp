#pragma once
#include <QMap>
#include <QRegularExpression>
#include <QSqlDatabase>
#include <QString>
#include <QVariant>
#include <optional>
namespace linaccess {
inline std::optional<QString> nativeDateDefaultSql(const QString& expression) {
    static const QRegularExpression shape(R"(^\s*=?\s*(Date|Now|Time)\s*\(\s*\)\s*$)",QRegularExpression::CaseInsensitiveOption);
    const auto match=shape.match(expression);if(!match.hasMatch()) return {};return "linaccess_jet_default_"+match.captured(1).toLower()+"()";
}
inline std::optional<QString> nativeDateDefaultExpression(QString sql) {
    QString token;for(const auto character:sql) if(!character.isSpace()) token+=character;token=token.toLower();while(token.startsWith('(') && token.endsWith(')')) token=token.mid(1,token.size()-2);
    if(token=="linaccess_jet_default_date()") return "Date()";if(token=="linaccess_jet_default_now()") return "Now()";if(token=="linaccess_jet_default_time()") return "Time()";return {};
}
// Bounded ISO #date# literal, optionally prefixed by =; canonical SQL string.
std::optional<QString> nativeDateLiteralSql(const QString& expression);
// Evaluates only literal Date defaults and the three owned native clock defaults.
// Never executes counter generators, VBA or an arbitrary SQL default expression.
QMap<QString,QVariant> nativeTableDateDefaults(QSqlDatabase database,const QString& table);
}
