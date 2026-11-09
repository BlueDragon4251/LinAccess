#pragma once
#include "vba_expression.hpp"
#include <QString>
#include <QRegularExpression>
#include <cmath>
namespace linaccess {
inline VbaValue nativeControlDefaultLiteral(QString expression) {
    expression=expression.trimmed(); if(expression.startsWith('=')) expression=expression.mid(1).trimmed();
    if(expression.compare("Null",Qt::CaseInsensitive)==0) return {};
    if(expression.compare("True",Qt::CaseInsensitive)==0 || expression.compare("Yes",Qt::CaseInsensitive)==0) return linaccess::VbaValue::logical(true);
    if(expression.compare("False",Qt::CaseInsensitive)==0 || expression.compare("No",Qt::CaseInsensitive)==0) return linaccess::VbaValue::logical(false);
    static const QRegularExpression quoted(R"literal(^"(?:[^"]|"")*"$)literal");
    static const QRegularExpression number(R"(^[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?$)");
    if(quoted.match(expression).hasMatch()) return linaccess::VbaValue::string(expression.mid(1,expression.size()-2).replace("\"\"","\"").toUtf8().toStdString());
    if(number.match(expression).hasMatch()) {
        bool valid=false; const auto value=expression.toDouble(&valid);
        if(!valid || !std::isfinite(value)) throw linaccess::VbaRuntimeError(6,"Default numeric literal overflow");
        return linaccess::VbaValue::numeric(value);
    }
    throw linaccess::VbaRuntimeError(3251,"Computed/default expression unsupported");
}
}
