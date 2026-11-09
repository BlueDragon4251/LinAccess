#pragma once
#include "vba_runtime.hpp"
#include <QString>
#include <QVector>
#include <QRegularExpression>
#include <cmath>
namespace linaccess {
// Shared native designer/runtime interpretation: bare values are twips.
inline QVector<int> nativeColumnWidths(const QString& text) {
    QVector<int> widths;
    const auto specifications=text.split(';');
    if(specifications.size()>255) throw VbaRuntimeError(5,"Too many column widths");
    for(auto specification:specifications) {
        specification=specification.trimmed();
        if(specification.isEmpty()) {widths.append(-1);continue;}
        static const QRegularExpression pattern(R"(^([0-9]+(?:[.,][0-9]+)?)\s*(cm|mm|in|pt)?$)",QRegularExpression::CaseInsensitiveOption);
        const auto match=pattern.match(specification);
        if(!match.hasMatch()) throw VbaRuntimeError(5,"Invalid column width");
        auto decimal=match.captured(1);decimal.replace(',','.');
        const auto amount=decimal.toDouble();const auto unit=match.captured(2).toLower();
        const auto scale=unit=="cm" ? 96.0/2.54 : unit=="mm" ? 96.0/25.4 : unit=="in" ? 96.0 : unit=="pt" ? 96.0/72.0 : 1.0/15.0;
        if(!std::isfinite(amount) || amount*scale>32767) throw VbaRuntimeError(5,"Column width out of range");
        widths.append(qRound(amount*scale));
    }
    return widths;
}
}
