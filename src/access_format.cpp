#include "access_format.hpp"
#include <QDate>
#include <QTime>
#include <QLocale>
#include <QRegularExpression>
#include <cmath>
namespace linaccess {
VbaValue formatAccessValue(const VbaValue& value,std::string_view mask,bool stringResult) {
    if (value.kind==VbaValue::Kind::Null && stringResult)
        throw VbaRuntimeError(94,"Format$: invalid use of Null");
    const auto pattern=QString::fromUtf8(mask.data(),mask.size()).trimmed().toCaseFolded();
    if (pattern!="yyyy" && pattern!="mm" && pattern!="dd" && pattern!="long date")
        throw VbaRuntimeError(5,"Access date format mask unsupported");
    if (value.kind==VbaValue::Kind::Null) return {};
    QDate date;
    if (value.kind==VbaValue::Kind::Number || value.kind==VbaValue::Kind::Date || value.kind==VbaValue::Kind::Empty) {
        const auto number=value.kind==VbaValue::Kind::Empty ? 0 : value.number;
        const QDate epoch(1899,12,30);
        const auto days=std::trunc(number); // negative OA fractions remain on the integral day
        if (!std::isfinite(days) || days<epoch.daysTo(QDate(100,1,1)) || days>epoch.daysTo(QDate(9999,12,31)))
            throw VbaRuntimeError(6,"Access date format overflow");
        date=epoch.addDays(static_cast<qint64>(days));
    } else if (value.kind==VbaValue::Kind::Text) {
        const auto input=QString::fromUtf8(value.text.data(),value.text.size());
        static const QRegularExpression iso("^\\d{4}-\\d{2}-\\d{2}(?:[ T]\\d{2}:\\d{2}:\\d{2}(?:\\.\\d{1,3})?)?$");
        if (input.size()<=23 && iso.match(input).hasMatch()) {
            date=QDate::fromString(input.left(10),Qt::ISODate);
            // Jet dates have no time zone. A valid wall-clock time must not be
            // rejected because it falls into the host's daylight-saving gap.
            if (input.size()>10 && !QTime::fromString(input.mid(11,8),"HH:mm:ss").isValid()) date={};
        }
    }
    if (!date.isValid() || date.year()<100 || date.year()>9999)
        throw VbaRuntimeError(13,"Access date format requires an ISO date or OA day value");
    const auto result=pattern=="long date" ? QLocale().toString(date,QLocale::LongFormat)
        : date.toString(pattern=="yyyy" ? "yyyy" : pattern=="mm" ? "MM" : "dd");
    return VbaValue::string(result.toUtf8().toStdString());
}
}
