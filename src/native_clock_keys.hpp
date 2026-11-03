#pragma once
#include <QDateTime>
#include <QKeyEvent>
#include <optional>
namespace linaccess {
inline std::optional<QString> nativeClockKeyText(const QKeyEvent& event) {
    const auto modifiers=event.modifiers() & ~Qt::KeypadModifier;
    const bool date=event.key()==Qt::Key_Semicolon && modifiers==Qt::ControlModifier;
    const bool time=(event.key()==Qt::Key_Semicolon || event.key()==Qt::Key_Colon)
        && modifiers==(Qt::ControlModifier|Qt::ShiftModifier);
    if(!date && !time) return {};
    return QDateTime::currentDateTime().toString(date ? "yyyy-MM-dd" : "HH:mm:ss");
}
}
