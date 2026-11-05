#pragma once
#include "vba_expression.hpp"
#include <QApplication>
#include <QCursor>
#include <QPainter>
#include <QPixmap>
#include <QVariant>

namespace linaccess {
inline bool ownsNativeBusyCursor() { return qApp && qApp->property("la_native_busy_cursor").isValid(); }
inline bool nativeBusyCursorOnTop() {
    return ownsNativeBusyCursor() && QApplication::overrideCursor()
        && *QApplication::overrideCursor()==qApp->property("la_native_busy_cursor").value<QCursor>();
}
inline void nativeHourglass(bool enabled) {
    if(!qApp) throw VbaRuntimeError(3251,"Native cursor requires a GUI application");
    if(ownsNativeBusyCursor()) {
        if(!nativeBusyCursorOnTop()) throw VbaRuntimeError(3251,"Another Qt cursor override is active");
        if(!enabled) { QApplication::restoreOverrideCursor(); qApp->setProperty("la_native_busy_cursor",QVariant{}); }
        return;
    }
    if(!enabled) return;
    if(QApplication::overrideCursor()) throw VbaRuntimeError(3251,"Another Qt cursor override is active");
    QPixmap image(16,16); image.fill(Qt::transparent);
    {
        QPainter painter(&image); painter.setPen(Qt::black); painter.setBrush(Qt::white);
        painter.drawPolygon(QPolygon{QPoint(3,1),QPoint(12,1),QPoint(12,4),QPoint(9,7),QPoint(9,8),QPoint(12,11),QPoint(12,14),QPoint(3,14),QPoint(3,11),QPoint(6,8),QPoint(6,7),QPoint(3,4)});
        painter.fillRect(2,0,12,2,Qt::black); painter.fillRect(2,14,12,2,Qt::black);
        painter.setPen(Qt::NoPen); painter.setBrush(QColor(150,110,40));
        painter.drawPolygon(QPolygon{QPoint(5,3),QPoint(10,3),QPoint(8,6),QPoint(7,6)});
        painter.drawPolygon(QPolygon{QPoint(5,12),QPoint(10,12),QPoint(8,10),QPoint(7,10)});
    }
    const QCursor cursor(image,8,8);
    QApplication::setOverrideCursor(cursor); qApp->setProperty("la_native_busy_cursor",QVariant::fromValue(cursor));
}
inline VbaValue nativeMousePointer() {
    if(nativeBusyCursorOnTop()) return VbaValue::longInteger(11);
    if(ownsNativeBusyCursor() || QApplication::overrideCursor()) throw VbaRuntimeError(3251,"Foreign Qt cursor mapping is unsupported");
    return VbaValue::longInteger(0);
}
inline void resetNativeHourglass() noexcept {
    if(nativeBusyCursorOnTop()) { QApplication::restoreOverrideCursor(); qApp->setProperty("la_native_busy_cursor",QVariant{}); }
}
} // namespace linaccess
