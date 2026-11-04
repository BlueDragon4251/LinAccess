#pragma once

#include <QApplication>
#include <QFontDatabase>
#include <QPalette>
#include <QSettings>
#include <QStyle>
#include <QStyleFactory>
#include <QVariant>

namespace linaccess {
inline bool classicStyleAvailable() {
    return QStyleFactory::keys().contains("Windows",Qt::CaseInsensitive);
}
inline bool setClassicStyle(bool enabled) {
    auto* application=qobject_cast<QApplication*>(QCoreApplication::instance());
    if (!application) return false;
    if (!application->property("la_original_style").isValid()) {
        application->setProperty("la_original_style",application->style()->objectName());
        application->setProperty("la_original_palette",QVariant::fromValue(application->palette()));
        application->setProperty("la_original_font",QVariant::fromValue(application->font()));
    }
    const auto name=enabled ? QString("Windows") : application->property("la_original_style").toString();
    auto* style=QStyleFactory::create(name);
    if (!style) return false;
    application->setStyle(style);
    application->setPalette(enabled ? style->standardPalette()
        : application->property("la_original_palette").value<QPalette>());
    auto font=application->property("la_original_font").value<QFont>();
    if (enabled) {
        if (QFontDatabase::families().contains("Tahoma",Qt::CaseInsensitive)) font.setFamily("Tahoma");
        font.setPointSize(8);
    }
    const auto configuredFont=QSettings().value("ui/applicationFont");
    if(configuredFont.canConvert<QFont>()) font=configuredFont.value<QFont>();
    application->setFont(font);
    application->setProperty("la_classic_style",enabled);
    return true;
}
inline void initializeClassicStyle() {
    if (QSettings().value("ui/classicStyle",true).toBool()) (void)setClassicStyle(true);
    const auto configuredFont=QSettings().value("ui/applicationFont");
    if(configuredFont.canConvert<QFont>()) qApp->setFont(configuredFont.value<QFont>());
}
}
