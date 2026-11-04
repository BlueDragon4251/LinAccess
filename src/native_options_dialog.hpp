#pragma once
#include "native_classic_style.hpp"
#include <QAction>
#include <QCheckBox>
#include <QDialog>
#include <QDialogButtonBox>
#include <QFontComboBox>
#include <QFormLayout>
#include <QMessageBox>
#include <QPushButton>
#include <QSignalBlocker>
#include <QSpinBox>

namespace linaccess {
inline void showNativeOptions(QWidget* parent,QAction* classicAction) {
    QDialog dialog(parent); dialog.setObjectName("application-options-dialog");
    dialog.setWindowTitle("Optionen"); auto* layout=new QFormLayout(&dialog);
    auto* classic=new QCheckBox("Klassische Oberfläche",&dialog);
    classic->setObjectName("options-classic-style");
    classic->setChecked(qApp->property("la_classic_style").toBool());
    classic->setEnabled(classicStyleAvailable()); layout->addRow(classic);
    auto* family=new QFontComboBox(&dialog); family->setObjectName("options-font-family");
    family->setCurrentFont(qApp->font()); layout->addRow("&Schriftart:",family);
    auto* size=new QSpinBox(&dialog); size->setObjectName("options-font-size");
    size->setRange(6,72); size->setValue(qApp->font().pointSize()>0 ? qApp->font().pointSize() : 8);
    layout->addRow("Schrift&größe:",size);
    auto* buttons=new QDialogButtonBox(QDialogButtonBox::Ok|QDialogButtonBox::Cancel|QDialogButtonBox::Apply,&dialog);
    layout->addRow(buttons);
    const auto apply=[&] {
        if(!setClassicStyle(classic->isChecked())) {
            QMessageBox::warning(&dialog,"Optionen","Der gewählte Oberflächenstil ist nicht verfügbar."); return false;
        }
        auto font=family->currentFont(); font.setPointSize(size->value()); qApp->setFont(font);
        QSettings settings; settings.setValue("ui/classicStyle",classic->isChecked());
        settings.setValue("ui/applicationFont",font); settings.sync();
        const QSignalBlocker block(classicAction); classicAction->setChecked(classic->isChecked());
        if(settings.status()!=QSettings::NoError)
            QMessageBox::warning(&dialog,"Optionen","Die Einstellungen gelten für diese Sitzung, konnten aber nicht gespeichert werden.");
        return true;
    };
    QObject::connect(buttons->button(QDialogButtonBox::Apply),&QPushButton::clicked,&dialog,[&] { (void)apply(); });
    QObject::connect(buttons,&QDialogButtonBox::accepted,&dialog,[&] { if(apply()) dialog.accept(); });
    QObject::connect(buttons,&QDialogButtonBox::rejected,&dialog,&QDialog::reject);
    dialog.exec();
}
}
