#pragma once
#include "native_find_value.hpp"
#include <QCheckBox>
#include <QComboBox>
#include <QDialog>
#include <QFormLayout>
#include <QLabel>
#include <QLineEdit>
#include <QPointer>
#include <QPushButton>
#include <QVBoxLayout>
#include <QVector>
#include <exception>
#include <functional>
#include <limits>
namespace linaccess {
struct NativeTableFindArguments {
    QString text;
    int column{-1},mode{};
    bool sensitive{};
    int origin{}; // 0: continue dialog search, 1: first record, 2: next record.
    int direction{2}; // AcSearchDirection: Up=0, Down=1, All=2.
    std::optional<VbaValue> value;
};
inline QDialog* nativeTableFindDialog(QWidget* owner,const QVector<QPair<QString,int>>& columns,const QMap<int,QString>& typedHints,
                                     NativeTableFindArguments initial,std::function<QString(NativeTableFindArguments)> search) {
    auto* dialog=new QDialog(owner);
    dialog->setObjectName("native-table-find-dialog"); dialog->setWindowTitle("Suchen"); dialog->setAttribute(Qt::WA_DeleteOnClose);
    auto* box=new QVBoxLayout(dialog); auto* fields=new QFormLayout;
    auto* text=new QLineEdit(dialog); text->setObjectName("native-table-find-text"); text->setMaxLength(std::numeric_limits<int>::max()); text->setText(initial.text);
    auto* scope=new QComboBox(dialog); scope->setObjectName("native-table-find-field");
    if(typedHints.size()<columns.size()) scope->addItem("Alle sichtbaren Textfelder",-1);
    for(const auto& column:columns) scope->addItem(column.first,column.second);
    const int selected=scope->findData(initial.column); if(selected>=0) scope->setCurrentIndex(selected);
    auto* match=new QComboBox(dialog); match->setObjectName("native-table-find-match"); match->addItems({"Teil des Feldinhalts","Ganzes Feld","Feldanfang"}); match->setCurrentIndex(initial.mode);
    auto* sensitive=new QCheckBox("Groß-/Kleinschreibung beachten",dialog); sensitive->setObjectName("native-table-find-case"); sensitive->setChecked(initial.sensitive);
    auto* direction=new QComboBox(dialog); direction->setObjectName("native-table-find-direction");
    direction->addItem("Alle Datensätze",2); direction->addItem("Nach unten",1); direction->addItem("Nach oben",0); direction->setCurrentIndex(direction->findData(initial.direction));
    fields->addRow("Suchen nach:",text); fields->addRow("Suchen in:",scope); fields->addRow("Vergleichen:",match);
    fields->addRow("Suchrichtung:",direction);
    box->addLayout(fields); box->addWidget(sensitive);
    auto* hint=new QLabel(dialog); hint->setObjectName("native-table-find-hint"); hint->setWordWrap(true); box->addWidget(hint);
    const auto updateField=[scope,match,sensitive,hint,typedHints] {
        const auto column=scope->currentData().toInt(); const bool exact=typedHints.contains(column);
        if(exact) match->setCurrentIndex(1);
        match->setEnabled(!exact); sensitive->setEnabled(!exact);
        hint->setText(exact ? typedHints.value(column) : QStringLiteral("Text wird wörtlich gesucht."));
    };
    QObject::connect(scope,&QComboBox::currentIndexChanged,dialog,[updateField](int) { updateField(); }); updateField();
    auto* result=new QLabel(dialog); result->setObjectName("native-table-find-result"); result->setWordWrap(true); box->addWidget(result);
    auto* buttons=new QHBoxLayout; auto* next=new QPushButton("&Weitersuchen",dialog); next->setObjectName("native-table-find-next"); next->setDefault(true);
    auto* close=new QPushButton("&Schließen",dialog); close->setObjectName("native-table-find-close"); buttons->addWidget(next); buttons->addWidget(close); box->addLayout(buttons);
    next->setEnabled(!initial.text.isEmpty());
    QObject::connect(text,&QLineEdit::textChanged,next,[next](const QString& value) { next->setEnabled(!value.isEmpty()); });
    QObject::connect(close,&QPushButton::clicked,dialog,&QDialog::reject);
    const QPointer<QDialog> guarded=dialog;
    QObject::connect(next,&QPushButton::clicked,dialog,[guarded,text,scope,match,sensitive,direction,result,search] {
        try {
            const auto message=search({text->text(),scope->currentData().toInt(),match->currentIndex(),sensitive->isChecked(),0,direction->currentData().toInt()});
            if(guarded) result->setText(message);
        } catch(const std::exception& error) { if(guarded) result->setText(QString::fromUtf8(error.what())); }
    });
    const QPointer<QLineEdit> guardedText=text;
    dialog->show(); if(!guarded) return nullptr;
    dialog->activateWindow(); if(!guarded) return nullptr;
    text->setFocus(); if(guardedText) guardedText->selectAll();
    return guarded.data();
}
}
