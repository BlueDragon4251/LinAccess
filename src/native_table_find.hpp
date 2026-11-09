#pragma once
#include "native_find_value.hpp"
#include <QCheckBox>
#include <QAbstractItemView>
#include <QStyledItemDelegate>
#include <QElapsedTimer>
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
    bool formatted{},wildcards{};
};
inline QString nativeVisibleFindText(const QAbstractItemView& view,const QModelIndex& index) {
    const auto* delegate=qobject_cast<const QStyledItemDelegate*>(view.itemDelegateForIndex(index));
    if(!delegate) throw VbaRuntimeError(3251,"Formatted search requires the actual native text delegate");
    const auto text=delegate->displayText(index.data(Qt::DisplayRole),view.locale());
    if(text.contains(QChar(0)) || !text.isValidUtf16()) throw VbaRuntimeError(3251,"Formatted search text is not representable");
    return text;
}
inline QDialog* nativeTableFindDialog(QWidget* owner,const QVector<QPair<QString,int>>& columns,const QMap<int,QString>& typedHints,
                                     NativeTableFindArguments initial,std::function<QString(NativeTableFindArguments)> search,
                                     std::function<QString(NativeTableFindArguments,QString)> replace={},
                                     std::function<QString(NativeTableFindArguments,QString)> replaceAll={}) {
    auto* dialog=new QDialog(owner);
    dialog->setObjectName(replace ? "native-table-replace-dialog" : "native-table-find-dialog"); dialog->setWindowTitle(replace ? "Suchen und Ersetzen" : "Suchen"); dialog->setAttribute(Qt::WA_DeleteOnClose);
    auto* box=new QVBoxLayout(dialog); auto* fields=new QFormLayout;
    auto* text=new QLineEdit(dialog); text->setObjectName("native-table-find-text"); text->setMaxLength(std::numeric_limits<int>::max()); text->setText(initial.text);
    auto* scope=new QComboBox(dialog); scope->setObjectName("native-table-find-field");
    if(!replace && (initial.formatted || typedHints.size()<columns.size())) scope->addItem(initial.formatted ? "Alle sichtbaren Felder" : "Alle sichtbaren Textfelder",-1);
    for(const auto& column:columns) scope->addItem(column.first,column.second);
    const int selected=scope->findData(initial.column); if(selected>=0) scope->setCurrentIndex(selected);
    auto* match=new QComboBox(dialog); match->setObjectName("native-table-find-match"); match->addItems({"Teil des Feldinhalts","Ganzes Feld","Feldanfang"}); match->setCurrentIndex(initial.mode);
    auto* sensitive=new QCheckBox("Groß-/Kleinschreibung beachten",dialog); sensitive->setObjectName("native-table-find-case"); sensitive->setChecked(initial.sensitive);
    auto* direction=new QComboBox(dialog); direction->setObjectName("native-table-find-direction");
    direction->addItem("Alle Datensätze",2); direction->addItem("Nach unten",1); direction->addItem("Nach oben",0); direction->setCurrentIndex(direction->findData(initial.direction));
    fields->addRow("Suchen nach:",text); fields->addRow("Suchen in:",scope); fields->addRow("Vergleichen:",match);
    fields->addRow("Suchrichtung:",direction);
    QLineEdit* replacement=nullptr;
    if(replace) {
        replacement=new QLineEdit(dialog); replacement->setObjectName("native-table-replace-with"); replacement->setMaxLength(1000000);
        fields->addRow("Ersetzen durch:",replacement);
    }
    box->addLayout(fields); box->addWidget(sensitive);
    auto* formatted=new QCheckBox("Formatierte Anzeige durchsuchen",dialog);formatted->setObjectName("native-table-find-formatted");formatted->setChecked(!replace && initial.formatted);formatted->setVisible(!replace);box->addWidget(formatted);
    auto* wildcards=new QCheckBox("Platzhalter verwenden (* ? # [ ])",dialog);wildcards->setObjectName("native-table-find-wildcards");wildcards->setChecked(!replace && initial.wildcards);wildcards->setVisible(!replace);box->addWidget(wildcards);
    auto* hint=new QLabel(dialog); hint->setObjectName("native-table-find-hint"); hint->setWordWrap(true); box->addWidget(hint);
    const auto updateField=[scope,match,sensitive,hint,typedHints,formatted,wildcards,replacing=bool(replace)] {
        const auto column=scope->currentData().toInt(); const bool exact=!formatted->isChecked() && !wildcards->isChecked() && typedHints.contains(column);
        if(exact) match->setCurrentIndex(1);
        match->setEnabled(!exact); sensitive->setEnabled(!exact);
        hint->setText(wildcards->isChecked() ? QStringLiteral("Platzhalter: * beliebig, ? ein Zeichen, # Ziffer, [abc] Zeichenklasse.") : formatted->isChecked() ? QStringLiteral("Die tatsächlich formatierte Zellanzeige wird wörtlich durchsucht.") : exact ? typedHints.value(column) : replacing ? QStringLiteral("Text wird wörtlich ersetzt: aktuell erstes Vorkommen, Alle ersetzt jedes Vorkommen.") : QStringLiteral("Text wird wörtlich gesucht."));
    };
    QObject::connect(wildcards,&QCheckBox::toggled,dialog,[updateField] {updateField();});
    QObject::connect(scope,&QComboBox::currentIndexChanged,dialog,[updateField](int) { updateField(); }); updateField();
    QObject::connect(formatted,&QCheckBox::toggled,dialog,[scope,updateField,hasText=typedHints.size()<columns.size()](bool checked) {int all=scope->findData(-1);if(checked && all<0) {scope->insertItem(0,"Alle sichtbaren Felder",-1);all=0;}if(!checked && !hasText && all>=0) scope->removeItem(all);else if(all>=0) scope->setItemText(all,checked ? "Alle sichtbaren Felder" : "Alle sichtbaren Textfelder");updateField();});
    auto* result=new QLabel(dialog); result->setObjectName("native-table-find-result"); result->setWordWrap(true); box->addWidget(result);
    auto* buttons=new QHBoxLayout; auto* next=new QPushButton("&Weitersuchen",dialog); next->setObjectName("native-table-find-next"); next->setDefault(true);
    auto* close=new QPushButton("&Schließen",dialog); close->setObjectName("native-table-find-close"); buttons->addWidget(next); buttons->addWidget(close); box->addLayout(buttons);
    next->setEnabled(!initial.text.isEmpty());
    QObject::connect(text,&QLineEdit::textChanged,next,[next](const QString& value) { next->setEnabled(!value.isEmpty()); });
    QObject::connect(close,&QPushButton::clicked,dialog,&QDialog::reject);
    const QPointer<QDialog> guarded=dialog;
    QObject::connect(next,&QPushButton::clicked,dialog,[guarded,text,scope,match,sensitive,direction,formatted,wildcards,result,search] {
        try {
            NativeTableFindArguments arguments{text->text(),scope->currentData().toInt(),match->currentIndex(),sensitive->isChecked(),0,direction->currentData().toInt()};arguments.formatted=formatted->isChecked();arguments.wildcards=wildcards->isChecked();
            const auto message=search(std::move(arguments));
            if(guarded) result->setText(message);
        } catch(const std::exception& error) { if(guarded) result->setText(QString::fromUtf8(error.what())); }
    });
    if(replace) {
        auto* apply=new QPushButton("&Ersetzen",dialog); apply->setObjectName("native-table-replace-current"); buttons->insertWidget(1,apply);
        apply->setEnabled(!text->text().isEmpty());
        QObject::connect(text,&QLineEdit::textChanged,apply,[apply](const QString& value) { apply->setEnabled(!value.isEmpty()); });
        QObject::connect(apply,&QPushButton::clicked,dialog,[guarded,text,scope,match,sensitive,direction,replacement,result,replace] {
            try {
                const auto message=replace({text->text(),scope->currentData().toInt(),match->currentIndex(),sensitive->isChecked(),0,direction->currentData().toInt()},replacement->text());
                if(guarded) result->setText(message);
            } catch(const std::exception& error) { if(guarded) result->setText(QString::fromUtf8(error.what())); }
        });
    }
    if(replaceAll && replacement) {
        auto* all=new QPushButton("&Alle ersetzen",dialog); all->setObjectName("native-table-replace-all"); buttons->insertWidget(2,all);
        const auto enabled=[all,text,direction] { all->setEnabled(!text->text().isEmpty() && direction->currentData().toInt()==2); };
        QObject::connect(text,&QLineEdit::textChanged,dialog,[enabled] { enabled(); });
        QObject::connect(direction,&QComboBox::currentIndexChanged,dialog,[enabled] { enabled(); }); enabled();
        QObject::connect(all,&QPushButton::clicked,dialog,[guarded,text,scope,match,sensitive,replacement,result,replaceAll] {
            try {
                const auto message=replaceAll({text->text(),scope->currentData().toInt(),match->currentIndex(),sensitive->isChecked(),0,2},replacement->text());
                if(guarded) result->setText(message);
            } catch(const std::exception& error) { if(guarded) result->setText(QString::fromUtf8(error.what())); }
        });
    }
    const QPointer<QLineEdit> guardedText=text;
    dialog->show(); if(!guarded) return nullptr;
    dialog->activateWindow(); if(!guarded) return nullptr;
    text->setFocus(); if(guardedText) guardedText->selectAll();
    return guarded.data();
}
}
