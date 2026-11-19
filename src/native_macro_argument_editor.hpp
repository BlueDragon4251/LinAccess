#pragma once
#include "native_macro_editor_draft.hpp"
#include <QCheckBox>
#include <QDialog>
#include <QDialogButtonBox>
#include <QFormLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QLineEdit>
#include <QPlainTextEdit>
namespace linaccess {
inline std::optional<NativeMacroRowDraft> editNativeMacroArguments(QWidget* parent,NativeMacroRowDraft draft) {
    const NativeMacroSignature* signature=nullptr;for(const auto& candidate:nativeMacroSignatures()) if(candidate.name==draft.action) signature=&candidate;if(!signature) return {};
    QDialog editor(parent);editor.setObjectName("native-macro-argument-dialog");editor.setWindowTitle(draft.action+" – Argumente");editor.resize(720,500);auto* fields=new QFormLayout(&editor);
    auto* help=new QLabel("Nicht übergebene Argumente verwenden den Standardwert. Mit Übergeben kann auch ein leerer Text ausdrücklich übergeben werden.",&editor);help->setWordWrap(true);fields->addRow(help);
    QStringList values;for(int index=0;index<3;++index) values.push_back(draft.cells[index]);values.append(draft.extraArguments);
    int count=signature->types.size();for(int index=count;index<qMax(values.size(),draft.presence.size());++index) if(!values.value(index).isEmpty() || draft.presence.value(index,-1)==1) count=index+1;
    QVector<QWidget*> inputs;QVector<QCheckBox*> supplied;
    const QStringList messageLabels{"Meldung","Signalton (Ja/Nein)","Typ (None/Critical/Warning?/Warning!/Information)","Titel"};
    const QStringList formLabels{"Formularname","Ansicht","Filtername","WhereCondition","Datenmodus","Fenstermodus","OpenArgs"};
    const QStringList setValueLabels{"Ziel (Steuerelement/Eigenschaft)","Ausdruck (ohne führendes =)"};
    const QStringList reportLabels{"Berichtsname","Ansicht","Filtername","WhereCondition","Fenstermodus","OpenArgs"};
    const QStringList moveSizeLabels{"Abstand links in Twips","Abstand oben in Twips","Breite in Twips","Höhe in Twips"};
    for(int index=0;index<count;++index) {
        auto* row=new QWidget(&editor);auto* layout=new QHBoxLayout(row);layout->setContentsMargins(0,0,0,0);QWidget* input=nullptr;const auto type=signature->types.value(index);
        if(type!="long" && type!="boolean") {auto* text=new QPlainTextEdit(values.value(index),row);text->setMaximumHeight(85);input=text;}else input=new QLineEdit(values.value(index),row);
        input->setObjectName("native-macro-argument-"+QString::number(index+1));inputs.push_back(input);layout->addWidget(input,1);
        auto* present=new QCheckBox("Übergeben",row);present->setObjectName("native-macro-argument-supplied-"+QString::number(index+1));const bool required=index<signature->required;present->setChecked(required || draft.presence.value(index,-1)==1 || (draft.presence.value(index,-1)==-1 && !values.value(index).isEmpty()));present->setEnabled(!required);input->setEnabled(present->isChecked());supplied.push_back(present);layout->addWidget(present);
        QObject::connect(present,&QCheckBox::toggled,input,&QWidget::setEnabled);
        if(auto* text=qobject_cast<QPlainTextEdit*>(input)) QObject::connect(text,&QPlainTextEdit::textChanged,present,[text,present] {if(!text->toPlainText().isEmpty()) present->setChecked(true);});else if(auto* line=qobject_cast<QLineEdit*>(input)) QObject::connect(line,&QLineEdit::textChanged,present,[present](const QString& text) {if(!text.isEmpty()) present->setChecked(true);});
        auto label=draft.action=="SetValue" ? setValueLabels.value(index) : draft.action=="MessageBox" ? messageLabels.value(index) : draft.action=="OpenForm" ? formLabels.value(index) : draft.action=="OpenReport" ? reportLabels.value(index) : draft.action=="MoveSize" ? moveSizeLabels.value(index) : QString("Argument %1").arg(index+1);label+=QString(" (%1%2)").arg(type.isEmpty() ? QString("nicht zulässig") : type,required ? ", erforderlich" : ", optional");fields->addRow(label,row);
    }
    if(!count) fields->addRow(new QLabel("Diese Aktion hat keine Argumente.",&editor));
    auto* buttons=new QDialogButtonBox(QDialogButtonBox::Ok|QDialogButtonBox::Cancel,&editor);fields->addRow(buttons);QObject::connect(buttons,&QDialogButtonBox::accepted,&editor,&QDialog::accept);QObject::connect(buttons,&QDialogButtonBox::rejected,&editor,&QDialog::reject);
    if(editor.exec()!=QDialog::Accepted) return {};QStringList edited;draft.presence.clear();
    for(int index=0;index<inputs.size();++index) {const bool present=supplied[index]->isChecked();draft.presence.push_back(present ? 1 : 0);auto* input=inputs[index];edited.push_back(!present ? QString{} : qobject_cast<QPlainTextEdit*>(input) ? qobject_cast<QPlainTextEdit*>(input)->toPlainText() : qobject_cast<QLineEdit*>(input)->text());}
    for(int index=0;index<3;++index) draft.cells[index]=edited.value(index);draft.extraArguments=edited.mid(3);return draft;
}
}
