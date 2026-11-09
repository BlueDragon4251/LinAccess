#pragma once
#include "native_form_events.hpp"
#include <QComboBox>
#include <QDialog>
#include <QDialogButtonBox>
#include <QFormLayout>
#include <QGroupBox>
#include <QLabel>
#include <QMessageBox>
#include <QSignalBlocker>
#include <QSpinBox>
#include <QVBoxLayout>
#include <functional>
#include <memory>
namespace linaccess {
inline bool editNativeFormEvents(QWidget* parent,QJsonObject& formEvents,QMap<QString,QJsonObject>& fieldEvents,int& timerInterval,const QStringList& fields,const QStringList& macroNames) {
    auto form=nativeEventObject(nativeEventOverrides(formEvents,nativeFormMacroEvents()));auto values=fieldEvents;
    QDialog dialog(parent);dialog.setObjectName("native-form-event-editor");dialog.setWindowTitle("Ereignisse zuordnen");dialog.resize(760,480);auto* layout=new QVBoxLayout(&dialog);
    auto* help=new QLabel("Automatisch verwendet vorhandene VBA-Prozeduren. Makro verbindet einen gespeicherten Makronamen. Keine Bindung deaktiviert dieses Ereignis ausdrücklich.",&dialog);help->setWordWrap(true);layout->addWidget(help);
    const auto row=[&](QGroupBox* group,const QString& prefix,const QStringList& events,std::function<QJsonObject()> read,std::function<void(QJsonObject)> write) {
        auto* formLayout=qobject_cast<QFormLayout*>(group->layout());
        auto* event=new QComboBox(group);event->setObjectName(prefix+"-event");event->addItems(events);formLayout->addRow("Ereignis",event);
        auto* mode=new QComboBox(group);mode->setObjectName(prefix+"-mode");mode->addItems({"Automatisch aus VBA","Makro","Funktion","VBA-Ereignisprozedur","Keine Bindung"});formLayout->addRow("Bindung",mode);
        auto* target=new QComboBox(group);target->setObjectName(prefix+"-target");target->setEditable(true);target->addItems(macroNames);formLayout->addRow("Makroname / Funktion",target);
        auto filling=std::make_shared<bool>(false);
        const auto enabled=[mode,target] {target->setEnabled(mode->currentIndex()==1 || mode->currentIndex()==2);};
        const auto refresh=[event,mode,target,read,filling,enabled] {
            *filling=true;const QSignalBlocker blockMode(mode),blockTarget(target);const auto saved=read().value(event->currentText());const auto binding=saved.toString();
            const int type=saved.isUndefined() ? 0 : binding.isEmpty() ? 4 : binding=="[Event Procedure]" ? 3 : binding.startsWith('=') ? 2 : 1;
            mode->setCurrentIndex(type);target->setEditText(type==1 || type==2 ? binding : QString{});enabled();*filling=false;
        };
        const auto store=[event,mode,target,read,write,filling,enabled] {
            if(*filling) return;auto object=read();const auto name=event->currentText();const int type=mode->currentIndex();
            if(type==0) object.remove(name);
            else if(type==3) object.insert(name,"[Event Procedure]");
            else if(type==4) object.insert(name,QStringLiteral(""));
            else {auto text=target->currentText().trimmed();if(text.isEmpty() || (type==1 && !nativeMacroCallNameValid(text))) text=QString(QChar(0));else if(type==2 && !text.startsWith('=')) text.prepend('=');object.insert(name,text);}
            write(object);enabled();
        };
        QObject::connect(mode,&QComboBox::currentIndexChanged,&dialog,store);QObject::connect(target,&QComboBox::currentTextChanged,&dialog,store);QObject::connect(event,&QComboBox::currentIndexChanged,&dialog,refresh);refresh();return refresh;
    };
    auto* formGroup=new QGroupBox("Formular",&dialog);formGroup->setLayout(new QFormLayout);layout->addWidget(formGroup);
    row(formGroup,"native-form-event",nativeFormMacroEvents(),[&] {return form;},[&](QJsonObject next) {form=std::move(next);});
    auto* fieldGroup=new QGroupBox("Formularfeld",&dialog);auto* fieldLayout=new QFormLayout(fieldGroup);layout->addWidget(fieldGroup);
    auto* field=new QComboBox(fieldGroup);field->setObjectName("native-field-event-field");field->addItems(fields);fieldLayout->addRow("Feld",field);fieldGroup->setEnabled(!fields.isEmpty());
    const auto refresh=row(fieldGroup,"native-field-event",nativeFieldMacroEvents(),[&] {return values.value(field->currentText().toCaseFolded());},[&](QJsonObject next) {values.insert(field->currentText().toCaseFolded(),std::move(next));});QObject::connect(field,&QComboBox::currentIndexChanged,&dialog,refresh);
    auto* interval=new QSpinBox(&dialog);interval->setObjectName("native-form-event-interval");interval->setRange(0,std::numeric_limits<qint32>::max());interval->setValue(timerInterval);interval->setPrefix("Timer: ");interval->setSuffix(" ms (0 = aus)");layout->addWidget(interval);
    auto* buttons=new QDialogButtonBox(QDialogButtonBox::Ok|QDialogButtonBox::Cancel,&dialog);layout->addWidget(buttons);QObject::connect(buttons,&QDialogButtonBox::rejected,&dialog,&QDialog::reject);
    QObject::connect(buttons,&QDialogButtonBox::accepted,&dialog,[&] {
        try {
            const auto normalized=nativeEventObject(nativeEventOverrides(form,nativeFormMacroEvents()));QJsonObject all;for(auto it=values.begin();it!=values.end();++it) if(fields.contains(it.key(),Qt::CaseInsensitive)) all.insert(it.key(),it.value());
            const auto checked=nativeFieldEventOverrides(all,fields);QMap<QString,QJsonObject> normalizedFields;for(auto it=checked.begin();it!=checked.end();++it) if(!it.value().isEmpty()) normalizedFields.insert(it.key(),nativeEventObject(it.value()));
            formEvents=normalized;fieldEvents=std::move(normalizedFields);timerInterval=interval->value();dialog.accept();
        } catch(const std::exception& error) {QMessageBox::critical(&dialog,"Ereignisbindung",QString::fromUtf8(error.what()));}
    });
    return dialog.exec()==QDialog::Accepted;
}
}
