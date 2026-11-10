#pragma once
#include <QDialog>
#include <QVariant>
#include <QCloseEvent>
#include <QMessageBox>
#include <QPushButton>
#include <QPointer>
#include <QScopeGuard>
class NativeModuleEditorDialog final : public QDialog {
public:
    explicit NativeModuleEditorDialog(QWidget* owner):QDialog(owner) {}
    void reject() override {
        if(property("la_module_modeless").toBool() && !closing_) {close();return;}
        QDialog::reject();
    }
protected:
    void closeEvent(QCloseEvent* event) override {
        if(closing_ || property("la_module_busy").toInt()>0) {event->ignore();return;}
        closing_=true;const auto reset=qScopeGuard([this] {closing_=false;});
        if(property("la_module_modeless").toBool() && isWindowModified() && !property("la_module_discard_close").toBool()) {
            const auto answer=QMessageBox::question(this,"Modul schließen","Änderungen vor dem Schließen speichern?",QMessageBox::Save|QMessageBox::Discard|QMessageBox::Cancel,QMessageBox::Save);
            if(answer!=QMessageBox::Save && answer!=QMessageBox::Discard) {event->ignore();return;}
            if(answer==QMessageBox::Save) {
                auto* save=findChild<QPushButton*>("native-module-save");
                if(!save || !save->isEnabled()) {event->ignore();return;}
                save->click();if(isWindowModified()) {event->ignore();return;}
            }
        }
        QDialog::closeEvent(event);
    }
private:
    bool closing_{};
};
class NativeModuleEditorBusy {
public:
    explicit NativeModuleEditorBusy(QDialog* dialog):dialog_(dialog),depth_(dialog->property("la_module_busy").toInt()) {dialog->setProperty("la_module_busy",depth_+1);}
    ~NativeModuleEditorBusy() {if(dialog_) dialog_->setProperty("la_module_busy",depth_);}
private:
    QPointer<QDialog> dialog_;int depth_{};
};
