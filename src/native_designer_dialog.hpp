#pragma once
#include <QApplication>
#include <QDialog>
#include <QVariant>
#include <QCloseEvent>
#include <QMessageBox>
#include <QPushButton>
#include <QPointer>
#include <QScopeGuard>
#include <QTableWidget>
class NativeDesignerDialog final : public QDialog {
public:
    explicit NativeDesignerDialog(QWidget* owner):QDialog(owner) {}

    void setSaveButton(QPushButton* button) { save_=button; }

    void saveKeepingOpen() {
        if(!save_ || !save_->isEnabled()) return;
        const QPointer<NativeDesignerDialog> dialog=this;
        const auto previous=property("la_editor_keep_open");
        setProperty("la_editor_keep_open",true);
        const auto reset=qScopeGuard([dialog,previous] {
            if(dialog) dialog->setProperty("la_editor_keep_open",previous);
        });
        save_->click();
    }

    bool prepareProjectClose() {
        if(!confirmClose()) return false;
        prepared_=true;
        wasDisabled_=testAttribute(Qt::WA_ForceDisabled);
        setEnabled(false);
        return true;
    }

    void finishProjectClose(bool proceed) {
        const QPointer<NativeDesignerDialog> dialog=this;
        if(proceed) QDialog::reject();
        if(!dialog) return;
        prepared_=false;
        setEnabled(!wasDisabled_);
    }

    void reject() override { close(); }

protected:
    void closeEvent(QCloseEvent* event) override {
        if(!confirmClose()) {
            event->ignore();
            return;
        }
        if(isVisible()) QDialog::reject();
    }

private:
    bool confirmClose() {
        if(prompting_ || prepared_ || property("la_module_busy").toInt()>0) return false;
        for(auto* modal=QApplication::activeModalWidget();modal && modal!=this;modal=modal->parentWidget())
            if(modal->parentWidget()==this) return false;
        auto* focus=QApplication::focusWidget();
        if(focus && save_ && save_->isEnabled())
            for(auto* table:findChildren<QTableWidget*>())
                if(table->viewport()->isAncestorOf(focus)) {
                    // The delegate commits its live cell editor on focus loss.
                    save_->setFocus(Qt::OtherFocusReason);
                    break;
                }
        if(!isWindowModified() || property("la_module_discard_close").toBool()) return true;
        const QPointer<NativeDesignerDialog> dialog=this;
        prompting_=true;
        const auto reset=qScopeGuard([dialog] { if(dialog) dialog->prompting_=false; });
        const auto name=property("la_project_object_name").toString();
        QMessageBox prompt(QMessageBox::Question,"Entwurf schließen",
                           name.isEmpty() ? QStringLiteral("Änderungen vor dem Schließen speichern?")
                                          : QString("Änderungen an „%1“ vor dem Schließen speichern?").arg(name),
                           QMessageBox::Save|QMessageBox::Discard|QMessageBox::Cancel,this);
        prompt.setObjectName("native-designer-close");
        prompt.setTextFormat(Qt::PlainText);
        prompt.setDefaultButton(QMessageBox::Save);
        prompt.exec();
        if(!dialog) return false;
        const auto answer=prompt.standardButton(prompt.clickedButton());
        if(answer==QMessageBox::Discard) return true;
        if(answer!=QMessageBox::Save) return false;
        saveKeepingOpen();
        return dialog && !dialog->isWindowModified();
    }

    QPointer<QPushButton> save_;
    bool prompting_{};
    bool prepared_{};
    bool wasDisabled_{};
};
class NativeModuleEditorBusy {
public:
    explicit NativeModuleEditorBusy(QDialog* dialog):dialog_(dialog),depth_(dialog->property("la_module_busy").toInt()) {dialog->setProperty("la_module_busy",depth_+1);}
    ~NativeModuleEditorBusy() {if(dialog_) dialog_->setProperty("la_module_busy",depth_);}
private:
    QPointer<QDialog> dialog_;int depth_{};
};
