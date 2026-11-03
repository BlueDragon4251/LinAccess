#pragma once
#include <QDialog>
#include <QDialogButtonBox>
#include <QPlainTextEdit>
#include <QPointer>
#include <QPushButton>
#include <QVBoxLayout>
#include <optional>
namespace linaccess {
struct NativeCellZoomResult { QString text; bool modified; };
inline std::optional<NativeCellZoomResult> nativeCellZoom(QWidget* owner,const QString& text,bool readOnly) {
    const QPointer<QWidget> focusOwner=owner;
    QPointer<QDialog> dialog=new QDialog(owner);
    dialog->setObjectName("native-cell-zoom"); dialog->setWindowTitle("Zoom"); dialog->resize(640,400);
    auto* layout=new QVBoxLayout(dialog);
    auto* editor=new QPlainTextEdit(dialog); editor->setObjectName("native-cell-zoom-text");
    editor->setPlainText(text); editor->setReadOnly(readOnly); editor->document()->setModified(false); layout->addWidget(editor);
    const auto normalizedOriginal=editor->toPlainText();
    auto* buttons=new QDialogButtonBox(readOnly ? QDialogButtonBox::Close : QDialogButtonBox::Ok|QDialogButtonBox::Cancel,dialog);
    if(auto* button=buttons->button(QDialogButtonBox::Cancel)) button->setText("Abbrechen");
    if(auto* button=buttons->button(QDialogButtonBox::Close)) button->setText("Schließen");
    layout->addWidget(buttons);
    QObject::connect(buttons,&QDialogButtonBox::accepted,dialog,&QDialog::accept);
    QObject::connect(buttons,&QDialogButtonBox::rejected,dialog,&QDialog::reject);
    editor->setFocus();
    const auto accepted=dialog->exec()==QDialog::Accepted;
    if(!dialog) return std::nullopt;
    std::optional<NativeCellZoomResult> result;
    if(accepted && !readOnly) {
        const auto current=editor->toPlainText();
        const bool modified=editor->document()->isModified() || current!=normalizedOriginal;
        result=NativeCellZoomResult{modified ? current : text,modified};
    }
    delete dialog.data();
    if(focusOwner && focusOwner->isVisible()) focusOwner->setFocus(Qt::OtherFocusReason);
    return result;
}
}
