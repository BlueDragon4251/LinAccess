#include "native_input_box.hpp"
#include <QApplication>
#include <QAction>
#include <QClipboard>
#include <QContextMenuEvent>
#include <QDialog>
#include <QDropEvent>
#include <QGridLayout>
#include <QInputMethodEvent>
#include <QKeyEvent>
#include <QLabel>
#include <QLineEdit>
#include <QMenu>
#include <QMimeData>
#include <QPushButton>
#include <QScreen>

namespace linaccess {
namespace {
QString firstLine(QString text) {
    const auto cr=text.indexOf('\r'),lf=text.indexOf('\n');
    auto end=cr<0 ? lf : lf<0 ? cr : qMin(cr,lf);
    return end<0 ? text : text.left(end);
}
class InputEditor final : public QLineEdit {
public:
    explicit InputEditor(QWidget* owner) : QLineEdit(owner) { setMaxLength(255); }
protected:
    void keyPressEvent(QKeyEvent* event) override {
        if (event->matches(QKeySequence::Paste)) { pasteLine(); event->accept(); return; }
        QLineEdit::keyPressEvent(event);
    }
    void contextMenuEvent(QContextMenuEvent* event) override {
        QMenu menu(this);
        auto* undoAction=menu.addAction("Rückgängig",this,&QLineEdit::undo); undoAction->setEnabled(isUndoAvailable());
        auto* redoAction=menu.addAction("Wiederholen",this,&QLineEdit::redo); redoAction->setEnabled(isRedoAvailable());
        menu.addSeparator();
        auto* cutAction=menu.addAction("Ausschneiden",this,&QLineEdit::cut); cutAction->setEnabled(hasSelectedText());
        auto* copyAction=menu.addAction("Kopieren",this,&QLineEdit::copy); copyAction->setEnabled(hasSelectedText());
        menu.addAction("Einfügen",this,[this] { pasteLine(); });
        menu.addSeparator(); menu.addAction("Alles markieren",this,&QLineEdit::selectAll);
        menu.exec(event->globalPos());
    }
    void dropEvent(QDropEvent* event) override {
        if (event->source()!=this && event->mimeData()->hasText()) {
            setCursorPosition(cursorPositionAt(event->position().toPoint()));
            insert(firstLine(event->mimeData()->text())); event->acceptProposedAction(); return;
        }
        QLineEdit::dropEvent(event);
    }
    void inputMethodEvent(QInputMethodEvent* event) override {
        QInputMethodEvent normalized(event->preeditString(),event->attributes());
        normalized.setCommitString(firstLine(event->commitString()),event->replacementStart(),event->replacementLength());
        QLineEdit::inputMethodEvent(&normalized); event->setAccepted(normalized.isAccepted());
    }
private:
    void pasteLine() { insert(firstLine(QApplication::clipboard()->text())); }
};
}
std::string nativeInputBox(QWidget* owner,const VbaInputBoxArguments& arguments) {
    QDialog dialog(owner); dialog.setObjectName("access-input-box");
    dialog.setWindowTitle(arguments.title ? QString::fromStdString(*arguments.title) : QStringLiteral("LinAccess"));
    dialog.setWindowModality(Qt::ApplicationModal);
    dialog.setWindowFlag(Qt::WindowContextHelpButtonHint,false);
    auto* layout=new QGridLayout(&dialog);
    auto* prompt=new QLabel(QString::fromStdString(arguments.prompt),&dialog);
    prompt->setTextFormat(Qt::PlainText); prompt->setWordWrap(true);
    prompt->setAlignment(Qt::AlignLeft|Qt::AlignTop); prompt->setMinimumWidth(260);
    auto* editor=new InputEditor(&dialog); editor->setObjectName("access-input-box-editor");
    editor->setText(firstLine(QString::fromStdString(arguments.defaultValue))); editor->selectAll();
    auto* ok=new QPushButton("OK",&dialog); ok->setObjectName("access-input-box-ok"); ok->setDefault(true);
    auto* cancel=new QPushButton("Abbrechen",&dialog); cancel->setObjectName("access-input-box-cancel");
    layout->addWidget(prompt,0,0,2,1); layout->addWidget(ok,0,1); layout->addWidget(cancel,1,1);
    layout->addWidget(editor,2,0,1,2); layout->setRowStretch(0,1);
    dialog.setMinimumSize(390,140); dialog.adjustSize();
    QObject::connect(ok,&QPushButton::clicked,&dialog,&QDialog::accept);
    QObject::connect(cancel,&QPushButton::clicked,&dialog,&QDialog::reject);
    QWidget::setTabOrder(editor,ok); QWidget::setTabOrder(ok,cancel); editor->setFocus();
    if (auto* screen=owner ? owner->screen() : QApplication::primaryScreen()) {
        const auto available=screen->availableGeometry(),full=screen->geometry();
        const int x=arguments.xTwips ? full.x()+qRound(*arguments.xTwips*screen->logicalDotsPerInchX()/1440.0)
            : available.x()+(available.width()-dialog.width())/2;
        const int y=arguments.yTwips ? full.y()+qRound(*arguments.yTwips*screen->logicalDotsPerInchY()/1440.0)
            : available.y()+(available.height()-dialog.height())/3;
        dialog.move(x,y);
    }
    return dialog.exec()==QDialog::Accepted ? editor->text().left(254).toStdString() : std::string{};
}
}
