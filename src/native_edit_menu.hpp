#pragma once
#include <QAction>
#include "access_table_view.hpp"
#include <QApplication>
#include <QClipboard>
#include <QKeyEvent>
#include <QLineEdit>
#include <QMenuBar>
#include <QMenu>
#include <QMimeData>
#include <QPlainTextEdit>
#include <QPointer>
#include <QTextEdit>
#include <QTextDocument>
#include <memory>
namespace linaccess {
inline QMenu* addNativeEditMenu(QMenuBar* bar,QWidget* owner) {
    auto* menu=bar->addMenu("&Bearbeiten"); menu->setObjectName("native-edit-menu");
    auto target=std::make_shared<QPointer<QWidget>>();
    struct Command { const char* id; const char* label; Qt::Key key; Qt::KeyboardModifiers modifiers; };
    const Command commands[]{
        {"undo","&Rückgängig\tStrg+Z",Qt::Key_Z,Qt::ControlModifier},
        {"redo","&Wiederholen\tStrg+Y",Qt::Key_Y,Qt::ControlModifier},
        {"cut","&Ausschneiden\tStrg+X",Qt::Key_X,Qt::ControlModifier},
        {"copy","&Kopieren\tStrg+C",Qt::Key_C,Qt::ControlModifier},
        {"paste","&Einfügen\tStrg+V",Qt::Key_V,Qt::ControlModifier},
        {"delete","&Löschen\tEntf",Qt::Key_Delete,Qt::NoModifier},
        {"select-all","&Alles markieren\tStrg+A",Qt::Key_A,Qt::ControlModifier}
    };
    for(const auto& command:commands) {
        if(QByteArray(command.id)=="cut" || QByteArray(command.id)=="select-all") menu->addSeparator();
        auto* action=menu->addAction(QString::fromUtf8(command.label));
        action->setObjectName("edit-"+QString::fromLatin1(command.id)); action->setEnabled(false);
        // Display shortcuts without capturing the controls' actual key routing.
        QObject::connect(action,&QAction::triggered,menu,[owner,target,command] {
            QPointer<QWidget> receiver=*target;
            if(!receiver || !owner->isAncestorOf(receiver) || !receiver->isEnabled()) return;
            QKeyEvent press(QEvent::KeyPress,command.key,command.modifiers);
            QApplication::sendEvent(receiver,&press);
            if(receiver) { QKeyEvent release(QEvent::KeyRelease,command.key,command.modifiers); QApplication::sendEvent(receiver,&release); }
        });
    }
    QObject::connect(menu,&QMenu::aboutToShow,menu,[owner,menu,target] {
        auto* focus=QApplication::focusWidget();
        if(!focus || !menu->isAncestorOf(focus)) *target=focus && owner->isAncestorOf(focus) ? focus : nullptr;
        auto* widget=target->data();
        bool readable=false,writable=false,selection=false,undo=false,redo=false,content=false,forward=false,tablePaste=false;
        if(auto* line=qobject_cast<QLineEdit*>(widget)) {
            readable=line->echoMode()==QLineEdit::Normal; writable=!line->isReadOnly();
            selection=line->hasSelectedText(); undo=line->isUndoAvailable(); redo=line->isRedoAvailable();
            content=!line->text().isEmpty(); forward=line->cursorPosition()<line->text().size();
        } else if(auto* plain=qobject_cast<QPlainTextEdit*>(widget)) {
            readable=true; writable=!plain->isReadOnly(); selection=plain->textCursor().hasSelection();
            undo=plain->document()->isUndoAvailable(); redo=plain->document()->isRedoAvailable();
            content=!plain->document()->isEmpty(); forward=!plain->textCursor().atEnd();
        } else if(auto* rich=qobject_cast<QTextEdit*>(widget)) {
            readable=true; writable=!rich->isReadOnly(); selection=rich->textCursor().hasSelection();
            undo=rich->document()->isUndoAvailable(); redo=rich->document()->isRedoAvailable();
            content=!rich->document()->isEmpty(); forward=!rich->textCursor().atEnd();
        }
        if(auto* table=dynamic_cast<AccessTableView*>(widget);table && table->model()) {
            readable=true; selection=table->selectionModel() && table->selectionModel()->hasSelection();
            content=table->model()->rowCount()>0 && table->model()->columnCount()>0;
            tablePaste=table->canPasteSelection();
        }
        const auto* clipboard=QApplication::clipboard()->mimeData();
        for(auto* action:menu->actions()) {
            const auto id=action->objectName();
            const bool enabled=id=="edit-undo" ? writable && undo : id=="edit-redo" ? writable && redo
                : id=="edit-copy" ? readable && selection : id=="edit-cut" ? readable && writable && selection
                : id=="edit-paste" ? (writable || tablePaste) && clipboard && clipboard->hasText()
                : id=="edit-delete" ? writable && (selection || forward) : id=="edit-select-all" && content;
            action->setEnabled(widget && widget->isEnabled() && enabled);
        }
    });
    return menu;
}
}
