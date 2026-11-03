#pragma once
#include <QStyledItemDelegate>
#include "native_clock_keys.hpp"
#include <QLineEdit>
#include <QPersistentModelIndex>
#include <QVariant>
#include <limits>
namespace linaccess {
// QString EditRole keeps Currency out of the default two-decimal spin editor.
class AccessTableDelegate final : public QStyledItemDelegate {
public:
    using QStyledItemDelegate::QStyledItemDelegate;
    bool eventFilter(QObject* watched,QEvent* event) override {
        if(event->type()==QEvent::KeyPress) if(const auto text=nativeClockKeyText(*static_cast<QKeyEvent*>(event))) {
            auto* editor=qobject_cast<QLineEdit*>(watched);
            const auto index=watched->property("la_cell_index").value<QPersistentModelIndex>();
            if(editor && editor->isEnabled() && !editor->isReadOnly() && index.isValid() && index.flags().testFlag(Qt::ItemIsEditable)
                && editor->text().size()-editor->selectedText().size()+text->size()<=editor->maxLength()) editor->insert(*text);
            return true;
        }
        return QStyledItemDelegate::eventFilter(watched,event);
    }
    QWidget* createEditor(QWidget* parent,const QStyleOptionViewItem& option,const QModelIndex& index) const override {
        auto* editor=QStyledItemDelegate::createEditor(parent,option,index);
        if(editor) editor->setProperty("la_cell_index",QVariant::fromValue(QPersistentModelIndex(index)));
        if(auto* text=qobject_cast<QLineEdit*>(editor)) text->setMaxLength(std::numeric_limits<int>::max());
        return editor;
    }
};
}
