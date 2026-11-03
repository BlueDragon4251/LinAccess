#pragma once
#include <QStyledItemDelegate>
#include "native_clock_keys.hpp"
#include "native_cell_editor.hpp"
#include "native_cell_zoom.hpp"
#include "access_table_model.hpp"
#include <QTableView>
#include <QPointer>
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
        if(event->type()==QEvent::KeyPress) {
            const auto* key=static_cast<QKeyEvent*>(event);
            const auto modifiers=key->modifiers() & ~Qt::KeypadModifier;
            const bool advance=(modifiers==Qt::NoModifier && (key->key()==Qt::Key_Return || key->key()==Qt::Key_Enter))
                || ((key->key()==Qt::Key_Tab || key->key()==Qt::Key_Backtab) && (modifiers==Qt::NoModifier || modifiers==Qt::ShiftModifier));
            if(advance) {
                QPointer<QWidget> editor=qobject_cast<QWidget*>(watched);
                const auto target=watched->property("la_cell_index").value<QPersistentModelIndex>();
                if(editor && target.isValid() && qobject_cast<const AccessTableModel*>(target.model())) {
                    if(!commitNativeCellEditor(editor,target) || !editor) return true;
                }
            }
            if(key->key()==Qt::Key_F2 && key->modifiers()==Qt::ShiftModifier) {
                QPointer<QLineEdit> editor=qobject_cast<QLineEdit*>(watched);
                const auto target=watched->property("la_cell_index").value<QPersistentModelIndex>();
                if(editor && target.isValid()) {
                    const auto result=nativeCellZoom(editor,editor->text(),editor->isReadOnly() || !target.flags().testFlag(Qt::ItemIsEditable));
                    if(editor && target.isValid() && result && (result->modified || result->text!=editor->text())) {
                        editor->setText(result->text); editor->setModified(true);
                    }
                    return true;
                }
            }
            if(key->key()==Qt::Key_F2 && key->modifiers()==Qt::NoModifier) {
                QPointer<QWidget> editor=qobject_cast<QWidget*>(watched);
                const auto target=watched->property("la_cell_index").value<QPersistentModelIndex>();
                if(editor && target.isValid() && qobject_cast<const AccessTableModel*>(target.model())) {
                    if(commitNativeCellEditor(editor,target) && editor) emit closeEditor(editor,QAbstractItemDelegate::NoHint);
                    return true;
                }
            }
            if(key->key()==Qt::Key_Apostrophe && (key->modifiers() & ~Qt::KeypadModifier)==Qt::ControlModifier) {
                QPointer<QWidget> editor=qobject_cast<QLineEdit*>(watched);
                const auto target=watched->property("la_cell_index").value<QPersistentModelIndex>();
                auto* typed=target.isValid() ? qobject_cast<AccessTableModel*>(const_cast<QAbstractItemModel*>(target.model())) : nullptr;
                auto* line=qobject_cast<QLineEdit*>(editor.data());
                QTableView* view=nullptr;
                for(auto* parent=editor ? editor->parentWidget() : nullptr;parent;parent=parent->parentWidget())
                    if((view=qobject_cast<QTableView*>(parent))) break;
                if(line && line->isEnabled() && !line->isReadOnly() && typed && view && target.row()>0
                    && !view->isRowHidden(target.row()-1) && !view->isColumnHidden(target.column())
                    && typed->repeatPreviousValue(target) && editor) setEditorData(editor,target);
                return true;
            }
        }
        if(event->type()==QEvent::KeyPress) if(const auto text=nativeClockKeyText(*static_cast<QKeyEvent*>(event))) {
            auto* editor=qobject_cast<QLineEdit*>(watched);
            const auto index=watched->property("la_cell_index").value<QPersistentModelIndex>();
            if(editor && editor->isEnabled() && !editor->isReadOnly() && index.isValid() && index.flags().testFlag(Qt::ItemIsEditable)
                && editor->text().size()-editor->selectedText().size()+text->size()<=editor->maxLength()) editor->insert(*text);
            return true;
        }
        return QStyledItemDelegate::eventFilter(watched,event);
    }
    void setModelData(QWidget* editor,QAbstractItemModel* model,const QModelIndex& index) const override {
        if(nativeCellEditorChanged(editor,index)) QStyledItemDelegate::setModelData(editor,model,index);
    }
    QWidget* createEditor(QWidget* parent,const QStyleOptionViewItem& option,const QModelIndex& index) const override {
        auto* editor=QStyledItemDelegate::createEditor(parent,option,index);
        if(editor) editor->setProperty("la_cell_index",QVariant::fromValue(QPersistentModelIndex(index)));
        if(auto* text=qobject_cast<QLineEdit*>(editor)) text->setMaxLength(std::numeric_limits<int>::max());
        return editor;
    }
};
}
