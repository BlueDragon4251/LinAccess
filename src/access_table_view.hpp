#pragma once
#include "access_table_model.hpp"
#include "native_clock_keys.hpp"
#include <QLineEdit>
#include <QTableView>
#include <QHeaderView>
#include <QApplication>
#include <QSignalBlocker>
#include <QMetaProperty>
#include <QPersistentModelIndex>
#include <QKeyEvent>
namespace linaccess {
class AccessTableView final : public QTableView {
public:
    explicit AccessTableView(QWidget* parent=nullptr):QTableView(parent) {
        // Connect before enabling Qt's header sorting: an editor may contain
        // text that has not reached the model's manual-submit cache yet.
        connect(horizontalHeader(),&QHeaderView::sortIndicatorChanged,this,[this] {
            auto* typed=qobject_cast<AccessTableModel*>(model()); if(!typed) return;
            if(activeEditor()) typed->protectNextSort();
        });
    }
    bool hasPendingEditor() const {
        auto* editor=activeEditor(); if(!editor) return false;
        const auto property=editor->metaObject()->userProperty();
        const auto stored=editor->property("la_cell_index").value<QPersistentModelIndex>();
        const QModelIndex target=stored.isValid() ? QModelIndex(stored) : currentIndex();
        if(!property.isValid() || target.model()!=model()) return true;
        const auto value=property.read(editor),original=target.data(Qt::EditRole);
        return value!=original && value.toString()!=original.toString();
    }
    bool finishEditing() {
        auto* editor=activeEditor(); if(!editor) return true;
        const auto property=editor->metaObject()->userProperty(); if(!property.isValid() || !model()) return false;
        const auto stored=editor->property("la_cell_index").value<QPersistentModelIndex>();
        const QModelIndex target=stored.isValid() ? QModelIndex(stored) : currentIndex();
        if(target.model()!=model() || !model()->setData(target,property.read(editor),Qt::EditRole)) return false;
        QTableView::closeEditor(editor,QAbstractItemDelegate::NoHint); return true;
    }
    bool insertRecord() {
        auto* typed=qobject_cast<AccessTableModel*>(model());
        if(!typed || !finishEditing()) return false;
        const auto row=typed->rowCount();
        if(!typed->insertRow(row)) return false;
        setCurrentIndex(typed->index(row,0)); scrollTo(currentIndex()); setFocus();
        return true;
    }
    bool deleteCurrentRecord() {
        auto* typed=qobject_cast<AccessTableModel*>(model());
        if(!typed || !currentIndex().isValid() || !finishEditing()) return false;
        const auto row=currentIndex().row(),column=currentIndex().column();
        if(!typed->removeRow(row)) return false;
        if(typed->rowCount()) setCurrentIndex(typed->index(qMin(row,typed->rowCount()-1),column));
        return true;
    }
    void setModel(QAbstractItemModel* model) override {
        disconnect(restore_); QTableView::setModel(model);
        if(auto* typed=qobject_cast<AccessTableModel*>(model))
            restore_=connect(typed,&AccessTableModel::sortRestored,this,[this](int column,Qt::SortOrder order) {
                const QSignalBlocker blocked(horizontalHeader()); horizontalHeader()->setSortIndicator(column,order);
            });
    }
protected:
    void keyPressEvent(QKeyEvent* event) override {
        if(const auto text=nativeClockKeyText(*event)) {
            if(currentIndex().isValid() && currentIndex().flags().testFlag(Qt::ItemIsEditable)) {
                const bool fresh=!activeEditor();
                if(fresh) edit(currentIndex());
                if(auto* editor=qobject_cast<QLineEdit*>(activeEditor());editor && editor->isEnabled() && !editor->isReadOnly()) {
                    if(fresh) editor->selectAll();
                    if(editor->text().size()-editor->selectedText().size()+text->size()<=editor->maxLength()) editor->insert(*text);
                }
            }
            event->accept(); return;
        }
        const auto modifiers=event->modifiers() & ~Qt::KeypadModifier;
        if((modifiers==Qt::ControlModifier || modifiers==(Qt::ControlModifier|Qt::ShiftModifier))
            && (event->key()==Qt::Key_Plus || (event->key()==Qt::Key_Equal && modifiers.testFlag(Qt::ShiftModifier)))) {
            (void)insertRecord(); event->accept(); return;
        }
        if(modifiers==Qt::ControlModifier && event->key()==Qt::Key_Minus) {
            (void)deleteCurrentRecord(); event->accept(); return;
        }
        QTableView::keyPressEvent(event);
    }
private:
    QWidget* activeEditor() const {
        auto* editor=QApplication::focusWidget();
        while(editor && editor->parentWidget()!=viewport()) editor=editor->parentWidget();
        if(editor && editor->isVisible()) return editor;
        for(auto* candidate:viewport()->findChildren<QWidget*>(QString{},Qt::FindDirectChildrenOnly))
            if(candidate->isVisible() && candidate->property("la_cell_index").value<QPersistentModelIndex>().isValid()) return candidate;
        return nullptr;
    }
    QMetaObject::Connection restore_;
};
}
