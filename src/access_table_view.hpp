#pragma once
#include "access_table_model.hpp"
#include <QTableView>
#include <QHeaderView>
#include <QApplication>
#include <QSignalBlocker>
#include <QMetaProperty>
#include <QPersistentModelIndex>
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
    void setModel(QAbstractItemModel* model) override {
        disconnect(restore_); QTableView::setModel(model);
        if(auto* typed=qobject_cast<AccessTableModel*>(model))
            restore_=connect(typed,&AccessTableModel::sortRestored,this,[this](int column,Qt::SortOrder order) {
                const QSignalBlocker blocked(horizontalHeader()); horizontalHeader()->setSortIndicator(column,order);
            });
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
