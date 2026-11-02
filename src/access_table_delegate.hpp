#pragma once
#include <QStyledItemDelegate>
#include <QLineEdit>
#include <QPersistentModelIndex>
#include <QVariant>
#include <limits>
namespace linaccess {
// QString EditRole keeps Currency out of the default two-decimal spin editor.
class AccessTableDelegate final : public QStyledItemDelegate {
public:
    using QStyledItemDelegate::QStyledItemDelegate;
    QWidget* createEditor(QWidget* parent,const QStyleOptionViewItem& option,const QModelIndex& index) const override {
        auto* editor=QStyledItemDelegate::createEditor(parent,option,index);
        if(editor) editor->setProperty("la_cell_index",QVariant::fromValue(QPersistentModelIndex(index)));
        if(auto* text=qobject_cast<QLineEdit*>(editor)) text->setMaxLength(std::numeric_limits<int>::max());
        return editor;
    }
};
}
