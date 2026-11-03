#pragma once
#include <QLineEdit>
#include <QMetaProperty>
#include <QModelIndex>
namespace linaccess {
inline bool nativeCellEditorChanged(QWidget* editor,const QModelIndex& target) {
    if(!editor || !target.isValid()) return true;
    const auto property=editor->metaObject()->userProperty();
    if(!property.isValid()) return true;
    const auto value=property.read(editor),original=target.data(Qt::EditRole);
    // An untouched empty editor must retain SQL NULL. Explicit user editing to
    // empty text can still request a distinct empty string for a text column.
    if(auto* line=qobject_cast<QLineEdit*>(editor);line && original.isNull() && line->isModified()) return true;
    return value!=original && value.toString()!=original.toString();
}
}
