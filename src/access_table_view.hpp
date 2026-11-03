#pragma once
#include "access_table_model.hpp"
#include "vba_runtime.hpp"
#include "native_clock_keys.hpp"
#include <QLineEdit>
#include <QClipboard>
#include <QItemSelectionModel>
#include <limits>
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
    bool copySelection() {
        auto* typed=qobject_cast<AccessTableModel*>(model());
        if(!typed || !selectionModel() || !finishEditing()) return false;
        const auto selected=selectionModel()->selectedIndexes();
        if(selected.isEmpty()) return false;
        if(selected.size()>1000000) return refuse(3251,"Die Auswahl ist für die Textzwischenablage zu groß.");
        int top=std::numeric_limits<int>::max(),left=top,bottom=-1,right=-1;
        for(const auto& index:selected) {
            if(isRowHidden(index.row()) || isColumnHidden(index.column())) return refuse(3251,"Ausgeblendete Zellen werden nicht in die Textzwischenablage übernommen.");
            top=qMin(top,index.row()); bottom=qMax(bottom,index.row()); left=qMin(left,index.column()); right=qMax(right,index.column());
        }
        if(static_cast<qint64>(bottom-top+1)*(right-left+1)!=selected.size())
            return refuse(3251,"Zum Kopieren einen zusammenhängenden rechteckigen Bereich auswählen.");
        QString output;
        try {
            for(int row=top;row<=bottom;++row) {
                if(row>top) output+="\r\n";
                for(int column=left;column<=right;++column) {
                    if(column>left) output+='\t';
                    auto text=typed->clipboardText(typed->index(row,column));
                    if(text.contains(QChar(0))) return refuse(3251,"Nullzeichen werden in der Textzwischenablage noch nicht unterstützt.");
                    if(text.contains('\t') || text.contains('\r') || text.contains('\n') || text.contains('"')) {
                        text.replace("\"","\"\""); text='"'+text+'"';
                    }
                    if(output.size()+text.size()>4*1024*1024) return refuse(3251,"Der Text ist für die Zwischenablage zu groß.");
                    output+=text;
                }
            }
        } catch(const VbaRuntimeError& error) { return refuse(error.number,"Dieser Zellwert kann noch nicht verlustfrei als Text kopiert werden."); }
        QApplication::clipboard()->setText(output);
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
        const auto selectionModifiers=event->modifiers() & ~Qt::KeypadModifier;
        if(event->key()==Qt::Key_C && selectionModifiers==Qt::ControlModifier) {
            (void)copySelection(); event->accept(); return;
        }
        if((event->key()==Qt::Key_Space && (selectionModifiers==Qt::ShiftModifier || selectionModifiers==Qt::ControlModifier))
            || (event->key()==Qt::Key_A && selectionModifiers==Qt::ControlModifier)) {
            if(finishEditing() && selectionModel() && model()) {
                if(event->key()==Qt::Key_Space && selectionModifiers==Qt::ShiftModifier) {
                    if(currentIndex().isValid()) selectionModel()->select(currentIndex(),QItemSelectionModel::ClearAndSelect|QItemSelectionModel::Rows);
                } else if(fetchSelectionRows()) {
                    if(event->key()==Qt::Key_A) selectAll();
                    else if(currentIndex().isValid()) selectionModel()->select(currentIndex(),QItemSelectionModel::ClearAndSelect|QItemSelectionModel::Columns);
                }
            }
            event->accept(); return;
        }
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
    bool refuse(int number,const QString& message) const {
        if(auto* typed=qobject_cast<AccessTableModel*>(model())) emit typed->editRejected(number,message);
        return false;
    }
    bool fetchSelectionRows() {
        while(model() && model()->canFetchMore({})) {
            if(static_cast<qint64>(model()->rowCount())*qMax(1,model()->columnCount())>=1000000)
                return refuse(3251,"Die vollständige Auswahl überschreitet den unterstützten Umfang.");
            const auto count=model()->rowCount(); model()->fetchMore({});
            if(model()->rowCount()==count) return refuse(3251,"Die vollständige Auswahl konnte nicht geladen werden.");
        }
        if(model() && static_cast<qint64>(model()->rowCount())*qMax(1,model()->columnCount())>1000000)
            return refuse(3251,"Die vollständige Auswahl überschreitet den unterstützten Umfang.");
        return model()!=nullptr;
    }
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
