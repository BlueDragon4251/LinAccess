#pragma once

#include <QKeyEvent>
#include <QItemSelectionModel>
#include <QPointer>
#include <QMouseEvent>
#include <QScopedValueRollback>
#include <QTreeView>
#include <QVariant>
#include <functional>

// A persistent single-selection row-source view. It shares the typed model
// roles with native combos, but has no popup or editable text surrogate.
class AccessListBox final:public QTreeView {
public:
    explicit AccessListBox(QWidget* parent=nullptr):QTreeView(parent) {
        setRootIsDecorated(false); setHeaderHidden(true); setItemsExpandable(false);
        setSelectionBehavior(QAbstractItemView::SelectRows);
        setSelectionMode(QAbstractItemView::SingleSelection);
        setEditTriggers(QAbstractItemView::NoEditTriggers);
        setUniformRowHeights(true); setAllColumnsShowFocus(true);
    }
    std::function<void()> selectionEdited;
    void setModel(QAbstractItemModel* next) override {
        QScopedValueRollback<bool> guard(assigning_,true);
        auto* previous=model(); disconnect(selectionConnection_);
        const QPointer<QItemSelectionModel> previousSelection=selectionModel();
        QTreeView::setModel(next);
        if(previousSelection && previousSelection!=selectionModel() && previousSelection->parent()==this) previousSelection->deleteLater();
        if(previous && previous!=next && previous->parent()==this) delete previous;
        if(selectionModel()) selectionConnection_=connect(selectionModel(),&QItemSelectionModel::selectionChanged,this,[this] {
            if(assigning_) return;
            const auto rows=selectionModel()->selectedRows();
            valueText_=rows.isEmpty() ? QString{} : boundData(rows.front().row()).toString();
            if(rows.isEmpty()) { QScopedValueRollback<bool> guard(assigning_,true); setCurrentIndex({}); }
            if(selectionEdited) selectionEdited();
        });
    }
    int rowCount() const { return model() ? model()->rowCount() : 0; }
    QVariant boundData(int row) const { return model() ? model()->index(row,0).data(Qt::UserRole) : QVariant{}; }
    QVariantList rowValues(int row) const { return model() ? model()->index(row,0).data(Qt::UserRole+1).toList() : QVariantList{}; }
    QString valueText() const { return valueText_; }
    void setDisplayColumn(int column) { displayColumn_=column; }
    void setValueText(const QString& text,bool isNull) {
        QScopedValueRollback<bool> guard(assigning_,true); valueText_=text;
        if(!isNull) for(int row=0;row<rowCount();++row) {
            const auto value=boundData(row);
            if(!value.isNull() && value.toString()==text) {
                setCurrentIndex(model()->index(row,displayColumn_));
                scrollTo(currentIndex()); return;
            }
        }
        setCurrentIndex({}); clearSelection();
    }
protected:
    void keyPressEvent(QKeyEvent* event) override {
        if(property("la_access_input_readonly").toBool() && event->key()!=Qt::Key_Tab && event->key()!=Qt::Key_Backtab) { event->accept(); return; }
        QTreeView::keyPressEvent(event);
    }
    void mousePressEvent(QMouseEvent* event) override {
        if(property("la_access_input_readonly").toBool()) { setFocus(Qt::MouseFocusReason); event->accept(); return; }
        QTreeView::mousePressEvent(event);
    }
    void mouseDoubleClickEvent(QMouseEvent* event) override {
        if(property("la_access_input_readonly").toBool()) { event->accept(); return; }
        QTreeView::mouseDoubleClickEvent(event);
    }
private:
    QMetaObject::Connection selectionConnection_;
    QString valueText_;
    int displayColumn_{};
    bool assigning_{};
};
