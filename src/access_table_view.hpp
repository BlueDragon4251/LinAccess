#pragma once
#include "access_table_model.hpp"
#include <QTableView>
#include <QHeaderView>
#include <QApplication>
#include <QSignalBlocker>
namespace linaccess {
class AccessTableView final : public QTableView {
public:
    explicit AccessTableView(QWidget* parent=nullptr):QTableView(parent) {
        // Connect before enabling Qt's header sorting: an editor may contain
        // text that has not reached the model's manual-submit cache yet.
        connect(horizontalHeader(),&QHeaderView::sortIndicatorChanged,this,[this] {
            auto* typed=qobject_cast<AccessTableModel*>(model()); if(!typed) return;
            auto* editor=QApplication::focusWidget();
            while(editor && editor->parentWidget()!=viewport()) editor=editor->parentWidget();
            if(editor && editor->isVisible()) typed->protectNextSort();
        });
    }
    void setModel(QAbstractItemModel* model) override {
        disconnect(restore_); QTableView::setModel(model);
        if(auto* typed=qobject_cast<AccessTableModel*>(model))
            restore_=connect(typed,&AccessTableModel::sortRestored,this,[this](int column,Qt::SortOrder order) {
                const QSignalBlocker blocked(horizontalHeader()); horizontalHeader()->setSortIndicator(column,order);
            });
    }
private:
    QMetaObject::Connection restore_;
};
}
