#pragma once
#include "vba_runtime.hpp"
#include <QTableView>
#include <QHeaderView>
#include <QItemSelectionModel>
#include <limits>

namespace linaccess {
// Build everything before changing the clipboard. Never export hidden cells or
// turn a nonrectangular selection into an apparently complete data block.
template<class CellText>
QString nativeGridClipboardText(const QTableView& view,CellText cellText) {
    if(!view.selectionModel()) throw VbaRuntimeError(2046,"No grid selection");
    qint64 cells=0;
    for(const auto& range:view.selectionModel()->selection()) {
        const qint64 count=static_cast<qint64>(range.height())*range.width();
        if(count>1000000-cells) throw VbaRuntimeError(3251,"Clipboard selection exceeds the cell limit");
        cells+=count;
    }
    const auto selected=view.selectionModel()->selectedIndexes();
    if(selected.isEmpty()) throw VbaRuntimeError(2046,"No selected cells");
    if(selected.size()>1000000) throw VbaRuntimeError(3251,"Clipboard selection exceeds the cell limit");
    QVector<int> columns,ranks(view.model()->columnCount(),-1);
    for(int visual=0;visual<view.horizontalHeader()->count();++visual) {const int column=view.horizontalHeader()->logicalIndex(visual);if(column>=0 && column<ranks.size() && !view.isColumnHidden(column)) {ranks[column]=columns.size();columns.push_back(column);}}
    int top=std::numeric_limits<int>::max(),left=top,bottom=-1,right=-1;qsizetype visibleCells=0;
    for(const auto& index:selected) {
        if(view.isRowHidden(index.row())) throw VbaRuntimeError(3251,"Hidden clipboard rows are unsupported");
        if(view.isColumnHidden(index.column())) {if(view.property("la_projected_query").toBool()) continue;throw VbaRuntimeError(3251,"Hidden clipboard cells are unsupported");}
        const int rank=ranks[index.column()];if(rank<0) throw VbaRuntimeError(3251,"Clipboard column layout changed");
        ++visibleCells;top=qMin(top,index.row());bottom=qMax(bottom,index.row());left=qMin(left,rank);right=qMax(right,rank);
    }
    if(!visibleCells) throw VbaRuntimeError(2046,"No visible clipboard selection");
    if(static_cast<qint64>(bottom-top+1)*(right-left+1)!=visibleCells) throw VbaRuntimeError(3251,"Select a contiguous visible rectangle");
    QString output;
    for(int row=top;row<=bottom;++row) {
        if(row>top) output+="\r\n";
        for(int column=left;column<=right;++column) {
            if(column>left) output+='\t';
            auto text=cellText(view.model()->index(row,columns[column]));
            if(text.contains(QChar(0))) throw VbaRuntimeError(3251,"Clipboard NUL characters are unsupported");
            if(text.contains('\t') || text.contains('\r') || text.contains('\n') || text.contains('"')) {
                text.replace("\"","\"\""); text='"'+text+'"';
            }
            if(output.size()+text.size()>4*1024*1024) throw VbaRuntimeError(3251,"Clipboard text exceeds the size limit");
            output+=text;
        }
    }
    return output;
}
}
