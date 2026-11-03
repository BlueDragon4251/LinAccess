#pragma once
#include "vba_runtime.hpp"
#include <QTableView>
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
    int top=std::numeric_limits<int>::max(),left=top,bottom=-1,right=-1;
    for(const auto& index:selected) {
        if(view.isRowHidden(index.row()) || view.isColumnHidden(index.column())) throw VbaRuntimeError(3251,"Hidden clipboard cells are unsupported");
        top=qMin(top,index.row()); bottom=qMax(bottom,index.row()); left=qMin(left,index.column()); right=qMax(right,index.column());
    }
    if(static_cast<qint64>(bottom-top+1)*(right-left+1)!=selected.size()) throw VbaRuntimeError(3251,"Select a contiguous rectangle");
    QString output;
    for(int row=top;row<=bottom;++row) {
        if(row>top) output+="\r\n";
        for(int column=left;column<=right;++column) {
            if(column>left) output+='\t';
            auto text=cellText(view.model()->index(row,column));
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
