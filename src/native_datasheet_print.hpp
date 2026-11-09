#pragma once
#include "access_table_model.hpp"
#include "access_query.hpp"
#include "vba_runtime.hpp"
#include <QHeaderView>
#include <QPointer>
#include <QTableView>
#include <QTextDocument>
#include <QTextCursor>
#include <QTextTable>
#include <QPdfWriter>
#include <QSaveFile>
#include <QFileInfo>
#include <QPainter>
#include <QImage>
#include <QAbstractTextDocumentLayout>
#include <QSqlError>
#include <QScopeGuard>
#include <cmath>
#include <memory>

namespace linaccess {
// Own text/widths/row ordinals only; retain no live SQL cursor or draft.
struct NativeDatasheetSnapshot {
    QString title;
    QStringList headers;
    QVector<int> widths;
    QVector<QStringList> rows;
    QVector<int> sourceRows;
};
inline NativeDatasheetSnapshot nativeDatasheetSnapshot(QTableView& grid,const QString& title) {
    const QPointer<QTableView> view=&grid;
    const QPointer<QAbstractItemModel> model=grid.model();
    auto* table=qobject_cast<AccessTableModel*>(model.data());
    auto* query=dynamic_cast<AccessQueryModel*>(model.data());
    if(!table && !query) throw VbaRuntimeError(3251,"Kein unterstütztes natives Datenblatt.");
    if(table && table->isDirty()) throw VbaRuntimeError(2501,"Vor der Ausgabe Tabellenentwürfe speichern oder verwerfen.");
    bool changed=false;
    QObject observer;
    const auto invalidate=[&] { changed=true; };
    QObject::connect(model,&QAbstractItemModel::modelReset,&observer,invalidate);
    QObject::connect(model,&QAbstractItemModel::dataChanged,&observer,invalidate);
    QObject::connect(model,&QAbstractItemModel::layoutChanged,&observer,invalidate);
    QObject::connect(model,&QAbstractItemModel::rowsRemoved,&observer,invalidate);
    QObject::connect(model,&QAbstractItemModel::columnsInserted,&observer,invalidate);
    QObject::connect(model,&QAbstractItemModel::columnsRemoved,&observer,invalidate);
    const auto check=[&] {
        if(!view || !model || view->model()!=model || changed) throw VbaRuntimeError(2501,"Datenblatt wurde während der Ausgabevorbereitung geändert oder geschlossen.");
    };
    NativeDatasheetSnapshot snapshot;
    QVector<int> columns;
    qsizetype characters=0;
    const auto checkedText=[&](QString text) {
        if(text.contains(QChar(0)) || !text.isValidUtf16()) throw VbaRuntimeError(3251,"Nicht darstellbarer Text im Datenblatt.");
        if(text.size()>16*1024*1024-characters) throw VbaRuntimeError(3251,"Drucktext überschreitet die native Grenze von 16 Mi UTF-16-Einheiten.");
        characters+=text.size(); return text;
    };
    snapshot.title=checkedText(title);
    for(int visual=0;visual<model->columnCount();++visual) {
        const int column=grid.horizontalHeader()->logicalIndex(visual);
        if(grid.isColumnHidden(column)) continue;
        columns.append(column);
        snapshot.headers.append(checkedText(model->headerData(column,Qt::Horizontal).toString()));
        snapshot.widths.append(qMax(1,grid.columnWidth(column)));
    }
    if(columns.isEmpty()) throw VbaRuntimeError(2046,"Keine sichtbaren Datenblattspalten.");
    while(model->canFetchMore({})) {
        if(static_cast<qint64>(model->rowCount())*columns.size()>=1000000) throw VbaRuntimeError(3251,"Druckumfang überschreitet eine Million Zellen.");
        const int before=model->rowCount(); model->fetchMore({}); check();
        if(model->rowCount()<=before) throw VbaRuntimeError(3251,"Datenblatt konnte nicht vollständig geladen werden.");
    }
    check();
    if(table && table->lastError().isValid()) throw VbaRuntimeError(3146,"Datenblatt konnte nicht fehlerfrei geladen werden.");
    if(static_cast<qint64>(model->rowCount())*columns.size()>1000000) throw VbaRuntimeError(3251,"Druckumfang überschreitet eine Million Zellen.");
    // Vertical visual order matters too when record headers have been moved.
    for(int visual=0;visual<model->rowCount();++visual) {
        const int row=grid.verticalHeader()->logicalIndex(visual);
        if(grid.isRowHidden(row)) continue;
        QStringList record;
        for(int column:columns) {
            const auto index=model->index(row,column);
            record.append(checkedText(table ? table->formattedText(index) : query->clipboardText(index)));
            check();
        }
        snapshot.rows.append(std::move(record)); snapshot.sourceRows.append(row);
    }
    check();
    if(table && table->isDirty()) throw VbaRuntimeError(2501,"Datenblatt enthält neue Entwürfe.");
    return snapshot;
}
inline std::unique_ptr<QTextDocument> nativeDatasheetDocument(const NativeDatasheetSnapshot& snapshot) {
    auto doc=std::make_unique<QTextDocument>();
    QFont font; font.setPointSizeF(9); doc->setDefaultFont(font);
    doc->setDocumentMargin(2); // Reserve the table border inside the printable area.
    auto options=doc->defaultTextOption(); options.setWrapMode(QTextOption::WrapAtWordBoundaryOrAnywhere); doc->setDefaultTextOption(options);
    QTextCursor cursor(doc.get()); cursor.insertText(snapshot.title); cursor.insertBlock();
    QTextTableFormat format; format.setMargin(0); format.setHeaderRowCount(1); format.setCellPadding(3);
    format.setCellSpacing(0); format.setBorder(0.5);
    format.setWidth(QTextLength(QTextLength::PercentageLength,100));
    qint64 total=0; for(int width:snapshot.widths) total+=width;
    QList<QTextLength> widths;
    for(int width:snapshot.widths) widths.append(QTextLength(QTextLength::PercentageLength,100.0*width/total));
    format.setColumnWidthConstraints(widths);
    auto* table=cursor.insertTable(snapshot.rows.size()+1,snapshot.headers.size(),format);
    for(int col=0;col<snapshot.headers.size();++col) {
        auto cell=table->cellAt(0,col); auto style=cell.format(); style.setBackground(QColor(230,230,230)); cell.setFormat(style);
        auto text=cell.firstCursorPosition(); QTextCharFormat bold; bold.setFontWeight(QFont::Bold); text.insertText(snapshot.headers[col],bold);
    }
    for(int row=0;row<snapshot.rows.size();++row) for(int col=0;col<snapshot.headers.size();++col)
        table->cellAt(row+1,col).firstCursorPosition().insertText(snapshot.rows[row][col]);
    return doc;
}
// Render explicitly to observe painter/page failures; QTextDocument::print is void.
inline void renderNativeDatasheet(const NativeDatasheetSnapshot& snapshot,QPagedPaintDevice& device,int first=1,int last=0,bool reverse=false,int copies=1,bool collated=true) {
    if(copies<1 || copies>100) throw VbaRuntimeError(3251,"Datenblatt unterstützt 1 bis 100 Druckexemplare.");
    QImage pointMetrics(1,1,QImage::Format_ARGB32);
    pointMetrics.setDotsPerMeterX(2835); pointMetrics.setDotsPerMeterY(2835); // 72 dpi, rounded to integral dots/metre.
    auto doc=nativeDatasheetDocument(snapshot);
    doc->documentLayout()->setPaintDevice(&pointMetrics);
    QPainter painter;
    if(!painter.begin(&device)) throw VbaRuntimeError(3146,"Druckgerät konnte nicht geöffnet werden.");
    const auto finish=qScopeGuard([&] { if(painter.isActive()) painter.end(); });
    // The document uses typographic points independent of printer/PDF resolution.
    const qreal scaleX=device.logicalDpiX()/72.0,scaleY=device.logicalDpiY()/72.0;
    const QSizeF size(device.width()/scaleX,device.height()/scaleY);
    if(size.width()<72 || size.height()<72) throw VbaRuntimeError(3251,"Druckfläche ist zu klein.");
    doc->setPageSize(size);
    // Qt positions the table grid after its outer cell padding/border; a 100%
    // table can therefore exceed the page even when its reported width fits.
    // Reserve both edges explicitly in points, retaining proportional columns.
    for(auto* frame:doc->rootFrame()->childFrames()) if(auto* table=qobject_cast<QTextTable*>(frame)) {
        auto format=table->format();
        const qreal width=size.width()-2*doc->documentMargin()-2*(format.cellPadding()+format.border());
        if(width<=0) throw VbaRuntimeError(3251,"Keine verfügbare Datenblatt-Druckbreite.");
        format.setWidth(QTextLength(QTextLength::FixedLength,width)); table->setFormat(format);
    }
    const int count=doc->pageCount();
    for(auto* frame:doc->rootFrame()->childFrames()) {
        const auto bounds=doc->documentLayout()->frameBoundingRect(frame);
        if(bounds.left()< -0.5 || bounds.right()>size.width()+0.5)
            throw VbaRuntimeError(3251,QString("Datenblattbreite %1 pt (links %3, rechts %4) passt nicht in %2 pt Druckfläche; Spalten ausblenden oder Seiteneinstellungen ändern.").arg(bounds.width(),0,'f',1).arg(size.width(),0,'f',1).arg(bounds.left(),0,'f',1).arg(bounds.right(),0,'f',1).toStdString());
    }
    const int begin=qMax(1,first),end=last>0 ? qMin(count,last) : count;
    if(begin>end) throw VbaRuntimeError(5,"Der gewählte Seitenbereich enthält keine Datenblattseite.");
    if(copies>1 && static_cast<qint64>(end-begin+1)*copies>10000) throw VbaRuntimeError(3251,"Mehrfache Datenblattausgabe überschreitet 10000 Seiten.");
    int printed=0;
    const auto draw=[&](int page) {
        if(printed++ && !device.newPage()) throw VbaRuntimeError(3146,"Neue Druckseite konnte nicht angelegt werden.");
        painter.save(); painter.scale(scaleX,scaleY);
        painter.setClipRect(QRectF(0,0,size.width(),size.height()));
        painter.translate(0,-(page-1)*size.height());
        QAbstractTextDocumentLayout::PaintContext context;
        context.clip=QRectF(0,(page-1)*size.height(),size.width(),size.height());
        doc->documentLayout()->draw(&painter,context); painter.restore();
    };
    if(collated) for(int copy=0;copy<copies;++copy) for(int offset=0;offset<=end-begin;++offset) draw(reverse ? end-offset : begin+offset);
    else for(int offset=0;offset<=end-begin;++offset) for(int copy=0;copy<copies;++copy) draw(reverse ? end-offset : begin+offset);
    if(!painter.end()) throw VbaRuntimeError(3146,"Druckausgabe konnte nicht abgeschlossen werden.");
}
inline void exportNativeDatasheetPdf(const NativeDatasheetSnapshot& snapshot,const QString& path,
    const QPageLayout& layout=QPageLayout(QPageSize(QPageSize::A4),QPageLayout::Portrait,QMarginsF(15,15,15,15)),
    int first=1,int last=0,bool reverse=false,int copies=1,bool collated=true) {
    if(path.contains(QChar(0))) throw VbaRuntimeError(5,"Ungültiger PDF-Zielname.");
    if(QFileInfo(path).suffix().compare("pdf",Qt::CaseInsensitive)!=0) throw VbaRuntimeError(5,"PDF-Ausgabe benötigt eine .pdf-Zieldatei.");
    if(QFileInfo(path).isSymLink()) throw VbaRuntimeError(75,"PDF-Ziel darf kein symbolischer Link sein.");
    QSaveFile output(path); output.setDirectWriteFallback(false);
    if(!output.open(QIODevice::WriteOnly)) throw VbaRuntimeError(75,"PDF-Zieldatei konnte nicht vorbereitet werden.");
    {
        QPdfWriter writer(&output); writer.setTitle(snapshot.title); writer.setCreator("LinAccess");
        if(!writer.setPageLayout(layout)) throw VbaRuntimeError(3251,"PDF-Seiteneinstellungen werden nicht unterstützt.");
        renderNativeDatasheet(snapshot,writer,first,last,reverse,copies,collated);
    }
    if(output.error()!=QFileDevice::NoError || !output.commit()) throw VbaRuntimeError(75,"PDF-Zieldatei konnte nicht atomar gespeichert werden.");
}
}
