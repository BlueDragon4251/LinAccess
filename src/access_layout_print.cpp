#include "access_form_preview.hpp"
#include "native_report_text_layout.hpp"
#include <QAction>
#include <QFileDialog>
#include <QFile>
#include <QFileInfo>
#include <QMessageBox>
#include <QPainter>
#include <QPrintDialog>
#include <QPrinter>
#include <QPrintPreviewDialog>
#include <QSaveFile>
#include <QTemporaryDir>
#include <QToolBar>
#include <QScopeGuard>
#include <QScopedValueRollback>
#include <QRegion>
#include <QPicture>
#include <QtMath>
#include <QVBoxLayout>
#include <QPrintPreviewWidget>
#include <QTabWidget>
#include <QHBoxLayout>
#include <QPushButton>
#include <QSpinBox>
#include <QLabel>
#include <QSignalBlocker>

namespace {
class NativeReportPageView final : public QWidget {
public:
    NativeReportPageView(std::function<bool(QPrinter&)> render,std::function<void()> changed,
                         std::function<void()> refreshRequested,QWidget* parent)
        : QWidget(parent),printer_(QPrinter::HighResolution) {
        setObjectName("access-report-page-view");
        auto* layout=new QVBoxLayout(this);
        auto* navigation=new QHBoxLayout;
        auto* previous=new QPushButton("Zurück",this);
        auto* next=new QPushButton("Weiter",this);
        auto* number=new QSpinBox(this); number->setRange(1,32767);
        number->setObjectName("access-report-page-number");
        auto* total=new QLabel(this);
        auto* smaller=new QPushButton("−",this);
        auto* larger=new QPushButton("+",this);
        smaller->setAccessibleName("Verkleinern"); larger->setAccessibleName("Vergrößern");
        auto* refresh=new QPushButton("Aktualisieren",this);
        for (auto* button:{previous,next}) navigation->addWidget(button);
        navigation->addWidget(number); navigation->addWidget(total); navigation->addStretch();
        navigation->addWidget(smaller); navigation->addWidget(larger); navigation->addWidget(refresh);
        layout->addLayout(navigation);
        pages_=new QPrintPreviewWidget(&printer_,this);
        pages_->setObjectName("access-report-pages");
        layout->addWidget(pages_,1);
        connect(pages_,&QPrintPreviewWidget::paintRequested,this,[this,render](QPrinter* printer) {
            setProperty("la_report_pages_valid",false);
            setProperty("la_report_pages_valid",render(*printer));
        });
        connect(pages_,&QPrintPreviewWidget::previewChanged,this,[this,previous,next,number,total,changed] {
            const int count=pages_->pageCount(),page=pages_->currentPage();
            const QSignalBlocker blocked(number);
            number->setMaximum(qMax(1,count)); number->setValue(qMax(1,page));
            total->setText("von "+QString::number(count));
            previous->setEnabled(page>1); next->setEnabled(page<count);
            changed();
        });
        connect(previous,&QPushButton::clicked,pages_,[this] { pages_->setCurrentPage(pages_->currentPage()-1); });
        connect(next,&QPushButton::clicked,pages_,[this] { pages_->setCurrentPage(pages_->currentPage()+1); });
        connect(number,&QSpinBox::valueChanged,pages_,&QPrintPreviewWidget::setCurrentPage);
        connect(smaller,&QPushButton::clicked,pages_,[this] { pages_->zoomOut(); });
        connect(larger,&QPushButton::clicked,pages_,[this] { pages_->zoomIn(); });
        connect(refresh,&QPushButton::clicked,this,[refreshRequested] { refreshRequested(); });
    }
    ~NativeReportPageView() override { delete pages_; } // printer outlives its preview.
    bool refreshPages() {
        setProperty("la_report_pages_valid",false);
        pages_->updatePreview();
        return property("la_report_pages_valid").toBool() && pages_->pageCount()>0;
    }
private:
    QPrinter printer_;
    QPrintPreviewWidget* pages_{};
};
void synchronizeReportView(QWidget* document,NativeReportPageView* pages) {
    auto* tabs=document->findChild<QTabWidget*>("access-report-view-tabs");
    auto* notice=document->findChild<QLabel*>("access-report-view-notice");
    auto* view=pages ? pages->findChild<QPrintPreviewWidget*>("access-report-pages") : nullptr;
    if (!tabs || !notice || !view) return;
    const bool valid=pages->property("la_report_pages_valid").toBool() && view->pageCount()>0;
    tabs->setTabEnabled(0,valid); notice->setVisible(!valid);
    if (!valid) { tabs->setProperty("la_report_fallback",true); tabs->setCurrentIndex(1); }
    else if (tabs->property("la_report_fallback").toBool()) {
        tabs->setProperty("la_report_fallback",false); tabs->setCurrentIndex(0);
    }
}
}

void AccessFormPreview::buildReportPageView(QVBoxLayout* layout,QWidget* layoutView) {
    if (reportPageView_ || closed_ || openCancelled_ || layout_.kind!=linaccess::AccessDocumentKind::Report) return;
    auto* holder=new QWidget(this);
    auto* old=layout->replaceWidget(layoutView,holder); delete old;
    auto* content=new QVBoxLayout(holder); content->setContentsMargins(0,0,0,0);
    auto* notice=new QLabel("Nur Layoutansicht: Die vollständige Berichtsausgabe ist für dieses Layout noch nicht unterstützt.",holder);
    notice->setObjectName("access-report-view-notice"); notice->setWordWrap(true);
    content->addWidget(notice);
    auto* tabs=new QTabWidget(holder); tabs->setObjectName("access-report-view-tabs");
    content->addWidget(tabs,1);
    auto* pages=new NativeReportPageView([this](QPrinter& printer) { return renderReport(printer); },
        [this] {
            if (!reportAggregateCaching_ && !closed_) {
                synchronizeReportView(this,dynamic_cast<NativeReportPageView*>(reportPageView_.data()));
                refreshCalculatedControls();
            }
        },[this] { refreshReportPageView(); },tabs);
    reportPageView_=pages;
    tabs->addTab(pages,"Bericht"); tabs->addTab(layoutView,"Aktuelles Layout");
    refreshReportPageView();
}
void AccessFormPreview::refreshReportPageView() {
    if (!reportPageView_ || reportAggregateCaching_ || closed_ || openCancelled_) return;
    auto* pages=dynamic_cast<NativeReportPageView*>(reportPageView_.data());
    if (!pages) return;
    pages->refreshPages(); synchronizeReportView(this,pages);
    refreshCalculatedControls();
}


void AccessFormPreview::buildPrintToolbar(QVBoxLayout* layout) {
    printToolbar_ = new QToolBar("Aktuelles Layout drucken",this);
    printToolbar_->setObjectName("access-print-preview-toolbar");
    layout->addWidget(printToolbar_);
    auto* preview = printToolbar_->addAction("Druckvorschau: aktuelles Layout");
    connect(preview,&QAction::triggered,this,[this] {
        QPrinter printer(QPrinter::HighResolution);
        QPrintPreviewDialog dialog(&printer,this);
        dialog.setWindowTitle("Aktuelles Layout — eine Seite");
        connect(&dialog,&QPrintPreviewDialog::paintRequested,this,[this](QPrinter* target) {
            if (!renderLayout(*target)) QMessageBox::warning(this,"Druckvorschau","Das aktuelle Layout konnte nicht ausgegeben werden.");
        });
        dialog.exec();
    });
    auto* print = printToolbar_->addAction("Drucken …");
    connect(print,&QAction::triggered,this,[this] {
        QPrinter printer(QPrinter::HighResolution);
        QPrintDialog dialog(&printer,this);
        if (dialog.exec()==QDialog::Accepted && !renderLayout(printer))
            QMessageBox::warning(this,"Drucken","Das aktuelle Layout konnte nicht ausgegeben werden.");
    });
    auto* pdf = printToolbar_->addAction("Aktuelles Layout als PDF …");
    connect(pdf,&QAction::triggered,this,[this] {
        const auto path = QFileDialog::getSaveFileName(this,"Aktuelles Layout als PDF",{},"PDF (*.pdf)");
        if (!path.isEmpty() && !exportLayoutPdf(path))
            QMessageBox::warning(this,"PDF","Das aktuelle Layout konnte nicht gespeichert werden.");
    });
    auto* reportPreview=printToolbar_->addAction("Berichtsvorschau …");
    reportPreview->setObjectName("access-report-preview");
    connect(reportPreview,&QAction::triggered,this,[this] {
        QPrinter printer(QPrinter::HighResolution);
        QPrintPreviewDialog dialog(&printer,this);
        dialog.setWindowTitle("Berichtsvorschau");
        bool failed=false;
        connect(&dialog,&QPrintPreviewDialog::paintRequested,this,[this,&dialog,&failed](QPrinter* target) {
            if (!renderReport(*target)) {
                failed=true;
                QMetaObject::invokeMethod(&dialog,&QDialog::reject,Qt::QueuedConnection);
            }
        });
        dialog.exec();
        if (failed) QMessageBox::warning(this,"Berichtsvorschau",
            "Dieser Bericht enthält noch nicht unterstützte Druckregeln oder konnte nicht vollständig dargestellt werden.");
    });
    auto* reportPrint=printToolbar_->addAction("Bericht drucken …");
    reportPrint->setObjectName("access-report-print");
    connect(reportPrint,&QAction::triggered,this,[this] {
        QPrinter printer(QPrinter::HighResolution);
        QPrintDialog dialog(&printer,this);
        dialog.setOption(QAbstractPrintDialog::PrintPageRange,false);
        dialog.setOption(QAbstractPrintDialog::PrintSelection,false);
        if (dialog.exec()==QDialog::Accepted && !renderReport(printer))
            QMessageBox::warning(this,"Bericht drucken",
                "Dieser Bericht enthält noch nicht unterstützte Druckregeln oder konnte nicht vollständig ausgegeben werden.");
    });
    auto* reportPdf=printToolbar_->addAction("Bericht als PDF …");
    connect(reportPdf,&QAction::triggered,this,[this] {
        const auto path=QFileDialog::getSaveFileName(this,"Bericht als PDF",{},"PDF (*.pdf)");
        if (!path.isEmpty() && !exportReportPdf(path))
            QMessageBox::warning(this,"Bericht als PDF",
                "Dieser Bericht konnte nicht vollständig ausgegeben werden. "
                "Unterstützt werden feste Kopf-, Detail- und Fußbereiche mit direkten Datenfeldern. "
                "Gruppierung, wachsende Bereiche und Seitenereignisse sind noch nicht unterstützt.");
    });
    auto* close = printToolbar_->addAction("&Schließen");
    connect(close,&QAction::triggered,this,&AccessFormPreview::reject);

}
bool AccessFormPreview::renderLayout(QPrinter& printer) {
    auto* canvas = findChild<QWidget*>("access-form-canvas");
    if (!canvas || canvas->width()<=0 || canvas->height()<=0) return false;
    QPainter painter;
    if (!painter.begin(&printer)) return false;
    const auto page = printer.pageRect(QPrinter::DevicePixel);
    const auto scale = qMin(page.width()/canvas->width(),page.height()/canvas->height());
    painter.translate(page.topLeft());
    painter.scale(scale,scale);
    canvas->render(&painter);
    return painter.end();
}
bool AccessFormPreview::renderReport(QPrinter& printer) {
    for (const auto* event:{"OnOpen","OnLoad","OnActivate","OnCurrent","OnNoData"})
        if (const auto result=eventResult(event);result && (result->status==linaccess::VbaExecutionStatus::Error
            || result->status==linaccess::VbaExecutionStatus::Unsupported)) return false;
    if (reportAggregateCaching_ || layout_.kind!=linaccess::AccessDocumentKind::Report || !dataSourceAvailable_
        || dataTruncated_ || !layout_.reportPrintLimitations.isEmpty()) return false;
    auto* canvas=findChild<QWidget*>("access-form-canvas");
    if (!canvas || layout_.width<=0 || layout_.height<=0
        || layout_.width/15.0>canvas->width() || layout_.height/15.0>canvas->height()) return false;
    QMap<QString,QStringList> calculated;
    QStringList growers;
    for (const auto& control:layout_.controls) {
        if (control.type!="Label" && control.type!="TextBox" && control.type!="Rectangle"
            && control.type!="Line") return false;
        if (control.canGrow) {
            if(control.type!="TextBox" || control.section!="Section" || control.controlSource.trimmed().startsWith('=')) return false;
            growers.append(control.name.toCaseFolded());
        }
        if (control.controlSource.trimmed().startsWith('=')) {
            calculated[control.section].append(control.name);
        } else if (!control.controlSource.isEmpty()) {
            auto field=control.controlSource.trimmed();
            if (field.startsWith('[') && field.endsWith(']')) field=field.mid(1,field.size()-2);
            bool found=false;
            for (const auto& column:dataColumns_)
                if (column.compare(field,Qt::CaseInsensitive)==0) { found=true; break; }
            if (!found) return false;
        }
        const auto section=layout_.sectionHeights.constFind(control.section);
        if (section==layout_.sectionHeights.cend()) return false;
        if (const auto* widget=controls_.value(control.name.toCaseFolded()).data()) {
            const int top=qRound(layout_.sectionOffsets.value(control.section)/15.0);
            const int bottom=qRound((layout_.sectionOffsets.value(control.section)+section.value())/15.0);
            if (widget->x()<0 || widget->geometry().right()>=qRound(layout_.width/15.0)
                || widget->y()<top || widget->geometry().bottom()>=bottom) return false;
        } else return false;
    }
    const auto height=[&](const QString& section) {
        return qRound(layout_.sectionHeights.value(section)/15.0);
    };
    const int header=height("PageHeader"),footer=height("PageFooter"),detail=height("Section");
    const double scaleX=printer.logicalDpiX()/96.0,scaleY=printer.logicalDpiY()/96.0;
    const auto page=printer.pageRect(QPrinter::DevicePixel);
    const int pageHeight=qFloor(page.height()/scaleY);
    if (layout_.width/15.0>page.width()/scaleX || pageHeight<=header+footer
        || (!recordRows_.isEmpty() && (detail<=0 || detail>3500))
        || height("FormHeader")>pageHeight-header-footer
        || height("FormFooter")>pageHeight-header-footer) return false;
    if (!growers.isEmpty() && layout_.sectionCanGrow.contains("Section") && !layout_.sectionCanGrow.value("Section")) return false;
    struct SavedControl { QPointer<QWidget> widget; QRect geometry; bool hidden; };
    QMap<QString,SavedControl> saved;
    for (const auto& control:layout_.controls) {
        auto* widget=controls_.value(control.name.toCaseFolded()).data();
        saved.insert(control.name.toCaseFolded(),{widget,widget->geometry(),widget->isHidden()});
    }
    const auto resetControls=[&] {
        for (const auto& state:saved) if(state.widget) {
            state.widget->setGeometry(state.geometry); state.widget->setVisible(!state.hidden);
        }
    };
    QScopedValueRollback<bool> caching(reportAggregateCaching_,true);
    QScopedValueRollback<bool> loading(loading_,true);
    reportAggregateValues_.clear();
    const auto clearAggregates=qScopeGuard([&] { reportAggregateValues_.clear(); });
    const int savedRecord=currentRecordIndex_,savedPage=reportPrintPage_,savedPages=reportPrintPages_;
    const auto canvasSize=canvas->size();
    const auto oldBand=canvas->property("la_print_band_rect"),oldColor=canvas->property("la_print_band_color");
    const auto restore=qScopeGuard([&] {
        resetControls(); canvas->setFixedSize(canvasSize);
        canvas->setProperty("la_print_band_rect",oldBand); canvas->setProperty("la_print_band_color",oldColor);
        currentRecordIndex_=savedRecord; reportPrintPage_=savedPage; reportPrintPages_=savedPages;
        loadRecord();
    });
    struct Growth { QString name; int delta; };
    struct GrowthGroup { int top,bottom,delta; };
    QVector<int> detailHeights(recordRows_.size(),detail);
    QVector<QVector<Growth>> growth(recordRows_.size());
    QVector<QVector<GrowthGroup>> growthGroups(recordRows_.size());
    if (!growers.isEmpty()) {
        reportPrintPage_=1; reportPrintPages_=1;
        for (int row=0;row<recordRows_.size();++row) {
            resetControls(); currentRecordIndex_=row+1; loadRecord();
            for (const auto& name:growers) {
                const auto state=saved.value(name); if(state.hidden) continue;
                const int wanted=state.widget->sizeHint().height();
                const int delta=qMax(0,wanted-state.geometry.height());
                if (!delta) continue;
                if (delta>3500-detail) return false;
                for (const auto& other:layout_.controls) {
                    const auto key=other.name.toCaseFolded(); const auto peer=saved.value(key);
                    if(other.section=="Section" && key!=name && !peer.hidden && peer.geometry.intersects(state.geometry)) return false;
                }
                bool grouped=false;
                for (auto& group:growthGroups[row]) {
                    if(group.top==state.geometry.top() && group.bottom==state.geometry.bottom()) {
                        group.delta=qMax(group.delta,delta); grouped=true; break;
                    }
                    // Aligned parallel columns or disjoint vertical groups
                    // have known flow rules; staggered spans remain explicit.
                    if(state.geometry.top()<=group.bottom && state.geometry.bottom()>=group.top) return false;
                }
                if(!grouped) growthGroups[row].append({state.geometry.top(),state.geometry.bottom(),delta});
                growth[row].append({name,delta});
            }
            int extra=0;
            for (const auto& group:growthGroups[row]) {
                if(group.delta>3500-detail-extra) return false;
                extra+=group.delta;
            }
            detailHeights[row]=detail+extra;
        }
    }
    const auto applyGrowth=[&](int row) {
        for(const auto& control:layout_.controls) {
            if(control.section!="Section") continue;
            const auto key=control.name.toCaseFolded(); const auto state=saved.value(key);
            auto rectangle=state.geometry; int shift=0;
            for(const auto& group:growthGroups[row])
                if(state.geometry.top()>group.bottom) shift+=group.delta;
            rectangle.translate(0,shift);
            for(const auto& change:growth[row])
                if(change.name==key) rectangle.setHeight(rectangle.height()+change.delta);
            state.widget->setGeometry(rectangle);
        }
    };
    struct Band { QString section; int y,record,page,height,slice; QPicture commands; };
    QVector<Band> bands;
    int pageNumber=1;
    int record=recordRows_.isEmpty() ? 0 : 1;
    const auto append=[&](const QString& section,int y,int row) {
        const int size=section=="Section" && row>0 ? detailHeights[row-1] : height(section);
        if (size) bands.append({section,y,row,pageNumber,size,0,{}});
    };
    append("FormHeader",0,record);
    append("PageHeader",height("FormHeader"),record);
    int y=height("FormHeader")+header;
    const auto finishPage=[&] { append("PageFooter",pageHeight-footer,record); };
    const auto startPage=[&](int row) {
        ++pageNumber; record=row; append("PageHeader",0,record); y=header;
    };
    for (int row=0;row<recordRows_.size();++row) {
        const int size=detailHeights[row];
        if (size<=pageHeight-header-footer && y+size>pageHeight-footer) { finishPage(); startPage(row+1); }
        record=row+1;
        if(size<=pageHeight-header-footer) { append("Section",y,record); y+=size; continue; }
        resetControls(); applyGrowth(row); currentRecordIndex_=record; loadRecord();
        QVector<QPair<int,int>> forbidden;
        const int origin=qRound(layout_.sectionOffsets.value("Section")/15.0);
        for(const auto& control:layout_.controls) {
            if(control.section!="Section") continue;
            const auto state=saved.value(control.name.toCaseFolded());
            if(state.hidden) continue;
            if(const auto* text=control.canGrow ? dynamic_cast<const NativeReportTextLayout*>(state.widget.data()) : nullptr) {
                for(const auto& line:text->reportTextLines()) {
                    const int top=qFloor(state.widget->y()-origin+line.top());
                    const int bottom=qCeil(state.widget->y()-origin+line.bottom());
                    // Fixed-height text remains clipped as designed.
                    const int limit=state.widget->geometry().bottom()+1-origin;
                    if(top<limit) forbidden.append({top,qMin(bottom,limit)});
                }
            } else forbidden.append({state.widget->y()-origin,state.widget->geometry().bottom()+1-origin});
        }
        int slice=0;
        while(slice<size) {
            const int available=pageHeight-footer-y;
            int end=qMin(size,slice+available);
            // Moving a cut upward can enter another column's line. Iterate
            // until the boundary lies outside every native text line/control.
            bool moved=true;
            while(moved) {
                moved=false;
                for(const auto& span:forbidden) if(end>span.first && end<span.second) {
                    end=span.first; moved=true;
                }
            }
            if(end<=slice) {
                if(y>header) { finishPage(); startPage(record); continue; }
                return false;
            }
            bands.append({"Section",y,record,pageNumber,end-slice,slice,{}});
            y+=end-slice; slice=end;
            if(slice<size) { finishPage(); startPage(record); }
        }
    }
    if (y+height("FormFooter")>pageHeight-footer) { finishPage(); startPage(record); }
    append("FormFooter",y,record); finishPage();
    if (pageNumber>32767) return false;
    reportPrintPages_=pageNumber;
    for (auto& band:bands) {
        resetControls();
        if(band.section=="Section" && band.record>0) applyGrowth(band.record-1);
        currentRecordIndex_=band.record; reportPrintPage_=band.page; loadRecord();
        for (const auto& name:calculated.value(band.section)) {
            const auto* widget=controls_.value(name.toCaseFolded()).data();
            if (!widget || widget->property("la_calculation_error").toInt()!=0) return false;
        }
        const int offset=qRound(layout_.sectionOffsets.value(band.section)/15.0)+band.slice;
        const QRect fragment(0,offset,canvas->width(),band.height);
        QVector<NativeReportTextLayout*> clipped;
        const auto clearClips=qScopeGuard([&] { for(auto* text:clipped) text->setReportTextClip({}); });
        for (const auto& control:layout_.controls) {
            const auto state=saved.value(control.name.toCaseFolded());
            const bool visible=!state.hidden && control.section==band.section && state.widget->geometry().intersects(fragment);
            state.widget->setVisible(visible);
            if(visible) if(auto* text=dynamic_cast<NativeReportTextLayout*>(state.widget.data())) {
                text->setReportTextClip(fragment.translated(-state.widget->pos())); clipped.append(text);
            }
        }
        if(offset+band.height>3500) return false;
        canvas->setFixedHeight(qMax(canvasSize.height(),offset+band.height));
        canvas->setProperty("la_print_band_rect",QRect(0,offset,qRound(layout_.width/15.0),band.height));
        canvas->setProperty("la_print_band_color",layout_.sectionBackColors.contains(band.section)
            ? resolveOleColor(layout_.sectionBackColors.value(band.section)) : QColor(Qt::white));
        QPainter recording(&band.commands);
        recording.setClipRect(QRect(0,0,canvas->width(),band.height));
        canvas->render(&recording,QPoint(),QRegion(0,offset,canvas->width(),band.height));
        if (!recording.end()) return false;
    }

    QPainter painter;
    if (!painter.begin(&printer)) return false;
    painter.scale(scaleX,scaleY);
    int outputPage=1;
    for (auto& band:bands) {
        while (outputPage<band.page) {
            if (!printer.newPage()) { painter.end(); return false; }
            ++outputPage;
        }
        painter.save(); painter.translate(0,band.y);
        painter.setClipRect(QRect(0,0,canvas->width(),band.height));
        // QPicture playback also converts device DPI; the physical TWIP scale
        // is already active and must not be applied a second time.
        painter.scale(double(band.commands.logicalDpiX())/printer.logicalDpiX(),
                      double(band.commands.logicalDpiY())/printer.logicalDpiY());
        if (!band.commands.play(&painter)) { painter.restore(); painter.end(); return false; }
        painter.restore();
    }
    while (outputPage<pageNumber) {
        if (!printer.newPage()) { painter.end(); return false; }
        ++outputPage;
    }
    return painter.end();
}

bool AccessFormPreview::exportLayoutPdf(const QString& path) { return writePdf(path,false); }
bool AccessFormPreview::exportReportPdf(const QString& path) { return writePdf(path,true); }
bool AccessFormPreview::writePdf(const QString& path,bool report) {
    if(openCancelled_ || closed_) return false;
    if (QFileInfo(path).suffix().compare("pdf",Qt::CaseInsensitive)!=0) return false;
    QTemporaryDir temporary;
    if (!temporary.isValid()) return false;
    const auto generated = temporary.filePath("layout.pdf");
    {
        QPrinter printer(QPrinter::HighResolution);
        printer.setOutputFormat(QPrinter::PdfFormat);
        printer.setOutputFileName(generated);
        if (!(report ? renderReport(printer) : renderLayout(printer))) return false;
    }
    QFile source(generated);
    QSaveFile target(path);
    if (!source.open(QIODevice::ReadOnly) || !target.open(QIODevice::WriteOnly)) return false;
    while (!source.atEnd()) {
        const auto bytes = source.read(65536);
        if (bytes.isEmpty() && source.error()!=QFile::NoError) return false;
        if (target.write(bytes)!=bytes.size()) return false;
    }
    return target.commit();
}
