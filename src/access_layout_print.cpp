#include "access_form_preview.hpp"
#include "access_format.hpp"
#include "native_report_text_layout.hpp"
#include "native_report_aggregate.hpp"
#include "native_report_grouping.hpp"
#include "vba_qt_value.hpp"
#include <QAction>
#include <QFileDialog>
#include <QFile>
#include <QFileInfo>
#include <QMessageBox>
#include <QPainter>
#include <QPrintDialog>
#include "native_report_page_setup.hpp"
#include "native_report_page_store.hpp"
#include <QPageLayout>
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
#include <QShortcut>
#include <QGraphicsView>
#include <utility>

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
        auto* number=new QSpinBox(this); number->setRange(1,32767); number->setKeyboardTracking(false);
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
        auto* pageNumberKey=new QShortcut(QKeySequence(Qt::Key_F5),this);
        pageNumberKey->setContext(Qt::WidgetWithChildrenShortcut);
        connect(pageNumberKey,&QShortcut::activated,number,[number] { number->setFocus(); number->selectAll(); });
        auto* zoomKey=new QShortcut(QKeySequence(Qt::Key_Z),this);
        zoomKey->setContext(Qt::WidgetWithChildrenShortcut);
        connect(zoomKey,&QShortcut::activated,pages_,[this] {
            if(pages_->zoomMode()==QPrintPreviewWidget::CustomZoom) pages_->setZoomMode(QPrintPreviewWidget::FitInView);
            else pages_->setZoomFactor(1.0);
        });
        if(auto* surface=pages_->findChild<QGraphicsView*>()) {
            const std::pair<Qt::Key,int> commands[]{{Qt::Key_PageDown,1},{Qt::Key_Down,1},{Qt::Key_PageUp,-1},{Qt::Key_Up,-1}};
            for(const auto& [key,direction]:commands) {
                auto* shortcut=new QShortcut(QKeySequence(key),surface);
                shortcut->setContext(Qt::WidgetWithChildrenShortcut);
                shortcut->setEnabled(pages_->zoomMode()==QPrintPreviewWidget::FitInView);
                pageKeys_.append(shortcut);
                connect(shortcut,&QShortcut::activated,pages_,[this,direction] {
                    pages_->setCurrentPage(qBound(1,pages_->currentPage()+direction,qMax(1,pages_->pageCount())));
                });
            }
        }
        connect(pages_,&QPrintPreviewWidget::paintRequested,this,[this,render](QPrinter* printer) {
            setProperty("la_report_pages_valid",false);
            if(property("la_page_setup_error").toString().isEmpty()) setProperty("la_report_pages_valid",render(*printer));
        });
        connect(pages_,&QPrintPreviewWidget::previewChanged,this,[this,previous,next,number,total,changed] {
            const int count=pages_->pageCount(),page=pages_->currentPage();
            for(const auto& key:pageKeys_) if(key) key->setEnabled(pages_->zoomMode()==QPrintPreviewWidget::FitInView);
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
    QPageLayout pageLayout() const { return printer_.pageLayout(); }
    bool setPageLayout(const QPageLayout& layout) { return printer_.setPageLayout(layout); }
    bool setupPage(bool canSave,const std::function<void(const QPageLayout&)>& save) {
        const auto previous=printer_.pageLayout();
        NativeReportPageSetupDialog dialog(previous,this,canSave);
        if(dialog.exec()!=QDialog::Accepted) return false;
        if(!printer_.setPageLayout(dialog.pageLayout())) {
            printer_.setPageLayout(previous);
            QMessageBox::warning(this,"Seite einrichten","Der Drucker kann diese Seiteneinrichtung nicht übernehmen."); return false;
        }
        try { if(dialog.saveToDesign()) save(dialog.pageLayout()); }
        catch(...) { printer_.setPageLayout(previous); throw; }
        setProperty("la_page_setup_error",QString());
        refreshPages(); return true;
    }
    bool refreshPages() {
        setProperty("la_report_pages_valid",false);
        pages_->updatePreview();
        return property("la_report_pages_valid").toBool() && pages_->pageCount()>0;
    }
private:
    QPrinter printer_;
    QPrintPreviewWidget* pages_{};
    QVector<QPointer<QShortcut>> pageKeys_;
};
bool applyReportPageLayout(QWidget* view,QPrinter& printer) {
    const auto* pages=dynamic_cast<NativeReportPageView*>(view);
    return !pages || (pages->property("la_page_setup_error").toString().isEmpty() && printer.setPageLayout(pages->pageLayout()));
}
void synchronizeReportView(QWidget* document,NativeReportPageView* pages) {
    auto* tabs=document->findChild<QTabWidget*>("access-report-view-tabs");
    auto* notice=document->findChild<QLabel*>("access-report-view-notice");
    auto* view=pages ? pages->findChild<QPrintPreviewWidget*>("access-report-pages") : nullptr;
    if (!tabs || !notice || !view) return;
    const bool valid=pages->property("la_report_pages_valid").toBool() && view->pageCount()>0;
    const auto pageError=pages->property("la_page_setup_error").toString();
    notice->setText(pageError.isEmpty()
        ? QString("Nur Layoutansicht: Die vollständige Berichtsausgabe ist für dieses Layout noch nicht unterstützt.")
        : pageError+" Öffnen Sie Seite einrichten, um gültige Einstellungen zu wählen.");
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
    try {
        if(const auto saved=linaccess::report_page_store::load(db_,formName_);saved && !pages->setPageLayout(*saved))
            pages->setProperty("la_page_setup_error","Die gespeicherte Seiteneinrichtung kann nicht übernommen werden.");
    } catch(const std::exception&) {
        pages->setProperty("la_page_setup_error","Die gespeicherte Seiteneinrichtung konnte nicht geladen werden.");
    }
    auto* printKey=new QShortcut(QKeySequence(Qt::Key_P),pages);
    printKey->setContext(Qt::WidgetWithChildrenShortcut);
    connect(printKey,&QShortcut::activated,this,[this] {
        if(auto* action=findChild<QAction*>("access-report-print");action && action->isVisible() && action->isEnabled()) action->trigger();
    });
    auto* setupKey=new QShortcut(QKeySequence(Qt::Key_S),pages);
    setupKey->setContext(Qt::WidgetWithChildrenShortcut);
    connect(setupKey,&QShortcut::activated,this,[this] {
        if(auto* action=findChild<QAction*>("access-report-page-setup");action && action->isVisible() && action->isEnabled()) action->trigger();
    });
    auto* closeKey=new QShortcut(QKeySequence(Qt::Key_C),pages);
    closeKey->setContext(Qt::WidgetWithChildrenShortcut);
    connect(closeKey,&QShortcut::activated,this,&AccessFormPreview::reject);
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
    print->setObjectName("access-layout-print");
    print->setToolTip("Aktuelles Formularlayout drucken (Strg+P)");
    // The form key handler runs Access KeyDown/KeyPreview before this command.
    connect(print,&QAction::triggered,this,[this] { printCurrentDocument(); });
    auto* pdf = printToolbar_->addAction("Aktuelles Layout als PDF …");
    connect(pdf,&QAction::triggered,this,[this] {
        const auto path = QFileDialog::getSaveFileName(this,"Aktuelles Layout als PDF",{},"PDF (*.pdf)");
        if (!path.isEmpty() && !exportLayoutPdf(path))
            QMessageBox::warning(this,"PDF","Das aktuelle Layout konnte nicht gespeichert werden.");
    });
    auto* setup=printToolbar_->addAction("Seite einrichten …");
    setup->setObjectName("access-report-page-setup");
    connect(setup,&QAction::triggered,this,[this] {
        if(reportAggregateCaching_) return;
        auto* pages=dynamic_cast<NativeReportPageView*>(reportPageView_.data());
        try {
            const bool canSave=linaccess::report_page_store::definition(db_,formName_).has_value();
            if(pages && pages->setupPage(canSave,[this](const QPageLayout& layout) { linaccess::report_page_store::save(db_,formName_,layout); })) {
                synchronizeReportView(this,pages); refreshCalculatedControls();
            }
        } catch(const std::exception&) {
            QMessageBox::warning(this,"Seite einrichten","Die Seiteneinrichtung konnte nicht im Berichtsentwurf gespeichert werden.");
        }
    });
    auto* reportPreview=printToolbar_->addAction("Berichtsvorschau …");
    reportPreview->setObjectName("access-report-preview");
    connect(reportPreview,&QAction::triggered,this,[this] {
        QPrinter printer(QPrinter::HighResolution);
        if(!applyReportPageLayout(reportPageView_.data(),printer)) {
            QMessageBox::warning(this,"Berichtsvorschau","Die Seiteneinrichtung kann für diese Ausgabe nicht verwendet werden."); return;
        }
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
    reportPrint->setShortcut(QKeySequence::Print);
    reportPrint->setShortcutContext(Qt::WindowShortcut);
    connect(reportPrint,&QAction::triggered,this,[this] { printCurrentDocument(); });
    auto* reportPdf=printToolbar_->addAction("Bericht als PDF …");
    connect(reportPdf,&QAction::triggered,this,[this] {
        const auto path=QFileDialog::getSaveFileName(this,"Bericht als PDF",{},"PDF (*.pdf)");
        if (!path.isEmpty() && !exportReportPdf(path))
            QMessageBox::warning(this,"Bericht als PDF",
                "Dieser Bericht konnte nicht vollständig ausgegeben werden. "
                "Prüfen Sie die Druckregeln und die Datenquelle.");
    });
    auto* close = printToolbar_->addAction("&Schließen");
    connect(close,&QAction::triggered,this,&AccessFormPreview::reject);

}
void AccessFormPreview::showDocumentPrintDialog() {
    const bool report=layout_.kind==linaccess::AccessDocumentKind::Report;
    QPrinter printer(QPrinter::HighResolution);
    if(report && !applyReportPageLayout(reportPageView_.data(),printer)) {
        QMessageBox::warning(this,"Bericht drucken","Die Seiteneinrichtung kann für diese Ausgabe nicht verwendet werden."); return;
    }
    QPrintDialog dialog(&printer,this);
    dialog.setObjectName(report ? "access-report-print-dialog" : "access-layout-print-dialog");
    dialog.setWindowTitle(report ? "Bericht drucken" : "Aktuelles Formularlayout drucken");
    dialog.setOption(QAbstractPrintDialog::PrintPageRange,report);
    dialog.setOption(QAbstractPrintDialog::PrintCurrentPage,false);
    if(report) if(const auto* pages=findChild<QPrintPreviewWidget*>("access-report-pages")) dialog.setMinMax(1,qMax(1,pages->pageCount()));
    dialog.setOption(QAbstractPrintDialog::PrintSelection,false);
    if(dialog.exec()!=QDialog::Accepted || closed_ || openCancelled_) return;
    const bool printed=printer.outputFormat()==QPrinter::PdfFormat
        ? writePdf(printer.outputFileName(),report,&printer)
        : (printer.outputFileName().isEmpty() && (report ? renderReport(printer) : renderLayout(printer)));
    if(!printed || printer.printerState()==QPrinter::Error)
        QMessageBox::warning(this,"Drucken","Dieses Dokument konnte nicht vollständig ausgegeben werden.");
}
bool AccessFormPreview::renderLayout(QPrinter& printer) {
    for(const auto& widget:controls_) if(widget && widget->property("la_format_error").toInt()) return false;
    for(const auto& control:layout_.controls) if(control.type=="TextBox" && !control.format.isEmpty()) {
        try {const auto actual=getFieldValue(control.name.toStdString());if(actual.kind!=linaccess::VbaValue::Kind::Null && actual.kind!=linaccess::VbaValue::Kind::Error) (void)linaccess::formatAccessValue(actual,control.format.toStdString(),true);}catch(const std::exception&) {return false;}
    }
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
        || dataTruncated_) return false;
    std::optional<linaccess::NativeReportGroupPlan> grouping;
    try {grouping=linaccess::nativeReportGrouping(layout_,dataColumns_,dataTypes_,recordRows_);}catch(const std::exception&) {return false;}
    if(!grouping) return false;
    for(const auto& limitation:layout_.reportPrintLimitations)
        if(grouping->levels.isEmpty() || (limitation!="GroupLevel" && limitation!="BreakLevel" && limitation!="BreakSection")) return false;
    auto* canvas=findChild<QWidget*>("access-form-canvas");
    if (!canvas || layout_.width<=0 || layout_.height<=0
        || layout_.width/15.0>canvas->width() || layout_.height/15.0>canvas->height()) return false;
    int runningGroup=-1;
    for(int level=0;level<grouping->levels.size();++level)
        if(!grouping->levels[level].header.isEmpty() || !grouping->levels[level].footer.isEmpty()) runningGroup=level;
    QMap<QString,QStringList> calculated;
    QStringList growers;
    for (const auto& control:layout_.controls) {
        if (control.type!="Label" && control.type!="TextBox" && control.type!="Rectangle"
            && control.type!="Line") return false;
        if(control.runningSum && ((control.runningSum!=1 && control.runningSum!=2) || (control.runningSum==1 && runningGroup<0) || control.type!="TextBox" || control.section!="Section" ||
            control.canGrow || (control.controlSource.trimmed().startsWith('=') && control.controlSource.trimmed()!="=1"))) return false;
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
    const int detail=height("Section");
    const bool hasReportHeader=height("FormHeader")>0,hasReportFooter=height("FormFooter")>0;
    const bool separateFooter=hasReportFooter && ((layout_.reportPageHeader|layout_.reportPageFooter)&2);
    const bool footerOnlyFirst=separateFooter && !hasReportHeader && recordRows_.isEmpty();
    const auto pageBandHeight=[&](const QString& section,int setting,bool first,bool reportFooter) {
        return ((first && hasReportHeader && (setting&1)) || (reportFooter && (setting&2))) ? 0 : height(section);
    };
    int header=pageBandHeight("PageHeader",layout_.reportPageHeader,true,footerOnlyFirst);
    int footer=pageBandHeight("PageFooter",layout_.reportPageFooter,true,footerOnlyFirst);
    const double scaleX=printer.logicalDpiX()/96.0,scaleY=printer.logicalDpiY()/96.0;
    const auto page=printer.pageRect(QPrinter::DevicePixel);
    const int pageHeight=qFloor(page.height()/scaleY);
    if (layout_.width/15.0>page.width()/scaleX || pageHeight<=header+footer
        || (!recordRows_.isEmpty() && (detail<=0 || detail>3500))
        || height("FormHeader")>pageHeight-header-footer) return false;
    if((layout_.reportHeaderForceNewPage && height("FormHeader")==0)
        || (layout_.reportFooterForceNewPage && height("FormFooter")==0)) return false;
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
    const auto originalRows=recordRows_;const auto originalIds=rowIds_;const auto originalSnapshot=recordsetSnapshot_;
    const auto originalGroupRange=reportGroupRange_;reportGroupRange_.reset();
    const auto canvasSize=canvas->size();
    const auto oldBand=canvas->property("la_print_band_rect"),oldColor=canvas->property("la_print_band_color");
    const auto restore=qScopeGuard([&] {
        resetControls(); canvas->setFixedSize(canvasSize);
        canvas->setProperty("la_print_band_rect",oldBand); canvas->setProperty("la_print_band_color",oldColor);
        reportRunningValues_.clear();
        if(!grouping->levels.isEmpty()) {recordRows_=originalRows;rowIds_=originalIds;recordsetSnapshot_=originalSnapshot;}
        reportGroupRange_=originalGroupRange;if(!grouping->levels.isEmpty()) reportAggregateValues_.clear();
        currentRecordIndex_=savedRecord; reportPrintPage_=savedPage; reportPrintPages_=savedPages;
        loadRecord();
    });
    if(!grouping->levels.isEmpty()) {
        if(!rowIds_.isEmpty() && rowIds_.size()!=recordRows_.size()) return false;
        QVector<QVariantList> sorted;QVector<QVariant> identities;sorted.reserve(recordRows_.size());identities.reserve(rowIds_.size());
        for(const int row:grouping->order) {sorted.append(recordRows_[row]);if(!rowIds_.isEmpty()) identities.append(rowIds_[row]);}
        recordRows_=std::move(sorted);rowIds_=std::move(identities);recordsetSnapshot_.reset();
    }
    // Build each record prefix once. Pagination and repeated preview paints only
    // read these values, so split sections never accumulate the same row twice.
    try {
        qsizetype prefixCells=0;
        for(const auto& control:layout_.controls) {
            if(!control.runningSum) continue;
            if(recordRows_.size()>1000000-prefixCells) return false;
            prefixCells+=recordRows_.size();
            auto field=control.controlSource.trimmed();
            const bool counter=field=="=1";
            if(field.startsWith('[') && field.endsWith(']')) field=field.mid(1,field.size()-2);
            int column=-1;
            for(int i=0;i<dataColumns_.size();++i) if(dataColumns_[i].compare(field,Qt::CaseInsensitive)==0) {column=i;break;}
            if(!counter && column<0) return false;
            auto& values=reportRunningValues_[control.name.toCaseFolded()];
            linaccess::NativeReportAggregate sum("sum");
            for(int position=0;position<recordRows_.size();++position) {
                if(control.runningSum==1 && position==grouping->ranges[runningGroup][position].first)
                    sum=linaccess::NativeReportAggregate("sum");
                const auto& row=recordRows_[position];
                const auto value=counter ? linaccess::VbaValue::integer(1) :
                    linaccess::qtFieldValue(row[column],dataTypes_.value(dataColumns_[column].toCaseFolded()));
                if(value.kind!=linaccess::VbaValue::Kind::Number && value.kind!=linaccess::VbaValue::Kind::Currency && value.kind!=linaccess::VbaValue::Kind::Decimal) return false;
                sum.add(value);values.append(std::make_shared<linaccess::VbaValue>(sum.result()));
            }
        }
    } catch(const std::exception&) { return false; }
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
    struct Band { QString section; int y,record,page,height,slice; QPicture commands;std::optional<QPair<int,int>> group; };
    QVector<Band> bands;
    int pageNumber=1;
    int record=recordRows_.isEmpty() ? 0 : 1;
    const auto append=[&](const QString& section,int y,int row,std::optional<QPair<int,int>> scope={}) {
        const int size=section=="Section" && row>0 ? detailHeights[row-1] : height(section);
        if (size) bands.append({section,y,row,pageNumber,size,0,{},scope});
    };
    append("FormHeader",0,record);
    if(header) append("PageHeader",height("FormHeader"),record);
    int y=height("FormHeader")+header;
    int bodyTop=header;QVector<std::optional<QPair<int,int>>> activeHeaders(grouping->levels.size());
    const auto finishPage=[&] {if(footer) append("PageFooter",pageHeight-footer,record);};
    const auto startPage=[&](int row,bool reportFooter=false) -> bool {
        ++pageNumber;record=row;
        header=pageBandHeight("PageHeader",layout_.reportPageHeader,false,reportFooter);
        footer=pageBandHeight("PageFooter",layout_.reportPageFooter,false,reportFooter);
        if(header) append("PageHeader",0,record);y=header;
        if(pageHeight<=header+footer) return false;
        if(!reportFooter) for(int level=0;level<grouping->levels.size();++level) {
            const auto& group=grouping->levels[level];const auto scope=activeHeaders[level];
            if(!group.repeatHeader || !scope || row<scope->first+1 || row>scope->second) continue;
            const int size=height(group.header);if(y+size>=pageHeight-footer) return false;
            append(group.header,y,scope->first+1,*scope);y+=size;
        }
        bodyTop=y;return true;
    };
    bool pendingBodyBreak=(layout_.reportHeaderForceNewPage&2)!=0;
    const auto appendGroup=[&](const QString& section,int row,QPair<int,int> scope) {
        const int size=height(section);if(section.isEmpty() || size==0) return true;
        if(pendingBodyBreak && y>bodyTop) {finishPage();if(!startPage(row)) return false;}pendingBodyBreak=false;
        if(size>pageHeight-bodyTop-footer) return false;
        if(y+size>pageHeight-footer) {finishPage();if(!startPage(row)) return false;}
        if(pageHeight<=header+footer || size>pageHeight-bodyTop-footer) return false;
        append(section,y,row,scope);y+=size;return true;
    };
    const auto closeGroups=[&](int row) {
        for(int level=grouping->levels.size()-1;level>=0;--level) if(grouping->ranges[level][row].second==row+1) {
            if(!appendGroup(grouping->levels[level].footer,row+1,grouping->ranges[level][row])) return false;
            activeHeaders[level].reset();
        }
        return true;
    };
    for (int row=0;row<recordRows_.size();++row) {
        if(pendingBodyBreak && y>bodyTop) {finishPage();if(!startPage(row+1)) return false;}pendingBodyBreak=false;
        for(int level=0;level<grouping->levels.size();++level) if(grouping->ranges[level][row].first==row) {
            if(!appendGroup(grouping->levels[level].header,row+1,grouping->ranges[level][row])) return false;
            if(!grouping->levels[level].header.isEmpty()) activeHeaders[level]=grouping->ranges[level][row];
        }
        if((layout_.detailForceNewPage&1) && y>bodyTop) {finishPage();if(!startPage(row+1)) return false;}
        const int size=detailHeights[row];
        if(pageHeight<=header+footer) return false;
        if (size<=pageHeight-bodyTop-footer && y+size>pageHeight-footer) { finishPage(); if(!startPage(row+1)) return false; }
        record=row+1;
        if(pageHeight<=header+footer) return false;
        if(size<=pageHeight-bodyTop-footer) {
            append("Section",y,record); y+=size;
            pendingBodyBreak=(layout_.detailForceNewPage&2)!=0;if(!closeGroups(row)) return false;continue;
        }
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
            if(pageHeight<=header+footer) return false;
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
                if(y>bodyTop) { finishPage(); if(!startPage(record)) return false; continue; }
                return false;
            }
            bands.append({"Section",y,record,pageNumber,end-slice,slice,{}});
            y+=end-slice; slice=end;
            if(slice<size) { finishPage(); if(!startPage(record)) return false; }
        }
        pendingBodyBreak=(layout_.detailForceNewPage&2)!=0;if(!closeGroups(row)) return false;
    }
    if((pendingBodyBreak || (layout_.reportFooterForceNewPage&1) || separateFooter) && hasReportFooter && y>bodyTop) {finishPage();if(!startPage(record,true)) return false;}
    if (y+height("FormFooter")>pageHeight-footer) {finishPage();if(!startPage(record,true)) return false;}
    if(pageHeight<=header+footer || y+height("FormFooter")>pageHeight-footer) return false;
    append("FormFooter",y,record);finishPage();
    if (pageNumber>32767) return false;
    reportPrintPages_=pageNumber;
    for (auto& band:bands) {
        resetControls();
        if(band.section=="Section" && band.record>0) applyGrowth(band.record-1);
        reportGroupRange_=band.group;
        currentRecordIndex_=band.record; reportPrintPage_=band.page; loadRecord();
        for (const auto& name:calculated.value(band.section)) {
            const auto* widget=controls_.value(name.toCaseFolded()).data();
            if (!widget || widget->property("la_calculation_error").toInt()!=0) return false;
        }
        for(const auto& control:layout_.controls) if(control.section==band.section) {const auto widget=controls_.value(control.name.toCaseFolded());if(widget && widget->property("la_format_error").toInt()) return false;}
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

    if(printer.printRange()!=QPrinter::AllPages && printer.printRange()!=QPrinter::PageRange) return false;
    const int first=printer.printRange()==QPrinter::PageRange ? printer.fromPage() : 1;
    const int requestedLast=printer.printRange()==QPrinter::PageRange ? printer.toPage() : pageNumber;
    const int last=qMin(pageNumber,requestedLast);
    const int requestedCopies=printer.copyCount();
    if(first<1 || requestedLast<first || first>last || requestedCopies<1 || requestedCopies>100) return false;
    const int copies=printer.outputFormat()==QPrinter::PdfFormat || !printer.supportsMultipleCopies() ? requestedCopies : 1;
    if(copies>1 && static_cast<qint64>(last-first+1)*copies>10000) return false;
    const bool collated=printer.collateCopies(),reverse=printer.pageOrder()==QPrinter::LastPageFirst;
    // Record each source page once; copies never rerun calculations or events.
    QVector<QVector<int>> pageBands(pageNumber+1);
    for(int i=0;i<bands.size();++i) pageBands[bands[i].page].append(i);
    if(copies>1) printer.setCopyCount(1);
    const auto restoreCopies=qScopeGuard([&] {if(copies>1) printer.setCopyCount(requestedCopies);});
    QPainter painter;
    if (!painter.begin(&printer)) return false;
    painter.scale(scaleX,scaleY);
    int printed=0;
    const auto draw=[&](int sourcePage) {
        if(printed++ && !printer.newPage()) return false;
        for(const int index:pageBands[sourcePage]) {
            auto& band=bands[index];
            painter.save(); painter.translate(0,band.y);
            painter.setClipRect(QRect(0,0,canvas->width(),band.height));
            painter.scale(double(band.commands.logicalDpiX())/printer.logicalDpiX(),
                          double(band.commands.logicalDpiY())/printer.logicalDpiY());
            const bool painted=band.commands.play(&painter);painter.restore();if(!painted) return false;
        }
        return true;
    };
    if(collated) {
        for(int copy=0;copy<copies;++copy) for(int offset=0;offset<=last-first;++offset)
            if(!draw(reverse ? last-offset : first+offset)) {painter.end();return false;}
    } else {
        for(int offset=0;offset<=last-first;++offset) for(int copy=0;copy<copies;++copy)
            if(!draw(reverse ? last-offset : first+offset)) {painter.end();return false;}
    }
    return painter.end();
}

bool AccessFormPreview::exportLayoutPdf(const QString& path) { return writePdf(path,false); }
bool AccessFormPreview::exportReportPdf(const QString& path) { return writePdf(path,true); }
bool AccessFormPreview::writePdf(const QString& path,bool report,const QPrinter* settings) {
    if(openCancelled_ || closed_) return false;
    if(path.contains(QChar(0)) || QFileInfo(path).isSymLink() || QFileInfo(path).suffix().compare("pdf",Qt::CaseInsensitive)!=0) return false;
    const auto destination=QFileInfo(path).canonicalFilePath(),database=QFileInfo(db_.databaseName()).canonicalFilePath();
    if(!destination.isEmpty() && !database.isEmpty() && destination==database) return false;
    QTemporaryDir temporary;
    if (!temporary.isValid()) return false;
    const auto generated = temporary.filePath("layout.pdf");
    {
        QPrinter printer(QPrinter::HighResolution);
        printer.setOutputFormat(QPrinter::PdfFormat);
        printer.setOutputFileName(generated);
        if(settings) {
            if(!printer.setPageLayout(settings->pageLayout())) return false;
            printer.setFromTo(settings->fromPage(),settings->toPage());printer.setPrintRange(settings->printRange());
            printer.setPageOrder(settings->pageOrder());printer.setCopyCount(settings->copyCount());printer.setCollateCopies(settings->collateCopies());
        } else if(report && !applyReportPageLayout(reportPageView_.data(),printer)) return false;
        if (!(report ? renderReport(printer) : renderLayout(printer))) return false;
        if(printer.printerState()==QPrinter::Error) return false;
    }
    QFile source(generated);
    QSaveFile target(path); target.setDirectWriteFallback(false);
    if (!source.open(QIODevice::ReadOnly) || !target.open(QIODevice::WriteOnly)) return false;
    while (!source.atEnd()) {
        const auto bytes = source.read(65536);
        if (bytes.isEmpty() && source.error()!=QFile::NoError) return false;
        if (target.write(bytes)!=bytes.size()) return false;
    }
    return target.commit();
}
