#include "access_form_preview.hpp"
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
    if (reportAggregateCaching_ || layout_.kind!=linaccess::AccessDocumentKind::Report || !dataSourceAvailable_
        || dataTruncated_ || !layout_.reportPrintLimitations.isEmpty()) return false;
    auto* canvas=findChild<QWidget*>("access-form-canvas");
    if (!canvas || layout_.width<=0 || layout_.height<=0
        || layout_.width/15.0>canvas->width() || layout_.height/15.0>canvas->height()) return false;
    QStringList calculated;
    for (const auto& control:layout_.controls) {
        if (control.type!="Label" && control.type!="TextBox" && control.type!="Rectangle"
            && control.type!="Line") return false;
        if (control.controlSource.trimmed().startsWith('=')) {
            calculated.append(control.name);
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
        || (!recordRows_.isEmpty() && (detail<=0 || detail>pageHeight-header-footer))
        || height("FormHeader")>pageHeight-header-footer
        || height("FormFooter")>pageHeight-header-footer) return false;
    QScopedValueRollback<bool> caching(reportAggregateCaching_,true);
    reportAggregateValues_.clear();
    const auto clearAggregates=qScopeGuard([&] { reportAggregateValues_.clear(); });
    const int savedRecord=currentRecordIndex_;
    const auto restore=qScopeGuard([&] { currentRecordIndex_=savedRecord; loadRecord(); });
    // Validate every row before opening the output device: an unsupported
    // calculated field must not produce an apparently complete partial report.
    if (!calculated.isEmpty()) {
        for (int row=0;row<qMax(1,int(recordRows_.size()));++row) {
            currentRecordIndex_=recordRows_.isEmpty() ? 0 : row+1;
            loadRecord();
            for (const auto& name:calculated) {
                const auto* widget=controls_.value(name.toCaseFolded()).data();
                if (!widget || widget->property("la_calculation_error").toInt()!=0) return false;
            }
        }
    }
    QPainter painter;
    if (!painter.begin(&printer)) return false;
    painter.scale(scaleX,scaleY);
    const auto band=[&](const QString& section,int y) {
        const int size=height(section);
        if (!size) return;
        painter.save(); painter.translate(0,y);
        painter.setClipRect(QRect(0,0,canvas->width(),size));
        QPicture commands;
        QPainter recording(&commands);
        const int offset=qRound(layout_.sectionOffsets.value(section)/15.0);
        canvas->render(&recording,QPoint(),QRegion(0,offset,canvas->width(),size));
        recording.end();
        // QPicture playback applies its own device-DPI ratio. Compensate it
        // because the report's physical TWIP scale is already installed above.
        painter.scale(double(commands.logicalDpiX())/printer.logicalDpiX(),
                      double(commands.logicalDpiY())/printer.logicalDpiY());
        commands.play(&painter);
        painter.restore();
    };
    const auto pageStart=[&] { band("PageHeader",0); };
    const auto pageEnd=[&] { band("PageFooter",pageHeight-footer); };
    const auto nextPage=[&](int record) {
        pageEnd();
        if (!printer.newPage()) return false;
        currentRecordIndex_=record; loadRecord();
        pageStart(); return true;
    };
    if (!recordRows_.isEmpty()) { currentRecordIndex_=1; loadRecord(); }
    band("FormHeader",0);
    band("PageHeader",height("FormHeader"));
    int y=height("FormHeader")+header;
    for (int row=0;row<recordRows_.size();++row) {
        if (y+detail>pageHeight-footer) {
            if (!nextPage(row+1)) { painter.end(); return false; }
            y=header;
        }
        currentRecordIndex_=row+1; loadRecord();
        band("Section",y); y+=detail;
    }
    if (y+height("FormFooter")>pageHeight-footer) {
        if (!nextPage(currentRecordIndex_)) { painter.end(); return false; }
        y=header;
    }
    band("FormFooter",y); pageEnd();
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
