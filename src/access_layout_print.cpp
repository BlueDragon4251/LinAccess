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
bool AccessFormPreview::exportLayoutPdf(const QString& path) {
    if(openCancelled_ || closed_) return false;
    if (QFileInfo(path).suffix().compare("pdf",Qt::CaseInsensitive)!=0) return false;
    QTemporaryDir temporary;
    if (!temporary.isValid()) return false;
    const auto generated = temporary.filePath("layout.pdf");
    {
        QPrinter printer(QPrinter::HighResolution);
        printer.setOutputFormat(QPrinter::PdfFormat);
        printer.setOutputFileName(generated);
        if (!renderLayout(printer)) return false;
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
