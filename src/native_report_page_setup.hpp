#pragma once
#include <QDialog>
#include <QDialogButtonBox>
#include <QComboBox>
#include <QCheckBox>
#include <QDoubleSpinBox>
#include <QFormLayout>
#include <QGroupBox>
#include <QMessageBox>
#include <QPageLayout>
#include <QPageSize>
#include <QPushButton>
#include <QVBoxLayout>
#include <QtMath>

// Actual document page geometry; works with PDF and physical QPrinter backends.
class NativeReportPageSetupDialog final : public QDialog {
public:
    NativeReportPageSetupDialog(QPageLayout initial,QWidget* parent,bool canSave=false)
        : QDialog(parent),layout_(std::move(initial)) {
        setObjectName("access-report-page-setup-dialog"); setWindowTitle("Seite einrichten");
        layout_.setUnits(QPageLayout::Millimeter);
        auto* outer=new QVBoxLayout(this); auto* form=new QFormLayout;
        paper_=new QComboBox(this); paper_->setObjectName("report-page-paper");
        for(int id=0;id<=QPageSize::LastPageSize;++id) {
            if(id==QPageSize::Custom) continue;
            const QPageSize size(static_cast<QPageSize::PageSizeId>(id));
            if(size.isValid()) paper_->addItem(size.name(),id);
        }
        paper_->addItem("Benutzerdefiniert",QPageSize::Custom);
        paper_->setCurrentIndex(paper_->findData(layout_.pageSize().isValid() ? layout_.pageSize().id() : QPageSize::A4));
        form->addRow("Papierformat",paper_);
        width_=dimension(this,"report-page-width"); height_=dimension(this,"report-page-height");
        const auto size=layout_.pageSize().isValid() ? layout_.pageSize().size(QPageSize::Millimeter) : QSizeF(210,297);
        width_->setValue(size.width()); height_->setValue(size.height());
        form->addRow("Papierbreite",width_); form->addRow("Papierhöhe",height_);
        orientation_=new QComboBox(this); orientation_->setObjectName("report-page-orientation");
        orientation_->addItems({"Hochformat","Querformat"}); orientation_->setCurrentIndex(layout_.orientation()==QPageLayout::Landscape ? 1 : 0);
        form->addRow("Ausrichtung",orientation_); outer->addLayout(form);
        const auto updatePaper=[this] {
            const bool custom=paper_->currentData().toInt()==QPageSize::Custom;
            width_->setEnabled(custom); height_->setEnabled(custom);
            if(!custom) {
                const QPageSize page(static_cast<QPageSize::PageSizeId>(paper_->currentData().toInt()));
                const auto size=page.size(QPageSize::Millimeter); width_->setValue(size.width()); height_->setValue(size.height());
            }
        };
        connect(paper_,&QComboBox::currentIndexChanged,this,[updatePaper] { updatePaper(); }); updatePaper();
        auto* group=new QGroupBox("Ränder",this); auto* margins=new QFormLayout(group);
        const auto current=layout_.margins(),minimum=layout_.minimumMargins();
        const auto addMargin=[&](const QString& label,const QString& name,double value,double lower) {
            auto* spin=dimension(group,name); spin->setMinimum(qMax(0.0,qCeil(lower*100)/100.0)); spin->setValue(value);
            margins->addRow(label,spin); return spin;
        };
        left_=addMargin("Links","report-page-left",current.left(),minimum.left());
        right_=addMargin("Rechts","report-page-right",current.right(),minimum.right());
        top_=addMargin("Oben","report-page-top",current.top(),minimum.top());
        bottom_=addMargin("Unten","report-page-bottom",current.bottom(),minimum.bottom());
        outer->addWidget(group);
        save_=new QCheckBox("Im Berichtsentwurf speichern",this); save_->setObjectName("report-page-save");
        save_->setEnabled(canSave);
        if(!canSave) save_->setToolTip("Für gespeicherte, in LinAccess angelegte Berichte verfügbar.");
        outer->addWidget(save_);
        auto* buttons=new QDialogButtonBox(QDialogButtonBox::Ok|QDialogButtonBox::Cancel,this);
        buttons->setObjectName("report-page-buttons"); buttons->button(QDialogButtonBox::Cancel)->setText("Abbrechen");
        connect(buttons,&QDialogButtonBox::accepted,this,&NativeReportPageSetupDialog::accept);
        connect(buttons,&QDialogButtonBox::rejected,this,&QDialog::reject); outer->addWidget(buttons);
    }
    QPageLayout pageLayout() const { return layout_; }
    bool saveToDesign() const { return save_->isEnabled() && save_->isChecked(); }
    void accept() override {
        auto candidate=layout_;
        const auto id=static_cast<QPageSize::PageSizeId>(paper_->currentData().toInt());
        const auto paper=id==QPageSize::Custom
            ? QPageSize(QSizeF(width_->value(),height_->value()),QPageSize::Millimeter,"Benutzerdefiniert",QPageSize::ExactMatch)
            : QPageSize(id);
        candidate.setPageSize(paper,candidate.minimumMargins());
        candidate.setOrientation(orientation_->currentIndex()==1 ? QPageLayout::Landscape : QPageLayout::Portrait);
        if(!paper.isValid() || !candidate.setMargins(QMarginsF(left_->value(),top_->value(),right_->value(),bottom_->value()))
            || !candidate.isValid() || candidate.paintRect(QPageLayout::Millimeter).isEmpty()) {
            QMessageBox::warning(this,"Seite einrichten","Papierformat und Ränder müssen einen gültigen Druckbereich ergeben."); return;
        }
        layout_=candidate; QDialog::accept();
    }
private:
    static QDoubleSpinBox* dimension(QWidget* parent,const QString& name) {
        auto* spin=new QDoubleSpinBox(parent); spin->setObjectName(name); spin->setDecimals(2);
        spin->setRange(0,5000); spin->setSuffix(" mm"); return spin;
    }
    QPageLayout layout_;
    QCheckBox* save_{};
    QComboBox *paper_{},*orientation_{};
    QDoubleSpinBox *width_{},*height_{},*left_{},*right_{},*top_{},*bottom_{};
};
