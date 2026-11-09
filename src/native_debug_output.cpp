#include "native_debug_output.hpp"
#include "vba_debug_output.hpp"
#include <QFontDatabase>
#include <QLabel>
#include <QPlainTextEdit>
#include <QPushButton>
#include <QScrollBar>
#include <QShowEvent>
#include <QTextCursor>
#include <QTimer>
#include <QVBoxLayout>
#include <limits>
namespace {
class OutputPanel final:public QWidget {
public:
    OutputPanel(QWidget* parent,const QString& prefix):QWidget(parent){setObjectName(prefix);auto* layout=new QVBoxLayout(this);text_=new QPlainTextEdit(this);text_->setObjectName(prefix+"-text");text_->setReadOnly(true);text_->setUndoRedoEnabled(false);text_->setMaximumBlockCount(8193);text_->setLineWrapMode(QPlainTextEdit::NoWrap);text_->setFont(QFontDatabase::systemFont(QFontDatabase::FixedFont));layout->addWidget(text_,1);auto* row=new QHBoxLayout;status_=new QLabel(this);status_->setObjectName(prefix+"-status");status_->setTextFormat(Qt::PlainText);row->addWidget(status_,1);auto* clear=new QPushButton("Ausgabe leeren",this);clear->setObjectName(prefix+"-clear");row->addWidget(clear);layout->addLayout(row);
        QObject::connect(clear,&QPushButton::clicked,this,[this]{linaccess::clearVbaDebugOutput();refresh();});auto* timer=new QTimer(this);timer->setInterval(100);QObject::connect(timer,&QTimer::timeout,this,[this]{refresh();});timer->start();}
protected:void showEvent(QShowEvent* event)override{QWidget::showEvent(event);refresh();}
private:
    void refresh(){if(!isVisible() || revision_==linaccess::vbaDebugOutputRevision())return;const auto snapshot=linaccess::vbaDebugOutputSnapshot();const auto value=QString::fromUtf8(snapshot.text.data(),static_cast<qsizetype>(snapshot.text.size()));const auto previous=text_->toPlainText();auto* scroll=text_->verticalScrollBar();const bool following=scroll->value()==scroll->maximum();const auto position=scroll->value();
        if(value.startsWith(previous)){QTextCursor insertion(text_->document());insertion.movePosition(QTextCursor::End);insertion.insertText(value.mid(previous.size()));}else text_->setPlainText(value);
        status_->setText(snapshot.truncated ? "Ältere Ausgabe begrenzt (1 MiB / 8192 Zeilenumbrüche)." : "Debug.Print — nur Arbeitsspeicher");scroll->setValue(following ? scroll->maximum() : position);revision_=snapshot.revision;}
    QPlainTextEdit* text_{};QLabel* status_{};std::uint64_t revision_{std::numeric_limits<std::uint64_t>::max()};
};
}
QWidget* createNativeDebugOutputView(QWidget* parent,const QString& prefix){return new OutputPanel(parent,prefix);}
