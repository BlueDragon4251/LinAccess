#include "native_macro_message_box.hpp"
#include <QApplication>
#include <QMessageBox>
#include <QPointer>
namespace linaccess {
void nativeMacroMessageBox(QWidget* owner,const std::vector<VbaValue>& args,
        const std::function<VbaValue(const QString&)>& evaluate,const std::function<void()>& check) {
    check();const QPointer<QWidget> alive(owner);
    auto message=QString::fromStdString(args.at(0).asText());
    if(message.startsWith('=')) {message=QString::fromStdString(coerceVbaModuleValue(evaluate(message.mid(1)),"string").asText());check();}
    if(!alive) throw VbaRuntimeError(3420,"MessageBox owner is closed");
    if(message.size()>262144 || !message.isValidUtf16() || message.contains(QChar(0))) throw VbaRuntimeError(5,"MessageBox result text is invalid or too large");
    const auto type=args.size()>2 ? QString::fromStdString(args[2].asText()).toCaseFolded() : QString{};
    QMessageBox box(owner);box.setObjectName("native-macro-message-box");box.setWindowTitle(args.size()>3 && !args[3].asText().empty() ? QString::fromStdString(args[3].asText()) : QString("LinAccess"));
    box.setIcon(type=="critical" ? QMessageBox::Critical : type=="warning?" ? QMessageBox::Question : type=="warning!" ? QMessageBox::Warning : type=="information" ? QMessageBox::Information : QMessageBox::NoIcon);
    const auto sections=message.split('@');
    if(sections.size()>1) {box.setTextFormat(Qt::RichText);QString text="<b>"+sections[0].toHtmlEscaped().replace('\n',"<br>")+"</b><br>"+sections[1].toHtmlEscaped().replace('\n',"<br>");if(sections.size()>2) text+="<br><br>"+sections.mid(2).join('@').toHtmlEscaped().replace('\n',"<br>");box.setText(text);}
    else {box.setTextFormat(Qt::PlainText);box.setText(message);}
    box.setStandardButtons(QMessageBox::Ok);box.setDefaultButton(QMessageBox::Ok);
    if(args.size()<2 || args[1].boolean) QApplication::beep();
    box.exec();check();
}
}
