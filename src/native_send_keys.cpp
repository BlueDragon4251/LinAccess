#include "native_send_keys.hpp"
#include "vba_runtime.hpp"
#include <QApplication>
#include <QKeyEvent>
#include <QPointer>
#include <QWidget>
#include <QMap>
#include <QVector>

namespace linaccess {
namespace {
struct Key { int code; QString text; Qt::KeyboardModifiers modifiers; };
class KeyParser {
public:
    explicit KeyParser(QString source) : source_(std::move(source)) {}
    QVector<Key> parse() { group({},false,0); return result_; }
private:
    void append(int code,QString text,Qt::KeyboardModifiers modifiers,int count=1) {
        if (count<1) throw VbaRuntimeError(5,"Invalid SendKeys repeat count");
        if (count>1000 || result_.size()+count>4096) throw VbaRuntimeError(3035,"SendKeys event budget exceeded");
        for (int i=0;i<count;++i) result_.append({code,text,modifiers});
    }
    void literal(QString value,Qt::KeyboardModifiers modifiers) {
        if (value.size()==1 && (value[0]=='\r' || value[0]=='\n')) { append(Qt::Key_Return,"\r",modifiers); return; }
        if (value=="\t") { append(Qt::Key_Tab,"\t",modifiers); return; }
        if (value==QString(QChar(27))) { append(Qt::Key_Escape,{},modifiers); return; }
        if (value==QString(QChar(8))) { append(Qt::Key_Backspace,{},modifiers); return; }
        const bool ascii=value.size()==1 && value[0].unicode()>=32 && value[0].unicode()<127;
        if ((modifiers & (Qt::ControlModifier|Qt::AltModifier)) && !ascii)
            throw VbaRuntimeError(3251,"Non-ASCII shortcut layout unsupported");
        if (modifiers & Qt::ShiftModifier) {
            if (!ascii || (!value[0].isLetter() && value!=" ")) throw VbaRuntimeError(3251,"Shifted character layout unsupported");
            value=value.toUpper();
        }
        if (value.size()==1 && value[0].unicode()<32) throw VbaRuntimeError(3251,"Control character unsupported");
        const int code=ascii ? value.toUpper()[0].unicode() : Qt::Key_unknown;
        append(code,value,modifiers);
    }
    void brace(Qt::KeyboardModifiers modifiers) {
        ++position_;
        if (position_+1<source_.size() && (source_[position_]=='{' || source_[position_]=='}') && source_[position_+1]=='}') {
            literal(source_.mid(position_,1),modifiers); position_+=2; return;
        }
        const auto end=source_.indexOf('}',position_);
        if (end<0) throw VbaRuntimeError(5,"Unclosed SendKeys key token");
        const auto token=source_.mid(position_,end-position_); position_=end+1;
        if (token.size()==1 && QString("+^%~()[]{}").contains(token)) { literal(token,modifiers); return; }
        const auto parts=token.simplified().split(' ');
        if (parts.empty() || parts.size()>2) throw VbaRuntimeError(5,"Invalid SendKeys key token");
        int count=1;
        if (parts.size()==2) { bool valid=false; count=parts[1].toInt(&valid); if (!valid) throw VbaRuntimeError(5,"Invalid SendKeys repeat count"); }
        const auto name=parts[0].toUpper();
        static const QMap<QString,int> named{{"ENTER",Qt::Key_Return},{"RETURN",Qt::Key_Return},{"TAB",Qt::Key_Tab},
            {"ESC",Qt::Key_Escape},{"ESCAPE",Qt::Key_Escape},{"UP",Qt::Key_Up},{"DOWN",Qt::Key_Down},
            {"LEFT",Qt::Key_Left},{"RIGHT",Qt::Key_Right},{"HOME",Qt::Key_Home},{"END",Qt::Key_End},
            {"PGUP",Qt::Key_PageUp},{"PGDN",Qt::Key_PageDown},{"BS",Qt::Key_Backspace},
            {"BACKSPACE",Qt::Key_Backspace},{"BKSP",Qt::Key_Backspace},{"DEL",Qt::Key_Delete},
            {"DELETE",Qt::Key_Delete},{"INS",Qt::Key_Insert},{"INSERT",Qt::Key_Insert},{"SPACE",Qt::Key_Space}};
        int code=0;
        if (named.contains(name)) code=named.value(name);
        else if (name.startsWith('F')) {
            bool valid=false; const auto number=name.mid(1).toInt(&valid);
            if (!valid || number<1 || number>24) throw VbaRuntimeError(3251,"SendKeys key token unsupported");
            code=Qt::Key_F1+number-1;
        } else throw VbaRuntimeError(3251,"SendKeys key token unsupported");
        const auto text=code==Qt::Key_Return ? QString("\r") : code==Qt::Key_Tab ? QString("\t") : code==Qt::Key_Space ? QString(" ") : QString{};
        append(code,text,modifiers,count);
    }
    void group(Qt::KeyboardModifiers inherited,bool nested,int depth) {
        if (depth>16) throw VbaRuntimeError(3035,"SendKeys group nesting exceeded");
        while (position_<source_.size()) {
            if (source_[position_]==')') {
                if (!nested) throw VbaRuntimeError(5,"Unexpected SendKeys group terminator");
                ++position_; return;
            }
            auto modifiers=inherited;
            while (position_<source_.size() && QString("+^%").contains(source_[position_])) {
                const auto prefix=source_[position_++];
                modifiers|=prefix=='+' ? Qt::ShiftModifier : prefix=='^' ? Qt::ControlModifier : Qt::AltModifier;
            }
            if (position_==source_.size()) throw VbaRuntimeError(5,"SendKeys modifier has no key");
            if (source_[position_]=='(') { ++position_; group(modifiers,true,depth+1); }
            else if (source_[position_]=='{') brace(modifiers);
            else if (source_[position_]=='~') { ++position_; append(Qt::Key_Return,"\r",modifiers); }
            else {
                int size=1;
                if (source_[position_].isHighSurrogate() && position_+1<source_.size() && source_[position_+1].isLowSurrogate()) size=2;
                const auto value=source_.mid(position_,size); position_+=size; literal(value,modifiers);
            }
        }
        if (nested) throw VbaRuntimeError(5,"Unclosed SendKeys modifier group");
    }
    QString source_;
    qsizetype position_{};
    QVector<Key> result_;
};
QWidget* focused(QWidget* scope,const std::function<QWidget*()>& resolver) {
    auto* widget=resolver ? resolver() : QApplication::focusWidget();
    if (!scope || (!resolver && !scope->isVisible()) || !widget || !widget->isVisible() || !widget->isEnabled()
        || (!resolver && widget!=scope && !scope->isAncestorOf(widget))) throw VbaRuntimeError(2110,"Native keyboard focus unavailable");
    return widget;
}
void deliver(QWidget* scope,const QVector<Key>& keys,const std::function<QWidget*()>& resolver) {
    const QPointer<QWidget> guardedScope(scope);
    for (const auto& key:keys) {
        QPointer<QWidget> target=focused(guardedScope,resolver);
        QKeyEvent press(QEvent::KeyPress,key.code,key.modifiers,key.text);
        QApplication::sendEvent(target,&press);
        if (target) { QKeyEvent release(QEvent::KeyRelease,key.code,key.modifiers,key.text); QApplication::sendEvent(target,&release); }
    }
}
}
void sendNativeKeys(QWidget* scope,const QString& source,bool wait,std::function<void(int)> failure,std::function<QWidget*()> resolver) {
    const auto keys=KeyParser(source).parse(); // Validate the whole sequence before emitting anything.
    if (keys.isEmpty()) return;
    (void)focused(scope,resolver);
    if (wait) { deliver(scope,keys,resolver); return; }
    const QPointer<QWidget> guarded(scope);
    QMetaObject::invokeMethod(scope,[guarded,keys,failure=std::move(failure),resolver=std::move(resolver)] {
        if (!guarded) return;
        try { deliver(guarded,keys,resolver); }
        catch(const VbaRuntimeError& error) { if (failure) failure(error.number); }
    },Qt::QueuedConnection);
}
}
