#include "native_send_keys.hpp"
#include "vba_runtime.hpp"
#include <QApplication>
#include <QKeyEvent>
#include <QPointer>
#include <QWidget>
#include <QMap>
#include <QVector>
#include <QEventLoop>
#include <QScopeGuard>
#include <utility>
#include <deque>
#include <exception>
#include <memory>
#include <QTimer>
#include <QElapsedTimer>

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
struct KeyBatch {
    QPointer<QWidget> scope;
    QVector<Key> keys;
    std::function<void(int)> failure;
    std::function<QWidget*()> resolver;
    std::uint64_t epoch{};
    qsizetype next{};
    int inFlight{};
    bool finished{},scheduled{},synchronous{};
    std::exception_ptr error;
    QPointer<QEventLoop> waiter;
    QElapsedTimer focusWait;
};
using Batch=std::shared_ptr<KeyBatch>;
thread_local std::deque<Batch> pending;
thread_local Batch issuing;
thread_local bool kickScheduled=false;
thread_local qsizetype pendingKeys=0;
thread_local std::uint64_t dispatchedKeys=0;
thread_local int keyHandlerDepth=0;
constexpr qsizetype queueLimit=65536;
void schedule(const Batch& batch);
void kick();
void complete(const Batch& batch) {
    if(batch->finished || batch->inFlight || batch->next<batch->keys.size()) return;
    batch->finished=true;
    if(batch->waiter) batch->waiter->quit();
    if(!batch->synchronous && batch->scope && batch->error) {
        try { std::rethrow_exception(batch->error); }
        catch(const VbaExecutionStopped&) {}
        catch(const VbaRuntimeError& error) { if(batch->failure) batch->failure(error.number); }
        catch(...) { if(batch->failure) batch->failure(3251); }
    }
}
void release(const Batch& batch) {
    if(issuing==batch && batch->next>=batch->keys.size()) { issuing.reset(); kick(); }
}
void abort(const Batch& batch,std::exception_ptr error) {
    if(batch->finished) return;
    pendingKeys-=batch->keys.size()-batch->next;
    batch->next=batch->keys.size(); batch->error=std::move(error);
    release(batch); complete(batch);
}
void step(const Batch& batch) {
    if(batch->finished) return;
    if(batch->epoch!=vbaExecutionEpoch() || vbaExecutionStopRequested()) { abort(batch,std::make_exception_ptr(VbaExecutionStopped{})); return; }
    if(!batch->scope) {
        abort(batch,batch->synchronous ? std::make_exception_ptr(VbaRuntimeError(2110,"Native keyboard scope closed")) : std::exception_ptr{}); return;
    }
    if(keyHandlerDepth>=32) { abort(batch,std::make_exception_ptr(VbaRuntimeError(3035,"SendKeys delivery nesting exceeded"))); return; }
    if(batch->next>=batch->keys.size()) { complete(batch); return; }
    ++batch->inFlight; ++keyHandlerDepth;
    const auto leave=qScopeGuard([batch] { --keyHandlerDepth; --batch->inFlight; complete(batch); });
    try {
        QPointer<QWidget> target;
        try { target=focused(batch->scope,batch->resolver); }
        catch(const VbaRuntimeError& error) {
            if(error.number!=2110) throw;
            if(!batch->focusWait.isValid()) batch->focusWait.start();
            if(batch->focusWait.elapsed()>=500) throw;
            QTimer::singleShot(1,QCoreApplication::instance(),[batch] { schedule(batch); });
            return;
        }
        batch->focusWait.invalidate();
        const auto key=batch->keys[batch->next++]; --pendingKeys; ++dispatchedKeys;
        // Queue the continuation before delivery: a modal key handler can run
        // its own Qt loop, where the remaining keys must still be available.
        if(batch->next<batch->keys.size()) schedule(batch); else release(batch);
        QKeyEvent press(QEvent::KeyPress,key.code,key.modifiers,key.text);
        QApplication::sendEvent(target,&press);
        if(target) { QKeyEvent up(QEvent::KeyRelease,key.code,key.modifiers,key.text); QApplication::sendEvent(target,&up); }
        if(batch->epoch!=vbaExecutionEpoch() || vbaExecutionStopRequested()) throw VbaExecutionStopped{};
    } catch(...) { abort(batch,std::current_exception()); }
}
void schedule(const Batch& batch) {
    if(batch->scheduled || batch->finished) return;
    batch->scheduled=true;
    // Give platform activation/focus events a turn before resolving the next
    // receiver. This also works inside native modal event loops.
    QTimer::singleShot(0,QCoreApplication::instance(),[batch] {
        QTimer::singleShot(0,QCoreApplication::instance(),[batch] { batch->scheduled=false; step(batch); });
    });
}
void kick() {
    if(kickScheduled || issuing || pending.empty()) return;
    kickScheduled=true;
    QTimer::singleShot(0,QCoreApplication::instance(),[] {
        kickScheduled=false;
        if(issuing) return;
        while(!pending.empty()) {
            auto batch=pending.front(); pending.pop_front();
            if(batch->finished) continue;
            if(batch->epoch!=vbaExecutionEpoch()) { abort(batch,std::make_exception_ptr(VbaExecutionStopped{})); continue; }
            if(!batch->scope) { abort(batch,{}); continue; }
            issuing=batch; schedule(batch); return;
        }
    });
}
}
void sendNativeKeys(QWidget* scope,const QString& source,bool wait,std::function<void(int)> failure,std::function<QWidget*()> resolver) {
    auto keys=KeyParser(source).parse(); // Validate the whole sequence before emitting anything.
    if(keys.isEmpty()) return;
    if(!scope) throw VbaRuntimeError(2110,"Native keyboard scope unavailable");
    try { (void)focused(scope,resolver); }
    catch(const VbaRuntimeError& error) { if(error.number!=2110 || !resolver) throw; }
    if(pendingKeys+keys.size()>queueLimit) throw VbaRuntimeError(3035,"SendKeys queue budget exceeded");
    auto batch=std::make_shared<KeyBatch>();
    batch->scope=scope; batch->keys=std::move(keys); batch->failure=std::move(failure); batch->resolver=std::move(resolver);
    batch->epoch=vbaExecutionEpoch(); batch->synchronous=wait;
    pendingKeys+=batch->keys.size();
    if(!wait) { pending.push_back(batch); kick(); return; }
    QEventLoop loop; batch->waiter=&loop;
    // Preserve immediate first-key delivery when nothing precedes it. Further
    // keys use the same ordered queue as asynchronous batches.
    if(!issuing && pending.empty()) { issuing=batch; step(batch); }
    else { pending.push_back(batch); kick(); }
    if(!batch->finished) loop.exec(QEventLoop::ExcludeUserInputEvents);
    batch->waiter=nullptr;
    if(!batch->finished) abort(batch,std::make_exception_ptr(VbaExecutionStopped{}));
    if(batch->error) std::rethrow_exception(batch->error);
}
void drainNativeKeys() {
    const auto start=dispatchedKeys;
    while(pendingKeys) {
        if(vbaExecutionStopRequested()) throw VbaExecutionStopped{};
        if(dispatchedKeys-start>static_cast<std::uint64_t>(queueLimit)) {
            const auto error=std::make_exception_ptr(VbaRuntimeError(3035,"SendKeys DoEvents budget exceeded"));
            auto current=issuing; if(current) abort(current,error);
            auto queued=std::move(pending); pending.clear();
            for(const auto& batch:queued) abort(batch,error);
            std::rethrow_exception(error);
        }
        QCoreApplication::processEvents(QEventLoop::AllEvents);
    }
}
}
