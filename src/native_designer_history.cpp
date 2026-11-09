#include "native_designer_history.hpp"
#include <QDialog>
#include <QBoxLayout>
#include <QPushButton>
#include <QLabel>
#include <QJsonDocument>
#include <QShortcut>
#include <QKeySequence>
#include <QScopedValueRollback>
#include <exception>
namespace linaccess {
NativeDesignerHistory::NativeDesignerHistory(QDialog& dialog,QBoxLayout& buttons,Capture capture,Restore restore)
    :QObject(&dialog),dialog_(&dialog),capture_(std::move(capture)),restore_(std::move(restore)),fresh_(dialog.property("la_designer_new").toBool()) {
    undo_=new QPushButton("Rückgängig",&dialog);undo_->setObjectName("native-designer-undo");buttons.addWidget(undo_);
    redo_=new QPushButton("Wiederholen",&dialog);redo_->setObjectName("native-designer-redo");buttons.addWidget(redo_);
    status_=new QLabel(&dialog);status_->setObjectName("native-designer-history-status");buttons.addWidget(status_);buttons.addStretch();
    timer_.setSingleShot(true);timer_.setInterval(250);
    connect(&timer_,&QTimer::timeout,this,[this] {flush();});
    connect(undo_,&QPushButton::clicked,this,[this] {move(-1);});
    connect(redo_,&QPushButton::clicked,this,[this] {move(1);});
    auto* undoKey=new QShortcut(QKeySequence::Undo,&dialog);undoKey->setContext(Qt::WidgetWithChildrenShortcut);
    auto* redoKey=new QShortcut(QKeySequence::Redo,&dialog);redoKey->setContext(Qt::WidgetWithChildrenShortcut);
    connect(undoKey,&QShortcut::activated,this,[this] {move(-1);});connect(redoKey,&QShortcut::activated,this,[this] {move(1);});
    // Text widgets keep their ordinary local Undo through ShortcutOverride.
    try {auto state=QJsonDocument(capture_()).toJson(QJsonDocument::Compact);states_.push_back(state);bytes_=state.size();saved_=state;if(bytes_>8*1024*1024) disable();}
    catch(const std::exception&) {disable();}
    update();
}
void NativeDesignerHistory::disable() {
    available_=false;timer_.stop();states_.clear();saved_.clear();bytes_=0;
    status_->setText("Rückgängig-Verlauf für diesen Umfang nicht verfügbar.");
    undo_->setEnabled(false);redo_->setEnabled(false);
}
void NativeDesignerHistory::update() {
    if(!dialog_ || !available_) return;
    undo_->setEnabled(cursor_>0 || timer_.isActive());redo_->setEnabled(cursor_+1<states_.size());
    if(!timer_.isActive()) {
        const bool dirty=fresh_ || states_[cursor_]!=saved_;dialog_->setWindowModified(dirty);
        if(auto* label=dialog_->findChild<QLabel*>("native-designer-save-status")) label->setText(dirty ? "Nicht gespeicherte Änderungen" : "Gespeichert");
    }
}
void NativeDesignerHistory::record() {
    if(restoring_ || !available_ || !dialog_) return;
    timer_.start();dialog_->setWindowModified(true);update();
}
void NativeDesignerHistory::flush() {
    timer_.stop();if(restoring_ || !available_ || !dialog_) return;
    try {
        auto state=QJsonDocument(capture_()).toJson(QJsonDocument::Compact);
        if(state.size()>8*1024*1024) {disable();return;}
        if(state!=states_[cursor_]) {
            while(states_.size()>cursor_+1) {bytes_-=states_.back().size();states_.pop_back();}
            bytes_+=state.size();states_.push_back(std::move(state));++cursor_;
            while(states_.size()>65 || bytes_>32*1024*1024) {bytes_-=states_.front().size();states_.removeFirst();--cursor_;}
        }
        update();
    } catch(const std::exception&) {disable();}
}
void NativeDesignerHistory::move(int direction) {
    if(!available_ || !dialog_ || restoring_) return;
    flush();if(!available_) return;
    const auto target=cursor_+direction;if(target<0 || target>=states_.size()) return;
    try {
        const QScopedValueRollback<bool> restoring(restoring_,true);
        restore_(QJsonDocument::fromJson(states_[target]).object());cursor_=target;
    } catch(const std::exception&) {disable();if(dialog_) dialog_->setWindowModified(true);return;}
    update();
}
void NativeDesignerHistory::markSaved() {
    if(!available_ || !dialog_) return;
    flush();if(!available_) return;
    saved_=states_[cursor_];fresh_=false;update();
}
}
