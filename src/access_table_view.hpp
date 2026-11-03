#pragma once
#include "access_table_model.hpp"
#include "vba_runtime.hpp"
#include "csv.hpp"
#include "native_clock_keys.hpp"
#include "native_cell_editor.hpp"
#include "native_cell_zoom.hpp"
#include "native_table_find.hpp"
#include <QLineEdit>
#include <QSqlError>
#include <QScopeGuard>
#include <QClipboard>
#include <QMimeData>
#include <new>
#include <QItemSelectionModel>
#include <limits>
#include <QTableView>
#include <QMenu>
#include <QContextMenuEvent>
#include <QHeaderView>
#include <QApplication>
#include <QSignalBlocker>
#include <QMetaProperty>
#include <QPersistentModelIndex>
#include <QKeyEvent>
namespace linaccess {
class AccessTableView final : public QTableView {
public:
    explicit AccessTableView(QWidget* parent=nullptr):QTableView(parent) {
        // Connect before enabling Qt's header sorting: an editor may contain
        // text that has not reached the model's manual-submit cache yet.
        connect(horizontalHeader(),&QHeaderView::sortIndicatorChanged,this,[this] {
            auto* typed=qobject_cast<AccessTableModel*>(model()); if(!typed) return;
            if(activeEditor()) typed->protectNextSort();
        });
    }
    void setRecordNavigationGuard(std::function<bool()> guard) { recordNavigationGuard_=std::move(guard); }
    bool openFindDialog() {
        QPointer<AccessTableView> self=this;
        QPointer<AccessTableModel> typed=qobject_cast<AccessTableModel*>(model());
        if(!typed || !finishEditing() || !self || !typed) return false;
        QVector<QPair<QString,int>> columns;
        for(int column=0;column<typed->columnCount();++column)
            if(!isColumnHidden(column) && typed->isTextColumn(column)) columns.push_back({typed->headerData(column,Qt::Horizontal).toString(),column});
        if(columns.isEmpty()) return refuse(3251,"Keine durchsuchbaren Textfelder vorhanden.");
        if(findDialog_ && findDialog_->isVisible()) { findDialog_->raise(); findDialog_->activateWindow(); return true; }
        auto initial=findArguments_;
        if(initial.text.isEmpty() && currentIndex().isValid() && typed->isTextColumn(currentIndex().column())) initial.column=currentIndex().column();
        QPointer<QDialog> dialog=nativeTableFindDialog(this,columns,initial,[self](NativeTableFindArguments arguments) {
            if(!self) return QString{};
            self->findText(std::move(arguments));
            return self ? self->findMessage_ : QString{};
        });
        if(!self) return false;
        findDialog_=dialog;
        return bool(dialog);
    }
    bool hasFindText() const { return !findArguments_.text.isEmpty(); }
    bool repeatFind() {
        const QPointer<AccessTableView> self=this;
        const bool completed=findText(findArguments_)>=0;
        if(!self) return false;
        if(findDialog_) if(auto* label=findDialog_->findChild<QLabel*>("native-table-find-result")) label->setText(findMessage_);
        return completed;
    }
    int findText(NativeTableFindArguments arguments) {
        QPointer<AccessTableView> self=this;
        QPointer<AccessTableModel> typed=qobject_cast<AccessTableModel*>(model());
        const auto fail=[&](int number,const QString& message) {
            if(self) { findMessage_=message; (void)refuse(number,message); }
            return -1;
        };
        if(!typed) return fail(2046,"Keine native Tabelle für die Suche.");
        if(finding_) return fail(3251,"Eine Suche wird bereits ausgeführt.");
        finding_=true; const auto searching=qScopeGuard([self] { if(self) self->finding_=false; });
        if(arguments.text.isEmpty()) return fail(2046,"Keine vorherige Suche vorhanden.");
        if(arguments.mode<0 || arguments.mode>2 || arguments.column< -1 || arguments.column>=typed->columnCount() || arguments.origin<0 || arguments.origin>2 || arguments.direction<0 || arguments.direction>2) return fail(5,"Ungültige Sucheinstellung.");
        if(!finishEditing()) { if(self && typed) findMessage_=typed->lastError().text(); return -1; }
        if(!self || !typed) return -1;
        if(recordNavigationGuard_) {
            const auto guard=recordNavigationGuard_;
            if(!guard()) { if(self && typed) findMessage_=typed->lastError().text(); return -1; }
        } else if(typed->isDirty()) return fail(3251,"Zuerst Tabellenentwürfe vor der Suche speichern.");
        if(!self || !typed) return -1;
        QVector<int> columns;
        for(int column=0;column<typed->columnCount();++column)
            if((arguments.column<0 || arguments.column==column) && !isColumnHidden(column) && typed->isTextColumn(column)) columns.push_back(column);
        if(columns.isEmpty()) return fail(3265,"Das Suchfeld ist nicht verfügbar.");
        constexpr qint64 cellLimit=1000000;
        while(typed->canFetchMore()) {
            if(static_cast<qint64>(typed->rowCount())*columns.size()>=cellLimit) return fail(3251,"Suchumfang überschreitet die native Grenze.");
            const int before=typed->rowCount(); typed->fetchMore();
            if(!self || !typed) return -1;
            if(typed->rowCount()==before && typed->canFetchMore()) return fail(3251,"Suchdaten konnten nicht vollständig geladen werden.");
        }
        if(static_cast<qint64>(typed->rowCount())*columns.size()>cellLimit) return fail(3251,"Suchumfang überschreitet die native Grenze.");
        if(findArguments_.text!=arguments.text || findArguments_.column!=arguments.column || findArguments_.mode!=arguments.mode || findArguments_.sensitive!=arguments.sensitive || findArguments_.direction!=arguments.direction) findAdvance_=false;
        findArguments_=arguments; findArguments_.origin=0;
        const int count=typed->rowCount();
        if(!count) { findMessage_="Keine Datensätze vorhanden."; return 0; }
        const int step=arguments.direction==0 ? -1 : 1;
        const int start=arguments.origin==1 ? 0 : currentIndex().isValid() ? currentIndex().row()+((arguments.origin==2 || findAdvance_) ? step : 0) : 0;
        const auto cs=arguments.sensitive ? Qt::CaseSensitive : Qt::CaseInsensitive;
        for(int offset=0;offset<count;++offset) {
            const qint64 candidateRow=static_cast<qint64>(start)+static_cast<qint64>(step)*offset;
            if(arguments.direction!=2 && (candidateRow<0 || candidateRow>=count)) break;
            const int row=static_cast<int>((candidateRow%count+count)%count); if(isRowHidden(row)) continue;
            for(int column:columns) {
                const auto candidate=typed->index(row,column);
                const auto value=candidate.data(Qt::EditRole);
                if(value.isNull() || value.metaType().id()==QMetaType::QByteArray) continue;
                const auto text=value.toString();
                const bool matches=arguments.mode==1 ? text.compare(arguments.text,cs)==0 : arguments.mode==2 ? text.startsWith(arguments.text,cs) : text.contains(arguments.text,cs);
                if(matches) {
                    setCurrentIndex(candidate);
                    if(!self || !typed) return -1;
                    scrollTo(candidate); findAdvance_=true;
                    findMessage_=QString("Übereinstimmung in Datensatz %1.").arg(row+1); return 1;
                }
            }
        }
        findMessage_="Keine Übereinstimmung gefunden."; return 0;
    }
    bool hasPendingEditor() const {
        auto* editor=activeEditor(); if(!editor) return false;
        const auto property=editor->metaObject()->userProperty();
        const auto stored=editor->property("la_cell_index").value<QPersistentModelIndex>();
        const QModelIndex target=stored.isValid() ? QModelIndex(stored) : currentIndex();
        if(!property.isValid() || target.model()!=model()) return true;
        return nativeCellEditorChanged(editor,target);
    }
    bool finishEditing() {
        auto* editor=activeEditor(); if(!editor) return true;
        const auto property=editor->metaObject()->userProperty(); if(!property.isValid() || !model()) return false;
        const auto stored=editor->property("la_cell_index").value<QPersistentModelIndex>();
        const QModelIndex target=stored.isValid() ? QModelIndex(stored) : currentIndex();
        if(target.model()!=model()) return false;
        if(!commitNativeCellEditor(editor,target)) return false;
        QTableView::closeEditor(editor,QAbstractItemDelegate::NoHint); return true;
    }
    bool insertRecord() {
        auto* typed=qobject_cast<AccessTableModel*>(model());
        if(!typed || !finishEditing()) return false;
        const auto row=typed->rowCount();
        if(!typed->insertRow(row)) return false;
        setCurrentIndex(typed->index(row,0)); scrollTo(currentIndex()); setFocus();
        return true;
    }
    bool deleteCurrentRecord() {
        auto* typed=qobject_cast<AccessTableModel*>(model());
        if(!typed || !currentIndex().isValid() || !finishEditing()) return false;
        const auto row=currentIndex().row(),column=currentIndex().column();
        if(!typed->removeRow(row)) return false;
        if(typed->rowCount()) setCurrentIndex(typed->index(qMin(row,typed->rowCount()-1),column));
        return true;
    }
    bool copySelection() {
        auto* typed=qobject_cast<AccessTableModel*>(model());
        if(!typed || !selectionModel() || !finishEditing()) return false;
        const auto selected=selectionModel()->selectedIndexes();
        if(selected.isEmpty()) return false;
        if(selected.size()>1000000) return refuse(3251,"Die Auswahl ist für die Textzwischenablage zu groß.");
        int top=std::numeric_limits<int>::max(),left=top,bottom=-1,right=-1;
        for(const auto& index:selected) {
            if(isRowHidden(index.row()) || isColumnHidden(index.column())) return refuse(3251,"Ausgeblendete Zellen werden nicht in die Textzwischenablage übernommen.");
            top=qMin(top,index.row()); bottom=qMax(bottom,index.row()); left=qMin(left,index.column()); right=qMax(right,index.column());
        }
        if(static_cast<qint64>(bottom-top+1)*(right-left+1)!=selected.size())
            return refuse(3251,"Zum Kopieren einen zusammenhängenden rechteckigen Bereich auswählen.");
        QString output;
        try {
            for(int row=top;row<=bottom;++row) {
                if(row>top) output+="\r\n";
                for(int column=left;column<=right;++column) {
                    if(column>left) output+='\t';
                    auto text=typed->clipboardText(typed->index(row,column));
                    if(text.contains(QChar(0))) return refuse(3251,"Nullzeichen werden in der Textzwischenablage noch nicht unterstützt.");
                    if(text.contains('\t') || text.contains('\r') || text.contains('\n') || text.contains('"')) {
                        text.replace("\"","\"\""); text='"'+text+'"';
                    }
                    if(output.size()+text.size()>4*1024*1024) return refuse(3251,"Der Text ist für die Zwischenablage zu groß.");
                    output+=text;
                }
            }
        } catch(const VbaRuntimeError& error) { return refuse(error.number,"Dieser Zellwert kann noch nicht verlustfrei als Text kopiert werden."); }
        QApplication::clipboard()->setText(output);
        return true;
    }
    bool zoomCurrentCell() {
        QPointer<AccessTableView> self=this;
        QPointer<AccessTableModel> typed=qobject_cast<AccessTableModel*>(model());
        if(!typed || !finishEditing()) return false;
        const QPersistentModelIndex target=currentIndex();
        if(!target.isValid()) return false;
        QString text;
        try { text=typed->clipboardText(target); }
        catch(const VbaRuntimeError& error) { return refuse(error.number,"Dieser Zellwert kann noch nicht im Text-Zoom bearbeitet werden."); }
        const bool writable=target.flags().testFlag(Qt::ItemIsEditable) && typed->editStrategy()==QSqlTableModel::OnManualSubmit;
        const auto result=nativeCellZoom(this,text,!writable);
        if(!self || !typed || !target.isValid()) return false;
        if(!result) return true; // Cancel/readonly Close completes the command without mutation.
        if(result->text==text && !(target.data(Qt::EditRole).isNull() && result->modified)) return true;
        return typed->setData(target,result->text,Qt::EditRole);
    }
    bool canPasteSelection() const {
        auto* typed=qobject_cast<AccessTableModel*>(model());
        return typed && typed->editStrategy()==QSqlTableModel::OnManualSubmit && !typed->isDirty()
            && !hasPendingEditor() && currentIndex().isValid() && currentIndex().flags().testFlag(Qt::ItemIsEditable);
    }
    bool pasteSelection() {
        auto* typed=qobject_cast<AccessTableModel*>(model());
        if(!typed || !finishEditing()) return false;
        if(typed->isDirty()) return refuse(2501,"Vor dem Einfügen eines Bereichs Änderungen speichern oder verwerfen.");
        if(typed->editStrategy()!=QSqlTableModel::OnManualSubmit) return refuse(3251,"Bereichseinfügen ist in dieser Ansicht nicht verfügbar.");
        const auto first=currentIndex();
        if(!first.isValid()) return false;
        if(!first.flags().testFlag(Qt::ItemIsEditable)) return refuse(2448,"Die Zielzelle ist schreibgeschützt.");
        const auto* mime=QApplication::clipboard()->mimeData();
        if(!mime || !mime->hasText()) return false;
        const auto text=QApplication::clipboard()->text();
        if(text.size()>4*1024*1024 || text.contains(QChar(0))) return refuse(3251,"Dieser Zwischenablageumfang wird noch nicht unterstützt.");
        QVector<QVector<QString>> values;
        try {
            const auto utf8=text.toUtf8();
            if(QString::fromUtf8(utf8)!=text) return refuse(13,"Der Zwischenablagetext enthält nicht unterstützte Unicode-Zeichen.");
            const auto parsed=parseTsvDetailed(std::string_view(utf8.constData(),utf8.size()));
            if(parsed.empty()) values={{QString{}}};
            else {
                const auto columns=parsed.front().size();
                if(!columns) return refuse(5,"Die Zwischenablage enthält keinen rechteckigen Tabellenbereich.");
                for(const auto& row:parsed) {
                    if(row.size()!=columns) return refuse(5,"Die Zwischenablage enthält keinen rechteckigen Tabellenbereich.");
                    QVector<QString> cells; cells.reserve(columns);
                    for(const auto& cell:row) cells.push_back(QString::fromUtf8(cell.value.data(),cell.value.size()));
                    values.push_back(std::move(cells));
                }
            }
        } catch(const std::bad_alloc&) { return refuse(7,"Der Tabellenbereich konnte nicht im Speicher vorbereitet werden."); }
          catch(const std::length_error&) { return refuse(3251,"Die Zwischenablage enthält zu viele Zellen."); }
          catch(const std::exception&) { return refuse(5,"Die Zwischenablage enthält kein gültiges Tabellenformat."); }
        const auto endRow=static_cast<qint64>(first.row())+values.size();
        if(endRow>std::numeric_limits<int>::max() || static_cast<qint64>(first.column())+values.front().size()>typed->columnCount())
            return refuse(5,"Der einzufügende Bereich passt nicht in die vorhandenen Spalten.");
        while(typed->rowCount()<endRow && typed->canFetchMore()) {
            const auto count=typed->rowCount(); typed->fetchMore(); if(typed->rowCount()==count) break;
        }
        if(typed->rowCount()<endRow) return refuse(3251,"Zum Einfügen sind noch nicht genügend vorhandene Datensätze verfügbar.");
        if(!typed->setDataBlock(first,values)) return false;
        if(selectionModel()) selectionModel()->select(QItemSelection(first,typed->index(endRow-1,first.column()+values.front().size()-1)),QItemSelectionModel::ClearAndSelect);
        return true;
    }
    void setModel(QAbstractItemModel* model) override {
        disconnect(restore_); QTableView::setModel(model);
        if(auto* typed=qobject_cast<AccessTableModel*>(model))
            restore_=connect(typed,&AccessTableModel::sortRestored,this,[this](int column,Qt::SortOrder order) {
                const QSignalBlocker blocked(horizontalHeader()); horizontalHeader()->setSortIndicator(column,order);
            });
    }
protected:
    void contextMenuEvent(QContextMenuEvent* event) override {
        if(!model() || !selectionModel() || !finishEditing()) { event->accept(); return; }
        if(event->reason()==QContextMenuEvent::Mouse) {
            const auto clicked=indexAt(event->pos());
            if(clicked.isValid() && !selectionModel()->isSelected(clicked)) setCurrentIndex(clicked);
        }
        QPointer<AccessTableView> self=this;
        QMenu menu; menu.setObjectName("native-datasheet-context-menu");
        connect(this,&QObject::destroyed,&menu,&QMenu::close);
        const auto add=[&](const char* name,const QString& label,bool enabled) {
            auto* action=menu.addAction(label); action->setObjectName(QString("datasheet-")+name); action->setEnabled(enabled); return action;
        };
        const bool native=qobject_cast<AccessTableModel*>(model())!=nullptr;
        auto* copy=add("copy","&Kopieren\tStrg+C",native && selectionModel()->hasSelection());
        const auto* clipboard=QApplication::clipboard()->mimeData();
        auto* paste=add("paste","&Einfügen\tStrg+V",canPasteSelection() && clipboard && clipboard->hasText());
        menu.addSeparator();
        auto* zoom=add("zoom","&Zoom...\tUmschalt+F2",native && currentIndex().isValid());
        menu.addSeparator();
        auto* row=add("select-row","&Datensatz markieren\tUmschalt+Leertaste",currentIndex().isValid());
        auto* column=add("select-column","&Spalte markieren\tStrg+Leertaste",currentIndex().isValid());
        auto* all=add("select-all","&Alles markieren\tStrg+A",model()->rowCount()>0 && model()->columnCount()>0);
        const auto position=event->reason()==QContextMenuEvent::Keyboard && currentIndex().isValid()
            ? viewport()->mapToGlobal(visualRect(currentIndex()).center()) : event->globalPos();
        auto* chosen=menu.exec(position);
        if(!self) return;
        if(chosen==copy) (void)copySelection();
        else if(chosen==paste) (void)pasteSelection();
        else if(chosen==zoom) (void)zoomCurrentCell();
        else if(chosen==row && currentIndex().isValid()) selectionModel()->select(currentIndex(),QItemSelectionModel::ClearAndSelect|QItemSelectionModel::Rows);
        else if(chosen==column && currentIndex().isValid() && fetchSelectionRows()) selectionModel()->select(currentIndex(),QItemSelectionModel::ClearAndSelect|QItemSelectionModel::Columns);
        else if(chosen==all && fetchSelectionRows()) selectAll();
        event->accept();
    }
    void closeEditor(QWidget* widget,QAbstractItemDelegate::EndEditHint hint) override {
        QPointer<AccessTableView> self=this;
        QPointer<QWidget> editor=widget;
        const auto target=widget ? widget->property("la_cell_index").value<QPersistentModelIndex>() : QPersistentModelIndex{};
        if(hint!=QAbstractItemDelegate::RevertModelCache && editor && state()==QAbstractItemView::EditingState && target.isValid() && target.model()==model()) {
            if(!commitNativeCellEditor(editor,target)) {
                if(!self || !editor) return;
                if(selectionModel() && target.isValid()) {
                    const QSignalBlocker blocked(selectionModel());
                    selectionModel()->setCurrentIndex(target,QItemSelectionModel::ClearAndSelect);
                }
                // Restore focus after the mouse/focus transition completes;
                // keeping the editor alive retains the rejected text.
                QMetaObject::invokeMethod(editor,[editor] {
                    if(editor && editor->isVisible() && editor->isEnabled() && editor->window()->isActiveWindow())
                        editor->setFocus(Qt::OtherFocusReason);
                },Qt::QueuedConnection);
                return;
            }
            if(!self || !editor) return;
        }
        QTableView::closeEditor(widget,hint);
    }
    void keyPressEvent(QKeyEvent* event) override {
        const auto selectionModifiers=event->modifiers() & ~Qt::KeypadModifier;
        if(event->key()==Qt::Key_End && selectionModifiers==Qt::ControlModifier) {
            if(finishEditing() && fetchSelectionRows()) QTableView::keyPressEvent(event);
            event->accept(); return;
        }
        if(event->key()==Qt::Key_F2 && selectionModifiers==Qt::ShiftModifier) {
            (void)zoomCurrentCell(); event->accept(); return;
        }
        if(event->key()==Qt::Key_F2 && selectionModifiers==Qt::NoModifier) {
            if(activeEditor()) { if(finishEditing()) setFocus(); }
            else if(currentIndex().isValid() && currentIndex().flags().testFlag(Qt::ItemIsEditable)) {
                edit(currentIndex());
                if(auto* line=qobject_cast<QLineEdit*>(activeEditor())) line->setCursorPosition(line->text().size());
            }
            event->accept(); return;
        }
        if(event->key()==Qt::Key_Apostrophe && selectionModifiers==Qt::ControlModifier) {
            if(auto* typed=qobject_cast<AccessTableModel*>(model());typed && finishEditing()) {
                const auto target=currentIndex();
                if(target.isValid() && target.row()>0 && !isRowHidden(target.row()-1) && !isColumnHidden(target.column()))
                    (void)typed->repeatPreviousValue(target);
            }
            event->accept(); return;
        }
        if(event->key()==Qt::Key_C && selectionModifiers==Qt::ControlModifier) {
            (void)copySelection(); event->accept(); return;
        }
        if(event->key()==Qt::Key_V && selectionModifiers==Qt::ControlModifier) {
            (void)pasteSelection(); event->accept(); return;
        }
        if((event->key()==Qt::Key_Space && (selectionModifiers==Qt::ShiftModifier || selectionModifiers==Qt::ControlModifier))
            || (event->key()==Qt::Key_A && selectionModifiers==Qt::ControlModifier)) {
            if(finishEditing() && selectionModel() && model()) {
                if(event->key()==Qt::Key_Space && selectionModifiers==Qt::ShiftModifier) {
                    if(currentIndex().isValid()) selectionModel()->select(currentIndex(),QItemSelectionModel::ClearAndSelect|QItemSelectionModel::Rows);
                } else if(fetchSelectionRows()) {
                    if(event->key()==Qt::Key_A) selectAll();
                    else if(currentIndex().isValid()) selectionModel()->select(currentIndex(),QItemSelectionModel::ClearAndSelect|QItemSelectionModel::Columns);
                }
            }
            event->accept(); return;
        }
        if(const auto text=nativeClockKeyText(*event)) {
            if(currentIndex().isValid() && currentIndex().flags().testFlag(Qt::ItemIsEditable)) {
                const bool fresh=!activeEditor();
                if(fresh) edit(currentIndex());
                if(auto* editor=qobject_cast<QLineEdit*>(activeEditor());editor && editor->isEnabled() && !editor->isReadOnly()) {
                    if(fresh) editor->selectAll();
                    if(editor->text().size()-editor->selectedText().size()+text->size()<=editor->maxLength()) editor->insert(*text);
                }
            }
            event->accept(); return;
        }
        const auto modifiers=event->modifiers() & ~Qt::KeypadModifier;
        if((modifiers==Qt::ControlModifier || modifiers==(Qt::ControlModifier|Qt::ShiftModifier))
            && (event->key()==Qt::Key_Plus || (event->key()==Qt::Key_Equal && modifiers.testFlag(Qt::ShiftModifier)))) {
            (void)insertRecord(); event->accept(); return;
        }
        if(modifiers==Qt::ControlModifier && event->key()==Qt::Key_Minus) {
            (void)deleteCurrentRecord(); event->accept(); return;
        }
        if(event->key()==Qt::Key_F && modifiers==Qt::ControlModifier) { (void)openFindDialog(); event->accept(); return; }
        if(event->key()==Qt::Key_F4 && modifiers==Qt::ShiftModifier) { (void)repeatFind(); event->accept(); return; }
        QTableView::keyPressEvent(event);
    }
private:
    bool refuse(int number,const QString& message) const {
        if(auto* typed=qobject_cast<AccessTableModel*>(model())) return typed->rejectEdit(number,message);
        return false;
    }
    bool fetchSelectionRows() {
        while(model() && model()->canFetchMore({})) {
            if(static_cast<qint64>(model()->rowCount())*qMax(1,model()->columnCount())>=1000000)
                return refuse(3251,"Die vollständige Auswahl überschreitet den unterstützten Umfang.");
            const auto count=model()->rowCount(); model()->fetchMore({});
            if(model()->rowCount()==count) return refuse(3251,"Die vollständige Auswahl konnte nicht geladen werden.");
        }
        if(model() && static_cast<qint64>(model()->rowCount())*qMax(1,model()->columnCount())>1000000)
            return refuse(3251,"Die vollständige Auswahl überschreitet den unterstützten Umfang.");
        return model()!=nullptr;
    }
    QWidget* activeEditor() const {
        if(state()!=QAbstractItemView::EditingState) return nullptr;
        auto* editor=QApplication::focusWidget();
        while(editor && editor->parentWidget()!=viewport()) editor=editor->parentWidget();
        if(editor && !editor->isHidden()) return editor;
        for(auto* candidate:viewport()->findChildren<QWidget*>(QString{},Qt::FindDirectChildrenOnly))
            if(!candidate->isHidden() && candidate->property("la_cell_index").value<QPersistentModelIndex>().isValid()) return candidate;
        return nullptr;
    }
    std::function<bool()> recordNavigationGuard_;
    QPointer<QDialog> findDialog_;
    NativeTableFindArguments findArguments_;
    QString findMessage_;
    bool findAdvance_{},finding_{};
    QMetaObject::Connection restore_;
};
}
