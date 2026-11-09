#pragma once
#include "access_table_model.hpp"
#include "native_grid_clipboard.hpp"
#include "native_grid_mime.hpp"
#include "native_text_replace.hpp"
#include "vba_runtime.hpp"
#include "csv.hpp"
#include "native_clock_keys.hpp"
#include "native_cell_editor.hpp"
#include "native_cell_zoom.hpp"
#include "native_table_find.hpp"
#include <QLineEdit>
#include <QSqlError>
#include <QSqlIndex>
#include <QSqlRecord>
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
#include <QShortcut>
namespace linaccess {
class AccessTableView final : public QTableView {
public:
    enum class RecordMove { Previous,Next,First,Last };
    explicit AccessTableView(QWidget* parent=nullptr):QTableView(parent) {
        // Connect before enabling Qt's header sorting: an editor may contain
        // text that has not reached the model's manual-submit cache yet.
        auto* replace=new QShortcut(QKeySequence(Qt::CTRL|Qt::Key_H),this);
        replace->setContext(Qt::WidgetWithChildrenShortcut);
        connect(replace,&QShortcut::activated,this,[this] { (void)openReplaceDialog(); });
        connect(horizontalHeader(),&QHeaderView::sortIndicatorChanged,this,[this] {
            auto* typed=qobject_cast<AccessTableModel*>(model()); if(!typed) return;
            if(activeEditor()) typed->protectNextSort();
        });
    }
    void setRecordNavigationGuard(std::function<bool()> guard) { recordNavigationGuard_=std::move(guard); }
    void setRecordNavigationHandler(std::function<bool(RecordMove,int)> handler) { recordNavigationHandler_=std::move(handler); }
    bool openReplaceDialog() { return openFindDialog(true); }
    bool openFindDialog(bool replace=false) {
        QPointer<AccessTableView> self=this;
        QPointer<AccessTableModel> typed=qobject_cast<AccessTableModel*>(model());
        if(!typed || !finishEditing() || !self || !typed) return false;
        if(replace && typed->isReadOnly()) return refuse(2046,"Das Datenblatt ist schreibgeschützt.");
        QVector<QPair<QString,int>> columns; QMap<int,QString> typedHints;
        for(int column=0;column<typed->columnCount();++column) if(!isColumnHidden(column) && (replace ? typed->isTextColumn(column) : typed->isSearchableColumn(column))) {
            columns.push_back({typed->headerData(column,Qt::Horizontal).toString(),column});
            if(!typed->isTextColumn(column)) typedHints.insert(column,nativeFindInputHint(typed->fieldSqlType(column)));
        }
        if(columns.isEmpty()) return refuse(3251,"Keine durchsuchbaren Felder vorhanden.");
        auto existing=replace ? replaceDialog_ : findDialog_;
        if(existing && existing->isVisible()) { existing->raise(); existing->activateWindow(); return true; }
        auto initial=findArguments_;
        if(initial.text.isEmpty() && currentIndex().isValid() && typed->isSearchableColumn(currentIndex().column())) initial.column=currentIndex().column();
        if(replace) { initial.mode=1; initial.value.reset();initial.formatted=false;initial.wildcards=false; }
        std::function<QString(NativeTableFindArguments,QString)> replaceAction;
        if(replace) replaceAction=[self](NativeTableFindArguments arguments,QString value) {
            if(self) self->replaceCurrentText(std::move(arguments),std::move(value));
            return self ? self->findMessage_ : QString{};
        };
        QPointer<QDialog> dialog=nativeTableFindDialog(this,columns,typedHints,initial,[self](NativeTableFindArguments arguments) {
            if(!self) return QString{};
            try {
                auto* model=qobject_cast<AccessTableModel*>(self->model());
                if(model && !arguments.formatted && !arguments.wildcards && arguments.column>=0 && !model->isTextColumn(arguments.column)) {
                    arguments.value=nativeFindInput(arguments.text,model->fieldSqlType(arguments.column));
                    arguments.text=nativeFindValueText(*arguments.value); arguments.mode=1;
                }
                self->findText(std::move(arguments));
            } catch(const VbaRuntimeError& error) {
                if(self) { self->findMessage_=QString::fromUtf8(error.what()); (void)self->refuse(error.number,self->findMessage_); }
            }
            return self ? self->findMessage_ : QString{};
        },std::move(replaceAction),replace ? std::function<QString(NativeTableFindArguments,QString)>([self](NativeTableFindArguments arguments,QString value) {
            if(self) self->replaceAllText(std::move(arguments),std::move(value));
            return self ? self->findMessage_ : QString{};
        }) : std::function<QString(NativeTableFindArguments,QString)>{});
        if(!self) return false;
        if(replace) replaceDialog_=dialog; else findDialog_=dialog;
        return bool(dialog);
    }
    bool replaceCurrentText(NativeTableFindArguments arguments,QString replacement) {
        const QPointer<AccessTableView> self=this;
        QPointer<AccessTableModel> typed=qobject_cast<AccessTableModel*>(model());
        const auto fail=[&](int number,const QString& message) { if(self) { findMessage_=message; (void)refuse(number,message); } return false; };
        if(!typed) return fail(2046,"Kein natives Datenblatt zum Ersetzen.");
        if(replacing_ || finding_) return fail(3251,"Eine Such- oder Ersetzungsaktion läuft bereits.");
        replacing_=true; const auto finished=qScopeGuard([self] { if(self) self->replacing_=false; });
        if(typed->isReadOnly()) return fail(2046,"Das Datenblatt ist schreibgeschützt.");
        if(arguments.formatted || arguments.wildcards || arguments.value || arguments.mode<0 || arguments.mode>2 || arguments.column<0 || arguments.column>=typed->columnCount()
            || isColumnHidden(arguments.column) || !typed->isTextColumn(arguments.column))
            return fail(3251,"Ersetzen erfordert ein ausgewähltes Textfeld und einen gültigen Vergleichsmodus.");
        if(arguments.text.isEmpty()) return fail(5,"Der Suchwert darf nicht leer sein.");
        try {(void)replaceNativeText({},arguments.text,replacement,arguments.mode,arguments.sensitive ? Qt::CaseSensitive : Qt::CaseInsensitive,false);}catch(const VbaRuntimeError& error) {return fail(error.number,QString::fromUtf8(error.what()));}
        if(arguments.text.size()>1000000 || replacement.size()>1000000) return fail(3251,"Der Text überschreitet den unterstützten Umfang.");
        if(!finishEditing() || !self || !typed) return false;
        if(!currentIndex().isValid()) return fail(2046,"Kein aktueller Datensatz.");
        const auto primary=typed->primaryKey();
        if(primary.isEmpty()) return fail(3251,"Ersetzen erfordert einen gespeicherten Datensatz mit Primärschlüssel.");
        const auto before=typed->record(currentIndex().row());
        QMap<QString,QVariant> expectedKey;
        for(int field=0;field<primary.count();++field) {
            const auto name=primary.fieldName(field); const auto value=before.value(name);
            if(!value.isValid() || value.isNull()) return fail(3251,"Den neuen Datensatz zuerst speichern.");
            expectedKey.insert(name,value);
        }
        const auto tableName=typed->tableName(),fieldName=before.fieldName(arguments.column);
        if(recordNavigationGuard_) {
            const auto guard=recordNavigationGuard_; if(!guard()) { if(self && typed) findMessage_=typed->lastError().text(); return false; }
        } else if(typed->isDirty()) return fail(3251,"Zuerst bestehende Änderungen speichern.");
        if(!self || !typed) return false;
        if(model()!=typed || typed->isReadOnly() || !currentIndex().isValid() || isColumnHidden(arguments.column) || !typed->isTextColumn(arguments.column)) return fail(2046,"Kein bearbeitbares aktuelles Feld.");
        const auto after=typed->record(currentIndex().row());
        if(typed->tableName()!=tableName || after.fieldName(arguments.column)!=fieldName)
            return fail(2501,"Die Datenquelle wurde während der Validierung geändert.");
        for(auto field=expectedKey.cbegin();field!=expectedKey.cend();++field)
            if(after.value(field.key())!=field.value()) return fail(2501,"Der aktuelle Datensatz wurde während der Validierung geändert.");
        const QPersistentModelIndex target=typed->index(currentIndex().row(),arguments.column);
        if(!target.isValid() || !(typed->flags(target)&Qt::ItemIsEditable)) return fail(2046,"Das Feld kann nicht bearbeitet werden.");
        const auto current=target.data(Qt::EditRole);
        NativeTextReplacement output;
        try {if(!current.isNull()) output=replaceNativeText(current.toString(),arguments.text,replacement,arguments.mode,arguments.sensitive ? Qt::CaseSensitive : Qt::CaseInsensitive,false);}
        catch(const VbaRuntimeError& error) {return fail(error.number,QString::fromUtf8(error.what()));}
        if(!output.matched) {findMessage_="Der aktuelle Feldwert stimmt nicht mit dem Suchwert überein.";return false;}
        const bool changed=typed->setData(target,output.value,Qt::EditRole);
        if(!self || !typed) return false;
        if(!changed) { findMessage_=typed->lastError().text(); return false; }
        findArguments_=std::move(arguments); findArguments_.origin=0; findAdvance_=true;
        findMessage_="Feldwert ersetzt.";
        return true;
    }
    bool replaceAllText(NativeTableFindArguments arguments,QString replacement) {
        const QPointer<AccessTableView> self=this;
        QPointer<AccessTableModel> typed=qobject_cast<AccessTableModel*>(model());
        const auto fail=[&](int number,const QString& message) { if(self) { findMessage_=message; (void)refuse(number,message); } return false; };
        if(!typed) return fail(2046,"Kein natives Datenblatt zum Ersetzen.");
        if(replacing_ || finding_) return fail(3251,"Eine Such- oder Ersetzungsaktion läuft bereits.");
        replacing_=true; const auto finished=qScopeGuard([self] { if(self) self->replacing_=false; });
        if(typed->isReadOnly()) return fail(2046,"Das Datenblatt ist schreibgeschützt.");
        if(arguments.formatted || arguments.wildcards || arguments.value || arguments.mode<0 || arguments.mode>2 || arguments.direction!=2 || arguments.column<0 || arguments.column>=typed->columnCount()
            || isColumnHidden(arguments.column) || !typed->isTextColumn(arguments.column))
            return fail(3251,"Alle ersetzen erfordert ein Textfeld, einen gültigen Vergleichsmodus und die Suchrichtung Alle.");
        if(arguments.text.isEmpty()) return fail(5,"Der Suchwert darf nicht leer sein.");
        try {(void)replaceNativeText({},arguments.text,replacement,arguments.mode,arguments.sensitive ? Qt::CaseSensitive : Qt::CaseInsensitive,false);}catch(const VbaRuntimeError& error) {return fail(error.number,QString::fromUtf8(error.what()));}
        if(!finishEditing() || !self || !typed) return false;
        const auto table=typed->tableName(),field=typed->record().fieldName(arguments.column);
        const auto primary=typed->primaryKey(); QMap<QString,QVariant> key;
        if(currentIndex().isValid()) for(int i=0;i<primary.count();++i) key.insert(primary.fieldName(i),typed->record(currentIndex().row()).value(primary.fieldName(i)));
        if(recordNavigationGuard_) {
            const auto guard=recordNavigationGuard_; if(!guard()) { if(self && typed) findMessage_=typed->lastError().text(); return false; }
        } else if(typed->isDirty()) return fail(2501,"Zuerst bestehende Änderungen speichern.");
        if(!self || !typed) return false;
        if(model()!=typed || typed->tableName()!=table || typed->record().fieldName(arguments.column)!=field || isColumnHidden(arguments.column))
            return fail(2501,"Die Datenquelle wurde während der Validierung geändert.");
        for(auto it=key.cbegin();it!=key.cend();++it)
            if(!currentIndex().isValid() || typed->record(currentIndex().row()).value(it.key())!=it.value())
                return fail(2501,"Der aktuelle Datensatz wurde während der Validierung geändert.");
        int count=0;
        if(!typed->replaceAllLiteralText(arguments.column,arguments.text,replacement,arguments.mode,arguments.sensitive ? Qt::CaseSensitive : Qt::CaseInsensitive,&count)) {
            if(self && typed) findMessage_=typed->lastError().text(); return false;
        }
        if(!self || !typed) return false;
        bool restored=key.isEmpty();
        for(int row=0;!restored && row<1000000;++row) {
            if(row>=typed->rowCount()) { if(typed->canFetchMore({})) typed->fetchMore({}); else break; }
            if(row>=typed->rowCount()) break;
            const auto record=typed->record(row); bool same=true;
            for(auto it=key.cbegin();it!=key.cend();++it) same&=record.value(it.key())==it.value();
            if(same) { setCurrentIndex(typed->index(row,arguments.column)); scrollTo(currentIndex()); restored=true; }
        }
        findArguments_=std::move(arguments); findArguments_.origin=0; findAdvance_=false;
        findMessage_=count ? QString::number(count)+" Feldwerte ersetzt und gespeichert." : QStringLiteral("Keine passenden Feldwerte gefunden.");
        if(!restored) findMessage_+=" Der zuvor ausgewählte Datensatz ist in dieser Ansicht nicht mehr sichtbar.";
        return true;
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
        try {
        if(!typed) return fail(2046,"Keine native Tabelle für die Suche.");
        if(finding_ || replacing_) return fail(3251,"Eine Such- oder Ersetzungsaktion läuft bereits.");
        finding_=true; const auto searching=qScopeGuard([self] { if(self) self->finding_=false; });
        if(arguments.text.isEmpty() && !arguments.value) return fail(2046,"Keine vorherige Suche vorhanden.");
        if(arguments.formatted && (arguments.text.size()>16*1024*1024 || arguments.text.contains(QChar(0)) || !arguments.text.isValidUtf16())) return fail(3251,"Formatierter Suchtext ist nicht darstellbar.");
        if(arguments.formatted && arguments.value) return fail(3251,"Formatierte Suche benötigt wörtlichen Text.");
        if(arguments.value) {
            arguments.value=normalizeNativeFindValue(*arguments.value);
            if(arguments.mode!=1) return fail(3251,"Typisierte Suche unterstützt derzeit nur ganze Feldwerte.");
        }
        if(arguments.mode<0 || arguments.mode>2 || arguments.column< -1 || arguments.column>=typed->columnCount() || arguments.origin<0 || arguments.origin>2 || arguments.direction<0 || arguments.direction>2) return fail(5,"Ungültige Sucheinstellung.");
        if(arguments.wildcards) {if(arguments.value) return fail(3251,"Platzhaltersuche benötigt String-Eingabe.");(void)nativeFindTextMatches({},arguments.text,arguments.mode,arguments.sensitive,true);}
        if(!finishEditing()) { if(self && typed) findMessage_=typed->lastError().text(); return -1; }
        if(!self || !typed) return -1;
        if(recordNavigationGuard_) {
            const auto guard=recordNavigationGuard_;
            if(!guard()) { if(self && typed) findMessage_=typed->lastError().text(); return -1; }
        } else if(typed->isDirty()) return fail(3251,"Zuerst Tabellenentwürfe vor der Suche speichern.");
        if(!self || !typed) return -1;
        QVector<int> columns;
        for(int column=0;column<typed->columnCount();++column)
            if((arguments.column<0 || arguments.column==column) && !isColumnHidden(column) && (arguments.formatted ? typed->isSearchableColumn(column) && nativeFindFieldKind(typed->fieldSqlType(column))!=NativeFindFieldKind::Boolean : arguments.value ? nativeFindSupports(*arguments.value,typed->fieldSqlType(column)) : typed->isTextColumn(column))) columns.push_back(column);
        if(columns.isEmpty()) return fail(3265,"Das Suchfeld ist nicht verfügbar.");
        constexpr qint64 cellLimit=1000000;
        while(typed->canFetchMore()) {
            if(static_cast<qint64>(typed->rowCount())*columns.size()>=cellLimit) return fail(3251,"Suchumfang überschreitet die native Grenze.");
            const int before=typed->rowCount(); typed->fetchMore();
            if(!self || !typed) return -1;
            if(typed->rowCount()==before && typed->canFetchMore()) return fail(3251,"Suchdaten konnten nicht vollständig geladen werden.");
        }
        if(static_cast<qint64>(typed->rowCount())*columns.size()>cellLimit) return fail(3251,"Suchumfang überschreitet die native Grenze.");
        if(findArguments_.text!=arguments.text || findArguments_.column!=arguments.column || findArguments_.mode!=arguments.mode || findArguments_.sensitive!=arguments.sensitive || findArguments_.direction!=arguments.direction || findArguments_.formatted!=arguments.formatted || findArguments_.wildcards!=arguments.wildcards || !sameNativeFindValue(findArguments_.value,arguments.value)) findAdvance_=false;
        findArguments_=arguments; findArguments_.origin=0;
        const int count=typed->rowCount();
        if(!count) { findMessage_="Keine Datensätze vorhanden."; return 0; }
        const int step=arguments.direction==0 ? -1 : 1;
        const int start=arguments.origin==1 ? 0 : currentIndex().isValid() ? currentIndex().row()+((arguments.origin==2 || findAdvance_) ? step : 0) : 0;
        const auto cs=arguments.sensitive ? Qt::CaseSensitive : Qt::CaseInsensitive;
        QElapsedTimer clock;clock.start();
        for(int offset=0;offset<count;++offset) {
            if((arguments.formatted || arguments.wildcards) && clock.hasExpired(5000)) return fail(3035,"Formatierte Suche überschreitet ihr Zeitlimit.");
            const qint64 candidateRow=static_cast<qint64>(start)+static_cast<qint64>(step)*offset;
            if(arguments.direction!=2 && (candidateRow<0 || candidateRow>=count)) break;
            const int row=static_cast<int>((candidateRow%count+count)%count); if(isRowHidden(row)) continue;
            for(int column:columns) {
                if((arguments.formatted || arguments.wildcards) && clock.hasExpired(5000)) return fail(3035,"Formatierte Suche überschreitet ihr Zeitlimit.");
                const auto candidate=typed->index(row,column);
                const auto value=candidate.data(Qt::EditRole);
                if(value.isNull() || value.metaType().id()==QMetaType::QByteArray) continue;
                if(arguments.formatted) (void)typed->formattedText(candidate); // Validate actual type/mask; never search error placeholders.
                const auto text=arguments.formatted ? nativeVisibleFindText(*this,candidate) : value.toString();
                const bool matches=arguments.value ? nativeFindEqual(qtFieldValue(value,typed->fieldSqlType(column)),*arguments.value)
                    : nativeFindTextMatches(text,arguments.text,arguments.mode,arguments.sensitive,arguments.wildcards);
                if(matches) {
                    setCurrentIndex(candidate);
                    if(!self || !typed) return -1;
                    scrollTo(candidate); findAdvance_=true;
                    findMessage_=QString("Übereinstimmung in Datensatz %1.").arg(row+1); return 1;
                }
            }
        }
        findMessage_="Keine Übereinstimmung gefunden."; return 0;
        } catch(const VbaRuntimeError& error) { return fail(error.number,QString::fromUtf8(error.what())); }
        catch(const std::exception& error) { return fail(3146,QString::fromUtf8(error.what())); }
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
        const QPointer<AccessTableView> self=this;
        QPointer<AccessTableModel> typed=qobject_cast<AccessTableModel*>(model());
        if(!typed || !finishEditing() || !self || !typed) return false;
        while(typed->canFetchMore()) {
            if(typed->rowCount()>=1000000) return refuse(3251,"Die neue Zeile überschreitet die native Datensatzgrenze.");
            const auto count=typed->rowCount(); typed->fetchMore();
            if(!self || !typed) return false;
            if(typed->rowCount()==count) return refuse(3251,"Die vorhandenen Datensätze konnten nicht vollständig geladen werden.");
        }
        if(typed->rowCount()>=1000000) return refuse(3251,"Die neue Zeile überschreitet die native Datensatzgrenze.");
        const auto row=typed->rowCount();
        if(!typed->insertRow(row)) return false;
        setCurrentIndex(typed->index(row,0)); scrollTo(currentIndex()); setFocus();
        return true;
    }
    bool deleteCurrentRecord() {
        const QPointer<AccessTableView> self=this;
        QPointer<AccessTableModel> typed=qobject_cast<AccessTableModel*>(model());
        if(!typed || !currentIndex().isValid()) return false;
        if(deletingRecords_) return refuse(2501,"Eine Datensatzlöschung läuft bereits.");
        deletingRecords_=true; const auto finished=qScopeGuard([self] { if(self) self->deletingRecords_=false; });
        const auto selected=selectionModel() ? selectionModel()->selectedRows() : QModelIndexList{};
        if(selected.size()<=1) {
            if(selectionModel()) for(const auto& index:selectionModel()->selectedIndexes())
                if(index.row()!=currentIndex().row()) return refuse(3251,"Zum Löschen mehrerer Datensätze vollständige Zeilen auswählen.");
            if(!finishEditing() || !self || !typed || model()!=typed || !currentIndex().isValid()) return false;
            const auto row=currentIndex().row(),column=currentIndex().column();
            if(!typed->removeRow(row) || !self || !typed) return false;
            if(typed->rowCount()) setCurrentIndex(typed->index(qMin(row,typed->rowCount()-1),column));
            return true;
        }
        if(selected.size()>1000000 || typed->editStrategy()!=QSqlTableModel::OnManualSubmit || typed->primaryKey().isEmpty())
            return refuse(3251,"Mehrfachlöschung erfordert gespeicherte Datensätze mit Primärschlüssel.");
        QSet<int> rows; QVector<QPersistentModelIndex> anchors;
        const auto primary=typed->primaryKey(); const auto source=typed->tableName();
        for(const auto& index:selected) {
            rows.insert(index.row()); anchors.append(index);
            const auto record=typed->record(index.row());
            for(int field=0;field<primary.count();++field)
                if(record.value(primary.fieldName(field)).isNull()) return refuse(3251,"Neue Datensätze zuerst speichern oder verwerfen.");
        }
        if(selectionModel()) for(const auto& index:selectionModel()->selectedIndexes())
            if(!rows.contains(index.row())) return refuse(3251,"Zum Löschen vollständige Zeilen ohne zusätzliche Zellbereiche auswählen.");
        if(!finishEditing() || !self || !typed) return false;
        if(model()!=typed || typed->tableName()!=source) return refuse(2501,"Die Datenquelle wurde während der Validierung geändert.");
        if(typed->isDirty()) return refuse(2501,"Vor der Mehrfachlöschung bestehende Änderungen speichern oder verwerfen.");
        for(const auto& anchor:anchors) if(!anchor.isValid() || anchor.model()!=typed) return refuse(2501,"Die Datensatzauswahl wurde während der Validierung geändert.");
        std::sort(anchors.begin(),anchors.end(),[](const auto& left,const auto& right) { return left.row()>right.row(); });
        const int column=currentIndex().column(),first=anchors.back().row();
        for(const auto& anchor:anchors) {
            const bool removed=typed->removeRow(anchor.row());
            if(!self || !typed) return false;
            if(removed) continue;
            const auto problem=typed->lastError();
            for(const auto& staged:anchors) {
                if(!self || !typed) return false;
                if(staged.isValid() && typed->hasPendingRecord(staged.row())) (void)typed->undoPendingRecord(staged.row());
            }
            if(!self || !typed) return false;
            bool valid=false; const int number=problem.nativeErrorCode().toInt(&valid);
            return typed->rejectEdit(problem.driverText()=="LinAccess" && valid ? number : 2501,problem.text());
        }
        if(!self || !typed) return false;
        clearSelection();
        if(typed->rowCount()) setCurrentIndex(typed->index(qMin(first,typed->rowCount()-1),qMax(0,column)));
        return true;
    }
    bool undoCurrentRecord() {
        const QPointer<AccessTableView> self=this;
        QPointer<AccessTableModel> typed=qobject_cast<AccessTableModel*>(model());
        if(!typed || !currentIndex().isValid()) return refuse(2046,"Kein aktueller Datensatz.");
        if(activeEditor()) return refuse(2501,"Zuerst die aktuelle Zelleingabe mit Esc abbrechen.");
        const int row=currentIndex().row(),column=currentIndex().column();
        const QPersistentModelIndex original=currentIndex();
        if(!typed->undoPendingRecord(row) || !self || !typed) return false;
        if(model()!=typed) return false;
        if(original.isValid()) setCurrentIndex(original);
        else if(typed->rowCount() && typed->columnCount()) setCurrentIndex(typed->index(qMin(row,typed->rowCount()-1),qMin(column,typed->columnCount()-1)));
        return true;
    }
    bool moveToRecordBoundary(bool last) {
        const QPointer<AccessTableView> self=this;
        QPointer<AccessTableModel> typed=qobject_cast<AccessTableModel*>(model());
        if(!typed || !recordNavigationHandler_ || !finishEditing() || !self || !typed) return false;
        const bool sameRecord=currentIndex().isValid() && (last
            ? currentIndex().row()==typed->rowCount()-1 && !typed->canFetchMore()
            : currentIndex().row()==0);
        if(!sameRecord) {
            const auto navigate=recordNavigationHandler_;
            if(!navigate(last ? RecordMove::Last : RecordMove::First,1) || !self || !typed) return false;
        }
        if(model()!=typed || !currentIndex().isValid()) return false;
        auto* header=horizontalHeader(); int column=-1;
        for(int offset=0;offset<header->count();++offset) {
            const int visual=last ? header->count()-1-offset : offset;
            const int logical=header->logicalIndex(visual);
            if(logical>=0 && !isColumnHidden(logical)) { column=logical; break; }
        }
        if(column<0) return refuse(3265,"Kein sichtbares Feld verfügbar.");
        setCurrentIndex(typed->index(currentIndex().row(),column));
        if(!self || !typed || model()!=typed) return false;
        scrollTo(currentIndex()); setFocus();
        return true;
    }
    bool moveByRecordPage(bool down) {
        const QPointer<AccessTableView> self=this;
        QPointer<AccessTableModel> typed=qobject_cast<AccessTableModel*>(model());
        if(!typed || !recordNavigationHandler_ || !finishEditing() || !self || !typed || !currentIndex().isValid()) return false;
        const auto origin=currentIndex();
        const auto candidate=QTableView::moveCursor(down ? MovePageDown : MovePageUp,Qt::NoModifier);
        if(!self || !typed || model()!=typed) return false;
        if(!candidate.isValid() || candidate.model()!=typed) return refuse(2105,"Keine verfügbare Datensatzseite.");
        const int distance=down ? candidate.row()-origin.row() : origin.row()-candidate.row();
        if(!distance) {
            if(down && typed->canFetchMore()) return refuse(3251,"Seitenwechsel über den geladenen Bereich hinaus ist noch nicht implementiert.");
            return true;
        }
        if(distance<0) return refuse(3251,"Die vorgeschlagene Seitenrichtung wird nicht unterstützt.");
        const auto navigate=recordNavigationHandler_;
        const bool moved=navigate(down ? RecordMove::Next : RecordMove::Previous,distance);
        if(!self || !typed || model()!=typed) return false;
        if(currentIndex().isValid()) scrollTo(currentIndex());
        return moved;
    }
    bool undoPastedBlock(bool redo=false) {
        if(activeEditor()) return refuse(2501,"Zuerst die aktuelle Zelleingabe abschließen.");
        const QPointer<AccessTableView> self=this;QPointer<AccessTableModel> typed=qobject_cast<AccessTableModel*>(model());if(!typed) return false;
        const int row=currentIndex().row(),column=currentIndex().column();
        const bool accepted=redo ? typed->redoBlock() : typed->undoBlock();
        if(accepted && self && typed && model()==typed && typed->rowCount() && typed->columnCount()) {
            const int target=column>=0 ? qMin(column,typed->columnCount()-1) : qMax(0,firstPasteColumn(typed));
            setCurrentIndex(typed->index(qBound(0,row,typed->rowCount()-1),target));
        }
        return accepted;
    }
    bool canCutSelection() const {
        const auto* typed=qobject_cast<AccessTableModel*>(model());
        if(!typed || typed->isReadOnly() || typed->isDirty() || typed->editStrategy()!=QSqlTableModel::OnManualSubmit || hasPendingEditor() || !selectionModel() || !selectionModel()->hasSelection()) return false;
        for(const auto& range:selectionModel()->selection()) for(int column=range.left();column<=range.right();++column)
            if(isColumnHidden(column) || !typed->isPasteColumn(column,true)) return false;
        return true;
    }
    bool cutSelection() {
        const QPointer<AccessTableView> self=this;QPointer<AccessTableModel> typed=qobject_cast<AccessTableModel*>(model());
        if(!typed || !finishEditing() || !self || !typed) return false;
        if(typed->isReadOnly()) return refuse(2448,"Die Datenquelle ist schreibgeschützt.");
        if(typed->isDirty()) return refuse(2501,"Vor dem Ausschneiden eines Bereichs Änderungen speichern oder verwerfen.");
        if(typed->editStrategy()!=QSqlTableModel::OnManualSubmit) return refuse(3251,"Bereichsausschneiden ist in dieser Ansicht nicht verfügbar.");
        QVector<QVector<QVariant>> cleared;QVector<QVariant> cells;int previous=-1;QPersistentModelIndex first;
        try {
            (void)nativeGridClipboardText(*this,[&](const QModelIndex& index) {
                if(!typed->isPasteColumn(index.column(),true)) throw VbaRuntimeError(2448,"Cut target is protected");
                if(!first.isValid()) first=index;
                if(previous>=0 && previous!=index.row()) {cleared.push_back(cells);cells.clear();}
                previous=index.row();cells.push_back(QVariant{});return QString{};
            });
            if(!cells.isEmpty()) cleared.push_back(cells);
        } catch(const VbaRuntimeError& error) {return refuse(error.number,"Die Auswahl kann nicht als zusammenhängender Bereich ausgeschnitten werden.");}
          catch(const std::bad_alloc&) {return refuse(7,"Der Bereich konnte nicht im Speicher vorbereitet werden.");}
        if(!copySelection() || !self || !typed || !first.isValid()) return false;
        return typed->setTypedDataBlock(first.row(),first.column(),cleared);
    }
    bool copySelection() {
        auto* typed=qobject_cast<AccessTableModel*>(model());
        if(!typed || !selectionModel() || !finishEditing()) return false;
        try {
            QJsonArray rows,cells,fields;int previous=-1;
            const auto output=nativeGridClipboardText(*this,[&](const QModelIndex& index) {
                if(previous>=0 && previous!=index.row()) {rows.append(cells);cells={};}
                if(rows.isEmpty() && (previous<0 || previous==index.row())) fields.append(typed->record().fieldName(index.column()));
                previous=index.row();const auto raw=typed->QSqlTableModel::data(index,Qt::EditRole);
                cells.append(nativeGridEncode(raw));
                if(raw.metaType().id()==QMetaType::QByteArray) return QStringLiteral("0x")+QString::fromLatin1(raw.toByteArray().toHex());
                const auto type=typed->fieldSqlType(index.column());
                if(raw.metaType().id()==QMetaType::LongLong && (type=="INTEGER" || type=="INT" || type=="BIGINT")) return QString::number(raw.toLongLong());
                return typed->clipboardText(index);
            });
            if(!cells.isEmpty()) rows.append(cells);
            const auto data=QJsonDocument(QJsonObject{{"version",1},{"rows",rows},{"fields",fields}}).toJson(QJsonDocument::Compact);
            if(data.size()>nativeGridMimeLimit) throw VbaRuntimeError(3035,"Typed grid clipboard size limit exceeded");
            auto* mime=new QMimeData;mime->setText(output);mime->setData(nativeGridMime,data);QApplication::clipboard()->setMimeData(mime);
        } catch(const VbaRuntimeError& error) { return refuse(error.number,"Die Auswahl kann nicht verlustfrei kopiert werden."); }
          catch(const std::bad_alloc&) {return refuse(7,"Die Auswahl konnte nicht im Speicher vorbereitet werden.");}
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
    bool selectAllRecords() {
        QPointer<AccessTableView> self=this;
        if(!finishEditing() || !self) return false;
        const QPersistentModelIndex current=currentIndex();
        if(!fetchSelectionRows() || !self || !selectionModel()) return false;
        selectAll();
        if(self && current.isValid() && selectionModel())
            selectionModel()->setCurrentIndex(current,QItemSelectionModel::NoUpdate);
        return true;
    }
    bool canPasteSelection() const {
        auto* typed=qobject_cast<AccessTableModel*>(model());
        const bool binary=QApplication::clipboard()->mimeData() && QApplication::clipboard()->mimeData()->hasFormat(nativeGridMime);
        return typed && typed->editStrategy()==QSqlTableModel::OnManualSubmit && typed->canPasteOverUpdates()
            && !hasPendingEditor() && ((currentIndex().isValid() && (currentIndex().flags().testFlag(Qt::ItemIsEditable) || (binary && typed->isPasteColumn(currentIndex().column(),true))))
                || (typed->rowCount()==0 && firstPasteColumn(typed,binary)>=0));
    }
    bool canAppendRecords() const {
        const auto* typed=qobject_cast<AccessTableModel*>(model());const auto* mime=QApplication::clipboard()->mimeData();
        if(!typed || typed->isReadOnly() || typed->editStrategy()!=QSqlTableModel::OnManualSubmit || !typed->canPasteOverUpdates() || hasPendingEditor()) return false;
        if(canPasteSelection()) return true;
        return mime && mime->hasFormat(nativeGridMime) && mime->data(nativeGridMime).size()<=nativeGridMimeLimit && QJsonDocument::fromJson(mime->data(nativeGridMime)).object().value("fields").toArray().size()==typed->columnCount();
    }
    int firstPasteColumn(const AccessTableModel* typed,bool binary=false) const {
        for(int column=0;column<typed->columnCount();++column) if(!isColumnHidden(column) && typed->isPasteColumn(column,binary)) return column;
        return -1;
    }
    bool pasteSelection(bool append=false) {
        auto* typed=qobject_cast<AccessTableModel*>(model());
        if(!typed || !finishEditing()) return false;
        if(!typed->canPasteOverUpdates()) return refuse(3251,"Der vorhandene Entwurfsumfang überschreitet die native Bereichsgrenze.");
        if(typed->editStrategy()!=QSqlTableModel::OnManualSubmit) return refuse(3251,"Bereichseinfügen ist in dieser Ansicht nicht verfügbar.");
        const auto* mime=QApplication::clipboard()->mimeData();
        const bool typedMime=mime && mime->hasFormat(nativeGridMime);
        if(typedMime && mime->data(nativeGridMime).size()>nativeGridMimeLimit) return refuse(3035,"Die typisierte Zwischenablage überschreitet die native Größenbegrenzung.");
        QStringList recordFields;bool fullRecord=false;
        if(append && typedMime) {
            const auto fields=QJsonDocument::fromJson(mime->data(nativeGridMime)).object().value("fields");
            if(fields.isArray() && fields.toArray().size()==typed->columnCount()) {
                fullRecord=true;for(const auto& value:fields.toArray()) {if(!value.isString()) return refuse(13,"Ungültige Datensatzfelder in der Zwischenablage.");recordFields<<value.toString();}
                for(int column=0;column<recordFields.size();++column) if(recordFields[column].compare(typed->record().fieldName(column),Qt::CaseInsensitive)!=0) return refuse(13,"Die Datensatzfelder stimmen nicht mit der Zieltabelle überein.");
            }
        }
        const auto current=currentIndex();
        const int firstColumn=fullRecord ? 0 : current.isValid() ? current.column() : firstPasteColumn(typed,typedMime);
        if(firstColumn<0 || (!fullRecord && !typed->isPasteColumn(firstColumn,typedMime))) return refuse(2448,"Die Zielspalte ist schreibgeschützt.");
        if(!append && current.isValid() && !current.flags().testFlag(Qt::ItemIsEditable) && !(typedMime && typed->isPasteColumn(firstColumn,true))) return refuse(2448,"Die Zielzelle ist schreibgeschützt.");
        if(append) while(typed->canFetchMore()) {
            if(typed->rowCount()>=1000000) return refuse(3251,"Die Tabelle überschreitet die native Anfügegrenze.");
            const int previous=typed->rowCount();typed->fetchMore();if(previous==typed->rowCount()) return refuse(3251,"Die vorhandenen Datensätze konnten nicht vollständig geladen werden.");
        }
        const int firstRow=append ? typed->rowCount() : current.isValid() ? current.row() : 0;
        if(!mime || (!mime->hasText() && !typedMime)) return refuse(2046,"Die Zwischenablage enthält keinen einfügbaren Text.");
        if(mime->hasFormat("application/x-linaccess-record-v1")) return refuse(3251,"Typisierte Formulardatensätze können hier noch nicht eingefügt werden.");
        const auto text=QApplication::clipboard()->text();
        if(text.size()>4*1024*1024 || text.contains(QChar(0))) return refuse(3251,"Dieser Zwischenablageumfang wird noch nicht unterstützt.");
        QVector<QVector<QString>> values;QVector<QVector<QVariant>> typedValues;
        try {
            if(typedMime) typedValues=nativeGridDecodeBlock(mime->data(nativeGridMime));
            else {
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
            }
        } catch(const VbaRuntimeError& error) {return refuse(error.number,"Die typisierte Zwischenablage ist ungültig.");}
          catch(const std::bad_alloc&) { return refuse(7,"Der Tabellenbereich konnte nicht im Speicher vorbereitet werden."); }
          catch(const std::length_error&) { return refuse(3251,"Die Zwischenablage enthält zu viele Zellen."); }
          catch(const std::exception&) { return refuse(5,"Die Zwischenablage enthält kein gültiges Tabellenformat."); }
        const auto rows=typedMime ? typedValues.size() : values.size();
        const auto columns=typedMime ? typedValues.front().size() : values.front().size();
        const auto endRow=static_cast<qint64>(firstRow)+rows;
        if(endRow>std::numeric_limits<int>::max() || static_cast<qint64>(firstColumn)+columns>typed->columnCount())
            return refuse(5,"Der einzufügende Bereich passt nicht in die vorhandenen Spalten.");
        while(typed->rowCount()<endRow && typed->canFetchMore()) {
            const auto count=typed->rowCount(); typed->fetchMore(); if(typed->rowCount()==count) break;
        }
        for(int column=firstColumn;column<firstColumn+columns;++column) if(isColumnHidden(column)) return refuse(3251,"Verdeckte Zielspalten werden nicht überschrieben.");
        for(int row=firstRow;row<qMin<qint64>(endRow,typed->rowCount());++row) if(isRowHidden(row)) return refuse(3251,"Verdeckte Zieldatensätze werden nicht überschrieben.");
        if(!(fullRecord ? typed->setTypedRecordBlock(firstRow,recordFields,typedValues) : typedMime ? typed->setTypedDataBlock(firstRow,firstColumn,typedValues) : typed->setDataBlock(firstRow,firstColumn,values))) return false;
        const auto first=typed->index(firstRow,firstColumn);
        setCurrentIndex(first);
        if(selectionModel()) selectionModel()->select(QItemSelection(first,typed->index(endRow-1,firstColumn+columns-1)),QItemSelectionModel::ClearAndSelect);
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
        auto* native=qobject_cast<AccessTableModel*>(model());
        auto* copy=add("copy","&Kopieren\tStrg+C",native && selectionModel()->hasSelection());
        auto* cut=add("cut","&Ausschneiden\tStrg+X",canCutSelection());
        const auto* clipboard=QApplication::clipboard()->mimeData();
        auto* paste=add("paste","&Einfügen\tStrg+V",canPasteSelection() && clipboard && clipboard->hasText());
        auto* append=add("paste-append","Datensätze &anfügen",canAppendRecords() && clipboard && clipboard->hasText());
        menu.addSeparator();
        auto* zoom=add("zoom","&Zoom...\tUmschalt+F2",native && currentIndex().isValid());
        auto* undoRecord=add("undo-record","Datensatzänderungen &verwerfen\tEsc",native && currentIndex().isValid() && !activeEditor() && native->hasPendingRecord(currentIndex().row()));
        menu.addSeparator();
        auto* row=add("select-row","&Datensatz markieren\tUmschalt+Leertaste",currentIndex().isValid());
        auto* column=add("select-column","&Spalte markieren\tStrg+Leertaste",currentIndex().isValid());
        auto* all=add("select-all","&Alles markieren\tStrg+A",model()->rowCount()>0 && model()->columnCount()>0);
        const auto position=event->reason()==QContextMenuEvent::Keyboard && currentIndex().isValid()
            ? viewport()->mapToGlobal(visualRect(currentIndex()).center()) : event->globalPos();
        auto* chosen=menu.exec(position);
        if(!self) return;
        if(chosen==copy) (void)copySelection();
        else if(chosen==cut) (void)cutSelection();
        else if(chosen==paste) (void)pasteSelection();
        else if(chosen==append) (void)pasteSelection(true);
        else if(chosen==zoom) (void)zoomCurrentCell();
        else if(chosen==undoRecord) (void)undoCurrentRecord();
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
    QModelIndex moveCursor(CursorAction action,Qt::KeyboardModifiers modifiers) override {
        const auto origin=currentIndex();
        const auto candidate=QTableView::moveCursor(action,modifiers);
        const auto keys=modifiers & ~Qt::KeypadModifier;
        if(movingRecord_ || !recordNavigationHandler_ || (action!=MoveNext && action!=MovePrevious)
            || (keys!=Qt::NoModifier && keys!=Qt::ShiftModifier) || !origin.isValid() || !candidate.isValid()
            || origin.model()!=model() || candidate.model()!=model()) return candidate;
        const int step=candidate.row()-origin.row();
        if(step!=(action==MoveNext ? 1 : -1)) return candidate;
        const int column=candidate.column();
        const QPointer<AccessTableView> self=this;
        QPointer<AccessTableModel> typed=qobject_cast<AccessTableModel*>(model());
        if(!typed) return candidate;
        movingRecord_=true; const auto finished=qScopeGuard([self] { if(self) self->movingRecord_=false; });
        const auto navigate=recordNavigationHandler_;
        const bool moved=navigate(step<0 ? RecordMove::Previous : RecordMove::Next,1);
        if(!self || !typed || model()!=typed) return {};
        if(moved && currentIndex().isValid() && column>=0 && column<typed->columnCount())
            return typed->index(currentIndex().row(),column);
        return currentIndex();
    }
    void keyPressEvent(QKeyEvent* event) override {
        const auto selectionModifiers=event->modifiers() & ~Qt::KeypadModifier;
        if(event->key()==Qt::Key_Escape && selectionModifiers==Qt::NoModifier && !activeEditor()) {
            auto* typed=qobject_cast<AccessTableModel*>(model());
            if(typed && currentIndex().isValid() && typed->hasPendingRecord(currentIndex().row())) {
                (void)undoCurrentRecord(); event->accept(); return;
            }
        }
        if((event->key()==Qt::Key_Up || event->key()==Qt::Key_Down) && selectionModifiers==Qt::NoModifier
            && recordNavigationHandler_ && model() && currentIndex().isValid()) {
            const int row=currentIndex().row(); const bool up=event->key()==Qt::Key_Up;
            if(up ? row>0 : row+1<model()->rowCount() || model()->canFetchMore({})) {
                const auto navigate=recordNavigationHandler_;
                (void)navigate(up ? RecordMove::Previous : RecordMove::Next,1); event->accept(); return;
            }
        }
        if((event->key()==Qt::Key_PageUp || event->key()==Qt::Key_PageDown) && selectionModifiers==Qt::NoModifier && recordNavigationHandler_) {
            (void)moveByRecordPage(event->key()==Qt::Key_PageDown); event->accept(); return;
        }
        if((event->key()==Qt::Key_Home || event->key()==Qt::Key_End) && selectionModifiers==Qt::ControlModifier && recordNavigationHandler_) {
            (void)moveToRecordBoundary(event->key()==Qt::Key_End); event->accept(); return;
        }
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
        if((event->key()==Qt::Key_Z || event->key()==Qt::Key_Y) && selectionModifiers==Qt::ControlModifier) {
            (void)undoPastedBlock(event->key()==Qt::Key_Y);event->accept();return;
        }
        if((event->key()==Qt::Key_X && selectionModifiers==Qt::ControlModifier) || (event->key()==Qt::Key_Delete && selectionModifiers==Qt::ShiftModifier)) {
            (void)cutSelection();event->accept();return;
        }
        if((event->key()==Qt::Key_C || event->key()==Qt::Key_Insert) && selectionModifiers==Qt::ControlModifier) {
            (void)copySelection(); event->accept(); return;
        }
        if((event->key()==Qt::Key_V && selectionModifiers==Qt::ControlModifier)
            || (event->key()==Qt::Key_Insert && selectionModifiers==Qt::ShiftModifier)) {
            (void)pasteSelection(); event->accept(); return;
        }
        if((event->key()==Qt::Key_Space && (selectionModifiers==Qt::ShiftModifier || selectionModifiers==Qt::ControlModifier))
            || (event->key()==Qt::Key_A && selectionModifiers==Qt::ControlModifier)) {
            if(event->key()==Qt::Key_A) { (void)selectAllRecords(); event->accept(); return; }
            if(finishEditing() && selectionModel() && model()) {
                if(event->key()==Qt::Key_Space && selectionModifiers==Qt::ShiftModifier) {
                    if(currentIndex().isValid()) selectionModel()->select(currentIndex(),QItemSelectionModel::ClearAndSelect|QItemSelectionModel::Rows);
                } else if(fetchSelectionRows()) {
                    if(currentIndex().isValid()) selectionModel()->select(currentIndex(),QItemSelectionModel::ClearAndSelect|QItemSelectionModel::Columns);
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
    std::function<bool(RecordMove,int)> recordNavigationHandler_;
    QPointer<QDialog> findDialog_,replaceDialog_;
    NativeTableFindArguments findArguments_;
    QString findMessage_;
    bool findAdvance_{},finding_{},replacing_{},movingRecord_{},deletingRecords_{};
    QMetaObject::Connection restore_;
};
}
