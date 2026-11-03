#pragma once
#include "access_query.hpp"
#include "native_table_find.hpp"
#include <QTableView>
#include <QScopeGuard>

namespace linaccess {
class NativeQueryView final : public QTableView {
public:
    explicit NativeQueryView(QWidget* parent=nullptr):QTableView(parent) {}
    void setModel(QAbstractItemModel* model) override {
        const QPointer<NativeQueryView> self=this;
        const auto dialog=findDialog_; findDialog_=nullptr;
        if(dialog) dialog->close();
        if(!self) return;
        findArguments_={}; findAdvance_=false; findMessage_.clear();
        QTableView::setModel(model);
    }
    bool hasFindText() const { return !findArguments_.text.isEmpty(); }
    QString findMessage() const { return findMessage_; }
    bool openFindDialog() {
        auto* result=dynamic_cast<AccessQueryModel*>(model());
        if(!result) throw VbaRuntimeError(2046,"Kein natives Abfrageergebnis.");
        if(findDialog_ && findDialog_->isVisible()) { findDialog_->raise(); findDialog_->activateWindow(); return true; }
        QVector<QPair<QString,int>> columns; QMap<int,QString> hints;
        for(int column=0;column<result->columnCount();++column) {
            const auto type=result->fieldSqlType(column);
            if(isColumnHidden(column) || nativeFindFieldKind(type)==NativeFindFieldKind::Unsupported) continue;
            columns.push_back({result->headerData(column,Qt::Horizontal).toString(),column});
            if(!nativeFindTextType(type)) hints.insert(column,nativeFindInputHint(type));
        }
        if(columns.isEmpty()) throw VbaRuntimeError(3251,"Keine Abfragespalten mit nachgewiesenem Suchtyp.");
        auto initial=findArguments_;
        if(initial.text.isEmpty()) {
            initial.column=columns.front().second;
            for(const auto& column:columns) if(column.second==currentIndex().column()) initial.column=column.second;
        }
        const QPointer<NativeQueryView> self=this;
        auto* dialog=nativeTableFindDialog(this,columns,hints,initial,[self](NativeTableFindArguments arguments) {
            if(!self) return QString{};
            try {
                auto* result=dynamic_cast<AccessQueryModel*>(self->model());
                if(!result) throw VbaRuntimeError(2046,"Abfrageergebnis wurde geschlossen.");
                if(arguments.column>=0 && !nativeFindTextType(result->fieldSqlType(arguments.column))) {
                    arguments.value=nativeFindInput(arguments.text,result->fieldSqlType(arguments.column));
                    arguments.text=nativeFindValueText(*arguments.value); arguments.mode=1;
                }
                self->findText(std::move(arguments));
            } catch(const VbaRuntimeError& error) { if(self) self->findMessage_=QString::fromUtf8(error.what())+" ("+QString::number(error.number)+")"; }
            return self ? self->findMessage_ : QString{};
        });
        if(self) findDialog_=dialog;
        return self && dialog;
    }
    int repeatFind() {
        const QPointer<NativeQueryView> self=this;
        const int result=findText(findArguments_);
        if(self && findDialog_) if(auto* label=findDialog_->findChild<QLabel*>("native-table-find-result")) label->setText(findMessage_);
        return result;
    }
    int findText(NativeTableFindArguments arguments) {
        const QPointer<NativeQueryView> self=this;
        const QPointer<AccessQueryModel> result=dynamic_cast<AccessQueryModel*>(model());
        if(!result) throw VbaRuntimeError(2046,"Kein natives Abfrageergebnis.");
        if(finding_) throw VbaRuntimeError(3251,"Abfragesuche wird bereits ausgeführt.");
        finding_=true; const auto searching=qScopeGuard([self] { if(self) self->finding_=false; });
        if(arguments.text.isEmpty() && !arguments.value) throw VbaRuntimeError(2046,"Keine vorherige Suche vorhanden.");
        if(arguments.mode<0 || arguments.mode>2 || arguments.column< -1 || arguments.column>=result->columnCount()
            || arguments.origin<0 || arguments.origin>2 || arguments.direction<0 || arguments.direction>2) throw VbaRuntimeError(5,"Ungültige Sucheinstellung.");
        if(arguments.value) {
            arguments.value=normalizeNativeFindValue(*arguments.value);
            if(arguments.mode!=1) throw VbaRuntimeError(3251,"Typisierte Suche erfordert ganze Feldwerte.");
        }
        QVector<int> columns;
        for(int column=0;column<result->columnCount();++column)
            if((arguments.column<0 || arguments.column==column) && !isColumnHidden(column)
                && (arguments.value ? nativeFindSupports(*arguments.value,result->fieldSqlType(column)) : nativeFindTextType(result->fieldSqlType(column)))) columns.push_back(column);
        if(columns.isEmpty()) throw VbaRuntimeError(3265,"Suchfeld ist nicht verfügbar oder sein Typ nicht nachgewiesen.");
        if(static_cast<qint64>(result->rowCount())*columns.size()>1000000) throw VbaRuntimeError(3251,"Suchumfang überschreitet die native Grenze.");
        if(findArguments_.text!=arguments.text || findArguments_.column!=arguments.column || findArguments_.mode!=arguments.mode
            || findArguments_.sensitive!=arguments.sensitive || findArguments_.direction!=arguments.direction || !sameNativeFindValue(findArguments_.value,arguments.value)) findAdvance_=false;
        findArguments_=arguments; findArguments_.origin=0;
        const int count=result->rowCount();
        if(!count) { findMessage_="Keine Datensätze vorhanden."; return 0; }
        const int step=arguments.direction==0 ? -1 : 1;
        const int start=arguments.origin==1 ? 0 : currentIndex().isValid() ? currentIndex().row()+((arguments.origin==2 || findAdvance_) ? step : 0) : 0;
        const auto cs=arguments.sensitive ? Qt::CaseSensitive : Qt::CaseInsensitive;
        for(int offset=0;offset<count;++offset) {
            const qint64 candidateRow=static_cast<qint64>(start)+static_cast<qint64>(step)*offset;
            if(arguments.direction!=2 && (candidateRow<0 || candidateRow>=count)) break;
            const int row=static_cast<int>((candidateRow%count+count)%count); if(isRowHidden(row)) continue;
            for(int column:columns) {
                const auto candidate=result->index(row,column); const auto value=candidate.data(Qt::EditRole);
                if(value.isNull() || value.metaType().id()==QMetaType::QByteArray) continue;
                const auto text=value.toString();
                const bool matches=arguments.value ? nativeFindEqual(qtFieldValue(value,result->fieldSqlType(column)),*arguments.value)
                    : arguments.mode==1 ? text.compare(arguments.text,cs)==0 : arguments.mode==2 ? text.startsWith(arguments.text,cs) : text.contains(arguments.text,cs);
                if(matches) {
                    setCurrentIndex(candidate);
                    if(!self || !result || self->model()!=result) return 0;
                    scrollTo(candidate); findAdvance_=true; findMessage_=QString("Übereinstimmung in Datensatz %1.").arg(row+1); return 1;
                }
            }
        }
        findMessage_="Keine Übereinstimmung gefunden."; return 0;
    }
private:
    QPointer<QDialog> findDialog_;
    NativeTableFindArguments findArguments_;
    QString findMessage_;
    bool findAdvance_{},finding_{};
};
}
