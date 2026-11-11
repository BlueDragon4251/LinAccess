#include "native_schema_editor.hpp"
#include "access_format.hpp"
#include "native_field_defaults.hpp"
#include "dao_database.hpp"
#include "access_table_model.hpp"
#include "access_autonumber.hpp"
#include "vba_qt_value.hpp"
#include "native_find_value.hpp"
#include "native_text_replace.hpp"
#include "native_grid_mime.hpp"
#include "vba_currency.hpp"
#include <QSqlError>
#include <QElapsedTimer>
#include <QSqlRecord>
#include <QSqlIndex>
#include <limits>
#include <QSignalBlocker>
#include <QScopedValueRollback>
#include <QSqlQuery>
#include <QSqlDriver>
#include <algorithm>
namespace linaccess {
namespace {
bool floatingDisplay(const QString& type) { return type!="CURRENCY_REAL" && (type.contains("REAL",Qt::CaseInsensitive) || type.contains("FLOA",Qt::CaseInsensitive) || type.contains("DOUB",Qt::CaseInsensitive)); }
bool checkedDisplay(const QString& type) { return type=="DECIMAL_EXACT" || type=="CURRENCY_TICKS" || type=="CURRENCY_REAL" || type=="INTEGER1" || type=="INTEGER2" || type=="INTEGER4" || floatingDisplay(type); }
QVariant checkedInput(const QString& type,const QVariant& input) {
    const bool emptyNumeric=input.metaType().id()==QMetaType::QString && input.toString().isEmpty()
        && (type=="CURRENCY_REAL" || type=="CURRENCY_TICKS" || type.contains("INT",Qt::CaseInsensitive) || type.contains("REAL",Qt::CaseInsensitive) || type.contains("DATE",Qt::CaseInsensitive));
    const auto scalar=emptyNumeric ? VbaValue{} : qtFieldValue(input,(type.contains("BOOL",Qt::CaseInsensitive) || type.contains("DATE",Qt::CaseInsensitive))?type:QString{});
    return sqlFieldValue(scalar,type);
}
QMap<QString,QVariant> pendingScalarDefaults(QSqlDatabase db,const QString& table,const QMap<QString,QString>& types) {
    auto result=nativeTableDateDefaults(db,table);
    if(types.values().contains("DECIMAL_EXACT")) {
        const auto defaults=nativeTableLiteralDefaults(db,table);
        for(auto it=defaults.cbegin();it!=defaults.cend();++it) if(types.value(it.key())=="DECIMAL_EXACT") result.insert(it.key(),it.value());
    }
    return result;
}
QString editMessage(int number) {
    if(number==6) return QStringLiteral("Zahlenwert außerhalb des zulässigen Bereichs.");
    if(number==13) return QStringLiteral("Der Wert passt nicht zum Spaltentyp.");
    if(number==2501) return QStringLiteral("Zuerst die Zelleingabe abschließen und Änderungen speichern oder verwerfen.");
    if(number==2448) return QStringLiteral("Dieses Feld ist schreibgeschützt.");
    if(number==3317) return QStringLiteral("Der Wert verletzt die Gültigkeitsregel des Feldes.");
    if(number==3251) return QStringLiteral("Dieser Wert wird für den Spaltentyp noch nicht unterstützt.");
    return QStringLiteral("Der Wert kann nicht übernommen werden.");
}
}
AccessTableModel::AccessTableModel(QObject* parent,QSqlDatabase database):QSqlTableModel(parent,database) {
    connect(this,&QAbstractItemModel::dataChanged,this,[this](const QModelIndex& first,const QModelIndex& last) {
        if(!blockMutation_) clearBlockHistory();
        if(editStrategy()==OnManualSubmit) for(int row=first.row();row<=last.row();++row) dirtyRows_.insert(row);
    });
    connect(this,&QAbstractItemModel::rowsInserted,this,[this] {if(!blockMutation_) clearBlockHistory();});
    connect(this,&QAbstractItemModel::rowsRemoved,this,[this] {if(!blockMutation_) clearBlockHistory();});
}
bool AccessTableModel::rejectEdit(int number,const QString& message) {
    setLastError(QSqlError("LinAccess",message,QSqlError::StatementError,QString::number(number)));
    emit editRejected(number,message); return false;
}
bool AccessTableModel::reject(int number) { return rejectEdit(number,editMessage(number)); }
bool AccessTableModel::refreshTypes() {
    types_.clear();decimalTypes_.clear();captions_.clear();descriptions_.clear();formats_.clear(); generated_.clear(); metadataError_=0;
    try { types_=sqliteColumnTypes(database(),tableName());if(types_.values().contains("DECIMAL_EXACT")) decimalTypes_=nativeTableDecimalTypes(database(),tableName()); generated_=generatedAccessColumns(database(),tableName());const auto properties=nativeTableFieldDisplayProperties(database(),tableName());for(auto it=properties.cbegin();it!=properties.cend();++it) {if(it.value().caption) captions_.insert(it.key(),*it.value().caption);if(it.value().description) descriptions_.insert(it.key(),*it.value().description);if(it.value().format) formats_.insert(it.key(),*it.value().format);}return true; }
    catch(const VbaRuntimeError& error) { metadataError_=error.number; return reject(error.number); }
}
QVariant AccessTableModel::checkedColumnInput(int column,const QVariant& input) const {
    const auto type=columnType(column);
    if(type=="CURRENCY_TICKS") {
        if(input.metaType()==QMetaType::fromType<NativeGridCurrency>()) return sqlFieldValue(VbaValue::currency(input.value<NativeGridCurrency>().units),type);
        if(replayingHistory_ && (input.metaType().id()==QMetaType::LongLong || input.metaType().id()==QMetaType::Int)) return sqlFieldValue(qtFieldValue(input,type),type);
        return checkedInput(type,input);
    }
    if(type!="DECIMAL_EXACT") return checkedInput(type,input);
    const auto shape=decimalTypes_.constFind(record().fieldName(column).toCaseFolded());if(shape==decimalTypes_.cend()) throw VbaRuntimeError(3251,"Decimal column write context is not proven");
    VbaValue value;
    if(input.isNull() || (input.metaType().id()==QMetaType::QString && input.toString().isEmpty())) value=VbaValue{};
    else if(input.metaType()==QMetaType::fromType<NativeGridCurrency>()) value=VbaValue::currency(input.value<NativeGridCurrency>().units);
    else if(input.metaType().id()==QMetaType::QString) value=VbaValue::string(input.toString().toStdString());
    else if(input.metaType().id()==QMetaType::Bool) value=VbaValue::logical(input.toBool());
    else if(input.metaType().id()==QMetaType::Double || input.metaType().id()==QMetaType::Float) value=VbaValue::numeric(input.toDouble());
    else if(input.metaType().id()==QMetaType::LongLong || input.metaType().id()==QMetaType::ULongLong || input.metaType().id()==QMetaType::Int || input.metaType().id()==QMetaType::UInt) value=VbaValue::string(input.toString().toStdString());
    else if(replayingHistory_ && input.metaType().id()==QMetaType::QByteArray) value=qtFieldValue(input,"DECIMAL_EXACT");
    else throw VbaRuntimeError(13,"Decimal input requires a supported numeric/text value");
    return sqlDecimalFieldValue(value,shape->first,shape->second);
}
void AccessTableModel::setTable(const QString& table) {
    if(saving_) { (void)reject(2501); return; }
    queryAliases_.clear();sourceSort_.clear();clearBlockHistory();dirtyRows_.clear(); deletedRows_.clear(); submittedKeys_.clear(); insertedKey_.clear(); QSqlTableModel::setTable(table); (void)refreshTypes();
}
bool AccessTableModel::select() {
    if(saving_) return reject(2501);
    clearBlockHistory();const bool result=refreshTypes() && QSqlTableModel::select();
    if(result) { dirtyRows_.clear(); deletedRows_.clear(); }
    return result;
}
void AccessTableModel::clear() {
    if(saving_) { (void)reject(2501); return; }
    queryAliases_.clear();sourceSort_.clear();clearBlockHistory();QSqlTableModel::clear();decimalTypes_.clear();captions_.clear();descriptions_.clear(); dirtyRows_.clear(); deletedRows_.clear(); submittedKeys_.clear(); insertedKey_.clear();
}
QVector<QPersistentModelIndex> AccessTableModel::rowAnchors(const QSet<int>& rows) const {
    QVector<QPersistentModelIndex> result; result.reserve(rows.size());
    for(int row:rows) if(row>=0 && row<rowCount()) result.push_back(index(row,0));
    return result;
}
QSet<int> AccessTableModel::anchoredRows(const QVector<QPersistentModelIndex>& anchors) const {
    QSet<int> result; for(const auto& anchor:anchors) if(anchor.isValid()) result.insert(anchor.row()); return result;
}
void AccessTableModel::revertRow(int row) {
    if(!blockMutation_) clearBlockHistory();
    if(saving_) { (void)reject(2501); return; }
    const QScopedValueRollback<bool> changing(changing_,true);
    auto dirty=dirtyRows_,deleted=deletedRows_; dirty.remove(row); deleted.remove(row);
    const auto dirtyAnchors=rowAnchors(dirty),deleteAnchors=rowAnchors(deleted);
    QSqlTableModel::revertRow(row);
    dirtyRows_=anchoredRows(dirtyAnchors); deletedRows_=anchoredRows(deleteAnchors);
}
bool AccessTableModel::hasPendingRecord(int row) const {
    if(editStrategy()!=OnManualSubmit || row<0 || row>=rowCount()) return false;
    if(deletedRows_.contains(row) || !indexInQuery(index(row,0)).isValid()) return true;
    for(int column=0;column<columnCount();++column) if(isDirty(index(row,column))) return true;
    return false;
}
bool AccessTableModel::undoPendingRecord(int row) {
    if(saving_ || changing_) return reject(2501);
    if(!hasPendingRecord(row)) return rejectEdit(2046,"Keine ungespeicherten Änderungen im aktuellen Datensatz.");
    revertRow(row);
    setLastError({});
    return true;
}
bool AccessTableModel::insertRows(int row,int count,const QModelIndex& parent) {
    if(saving_ || changing_) return reject(2501);
    if(readOnly_) return reject(2448);
    const QScopedValueRollback<bool> changing(changing_,true);
    QMap<QString,QVariant> dateDefaults;if(count>0 && editStrategy()==OnManualSubmit) {try {dateDefaults=pendingScalarDefaults(database(),tableName(),types_);}catch(const VbaRuntimeError& error) {return reject(error.number);}}
    const auto dirtyAnchors=rowAnchors(dirtyRows_),deleteAnchors=rowAnchors(deletedRows_);
    const bool accepted=QSqlTableModel::insertRows(row,count,parent);
    if(accepted && !dateDefaults.isEmpty()) for(int inserted=row;inserted<row+count;++inserted) for(auto it=dateDefaults.cbegin();it!=dateDefaults.cend();++it) {const int column=fieldIndex(it.key());if(column>=0) QSqlTableModel::setData(index(inserted,column),it.value(),Qt::EditRole);}
    dirtyRows_=anchoredRows(dirtyAnchors); deletedRows_=anchoredRows(deleteAnchors);
    if(accepted && editStrategy()==OnManualSubmit) for(int i=0;i<count;++i) dirtyRows_.insert(row+i);
    return accepted;
}
bool AccessTableModel::removeRows(int row,int count,const QModelIndex& parent) {
    if(saving_ || changing_) return reject(2501);
    if(readOnly_) return reject(2448);
    const QScopedValueRollback<bool> changing(changing_,true);
    const auto dirtyAnchors=rowAnchors(dirtyRows_),deleteAnchors=rowAnchors(deletedRows_);
    QVector<QPersistentModelIndex> removed;
    if(!parent.isValid() && row>=0 && count>0 && static_cast<qint64>(row)+count<=rowCount())
        for(int i=0;i<count;++i) if(indexInQuery(index(row+i,0)).isValid()) removed.push_back(index(row+i,0));
    const bool accepted=QSqlTableModel::removeRows(row,count,parent);
    dirtyRows_=anchoredRows(dirtyAnchors); deletedRows_=anchoredRows(deleteAnchors);
    if(accepted && editStrategy()==OnManualSubmit) for(const auto& anchor:removed) if(anchor.isValid()) {
        dirtyRows_.insert(anchor.row()); deletedRows_.insert(anchor.row());
    }
    return accepted;
}
bool AccessTableModel::stageRecordDeletions(const QVector<int>& rows) {
    if(saving_ || changing_) return reject(2501);
    if(readOnly_) return reject(2448);
    if(editStrategy()!=OnManualSubmit || primaryKey().isEmpty() || rows.isEmpty() || rows.size()>1000000) return reject(3251);
    if(!canPasteOverUpdates()) return reject(3251);
    auto ordered=rows;std::sort(ordered.begin(),ordered.end(),std::greater<int>());
    for(int i=0;i<ordered.size();++i) {
        const int row=ordered[i];
        if(row<0 || row>=rowCount() || (i && ordered[i-1]==row) || deletedRows_.contains(row)) return reject(3251);
        if(!isStoredRecord(row)) continue;
        const auto values=record(row);const auto key=primaryKey();
        for(int field=0;field<key.count();++field) if(values.value(key.fieldName(field)).isNull()) return reject(3251);
    }
    BlockHistory history;history.originalRows=rowCount();history.applied=true;history.deletions=ordered;
    QVector<BlockHistory> nextHistory;
    QVector<int> staged;
    try {
        history.before=captureUpdates();
        for(const int row:ordered) if(!isStoredRecord(row)) ++history.removedInserts;
        if(!replayingHistory_) nextHistory=prepareBlockHistory(history);
        staged.reserve(ordered.size());
    }
    catch(const VbaRuntimeError& error) {return rejectEdit(error.number,QString::fromUtf8(error.what()));}
    catch(const std::bad_alloc&) {return reject(7);}
    const QScopedValueRollback<bool> mutation(blockMutation_,true);
    for(const int row:ordered) {
        if(!removeRows(row,1)) {
            const auto problem=lastError();
            if(history.before.isEmpty()) {for(const int undone:staged) revertRow(undone);}
            else if(!restoreUpdates(history.before,history.originalRows)) {clearBlockHistory();return reject(3251);}
            setLastError(problem);return false;
        }
        staged.push_back(row);
    }
    if(!replayingHistory_) {undoHistory_.swap(nextHistory);redoHistory_.clear();}
    setLastError({});return true;
}
bool AccessTableModel::setReadOnly(bool value) {
    if(saving_ || changing_) return reject(2501);
    if(readOnly_==value) return true;
    if(isDirty()) return reject(2501);
    readOnly_=value;
    if(rowCount()>0 && columnCount()>0) emit dataChanged(index(0,0),index(rowCount()-1,columnCount()-1));
    return true;
}
bool AccessTableModel::canSubmitCurrentRecordAlone(int row,bool allowInsert) const {
    if(row<0 || row>=rowCount() || primaryKey().isEmpty() || deletedRows_.contains(row)) return false;
    if(!indexInQuery(index(row,0)).isValid()) {
        if(!allowInsert) return false;
        const auto values=record(row);bool supplied=false;
        for(int column=0;column<columnCount();++column)
            if(values.isGenerated(column) && !generated_.contains(values.fieldName(column).toCaseFolded()) && !QSqlTableModel::data(index(row,column),Qt::EditRole).isNull()) supplied=true;
        if(!supplied) return false;
    }
    bool currentDirty=false;
    for(int column=0;column<columnCount();++column) currentDirty|=isDirty(index(row,column));
    if(!currentDirty) return false;
    auto candidates=dirtyRows_; candidates.unite(deletedRows_);
    for(int candidate:candidates) if(candidate!=row && candidate>=0 && candidate<rowCount())
        for(int column=0;column<columnCount();++column) if(isDirty(index(candidate,column))) return false;
    return true;
}
bool AccessTableModel::submitExistingRecord(int row) {
    if(saving_ || changing_) return reject(2501);
    if(readOnly_) return reject(2448);
    if(metadataError_) return reject(metadataError_);
    if(editStrategy()!=OnManualSubmit || row<0 || row>=rowCount() || deletedRows_.contains(row)
        || !indexInQuery(index(row,0)).isValid() || primaryKey().isEmpty()) return reject(3251);
    if(!hasPendingRecord(row)) {setLastError({});return true;}
    const auto primary=primaryValues(row);
    for(int field=0;field<primary.count();++field) {
        const int column=fieldIndex(primary.fieldName(field));
        if(column<0 || primary.value(field).isNull() || isDirty(index(row,column))) return reject(3251);
    }
    auto values=record(row);
    for(int column=0;column<values.count();++column) {
        const bool generated=values.isGenerated(column);
        values.setValue(column,QSqlTableModel::data(index(row,column),Qt::EditRole));values.setGenerated(column,generated);
    }
    int counterError=0;
    {
        const QScopedValueRollback<bool> saving(saving_,true);
        auto db=database();AccessCounterSqlScope counters(db);
        if(!db.transaction()) {setLastError(db.lastError());return false;}
        if(!updateRowInTable(row,values)) {const auto problem=lastError();db.rollback();setLastError(problem);return false;}
        // Refresh after commit must still address this original key. Reject
        // disappeared/trigger-rekeyed records before retiring any cached edit.
        QStringList predicates;QSqlQuery check(db);
        for(int field=0;field<primary.count();++field) predicates.push_back(db.driver()->escapeIdentifier(primary.fieldName(field),QSqlDriver::FieldName)+"=?");
        check.prepare("SELECT 1 FROM "+db.driver()->escapeIdentifier(tableName(),QSqlDriver::TableName)+" WHERE "+predicates.join(" AND "));
        for(int field=0;field<primary.count();++field) check.addBindValue(primary.value(field));
        if(!check.exec()) {const auto problem=check.lastError();db.rollback();setLastError(problem);return false;}
        if(!check.next()) {check.finish();db.rollback();return reject(3251);}check.finish();
        if(!db.commit()) {const auto problem=db.lastError();db.rollback();setLastError(problem);return false;}
        try {counters.complete(true);} catch(const VbaRuntimeError& error) {counterError=error.number;}
    }
    clearBlockHistory();submittedKeys_.clear();
    // Qt selectRow replaces only this cache record. No reselect/reordering or
    // restoration of unrelated update/insert/delete drafts is necessary.
    const bool refreshed=QSqlTableModel::selectRow(row);
    if(!refreshed) QSqlTableModel::revertRow(row); // Never retry a committed write.
    dirtyRows_.remove(row);deletedRows_.remove(row);
    if(!refreshed || counterError) {
        const auto problem=lastError();
        setLastError(QSqlError("LinAccess",QStringLiteral("Datensatz gespeichert; die Ansicht konnte nicht vollständig aktualisiert werden. ")+problem.text(),QSqlError::StatementError,counterError ? QString::number(counterError) : QStringLiteral("3146")));
        return false;
    }
    setLastError({});return true;
}
bool AccessTableModel::submitTransactionally() {
    submittedKeys_.clear();
    if(saving_ || changing_) return reject(2501);
    if(readOnly_) return reject(2448);
    if(metadataError_) return reject(metadataError_);
    if(editStrategy()!=OnManualSubmit) return reject(3251);
    if(!isDirty()) return select();
    struct Pending { int row; enum Kind { Update,Insert,Delete } kind; QSqlRecord values; };
    QVector<Pending> pending;
    auto candidates=dirtyRows_; candidates.unite(deletedRows_);
    auto rows=candidates.values(); std::sort(rows.begin(),rows.end());
    for(int row:rows) {
        if(row<0 || row>=rowCount()) continue;
        bool dirty=false; for(int col=0;col<columnCount();++col) dirty|=isDirty(index(row,col));
        if(!dirty) continue;
        const auto kind=deletedRows_.contains(row) ? Pending::Delete : !indexInQuery(index(row,0)).isValid() ? Pending::Insert : Pending::Update;
        auto values=record(row);
        for(int col=0;col<values.count();++col) {
            const bool generated=values.isGenerated(col);
            values.setValue(col,QSqlTableModel::data(index(row,col),Qt::EditRole)); values.setGenerated(col,generated);
        }
        pending.push_back({row,kind,std::move(values)});
    }
    if(pending.isEmpty()) return reject(3251);
    QMap<int,QMap<QString,QVariant>> committedKeys;
    {
        const QScopedValueRollback<bool> saving(saving_,true);
        auto db=database();
        AccessCounterSqlScope counterScope(db);
        if(!db.transaction()) { setLastError(db.lastError()); return false; }
        for(const auto& item:pending) {
            const bool accepted=item.kind==Pending::Insert ? insertRowIntoTable(item.values)
                : item.kind==Pending::Delete ? deleteRowFromTable(item.row) : updateRowInTable(item.row,item.values);
            if(!accepted) { const auto problem=lastError(); db.rollback(); setLastError(problem); return false; }
            if(item.kind==Pending::Insert && !insertedKey_.isEmpty()) committedKeys.insert(item.row,insertedKey_);
        }
        if(!db.commit()) { const auto problem=db.lastError(); db.rollback(); setLastError(problem); return false; }
        try {counterScope.complete(true);} catch(const VbaRuntimeError& error) {return reject(error.number);}
    }
    submittedKeys_=std::move(committedKeys);
    // No Qt cache entry is marked submitted before the database commit succeeds.
    // Once committed, retire the drafts even if metadata/query refresh fails:
    // otherwise retrying Save can execute an already committed INSERT again.
    QSqlTableModel::revertAll(); dirtyRows_.clear(); deletedRows_.clear();
    if(select()) return true;
    const auto problem=lastError();
    setLastError(QSqlError("LinAccess",QStringLiteral("Änderungen gespeichert; die Ansicht konnte nicht aktualisiert werden. ")+problem.text(),problem.type(),problem.nativeErrorCode()));
    return false;
}
bool AccessTableModel::replaceAllWholeText(int column,const QString& search,const QString& replacement,Qt::CaseSensitivity sensitivity,int* count) {
    return replaceAllLiteralText(column,search,replacement,1,sensitivity,count);
}
bool AccessTableModel::replaceAllLiteralText(int column,const QString& search,const QString& replacement,int mode,Qt::CaseSensitivity sensitivity,int* count) {
    return replaceAllLiteralTextColumns({column},search,replacement,mode,sensitivity,count);
}
bool AccessTableModel::replaceAllLiteralTextColumns(const QVector<int>& columns,const QString& search,const QString& replacement,int mode,Qt::CaseSensitivity sensitivity,int* count) {
    if(count) *count=0;
    if(saving_ || changing_ || isDirty()) return reject(2501);
    if(readOnly_) return reject(2448);
    if(editStrategy()!=OnManualSubmit || columns.isEmpty() || columns.size()>columnCount()) return reject(3251);
    if(metadataError_) return reject(metadataError_);
    if(search.isEmpty() || mode<0 || mode>2) return reject(5);
    if(search.size()>1000000 || replacement.size()>1000000) return reject(3251);
    const auto primary=primaryKey();
    if(primary.isEmpty()) return rejectEdit(3251,"Alle ersetzen erfordert eine Tabelle mit Primärschlüssel.");
    QSet<int> unique;int firstColumn=columnCount(),lastColumn=-1;
    for(const auto column:columns) {
        if(!isTextColumn(column) || unique.contains(column)) return reject(3251);
        if(isGeneratedColumn(column)) return reject(2448);
        unique.insert(column);firstColumn=std::min(firstColumn,column);lastColumn=std::max(lastColumn,column);
        const auto name=record().fieldName(column);
        for(int field=0;field<primary.count();++field)
            if(primary.fieldName(field).compare(name,Qt::CaseInsensitive)==0) return rejectEdit(3251,"Primärschlüsselfelder können hier nicht gemeinsam ersetzt werden.");
    }
    QVector<QPair<int,int>> matches;QSet<int> affectedRows;QVector<QVariant> prepared;qsizetype outputUnits=0;QElapsedTimer clock;clock.start();
    try {
        (void)replaceNativeText({},search,replacement,mode,sensitivity,true);
        setLastError({});
        while(canFetchMore({})) {
            if(clock.hasExpired(5000)) return reject(3035);
            if(rowCount()>1000000) return reject(3251);
            fetchMore({});
        }
        if(lastError().isValid()) return false;
        if(rowCount()>1000000) return reject(3251);
        for(int row=0;row<rowCount();++row) for(const auto column:columns) {
            const auto target=index(row,column); const auto value=QSqlTableModel::data(target,Qt::EditRole);
            if(clock.hasExpired(5000)) return reject(3035);
            if(value.isNull()) continue;
            const auto changed=replaceNativeText(value.toString(),search,replacement,mode,sensitivity,true);if(!changed.matched) continue;
            if(!(flags(target)&Qt::ItemIsEditable) || !indexInQuery(target).isValid()) return reject(2448);
            if(changed.value.size()>16*1024*1024-outputUnits) return reject(3251);outputUnits+=changed.value.size();
            if(matches.size()>=1000000) return reject(3251);
            matches.append({row,column});affectedRows.insert(row);prepared.append(checkedColumnInput(column,changed.value));
        }
    } catch(const VbaRuntimeError& error) { return reject(error.number); }
      catch(const std::bad_alloc&) { return reject(7); }
    if(matches.isEmpty()) { setLastError({}); return true; }
    int failure=0;
    {
        const QScopedValueRollback<bool> saving(saving_,true);
        const QScopedValueRollback<bool> changing(changing_,true);
        {
            const QSignalBlocker blocked(this);
            try {
                for(int cell=0;cell<matches.size();++cell) if(!QSqlTableModel::setData(index(matches[cell].first,matches[cell].second),prepared[cell],Qt::EditRole)) { failure=3251; break; }
            } catch(const std::bad_alloc&) { failure=7; }
              catch(const std::exception&) { failure=3251; }
            if(failure) {
                for(const auto row:affectedRows) QSqlTableModel::revertRow(row);
                dirtyRows_.clear(); deletedRows_.clear();
            }
        }
        emit dataChanged(index(matches.front().first,firstColumn),index(matches.back().first,lastColumn),{Qt::DisplayRole,Qt::EditRole});
    }
    if(failure) { dirtyRows_.clear(); deletedRows_.clear(); return reject(failure); }
    if(!submitTransactionally()) {
        const auto problem=lastError();
        // Entry was clean. A failed SQL transaction must leave neither partial
        // database updates nor replacement drafts for a later unrelated Save.
        if(isDirty()) {
            const QScopedValueRollback<bool> saving(saving_,true);
            // Qualified row rollback bypasses our saving guard; revertAll
            // dispatches virtual revertRow and would leave the drafts intact.
            for(const auto row:affectedRows) QSqlTableModel::revertRow(row);
            dirtyRows_.clear(); deletedRows_.clear();
        }
        setLastError(problem); return false;
    }
    if(count) *count=matches.size();
    return true;
}
bool AccessTableModel::setSourceSort(const QVector<QPair<int,Qt::SortOrder>>& ordering) {
    if(saving_ || changing_ || isDirty()) return reject(2501);
    if(ordering.isEmpty() || ordering.size()>10) return reject(3251);
    QSet<int> unique;
    for(const auto& item:ordering) {
        if(item.first<0 || item.first>=columnCount() || unique.contains(item.first) || (item.second!=Qt::AscendingOrder && item.second!=Qt::DescendingOrder)) return reject(3251);
        unique.insert(item.first);
    }
    sourceSort_=ordering;sortColumn_=ordering.front().first;sortOrder_=ordering.front().second;
    QSqlTableModel::setSort(sortColumn_,sortOrder_);return true;
}
void AccessTableModel::setSort(int column,Qt::SortOrder order) {
    if(!sourceSort_.isEmpty() && sourceSort_.front()!=qMakePair(column,order)) sourceSort_.clear();
    QSqlTableModel::setSort(column,order);
}
QString AccessTableModel::orderByClause() const {
    if(sourceSort_.isEmpty()) return QSqlTableModel::orderByClause();
    QStringList terms;const auto* driver=database().driver();
    for(const auto& item:sourceSort_) terms.push_back(driver->escapeIdentifier(tableName(),QSqlDriver::TableName)+"."+driver->escapeIdentifier(record().fieldName(item.first),QSqlDriver::FieldName)+(item.second==Qt::DescendingOrder ? " DESC" : " ASC"));
    return "ORDER BY "+terms.join(',');
}
void AccessTableModel::sort(int column,Qt::SortOrder order) {
    const bool protectedEditor=protectSort_; protectSort_=false;
    if(protectedEditor || isDirty()) {
        reject(2501); emit sortRestored(sortColumn_,sortOrder_); return;
    }
    sortColumn_=column; sortOrder_=order; QSqlTableModel::sort(column,order);
}
bool AccessTableModel::isSearchableColumn(int column) const {
    return column>=0 && column<columnCount() && !metadataError_ && nativeFindFieldKind(columnType(column))!=NativeFindFieldKind::Unsupported;
}
QString AccessTableModel::fieldSqlType(int column) const {
    if(column<0 || column>=columnCount()) throw VbaRuntimeError(5,"Invalid FindRecord column");
    if(metadataError_) throw VbaRuntimeError(metadataError_,"FindRecord field metadata is unavailable");
    return columnType(column);
}
bool AccessTableModel::isTextColumn(int column) const {
    if(column<0 || column>=columnCount() || metadataError_) return false;
    const auto type=columnType(column);
    return type.contains("TEXT",Qt::CaseInsensitive) || type.contains("CHAR",Qt::CaseInsensitive) || type.contains("CLOB",Qt::CaseInsensitive);
}
QString AccessTableModel::columnType(int column) const { return types_.value(record().fieldName(column).toCaseFolded()); }
bool AccessTableModel::setQueryColumnAliases(const QMap<int,QString>& aliases) {
    for(auto it=aliases.cbegin();it!=aliases.cend();++it) if(it.key()<0 || it.key()>=columnCount() || it.value().isEmpty()) return reject(3251);
    queryAliases_=aliases;if(columnCount()) emit headerDataChanged(Qt::Horizontal,0,columnCount()-1);return true;
}
QVariant AccessTableModel::headerData(int section,Qt::Orientation orientation,int role) const {
    if(orientation==Qt::Horizontal && role==Qt::DisplayRole && queryAliases_.contains(section)) return queryAliases_.value(section);
    if(orientation==Qt::Horizontal && section>=0 && section<columnCount()) {const auto field=record().fieldName(section),key=field.toCaseFolded();if(role==Qt::DisplayRole && !captions_.value(key).isEmpty()) return captions_.value(key);if(role==Qt::ToolTipRole && (captions_.contains(key) || descriptions_.contains(key))) return field+(descriptions_.value(key).isEmpty() ? QString{} : "\n"+descriptions_.value(key));if(role==Qt::StatusTipRole && descriptions_.contains(key)) return descriptions_.value(key);}return QSqlTableModel::headerData(section,orientation,role);
}
QVariant AccessTableModel::data(const QModelIndex& index,int role) const {
    if(index.isValid() && role==Qt::StatusTipRole) {const auto key=record().fieldName(index.column()).toCaseFolded();if(descriptions_.contains(key)) return descriptions_.value(key);}
    if(index.isValid() && columnType(index.column()).contains("BOOL",Qt::CaseInsensitive)) {
        if(role==Qt::CheckStateRole) {
            const auto value=QSqlTableModel::data(index,Qt::EditRole);
            if(value.isNull()) return Qt::PartiallyChecked;
            try { return qtFieldValue(value,"BOOLEAN").boolean ? Qt::Checked : Qt::Unchecked; }
            catch(const VbaRuntimeError&) { return {}; }
        }
        if(role==Qt::DisplayRole) return {};
    }
    const auto raw=QSqlTableModel::data(index,role);
    if(!index.isValid() || (role!=Qt::DisplayRole && role!=Qt::EditRole)) return raw;
    const auto name=record().fieldName(index.column()).toCaseFolded(),type=types_.value(name);
    if(generated_.contains(name) && (raw.isNull() || raw.metaType().id()==QMetaType::QString) && role==Qt::DisplayRole) return QStringLiteral("(AutoWert)");
    if(role==Qt::DisplayRole && !raw.isNull() && !formats_.value(name).isEmpty()) {
        try {return formattedText(index);}catch(const VbaRuntimeError& failure) {return QStringLiteral("#Fehler (%1)").arg(failure.number);}
    }
    // Typed null spin/date editors otherwise materialize their own defaults.
    // A null QString selects the text editor without losing SQL NULL identity.
    if(role==Qt::EditRole && raw.isNull() && !type.contains("BOOL",Qt::CaseInsensitive) && !type.contains("BLOB",Qt::CaseInsensitive))
        return QVariant(QMetaType::fromType<QString>());
    if(!checkedDisplay(type) || raw.isNull() || (role==Qt::EditRole && type!="DECIMAL_EXACT" && type!="CURRENCY_REAL" && type!="CURRENCY_TICKS" && !floatingDisplay(type))) return raw;
    try {
        const auto value=qtFieldValue(raw,type);
        if(floatingDisplay(type) && value.kind==VbaValue::Kind::Number)
            return QString::number(value.number,'g',std::numeric_limits<double>::max_digits10);
        return QString::fromStdString(value.asText());
    }
    catch(const VbaRuntimeError& error) { return QStringLiteral("#Fehler (%1)").arg(error.number); }
}
QString AccessTableModel::formattedText(const QModelIndex& index) const {
    if(!index.isValid() || index.model()!=this || index.row()<0 || index.row()>=rowCount() || index.column()<0 || index.column()>=columnCount()) throw VbaRuntimeError(5,"Formatted cell index is invalid");
    if(metadataError_) throw VbaRuntimeError(metadataError_,"Formatted column metadata is unavailable");
    const auto mask=formats_.value(record().fieldName(index.column()).toCaseFolded());
    const auto raw=QSqlTableModel::data(index,Qt::EditRole);
    if(mask.isEmpty() || raw.isNull() || columnType(index.column()).contains("BOOL",Qt::CaseInsensitive)) return clipboardText(index);
    if((raw.metaType().id()==QMetaType::QByteArray && columnType(index.column())!="DECIMAL_EXACT") || columnType(index.column()).contains("BLOB",Qt::CaseInsensitive)) throw VbaRuntimeError(3251,"Binary cells have no text format");
    return QString::fromStdString(formatAccessValue(qtFieldValue(raw,columnType(index.column())),mask.toUtf8().toStdString(),true).asText());
}
QString AccessTableModel::clipboardText(const QModelIndex& index) const {
    if(!index.isValid() || index.model()!=this) throw VbaRuntimeError(5,"Clipboard cell index is invalid");
    if(metadataError_) throw VbaRuntimeError(metadataError_,"Clipboard column metadata is unavailable");
    const auto raw=QSqlTableModel::data(index,Qt::EditRole);
    if(raw.isNull()) return {};
    const auto type=columnType(index.column());
    if((raw.metaType().id()==QMetaType::QByteArray && type!="DECIMAL_EXACT") || type.contains("BLOB",Qt::CaseInsensitive)) throw VbaRuntimeError(3251,"Binary cells cannot be copied as text");
    const auto value=qtFieldValue(raw,type);
    if(value.kind==VbaValue::Kind::Number)
        return QString::number(value.number,'g',std::numeric_limits<double>::max_digits10);
    return QString::fromStdString(value.asText());
}
Qt::ItemFlags AccessTableModel::flags(const QModelIndex& index) const {
    auto result=QSqlTableModel::flags(index); if(!index.isValid()) return result;
    if(readOnly_) return result & ~(Qt::ItemIsEditable|Qt::ItemIsUserCheckable);
    const auto type=columnType(index.column());
    if((type=="DECIMAL_EXACT" && !decimalTypes_.contains(record().fieldName(index.column()).toCaseFolded())) || metadataError_ || generated_.contains(record().fieldName(index.column()).toCaseFolded()) || (type.contains("BLOB",Qt::CaseInsensitive) && !binaryWriting_)) return result & ~Qt::ItemIsEditable;
    if(type.contains("BOOL",Qt::CaseInsensitive)) result |= Qt::ItemIsUserCheckable;
    if(checkedDisplay(type)) {
        try { (void)qtFieldValue(QSqlTableModel::data(index,Qt::EditRole),type); }
        catch(const VbaRuntimeError&) { result &= ~Qt::ItemIsEditable; }
    }
    return result;
}
bool AccessTableModel::setData(const QModelIndex& index,const QVariant& value,int role) {
    if(saving_ || changing_) return reject(2501);
    if(readOnly_) return reject(2448);
    if(!index.isValid() || index.model()!=this) return false;
    const bool checkbox=role==Qt::CheckStateRole && columnType(index.column()).contains("BOOL",Qt::CaseInsensitive);
    if(role!=Qt::EditRole && !checkbox) return QSqlTableModel::setData(index,value,role);
    if(!(flags(index)&Qt::ItemIsEditable)) return reject(2448);
    const QScopedValueRollback<bool> changing(changing_,true);
    try {
        const auto type=columnType(index.column());
        QVariant input=value;
        if(checkbox) {
            bool valid=false; const auto state=value.toInt(&valid);
            if(!valid || state<Qt::Unchecked || state>Qt::Checked) return reject(5);
            input=state==Qt::PartiallyChecked ? QVariant{} : QVariant(state==Qt::Checked);
        }
        const auto converted=checkedColumnInput(index.column(),input);
        const bool accepted=QSqlTableModel::setData(index,converted,Qt::EditRole);
        if(accepted) { setLastError({}); if(checkbox) emit dataChanged(index,index,{Qt::CheckStateRole}); }
        return accepted;
    } catch(const VbaRuntimeError& error) { return reject(error.number); }
}
bool AccessTableModel::repeatPreviousValue(const QModelIndex& target) {
    if(saving_ || changing_) return reject(2501);
    if(readOnly_) return reject(2448);
    if(editStrategy()!=OnManualSubmit) return reject(3251);
    if(!target.isValid() || target.model()!=this || target.row()==0) return reject(5);
    const auto source=index(target.row()-1,target.column());
    if(!flags(target).testFlag(Qt::ItemIsEditable) || !flags(source).testFlag(Qt::ItemIsEditable)) return reject(2448);
    try {
        const auto raw=QSqlTableModel::data(source,Qt::EditRole);
        // Checked textual conversion retains supported Currency precision and
        // Date/Boolean types; preserve SQL NULL independently of empty text.
        return setData(target,raw.isNull() ? QVariant{} : QVariant(clipboardText(source)),Qt::EditRole);
    } catch(const VbaRuntimeError& error) { return reject(error.number); }
}
bool AccessTableModel::setDataBlock(const QModelIndex& first,const QVector<QVector<QString>>& values) {
    if(!first.isValid() || first.model()!=this) return reject(5);
    return setDataBlock(first.row(),first.column(),values);
}
bool AccessTableModel::isGeneratedColumn(int column) const {
    return !metadataError_ && column>=0 && column<columnCount() && generated_.contains(record().fieldName(column).toCaseFolded());
}
bool AccessTableModel::isPasteColumn(int column,bool binary) const {
    return !readOnly_ && !metadataError_ && column>=0 && column<columnCount()
        && !generated_.contains(record().fieldName(column).toCaseFolded())
        && (columnType(column)!="DECIMAL_EXACT" || decimalTypes_.contains(record().fieldName(column).toCaseFolded()))
        && (binary || !columnType(column).contains("BLOB",Qt::CaseInsensitive));
}
bool AccessTableModel::setDataBlock(int firstRow,int firstColumn,const QVector<QVector<QString>>& values) {
    QVector<QVector<QVariant>> prepared;
    try {for(const auto& row:values) {QVector<QVariant> cells;cells.reserve(row.size());for(const auto& cell:row) cells.push_back(cell);prepared.push_back(std::move(cells));}}
    catch(const std::bad_alloc&) {return reject(7);}
    return applyDataBlock(firstRow,firstColumn,prepared,false);
}
bool AccessTableModel::setTypedDataBlock(int firstRow,int firstColumn,const QVector<QVector<QVariant>>& values) {
    return applyDataBlock(firstRow,firstColumn,values,true);
}
bool AccessTableModel::setDataBlock(int firstRow,const QVector<int>& columns,const QVector<QVector<QString>>& values) {
    if(columns.isEmpty()) return reject(5);QVector<QVector<QVariant>> prepared;
    try {for(const auto& row:values) {QVector<QVariant> cells;for(const auto& cell:row) cells.push_back(cell);prepared.push_back(std::move(cells));}}
    catch(const std::bad_alloc&) {return reject(7);}
    return applyDataBlock(firstRow,columns.front(),prepared,false,nullptr,false,nullptr,&columns);
}
bool AccessTableModel::setTypedDataBlock(int firstRow,const QVector<int>& columns,const QVector<QVector<QVariant>>& values) {
    if(columns.isEmpty()) return reject(5);return applyDataBlock(firstRow,columns.front(),values,true,nullptr,false,nullptr,&columns);
}
bool AccessTableModel::setTypedRecordBlock(int firstRow,const QStringList& fields,const QVector<QVector<QVariant>>& values) {
    if(firstRow!=rowCount() || canFetchMore()) return reject(3251);
    if(fields.isEmpty() || fields.size()>columnCount() || values.isEmpty()) return reject(13);
    if(values.size()>1000000/columnCount()) return reject(3251);
    try {
        QMap<QString,int> source,target;QSet<int> supplied;
        for(int column=0;column<columnCount();++column) {const auto name=record().fieldName(column).toCaseFolded();if(target.contains(name)) return reject(3251);target.insert(name,column);}
        for(int column=0;column<fields.size();++column) {const auto name=fields[column].toCaseFolded();if(source.contains(name) || !target.contains(name)) return reject(13);source.insert(name,column);supplied.insert(target.value(name));}
        QVector<QVector<QVariant>> mapped;mapped.reserve(values.size());
        for(const auto& row:values) {
            if(row.size()!=fields.size()) return reject(13);QVector<QVariant> cells(columnCount());
            for(auto field=source.cbegin();field!=source.cend();++field) cells[target.value(field.key())]=row[field.value()];
            mapped.push_back(std::move(cells));
        }
        return applyDataBlock(firstRow,0,mapped,true,nullptr,true,&supplied);
    } catch(const std::bad_alloc&) {return reject(7);}
}
bool AccessTableModel::applyDataBlock(int firstRow,int firstColumn,const QVector<QVector<QVariant>>& values,bool typed,const QMap<QString,QVariant>* restoredDefaults,bool skipGenerated,const QSet<int>* suppliedColumns,const QVector<int>* mappedColumns) {
    if(saving_ || changing_) return reject(2501);
    if(readOnly_) return reject(2448);
    const QScopedValueRollback<bool> changing(changing_,true);
    const QScopedValueRollback<bool> blockMutation(blockMutation_,true);
    if(editStrategy()!=QSqlTableModel::OnManualSubmit) return reject(3251);
    if(!canPasteOverUpdates()) return rejectEdit(3251,"Der vorhandene Entwurfsumfang überschreitet die native Bereichsgrenze.");
    if(skipGenerated && (!typed || firstColumn!=0 || firstRow!=rowCount() || canFetchMore())) return reject(3251);
    if(metadataError_) return reject(metadataError_);
    if(firstRow<0 || firstRow>rowCount() || firstColumn<0 || firstColumn>=columnCount() || values.isEmpty() || values.front().isEmpty()) return reject(5);
    const auto columns=values.front().size();
    if(mappedColumns) {if(skipGenerated || mappedColumns->size()!=columns) return reject(5);QSet<int> unique;for(const int column:*mappedColumns) {if(column<0 || column>=columnCount() || unique.contains(column)) return reject(5);unique.insert(column);}}
    const qint64 endRow=static_cast<qint64>(firstRow)+values.size();
    if(values.size()>1000000/columns || endRow>std::numeric_limits<int>::max()
        || (!mappedColumns && static_cast<qint64>(firstColumn)+columns>columnCount())) return reject(5);
    // Exhaust the source before appending: an unfetched record is never a new row.
    while(rowCount()<endRow && canFetchMore()) {
        const int previous=rowCount();fetchMore();if(rowCount()==previous) break;
    }
    const int originalRows=rowCount();
    const int appended=std::max(0,static_cast<int>(endRow)-originalRows);
    struct Cell { int row,column; QVariant value; };
    QVector<Cell> checked;QMap<QString,QVariant> dateDefaults;BlockHistory history;
    history.firstRow=firstRow;history.firstColumn=firstColumn;history.originalRows=originalRows;history.applied=true;history.skipGenerated=skipGenerated;if(mappedColumns) history.columns=*mappedColumns;if(suppliedColumns) history.suppliedColumns=*suppliedColumns;
    QVector<BlockHistory> nextHistory;
    try {
        history.before=captureUpdates();
        if(appended) dateDefaults=restoredDefaults ? *restoredDefaults : pendingScalarDefaults(database(),tableName(),types_);
        history.dateDefaults=dateDefaults;
        checked.reserve(values.size()*columns);
        for(qsizetype row=0;row<values.size();++row) {
            if(values[row].size()!=columns) return reject(5);
            QVector<QVariant> historyRow;historyRow.reserve(columns);
            for(qsizetype column=0;column<columns;++column) {
                const int targetRow=firstRow+row,targetColumn=mappedColumns ? (*mappedColumns)[column] : firstColumn+column;
                const auto type=columnType(targetColumn);
                if(skipGenerated && (generated_.contains(record().fieldName(targetColumn).toCaseFolded()) || (suppliedColumns && !suppliedColumns->contains(targetColumn)))) {historyRow.push_back(QVariant{});continue;}
                if(targetRow<originalRows) {
                    if(deletedRows_.contains(targetRow)) return reject(2448);
                    if(!(flags(index(targetRow,targetColumn))&Qt::ItemIsEditable) && !(typed && type.contains("BLOB",Qt::CaseInsensitive) && isPasteColumn(targetColumn,true))) return reject(2448);
                } else if(!isPasteColumn(targetColumn,typed)) return reject(2448);
                const auto input=values[row][column];QVariant converted;
                if(typed && input.isNull()) converted={};
                else if(type=="DECIMAL_EXACT" || type=="CURRENCY_TICKS") converted=checkedColumnInput(targetColumn,input);
                else if(typed && input.metaType()==QMetaType::fromType<NativeGridCurrency>()) {
                    const auto exact=VbaValue::currency(input.value<NativeGridCurrency>().units);
                    if(nativeFindTextType(type) || type=="CURRENCY_REAL") converted=sqlFieldValue(exact,type);
                    else if(type!="SINGLE_REAL" && (type.contains("REAL",Qt::CaseInsensitive) || type.contains("DOUB",Qt::CaseInsensitive) || type.contains("FLOA",Qt::CaseInsensitive))) converted=currencyToRealStorage(exact);
                    else return reject(3251);
                }
                else if(typed && type.contains("BLOB",Qt::CaseInsensitive)) {
                    if(input.metaType().id()!=QMetaType::QByteArray) return reject(13);
                    converted=input;
                } else if(typed && input.metaType().id()==QMetaType::QByteArray) return reject(13);
                else if(typed && (type=="INTEGER" || type=="INT" || type=="BIGINT") && (input.metaType().id()==QMetaType::LongLong || input.metaType().id()==QMetaType::Int)) converted=input;
                else if(typed && (input.metaType().id()==QMetaType::Double || input.metaType().id()==QMetaType::Bool)) converted=sqlFieldValue(input.metaType().id()==QMetaType::Bool ? VbaValue::logical(input.toBool()) : VbaValue::numeric(input.toDouble()),type);
                else converted=checkedColumnInput(targetColumn,input);
                checked.push_back({targetRow,targetColumn,converted});historyRow.push_back(converted);
            }
            history.values.push_back(std::move(historyRow));
        }
        if(!replayingHistory_) nextHistory=prepareBlockHistory(history);
    } catch(const VbaRuntimeError& error) { return rejectEdit(error.number,QString::fromUtf8(error.what())); }
      catch(const std::bad_alloc&) { return reject(7); }
    int failure=0;
    try {if(appended && !QSqlTableModel::insertRows(originalRows,appended)) failure=3251;}
    catch(const std::bad_alloc&) {failure=7;}
    catch(const std::exception&) {failure=3251;}
    {
        const QSignalBlocker blocked(this);
        const QScopedValueRollback<bool> binary(binaryWriting_,typed);
        try {
            if(!failure && appended) for(int row=originalRows;row<endRow;++row) for(auto it=dateDefaults.cbegin();it!=dateDefaults.cend();++it) {
                const int column=fieldIndex(it.key());
                if(column>=0 && !QSqlTableModel::setData(index(row,column),it.value(),Qt::EditRole)) {failure=3251;break;}
            }
            if(!failure) for(const auto& cell:checked) if(!QSqlTableModel::setData(index(cell.row,cell.column),cell.value,Qt::EditRole)) { failure=3251; break; }
        } catch(const std::bad_alloc&) { failure=7; }
          catch(const std::exception&) { failure=3251; }
        // The cache was clean on entry; revert inserted rows and changed cells together.
    }
    if(failure) {const bool restored=restoreUpdates(history.before,originalRows);if(!restored) clearBlockHistory();return reject(restored ? failure : 3251);}
    for(int row=originalRows;row<endRow;++row) dirtyRows_.insert(row);
    if(!replayingHistory_) {undoHistory_.swap(nextHistory);redoHistory_.clear();}
    const int firstChanged=mappedColumns ? *std::min_element(mappedColumns->cbegin(),mappedColumns->cend()) : firstColumn;
    const int lastChanged=mappedColumns ? *std::max_element(mappedColumns->cbegin(),mappedColumns->cend()) : firstColumn+columns-1;
    const auto last=index(firstRow+values.size()-1,lastChanged);
    if(!failure) setLastError({});
    emit dataChanged(index(firstRow,firstChanged),last,{Qt::DisplayRole,Qt::EditRole,Qt::CheckStateRole});
    emit headerDataChanged(Qt::Vertical,firstRow,last.row());
    if(failure) return reject(failure);
    return true;
}
bool AccessTableModel::canPasteOverUpdates() const {
    return !isDirty() || static_cast<qint64>(rowCount())*columnCount()<=1000000;
}
QVector<AccessTableModel::CachedUpdate> AccessTableModel::captureUpdates() const {
    QVector<CachedUpdate> updates;if(!isDirty()) return updates;
    if(!canPasteOverUpdates()) throw VbaRuntimeError(3251,"Pending cache exceeds the native range snapshot bound");
    for(int row=0;row<rowCount();++row) {
        CachedUpdate update;update.row=row;
        update.kind=deletedRows_.contains(row) ? CachedUpdate::Delete : !indexInQuery(index(row,0)).isValid() ? CachedUpdate::Insert : CachedUpdate::Update;
        const auto values=record(row);
        for(int column=0;column<columnCount();++column) {
            const bool include=update.kind==CachedUpdate::Insert ? values.isGenerated(column) : update.kind==CachedUpdate::Delete || isDirty(index(row,column));
            if(!include) continue;
            if(generated_.contains(record().fieldName(column).toCaseFolded())) {
                if(update.kind==CachedUpdate::Update) throw VbaRuntimeError(3251,"A protected pending field cannot be restored");
                continue;
            }
            update.columns.push_back(column);update.values.push_back(QSqlTableModel::data(index(row,column),Qt::EditRole));
        }
        if(update.kind!=CachedUpdate::Update || !update.columns.isEmpty()) updates.push_back(std::move(update));
    }
    return updates;
}
bool AccessTableModel::restoreUpdates(const QVector<CachedUpdate>& updates,int originalRows) {
    const auto inserted=std::count_if(updates.cbegin(),updates.cend(),[](const auto& update) {return update.kind==CachedUpdate::Insert;});
    if(inserted>originalRows) return false;const int physicalRows=originalRows-inserted;
    QSqlTableModel::revertAll();dirtyRows_.clear();deletedRows_.clear();
    while(rowCount()<physicalRows && canFetchMore()) {const auto before=rowCount();fetchMore();if(rowCount()==before) break;}
    if(rowCount()!=physicalRows) return false;
    for(const auto& update:updates) if(update.kind==CachedUpdate::Insert && !QSqlTableModel::insertRows(update.row,1)) return false;
    if(rowCount()!=originalRows) return false;
    {
        const QSignalBlocker blocked(this);const QScopedValueRollback<bool> binary(binaryWriting_,true);
        for(const auto& update:updates) {
            if(update.row<0 || update.row>=rowCount()) return false;
            for(int cell=0;cell<update.columns.size();++cell) if(!QSqlTableModel::setData(index(update.row,update.columns[cell]),update.values[cell],Qt::EditRole)) return false;
            dirtyRows_.insert(update.row);
        }
    }
    for(const auto& update:updates) if(update.kind==CachedUpdate::Delete) {
        if(!QSqlTableModel::removeRows(update.row,1)) return false;
        deletedRows_.insert(update.row);
    }
    if(!updates.isEmpty()) emit dataChanged(index(updates.front().row,0),index(updates.back().row,columnCount()-1),{Qt::DisplayRole,Qt::EditRole,Qt::CheckStateRole});
    return true;
}
void AccessTableModel::clearBlockHistory() {
    undoHistory_.clear();redoHistory_.clear();++historyGeneration_;
}
QVector<AccessTableModel::BlockHistory> AccessTableModel::prepareBlockHistory(const BlockHistory& history) const {
    constexpr qint64 budget=256LL*1024*1024;
    const auto cost=[&](const BlockHistory& action) {
        qint64 bytes=256+4LL*(action.deletions.size()+action.columns.size()+action.suppliedColumns.size());
        const auto valueCost=[](const QVariant& value)->qint64 {
            return 64+(value.metaType().id()==QMetaType::QString ? 2LL*value.toString().size() : value.metaType().id()==QMetaType::QByteArray ? value.toByteArray().size() : 0);
        };
        for(const auto& row:action.values) {bytes+=32;for(const auto& value:row) bytes+=valueCost(value);if(bytes>budget) return bytes;}
        for(const auto& row:action.before) {bytes+=64+4LL*row.columns.size();for(const auto& value:row.values) bytes+=valueCost(value);if(bytes>budget) return bytes;}
        for(auto it=action.dateDefaults.cbegin();it!=action.dateDefaults.cend();++it) bytes+=32+2LL*it.key().size()+valueCost(it.value());
        return bytes;
    };
    const auto newest=cost(history);
    if(newest>budget) throw VbaRuntimeError(3251,"Native range undo snapshot exceeds its memory budget");
    auto next=undoHistory_;
    if(!next.isEmpty()) {
        const auto& previous=next.back();
        const auto rows=qMax<qint64>(previous.originalRows,static_cast<qint64>(previous.firstRow)+previous.values.size())-previous.removedInserts;
        if(history.originalRows!=rows) next.clear(); // New physical fetch boundary.
    }
    next.append(history);
    while(next.size()>64) next.removeFirst();
    qint64 total=newest;int first=next.size()-1;
    while(first>0) {const auto previous=cost(next[first-1]);if(previous>budget-total) break;total+=previous;--first;}
    if(first) next.erase(next.begin(),next.begin()+first);
    return next;
}
bool AccessTableModel::canUndoBlock() const {
    if(readOnly_ || saving_ || changing_ || undoHistory_.isEmpty()) return false;
    const auto& history=undoHistory_.back();
    return rowCount()==qMax<qint64>(history.originalRows,static_cast<qint64>(history.firstRow)+history.values.size())-history.removedInserts;
}
bool AccessTableModel::canRedoBlock() const {
    if(readOnly_ || saving_ || changing_ || redoHistory_.isEmpty()) return false;
    const auto& history=redoHistory_.back();
    return isDirty()==!history.before.isEmpty() && rowCount()==history.originalRows;
}
bool AccessTableModel::undoBlock() {
    if(!canUndoBlock()) return rejectEdit(2046,"Keine unveränderte Bereichsaktion zum Rückgängigmachen vorhanden.");
    try {
        const auto history=undoHistory_.back();auto nextUndo=undoHistory_,nextRedo=redoHistory_;
        nextUndo.removeLast();nextRedo.append(history);const auto generation=historyGeneration_;
        const QScopedValueRollback<bool> changing(changing_,true),blockMutation(blockMutation_,true);
        if(!history.deletions.isEmpty() && history.before.isEmpty()) {
            for(const int row:history.deletions) {QSqlTableModel::revertRow(row);dirtyRows_.remove(row);deletedRows_.remove(row);}
        } else if(!restoreUpdates(history.before,history.originalRows)) {clearBlockHistory();return reject(3251);}
        if(generation!=historyGeneration_) return reject(2501);
        undoHistory_.swap(nextUndo);redoHistory_.swap(nextRedo);setLastError({});return true;
    } catch(const std::bad_alloc&) {return reject(7);}
}
bool AccessTableModel::redoBlock() {
    if(!canRedoBlock()) return rejectEdit(2046,"Keine Bereichsaktion zum Wiederherstellen vorhanden.");
    try {
        const auto history=redoHistory_.back();auto nextUndo=undoHistory_,nextRedo=redoHistory_;
        nextUndo.append(history);nextRedo.removeLast();const auto generation=historyGeneration_;
        const QScopedValueRollback<bool> replay(replayingHistory_,true);
        const bool accepted=!history.deletions.isEmpty() ? stageRecordDeletions(history.deletions)
            : applyDataBlock(history.firstRow,history.firstColumn,history.values,true,&history.dateDefaults,history.skipGenerated,history.skipGenerated ? &history.suppliedColumns : nullptr,history.columns.isEmpty() ? nullptr : &history.columns);
        if(!accepted) return false;
        if(generation!=historyGeneration_) return reject(2501);
        undoHistory_.swap(nextUndo);redoHistory_.swap(nextRedo);setLastError({});return true;
    } catch(const std::bad_alloc&) {return reject(7);}
}
bool AccessTableModel::changeDecimalRow(int row,const QSqlRecord* values) {
    // Qt captures original records through our textual EditRole. Bind raw query
    // values so that stored Decimal keys or Currency ticks never become strings in WHERE.
    if(readOnly_) return reject(2448);if(metadataError_) return reject(metadataError_);
    const auto original=primaryValues(row);if(original.isEmpty()) return reject(3251);
    const auto* driver=database().driver();const auto table=driver->escapeIdentifier(tableName(),QSqlDriver::TableName);
    QStringList predicates,assignments;QVector<QVariant> changed,originalValues;
    for(int field=0;field<original.count();++field) {
        const auto column=fieldIndex(original.fieldName(field));const auto cell=index(row,column);if(column<0 || !indexInQuery(cell).isValid()) return reject(3251);
        // primaryValues() has already captured our textual EditRole. Fetch the
        // raw original query value, bypassing the editable/formatting overrides.
        originalValues.append(QSqlQueryModel::data(cell,Qt::EditRole));
        predicates.append(driver->escapeIdentifier(original.fieldName(field),QSqlDriver::FieldName)+" IS ?");
    }
    if(values) for(int field=0;field<values->count();++field) if(values->isGenerated(field) && !generated_.contains(values->fieldName(field).toCaseFolded())) {if(columnType(field)=="DECIMAL_EXACT" && !decimalTypes_.contains(values->fieldName(field).toCaseFolded())) return reject(3251);assignments.append(driver->escapeIdentifier(values->fieldName(field),QSqlDriver::FieldName)+"=?");changed.append(values->value(field));}
    if(values && assignments.isEmpty()) return reject(3251);
    const auto where=predicates.join(" AND ");
    const auto sql=(values ? "UPDATE "+table+" SET "+assignments.join(',') : "DELETE FROM "+table)+" WHERE "+where+" AND (SELECT COUNT(*) FROM "+table+" WHERE "+where+")=1";
    QSqlQuery query(database());if(!query.prepare(sql)) {setLastError(query.lastError());return false;}
    for(const auto& value:changed) query.addBindValue(value);
    for(int pass=0;pass<2;++pass) for(const auto& value:originalValues) query.addBindValue(value);
    if(!query.exec()) {setLastError(query.lastError());return false;}
    if(query.numRowsAffected()!=1) return reject(3251);return true;
}
bool AccessTableModel::updateRowInTable(int row,const QSqlRecord& values) {
    if(!types_.values().contains("DECIMAL_EXACT") && !types_.values().contains("CURRENCY_TICKS")) return QSqlTableModel::updateRowInTable(row,values);
    auto changed=values;emit beforeUpdate(row,changed);return changeDecimalRow(row,&changed);
}
bool AccessTableModel::deleteRowFromTable(int row) {
    if(!types_.values().contains("DECIMAL_EXACT") && !types_.values().contains("CURRENCY_TICKS")) return QSqlTableModel::deleteRowFromTable(row);
    emit beforeDelete(row);return changeDecimalRow(row,nullptr);
}
bool AccessTableModel::insertRowIntoTable(const QSqlRecord& row) {
    insertedKey_.clear();
    auto values=row;
    for(int i=0;i<values.count();++i) if(generated_.contains(values.fieldName(i).toCaseFolded())) values.setGenerated(i,false);
    AccessCounterSqlScope counterScope(database());
    try {
        const bool inserted=QSqlTableModel::insertRowIntoTable(values);
        counterScope.complete(inserted);
        if(!inserted) return false;
    } catch(const VbaRuntimeError& error) {return reject(error.number);}
    // Capture actual generated keys on this connection inside the transaction.
    // SQLite restores last_insert_rowid after triggers; another connection
    // cannot change it. Never infer a key from MAX(), row order or draft text.
    const auto db=database(); const auto primary=primaryKey();
    if(db.driverName()!=QStringLiteral("QSQLITE") || primary.isEmpty()) return true;
    QString rowid;
    for(const auto& candidate:{QStringLiteral("_rowid_"),QStringLiteral("rowid"),QStringLiteral("oid")}) {
        bool shadowed=false;
        for(int field=0;field<values.count();++field) if(values.fieldName(field).compare(candidate,Qt::CaseInsensitive)==0) { shadowed=true; break; }
        if(!shadowed) { rowid=candidate; break; }
    }
    if(rowid.isEmpty()) return true;
    QStringList fields;
    for(int field=0;field<primary.count();++field) fields.push_back(db.driver()->escapeIdentifier(primary.fieldName(field),QSqlDriver::FieldName));
    QSqlQuery inserted(db);
    // Ignored inserts, WITHOUT ROWID tables and trigger-deleted rows yield no
    // bookmark; successful writes must not become failures of optional UI state.
    if(!inserted.exec("SELECT "+fields.join(',')+" FROM "+db.driver()->escapeIdentifier(tableName(),QSqlDriver::TableName)
        +" WHERE "+db.driver()->escapeIdentifier(rowid,QSqlDriver::FieldName)+"=last_insert_rowid() AND changes()=1") || !inserted.next()) return true;
    for(int field=0;field<primary.count();++field) {
        const auto value=inserted.value(field);
        if(value.isNull()) { insertedKey_.clear(); break; }
        insertedKey_.insert(primary.fieldName(field),value);
    }
    return true;
}
}
