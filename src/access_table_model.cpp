#include "native_field_defaults.hpp"
#include "access_table_model.hpp"
#include "access_autonumber.hpp"
#include "vba_qt_value.hpp"
#include "native_find_value.hpp"
#include <QSqlError>
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
bool checkedDisplay(const QString& type) { return type=="CURRENCY_REAL" || type=="INTEGER1" || type=="INTEGER2" || type=="INTEGER4" || floatingDisplay(type); }
QVariant checkedInput(const QString& type,const QVariant& input) {
    const bool emptyNumeric=input.metaType().id()==QMetaType::QString && input.toString().isEmpty()
        && (type=="CURRENCY_REAL" || type.contains("INT",Qt::CaseInsensitive) || type.contains("REAL",Qt::CaseInsensitive) || type.contains("DATE",Qt::CaseInsensitive));
    const auto scalar=emptyNumeric ? VbaValue{} : qtFieldValue(input,(type.contains("BOOL",Qt::CaseInsensitive) || type.contains("DATE",Qt::CaseInsensitive))?type:QString{});
    return sqlFieldValue(scalar,type);
}
QString editMessage(int number) {
    if(number==6) return QStringLiteral("Zahlenwert außerhalb des zulässigen Bereichs.");
    if(number==13) return QStringLiteral("Der Wert passt nicht zum Spaltentyp.");
    if(number==2501) return QStringLiteral("Zuerst die Zelleingabe abschließen und Änderungen speichern oder verwerfen.");
    if(number==2448) return QStringLiteral("Dieses Feld ist schreibgeschützt.");
    if(number==3251) return QStringLiteral("Dieser Wert wird für den Spaltentyp noch nicht unterstützt.");
    return QStringLiteral("Der Wert kann nicht übernommen werden.");
}
}
AccessTableModel::AccessTableModel(QObject* parent,QSqlDatabase database):QSqlTableModel(parent,database) {
    connect(this,&QAbstractItemModel::dataChanged,this,[this](const QModelIndex& first,const QModelIndex& last) {
        if(editStrategy()==OnManualSubmit) for(int row=first.row();row<=last.row();++row) dirtyRows_.insert(row);
    });
}
bool AccessTableModel::rejectEdit(int number,const QString& message) {
    setLastError(QSqlError("LinAccess",message,QSqlError::StatementError,QString::number(number)));
    emit editRejected(number,message); return false;
}
bool AccessTableModel::reject(int number) { return rejectEdit(number,editMessage(number)); }
bool AccessTableModel::refreshTypes() {
    types_.clear(); generated_.clear(); metadataError_=0;
    try { types_=sqliteColumnTypes(database(),tableName()); generated_=generatedAccessColumns(database(),tableName()); return true; }
    catch(const VbaRuntimeError& error) { metadataError_=error.number; return reject(error.number); }
}
void AccessTableModel::setTable(const QString& table) {
    if(saving_) { (void)reject(2501); return; }
    dirtyRows_.clear(); deletedRows_.clear(); submittedKeys_.clear(); insertedKey_.clear(); QSqlTableModel::setTable(table); (void)refreshTypes();
}
bool AccessTableModel::select() {
    if(saving_) return reject(2501);
    const bool result=refreshTypes() && QSqlTableModel::select();
    if(result) { dirtyRows_.clear(); deletedRows_.clear(); }
    return result;
}
void AccessTableModel::clear() {
    if(saving_) { (void)reject(2501); return; }
    QSqlTableModel::clear(); dirtyRows_.clear(); deletedRows_.clear(); submittedKeys_.clear(); insertedKey_.clear();
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
    if(saving_) { (void)reject(2501); return; }
    const QScopedValueRollback<bool> changing(changing_,true);
    auto dirty=dirtyRows_,deleted=deletedRows_; dirty.remove(row); deleted.remove(row);
    const auto dirtyAnchors=rowAnchors(dirty),deleteAnchors=rowAnchors(deleted);
    QSqlTableModel::revertRow(row);
    dirtyRows_=anchoredRows(dirtyAnchors); deletedRows_=anchoredRows(deleteAnchors);
}
bool AccessTableModel::hasPendingRecord(int row) const {
    if(editStrategy()!=OnManualSubmit || row<0 || row>=rowCount()) return false;
    if(deletedRows_.contains(row)) return true;
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
    QMap<QString,QVariant> dateDefaults;if(count>0 && editStrategy()==OnManualSubmit) {try {dateDefaults=nativeTableDateDefaults(database(),tableName());}catch(const VbaRuntimeError& error) {return reject(error.number);}}
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
bool AccessTableModel::setReadOnly(bool value) {
    if(saving_ || changing_) return reject(2501);
    if(readOnly_==value) return true;
    if(isDirty()) return reject(2501);
    readOnly_=value;
    if(rowCount()>0 && columnCount()>0) emit dataChanged(index(0,0),index(rowCount()-1,columnCount()-1));
    return true;
}
bool AccessTableModel::canSubmitCurrentRecordAlone(int row) const {
    if(row<0 || row>=rowCount() || primaryKey().isEmpty() || deletedRows_.contains(row) || !indexInQuery(index(row,0)).isValid()) return false;
    bool currentDirty=false;
    for(int column=0;column<columnCount();++column) currentDirty|=isDirty(index(row,column));
    if(!currentDirty) return false;
    auto candidates=dirtyRows_; candidates.unite(deletedRows_);
    for(int candidate:candidates) if(candidate!=row && candidate>=0 && candidate<rowCount())
        for(int column=0;column<columnCount();++column) if(isDirty(index(candidate,column))) return false;
    return true;
}
bool AccessTableModel::submitTransactionally() {
    submittedKeys_.clear();
    if(saving_ || changing_) return reject(2501);
    if(readOnly_) return reject(2448);
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
    if(count) *count=0;
    if(saving_ || changing_ || isDirty()) return reject(2501);
    if(readOnly_) return reject(2448);
    if(editStrategy()!=OnManualSubmit || !isTextColumn(column)) return reject(3251);
    if(metadataError_) return reject(metadataError_);
    if(search.isEmpty()) return reject(5);
    if(search.size()>1000000 || replacement.size()>1000000) return reject(3251);
    const auto primary=primaryKey(); const auto name=record().fieldName(column);
    if(primary.isEmpty()) return rejectEdit(3251,"Alle ersetzen erfordert eine Tabelle mit Primärschlüssel.");
    for(int field=0;field<primary.count();++field)
        if(primary.fieldName(field).compare(name,Qt::CaseInsensitive)==0) return rejectEdit(3251,"Primärschlüsselfelder können hier nicht gemeinsam ersetzt werden.");
    QVector<int> matches; QVariant converted;
    try {
        converted=checkedInput(columnType(column),replacement);
        setLastError({});
        while(canFetchMore({})) {
            if(rowCount()>1000000) return reject(3251);
            fetchMore({});
        }
        if(lastError().isValid()) return false;
        if(rowCount()>1000000) return reject(3251);
        for(int row=0;row<rowCount();++row) {
            const auto target=index(row,column); const auto value=QSqlTableModel::data(target,Qt::EditRole);
            if(value.isNull() || value.toString().compare(search,sensitivity)!=0) continue;
            if(!(flags(target)&Qt::ItemIsEditable) || !indexInQuery(target).isValid()) return reject(2448);
            matches.append(row);
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
                for(const auto row:matches) if(!QSqlTableModel::setData(index(row,column),converted,Qt::EditRole)) { failure=3251; break; }
            } catch(const std::bad_alloc&) { failure=7; }
              catch(const std::exception&) { failure=3251; }
            if(failure) {
                for(const auto row:matches) QSqlTableModel::revertRow(row);
                dirtyRows_.clear(); deletedRows_.clear();
            }
        }
        emit dataChanged(index(matches.front(),column),index(matches.back(),column),{Qt::DisplayRole,Qt::EditRole});
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
            for(const auto row:matches) QSqlTableModel::revertRow(row);
            dirtyRows_.clear(); deletedRows_.clear();
        }
        setLastError(problem); return false;
    }
    if(count) *count=matches.size();
    return true;
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
QVariant AccessTableModel::data(const QModelIndex& index,int role) const {
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
    // Typed null spin/date editors otherwise materialize their own defaults.
    // A null QString selects the text editor without losing SQL NULL identity.
    if(role==Qt::EditRole && raw.isNull() && !type.contains("BOOL",Qt::CaseInsensitive) && !type.contains("BLOB",Qt::CaseInsensitive))
        return QVariant(QMetaType::fromType<QString>());
    if(!checkedDisplay(type) || raw.isNull() || (role==Qt::EditRole && type!="CURRENCY_REAL" && !floatingDisplay(type))) return raw;
    try {
        const auto value=qtFieldValue(raw,type);
        if(floatingDisplay(type) && value.kind==VbaValue::Kind::Number)
            return QString::number(value.number,'g',std::numeric_limits<double>::max_digits10);
        return QString::fromStdString(value.asText());
    }
    catch(const VbaRuntimeError& error) { return QStringLiteral("#Fehler (%1)").arg(error.number); }
}
QString AccessTableModel::clipboardText(const QModelIndex& index) const {
    if(!index.isValid() || index.model()!=this) throw VbaRuntimeError(5,"Clipboard cell index is invalid");
    if(metadataError_) throw VbaRuntimeError(metadataError_,"Clipboard column metadata is unavailable");
    const auto raw=QSqlTableModel::data(index,Qt::EditRole);
    if(raw.isNull()) return {};
    const auto type=columnType(index.column());
    if(raw.metaType().id()==QMetaType::QByteArray || type.contains("BLOB",Qt::CaseInsensitive)) throw VbaRuntimeError(3251,"Binary cells cannot be copied as text");
    const auto value=qtFieldValue(raw,type);
    if(value.kind==VbaValue::Kind::Number)
        return QString::number(value.number,'g',std::numeric_limits<double>::max_digits10);
    return QString::fromStdString(value.asText());
}
Qt::ItemFlags AccessTableModel::flags(const QModelIndex& index) const {
    auto result=QSqlTableModel::flags(index); if(!index.isValid()) return result;
    if(readOnly_) return result & ~(Qt::ItemIsEditable|Qt::ItemIsUserCheckable);
    const auto type=columnType(index.column());
    if(metadataError_ || generated_.contains(record().fieldName(index.column()).toCaseFolded()) || type.contains("BLOB",Qt::CaseInsensitive)) return result & ~Qt::ItemIsEditable;
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
        const auto converted=checkedInput(type,input);
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
    if(saving_ || changing_) return reject(2501);
    if(readOnly_) return reject(2448);
    const QScopedValueRollback<bool> changing(changing_,true);
    if(editStrategy()!=QSqlTableModel::OnManualSubmit) return reject(3251);
    if(isDirty()) return reject(2501);
    if(metadataError_) return reject(metadataError_);
    if(!first.isValid() || first.model()!=this || values.isEmpty() || values.front().isEmpty()) return reject(5);
    const auto columns=values.front().size();
    if(values.size()>1000000/columns || static_cast<qint64>(first.row())+values.size()>rowCount()
        || static_cast<qint64>(first.column())+columns>columnCount()) return reject(5);
    struct Cell { QModelIndex index; QVariant value; };
    QVector<Cell> checked;
    try {
        checked.reserve(values.size()*columns);
        for(qsizetype row=0;row<values.size();++row) {
            if(values[row].size()!=columns) return reject(5);
            for(qsizetype column=0;column<columns;++column) {
                const auto target=index(first.row()+row,first.column()+column);
                if(!(flags(target)&Qt::ItemIsEditable)) return reject(2448);
                checked.push_back({target,checkedInput(columnType(target.column()),values[row][column])});
            }
        }
    } catch(const VbaRuntimeError& error) { return reject(error.number); }
      catch(const std::bad_alloc&) { return reject(7); }
    int failure=0;
    {
        const QSignalBlocker blocked(this);
        try {
            for(const auto& cell:checked) if(!QSqlTableModel::setData(cell.index,cell.value,Qt::EditRole)) { failure=3251; break; }
        } catch(const std::bad_alloc&) { failure=7; }
          catch(const std::exception&) { failure=3251; }
        // The cache was clean on entry; only this block can need reverting.
        if(failure) QSqlTableModel::revertAll();
    }
    const auto last=index(first.row()+values.size()-1,first.column()+columns-1);
    if(!failure) setLastError({});
    emit dataChanged(first,last,{Qt::DisplayRole,Qt::EditRole,Qt::CheckStateRole});
    emit headerDataChanged(Qt::Vertical,first.row(),last.row());
    if(failure) return reject(failure);
    return true;
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
