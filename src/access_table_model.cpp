#include "access_table_model.hpp"
#include "access_autonumber.hpp"
#include "vba_qt_value.hpp"
#include <QSqlError>
#include <QSqlRecord>
namespace linaccess {
namespace {
bool checkedDisplay(const QString& type) { return type=="CURRENCY_REAL" || type=="INTEGER1" || type=="INTEGER2" || type=="INTEGER4"; }
QString editMessage(int number) {
    if(number==6) return QStringLiteral("Zahlenwert außerhalb des zulässigen Bereichs.");
    if(number==13) return QStringLiteral("Der Wert passt nicht zum Spaltentyp.");
    if(number==2501) return QStringLiteral("Zuerst die Zelleingabe abschließen und Änderungen speichern oder verwerfen.");
    if(number==2448) return QStringLiteral("Dieses Feld ist schreibgeschützt.");
    if(number==3251) return QStringLiteral("Dieser Wert wird für den Spaltentyp noch nicht unterstützt.");
    return QStringLiteral("Der Wert kann nicht übernommen werden.");
}
}
AccessTableModel::AccessTableModel(QObject* parent,QSqlDatabase database):QSqlTableModel(parent,database) {}
bool AccessTableModel::reject(int number) {
    const auto message=editMessage(number);
    setLastError(QSqlError("LinAccess",message,QSqlError::StatementError,QString::number(number)));
    emit editRejected(number,message); return false;
}
bool AccessTableModel::refreshTypes() {
    types_.clear(); generated_.clear(); metadataError_=0;
    try { types_=sqliteColumnTypes(database(),tableName()); generated_=generatedAccessColumns(database(),tableName()); return true; }
    catch(const VbaRuntimeError& error) { metadataError_=error.number; return reject(error.number); }
}
void AccessTableModel::setTable(const QString& table) { QSqlTableModel::setTable(table); (void)refreshTypes(); }
bool AccessTableModel::select() { return refreshTypes() && QSqlTableModel::select(); }
void AccessTableModel::sort(int column,Qt::SortOrder order) {
    const bool protectedEditor=protectSort_; protectSort_=false;
    if(protectedEditor || isDirty()) {
        reject(2501); emit sortRestored(sortColumn_,sortOrder_); return;
    }
    sortColumn_=column; sortOrder_=order; QSqlTableModel::sort(column,order);
}
QString AccessTableModel::columnType(int column) const { return types_.value(record().fieldName(column).toCaseFolded()); }
QVariant AccessTableModel::data(const QModelIndex& index,int role) const {
    const auto raw=QSqlTableModel::data(index,role);
    if(!index.isValid() || (role!=Qt::DisplayRole && role!=Qt::EditRole)) return raw;
    const auto name=record().fieldName(index.column()).toCaseFolded(),type=types_.value(name);
    if(generated_.contains(name) && (raw.isNull() || raw.metaType().id()==QMetaType::QString) && role==Qt::DisplayRole) return QStringLiteral("(AutoWert)");
    if(!checkedDisplay(type) || raw.isNull() || (role==Qt::EditRole && type!="CURRENCY_REAL")) return raw;
    try { return QString::fromStdString(qtFieldValue(raw,type).asText()); }
    catch(const VbaRuntimeError& error) { return QStringLiteral("#Fehler (%1)").arg(error.number); }
}
Qt::ItemFlags AccessTableModel::flags(const QModelIndex& index) const {
    auto result=QSqlTableModel::flags(index); if(!index.isValid()) return result;
    const auto type=columnType(index.column());
    if(metadataError_ || generated_.contains(record().fieldName(index.column()).toCaseFolded()) || type.contains("BLOB",Qt::CaseInsensitive)) return result & ~Qt::ItemIsEditable;
    if(checkedDisplay(type)) {
        try { (void)qtFieldValue(QSqlTableModel::data(index,Qt::EditRole),type); }
        catch(const VbaRuntimeError&) { result &= ~Qt::ItemIsEditable; }
    }
    return result;
}
bool AccessTableModel::setData(const QModelIndex& index,const QVariant& value,int role) {
    if(role!=Qt::EditRole) return QSqlTableModel::setData(index,value,role);
    if(!index.isValid() || index.model()!=this) return false;
    if(!(flags(index)&Qt::ItemIsEditable)) return reject(2448);
    try {
        const auto type=columnType(index.column());
        const bool emptyNumeric=value.metaType().id()==QMetaType::QString && value.toString().isEmpty()
            && (type=="CURRENCY_REAL" || type.contains("INT",Qt::CaseInsensitive) || type.contains("REAL",Qt::CaseInsensitive) || type.contains("DATE",Qt::CaseInsensitive));
        const auto scalar=emptyNumeric ? VbaValue{} : qtFieldValue(value,(type.contains("BOOL",Qt::CaseInsensitive) || type.contains("DATE",Qt::CaseInsensitive))?type:QString{});
        const auto converted=sqlFieldValue(scalar,type);
        const bool accepted=QSqlTableModel::setData(index,converted,role);
        if(accepted) setLastError({});
        return accepted;
    } catch(const VbaRuntimeError& error) { return reject(error.number); }
}
bool AccessTableModel::insertRowIntoTable(const QSqlRecord& row) {
    auto values=row;
    for(int i=0;i<values.count();++i) if(generated_.contains(values.fieldName(i).toCaseFolded())) values.setGenerated(i,false);
    return QSqlTableModel::insertRowIntoTable(values);
}
}
