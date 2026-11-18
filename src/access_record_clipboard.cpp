#include "access_form_preview.hpp"
#include "dao_database.hpp"
#include "access_autonumber.hpp"
#include "jet_sql_functions.hpp"
#include "native_grid_mime.hpp"
#include "sqlite_execution.hpp"
#include "vba_qt_value.hpp"
#include <QApplication>
#include <QClipboard>
#include <QComboBox>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QLineEdit>
#include <QMimeData>
#include <QScopedValueRollback>
#include <QScopeGuard>
#include <QSet>
#include <QSignalBlocker>
#include <QSqlError>
#include <QSqlQuery>

namespace {
using linaccess::VbaRuntimeError;
QString identifier(QString name) { return '"'+name.replace('"',"\"\"")+'"'; }
QString tabCell(QString value) {
    if (value.contains('\t') || value.contains('\r') || value.contains('\n') || value.contains('"')) {
        value.replace('"',"\"\""); return '"'+value+'"';
    }
    return value;
}
}

void AccessFormPreview::copyCurrentRecord() {
    if (recordCommandRunning_) throw VbaRuntimeError(2046,"Copy during record command unavailable");
    if (!recordSelected_) {
        auto* editor=qobject_cast<QLineEdit*>(QApplication::focusWidget());
        if (editor && isAncestorOf(editor)) { editor->copy(); return; }
        throw VbaRuntimeError(2046,"Copy requires record or text selection");
    }
    if (layout_.kind!=linaccess::AccessDocumentKind::Form || !dataSourceAvailable_ || currentRecordIndex_<1 || currentRecordIndex_>recordRows_.size())
        throw VbaRuntimeError(2046,"Record copy unavailable");
    if (!saveRecord()) throw VbaRuntimeError(2501,"Copy save cancelled");
    const auto& row=recordRows_[currentRecordIndex_-1];
    if (dataColumns_.size()>255 || row.size()!=dataColumns_.size()) throw VbaRuntimeError(3035,"Clipboard column limit exceeded");
    QJsonArray fields; QStringList labels,values;
    for (qsizetype i=0;i<row.size();++i) {
        const auto& name=dataColumns_[i];
        const auto& value=row[i];
        const auto type=dataTypes_.value(name.toCaseFolded());
        auto storage=value;
        auto storageType=type;
        QString text;
        if(!value.isNull() && (type=="CURRENCY_REAL" || type=="CURRENCY_TICKS" || type=="DECIMAL_EXACT")) {
            const auto scalar=linaccess::qtFieldValue(value,type);
            text=QString::fromStdString(scalar.asText());
            if(type=="DECIMAL_EXACT") storage=text;
            if(type=="CURRENCY_REAL") {
                storage=QVariant::fromValue(static_cast<qlonglong>(scalar.currencyUnits));
                storageType="CURRENCY_TICKS";
            }
        } else {
            text=value.metaType().id()==QMetaType::QByteArray ? QStringLiteral("0x")+QString::fromLatin1(value.toByteArray().toHex()) : value.toString();
        }
        auto field=linaccess::nativeGridEncode(storage,storageType);
        field.insert("name",name);
        fields.append(field);
        labels.append(tabCell(name));
        values.append(tabCell(text));
    }
    const auto data=QJsonDocument(QJsonObject{{"version",1},{"fields",fields}}).toJson(QJsonDocument::Compact);
    if (data.size()>linaccess::nativeGridMimeLimit) throw VbaRuntimeError(3035,"Clipboard record limit exceeded");
    auto* mime=new QMimeData; mime->setData(linaccess::nativeRecordMime,data); mime->setText(labels.join('\t')+'\n'+values.join('\t'));
    QApplication::clipboard()->setMimeData(mime);
}

void AccessFormPreview::pasteAppendRecord() {
    using linaccess::VbaValue;
    if (recordCommandRunning_ || layout_.kind!=linaccess::AccessDocumentKind::Form || !layout_.allowAdditions || !directTable_ || !dataSourceAvailable_)
        throw VbaRuntimeError(2046,"Record append unavailable");
    for (const auto* event:{"beforeinsert","afterinsert","beforeupdate","afterupdate","ondirty"})
        if (!layout_.events.value(event).trimmed().isEmpty()) throw VbaRuntimeError(3251,"PasteAppend insertion event pipeline unsupported");
    const auto* mime=QApplication::clipboard()->mimeData();
    if (!mime || !mime->hasFormat(linaccess::nativeRecordMime)) throw VbaRuntimeError(2046,"No native record on clipboard");
    const auto block=linaccess::nativeRecordDecode(mime->data(linaccess::nativeRecordMime));
    if (block.fields.size()!=dataColumns_.size()) throw VbaRuntimeError(13,"Clipboard field shape mismatch");
    QMap<QString,QVariant> source;
    for (qsizetype i=0;i<block.fields.size();++i) source.insert(block.fields[i].toCaseFolded(),block.rows.front()[i]);
    QString table;
    for (const auto& candidate:db_.tables()) if (candidate.compare(layout_.recordSource.trimmed(),Qt::CaseInsensitive)==0) { table=candidate; break; }
    if (table.isEmpty()) throw VbaRuntimeError(3078,"Append table unavailable");
    checkPhysicalRecordSource();
    QSet<QString> counters;
    QSqlQuery metadata(db_); metadata.prepare("SELECT definition FROM main._la_table_schema WHERE table_name=?"); metadata.addBindValue(table);
    if (metadata.exec() && metadata.next())
        for (const auto& field:QJsonDocument::fromJson(metadata.value(0).toByteArray()).array())
            if (field.toObject()["auto_number"].toBool()) counters.insert(field.toObject()["name"].toString().toCaseFolded());
    metadata.finish();
    const auto generated=linaccess::generatedAccessColumns(db_,table);
    QStringList columns,placeholders; QVariantList values;
    for (const auto& column:dataColumns_) {
        const auto name=column.toCaseFolded();
        if (!source.contains(name)) throw VbaRuntimeError(3265,"Clipboard column unavailable");
        if (counters.contains(name) && !generated.contains(name)) throw VbaRuntimeError(3251,"Counter generation for this column unsupported");
        if (generated.contains(name)) continue;
        auto value=source.value(name); const auto type=dataTypes_.value(name);
        if (value.metaType()==QMetaType::fromType<linaccess::NativeGridCurrency>()) {
            if (type!="CURRENCY_TICKS" && type!="CURRENCY_REAL" && type!="TEXT") throw VbaRuntimeError(13,"Clipboard Currency column type mismatch");
            value=linaccess::sqlFieldValue(VbaValue::currency(value.value<linaccess::NativeGridCurrency>().units),type);
        } else if (!value.isNull() && type!="BLOB" && type!="TEXT" && !(type.contains("INT",Qt::CaseInsensitive) && value.metaType().id()==QMetaType::LongLong))
            value=linaccess::sqlFieldValue(linaccess::qtFieldValue(value,type),type);
        if (!value.isNull() && ((type=="BLOB")!=(value.metaType().id()==QMetaType::QByteArray))) throw VbaRuntimeError(13,"Clipboard column type mismatch");
        columns.append(identifier(column)); placeholders.append("?"); values.append(value);
    }
    if (!saveRecord()) throw VbaRuntimeError(2501,"Append save cancelled");
    QScopedValueRollback<bool> running(recordCommandRunning_,true);
    const auto snapshot=recordsetSnapshot_; const auto oldId=currentRecordIndex_>0 ? rowIds_.value(currentRecordIndex_-1) : QVariant{};
    static unsigned long long sequence=0; const auto savepoint="linaccess_append_"+QString::number(++sequence);
    linaccess::AccessCounterSqlScope counterScope(db_);
    QSqlQuery sql(db_); if (!sql.exec("SAVEPOINT "+savepoint)) throw VbaRuntimeError(3146,"Append transaction unavailable");
    bool released=false,displayed=false;
    const auto rollback=qScopeGuard([&] {
        if (released) return;
        QSqlQuery undo(db_); (void)undo.exec("ROLLBACK TO "+savepoint); (void)undo.exec("RELEASE "+savepoint);
        if (displayed) {
            loadRecords(); const int index=oldId.isNull() ? 0 : rowIds_.indexOf(oldId)+1;
            const QSignalBlocker blocked(recordSelector_); recordSelector_->setCurrentIndex(index); currentRecordIndex_=index; loadRecord(); recordsetSnapshot_=snapshot;
        }
    });
    const auto insert=columns.isEmpty() ? "INSERT INTO main."+identifier(table)+" DEFAULT VALUES RETURNING rowid" : "INSERT INTO main."+identifier(table)+"("+columns.join(',')+") VALUES("+placeholders.join(',')+") RETURNING rowid";
    {
        linaccess::SqliteExecutionBudget budget(db_);
        if (!sql.prepare(insert)) throw VbaRuntimeError(3146,"Append preparation failed");
        for (const auto& value:values) sql.addBindValue(value);
        if (!sql.exec()) {
            const auto message=sql.lastError().text().toCaseFolded();
            const auto typed=linaccess::jetSqlFunctionErrorNumber(message.toStdString());
            const int number=typed ? *typed : message.contains("unique constraint") || message.contains("primary key") ? 3022 : message.contains("not null") ? 3314 : message.contains("foreign key") ? 3201 : message.contains("interrupted") ? 3035 : 3146;
            throw VbaRuntimeError(number,"Record append failed");
        }
    }
    if(!sql.next()) throw VbaRuntimeError(3146,"Appended row identity unavailable");
    const auto id=sql.value(0);
    if(id.isNull() || sql.next() || sql.lastError().isValid()) throw VbaRuntimeError(3146,"Appended row identity is ambiguous");
    sql.finish();
    loadRecords(); displayed=true;
    if (!dataSourceAvailable_ || dataTruncated_) throw VbaRuntimeError(3035,"Complete appended record view unavailable");
    const int index=rowIds_.indexOf(id)+1;
    if (index<1) throw VbaRuntimeError(3251,"Appended record excluded by current form filter");
    if (!sql.exec("RELEASE "+savepoint)) throw VbaRuntimeError(3146,"Append commit failed");
    released=true;
    counterScope.complete(true);
    if (snapshot) { linaccess::refreshDaoSnapshot(snapshot,dataColumns_,dataTypes_,recordRows_,rowIds_); recordsetSnapshot_=snapshot; }
    { const QSignalBlocker blocked(recordSelector_); recordSelector_->setCurrentIndex(index); currentRecordIndex_=index; loadRecord(); }
    if (opened_ && initialCurrentDispatched_) (void)formEvent("OnCurrent","Form_Current");
}
