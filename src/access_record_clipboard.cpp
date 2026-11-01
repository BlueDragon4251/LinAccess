#include "access_form_preview.hpp"
#include "dao_database.hpp"
#include "access_autonumber.hpp"
#include "jet_sql_functions.hpp"
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
#include <cmath>

namespace {
using linaccess::VbaRuntimeError;
constexpr auto format="application/x-linaccess-record-v1";
constexpr qsizetype limit=8*1024*1024;
QString identifier(QString name) { return '"'+name.replace('"',"\"\"")+'"'; }
QJsonObject encode(const QString& name,const QVariant& value) {
    QJsonObject field{{"name",name}};
    if (value.isNull()) field["kind"]="null";
    else if (value.metaType().id()==QMetaType::Int || value.metaType().id()==QMetaType::LongLong) {
        field["kind"]="integer"; field["value"]=QString::number(value.toLongLong());
    } else if (value.metaType().id()==QMetaType::Double) {
        if (!std::isfinite(value.toDouble())) throw VbaRuntimeError(6,"Nonfinite clipboard value");
        field["kind"]="real"; field["value"]=value.toDouble();
    } else if (value.metaType().id()==QMetaType::Bool) {
        field["kind"]="boolean"; field["value"]=value.toBool();
    } else if (value.metaType().id()==QMetaType::QString) {
        if (value.toString().size()>limit/2) throw VbaRuntimeError(3035,"Clipboard value limit exceeded");
        field["kind"]="text"; field["value"]=value.toString();
    } else if (value.metaType().id()==QMetaType::QByteArray) {
        if (value.toByteArray().size()>limit) throw VbaRuntimeError(3035,"Clipboard value limit exceeded");
        field["kind"]="binary"; field["value"]=QString::fromLatin1(value.toByteArray().toBase64());
    } else throw VbaRuntimeError(13,"Clipboard value type unsupported");
    return field;
}
QVariant decode(const QJsonObject& field) {
    const auto kind=field["kind"].toString(); const auto value=field["value"];
    if (kind=="null") return {};
    if (kind=="integer" && value.isString()) {
        bool valid=false; const auto integer=value.toString().toLongLong(&valid);
        if (valid && QString::number(integer)==value.toString()) return QVariant::fromValue(integer);
    }
    if (kind=="real" && value.isDouble() && std::isfinite(value.toDouble())) return value.toDouble();
    if (kind=="boolean" && value.isBool()) return value.toBool();
    if (kind=="text" && value.isString()) return value.toString();
    if (kind=="binary" && value.isString()) {
        const auto bytes=QByteArray::fromBase64Encoding(value.toString().toLatin1(),QByteArray::AbortOnBase64DecodingErrors);
        if (bytes) return bytes.decoded;
    }
    throw VbaRuntimeError(13,"Invalid typed clipboard value");
}
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
        fields.append(encode(dataColumns_[i],row[i])); labels.append(tabCell(dataColumns_[i]));
        values.append(tabCell(row[i].metaType().id()==QMetaType::QByteArray ? QStringLiteral("0x")+QString::fromLatin1(row[i].toByteArray().toHex()) : row[i].toString()));
    }
    const auto data=QJsonDocument(QJsonObject{{"version",1},{"fields",fields}}).toJson(QJsonDocument::Compact);
    if (data.size()>limit) throw VbaRuntimeError(3035,"Clipboard record limit exceeded");
    auto* mime=new QMimeData; mime->setData(format,data); mime->setText(labels.join('\t')+'\n'+values.join('\t'));
    QApplication::clipboard()->setMimeData(mime);
}

void AccessFormPreview::pasteAppendRecord() {
    using linaccess::VbaValue;
    if (recordCommandRunning_ || layout_.kind!=linaccess::AccessDocumentKind::Form || !layout_.allowAdditions || !directTable_ || !dataSourceAvailable_)
        throw VbaRuntimeError(2046,"Record append unavailable");
    for (const auto* event:{"beforeinsert","afterinsert","beforeupdate","afterupdate","ondirty"})
        if (!layout_.events.value(event).trimmed().isEmpty()) throw VbaRuntimeError(3251,"PasteAppend insertion event pipeline unsupported");
    const auto* mime=QApplication::clipboard()->mimeData();
    if (!mime || !mime->hasFormat(format)) throw VbaRuntimeError(2046,"No native record on clipboard");
    const auto bytes=mime->data(format);
    if (bytes.size()>limit) throw VbaRuntimeError(3035,"Clipboard record limit exceeded");
    const auto document=QJsonDocument::fromJson(bytes);
    if (!document.isObject() || document.object()["version"]!=QJsonValue(1) || !document.object()["fields"].isArray())
        throw VbaRuntimeError(13,"Invalid clipboard record");
    const auto fields=document.object()["fields"].toArray();
    if (fields.size()>255 || fields.size()!=dataColumns_.size()) throw VbaRuntimeError(13,"Clipboard field shape mismatch");
    QMap<QString,QVariant> source;
    for (const auto& field:fields) {
        if (!field.isObject() || !field.toObject()["name"].isString()) throw VbaRuntimeError(13,"Invalid clipboard field");
        const auto object=field.toObject(); const auto name=object["name"].toString().toCaseFolded();
        if (name.isEmpty() || source.contains(name)) throw VbaRuntimeError(13,"Ambiguous clipboard field");
        source.insert(name,decode(object));
    }
    QString table;
    for (const auto& candidate:db_.tables()) if (candidate.compare(layout_.recordSource.trimmed(),Qt::CaseInsensitive)==0) { table=candidate; break; }
    if (table.isEmpty()) throw VbaRuntimeError(3078,"Append table unavailable");
    QSet<QString> counters;
    QSqlQuery metadata(db_); metadata.prepare("SELECT definition FROM _la_table_schema WHERE table_name=?"); metadata.addBindValue(table);
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
        if (!value.isNull() && type!="BLOB" && type!="TEXT" && !(type.contains("INT",Qt::CaseInsensitive) && value.metaType().id()==QMetaType::LongLong))
            value=linaccess::sqlFieldValue(linaccess::qtFieldValue(value,type),type);
        if (!value.isNull() && ((type=="BLOB")!=(value.metaType().id()==QMetaType::QByteArray))) throw VbaRuntimeError(13,"Clipboard column type mismatch");
        columns.append(identifier(column)); placeholders.append("?"); values.append(value);
    }
    if (!saveRecord()) throw VbaRuntimeError(2501,"Append save cancelled");
    QScopedValueRollback<bool> running(recordCommandRunning_,true);
    const auto snapshot=recordsetSnapshot_; const auto oldId=currentRecordIndex_>0 ? rowIds_.value(currentRecordIndex_-1) : QVariant{};
    static unsigned long long sequence=0; const auto savepoint="linaccess_append_"+QString::number(++sequence);
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
    const auto insert=columns.isEmpty() ? "INSERT INTO "+identifier(table)+" DEFAULT VALUES" : "INSERT INTO "+identifier(table)+"("+columns.join(',')+") VALUES("+placeholders.join(',')+")";
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
    if (sql.numRowsAffected()!=1) throw VbaRuntimeError(3146,"Append affected an unexpected row count");
    const auto id=sql.lastInsertId(); sql.finish();
    loadRecords(); displayed=true;
    if (!dataSourceAvailable_ || dataTruncated_) throw VbaRuntimeError(3035,"Complete appended record view unavailable");
    const int index=rowIds_.indexOf(id)+1;
    if (index<1) throw VbaRuntimeError(3251,"Appended record excluded by current form filter");
    if (!sql.exec("RELEASE "+savepoint)) throw VbaRuntimeError(3146,"Append commit failed");
    released=true;
    if (snapshot) { linaccess::refreshDaoSnapshot(snapshot,dataColumns_,dataTypes_,recordRows_,rowIds_); recordsetSnapshot_=snapshot; }
    { const QSignalBlocker blocked(recordSelector_); recordSelector_->setCurrentIndex(index); currentRecordIndex_=index; loadRecord(); }
    if (opened_ && initialCurrentDispatched_) (void)formEvent("OnCurrent","Form_Current");
}
