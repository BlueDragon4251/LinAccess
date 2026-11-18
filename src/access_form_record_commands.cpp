#include "access_form_preview.hpp"
#include "dao_database.hpp"
#include "vba_qt_value.hpp"
#include "native_access_warnings.hpp"
#include "sqlite_execution.hpp"
#include <QComboBox>
#include <QMessageBox>
#include <QScopedValueRollback>
#include <QScopeGuard>
#include <QSignalBlocker>
#include <QSqlQuery>
#include <QSqlError>

namespace {
QString identifier(QString name) { return '"'+name.replace('"',"\"\"")+'"'; }
}

linaccess::NativeGoToRecordArguments linaccess::nativeGoToRecordArguments(const std::vector<VbaValue>& values) {
    if(values.size()>4) throw VbaRuntimeError(450,"GoToRecord accepts at most four arguments");
    const auto omitted=[&](std::size_t index) {
        return index>=values.size() || values[index].kind==VbaValue::Kind::Missing || values[index].kind==VbaValue::Kind::Empty;
    };
    const auto number=[&](std::size_t index,int fallback) {
        return omitted(index) ? fallback : static_cast<int>(coerceVbaModuleValue(values[index],"long").number);
    };
    NativeGoToRecordArguments result;
    result.type=number(0,-1);
    if(!omitted(1)) result.name=QString::fromStdString(coerceVbaModuleValue(values[1],"string").asText());
    result.action=number(2,1);
    result.offset=number(3,1);
    if(result.type==3) throw VbaRuntimeError(3251,"GoToRecord does not navigate report records");
    if(result.type< -1 || result.type>2 || (result.type==-1 && !result.name.isEmpty()))
        throw VbaRuntimeError(3251,"GoToRecord object type/name unsupported");
    if(result.action<0 || result.action>5) throw VbaRuntimeError(5,"Invalid GoToRecord action");
    if((result.action==0 || result.action==1 || result.action==4) && result.offset<1)
        throw VbaRuntimeError(2105,"Invalid record navigation offset");
    return result;
}

void AccessFormPreview::navigateFormRecord(int action,int offset) {
    using linaccess::VbaRuntimeError;
    if(layout_.kind!=linaccess::AccessDocumentKind::Form)
        throw VbaRuntimeError(438,"Report does not support record navigation");
    if(closed_ || closing_ || openCancelled_)
        throw VbaRuntimeError(2450,"GoToRecord form is not open");
    if(recordCommandRunning_) throw VbaRuntimeError(3251,"Navigation during record command unsupported");
    if(!dataSourceAvailable_) throw VbaRuntimeError(dataSourceErrorNumber_ ? dataSourceErrorNumber_ : 2046,"Record source unavailable");
    const bool additions=layout_.allowAdditions && directTable_;
    if(action==5 && !additions) throw VbaRuntimeError(2046,"New records are unavailable");
    const auto requiresRecordEnd=[&] {
        return action==3 || action==5 || ((action==0 || action==1) && currentRecordIndex_==0);
    };
    if(dataTruncated_ && requiresRecordEnd()) throw VbaRuntimeError(3035,"Incomplete record navigation");

    const QPointer<AccessFormPreview> selected=this;
    if(!saveRecord()) throw VbaRuntimeError(2501,"Record navigation save cancelled");
    if(!selected || closed_ || closing_ || openCancelled_ || !dataSourceAvailable_)
        throw VbaRuntimeError(2501,"GoToRecord target became unavailable during save");

    const qint64 count=recordSelector_->count()-1;
    const qint64 current=currentRecordIndex_ ? currentRecordIndex_ : count+1;
    qint64 target;
    switch(action) {
    case 0: target=current-offset; break;
    case 1: target=current+offset; break;
    case 2: target=1; break;
    case 3: target=count; break;
    case 4: target=offset; break;
    case 5: target=count+1; break;
    default: throw VbaRuntimeError(5,"Invalid record navigation action");
    }
    if(dataTruncated_ && (requiresRecordEnd() || target>count))
        throw VbaRuntimeError(3035,"Incomplete record navigation");
    const bool newRecord=target==count+1 && (action==1 || action==5);
    if(newRecord && !(layout_.allowAdditions && directTable_))
        throw VbaRuntimeError(2046,"New records are unavailable");
    if(!newRecord && (target<1 || target>count))
        throw VbaRuntimeError(2105,"Cannot go to specified record");
    const int index=newRecord ? 0 : static_cast<int>(target);
    selectRecord(index);
    if(!selected || closed_ || currentRecordIndex_!=index)
        throw VbaRuntimeError(2501,"Record navigation cancelled");
}

void AccessFormPreview::clearRecordSelection() {
    if (!recordSelected_) return;
    recordSelected_=false; recordSelector_->setStyleSheet({});
}
void AccessFormPreview::selectCurrentRecordForCommand() {
    using linaccess::VbaRuntimeError;
    if (recordCommandRunning_ || layout_.kind!=linaccess::AccessDocumentKind::Form || !dataSourceAvailable_ || currentRecordIndex_<1)
        throw VbaRuntimeError(2046,"Record selection unavailable");
    if (!saveRecord()) throw VbaRuntimeError(2501,"Record selection save cancelled");
    recordSelector_->setFocus(); recordSelected_=true;
    recordSelector_->setStyleSheet("QComboBox { background-color: palette(highlight); color: palette(highlighted-text); }");
}
void AccessFormPreview::deleteCurrentRecord() {
    using linaccess::VbaRuntimeError; using linaccess::VbaValue;
    if (recordCommandRunning_ || layout_.kind!=linaccess::AccessDocumentKind::Form || !layout_.allowDeletions || !layout_.allowEdits || !directTable_ || !dataSourceAvailable_ || currentRecordIndex_<1 || currentRecordIndex_>rowIds_.size())
        throw VbaRuntimeError(2046,"Record deletion unavailable");
    if (!saveRecord()) throw VbaRuntimeError(2501,"Record deletion save cancelled");
    QScopedValueRollback<bool> running(recordCommandRunning_,true);
    const auto oldId=rowIds_[currentRecordIndex_-1]; const int oldIndex=currentRecordIndex_;
    const auto snapshot=recordsetSnapshot_;
    const auto callback=[&](const QString& property,const QString& name,std::vector<VbaValue>& arguments) {
        if (layout_.events.value(property.toCaseFolded()).compare("[Event Procedure]",Qt::CaseInsensitive)!=0) return;
        const auto result=executeProcedure(name,arguments); eventResults_.insert(property.toCaseFolded(),result);
        if (result.status!=linaccess::VbaExecutionStatus::Completed)
            throw VbaRuntimeError(result.errorNumber ? result.errorNumber : 3251,"Record deletion event failed");
        if (dirty_) throw VbaRuntimeError(3251,"Editing a record during deletion unsupported");
    };
    std::vector<VbaValue> cancel{VbaValue::numeric(0)};
    callback("OnDelete","Form_Delete",cancel);
    if (cancel.front().kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null delete cancellation");
    if (cancel.front().truthy()) throw VbaRuntimeError(2501,"Delete event cancelled");
    checkPhysicalRecordSource();
    static unsigned long long sequence=0;
    const auto savepoint="linaccess_delete_"+QString::number(++sequence);
    QSqlQuery sql(db_);
    if (!sql.exec("SAVEPOINT "+savepoint)) throw VbaRuntimeError(3146,"Delete transaction unavailable");
    bool released=false, displayed=false;
    const auto restore=[&] {
        loadRecords();
        const auto index=rowIds_.indexOf(oldId)+1;
        const QSignalBlocker blocked(recordSelector_);
        recordSelector_->setCurrentIndex(index); currentRecordIndex_=index; loadRecord();
        recordsetSnapshot_=snapshot;
    };
    const auto rollback=qScopeGuard([&] {
        if (released) return;
        QSqlQuery undo(db_); (void)undo.exec("ROLLBACK TO "+savepoint); (void)undo.exec("RELEASE "+savepoint);
        if (displayed) restore();
    });
    if (!sql.prepare("DELETE FROM main."+identifier(layout_.recordSource.trimmed())+" WHERE "+identifier(recordIdentity_)+"=?"))
        throw VbaRuntimeError(3146,"Delete preparation failed");
    sql.addBindValue(oldId);
    {
        linaccess::SqliteExecutionBudget budget(db_);
        if (!sql.exec()) {
            const auto message=sql.lastError().text().toCaseFolded();
            throw VbaRuntimeError(message.contains("foreign key") ? 3200 : message.contains("interrupted") ? 3035 : 3146,"Record deletion failed");
        }
    }
    if (sql.numRowsAffected()!=1) throw VbaRuntimeError(3021,"Record no longer available");
    sql.finish();
    loadRecords(); displayed=true;
    {
        const QSignalBlocker blocked(recordSelector_);
        const int index=qMin(oldIndex,static_cast<int>(recordRows_.size()));
        recordSelector_->setCurrentIndex(index); currentRecordIndex_=index; loadRecord();
    }
    if (opened_ && initialCurrentDispatched_) (void)formEvent("OnCurrent","Form_Current");
    std::vector<VbaValue> before{VbaValue::numeric(0),VbaValue::numeric(1)};
    callback("BeforeDelConfirm","Form_BeforeDelConfirm",before);
    if (before[0].kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null confirmation cancellation");
    int status=before[0].truthy() ? 1 : 0;
    if (!status) {
        if (before[1].kind==VbaValue::Kind::Null) throw VbaRuntimeError(94,"Invalid Null confirmation response");
        const auto response=linaccess::sqlFieldValue(before[1],"INTEGER").toLongLong();
        if (response!=0 && response!=1) throw VbaRuntimeError(5,"Invalid delete confirmation response");
        if (response==1 && linaccess::nativeWarningsEnabled()) {
            QMessageBox confirmation(QMessageBox::Question,"LinAccess","Diesen Datensatz löschen?",QMessageBox::Yes|QMessageBox::No,this);
            confirmation.setObjectName("access-delete-confirm"); confirmation.setTextFormat(Qt::PlainText);
            confirmation.setDefaultButton(QMessageBox::No);
            if (confirmation.exec()!=QMessageBox::Yes) status=2;
        }
    }
    if (status) {
        if (!sql.exec("ROLLBACK TO "+savepoint) || !sql.exec("RELEASE "+savepoint)) throw VbaRuntimeError(3146,"Delete rollback failed");
        released=true; restore();
    } else {
        if (!sql.exec("RELEASE "+savepoint)) throw VbaRuntimeError(3146,"Delete commit failed");
        released=true;
        dataEntryRows_.remove(oldId.toLongLong());
        if (snapshot && !dataTruncated_) {
            linaccess::refreshDaoSnapshot(snapshot,dataColumns_,dataTypes_,recordRows_,rowIds_); recordsetSnapshot_=snapshot;
        }
    }
    std::vector<VbaValue> after{VbaValue::numeric(status)};
    callback("AfterDelConfirm","Form_AfterDelConfirm",after);
    if (status) throw VbaRuntimeError(2501,"Record deletion cancelled");
}
