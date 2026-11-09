#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "vba_source_codec.hpp"
#include "vba_analyzer.hpp"
#include <QApplication>
#include <QCryptographicHash>
#include <QDir>
#include <QFile>
#include <QJsonArray>
#include <QJsonObject>
#include <QMap>
#include <QRegularExpression>
#include <QComboBox>
#include <QLineEdit>
#include <QPushButton>
#include <QAbstractButton>
#include <QMessageBox>
#include <QTimer>
#include <QScrollArea>
#include <QSqlQuery>
#include <QTest>
#include <QTreeView>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>
#include <algorithm>

void testRecordsetNavigation() {
    using namespace linaccess;
    auto db=QSqlDatabase::addDatabase("QSQLITE","recordset-navigation"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db); assert(query.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,label TEXT)"));
    assert(query.exec("INSERT INTO records VALUES(1,'first'),(2,'second'),(3,'third')"));
    const QByteArray layout=R"FORM(Begin Form
RecordSource ="records"
BeforeUpdate ="[Event Procedure]"
Begin Section
Height =1800
Begin TextBox
Name ="id"
ControlSource ="id"
Width =1200
Height =300
End
Begin TextBox
Name ="label"
ControlSource ="label"
Top =400
Width =2500
Height =300
End
Begin ComboBox
Name ="pick"
AfterUpdate ="[Event Procedure]"
RowSourceType ="Table/Query"
RowSource ="SELECT id,label FROM records ORDER BY id"
ColumnCount =2
ColumnWidths ="0;1440"
Top =900
Width =2500
Height =300
End
End
End
)FORM";
    const QByteArray code=R"VBA(Sub pick_AfterUpdate()
Dim rows As Object
Set rows=Me.RecordsetClone
rows.FindFirst "[id] = " & Str(Nz(Me!pick,0))
If Not rows.EOF Then Me.Bookmark=rows.Bookmark
End Sub
Sub Form_BeforeUpdate(Cancel As Integer)
If Me.label="blocked" Then Cancel=1
End Sub
Function RecordLabel() As String
RecordLabel=Me.Recordset.Fields(Index:="label").Value
End Function
Function BangLabel() As String
BangLabel=Me.Recordset!label
End Function
Function FieldTotal() As Long
Dim entry As DAO.Field
For Each entry In Me.Recordset.Fields
FieldTotal=FieldTotal+1
Next
End Function
)VBA";
    AccessFormPreview form("recordset navigation",layout,code,db); form.show(); form.activateWindow(); QApplication::processEvents();
    auto model=makePreviewFormObject(form);
    auto recordset=model->member("Recordset",{}).object;
    auto liveFields=recordset->member("Fields",{}).object;
    auto liveLabel=liveFields->member("Item",{VbaValue::string("label")}).object;
    assert(liveFields->isNativeType("dao.fields") && liveLabel->isNativeType("dao.field"));
    assert(liveFields->member("Count",{}).asNumber()==2 && liveLabel->scalarValue().asText()=="first");
    assert(liveLabel->member("Name",{}).asText()=="label" && liveLabel->member("OrdinalPosition",{}).asNumber()==1);
    assert(!liveLabel->member("DataUpdatable",{}).truthy());
    std::vector<VbaValue> callArguments;
    for(const auto* procedure:{"RecordLabel","BangLabel"}) {
        const auto result=form.executeProcedure(procedure,callArguments);
        assert(result.status==VbaExecutionStatus::Completed && result.typedReturnValue && result.typedReturnValue->asText()=="first");
    }
    const auto fieldTotal=form.executeProcedure("FieldTotal",callArguments);assert(fieldTotal.status==VbaExecutionStatus::Completed && fieldTotal.typedReturnValue->asNumber()==2);
    auto clone=model->objectMember("RecordsetClone",{}).object;
    auto other=model->member("RecordsetClone",{}).object;
    assert(other!=clone && other->isNativeType("dao.recordset"));
    int readonlyError=0;try {model->assign("RecordsetClone",VbaValue{});} catch(const VbaRuntimeError& error) {readonlyError=error.number;}
    assert(readonlyError==383);
    int argumentError=0;try {model->member("RecordsetClone",{VbaValue::numeric(1)});} catch(const VbaRuntimeError& error) {argumentError=error.number;}
    assert(argumentError==450);
    clone->member("FindFirst",{VbaValue::string("id=3")});
    const auto third=clone->member("Bookmark",{});
    const auto first=model->member("Bookmark",{});
    assert(model->item("id").asNumber()==1 && clone->item("id").asNumber()==3);
    auto* combo=form.findChild<QComboBox*>("access-control-pick"); assert(combo); combo->setFocus();
    QTest::keyClick(combo,Qt::Key_Down,Qt::AltModifier); QApplication::processEvents();
    auto* view=qobject_cast<QTreeView*>(combo->view()); assert(view && view->isVisible() && combo->count()==3);
    QTest::qWait(QApplication::doubleClickInterval()+20);
    const auto point=view->visualRect(combo->model()->index(1,1)).center();
    QTest::mouseMove(view->viewport(),point); QTest::mouseClick(view->viewport(),Qt::LeftButton,Qt::NoModifier,point); QApplication::processEvents();
    const auto result=form.eventResult("control:pick:AfterUpdate");
    if (!result || result->status!=VbaExecutionStatus::Completed || model->item("id").asNumber()!=2)
        std::cerr<<"Synthetic navigation diagnostic: event="<<bool(result)<<", status="<<(result ? static_cast<int>(result->status) : -1)
            <<", code="<<(result ? result->errorNumber : 0)<<", line="<<(result ? result->line : 0)<<", row="<<combo->currentIndex()<<", id="<<model->item("id").asNumber()<<'\n';
    assert(result && result->status==VbaExecutionStatus::Completed && model->item("id").asNumber()==2);
    assert(clone->item("id").asNumber()==3 && liveLabel->scalarValue().asText()=="second");
    model->assign("Bookmark",third); assert(model->item("label").asText()=="third");
    auto* edit=form.findChild<QLineEdit*>("access-control-label"); assert(edit); edit->setFocus(); edit->selectAll(); QTest::keyClicks(edit,"blocked");
    int number=0; try { model->assign("Bookmark",first); } catch (const VbaRuntimeError& failure) { number=failure.number; }
    assert(number==2501 && form.hasUnsavedChanges() && model->item("id").asNumber()==3);
    assert(liveLabel->scalarValue().asText()=="third"); // Stored snapshot excludes the blocked form draft.
    number=0;try {liveLabel->assign("Value",VbaValue::string("forbidden"));} catch(const VbaRuntimeError& error) {number=error.number;}assert(number==3251);
    model->member("Undo",{}); assert(!form.hasUnsavedChanges());
    edit->setFocus(); edit->selectAll(); QTest::keyClicks(edit,"updated");
    model->assign("Bookmark",first);
    assert(model->item("id").asNumber()==1 && !form.hasUnsavedChanges());
    assert(clone->item("id").asNumber()==3 && clone->item("label").asText()=="updated");
    model->assign("Bookmark",third); assert(model->item("id").asNumber()==3);
    assert(query.exec("SELECT label FROM records WHERE id=3") && query.next() && query.value(0).toString()=="updated"); query.finish();
    model->member("Requery",{});
    number=0;try {liveLabel->scalarValue();} catch(const VbaRuntimeError& error) {number=error.number;}assert(number==3420);
    number=0;try {liveFields->member("Count",{});} catch(const VbaRuntimeError& error) {number=error.number;}assert(number==3420);
    assert(recordset->member("Fields",{VbaValue::string("label")}).object->scalarValue().asText()=="updated");
    number=0; try { model->assign("Bookmark",first); } catch (const VbaRuntimeError& failure) { number=failure.number; }
    assert(number==3159);
    assert(form.close() && clone->item("id").asNumber()==3);
    number=0; try { recordset->member("Clone",{}); } catch (const VbaRuntimeError& failure) { number=failure.number; }
    assert(number==2450);
    number=0;try {liveLabel->member("Name",{});} catch(const VbaRuntimeError& error) {number=error.number;}assert(number==2450);

    const auto queryLayout=QByteArray("Begin Form\nRecordSource =\"SELECT id,label FROM records WHERE id>=2 ORDER BY id DESC\"\nBegin Section\nEnd\nEnd\n");
    AccessFormPreview queryForm("query clone",queryLayout,{},db);
    auto queryModel=makePreviewFormObject(queryForm);
    auto queryClone=queryModel->member("RecordsetClone",{}).object;
    assert(queryClone->member("RecordCount",{}).asNumber()==2);
    queryClone->member("MoveFirst",{});assert(queryClone->item("id").asNumber()==3);
    queryClone->member("MoveNext",{});assert(queryClone->item("id").asNumber()==2);
    assert(queryModel->item("id").asNumber()==3);
    queryModel->assign("Bookmark",queryClone->member("Bookmark",{}));
    assert(queryModel->item("id").asNumber()==2);
    assert(queryModel->member("Recordset",{}).object->defaultMember({VbaValue::string("id")},false).asNumber()==2);
    const auto oldQueryBookmark=queryClone->member("Bookmark",{});
    assert(!queryClone->member("Updatable",{}).truthy());
    bool writeDenied=false;try {queryClone->member("Edit",{});} catch(const std::runtime_error& error) {writeDenied=std::string(error.what()).find("read-only")!=std::string::npos;}
    assert(writeDenied);
    assert(query.exec("DELETE FROM records WHERE id>=2"));queryModel->member("Requery",{});
    auto emptyClone=queryModel->member("RecordsetClone",{}).object;
    assert(emptyClone->member("RecordCount",{}).asNumber()==0);
    assert(emptyClone->member("BOF",{}).truthy() && emptyClone->member("EOF",{}).truthy());
    number=0;try {queryModel->assign("Bookmark",oldQueryBookmark);} catch(const VbaRuntimeError& error) {number=error.number;}
    assert(number==3159 && queryClone->item("id").asNumber()==2);
    assert(queryForm.close() && queryClone->item("id").asNumber()==2);
    number=0;try {queryModel->member("RecordsetClone",{});} catch(const VbaRuntimeError& error) {number=error.number;}
    assert(number==2450);

    assert(query.exec("CREATE TABLE binary_record(id INTEGER PRIMARY KEY,payload BLOB)"));
    query.prepare("INSERT INTO binary_record VALUES(1,?)");query.addBindValue(QByteArray::fromHex("0001ff"));assert(query.exec());query.finish();
    AccessFormPreview binaryForm("binary fields","Begin Form\nRecordSource =\"binary_record\"\nBegin Section\nEnd\nEnd\n",{},db);
    auto binaryModel=makePreviewFormObject(binaryForm);auto binaryField=binaryModel->member("Recordset",{}).object->member("Fields",{VbaValue::string("payload")}).object;
    assert(binaryField->scalarValue().kind==VbaValue::Kind::ByteArray && binaryField->scalarValue().bytes==std::vector<unsigned char>({0,1,255}));
    assert(binaryField->member("FieldSize",{}).asNumber()==3);
    assert(binaryField->member("GetChunk",{VbaValue::numeric(1),VbaValue::numeric(2)}).bytes==std::vector<unsigned char>({1,255}));
    assert(binaryForm.close());
}

void testOriginalRecordsetSearch(const QDir& directory,const QJsonArray& objects) {
    using namespace linaccess;
    QByteArray raw,formBytes; QString formName;
    for (const auto& value:objects) {
        const auto object=value.toObject();
        if (object["kind"]!="vba" || object["sha256"].toString().toLower()!="dbd3623b9404e490a814b70de6318d2c0877969b7feef081696d3e235ae3b427") continue;
        QFile file(directory.filePath(object["file"].toString())); assert(file.open(QIODevice::ReadOnly)); raw=file.readAll();
        assert(QCryptographicHash::hash(raw,QCryptographicHash::Sha256).toHex()==object["sha256"].toString().toLatin1().toLower());
        formName=object["name"].toString().mid(5);
    }
    assert(!raw.isEmpty());
    for (const auto& value:objects) {
        const auto object=value.toObject();
        if (object["kind"]=="form" && object["name"].toString()==formName) {
            QFile file(directory.filePath(object["file"].toString())); assert(file.open(QIODevice::ReadOnly)); formBytes=file.readAll();
            assert(QCryptographicHash::hash(formBytes,QCryptographicHash::Sha256).toHex()==object["sha256"].toString().toLatin1().toLower());
        }
    }
    assert(!formBytes.isEmpty());
    const auto code=decodeVbaSource(raw); const auto lines=code.split('\n'); const auto analysis=analyzeVba(code.toStdString());
    struct Search { QString procedure,field,control; };
    QVector<Search> searches;
    const QRegularExpression criterion("FindFirst\\s+\"\\s*\\[([^\\]]+)\\]",QRegularExpression::CaseInsensitiveOption);
    const QRegularExpression control("Nz\\s*\\(\\s*Me\\s*!\\s*(?:\\[([^\\]]+)\\]|([\\p{L}_][\\p{L}\\p{N}_]*))",QRegularExpression::CaseInsensitiveOption);
    for (std::size_t i=0;i<analysis.procedures.size();++i) {
        const auto& procedure=analysis.procedures[i];
        const int end=i+1<analysis.procedures.size() ? analysis.procedures[i+1].line-1 : lines.size();
        const auto body=lines.mid(procedure.line-1,end-procedure.line+1).join('\n');
        if (!body.contains("FindFirst",Qt::CaseInsensitive)) continue;
        const auto field=criterion.match(body),name=control.match(body); assert(field.hasMatch() && name.hasMatch());
        searches.append({QString::fromStdString(procedure.name),field.captured(1),name.captured(1).isEmpty() ? name.captured(2) : name.captured(1)});
    }
    assert(searches.size()==3);
    const auto layout=parseAccessFormLayout(decodeVbaSource(formBytes));
    assert(!layout.recordSource.isEmpty() && !layout.recordSource.startsWith("SELECT",Qt::CaseInsensitive));
    QMap<QString,QString> columns;
    for (const auto& item:layout.controls) {
        auto field=item.controlSource.trimmed(); if (field.isEmpty() || field.startsWith('=')) continue;
        if (field.startsWith('[') && field.endsWith(']')) field=field.mid(1,field.size()-2);
        columns.insert(field,"TEXT");
    }
    for (const auto& item:searches) columns.insert(item.field,"INTEGER");
    const auto quoted=[](QString name) { return '"'+name.replace('"',"\"\"")+'"'; };
    const auto jetIdentifier=[](QString name) { return '['+name.replace(']',"]]")+']'; };
    auto db=QSqlDatabase::addDatabase("QSQLITE","original-recordset-search"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db); QStringList definitions,names,parameters;
    for (auto it=columns.cbegin();it!=columns.cend();++it) {
        definitions.append(quoted(it.key())+' '+it.value()); names.append(quoted(it.key())); parameters.append("?");
    }
    assert(query.exec("CREATE TABLE "+quoted(layout.recordSource)+"("+definitions.join(',')+")"));
    for (int id=1;id<=2;++id) {
        assert(query.prepare("INSERT INTO "+quoted(layout.recordSource)+"("+names.join(',')+") VALUES("+parameters.join(',')+")"));
        for (auto it=columns.cbegin();it!=columns.cend();++it) query.addBindValue(it.value()=="INTEGER" ? QVariant(id) : QVariant{});
        assert(query.exec());
    }
    // The full original layout/module run against a prepared synthetic table.
    // No original database rows, passwords or lookup source literals are copied.
    AccessFormPreview form(formName,formBytes,raw,db);
    if (form.cancelledByOpen()) {
        const auto result=form.eventResult("OnOpen");
        std::cerr<<"Original search initialization error code="<<(result ? result->errorNumber : 0)<<", line="<<(result ? result->line : 0)<<'\n';
    }
    assert(!form.cancelledByOpen() && form.hasDataSource());
    for (const auto& event:{"OnOpen","OnLoad"}) {
        const auto result=form.eventResult(event);
        if (result && result->status!=VbaExecutionStatus::Completed)
            std::cerr<<"Original search initialization event error="<<result->errorNumber<<", line="<<result->line<<'\n';
        assert(result && result->status==VbaExecutionStatus::Completed);
    }
    form.show(); form.activateWindow(); QApplication::processEvents(); auto model=makePreviewFormObject(form);
    int executed=0,native=0,missing=0,brokenNative=0,unbound=0;
    for (const auto& search:searches) {
        const auto reset=[&] {
            auto clone=model->member("Recordset",{}).object->member("Clone",{}).object;
            clone->member("FindFirst",{VbaValue::string((jetIdentifier(search.field)+"=1").toStdString())});
            model->assign("Bookmark",clone->member("Bookmark",{}));
        };
        const auto current=[&] {
            auto clone=model->member("Recordset",{}).object->member("Clone",{}).object;
            clone->assign("Bookmark",model->member("Bookmark",{})); return clone->item(search.field.toStdString()).asNumber();
        };
        const bool referenceExists=std::any_of(layout.controls.cbegin(),layout.controls.cend(),[&](const auto& definition) {
            return definition.name.compare(search.control,Qt::CaseInsensitive)==0;
        });
        if (referenceExists) {
            for (const auto input:{2,999,0}) {
                reset(); model->assignItem(search.control.toStdString(),input ? VbaValue::numeric(input) : VbaValue{});
                std::vector<VbaValue> arguments; const auto result=form.executeProcedure(search.procedure,arguments);
                if (result.status!=VbaExecutionStatus::Completed)
                    std::cerr<<"Original search execution error="<<result.errorNumber<<", line="<<result.line<<'\n';
                assert(result.status==VbaExecutionStatus::Completed && current()==(input==2 ? 2 : 1)); ++executed;
            }
        } else {
            // One original procedure refers to a control absent from the
            // export. Preserve that error rather than invent a renamed alias.
            reset(); std::vector<VbaValue> arguments; const auto result=form.executeProcedure(search.procedure,arguments);
            assert(result.status==VbaExecutionStatus::Error && result.errorNumber==2465 && current()==1); ++missing;
        }
        int bindings=0;
        for (const auto& definition:layout.controls) {
            auto prefix=definition.eventProcPrefix;
            if (prefix.isEmpty()) { prefix=definition.name; prefix.replace(QRegularExpression("[^\\p{L}\\p{N}_]"),"_"); }
            if (definition.afterUpdate.compare("[Event Procedure]",Qt::CaseInsensitive)!=0 || (prefix+"_AfterUpdate").compare(search.procedure,Qt::CaseInsensitive)!=0) continue;
            ++bindings;
            reset(); auto target=model->member("Controls",{VbaValue::string(definition.name.toStdString())}).object;
            // Prepare a non-sensitive lookup source through real runtime properties.
            target->assign("RowSource",VbaValue::string(("SELECT "+jetIdentifier(search.field)+','+jetIdentifier(search.field)+" FROM "+jetIdentifier(layout.recordSource)+" ORDER BY "+jetIdentifier(search.field)).toStdString()));
            target->assign("ColumnCount",VbaValue::numeric(2)); target->assign("BoundColumn",VbaValue::numeric(1));
            target->assign("ColumnWidths",VbaValue::string("0;1440")); target->assign("Visible",VbaValue::logical(true)); target->assign("Enabled",VbaValue::logical(true));
            target->member("Dropdown",{}); QApplication::processEvents();
            auto* combo=form.findChild<QComboBox*>("access-control-"+definition.name); assert(combo && combo->count()==2);
            auto* view=qobject_cast<QTreeView*>(combo->view()); assert(view && view->isVisible());
            QTest::qWait(QApplication::doubleClickInterval()+20);
            const auto point=view->visualRect(combo->model()->index(1,1)).center();
            QTest::mouseMove(view->viewport(),point); QTest::mouseClick(view->viewport(),Qt::LeftButton,Qt::NoModifier,point); QApplication::processEvents();
            const auto result=form.eventResult("control:"+definition.name+":AfterUpdate");
            if (referenceExists) { assert(result && result->status==VbaExecutionStatus::Completed && current()==2); ++native; }
            else { assert(result && result->status==VbaExecutionStatus::Error && result->errorNumber==2465 && current()==1); ++brokenNative; }
        }
        if (!bindings) ++unbound;
    }
    assert(executed==6 && native==2 && missing==1 && brokenNative==1 && unbound==0);
    QString deleteProcedure;
    const QRegularExpression deleteCall("DoMenuItem\\s+acFormBar\\s*,\\s*acEditMenu\\s*,\\s*6\\s*,",QRegularExpression::CaseInsensitiveOption);
    for (std::size_t i=0;i<analysis.procedures.size();++i) {
        const auto& procedure=analysis.procedures[i];
        const int end=i+1<analysis.procedures.size() ? analysis.procedures[i+1].line-1 : lines.size();
        if (deleteCall.match(lines.mid(procedure.line-1,end-procedure.line+1).join('\n')).hasMatch()) {
            assert(deleteProcedure.isEmpty()); deleteProcedure=QString::fromStdString(procedure.name);
        }
    }
    assert(!deleteProcedure.isEmpty());
    const auto remaining=[&] {
        assert(query.exec("SELECT count(*) FROM "+quoted(layout.recordSource)) && query.next());
        const int count=query.value(0).toInt(); query.finish(); return count;
    };
    int confirmations=0;
    const auto acceptDelete=[&] {
        const int expected=remaining()-1;
        auto* timer=new QTimer(&form); timer->setInterval(1);
        QObject::connect(timer,&QTimer::timeout,&form,[&,timer,expected] {
            auto* dialog=qobject_cast<QMessageBox*>(QApplication::activeModalWidget()); if (!dialog) return;
            assert(dialog->objectName()=="access-delete-confirm" && remaining()==expected);
            timer->stop(); timer->deleteLater(); ++confirmations;
            QTest::mouseClick(dialog->button(QMessageBox::Yes),Qt::LeftButton);
        }); timer->start();
    };
    assert(remaining()==2); acceptDelete(); std::vector<VbaValue> deleteArguments;
    const auto deleted=form.executeProcedure(deleteProcedure,deleteArguments);
    if (deleted.status!=VbaExecutionStatus::Completed) std::cerr<<"Original delete error="<<deleted.errorNumber<<", line="<<deleted.line<<'\n';
    assert(deleted.status==VbaExecutionStatus::Completed && remaining()==1 && confirmations==1);
    int deleteBindings=0;
    for (const auto& definition:layout.controls) {
        const auto prefix=definition.eventProcPrefix.isEmpty() ? definition.name : definition.eventProcPrefix;
        if (definition.type!="CommandButton" || definition.onClick.compare("[Event Procedure]",Qt::CaseInsensitive)!=0 || (prefix+"_Click").compare(deleteProcedure,Qt::CaseInsensitive)!=0) continue;
        auto control=model->member("Controls",{VbaValue::string(definition.name.toStdString())}).object;
        control->assign("Visible",VbaValue::logical(true)); control->assign("Enabled",VbaValue::logical(true));
        auto* button=form.findChild<QPushButton*>("access-control-"+definition.name); assert(button);
        for (auto* area:form.findChildren<QScrollArea*>()) if (area->widget() && area->widget()->isAncestorOf(button)) area->ensureWidgetVisible(button);
        QApplication::processEvents(); acceptDelete(); QTest::mouseClick(button,Qt::LeftButton); QApplication::processEvents();
        const auto result=form.eventResult("control:"+definition.name+":Click");
        assert(result && result->status==VbaExecutionStatus::Completed && remaining()==0 && confirmations==2); ++deleteBindings;
    }
    assert(deleteBindings==1 && form.close());
    std::cout<<"Original record deletion: one unchanged procedure, one explicit call and one native bound click, two confirmed synthetic-row deletions\n";
    std::cout<<"Original recordset search: "<<executed<<" explicit positive/missing/NULL calls, "<<native<<" native successful selections, "
        <<missing<<" missing control reference, "<<brokenNative<<" native reference errors, "<<unbound<<" absent bindings; synthetic records only\n";
}
