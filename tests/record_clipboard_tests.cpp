#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "access_catalog.hpp"
#include "access_table_import.hpp"
#include "vba_analyzer.hpp"
#include "vba_source_codec.hpp"
#include <QApplication>
#include <QClipboard>
#include <QCryptographicHash>
#include <QDir>
#include <QFile>
#include <QMessageBox>
#include <QAbstractButton>
#include <QPushButton>
#include <QRegularExpression>
#include <QScrollArea>
#include <QTimer>
#include <QComboBox>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QLineEdit>
#include <QMimeData>
#include <QSqlQuery>
#include <QTest>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>

void testRecordClipboard() {
    using namespace linaccess;
    auto db=QSqlDatabase::addDatabase("QSQLITE","record-clipboard"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY AUTOINCREMENT,label TEXT,n INTEGER,amount REAL,empty TEXT,absent TEXT,bytes BLOB)"));
    sql.prepare("INSERT INTO records VALUES(1,?,9223372036854775806,1.25,'',NULL,?)");
    sql.addBindValue(QString::fromUtf8("Grüße\t\"quoted\"\nsecond line")); sql.addBindValue(QByteArray::fromHex("0001ff80")); assert(sql.exec()); sql.finish();
    const QByteArray definition="Begin Form\nRecordSource =\"records\"\nBegin Section\nBegin TextBox\nName =\"label\"\nControlSource =\"label\"\nEnd\nEnd\nEnd\n";
    const QByteArray code=R"VBA(Sub Duplicate()
DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
DoCmd.DoMenuItem acFormBar, acEditMenu, acCopy, , acMenuVer70
DoCmd.DoMenuItem acFormBar, acEditMenu, 5, , acMenuVer70
End Sub
Sub Append()
Application.DoCmd.RunCommand acCmdPasteAppend
End Sub
)VBA";
    AccessFormPreview form("clipboard",definition,code,db); form.show(); form.activateWindow(); QApplication::processEvents();
    auto root=makePreviewFormObject(form),cmd=makePreviewApplicationObject(form)->member("DoCmd",{}).object;
    auto clone=root->member("Recordset",{}).object->member("Clone",{}).object; clone->member("MoveFirst",{});
    const auto bookmark=clone->member("Bookmark",{});
    std::vector<VbaValue> arguments; auto result=form.executeProcedure("Duplicate",arguments);
    assert(result.status==VbaExecutionStatus::Completed && root->member("NewRecord",{}).truthy()==false);
    assert(root->member("Recordset",{}).object->member("RecordCount",{}).asNumber()==2);
    clone->assign("Bookmark",bookmark); assert(clone->item("id").asNumber()==1 && clone->member("RecordCount",{}).asNumber()==2);
    assert(sql.exec("SELECT label,n,amount,empty,absent,bytes FROM records ORDER BY id"));
    assert(sql.next()); QVariantList expected; for (int i=0;i<6;++i) expected.append(sql.value(i));
    assert(sql.next()); for (int i=0;i<6;++i) assert(sql.value(i)==expected[i] && sql.value(i).isNull()==expected[i].isNull());
    assert(!sql.next()); sql.finish();
    assert(sql.exec("SELECT max(id) FROM records") && sql.next() && sql.value(0).toInt()==2); sql.finish();
    cmd->member("RunCommand",{VbaValue::numeric(50)});
    auto* selected=form.findChild<QComboBox*>("access-record-selector"); assert(selected);
    QApplication::clipboard()->clear(); QTest::keyClick(selected,Qt::Key_C,Qt::ControlModifier);
    assert(QApplication::clipboard()->mimeData()->hasFormat("application/x-linaccess-record-v1"));
    const auto original=QApplication::clipboard()->mimeData()->data("application/x-linaccess-record-v1");
    assert(!original.isEmpty() && QApplication::clipboard()->text().contains(QString::fromUtf8("Grüße")));
    const auto setClipboard=[&](const QByteArray& data) { auto* mime=new QMimeData; mime->setData("application/x-linaccess-record-v1",data); QApplication::clipboard()->setMimeData(mime); };
    root->assign("AllowAdditions",VbaValue::logical(false)); result=form.executeProcedure("Append",arguments); assert(result.errorNumber==2046);
    auto* selector=form.findChild<QComboBox*>("access-record-selector"); assert(selector && !(selector->model()->flags(selector->model()->index(0,0))&Qt::ItemIsEnabled));
    int error=0; try { cmd->member("GoToRecord",{VbaValue::empty(),VbaValue::empty(),VbaValue::numeric(5)}); } catch (const VbaRuntimeError& e) { error=e.number; } assert(error==2046);
    root->assign("AllowAdditions",VbaValue::logical(true));
    setClipboard("{}"); result=form.executeProcedure("Append",arguments); assert(result.errorNumber==13);
    auto data=QJsonDocument::fromJson(original).object(); auto fields=data["fields"].toArray(); fields[1]=fields[0]; data["fields"]=fields;
    setClipboard(QJsonDocument(data).toJson()); result=form.executeProcedure("Append",arguments); assert(result.errorNumber==13);
    setClipboard(QByteArray(8*1024*1024+1,'x')); result=form.executeProcedure("Append",arguments); assert(result.errorNumber==3035);
    QApplication::clipboard()->setText("external text"); result=form.executeProcedure("Append",arguments); assert(result.errorNumber==2046);
    setClipboard(original); root->assign("Filter",VbaValue::string("id=1")); root->assign("FilterOn",VbaValue::logical(true));
    result=form.executeProcedure("Append",arguments); assert(result.errorNumber==3251 && root->member("Recordset",{}).object->member("RecordCount",{}).asNumber()==1);
    assert(sql.exec("SELECT count(*) FROM records") && sql.next() && sql.value(0).toInt()==2); sql.finish();
    root->assign("FilterOn",VbaValue::logical(false));
    auto refreshed=root->member("Recordset",{}).object->member("Clone",{}).object; refreshed->member("MoveFirst",{}); root->assign("Bookmark",refreshed->member("Bookmark",{}));
    auto* editor=form.findChild<QLineEdit*>("access-control-label"); assert(editor); editor->setFocus(); editor->setText("text selection"); editor->setSelection(0,4);
    cmd->member("RunCommand",{VbaValue::numeric(190)}); assert(QApplication::clipboard()->text()=="text"); root->member("Undo",{});
    assert(form.close());

    assert(sql.exec("CREATE TABLE manual(id INTEGER PRIMARY KEY,label TEXT,n INTEGER,amount REAL,empty TEXT,absent TEXT,bytes BLOB)"));
    assert(sql.exec("INSERT INTO manual SELECT * FROM records WHERE id=1"));
    auto manualDefinition=definition; manualDefinition.replace("records","manual");
    AccessFormPreview manual("manual key",manualDefinition,code,db); manual.show(); manual.activateWindow(); QApplication::processEvents();
    result=manual.executeProcedure("Duplicate",arguments); assert(result.errorNumber==3022);
    assert(sql.exec("SELECT count(*) FROM manual") && sql.next() && sql.value(0).toInt()==1); sql.finish(); assert(manual.close());
    auto eventDefinition=definition; eventDefinition.replace("Begin Form\n","Begin Form\nBeforeInsert =\"[Event Procedure]\"\n");
    const auto eventCode=code+"Sub Form_BeforeInsert(Cancel As Integer)\nCurrentDb.Execute \"CREATE TABLE forbidden_event(value)\"\nEnd Sub\n";
    AccessFormPreview guarded("unsupported insert pipeline",eventDefinition,eventCode,db); setClipboard(original);
    result=guarded.executeProcedure("Append",arguments); assert(result.errorNumber==3251 && !db.tables().contains("forbidden_event")); assert(guarded.close());
    assert(sql.exec("CREATE TABLE counter_without_pk(id INTEGER NOT NULL,label TEXT,n INTEGER,amount REAL,empty TEXT,absent TEXT,bytes BLOB)"));
    assert(sql.exec("INSERT INTO counter_without_pk SELECT * FROM records WHERE id=1"));
    assert(sql.exec("CREATE TABLE _la_table_schema(table_name TEXT PRIMARY KEY,definition TEXT)"));
    sql.prepare("INSERT INTO _la_table_schema VALUES('counter_without_pk',?)");
    sql.addBindValue(QString::fromUtf8(QJsonDocument(QJsonArray{QJsonObject{{"name","id"},{"auto_number",true},{"auto_number_native",false}}}).toJson())); assert(sql.exec()); sql.finish();
    auto counterDefinition=definition; counterDefinition.replace("records","counter_without_pk");
    AccessFormPreview unsupportedCounter("counter metadata",counterDefinition,code,db); unsupportedCounter.show(); unsupportedCounter.activateWindow(); QApplication::processEvents();
    result=unsupportedCounter.executeProcedure("Duplicate",arguments); assert(result.errorNumber==3251);
    assert(sql.exec("SELECT count(*) FROM counter_without_pk") && sql.next() && sql.value(0).toInt()==1); sql.finish(); assert(unsupportedCounter.close());
    assert(sql.exec("CREATE TABLE append_target(id INTEGER PRIMARY KEY AUTOINCREMENT,label TEXT,n INTEGER,amount REAL,empty TEXT,absent TEXT,bytes BLOB)"));
    const QByteArray callerCode="Sub AppendActive()\nDoCmd.RunCommand acCmdPasteAppend\nMe.Tag=\"caller\"\nEnd Sub\n";
    AccessFormPreview caller("append caller","Begin Form\nEnd\n",callerCode,db); caller.show();
    auto targetDefinition=definition; targetDefinition.replace("records","append_target");
    AccessFormPreview target("append target",targetDefinition,{},db); target.show(); target.activateWindow(); QApplication::processEvents(); setClipboard(original);
    result=caller.executeProcedure("AppendActive",arguments); assert(result.status==VbaExecutionStatus::Completed);
    assert(sql.exec("SELECT id,n FROM append_target") && sql.next() && sql.value(0).toInt()==1 && sql.value(1).toLongLong()==9223372036854775806LL && !sql.next()); sql.finish();
    assert(makePreviewFormObject(caller)->member("Tag",{}).asText()=="caller");
    assert(sql.exec("SELECT count(*) FROM records") && sql.next() && sql.value(0).toInt()==2); sql.finish();
    assert(target.close() && caller.close());
    QApplication::clipboard()->clear();
}

void testOriginalClipboard(const QDir& directory,const QJsonArray& objects,const QString& source) {
    using namespace linaccess;
    QByteArray raw,formBytes; QString formName;
    for (const auto& value:objects) {
        const auto object=value.toObject();
        if (object["kind"]!="vba" || object["sha256"].toString().toLower()!="dbd3623b9404e490a814b70de6318d2c0877969b7feef081696d3e235ae3b427") continue;
        QFile file(directory.filePath(object["file"].toString())); assert(file.open(QIODevice::ReadOnly)); raw=file.readAll();
        assert(QCryptographicHash::hash(raw,QCryptographicHash::Sha256).toHex()==object["sha256"].toString().toLatin1().toLower());
        formName=object["name"].toString().mid(5);
    }
    for (const auto& value:objects) {
        const auto object=value.toObject(); if (object["kind"]!="form" || object["name"].toString()!=formName) continue;
        QFile file(directory.filePath(object["file"].toString())); assert(file.open(QIODevice::ReadOnly)); formBytes=file.readAll();
        assert(QCryptographicHash::hash(formBytes,QCryptographicHash::Sha256).toHex()==object["sha256"].toString().toLatin1().toLower());
    }
    assert(!raw.isEmpty() && !formBytes.isEmpty());
    const auto layout=parseAccessFormLayout(decodeVbaSource(formBytes));
    const auto schema=parseMdbSchema(AccessCatalog(source).tableSchema(layout.recordSource));
    assert(schema.columns.size()==14);
    int keys=0,counters=0;
    for (const auto& column:schema.columns) { if (column.primaryKeyOrder) ++keys; if (column.autoNumber) ++counters; }
    assert(keys==1 && counters==0); // Real source schema: never invent an AutoNumber.
    std::vector<std::vector<CsvField>> rows(2);
    for (const auto& column:schema.columns) {
        rows[0].push_back({column.name.toUtf8().toStdString(),true});
        if (!column.required && !column.primaryKeyOrder) rows[1].push_back({"\x1e",false});
        else if (column.sourceType=="BOOLEAN" || column.sourceType=="BIT") rows[1].push_back({"0",false});
        else if (column.sqliteType=="INTEGER") rows[1].push_back({"7",false});
        else if (column.sqliteType=="REAL") rows[1].push_back({"1.25",false});
        else if (column.sourceType=="DATE" || column.sourceType=="DATETIME") rows[1].push_back({"2026-11-01 00:00:00",true});
        else { assert(column.sqliteType=="TEXT"); rows[1].push_back({"synthetic",true}); }
    }
    auto db=QSqlDatabase::addDatabase("QSQLITE","original-clipboard"); db.setDatabaseName(":memory:"); assert(db.open()); QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    importAccessTable(db,"synthetic",layout.recordSource,schema,rows);
    const auto code=decodeVbaSource(raw); const auto lines=code.split('\n'); const auto analysis=analyzeVba(code.toStdString()); QString procedure;
    const QRegularExpression copy("DoMenuItem\\s+acFormBar\\s*,\\s*acEditMenu\\s*,\\s*2\\s*,",QRegularExpression::CaseInsensitiveOption);
    const QRegularExpression append("DoMenuItem\\s+acFormBar\\s*,\\s*acEditMenu\\s*,\\s*5\\s*,",QRegularExpression::CaseInsensitiveOption);
    for (std::size_t i=0;i<analysis.procedures.size();++i) {
        const auto& candidate=analysis.procedures[i]; const int end=i+1<analysis.procedures.size() ? analysis.procedures[i+1].line-1 : lines.size();
        const auto body=lines.mid(candidate.line-1,end-candidate.line+1).join('\n');
        if (copy.match(body).hasMatch() && append.match(body).hasMatch()) { assert(procedure.isEmpty()); procedure=QString::fromStdString(candidate.name); }
    }
    assert(!procedure.isEmpty());
    AccessFormPreview form(formName,formBytes,raw,db); assert(!form.cancelledByOpen() && form.hasDataSource()); form.show(); form.activateWindow(); QApplication::processEvents();
    auto root=makePreviewFormObject(form),command=makePreviewApplicationObject(form)->member("DoCmd",{}).object;
    command->member("RunCommand",{VbaValue::numeric(50)}); command->member("RunCommand",{VbaValue::numeric(190)});
    int number=0; try { command->member("RunCommand",{VbaValue::numeric(38)}); } catch (const VbaRuntimeError& e) { number=e.number; } assert(number==3022);
    int messages=0;
    const auto expectConflict=[&] {
        auto* timer=new QTimer(&form); timer->setInterval(1);
        QObject::connect(timer,&QTimer::timeout,&form,[&,timer] {
            auto* box=qobject_cast<QMessageBox*>(QApplication::activeModalWidget()); if (!box) return;
            assert(box->text().contains("Record append failed") && box->button(QMessageBox::Ok));
            timer->stop(); timer->deleteLater(); ++messages; QTest::mouseClick(box->button(QMessageBox::Ok),Qt::LeftButton);
        }); timer->start();
    };
    const auto unchanged=[&] {
        assert(root->member("Recordset",{}).object->member("RecordCount",{}).asNumber()==1);
        auto table=layout.recordSource; table.replace('"',"\"\""); assert(sql.exec("SELECT count(*) FROM \""+table+"\"") && sql.next() && sql.value(0).toInt()==1); sql.finish();
    };
    expectConflict(); std::vector<VbaValue> arguments; const auto result=form.executeProcedure(procedure,arguments);
    assert(result.status==VbaExecutionStatus::Completed && messages==1); unchanged();
    int bindings=0;
    for (const auto& definition:layout.controls) {
        const auto prefix=definition.eventProcPrefix.isEmpty() ? definition.name : definition.eventProcPrefix;
        if (definition.type!="CommandButton" || definition.onClick.compare("[Event Procedure]",Qt::CaseInsensitive)!=0 || (prefix+"_Click").compare(procedure,Qt::CaseInsensitive)!=0) continue;
        auto control=root->member("Controls",{VbaValue::string(definition.name.toStdString())}).object;
        control->assign("Visible",VbaValue::logical(true)); control->assign("Enabled",VbaValue::logical(true));
        auto* button=form.findChild<QPushButton*>("access-control-"+definition.name); assert(button);
        for (auto* area:form.findChildren<QScrollArea*>()) if (area->widget() && area->widget()->isAncestorOf(button)) area->ensureWidgetVisible(button);
        QApplication::processEvents(); expectConflict(); QTest::mouseClick(button,Qt::LeftButton); QApplication::processEvents();
        const auto outcome=form.eventResult("control:"+definition.name+":Click"); assert(outcome && outcome->status==VbaExecutionStatus::Completed && messages==2); unchanged(); ++bindings;
    }
    assert(bindings==1 && form.close()); QApplication::clipboard()->clear();
    std::cout<<"Original duplicate handler: actual manual-PK schema, direct append error 3022, one original invocation and one bound click handle the conflict; no inserted rows\n";
}
