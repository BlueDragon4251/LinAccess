#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QLineEdit>
#include <QApplication>
#include <QDir>
#include <QFile>
#include <QJsonArray>
#include <QJsonObject>
#include <QCryptographicHash>
#include <QRegularExpression>
#include <QPushButton>
#include "vba_source_codec.hpp"
#include "access_query.hpp"
#include <QSqlRecord>
#include <QMainWindow>
#include <QMenuBar>
#include <QMenu>
#include <iostream>
#include <QToolBar>
#include <QComboBox>
#include <QSqlQuery>
#include <QTemporaryDir>
#include <QProcess>
#include <QStandardPaths>
#include <QPrintPreviewDialog>
#include <QPrintPreviewWidget>
#include <QTimer>
#include <QAction>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

void testOpenReport();

void testReportEvents() {
    testOpenReport();
    auto db = QSqlDatabase::addDatabase("QSQLITE","report-event-tests");
    db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,name TEXT)"));
    assert(query.exec("INSERT INTO records VALUES(1,'first'),(2,'second')"));
    assert(query.exec("CREATE TABLE events(id INTEGER PRIMARY KEY,name TEXT)"));
    assert(query.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('query','ranked',?)");
    query.addBindValue(QByteArray("SELECT TOP 1 id,name FROM records ORDER BY id")); assert(query.exec()); query.finish();
    const QByteArray source = R"REPORT(Begin Report
Caption ="Report title"
RecordSource ="ranked"
OnOpen ="[Event Procedure]"
OnClose ="[Event Procedure]"
OnNoData ="[Event Procedure]"
Begin Section
Height =1500
Begin TextBox
Name ="Name"
ControlSource ="name"
End
Begin TextBox
Name ="Info"
Top =400
End
End
End
)REPORT";
    const QByteArray code = R"VBA(
Private retainedReport As Report
Sub Report_Open(Cancel As Integer)
Dim typedReport As Access.Report
Set typedReport=Me
Set retainedReport=typedReport
If Not TypeOf Me Is Access.Report Then Err.Raise 5
If TypeOf Me Is Access.Form Then Err.Raise 5
If Not TypeOf Reports Is Access.Reports Then Err.Raise 5
CurrentDb.Execute "INSERT INTO events(name) VALUES('open')"
Me!Info = CStr(Reports.Count) & ":" & CStr(Forms.Count)
End Sub
Sub Report_NoData(Cancel As Integer)
CurrentDb.Execute "INSERT INTO events(name) VALUES('nodata')"
Cancel = 0
End Sub
Sub Report_Close()
CurrentDb.Execute "INSERT INTO events(name) VALUES('close')"
End Sub
)VBA";
    const auto trace = [&] {
        QStringList rows;
        assert(query.exec("SELECT name FROM events ORDER BY id"));
        while (query.next()) rows.append(query.value(0).toString());
        query.finish(); return rows;
    };
    const auto clear = [&] { assert(query.exec("DELETE FROM events")); query.finish(); };
    auto bandedSource=source;
    bandedSource.replace("Begin Section\nHeight =1500",R"BANDS(Begin FormHeader
Height =300
End
Begin PageHeader
Height =600
Begin TextBox
Name ="PageTitle"
Top =150
End
End
Begin PageFooter
Height =450
Begin TextBox
Name ="PageNumber"
Top =150
End
End
Begin FormFooter
Height =300
Begin TextBox
Name ="ReportSummary"
Top =150
End
End
Begin Section
Height =1500)BANDS");
    const auto bandedLayout=linaccess::parseAccessFormLayout(QString::fromUtf8(bandedSource));
    assert(bandedLayout.sectionHeights.value("PageHeader")==600);
    assert(bandedLayout.sectionHeights.value("PageFooter")==450);
    assert(bandedLayout.sectionOffsets.value("Section")==900);
    assert(bandedLayout.sectionOffsets.value("FormFooter")==2400);
    assert(bandedLayout.sectionOffsets.value("PageFooter")==2700 && bandedLayout.height==3150);
    {
        AccessFormPreview banded("banded report",bandedSource,{},db);
        const auto field=[&](const char* name) {
            auto* widget=banded.findChild<QLineEdit*>(QString("access-control-")+name);
            assert(widget); return widget;
        };
        assert(field("Name")->text()=="first" && field("Name")->isReadOnly());
        assert(field("PageTitle")->y()==30 && field("Name")->y()==60);
        assert(field("ReportSummary")->y()==170 && field("PageNumber")->y()==190);
        QTemporaryDir output; assert(output.isValid());
        const auto single=output.filePath("single.pdf");
        // The summary extends outside its short section: do not silently clip it.
        assert(!banded.exportReportPdf(single));
        assert(field("Name")->text()=="first");
        auto pagedSource=bandedSource;
        pagedSource.replace("RecordSource =\"ranked\"","RecordSource =\"records\"");
        pagedSource.replace("Height =1500","Height =9000");
        pagedSource.replace("Height =300","Height =600");
        AccessFormPreview paged("paged report",pagedSource,{},db);
        const auto document=output.filePath("records.pdf");
        assert(paged.exportReportPdf(document));
        QFile pdf(document); assert(pdf.open(QIODevice::ReadOnly));
        const auto bytes=pdf.readAll();
        const QRegularExpression pagePattern("/Type\\s*/Page(?:\\s|/|>)");
        auto pages=pagePattern.globalMatch(QString::fromLatin1(bytes));
        int pageCount=0; while(pages.hasNext()) { pages.next(); ++pageCount; }
        assert(pageCount==2);
        const auto reader=QStandardPaths::findExecutable("pdftotext");
        if (!reader.isEmpty()) {
            QProcess textReader; textReader.start(reader,{"-layout",document,"-"});
            assert(textReader.waitForFinished(10000) && textReader.exitCode()==0);
            const auto text=QString::fromUtf8(textReader.readAllStandardOutput())
                .normalized(QString::NormalizationForm_KC);
            assert(text.count("first")==1 && text.count("second")==1);
            assert(text.indexOf("first")<text.indexOf("second"));
        }

        assert(paged.findChild<QLineEdit*>("access-control-Name")->text()=="first");
        auto* previewAction=paged.findChild<QAction*>("access-report-preview");
        assert(previewAction && paged.findChild<QAction*>("access-report-print"));
        bool previewChecked=false;
        QTimer::singleShot(0,[&] {
            auto* dialog=qobject_cast<QPrintPreviewDialog*>(QApplication::activeModalWidget());
            assert(dialog);
            auto* pages=dialog->findChild<QPrintPreviewWidget*>();
            assert(pages && pages->pageCount()==2);
            previewChecked=true; dialog->reject();
        });
        previewAction->trigger();
        assert(previewChecked && paged.findChild<QLineEdit*>("access-control-Name")->text()=="first");
        assert(paged.close());
        auto growingSource=pagedSource; growingSource.replace("Name =\"Name\"","Name =\"Name\"\nCanGrow =-1");
        AccessFormPreview growing("growing report",growingSource,{},db);
        assert(!growing.exportReportPdf(document));
        QFile unchanged(document); assert(unchanged.open(QIODevice::ReadOnly) && unchanged.readAll()==bytes);
        assert(growing.close());
        assert(banded.close());
    }
    AccessFormPreview report("report",source,code,db);
    assert(!report.cancelledByOpen() && report.hasDataSource());
    assert(report.eventResult("OnOpen") && report.eventResult("OnOpen")->status==linaccess::VbaExecutionStatus::Completed);
    assert(!report.eventResult("OnNoData"));
    auto* name = report.findChild<QLineEdit*>("access-control-Name");
    auto* info = report.findChild<QLineEdit*>("access-control-Info");
    auto* save = report.findChild<QPushButton*>("access-save-record");
    assert(name && name->text()=="first" && name->isReadOnly());
    assert(info && info->text()=="1:0"); assert(save && !save->isEnabled());
    assert(trace()==QStringList{"open"});
    auto reports = makePreviewReportsObject(report);
    assert(reports->member("Count",{}).asNumber()==1);
    assert(makePreviewFormsObject(report)->member("Count",{}).asNumber()==0);
    assert(reports->item("report").object->member("Name",{}).asText()=="report");
    bool readOnly = false;
    try { reports->item("report").object->assignItem("Name",linaccess::VbaValue::string("forbidden")); }
    catch(const linaccess::VbaRuntimeError& error) { readOnly=error.number==2448; }
    assert(readOnly && name->text()=="first");
    assert(query.exec("SELECT name FROM records WHERE id=1") && query.next() && query.value(0).toString()=="first"); query.finish();
    bool missing = false;
    try { (void)reports->item("missing"); } catch(const linaccess::VbaRuntimeError& error) { missing=error.number==2451; }
    assert(missing);
    assert(report.close());
    assert(report.eventResult("OnClose") && report.eventResult("OnClose")->status==linaccess::VbaExecutionStatus::Completed);
    assert(trace()==(QStringList{"open","close"}));
    clear();
    auto tableSource = source; tableSource.replace("RecordSource =\"ranked\"","RecordSource =\"records\"");
    AccessFormPreview tableReport("table report",tableSource,code,db);
    assert(tableReport.hasDataSource());
    assert(tableReport.findChild<QLineEdit*>("access-control-Name")->isReadOnly());
    assert(!tableReport.findChild<QPushButton*>("access-save-record")->isEnabled());
    assert(tableReport.close()); clear();
    auto cancelCode = code;
    cancelCode.replace("Sub Report_Open(Cancel As Integer)","Sub Report_Open(Cancel As Integer)\nCancel = -1");
    AccessFormPreview cancelled("cancelled",source,cancelCode,db);
    assert(cancelled.cancelledByOpen() && !cancelled.hasDataSource());
    assert(cancelled.eventResult("OnOpen")->cancelled && !cancelled.eventResult("OnNoData"));
    assert(trace()==QStringList{"open"}); clear();
    assert(query.exec("DELETE FROM records")); query.finish();
    AccessFormPreview empty("empty",source,code,db);
    assert(empty.hasDataSource() && !empty.cancelledByOpen());
    assert(empty.eventResult("OnNoData") && empty.eventResult("OnNoData")->status==linaccess::VbaExecutionStatus::Completed);
    assert(trace()==(QStringList{"open","nodata"}));
    assert(empty.close()); clear();
    auto noDataCancel = code; noDataCancel.replace("Cancel = 0","Cancel = -1");
    AccessFormPreview noRows("no rows",source,noDataCancel,db);
    assert(noRows.cancelledByOpen() && noRows.eventResult("OnNoData")->cancelled);
    assert(trace()==(QStringList{"open","nodata"})); clear();
    auto unavailable = source; unavailable.replace("RecordSource =\"ranked\"","RecordSource =\"missing\"");
    AccessFormPreview missingSource("missing source",unavailable,code,db);
    assert(!missingSource.hasDataSource() && !missingSource.eventResult("OnNoData"));
    assert(trace()==QStringList{"open"}); assert(missingSource.close());
}

void testOpenReport() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","open-report-command"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,name TEXT)"));
    assert(sql.exec("INSERT INTO records VALUES(1,'first'),(2,'second')"));
    assert(sql.exec("CREATE TABLE events(value TEXT)"));
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    const QByteArray layout=R"REPORT(Begin Report
RecordSource ="records"
OnOpen ="[Event Procedure]"
OnNoData ="[Event Procedure]"
OnClose ="[Event Procedure]"
Begin Section
Begin TextBox
Name ="value"
ControlSource ="name"
End
End
End
)REPORT";
    const QByteArray code=R"VBA(Sub Report_Open(Cancel As Integer)
CurrentDb.Execute "INSERT INTO events VALUES('open')"
If Me.Tag="cancel" Then Cancel=1
End Sub
Sub Report_NoData(Cancel As Integer)
CurrentDb.Execute "INSERT INTO events VALUES('nodata')"
Cancel=1
End Sub
Sub Report_Close()
CurrentDb.Execute "INSERT INTO events VALUES('close')"
End Sub
)VBA";
    const auto artifact=[&](const QString& kind,const QString& name,const QByteArray& bytes) {
        sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES(?,?,?)"); sql.addBindValue(kind); sql.addBindValue(name); sql.addBindValue(bytes); assert(sql.exec()); sql.finish();
    };
    for (const auto* name:{"summary","empty","cancelled"}) {
        auto definition=layout; if(QString::fromLatin1(name)=="cancelled") definition.replace("Begin Report\n","Begin Report\nTag =\"cancel\"\n");
        artifact("report",name,definition); artifact("vba","Report_"+QString::fromLatin1(name),code);
    }
    auto missingSource=layout; missingSource.replace("RecordSource =\"records\"","RecordSource =\"missing_table\"");
    artifact("report","missing_source",missingSource); artifact("vba","Report_missing_source",code);
    artifact("report","failed",layout); artifact("vba","Report_failed","Sub Report_Open(Cancel As Integer)\nErr.Raise 17\nEnd Sub\n");
    artifact("report","kind_mismatch","Begin Form\nEnd\n");
    const QByteArray controller=R"VBA(Sub ShowReport()
DoCmd.OpenReport "summary",acViewPreview,,"[id]=2"
Me.Tag="caller retained"
End Sub
Sub ShowAgain()
Application.DoCmd.OpenReport "SUMMARY",acViewPreview,,"[id]=2",acWindowNormal
End Sub
)VBA";
    QMainWindow project;
    AccessFormPreview caller("controller","Begin Form\nEnd\n",controller,db,&project); caller.show(); QApplication::setActiveWindow(&caller); QApplication::processEvents();
    auto callerRoot=makePreviewFormObject(caller); auto commands=makePreviewApplicationObject(caller)->member("DoCmd",{}).object;
    const auto eventCount=[&](const char* event) { sql.prepare("SELECT COUNT(*) FROM events WHERE value=?"); sql.addBindValue(event); assert(sql.exec() && sql.next()); const int n=sql.value(0).toInt(); sql.finish(); return n; };
    std::vector<VbaValue> none; assert(caller.executeProcedure("ShowReport",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(callerRoot->member("Tag",{}).asText()=="caller retained" && eventCount("open")==1);
    assert(makePreviewReportsObject(caller)->member("Count",{}).asNumber()==1);
    AccessFormPreview* report=nullptr;
    for(auto* widget:QApplication::topLevelWidgets()) if(auto* candidate=dynamic_cast<AccessFormPreview*>(widget);candidate && candidate!=&caller && candidate->isVisible() && makePreviewFormObject(*candidate)->member("Name",{}).asText()=="summary") { assert(!report); report=candidate; }
    assert(report && report->parentWidget()==&project && report->findChild<QToolBar*>("access-print-preview-toolbar"));
    auto reportRoot=makePreviewFormObject(*report); auto* field=report->findChild<QLineEdit*>("access-control-value"); assert(field && field->text()=="second" && field->isReadOnly());
    assert(reportRoot->member("Filter",{}).asText()=="[id]=2");
    assert(caller.executeProcedure("ShowAgain",none).status==linaccess::VbaExecutionStatus::Completed && eventCount("open")==1);
    const auto invokeError=[&](std::vector<VbaValue> args,int expected) {
        int error=0; try { commands->member("OpenReport",args); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
        if(error!=expected) std::cerr<<"Report error contract: expected="<<expected<<", actual="<<error<<'\n';
        assert(error==expected);
    };
    invokeError({},450); invokeError({VbaValue{}},94);
    invokeError({VbaValue::string("summary")},3251);
    invokeError({VbaValue::string("summary"),VbaValue::numeric(1)},3251);
    invokeError({VbaValue::string("summary"),VbaValue::numeric(2147483648.0)},6);
    invokeError({VbaValue::string("summary"),VbaValue::numeric(2),VbaValue::string("filter")},3251);
    invokeError({VbaValue::string("summary"),VbaValue::numeric(2),VbaValue::empty(),VbaValue::empty(),VbaValue::numeric(1)},3251);
    invokeError({VbaValue::string("missing"),VbaValue::numeric(2)},2103);
    assert(eventCount("open")==1);
    invokeError({VbaValue::string("cancelled"),VbaValue::numeric(2)},2501);
    invokeError({VbaValue::string("empty"),VbaValue::numeric(2),VbaValue::empty(),VbaValue::string("[id]=999")},2501);
    assert(eventCount("nodata")==1 && makePreviewReportsObject(caller)->member("Count",{}).asNumber()==1);
    invokeError({VbaValue::string("missing_source"),VbaValue::numeric(2)},3078);
    invokeError({VbaValue::string("failed"),VbaValue::numeric(2)},17);
    invokeError({VbaValue::string("kind_mismatch"),VbaValue::numeric(2)},2103);
    assert(eventCount("nodata")==1 && makePreviewReportsObject(caller)->member("Count",{}).asNumber()==1);

    assert(caller.close() && report->isVisible() && field->text()=="second");
    auto reportCommands=makePreviewApplicationObject(*report)->member("DoCmd",{}).object;
    reportCommands->member("Close",{VbaValue::numeric(3),VbaValue::string("summary")});
    assert(eventCount("close")==1);
    int closed=0; try { reportRoot->member("Name",{}); } catch(const linaccess::VbaRuntimeError& error) { closed=error.number; } assert(closed==2450);
    QApplication::processEvents();
}

void testOriginalOpenReport(const QDir& directory,const QJsonArray& objects,QSqlDatabase db) {
    const QString moduleHash="6e8e0a0e41e1a1768cff4a0b79a07205c7b0f799eb8b465af5397843ce048ed8";
    QByteArray code,layout; QString formName;
    const auto read=[&](const QJsonObject& artifact) {
        QFile file(directory.filePath(artifact["file"].toString())); assert(file.open(QIODevice::ReadOnly)); const auto bytes=file.readAll();
        assert(QCryptographicHash::hash(bytes,QCryptographicHash::Sha256).toHex()==artifact["sha256"].toString().toLatin1()); return bytes;
    };
    for(const auto& entry:objects) { const auto item=entry.toObject(); if(item["kind"]=="vba" && item["sha256"].toString()==moduleHash) { code=read(item); formName=item["name"].toString().mid(5); } }
    assert(!code.isEmpty());
    for(const auto& entry:objects) { const auto item=entry.toObject(); if(item["kind"]=="form" && item["name"].toString()==formName) layout=read(item); }
    assert(!layout.isEmpty());
    const QRegularExpression procedures("^\\s*(?:(?:Private|Public)\\s+)?Sub\\s+([\\p{L}_][\\p{L}\\p{N}_]*)\\s*\\([^\\n]*\\)([\\s\\S]*?)^\\s*End\\s+Sub",QRegularExpression::MultilineOption|QRegularExpression::CaseInsensitiveOption);
    const QRegularExpression preview("DoCmd\\s*\\.\\s*OpenReport\\s+([^,\\r\\n]+),\\s*acPreview\\b",QRegularExpression::CaseInsensitiveOption);
    QString procedure,target; auto matches=procedures.globalMatch(linaccess::decodeVbaSource(code));
    while(matches.hasNext()) {
        const auto candidate=matches.next(); const auto call=preview.match(candidate.captured(2)); if(!call.hasMatch()) continue;
        assert(procedure.isEmpty()); procedure=candidate.captured(1); const auto argument=call.captured(1).trimmed();
        if(argument.startsWith('"') && argument.endsWith('"')) target=argument.mid(1,argument.size()-2);
        else { const QRegularExpression value("\\b"+QRegularExpression::escape(argument)+"\\s*=\\s*\"([^\"]+)\"",QRegularExpression::CaseInsensitiveOption); const auto literal=value.match(candidate.captured(2)); assert(literal.hasMatch()); target=literal.captured(1); }
    }
    assert(!procedure.isEmpty() && !target.isEmpty());
    QByteArray reportBytes;
    for(const auto& entry:objects) { const auto item=entry.toObject(); if(item["kind"]=="report" && item["name"].toString()==target) {
        assert(item["sha256"].toString()=="30f1ebe8559d44c1ca9f965948b55e556460d560286f669cac8f43a298741a6f"); reportBytes=read(item);
    } }
    assert(!reportBytes.isEmpty()); const auto reportLayout=linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(reportBytes));
    const auto callerLayout=linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(layout)); QString buttonName;
    for(const auto& control:callerLayout.controls) {
        const auto prefix=control.eventProcPrefix.isEmpty() ? control.name : control.eventProcPrefix;
        if(control.type=="CommandButton" && control.onClick.compare("[Event Procedure]",Qt::CaseInsensitive)==0 && (prefix+"_Click").compare(procedure,Qt::CaseInsensitive)==0) buttonName=control.name;
    }
    assert(!buttonName.isEmpty()); QSqlQuery guard(db); assert(guard.exec("SAVEPOINT original_open_report"));
    QMainWindow project; auto* toolbar=project.addToolBar("Project"); toolbar->addAction("Test command"); project.menuBar()->addMenu("Project")->addAction(toolbar->actions().first());
    AccessFormPreview caller(formName,layout,code,db,&project); caller.show(); QApplication::setActiveWindow(&caller); QApplication::processEvents();
    const auto verify=[&] {
        assert(makePreviewReportsObject(caller)->member("Count",{}).asNumber()==1);
        AccessFormPreview* opened=nullptr;
        for(auto* widget:QApplication::topLevelWidgets()) if(auto* candidate=dynamic_cast<AccessFormPreview*>(widget); candidate && candidate!=&caller && candidate->isVisible() && makePreviewFormObject(*candidate)->member("Name",{}).asText()==target.toStdString()) { assert(!opened); opened=candidate; }
        assert(opened && opened->hasDataSource() && opened->findChild<QToolBar*>("access-print-preview-toolbar"));
        if(const auto event=opened->eventResult("OnOpen")) assert(event->status==linaccess::VbaExecutionStatus::Completed);
        QString querySource;
        if(db.tables().contains(reportLayout.recordSource,Qt::CaseInsensitive)) { auto table=reportLayout.recordSource; table.replace('"',"\"\""); querySource="SELECT * FROM \""+table+"\""; }
        else if(reportLayout.recordSource.startsWith("SELECT",Qt::CaseInsensitive) || reportLayout.recordSource.startsWith("PARAMETERS",Qt::CaseInsensitive)) querySource=reportLayout.recordSource;
        else querySource=linaccess::accessQuerySource(db,reportLayout.recordSource);
        const auto plan=linaccess::compileAccessQuery(querySource); QSqlQuery reference(db); assert(reference.exec(plan.sql) && reference.next());
        const auto columns=reference.record(); int cells=0;
        for(const auto& control:reportLayout.controls) {
            if(control.type!="TextBox") continue; auto fieldName=control.controlSource.trimmed();
            if(fieldName.startsWith('[') && fieldName.endsWith(']')) fieldName=fieldName.mid(1,fieldName.size()-2);
            const auto index=columns.indexOf(fieldName); if(index<0) continue;
            auto* field=opened->findChild<QLineEdit*>("access-control-"+control.name); assert(field && field->isReadOnly());
            const auto value=reference.value(index); assert(field->property("la_null").toBool()==value.isNull() && field->text()==value.toString()); if(!value.isNull()) ++cells;
        }
        assert(cells>0); reference.finish();
        makePreviewApplicationObject(caller)->member("DoCmd",{}).object->member("Close",{linaccess::VbaValue::numeric(3),linaccess::VbaValue::string(target.toStdString())});
    };
    std::vector<linaccess::VbaValue> none;
    const auto outcome=caller.executeProcedure(procedure,none);
    std::cerr<<"Original report preview: status="<<static_cast<int>(outcome.status)<<", error="<<outcome.errorNumber<<", line="<<outcome.line<<'\n';
    assert(outcome.status==linaccess::VbaExecutionStatus::Completed); verify();
    auto* button=caller.findChild<QPushButton*>("access-control-"+buttonName); assert(button);
    button->setEnabled(true); button->show(); button->click(); verify();
    const auto clicked=caller.eventResult("control:"+buttonName+":Click"); assert(clicked && clicked->status==linaccess::VbaExecutionStatus::Completed);
    assert(caller.close()); assert(guard.exec("ROLLBACK TO original_open_report") && guard.exec("RELEASE original_open_report"));
}
