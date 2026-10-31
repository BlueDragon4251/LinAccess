#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include <QLineEdit>
#include <QComboBox>
#include <QPushButton>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

void testReportEvents() {
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
Sub Report_Open(Cancel As Integer)
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
