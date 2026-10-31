#include "access_form_preview.hpp"
#include "access_form_import.hpp"
#include "access_table_import.hpp"
#include "access_query.hpp"
#include "sqlite_execution.hpp"
#include "vba_source_codec.hpp"
#include <QCoreApplication>
#include <QCryptographicHash>
#include <QDir>
#include <QFile>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QLineEdit>
#include <QComboBox>
#include <QPushButton>
#include <QRegularExpression>
#include <QSqlQuery>
#include <QSqlDriver>
#include <QSqlRecord>
#include <QSqlError>
#include <QTemporaryDir>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>
namespace {
QString failureCategory(const QString& error) {
    const auto text=error.toCaseFolded();
    if (text.contains("no such column")) return "unresolved_column";
    if (text.contains("no such table")) return "missing_dependency";
    if (text.contains("no such function")) return "unsupported_function";
    if (text.contains("ambiguous")) return "ambiguous_name";
    if (text.contains("aggregate")) return "aggregate_usage";
    if (text.contains("interrupted")) return "execution_budget";
    if (text.contains("unrecognized token")) {
        const auto match=QRegularExpression("unrecognized token: \"([^\"]*)\"").match(error);
        const auto token=match.captured(1);
        if (token.size()==1 && !token[0].isLetterOrNumber())
            return "unrecognized_token_u"+QString::number(token[0].unicode(),16).rightJustified(4,'0');
        return "unrecognized_token_length_"+QString::number(token.size());
    }
    if (text.contains("syntax")) return "unsupported_grammar";
    return "other_sql_failure";
}
}
void testReportCorpus(const QDir& directory,const QJsonArray& objects,const QString& mdb) {
    QTemporaryDir temporary; assert(temporary.isValid());
    auto db=QSqlDatabase::addDatabase("QSQLITE","report-corpus");
    db.setDatabaseName(temporary.filePath("copy.linaccess")); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
    assert(linaccess::importAccessDatabase(db,mdb)==29);
    assert(query.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    for (const auto& value : objects) {
        const auto object=value.toObject();
        if (object["kind"]!="query") continue;
        QFile file(directory.filePath(object["file"].toString())); assert(file.open(QIODevice::ReadOnly));
        const auto bytes=file.readAll();
        assert(QCryptographicHash::hash(bytes,QCryptographicHash::Sha256).toHex()==object["sha256"].toString().toLatin1().toLower());
        query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('query',?,?)");
        query.addBindValue(object["name"].toString()); query.addBindValue(bytes); assert(query.exec()); query.finish();
    }
    assert(linaccess::installAccessQueryViews(db)>0);
    int scanned=0,available=0,cells=0;
    QJsonArray inventory;
    QMap<QString,int> failures;
    for (const auto& value : objects) {
        const auto object=value.toObject(); if (object["kind"]!="report") continue;
        QFile file(directory.filePath(object["file"].toString())); assert(file.open(QIODevice::ReadOnly));
        const auto bytes=file.readAll();
        assert(QCryptographicHash::hash(bytes,QCryptographicHash::Sha256).toHex()==object["sha256"].toString().toLatin1().toLower());
        const auto layout=linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(bytes));
        assert(layout.kind==linaccess::AccessDocumentKind::Report);
        // Source modules are deliberately absent: measure actual record binding,
        // without confusing this structural/data test with VBA execution.
        AccessFormPreview report(object["name"].toString(),bytes,{},db);
        QString status;
        int reportCells=0;
        if (report.hasDataSource()) {
            ++available;
            const auto source=layout.recordSource.trimmed();
            linaccess::AccessQueryPlan plan;
            if (db.tables().contains(source,Qt::CaseInsensitive))
                plan.sql="SELECT * FROM "+db.driver()->escapeIdentifier(source,QSqlDriver::TableName);
            else plan=linaccess::compileAccessQuery(source.startsWith("SELECT",Qt::CaseInsensitive)
                ? source : linaccess::accessQuerySource(db,source));
            const auto bindings=linaccess::bindAccessQueryParameters(plan,{});
            linaccess::SqliteExecutionBudget budget(db);
            QSqlQuery reference(db); assert(reference.prepare(plan.sql));
            for (auto it=bindings.cbegin();it!=bindings.cend();++it) reference.bindValue(it.key(),it.value());
            assert(reference.exec()); const bool hasRow=reference.next();
            assert((report.findChild<QComboBox*>("access-record-selector")->count()>1)==hasRow);
            for (const auto& control : layout.controls) {
                if (control.controlSource.isEmpty() || control.controlSource.startsWith('=')) continue;
                if (auto* line=report.findChild<QLineEdit*>("access-control-"+control.name)) {
                    assert(line->isReadOnly());
                    auto field=control.controlSource.trimmed();
                    if (field.startsWith('[') && field.endsWith(']')) field=field.mid(1,field.size()-2);
                    const int column=reference.record().indexOf(field);
                    if (hasRow && column>=0) {
                        assert(line->text()==reference.value(column).toString());
                        assert(line->property("la_null").toBool()==reference.value(column).isNull());
                    }
                    if (!line->property("la_null").toBool()) ++reportCells;
                }
            }
            cells+=reportCells;
            status=reportCells ? "bound_cells_no_access_reference_oracle" : "source_executed_no_bound_cells";
        } else {
            status="unbound_source";
            try {
                const auto source=layout.recordSource.trimmed();
                const auto sql=source.startsWith("SELECT",Qt::CaseInsensitive) || source.startsWith("PARAMETERS",Qt::CaseInsensitive)
                    ? source : linaccess::accessQuerySource(db,source);
                const auto plan=linaccess::compileAccessQuery(sql);
                if (!plan.sql.trimmed().startsWith("SELECT",Qt::CaseInsensitive)) throw linaccess::VbaRuntimeError(3065,"Report source must be SELECT");
                const auto bindings=linaccess::bindAccessQueryParameters(plan,{});
                linaccess::SqliteExecutionBudget budget(db);
                QSqlQuery diagnostic(db);
                bool ok=diagnostic.prepare(plan.sql);
                if (ok) {
                    for (auto it=bindings.cbegin();it!=bindings.cend();++it) diagnostic.bindValue(it.key(),it.value());
                    ok=diagnostic.exec();
                }
                if (!ok) status=failureCategory(diagnostic.lastError().text());
                else {
                    while (diagnostic.next()) {}
                    if (diagnostic.lastError().isValid()) status=failureCategory(diagnostic.lastError().text());
                }
            } catch (const linaccess::VbaRuntimeError& error) { status="adapter_error_"+QString::number(error.number); }
            ++failures[status];
        }
        auto* save=report.findChild<QPushButton*>("access-save-record"); assert(save && !save->isEnabled());
        inventory.append(QJsonObject{{"artifact_sha256",object["sha256"]},{"status",status},{"bound_cells",reportCells},{"vba_executed",false}});
        assert(report.close()); ++scanned;
    }
    assert(scanned==9);
    QFile output(QCoreApplication::applicationDirPath()+"/private-report-inventory.json"); assert(output.open(QIODevice::WriteOnly));
    assert(output.write(QJsonDocument(inventory).toJson())>0);
    std::cout<<"Original reports: "<<scanned<<", executed sources: "<<available<<", non-NULL bound controls: "<<cells<<'\n';
    for (auto it=failures.cbegin();it!=failures.cend();++it) std::cout<<"Report source category "<<it.key().toStdString()<<": "<<it.value()<<'\n';
    assert(available>0 && cells>0);
}
