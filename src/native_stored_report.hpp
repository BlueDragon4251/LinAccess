#pragma once
#include "access_query.hpp"
#include "vba_analyzer.hpp"
#include "vba_runtime.hpp"
#include <QJsonDocument>
#include <QJsonObject>
#include <QSet>
#include <QSqlQuery>

namespace linaccess {
struct NativeStoredReport { QString name; QByteArray layout,code; QString recordSource; };
inline NativeStoredReport nativeStoredReport(QString name,const QJsonObject& definition,QSqlDatabase database) {
    const auto quoted=[](QString value) {
        if (value.contains('\r') || value.contains('\n') || value.contains(QChar(0)))
            throw VbaRuntimeError(3251,"Native report metadata contains unsupported control characters");
        value.replace('"',"\"\""); return QString("\"")+value+"\"";
    };
    if (!database.isOpen()) throw VbaRuntimeError(3146,"Native report database unavailable");
    const auto source=definition.value("sql").toString();
    if (source.contains(QChar(0))) throw VbaRuntimeError(3251,"Native report SQL contains NUL");
    const auto plan=compileAccessQuery(source);
    if (!plan.makeTable.isEmpty() || !plan.sql.trimmed().startsWith("SELECT",Qt::CaseInsensitive))
        throw VbaRuntimeError(3251,"Native report requires a read-only SELECT source");
    if (installAccessQueryViews(database)<0) throw VbaRuntimeError(3251,"Native report query views are busy");
    const auto columns=accessQueryColumns(database,plan); // prepare only, never step rows.
    if (columns.isEmpty() || columns.size()>32) throw VbaRuntimeError(3251,"Native report automatic layout supports 1-32 columns");
    QSet<QString> used;
    for (const auto& column:columns) {
        if (column.isEmpty() || used.contains(column.toCaseFolded())) throw VbaRuntimeError(3251,"Native report columns are empty or ambiguous");
        (void)quoted(column); used.insert(column.toCaseFolded());
    }
    const auto unique=[&used](QString base) {
        auto result=base; int suffix=1;
        while (used.contains(result.toCaseFolded())) result=base+QString::number(suffix++);
        used.insert(result.toCaseFolded()); return result;
    };
    const auto forceValue=definition.value("detail_force_new_page");
    const int force=forceValue.toInt(-1);
    if(!forceValue.isUndefined() && (!forceValue.isDouble() || force<0 || force>3 || forceValue.toDouble()!=force))
        throw VbaRuntimeError(5,"Invalid native report detail page break setting");
    const auto headerForceValue=definition.value("header_force_new_page");
    const int headerForce=headerForceValue.toInt(-1);
    if(!headerForceValue.isUndefined() && (!headerForceValue.isDouble() || headerForce<0 || headerForce>3 || headerForceValue.toDouble()!=headerForce))
        throw VbaRuntimeError(5,"Invalid native report header page break setting");
    const auto code=definition.value("vba").toString();
    const auto bytes=code.toUtf8(); const auto analysis=analyzeVba(std::string_view(bytes.constData(),bytes.size()));
    QSet<QString> procedures;
    for (const auto& procedure:analysis.procedures) procedures.insert(QString::fromStdString(procedure.name).toCaseFolded());
    const int width=qMax(900,9000/int(columns.size())),reportWidth=width*columns.size();
    QStringList lines{"Begin Report","Caption ="+quoted(definition.value("title").toString(name)),"Width ="+QString::number(reportWidth)};
    for (const auto& event:{QString("Open"),QString("Load"),QString("NoData"),QString("Close")})
        if (procedures.contains(("Report_"+event).toCaseFolded())) lines<<"On"+event+" =\"[Event Procedure]\"";
    lines<<"Begin FormHeader"<<"Height =600";
    if(!headerForceValue.isUndefined()) lines<<"ForceNewPage ="+QString::number(headerForce);
    lines<<"Begin Label"<<"Name ="+quoted(unique("native_report_title"))
        <<"Caption ="+quoted(definition.value("title").toString(name))<<"Width ="+QString::number(reportWidth-60)
        <<"Height =450"<<"FontWeight =700"<<"End"<<"End"<<"Begin PageHeader"<<"Height =450";
    for (int i=0;i<columns.size();++i)
        lines<<"Begin Label"<<"Name ="+quoted(unique("label_"+columns[i]))<<"Caption ="+quoted(columns[i])
            <<"Left ="+QString::number(i*width+30)<<"Width ="+QString::number(width-60)<<"Height =300"<<"End";
    lines<<"End"<<"Begin Section"<<"Height =360";
    if(!forceValue.isUndefined()) lines<<"ForceNewPage ="+QString::number(force);
    for (int i=0;i<columns.size();++i) {
        lines<<"Begin TextBox"<<"Name ="+quoted(columns[i])<<"ControlSource ="+quoted("["+columns[i]+"]");
        if(definition.value("can_grow").toBool()) lines<<"CanGrow =-1";
        lines<<"Left ="+QString::number(i*width+30)<<"Top =30"<<"Width ="+QString::number(width-60)<<"Height =270"<<"End";
    }
    lines<<"End"<<"Begin PageFooter"<<"Height =360"<<"Begin TextBox"<<"Name ="+quoted(unique("native_report_page"))
        <<"ControlSource =\"=CStr([Page]) & Chr(47) & CStr([Pages])\""<<"Top =30"<<"Width =900"<<"Height =270"<<"End"<<"End"<<"End";
    return {std::move(name),lines.join('\n').toUtf8(),bytes,source};
}
inline std::optional<NativeStoredReport> loadNativeStoredReport(QSqlDatabase database,const QString& name) {
    if (!database.isOpen() || !database.tables().contains("_la_objects",Qt::CaseInsensitive)) return {};
    QSqlQuery query(database); QString canonical; int rows=0;
    if (!query.exec("SELECT name FROM main._la_objects WHERE kind='report'")) throw VbaRuntimeError(3146,"Native report catalog query failed");
    while (query.next()) {
        if (++rows>1000000) throw VbaRuntimeError(3251,"Native report catalog limit exceeded");
        const auto candidate=query.value(0).toString();
        if (candidate.compare(name,Qt::CaseInsensitive)!=0) continue;
        if (!canonical.isEmpty()) throw VbaRuntimeError(5,"Native report name is ambiguous");
        canonical=candidate;
    }
    query.finish(); if (canonical.isEmpty()) return {};
    if (database.tables().contains("_la_access_artifacts",Qt::CaseInsensitive)) {
        if (!query.exec("SELECT name FROM main._la_access_artifacts WHERE kind='report'")) throw VbaRuntimeError(3146,"Imported report catalog query failed");
        while (query.next()) {
            if (++rows>1000000) throw VbaRuntimeError(3251,"Native report catalog limit exceeded");
            if (query.value(0).toString().compare(canonical,Qt::CaseInsensitive)==0) throw VbaRuntimeError(5,"Stored and imported report names conflict");
        }
        query.finish();
    }
    query.prepare("SELECT definition FROM main._la_objects WHERE kind='report' AND name=?"); query.addBindValue(canonical);
    if (!query.exec() || !query.next()) throw VbaRuntimeError(3146,"Native report definition query failed");
    QJsonParseError error; const auto document=QJsonDocument::fromJson(query.value(0).toByteArray(),&error); query.finish();
    if (error.error!=QJsonParseError::NoError || !document.isObject()) throw VbaRuntimeError(3251,"Native report definition is invalid");
    return nativeStoredReport(canonical,document.object(),database);
}
}
