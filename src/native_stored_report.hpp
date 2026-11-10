#pragma once
#include "access_query.hpp"
#include "vba_analyzer.hpp"
#include "vba_conditional.hpp"
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
    QVector<QString> columnTypes;
    const auto columns=accessQueryColumns(database,plan,&columnTypes); // prepare only, never step rows.
    if (columns.isEmpty() || columns.size()>32) throw VbaRuntimeError(3251,"Native report automatic layout supports 1-32 columns");
    const auto captions=definition.value("field_captions").toObject();
    if(definition.contains("field_formats") && !definition.value("field_formats").isObject()) throw VbaRuntimeError(5,"Invalid native field Format catalog");
    const auto formats=definition.value("field_formats").toObject();
    QSet<QString> used;
    for (const auto& column:columns) {
        if (column.isEmpty() || used.contains(column.toCaseFolded())) throw VbaRuntimeError(3251,"Native report columns are empty or ambiguous");
        (void)quoted(column); used.insert(column.toCaseFolded());
    }
    const auto runningValue=definition.value("running_sum_field");
    if(!runningValue.isUndefined() && !runningValue.isString()) throw VbaRuntimeError(5,"Invalid running sum field");
    const auto runningField=runningValue.toString().trimmed();
    if(!runningField.isEmpty()) {
        int index=-1;for(int i=0;i<columns.size();++i) if(columns[i].compare(runningField,Qt::CaseInsensitive)==0) {index=i;break;}
        if(index<0) throw VbaRuntimeError(3265,"Running sum field is not in the report source");
        static const QSet<QString> numeric{"INTEGER","INT","INTEGER1","INTEGER2","INTEGER4","REAL","DOUBLE","SINGLE_REAL","CURRENCY_REAL","CURRENCY_TICKS"};
        if(index>=columnTypes.size() || !numeric.contains(columnTypes[index].trimmed().toUpper()))
            throw VbaRuntimeError(3251,"Running sum requires a proven native numeric source field");
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
    const auto footerValue=definition.value("footer_expression");
    if(!footerValue.isUndefined() && !footerValue.isString()) throw VbaRuntimeError(5,"Invalid native report footer expression");
    const auto footerExpression=footerValue.toString().trimmed();
    if(!footerExpression.isEmpty() && (!footerExpression.startsWith('=') || footerExpression.size()>32768))
        throw VbaRuntimeError(5,"Report footer calculation must start with equals");
    (void)quoted(footerExpression);
    const auto footerForceValue=definition.value("footer_force_new_page");
    const int footerForce=footerForceValue.toInt(0);
    if(!footerForceValue.isUndefined() && (!footerForceValue.isDouble() || footerForce<0 || footerForce>3 || footerForceValue.toDouble()!=footerForce))
        throw VbaRuntimeError(5,"Invalid native report footer page break setting");
    if(footerExpression.isEmpty() && footerForce) throw VbaRuntimeError(5,"Report footer page break requires a footer calculation");
    const auto code=definition.value("vba").toString();
    const auto bytes=code.toUtf8(); const auto analysis=analyzeVba(preprocessVbaConditional(std::string_view(bytes.constData(),bytes.size())));
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
        lines<<"Begin Label"<<"Name ="+quoted(unique("label_"+columns[i]))<<"Caption ="+quoted(captions.value(columns[i].toCaseFolded()).toString(columns[i]))
            <<"Left ="+QString::number(i*width+30)<<"Width ="+QString::number(width-60)<<"Height =300"<<"End";
    lines<<"End"<<"Begin Section"<<"Height =360";
    if(!forceValue.isUndefined()) lines<<"ForceNewPage ="+QString::number(force);
    for (int i=0;i<columns.size();++i) {
        lines<<"Begin TextBox"<<"Name ="+quoted(columns[i])<<"ControlSource ="+quoted("["+columns[i]+"]");
        if(const auto format=formats.value(columns[i].toCaseFolded());!format.isUndefined()) {if(!format.isString()) throw VbaRuntimeError(5,"Invalid native report field Format");lines<<"Format ="+quoted(format.toString());}
        if(!runningField.isEmpty() && columns[i].compare(runningField,Qt::CaseInsensitive)==0) lines<<"RunningSum =2";
        else if(definition.value("can_grow").toBool()) lines<<"CanGrow =-1";
        lines<<"Left ="+QString::number(i*width+30)<<"Top =30"<<"Width ="+QString::number(width-60)<<"Height =270"<<"End";
    }
    lines<<"End";
    if(!footerExpression.isEmpty())
        lines<<"Begin FormFooter"<<"Height =450"<<"ForceNewPage ="+QString::number(footerForce)
            <<"Begin TextBox"<<"Name ="+quoted(unique("native_report_summary"))<<"ControlSource ="+quoted(footerExpression)
            <<"Left =30"<<"Top =30"<<"Width ="+QString::number(reportWidth-60)<<"Height =300"<<"End"<<"End";
    lines<<"Begin PageFooter"<<"Height =360"<<"Begin TextBox"<<"Name ="+quoted(unique("native_report_page"))
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
