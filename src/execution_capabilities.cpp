#include "artifact_execution_observer.hpp"
#include "native_macro.hpp"
#include "execution_capabilities.hpp"
#include "vba_analyzer.hpp"
#include "vba_source_codec.hpp"
#include "access_query.hpp"
#include "native_module_definition.hpp"
#include <QCryptographicHash>
#include <QDateTime>
#include <QJsonDocument>
#include <QSqlDriver>
#include <QSqlQuery>
#include <QSqlError>
#include <QThread>
#include <QVariant>
namespace linaccess {
namespace {
constexpr auto journalKey="la_execution_observations";
void check(QSqlDatabase db) {if(!db.isOpen() || !db.driver()) throw VbaRuntimeError(3420,"Capability report: project is closed");if(db.driver()->thread()!=QThread::currentThread()) throw VbaRuntimeError(3251,"Capability report requires its connection thread");}
QString digest(const QByteArray& bytes) {return QString::fromLatin1(QCryptographicHash::hash(bytes,QCryptographicHash::Sha256).toHex());}
QString label(VbaExecutionStatus status) {switch(status) {case VbaExecutionStatus::Completed:return "completed";case VbaExecutionStatus::ProcedureMissing:return "procedure_missing";case VbaExecutionStatus::Unsupported:return "unsupported";case VbaExecutionStatus::Error:return "error";}return "unknown";}
}
QJsonArray executionObservations(QSqlDatabase db) {check(db);if(db.driver()->property("la_execution_database").toString()!=db.databaseName()) return {};return db.driver()->property(journalKey).value<QJsonArray>();}
void clearExecutionObservations(QSqlDatabase db) {if(!db.driver()) return;if(db.driver()->thread()!=QThread::currentThread()) throw VbaRuntimeError(3251,"Observation reset requires its connection thread");db.driver()->setProperty(journalKey,QVariant{});db.driver()->setProperty("la_execution_database",QVariant{});}
void recordVbaObservation(QSqlDatabase db,std::string_view source,std::string_view procedure,const VbaExecutionResult& result,bool stopped,const QString& context) {
    check(db);auto rows=executionObservations(db);while(rows.size()>=1024) rows.removeAt(0);
    rows.append(QJsonObject{{"kind","vba"},{"source_sha256",digest(QByteArray(source.data(),source.size()))},{"procedure",QString::fromUtf8(procedure.data(),procedure.size()).left(512)},{"context",context},{"status",stopped ? "stopped" : label(result.status)},{"line",result.line},{"error",result.errorNumber},{"cancelled",result.cancelled},{"utc",QDateTime::currentDateTimeUtc().toString(Qt::ISODateWithMs)}});
    db.driver()->setProperty("la_execution_database",db.databaseName());db.driver()->setProperty(journalKey,QVariant::fromValue(rows));
}
void recordArtifactExecution(QSqlDatabase db,const QString& kind,const QString& name,const QByteArray& source,const QString& status,int error) {
    check(db);auto rows=executionObservations(db);while(rows.size()>=1024) rows.removeAt(0);rows.append(QJsonObject{{"kind",kind},{"name",name.left(512)},{"source_sha256",digest(source)},{"status",status},{"error",error},{"utc",QDateTime::currentDateTimeUtc().toString(Qt::ISODateWithMs)}});db.driver()->setProperty("la_execution_database",db.databaseName());db.driver()->setProperty(journalKey,QVariant::fromValue(rows));
}
QJsonObject projectCapabilityReport(QSqlDatabase db) {
    check(db);VbaEnumAliases aliases;int typeEnvironmentError=0;
    try {aliases=projectVbaEnumAliases(db);}catch(const VbaRuntimeError& failure) {typeEnvironmentError=failure.number;}catch(const std::exception&) {typeEnvironmentError=3251;}
    const auto observed=executionObservations(db);QJsonArray artifacts;
    const auto classify=[&](QString origin,QString kind,QString name,QString status,QByteArray bytes) {
        const auto hash=digest(bytes);QJsonObject item{{"origin",origin},{"kind",kind},{"name",name},{"stored_status",status},{"source_sha256",hash},{"capability","unclassified"},{"execution","not_observed"}};
        if(kind=="vba" || kind=="module") {
            const auto text=decodeVbaSource(bytes).toUtf8();const auto codeHash=digest(text);item["decoded_source_sha256"]=codeHash;QJsonArray procedures;
            if(bytes.isEmpty()) item["capability"]="source_unavailable";
            else if(text.size()>1024*1024) item["capability"]="source_limit_exceeded";
            else {const auto analysis=analyzeVba(text.toStdString());item["capability"]="procedure_inventory_only";
                for(const auto& procedure:analysis.procedures) {const auto executionName=vbaProcedureExecutionName(procedure);const auto validation=validateVbaProcedure(text.toStdString(),executionName,&aliases);QJsonObject entry{{"name",QString::fromStdString(procedure.name)},{"kind",QString::fromStdString(procedure.kind)},{"line",procedure.line},{"capability",validation.status==VbaExecutionStatus::Completed ? "syntax_accepted_runtime_unproven" : label(validation.status)},{"validation_line",validation.line},{"validation_error",validation.errorNumber},{"execution","not_observed"}};QJsonArray history;
                    for(const auto& value:observed) {const auto row=value.toObject();if(row["kind"]=="vba" && row["source_sha256"].toString()==codeHash && row["procedure"].toString().compare(QString::fromStdString(executionName),Qt::CaseInsensitive)==0) history.append(row);}
                    if(!history.isEmpty()) {entry["execution"]="matching_source_observed";entry["last_source_result"]=history.last().toObject()["status"];entry["attribution"]="same_source_and_procedure; object_instance_not_proven";entry["observations"]=history;}procedures.append(entry);
                }
            }item["procedures"]=procedures;item["execution"]="see_procedures";
        } else if(kind=="query") {try {(void)compileAccessQuery(decodeVbaSource(bytes));item["capability"]="sql_compiled_runtime_unproven";}catch(const std::exception&) {item["capability"]="sql_compile_failed";}}
        else if(kind=="form" || kind=="report") item["capability"]="document_requires_live_context";
        else if(kind=="macro") {if(origin=="native") {try {(void)parseNativeMacro(QJsonDocument::fromJson(bytes).object());item["capability"]="macro_syntax_accepted_runtime_unproven";}catch(const std::exception&) {item["capability"]="macro_definition_failed";}}else item["capability"]="macro_requires_project_context";}
        if(kind!="vba" && kind!="module") {QJsonArray history;for(const auto& value:observed) {const auto row=value.toObject();if(!bytes.isEmpty() && row["kind"]==kind && row["source_sha256"].toString()==hash) history.append(row);}if(!history.isEmpty()) {item["execution"]="matching_source_observed";item["last_source_result"]=history.last().toObject()["status"];item["observations"]=history;item["attribution"]="same_source; object_instance_not_proven";}}
        artifacts.append(item);
    };
    QSqlQuery query(db);if(!query.exec("SELECT kind,name,status,source_bytes FROM main._la_access_artifacts ORDER BY id")) throw VbaRuntimeError(3251,"Capability report cannot read imported artifact catalog");
    while(query.next()) classify("imported",query.value(0).toString(),query.value(1).toString(),query.value(2).toString(),query.value(3).toByteArray());query.finish();
    if(!query.exec("SELECT kind,name,definition FROM main._la_objects ORDER BY kind,name")) throw VbaRuntimeError(3251,"Capability report cannot read native object catalog");
    while(query.next()) {const auto kind=query.value(0).toString();const auto name=query.value(1).toString();auto bytes=query.value(2).toByteArray();const auto definition=QJsonDocument::fromJson(bytes).object();
        if(kind=="module") {try {bytes=nativeModuleSource(definition);}catch(const std::exception&) {artifacts.append(QJsonObject{{"origin","native"},{"kind",kind},{"name",name},{"source_sha256",digest(bytes)},{"capability","invalid_definition"},{"execution","not_observed"}});continue;}}
        else if(kind=="query") {if(!definition["sql"].isString()) {artifacts.append(QJsonObject{{"origin","native"},{"kind",kind},{"name",name},{"source_sha256",digest(bytes)},{"capability","invalid_definition"},{"execution","not_observed"}});continue;}bytes=definition["sql"].toString().toUtf8();}
        classify("native",kind,name,"native",bytes);
        if((kind=="form" || kind=="report") && definition["vba"].isString() && !definition["vba"].toString().isEmpty()) classify("native_codebehind","vba",name,"native",definition["vba"].toString().toUtf8());
    }
    query.finish();if(!query.exec("SELECT name,sql FROM main.sqlite_schema WHERE type='table' AND name NOT GLOB 'sqlite_*' AND name NOT GLOB '_la_*' ORDER BY name")) throw VbaRuntimeError(3251,"Capability report cannot read table catalog");
    while(query.next()) artifacts.append(QJsonObject{{"origin","physical"},{"kind","table"},{"name",query.value(0).toString()},{"source_sha256",digest(query.value(1).toByteArray())},{"capability","schema_present_data_unproven"},{"execution","not_observed"}});
    return {{"format","linaccess-capabilities-v1"},{"scope","current_project_session"},{"enum_type_environment",typeEnvironmentError ? "unavailable" : "available"},{"enum_type_environment_error",typeEnvironmentError},{"semantics","Syntax acceptance is not execution or Access compatibility. Observations cover actual attempted instrumented VBA, macro and SELECT calls only; at most 1024 retained. No business procedures were started by this report."},{"artifacts",artifacts},{"observations",observed}};
}
}
