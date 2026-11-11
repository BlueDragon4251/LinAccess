#include "native_class_definition.hpp"
#include "vba_module_context.hpp"
#include "vba_random.hpp"
#include "execution_capabilities.hpp"
#include "native_module_definition.hpp"
#include <QJsonDocument>
#include <QJsonObject>
#include <QSqlError>
#include <QElapsedTimer>
#include "vba_expression.hpp"
#include "vba_source_codec.hpp"
#include <QSqlDatabase>
#include <QSqlDriver>
#include <QSqlQuery>
#include <QVariant>
#include <QObject>
#include <QThread>
#include <QPointer>
#include <QRegularExpression>
#include <QSet>
#include <sqlite3.h>
#include <cctype>

namespace linaccess {
namespace {
class ProjectModuleContext final:public QObject {
public:
    explicit ProjectModuleContext(QSqlDatabase db):QObject(db.driver()),database(db.databaseName()) {}
    QString database;
    std::shared_ptr<VbaObject> error;
    std::shared_ptr<VbaRandomState> random;
    std::unordered_map<std::string,std::shared_ptr<VbaModuleState>> frames;
    VbaEnumAliases enums;sqlite3_int64 enumChanges{-1};int enumData{-1},enumSchema{-1},enumTemp{-1};
};
ProjectModuleContext& context(QSqlDatabase db) {
    if (!db.isOpen() || !db.driver()) throw VbaRuntimeError(3420,"VBA project is closed");
    if (QThread::currentThread()!=db.driver()->thread()) throw VbaRuntimeError(3251,"VBA project requires its connection thread");
    for (auto* child:db.driver()->children()) if (auto* state=dynamic_cast<ProjectModuleContext*>(child)) {
        if (state->database==db.databaseName()) return *state;
        delete state; break;
    }
    return *new ProjectModuleContext(db);
}
}
bool isVbaClassModule(std::string_view source) {
    static const QRegularExpression pattern("^\\s*VERSION[^\\r\\n]*\\bCLASS\\s*$",QRegularExpression::MultilineOption|QRegularExpression::CaseInsensitiveOption);
    return pattern.match(QString::fromUtf8(source.data(),source.size())).hasMatch();
}
std::shared_ptr<VbaModuleState> standardVbaModuleState(QSqlDatabase db,std::string_view source,const VbaEnumAliases* suppliedAliases) {
    if (isVbaClassModule(source)) throw VbaRuntimeError(3251,"CodeBehind state requires its document instance");
    const auto aliases=suppliedAliases ? *suppliedAliases : projectVbaEnumAliases(db);
    auto& frames=context(db).frames; const auto key=vbaModuleBindingKey(source,aliases);
    if (auto found=frames.find(key);found!=frames.end()) return found->second;
    auto state=createVbaModuleState(source,&aliases); frames.emplace(key,state); return state;
}
QMap<QString,QByteArray> standardVbaModuleSources(QSqlDatabase db,QMap<QString,QString>* canonicalNames,QSet<QString>* standaloneClasses) {
    (void)context(db);QElapsedTimer timer;timer.start();
    QMap<QString,QByteArray> sources;QMap<QString,QString> names;QSet<QString> classes,documents;
    QSqlQuery table(db);
    if(!table.exec("SELECT type FROM main.sqlite_schema WHERE name='_la_access_artifacts' COLLATE NOCASE")) throw VbaRuntimeError(3251,"VBA SQL metadata unavailable");
    if(table.next()) {
        if(table.value(0).toString()!="table") throw VbaRuntimeError(3251,"VBA SQL requires a physical source archive");
        table.finish();
        QSqlQuery any(db);
        if(!any.exec("SELECT 1 FROM main._la_access_artifacts WHERE kind='vba' LIMIT 1")) throw VbaRuntimeError(3251,"VBA SQL archive shape unsupported");
        const bool hasCode=any.next();if(any.lastError().isValid()) throw VbaRuntimeError(3146,"VBA source catalog fetch failed"); any.finish();
        if(hasCode) {
            QSqlQuery shadow(db);
            if(!shadow.exec("SELECT 1 FROM temp.sqlite_schema WHERE name='_la_access_artifacts' COLLATE NOCASE AND type IN ('table','view')")) throw VbaRuntimeError(3251,"VBA SQL shadow check failed");
            if(shadow.next()) throw VbaRuntimeError(3251,"VBA SQL archive shadow is unsupported");
            shadow.finish();
            QSqlQuery rows(db);
            if(!rows.exec("SELECT name,source_bytes FROM main._la_access_artifacts WHERE kind='vba' ORDER BY id DESC")) throw VbaRuntimeError(3251,"VBA SQL archive ordering unavailable");
            qsizetype bytes=0;
            qsizetype count=0;
            while(rows.next()) {
                if(++count>1000000 || timer.hasExpired(5000)) throw VbaRuntimeError(3251,"VBA source catalog scan limit exceeded");
                const auto name=rows.value(0).toString(),key=name.toCaseFolded();
                if(sources.contains(key)) continue;
                auto source=rows.value(1).toByteArray(); bytes+=source.size();
                if(sources.size()>=256 || bytes>16*1024*1024) throw VbaRuntimeError(3251,"VBA SQL module inventory limit exceeded");
                names.insert(key,name);sources.insert(key,std::move(source));
            }
            if(rows.lastError().isValid()) throw VbaRuntimeError(3146,"VBA SQL archive fetch failed");
            rows.finish();
            if(!rows.exec("SELECT kind,name FROM main._la_access_artifacts WHERE kind IN ('form','report')")) throw VbaRuntimeError(3146,"Class document association metadata unavailable");
            count=0;while(rows.next()) {
                if(++count>1000000 || timer.hasExpired(5000)) throw VbaRuntimeError(3251,"Class document association scan limit exceeded");
                documents.insert((rows.value(0).toString()+"_"+rows.value(1).toString()).toCaseFolded());
            }
            if(rows.lastError().isValid()) throw VbaRuntimeError(3146,"Class document association fetch failed");
        }
    }

    table.finish();
    if(!table.exec("SELECT type FROM main.sqlite_schema WHERE name='_la_objects' COLLATE NOCASE")) throw VbaRuntimeError(3251,"Native module catalog metadata unavailable");
    if(table.next()) {
        if(table.value(0).toString()!="table") throw VbaRuntimeError(3251,"Native modules require a physical project catalog");table.finish();
        QSqlQuery any(db);if(!any.exec("SELECT 1 FROM main._la_objects WHERE kind IN ('module','class') LIMIT 1")) throw VbaRuntimeError(3251,"Native module catalog shape unsupported");const bool present=any.next();if(any.lastError().isValid()) throw VbaRuntimeError(3146,"Native module catalog probe failed");any.finish();
        if(present) {
            QSqlQuery shadow(db);if(!shadow.exec("SELECT 1 FROM temp.sqlite_schema WHERE name='_la_objects' COLLATE NOCASE AND type IN ('table','view')")) throw VbaRuntimeError(3251,"Native module shadow check failed");if(shadow.next()) throw VbaRuntimeError(3251,"Native module catalog shadow is unsupported");shadow.finish();
            qsizetype bytes=0,count=0;for(const auto& source:sources) bytes+=source.size();
            QSqlQuery rows(db);rows.setForwardOnly(true);if(!rows.exec("SELECT kind,name,definition FROM main._la_objects WHERE kind IN ('module','class')")) throw VbaRuntimeError(3146,"Native module catalog unavailable");
            while(rows.next()) {
                if(++count>1000000 || timer.hasExpired(5000)) throw VbaRuntimeError(3251,"Native module catalog scan limit exceeded");
                const auto kind=rows.value(0).toString(),name=rows.value(1).toString(),key=name.toCaseFolded();
                if(!(kind=="class" ? nativeClassNameValid(name) : nativeModuleNameValid(name))) throw VbaRuntimeError(3251,"Invalid native standard module name");
                if(sources.contains(key)) throw VbaRuntimeError(5,"Ambiguous native/imported module name");
                const auto encoded=rows.value(2).toByteArray();if(encoded.size()>4*1024*1024) throw VbaRuntimeError(3035,"Native module definition limit exceeded");
                QJsonParseError error;const auto document=QJsonDocument::fromJson(encoded,&error);if(error.error!=QJsonParseError::NoError || !document.isObject()) throw VbaRuntimeError(13,"Invalid native module document");
                auto source=kind=="class" ? nativeClassModuleSource(document.object(),name) : nativeModuleSource(document.object());bytes+=source.size();if(sources.size()>=256 || bytes>16*1024*1024) throw VbaRuntimeError(3251,"VBA module inventory limit exceeded");
                if(kind=="class") classes.insert(key);names.insert(key,name);sources.insert(key,std::move(source));
            }
            if(rows.lastError().isValid()) throw VbaRuntimeError(3146,"Native module catalog fetch failed");
        }
    }
    for(auto it=sources.begin();it!=sources.end();++it) if(!classes.contains(it.key()) && !documents.contains(it.key()) && standaloneVbaClassHeader(names.value(it.key()),it.value())) classes.insert(it.key());
    if(canonicalNames) *canonicalNames=std::move(names);if(standaloneClasses) *standaloneClasses=std::move(classes);return sources;
}
VbaEnumAliases projectVbaEnumAliases(QSqlDatabase db) {
    auto& state=context(db);const auto handle=db.driver()->handle();
    if(!handle.isValid() || QByteArray(handle.typeName())!="sqlite3*") throw VbaRuntimeError(3251,"Project type binding requires native SQLite");
    auto* native=*static_cast<sqlite3* const*>(handle.constData());if(!native) throw VbaRuntimeError(3420,"Project type binding connection is closed");
    const auto pragma=[&](const char* command) {QSqlQuery query(db);if(!query.exec(QString::fromLatin1(command)) || !query.next()) throw VbaRuntimeError(3251,"Project type metadata revision unavailable");return query.value(0).toInt();};
    const auto changes=sqlite3_total_changes64(native);const int data=pragma("PRAGMA main.data_version"),schema=pragma("PRAGMA main.schema_version"),temporary=pragma("PRAGMA temp.schema_version");
    if(state.enumChanges==changes && state.enumData==data && state.enumSchema==schema && state.enumTemp==temporary) return state.enums;
    QMap<QString,QString> names;const auto sources=standardVbaModuleSources(db,&names);VbaEnumAliases aliases;
    std::unordered_map<std::string,int> bare;
    for(auto item=sources.cbegin();item!=sources.cend();++item) {
        const auto source=decodeVbaSource(item.value()).toUtf8().toStdString();
        auto owner=names.value(item.key()).toUtf8().toStdString();for(auto& ch:owner) ch=static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
        for(const auto& name:nativeVbaEnumTypeNames(source,true)) {
            const auto dot=name.find('.');
            if(dot!=std::string::npos) {if(name.substr(0,dot)!=owner) throw VbaRuntimeError(5,"Public Enum module identity does not match its project owner");continue;}
            ++bare[name];if(!owner.empty()) ++bare[owner+"."+name];
        }
    }
    for(const auto& [name,count]:bare) if(count==1 && (name.find('.')!=std::string::npos || !sources.contains(QString::fromStdString(name).toCaseFolded()))) aliases.insert(name);
    state.enums=std::move(aliases);state.enumChanges=changes;state.enumData=data;state.enumSchema=schema;state.enumTemp=temporary;
    return state.enums;
}
std::vector<std::shared_ptr<VbaModuleState>> publicVbaModuleStates(QSqlDatabase db,const VbaEnumAliases* suppliedAliases) {
    (void)context(db);
    std::vector<std::shared_ptr<VbaModuleState>> states;
    const auto aliases=suppliedAliases ? *suppliedAliases : projectVbaEnumAliases(db);
    const auto sources=standardVbaModuleSources(db);
    for (const auto& bytes:sources) {
        const auto source=decodeVbaSource(bytes).toUtf8().toStdString();
        if (isVbaClassModule(source)) continue;
        auto state=standardVbaModuleState(db,source,&aliases);
        if (!state->publicVariables.empty()) states.push_back(std::move(state));
    }
    return states;
}
std::shared_ptr<VbaRandomState> projectVbaRandomState(QSqlDatabase db) {
    auto& state=context(db);if(!state.random) state.random=std::make_shared<VbaRandomState>();return state.random;
}
std::shared_ptr<VbaObject> projectVbaErrorObject(QSqlDatabase db) {
    auto& state=context(db);
    if(!state.error) {
        const QPointer<ProjectModuleContext> owner=&state;
        const QPointer<QSqlDriver> driver=db.driver();
        const auto connection=db.connectionName(),database=db.databaseName();
        state.error=makeVbaErrorObject([owner,driver,connection,database] {
            if(!owner || !driver) throw VbaRuntimeError(3420,"VBA error environment is closed");
            if(QThread::currentThread()!=driver->thread()) throw VbaRuntimeError(3251,"VBA error environment requires its connection thread");
            if(!QSqlDatabase::contains(connection)) throw VbaRuntimeError(3420,"VBA error connection is closed");
            const auto current=QSqlDatabase::database(connection,false);
            if(!current.isOpen() || current.driver()!=driver || current.databaseName()!=database) throw VbaRuntimeError(3420,"VBA error environment has changed");
        });
    }
    return state.error;
}
void resetVbaProjectState(QSqlDatabase db) {
    clearExecutionObservations(db);
    if (!db.driver()) return;
    if (QThread::currentThread()!=db.driver()->thread()) throw VbaRuntimeError(3251,"VBA project reset requires its connection thread");
    const auto children=db.driver()->children();
    for (auto* child:children) if (auto* state=dynamic_cast<ProjectModuleContext*>(child)) delete state;
}
}
