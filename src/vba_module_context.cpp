#include "vba_module_context.hpp"
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
#include <QRegularExpression>
#include <QSet>

namespace linaccess {
namespace {
class ProjectModuleContext final:public QObject {
public:
    explicit ProjectModuleContext(QSqlDatabase db):QObject(db.driver()),database(db.databaseName()) {}
    QString database;
    std::unordered_map<std::string,std::shared_ptr<VbaModuleState>> frames;
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
std::shared_ptr<VbaModuleState> standardVbaModuleState(QSqlDatabase db,std::string_view source) {
    if (isVbaClassModule(source)) throw VbaRuntimeError(3251,"CodeBehind state requires its document instance");
    auto& frames=context(db).frames; const auto key=std::string(source);
    if (auto found=frames.find(key);found!=frames.end()) return found->second;
    auto state=createVbaModuleState(source); frames.emplace(key,state); return state;
}
QMap<QString,QByteArray> standardVbaModuleSources(QSqlDatabase db,QMap<QString,QString>* canonicalNames) {
    (void)context(db);QElapsedTimer timer;timer.start();
    QMap<QString,QByteArray> sources;QMap<QString,QString> names;
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
        }
    }

    table.finish();
    if(!table.exec("SELECT type FROM main.sqlite_schema WHERE name='_la_objects' COLLATE NOCASE")) throw VbaRuntimeError(3251,"Native module catalog metadata unavailable");
    if(table.next()) {
        if(table.value(0).toString()!="table") throw VbaRuntimeError(3251,"Native modules require a physical project catalog");table.finish();
        QSqlQuery any(db);if(!any.exec("SELECT 1 FROM main._la_objects WHERE kind='module' LIMIT 1")) throw VbaRuntimeError(3251,"Native module catalog shape unsupported");const bool present=any.next();if(any.lastError().isValid()) throw VbaRuntimeError(3146,"Native module catalog probe failed");any.finish();
        if(present) {
            QSqlQuery shadow(db);if(!shadow.exec("SELECT 1 FROM temp.sqlite_schema WHERE name='_la_objects' COLLATE NOCASE AND type IN ('table','view')")) throw VbaRuntimeError(3251,"Native module shadow check failed");if(shadow.next()) throw VbaRuntimeError(3251,"Native module catalog shadow is unsupported");shadow.finish();
            qsizetype bytes=0,count=0;for(const auto& source:sources) bytes+=source.size();
            QSqlQuery rows(db);rows.setForwardOnly(true);if(!rows.exec("SELECT name,definition FROM main._la_objects WHERE kind='module'")) throw VbaRuntimeError(3146,"Native module catalog unavailable");
            while(rows.next()) {
                if(++count>1000000 || timer.hasExpired(5000)) throw VbaRuntimeError(3251,"Native module catalog scan limit exceeded");
                const auto name=rows.value(0).toString(),key=name.toCaseFolded();
                if(!nativeModuleNameValid(name)) throw VbaRuntimeError(3251,"Invalid native standard module name");
                if(sources.contains(key)) throw VbaRuntimeError(5,"Ambiguous native/imported module name");
                const auto encoded=rows.value(1).toByteArray();if(encoded.size()>4*1024*1024) throw VbaRuntimeError(3035,"Native module definition limit exceeded");
                QJsonParseError error;const auto document=QJsonDocument::fromJson(encoded,&error);if(error.error!=QJsonParseError::NoError || !document.isObject()) throw VbaRuntimeError(13,"Invalid native module document");
                auto source=nativeModuleSource(document.object());bytes+=source.size();if(sources.size()>=256 || bytes>16*1024*1024) throw VbaRuntimeError(3251,"VBA module inventory limit exceeded");
                names.insert(key,name);sources.insert(key,std::move(source));
            }
            if(rows.lastError().isValid()) throw VbaRuntimeError(3146,"Native module catalog fetch failed");
        }
    }
    if(canonicalNames) *canonicalNames=std::move(names);return sources;
}
std::vector<std::shared_ptr<VbaModuleState>> publicVbaModuleStates(QSqlDatabase db) {
    (void)context(db);
    std::vector<std::shared_ptr<VbaModuleState>> states;
    const auto sources=standardVbaModuleSources(db);
    for (const auto& bytes:sources) {
        const auto source=decodeVbaSource(bytes).toUtf8().toStdString();
        if (isVbaClassModule(source)) continue;
        auto state=standardVbaModuleState(db,source);
        if (!state->publicVariables.empty()) states.push_back(std::move(state));
    }
    return states;
}
void resetVbaProjectState(QSqlDatabase db) {
    clearExecutionObservations(db);
    if (!db.driver()) return;
    if (QThread::currentThread()!=db.driver()->thread()) throw VbaRuntimeError(3251,"VBA project reset requires its connection thread");
    const auto children=db.driver()->children();
    for (auto* child:children) if (auto* state=dynamic_cast<ProjectModuleContext*>(child)) delete state;
}
}
