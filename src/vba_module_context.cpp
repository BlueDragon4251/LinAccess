#include "vba_module_context.hpp"
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
std::vector<std::shared_ptr<VbaModuleState>> publicVbaModuleStates(QSqlDatabase db) {
    (void)context(db);
    std::vector<std::shared_ptr<VbaModuleState>> states;
    QSqlQuery query(db);
    if (!query.exec("SELECT name,source_bytes FROM _la_access_artifacts WHERE kind='vba' ORDER BY id DESC")) return states;
    QSet<QString> names;
    for (;query.next();) {
        const auto name=query.value(0).toString().toCaseFolded(); if (names.contains(name)) continue; names.insert(name);
        const auto source=decodeVbaSource(query.value(1).toByteArray()).toUtf8().toStdString();
        if (isVbaClassModule(source)) continue;
        auto state=standardVbaModuleState(db,source);
        if (!state->publicVariables.empty()) states.push_back(std::move(state));
    }
    return states;
}
void resetVbaProjectState(QSqlDatabase db) {
    if (!db.driver()) return;
    if (QThread::currentThread()!=db.driver()->thread()) throw VbaRuntimeError(3251,"VBA project reset requires its connection thread");
    const auto children=db.driver()->children();
    for (auto* child:children) if (auto* state=dynamic_cast<ProjectModuleContext*>(child)) delete state;
}
}
