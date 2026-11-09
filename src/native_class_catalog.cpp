#include "native_class_catalog.hpp"
#include "native_class_definition.hpp"
#include "native_standard_module_host.hpp"
#include "vba_class.hpp"
#include "dao_database.hpp"
#include <QJsonDocument>
#include <QSqlDriver>
#include <QSqlQuery>
#include <QSqlError>
#include <QMainWindow>
#include <QThread>
namespace linaccess {
namespace {
struct ClassRevision {QString name;QByteArray encoded,source;};
std::optional<ClassRevision> revision(QSqlDatabase db,const QString& requested) {
    if(!db.isOpen() || !db.driver()) throw VbaRuntimeError(3420,"Class project is closed");
    if(db.driver()->thread()!=QThread::currentThread()) throw VbaRuntimeError(3251,"Class project requires its connection thread");
    QSqlQuery query(db);if(!query.exec("SELECT type FROM main.sqlite_schema WHERE name='_la_objects' COLLATE NOCASE")) throw VbaRuntimeError(3146,"Class catalog metadata unavailable");
    if(!query.next()) return {};if(query.value(0).toString()!="table") throw VbaRuntimeError(3251,"Class catalog must be physical main metadata");query.finish();
    if(!query.exec("SELECT 1 FROM temp.sqlite_schema WHERE name='_la_objects' COLLATE NOCASE")) throw VbaRuntimeError(3146,"Class catalog shadow check failed");if(query.next()) throw VbaRuntimeError(3251,"Shadowed class catalog is unsupported");query.finish();
    query.prepare("SELECT name,definition FROM main._la_objects WHERE kind='class' AND name=? COLLATE NOCASE");query.addBindValue(requested);if(!query.exec()) throw VbaRuntimeError(3146,"Class catalog unavailable");if(!query.next()) return {};
    ClassRevision row{query.value(0).toString(),query.value(1).toByteArray(),{}};if(query.next()) throw VbaRuntimeError(5,"Ambiguous native class name");if(query.lastError().isValid()) throw VbaRuntimeError(3146,"Class catalog fetch failed");query.finish();
    if(row.encoded.size()>4*1024*1024) throw VbaRuntimeError(3035,"Native class definition is too large");
    QJsonParseError error;const auto document=QJsonDocument::fromJson(row.encoded,&error);if(error.error!=QJsonParseError::NoError || !document.isObject()) throw VbaRuntimeError(13,"Invalid class JSON");row.source=nativeClassModuleSource(document.object(),row.name);return row;
}
std::shared_ptr<const NativeVbaClassDefinition> compiled(QSqlDatabase db,const ClassRevision& row) {
    const auto sources=standardVbaModuleSources(db);if(!sources.contains(row.name.toCaseFolded()) || sources.value(row.name.toCaseFolded())!=row.source) throw VbaRuntimeError(5,"Native class namespace is ambiguous");
    return compileNativeVbaClass(row.name.toStdString(),row.source.toStdString());
}
}
bool nativeProjectClassAvailable(QSqlDatabase db,std::string_view name) {const auto row=revision(db,QString::fromUtf8(name.data(),name.size()));if(!row) return false;(void)compiled(db,*row);return true;}
std::shared_ptr<VbaObject> constructNativeProjectClass(QSqlDatabase db,QWidget* caller,std::string_view name,std::function<void()> standaloneGuard) {
    const auto row=revision(db,QString::fromUtf8(name.data(),name.size()));if(!row) throw VbaRuntimeError(3251,"Native project class is unavailable");const auto definition=compiled(db,*row);
    QWidget* window=caller;bool project=false;
    for(auto* parent=caller;parent;parent=parent->parentWidget()) if(qobject_cast<QMainWindow*>(parent) && parent->property("la_project_generation").isValid()) {window=parent;project=true;break;}
    const QPointer<QWidget> owner=window;const auto generation=owner ? owner->property("la_project_generation") : QVariant{};const auto identity=db.databaseName(),connection=db.connectionName();
    auto database=makeDaoDatabase(db);auto valid=std::make_shared<bool>(true);const auto expected=*row;
    const auto check=[db,owner,generation,identity,connection,database,valid,expected,project,standaloneGuard] {
        if(!*valid) throw VbaRuntimeError(3420,"Native class instance was invalidated");
        try {
            if(!owner || !db.isOpen() || db.databaseName()!=identity || db.connectionName()!=connection || owner->property("la_project_generation")!=generation) throw VbaRuntimeError(3420,"Native class project lifetime ended");
            (void)database->member("Name",{});if(!project && standaloneGuard) standaloneGuard();
            const auto current=revision(db,expected.name);if(!current || current->name!=expected.name || current->encoded!=expected.encoded) throw VbaRuntimeError(3420,"Native class source revision changed");
        } catch(...) {*valid=false;throw;}
    };
    VbaClassContext context;context.check=check;context.host=[db,owner,check]() -> std::unique_ptr<VbaHost> {check();return std::make_unique<NativeStandardModuleHost>(owner,db,check);};
    return constructNativeVbaClass(definition,std::move(context));
}
}
