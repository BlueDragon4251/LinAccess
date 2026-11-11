#include "vba_project_types.hpp"
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
struct ClassRevision {QString name;QByteArray encoded,source;bool imported{};qint64 archiveId{};};
bool catalog(QSqlDatabase db,const QString& table) {
    QSqlQuery query(db);query.prepare("SELECT type FROM main.sqlite_schema WHERE name=? COLLATE NOCASE");query.addBindValue(table);
    if(!query.exec()) throw VbaRuntimeError(3146,"Class catalog metadata unavailable");if(!query.next()) return false;
    if(query.value(0).toString()!="table") throw VbaRuntimeError(3251,"Class catalog must be physical main metadata");query.finish();
    query.prepare("SELECT 1 FROM temp.sqlite_schema WHERE name=? COLLATE NOCASE");query.addBindValue(table);if(!query.exec()) throw VbaRuntimeError(3146,"Class catalog shadow check failed");if(query.next()) throw VbaRuntimeError(3251,"Shadowed class catalog is unsupported");return true;
}
std::optional<ClassRevision> revision(QSqlDatabase db,const QString& requested) {
    if(!db.isOpen() || !db.driver()) throw VbaRuntimeError(3420,"Class project is closed");
    if(db.driver()->thread()!=QThread::currentThread()) throw VbaRuntimeError(3251,"Class project requires its connection thread");
    if(!nativeClassNameValid(requested)) return {};
    QSqlQuery query(db);
    if(catalog(db,"_la_objects")) {
        query.prepare("SELECT name,definition FROM main._la_objects WHERE kind='class' AND name=? COLLATE NOCASE");query.addBindValue(requested);if(!query.exec()) throw VbaRuntimeError(3146,"Class catalog unavailable");
        if(query.next()) {
            ClassRevision row{query.value(0).toString(),query.value(1).toByteArray(),{}};
            if(query.next()) throw VbaRuntimeError(5,"Ambiguous native class name");if(query.lastError().isValid()) throw VbaRuntimeError(3146,"Class catalog fetch failed");query.finish();
            if(row.encoded.size()>4*1024*1024) throw VbaRuntimeError(3035,"Native class definition is too large");
            QJsonParseError error;const auto document=QJsonDocument::fromJson(row.encoded,&error);if(error.error!=QJsonParseError::NoError || !document.isObject()) throw VbaRuntimeError(13,"Invalid class JSON");row.source=nativeClassModuleSource(document.object(),row.name);return row;
        }
        if(query.lastError().isValid()) throw VbaRuntimeError(3146,"Class catalog fetch failed");query.finish();
    }
    if(!catalog(db,"_la_access_artifacts")) return {};
    query.prepare("SELECT id,name,source_bytes FROM main._la_access_artifacts WHERE kind='vba' AND name=? COLLATE NOCASE ORDER BY id DESC LIMIT 1");query.addBindValue(requested);
    if(!query.exec()) throw VbaRuntimeError(3146,"Imported class source unavailable");if(!query.next()) return {};
    ClassRevision row{query.value(1).toString(),query.value(2).toByteArray(),{},true,query.value(0).toLongLong()};query.finish();
    if(row.encoded.size()>1024*1024) throw VbaRuntimeError(3035,"Imported class source exceeds one MiB");
    if(!standaloneVbaClassHeader(row.name,row.encoded)) return {};
    QString kind,suffix;
    if(row.name.startsWith("Form_",Qt::CaseInsensitive)) {kind="form";suffix=row.name.mid(5);}
    else if(row.name.startsWith("Report_",Qt::CaseInsensitive)) {kind="report";suffix=row.name.mid(7);}
    if(!kind.isEmpty()) {
        query.prepare("SELECT 1 FROM main._la_access_artifacts WHERE kind=? AND name=? COLLATE NOCASE LIMIT 1");query.addBindValue(kind);query.addBindValue(suffix);if(!query.exec()) throw VbaRuntimeError(3146,"Imported class document association unavailable");if(query.next()) return {};query.finish();
    }
    const auto text=decodeVbaSource(row.encoded);
    const bool fallback=!row.encoded.startsWith("\xEF\xBB\xBF") && !row.encoded.startsWith("\xFF\xFE") && QString::fromUtf8(row.encoded).toUtf8()!=row.encoded;
    if(fallback && text.contains(QChar::ReplacementCharacter)) throw VbaRuntimeError(3251,"Imported class encoding has unrepresentable bytes");
    row.source=nativeClassModuleSource(QJsonObject{{"version",1},{"vba",text}},row.name);return row;
}
void proveNamespace(QSqlDatabase db,const ClassRevision& row) {
    QSet<QString> classes;const auto sources=standardVbaModuleSources(db,nullptr,&classes);if(!classes.contains(row.name.toCaseFolded()) || !sources.contains(row.name.toCaseFolded()) || sources.value(row.name.toCaseFolded())!=(row.imported ? row.encoded : row.source)) throw VbaRuntimeError(5,"Native class namespace is ambiguous");
}
std::shared_ptr<const NativeVbaClassDefinition> compiled(QSqlDatabase db,const ClassRevision& row) {
    proveNamespace(db,row);const auto aliases=projectVbaEnumAliases(db);auto result=compileNativeVbaClass(row.name.toStdString(),row.source.toStdString(),&aliases);
    for(const auto& type:nativeVbaClassDeclaredTypes(*result)) if(!nativeProjectClassTypeKnown(db,type)) throw VbaRuntimeError(3251,"Class declaration type is unavailable");
    return result;
}
}
std::string nativeProjectClassName(QSqlDatabase db,std::string_view name) {const auto row=revision(db,QString::fromUtf8(name.data(),name.size()));if(!row) throw VbaRuntimeError(3251,"Project class type is unavailable");proveNamespace(db,*row);return row->name.toStdString();}
bool nativeProjectClassTypeKnown(QSqlDatabase db,std::string_view name) {const auto row=revision(db,QString::fromUtf8(name.data(),name.size()));if(!row) return false;proveNamespace(db,*row);return true;}
void validateNativeProjectClassFields(QSqlDatabase db,const VbaModuleState& state) {for(const auto& [name,type]:state.types) if(isProjectClassType(type) && !nativeProjectClassTypeKnown(db,projectClassTypeName(type))) throw VbaRuntimeError(3251,"Module declaration type is unavailable");}
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
            const auto current=revision(db,expected.name);if(!current || current->name!=expected.name || current->encoded!=expected.encoded || current->imported!=expected.imported || current->archiveId!=expected.archiveId) throw VbaRuntimeError(3420,"Native class source revision changed");
        } catch(...) {*valid=false;throw;}
    };
    VbaClassContext context;context.check=check;context.host=[db,owner,check]() -> std::unique_ptr<VbaHost> {check();return std::make_unique<NativeStandardModuleHost>(owner,db,check);};
    return constructNativeVbaClass(definition,std::move(context));
}
}
