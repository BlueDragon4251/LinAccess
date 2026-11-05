#pragma once
#include "vba_analyzer.hpp"
#include "vba_expression.hpp"
#include "vba_module_context.hpp"
#include "vba_source_codec.hpp"
#include "vba_qt_value.hpp"
#include "vba_currency.hpp"
#include <QMap>
#include <QPointer>
#include <QRegularExpression>
#include <QSet>
#include <QScopeGuard>
#include <QSqlDatabase>
#include <QSqlDriver>
#include <QSqlError>
#include <QSqlQuery>
#include <QThread>
#include <QVariant>
#include <charconv>
#include <cmath>
#include <limits>
#include <sqlite3.h>

namespace linaccess::vba_sql {
struct Lifetime { bool alive{true}; };
struct Procedure {
    std::shared_ptr<const std::string> module;
    bool function{},ambiguous{};
    std::string resultType;
};
class Registry final : public QObject {
public:
    explicit Registry(QSqlDriver* owner):QObject(owner),driver(owner) {}
    QPointer<QSqlDriver> driver;
    QString connectionName,databaseName;
    std::shared_ptr<Lifetime> lifetime;
    QMap<QString,Procedure> procedures;
    QMap<QString,QByteArray> sources;
    QSet<QString> registered;
    sqlite3_int64 changes{-1}; int dataVersion{-1},schemaVersion{-1},tempSchemaVersion{-1};
    quint64 statementEpoch{},observedEpoch{std::numeric_limits<quint64>::max()}; int metadataDepth{};
    QSqlDatabase database(sqlite3* native) const {
        if(!driver) throw VbaRuntimeError(3420,"VBA SQL connection is closed");
        if(driver->thread()!=QThread::currentThread()) throw VbaRuntimeError(3251,"VBA SQL requires the connection thread");
        auto db=QSqlDatabase::database(connectionName,false);
        if(!db.isOpen() || db.driver()!=driver || db.databaseName()!=databaseName) throw VbaRuntimeError(3420,"VBA SQL connection identity changed");
        const auto handle=driver->handle();
        if(!handle.isValid() || QByteArray(handle.typeName())!="sqlite3*" || *static_cast<sqlite3* const*>(handle.constData())!=native)
            throw VbaRuntimeError(3420,"VBA SQL native connection changed");
        return db;
    }
};
class MetadataScope final {
public:
    explicit MetadataScope(Registry& registry):registry_(registry) { ++registry_.metadataDepth; }
    ~MetadataScope() { --registry_.metadataDepth; }
private: Registry& registry_;
};
struct TraceBinding { QPointer<Registry> registry; std::shared_ptr<Lifetime> lifetime; };
inline int pragma(QSqlDatabase db,const char* command) {
    QSqlQuery query(db);
    if(!query.exec(QString::fromLatin1(command)) || !query.next()) throw VbaRuntimeError(3251,"VBA SQL metadata revision unavailable");
    return query.value(0).toInt();
}
inline void refresh(QSqlDatabase db,sqlite3* native,Registry& registry) {
    MetadataScope metadata(registry); const auto epoch=registry.statementEpoch;
    const auto changes=sqlite3_total_changes64(native);
    const auto data=pragma(db,"PRAGMA main.data_version"),schema=pragma(db,"PRAGMA main.schema_version"),temporary=pragma(db,"PRAGMA temp.schema_version");
    if(registry.changes==changes && registry.dataVersion==data && registry.schemaVersion==schema && registry.tempSchemaVersion==temporary && registry.observedEpoch==epoch) return;
    QMap<QString,QByteArray> sources;
    QSqlQuery table(db);
    if(!table.exec("SELECT type FROM main.sqlite_schema WHERE name='_la_access_artifacts' COLLATE NOCASE")) throw VbaRuntimeError(3251,"VBA SQL metadata unavailable");
    if(table.next()) {
        if(table.value(0).toString()!="table") throw VbaRuntimeError(3251,"VBA SQL requires a physical source archive");
        table.finish();
        QSqlQuery any(db);
        if(!any.exec("SELECT 1 FROM main._la_access_artifacts WHERE kind='vba' LIMIT 1")) throw VbaRuntimeError(3251,"VBA SQL archive shape unsupported");
        const bool hasCode=any.next(); any.finish();
        if(hasCode) {
            QSqlQuery shadow(db);
            if(!shadow.exec("SELECT 1 FROM temp.sqlite_schema WHERE name='_la_access_artifacts' COLLATE NOCASE AND type IN ('table','view')")) throw VbaRuntimeError(3251,"VBA SQL shadow check failed");
            if(shadow.next()) throw VbaRuntimeError(3251,"VBA SQL archive shadow is unsupported");
            shadow.finish();
            QSqlQuery rows(db);
            if(!rows.exec("SELECT name,source_bytes FROM main._la_access_artifacts WHERE kind='vba' ORDER BY id DESC")) throw VbaRuntimeError(3251,"VBA SQL archive ordering unavailable");
            qsizetype bytes=0;
            while(rows.next()) {
                const auto key=rows.value(0).toString().toCaseFolded();
                if(sources.contains(key)) continue;
                auto source=rows.value(1).toByteArray(); bytes+=source.size();
                if(sources.size()>=256 || bytes>16*1024*1024) throw VbaRuntimeError(3251,"VBA SQL module inventory limit exceeded");
                sources.insert(key,std::move(source));
            }
            if(rows.lastError().isValid()) throw VbaRuntimeError(3146,"VBA SQL archive fetch failed");
        }
    }
    if(sources!=registry.sources) {
        QMap<QString,Procedure> procedures;
        static const QRegularExpression privateDeclaration("^\\s*(?:(?:Public|Static)\\s+)*(?:Private|Friend)\\b",QRegularExpression::CaseInsensitiveOption);
        static const QRegularExpression sqlName("^[A-Za-z_][A-Za-z0-9_]*$");
        qsizetype bytes=0;
        for(const auto& raw:sources) {
            const auto text=decodeVbaSource(raw);
            auto module=std::make_shared<const std::string>(text.toUtf8().toStdString()); bytes+=static_cast<qsizetype>(module->size());
            if(bytes>16*1024*1024) throw VbaRuntimeError(3251,"VBA SQL decoded module limit exceeded");
            if(isVbaClassModule(*module)) continue;
            const auto analysis=analyzeVba(*module); const auto lines=text.split('\n');
            for(const auto& item:analysis.procedures) {
                const auto name=QString::fromStdString(item.name),key=name.toCaseFolded();
                if(!sqlName.match(name).hasMatch() || name.size()>128 || key.startsWith("linaccess_") || key.startsWith("sqlite_")) continue;
                if(item.line<1 || item.line>lines.size() || privateDeclaration.match(lines[item.line-1]).hasMatch()) continue;
                const bool function=QString::fromStdString(item.kind).compare("function",Qt::CaseInsensitive)==0;
                if(auto found=procedures.find(key);found!=procedures.end()) { found->ambiguous=true; found->function|=function; }
                else {
                    const auto signature=function ? nativeVbaProcedureSignature(*module,item.name) : std::optional<VbaArgumentSignature>{};
                    procedures.insert(key,{module,function,false,signature && signature->function ? signature->resultType : std::string{}});
                }
                if(procedures.size()>4096) throw VbaRuntimeError(3251,"VBA SQL procedure inventory limit exceeded");
            }
        }
        registry.procedures=std::move(procedures); registry.sources=std::move(sources);
    }
    registry.changes=changes; registry.dataVersion=data; registry.schemaVersion=schema; registry.tempSchemaVersion=temporary; registry.observedEpoch=epoch;
}
class Host final : public VbaHost {
public:
    Host(QSqlDatabase db,Registry& registry):db_(std::move(db)),registry_(registry) {}
    std::optional<std::string> field(std::string_view) override { unavailable(); }
    void setField(std::string_view,std::optional<std::string>) override { unavailable(); }
    void message(std::string_view) override { unavailable(); }
    void saveRecord() override { unavailable(); }
    void quitApplication() override { unavailable(); }
    void maximizeForm() override { unavailable(); }
    void gotoLastRecord() override { unavailable(); }
    void gotoNextRecord() override { unavailable(); }
    void gotoNewRecord() override { unavailable(); }
    void requery() override { unavailable(); }
    std::shared_ptr<VbaObject> currentDatabase() override { unavailable(); }
    std::shared_ptr<VbaObject> databaseEngine() override { unavailable(); }
    std::shared_ptr<VbaObject> applicationObject() override { unavailable(); }
    std::shared_ptr<VbaObject> currentForm() override { unavailable(); }
    std::shared_ptr<VbaObject> reportsCollection() override { unavailable(); }
    std::shared_ptr<VbaModuleState> moduleState(std::string_view source) override {
        MetadataScope metadata(registry_);
        auto state=standardVbaModuleState(db_,source);
        for(const auto& [name,value]:state->values) requireScalar(value);
        for(const auto& [procedure,values]:state->procedureStatics)
            for(const auto& [name,value]:values) requireScalar(value);
        return state;
    }
    std::vector<std::shared_ptr<VbaModuleState>> globalModuleStates() override {
        MetadataScope metadata(registry_);
        auto states=publicVbaModuleStates(db_);
        for(const auto& state:states) for(const auto& name:state->publicVariables)
            requireScalar(state->values.at(name));
        return states;
    }
    std::optional<std::string> moduleForProcedure(std::string_view name) override {
        const auto found=registry_.procedures.constFind(QString::fromUtf8(name.data(),name.size()).toCaseFolded());
        if(found==registry_.procedures.cend()) return {};
        if(found->ambiguous) throw VbaRuntimeError(5,"Ambiguous public VBA SQL procedure");
        return *found->module;
    }
private:
    static void requireScalar(const VbaValue& value) {
        if((value.kind==VbaValue::Kind::Object && value.object) || value.kind==VbaValue::Kind::Array)
            throw VbaRuntimeError(3251,"VBA SQL object/array module state requires an unavailable execution context");
    }
    [[noreturn]] static void unavailable() { throw VbaRuntimeError(3251,"VBA SQL has no UI or DAO execution context"); }
    QSqlDatabase db_; Registry& registry_;
};
struct Binding { QPointer<Registry> registry; std::shared_ptr<Lifetime> lifetime; QString name; };
inline void failure(sqlite3_context* context,int number) noexcept {
    char text[80]="LinAccess SQL function error ";
    constexpr auto prefix=sizeof("LinAccess SQL function error ")-1;
    const auto converted=std::to_chars(text+prefix,text+sizeof(text),number);
    sqlite3_result_error(context,text,static_cast<int>(converted.ptr-text));
}
inline void invoke(sqlite3_context* context,int count,sqlite3_value** values) noexcept {
    try {
        const auto* binding=static_cast<Binding*>(sqlite3_user_data(context));
        if(!binding || !binding->registry || !binding->lifetime->alive || binding->registry->lifetime!=binding->lifetime)
            throw VbaRuntimeError(3420,"VBA SQL function connection expired");
        auto& registry=*binding->registry; auto* native=sqlite3_context_db_handle(context);
        auto db=registry.database(native); refresh(db,native,registry);
        const auto definition=registry.procedures.value(binding->name);
        if(!definition.module || !definition.function || definition.ambiguous) throw VbaRuntimeError(3251,"VBA SQL public function unavailable or ambiguous");
        if(count>60) throw VbaRuntimeError(450,"VBA SQL argument limit exceeded");
        std::vector<VbaValue> arguments; arguments.reserve(count);
        for(int index=0;index<count;++index) {
            switch(sqlite3_value_type(values[index])) {
            case SQLITE_NULL: arguments.emplace_back(); break;
            case SQLITE_INTEGER: {
                const auto value=sqlite3_value_int64(values[index]);
                if(value < -9007199254740992LL || value > 9007199254740992LL) throw VbaRuntimeError(3251,"VBA SQL integer transport exceeds exact Double range");
                arguments.push_back(VbaValue::numeric(static_cast<double>(value))); break;
            }
            case SQLITE_FLOAT: {
                const auto value=sqlite3_value_double(values[index]);
                if(!std::isfinite(value)) throw VbaRuntimeError(6,"VBA SQL nonfinite argument unsupported");
                arguments.push_back(VbaValue::numeric(value)); break;
            }
            case SQLITE_TEXT: {
                const auto* text=sqlite3_value_text(values[index]); const auto size=sqlite3_value_bytes(values[index]);
                if(!text) throw VbaRuntimeError(7,"VBA SQL text unavailable");
                std::string value=text ? std::string(reinterpret_cast<const char*>(text),size) : std::string{};
                (void)vbaStringCodeUnits(value); arguments.push_back(VbaValue::string(std::move(value))); break;
            }
            default: throw VbaRuntimeError(3251,"VBA SQL binary argument transport unsupported");
            }
        }
        Host host(db,registry);
        const auto result=executeVbaProcedureWithArguments(*definition.module,binding->name.toStdString(),host,arguments);
        if(result.status!=VbaExecutionStatus::Completed || !result.typedReturnValue)
            throw VbaRuntimeError(result.errorNumber ? result.errorNumber : 3251,"VBA SQL execution failed");
        const auto& value=*result.typedReturnValue;
        if(value.kind==VbaValue::Kind::Null) sqlite3_result_null(context);
        else if(value.kind==VbaValue::Kind::Boolean) sqlite3_result_int64(context,value.boolean ? -1 : 0);
        else if(value.kind==VbaValue::Kind::Currency && definition.resultType=="currency")
            sqlite3_result_double(context,currencyToRealStorage(value));
        else if(value.kind==VbaValue::Kind::Date && definition.resultType=="date") {
            const auto text=sqlFieldValue(value,"DATETIME").toString().toUtf8();
            sqlite3_result_text(context,text.constData(),text.size(),SQLITE_TRANSIENT);
        }
        else if(value.kind==VbaValue::Kind::Number) {
            if(!std::isfinite(value.number)) throw VbaRuntimeError(6,"VBA SQL numeric overflow");
            if(value.numericType==VbaValue::NumericType::Double || value.numericType==VbaValue::NumericType::Single) sqlite3_result_double(context,value.number);
            else sqlite3_result_int64(context,static_cast<sqlite3_int64>(value.number));
        } else if(value.kind==VbaValue::Kind::Text) {
            (void)vbaStringCodeUnits(value.text);
            if(value.text.size()>static_cast<std::size_t>(std::numeric_limits<int>::max())) throw VbaRuntimeError(6,"VBA SQL result length overflow");
            sqlite3_result_text(context,value.text.data(),static_cast<int>(value.text.size()),SQLITE_TRANSIENT);
        } else throw VbaRuntimeError(3251,"VBA SQL result type requires unavailable transport metadata");
    } catch(const VbaRuntimeError& error) { failure(context,error.number); }
      catch(...) { sqlite3_result_error(context,"Native VBA SQL execution failed",-1); }
}
inline void resultGuard(sqlite3_context* context,int count,sqlite3_value** values) noexcept {
    try {
        auto* binding=static_cast<TraceBinding*>(sqlite3_user_data(context));
        if(!binding || !binding->registry || !binding->lifetime->alive || binding->registry->lifetime!=binding->lifetime)
            throw VbaRuntimeError(3420,"VBA SQL result declaration expired");
        if(count!=3 || sqlite3_value_type(values[0])!=SQLITE_TEXT || sqlite3_value_type(values[1])!=SQLITE_TEXT)
            throw VbaRuntimeError(450,"VBA SQL result guard requires declaration metadata");
        auto& registry=*binding->registry; auto* native=sqlite3_context_db_handle(context);
        auto db=registry.database(native); refresh(db,native,registry);
        const auto name=QString::fromUtf8(reinterpret_cast<const char*>(sqlite3_value_text(values[0])),sqlite3_value_bytes(values[0])).toCaseFolded();
        const auto expected=QString::fromUtf8(reinterpret_cast<const char*>(sqlite3_value_text(values[1])),sqlite3_value_bytes(values[1])).toStdString();
        const auto procedure=registry.procedures.value(name);
        if(!procedure.module || !procedure.function || procedure.ambiguous || procedure.resultType!=expected || !registry.registered.contains(name))
            throw VbaRuntimeError(3251,"VBA SQL function result declaration changed; reprepare required");
        sqlite3_result_value(context,values[2]);
    } catch(const VbaRuntimeError& error) { failure(context,error.number); }
      catch(...) { sqlite3_result_error(context,"Native VBA SQL result declaration failed",-1); }
}
inline QMap<QString,QString> declaredTypes(QSqlDatabase db) {
    for(auto* child:db.driver()->children()) if(auto* registry=dynamic_cast<Registry*>(child)) {
        QMap<QString,QString> types;
        for(auto it=registry->procedures.cbegin();it!=registry->procedures.cend();++it)
            if(it->function && !it->ambiguous && !it->resultType.empty() && registry->registered.contains(it.key()))
                types.insert(it.key(),QString::fromStdString(it->resultType));
        return types;
    }
    return {};
}
inline void install(QSqlDatabase db,sqlite3* native) {
    Registry* registry=nullptr;
    for(auto* child:db.driver()->children()) if(auto* found=dynamic_cast<Registry*>(child)) { registry=found; break; }
    if(!registry) registry=new Registry(db.driver());
    QSqlQuery identity(db); const bool present=identity.prepare("SELECT linaccess_vba_registry_epoch()"); identity.clear();
    if(!present || !registry->lifetime || !registry->lifetime->alive) {
        if(present && !registry->lifetime) throw VbaRuntimeError(3251,"VBA SQL registry identity collision");
        if(registry->lifetime) resetVbaProjectState(db);
        registry->lifetime=std::make_shared<Lifetime>();
        auto* token=new TraceBinding{registry,registry->lifetime};
        if(sqlite3_create_function_v2(native,"linaccess_vba_registry_epoch",0,SQLITE_UTF8,token,
            [](sqlite3_context* context,int,sqlite3_value**) { sqlite3_result_int(context,1); },nullptr,nullptr,
            [](void* pointer) { auto* token=static_cast<TraceBinding*>(pointer); token->lifetime->alive=false; delete token; })!=SQLITE_OK)
            throw VbaRuntimeError(3251,"VBA SQL registry initialization failed");
        if(sqlite3_trace_v2(native,SQLITE_TRACE_STMT,
            [](unsigned event,void* pointer,void*,void*) -> int {
                const auto* binding=static_cast<TraceBinding*>(pointer);
                if(event==SQLITE_TRACE_STMT && binding->lifetime->alive && binding->registry
                    && binding->registry->lifetime==binding->lifetime && binding->registry->metadataDepth==0)
                    ++binding->registry->statementEpoch;
                return 0;
            },token)!=SQLITE_OK) {
            registry->lifetime->alive=false; throw VbaRuntimeError(3251,"VBA SQL statement generation unavailable");
        }
        auto* guard=new TraceBinding{registry,registry->lifetime};
        if(sqlite3_create_function_v2(native,"linaccess_vba_result",3,SQLITE_UTF8,guard,&resultGuard,nullptr,nullptr,
            [](void* pointer) { delete static_cast<TraceBinding*>(pointer); })!=SQLITE_OK)
            throw VbaRuntimeError(3251,"VBA SQL result guard registration failed");
        registry->statementEpoch=0; registry->observedEpoch=std::numeric_limits<quint64>::max();
        registry->connectionName=db.connectionName(); registry->databaseName=db.databaseName();
        registry->registered.clear(); registry->procedures.clear(); registry->sources.clear();
        registry->changes=-1; registry->dataVersion=-1; registry->schemaVersion=-1; registry->tempSchemaVersion=-1;
    }
    MetadataScope metadata(*registry);
    (void)registry->database(native); refresh(db,native,*registry);
    bool pending=false;
    for(auto it=registry->procedures.cbegin();it!=registry->procedures.cend();++it)
        if(it->function && !registry->registered.contains(it.key())) { pending=true; break; }
    if(!pending) return;
    QSet<QString> occupied;
    QSqlQuery functions(db);
    if(!functions.exec("PRAGMA function_list")) throw VbaRuntimeError(3251,"SQLite function inventory unavailable");
    const QString adapterPrefix="linaccess_jet_";
    while(functions.next()) {
        const auto name=functions.value(0).toString().toCaseFolded();
        occupied.insert(name);
        // Jet source names are rewritten to these native adapters. Their VBA
        // declaration must not become false result metadata for the builtin.
        if(name.startsWith(adapterPrefix)) occupied.insert(name.mid(adapterPrefix.size()));
    }
    functions.finish();
    for(auto it=registry->procedures.cbegin();it!=registry->procedures.cend();++it) {
        if(!it->function || registry->registered.contains(it.key()) || occupied.contains(it.key())) continue;
        const auto name=it.key().toUtf8(); auto* binding=new Binding{registry,registry->lifetime,it.key()};
        if(sqlite3_create_function_v2(native,name.constData(),-1,SQLITE_UTF8,binding,&invoke,nullptr,nullptr,
            [](void* pointer) { delete static_cast<Binding*>(pointer); })!=SQLITE_OK)
            throw VbaRuntimeError(3251,"VBA SQL function registration failed");
        registry->registered.insert(it.key());
    }
}
} // namespace linaccess::vba_sql
