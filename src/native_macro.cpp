#include "native_macro.hpp"
#include "native_document_open.hpp"
namespace linaccess {
const QVector<NativeMacroSignature>& nativeMacroSignatures() {
    static const QVector<NativeMacroSignature> signatures{
        {"Maximize",{},0},{"Minimize",{},0},{"Restore",{},0},{"ApplyFilter",{"text","text"},0},{"ShowAllRecords",{},0},{"RepaintObject",{"long","text"},0},{"SetValue",{"text","text"},2},{"GoToControl",{"text"},1},{"Hourglass",{"boolean"},1},{"SetWarnings",{"boolean"},1},{"FindRecord",{"text","long","boolean","long","boolean","long","boolean"},1},{"GoToRecord",{"long","text","long","long"},0},{"MessageBox",{"text","boolean","text","text"},1},{"RunSQL",{"text","boolean"},1},{"OpenTable",{"text","long","long"},1},
        {"OpenQuery",{"text","long","long"},1},{"OpenForm",{"text","long","text","text","long","long","text"},1},
        {"OpenReport",{"text","long","text","text","long","text"},1},{"OpenModule",{"text","text"},0},{"RunCommand",{"long"},1},
        {"SelectObject",{"long","text","boolean"},2},{"Close",{"long","text","long"},0},{"DeleteObject",{"long","text"},2},{"Rename",{"text","long","text"},3},{"RunCode",{"text"},1},{"Echo",{"boolean","text"},1},{"Requery",{"text"},0},{"StopAllMacros",{},0},{"StopMacro",{},0},{"CancelEvent",{},0},{"RunMacro",{"text","long","text"},1},
        {"MoveSize",{"long","long","long","long"},0}};
    return signatures;
}
bool nativeMacroNameValid(const QString& name) {
    return !name.trimmed().isEmpty() && name.size()<=64 && name.isValidUtf16() && !name.contains(QChar(0)) && !name.contains('.') && !name.trimmed().startsWith('=') && !name.trimmed().startsWith('[');
}
bool nativeMacroCallNameValid(const QString& name) {
    const auto parts=name.split('.');if(parts.isEmpty() || parts.size()>2) return false;
    for(const auto& part:parts) if(!nativeMacroNameValid(part)) return false;return true;
}
std::optional<QString> nativeMacroRepeatExpression(const QString& value) {
    const auto expression=value.trimmed();
    if(expression.size()>4096 || !expression.isValidUtf16() || expression.contains(QChar(0))) throw VbaRuntimeError(5,"Invalid RunMacro repeat expression text");
    if(expression.isEmpty()) return {};
    if(!validVbaExpression(expression.toUtf8().toStdString())) throw VbaRuntimeError(3251,"RunMacro repeat expression syntax unsupported");
    return expression;
}

std::array<std::optional<int>,4> nativeMacroMoveSizeArguments(const std::vector<VbaValue>& arguments) {
    std::array<std::optional<int>,4> twips;
    if(arguments.size()>twips.size()) throw VbaRuntimeError(450,"MoveSize accepts at most four arguments");
    bool supplied=false;
    for(std::size_t index=0;index<arguments.size();++index) {
        const auto& argument=arguments[index];
        if(argument.kind==VbaValue::Kind::Missing) continue;
        if(argument.kind!=VbaValue::Kind::Number || !std::isfinite(argument.number)
            || std::trunc(argument.number)!=argument.number)
            throw VbaRuntimeError(13,"MoveSize requires whole twips");
        if(argument.number<std::numeric_limits<qint16>::min() || argument.number>std::numeric_limits<qint16>::max())
            throw VbaRuntimeError(6,"MoveSize exceeds the supported twips range");
        if(index>=2 && argument.number<=0) throw VbaRuntimeError(5,"MoveSize dimensions must be positive");
        twips[index]=static_cast<int>(argument.number);
        supplied=true;
    }
    if(!supplied) throw VbaRuntimeError(450,"MoveSize requires at least one argument");
    return twips;
}

QVector<NativeMacroAction> parseNativeMacro(const QJsonObject& definition,const QString& requestedGroup) {
    if(definition.size()!=2) throw VbaRuntimeError(3251,"Unknown native macro attributes are not supported");
    if(QJsonDocument(definition).toJson(QJsonDocument::Compact).size()>1024*1024) throw VbaRuntimeError(3035,"Native macro size limit exceeded");
    if(definition.value("version")==QJsonValue(2)) {
        if(!definition.value("groups").isArray()) throw VbaRuntimeError(3251,"Native macro groups require an ordered list");
        const auto groups=definition.value("groups").toArray();if(groups.isEmpty() || groups.size()>64) throw VbaRuntimeError(3251,"Native macro group count requires 1 to 64");
        QStringList names;QVector<NativeMacroAction> selected;int total=0;
        for(const auto& value:groups) {
            if(!value.isObject()) throw VbaRuntimeError(13,"Invalid native submacro");const auto group=value.toObject();const auto name=group.value("name").toString();
            if(group.size()!=2 || name!=name.trimmed() || !nativeMacroNameValid(name) || !group.value("actions").isArray()) throw VbaRuntimeError(3251,"Invalid native submacro definition");
            if(names.contains(name,Qt::CaseInsensitive)) throw VbaRuntimeError(3251,"Duplicate native submacro name");
            QVector<NativeMacroAction> actions;try {actions=parseNativeMacro(QJsonObject{{"version",1},{"actions",group.value("actions")}});}catch(const VbaRuntimeError& error) {throw VbaRuntimeError(error.number,"Teilmakro "+name.toStdString()+": "+error.what(),error.source,error.helpFile,error.helpContext);}total+=actions.size();if(total>256) throw VbaRuntimeError(3251,"Native macro groups exceed 256 total actions");
            if((requestedGroup.isEmpty() && names.isEmpty()) || name.compare(requestedGroup,Qt::CaseInsensitive)==0) selected=actions;
            names.append(name);
        }
        if(selected.isEmpty()) throw VbaRuntimeError(2485,"Native submacro not found");return selected;
    }
    if(!requestedGroup.isEmpty()) throw VbaRuntimeError(2485,"Native macro has no named submacros");
    if(definition.value("version")!=QJsonValue(1) || !definition.value("actions").isArray()) throw VbaRuntimeError(3251,"Unsupported native macro definition");
    const auto entries=definition.value("actions").toArray();
    if(entries.isEmpty() || entries.size()>256) throw VbaRuntimeError(3251,"Native macros require 1 to 256 actions");
    QVector<NativeMacroAction> actions;bool conditionalBlock=false;
    for(const auto& entry:entries) {
        if(!entry.isObject()) throw VbaRuntimeError(13,"Invalid native macro action");
        const auto object=entry.toObject();if(object.size()<2 || object.size()>3 || (object.size()==3 && !object.contains("condition"))) throw VbaRuntimeError(3251,"Unknown native macro action attributes are not supported");const auto name=object.value("action").toString();
        const NativeMacroSignature* signature=nullptr;
        for(const auto& candidate:nativeMacroSignatures()) if(candidate.name.compare(name,Qt::CaseInsensitive)==0) {signature=&candidate;break;}
        if(!signature) throw VbaRuntimeError(3251,"Native macro action is not implemented");
        if(!object.value("arguments").isArray()) throw VbaRuntimeError(13,"Native macro arguments require a list");
        const auto arguments=object.value("arguments").toArray();
        if(arguments.size()<signature->required || arguments.size()>signature->types.size()) throw VbaRuntimeError(450,"Native macro action argument count invalid");
        NativeMacroAction action{signature->name,{},std::nullopt};
        if(object.contains("condition")) {
            if(!object.value("condition").isString()) throw VbaRuntimeError(13,"Native macro condition requires an expression string");
            const auto condition=object.value("condition").toString().trimmed();
            if(condition.size()>4096 || !condition.isValidUtf16() || condition.contains(QChar(0))) throw VbaRuntimeError(5,"Invalid native macro condition text");
            if(!condition.isEmpty()) {
                if(condition=="...") {if(!conditionalBlock) throw VbaRuntimeError(3251,"Macro condition continuation requires a preceding conditional block");}
                else if(!validVbaExpression(condition.toUtf8().toStdString())) throw VbaRuntimeError(3251,"Native macro condition syntax unsupported");
                action.condition=condition;
            }
        }
        for(int index=0;index<arguments.size();++index) {
            const auto value=arguments[index];const auto type=signature->types[index];
            if(index>=signature->required && value.isNull()) {action.arguments.push_back(VbaValue::missing());continue;}
            if(type=="text" && value.isString()) {
                const auto text=value.toString();if(text.size()>262144 || !text.isValidUtf16() || text.contains(QChar(0))) throw VbaRuntimeError(5,"Invalid native macro text");
                if(index==0 && signature->required && text.trimmed().isEmpty()) throw VbaRuntimeError(5,"Native macro target cannot be empty");
                if(signature->name=="MessageBox") {
                    if(index==0) {if(text.size()>255) throw VbaRuntimeError(5,"MessageBox message argument exceeds 255 characters");if(text.startsWith('=' ) && !validVbaExpression(text.mid(1).toUtf8().toStdString())) throw VbaRuntimeError(3251,"MessageBox expression syntax unsupported");}
                    if(index==2 && !QStringList{"none","critical","warning?","warning!","information",""}.contains(text.toCaseFolded())) throw VbaRuntimeError(5,"MessageBox requires a supported named message type");
                }
                if(signature->name=="ApplyFilter" && index==1 && text.size()>256) throw VbaRuntimeError(5,"ApplyFilter macro condition exceeds 256 characters");
                if(signature->name=="SetValue") {if(index==0) (void)nativeMacroSetValueTarget(text);else (void)nativeMacroSetValueExpression(text);}
                if(signature->name=="RunCode") (void)nativeMacroFunctionName(text);
                if(signature->name=="RunMacro" && index==2) (void)nativeMacroRepeatExpression(text);
                action.arguments.push_back(VbaValue::string(text.toUtf8().toStdString()));
            } else if(type=="boolean" && value.isBool()) action.arguments.push_back(VbaValue::logical(value.toBool()));
            else if(type=="long" && value.isDouble() && std::isfinite(value.toDouble()) && std::trunc(value.toDouble())==value.toDouble()
                && value.toDouble()>=std::numeric_limits<qint32>::min() && value.toDouble()<=std::numeric_limits<qint32>::max()) action.arguments.push_back(VbaValue::longInteger(static_cast<qint32>(value.toDouble())));
            else throw VbaRuntimeError(13,"Native macro action argument type invalid");
        }
        if(action.name=="MoveSize") (void)nativeMacroMoveSizeArguments(action.arguments);
        if(action.name=="OpenModule") {
            const auto request=nativeModuleOpenArguments(action.arguments);
            if(request.name.size()>64 || request.procedure.size()>255) throw VbaRuntimeError(5,"Invalid OpenModule name");
        }
        conditionalBlock=action.condition.has_value();actions.push_back(std::move(action));
    }
    return actions;
}
QString nativeMacroCatalogName(QSqlDatabase database,const QString& requested) {
    // Imported artifact-only databases have no owned-object catalog yet.
    // Its absence means no native macro; existing unreadable catalogs remain errors.
    QSqlQuery catalog(database);if(!catalog.exec("SELECT 1 FROM main.sqlite_schema WHERE name='_la_objects' COLLATE NOCASE LIMIT 1"))throw VbaRuntimeError(3146,"Native macro catalog unavailable");
    if(!catalog.next()){if(catalog.lastError().isValid())throw VbaRuntimeError(3146,"Native macro catalog read failed");return {};}
    catalog.finish();
    QSqlQuery names(database);names.setForwardOnly(true);if(!names.exec("SELECT name FROM main._la_objects WHERE kind='macro'")) throw VbaRuntimeError(3146,"Native macro catalog unavailable");
    QElapsedTimer timer;timer.start();int count=0;QString found;
    while(names.next()) {
        if(++count>1000000) throw VbaRuntimeError(3251,"Native macro catalog limit exceeded");
        if(timer.hasExpired(5000)) throw VbaRuntimeError(3035,"Native macro catalog deadline exceeded");
        const auto candidate=names.value(0).toString();
        if(candidate.compare(requested,Qt::CaseInsensitive)==0) {if(!found.isEmpty()) throw VbaRuntimeError(3251,"Native macro name is ambiguous");found=candidate;}
    }
    if(names.lastError().isValid()) throw VbaRuntimeError(3146,"Native macro catalog read failed");return found;
}
QVector<NativeMacroAction> loadNativeMacro(QSqlDatabase database,const QString& name) {
    if(!nativeMacroCallNameValid(name)) throw VbaRuntimeError(3251,"Invalid native macro call name");
    const auto parts=name.split('.');const auto canonical=nativeMacroCatalogName(database,parts[0]);if(canonical.isEmpty()) throw VbaRuntimeError(2485,"Native macro not found");
    QSqlQuery query(database);query.prepare("SELECT definition FROM main._la_objects WHERE kind='macro' AND name=? COLLATE BINARY LIMIT 2");query.addBindValue(canonical);
    if(!query.exec()) throw VbaRuntimeError(3146,"Native macro catalog unavailable");
    if(!query.next()) throw VbaRuntimeError(2485,"Native macro not found");
    const auto bytes=query.value(0).toByteArray();
    if(query.next()) throw VbaRuntimeError(3251,"Native macro name is ambiguous");
    if(query.lastError().isValid()) throw VbaRuntimeError(3146,"Native macro catalog read failed");
    if(bytes.size()>1024*1024) throw VbaRuntimeError(3035,"Native macro size limit exceeded");
    QJsonParseError error;const auto document=QJsonDocument::fromJson(bytes,&error);
    if(error.error!=QJsonParseError::NoError || !document.isObject()) throw VbaRuntimeError(13,"Invalid native macro document");
    return parseNativeMacro(document.object(),parts.size()==2 ? parts[1] : QString{});
}
}
