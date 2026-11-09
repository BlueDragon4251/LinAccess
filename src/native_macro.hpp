#pragma once
#include "vba_runtime.hpp"
#include "vba_expression.hpp"
#include <QJsonArray>
#include <QElapsedTimer>
#include <QJsonDocument>
#include <QJsonObject>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlError>
#include <QStringList>
#include <QVector>
#include <cmath>
#include <limits>
namespace linaccess {
struct NativeMacroSignature {QString name;QStringList types;int required;};
inline const QVector<NativeMacroSignature>& nativeMacroSignatures() {
    static const QVector<NativeMacroSignature> signatures{
        {"RunSQL",{"text","boolean"},1},{"OpenTable",{"text","long","long"},1},
        {"OpenQuery",{"text","long","long"},1},{"OpenForm",{"text","long"},1},
        {"OpenReport",{"text","long"},1},{"RunCommand",{"long"},1},
        {"Requery",{"text"},0},{"StopMacro",{},0},{"CancelEvent",{},0},{"RunMacro",{"text","long"},1}};
    return signatures;
}
inline bool nativeMacroNameValid(const QString& name) {
    return !name.trimmed().isEmpty() && name.size()<=64 && name.isValidUtf16() && !name.contains(QChar(0)) && !name.contains('.') && !name.trimmed().startsWith('=') && !name.trimmed().startsWith('[');
}
struct NativeMacroAction {QString name;std::vector<VbaValue> arguments;std::optional<QString> condition;};
inline QVector<NativeMacroAction> parseNativeMacro(const QJsonObject& definition) {
    if(definition.size()!=2) throw VbaRuntimeError(3251,"Unknown native macro attributes are not supported");
    if(definition.value("version")!=QJsonValue(1) || !definition.value("actions").isArray()) throw VbaRuntimeError(3251,"Unsupported native macro definition");
    if(QJsonDocument(definition).toJson(QJsonDocument::Compact).size()>1024*1024) throw VbaRuntimeError(3035,"Native macro size limit exceeded");
    const auto entries=definition.value("actions").toArray();
    if(entries.isEmpty() || entries.size()>256) throw VbaRuntimeError(3251,"Native macros require 1 to 256 actions");
    QVector<NativeMacroAction> actions;
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
            if(!condition.isEmpty()) {if(!validVbaExpression(condition.toUtf8().toStdString())) throw VbaRuntimeError(3251,"Native macro condition syntax unsupported");action.condition=condition;}
        }
        for(int index=0;index<arguments.size();++index) {
            const auto value=arguments[index];const auto type=signature->types[index];
            if(type=="text" && value.isString()) {
                const auto text=value.toString();if(text.size()>262144 || !text.isValidUtf16() || text.contains(QChar(0))) throw VbaRuntimeError(5,"Invalid native macro text");
                if(index==0 && signature->required && text.trimmed().isEmpty()) throw VbaRuntimeError(5,"Native macro target cannot be empty");
                action.arguments.push_back(VbaValue::string(text.toUtf8().toStdString()));
            } else if(type=="boolean" && value.isBool()) action.arguments.push_back(VbaValue::logical(value.toBool()));
            else if(type=="long" && value.isDouble() && std::isfinite(value.toDouble()) && std::trunc(value.toDouble())==value.toDouble()
                && value.toDouble()>=std::numeric_limits<qint32>::min() && value.toDouble()<=std::numeric_limits<qint32>::max()) action.arguments.push_back(VbaValue::longInteger(static_cast<qint32>(value.toDouble())));
            else throw VbaRuntimeError(13,"Native macro action argument type invalid");
        }
        actions.push_back(std::move(action));
    }
    return actions;
}
inline QString nativeMacroCatalogName(QSqlDatabase database,const QString& requested) {
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
inline QVector<NativeMacroAction> loadNativeMacro(QSqlDatabase database,const QString& name) {
    if(!nativeMacroNameValid(name)) throw VbaRuntimeError(3251,"Macro groups or invalid native macro names are not supported");
    const auto canonical=nativeMacroCatalogName(database,name);if(canonical.isEmpty()) throw VbaRuntimeError(2485,"Native macro not found");
    QSqlQuery query(database);query.prepare("SELECT definition FROM main._la_objects WHERE kind='macro' AND name=? COLLATE BINARY LIMIT 2");query.addBindValue(canonical);
    if(!query.exec()) throw VbaRuntimeError(3146,"Native macro catalog unavailable");
    if(!query.next()) throw VbaRuntimeError(2485,"Native macro not found");
    const auto bytes=query.value(0).toByteArray();
    if(query.next()) throw VbaRuntimeError(3251,"Native macro name is ambiguous");
    if(query.lastError().isValid()) throw VbaRuntimeError(3146,"Native macro catalog read failed");
    if(bytes.size()>1024*1024) throw VbaRuntimeError(3035,"Native macro size limit exceeded");
    QJsonParseError error;const auto document=QJsonDocument::fromJson(bytes,&error);
    if(error.error!=QJsonParseError::NoError || !document.isObject()) throw VbaRuntimeError(13,"Invalid native macro document");
    return parseNativeMacro(document.object());
}
}
