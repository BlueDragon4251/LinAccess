#pragma once
#include "native_macro.hpp"
#include <QMap>
namespace linaccess {
inline const QStringList& nativeFormMacroEvents() {
    static const QStringList events{"OnOpen","OnLoad","OnCurrent","OnActivate","OnDeactivate","OnUnload","OnClose","BeforeInsert","AfterInsert","BeforeUpdate","AfterUpdate","OnDirty","OnTimer"};return events;
}
inline const QStringList& nativeFieldMacroEvents() {
    static const QStringList events{"OnClick","OnDblClick","OnChange","BeforeUpdate","AfterUpdate","OnMouseDown","OnMouseUp"};return events;
}
inline QMap<QString,QString> nativeEventOverrides(const QJsonValue& value,const QStringList& allowed) {
    if(value.isUndefined()) return {};
    if(!value.isObject()) throw VbaRuntimeError(5,"Native event bindings require an object");
    const auto source=value.toObject();if(source.size()>allowed.size()) throw VbaRuntimeError(5,"Too many native event bindings");
    QMap<QString,QString> result;
    for(auto it=source.begin();it!=source.end();++it) {
        QString name;for(const auto& candidate:allowed) if(candidate.compare(it.key(),Qt::CaseInsensitive)==0) name=candidate;
        if(name.isEmpty() || result.contains(name)) throw VbaRuntimeError(3251,"Unknown or duplicate native event binding");
        if(!it.value().isString()) throw VbaRuntimeError(13,"Native event binding requires text");
        auto binding=it.value().toString().trimmed();
        if(binding.size()>4096 || !binding.isValidUtf16() || binding.contains(QChar(0))) throw VbaRuntimeError(5,"Invalid native event binding text");
        if(binding.compare("[Event Procedure]",Qt::CaseInsensitive)==0) binding="[Event Procedure]";
        else if(!binding.isEmpty() && !binding.startsWith('=') && !nativeMacroCallNameValid(binding)) throw VbaRuntimeError(3251,"Unsupported native event macro name");
        if(binding.startsWith('=') && binding.mid(1).trimmed().isEmpty()) throw VbaRuntimeError(5,"Native event expression cannot be empty");
        result.insert(name,binding);
    }
    return result;
}
inline QMap<QString,QMap<QString,QString>> nativeFieldEventOverrides(const QJsonValue& value,const QStringList& fields) {
    if(value.isUndefined()) return {};
    if(!value.isObject() || value.toObject().size()>255) throw VbaRuntimeError(5,"Invalid native field event catalog");
    QMap<QString,QMap<QString,QString>> result;const auto entries=value.toObject();
    for(auto it=entries.begin();it!=entries.end();++it) {
        QString field;for(const auto& candidate:fields) if(candidate.compare(it.key(),Qt::CaseInsensitive)==0) field=candidate.toCaseFolded();
        if(field.isEmpty() || result.contains(field)) throw VbaRuntimeError(3251,"Unknown or duplicate native event field");
        result.insert(field,nativeEventOverrides(it.value(),nativeFieldMacroEvents()));
    }
    return result;
}
inline int nativeFormTimerInterval(const QJsonValue& value) {
    if(value.isUndefined()) return 0;
    const auto number=value.toDouble(-1);
    if(!value.isDouble() || !std::isfinite(number) || std::trunc(number)!=number || number<0 || number>std::numeric_limits<qint32>::max()) throw VbaRuntimeError(5,"Invalid native form TimerInterval");
    return static_cast<int>(number);
}
inline QJsonObject nativeEventObject(const QMap<QString,QString>& bindings) {
    QJsonObject object;for(auto it=bindings.begin();it!=bindings.end();++it) object.insert(it.key(),it.value());return object;
}
}
