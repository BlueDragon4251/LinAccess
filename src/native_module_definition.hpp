#pragma once
#include "vba_runtime.hpp"
#include "vba_module_context.hpp"
#include <QJsonObject>
#include <QJsonDocument>
#include <QString>
namespace linaccess {
inline bool nativeModuleNameValid(const QString& name) {
    return !name.trimmed().isEmpty() && name==name.trimmed() && name.size()<=64 && name.isValidUtf16() && !name.contains(QChar(0)) && !name.contains('.') && !name.contains('!') && !name.contains('[') && !name.contains(']');
}
inline QByteArray nativeModuleSource(const QJsonObject& definition) {
    if(definition.size()!=2 || definition.value("version")!=QJsonValue(1) || !definition.value("vba").isString()) throw VbaRuntimeError(3251,"Unsupported native standard module definition");
    const auto text=definition.value("vba").toString();
    if(text.size()>1024*1024 || !text.isValidUtf16() || text.contains(QChar(0))) throw VbaRuntimeError(5,"Invalid native standard module text");
    const auto bytes=text.toUtf8();if(bytes.size()>1024*1024) throw VbaRuntimeError(3035,"Native standard module source limit exceeded");
    if(isVbaClassModule(std::string_view(bytes.constData(),static_cast<std::size_t>(bytes.size())))) throw VbaRuntimeError(3251,"Native standard module editor does not support class modules");
    return bytes;
}
}
