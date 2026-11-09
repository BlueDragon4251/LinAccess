#pragma once
#include "vba_runtime.hpp"
#include "vba_module_context.hpp"
#include <QJsonObject>
#include <QRegularExpression>
namespace linaccess {
inline bool nativeClassNameValid(const QString& name) {static const QRegularExpression valid("^[A-Za-z_][A-Za-z0-9_]{0,63}$");return valid.match(name).hasMatch();}
inline QByteArray nativeClassModuleSource(const QJsonObject& definition,const QString& name) {
    if(!nativeClassNameValid(name) || definition.size()!=2 || definition.value("version")!=QJsonValue(1) || !definition.value("vba").isString()) throw VbaRuntimeError(3251,"Invalid native class definition/name");
    const auto text=definition.value("vba").toString();if(!text.isValidUtf16() || text.contains(QChar(0))) throw VbaRuntimeError(13,"Native class source is not representable");
    const auto bytes=text.toUtf8();if(bytes.size()>1024*1024) throw VbaRuntimeError(3035,"Native class source exceeds one MiB");
    static const QRegularExpression attribute("^\\s*Attribute\\s+VB_Name\\s*=\\s*\"([^\"]+)\"\\s*$",QRegularExpression::MultilineOption|QRegularExpression::CaseInsensitiveOption);
    auto matches=attribute.globalMatch(text);if(!matches.hasNext()) throw VbaRuntimeError(3251,"Native class requires its VB_Name attribute");const auto declared=matches.next().captured(1);
    if(matches.hasNext() || declared.compare(name,Qt::CaseInsensitive)!=0 || !isVbaClassModule(std::string_view(bytes.constData(),bytes.size()))) throw VbaRuntimeError(3251,"Native class source header does not match its catalog name");
    return bytes;
}
}
