#pragma once
#include "native_macro.hpp"
#include <array>
namespace linaccess {
struct NativeMacroRowDraft {QString action{"RunSQL"};std::array<QString,4> cells;QStringList extraArguments;};
struct NativeMacroGroupDraft {QString name;QVector<NativeMacroRowDraft> rows;};
inline QVector<NativeMacroRowDraft> nativeMacroDraftRows(const QJsonArray& actions) {
    QVector<NativeMacroRowDraft> result;
    for(const auto& value:actions) {const auto entry=value.toObject();NativeMacroRowDraft row;row.action=entry.value("action").toString("RunSQL");const auto args=entry.value("arguments").toArray();for(int index=0;index<3;++index) {const auto arg=index<args.size() ? args[index] : QJsonValue{};row.cells[index]=arg.isString() ? arg.toString() : arg.isBool() ? (arg.toBool() ? QString("Ja") : QString("Nein")) : arg.isDouble() ? QString::number(arg.toDouble(),'g',17) : QString{};}for(int index=3;index<args.size();++index) {const auto arg=args[index];row.extraArguments.push_back(arg.isString() ? arg.toString() : arg.isBool() ? (arg.toBool() ? QString("Ja") : QString("Nein")) : arg.isDouble() ? QString::number(arg.toDouble(),'g',17) : QString{});}row.cells[3]=entry.value("condition").toString();result.push_back(row);}
    return result;
}
inline QJsonArray nativeMacroDraftActions(const QVector<NativeMacroRowDraft>& rows) {
    QJsonArray actions;
    for(const auto& row:rows) {
        const NativeMacroSignature* signature=nullptr;for(const auto& candidate:nativeMacroSignatures()) if(candidate.name==row.action) signature=&candidate;
        if(!signature) throw VbaRuntimeError(3251,"Unbekannte Makroaktion.");QStringList values;for(int index=0;index<3;++index) values.push_back(row.cells[index]);values.append(row.extraArguments);int count=signature->required;for(int index=0;index<values.size();++index) if(!values[index].isEmpty()) count=qMax(count,index+1);
        if(count>signature->types.size()) throw VbaRuntimeError(450,"Zu viele Argumente für die Makroaktion.");QJsonArray arguments;
        for(int index=0;index<count;++index) {
            const auto value=values.value(index),type=signature->types[index];
            if(((row.action=="RunMacro" && index==1) || (row.action=="GoToRecord" && type=="long") || (row.action=="FindRecord" && index>0)) && value.trimmed().isEmpty()) arguments.append(QJsonValue(QJsonValue::Null));
            else if(type=="text") arguments.append(value);
            else if(type=="boolean") {if(row.action=="MessageBox" && value.isEmpty()) arguments.append(true);else if(value.compare("Ja",Qt::CaseInsensitive)==0 || value.compare("True",Qt::CaseInsensitive)==0) arguments.append(true);else if(value.compare("Nein",Qt::CaseInsensitive)==0 || value.compare("False",Qt::CaseInsensitive)==0) arguments.append(false);else throw VbaRuntimeError(13,"Makroargument erfordert Ja oder Nein.");}
            else {bool valid=false;const auto integer=value.toLongLong(&valid);if(!valid || integer<std::numeric_limits<qint32>::min() || integer>std::numeric_limits<qint32>::max()) throw VbaRuntimeError(13,"Makroargument erfordert eine 32-Bit-Ganzzahl.");arguments.append(static_cast<int>(integer));}
        }
        QJsonObject entry{{"action",row.action},{"arguments",arguments}};const auto condition=row.cells[3].trimmed();if(!condition.isEmpty()) entry.insert("condition",condition);actions.append(entry);
    }
    return actions;
}
}
