#pragma once
#include "vba_expression.hpp"
#include <QJsonArray>
#include <QJsonObject>
#include <QSet>
#include <QStringList>
#include <QVector>
namespace linaccess {
struct NativeReportGroupDefinition {QString field;bool descending{},header{true};QString footer;bool repeatHeader{};};
inline QVector<NativeReportGroupDefinition> nativeReportGroups(const QJsonValue& value,const QStringList& columns,const QVector<QString>& types={}) {
    if(value.isUndefined()) return {};
    if(!value.isArray() || value.toArray().size()>10) throw VbaRuntimeError(3251,"Native report grouping requires at most ten levels");
    QVector<NativeReportGroupDefinition> groups;QSet<QString> fields;
    for(const auto& entry:value.toArray()) {
        if(!entry.isObject()) throw VbaRuntimeError(5,"Invalid native report group");const auto item=entry.toObject();
        if(!item.value("field").isString() || (item.contains("descending") && !item.value("descending").isBool())
            || (item.contains("header") && !item.value("header").isBool()) || (item.contains("footer_expression") && !item.value("footer_expression").isString()) || (item.contains("repeat_header") && !item.value("repeat_header").isBool())) throw VbaRuntimeError(5,"Invalid native report group properties");
        QString field;int position=-1;for(const auto& candidate:columns) if(candidate.compare(item.value("field").toString(),Qt::CaseInsensitive)==0) {if(!field.isEmpty()) throw VbaRuntimeError(5,"Ambiguous native report group field");field=candidate;position=columns.indexOf(candidate);}
        const auto type=types.value(position).toUpper();if(type=="BLOB" || type.contains("BINARY")) throw VbaRuntimeError(3251,"Binary report group keys are unsupported");
        if(field.isEmpty()) throw VbaRuntimeError(3265,"Native report group field is not in the source");
        if(field.contains(']') || field.contains('\n') || field.contains('\r') || field.contains(QChar(0))) throw VbaRuntimeError(3251,"Native report group field identifier unsupported");
        if(fields.contains(field.toCaseFolded())) throw VbaRuntimeError(5,"Duplicate native report group field");fields.insert(field.toCaseFolded());
        const auto expression=item.value("footer_expression").toString().trimmed();
        if((!expression.isEmpty() && !expression.startsWith('=')) || expression.size()>32768 || expression.contains('\n') || expression.contains('\r') || expression.contains(QChar(0))) throw VbaRuntimeError(5,"Group footer calculation must be a single expression starting with equals");
        const bool repeat=item.value("repeat_header").toBool();if(repeat && !item.value("header").toBool(true)) throw VbaRuntimeError(5,"Repeating group header requires a header");
        groups.append({field,item.value("descending").toBool(),item.value("header").toBool(true),expression,repeat});
    }
    return groups;
}
}
