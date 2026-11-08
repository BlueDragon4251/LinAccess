#include "access_index_metadata.hpp"
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QSet>
#include <stdexcept>
namespace linaccess {
QVector<AccessSourceIndex> checkedAccessSourceIndexes(const QByteArray& json,const QStringList& columns) {
    const auto fail=[] { throw std::runtime_error("Originale Indexmetadaten sind nicht eindeutig nachgewiesen"); };
    const auto validName=[](const QString& name) { return !name.isEmpty() && name.isValidUtf16() && !name.contains(QChar(0)); };
    if(json.isEmpty() || json.size()>1024*1024) fail();
    const auto document=QJsonDocument::fromJson(json);
    if(!document.isObject()) fail(); const auto root=document.object();
    if(root.value("version").toDouble(-1)!=1 || !root.value("source_table").isString() || !validName(root.value("source_table").toString())
        || !root.value("columns").isArray() || !root.value("indexes").isArray()) fail();
    const auto storedColumns=root.value("columns").toArray(),indexes=root.value("indexes").toArray();
    if(storedColumns.size()!=columns.size() || indexes.size()>4096) fail();
    QSet<QString> names;
    for(int i=0;i<columns.size();++i) {
        if(!validName(columns[i]) || names.contains(columns[i].toCaseFolded()) || !storedColumns[i].isString() || storedColumns[i].toString()!=columns[i]) fail();
        names.insert(columns[i].toCaseFolded());
    }
    QVector<AccessSourceIndex> result; QSet<QString> indexNames; int primary=0;
    for(const auto& value:indexes) {
        if(!value.isObject()) fail(); const auto object=value.toObject();
        if(!object.value("name").isString() || !object.value("kind").isDouble() || !object.value("flags").isDouble() || !object.value("fields").isArray()) fail();
        const auto name=object.value("name").toString(); const int kind=object.value("kind").toInt(-1),flags=object.value("flags").toInt(-1);
        if(!validName(name) || indexNames.contains(name.toCaseFolded()) || kind<0 || kind>2 || flags<0 || flags>255
            || object.value("kind").toDouble()!=kind || object.value("flags").toDouble()!=flags) fail();
        indexNames.insert(name.toCaseFolded());
        const auto fields=object.value("fields").toArray(); if(fields.size()>10 || (fields.isEmpty() && kind!=2) || (kind==1 && (++primary>1 || !(flags&1)))) fail();
        AccessSourceIndex index; index.name=name; index.kind=kind; index.flags=flags; QSet<QString> seen;
        for(const auto& field:fields) {
            if(!field.isObject()) fail(); const auto entry=field.toObject();
            if(!entry.value("name").isString() || !entry.value("descending").isBool()) fail(); const auto column=entry.value("name").toString();
            if(!columns.contains(column) || seen.contains(column.toCaseFolded())) fail(); seen.insert(column.toCaseFolded());
            index.columns.append(column); index.descending.append(entry.value("descending").toBool());
        }
        result.append(index);
    }
    return result;
}
}
