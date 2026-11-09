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
std::optional<QMap<QString,AccessSourceFieldDisplay>> checkedAccessSourceFieldDisplayProperties(const QByteArray& json,const QStringList& columns) {
    const auto fail=[] {throw std::runtime_error("Originale Spaltenbeschriftungen sind nicht eindeutig nachgewiesen");};if(json.isEmpty()) return {};if(json.size()>1024*1024) fail();const auto document=QJsonDocument::fromJson(json);if(!document.isObject()) fail();const auto root=document.object();if(!root.contains("field_properties")) return {};if(root.value("field_properties_format").toString()!="JET4" || root.value("version").toDouble(-1)!=1 || !root.value("source_table").isString() || !root.value("library").isString() || !root.value("source_sha256").isString() || root.value("source_sha256").toString().size()!=64 || !root.value("field_properties").isArray()) fail();
    const auto digest=root.value("source_sha256").toString().toLatin1();for(const auto character:digest) if(!((character>='0' && character<='9') || (character>='a' && character<='f'))) fail();(void)checkedAccessSourceIndexes(json,columns);const auto entries=root.value("field_properties").toArray();if(entries.size()!=columns.size()) fail();QMap<QString,AccessSourceFieldDisplay> result;
    for(int index=0;index<columns.size();++index) {if(!entries[index].isObject()) fail();const auto object=entries[index].toObject();if(!object.value("name").isString() || object.value("name").toString()!=columns[index] || result.contains(columns[index].toCaseFolded())) fail();AccessSourceFieldDisplay properties;if(object.contains("format_proven")) {if(!object.value("format_proven").isBool()) fail();properties.formatProven=object.value("format_proven").toBool();}if(object.contains("format") && !properties.formatProven) fail();for(const auto& name:{QString("caption"),QString("description"),QString("format")}) if(object.contains(name)) {const auto value=object.value(name);const auto text=value.toString();if(!value.isString() || text.size()>4096 || !text.isValidUtf16() || text.contains(QChar(0))) fail();if(name=="caption") properties.caption=text;else if(name=="description") properties.description=text;else properties.format=text;}for(const auto& key:{QString("required"),QString("allow_zero_length")}) if(object.contains(key)) {if(!object.value(key).isBool()) fail();if(key=="required") properties.required=object.value(key).toBool();else properties.allowZeroLength=object.value(key).toBool();}result.insert(columns[index].toCaseFolded(),properties);}return result;
}

}
