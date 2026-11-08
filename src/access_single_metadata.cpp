#include "access_single_metadata.hpp"
#include <QCryptographicHash>
#include <QFile>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QSet>
#include <bit>
#include <cmath>
#include <limits>
#include <stdexcept>
namespace linaccess {
QByteArray accessSourceDigest(const QString& path) {
    QFile source(path); if(!source.open(QIODevice::ReadOnly)) throw std::runtime_error("Single-Quelle kann nicht gelesen werden");
    QCryptographicHash hash(QCryptographicHash::Sha256); if(!hash.addData(&source)) throw std::runtime_error("Single-Quellhash konnte nicht geprüft werden");
    return hash.result().toHex();
}
AccessSingleValues checkedAccessSingleValues(const QByteArray& json,const QStringList& names,const QString& sourceTable) {
    static_assert(sizeof(float)==4 && std::numeric_limits<float>::is_iec559);
    const auto fail=[] { throw std::runtime_error("Exakte Single-Quelldaten sind unvollständig oder nicht unterstützt"); };
    if(json.isEmpty() || json.size()>16*1024*1024 || names.isEmpty() || names.size()>32767 || sourceTable.isEmpty() || !sourceTable.isValidUtf16() || sourceTable.contains(QChar(0))) fail();
    const auto document=QJsonDocument::fromJson(json); if(!document.isObject()) fail(); const auto root=document.object();
    if(root.value("version").toDouble(-1)!=1 || !root.value("columns").isArray() || !root.value("rows").isArray()
        || !root.value("source_table").isString() || root.value("source_table").toString()!=sourceTable || !root.value("source_sha256").isString()) fail();
    const auto columns=root.value("columns").toArray(),rows=root.value("rows").toArray();
    if(columns.size()!=names.size() || rows.size()>1000000/names.size()) fail();
    const auto hex=[](const QString& value,qsizetype size) {
        if(value.size()!=size) return false;
        for(const auto ch:value) if(!((ch>='0' && ch<='9') || (ch>='a' && ch<='f') || (ch>='A' && ch<='F'))) return false;
        return true;
    };
    AccessSingleValues result; result.rowCount=rows.size();
    const auto digest=root.value("source_sha256").toString(); if(!hex(digest,64)) fail(); result.sourceDigest=digest.toLatin1().toLower();
    QSet<QString> seen;
    for(int i=0;i<names.size();++i) {
        if(names[i].isEmpty() || !names[i].isValidUtf16() || names[i].contains(QChar(0)) || seen.contains(names[i].toCaseFolded())
            || !columns[i].isString() || columns[i].toString()!=names[i]) fail();
        seen.insert(names[i].toCaseFolded()); result.columns.insert(names[i].toCaseFolded(),{});
    }
    for(const auto& value:rows) {
        if(!value.isArray()) fail(); const auto cells=value.toArray(); if(cells.size()!=names.size()) fail();
        for(int i=0;i<names.size();++i) {
            auto& values=result.columns[names[i].toCaseFolded()];
            if(cells[i].isNull()) { values.append(std::nullopt); continue; }
            if(!cells[i].isString() || !hex(cells[i].toString(),8)) fail();
            bool valid=false; const auto bits=cells[i].toString().toUInt(&valid,16); if(!valid) fail();
            const auto number=std::bit_cast<float>(static_cast<quint32>(bits));
            if(!std::isfinite(number) || (number==0 && std::signbit(number))) fail();
            values.append(number);
        }
    }
    return result;
}
}
