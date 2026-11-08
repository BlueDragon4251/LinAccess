#include "access_floating_metadata.hpp"
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
#include <type_traits>
#include <sstream>
#include <iomanip>
#include <locale>
#include <QRegularExpression>
namespace linaccess {
QByteArray accessSourceDigest(const QString& path) {
    QFile source(path); if(!source.open(QIODevice::ReadOnly)) throw std::runtime_error("Single-Quelle kann nicht gelesen werden");
    QCryptographicHash hash(QCryptographicHash::Sha256); if(!hash.addData(&source)) throw std::runtime_error("Single-Quellhash konnte nicht geprüft werden");
    return hash.result().toHex();
}
template<class T> AccessFloatingValues<T> checkedFloatingValues(const QByteArray& json,const QStringList& names,const QString& sourceTable) {
    static_assert((sizeof(T)==4 || sizeof(T)==8) && std::numeric_limits<T>::is_iec559);
    using Bits=std::conditional_t<sizeof(T)==4,quint32,quint64>;
    const auto fail=[] { throw std::runtime_error("Exakte Gleitkomma-Quelldaten sind unvollständig oder nicht unterstützt"); };
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
    AccessFloatingValues<T> result; result.rowCount=rows.size();
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
            if(!cells[i].isString() || !hex(cells[i].toString(),sizeof(T)*2)) fail();
            bool valid=false; const auto bits=cells[i].toString().toULongLong(&valid,16); if(!valid) fail();
            const auto number=std::bit_cast<T>(static_cast<Bits>(bits));
            if(!std::isfinite(number) || (number==0 && std::signbit(number))) fail();
            values.append(number);
        }
    }
    return result;
}
AccessSingleValues checkedAccessSingleValues(const QByteArray& json,const QStringList& names,const QString& table) { return checkedFloatingValues<float>(json,names,table); }
AccessDoubleValues checkedAccessDoubleValues(const QByteArray& json,const QStringList& names,const QString& table) { return checkedFloatingValues<double>(json,names,table); }
bool mdbFloatingExportMatches(const QString& csv,double value,int precision) {
    if(!std::isfinite(value) || (precision!=8 && precision!=16)) return false;
    const auto canonical=[](const QString& input)->std::optional<QString> {
        const auto text=input.trimmed(); if(text.size()>128) return {};
        static const QRegularExpression number(R"(^([+-]?)([0-9]+)(?:\.([0-9]*))?(?:[eE]([+-]?[0-9]+))?$)");
        const auto match=number.match(text); if(!match.hasMatch()) return {};
        bool valid=true; int exponent=match.captured(4).isEmpty() ? 0 : match.captured(4).toInt(&valid);
        if(!valid || exponent< -4096 || exponent>4096) return {};
        auto digits=match.captured(2)+match.captured(3); exponent-=match.capturedLength(3)>0 ? match.capturedLength(3) : 0;
        while(digits.startsWith('0')) digits.remove(0,1);
        if(digits.isEmpty()) return QString("0");
        while(digits.endsWith('0')) { digits.chop(1); ++exponent; }
        return (match.captured(1)=="-" ? QString("-") : QString{})+digits+"e"+QString::number(exponent);
    };
    std::ostringstream formatted; formatted.imbue(std::locale::classic());
    formatted<<std::setprecision(precision)<<std::defaultfloat<<value;
    const auto expected=canonical(QString::fromStdString(formatted.str())),actual=canonical(csv);
    return expected && actual && *expected==*actual;
}
}
