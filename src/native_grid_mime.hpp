#pragma once
#include "vba_runtime.hpp"
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QVariant>
#include <cmath>
namespace linaccess {
inline constexpr auto nativeGridMime="application/x-linaccess-grid-v1";
inline constexpr qsizetype nativeGridMimeLimit=8*1024*1024;
inline QJsonObject nativeGridEncode(const QVariant& value) {
    if(value.isNull()) return {{"kind","null"}};
    const auto kind=value.metaType().id();
    if(kind==QMetaType::Int || kind==QMetaType::LongLong) return {{"kind","integer"},{"value",QString::number(value.toLongLong())}};
    if(kind==QMetaType::Double) {
        if(!std::isfinite(value.toDouble()) || (value.toDouble()==0 && std::signbit(value.toDouble()))) throw VbaRuntimeError(6,"Nonfinite grid clipboard number");
        return {{"kind","real"},{"value",value.toDouble()}};
    }
    if(kind==QMetaType::Bool) return {{"kind","boolean"},{"value",value.toBool()}};
    if(kind==QMetaType::QString) {
        const auto text=value.toString();
        if(text.size()>nativeGridMimeLimit/2 || !text.isValidUtf16() || text.contains(QChar(0))) throw VbaRuntimeError(3251,"Grid clipboard text is not representable");
        return {{"kind","text"},{"value",text}};
    }
    if(kind==QMetaType::QByteArray) {
        if(value.toByteArray().size()>nativeGridMimeLimit) throw VbaRuntimeError(3035,"Grid clipboard binary limit exceeded");
        return {{"kind","binary"},{"value",QString::fromLatin1(value.toByteArray().toBase64())}};
    }
    throw VbaRuntimeError(13,"Grid clipboard storage type is unsupported");
}
inline QVariant nativeGridDecode(const QJsonValue& cell) {
    if(!cell.isObject()) throw VbaRuntimeError(13,"Invalid typed grid cell");
    const auto object=cell.toObject();const auto kind=object.value("kind").toString();const auto value=object.value("value");
    if(kind=="null" && value.isUndefined()) return {};
    if(kind=="integer" && value.isString()) {
        bool valid=false;const auto integer=value.toString().toLongLong(&valid);
        if(valid && QString::number(integer)==value.toString()) return QVariant::fromValue(integer);
    }
    if(kind=="real" && value.isDouble() && std::isfinite(value.toDouble())) return value.toDouble();
    if(kind=="boolean" && value.isBool()) return value.toBool();
    if(kind=="text" && value.isString()) {
        const auto text=value.toString();if(text.isValidUtf16() && !text.contains(QChar(0))) return text;
    }
    if(kind=="binary" && value.isString()) {
        const auto base64=value.toString().toLatin1();
        const auto decoded=QByteArray::fromBase64Encoding(base64,QByteArray::AbortOnBase64DecodingErrors);
        if(decoded && decoded.decoded.toBase64()==base64) return decoded.decoded.isNull() ? QByteArray("",0) : decoded.decoded;
    }
    throw VbaRuntimeError(13,"Invalid typed grid value");
}
inline QVector<QVector<QVariant>> nativeGridDecodeBlock(const QByteArray& bytes) {
    if(bytes.size()>nativeGridMimeLimit) throw VbaRuntimeError(3035,"Grid clipboard size limit exceeded");
    QJsonParseError error;const auto document=QJsonDocument::fromJson(bytes,&error);
    if(error.error!=QJsonParseError::NoError || !document.isObject()) throw VbaRuntimeError(13,"Invalid grid clipboard document");
    const auto root=document.object();
    if(root.value("version")!=QJsonValue(1) || !root.value("rows").isArray()) throw VbaRuntimeError(13,"Invalid grid clipboard version");
    const auto rows=root.value("rows").toArray();QVector<QVector<QVariant>> result;qsizetype columns=0;
    if(rows.isEmpty()) throw VbaRuntimeError(13,"Empty grid clipboard block");
    for(const auto& row:rows) {
        if(!row.isArray()) throw VbaRuntimeError(13,"Invalid grid clipboard row");
        const auto cells=row.toArray();if(!columns) columns=cells.size();
        if(!columns || cells.size()!=columns || rows.size()>1000000/columns) throw VbaRuntimeError(13,"Invalid grid clipboard shape");
        QVector<QVariant> values;values.reserve(columns);for(const auto& cell:cells) values.push_back(nativeGridDecode(cell));result.push_back(std::move(values));
    }
    return result;
}
}
