#pragma once
#include "vba_runtime.hpp"
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QSet>
#include <QStringList>
#include <QVariant>
#include <cmath>
namespace linaccess {struct NativeGridCurrency {qint64 units{};};}
Q_DECLARE_METATYPE(linaccess::NativeGridCurrency)
namespace linaccess {
inline constexpr auto nativeGridMime="application/x-linaccess-grid-v1";
inline constexpr auto nativeRecordMime="application/x-linaccess-record-v1";
inline constexpr qsizetype nativeGridMimeLimit=8*1024*1024;
inline QJsonObject nativeGridEncode(const QVariant& value,const QString& type={}) {
    if(value.isNull()) return {{"kind","null"}};
    const auto kind=value.metaType().id();
    if(type=="CURRENCY_TICKS") {
        if(kind!=QMetaType::Int && kind!=QMetaType::LongLong) throw VbaRuntimeError(13,"Exact Currency clipboard requires integer ticks");
        return {{"kind","currency"},{"value",QString::number(value.toLongLong())}};
    }
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
    if((kind=="integer" || kind=="currency") && value.isString()) {
        bool valid=false;const auto integer=value.toString().toLongLong(&valid);
        if(valid && QString::number(integer)==value.toString()) return kind=="currency" ? QVariant::fromValue(NativeGridCurrency{integer}) : QVariant::fromValue(integer);
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
struct NativeClipboardBlock {
    QStringList fields;
    QVector<QVector<QVariant>> rows;
};

inline QJsonObject nativeClipboardObject(const QByteArray& bytes) {
    if(bytes.size()>nativeGridMimeLimit) throw VbaRuntimeError(3035,"Typed clipboard size limit exceeded");
    QJsonParseError error;
    const auto document=QJsonDocument::fromJson(bytes,&error);
    if(error.error!=QJsonParseError::NoError || !document.isObject()) throw VbaRuntimeError(13,"Invalid typed clipboard document");
    const auto root=document.object();
    if(root.value("version")!=QJsonValue(1)) throw VbaRuntimeError(13,"Invalid typed clipboard version");
    return root;
}

inline QString nativeClipboardField(const QJsonValue& value,QSet<QString>& seen) {
    const auto field=value.toString();
    const auto key=field.toCaseFolded();
    if(!value.isString() || field.isEmpty() || !field.isValidUtf16() || field.contains(QChar(0)) || seen.contains(key))
        throw VbaRuntimeError(13,"Invalid or ambiguous clipboard field name");
    seen.insert(key);
    return field;
}

inline QVector<QVector<QVariant>> nativeGridDecodeBlock(const QByteArray& bytes,QStringList* fieldNames=nullptr) {
    const auto root=nativeClipboardObject(bytes);
    if(!root.value("rows").isArray()) throw VbaRuntimeError(13,"Invalid grid clipboard rows");
    const auto rows=root.value("rows").toArray();
    QVector<QVector<QVariant>> result;
    qsizetype columns=0;
    if(rows.isEmpty()) throw VbaRuntimeError(13,"Empty grid clipboard block");
    for(const auto& row:rows) {
        if(!row.isArray()) throw VbaRuntimeError(13,"Invalid grid clipboard row");
        const auto cells=row.toArray();if(!columns) columns=cells.size();
        if(!columns || cells.size()!=columns || rows.size()>1000000/columns) throw VbaRuntimeError(13,"Invalid grid clipboard shape");
        QVector<QVariant> values;
        values.reserve(columns);
        for(const auto& cell:cells) values.push_back(nativeGridDecode(cell));
        result.push_back(std::move(values));
    }
    if(fieldNames) {
        QStringList names;
        const auto fields=root.value("fields");
        if(!fields.isUndefined()) {
            if(!fields.isArray() || fields.toArray().size()!=columns) throw VbaRuntimeError(13,"Grid clipboard field shape mismatch");
            QSet<QString> seen;
            for(const auto& field:fields.toArray()) names.append(nativeClipboardField(field,seen));
        }
        *fieldNames=std::move(names);
    }
    return result;
}

inline NativeClipboardBlock nativeRecordDecode(const QByteArray& bytes) {
    const auto root=nativeClipboardObject(bytes);
    const auto fields=root.value("fields").toArray();
    if(fields.isEmpty() || fields.size()>255) throw VbaRuntimeError(13,"Invalid form clipboard field shape");
    NativeClipboardBlock result;
    QVector<QVariant> row;
    QSet<QString> seen;
    for(const auto& field:fields) {
        if(!field.isObject()) throw VbaRuntimeError(13,"Invalid form clipboard field");
        result.fields.append(nativeClipboardField(field.toObject().value("name"),seen));
        row.append(nativeGridDecode(field));
    }
    result.rows.append(std::move(row));
    return result;
}
}
