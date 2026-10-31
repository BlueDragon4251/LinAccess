#include "vba_qt_value.hpp"
#include <QJsonArray>
#include <QDateTime>
#include <QJsonDocument>
#include <QJsonObject>
#include <QSqlQuery>
#include <cmath>
#include <limits>
namespace linaccess {
QMap<QString,QString> sqliteColumnTypes(QSqlDatabase db,const QString& table) {
    QMap<QString,QString> types;
    QSqlQuery query(db);
    auto name = table; name.replace('"',"\"\"");
    if (query.exec("PRAGMA table_info(\"" + name + "\")"))
        while (query.next()) types.insert(query.value(1).toString().toCaseFolded(),query.value(2).toString().toUpper());
    query.prepare("SELECT definition FROM _la_table_schema WHERE table_name=?"); query.addBindValue(table);
    if (query.exec() && query.next()) {
        const auto metadata = QJsonDocument::fromJson(query.value(0).toByteArray()).array();
        for (const auto& value : metadata) {
            const auto column = value.toObject();
            const auto source = column["source_type"].toString();
            if (source == "BOOLEAN" || source == "BIT") types[column["name"].toString().toCaseFolded()] = "BOOLEAN";
            else if (source == "DATETIME" || source == "DATE") types[column["name"].toString().toCaseFolded()] = "DATETIME";
        }
    }
    return types;
}
namespace {
QString dateText(const QVariant& value) {
    if (value.metaType().id() == QMetaType::QDateTime)
        return value.toDateTime().toString("yyyy-MM-dd HH:mm:ss");
    for (const auto& format : QStringList{"yyyy-MM-dd HH:mm:ss","yyyy-MM-dd"}) {
        const auto parsed = QDateTime::fromString(value.toString(),format);
        if (parsed.isValid()) return parsed.toString("yyyy-MM-dd HH:mm:ss");
    }
    const auto time = QTime::fromString(value.toString(),"HH:mm:ss");
    if (time.isValid()) return "1899-12-30 " + time.toString("HH:mm:ss");
    throw VbaRuntimeError(13,"Date field requires an ISO date/time value");
}
bool integerType(const QString& type) { return type.contains("INT",Qt::CaseInsensitive); }
bool realType(const QString& type) {
    return type.contains("REAL",Qt::CaseInsensitive) || type.contains("FLOA",Qt::CaseInsensitive)
        || type.contains("DOUB",Qt::CaseInsensitive) || type.contains("NUM",Qt::CaseInsensitive)
        || type.contains("DEC",Qt::CaseInsensitive);
}
bool boolean(const QVariant& value) {
    if (value.metaType().id() == QMetaType::QString) {
        const auto string = value.toString();
        if (string.compare("true",Qt::CaseInsensitive) == 0) return true;
        if (string.compare("false",Qt::CaseInsensitive) == 0) return false;
    }
    bool valid = false; const auto number = value.toDouble(&valid);
    if (!valid || !std::isfinite(number)) throw VbaRuntimeError(13,"Boolean field type mismatch");
    return number != 0;
}
}
VbaValue qtFieldValue(const QVariant& value,const QString& type) {
    if (value.isNull()) return {};
    if ((integerType(type) || realType(type) || type.contains("BOOL",Qt::CaseInsensitive)) && value.toString().isEmpty()) return {};
    if (type.contains("DATE",Qt::CaseInsensitive)) return VbaValue::string(dateText(value).toUtf8().toStdString());
    if (type.contains("BOOL",Qt::CaseInsensitive) || value.metaType().id() == QMetaType::Bool)
        return VbaValue::logical(boolean(value));
    if (integerType(type) || value.metaType().id() == QMetaType::Int || value.metaType().id() == QMetaType::LongLong) {
        bool valid = false; const auto integer = value.toLongLong(&valid);
        if (!valid) throw VbaRuntimeError(13,"Integer field type mismatch");
        // A Double-only numeric Variant must not silently round a 64-bit integer.
        if (integer < -9007199254740992LL || integer > 9007199254740992LL)
            throw VbaRuntimeError(6,"Exact 64-bit integer Variant not implemented");
        return VbaValue::numeric(static_cast<double>(integer));
    }
    if (realType(type) || value.metaType().id() == QMetaType::Double) {
        bool valid = false; const auto number = value.toDouble(&valid);
        if (!valid || !std::isfinite(number)) throw VbaRuntimeError(13,"Numeric field type mismatch");
        return VbaValue::numeric(number);
    }
    if (value.metaType().id() == QMetaType::QByteArray) throw VbaRuntimeError(13,"Binary Variant not implemented");
    return VbaValue::string(value.toString().toUtf8().toStdString());
}
QVariant sqlFieldValue(const VbaValue& value,const QString& type) {
    if (value.kind == VbaValue::Kind::Null) return {};
    if (type.contains("DATE",Qt::CaseInsensitive)) {
        if (value.kind != VbaValue::Kind::Text) throw VbaRuntimeError(13,"Date serial Variant conversion not implemented");
        return dateText(QString::fromUtf8(value.text));
    }
    if (type.contains("BOOL",Qt::CaseInsensitive)) {
        const auto converted = value.kind == VbaValue::Kind::Text ? QVariant(QString::fromStdString(value.text)) : QVariant(value.asNumber());
        return boolean(converted) ? -1 : 0;
    }
    if (integerType(type)) {
        const auto number = value.asNumber();
        const auto lower = std::floor(static_cast<long double>(number));
        const auto fraction = static_cast<long double>(number)-lower;
        const auto rounded = fraction < 0.5L ? lower : fraction > 0.5L ? lower+1
            : std::fmod(std::abs(lower),2.0L) == 0 ? lower : lower+1;
        if (!std::isfinite(rounded) || rounded < std::numeric_limits<qlonglong>::min()
            || rounded > std::numeric_limits<qlonglong>::max()) throw VbaRuntimeError(6,"Integer field overflow");
        return static_cast<qlonglong>(rounded);
    }
    if (realType(type)) {
        const auto number = value.asNumber();
        if (!std::isfinite(number)) throw VbaRuntimeError(6,"Numeric field overflow");
        return number;
    }
    if (type.contains("TEXT",Qt::CaseInsensitive) || type.contains("CHAR",Qt::CaseInsensitive)
        || type.contains("CLOB",Qt::CaseInsensitive)) return QString::fromUtf8(value.asText());
    switch (value.kind) {
    case VbaValue::Kind::Empty: return QString{};
    case VbaValue::Kind::Text: return QString::fromUtf8(value.text);
    case VbaValue::Kind::Currency: return QString::fromStdString(value.asText());
    case VbaValue::Kind::Number: return value.number;
    case VbaValue::Kind::Boolean: return value.boolean ? -1 : 0;
    default: throw VbaRuntimeError(13,"Object cannot be stored in a field");
    }
}
}
