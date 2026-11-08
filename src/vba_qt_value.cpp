#include "vba_qt_value.hpp"
#include "access_autonumber.hpp"
#include "vba_currency.hpp"
#include "vba_calendar.hpp"
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
            else if(source=="CURRENCY") {
                const auto name=column["name"].toString().toCaseFolded();
                if(types.value(name)!="REAL") throw VbaRuntimeError(3251,"Currency source requires the supported REAL backing");
                types[name]="CURRENCY_REAL";
            }
            else if (source == "TINYINT") types[column["name"].toString().toCaseFolded()] = "INTEGER1";
            else if (source == "SMALLINT") types[column["name"].toString().toCaseFolded()] = "INTEGER2";
            else if (source == "INT") types[column["name"].toString().toCaseFolded()] = "INTEGER4";
        }
    }
    query.finish();
    const auto counters=jetLongAccessColumns(db,table);
    for(const auto& column:counters) {
        if(!types.contains(column) || (types.value(column)!="INTEGER" && types.value(column)!="INT" && types.value(column)!="INTEGER4"))
            throw VbaRuntimeError(3251,"Jet Long generator requires integer storage");
        types[column]="INTEGER4";
    }
    return types;
}
namespace {
std::string_view integerSubtype(const QString& type) {
    const auto name=type.trimmed().toUpper();
    return name=="INTEGER1" ? "byte" : name=="INTEGER2" ? "integer" : name=="INTEGER4" ? "long" : "";
}
QString dateText(const QVariant& value) {
    if (value.metaType().id() == QMetaType::QDateTime) {
        const auto date=value.toDateTime();
        if(!date.isValid() || date.date().year()<100 || date.date().year()>9999) throw VbaRuntimeError(13,"Date field year out of range");
        return date.toString("yyyy-MM-dd HH:mm:ss");
    }
    const auto text=value.toString();
    if(text.size()==10 || ((text.size()==16 || text.size()==19) && (text[10]==' ' || text[10]=='T'))) {
        const auto date=QDate::fromString(text.left(10),Qt::ISODate);
        const auto time=text.size()==10 ? QTime(0,0) : QTime::fromString(text.mid(11),text.size()==16 ? "HH:mm" : "HH:mm:ss");
        if(date.isValid() && date.year()>=100 && date.year()<=9999 && time.isValid())
            return date.toString(Qt::ISODate)+" "+time.toString("HH:mm:ss");
    }
    const auto time = QTime::fromString(value.toString(),text.size()==5 ? "HH:mm" : "HH:mm:ss");
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
    if(type=="CURRENCY_REAL") {
        if(value.metaType().id()==QMetaType::QString) {
            const auto exact=convertCurrency(VbaValue::string(value.toString().toUtf8().toStdString()));
            (void)currencyToRealStorage(exact); return exact;
        }
        if(value.metaType().id()==QMetaType::QByteArray) throw VbaRuntimeError(13,"Binary Currency field storage is unsupported");
        bool valid=false; const auto number=value.toDouble(&valid);
        if(!valid) throw VbaRuntimeError(13,"Currency field type mismatch");
        return currencyFromRealStorage(number);
    }
    if (type.contains("DATE",Qt::CaseInsensitive)) return vbaDateValue(VbaValue::string(dateText(value).toUtf8().toStdString()));
    if (type.contains("BOOL",Qt::CaseInsensitive) || value.metaType().id() == QMetaType::Bool)
        return VbaValue::logical(boolean(value));
    if(const auto subtype=integerSubtype(type); !subtype.empty()) {
        if(value.metaType().id()==QMetaType::Double || value.metaType().id()==QMetaType::Float) {
            const auto number=value.toDouble();
            if(!std::isfinite(number) || std::trunc(number)!=number) throw VbaRuntimeError(13,"Integer field storage is not integral");
        }
        if(value.metaType().id()==QMetaType::ULongLong && value.toULongLong()>2147483647ULL) throw VbaRuntimeError(6,"Integer field overflow");
        if(value.metaType().id()==QMetaType::QByteArray) throw VbaRuntimeError(13,"Binary integer field storage is unsupported");
        bool valid=false; const auto integer=value.toLongLong(&valid);
        if(!valid) throw VbaRuntimeError(13,"Integer field type mismatch");
        return coerceVbaModuleValue(VbaValue::numeric(static_cast<double>(integer)),subtype);
    }
    if (integerType(type) || value.metaType().id() == QMetaType::Int || value.metaType().id() == QMetaType::LongLong) {
        bool valid = false; const auto integer = value.toLongLong(&valid);
        if (!valid) throw VbaRuntimeError(13,"Integer field type mismatch");
        // Generic SQLite integers still use Double; do not silently round 64-bit data.
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
QString qtFieldText(const QVariant& value,const QString& type) {
    if(value.isNull()) return {};
    return type=="CURRENCY_REAL" ? QString::fromStdString(qtFieldValue(value,type).asText()) : value.toString();
}
QVariant sqlFieldValue(const VbaValue& value,const QString& type) {
    if (value.kind == VbaValue::Kind::Null) return {};
    if(type=="CURRENCY_REAL") return currencyToRealStorage(value);
    if (type.contains("DATE",Qt::CaseInsensitive)) {
        const auto converted=value.kind==VbaValue::Kind::Text ? value : vbaDateValue(value);
        if(converted.kind==VbaValue::Kind::Date) {
            const auto seconds=std::abs(converted.number-std::trunc(converted.number))*86400.0;
            if(std::abs(seconds-std::round(seconds))>0.00005) throw VbaRuntimeError(3251,"Subsecond Date storage is not implemented");
        }
        return dateText(QString::fromStdString(converted.asText()));
    }
    if (type.contains("BOOL",Qt::CaseInsensitive)) {
        const auto converted = value.kind == VbaValue::Kind::Text ? QVariant(QString::fromStdString(value.text)) : QVariant(value.asNumber());
        return boolean(converted) ? -1 : 0;
    }
    if(const auto subtype=integerSubtype(type); !subtype.empty()) {
        const auto converted=coerceVbaModuleValue(value,subtype);
        return static_cast<qlonglong>(converted.number);
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
    case VbaValue::Kind::Date: return QString::fromStdString(value.asText());
    case VbaValue::Kind::Currency: return QString::fromStdString(value.asText());
    case VbaValue::Kind::Number: return value.number;
    case VbaValue::Kind::Boolean: return value.boolean ? -1 : 0;
    default: throw VbaRuntimeError(13,"Object cannot be stored in a field");
    }
}
}
