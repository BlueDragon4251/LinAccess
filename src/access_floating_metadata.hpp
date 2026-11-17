#pragma once
#include <QByteArray>
#include <QMap>
#include <QStringList>
#include <QVector>
#include <optional>
namespace linaccess {
template<class T> struct AccessRawValues {
    QByteArray sourceDigest;
    qsizetype rowCount{};
    QMap<QString,QVector<std::optional<T>>> columns;
};
template<class T> using AccessFloatingValues=AccessRawValues<T>;
using AccessSingleValues=AccessFloatingValues<float>;
using AccessDoubleValues=AccessFloatingValues<double>;
using AccessCurrencyValues=AccessRawValues<qint64>;
AccessSingleValues checkedAccessSingleValues(const QByteArray& json,const QStringList& names,const QString& sourceTable);
AccessDoubleValues checkedAccessDoubleValues(const QByteArray& json,const QStringList& names,const QString& sourceTable);
AccessCurrencyValues checkedAccessCurrencyValues(const QByteArray& json,const QStringList& names,const QString& sourceTable,const QString& expectedFormat={});
std::optional<qint64> accessCurrencyLiteralTicks(const QString& text);
bool mdbCurrencyExportMatches(const QString& csv,qint64 ticks);
struct AccessDecimalValues {
    QByteArray sourceDigest;
    qsizetype rowCount{};
    QMap<QString,QVector<std::optional<QString>>> columns;
};
QMap<QString,QPair<int,int>> checkedAccessDecimalColumnTypes(const QByteArray& json,const QStringList& names);
AccessDecimalValues checkedAccessDecimalValues(const QByteArray& json,const QStringList& names,const QString& sourceTable,const QMap<QString,QPair<int,int>>& types);
bool mdbFloatingExportMatches(const QString& csv,double value,int precision);
QByteArray accessSourceDigest(const QString& path);
}
