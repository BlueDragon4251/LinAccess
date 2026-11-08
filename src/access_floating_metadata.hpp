#pragma once
#include <QByteArray>
#include <QMap>
#include <QStringList>
#include <QVector>
#include <optional>
namespace linaccess {
template<class T> struct AccessFloatingValues {
    QByteArray sourceDigest;
    qsizetype rowCount{};
    QMap<QString,QVector<std::optional<T>>> columns;
};
using AccessSingleValues=AccessFloatingValues<float>;
using AccessDoubleValues=AccessFloatingValues<double>;
AccessSingleValues checkedAccessSingleValues(const QByteArray& json,const QStringList& names,const QString& sourceTable);
AccessDoubleValues checkedAccessDoubleValues(const QByteArray& json,const QStringList& names,const QString& sourceTable);
bool mdbFloatingExportMatches(const QString& csv,double value,int precision);
QByteArray accessSourceDigest(const QString& path);
}
