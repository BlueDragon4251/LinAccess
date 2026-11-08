#pragma once
#include <QByteArray>
#include <QMap>
#include <QStringList>
#include <QVector>
#include <optional>
namespace linaccess {
struct AccessSingleValues {
    QByteArray sourceDigest;
    qsizetype rowCount{};
    QMap<QString,QVector<std::optional<float>>> columns;
};
AccessSingleValues checkedAccessSingleValues(const QByteArray& json,const QStringList& names,const QString& sourceTable);
QByteArray accessSourceDigest(const QString& path);
}
