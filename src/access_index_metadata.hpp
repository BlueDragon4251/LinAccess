#pragma once
#include <QByteArray>
#include <QStringList>
#include <QVector>
#include <QMap>
#include <optional>
namespace linaccess {
struct AccessSourceIndex {
    QString name;
    int kind{},flags{};
    QStringList columns;
    QVector<bool> descending;
    bool operator==(const AccessSourceIndex&) const = default;
};
struct AccessSourceFieldDisplay {std::optional<QString> caption,description;std::optional<bool> required,allowZeroLength;};
// Missing component means unknown, not a proven empty property catalog.
std::optional<QMap<QString,AccessSourceFieldDisplay>> checkedAccessSourceFieldDisplayProperties(const QByteArray& json,const QStringList& columns);
// Strict versioned metadata produced by the read-only native MDB helper.
// Caller supplies the actual ordered column names; no index names are inferred.
QVector<AccessSourceIndex> checkedAccessSourceIndexes(const QByteArray& json,const QStringList& columns);
}
