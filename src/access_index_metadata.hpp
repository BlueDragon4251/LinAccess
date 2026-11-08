#pragma once
#include <QByteArray>
#include <QStringList>
#include <QVector>
namespace linaccess {
struct AccessSourceIndex {
    QString name;
    int kind{},flags{};
    QStringList columns;
    QVector<bool> descending;
    bool operator==(const AccessSourceIndex&) const = default;
};
// Strict versioned metadata produced by the read-only native MDB helper.
// Caller supplies the actual ordered column names; no index names are inferred.
QVector<AccessSourceIndex> checkedAccessSourceIndexes(const QByteArray& json,const QStringList& columns);
}
