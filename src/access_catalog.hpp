#pragma once

#include <QByteArray>
#include <QString>
#include <QStringList>
#include <QVector>

namespace linaccess {
struct AccessCatalogEntry {
    QString kind;
    QString name;
    bool internal{};
};

struct AccessRelation {
    QString childTable;
    QString childColumn;
    QString parentTable;
    QString parentColumn;
};

// Only reads the Access file. No VBA is loaded or executed, and no file is altered.
class AccessCatalog {
public:
    explicit AccessCatalog(QString file);
    QString formatVersion() const;
    QVector<AccessCatalogEntry> objects() const;
    QVector<AccessRelation> relationships() const;
    QString querySql(const QString& name) const;
    QStringList missingQuerySources(const QString& sql) const;
    QString tableSchema(const QString& name) const;
    static QByteArray tool(const QString& executable, const QStringList& args,
                           const QByteArray& input = {});
private:
    QString file_;
};
}
