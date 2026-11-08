#pragma once

#include <QByteArray>
#include <QString>
#include <QStringList>
#include <QVector>
#include <QMap>

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
    QString name;
    quint32 flags{};
    int ordinal{};
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
    QByteArray singleValues(const QString& name) const;
    QByteArray doubleValues(const QString& name) const;
    // Decode exact Currency/Single/Double labels; storage/coercion remains separate.
    static QMap<QString,QString> floatingSourceTypes(const QString& accessSchema);
    // Explicit unprefixed hexadecimal MDBTools binary serialization when requested.
    static QByteArray exportTable(const QString& source, const QString& table,bool binaryHex=false);
    static QByteArray tool(const QString& executable, const QStringList& args,
                           const QByteArray& input = {});
private:
    QString file_;
};
}
