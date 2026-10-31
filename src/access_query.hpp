#pragma once
#include <QMap>
#include <QSqlDatabase>
#include <QString>
#include <QVariant>
#include <QVector>

namespace linaccess {
struct AccessQueryParameter {
    QString name;
    QString type;
    QString binding;
    bool likePattern{};
    bool direct{};
};
struct AccessQueryPlan {
    QString sql;
    QVector<AccessQueryParameter> parameters;
    int topLimit{};
    bool hiddenRank{};
};
// Jet ANSI-89 strings, named PARAMETERS and basic LIKE patterns. Unsupported
// grammar remains an explicit error; this is not a complete Jet SQL translator.
AccessQueryPlan compileAccessQuery(const QString& source);
int installAccessQueryViews(QSqlDatabase db, bool freshConnection = false);
QString accessQuerySource(QSqlDatabase db, const QString& name);
QMap<QString,QVariant> bindAccessQueryParameters(const AccessQueryPlan& plan,
                                                const QMap<QString,QVariant>& values);
}
