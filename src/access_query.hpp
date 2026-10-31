#pragma once
#include <QAbstractTableModel>
#include <QStringList>
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
    QString makeTable;
    QVector<AccessQueryParameter> parameters;
    int topLimit{};
    bool hiddenRank{};
    QString topOneSql;
    QString topOneTable;
    QString topOneColumn;
    QString rightJoinSql;
    QStringList rightJoinSources;
};
struct AccessQueryResult {
    QStringList columns;
    QVector<QVariantList> rows;
};
AccessQueryResult executeAccessQueryPlan(QSqlDatabase db, const AccessQueryPlan& plan,
                                       const QMap<QString,QVariant>& bindings);
class AccessQueryModel final : public QAbstractTableModel {
public:
    explicit AccessQueryModel(QObject* parent = nullptr) : QAbstractTableModel(parent) {}
    void execute(QSqlDatabase db, const QString& source, const QMap<QString,QVariant>& parameters = {});
    int rowCount(const QModelIndex& parent = {}) const override { return parent.isValid() ? 0 : result_.rows.size(); }
    int columnCount(const QModelIndex& parent = {}) const override { return parent.isValid() ? 0 : result_.columns.size(); }
    QVariant data(const QModelIndex& index, int role = Qt::DisplayRole) const override;
    QVariant headerData(int section, Qt::Orientation orientation, int role = Qt::DisplayRole) const override;
private:
    AccessQueryResult result_;
};
// Jet ANSI-89 strings, named PARAMETERS and basic LIKE patterns. Unsupported
// grammar remains an explicit error; this is not a complete Jet SQL translator.
QString accessLikePattern(const QString& value);
AccessQueryPlan compileAccessQuery(const QString& source);
AccessQueryPlan filterAccessQueryPlan(QSqlDatabase db,const AccessQueryPlan& source,const QString& criteria);
int installAccessQueryViews(QSqlDatabase db, bool freshConnection = false);
QString accessQuerySource(QSqlDatabase db, const QString& name);
QMap<QString,QVariant> bindAccessQueryParameters(const AccessQueryPlan& plan,
                                                const QMap<QString,QVariant>& values);
}
