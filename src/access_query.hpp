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
    QString updateTarget;
    QString updateIdentity;
    QString updateSource;
    qsizetype updateSourceOffset{};
    QVector<QString> resultTypes; // Proven projection types; empty entries are unknown.
    QString source;
    QString compiledSql;
    bool currencyContextResolved{};
};
struct AccessQueryResult {
    QStringList columns;
    QVector<QVariantList> rows;
    QMap<QString,QString> types; // Only proven, unambiguous output-column types.
    QVector<QString> columnTypes; // Per-position proven UI types; preserves the DAO type contract.
};
AccessQueryResult executeAccessQueryPlan(QSqlDatabase db, const AccessQueryPlan& plan,
                                       const QMap<QString,QVariant>& bindings);
// Schema only: prepare without stepping records or evaluating user expressions.
QStringList accessQueryColumns(QSqlDatabase db,const AccessQueryPlan& plan);
class AccessQueryModel final : public QAbstractTableModel {
public:
    explicit AccessQueryModel(QObject* parent = nullptr) : QAbstractTableModel(parent) {}
    void execute(QSqlDatabase db, const QString& source, const QMap<QString,QVariant>& parameters = {});
    int rowCount(const QModelIndex& parent = {}) const override { return parent.isValid() ? 0 : result_.rows.size(); }
    int columnCount(const QModelIndex& parent = {}) const override { return parent.isValid() ? 0 : result_.columns.size(); }
    QVariant data(const QModelIndex& index, int role = Qt::DisplayRole) const override;
    QString clipboardText(const QModelIndex& index) const;
    QString fieldSqlType(int column) const;
    QVariant headerData(int section, Qt::Orientation orientation, int role = Qt::DisplayRole) const override;
private:
    AccessQueryResult result_;
};
// Jet ANSI-89 strings, named PARAMETERS and basic LIKE patterns. Unsupported
// grammar remains an explicit error; this is not a complete Jet SQL translator.
QString accessLikePattern(const QString& value);
AccessQueryPlan compileAccessQuery(const QString& source);
// Resolves compiled action metadata against the live writable destination.
QString accessActionSql(QSqlDatabase db,const AccessQueryPlan& plan);
AccessQueryPlan filterAccessQueryPlan(QSqlDatabase db,const AccessQueryPlan& source,const QString& criteria);
QString accessOrderByClause(const QString& orderBy,const QStringList& columns,const QString& tableQualifier={});
AccessQueryPlan sortAccessQueryPlan(QSqlDatabase db,const AccessQueryPlan& source,const QString& orderBy);
int installAccessQueryViews(QSqlDatabase db, bool freshConnection = false);
QString accessQuerySource(QSqlDatabase db, const QString& name);
// Strict editable subset: an unparameterized SELECT * over one keyed main table.
QString editableAccessQueryTable(QSqlDatabase db,const QString& source);
QMap<QString,QVariant> bindAccessQueryParameters(const AccessQueryPlan& plan,
                                                const QMap<QString,QVariant>& values);
}
