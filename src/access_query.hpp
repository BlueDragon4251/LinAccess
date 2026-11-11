#pragma once
#include <QAbstractTableModel>
#include <QStringList>
#include <QMap>
#include <QSqlDatabase>
#include <QString>
#include <QVariant>
#include <QVector>
#include <QPair>

namespace linaccess {
// Strict native-token comparison for guarded schema reconstruction; no SQL execution.
bool equivalentNativeSchemaSql(const QString& left,const QString& right);
bool nativeSchemaUsesIndexHints(const QString& sql);
struct AccessQueryParameter {
    QString name;
    QString type;
    QString binding;
    bool likePattern{};
    bool direct{};
    bool currencyTicks{}; // Strict table-free bare Currency-parameter projections only.
};
struct AccessCounterDeclaration {
    QString table,column;
    qint32 seed{1},increment{1};
    bool reseed{};
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
    QVector<AccessCounterDeclaration> counters;
    QByteArray currencyParameterRootHash; // Original pure binding SQL, retained across wrappers.
    bool currencyParameterDerived{}; // Checked wrapper retaining a pure full-range Currency binding source.
    bool deferredDecimalAction{}; // Must pass a proven physical action adapter before execution.
    int daoQueryType{-1}; // Native DAO source kind, unknown until classified by the compiler.
};
struct AccessQueryResult {
    QStringList columns;
    QVector<QVariantList> rows;
    QMap<QString,QString> types; // Only proven, unambiguous output-column types.
    QVector<QString> columnTypes; // Per-position proven UI types; preserves the DAO type contract.
    QVector<QPair<QString,QString>> fieldOrigins; // Proven main-schema physical table/field; absent entries are unknown.
};
AccessQueryResult executeAccessQueryPlan(QSqlDatabase db, const AccessQueryPlan& plan,
                                       const QMap<QString,QVariant>& bindings);
// Schema only: prepare without stepping records or evaluating user expressions.
QStringList accessQueryColumns(QSqlDatabase db,const AccessQueryPlan& plan,QVector<QString>* columnTypes=nullptr,QVector<QPair<QString,QString>>* fieldOrigins=nullptr);
class AccessQueryModel final : public QAbstractTableModel {
public:
    explicit AccessQueryModel(QObject* parent = nullptr) : QAbstractTableModel(parent) {}
    void execute(QSqlDatabase db, const QString& source, const QMap<QString,QVariant>& parameters = {});
    void executeFiltered(QSqlDatabase db,const QString& source,const QMap<QString,QVariant>& parameters,const QString& criteria);
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
// Checked single predicate compiled through the Jet adapter, without extra clauses.
QString accessWhereClause(const QString& criteria);
// Bounded scalar domain expression; no clauses, aggregates or unregistered calls.
QString accessScalarDomainExpression(const QString& expression,bool deterministicOnly=false);
AccessQueryPlan compileAccessQuery(const QString& source);
// Bind an already-known native recordset schema; no result-value type inference.
AccessQueryPlan compileAccessFindCriteria(const QString& criteria,const QMap<QString,QString>& fieldTypes);
// Resolves compiled action metadata against the live writable destination.
QString accessActionSql(QSqlDatabase db,const AccessQueryPlan& plan);
// Recover the SELECT source through SQL tokens, preserving parameter declarations.
AccessQueryPlan accessMakeTableSource(QSqlDatabase db,const AccessQueryPlan& plan);
AccessQueryPlan filterAccessQueryPlan(QSqlDatabase db,const AccessQueryPlan& source,const QString& criteria);
QString accessOrderByClause(const QString& orderBy,const QStringList& columns,const QString& tableQualifier={});
AccessQueryPlan sortAccessQueryPlan(QSqlDatabase db,const AccessQueryPlan& source,const QString& orderBy);
int installAccessQueryViews(QSqlDatabase db, bool freshConnection = false);
QString accessQuerySource(QSqlDatabase db, const QString& name,QString* canonical=nullptr,bool* owned=nullptr,bool allowEmpty=false);
// Editable keyed single-table sources. Explicit field projections require an output mapping.
struct AccessQueryColumn {int column;QString alias;};
QString editableAccessQueryTable(QSqlDatabase db,const QString& source,QString* filter=nullptr,QVector<QPair<int,Qt::SortOrder>>* sort=nullptr,QVector<AccessQueryColumn>* projection=nullptr);
QMap<QString,QVariant> bindAccessQueryParameters(const AccessQueryPlan& plan,
                                                const QMap<QString,QVariant>& values);
}
