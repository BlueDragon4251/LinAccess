#pragma once
#include "access_query.hpp"
#include "vba_qt_value.hpp"
#include "vba_runtime.hpp"
#include <QSqlDatabase>
#include <QStringList>

namespace linaccess {
// Field-only, typed lookup. Expressions needing further binding are refused.
inline VbaValue nativeDomainLookup(QSqlDatabase database,const std::vector<VbaValue>& args) {
    if(args.size()<2 || args.size()>3) throw VbaRuntimeError(450,"DLookup requires two or three arguments");
    if(args[0].kind==VbaValue::Kind::Null || args[1].kind==VbaValue::Kind::Null)
        throw VbaRuntimeError(94,"Invalid Null DLookup argument");
    if(args[0].kind!=VbaValue::Kind::Text || args[1].kind!=VbaValue::Kind::Text)
        throw VbaRuntimeError(13,"DLookup field and domain require strings");
    auto field=QString::fromStdString(args[0].text).trimmed();
    const auto requested=QString::fromStdString(args[1].text);
    if(field.isEmpty() || requested.isEmpty() || field.contains(QChar(0)) || requested.contains(QChar(0)))
        throw VbaRuntimeError(5,"Invalid DLookup field/domain");
    QString criteria;
    if(args.size()==3 && args[2].kind!=VbaValue::Kind::Missing && args[2].kind!=VbaValue::Kind::Empty) {
        if(args[2].kind==VbaValue::Kind::Null) throw VbaRuntimeError(3251,"Explicit Null DLookup criteria requires reference validation");
        if(args[2].kind!=VbaValue::Kind::Text) throw VbaRuntimeError(13,"DLookup criteria requires a string");
        criteria=QString::fromStdString(args[2].text).trimmed();
        if(criteria.contains(QChar(0))) throw VbaRuntimeError(5,"DLookup criteria contains NUL");
    }
    if(!database.isOpen()) throw VbaRuntimeError(3146,"DLookup project database unavailable");
    (void)installAccessQueryViews(database);
    QString domain;
    for(const auto& candidate:database.tables(QSql::AllTables))
        if(candidate.compare(requested,Qt::CaseInsensitive)==0) {
            if(!domain.isEmpty() && domain!=candidate) throw VbaRuntimeError(5,"DLookup domain name ambiguous");
            domain=candidate;
        }
    if(domain.isEmpty()) throw VbaRuntimeError(3078,"DLookup domain unavailable");
    const auto quoted=[](QString name) { name.replace(']',"]]"); return '['+name+']'; };
    if(field.startsWith('[') && field.endsWith(']')) {
        const auto inner=field.mid(1,field.size()-2); QString decoded;
        for(qsizetype index=0;index<inner.size();++index) {
            if(inner[index]==']') {
                if(index+1>=inner.size() || inner[index+1]!=']') throw VbaRuntimeError(5,"DLookup field delimiter invalid");
                ++index;
            }
            decoded+=inner[index];
        }
        field=decoded;
    }
    const auto columns=accessQueryColumns(database,compileAccessQuery("SELECT * FROM "+quoted(domain)));
    QString canonical;
    for(const auto& column:columns) if(column.compare(field,Qt::CaseInsensitive)==0) {
        if(!canonical.isEmpty()) throw VbaRuntimeError(5,"DLookup field ambiguous");
        canonical=column;
    }
    if(canonical.isEmpty()) throw VbaRuntimeError(3251,"DLookup requires an existing field; calculated expressions are unsupported");
    QString source="SELECT TOP 1 "+quoted(canonical)+" FROM "+quoted(domain);
    if(!criteria.isEmpty()) source+=" WHERE "+criteria;
    const auto plan=compileAccessQuery(source);
    if(!plan.makeTable.isEmpty() || plan.hiddenRank || !plan.sql.trimmed().startsWith("SELECT",Qt::CaseInsensitive))
        throw VbaRuntimeError(3251,"DLookup requires a bounded read-only field selection");
    const auto result=executeAccessQueryPlan(database,plan,bindAccessQueryParameters(plan,{}));
    if(result.columns.size()!=1 || result.rows.size()>1 || (!result.rows.isEmpty() && result.rows.front().size()!=1))
        throw VbaRuntimeError(3146,"DLookup result shape invalid");
    if(result.rows.isEmpty() || result.rows.front().front().isNull()) return {};
    const auto type=result.columnTypes.value(0);
    if(type.isEmpty()) throw VbaRuntimeError(3251,"DLookup field type is not proven");
    return qtFieldValue(result.rows.front().front(),type);
}
}
