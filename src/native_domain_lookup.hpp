#pragma once
#include "access_query.hpp"
#include "native_report_aggregate.hpp"
#include "vba_qt_value.hpp"
#include "vba_runtime.hpp"
#include <QSqlDatabase>
#include <QStringList>

namespace linaccess {
// Existing SQL-expression Sum/Count path; preserve proven scalar result types.
inline VbaValue nativeSqlDomainAggregate(QSqlDatabase database,std::string_view function,const std::vector<VbaValue>& args) {
    const auto quoted=[](QString name) { name.replace(']',"]]"); return '['+name+']'; };
    if (args.size()<2 || args.size()>3) throw linaccess::VbaRuntimeError(450,"Invalid domain aggregate arguments");
    if (args[0].kind==linaccess::VbaValue::Kind::Null || args[1].kind==linaccess::VbaValue::Kind::Null)
        throw linaccess::VbaRuntimeError(94,"Invalid Null domain aggregate argument");
    const auto expression=QString::fromStdString(args[0].asText()).trimmed();
    const auto domain=QString::fromStdString(args[1].asText());
    if (expression.isEmpty() || domain.isEmpty()) throw linaccess::VbaRuntimeError(5,"Empty domain aggregate argument");
    // Domain is an object name, never executable SQL. Saved non-parameter
    // queries are installed as views on the same native connection.
    linaccess::installAccessQueryViews(database);
    if (!database.tables(QSql::AllTables).contains(domain,Qt::CaseInsensitive))
        throw linaccess::VbaRuntimeError(3078,"Domain object unavailable");
    QString source="SELECT "+QString(function=="dsum" ? "Sum" : "Count")+"("+expression+") FROM "+quoted(domain);
    if (args.size()==3 && args[2].kind!=linaccess::VbaValue::Kind::Null) {
        const auto criteria=QString::fromStdString(args[2].asText()).trimmed();
        if (!criteria.isEmpty()) source+=" WHERE "+criteria;
    }
    const auto plan=linaccess::compileAccessQuery(source);
    try {
        const auto result=linaccess::executeAccessQueryPlan(database,plan,linaccess::bindAccessQueryParameters(plan,{}));
        if (result.rows.size()!=1 || result.rows.front().size()!=1)
            throw linaccess::VbaRuntimeError(3146,"Invalid domain aggregate result");
        return linaccess::qtFieldValue(result.rows.front().front(),result.columnTypes.value(0));
    } catch(const linaccess::VbaRuntimeError&) { throw; }
    catch(...) { throw linaccess::VbaRuntimeError(3146,"Domain aggregate execution failed"); }
}

// Field-only typed lookup/extrema/mean. Further expression binding is refused.
inline VbaValue nativeDomainValue(QSqlDatabase database,const std::vector<VbaValue>& args,std::string_view function="dlookup") {
    if(function=="dsum" || function=="dcount") return nativeSqlDomainAggregate(database,function,args);
    const bool lookup=function=="dlookup" || function=="dfirst";
    const bool selector=lookup || function=="dlast";
    if(!selector && function!="dmin" && function!="dmax" && function!="davg")
        throw VbaRuntimeError(3251,"Domain function unsupported");
    if(args.size()<2 || args.size()>3) throw VbaRuntimeError(450,"Domain function requires two or three arguments");
    if(args[0].kind==VbaValue::Kind::Null || args[1].kind==VbaValue::Kind::Null)
        throw VbaRuntimeError(94,"Invalid Null Domain function argument");
    if(args[0].kind!=VbaValue::Kind::Text || args[1].kind!=VbaValue::Kind::Text)
        throw VbaRuntimeError(13,"Domain function field and domain require strings");
    auto field=QString::fromStdString(args[0].text).trimmed();
    const auto requested=QString::fromStdString(args[1].text);
    if(field.isEmpty() || requested.isEmpty() || field.contains(QChar(0)) || requested.contains(QChar(0)))
        throw VbaRuntimeError(5,"Invalid Domain function field/domain");
    QString criteria;
    if(args.size()==3 && args[2].kind!=VbaValue::Kind::Missing && args[2].kind!=VbaValue::Kind::Empty) {
        if(args[2].kind==VbaValue::Kind::Null) throw VbaRuntimeError(3251,"Explicit Null Domain function criteria requires reference validation");
        if(args[2].kind!=VbaValue::Kind::Text) throw VbaRuntimeError(13,"Domain function criteria requires a string");
        criteria=QString::fromStdString(args[2].text).trimmed();
        if(criteria.contains(QChar(0))) throw VbaRuntimeError(5,"Domain function criteria contains NUL");
    }
    if(!database.isOpen()) throw VbaRuntimeError(3146,"Domain function project database unavailable");
    (void)installAccessQueryViews(database);
    QString domain;
    for(const auto& candidate:database.tables(QSql::AllTables))
        if(candidate.compare(requested,Qt::CaseInsensitive)==0) {
            if(!domain.isEmpty() && domain!=candidate) throw VbaRuntimeError(5,"Domain function domain name ambiguous");
            domain=candidate;
        }
    if(domain.isEmpty()) throw VbaRuntimeError(3078,"Domain function domain unavailable");
    const auto quoted=[](QString name) { name.replace(']',"]]"); return '['+name+']'; };
    if(field.startsWith('[') && field.endsWith(']')) {
        const auto inner=field.mid(1,field.size()-2); QString decoded;
        for(qsizetype index=0;index<inner.size();++index) {
            if(inner[index]==']') {
                if(index+1>=inner.size() || inner[index+1]!=']') throw VbaRuntimeError(5,"Domain function field delimiter invalid");
                ++index;
            }
            decoded+=inner[index];
        }
        field=decoded;
    }
    const auto columns=accessQueryColumns(database,compileAccessQuery("SELECT * FROM "+quoted(domain)));
    QString canonical;
    for(const auto& column:columns) if(column.compare(field,Qt::CaseInsensitive)==0) {
        if(!canonical.isEmpty()) throw VbaRuntimeError(5,"Domain function field ambiguous");
        canonical=column;
    }
    if(canonical.isEmpty()) throw VbaRuntimeError(3251,"Domain function requires an existing field; calculated expressions are unsupported");
    QString source=QString(lookup ? "SELECT TOP 1 " : "SELECT ")+quoted(canonical)+" FROM "+quoted(domain);
    if(!criteria.isEmpty()) source+=" WHERE "+criteria;
    const auto plan=compileAccessQuery(source);
    if(!plan.makeTable.isEmpty() || plan.hiddenRank || !plan.sql.trimmed().startsWith("SELECT",Qt::CaseInsensitive))
        throw VbaRuntimeError(3251,"Domain function requires a bounded read-only field selection");
    if(!selector) {
        QVector<QString> types;
        const auto selected=accessQueryColumns(database,plan,&types);
        const auto type=types.value(0).toUpper();
        const bool numeric=type.contains("INT") || type.contains("REAL") || type.contains("FLOA") || type.contains("DOUB");
        const bool typedExtremum=function!="davg" && (type=="CURRENCY_REAL" || type.contains("DATE"));
        if(selected.size()!=1 || type.contains("BOOL") || (!numeric && !typedExtremum)
            || (function=="davg" && type=="CURRENCY_REAL"))
            throw VbaRuntimeError(3251,"Domain aggregate field type unsupported or unproven");
        const auto result=executeAccessQueryPlan(database,plan,bindAccessQueryParameters(plan,{}));
        if(result.columns.size()!=1 || result.columnTypes.value(0).toUpper()!=type)
            throw VbaRuntimeError(3146,"Domain aggregate result type changed");
        NativeReportAggregate aggregate(function=="davg" ? "avg" : function=="dmin" ? "min" : "max");
        for(const auto& row:result.rows) {
            if(row.size()!=1) throw VbaRuntimeError(3146,"Domain aggregate result shape invalid");
            aggregate.add(qtFieldValue(row.front(),type));
        }
        return aggregate.result();
    }
    const auto result=executeAccessQueryPlan(database,plan,bindAccessQueryParameters(plan,{}));
    if(result.columns.size()!=1 || (lookup && result.rows.size()>1)
        || (!result.rows.isEmpty() && (function=="dlast" ? result.rows.back() : result.rows.front()).size()!=1))
        throw VbaRuntimeError(3146,"Domain function result shape invalid");
    if(result.rows.isEmpty()) return {};
    const auto& selected=function=="dlast" ? result.rows.back() : result.rows.front();
    if(selected.front().isNull()) return {};
    const auto type=result.columnTypes.value(0);
    if(type.isEmpty()) throw VbaRuntimeError(3251,"Domain function field type is not proven");
    return qtFieldValue(selected.front(),type);
}
}
