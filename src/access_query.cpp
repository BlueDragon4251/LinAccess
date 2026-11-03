#include "access_query.hpp"
#include <functional>
#include <algorithm>
#include "vba_qt_value.hpp"
#include "jet_sql_functions.hpp"
#include "sqlite_execution.hpp"
#include <QSqlError>
#include <QSqlRecord>
#include "vba_runtime.hpp"
#include "vba_source_codec.hpp"
#include <QSqlQuery>
#include <QSqlDriver>
#include <QSqlResult>
#include <QScopeGuard>
#include <sqlite3.h>
#include <QLibrary>
struct sqlite3;
struct sqlite3_stmt;
#include <QJsonDocument>
#include <QCryptographicHash>
#include <QSet>
#include <QDateTime>
#include <cmath>
#include <limits>

namespace linaccess {
namespace {
struct Token { enum Kind { Word, Identifier, String, Other } kind; QString value; };
void require(bool value, const char* reason) { if (!value) throw VbaRuntimeError(3085, reason); }
QString canonicalDate(const QString& input, bool literalDate = false) {
    const QStringList formats = literalDate
        ? QStringList{"M/d/yyyy H:mm:ss","M/d/yyyy H:mm","M/d/yyyy h:mm:ss AP","yyyy-MM-dd HH:mm:ss","yyyy-MM-dd" ,"M/d/yyyy"}
        : QStringList{"yyyy-MM-dd HH:mm:ss","yyyy-MM-dd"};
    for (const auto& format : formats) {
        const auto date = QDateTime::fromString(input.trimmed(),format);
        if (date.isValid() && date.date().year() >= 100)
            return date.toString("yyyy-MM-dd HH:mm:ss");
    }
    throw VbaRuntimeError(13,"Jet SQL: unambiguous four-digit-year date required");
}
QString literal(QString value) { return '\'' + value.replace('\'', "''") + '\''; }
QString identifier(QString value) { return '"' + value.replace('"', "\"\"") + '"'; }
QVector<Token> tokenize(const QString& sql,bool nativeIdentifiers=false) {
    QVector<Token> result;
    for (qsizetype i = 0; i < sql.size();) {
        const auto ch = sql[i];
        if (ch.isSpace()) { ++i; continue; }
        if (ch == '-' && i + 1 < sql.size() && sql[i + 1] == '-') {
            while (i < sql.size() && sql[i] != '\n') ++i;
            continue;
        }
        if (ch == '/' && i + 1 < sql.size() && sql[i + 1] == '*') {
            const auto end = sql.indexOf("*/",i+2);
            require(end >= 0,"Jet SQL: unterminated comment"); i = end + 2; continue;
        }
        if (ch == '#') {
            const auto end = sql.indexOf('#',i+1);
            require(end >= 0,"Jet SQL: unterminated date literal");
            result.append({Token::String,canonicalDate(sql.mid(i+1,end-i-1),true)});
            i = end + 1; continue;
        }
        if (nativeIdentifiers && ch==':') {
            const auto start=i++;
            while (i<sql.size() && (sql[i].isLetterOrNumber() || sql[i]=='_')) ++i;
            result.append({Token::Other,sql.mid(start,i-start)}); continue;
        }
        if (ch == '\'' || ch == '"' || ch == '[') {
            if (ch=='[') {
                auto begin=i+1;
                while (begin<sql.size() && sql[begin].isSpace()) ++begin;
                if (sql.mid(begin,6).compare("SELECT",Qt::CaseInsensitive)==0
                    && begin+6<sql.size() && sql[begin+6].isSpace()) {
                    // Jet exports derived tables as [SELECT ...]. AS alias.
                    // Nested identifier brackets belong to the inner query.
                    int depth=1; auto end=i+1; QChar quote;
                    for (;end<sql.size();++end) {
                        const auto current=sql[end];
                        if (!quote.isNull()) {
                            if (current==quote) {
                                if (end+1<sql.size() && sql[end+1]==quote) ++end;
                                else quote={};
                            }
                        } else if (current=='\'' || current=='"') quote=current;
                        else if (current=='[') ++depth;
                        else if (current==']' && --depth==0) break;
                    }
                    auto dot=end+1;
                    while (dot<sql.size() && sql[dot].isSpace()) ++dot;
                    auto as=dot+1;
                    while (as<sql.size() && sql[as].isSpace()) ++as;
                    if (depth==0 && dot<sql.size() && sql[dot]=='.'
                        && sql.mid(as,2).compare("AS",Qt::CaseInsensitive)==0
                        && as+2<sql.size() && sql[as+2].isSpace()) {
                        static thread_local int nesting=0;
                        require(nesting<16,"Jet SQL: bracketed subquery nesting limit");
                        struct Guard { int& depth; explicit Guard(int& value):depth(value) { ++depth; } ~Guard() { --depth; } } guard(nesting);
                        const auto inner=compileAccessQuery(sql.mid(i+1,end-i-1));
                        require(inner.parameters.empty() && inner.topLimit==0,"Jet SQL: bracketed subquery parameter/TOP materialization unsupported");
                        require(inner.sql.trimmed().startsWith("SELECT",Qt::CaseInsensitive),"Jet SQL: bracketed subquery requires SELECT");
                        result.append({Token::Other,"("+inner.sql+")"});
                        i=dot+1;
                        continue;
                    }
                }
            }
            const auto close = ch == '[' ? QChar(']') : ch;
            QString value; bool closed = false; ++i;
            while (i < sql.size()) {
                if (sql[i] == close) {
                    ++i;
                    if (i < sql.size() && sql[i] == close) { value += close; ++i; }
                    else { closed = true; break; }
                } else value += sql[i++];
            }
            require(closed, "Jet SQL: unterminated string or identifier");
            result.append({ch == '[' || (nativeIdentifiers && ch=='"') ? Token::Identifier : Token::String, value});
        } else if (ch.isLetter() || ch == '_') {
            const auto start = i++;
            while (i < sql.size() && (sql[i].isLetterOrNumber() || sql[i] == '_')) ++i;
            if (i<sql.size() && sql[i]=='$') ++i;
            result.append({Token::Word,sql.mid(start,i-start)});
        } else if (ch.isDigit() || (ch=='.' && i+1<sql.size() && sql[i+1].isDigit())) {
            const auto start = i++;
            while (i < sql.size() && (sql[i].isDigit() || sql[i] == '.')) ++i;
            if (i < sql.size() && (sql[i] == 'e' || sql[i] == 'E')) {
                ++i; if (i < sql.size() && (sql[i] == '+' || sql[i] == '-')) ++i;
                while (i < sql.size() && sql[i].isDigit()) ++i;
            }
            result.append({Token::Other,sql.mid(start,i-start)});
        } else {
            QString op(ch); ++i;
            if (i < sql.size() && ((ch == '<' && (sql[i] == '>' || sql[i] == '=')) || (ch == '>' && sql[i] == '='))) op += sql[i++];
            result.append({Token::Other,op});
        }
    }
    return result;
}
// Wrap concatenation operands before SQL emission. Arithmetic stays inside each
// operand; comparison, Boolean and clause boundaries remain outside. Parentheses
// are scanned structurally, including nested function arguments and subqueries.
bool symbol(const Token& token,const char* value) { return token.kind==Token::Other && token.value==value; }
bool expressionBoundary(const QVector<Token>& tokens,qsizetype index) {
    const auto& token=tokens[index];
    if (token.kind==Token::Other) {
        static const QSet<QString> symbols{",",";","&","=","<",">","<=",">=","<>"};
        return symbols.contains(token.value);
    }
    if (token.kind!=Token::Word) return false;
    // Reserved words used as function names are atomic with their arguments.
    if ((token.value.compare("Left",Qt::CaseInsensitive)==0 || token.value.compare("Right",Qt::CaseInsensitive)==0)
        && index+1<tokens.size() && symbol(tokens[index+1],"(")) return false;
    static const QSet<QString> words{"select","from","where","group","order","having","by","as","on",
        "join","left","right","inner","outer","full","cross","union","except","intersect","limit","offset",
        "when","then","else","end","case","set","values","into","update","delete","insert","top","distinct",
        "distinctrow","parameters","pivot","transform","asc","desc","and","or","xor","eqv","imp","not","is","like","in","between"};
    return words.contains(token.value.toCaseFolded());
}
qsizetype operandEnd(const QVector<Token>& tokens,qsizetype start) {
    int depth=0;
    auto end=start;
    while (end<tokens.size()) {
        if (symbol(tokens[end],"(")) ++depth;
        else if (symbol(tokens[end],")")) { if (depth==0) break; --depth; }
        else if (depth==0 && expressionBoundary(tokens,end)) break;
        ++end;
    }
    return end;
}
void replaceTokens(QVector<Token>& tokens,qsizetype first,qsizetype last,const QVector<Token>& replacement) {
    tokens.erase(tokens.begin()+first,tokens.begin()+last);
    for (qsizetype i=replacement.size();i>0;--i) tokens.insert(first,replacement[i-1]);
}
void adaptConcatenation(QVector<Token>& tokens,qsizetype start) {
    int count=0;
    while (true) {
        qsizetype amp=start;
        while (amp<tokens.size() && !symbol(tokens[amp],"&")) ++amp;
        if (amp==tokens.size()) break;
        require(++count<=128,"Jet SQL: concatenation expression limit exceeded");
        qsizetype first=amp;
        int depth=0;
        for (qsizetype i=amp-1;i>=start;--i) {
            if (symbol(tokens[i],")")) ++depth;
            else if (symbol(tokens[i],"(")) { if (depth==0) break; --depth; }
            else if (depth==0 && expressionBoundary(tokens,i)) break;
            first=i;
        }
        const auto last=operandEnd(tokens,amp+1);
        require(first<amp && last>amp+1,"Jet SQL: concatenation operand missing");
        QVector<Token> replacement{{Token::Word,"linaccess_jet_concat"},{Token::Other,"("}};
        replacement+=tokens.mid(first,amp-first);
        replacement.append({Token::Other,","});
        replacement+=tokens.mid(amp+1,last-amp-1);
        replacement.append({Token::Other,")"});
        replaceTokens(tokens,first,last,replacement);
    }
}
void adaptDynamicLike(QVector<Token>& tokens,qsizetype start,const QMap<QString,int>& parameters) {
    for (qsizetype i=tokens.size()-1;i>=start;--i) {
        if (tokens[i].kind!=Token::Word || tokens[i].value.compare("LIKE",Qt::CaseInsensitive)!=0) continue;
        const auto last=operandEnd(tokens,i+1);
        require(last>i+1,"Jet SQL: LIKE operand missing");
        const auto& first=tokens[i+1];
        const bool simple=last==i+2 && (first.kind==Token::String ||
            ((first.kind==Token::Identifier || first.kind==Token::Word) && parameters.contains(first.value.toCaseFolded())));
        if (simple) continue;
        QVector<Token> replacement{{Token::Word,"linaccess_jet_like_pattern"},{Token::Other,"("}};
        replacement+=tokens.mid(i+1,last-i-1);
        replacement.append({Token::Other,")"});
        replacement.append({Token::Word,"ESCAPE"});
        replacement.append({Token::String,"\\"});
        replaceTokens(tokens,i+1,last,replacement);
    }
}
void adaptObjectReferences(QVector<Token>& tokens,qsizetype start) {
    const auto name=[](const Token& token) { return token.kind==Token::Word || token.kind==Token::Identifier; };
    for (qsizetype i=start;i+4<tokens.size();++i) {
        if (!name(tokens[i])) continue;
        const auto root=tokens[i].value.toCaseFolded();
        const bool forms=root=="forms" || root=="formulare";
        const bool reports=root=="reports" || root=="berichte";
        if ((!forms && !reports) || !symbol(tokens[i+1],"!") || !name(tokens[i+2]) || !symbol(tokens[i+3],"!") || !name(tokens[i+4])) continue;
        require(i==start || (!symbol(tokens[i-1],".") && !symbol(tokens[i-1],"!")),"Jet SQL: qualified object context unsupported");
        auto last=i+5;
        if (last+1<tokens.size() && symbol(tokens[last],".") && name(tokens[last+1]) && tokens[last+1].value.compare("Value",Qt::CaseInsensitive)==0) last+=2;
        require(last==tokens.size() || (!symbol(tokens[last],".") && !symbol(tokens[last],"!") && !symbol(tokens[last],"(")),"Jet SQL: nested object reference unsupported");
        const QVector<Token> replacement{{Token::Word,"linaccess_jet_object_value"},{Token::Other,"("},
            {Token::String,forms ? "forms" : "reports"},{Token::Other,","},{Token::String,tokens[i+2].value},
            {Token::Other,","},{Token::String,tokens[i+4].value},{Token::Other,")"}};
        replaceTokens(tokens,i,last,replacement);
        i+=replacement.size()-1;
    }
}
QVector<QString> currencyProjectionTypes(QVector<Token>& tokens,qsizetype start,const QVector<AccessQueryParameter>& parameters,const QMap<QString,QString>& fields) {
    const auto word=[&](qsizetype i,const char* text) { return i<tokens.size() && tokens[i].kind==Token::Word && tokens[i].value.compare(text,Qt::CaseInsensitive)==0; };
    if(!word(start,"SELECT")) return {};
    int depth=0;
    for(qsizetype i=start+1;i<tokens.size();++i) {
        if(symbol(tokens[i],"(")) ++depth;
        else if(symbol(tokens[i],")")) --depth;
        else if(depth==0 && (word(i,"UNION") || word(i,"INTERSECT") || word(i,"EXCEPT"))) return {};
    }
    const auto name=[](const Token& value) { return value.kind==Token::Word || value.kind==Token::Identifier; };
    const auto precedence=[](const Token& token) {
        const auto value=token.value.toCaseFolded();
        if(token.kind!=Token::Other && token.kind!=Token::Word) return 0;
        if(value=="or" || value=="xor" || value=="eqv" || value=="imp") return 1;
        if(value=="and") return 2;
        if(value=="=" || value=="<>" || value=="!=" || value=="<" || value==">" || value=="<=" || value==">=" || value=="like" || value=="is" || value=="in") return 3;
        if(value=="&" || value=="||") return 4;
        if(value=="+" || value=="-") return 5;
        if(value=="*" || value=="/" || value=="\\" || value=="%" || value=="mod") return 6;
        if(value=="^") return 7;
        return 0;
    };
    struct Expression { QString type; QVector<Token> code; };
    const auto call=[](const QString& function,const QVector<Token>& first,const QVector<Token>& second=QVector<Token>{}) {
        QVector<Token> code{{Token::Word,function},{Token::Other,"("}}; code+=first;
        if(!second.empty()) { code.append({Token::Other,","}); code+=second; }
        code.append({Token::Other,")"}); return code;
    };
    std::function<Expression(QVector<Token>,int)> parse=[&](QVector<Token> expression,int recursion)->Expression {
        const Expression unknown{{},expression};
        if(recursion>64 || expression.isEmpty()) return unknown;
        if(expression.size()>=2 && symbol(expression.front(),"(")) {
            int nesting=0; qsizetype close=-1;
            for(qsizetype i=0;i<expression.size();++i) {
                if(symbol(expression[i],"(")) ++nesting;
                else if(symbol(expression[i],")") && --nesting==0) { close=i; break; }
            }
            if(close==expression.size()-1) {
                auto inner=parse(expression.mid(1,expression.size()-2),recursion+1);
                if(inner.type.isEmpty()) return unknown;
                QVector<Token> code{{Token::Other,"("}}; code+=inner.code; code.append({Token::Other,")"}); return {inner.type,code};
            }
        }
        qsizetype operation=-1; int lowest=100,nesting=0;
        for(qsizetype i=0;i<expression.size();++i) {
            if(symbol(expression[i],"(")) { ++nesting; continue; }
            if(symbol(expression[i],")")) { --nesting; continue; }
            if(nesting) continue;
            const auto rank=precedence(expression[i]); if(!rank) continue;
            if((symbol(expression[i],"+") || symbol(expression[i],"-")) && (i==0 || precedence(expression[i-1]) || symbol(expression[i-1],"(" ) || symbol(expression[i-1],","))) continue;
            if(rank<=lowest) { lowest=rank; operation=i; }
        }
        if(operation>=0) {
            auto left=parse(expression.mid(0,operation),recursion+1),right=parse(expression.mid(operation+1),recursion+1);
            const auto op=expression[operation].value;
            if(left.type!="CURRENCY_REAL" || right.type!="CURRENCY_REAL" || (op!="+" && op!="-" && op!="*")) return unknown;
            return {"CURRENCY_REAL",call(op=="+" ? "linaccess_jet_currency_add" : op=="-" ? "linaccess_jet_currency_subtract" : "linaccess_jet_currency_multiply",left.code,right.code)};
        }
        if(expression.size()>1 && (symbol(expression.front(),"+") || symbol(expression.front(),"-"))) {
            auto inner=parse(expression.mid(1),recursion+1); if(inner.type.isEmpty()) return unknown;
            QVector<Token> code{expression.front(),{Token::Other,"("}}; code+=inner.code; code.append({Token::Other,")"}); return {inner.type,code};
        }
        if(expression.size()==1 && name(expression.front())) {
            for(const auto& parameter:parameters) if(parameter.type=="currency" && parameter.name.compare(expression.front().value,Qt::CaseInsensitive)==0) return {"CURRENCY_REAL",expression};
            if(fields.value(expression.front().value.toCaseFolded())=="CURRENCY_REAL") return {"CURRENCY_REAL",expression};
        }
        if(expression.size()==3 && name(expression.front()) && symbol(expression[1],".") && name(expression.back()) && fields.value((expression.front().value+'.'+expression.back().value).toCaseFolded())=="CURRENCY_REAL") return {"CURRENCY_REAL",expression};
        if(expression.size()<4 || expression.front().kind!=Token::Word || !symbol(expression[1],"(") || !symbol(expression.back(),")")) return unknown;
        const bool currency=expression.front().value.compare("CCur",Qt::CaseInsensitive)==0,sum=expression.front().value.compare("Sum",Qt::CaseInsensitive)==0;
        if(!currency && !sum) return unknown;
        nesting=0;
        for(qsizetype i=1;i<expression.size();++i) {
            if(symbol(expression[i],"(")) ++nesting;
            else if(symbol(expression[i],")")) { if(--nesting==0 && i!=expression.size()-1) return unknown; }
            else if(nesting==1 && symbol(expression[i],",")) return unknown;
        }
        if(nesting) return unknown;
        auto argument=parse(expression.mid(2,expression.size()-3),recursion+1);
        if(sum && argument.type!="CURRENCY_REAL") return unknown;
        return {"CURRENCY_REAL",call(sum ? QString("linaccess_jet_currency_sum") : expression.front().value,argument.code)};
    };
    qsizetype first=start+1;
    if(word(first,"DISTINCT") || word(first,"DISTINCTROW") || word(first,"ALL")) ++first;
    if(word(first,"TOP")) first+=2;
    struct Rewrite { qsizetype first,last; QVector<Token> code; };
    QVector<Rewrite> rewrites; QVector<QString> result; depth=0;
    for(qsizetype i=first;i<=tokens.size();++i) {
        const bool end=i==tokens.size() || (depth==0 && (word(i,"FROM") || word(i,"INTO") || word(i,"WHERE") || word(i,"ORDER") || word(i,"GROUP") || word(i,"HAVING") || symbol(tokens[i],";")));
        if(end || (depth==0 && symbol(tokens[i],","))) {
            auto expression=tokens.mid(first,i-first); QVector<Token> alias;
            if(expression.size()>=3 && expression[expression.size()-2].kind==Token::Word && expression[expression.size()-2].value.compare("AS",Qt::CaseInsensitive)==0 && name(expression.back())) { alias=expression.mid(expression.size()-2); expression.resize(expression.size()-2); }
            if((expression.size()==1 && symbol(expression.front(),"*")) || (expression.size()>=3 && symbol(expression.back(),"*") && symbol(expression[expression.size()-2],"."))) return {};
            auto parsed=parse(expression,0); result.append(parsed.type); parsed.code+=alias; rewrites.append({first,i,parsed.code}); first=i+1;
            if(end) break;
        } else if(symbol(tokens[i],"(")) ++depth;
        else if(symbol(tokens[i],")")) --depth;
    }
    for(auto it=rewrites.crbegin();it!=rewrites.crend();++it) replaceTokens(tokens,it->first,it->last,it->code);
    return result;
}

void preserveQualifiedProjectionNames(QVector<Token>& tokens,qsizetype start) {
    const auto word=[&](qsizetype i,const char* value) {
        return i<tokens.size() && tokens[i].kind==Token::Word && tokens[i].value.compare(value,Qt::CaseInsensitive)==0;
    };
    if (!word(start,"SELECT")) return;
    qsizetype first=start+1;
    if (word(first,"DISTINCT") || word(first,"DISTINCTROW") || word(first,"ALL")) ++first;
    if (word(first,"TOP")) first+=2;
    QVector<QPair<qsizetype,QString>> aliases;
    int depth=0;
    const auto projection=[&](qsizetype end) {
        const auto name=[](const Token& token) { return token.kind==Token::Word || token.kind==Token::Identifier; };
        if (end-first!=3 && end-first!=5) return;
        for (auto i=first;i<end;++i) {
            if ((i-first)%2==0 ? !name(tokens[i]) : !symbol(tokens[i],".")) return;
        }
        aliases.append({end,tokens[end-1].value});
    };
    for (qsizetype i=first;i<=tokens.size();++i) {
        if (i==tokens.size() || (depth==0 && (word(i,"FROM") || word(i,"WHERE") || word(i,"ORDER") || word(i,"UNION") || symbol(tokens[i],";")))) {
            projection(i); break;
        }
        if (depth==0 && symbol(tokens[i],",")) { projection(i); first=i+1; }
        else if (symbol(tokens[i],"(")) ++depth;
        else if (symbol(tokens[i],")")) --depth;
    }
    // Explicit labels prevent SQLite's nested RIGHT JOIN rewrite from leaking
    // private duplicate-column suffixes (e.g. id:1) through saved-query views.
    for (auto i=aliases.size();i>0;--i) {
        const auto& alias=aliases[i-1];
        tokens.insert(alias.first,{Token::Word,"AS"});
        tokens.insert(alias.first+1,{Token::Identifier,alias.second});
    }
}
struct TopOneCandidate { QString source; QString table; QString column; };
QString renderJetTokens(const QVector<Token>& tokens) {
    QStringList text;
    for (const auto& token : tokens) {
        if (token.kind==Token::String) text.append(literal(token.value));
        else if (token.kind==Token::Identifier) { auto value=token.value; value.replace(']',"]]"); text.append('['+value+']'); }
        else text.append(token.value);
    }
    return text.join(' ');
}
struct ColumnOrigin { QString label,table,column; };
// Inspect stored view text before asking SQLite to expand it. SQLite 3.45 can
// crash while preparing nested outer joins with an unresolved dependency.
// This check does not prepare a view or evaluate any user expression.
bool missingSqlDependency(QSqlDatabase db,const QString& sql,QSet<QString> visiting={});
bool missingSourceDependency(QSqlDatabase db,const QString& source,QSet<QString> visiting={}) {
    const auto key=source.toCaseFolded();
    if (visiting.contains(key) || visiting.size()>=64) return false;
    visiting.insert(key);
    if (key=="sqlite_master" || key=="sqlite_schema" || key=="sqlite_temp_master" || key=="sqlite_temp_schema") return false;
    QSqlQuery schema(db);
    schema.prepare("SELECT type,sql FROM temp.sqlite_master WHERE name=? COLLATE NOCASE UNION ALL SELECT type,sql FROM main.sqlite_master WHERE name=? COLLATE NOCASE LIMIT 1");
    schema.addBindValue(source); schema.addBindValue(source);
    if (!schema.exec()) return false;
    if (!schema.next()) return true;
    const auto type=schema.value(0).toString(), text=schema.value(1).toString();
    schema.finish();
    return type=="view" && missingSqlDependency(db,text,visiting);
}
bool missingSqlDependency(QSqlDatabase db,const QString& sql,QSet<QString> visiting) {
    const auto tokens=tokenize(sql,true);
    // CTE names have lexical scope. Leave those statements to SQLite rather
    // than mistake a local relation for an absent persistent table.
    for (const auto& token:tokens)
        if (token.kind==Token::Word && token.value.compare("WITH",Qt::CaseInsensitive)==0) return false;
    for (qsizetype i=0;i<tokens.size();++i) {
        if (tokens[i].kind!=Token::Word || (tokens[i].value.compare("FROM",Qt::CaseInsensitive)!=0 && tokens[i].value.compare("JOIN",Qt::CaseInsensitive)!=0)) continue;
        auto next=i+1;
        while (next<tokens.size() && symbol(tokens[next],"(")) ++next;
        if (next>=tokens.size() || (tokens[next].kind!=Token::Word && tokens[next].kind!=Token::Identifier)) continue;
        if (tokens[next].kind==Token::Word && tokens[next].value.compare("SELECT",Qt::CaseInsensitive)==0) continue;
        // Qualified sources and table-valued functions require different
        // resolution rules; do not infer an absent relation for them.
        if (next+1<tokens.size() && (symbol(tokens[next+1],".") || symbol(tokens[next+1],"("))) continue;
        if (missingSourceDependency(db,tokens[next].value,visiting)) return true;
    }
    return false;
}
QVector<ColumnOrigin> columnOrigins(QSqlDatabase db,const QString& source) {
    if (missingSourceDependency(db,source)) return {};
    static QLibrary library("sqlite3",0);
    library.setLoadHints(QLibrary::PreventUnloadHint);
    const auto prepare=reinterpret_cast<int (*)(sqlite3*,const char*,int,sqlite3_stmt**,const char**)>(library.resolve("sqlite3_prepare_v2"));
    const auto finish=reinterpret_cast<int (*)(sqlite3_stmt*)>(library.resolve("sqlite3_finalize"));
    const auto count=reinterpret_cast<int (*)(sqlite3_stmt*)>(library.resolve("sqlite3_column_count"));
    const auto label=reinterpret_cast<const char* (*)(sqlite3_stmt*,int)>(library.resolve("sqlite3_column_name"));
    const auto table=reinterpret_cast<const char* (*)(sqlite3_stmt*,int)>(library.resolve("sqlite3_column_table_name"));
    const auto column=reinterpret_cast<const char* (*)(sqlite3_stmt*,int)>(library.resolve("sqlite3_column_origin_name"));
    const auto handle=db.driver()->handle();
    if (!prepare || !finish || !count || !label || !table || !column || !handle.isValid() || QByteArray(handle.typeName())!="sqlite3*") return {};
    const auto connection=*static_cast<sqlite3* const*>(handle.constData());
    if (!connection) return {};
    const auto sql=("SELECT * FROM "+identifier(source)+" LIMIT 0").toUtf8();
    sqlite3_stmt* statement=nullptr;
    const auto result=prepare(connection,sql.constData(),sql.size(),&statement,nullptr);
    struct Guard { sqlite3_stmt* statement; int (*finish)(sqlite3_stmt*); ~Guard() { if (statement) finish(statement); } } guard{statement,finish};
    if (result!=0 || !statement) return {};
    QVector<ColumnOrigin> origins;
    for (int i=0;i<count(statement);++i) origins.append({QString::fromUtf8(label(statement,i)),QString::fromUtf8(table(statement,i)),QString::fromUtf8(column(statement,i))});
    // No sqlite3_step: obtaining lineage must never evaluate user expressions.
    return origins;
}
QString resolveQueryOrigins(QSqlDatabase db,const QString& sql) {
    auto tokens=tokenize(sql,true);
    const auto word=[&](qsizetype i,const char* text) { return i>=0 && i<tokens.size() && tokens[i].kind==Token::Word && tokens[i].value.compare(text,Qt::CaseInsensitive)==0; };
    const auto name=[](const Token& token) { return token.kind==Token::Word || token.kind==Token::Identifier; };
    if (!word(0,"SELECT")) return sql;
    int depth=0; qsizetype from=-1,end=tokens.size();
    for (qsizetype i=1;i<tokens.size();++i) {
        if (word(i,"SELECT")) return sql; // Nested scopes require independent lineage.
        if (symbol(tokens[i],"(")) ++depth;
        else if (symbol(tokens[i],")")) --depth;
        else if (depth==0 && word(i,"FROM")) { from=i; break; }
    }
    if (from<0 || from+1>=tokens.size()) return sql;
    for (auto i=from+2;i<tokens.size();++i) {
        if (word(i,"WHERE") || word(i,"GROUP") || word(i,"HAVING") || word(i,"ORDER") || word(i,"LIMIT") || symbol(tokens[i],";")) { end=i; break; }
    }
    for (auto i=from+1;i<end;++i) if (symbol(tokens[i],",")) return sql;
    struct Source { QString name,alias; QVector<ColumnOrigin> origins; };
    QVector<Source> sources;
    const QSet<QString> reserved{"inner","left","right","outer","full","cross","join","on","using"};
    for (auto i=from+1;i<end;++i) {
        if (i!=from+1 && !word(i-1,"JOIN")) continue;
        while (i<end && symbol(tokens[i],"(")) ++i;
        if (i>=end || !name(tokens[i]) || (i+1<end && symbol(tokens[i+1],"."))) return sql;
        Source source{tokens[i].value,tokens[i].value,{}};
        if (i+2<end && word(i+1,"AS") && name(tokens[i+2])) source.alias=tokens[i+2].value;
        else if (i+1<end && name(tokens[i+1]) && !reserved.contains(tokens[i+1].value.toCaseFolded())) source.alias=tokens[i+1].value;
        source.origins=columnOrigins(db,source.name);
        if (source.origins.empty()) return sql;
        sources.append(source);
    }
    if (sources.empty()) return sql;
    QSet<QString> fields,visibleOwners;
    for (const auto& source : sources) {
        visibleOwners.insert(source.alias.toCaseFolded());
        for (const auto& origin : source.origins) fields.insert(origin.label.toCaseFolded());
    }
    bool changed=false;
    for (qsizetype i=1;i+2<tokens.size();++i) {
        if (!name(tokens[i]) || !symbol(tokens[i+1],".") || !name(tokens[i+2])) continue;
        const auto owner=tokens[i].value;
        if (visibleOwners.contains(owner.toCaseFolded())) continue;
        QVector<QPair<QString,QString>> matches;
        for (const auto& source : sources) for (const auto& origin : source.origins)
            if (!origin.table.isEmpty() && origin.table.compare(owner,Qt::CaseInsensitive)==0 && origin.column.compare(tokens[i+2].value,Qt::CaseInsensitive)==0) matches.append({source.alias,origin.label});
        if (matches.size()!=1) continue; // Never guess across ambiguous origins.
        tokens[i]={Token::Identifier,matches.front().first}; tokens[i+2]={Token::Identifier,matches.front().second}; changed=true;
    }
    // Jet permits referring to earlier computed projection aliases. Expand only
    // known pure expressions and only when no source column shadows the alias.
    QMap<QString,QVector<Token>> aliases;
    qsizetype first=(word(1,"DISTINCT") || word(1,"ALL")) ? 2 : 1;
    QVector<Token> output=tokens.mid(0,first);
    depth=0;
    for (qsizetype i=first;i<=from;++i) {
        if (i!=from && symbol(tokens[i],"(")) ++depth;
        else if (i!=from && symbol(tokens[i],")")) --depth;
        if (i!=from && !(depth==0 && symbol(tokens[i],","))) continue;
        auto projection=tokens.mid(first,i-first);
        const bool explicitAlias=projection.size()>=3 && projection[projection.size()-2].kind==Token::Word && projection[projection.size()-2].value.compare("AS",Qt::CaseInsensitive)==0 && name(projection.last());
        const auto expressionEnd=projection.size()-(explicitAlias ? 2 : 0);
        QVector<Token> expanded;
        for (qsizetype j=0;j<expressionEnd;++j) {
            const auto key=projection[j].value.toCaseFolded();
            const bool qualified=(j>0 && symbol(projection[j-1],".")) || (j+1<expressionEnd && symbol(projection[j+1],"."));
            if (name(projection[j]) && !qualified && !fields.contains(key) && aliases.contains(key) && !(j+1<expressionEnd && symbol(projection[j+1],"("))) {
                expanded.append({Token::Other,"("}); expanded+=aliases[key]; expanded.append({Token::Other,")"}); changed=true;
            } else expanded.append(projection[j]);
        }
        if (expanded.size()>4096 || output.size()+expanded.size()>16384) return sql;
        bool pure=true;
        static const QSet<QString> functions{"linaccess_jet_iif","linaccess_jet_format","linaccess_jet_format_string","linaccess_jet_concat","abs","round","linaccess_jet_round","ccur","linaccess_jet_ccur","linaccess_jet_currency_add","linaccess_jet_currency_subtract","linaccess_jet_currency_multiply","space","linaccess_jet_space","string","linaccess_jet_string","asc","linaccess_jet_asc","ascw","linaccess_jet_ascw","chr","linaccess_jet_chr","chrw","linaccess_jet_chrw","instrrev","linaccess_jet_instrrev","strreverse","linaccess_jet_strreverse","datediff","linaccess_jet_datediff","year","linaccess_jet_year","month","linaccess_jet_month","day","linaccess_jet_day","hour","linaccess_jet_hour","minute","linaccess_jet_minute","second","linaccess_jet_second","weekday","linaccess_jet_weekday","len","linaccess_jet_len","rgb","linaccess_jet_rgb","qbcolor","linaccess_jet_qbcolor","length","lower","upper","coalesce"};
        for (qsizetype j=0;j+1<expanded.size();++j)
            if (name(expanded[j]) && symbol(expanded[j+1],"(") && !functions.contains(expanded[j].value.toCaseFolded())) pure=false;
        if (explicitAlias && pure) aliases.insert(projection.last().value.toCaseFolded(),expanded);
        output+=expanded;
        if (explicitAlias) output+=projection.mid(expressionEnd);
        if (i!=from) output.append(tokens[i]);
        first=i+1;
    }
    if (!changed) return sql;
    output+=tokens.mid(from);
    return renderJetTokens(output);
}
struct RightJoinCandidate { QString source; QStringList tables; };
RightJoinCandidate simpleRightJoin(const QVector<Token>& tokens,qsizetype start) {
    const auto name=[](const Token& token) { return token.kind==Token::Word || token.kind==Token::Identifier; };
    const auto word=[&](qsizetype i,const char* value) { return i<tokens.size() && tokens[i].kind==Token::Word && tokens[i].value.compare(value,Qt::CaseInsensitive)==0; };
    if (!word(start,"SELECT")) return {};
    const QSet<QString> scalar{"iif","format","format$","abs","round","linaccess_jet_round","ccur","linaccess_jet_ccur","linaccess_jet_currency_add","linaccess_jet_currency_subtract","linaccess_jet_currency_multiply","coalesce","ifnull","nullif","space","linaccess_jet_space","string","linaccess_jet_string","asc","linaccess_jet_asc","ascw","linaccess_jet_ascw","chr","linaccess_jet_chr","chrw","linaccess_jet_chrw","instrrev","linaccess_jet_instrrev","strreverse","linaccess_jet_strreverse","datediff","linaccess_jet_datediff","year","linaccess_jet_year","month","linaccess_jet_month","day","linaccess_jet_day","hour","linaccess_jet_hour","minute","linaccess_jet_minute","second","linaccess_jet_second","weekday","linaccess_jet_weekday","len","linaccess_jet_len","rgb","linaccess_jet_rgb","qbcolor","linaccess_jet_qbcolor","length","lower","upper","substr","trim","ltrim","rtrim","replace","linaccess_jet_iif","linaccess_jet_format","linaccess_jet_format_string","linaccess_jet_concat"};
    const QSet<QString> syntax{"where","on","and","or","not","in","is","by"};
    qsizetype from=-1,right=-1,join=-1,on=-1,end=-1; int selects=0,joins=0,depth=0;
    for (qsizetype i=start;i<tokens.size();++i) {
        if (word(i,"SELECT") && ++selects>1) return {};
        if (word(i,"JOIN")) ++joins;
        if (word(i,"GROUP") || word(i,"HAVING") || word(i,"UNION") || symbol(tokens[i],"*")) return {};
        if (i+1<tokens.size() && symbol(tokens[i+1],"(") && tokens[i].kind!=Token::Other && !scalar.contains(tokens[i].value.toCaseFolded()) && !(tokens[i].kind==Token::Word && syntax.contains(tokens[i].value.toCaseFolded()))) return {};
        if (depth==0) {
            if (word(i,"FROM")) from=i;
            if (word(i,"RIGHT")) right=i;
            if (word(i,"JOIN")) join=i;
            if (word(i,"ON")) on=i;
            if (on>=0 && end<0 && (word(i,"WHERE") || word(i,"ORDER") || symbol(tokens[i],";"))) end=i;
        }
        if (symbol(tokens[i],"(")) ++depth;
        else if (symbol(tokens[i],")")) --depth;
    }
    if (from<0 || right<=from || join!=right+1 || on<=join || joins!=1) return {};
    if (end<0) end=tokens.size();
    const auto left=tokens.mid(from+1,right-from-1),rhs=tokens.mid(join+1,on-join-1);
    const auto relation=[&](const QVector<Token>& values) {
        if (values.empty() || values.size()>3 || !name(values[0])) return false;
        if (values.size()==2) return name(values[1]);
        return values.size()==1 || (values[1].kind==Token::Word && values[1].value.compare("AS",Qt::CaseInsensitive)==0 && name(values[2]));
    };
    if (!relation(left) || !relation(rhs)) return {};
    auto condition=tokens.mid(on+1,end-on-1);
    while (condition.size()>2 && symbol(condition.first(),"(") && symbol(condition.last(),")")) {
        int level=0; bool enclosing=true;
        for (qsizetype i=0;i<condition.size()-1;++i) { if (symbol(condition[i],"(")) ++level; else if (symbol(condition[i],")")) --level; if (level==0) enclosing=false; }
        if (!enclosing) break;
        condition.removeFirst(); condition.removeLast();
    }
    if (condition.size()!=7 || !symbol(condition[3],"=")) return {};
    for (const auto index : {0,2,4,6}) if (!name(condition[index])) return {};
    if (!symbol(condition[1],".") || !symbol(condition[5],".")) return {};
    auto rewritten=tokens.mid(0,from+1); rewritten+=rhs;
    rewritten.append({Token::Word,"LEFT"}); rewritten.append({Token::Word,"JOIN"}); rewritten+=left; rewritten+=tokens.mid(on);
    return {renderJetTokens(rewritten),{left[0].value,rhs[0].value}};
}
bool stableTopSource(QSqlDatabase db,const QString& source,QSet<QString> visiting={}) {
    const auto key=source.toCaseFolded();
    if (visiting.contains(key) || visiting.size()>=32) return false;
    visiting.insert(key);
    QSqlQuery schema(db);
    schema.prepare("SELECT type,sql FROM temp.sqlite_master WHERE name=? COLLATE NOCASE UNION ALL SELECT type,sql FROM main.sqlite_master WHERE name=? COLLATE NOCASE LIMIT 1");
    schema.addBindValue(source); schema.addBindValue(source);
    if (!schema.exec() || !schema.next()) return false;
    const auto type=schema.value(0).toString(); const auto sql=schema.value(1).toString(); schema.finish();
    if (type=="table") return !sql.trimmed().startsWith("CREATE VIRTUAL",Qt::CaseInsensitive);
    if (type!="view") return false;
    const auto tokens=tokenize(sql);
    const QSet<QString> pure{"abs","coalesce","ifnull","nullif","iif","sum","linaccess_jet_currency_sum","total","count","min","max","avg","round","linaccess_jet_round","ccur","linaccess_jet_ccur","linaccess_jet_currency_add","linaccess_jet_currency_subtract","linaccess_jet_currency_multiply","space","linaccess_jet_space","string","linaccess_jet_string","asc","linaccess_jet_asc","ascw","linaccess_jet_ascw","chr","linaccess_jet_chr","chrw","linaccess_jet_chrw","instrrev","linaccess_jet_instrrev","strreverse","linaccess_jet_strreverse","datediff","linaccess_jet_datediff","year","linaccess_jet_year","month","linaccess_jet_month","day","linaccess_jet_day","hour","linaccess_jet_hour","minute","linaccess_jet_minute","second","linaccess_jet_second","weekday","linaccess_jet_weekday","len","linaccess_jet_len","rgb","linaccess_jet_rgb","qbcolor","linaccess_jet_qbcolor","length","lower","upper","substr","substring","trim","ltrim","rtrim","replace","cast","typeof","rank",
        "linaccess_jet_iif","linaccess_jet_first","linaccess_jet_last","linaccess_jet_format","linaccess_jet_format_string","linaccess_jet_concat","linaccess_jet_like_pattern"};
    const QSet<QString> syntax{"from","join","as","in","not","exists","over","by","and","or","on","where","having","is","when","then","else","case"};
    QSet<QString> dependencies; int selects=0; bool inFrom=false;
    for (qsizetype i=0;i<tokens.size();++i) {
        const auto word=tokens[i].value.toCaseFolded();
        if (tokens[i].kind==Token::Word && word=="select" && ++selects>1) return false;
        if (i+1<tokens.size() && symbol(tokens[i+1],"(") && tokens[i].kind!=Token::Other && !pure.contains(word) && !(tokens[i].kind==Token::Word && syntax.contains(word))) return false;
        if (tokens[i].kind==Token::Word && (word=="current_timestamp" || word=="current_date" || word=="current_time")) return false;
        if (tokens[i].kind==Token::Word && (word=="where" || word=="group" || word=="having" || word=="order" || word=="limit" || word=="union")) inFrom=false;
        if (inFrom && symbol(tokens[i],",")) return false;
        if (tokens[i].kind==Token::Word && (word=="from" || word=="join")) {
            inFrom=true;
            auto next=i+1;
            while (next<tokens.size() && symbol(tokens[next],"(")) ++next;
            if (next>=tokens.size() || tokens[next].kind==Token::Other) return false;
            auto name=tokens[next].value;
            if (next+2<tokens.size() && symbol(tokens[next+1],".")) {
                if (name.compare("main",Qt::CaseInsensitive)!=0 && name.compare("temp",Qt::CaseInsensitive)!=0) return false;
                name=tokens[next+2].value;
            }
            dependencies.insert(name);
        }
    }
    for (const auto& dependency : dependencies) if (!stableTopSource(db,dependency,visiting)) return false;
    return true;
}
bool useTopOneBoundary(QSqlDatabase db,const AccessQueryPlan& plan) {
    return !plan.topOneSql.isEmpty() && stableTopSource(db,plan.topOneTable) && db.record(plan.topOneTable).indexOf(plan.topOneColumn)>=0;
}
bool useLeftJoin(QSqlDatabase db,const AccessQueryPlan& plan) {
    if (plan.rightJoinSql.isEmpty() || plan.rightJoinSources.size()!=2) return false;
    const auto tables=db.tables(QSql::Tables);
    for (const auto& source : plan.rightJoinSources)
        if (!tables.contains(source,Qt::CaseInsensitive) || !stableTopSource(db,source)) return false;
    return true;
}
TopOneCandidate topOneCandidate(const QVector<Token>& tokens,qsizetype start,qsizetype from,qsizetype order,const QMap<QString,QVector<Token>>& aliases) {
    const auto name=[](const Token& token) { return token.kind==Token::Word || token.kind==Token::Identifier; };
    const auto word=[&](qsizetype i,const char* value) { return i<tokens.size() && tokens[i].kind==Token::Word && tokens[i].value.compare(value,Qt::CaseInsensitive)==0; };
    auto key=tokens.mid(order+2);
    if (!key.empty() && key.last().kind==Token::Word && (key.last().value.compare("ASC",Qt::CaseInsensitive)==0 || key.last().value.compare("DESC",Qt::CaseInsensitive)==0)) key.removeLast();
    if (key.size()==1 && name(key[0]) && aliases.contains(key[0].value.toCaseFolded())) key=aliases.value(key[0].value.toCaseFolded());
    if (key.empty() || key.size()>3 || key.size()%2==0) return {};
    for (qsizetype i=0;i<key.size();++i) if (i%2==0 ? !name(key[i]) : !symbol(key[i],".")) return {};
    qsizetype where=-1;
    for (auto i=from+1;i<order;++i) {
        if (word(i,"WHERE") && where<0) where=i;
        if (word(i,"GROUP") || word(i,"HAVING") || word(i,"SELECT")) return {};
        if (where>=0 && i>where && tokens[i].kind==Token::Word && i+1<order && symbol(tokens[i+1],"(") && !word(i,"AND") && !word(i,"OR") && !word(i,"NOT") && !word(i,"IN")) return {};
    }
    const auto relationEnd=where>=0 ? where : order;
    const auto relation=tokens.mid(from+1,relationEnd-from-1);
    if (relation.empty() || relation.size()>3 || !name(relation[0])) return {};
    if (relation.size()==2 && !name(relation[1])) return {};
    if (relation.size()==3 && (!word(from+2,"AS") || !name(relation[2]))) return {};
    // Only direct field/wildcard projections: evaluating user functions twice
    // would not preserve their behavior. Complex sorts/joins retain RANK.
    qsizetype first=start+1;
    for (auto i=first;i<=from;++i) if (i==from || symbol(tokens[i],",")) {
        auto projection=tokens.mid(first,i-first);
        if (projection.size()>=3 && projection[projection.size()-2].kind==Token::Word && projection[projection.size()-2].value.compare("AS",Qt::CaseInsensitive)==0) projection.resize(projection.size()-2);
        if (projection.empty() || projection.size()>3 || projection.size()%2==0) return {};
        for (qsizetype j=0;j<projection.size();++j)
            if (j%2==0 ? !(name(projection[j]) || (j==projection.size()-1 && symbol(projection[j],"*"))) : !symbol(projection[j],".")) return {};
        first=i+1;
    }
    const auto source=tokens.mid(from,order-from);
    QVector<Token> cutoff{{Token::Word,"SELECT"}}; cutoff+=key; cutoff+=source;
    cutoff.append({Token::Word,"ORDER"}); cutoff.append({Token::Word,"BY"}); cutoff+=key;
    if (tokens.last().kind==Token::Word && (tokens.last().value.compare("ASC",Qt::CaseInsensitive)==0 || tokens.last().value.compare("DESC",Qt::CaseInsensitive)==0)) cutoff.append(tokens.last());
    cutoff.append({Token::Word,"LIMIT"}); cutoff.append({Token::Other,"1"});
    QVector<Token> candidate=tokens.mid(0,where>=0 ? where : order);
    candidate.append({Token::Word,"WHERE"});
    if (where>=0) {
        candidate.append({Token::Other,"("}); candidate+=tokens.mid(where+1,order-where-1);
        candidate.append({Token::Other,")"}); candidate.append({Token::Word,"AND"});
    }
    candidate+=key; candidate.append({Token::Word,"IS"}); candidate.append({Token::Other,"("}); candidate+=cutoff;
    candidate.append({Token::Other,")"}); candidate+=tokens.mid(order);
    return {renderJetTokens(candidate),relation[0].value,key.last().value};
}
QString pattern(const QString& input) {
    require(!input.contains('[') && !input.contains(']') && !input.contains('#'), "Jet SQL: LIKE character classes unsupported");
    QString result;
    for (const auto ch : input) {
        if (ch == '*') result += '%';
        else if (ch == '?') result += '_';
        else { if (ch == '%' || ch == '_' || ch == '\\') result += '\\'; result += ch; }
    }
    return result;
}
}
namespace {
QString updateSourceText(const QVector<Token>& tokens,qsizetype begin,qsizetype end) {
    QStringList parts;
    for (auto i=begin;i<end;++i) {
        const auto& token=tokens[i];
        if (token.kind==Token::Identifier) { auto name=token.value; parts.append('['+name.replace(']',"]]")+']'); }
        else if (token.kind==Token::String) { auto value=token.value; parts.append('"'+value.replace('"',"\"\"")+'"'); }
        else parts.append(token.value);
    }
    return parts.join(' ');
}
AccessQueryPlan compileJoinedUpdate(const QVector<Token>& tokens,qsizetype start) {
    const auto word=[&](qsizetype index,const char* value) { return index<tokens.size() && tokens[index].kind==Token::Word && tokens[index].value.compare(value,Qt::CaseInsensitive)==0; };
    const auto name=[&](qsizetype index) { return index<tokens.size() && (tokens[index].kind==Token::Word || tokens[index].kind==Token::Identifier); };
    const auto check=[](bool valid) { if (!valid) throw VbaRuntimeError(3251,"Joined UPDATE shape unsupported"); };
    qsizetype set=start+1,end=tokens.size();
    if (end>start && symbol(tokens[end-1],";")) --end;
    while (set<end && !word(set,"SET")) ++set;
    check(set<end);
    qsizetype joins=0; for (auto i=start+1;i<set;++i) if (word(i,"JOIN")) ++joins;
    check(joins==1);
    struct Relation { QString table,alias; };
    auto position=start+1;
    while (position<set && symbol(tokens[position],"(")) ++position;
    const auto relation=[&]() {
        check(position<set && name(position)); Relation result{tokens[position].value,tokens[position].value}; ++position;
        if (word(position,"AS")) { ++position; check(name(position)); result.alias=tokens[position++].value; }
        return result;
    };
    const auto left=relation();
    enum class JoinSide { Inner,Left,Right };
    auto side=JoinSide::Inner;
    if (word(position,"INNER")) ++position;
    else if (word(position,"LEFT") || word(position,"RIGHT")) {
        side=word(position,"LEFT") ? JoinSide::Left : JoinSide::Right; ++position;
        if (word(position,"OUTER")) ++position;
    }
    check(word(position,"JOIN")); ++position;
    const auto right=relation(); check(word(position,"ON"));
    check(left.alias.compare(right.alias,Qt::CaseInsensitive)!=0);
    qsizetype where=set+1; int depth=0;
    for (;where<end;++where) {
        if (symbol(tokens[where],"(")) ++depth;
        else if (symbol(tokens[where],")")) --depth;
        if (!depth && word(where,"WHERE")) break;
    }
    struct Assignment { QString column; QString expression; };
    QVector<Assignment> assignments; QString targetAlias; QSet<QString> assigned;
    auto first=set+1; depth=0;
    for (auto i=first;i<=where;++i) {
        if (i<where && symbol(tokens[i],"(")) ++depth;
        else if (i<where && symbol(tokens[i],")")) --depth;
        if (i!=where && (depth || !symbol(tokens[i],","))) continue;
        check(i-first>=5 && name(first) && symbol(tokens[first+1],".") && name(first+2) && symbol(tokens[first+3],"="));
        if (targetAlias.isEmpty()) targetAlias=tokens[first].value;
        check(targetAlias.compare(tokens[first].value,Qt::CaseInsensitive)==0);
        const auto column=tokens[first+2].value; check(!assigned.contains(column.toCaseFolded())); assigned.insert(column.toCaseFolded());
        assignments.append({column,updateSourceText(tokens,first+4,i)}); first=i+1;
    }
    check(!assignments.isEmpty());
    const auto* target=targetAlias.compare(left.alias,Qt::CaseInsensitive)==0 ? &left : targetAlias.compare(right.alias,Qt::CaseInsensitive)==0 ? &right : nullptr;
    check(target);
    check((side!=JoinSide::Left || target==&left) && (side!=JoinSide::Right || target==&right));
    const auto key=QString::fromLatin1(QCryptographicHash::hash(updateSourceText(tokens,0,end).toUtf8(),QCryptographicHash::Sha256).toHex());
    const auto cte="__la_update_"+key,identity="__la_identity_"+key,row="__la_row_"+key;
    auto accessName=[](QString value) { return '['+value.replace(']',"]]")+']'; };
    QStringList projection; projection.append(accessName(targetAlias)+'.'+accessName(identity)+" AS "+accessName(row));
    for (qsizetype i=0;i<assignments.size();++i) projection.append(assignments[i].expression+" AS "+accessName("__la_value_"+QString::number(i)));
    const auto select=updateSourceText(tokens,0,start)+" SELECT "+projection.join(',')+" FROM "+updateSourceText(tokens,start+1,set)+(where<end ? " WHERE "+updateSourceText(tokens,where+1,end) : QString{});
    auto plan=compileAccessQuery(select);
    check(!plan.hiddenRank && !plan.topLimit && plan.makeTable.isEmpty());
    plan.updateTarget=target->table; plan.updateIdentity=identity; plan.updateSource=plan.sql;
    const auto prefix="WITH "+identifier(cte)+" AS MATERIALIZED ("; plan.updateSourceOffset=prefix.size();
    QStringList changes;
    const auto id=identifier(targetAlias)+'.'+identifier(identity);
    for (qsizetype i=0;i<assignments.size();++i)
        changes.append(identifier(assignments[i].column)+"=(SELECT linaccess_jet_update_value("+identifier("__la_value_"+QString::number(i))+") FROM "+identifier(cte)+" WHERE "+identifier(row)+'='+id+')');
    plan.sql=prefix+plan.updateSource+") UPDATE "+identifier(target->table)+" AS "+identifier(targetAlias)+" SET "+changes.join(',')+" WHERE "+id+" IN (SELECT "+identifier(row)+" FROM "+identifier(cte)+')';
    return plan;
}
}
QString accessLikePattern(const QString& value) { return pattern(value); }
static AccessQueryPlan compileCurrencyQuery(const QString& source,const QMap<QString,QString>& fields);
AccessQueryPlan compileAccessQuery(const QString& source) { return compileCurrencyQuery(source,{}); }
static AccessQueryPlan compileCurrencyQuery(const QString& source,const QMap<QString,QString>& fields) {
    auto tokens = tokenize(source);
    require(!tokens.empty(), "Jet SQL: empty query");
    AccessQueryPlan plan;
    QMap<QString,int> parameters;
    qsizetype pos = 0;
    if (tokens[0].kind == Token::Word && tokens[0].value.compare("PARAMETERS",Qt::CaseInsensitive) == 0) {
        ++pos;
        while (pos < tokens.size()) {
            require(tokens[pos].kind == Token::Word || tokens[pos].kind == Token::Identifier, "Jet SQL: parameter name required");
            const auto name = tokens[pos++].value;
            require(pos < tokens.size() && tokens[pos].kind == Token::Word, "Jet SQL: parameter type required");
            auto type = tokens[pos++].value.toCaseFolded();
            static const QMap<QString,QString> aliases{
                {"money","currency"},
                {"bit","boolean"},{"logical","boolean"},{"logical1","boolean"},
                {"integer","long"},{"int","long"},{"integer4","long"},
                {"smallint","short"},{"integer2","short"},
                {"tinyint","byte"},{"integer1","byte"},
                {"real","single"},{"float4","single"},{"ieeesingle","single"},
                {"float","double"},{"float8","double"},{"ieeedouble","double"},{"number","double"}};
            type=aliases.value(type,type);
            const QSet<QString> types{"long","integer","short","byte","single","double","currency","text","varchar","char","boolean","yesno","datetime","date","value"};
            require(types.contains(type), "Jet SQL: parameter type unsupported");
            if (pos < tokens.size() && tokens[pos].value == "(") {
                require(type!="currency","Jet SQL: Currency parameter width unsupported");
                ++pos; require(pos < tokens.size() && tokens[pos].kind == Token::Other, "Jet SQL: type width required");
                bool valid = false; const auto width = tokens[pos++].value.toInt(&valid);
                require(valid && width > 0 && pos < tokens.size() && tokens[pos++].value == ")", "Jet SQL: invalid type width");
            }
            require(!parameters.contains(name.toCaseFolded()), "Jet SQL: duplicate parameter");
            const auto binding = ":la_parameter_" + QString::number(plan.parameters.size());
            parameters.insert(name.toCaseFolded(),plan.parameters.size());
            plan.parameters.append({name,type,binding,false});
            require(pos < tokens.size(), "Jet SQL: PARAMETERS terminator required");
            const auto separator = tokens[pos++].value;
            if (separator == ";") break;
            require(separator == ",", "Jet SQL: invalid parameter separator");
        }
    }
    require(pos < tokens.size(), "Jet SQL: query body required");
    plan.resultTypes=currencyProjectionTypes(tokens,pos,plan.parameters,fields);
    if (tokens[pos].kind==Token::Word && tokens[pos].value.compare("UPDATE",Qt::CaseInsensitive)==0) {
        for (auto i=pos+1;i<tokens.size();++i) {
            if (tokens[i].kind==Token::Word && tokens[i].value.compare("SET",Qt::CaseInsensitive)==0) break;
            if (tokens[i].kind==Token::Word && tokens[i].value.compare("JOIN",Qt::CaseInsensitive)==0) return compileJoinedUpdate(tokens,pos);
        }
    }
    // Jet make-table queries place INTO between the projection and FROM.
    // Inspect tokens at the outer query depth, never text inside a literal.
    if (tokens[pos].kind==Token::Word && tokens[pos].value.compare("SELECT",Qt::CaseInsensitive)==0) {
        int depth=0;
        for (qsizetype i=pos+1;i<tokens.size();++i) {
            if (symbol(tokens[i],"(")) { ++depth; continue; }
            if (symbol(tokens[i],")")) { --depth; continue; }
            if (depth!=0 || tokens[i].kind!=Token::Word) continue;
            if (tokens[i].value.compare("FROM",Qt::CaseInsensitive)==0) break;
            if (tokens[i].value.compare("INTO",Qt::CaseInsensitive)!=0) continue;
            if (i+2>=tokens.size() || (tokens[i+1].kind!=Token::Word && tokens[i+1].kind!=Token::Identifier)
                || tokens[i+1].value.isEmpty() || tokens[i+2].kind!=Token::Word
                || tokens[i+2].value.compare("FROM",Qt::CaseInsensitive)!=0)
                throw VbaRuntimeError(3251,"Make-table destination must be a local table followed by FROM");
            plan.makeTable=tokens[i+1].value;
            tokens.remove(i,2); break;
        }
    }
    adaptObjectReferences(tokens,pos);
    adaptConcatenation(tokens,pos);
    adaptDynamicLike(tokens,pos,parameters);
    preserveQualifiedProjectionNames(tokens,pos);
    const auto word = [&](qsizetype index, const char* value) {
        return index >= 0 && index < tokens.size() && tokens[index].kind == Token::Word
            && tokens[index].value.compare(value,Qt::CaseInsensitive) == 0;
    };
    if (word(pos,"SELECT")) {
        auto modifier = pos + 1;
        const bool distinctProjection = word(modifier,"DISTINCT");
        if (word(modifier,"DISTINCTROW")) {
            int depth = 0; bool from = false; bool single = true;
            for (qsizetype i = modifier + 1; i < tokens.size(); ++i) {
                if (tokens[i].value == "(") { if (from && depth == 0) single = false; ++depth; }
                else if (tokens[i].value == ")") --depth;
                if (depth != 0) continue;
                if (word(i,"FROM")) { from = true; continue; }
                if (from && (word(i,"WHERE") || word(i,"GROUP") || word(i,"HAVING") || word(i,"ORDER"))) break;
                if (from && (word(i,"JOIN") || tokens[i].value == ",")) single = false;
            }
            const bool allColumns=modifier+2<tokens.size() && symbol(tokens[modifier+1],"*") && word(modifier+2,"FROM");
            require((from && single) || allColumns,"Jet SQL: multitable DISTINCTROW requires record identity mapping");
            // Jet ignores DISTINCTROW for one source or when all source-table
            // fields are output. SELECT DISTINCT would erase valid duplicates.
            tokens.removeAt(modifier);
        } else if (word(modifier,"DISTINCT")) ++modifier;
        if (word(modifier,"TOP")) {
            require(modifier + 1 < tokens.size(),"Jet SQL: TOP count required");
            bool valid = false; plan.topLimit = tokens[modifier+1].value.toInt(&valid);
            require(valid && plan.topLimit > 0,"Jet SQL: TOP requires a positive integer");
            require(!word(modifier+2,"PERCENT"),"Jet SQL: TOP PERCENT unsupported");
            tokens.removeAt(modifier); tokens.removeAt(modifier);
            while (!tokens.empty() && tokens.last().value == ";") tokens.removeLast();
            int depth = 0; qsizetype from = -1, order = -1;
            for (qsizetype i = pos; i < tokens.size(); ++i) {
                if (tokens[i].value == "(") ++depth;
                else if (tokens[i].value == ")") --depth;
                if (depth != 0) continue;
                require(!word(i,"UNION") && !word(i,"LIMIT"),"Jet SQL: TOP with UNION/LIMIT unsupported");
                if (from < 0 && word(i,"FROM")) from = i;
                if (word(i,"ORDER") && word(i+1,"BY")) order = i;
            }
            if (order >= 0) {
                require(from >= 0,"Jet SQL: ordered TOP requires FROM");
                require(!distinctProjection,"Jet SQL: ordered TOP DISTINCT requires a deduplication stage");
                QMap<QString,QVector<Token>> aliases;
                int expressionDepth = 0; qsizetype expressionStart = pos + 1;
                for (qsizetype i = expressionStart; i <= from; ++i) {
                    const bool end = i == from || (tokens[i].value == "," && expressionDepth == 0);
                    if (end) {
                        for (qsizetype j = expressionStart; j + 1 < i; ++j) {
                            if (word(j,"AS") && j + 2 == i) {
                                QVector<Token> expression;
                                for (qsizetype k = expressionStart; k < j; ++k) expression.append(tokens[k]);
                                aliases.insert(tokens[j+1].value.toCaseFolded(),expression);
                            }
                        }
                        expressionStart = i + 1;
                    } else if (tokens[i].value == "(") ++expressionDepth;
                    else if (tokens[i].value == ")") --expressionDepth;
                }
                if (plan.topLimit==1) {
                    const auto candidate=topOneCandidate(tokens,pos,from,order,aliases);
                    if (!candidate.source.isEmpty()) {
                        plan.topOneSql=compileAccessQuery(candidate.source).sql;
                        plan.topOneTable=candidate.table; plan.topOneColumn=candidate.column;
                    }
                }
                QVector<Token> rank{{Token::Other,","},{Token::Word,"RANK"},{Token::Other,"("},{Token::Other,")"},
                    {Token::Word,"OVER"},{Token::Other,"("},{Token::Word,"ORDER"},{Token::Word,"BY"}};
                bool termStart = true; int orderDepth = 0;
                for (qsizetype i = order + 2; i < tokens.size(); ++i) {
                    bool ordinal = false; tokens[i].value.toInt(&ordinal);
                    require(!(termStart && ordinal),"Jet SQL: TOP ORDER BY ordinals require expression binding");
                    const bool qualified = (i > order+2 && tokens[i-1].value == ".")
                        || (i+1 < tokens.size() && tokens[i+1].value == ".");
                    const auto alias = aliases.constFind(tokens[i].value.toCaseFolded());
                    if (!qualified && alias != aliases.cend() && tokens[i].kind != Token::String) {
                        rank.append({Token::Other,"("});
                        rank += alias.value(); rank.append({Token::Other,")"});
                    } else rank.append(tokens[i]);
                    if (tokens[i].value == "(") ++orderDepth;
                    else if (tokens[i].value == ")") --orderDepth;
                    termStart = orderDepth == 0 && tokens[i].value == ",";
                }
                rank.append({Token::Other,")"}); rank.append({Token::Word,"AS"}); rank.append({Token::Identifier,"__la_top_rank"});
                for (qsizetype i = rank.size(); i > 0; --i) tokens.insert(from,rank[i-1]);
                plan.hiddenRank = true;
            } else {
                tokens.append({Token::Word,"LIMIT"}); tokens.append({Token::Other,QString::number(plan.topLimit)});
            }
        }
    }
    QStringList output;
    for (qsizetype i = pos; i < tokens.size(); ++i) {
        const auto& token = tokens[i];
        if (token.kind == Token::Word && i + 2 < tokens.size() && tokens[i+1].value == "(" && tokens[i+2].value == ")") {
            const auto function = token.value.toCaseFolded();
            if (function == "date" || function == "now" || function == "time") {
                output << (function == "date" ? "datetime('now','localtime','start of day')"
                    : function == "now" ? "datetime('now','localtime')" : "strftime('1899-12-30 %H:%M:%S','now','localtime')");
                i += 2; continue;
            }
        }
        const bool like = i > pos && tokens[i-1].kind == Token::Word && tokens[i-1].value.compare("LIKE",Qt::CaseInsensitive) == 0;
        const bool qualified = (i > pos && tokens[i-1].value == ".") || (i + 1 < tokens.size() && tokens[i+1].value == ".");
        const auto param = parameters.constFind(token.value.toCaseFolded());
        if ((token.kind == Token::Word || token.kind == Token::Identifier) && !qualified && param != parameters.cend()
            && !(i+1<tokens.size() && symbol(tokens[i+1],"("))) {
            plan.parameters[*param].likePattern |= like;
            plan.parameters[*param].direct |= !like;
            output << plan.parameters[*param].binding + (like ? "_like" : "");
            if (like) output << "ESCAPE '\\'";
        } else if (token.kind == Token::String) {
            output << literal(like ? pattern(token.value) : token.value);
            if (like) output << "ESCAPE '\\'";
        } else if (token.kind == Token::Identifier) output << identifier(token.value);
        else if (token.kind==Token::Word && i+1<tokens.size() && tokens[i+1].value=="("
                 && (token.value.compare("Format$",Qt::CaseInsensitive)==0 || token.value.compare("Format",Qt::CaseInsensitive)==0))
            output << (token.value.endsWith('$') ? "linaccess_jet_format_string" : "linaccess_jet_format");
        else if(token.kind==Token::Word && i+1<tokens.size() && tokens[i+1].value=="("
                && QSet<QString>{"space","string","asc","ascw","chr","chrw","instrrev","strreverse","datediff","year","month","day","hour","minute","second","weekday"}.contains(token.value.toCaseFolded()))
            output << "linaccess_jet_"+token.value.toCaseFolded();
        else if (token.kind==Token::Word && i+1<tokens.size() && tokens[i+1].value=="(" && token.value.compare("Len",Qt::CaseInsensitive)==0)
            output << "linaccess_jet_len";
        else if (token.kind==Token::Word && i+1<tokens.size() && tokens[i+1].value=="(" && (token.value.compare("RGB",Qt::CaseInsensitive)==0 || token.value.compare("QBColor",Qt::CaseInsensitive)==0 || token.value.compare("Round",Qt::CaseInsensitive)==0 || token.value.compare("CCur",Qt::CaseInsensitive)==0))
            output << "linaccess_jet_"+token.value.toCaseFolded();
        else if (token.kind == Token::Word && i + 1 < tokens.size() && tokens[i+1].value == "("
                 && (token.value.compare("IIf",Qt::CaseInsensitive)==0 || token.value.compare("First",Qt::CaseInsensitive)==0
                     || token.value.compare("Last",Qt::CaseInsensitive)==0)) output << "linaccess_jet_" + token.value.toCaseFolded();
        else if (token.kind == Token::Word && token.value.compare("TRUE",Qt::CaseInsensitive) == 0) output << "-1";
        else if (token.kind == Token::Word && token.value.compare("FALSE",Qt::CaseInsensitive) == 0) output << "0";
        else output << token.value;
    }
    plan.sql = output.join(' ');
    if (plan.topLimit==0) {
        const auto candidate=simpleRightJoin(tokens,pos);
        if (!candidate.source.isEmpty()) {
            plan.rightJoinSql=compileAccessQuery(candidate.source).sql;
            plan.rightJoinSources=candidate.tables;
        }
    }
    if (!plan.makeTable.isEmpty()) {
        if (plan.topLimit || !plan.rightJoinSql.isEmpty()) throw VbaRuntimeError(3251,"Make-table TOP/RIGHT JOIN materialization unsupported");
        plan.sql="CREATE TABLE "+identifier(plan.makeTable)+" AS "+plan.sql;
    }
    plan.source=source; plan.compiledSql=plan.sql;
    return plan;
}
static AccessQueryPlan resolveCurrencyQuery(QSqlDatabase db,const AccessQueryPlan& input) {
    if(input.source.isEmpty()) return input;
    if(input.sql!=input.compiledSql) { auto changed=input; changed.resultTypes.clear(); changed.source.clear(); return changed; }
    const auto source=input.currencyContextResolved ? compileAccessQuery(input.source) : input;
    if(!source.sql.trimmed().startsWith("SELECT",Qt::CaseInsensitive)) return source;
    const auto tokens=tokenize(source.source);
    const auto word=[&](qsizetype i,const char* value) { return i<tokens.size() && tokens[i].kind==Token::Word && tokens[i].value.compare(value,Qt::CaseInsensitive)==0; };
    const auto name=[&](qsizetype i) { return i<tokens.size() && (tokens[i].kind==Token::Word || tokens[i].kind==Token::Identifier); };
    qsizetype body=0;
    if(word(0,"PARAMETERS")) { while(body<tokens.size() && !symbol(tokens[body],";")) ++body; ++body; }
    if(!word(body,"SELECT")) return source;
    qsizetype from=-1; int depth=0;
    for(qsizetype i=body+1;i<tokens.size();++i) {
        if(word(i,"SELECT") || word(i,"HAVING") || word(i,"UNION") || word(i,"INTERSECT") || word(i,"EXCEPT")) return source;
        if(symbol(tokens[i],"(")) ++depth;
        else if(symbol(tokens[i],")")) --depth;
        else if(depth==0 && word(i,"FROM")) { if(from>=0) return source; from=i; }
        if(from>=0 && word(i,"SUM") && i+1<tokens.size() && symbol(tokens[i+1],"(")) return source;
    }
    if(from<0 || !name(from+1)) return source;
    const auto boundary=[&](qsizetype i) { return i>=tokens.size() || word(i,"WHERE") || word(i,"GROUP") || word(i,"ORDER") || word(i,"LIMIT") || symbol(tokens[i],";"); };
    qsizetype next=from+2; QString alias;
    if(word(next,"AS")) { ++next; if(!name(next)) return source; alias=tokens[next++].value; }
    else if(!boundary(next) && name(next)) alias=tokens[next++].value;
    if(!boundary(next)) return source;
    QString table;
    for(const auto& candidate:db.tables()) if(candidate.compare(tokens[from+1].value,Qt::CaseInsensitive)==0) { table=candidate; break; }
    if(table.isEmpty()) return source;
    QMap<QString,QString> fields;
    const auto types=sqliteColumnTypes(db,table);
    for(auto it=types.cbegin();it!=types.cend();++it) if(it.value()=="CURRENCY_REAL") {
        fields.insert(it.key(),it.value()); fields.insert(table.toCaseFolded()+'.'+it.key(),it.value());
        if(!alias.isEmpty()) fields.insert(alias.toCaseFolded()+'.'+it.key(),it.value());
    }
    if(fields.isEmpty()) return source;
    auto resolved=compileCurrencyQuery(source.source,fields); resolved.currencyContextResolved=true; return resolved;
}
static bool checkedSqlStorageType(const QString& type) {
    static const QSet<QString> types{"CURRENCY_REAL","INTEGER1","INTEGER2","INTEGER4","BOOL","BOOLEAN","DATE","DATETIME"};
    return types.contains(type);
}
static QString checkedSqlStorageValue(const QString& type,const QString& value) {
    if(type=="CURRENCY_REAL") return "linaccess_jet_currency_store("+value+')';
    if(checkedSqlStorageType(type)) return "linaccess_jet_field_store("+literal(type)+','+value+')';
    return value;
}
static QString currencyParameterHeader(const AccessQueryPlan& plan) {
    QStringList declarations;
    for(const auto& parameter:plan.parameters) declarations.append(renderJetTokens({{Token::Identifier,parameter.name}})+' '+parameter.type);
    return declarations.isEmpty()?QString{}:"PARAMETERS "+declarations.join(',')+"; ";
}
static QString currencyActionExpression(const AccessQueryPlan& plan,const QVector<Token>& code,const QMap<QString,QString>& fields={}) {
    const auto compiled=compileCurrencyQuery(currencyParameterHeader(plan)+"SELECT "+renderJetTokens(code)+" AS [__la_currency_expr]",fields).sql;
    const QString prefix="SELECT ",suffix=" AS "+identifier("__la_currency_expr");
    if(!compiled.startsWith(prefix) || !compiled.endsWith(suffix)) throw VbaRuntimeError(3251,"Typed action expression shape unsupported");
    return compiled.mid(prefix.size(),compiled.size()-prefix.size()-suffix.size());
}

static QString checkedInsertSql(QSqlDatabase db,const AccessQueryPlan& plan) {
    if(plan.source.isEmpty() || plan.sql!=plan.compiledSql) return plan.sql;
    const auto tokens=tokenize(plan.source);
    const auto word=[&](qsizetype i,const char* value) { return i<tokens.size() && tokens[i].kind==Token::Word && tokens[i].value.compare(value,Qt::CaseInsensitive)==0; };
    const auto name=[&](qsizetype i) { return i<tokens.size() && (tokens[i].kind==Token::Word || tokens[i].kind==Token::Identifier); };
    qsizetype pos=0;
    if(word(0,"PARAMETERS")) { while(pos<tokens.size() && !symbol(tokens[pos],";")) ++pos; ++pos; }
    if(!word(pos,"INSERT") || !word(pos+1,"INTO") || !name(pos+2)) return plan.sql;
    const auto requested=tokens[pos+2].value; pos+=3;
    QString table;
    for(const auto& candidate:db.tables(QSql::Tables)) if(candidate.compare(requested,Qt::CaseInsensitive)==0) { table=candidate; break; }
    if(table.isEmpty()) return plan.sql;
    QStringList columns;
    if(pos<tokens.size() && symbol(tokens[pos],"(")) {
        ++pos;
        while(true) {
            if(!name(pos)) return plan.sql;
            columns.append(tokens[pos++].value);
            if(pos>=tokens.size()) return plan.sql;
            if(symbol(tokens[pos],")")) { ++pos; break; }
            if(!symbol(tokens[pos++],",")) return plan.sql;
        }
    } else {
        const auto record=db.record(table);
        for(int i=0;i<record.count();++i) columns.append(record.fieldName(i));
    }
    const auto types=sqliteColumnTypes(db,table); bool checked=false;
    for(const auto& column:columns) checked|=checkedSqlStorageType(types.value(column.toCaseFolded()));
    if(!checked) return plan.sql;
    if(word(pos,"SELECT")) {
        // Keep append inference bounded to one physical SELECT source. The CTE
        // fixes positional labels (including duplicate output names) and freezes
        // self-appends before conversion/insertion starts.
        qsizetype from=-1; int depth=0;
        for(qsizetype i=pos+1;i<tokens.size();++i) {
            if(word(i,"SELECT") || word(i,"TOP") || word(i,"JOIN") || word(i,"UNION") || word(i,"INTERSECT") || word(i,"EXCEPT") || word(i,"HAVING")) return plan.sql;
            if(symbol(tokens[i],"(")) ++depth;
            else if(symbol(tokens[i],")")) --depth;
            else if(depth==0 && word(i,"FROM")) { if(from>=0) return plan.sql; from=i; }
            if(depth<0 || (symbol(tokens[i],";") && i!=tokens.size()-1)) return plan.sql;
        }
        if(depth!=0) return plan.sql;
        if(from>=0) {
            if(!name(from+1) || !db.tables(QSql::Tables).contains(tokens[from+1].value,Qt::CaseInsensitive)) return plan.sql;
            const auto boundary=[&](qsizetype i) { return i>=tokens.size() || word(i,"WHERE") || word(i,"GROUP") || word(i,"ORDER") || symbol(tokens[i],";"); };
            qsizetype next=from+2;
            if(word(next,"AS")) { ++next; if(!name(next)) return plan.sql; ++next; }
            else if(!boundary(next) && name(next)) ++next;
            if(!boundary(next)) return plan.sql;
        }
        auto select=resolveCurrencyQuery(db,compileAccessQuery(currencyParameterHeader(plan)+renderJetTokens(tokens.mid(pos))));
        if(accessQueryColumns(db,select).size()!=columns.size()) return plan.sql;
        auto source=resolveQueryOrigins(db,select.sql).trimmed(); if(source.endsWith(';')) source.chop(1);
        QString cte="__la_typed_append";
        while(std::any_of(tokens.cbegin(),tokens.cend(),[&](const Token& token) { return token.value.compare(cte,Qt::CaseInsensitive)==0; })) cte+='_';
        QStringList targets,labels,values;
        for(qsizetype i=0;i<columns.size();++i) {
            targets.append(identifier(columns[i])); const auto label=identifier("c"+QString::number(i)); labels.append(label);
            values.append(checkedSqlStorageValue(types.value(columns[i].toCaseFolded()),label));
        }
        installJetSqlFunctions(db);
        return "WITH "+identifier(cte)+'('+labels.join(',')+") AS MATERIALIZED ("+source+") INSERT INTO "+identifier(table)+'('+targets.join(',')+") SELECT "+values.join(',')+" FROM "+identifier(cte);
    }
    if(!word(pos,"VALUES") || pos+1>=tokens.size() || !symbol(tokens[pos+1],"(")) return plan.sql;
    pos+=2; auto start=pos; int depth=0; bool closed=false; QVector<QVector<Token>> expressions;
    for(;pos<tokens.size();++pos) {
        if(word(pos,"SELECT")) return plan.sql;
        if(symbol(tokens[pos],"(")) ++depth;
        else if(symbol(tokens[pos],")")) {
            if(depth==0) { if(pos==start) return plan.sql; expressions.append(tokens.mid(start,pos-start)); ++pos; closed=true; break; }
            --depth;
        } else if(depth==0 && symbol(tokens[pos],",")) {
            if(pos==start) return plan.sql;
            expressions.append(tokens.mid(start,pos-start)); start=pos+1;
        }
    }
    if(pos<tokens.size() && symbol(tokens[pos],";")) ++pos;
    if(!closed || pos!=tokens.size() || expressions.size()!=columns.size()) return plan.sql;
    QStringList targets,values;
    for(qsizetype i=0;i<columns.size();++i) {
        targets.append(identifier(columns[i])); auto value=currencyActionExpression(plan,expressions[i]);
        value=checkedSqlStorageValue(types.value(columns[i].toCaseFolded()),value);
        values.append(value);
    }
    installJetSqlFunctions(db);
    return "INSERT INTO "+identifier(table)+'('+targets.join(',')+") VALUES ("+values.join(',')+')';
}

// Only direct physical-table UPDATEs enter this path. Joined actions retain
// their existing materialized-source/identity handling.
static QString checkedUpdateSql(QSqlDatabase db,const AccessQueryPlan& plan) {
    if(plan.source.isEmpty() || plan.sql!=plan.compiledSql) return plan.sql;
    const auto tokens=tokenize(plan.source);
    const auto word=[&](qsizetype i,const char* value) { return i<tokens.size() && tokens[i].kind==Token::Word && tokens[i].value.compare(value,Qt::CaseInsensitive)==0; };
    const auto name=[&](qsizetype i) { return i<tokens.size() && (tokens[i].kind==Token::Word || tokens[i].kind==Token::Identifier); };
    qsizetype pos=0;
    if(word(0,"PARAMETERS")) { while(pos<tokens.size() && !symbol(tokens[pos],";")) ++pos; ++pos; }
    if(!word(pos,"UPDATE") || !name(pos+1)) return plan.sql;
    const auto requested=tokens[++pos].value; ++pos;
    QString alias;
    if(word(pos,"AS")) { ++pos; if(!name(pos)) return plan.sql; alias=tokens[pos++].value; }
    else if(!word(pos,"SET") && name(pos)) alias=tokens[pos++].value;
    if(!word(pos,"SET")) return plan.sql;
    const auto start=++pos;
    qsizetype end=tokens.size(),where=-1; int depth=0;
    QVector<qsizetype> separators;
    for(qsizetype i=start;i<tokens.size();++i) {
        if(word(i,"SELECT") || word(i,"FROM") || word(i,"RETURNING")) return plan.sql;
        if(symbol(tokens[i],"(")) ++depth;
        else if(symbol(tokens[i],")")) --depth;
        else if(depth==0 && word(i,"WHERE")) { if(where>=0) return plan.sql; where=i; }
        else if(depth==0 && symbol(tokens[i],",") && where<0) separators.append(i);
        else if(depth==0 && symbol(tokens[i],";")) { if(i!=tokens.size()-1) return plan.sql; end=i; }
        if(depth<0) return plan.sql;
    }
    if(depth!=0) return plan.sql;
    QString table;
    for(const auto& candidate:db.tables(QSql::Tables)) if(candidate.compare(requested,Qt::CaseInsensitive)==0) { table=candidate; break; }
    if(table.isEmpty()) return plan.sql;
    const auto types=sqliteColumnTypes(db,table);
    QMap<QString,QString> fields;
    for(auto it=types.cbegin();it!=types.cend();++it) if(it.value()=="CURRENCY_REAL") {
        fields.insert(it.key(),it.value()); fields.insert(table.toCaseFolded()+'.'+it.key(),it.value());
        if(!alias.isEmpty()) fields.insert(alias.toCaseFolded()+'.'+it.key(),it.value());
    }
    struct Assignment { QString column; QVector<Token> expression; QString targetType; };
    QVector<Assignment> assignments; bool checkedTarget=false;
    separators.append(where>=0 ? where : end);
    for(const auto stop:separators) {
        QString column;
        if(name(pos) && pos+1<stop && symbol(tokens[pos+1],"=")) { column=tokens[pos].value; pos+=2; }
        else if(name(pos) && pos+3<stop && symbol(tokens[pos+1],".") && name(pos+2) && symbol(tokens[pos+3],"=")) {
            const auto qualifier=tokens[pos].value;
            if(qualifier.compare(alias.isEmpty()?table:alias,Qt::CaseInsensitive)!=0) return plan.sql;
            column=tokens[pos+2].value; pos+=4;
        } else return plan.sql;
        if(pos>=stop) return plan.sql;
        const auto targetType=types.value(column.toCaseFolded());
        checkedTarget|=checkedSqlStorageType(targetType); assignments.append({column,tokens.mid(pos,stop-pos),targetType}); pos=stop+1;
    }
    if(!checkedTarget) return plan.sql;
    const auto expression=[&](const QVector<Token>& code) { return currencyActionExpression(plan,code,fields); };
    QStringList set;
    for(const auto& assignment:assignments) {
        auto value=expression(assignment.expression);
        value=checkedSqlStorageValue(assignment.targetType,value);
        set.append(identifier(assignment.column)+'='+value);
    }
    auto result="UPDATE "+identifier(table)+(alias.isEmpty()?QString{}:" AS "+identifier(alias))+" SET "+set.join(',');
    if(where>=0) result+=" WHERE "+expression(tokens.mid(where+1,end-where-1));
    installJetSqlFunctions(db);
    return result;
}
QString accessActionSql(QSqlDatabase db,const AccessQueryPlan& plan) {
    if (plan.updateTarget.isEmpty()) {
        const auto insert=checkedInsertSql(db,plan);
        return insert==plan.sql ? checkedUpdateSql(db,plan) : insert;
    }
    if (!db.tables(QSql::Tables).contains(plan.updateTarget,Qt::CaseInsensitive)) throw VbaRuntimeError(3251,"Joined UPDATE requires a writable table target");
    if (installAccessQueryViews(db)<0) throw VbaRuntimeError(3251,"Joined UPDATE query dependencies unavailable");
    const auto record=db.record(plan.updateTarget); QString identity;
    for (const auto* candidate:{"rowid","_rowid_","oid"}) {
        bool shadowed=false; for (int i=0;i<record.count();++i) shadowed|=record.fieldName(i).compare(candidate,Qt::CaseInsensitive)==0;
        if (shadowed) continue;
        QSqlQuery probe(db); if (probe.prepare("SELECT "+identifier(plan.updateTarget)+'.'+identifier(candidate)+" FROM "+identifier(plan.updateTarget)+" LIMIT 0")) { identity=candidate; break; }
    }
    if (identity.isEmpty()) throw VbaRuntimeError(3251,"Joined UPDATE row identity unavailable");
    if (plan.updateSourceOffset<0 || plan.sql.mid(plan.updateSourceOffset,plan.updateSource.size())!=plan.updateSource) throw VbaRuntimeError(3251,"Invalid joined UPDATE plan");
    auto sql=plan.sql;
    sql.replace(plan.updateSourceOffset,plan.updateSource.size(),resolveQueryOrigins(db,plan.updateSource));
    sql.replace(identifier(plan.updateIdentity),identifier(identity));
    return sql;
}
AccessQueryPlan filterAccessQueryPlan(QSqlDatabase db,const AccessQueryPlan& untyped,const QString& criteria) {
    const auto source=resolveCurrencyQuery(db,untyped);
    if (criteria.trimmed().isEmpty()) return source;
    if (!source.sql.trimmed().startsWith("SELECT",Qt::CaseInsensitive)) throw VbaRuntimeError(3065,"Filter source requires SELECT");
    installJetSqlFunctions(db);
    const bool topOne=useTopOneBoundary(db,source);
    if (source.hiddenRank && !topOne) throw VbaRuntimeError(3251,"Filtering an inline ranked TOP source is unsupported");
    auto sql=resolveQueryOrigins(db,topOne ? source.topOneSql : useLeftJoin(db,source) ? source.rightJoinSql : source.sql).trimmed();
    while (sql.endsWith(';')) sql.chop(1);
    auto filtered=compileAccessQuery("SELECT * FROM [__la_filter_source] WHERE "+criteria);
    const auto marker="FROM "+identifier("__la_filter_source");
    filtered.sql.replace(filtered.sql.indexOf(marker),marker.size(),"FROM ("+sql+") AS "+identifier("__la_filter_source"));
    filtered.parameters=source.parameters;
    filtered.resultTypes=source.resultTypes;
    filtered.source.clear(); filtered.compiledSql.clear();
    return filtered;
}
QString accessOrderByClause(const QString& orderBy,const QStringList& columns,const QString& tableQualifier) {
    const auto tokens=tokenize(orderBy); if(tokens.isEmpty()) return {};
    QStringList parts; qsizetype pos=0;
    while(pos<tokens.size()) {
        const auto& token=tokens[pos++];
        if(token.kind!=Token::Word && token.kind!=Token::Identifier) throw VbaRuntimeError(3251,"OrderBy requires output field names");
        auto requested=token.value;
        if(pos<tokens.size() && symbol(tokens[pos],".")) {
            if(tableQualifier.isEmpty() || requested.compare(tableQualifier,Qt::CaseInsensitive)!=0 || pos+1>=tokens.size()
                || (tokens[pos+1].kind!=Token::Word && tokens[pos+1].kind!=Token::Identifier)) throw VbaRuntimeError(3251,"OrderBy qualifier unavailable");
            requested=tokens[pos+1].value; pos+=2;
        }
        QString column; int matches=0;
        for(const auto& candidate:columns) if(candidate.compare(requested,Qt::CaseInsensitive)==0) { column=candidate; ++matches; }
        if(matches!=1) throw VbaRuntimeError(3251,"OrderBy field unavailable or ambiguous");
        auto part=identifier(column);
        if(pos<tokens.size() && tokens[pos].kind==Token::Word && (tokens[pos].value.compare("ASC",Qt::CaseInsensitive)==0 || tokens[pos].value.compare("DESC",Qt::CaseInsensitive)==0)) part+=' '+tokens[pos++].value.toUpper();
        parts.append(part);
        if(pos==tokens.size()) break;
        if(!symbol(tokens[pos++],",") || pos==tokens.size()) throw VbaRuntimeError(3251,"OrderBy field list unsupported");
    }
    return " ORDER BY "+parts.join(',');
}
AccessQueryPlan sortAccessQueryPlan(QSqlDatabase db,const AccessQueryPlan& source,const QString& orderBy) {
    if(orderBy.trimmed().isEmpty()) return source;
    const auto clause=accessOrderByClause(orderBy,accessQueryColumns(db,source));
    auto sorted=filterAccessQueryPlan(db,source,"True"); sorted.sql+=clause;
    return sorted;
}
int installAccessQueryViews(QSqlDatabase db, bool freshConnection) {
    auto* driver = db.driver();
    if (freshConnection) {
        driver->setProperty("la_query_views",QStringList{});
        driver->setProperty("la_query_internal_views",QStringList{});
        driver->setProperty("la_query_rank_pending",false);
        driver->setProperty("la_query_revision",QString{});
    }
    QSqlQuery query(db);
    if (!query.exec("SELECT name,source_bytes FROM _la_access_artifacts WHERE kind='query' ORDER BY id")) return 0;
    QMap<QString,QPair<QString,QString>> definitions;
    QCryptographicHash revision(QCryptographicHash::Sha256);
    while (query.next()) {
        const auto name = query.value(0).toString();
        const auto bytes = query.value(1).toByteArray();
        definitions.insert(name.toCaseFolded(),{name,decodeVbaSource(bytes)});
        const auto encodedName = name.toUtf8();
        revision.addData(QByteArray::number(encodedName.size()) + ':' + encodedName);
        revision.addData(QByteArray::number(bytes.size()) + ':' + bytes);
    }
    query.finish();
    const auto key = db.databaseName() + ':' + QString::fromLatin1(revision.result().toHex());
    auto owned = driver->property("la_query_views").toStringList();
    auto internal = driver->property("la_query_internal_views").toStringList();
    const auto existingViews = db.tables(QSql::Views);
    bool present = true;
    for (const auto& name : owned+internal) present &= existingViews.contains(name,Qt::CaseInsensitive);
    if (present && !driver->property("la_query_rank_pending").toBool() && driver->property("la_query_revision").toString() == key) return owned.size();
    installJetSqlFunctions(db);
    if (!query.exec("SAVEPOINT la_query_views")) return -1;
    const auto abort = [&] {
        query.exec("ROLLBACK TO la_query_views"); query.exec("RELEASE la_query_views"); return -1;
    };
    if (driver->property("la_query_database").toString() == db.databaseName())
        for (const auto& name : owned+internal)
            if (!query.exec("DROP VIEW IF EXISTS temp." + identifier(name))) return abort();
    owned.clear(); internal.clear();
    struct RankedView { QString name; QString raw; int limit; AccessQueryPlan plan; };
    QVector<RankedView> pending;
    QMap<QString,QString> ordinarySql;
    const auto tables = db.tables(QSql::Tables);
    const auto views = db.tables(QSql::Views);
    for (auto it = definitions.cbegin(); it != definitions.cend(); ++it) {
        const auto& name = it.value().first;
        if (tables.contains(name,Qt::CaseInsensitive) || views.contains(name,Qt::CaseInsensitive)) continue;
        try {
            auto plan = compileAccessQuery(it.value().second);
            if (useLeftJoin(db,plan)) plan.sql=plan.rightJoinSql;
            if (!plan.parameters.empty() || !plan.sql.startsWith("SELECT ",Qt::CaseInsensitive)) continue;
            if (plan.hiddenRank) {
                const auto raw="__la_top_"+QString::fromLatin1(QCryptographicHash::hash(name.toUtf8(),QCryptographicHash::Sha256).toHex());
                if (tables.contains(raw,Qt::CaseInsensitive) || views.contains(raw,Qt::CaseInsensitive)) return abort();
                if (query.exec("CREATE TEMP VIEW "+identifier(raw)+" AS "+plan.sql)) {
                    internal.append(raw); pending.append({name,raw,plan.topLimit,plan});
                }
            } else if (query.exec("CREATE TEMP VIEW " + identifier(name) + " AS " + plan.sql)) {
                owned << name; ordinarySql.insert(name,plan.sql);
            }
        } catch (const VbaRuntimeError&) { /* Source remains stored and explicitly unsupported. */ }
    }
    // Ordinary views can forward-reference a ranked query. Resolve ranked
    // dependency chains after all raw/ordinary definitions exist. LIMIT 0 reads
    // schema only, without running expressions or leaking a hidden rank column.
    const auto repairOrdinary=[&]() {
        for (int pass=0;pass<ordinarySql.size();++pass) {
            bool changed=false;
            for (auto it=ordinarySql.begin();it!=ordinarySql.end();++it) {
                const auto resolved=resolveQueryOrigins(db,it.value());
                if (resolved==it.value()) continue;
                if (!query.exec("DROP VIEW temp."+identifier(it.key())) || !query.exec("CREATE TEMP VIEW "+identifier(it.key())+" AS "+resolved)) return false;
                it.value()=resolved; changed=true;
            }
            if (!changed) break;
        }
        return true;
    };
    if (!repairOrdinary()) return abort();
    bool progressed=true;
    while (progressed && !pending.empty()) {
        progressed=false;
        if (!repairOrdinary()) return abort();
        for (qsizetype i=0;i<pending.size();) {
            const auto item=pending[i];
            if (useTopOneBoundary(db,item.plan)) {
                if (!query.exec("CREATE TEMP VIEW "+identifier(item.name)+" AS "+item.plan.topOneSql) ||
                    !query.exec("DROP VIEW temp."+identifier(item.raw))) return abort();
                internal.removeAll(item.raw); owned.append(item.name); pending.removeAt(i); progressed=true; continue;
            }
            QSqlQuery metadata(db);
            if (missingSourceDependency(db,item.raw) || !metadata.exec("SELECT * FROM "+identifier(item.raw)+" LIMIT 0")) { ++i; continue; }
            const auto record=metadata.record(); metadata.finish();
            if (record.count()<2) { ++i; continue; }
            QStringList columns;
            for (int column=0;column<record.count()-1;++column) columns.append(identifier(record.fieldName(column)));
            const auto sql="CREATE TEMP VIEW "+identifier(item.name)+" AS SELECT "+columns.join(',')+
                " FROM "+identifier(item.raw)+" WHERE "+identifier(record.fieldName(record.count()-1))+"<="+QString::number(item.limit);
            if (!query.exec(sql)) return abort();
            owned.append(item.name); pending.removeAt(i); progressed=true;
        }
    }
    if (!query.exec("RELEASE la_query_views")) return abort();
    driver->setProperty("la_query_database",db.databaseName());
    driver->setProperty("la_query_views",owned);
    driver->setProperty("la_query_internal_views",internal);
    driver->setProperty("la_query_rank_pending",!pending.empty());
    driver->setProperty("la_query_revision",key);
    return owned.size();
}
QString accessQuerySource(QSqlDatabase db, const QString& name) {
    QSqlQuery query(db);
    query.prepare("SELECT source_bytes FROM _la_access_artifacts WHERE kind='query' AND name=? COLLATE NOCASE ORDER BY id DESC LIMIT 1");
    query.addBindValue(name);
    if (!query.exec() || !query.next()) throw VbaRuntimeError(3265,"DAO: QueryDef not found");
    return decodeVbaSource(query.value(0).toByteArray());
}
QMap<QString,QVariant> bindAccessQueryParameters(const AccessQueryPlan& plan, const QMap<QString,QVariant>& values) {
    QMap<QString,QVariant> result;
    for (const auto& parameter : plan.parameters) {
        const auto found = values.constFind(parameter.name.toCaseFolded());
        if (found == values.cend()) throw VbaRuntimeError(3061,"DAO: required query parameter is unset");
        const auto& value = found.value();
        if (value.isNull()) {
            if (parameter.direct) result.insert(parameter.binding,{});
            if (parameter.likePattern) result.insert(parameter.binding + "_like",{});
            continue;
        }
        QVariant converted = value;
        const auto& type = parameter.type;
        if(type=="currency") {
            if(parameter.likePattern) throw VbaRuntimeError(3251,"Currency LIKE parameter conversion is not implemented");
            converted=sqlFieldValue(qtFieldValue(value,{}),"CURRENCY_REAL");
        } else if (type=="value") {
            const auto kind=value.metaType().id();
            if (kind==QMetaType::Bool) converted=value.toBool() ? -1 : 0;
            else if (kind==QMetaType::QDateTime) {
                const auto date=value.toDateTime();
                if (!date.isValid()) throw VbaRuntimeError(13,"DAO: parameter date invalid");
                converted=date.toString("yyyy-MM-dd HH:mm:ss");
            } else if (kind==QMetaType::QDate) {
                const auto date=value.toDate();
                if (!date.isValid()) throw VbaRuntimeError(13,"DAO: parameter date invalid");
                converted=date.toString("yyyy-MM-dd");
            } else if (kind==QMetaType::QString || kind==QMetaType::QByteArray || kind==QMetaType::Int || kind==QMetaType::UInt || kind==QMetaType::LongLong) converted=value;
            else if (kind==QMetaType::ULongLong) {
                if (value.toULongLong()>static_cast<qulonglong>(std::numeric_limits<qlonglong>::max())) throw VbaRuntimeError(6,"DAO: parameter integer overflow");
                converted=static_cast<qlonglong>(value.toULongLong());
            } else if (kind==QMetaType::Float || kind==QMetaType::Double) {
                if (!std::isfinite(value.toDouble())) throw VbaRuntimeError(6,"DAO: parameter number overflow");
                converted=value;
            } else throw VbaRuntimeError(13,"DAO: parameter scalar type required");
        } else if (type == "long" || type == "integer" || type == "short" || type == "byte") {
            bool valid = false; const auto real = value.toDouble(&valid);
            const auto floor = std::floor(real); const auto fraction = real - floor;
            const auto rounded = fraction < 0.5 ? floor : fraction > 0.5 ? floor+1
                : std::fmod(std::abs(floor),2.0) == 0 ? floor : floor+1;
            const qlonglong minimum = type == "byte" ? 0 : (type == "long" || type == "integer") ? -2147483648LL : -32768;
            const qlonglong maximum = type == "byte" ? 255 : (type == "long" || type == "integer") ? 2147483647LL : 32767;
            if (!valid || !std::isfinite(rounded) || rounded < minimum || rounded > maximum) throw VbaRuntimeError(13,"DAO: parameter type mismatch");
            converted = static_cast<qlonglong>(rounded);
        } else if (type == "single" || type == "double") {
            bool valid = false; const auto number = value.toDouble(&valid);
            if (!valid || !std::isfinite(number)) throw VbaRuntimeError(13,"DAO: parameter type mismatch");
            if (type=="single") {
                if (std::abs(number)>std::numeric_limits<float>::max()) throw VbaRuntimeError(13,"DAO: Single parameter overflow");
                if (number!=0 && std::abs(number)<std::numeric_limits<float>::denorm_min()) throw VbaRuntimeError(3251,"Single parameter underflow is not validated");
                converted=static_cast<double>(static_cast<float>(number));
            } else converted=number;
        } else if (type == "datetime" || type == "date") converted = canonicalDate(value.toString());
        else if (type == "boolean" || type == "yesno") {
            const auto spelling = value.toString().trimmed();
            if (spelling.compare("true",Qt::CaseInsensitive)==0) converted=-1;
            else if (spelling.compare("false",Qt::CaseInsensitive)==0) converted=0;
            else {
                bool valid=false;
                const auto number=value.toDouble(&valid);
                if (!valid || !std::isfinite(number)) throw VbaRuntimeError(13,"DAO: parameter type mismatch");
                converted=number!=0 ? -1 : 0;
            }
        }
        else converted = value.toString();
        if (parameter.direct) result.insert(parameter.binding,converted);
        if (parameter.likePattern) result.insert(parameter.binding + "_like",pattern(converted.toString()));
    }
    return result;
}
[[noreturn]] static void failAccessQuery(const QSqlQuery& query) {
        const auto error = query.lastError().text().toCaseFolded();
        if (const auto code=jetSqlFunctionErrorNumber(error.toUtf8().toStdString()))
            throw VbaRuntimeError(*code,"Jet SQL function argument invalid");
        if (error.contains("no such function")) throw VbaRuntimeError(3085,"SQL function unsupported");
        if (error.contains("no such table")) throw VbaRuntimeError(3078,"SQL dependency unavailable");
        if (error.contains("interrupted")) throw VbaRuntimeError(3035,"Query execution budget exceeded");
        if ((error.contains("syntax error") || error.contains("incomplete input")) || error.contains("circularly defined"))
            throw VbaRuntimeError(3075,"SQL grammar unsupported");
        throw VbaRuntimeError(3146,"SQL operation failed");
}
namespace {
struct PreparedAccessQuery { QSqlQuery query; bool hiddenRank; QVector<QString> types; };
PreparedAccessQuery prepareAccessQuery(QSqlDatabase db,const AccessQueryPlan& untyped) {
    const auto plan=resolveCurrencyQuery(db,untyped);
    if (!plan.sql.trimmed().startsWith("SELECT",Qt::CaseInsensitive))
        throw VbaRuntimeError(3065,"Query result requires SELECT");
    const bool topOne=useTopOneBoundary(db,plan);
    const auto source=topOne ? plan.topOneSql : useLeftJoin(db,plan) ? plan.rightJoinSql : plan.sql;
    if (missingSqlDependency(db,source)) throw VbaRuntimeError(3078,"SQL dependency unavailable");
    QSqlQuery query(db);
    if (!query.prepare(resolveQueryOrigins(db,source))) failAccessQuery(query);
    return {std::move(query),plan.hiddenRank && !topOne,plan.resultTypes};
}
}
QStringList accessQueryColumns(QSqlDatabase db,const AccessQueryPlan& plan) {
    SqliteExecutionBudget budget(db);
    const auto databaseHandle=db.driver()->handle();
    if (!databaseHandle.isValid() || QByteArray(databaseHandle.typeName())!="sqlite3*")
        throw VbaRuntimeError(3251,"SQLite metadata connection unavailable");
    auto* connection=*static_cast<sqlite3* const*>(databaseHandle.constData());
    int previousDqs=0;
    if (!connection || sqlite3_db_config(connection,SQLITE_DBCONFIG_DQS_DML,-1,&previousDqs)!=SQLITE_OK
        || sqlite3_db_config(connection,SQLITE_DBCONFIG_DQS_DML,0,nullptr)!=SQLITE_OK)
        throw VbaRuntimeError(3251,"Strict query metadata unavailable");
    const auto restore=qScopeGuard([&] { sqlite3_db_config(connection,SQLITE_DBCONFIG_DQS_DML,previousDqs,nullptr); });
    auto prepared=prepareAccessQuery(db,plan);
    const auto handle=prepared.query.result()->handle();
    if (!handle.isValid() || QByteArray(handle.typeName())!="sqlite3_stmt*")
        throw VbaRuntimeError(3251,"SQLite statement metadata unavailable");
    auto* statement=*static_cast<sqlite3_stmt* const*>(handle.constData());
    if (!statement) throw VbaRuntimeError(3251,"SQLite statement unavailable");
    QStringList columns;
    const int count=sqlite3_column_count(statement)-(prepared.hiddenRank ? 1 : 0);
    for (int i=0;i<count;++i) columns.append(QString::fromUtf8(sqlite3_column_name(statement,i)));
    return columns;
}
AccessQueryResult executeAccessQueryPlan(QSqlDatabase db, const AccessQueryPlan& plan,
                                       const QMap<QString,QVariant>& bindings) {
    if (!plan.sql.trimmed().startsWith("SELECT", Qt::CaseInsensitive))
        throw VbaRuntimeError(3065,"Query result requires SELECT");
    SqliteExecutionBudget budget(db);
    auto prepared=prepareAccessQuery(db,plan);
    auto& query=prepared.query;
    for (auto it = bindings.cbegin(); it != bindings.cend(); ++it) query.bindValue(it.key(),it.value());
    if (!query.exec()) failAccessQuery(query);
    AccessQueryResult result;
    const int count = query.record().count() - (prepared.hiddenRank ? 1 : 0);
    for (int i = 0; i < count; ++i) result.columns.append(query.record().fieldName(i));
    if(prepared.types.size()==count) {
        QMap<QString,int> names;
        for(const auto& column:result.columns) ++names[column.toCaseFolded()];
        for(int i=0;i<count;++i) if(!prepared.types[i].isEmpty() && names.value(result.columns[i].toCaseFolded())==1)
            result.types.insert(result.columns[i].toCaseFolded(),prepared.types[i]);
    }
    while (query.next()) {
        if (prepared.hiddenRank && query.value(count).toLongLong() > plan.topLimit) break;
        if (result.rows.size() >= 100000) throw VbaRuntimeError(3035,"Query row limit exceeded");
        QVariantList row;
        for (int i = 0; i < count; ++i) row.append(query.value(i));
        result.rows.append(row);
    }
    if (query.lastError().isValid()) failAccessQuery(query);
    return result;
}
void AccessQueryModel::execute(QSqlDatabase db, const QString& source, const QMap<QString,QVariant>& parameters) {
    beginResetModel(); result_ = {}; endResetModel();
    const auto plan = compileAccessQuery(source);
    auto result = executeAccessQueryPlan(db,plan,bindAccessQueryParameters(plan,parameters));
    beginResetModel(); result_ = std::move(result); endResetModel();
}
QVariant AccessQueryModel::data(const QModelIndex& index, int role) const {
    if (!index.isValid() || index.row() < 0 || index.row() >= rowCount() ||
        index.column() < 0 || index.column() >= columnCount() ||
        (role != Qt::DisplayRole && role != Qt::EditRole)) return {};
    const auto value=result_.rows[index.row()][index.column()];
    const auto type=result_.types.value(result_.columns[index.column()].toCaseFolded());
    if(role==Qt::DisplayRole && type=="CURRENCY_REAL") {
        try { return qtFieldText(value,type); }
        catch(const VbaRuntimeError& error) { return QString("#Fehler (%1)").arg(error.number); }
    }
    return value;
}
QVariant AccessQueryModel::headerData(int section, Qt::Orientation orientation, int role) const {
    if (role != Qt::DisplayRole || section < 0) return {};
    if (orientation == Qt::Vertical) return section < rowCount() ? QVariant(section+1) : QVariant{};
    return section < columnCount() ? QVariant(result_.columns[section]) : QVariant{};
}

}
