#include "access_query.hpp"
#include "jet_sql_functions.hpp"
#include "sqlite_execution.hpp"
#include <QSqlError>
#include <QSqlRecord>
#include "vba_runtime.hpp"
#include "vba_source_codec.hpp"
#include <QSqlQuery>
#include <QSqlDriver>
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
QVector<Token> tokenize(const QString& sql) {
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
        if (ch == '\'' || ch == '"' || ch == '[') {
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
            result.append({ch == '[' ? Token::Identifier : Token::String, value});
        } else if (ch.isLetter() || ch == '_') {
            const auto start = i++;
            while (i < sql.size() && (sql[i].isLetterOrNumber() || sql[i] == '_')) ++i;
            if (i<sql.size() && sql[i]=='$') ++i;
            result.append({Token::Word,sql.mid(start,i-start)});
        } else if (ch.isDigit()) {
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
QString accessLikePattern(const QString& value) { return pattern(value); }
AccessQueryPlan compileAccessQuery(const QString& source) {
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
            const auto type = tokens[pos++].value.toCaseFolded();
            const QSet<QString> types{"long","integer","short","byte","single","double","text","varchar","char","boolean","yesno","datetime","date","value"};
            require(types.contains(type), "Jet SQL: parameter type unsupported");
            if (pos < tokens.size() && tokens[pos].value == "(") {
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
    adaptObjectReferences(tokens,pos);
    adaptConcatenation(tokens,pos);
    adaptDynamicLike(tokens,pos,parameters);
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
            require(from && single,"Jet SQL: multitable DISTINCTROW requires record identity mapping");
            tokens.removeAt(modifier); // Single-table DISTINCTROW is ignored by Jet; DISTINCT would be wrong.
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
        else if (token.kind == Token::Word && i + 1 < tokens.size() && tokens[i+1].value == "("
                 && (token.value.compare("IIf",Qt::CaseInsensitive)==0 || token.value.compare("First",Qt::CaseInsensitive)==0
                     || token.value.compare("Last",Qt::CaseInsensitive)==0)) output << "linaccess_jet_" + token.value.toCaseFolded();
        else if (token.kind == Token::Word && token.value.compare("TRUE",Qt::CaseInsensitive) == 0) output << "-1";
        else if (token.kind == Token::Word && token.value.compare("FALSE",Qt::CaseInsensitive) == 0) output << "0";
        else output << token.value;
    }
    plan.sql = output.join(' ');
    return plan;
}
int installAccessQueryViews(QSqlDatabase db, bool freshConnection) {
    auto* driver = db.driver();
    if (freshConnection) {
        driver->setProperty("la_query_views",QStringList{});
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
    const auto existingViews = db.tables(QSql::Views);
    bool present = true;
    for (const auto& name : owned) present &= existingViews.contains(name,Qt::CaseInsensitive);
    if (present && driver->property("la_query_revision").toString() == key) return owned.size();
    if (!query.exec("SAVEPOINT la_query_views")) return -1;
    const auto abort = [&] {
        query.exec("ROLLBACK TO la_query_views"); query.exec("RELEASE la_query_views"); return -1;
    };
    if (driver->property("la_query_database").toString() == db.databaseName())
        for (const auto& name : owned)
            if (!query.exec("DROP VIEW IF EXISTS temp." + identifier(name))) return abort();
    owned.clear();
    const auto tables = db.tables(QSql::Tables);
    const auto views = db.tables(QSql::Views);
    for (auto it = definitions.cbegin(); it != definitions.cend(); ++it) {
        const auto& name = it.value().first;
        if (tables.contains(name,Qt::CaseInsensitive) || views.contains(name,Qt::CaseInsensitive)) continue;
        try {
            const auto plan = compileAccessQuery(it.value().second);
            if (!plan.parameters.empty() || plan.hiddenRank || !plan.sql.startsWith("SELECT ",Qt::CaseInsensitive)) continue;
            if (query.exec("CREATE TEMP VIEW " + identifier(name) + " AS " + plan.sql)) owned << name;
        } catch (const VbaRuntimeError&) { /* Source remains stored and explicitly unsupported. */ }
    }
    if (!query.exec("RELEASE la_query_views")) return abort();
    driver->setProperty("la_query_database",db.databaseName());
    driver->setProperty("la_query_views",owned);
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
        if (type=="value") {
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
            const qlonglong minimum = type == "byte" ? 0 : type == "long" ? -2147483648LL : -32768;
            const qlonglong maximum = type == "byte" ? 255 : type == "long" ? 2147483647LL : 32767;
            if (!valid || !std::isfinite(rounded) || rounded < minimum || rounded > maximum) throw VbaRuntimeError(13,"DAO: parameter type mismatch");
            converted = static_cast<qlonglong>(rounded);
        } else if (type == "single" || type == "double") {
            bool valid = false; converted = value.toDouble(&valid);
            if (!valid) throw VbaRuntimeError(13,"DAO: parameter type mismatch");
        } else if (type == "datetime" || type == "date") converted = canonicalDate(value.toString());
        else if (type == "boolean" || type == "yesno") converted = value.toBool() ? -1 : 0;
        else converted = value.toString();
        if (parameter.direct) result.insert(parameter.binding,converted);
        if (parameter.likePattern) result.insert(parameter.binding + "_like",pattern(converted.toString()));
    }
    return result;
}
AccessQueryResult executeAccessQueryPlan(QSqlDatabase db, const AccessQueryPlan& plan,
                                       const QMap<QString,QVariant>& bindings) {
    if (!plan.sql.trimmed().startsWith("SELECT", Qt::CaseInsensitive))
        throw VbaRuntimeError(3065,"Query result requires SELECT");
    SqliteExecutionBudget budget(db);
    QSqlQuery query(db);
    const auto fail = [&]() {
        const auto error = query.lastError().text().toCaseFolded();
        if (const auto code=jetSqlFunctionErrorNumber(error.toUtf8().toStdString()))
            throw VbaRuntimeError(*code,"Jet SQL function argument invalid");
        if (error.contains("no such function")) throw VbaRuntimeError(3085,"SQL function unsupported");
        if (error.contains("no such table")) throw VbaRuntimeError(3078,"SQL dependency unavailable");
        if (error.contains("interrupted")) throw VbaRuntimeError(3035,"Query execution budget exceeded");
        if (error.contains("syntax error") || error.contains("circularly defined"))
            throw VbaRuntimeError(3075,"SQL grammar unsupported");
        throw VbaRuntimeError(3146,"SQL operation failed");
    };
    if (!query.prepare(plan.sql)) fail();
    for (auto it = bindings.cbegin(); it != bindings.cend(); ++it) query.bindValue(it.key(),it.value());
    if (!query.exec()) fail();
    AccessQueryResult result;
    const int count = query.record().count() - (plan.hiddenRank ? 1 : 0);
    for (int i = 0; i < count; ++i) result.columns.append(query.record().fieldName(i));
    while (query.next()) {
        if (plan.hiddenRank && query.value(count).toLongLong() > plan.topLimit) break;
        if (result.rows.size() >= 100000) throw VbaRuntimeError(3035,"Query row limit exceeded");
        QVariantList row;
        for (int i = 0; i < count; ++i) row.append(query.value(i));
        result.rows.append(row);
    }
    if (query.lastError().isValid()) fail();
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
    return result_.rows[index.row()][index.column()];
}
QVariant AccessQueryModel::headerData(int section, Qt::Orientation orientation, int role) const {
    if (role != Qt::DisplayRole || section < 0) return {};
    if (orientation == Qt::Vertical) return section < rowCount() ? QVariant(section+1) : QVariant{};
    return section < columnCount() ? QVariant(result_.columns[section]) : QVariant{};
}

}
