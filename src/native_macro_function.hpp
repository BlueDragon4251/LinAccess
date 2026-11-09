#pragma once
#include "vba_expression.hpp"
#include <QString>
#include <QRegularExpression>
namespace linaccess {
// RunCode accepts exactly one unqualified Function call, never an arbitrary sum
// of expressions. The normal expression parser checks the argument grammar.
inline QString nativeMacroFunctionName(const QString& expression) {
    const auto source=expression.trimmed();
    if(source.size()>65536 || !source.isValidUtf16() || source.contains(QChar(0))) throw VbaRuntimeError(5,"Invalid RunCode expression text");
    static const QRegularExpression prefix("^([A-Za-z_][A-Za-z0-9_]*)\\s*\\(");
    const auto match=prefix.match(source);
    if(!match.hasMatch()) throw VbaRuntimeError(3251,"RunCode requires one unqualified Function call with parentheses");
    int depth=0;bool quoted=false,bracket=false,dated=false,closed=false;
    for(qsizetype index=match.capturedEnd()-1;index<source.size();++index) {
        const auto ch=source[index];
        if(quoted) {if(ch=='"') {if(index+1<source.size() && source[index+1]=='"') ++index;else quoted=false;}continue;}
        if(bracket) {if(ch==']') {if(index+1<source.size() && source[index+1]==']') ++index;else bracket=false;}continue;}
        if(dated) {if(ch=='#') dated=false;continue;}
        if(ch=='"') quoted=true;else if(ch=='[') bracket=true;else if(ch=='#') dated=true;
        else if(ch=='(') ++depth;
        else if(ch==')' && --depth==0) {closed=index==source.size()-1;break;}
    }
    if(!closed || !validVbaExpression(source.toUtf8().toStdString())) throw VbaRuntimeError(3251,"RunCode Function expression syntax unsupported");
    return match.captured(1);
}
}
