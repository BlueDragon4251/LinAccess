#pragma once
#include "vba_expression.hpp"
#include <QString>
#include <QStringList>
#include <QRegularExpression>
namespace linaccess {
inline QString nativeMacroSetValueExpression(const QString& input) {
    const auto text=input.trimmed();if(text.isEmpty() || text.size()>4096 || !text.isValidUtf16() || text.contains(QChar(0))) throw VbaRuntimeError(5,"Invalid SetValue expression text");
    const auto expression=text.startsWith('=') ? text.mid(1).trimmed() : text;
    if(expression.isEmpty() || !validVbaExpression(expression.toUtf8().toStdString())) throw VbaRuntimeError(3251,"SetValue expression syntax unsupported");return text;
}
struct NativeMacroSetValueTarget {QString object;std::optional<QString> property;};
inline NativeMacroSetValueTarget nativeMacroSetValueTarget(const QString& input) {
    const auto text=nativeMacroSetValueExpression(input);if(text.startsWith('=')) throw VbaRuntimeError(3251,"Dynamic SetValue Item targets are not implemented");int depth=0,dot=-1;bool quoted=false,bracket=false,date=false;
    for(int index=0;index<text.size();++index) {
        const auto ch=text[index];
        if(quoted) {if(ch=='"') {if(index+1<text.size() && text[index+1]=='"') ++index;else quoted=false;}continue;}
        if(bracket) {if(ch==']') {if(index+1<text.size() && text[index+1]==']') ++index;else bracket=false;}continue;}
        if(date) {if(ch=='#') date=false;continue;}
        if(ch=='"') quoted=true;else if(ch=='[') bracket=true;else if(ch=='#') date=true;else if(ch=='(') ++depth;else if(ch==')') --depth;else if(ch=='.' && depth==0) dot=index;
    }
    static const QRegularExpression identifier("^[\\p{L}_][\\p{L}\\p{N}_]*$",QRegularExpression::UseUnicodePropertiesOption);
    static const QRegularExpression bracketed("^\\[(?:[^\\]]|\\]\\])+\\]$");
    NativeMacroSetValueTarget target{text,{}};
    if(dot>=0) {auto member=text.mid(dot+1).trimmed();if(identifier.match(member).hasMatch() || bracketed.match(member).hasMatch()) {if(member.startsWith('[')) {member=member.mid(1,member.size()-2);member.replace("]]","]");}target={text.left(dot).trimmed(),member};}}
    auto name=target.object.trimmed();if(identifier.match(name).hasMatch() || bracketed.match(name).hasMatch()) {if(name.startsWith('[')) {name=name.mid(1,name.size()-2);name.replace("]]","]");}if(!QStringList{"me","forms","reports","application","screen","currentproject","currentdata","currentdb","docmd"}.contains(name.toCaseFolded())) {name.replace("]","]]");target.object="Me!["+name+"]";}}
    return target;
}
}
