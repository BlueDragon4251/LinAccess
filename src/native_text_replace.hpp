#pragma once
#include "vba_runtime.hpp"
#include <QString>
namespace linaccess {
struct NativeTextReplacement {bool matched{};QString value;};
inline NativeTextReplacement replaceNativeText(const QString& input,const QString& search,QString replacement,int mode,Qt::CaseSensitivity sensitivity,bool all) {
    if(mode<0 || mode>2 || search.isEmpty()) throw VbaRuntimeError(5,"Invalid literal replacement setting");
    if(input.size()>16*1024*1024 || search.size()>1000000 || replacement.size()>1000000 || input.contains(QChar(0)) || search.contains(QChar(0)) || replacement.contains(QChar(0)) || !input.isValidUtf16() || !search.isValidUtf16() || !replacement.isValidUtf16()) throw VbaRuntimeError(3251,"Literal replacement text is not representable");
    if(replacement.isNull()) replacement=QStringLiteral("");
    if(mode==1) return input.compare(search,sensitivity)==0 ? NativeTextReplacement{true,replacement} : NativeTextReplacement{false,input};
    const auto first=input.indexOf(search,0,sensitivity);
    if(first<0 || (mode==2 && first!=0)) return {false,input};
    QString output=QStringLiteral("");qsizetype cursor=0,position=first;
    while(position>=0) {
        if(input.mid(position,search.size()).compare(search,sensitivity)!=0) throw VbaRuntimeError(3251,"Length-changing replacement collation is not supported");
        if(static_cast<qint64>(output.size())+position-cursor+replacement.size()>16*1024*1024) throw VbaRuntimeError(3251,"Literal replacement output exceeds its native bound");
        output+=input.mid(cursor,position-cursor);output+=replacement;cursor=position+search.size();
        if(!all || mode==2) break;
        position=input.indexOf(search,cursor,sensitivity);
    }
    if(static_cast<qint64>(output.size())+input.size()-cursor>16*1024*1024) throw VbaRuntimeError(3251,"Literal replacement output exceeds its native bound");
    output+=input.mid(cursor);if(!output.isValidUtf16()) throw VbaRuntimeError(3251,"Literal replacement would split a Unicode character");
    return {true,output};
}
}
