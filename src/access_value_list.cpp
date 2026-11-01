#include "access_value_list.hpp"
#include "vba_runtime.hpp"
namespace linaccess {
AccessQueryResult parseAccessValueList(const QString& source,int columns) {
    if (columns<1 || columns>255) throw VbaRuntimeError(5,"Invalid value-list column count");
    AccessQueryResult result;
    for (int i=0;i<columns;++i) result.columns.append("Column "+QString::number(i+1));
    if (source.isEmpty()) return result;
    QVariantList row;
    qsizetype position=0;
    while (true) {
        while (position<source.size() && source[position].isSpace()) ++position;
        const bool quoted=position<source.size() && source[position]=='"';
        QString value=QStringLiteral("");
        if (quoted) {
            ++position; bool closed=false;
            while (position<source.size()) {
                const auto ch=source[position++];
                if (ch!='"') { value+=ch; continue; }
                if (position<source.size() && source[position]=='"') { value+='"'; ++position; }
                else { closed=true; break; }
            }
            if (!closed) throw VbaRuntimeError(5,"Unterminated value-list item");
            while (position<source.size() && source[position].isSpace()) ++position;
            if (position<source.size() && source[position]!=';') throw VbaRuntimeError(5,"Invalid value-list delimiter");
        } else {
            const auto start=position;
            while (position<source.size() && source[position]!=';') {
                if (source[position]=='"') throw VbaRuntimeError(5,"Invalid value-list quote");
                ++position;
            }
            value=source.mid(start,position-start).trimmed();
        }
        row.append(!quoted && value.isEmpty() ? QVariant{} : QVariant(value));
        if (row.size()==columns || position==source.size()) {
            if (result.rows.size()>=10000) throw VbaRuntimeError(3035,"Control row limit exceeded");
            while (row.size()<columns) row.append(QVariant{});
            result.rows.append(std::move(row)); row=QVariantList{};
        }
        if (position==source.size()) break;
        ++position;
    }
    return result;
}
}
