#pragma once
#include "access_form_import.hpp"
#include "vba_qt_value.hpp"
#include "vba_decimal.hpp"
#include <QSet>
#include <algorithm>
#include <cmath>
#include <numeric>
#include <optional>
namespace linaccess {
struct NativeReportGroupPlan {
    struct Level {int column{};bool descending{};QString header,footer;bool repeatHeader{};};
    QVector<Level> levels;
    QVector<int> order;
    QVector<QVector<QPair<int,int>>> ranges; // Sorted row, half-open scope per level.
};
inline int nativeReportGroupCompare(const VbaValue& a,const VbaValue& b) {
    if(a.kind==VbaValue::Kind::Null || b.kind==VbaValue::Kind::Null)
        return a.kind==b.kind ? 0 : a.kind==VbaValue::Kind::Null ? -1 : 1;
    if(a.kind!=b.kind) throw VbaRuntimeError(3251,"Mixed report group key types are unsupported");
    if(a.kind==VbaValue::Kind::Text) return QString::compare(QString::fromStdString(a.text),QString::fromStdString(b.text),Qt::CaseInsensitive);
    if(a.kind==VbaValue::Kind::Currency) return (a.currencyUnits>b.currencyUnits)-(a.currencyUnits<b.currencyUnits);
    if(a.kind==VbaValue::Kind::Decimal) return decimalCompare(a,b);
    if(a.kind==VbaValue::Kind::Date) {
        const auto civil=[](double value) {if(!std::isfinite(value)) throw VbaRuntimeError(3251,"Non-finite report Date group key");const auto day=std::trunc(value);return day+std::abs(value-day);};
        const auto x=civil(a.number),y=civil(b.number);return (x>y)-(x<y);
    }
    if(a.kind==VbaValue::Kind::Number || a.kind==VbaValue::Kind::Boolean) {
        const auto x=a.asNumber(),y=b.asNumber();if(!std::isfinite(x) || !std::isfinite(y)) throw VbaRuntimeError(3251,"Non-finite report group key");return (x>y)-(x<y);
    }
    throw VbaRuntimeError(3251,"Report group key type unsupported");
}
inline std::optional<NativeReportGroupPlan> nativeReportGrouping(const AccessFormLayout& layout,const QStringList& columns,
    const QMap<QString,QString>& types,const QVector<QVariantList>& rows) {
    NativeReportGroupPlan plan;bool inactive=false;QSet<QString> sections;
    for(int index=0;index<layout.reportGroups.size();++index) {
        const auto& source=layout.reportGroups[index];const auto expression=source.controlSource.value_or(QString()).trimmed();
        if(expression.isEmpty()) {if(source.groupHeader.value_or(false) || source.groupFooter.value_or(false)) return {};inactive=true;continue;}
        // No missing source properties are replaced with guessed Access defaults.
        if(inactive || !source.sortOrder || !source.groupHeader || !source.groupFooter || !source.groupOn || !source.groupInterval || !source.keepTogether
            || *source.groupOn!=0 || *source.groupInterval!=1 || *source.keepTogether!=0) return {};
        auto field=expression;if(field.startsWith('[') && field.endsWith(']')) {field=field.mid(1,field.size()-2);field.replace("]]","]");}
        int column=-1;for(int i=0;i<columns.size();++i) if(columns[i].compare(field,Qt::CaseInsensitive)==0) {if(column>=0) return {};column=i;}
        if(column<0) return {};
        const auto type=types.value(columns[column].toCaseFolded()).toUpper();if(type=="BLOB" || type.contains("BINARY")) return {};
        NativeReportGroupPlan::Level level{column,*source.sortOrder,{},{}};
        for(const auto& band:layout.reportGroupSections) if(band.level==index) {
            if(!band.levelProven) return {};
            auto& target=band.header ? level.header : level.footer;
            if(!target.isEmpty() || !layout.sectionHeights.contains(band.section)) return {};
            if(band.header && band.repeatHeader) {
                // A repeated native header shows the representative group key.
                // Arbitrary row-dependent imported header expressions need an oracle.
                for(const auto& control:layout.controls) if(control.section==band.section && control.type=="TextBox") {
                    auto field=control.controlSource.trimmed();if(field.startsWith('[') && field.endsWith(']')) field=field.mid(1,field.size()-2);
                    if(field.compare(columns[column],Qt::CaseInsensitive)!=0 || control.runningSum || control.canGrow) return {};
                }
                level.repeatHeader=true;
            }
            target=band.section;sections.insert(band.section);
        }
        if(*source.groupHeader!=!level.header.isEmpty() || *source.groupFooter!=!level.footer.isEmpty()) return {};
        plan.levels.append(level);
    }
    if(sections.size()!=layout.reportGroupSections.size()) return {};
    if(plan.levels.isEmpty()) return plan;
    if(rows.size()>1000000/plan.levels.size()) throw VbaRuntimeError(3251,"Report group key bound exceeded");
    QVector<QVector<VbaValue>> keys;keys.reserve(rows.size());qsizetype textBytes=0;
    for(const auto& row:rows) {
        QVector<VbaValue> key;key.reserve(plan.levels.size());
        for(const auto& level:plan.levels) {
            if(level.column>=row.size() || row[level.column].metaType().id()==QMetaType::QByteArray) return {};
            auto value=qtFieldValue(row[level.column],types.value(columns[level.column].toCaseFolded()));
            if(value.kind==VbaValue::Kind::Text) {textBytes+=value.text.size();if(textBytes>16*1024*1024) throw VbaRuntimeError(3251,"Report group text bound exceeded");}
            (void)nativeReportGroupCompare(value,value);key.append(std::move(value));
        }
        keys.append(std::move(key));
    }
    plan.order.resize(rows.size());std::iota(plan.order.begin(),plan.order.end(),0);
    std::stable_sort(plan.order.begin(),plan.order.end(),[&](int a,int b) {
        for(int i=0;i<plan.levels.size();++i) {const auto order=nativeReportGroupCompare(keys[a][i],keys[b][i]);if(order) return plan.levels[i].descending ? order>0 : order<0;}return false;
    });
    plan.ranges.resize(plan.levels.size());
    for(int level=0;level<plan.levels.size();++level) {
        auto& ranges=plan.ranges[level];ranges.resize(rows.size());int begin=0;
        for(int position=1;position<=rows.size();++position) {
            bool changed=position==rows.size();
            if(!changed) for(int key=0;key<=level;++key) if(nativeReportGroupCompare(keys[plan.order[position-1]][key],keys[plan.order[position]][key])) {changed=true;break;}
            if(changed) {for(int row=begin;row<position;++row) ranges[row]={begin,position};begin=position;}
        }
    }
    return plan;
}
}
