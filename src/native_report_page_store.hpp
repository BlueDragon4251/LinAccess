#pragma once
#include "vba_runtime.hpp"
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QPageLayout>
#include <QPageSize>
#include <optional>
#include <cmath>

namespace linaccess::report_page_store {
struct Definition { QString name,raw; QJsonObject object; };
inline std::optional<Definition> definition(QSqlDatabase database,const QString& name) {
    if(!database.isOpen()) return {};
    QSqlQuery query(database);
    if(!query.exec("SELECT 1 FROM main.sqlite_schema WHERE type='table' AND name='_la_objects'"))
        throw VbaRuntimeError(3146,"Report design catalog unavailable");
    const bool catalog=query.next(); query.finish(); if(!catalog) return {};
    if(!query.exec("SELECT name,definition FROM main._la_objects WHERE kind='report'"))
        throw VbaRuntimeError(3146,"Report design lookup failed");
    std::optional<Definition> result; int rows=0;
    while(query.next()) {
        if(++rows>1000000) throw VbaRuntimeError(3251,"Report design catalog limit exceeded");
        if(query.value(0).toString().compare(name,Qt::CaseInsensitive)!=0) continue;
        if(result) throw VbaRuntimeError(5,"Report design name ambiguous");
        const auto raw=query.value(1).toString(); QJsonParseError error;
        const auto document=QJsonDocument::fromJson(raw.toUtf8(),&error);
        if(error.error!=QJsonParseError::NoError || !document.isObject())
            throw VbaRuntimeError(3251,"Report design metadata invalid");
        result=Definition{query.value(0).toString(),raw,document.object()};
    }
    return result;
}
inline std::optional<QPageLayout> load(QSqlDatabase database,const QString& name) {
    const auto row=definition(database,name); if(!row) return {};
    const auto value=row->object.value("page_setup"); if(value.isUndefined()) return {};
    if(!value.isObject()) throw VbaRuntimeError(3251,"Report page setup invalid");
    const auto settings=value.toObject();
    const auto number=[](const QJsonValue& value,double minimum,double maximum) {
        const double result=value.toDouble();
        if(!value.isDouble() || !std::isfinite(result) || result<minimum || result>maximum)
            throw VbaRuntimeError(3251,"Report page geometry invalid");
        return result;
    };
    if(number(settings.value("version"),1,1)!=1) throw VbaRuntimeError(3251,"Report page setup version unsupported");
    const auto width=number(settings.value("width_mm"),0,5000),height=number(settings.value("height_mm"),0,5000);
    if(width<=0 || height<=0) throw VbaRuntimeError(3251,"Report paper dimensions invalid");
    const auto orientation=number(settings.value("orientation"),0,1);
    if(orientation!=std::floor(orientation)) throw VbaRuntimeError(3251,"Report orientation invalid");
    const auto margins=settings.value("margins_mm");
    if(!margins.isArray() || margins.toArray().size()!=4) throw VbaRuntimeError(3251,"Report margins invalid");
    const auto edges=margins.toArray();
    const QPageSize paper(QSizeF(width,height),QPageSize::Millimeter,{},QPageSize::ExactMatch);
    QPageLayout layout(paper,orientation==1 ? QPageLayout::Landscape : QPageLayout::Portrait,QMarginsF(),QPageLayout::Millimeter);
    if(!layout.setMargins(QMarginsF(number(edges[0],0,5000),number(edges[1],0,5000),number(edges[2],0,5000),number(edges[3],0,5000)))
        || !layout.isValid() || layout.paintRect(QPageLayout::Millimeter).isEmpty())
        throw VbaRuntimeError(3251,"Report printable area invalid");
    return layout;
}
inline void save(QSqlDatabase database,const QString& name,const QPageLayout& layout) {
    const auto row=definition(database,name);
    if(!row) throw VbaRuntimeError(3251,"Only saved native report designs can store page setup");
    const auto size=layout.pageSize().size(QPageSize::Millimeter); const auto margins=layout.margins(QPageLayout::Millimeter);
    auto object=row->object;
    object.insert("page_setup",QJsonObject{{"version",1},{"width_mm",size.width()},{"height_mm",size.height()},
        {"orientation",layout.orientation()==QPageLayout::Landscape ? 1 : 0},
        {"margins_mm",QJsonArray{margins.left(),margins.top(),margins.right(),margins.bottom()}}});
    QSqlQuery query(database);
    if(!query.prepare("UPDATE main._la_objects SET definition=? WHERE kind='report' AND name=? AND definition=?"))
        throw VbaRuntimeError(3146,"Report page setup update unavailable");
    query.addBindValue(QString::fromUtf8(QJsonDocument(object).toJson(QJsonDocument::Compact)));
    query.addBindValue(row->name); query.addBindValue(row->raw);
    if(!query.exec()) throw VbaRuntimeError(3146,"Report page setup could not be saved");
    if(query.numRowsAffected()!=1) throw VbaRuntimeError(3251,"Report design changed; reopen page setup");
}
}
