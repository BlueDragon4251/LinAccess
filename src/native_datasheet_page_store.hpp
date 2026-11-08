#pragma once
#include <QSizeF>
#include <QRectF>
#include "vba_runtime.hpp"
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QVariant>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QPageLayout>
#include <QPageSize>
#include <cmath>

namespace linaccess::datasheet_page_store {
struct Row { bool present{}; QString raw; };
inline QPageLayout defaults() { return {QPageSize(QPageSize::A4),QPageLayout::Portrait,QMarginsF(15,15,15,15),QPageLayout::Millimeter}; }
inline void schema(QSqlDatabase db) {
    QSqlQuery query(db);
    if(!query.exec("PRAGMA main.table_info('_la_datasheet_page_setup')")) throw VbaRuntimeError(3146,"Datasheet page catalog unavailable");
    const QStringList names{"kind","name","layout"}; int index=0;
    while(query.next()) {
        if(index>=3 || query.value(1).toString()!=names[index] || query.value(2).toString().toUpper()!="TEXT"
            || query.value(3).toInt()!=1 || query.value(5).toInt()!=(index<2 ? index+1 : 0))
            throw VbaRuntimeError(3251,"Datasheet page catalog schema invalid");
        ++index;
    }
    if(index!=3) throw VbaRuntimeError(3251,"Datasheet page catalog schema invalid");
}
inline Row read(QSqlDatabase db,const QString& kind,const QString& name) {
    QSqlQuery query(db);
    if(!query.exec("SELECT type FROM main.sqlite_schema WHERE name='_la_datasheet_page_setup'")) throw VbaRuntimeError(3146,"Datasheet page catalog unavailable");
    if(!query.next()) return {};
    if(query.value(0).toString()!="table" || query.next()) throw VbaRuntimeError(3251,"Datasheet page catalog invalid");
    query.finish(); schema(db);
    query.prepare("SELECT layout FROM main._la_datasheet_page_setup WHERE kind=? AND name=?");
    query.addBindValue(kind); query.addBindValue(name.toCaseFolded());
    if(!query.exec()) throw VbaRuntimeError(3146,"Datasheet page lookup failed");
    if(!query.next()) return {};
    Row result{true,query.value(0).toString()};
    if(result.raw.size()>8192 || query.next()) throw VbaRuntimeError(3251,"Datasheet page settings invalid");
    return result;
}
inline QPageLayout decode(const Row& row) {
    if(!row.present) return defaults();
    QJsonParseError error; const auto document=QJsonDocument::fromJson(row.raw.toUtf8(),&error);
    if(error.error!=QJsonParseError::NoError || !document.isObject()) throw VbaRuntimeError(3251,"Datasheet page settings invalid");
    const auto settings=document.object();
    const auto number=[](QJsonValue value,double minimum,double maximum) {
        const auto result=value.toDouble();
        if(!value.isDouble() || !std::isfinite(result) || result<minimum || result>maximum) throw VbaRuntimeError(3251,"Datasheet page geometry invalid");
        return result;
    };
    (void)number(settings.value("version"),1,1);
    const auto width=number(settings.value("width_mm"),0,5000),height=number(settings.value("height_mm"),0,5000);
    const auto orientation=number(settings.value("orientation"),0,1);
    const auto edges=settings.value("margins_mm").toArray();
    if(width<=0 || height<=0 || orientation!=std::floor(orientation) || edges.size()!=4) throw VbaRuntimeError(3251,"Datasheet page geometry invalid");
    QPageLayout layout(QPageSize(QSizeF(width,height),QPageSize::Millimeter,{},QPageSize::ExactMatch),
        orientation==1 ? QPageLayout::Landscape : QPageLayout::Portrait,QMarginsF(),QPageLayout::Millimeter);
    if(!layout.setMargins({number(edges[0],0,5000),number(edges[1],0,5000),number(edges[2],0,5000),number(edges[3],0,5000)})
        || !layout.isValid() || layout.paintRect(QPageLayout::Millimeter).isEmpty()) throw VbaRuntimeError(3251,"Datasheet printable area invalid");
    return layout;
}
inline void save(QSqlDatabase db,const QString& kind,const QString& name,const Row& before,const QPageLayout& layout) {
    if((kind!="table" && kind!="query") || name.isEmpty()) throw VbaRuntimeError(3251,"Only named datasheets can store page settings");
    const auto size=layout.pageSize().size(QPageSize::Millimeter); const auto margins=layout.margins(QPageLayout::Millimeter);
    const auto raw=QString::fromUtf8(QJsonDocument(QJsonObject{{"version",1},{"width_mm",size.width()},{"height_mm",size.height()},
        {"orientation",layout.orientation()==QPageLayout::Landscape ? 1 : 0},
        {"margins_mm",QJsonArray{margins.left(),margins.top(),margins.right(),margins.bottom()}}}).toJson(QJsonDocument::Compact));
    (void)decode({true,raw});
    QSqlQuery guard(db);
    if(!guard.exec("SAVEPOINT la_datasheet_page_setup")) throw VbaRuntimeError(3146,"Datasheet page save unavailable");
    try {
        if(!guard.exec("CREATE TABLE IF NOT EXISTS main._la_datasheet_page_setup(kind TEXT NOT NULL,name TEXT NOT NULL,layout TEXT NOT NULL,PRIMARY KEY(kind,name))")) throw VbaRuntimeError(3146,"Datasheet page catalog creation failed");
        schema(db); QSqlQuery query(db);
        query.prepare(before.present ? "UPDATE main._la_datasheet_page_setup SET layout=? WHERE kind=? AND name=? AND layout=?"
            : "INSERT OR IGNORE INTO main._la_datasheet_page_setup(layout,kind,name) VALUES(?,?,?)");
        query.addBindValue(raw); query.addBindValue(kind); query.addBindValue(name.toCaseFolded()); if(before.present) query.addBindValue(before.raw);
        if(!query.exec()) throw VbaRuntimeError(3146,"Datasheet page save failed");
        if(query.numRowsAffected()!=1) throw VbaRuntimeError(3251,"Datasheet page settings changed; reopen setup");
        if(read(db,kind,name).raw!=raw) throw VbaRuntimeError(3251,"Datasheet page save did not retain settings");
        if(!guard.exec("RELEASE la_datasheet_page_setup")) throw VbaRuntimeError(3146,"Datasheet page save failed");
    } catch(...) { guard.exec("ROLLBACK TO la_datasheet_page_setup"); guard.exec("RELEASE la_datasheet_page_setup"); throw; }
}
}
