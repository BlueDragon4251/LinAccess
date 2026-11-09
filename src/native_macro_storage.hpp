#pragma once
#include "native_macro_runner.hpp"
#include <QVariant>
namespace linaccess {
struct NativeMacroRevision {QString name;QVariant definition;};
[[noreturn]] inline void nativeMacroStorageError(const QSqlQuery& query) {
    const int code=query.lastError().nativeErrorCode().toInt()&255;
    throw VbaRuntimeError(code==8 ? 3027 : code==5 || code==6 ? 3260 : code==19 ? 3010 : 3146,"Makro-Katalog konnte nicht geschrieben werden.");
}
inline void checkNativeMacroStorage(QSqlDatabase db,bool writing) {
    if(!db.isOpen()) throw VbaRuntimeError(3420,"Makroprojekt ist geschlossen.");QSqlQuery query(db);
    if(!query.exec("SELECT type FROM main.sqlite_schema WHERE name='_la_objects' COLLATE NOCASE") || !query.next() || query.value(0).toString()!="table") throw VbaRuntimeError(3251,"Physischer Makrokatalog ist nicht verfügbar.");query.finish();
    if(!query.exec("SELECT 1 FROM temp.sqlite_schema WHERE name='_la_objects' COLLATE NOCASE AND type IN ('table','view') LIMIT 1")) throw VbaRuntimeError(3146,"Makrokatalog konnte nicht geprüft werden.");if(query.next()) throw VbaRuntimeError(3251,"Temporärer Schatten des Makrokatalogs wird nicht unterstützt.");query.finish();
    if(writing) for(const auto* schema:{"main","temp"}) {if(!query.exec(QString("SELECT 1 FROM ")+schema+".sqlite_schema WHERE type='trigger' AND tbl_name='_la_objects' COLLATE NOCASE LIMIT 1")) throw VbaRuntimeError(3146,"Makrokatalog konnte nicht geprüft werden.");if(query.next()) throw VbaRuntimeError(3251,"Katalogänderungen mit zusätzlichen Triggern werden nicht unterstützt.");query.finish();}
}
inline std::optional<NativeMacroRevision> readNativeMacroRevision(QSqlDatabase db,const QString& name) {
    checkNativeMacroStorage(db,false);if(!nativeMacroNameValid(name)) throw VbaRuntimeError(3251,"Hier ist der Name eines gespeicherten Makroobjekts erforderlich.");
    QSqlQuery query(db);query.prepare("SELECT name,length(CAST(definition AS BLOB)),CASE WHEN length(CAST(definition AS BLOB))<=1048576 THEN definition ELSE NULL END FROM main._la_objects WHERE kind='macro' AND name=? COLLATE BINARY LIMIT 2");query.addBindValue(name);
    if(!query.exec()) throw VbaRuntimeError(3146,"Makrorevision konnte nicht gelesen werden.");if(!query.next()) {if(query.lastError().isValid()) throw VbaRuntimeError(3146,"Makrorevision konnte nicht gelesen werden.");return {};}
    if(query.value(1).toLongLong()>1024*1024) throw VbaRuntimeError(3035,"Makrorevision überschreitet das Größenlimit.");NativeMacroRevision revision{query.value(0).toString(),query.value(2)};
    if(query.next()) throw VbaRuntimeError(3251,"Makroobjekt ist nicht eindeutig.");if(query.lastError().isValid()) throw VbaRuntimeError(3146,"Makrorevision konnte nicht gelesen werden.");return revision;
}
inline bool importedMacroNameExists(QSqlDatabase db,const QString& name) {
    QSqlQuery query(db);if(!query.exec("SELECT type FROM main.sqlite_schema WHERE name='_la_access_artifacts' COLLATE NOCASE")) throw VbaRuntimeError(3146,"Importkatalog konnte nicht geprüft werden.");if(!query.next()) {if(query.lastError().isValid()) throw VbaRuntimeError(3146,"Importkatalog konnte nicht gelesen werden.");return false;}if(query.value(0).toString()!="table") throw VbaRuntimeError(3251,"Physischer Importkatalog ist nicht verfügbar.");query.finish();query.setForwardOnly(true);
    if(!query.exec("SELECT name FROM main._la_access_artifacts WHERE kind='macro'")) throw VbaRuntimeError(3146,"Importierte Makronamen konnten nicht gelesen werden.");QElapsedTimer timer;timer.start();int count=0;
    while(query.next()) {if(++count>1000000 || timer.hasExpired(5000)) throw VbaRuntimeError(3035,"Importkatalog-Limit überschritten.");if(query.value(0).toString().compare(name,Qt::CaseInsensitive)==0) return true;}if(query.lastError().isValid()) throw VbaRuntimeError(3146,"Importierte Makronamen konnten nicht gelesen werden.");return false;
}
template<class Action> inline auto nativeMacroWrite(QSqlDatabase db,Action action) {
    if(!db.isOpen()) throw VbaRuntimeError(3420,"Makroprojekt ist geschlossen.");
    static thread_local quint64 sequence=0;const auto savepoint=QString("la_macro_catalog_%1").arg(++sequence);QSqlQuery query(db);if(!query.exec("SAVEPOINT "+savepoint)) nativeMacroStorageError(query);
    try {checkNativeMacroStorage(db,true);auto result=action();if(!query.exec("RELEASE "+savepoint)) nativeMacroStorageError(query);return result;}
    catch(...) {const bool rollback=query.exec("ROLLBACK TO "+savepoint),release=query.exec("RELEASE "+savepoint);if(!rollback || !release) throw VbaRuntimeError(3146,"Makro-Katalogtransaktion konnte nicht zurückgenommen werden.");throw;}
}
inline NativeMacroRevision storeNativeMacro(QSqlDatabase db,const QString& name,const QJsonObject& definition,const std::optional<NativeMacroRevision>& previous={}) {
    if(!nativeMacroNameValid(name) || name!=name.trimmed()) throw VbaRuntimeError(5,"Ungültiger Makro-Objektname.");(void)parseNativeMacro(definition);
    if(previous && previous->name!=name) throw VbaRuntimeError(3251,"Makro-Objektumbenennung ist hier nicht implementiert.");
    return nativeMacroWrite(db,[&] {
        if(nativeMacroObjectActive(db,name)) throw VbaRuntimeError(3251,"Ein laufendes Makro kann nicht ersetzt werden.");const auto canonical=nativeMacroCatalogName(db,name);
        if(!previous && (!canonical.isEmpty() || importedMacroNameExists(db,name))) throw VbaRuntimeError(3010,"Makroname ist bereits vorhanden. Vorhandenes natives Makro öffnen oder anderen Namen wählen.");
        if(previous && canonical!=name) throw VbaRuntimeError(3251,"Makroobjekt wurde zwischenzeitlich geändert. Neu öffnen.");
        const auto encoded=QString::fromUtf8(QJsonDocument(definition).toJson(QJsonDocument::Compact));QSqlQuery query(db);
        if(previous) {query.prepare("UPDATE main._la_objects SET definition=? WHERE kind='macro' AND name=? COLLATE BINARY AND definition COLLATE BINARY IS ?");query.addBindValue(encoded);query.addBindValue(name);query.addBindValue(previous->definition);}
        else {query.prepare("INSERT INTO main._la_objects(kind,name,definition) VALUES('macro',?,?)");query.addBindValue(name);query.addBindValue(encoded);}
        if(!query.exec()) nativeMacroStorageError(query);if(query.numRowsAffected()!=1) throw VbaRuntimeError(3251,"Makro wurde zwischenzeitlich geändert. Neu öffnen.");return NativeMacroRevision{name,encoded};
    });
}
inline void deleteNativeMacro(QSqlDatabase db,const NativeMacroRevision& revision) {
    (void)nativeMacroWrite(db,[&] {if(nativeMacroObjectActive(db,revision.name)) throw VbaRuntimeError(3251,"Ein laufendes Makro kann nicht gelöscht werden.");QSqlQuery query(db);query.prepare("DELETE FROM main._la_objects WHERE kind='macro' AND name=? COLLATE BINARY AND definition COLLATE BINARY IS ?");query.addBindValue(revision.name);query.addBindValue(revision.definition);if(!query.exec()) nativeMacroStorageError(query);if(query.numRowsAffected()!=1) throw VbaRuntimeError(3251,"Makro wurde zwischenzeitlich geändert. Neu öffnen.");return true;});
}
}
