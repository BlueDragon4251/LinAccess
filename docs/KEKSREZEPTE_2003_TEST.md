# Reale MDB-Regression: `Keksrezepte.mdb` (Access 2003)

Stand: LinAccess 0.4.0. Es wurde eine vom Anwender bereitgestellte MDB-Datei
überprüft, ohne sie zu verändern oder mit der Software weiterzugeben.

## Nachgewiesene Objekte

| Jet-Objekt | Anzahl | Inhalt / Status |
|---|---:|---|
| Tabelle | 2 | `3_tbl_Rezepte` (4 Zeilen); `33_tbl_ZutatenRezept` (35 Zeilen) |
| Formular | 4 | `2_frm_Nährwerte_Eingabe_Neu`, `3_frm_Rezepte`, `33_Ufo_frm_ZutatenRezept`, `333_Ufo_Rezept_NWProRezProPerson` |
| Bericht | 1 | `Rezeptsammlung` |
| Sichtbare Abfrage | 3 | `3_qry_Rezepte`, `33_qry_ZutatenRezept`, `qry_Rezept für Bericht` |
| Verborgene Systemabfrage | 16 | `~sq_...`-Objekte, nicht mit regulären Benutzerabfragen verwechseln |
| Beziehung | 1 | `33_tbl_ZutatenRezept.Rezept` → `3_tbl_Rezepte.RezeptNr` |
| Eigenständiges VBA-Modul | 0 im MSysObjects-Katalog | **Keine Aussage** über eingebetteten Code in Formularen oder Berichten |

MDBTools meldet `JET4`. Dies ist ein Test einer echten alten MDB, **kein**
Nachweis für die Kompatibilität beliebiger Access-2003-Dateien.

## Automatisiert getestet

1. Objekttypen und Zählung über die Jet-Systemtabelle `MSysObjects`.
2. Quell-SQL der Abfragen über `mdb-queries` gelesen.
3. Eine Beziehung aus `MSysRelationships` gelesen.
4. Tabellenschema und CSV-Daten lesend aus MDBTools extrahiert.
5. Beide Tabellen in einer **separaten** SQLite-Datei angelegt; numerische
   Datentypen und `NULL` aus den Quellfeldern korrekt typisiert.
6. Datenmengen **4 + 35** und ausgewählte Werte/SQLite-Typen geprüft.
7. Doppelte Tabellenimporte werden abgelehnt; Importprotokoll unverändert.
8. SHA-256 des Original-MDB vor und nach dem Test identisch.
9. Qt-Offscreen-Oberfläche: 11 sichtbare Katalogobjekte erkannt, die 11
   Metadateneinträge transaktional gespeichert und Quell-SQL wieder ausgelesen.

## Wichtige Inkompatibilitäten gerade dieser Datenbank

- Die Abfrage `3_qry_Rezepte` referenziert `2_tbl_Nährwerte`;
  diese Tabelle ist **nicht** Bestandteil der bereitgestellten Datei.
- Die Bericht-Abfrage referenziert weitere abwesende Tabellen/Abfragen,
  unter anderem `1_qry_Kategorie`, `11_qry_Unterkategorie` und
  `333_qry_NährwerteProRezProPerson`.
- Jet-SQL-Ausdrücke wie `IIf(...)` oder `First(...)` werden nicht automatisch
  in SQLite-Ausdrücke übersetzt oder ausgeführt.
- Die originalen Formulare, Unterformulare, der Bericht und deren eventuell
  eingebettete VBA-Ereignisprozeduren können derzeit **nur mit Namen**
  katalogisiert werden. Weder Layout noch VBA-Quellcode wurden extrahiert.
- Insbesondere ist die Anzahl der **tatsächlich vorhandenen VBA-Prozeduren**
  unbekannt. Null eigenständige Module bedeutet **nicht**, dass keine
  Formular-/Berichts-VBA-Prozeduren existieren.

## Wiederholen

Die Datei ist bewusst keine öffentliche Test-Fixture. Per
`-DLINACCESS_REAL_MDB_FILE=/pfad/zur/mdb` kann jeder Entwickler dieselbe
Regression ausführen. Ohne diese CMake-Option werden nur die quelloffenen,
synthetischen Parser- und Bridge-Tests registriert.
