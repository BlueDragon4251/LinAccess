# Reale Access-2003-Kassendatenbank: Bridge-Kompatibilitätsprüfung

Stand: LinAccess Native 0.7.0

## Quelle und Integritätsprüfung

- Quelle: vom Anwender bereitgestellte `Kassa_CSE_V2.0.mdb` (JET4, ca. 31 MB).
- Windows-Bridge: `Export-LinAccessBridge.ps1`, **Microsoft Access 11.0**
  unter 32-Bit-PowerShell (Access 2003). Der Export wurde vom Anwender
  ausgeführt; die Linux-Seite hat das Ergebnis überprüft.
- Die 192 Einträge des Manifests sind sämtlich `source_only`. Keine
  `failed`- oder `unavailable`-Einträge.
- Der vollständige ZIP-Test und die Prüfung **aller 192 SHA-256-Dateihashes**
  waren erfolgreich. Die Original-MDB wurde ausschließlich lesend untersucht.

| Quellobjekt | Anzahl |
|---|---:|
| Access-Formulare (`SaveAsText`) | 47 |
| Access-Berichte (`SaveAsText`) | 9 |
| Access-Abfragen (Jet-SQL) | 100 |
| Access-Module (`SaveAsText`) | 4 |
| VBA-Dateien (`VBIDE.Export`) | 32 |
| **Insgesamt** | **192** |

Die 32 VBA-Dateien umfassen rund 4.562 Textzeilen. Der statische C++-Analysator
findet darin **187 Prozeduren**. Das zählt Deklarationen, keine erfolgreich
ausgeführten VBA-Programme. Auch deklarierte, aber leere Prozeduren werden
gezählt.

## Formular-Datenquellen

Aus den 47 Formularen lassen sich folgende `RecordSource`-Typen erkennen:

| RecordSource | Formulare | Derzeitige Linux-Umsetzung |
|---|---:|---|
| Direkter Tabellenname | 10 | Nur nach separatem, verlustbehaftetem SQLite-Tabellenimport teilkonvertierbar |
| Gespeicherter Abfragename | 4 | Kein generischer Jet-SQL-Konverter |
| Eingebettetes `SELECT`-SQL | 23 | Noch kein Formular-RecordSource-Übersetzer |
| Andere Referenz | 1 | Manuelle Prüfung nötig |
| Keine RecordSource | 9 | Ungebundene Access-Formulare erfordern eigene UI- und Ereignislogik |

21 der 47 Formulare haben ein passendes `Form_...`-VBA-Klassenmodul im Export.
Die ältere Formular-Teilkonvertierung bewahrt den Access-Designer nicht.
Ab 0.7 ist zusätzlich eine Qt-Layoutansicht vorhanden: Der SaveAsText-Parser
hat **685 reale Steuerelementdefinitionen** aus allen 47 Formularen gelesen,
darunter **43 CommandButtons**. Layout und Standardcontrols sind nur angenähert,
nicht pixel-/verhaltensidentisch mit Access 2003.

## VBA-Kompatibilitätsgrenzen

Die Quelldateien enthalten unter anderem:

- über 100 `DoCmd`-Verwendungen, etwa Formulare öffnen/schließen,
  `Requery`, Fokusnavigation und Druck-/Berichtsaktionen;
- häufig `CurrentDb`, `DAO.Recordset`, Datensatzschleifen und gebundene
  Steuerelemente (`Me.`, `Me!`, `Forms!`);
- Fehlerbehandlung (`On Error`), `For Each`, `Do While` und weitere
  bisher nicht vollständig implementierte VBA-Sprachkonstrukte;
- mutierende SQL-Anweisungen, darunter `INSERT`, `DELETE` und `DROP TABLE`.

Die Tests führen weiterhin **keine Original-Kassen-VBA-Prozedur** aus.
Version 0.7 bindet aber `CurrentDb.Execute` und `DoCmd.RunSQL` tatsächlich an
SQLite: Auch `DELETE`, `DROP TABLE`, `UPDATE` und `INSERT` werden auf einem
geöffneten LinAccess-Projekt ohne LinAccess-spezifische Zusatzbestätigung
ausgeführt, **sofern** der Interpreter die ganze aufgerufene Prozedur
versteht. Unbekannte Anweisungen erzeugen einen Fehler. Es gibt weiterhin
keine allgemeine DAO-/Jet-/VBA-Kompatibilität. In der Testumgebung wurde
eine getrennte SQLite-Kopie benutzt, nie das Original-MDB beschrieben.

In Version 0.6.1 wurde anhand dieses Exports die Erkennung von
**Umlauten in VBA-Prozedurnamen** korrigiert. Beispielsweise wird
`Mengenänderung_Click` nun erkannt. Eine eigene Testprozedur mit diesem Namen
wird auch von der eingeschränkten C++-Laufzeit ausgeführt, sofern ihr Inhalt
in der unterstützten Teilmenge liegt. Das bedeutet ausdrücklich **nicht**, dass
das originale Kassen-Ereignis bereits ausführbar ist.

Version 0.7 verfügt über einen Sammelimport und hat die eigentliche
`Kassa_CSE_V2.0.mdb` zusätzlich mit MDBTools gelesen. **29 von 29 Tabellen**
wurden in ein temporäres SQLite-Projekt kopiert. Die `Foto`-BLOB-Spalte
enthielt hier keine nicht-NULL-Werte; nicht-NULL-OLE/BLOB bleibt ausdrücklich
nicht unterstützt. Primär-/Fremdschlüssel, Indizes, Standardwerte und
Jet-spezifische Datentypsemantik fehlen nach wie vor. Ein synthetisches
Qt-Button-Ereignis hat danach real einen SQL-INSERT ausgeführt.

## Reproduktion mit privaten Testdateien

```bash
cmake -S . -B build -DCMAKE_BUILD_TYPE=Release \
  -DLINACCESS_COMPLEX_BRIDGE_DIR=/absoluter/pfad/Kassa_CSE_V2.0-Bridge-2003-Retry \
  -DLINACCESS_COMPLEX_MDB_FILE=/absoluter/pfad/Kassa_CSE_V2.0.mdb
cmake --build build -j2
ctest --test-dir build -R 'complex_access_bridge|access_2003_layouts|complex_real_mdb_table_copy' --output-on-failure
```

Der Test startet die Qt-Importoberfläche ohne Displayserver, verifiziert
den vollständigen Import in eine temporäre SQLite-Datenbank, zählt die
Objektarten und analysiert die 32 gespeicherten VBA-Dateien **ohne Ausführung**.
Die private Kassendatenbank und der persönliche Bridge-Export gehören
ausdrücklich nicht zum veröffentlichten Quellcode oder Debian-Paket.

## Weitere Arbeit

1. Jet-/ACE-Abfragen sicher in ein separates SQLite-Modell konvertieren.
2. Formulare mit SQL-RecordSources, Controls und Unterformularen abbilden.
3. VBA-Sprachumfang erweitern (Parameter, Objekte, Fehlerbehandlung).
4. Access-Objektmodell und DAO/Recordsets mit klaren Datenbankrechten und
   vollständigen Regressionstests implementieren.
5. Erst dann komplexe Bestandsanwendungen mit praktischen, kontrollierten
   Funktionstests abnehmen. Ohne Nachweise keine 1:1-Kompatibilitätszusage.
