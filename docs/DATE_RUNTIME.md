# Begrenzte native Datumslaufzeit

Stand: 2026-11-03. Dies beschreibt implementierte Teilfunktionen, keine vollständige
Access-2003-Datums- oder Gebietsschemakompatibilität.

Der C++-Laufzeitkern führt einen eigenen `Date`-Variant mit OLE-Serienwert. `VarType`
liefert 7, `TypeName` liefert `Date`; DAO-Datumsfelder haben den Datentypcode 8.
Date-Lokale, Modulfelder, Parameter und Funktionsrückgaben behalten ihren Typ.
`Len(Date)` liefert 8, `IsNumeric(Date)` liefert False. Ein überlaufender
Date-Variant kann bei Addition/Subtraktion zu Double werden; eine anschließende
Zuweisung an einen deklarierten Date-Wert bleibt bereichsgeprüft.

`Now`, `Date` und `Time` liefern Date-Variants. `CDate` unterstützt numerische
Serienwerte und eindeutigen gregorianischen ISO-Text. `DateSerial` und `TimeSerial`
normalisieren relative Monats-, Tages- und Zeitangaben. Zweistellige
DateSerial-Jahre benötigen noch die regionale Jahresgrenze und werden mit 3251
abgewiesen. `DateAdd` unterstützt Jahr, Quartal, Monat, Tag, Woche, Stunde, Minute
und Sekunde; Monats-/Jahreswechsel begrenzen den Tag auf das Monatsende. Das
Intervall `w` zählt alle Kalendertage. DateAdd normalisiert auch negative
OLE-Bruchteile über das tatsächliche Kalenderdatum.

Gebundene Datumsfelder und deklarierte DAO-Datumsparameter lesen echte
Date-Variants. Numerische Feldzuweisungen werden über den Serienwert konvertiert.
Die unterstützte SQLite-Speicherung verwendet ISO-Zeitstempel mit ganzen Sekunden;
feinere Date-Werte werden mit 3251 abgewiesen, bevor ein Feld verändert wird.
Interne und ungebundene Date-Werte behalten ihren Double-Serienwert. Die bisherige
ISO-Anzeige ist keine Nachbildung regionaler Access-Formatierung.

Der SQL-Adapter führt CDate, DateSerial, TimeSerial und DateAdd nativ aus. Direkte
Konstruktor- und Uhrfunktionsprojektionen behalten DATETIME-Metadaten. Bei
Datumsfeldern nutzt die Ergebnisbindung die physische SQLite-Spaltenherkunft
im Hauptprojekt, nicht den Textinhalt einer Zelle. Geprüfte einfache
SELECT-Projektionen unterstützen typisierte Addition/Subtraktion, Tagesdifferenzen,
Null und verschachtelte gewöhnliche Funktionsargumente. Numerische SUM-Ausdrücke
behalten native Datumsdifferenzen; SUM über Date-Werte selbst bleibt mit 3251
abgewiesen, auch bei leerer Eingabe. Fest als Date deklarierte SQL-Projektionen
weisen einen Bereichsüberlauf ab.

Offen bleiben unter anderem regionale Eingabe/Anzeige und Jahresfenster, Hijri,
vollständige Date-Literale, DateValue/TimeValue, sämtliche OLE-/Variant-Promotionen,
Untersekundenspeicherung, der vollständige Kalender-Control-Adapter sowie alle
SQL-Operatoren, Filter-/Aktions-/CASE-/CAST-/Aggregat- und komplexen Quelltypen.
DatePart/DateDiff haben weiterhin die in COMPATIBILITY.md dokumentierten
Intervall-/Wochen-/Systemeinstellungsgrenzen. Direkte negative gebrochene
Date-Arithmetik bleibt begrenzt; DateAdd ist davon getrennt implementiert.

Verifikation: bestehende Core-, DAO-, Qt-Feld- und Query-Modell-Ziele prüfen echte
Werte und SQL-Speicherung. Der letzte gemeinsame Stand a8d3061 bestand 48/48
CTests, GUI-freien Build, Paket- und vier Originalfixture-Prüfungen. Nachfolgende
Inkremente und ihre gezielten Ergebnisse stehen in PROGRESS.md. Es wurden keine
zusätzlichen CTest-Ziele für diese Datumserweiterungen angelegt.
