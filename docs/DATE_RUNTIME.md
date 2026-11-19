# Begrenzte native Datumslaufzeit

Stand: 2026-11-03. Dies beschreibt implementierte Teilfunktionen, keine vollständige
Access-2003-Datums- oder Gebietsschemakompatibilität.

Der C++-Laufzeitkern führt einen eigenen `Date`-Variant mit OLE-Serienwert. `VarType`
liefert 7, `TypeName` liefert `Date`; DAO-Datumsfelder haben den Datentypcode 8.
Date-Lokale, Modulfelder, Parameter und Funktionsrückgaben behalten ihren Typ.
`Len(Date)` liefert 8, `IsNumeric(Date)` liefert False. Ein überlaufender
Date-Variant kann bei Addition/Subtraktion zu Double werden; eine anschließende
Zuweisung an einen deklarierten Date-Wert bleibt bereichsgeprüft.

VBA-Datumsliterale mit vierstelligen ISO-/US-Daten und eindeutigen 24-Stunden-
Zeiten werden als Date-Konstanten gelesen, auch in optionalen Date-Vorgaben.
Ungültige Literale scheitern vor Verfahrensausführung; weitere regionale und
zweistellige Schreibweisen bleiben offen.

`DateValue` und `TimeValue` trennen Datum und Uhrzeit für Date/ISO-Werte;
TimeValue übernimmt Null, DateValue meldet dafür 94. Eindeutige 24-Stunden-
Minutenangaben werden normalisiert. DateDiff mit `w` zählt ganze Wochenabstände
vom Ausgangswochentag; System-/Wochenjahrgrenzen bleiben begrenzt.

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

Ungebundene native MSCAL-Kalender liefern Date-Werte und akzeptieren ISO-Text,
Date oder numerische Serienwerte für reine Kalenderdaten. Gebundene Kalender,
Zeitanteile und vollständige ActiveX-Kompatibilität bleiben begrenzt. Der
Datumsfeld-Textparser prüft Datum und Uhrzeit getrennt von der Systemzeitzone;
eine Sommerzeit-Lücke macht einen gültigen zivilen Access-Wert nicht ungültig.

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
weitere Date-Literal-Schreibweisen, regionale DateValue/TimeValue-Eingaben, sämtliche OLE-/Variant-Promotionen,
Untersekundenspeicherung, der vollständige Kalender-Control-Adapter sowie alle
SQL-Operatoren, Filter-/Aktions-/CASE-/CAST-/Aggregat- und komplexen Quelltypen.
DatePart/DateDiff haben weiterhin die in COMPATIBILITY.md dokumentierten
Intervall-/Wochen-/Systemeinstellungsgrenzen. Die folgenden Access-2003-Messwerte
beschreiben die VBA-Operatoren mit negativen OLE-Bruchteilen.

Verifikation: bestehende Core-, DAO-, Qt-Feld- und Query-Modell-Ziele prüfen echte
Werte und SQL-Speicherung. Der letzte gemeinsame Stand a8d3061 bestand 48/48
CTests, GUI-freien Build, Paket- und vier Originalfixture-Prüfungen. Nachfolgende
Inkremente und ihre gezielten Ergebnisse stehen in PROGRESS.md. Es wurden keine
zusätzlichen CTest-Ziele für diese Datumserweiterungen angelegt.


VBA Timer, Timer() and VBA.Timer() read the actual Linux system clock and return
Single seconds since local midnight, including a fractional second. The sample
is shared for its whole/fractional components. Like a wall clock, it wraps at
midnight and can change with system-clock adjustments; it is not a monotonic
stopwatch. Windows tick-resolution equivalence and native SQL Timer translation
are not claimed. Arguments are rejected with error 450.

Eigene native Date-Felder unterstützen jetzt Date()/Now()/Time() als begrenzte
Tabellenstandardwerte. Neue DAO-/Datenblatt-/Formularpuffer erhalten tatsächliche
Uhrwerte und speichern den initialisierten Wert, auch bei ungebundenen
Datumsfeldern im Formular. Neue Date-Schemas erzwingen kanonische Zeitstempel;
ältere Definitionen werden bei Migration nicht still verändert. Einzelheiten
und weiterhin offene Grenzen: SOURCE_LITERAL_DEFAULTS.md, Abschnitt vom2026-11-08.


Typisierte native SQL-Kalenderbindung (2026-11-10)
------------------------------------------------
DateValue/TimeValue/DateAdd/DateSerial/TimeSerial/CDate, DatePart/DateDiff,
Year/Month/Day/Hour/Minute/Second/Weekday und MonthName/WeekdayName erhalten
in nachgewiesenen einfachen Quellen die tatsächlichen Argumenttypen. Date,
Boolean, Single und Currency werden über die gemeinsame Feldbindung gelesen;
ein ISO-Text in einer TEXT-Spalte wird dadurch nicht als Date deklariert.
Verschachtelte Projektionen sowie unterstützte WHERE-/GROUP-BY-/ORDER-BY-
Ausdrücke behalten diese Bindung und die tatsächlichen Ergebnistypen.

Vergleiche zwischen zwei nachgewiesenen Date-Ausdrücken verwenden native
Date-Werte; Datumsliterale behalten ihre Herkunft bereits im Lexer. Die sechs
Vergleichsoperatoren liefern Boolean (-1/0), NULL bleibt NULL. Date-Ausdrücke
mit IS NULL/IS NOT NULL erhalten die gebundenen Kinderausdrücke. Datumsfelder
mit reinem Datum werden dabei normalisiert, ungültige Date-Speicherung wird
abgewiesen. Andere gemischte Vergleichstypen sind damit nicht vollständig
implementiert. Der allgemeine Jet-Übersetzer F067 bleibt offen.

Die bestehenden Grenzen für Gregorian/ISO, System-Wochenregeln, regionale
Eingabe und Ausgabe, Untersekundenspeicherung und komplexe Quellen gelten
weiter. Die gemeinsamen SQL-Scalar-Adapter lehnen ungültiges UTF-8 vor einer
ersetzenden Textkonvertierung ab. Dies ist ein nativer Implementierungsbeleg,
kein neuer Windows-/Access-Ergebnisvergleich.

## VBA-Arithmetik mit negativen OLE-Bruchteilen (2026-11-19)

Access 2003 (11.0, MSACCESS.EXE 11.0.8166) führte die Ausdrücke in einem
VBA-Standardmodul einer frischen temporären Jet4-Datenbank aus. Die versteckte
32-Bit-COM-Sitzung rief die Funktion über `Application.Run` auf. Die Funktion
wies das Ergebnis einem Variant zu und erfasste `VarType`, `TypeName`, den
numerischen Serienwert über `CDbl` und bei Fehlern `Err.Number`. Die Datenbank
enthielt ausschließlich das synthetische Probe-Modul. `dateVariant` war als
Variant deklariert und enthielt `CDate(-1.25)`. `dt` und `epochDt` waren als
Date deklariert und enthielten `CDate(-1.25)` und `CDate(-1.75)`.

| Access-2003-Ausdruck | Ergebnis | Serienwert |
| --- | --- | ---: |
| `CDate(-1.25)+0.25` | Date (7) | -1 |
| `CDate(-1.25)+0.5` | Date (7) | -0.75 |
| `CDate(-1.25)-0.25` | Date (7) | -1.5 |
| `CDate(-1.25)+1` | Date (7) | -0.25 |
| `CDate(-0.25)+0.25` | Date (7) | 0 |
| `CDate(0)-0.25` | Date (7) | -0.25 |
| `0.5-CDate(-1.25)` | Date (7) | 1.75 |
| `CDate(-1.25)+CDate(-0.25)` | Date (7) | -1.5 |
| `CDate(-1.25)-CDate(-1.75)` | Double (5) | 0.5 |
| `CDate(-0.25)-CDate(0.25)` | Double (5) | -0.5 |
| `dt+0.5` | Date (7) | -0.75 |
| `dt-epochDt` | Double (5) | 0.5 |
| `dateVariant+Empty` | Date (7) | -1.25 |
| `dateVariant+Null` | Null (1) | |
| `dateVariant-CVErr(5)` | Fehler 13 | |
| `CDate(-657434)-0.25` | Date (7) | -657434.25 |
| `CDate(-657434.25)-0.75` | Fehler 6 | |
| `CVDate(-657434.25)-0.75` | Double (5) | -657435 |
| `DateAdd("d",1,CDate(-1.25))` | Date (7) | 0.25 |

LinAccess verwendet für die geprüften VBA-Operatoren jetzt die OLE-Serienwerte
direkt. `Date - Date` ergibt die numerische Differenz als Double. `DateAdd`
verschiebt das Kalenderdatum und behält bei einer Verschiebung um einen Tag
die Uhrzeit. Beide Ergebnisse im Beispiel mit einem Tag zeigen 06:00 Uhr am
30.12.1899 an und haben unterschiedliche Serienwerte.

Die Core-Regression prüft auch deklarierte Date-Werte, Varianten,
Ergebnisdarstellung und den unteren Datumsbereich. Regionale
String/Date-Konvertierung, allgemeine Variant-Koerzierung und Jet-SQL-Operatoren
bleiben außerhalb dieser VBA-Messung. F041 bleibt offen.
