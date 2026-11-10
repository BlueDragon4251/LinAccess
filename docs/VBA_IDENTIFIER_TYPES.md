# Typzeichen an nativen VBA-Bezeichnern

Die native Laufzeit unterstützt `%` (Integer), `&` (Long), `!` (Single), `#`
(Double), `@` (Currency) und `$` (String) bei skalaren lokalen Dim-/Static-
Deklarationen, skalaren Parametern, skalaren Modulvariablen und Const-
Deklarationen. Der Name ohne Typzeichen bezeichnet dieselbe Variable.
Ein zusätzliches `As` an einer solchen Deklaration ist unzulässig.

Der Prozedurrahmen löst typisierte einfache Variablenreferenzen anhand seiner
Deklarationsmetadaten auf. Der gegenwärtige Wert einer Variant-Variablen ersetzt
keine passende Deklaration. Ein nicht passendes Typzeichen wird verweigert.
Skalare Zuweisungen und ByRef-Argumente benutzen die kanonische Variablenbindung;
Strings, Kommentare, geklammerte Feldnamen und Objektmitglieder werden nicht
global umgeschrieben. Die Auflösung bleibt pro Aufruf; der Syntaxcache erhält
keine veränderlichen Rahmen- oder Hostdaten.

Die Lexik erhält Bang-Feldzugriffe und bisherige unmittelbar angrenzende
Verkettung. Mehrdeutige Verkettung sollte mit Leerzeichen geschrieben werden.
Der Access-/Jet-Ausdrucksmodus wird dadurch nicht auf VBA-Typzeichen umgestellt.
Bestehende String-Intrinsics mit `$` bleiben eigenständige Intrinsics, sofern
keine deklarierte lokale Variable die Namensauflösung übernimmt.

Noch offen sind insbesondere die vollständige frühe Bindung aller Projekt-/Library-/Objektnamensräume und
sämtliche Compiler-/Deklarationsregeln. `^`/LongLong
gehört nicht zum hier unterstützten Access-2003-Profil. Dies ist kein Nachweis
vollständiger VBA- oder Access-Parität. Die feste Katalogzahl bleibt 456;
Teilarbeiten ändern die 386 bestätigten Teilumfänge/70 offenen Vollanforderungen
nicht automatisch.

Semantische Grundlage: MS-VBAL 3.3.5.3 (Typzeichen), Deklarationsregeln und
5.6.10 (Abgleich des Typzeichens mit dem deklarierten Typ). Die vorhandene
Kernprüfung wird gebündelt verwendet; es entsteht kein zusätzlicher CTest-Eintrag.


## Funktionsnamen

Lokale Function-Deklarationen und deren Rückgabevariable unterstützen dieselben
sechs Typzeichen. Aufruf und Metadateninventar verwenden den kanonischen Namen;
Aufrufe mit Typzeichen müssen zum deklarierten Rückgabetyp passen. Direkte
Ausführung mit typisiertem Namen validiert den Typ ebenfalls. Sub-Deklarationen
mit Typzeichen und Function-Typzeichen zusammen mit As werden verweigert.
Bei Long-Funktionsaufrufen unmittelbar vor einer Argumentklammer bindet erst die
bekannte lokale Signatur das `&` als Typzeichen; die allgemeine Tokenisierung
bleibt unabhängig vom Prozedurrahmen und erhält Stringverkettung. Keine Aussage
zu sämtlichen Projekt-/Library-Namensräumen oder Property-Accessoren.
Vorhandenes core_tests:1/1 in0,57s; keine zusätzlichen CTest-Einträge.


## Implizite Variablen und Arrays

Ohne Option Explicit können ungebundene skalare Namen mit Typzeichen eine
passend typisierte lokale Variable anlegen; das Typzeichen hat Vorrang vor
Def-Direktiven. Option Explicit bleibt wirksam. Vorhandene lokale, Modul-,
Instanz- und Prozedurbindungen werden dabei nicht durch einen neuen Namen
ersetzt. Statische Verfahren behalten implizite Variablen, und ein Aufruf mit
oder ohne Funktions-Typzeichen verwendet denselben statischen Speicherbereich.
Typzeichen nach Literalbezeichnern True/False/Null/Empty/Nothing werden gemäß
MS-VBAL ignoriert; der Jet-/Access-Ausdruckspfad wird nicht umgestellt.

Lokale und Modul-Arrays unterstützen Typzeichen bei Deklaration, Elementzugriff,
ByRef-Arrayparametern, ReDim/Preserve und Erase. Elementtyp und kanonische
Arraybindung werden gemeinsam verwendet; vorhandene Allokations-, Bounds-,
Preserve-, ByRef- und Schleifenwachen bleiben erhalten. Ein bislang ungebundener
ReDim-Name deklariert wie bisher prozedurweit ein Array, auch mit Option Explicit;
der Typzeichenpfad verwendet nun ebenfalls dessen kanonischen Namen und Typ.
Typzeichen zusammen mit As werden in diesen Deklarationen verweigert.
Neue Arrayziele werden ausschließlich als einfache native Variablennamen
analysiert; Objektmitglieder werden nicht umgeschrieben.

Die bestehende Kernprüfung mit impliziten skalaren Variablen und Arrays bestand
in0,52s. Die nachfolgende gemeinsame Anwendungs-/Kern-/DAO-Prüfung enthält
zusätzlich die Korrekturen für statische Funktionsidentität und neue typisierte
ReDim-Ziele; Ergebnis: Anwendung, Kern- und DAO-Ziele gebaut; core_tests0,57s und
dao_execution2,97s, gemeinsam2/2 in3,54s bestanden.


## Anweisungspositionen

For/Next, For Each im vorhandenen unterstützten Sammlungspfad und direkte
lokale Function-Aufrufe mit Call oder als Anweisung binden dieselben kanonischen
Namen. Typisierte Arrayquellen behalten die bestehende Live-Mutationswache bei
For Each. Bei Arrayiteration bleibt die Voraussetzung eines Variant-Steuerelements
bestehen. Mid$-Zuweisung und LSet/RSet akzeptieren typisierte einfache Stringziele.
Die Signaturprüfung betrifft vorhandene lokale Prozeduren; eingebaute `$`-
Funktionen werden dadurch nicht zu vermeintlichen lokalen Verfahren.

Die Quellanweisungstrennung unterscheidet ein angehängtes Double-Typzeichen
von einem Datumsliteral. Doppelpunkt-Anweisungen nach Bezeichnern und Zahlen mit
`#` werden getrennt; Uhrzeit-Doppelpunkte innerhalb echter Datumsliterale bleiben
im Literal. Vorhandenes core_tests nach Korrektur des lokalen Signaturfilters:
1/1 in0,53s. Keine neuen CTest-Einträge und kein vollständiger Compilernachweis.


## Konstanten und Properties

Konstantenwerte besitzen in der quellgebundenen Übersetzung getrennte
Deklarationsmetadaten. Typzeichen in lokalen/Modul-Konstanteninitialisierern,
Enum-Werten, Optional-Standardwerten, festen Arraygrenzen und Stringlängen
werden dagegen geprüft. Auch ein momentan integraler Variant-Konstantenwert
begründet keine Integer-Deklaration. Inline-Übersetzungen übernehmen Werte und
Typmetadaten gemeinsam; es wird kein veränderlicher Rahmen im Syntaxcache
abgelegt. Bekannte qualifizierte Enum-Konstanten verwenden denselben Abgleich.

Property Get kann einen typisierten Funktionsnamen deklarieren; Property Let
verwendet weiterhin einen untypisierten Deklarationsnamen und den tatsächlichen
Typ des Wertparameters. Passende Get-/Let-Signaturen und Referenztypzeichen
werden geprüft. Kanonische interne Selektoren erhalten den Rückgabewert und die
Zuweisung an dieselbe Property. Der Schreibpfad prüft das echte Inventar unter
diesem kanonischen Namen. Die reine Syntaxprüfung akzeptiert auch `Name&()`;
erst die Laufzeit bindet diese Form an eine nachgewiesene Long-Signatur/Arraybindung.

Der Konstantenschreibschutz erkennt auch typisierte einfache Ziele und Mid$-
Ziele vor der Ausführung; ein rechter Funktionsausdruck wird dabei nicht erst
ausgeführt. Die Loop-Vorprüfung verwendet dieselben Namen/Typen wie die Laufzeit,
einschließlich Integer-Steuerungen unter DefStr A-Z.

Bestehendes core_tests nach gezielter Korrektur des Property-Inventarfilters und
der Syntaxprüfung parameterloser Long-Aufrufe1/1 in0,54s bestanden. Es wurden
keine CTest-/Katalog-IDs ergänzt. Der Paket-/Original-/Inventarbeleg des vorherigen
veröffentlichten Stands53e91932 ist in CODEX_HANDOFF.md dokumentiert; er liegt
vor dieser Konstanten-/Property-Erweiterung.
