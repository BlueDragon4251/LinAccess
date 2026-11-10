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

Noch offen sind insbesondere Typzeichen an Property-Accessoren, Arraydeklarationen
und ReDim-Zielen, Schleifensteuerungen, implizite typisierte Variablen und die
vollständige frühe Bindung qualifizierter Projekt-/Objektmitglieder. `^`/LongLong
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
