# Boolean-Bedingungen im nativen Interpreter

If/ElseIf, Do While/Until, Loop While/Until, While/Wend, IIf und native
Makrobedingungen/-wiederholungsbedingungen benutzen dieselbe Boolean-Konvertierung:
Ein Objekt-Standardwert wird einmal aufgelöst. Null ergibt in einer Bedingung
False; alle anderen skalaren Werte werden wie bei Boolean-Zuweisung/CBool
konvertiert. Stringwerte True/False und die exakten invarianten Schreibweisen
#TRUE#/#FALSE# sind damit ebenfalls verwendbar. Letztere sind case-sensitive;
#true#/#false# werden ausdrücklich nicht als numerische Strings angenommen.
CBool(Null) selbst behält Fehler94; eine Null-Bedingung wird getrennt behandelt.
Fehlerwerte, Arrays und Nothing werden nicht als erfolgreiche True-Bedingung
verschluckt. Inaktive Zweige lesen keine Standardwerte. IIf behält seine bisherige
eager Auswertung beider Ergebnisargumente.

Belegt durch MS-VBAL v20260818,5.6.16.3 und die String→Boolean-Konvertierung in
5.5.1.2.4, ergänzend5.4.2.8/9 für If-Formen. Dies ist ein Sprachspezifikationsbeleg,
kein neu ausgeführter Windows-Access-2003-Vergleich. Die bisherige native Annahme
von umgebendem Whitespace bei True/False und invarianten numerischen Strings
bleibt erhalten; vollständige regionale Let-Konvertierung ist weiterhin offen.

Der strukturierte Parser akzeptiert nun auch durch Klammern begrenzte
Schlüsselwörter wie If(v)Then, ElseIf(v)Then, While(v), Do Until(v), Loop While(v),
Select Case(v), Case(1) und With(v). Wortgrenzen verhindern die Umdeutung längerer
Bezeichner. EndIf und End mit mehreren Leerzeichen/Tabulatoren vor If/With sind
unterstützt. EndIf mit nachfolgender Doppelpunkt-Anweisung bleibt ein Abschluss,
kein Label. Exit Do akzeptiert die normalen Tokenabstände. Leere Then-/Else-Teile
in einzeiligen If-Anweisungen sind zulässig; Case ohne Ausdruck wird abgelehnt.

Die bestehenden Kern- und GUI-Ziele prüfen tatsächliche Interpreterausführung,
alle Bedingungspositionen, skalare Typen/Decimal/Currency, Null-Standardwerte,
Grenzen, inaktive Zweige, einmalige Getter, wiederholendes Resume nach Bedingungsfehler und echte
SQL-Effekte von Makrobedingungen und Wiederholungen. Keine neuen CTest-IDs.

F045/F041/F051/F088 bleiben breite offene Anforderungen: vollständige Grammatik,
Bindung, Fehlerreihenfolge/-fortsetzung, Locale und Windows-Oracle werden durch
diesen Umfang nicht als erledigt behauptet. Vorhandene Ausführungsbudgets gelten
weiterhin. Insbesondere ist eine pauschale Access-Parität nicht nachgewiesen.

Bei einem Fehler im If-/ElseIf-Kopf bleibt ein eigener ausstehender Bedingungsrahmen
bestehen. On Error Resume Next, explizites Resume Next und das Debugger-
Überspringen führen keine Then-/ElseIf-/Else-Nebeneffekte aus und beschädigen
keinen äußeren If-Rahmen. Wiederholendes Resume wertet denselben Kopf erneut aus;
der Rahmen wird dabei nicht dupliziert. Err bleibt bei impliziter Fortsetzung
sichtbar; explizites Resume behält die vorhandene Err-Rücksetzung. Das ist der
nachgewiesene native Umfang, kein vollständiger Fehlerreihenfolge-Oracle.
