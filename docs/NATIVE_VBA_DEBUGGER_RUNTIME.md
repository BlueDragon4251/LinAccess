# Nativer VBA-Debugger

Der Modul-Editor führt den tatsächlich gespeicherten Modulstand im vorhandenen
Interpreter aus. F9 setzt/entfernt einen Haltepunkt an einer ausführbaren Zeile;
F5 startet bzw. setzt fort. F8 startet im Einzelschritt, Umschalt+F8 geht über
einen Aufruf, Strg+Umschalt+F8 verlässt den aktuellen Aufruf. Umschalt+F5 beendet
den Lauf. Strg+Pause fordert während der Ausführung eine Unterbrechung an.
Haltepunkte gelten für die konkrete Quellrevision und werden bei Bearbeitung
dieser Revision entfernt. Sie verändern weder Modulquelle noch Projektdatei.

Die angehaltene Ansicht zeigt wirkliche Quellzeilen, verschachtelte Aufrufrahmen,
aktuelle lokale Werte einschließlich ByRef-Rückbindungen und Err-Metadaten.
Variablen und Arrayelemente werden seitenweise gelesen. Arrays zeigen ihre
tatsächlichen Grenzen und Elemente, Strings lassen sich vollständig bis zur
bestehenden Laufzeitgrenze von16MiB öffnen. Objektinspektion zeigt Typinformation;
sie führt keine Property-Getter oder Default-Member mit möglichen Nebenwirkungen
aus. Ein nicht dimensioniertes Array wird ausdrücklich als solches angezeigt.

Auch F5 ohne Haltepunkte hält bei einem tatsächlich unbehandelten Laufzeitfehler
an. Der zuvor ausgeführte SQL-Abschnitt wird dafür nicht erneut gestartet.
Vorhandene On-Error-Handler behalten Vorrang. Fortsetzen propagiert den Fehler;
bei einem fortsetzbaren Anweisungsfehler bietet „Fehler überspringen“ die echte
Resume-Next-Fortsetzung der betroffenen Anweisung. Err bleibt beobachtbar.
Beenden stoppt den bestehenden Interpreter und wird ausdrücklich als Abbruch
angezeigt. Bereits ausgeführte Datenbankaktionen werden dadurch nicht rückgängig
gemacht; vorhandene Transaktionsregeln bleiben maßgeblich.

Während einer Unterbrechung sind Quelle/Speichern und Projektwechsel geschützt.
Inspektion ist ausschließlich bei lebenden angehaltenen Rahmen zulässig. Nach
Fehlerpropagation bleibt eine eingefrorene letzte Ansicht bis Schließen sichtbar;
sie enthält keine aufrufbaren Referenzen auf beendete lokale Rahmen. Nach einem
Lauf funktionieren Editorfokus und F9 wieder. Der Kern bleibt Qt-frei.

## Nachweis und Grenzen

Die bestehenden Kern- und GUI-Ziele prüfen StepInto/Over/Out, Haltepunkte auch
hinter Dim auf derselben Zeile, inaktive Zweige, verschachtelte Rahmen, ByRef,
Arrayseiten, vollständige Strings, nebenwirkungsfreie Objektinspektion,
On-Error-Vorrang, tatsächliche Fehler11/Err, Fortsetzen/Überspringen/Stop sowie
SQL-Seiteneffekte vor und nach der Fehlerstelle. Eine echte gespeicherte native
Modulquelle wird über den realen Editor bedient; GUI-Nachweise laufen offscreen.
Keine neue CTest-ID und keine zusätzliche Katalog-ID.

Dies belegt den beschriebenen nativen Debuggerumfang von F060. Vollständige
Access-IDE-, Makrodebugger-, VBA-Sprach-, COM-, Ereignis- und Windows-Parität
folgt daraus nicht. Der Einstieg liegt im nativen Standardmodul-Editor; importierte
und native gerufene Prozeduren verwenden dieselben Interpreterrahmen, soweit
sie vom vorhandenen Interpreter unterstützt werden. Unbekannte Modulquellen
werden als unbekannt beschriftet. Sprach- und Ausführungsgrenzen bleiben explizit;
insbesondere ist der ursprüngliche Kassa-F129-Fehler3075 weiterhin offen.


## Explicit Stop and Debug.Assert

Stop and Debug.Assert BooleanExpression are native interpreter statements.
An active debugger suspends at the original statement line with unchanged live
frames/variables. Continue resumes after the statement; Step/Stop use the existing
session controls. A true assertion evaluates once without a halt. A false or Null
assertion halts; the whole expression is evaluated eagerly, including the second
operand of And/Or. Scalar default getters are evaluated once. These halts are not
runtime errors, do not overwrite Err, and remain visible under On Error Resume
Next. They have distinct Stop/Assertion reasons, with existing enum values intact.
The actual debugger dialog labels both cases and disables error-only continuation.

Without an active native debugger, Stop or a failed assertion raises explicit3251
rather than silently continuing. The saved-module editor already creates a debugger
session for Run/F5; this does not automatically create one for every form event or
standalone host. Debug.Print output is described below; an interactive Immediate evaluator, complete IDE lifecycle,
Windows modal/event oracle and general VBA parity remain unproven. Malformed assertions fail normal
whole-procedure preflight before host effects. MS-VBAL5.4.2.11/5.4.2.23 are the
normative basis; Stop resumption is explicitly implementation-defined there.

Existing core tests cover live variables, Continue/Stop, eager ByRef operand effects,
inactive branches, default getter identity, Err preservation and unavailable-host
errors. The existing document test drives the real Stop/Assert dialog, reads the
actual local variable, resumes and checks actual SQL before rollback. Final targeted
core/GUI2/2 passed (14.49s); no new CTest ID, Windows or original-workflow claim.
F040/F045/F051/F060 remain at their existing documented catalog statuses.


## Native Debug.Print output

Debug.Print and Debug.? execute in the actual interpreter, including outside an
active debugger. Expression lists use comma zones of14 UTF-16 units, semicolon
continuation, Spc and Tab, and native scalar formatting. Expressions evaluate once
from left to right; output already emitted survives a later expression error.
Malformed lists fail procedure preflight. Numeric leading/trailing spaces and
separator rules follow MS-VBAL5.4.5.8/5.4.5.8.1. Native date/locale conversion
boundaries remain; this is not a Windows output-format oracle.

The real View-menu Debug-Ausgabe dock and the debugger's output pane display the
same thread-local buffer. Output remains readable during Stop and after Continue.
Clear resets it explicitly. Reading output never invokes property getters or
re-evaluates expressions. Hidden/unchanged views skip refresh; appends preserve
manual scroll position. No output is written to stdout, files or network.

Limits: retained text1MiB/8192 line breaks, statement output1MiB, expression list
64KiB/256 items/64 nesting levels. Retention drops complete UTF-8 sequences and
reports truncation. Position follows UTF-16 units; full glyph/display-width parity
is not claimed. Print#, Me.Print, graphics and interactive Immediate evaluation
remain outside this scope.

Application built; existing core/GUI bundle2/2 passed (17.43s), including real
Stop/Continue output, separators, ByRef single evaluation, partial error output,
UTF-8 retention and clear. The core binary links without Qt. No new CTest IDs.
F060 remains a confirmed partial scope; catalog456/386/70 is unchanged.
