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
