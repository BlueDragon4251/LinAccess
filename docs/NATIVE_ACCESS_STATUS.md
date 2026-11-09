# Native Access status messages and progress meter

`Application.SysCmd` and unqualified `SysCmd` now execute the status actions
acSysCmdInitMeter1, UpdateMeter2, RemoveMeter3, SetStatus4 and ClearStatus5.
Both function expressions and statement calls use the real native Application
object; named Action/Argument2/Argument3 use the normal binder. Local and external
VBA procedure declarations retain priority for statement calls; local functions
retain expression-call priority. Syntax validation does not invoke GUI actions.
Calculated-expression/SQL SysCmd is not added.

The real QMainWindow status bar displays the message and QProgressBar; a standalone
native document obtains an actual status bar in its layout. Maximum200/update100
produces an actual half-filled meter. SetStatus removes the meter, as documented.
RemoveMeter removes its widget and message; ClearStatus clears the message.
Successful status actions return a Null Variant. No fake Access version, registry,
ActiveX or compiler result is returned for other unimplemented actions3251.

Bounded validation happens before changing existing message/meter values: correct
arity450, required arguments, normal VBA String/Long conversions, Null94,
nonempty single-line/NUL-free text5, positive maximum5, update0..maximum5.
Updating without an initialized native meter is unsupported3251. Empty text is
rejected; a single space can replace the status text. Native Qt font/width,
clipping, maximum/range and exact Windows error precedence lack an Access2003
oracle. Status widgets belong to their Qt surface; Application objects still
check the originating document's weak context and fail2450 after closure.
These are actual status actions, not a complete Application/SysCmd implementation.

Existing core/layout2/2 (13.17 s), fresh GUI-free1/1 (1.75 s), current Debian
package/build-ID/SQLite3.54.1 and four original fixture hashes passed. The actual
VBA/UI case observes the maximum, initial/update values, unchanged value after an
invalid update, SetStatus meter removal, ClearStatus and closed-context failure.
No new CTest target or catalog ID. Last full baseline48/48 (83.85 s); no full
rerun claimed after this bounded extension. Existing F068 remains open;456/384/72.

References: [Application.SysCmd](https://learn.microsoft.com/en-us/office/vba/api/access.application.syscmd),
[AcSysCmdAction](https://learn.microsoft.com/en-us/office/vba/api/access.acsyscmdaction).
Only the five longstanding status actions are exposed; newer Office constants
from the evolving online documentation are not added as Access2003 features.


## Native GetObjectState

SysCmd action10 / acSysCmdGetObjectState now reports actual native object state
for table0, query1, form2 and report3. Open bit1 comes from the existing project
registry and actual registered tab/dialog membership, including hidden native
runtime documents. Modified bit2 is set only by tracked design editors; pending
data-record changes do not become design changes. New bit4 is available only for
tracked registered design identities. Unsaved anonymous native editors do not
reserve an Access object name, so complete Access temporary naming/new-object
state semantics remain open. Closed or nonexistent native object names return0.
The named action/object-type/object-name arguments use Argument2/Argument3.

The shared checked state reader also backs existing AccessObject.IsLoaded;
connection, database identity, project-generation and Qt owner guards are retained.
Other Access object types explicitly3251. No generic all-object COM/VBE state,
external-instance state, original Access designer bits or complete Windows error
precedence claim. Native acObjStateOpen/Dirty/New constants expose1/2/4.

Existing real mainwindow case checks runtime form open, closed report/nonexistent
object, real query designer open1/modified3/saved1/closed0, and pending editable
query data remaining open1. Targeted core/mainwindow/layout3/3 (15.13 s), fresh
headless1/1 (1.76 s), current package/build-ID/SQLite3.54.1/four hashes pass.
Last full baseline48/48 (83.85 s); no full rerun claimed.456/384/72 unchanged.


Berechnete Felder im visuellen Abfrageentwurf: echte skalare Ausdrücke mit Alias,
Ausgabe, Sortierung und UND/ODER-Kriterien; Hinzufügen/Entfernen betrifft nur den
Entwurf. Gespeicherte Ausdrücke öffnen wieder korrekt. Begrenzte native skalare
Funktionen; Aggregate, Unterabfragen, unbekannte Callbacks und externe/lebende
Werte werden in diesem Entwurf ausdrücklich begrenzt. Keine Zeilenauswertung in
Vorschau/Übernahme. Details: docs/VISUAL_QUERY_DESIGNER.md. Anwendung gebaut;
bestehendes GUI-/SQL-Bündel2/2 bestanden (15,09s), inklusive wirklicher berechneter
Ergebnisse gegen unabhängiges SQL, Sortierung und Wiederöffnung. Keine neuen
CTest-IDs oder Windows-/Original-/Gesamtsuiten-/Paketnachweise. Parameter/Oder-
Etappe als b1b5858 veröffentlicht; letzte geprüfteDEB weiterhin1dafedd.
Katalog456/386/70 unverändert; keine vollständige Abfrageentwurfsparität behauptet.


Gruppierung im visuellen Abfrageentwurf: pro Feld Group By/Where/Count/Sum/Avg/
Min/Max/Var/StDev, echte GROUP BY-/WHERE-/HAVING-Klauseln mit getrennten ODER-
Gruppen; Where-Felder nur ausgeblendet/unsortiert. Entwurfsmetadaten behalten
Funktionen und Modus. Nachträglich hinzugekommene Quellspalten verändern keine
bestehende Gruppierung. Details: docs/VISUAL_QUERY_DESIGNER.md. Bestehendes GUI-/
SQL-Bündel2/2 (14,62s) bestanden; nach Korrektur der Schemaerweiterung nur betroffene
GUI-Prüfung1/1 (14,36s) wiederholt. Anwendung gebaut; keine neuen CTest-IDs oder
Windows-/Original-/Gesamtsuiten-/Paketnachweise. Berechnete-Felder-Etappe als
cee05d8 veröffentlicht/verifiziert; letzte geprüfteDEB weiterhin1dafedd.
F067/F068/F080/F148 sowie fester Katalog456/386/70 bleiben unverändert offen.


Visueller Datenquellenentwurf im nativen Berichtseditor: derselbe tatsächliche
SQL-Entwurf einschließlich berechneter Felder/Gruppierung; Übernahme als ein
Text-Undo, Abbrechen unverändert, Metadaten nur bei passendem SQL-Hash. Speichern/
Wiederöffnen und Invalidierung bei manueller SQL-Änderung nachgewiesen. Keine
Parameterdeklarationen in diesem Berichtsweg, da Berichtsparameterbindung noch
fehlt. Projektidentität/Generation wird bei Entwurf/Definition/Speichern/Ausgabe
geprüft. Bestehendes GUI-/SQL-Bündel2/2 (15,31s), Anwendung gebaut. Tatsächlicher
Callbackzähler: Entwurf und Speichern ohne Auswertung, explizite Ausführung liest
Zeilen. Keine neuen CTests oder Windows-/Original-/Gesamtsuiten-/Paketnachweise.
Details: docs/VISUAL_QUERY_DESIGNER.md. Gruppierungs-Etappe als31c50d2 veröffentlicht;
letzte geprüfteDEB weiterhin1dafedd. F080/F085/F087/F148 und456/386/70 unverändert.


Gespeicherte Auswahlabfragen als visuelle Datenquelle: tatsächliche native Query-
Views werden vor Entwurfsöffnung installiert und zusammen mit Tabellen angeboten;
Metadatenvorschau bleibt ohne Zeilenauswertung. Tatsächliche verschachtelte Abfrage
auf gespeichertem Entwurf inklusive berechnetem Feld/Kriterium/Sortierung und
Wiederöffnung geprüft; Callbackzähler trennt Vorschau von Ausführung. Native
Abhängigkeits-/Typ-/Parametergrenzen bleiben ausdrücklich erhalten. Anwendung
gebaut, bestehendes GUI-/SQL-Bündel2/2 (15,84s) bestanden. Ein Qt-Enum-Baufehler
wurde vor dem Prüflauf korrigiert. Keine neuen CTests oder Windows-/Original-/
Gesamtsuiten-/Paketnachweise. Berichtsentwurf als4e20658 veröffentlicht; letzte
geprüfteDEB weiterhin1dafedd. Details VISUAL_QUERY_DESIGNER.md;456/386/70 unverändert.


Visuelle Verknüpfungen: bis8 geordnete INNER-/LEFT-Equijoins auf tatsächlichen
Tabellen/gespeicherten Auswahlabfragen, pro neuer Quelle vorhandenes Feld = neues
Feld. Vollqualifizierte Feldidentität und persistierte Quellzuordnung; echte
Nullwerte bei fehlendem LEFT-Treffer. Entfernen betrifft nur Entwurfsfelder;
Selbstverknüpfungen/Aliase/zusammengesetzte Vergleichsbedingungen bleiben im
SQL-Editor mit bestehenden Grenzen. Strikte Metadatenprüfung; fehlgeschlagenes
Einlesen hinterlässt keine halbfertigen Feldzeilen. Bestehendes GUI-/SQL-Bündel2/2
(14,92s), nach Metadatenguards nur GUI1/1 (14,90s) wiederholt; Anwendung gebaut.
Tatsächliche zweistufige LEFT-Verknüpfung gegen unabhängiges SQL verglichen,
Wiederöffnung/Entfernen/INNER-Wechsel und kaputte Metadaten geprüft. Keine neuen
CTests oder Windows-/Original-/Gesamtsuiten-/Paketnachweise. Details in
VISUAL_QUERY_DESIGNER.md. Gespeicherte-Abfragequellen-Etappe28e3302 veröffentlicht;
letzte geprüfteDEB weiterhin1dafedd. F067/F068/F080/F148 und456/386/70 unverändert.


Visueller Formulardatenquellenentwurf: derselbe echte SQL-Entwurf erzeugt SELECT-
Quellen mit vorbereiteten Feldmetadaten; keine Parameterbindung und weiterhin
schreibgeschützte Abfrageformulare. Passende SQL-Hash-/Entwurfsmetadaten werden
mitgespeichert und im bestehenden Formular-Undo/Redo erfasst/wiederhergestellt.
Manuelle Quelländerung entfernt veraltete Metadaten. Projektidentitäts-/Generations-
prüfung bei Entwurf/Speichern/Verlauf. Bestehendes GUI-/SQL-Bündel2/2 bestanden
(15,47s), Anwendung gebaut; tatsächliches gespeichertes Formular mit echten
Ergebniswerten/schreibgeschütztem Zustand geöffnet. Callback zählt keine Auswertung
bei Entwurf/Speichern/Wiederöffnung, erst tatsächliche Formularausführung liest.
Keine neuen CTests oder Windows-/Original-/Gesamtsuitennachweise.

Verknüpfungs-Etappe alsfe01e1d veröffentlicht/verifiziert. Geprüftes DEB dieses
Standes unter /workspace/work/visual-query-joins/package/linaccess-0.7.0-Linux.deb:
Build-ID passt, MDBTools1.0.1/SQLite3.54.1, Payload/Berechtigungen/keine Fixtures
bestanden. Paket enthält noch nicht die anschließende Formularintegration. Vier
Fixture- und drei Upload-Originalhashes erneut unverändert geprüft. Details in
VISUAL_QUERY_DESIGNER.md/NATIVE_QUERY_FORM_DESIGN.md. F050/F080/F148 und fester
Katalog456/386/70 unverändert. Nächste unabhängige Untersuchung: ursprüngliche
3146/1023-Grenze beim Öffnen des Originalformulars, keine erneute Makro-Rückfrage.


Korrektur der bisherigen Original-3146/1023-Deutung: private Diagnose zeigte
Document initialization failed: Native macro catalog unavailable bei source-error0.
Die Import-Testkopie besitzt keinen optionalen main._la_objects-Katalog. Native
Makrosuche behandelt dessen Abwesenheit jetzt wie einen leeren Katalog (2485),
ohne Tabellen anzulegen oder temporäre Schatten zu lesen; vorhandene defekte
Kataloge bleiben3146. Beide tatsächlichen Tastaturhandler erreichen nun den
fehlenden Makroquelltext2485 in Zeile1023. Das ist keine offene Datensatzquellen-
SQL-Operation. Getaggter Click erreicht separat ebenfalls fehlenden Makroquelltext.
Bestehendes GUI-/Original-Bündel2/2 bestanden (125,46s; GUI15,85s); Anwendung gebaut.
Isolierte Prüfungen für fehlenden/temporären/defekten/leeren Hauptkatalog sind in
bestehendem Ziel enthalten. Keine neuen CTests, Windows- oder positive Original-
Verkaufs-/Makrolaufnachweise. Die diagnostischen Instrumentierungen/privaten Details
liegen ausschließlich außerhalb Git; sie sind keine zusätzlichen erfolgreichen
Prüfziele. Native Makroquelle fehlt weiterhin im Bridge-Export; vorhandene
Codex-Call-Rückfrage dazu blieb unbeantwortet, kein wiederholter Anruf nötig.
Formularintegration als3b18277 veröffentlicht; letzte geprüfteDEB fe01e1d vor dieser
Integration. Details NATIVE_MACROS_RUNTIME.md/READONLY_FORM_REQUERY_RUNTIME.md.
F129-Text korrigiert, fester Katalog456/386/70 unverändert. Weiter mit echten
Quellenaliasen und Selbstverknüpfungen im visuellen Entwurf.


Quellenaliase/Selbstverknüpfungen im visuellen Entwurf: wiederholte echte Tabellen/
Query-Views mit eindeutigem explizitem Alias, getrennte physische Herkunft und
SQL-Feldidentität, Metadaten immer von der tatsächlichen Quelle. Linker Selektor
bietet vorhandene Aliase und zugehörige Felder, Tooltip zeigt Herkunft. Summen-
dialog qualifiziert doppelte Feldnamen. Reale Eltern-Kind-LEFT-Verknüpfung gegen
unabhängiges SQL verglichen, Nullfälle/Alias mit Leerzeichen/Wiederöffnung und
Ablehnung mehrdeutiger aliasloser Wiederverwendung geprüft; Metadaten führen
keine Zeilenberechnung aus. Anwendung gebaut, bestehendes GUI-/SQL-Bündel2/2
(16,45s) bestanden ohne weitere Wiederholung. Keine neuen CTests oder neue
Original-/Windows-/Gesamtsuiten-/Paketnachweise. Details VISUAL_QUERY_DESIGNER.md;
F080-Beschreibung aktualisiert, Katalog456/386/70 unverändert. Makrokatalog-Fix
als657d6b3 veröffentlicht/verifiziert; letzte geprüfteDEB weiterhinfe01e1d.


Visuelle Verknüpfungen unterstützen nun INNER/LEFT/RIGHT mit1..16 UND-verknüpften
Feldvergleichen (=, <>, <, <=, >, >=) pro neuer Quelle, auch von vorhandenen
Aliasen. Zusammengesetzte Schlüssel werden tatsächlich als ON-Vergleiche erzeugt;
RIGHT erhält die Zeilen der neuen Quelle mit echten Nullwerten links. Vorgemerkte
Vergleiche plus vollständiges aktuelles Paar bilden die Bedingung; unvollständige/
unhinzugefügte Entwürfe müssen ergänzt oder explizit verworfen werden. Visuelle
Metadaten schreibenVersion2, Version1/alte Einfachvergleiche bleiben lesbar.
Anwendung gebaut; bestehendes GUI-/SQL-Bündel2/2 (15,54s) ohne Nachlauf bestanden.
Tatsächlicher nativer executeAccessQueryPlan-Output für zusammengesetzten RIGHT
JOIN gegen umgekehrten unabhängigen LEFT JOIN inklusive Null verglichen; <=-INNER
nach Wiederöffnung ebenfalls unabhängig geprüft. Keine neuen CTests oder neue
Original-/Windows-/Gesamtsuiten-/Paketnachweise. Details VISUAL_QUERY_DESIGNER.md.
Alias-Etappe9953a83 veröffentlicht; letzte geprüfteDEB weiterhinfe01e1d; jetzt
aktuelles Paket nach dieser zusammenhängenden Entwurfserweiterung prüfen.
F067/F068/F080/F148 und fester Katalog456/386/70 unverändert.


Visuelle Aktionsabfragen implementiert: Auswahl/Tabellenerstellung/Anfügen sowie
Aktualisieren/Löschen für einzelne physische Tabellen ohne Joins/Summen/DISTINCT/
Sortierung. Expliziter freier neuer Tabellenname bzw. reale Zielspaltenzuordnung;
Updatefelder mit begrenzten skalaren SQL-Ausdrücken und deklarierten Parametern,
WHERE aus tatsächlichen Kriterien. Vorschau/Übernahme erzeugen nur SQL und
Metadaten; ausschließlich Ausführen verwendet den bestehenden nativen DAO-
QueryDef-/Bindungs-/Aktionspfad. Reale Tabellen/Zeilen/Anfüge-Standardwerte sowie
Parameteraktualisierung und kriteriumsgebundene Löschung nachgewiesen. Speichern/
Wiederöffnung behalten Modus/Ziel/Zuordnung/Ausdrücke; Formular-/Berichtsquellen
bleiben Auswahl-only. Entwurfsformat schreibtVersion3, liestVersion1/2/3.
Initiales bestehendes GUI-/SQL-Bündel2/2 (16,23s) für Make/Append bestanden. Nach
Update/Delete nur betroffenes GUI-Ziel: Qt-Iterator-Baufehler und DELETE-*-Vorschau
korrigiert; endgültig1/1 (15,79s), Anwendung gebaut. Der zunächst fälschlich dem
Update zugeordnete Fehler lag bei der DELETE-Vorbereitung. Keine neuen CTests
oder neue Kern-/Original-/Windows-/Gesamtsuiten-/Paketnachweise für diese Stufe.
Details VISUAL_QUERY_DESIGNER.md; native Laufzeit-/Typ-/Metadatengrenzen bleiben.

Zusammengesetzte Joins als44b1f70 veröffentlicht/verifiziert. Geprüftes DEB dieses
Standes /workspace/work/visual-query-composite/package/linaccess-0.7.0-Linux.deb:
Build-ID, MDBTools1.0.1/SQLite3.54.1, Payload/Berechtigungen/keine Fixtures bestanden;
dieses Paket enthält noch nicht die anschließenden Aktionsentwürfe. F067/F068/
F080/F148 und fester Katalog456/386/70 unverändert. Weiter mit nativer TOP-Auswahl
im visuellen Entwurf; kein Gesamtabschluss, kein Abschlussanruf.


Native TOP-Steuerung im visuellen Auswahlentwurf: positive32-Bit-Zahl oder Alle
Datensätze(0), tatsächlicher nativer RANK-/LIMIT-Pfad einschließlich Grenzwert-
Gleichständen und ohne sichtbare Rangspalte. Gespeicherter Wert öffnet wieder,
Bruch-/negative/überlaufende Metadaten werden abgelehnt. Rohansichten __la_top_
werden aus Quellen-/Join-/Ziellisten ausgeschlossen; öffentliche gespeicherte
Abfrage bleibt verfügbar. TOP in diesem visuellen Pfad nur Auswahl; Aktionsmodus
darf den Wert nicht stillschweigend ignorieren. Native PERCENT-/DISTINCT-/Budget-
Grenzen bleiben. EntwurfsformatVersion4 schreibt, Version1/2/3/4 liest.
Anwendung gebaut; nur bestehendes betroffenes GUI1/1 (16,14s) bestanden. Tatsächliche
TOP2-Ergebnisse mit3 Grenzwert-Zeilen gegen unabhängige nth-Wert-SQL verglichen,
Ausgabe/Metadaten/Begrenzungslöschung/Callback-Nichtauswertung geprüft. Keine neuen
CTests oder zusätzliche Kern-/Original-/Windows-/Gesamtsuiten-/Paketnachweise.
Aktionsstufe als1481810 veröffentlicht/verifiziert; letzte geprüfteDEB44b1f70
vor Aktions-/TOP-Erweiterung. Details VISUAL_QUERY_DESIGNER.md. F067/F068/F080/
F148 und fester Katalog456/386/70 unverändert; Arbeit fortsetzen.
