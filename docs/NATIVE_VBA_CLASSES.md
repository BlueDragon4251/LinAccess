# Native authored class instances

Native owned class modules are saved as version1 UTF8 VBA documents in the real
project catalog (`kind=class`). The class editor and explorer create/save/reopen/
delete that source. A first body-only save writes a real CLASS/VB_Name header;
subsequent edits preserve the actual source text and line numbering. Readonly
imported VBA bytes are not rewritten. Native class names are ASCII identifiers
up to64 characters. The shared source namespace rejects collisions with ordinary
native/imported module names. AllModules includes actual owned classes, and its
existing editor-loaded state recognizes the real class editor. Public SQL/VBA
standard-module discovery skips class code; class methods are not registered as
free SQL functions or public global variables.

New creates a real independent instance through the native runtime. Module fields,
arrays and Static locals belong to that instance, not to a singleton standard
module frame. A private parameterless Class_Initialize Sub runs once per successful
construction. Ordinary public Functions/Subs, public fields/constants, Property
Get/Let/Set, native array returns, indexed reads/writes and native returned-object
default forwarding use the existing compiler, type and argument binder. Object
identity, Me identity and actual object references are retained. A Sub is permitted
in an object statement, and refused before its body when used as a value. External
class arguments currently require explicit ByVal; direct unqualified owning-class
ByRef calls retain the real existing aliases. No detached value-copy writeback is
presented as ByRef.

The lexical execution context tracks the actual source and host for every VM
entry. Standalone/nested expressions through another host do not inherit the
parent class's private-member privileges. Reentry into another procedure/source
uses that procedure's lexical scope; returning restores the caller's scope.
Private members remain available inside their class, including same-class instance
access, and unavailable from other modules/callback expression hosts. Constructor
references are bound before procedure effects; source strings and raw control
names containing New are not treated as constructors.

Class methods use genuine project services. CurrentDb/DBEngine, public standard
module calls, supported formatting and application commands use an owned service
host; no fabricated hidden Form exists. A class tied to an actual project window
outlives its creating editor/form. A standalone document uses its guarded document
lifetime. Database/workspace identity, project generation and source revision are
checked before/after method calls and before field reads/writes. Once an observed
revision/deletion/context change invalidates an instance, restoring that revision
does not revive it. Unobserved byte-identical delete/recreate, all external editing
orders and full Windows lifetime equivalence remain outside this proof.

Explicit native boundaries: references are held As Object/Variant; named custom
As types/As New, imported generic class construction, Implements/WithEvents,
default/predeclared instances, custom default-member attributes, COM/ActiveX and
complete external ByRef are not claimed. Class_Terminate is rejected before
construction instead of silently pretending cleanup ran. Bang/default member
binding and array-valued getter mutation retain explicit refusals. The existing
one-MiB source/array/procedure limits apply; live instances have a4096 count and
16MiB retained-source budget. This does not complete F040/F042/F043/F050 or claim
full Access2003 class parity. Fixed catalog456/386/70 remains unchanged.

Verification: application and the existing core/GUI/DAO/Original targets built.
First core/DAO passed5.20s/3.49s; Original exposed the Def preflight regression
and GUI hit its20-second watchdog under the concurrent block. The Def binding
was restricted/caught, constructor discovery excludes raw control names, and
private lexical scopes were completed before publication. The final existing
core/GUI/Original block passed3/3 in143.28s: core4.08s, GUI17.58s, Original121.62s.
Only affected targets were repeated; no new CTest IDs or broad full-suite run.
The GUI editor diagnostic observed no unexpected error dialog in the successful
run. All seven actual upload/fixture hashes remain unchanged and core remains
Qt-free. Native core/GUI cases cover independent initialization/state, Sub versus
value calls, fields/constants, properties/indexing/arrays, real ByRef internally
and explicit external refusal, named optional calls, Me identity, object default
forwarding, private callback/reentrant scopes, actual CurrentDb SQL writes,
class SQL-function exclusion, real editor Ctrl+S/header/catalog state, instances
outliving caller windows, sticky revision invalidation and TEMP/deletion guards.
The Original event path retains its known missing-macro boundary, not a positive
initialized sale proof. No Windows oracle or fresh class package proof. Last
checked Debian3fe7245 precedes this scope.

## Ordinary classes in the imported source archive

New also resolves a checked ordinary class from the latest actual archived VBA
revision. A CLASS header and matching VB_Name must agree with its canonical
archive name. VB_Base document classes and Form_/Report_ classes associated with
actual archived document names are excluded. A mere Form_ prefix without a document
association is not guessed to be CodeBehind. The same source-only classification
feeds AllModules; ordinary imported classes join authored classes, while actual
form/report modules remain excluded. Native/imported namespace collisions fail5.

Raw UTF8/BOM, UTF16LE and the existing CP1252 fallback remain archived unchanged;
only the execution view is decoded. Undefined fallback bytes that would produce
replacement characters refuse3251. Runtime instances retain the actual origin,
archive ID and raw revision bytes, not just decoded text. A newer revision with
equivalent decoded text but different raw encoding still invalidates the old
instance. Owned/imported origin changes, TEMP catalog shadows, project/database
lifetime and observed-source deletion retain explicit guards. No source is renamed,
rewritten or automatically promoted from document CodeBehind.

First existing GUI/DAO bundle2/2 passed20.44s (17.74/2.70s), app built. Actual New,
methods/properties/CurrentDb SQL writes, ANSI umlauts, UTF16 revision replacement,
raw-byte retention, AllModules, associated CodeBehind and VB_Base exclusions,
unassociated prefix classes, name/encoding failures, TEMP shadows and ambiguous
owned/imported names are exercised in separate technical archive records. The
pre-existing archive hash is identical after retiring only those test records.
These technical sources are not claimed as a Windows class-export oracle.
The existing Kassa Bridge contains28 class files, all28 declared document-module
names and no ordinary class candidate; it was only inspected as aggregate metadata.
No new CTest or additional Original/Windows/full-suite run. Class core is unchanged
from the preceding verified scope; complete F042/F043/F050 remain open. The prior
"imported generic class construction" boundary above is superseded for this checked
ordinary-class path. Other typed/default/reference/lifetime boundaries remain.
Latest checked Debian7f40893 includes authored classes, before this import extension.


## Benannte Projektklassen-Typbindungen (2026-11-09)

Benannte skalare Klassen-Typen binden nun an tatsächlich vorhandene gewöhnliche
Projektklassen: eigene und geprüfte archivierte Klassen, Felder/Dim, ByVal und
owning-module ByRef, optionale Nothing-Parameter, Function-/Property-Get-Rückgaben,
Property Set und TypeOf. Selbstreferenzen prüfen den Quellkatalog ohne rekursive
Konstruktion. Nothing bleibt erlaubt; falsche reale Instanztypen ergeben13,
unbekannte Klassennamen3251 vor Prozedureffekten. Lokale Enum-Aliase werden nicht
als Klasse erfunden. Direkte öffentliche Formularfelder validieren den echten
Namespace ebenfalls. Klassenwertige SQL-Funktionen werden vor ihrem Body3251
abgewiesen; der tatsächliche gemeinsame Aufrufzähler bleibt0. Interne Typmarkierungen
werden im Debugger nicht als Nutzer-Typnamen angezeigt.

Dies ersetzt die frühere Grenze „nur As Object/Variant“ für benannte skalare
Projektklassen. Klassen-Arrays, As New, Interfaces/WithEvents/Defaultinstanzen,
vollständiges externes ByRef und vollständige projektübergreifende Typidentität
bleiben offen. Technische gemeinsame Kern-/GUI-Fälle sind kein Windows-Oracle.
Katalog456/386/70 unverändert; F042/F043 bleiben vollständige offene Anforderungen.

Nachweise dieses Blocks: native Kernprüfung nach If-Typbindungsfix1/1 (4,15s),
mit OptionExplicit-Fixture erneut1/1 (4,04s); Original-Ereignispfad1/1 (123,33s)
für die native VM, anschließend nur noch Qt-Formularfeld-Set ergänzt. Finales
vorhandenes GUI-Ziel1/1 (17,80s), App gebaut. Erster GUI-Basisblock17,88s bestand;
Erweiterungen korrigierten technische SQL-Registrierung/Deklarationsreihenfolge
und fanden den fehlenden echten Formularfeld-Set-Pfad. Betroffene Ziele gezielt
wiederholt; keine neue CTest-ID oder Vollsuite. Sieben Upload-/Fixturehashes
unverändert. Kein neuer Windows-/Distributions-/DEB-Nachweis; letzter geprüfter
DEB-Stand7f40893 liegt vor archivierten und typisierten Klassenerweiterungen.


## Benannte Projektklassen-Arrays (2026-11-09)

Feste/dynamische lokale und Modul-/Klassenfelder, ReDim/Preserve, geprüfte Set-
Elementzuweisungen, gleiche ByRef-Arrayparameter und Function-/Get-Rückgaben
verwenden jetzt die reale gebundene Projektklasse als Elementtyp. Initialisierte
und durch Erase zurückgesetzte Plätze enthalten Nothing; falsche Objekt- oder
Arraytypen scheitern13. Kopien behalten getrennte Arraywerte und gemeinsame reale
Objektreferenzen. Klassenname/Enum-Grenzen werden einschließlich expliziter ReDim-
Typangaben vor Effekten gebunden. Ein Array wird nicht als skalares Klassenobjekt
behandelt. VarType ist8201; VBA-TypeName löst den kanonischen tatsächlichen Namen
im Projekt auf, auch für undimensionierte Rückgaben. Debugger-/C++-Typanzeigen
behalten die normalisierte Deklarationsschreibweise, ohne interne Typpräfixe.

For Each behält die normativ erforderliche Variant-Schleifenvariable für Arrays;
die Klassentypisierung der Elemente hebt diese Regel nicht auf (MS-VBAL5.4.2.4).
Reale Rückgabearrays werden einmal enumeriert. Bestehende Rang-/Bounds-/Preserve-/
Speicher-/ByRef-Element- und Mutationsgrenzen bleiben erhalten. Keine allgemeinen
Collection-/DAO-Typarrays, AsNew, Interfaces oder vollständige Referenzparität.
Dies ersetzt die vorherige Klassenarray-Grenze für den beschriebenen Umfang.

App gebaut; vorhandenes Kern-/GUI-Bündel nach Fixture-/Binderkorrektur2/2 bestanden
(22,53s:4,01/18,52s). Erste Prüfung fand die falsche ForEach-Fixture-Deklaration;
danach war der skalare Objektbindungszweig für ganze Klassenarrays zu korrigieren.
Nur betroffene bestehende Ziele wiederholt, keine neue CTest-ID/Original-/Windows-
oder Vollsuite. Der GUI-Archivtest umfasst gewöhnliche importierte Klassen mit
solchem Arrayfeld und erhält den ursprünglichen Archivhash. Kein neuer DEB-Nachweis.
F042/F043 bleiben offen; Katalog456/386/70 unverändert. Ohne Zwischenabschluss an
nachgewiesenen benutzerdefinierten Klassen-Standardmitgliedern und weiteren offenen
Anforderungen weiterarbeiten.


## Deklarierte Klassen-Standardmitglieder (2026-11-09)

Ein tatsächlich vorhandenes Memberattribut VB_UserMemId=0 bindet jetzt einen
öffentlichen Property Get oder eine öffentliche Function als Klassenstandard.
Die ursprünglichen Attributzeilen/Zeilennummern bleiben erhalten; der Parser
überspringt nur den passenden geschlossenen0-Attributsatz der gewählten Prozedur.
Unbekannte/private/nicht lesbare Standardmitglieder, widersprüchliche Namen und
nicht implementierte andere Member-IDs/Enumeratoren werden vor Konstruktion
abgewiesen. Ein unannotiertes Item wird nicht zum Klassenstandard erfunden.
Ohne deklarierten Standard schlägt ein direkter Standardzugriff438 fehl.

Implizite skalare Auswertung verwendet nur einen Standardleser, der ohne
Pflichtargument aufrufbar ist. Getterfehler und die vorhandene16-stufige Grenze
rekursiver Standardwerte bleiben erhalten. Explizite Positions-/Namensargumente,
Missing/optionale Indizes und vorhandene Property-Schreibbinder werden verwendet;
Standardargumente bekommen die echte Signatur, unabhängig von einem anderen
öffentlichen Item-Member. Namensfehler448 werden vor Getteraufruf erkannt.
Standard-Schreiben verwendet tatsächliche Let/Set-Accessoren beziehungsweise
vorhandene geprüfte Rückgabeweiterleitung. Benannte Setterindizes, Klassen-Bang-
Syntax, vollständiges externes ByRef, allgemeine COM-Attribute/-IDs und
Enumeratoren bleiben offen; objektwertige Standard-Set-Kombinationen sind kein
separater Windows-Nachweis. Diese Umsetzung ersetzt die frühere pauschale
Standardmitglied-Attributgrenze nur für den beschriebenen0-Pfad.

App gebaut, erstes Laufzeitbündel der vorhandenen Kern-/GUI-Ziele2/2 bestanden
(22,36s:4,07/18,29s). Gemeinsame technische Quellen prüfen Default-Get/Let,
impliziten Skalarwert, optionale/namentliche Indizes, unverwechselbares Item,
private Zustände und widersprüchliche/andereIDs. Der echte Projektpfad sowie eine
gewöhnliche archivierte Defaultklasse und die unveränderte ursprüngliche
Archivsignatur sind geprüft. Keine neue CTest-ID/Original-/Windows-/Vollsuite.
Katalog456/386/70 unverändert, F042/F043 bleiben offen. Nach privater Veröffentlichung
aktuelles Paket prüfen und ohne Zwischenabschluss weitere Anforderungen bearbeiten.
