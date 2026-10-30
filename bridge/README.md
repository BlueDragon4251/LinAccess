# Access-2003-VBA aus einer MDB exportieren

Die originale `.mdb` enthält VBA gegebenenfalls in Formular-/Berichtsmodulen.
LinAccess kann diese *eingebetteten VBA-Streams* derzeit **nicht direkt unter
Linux** extrahieren. Der portable VBA-Quelltext muss zunächst auf einem
Windows-Rechner mit installiertem Desktop-Microsoft-Access exportiert werden.

1. LinAccess-Quellarchiv auf dem Windows-Rechner entpacken.
2. Die `.mdb` dort bereitstellen. Nur mit eigenen/vertrauenswürdigen Datenbanken
   arbeiten; die Exportbrücke öffnet eine temporäre Kopie.
3. In PowerShell (5.1+) aus dem entpackten Quellordner ausführen:

   ```powershell
   powershell.exe -NoProfile -File .\bridge\Export-LinAccessBridge.ps1 `
       -DatabasePath "C:\Daten\Keksrezepte.mdb" `
       -OutputDirectory "C:\Daten\Keksrezepte-Bridge"
   ```

   Wenn **ausdrücklich Microsoft Access 2003** und nicht das meist als
   Standard registrierte Access 365 verwendet werden soll:

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\bridge\Export-LinAccessBridge.ps1 `
       -DatabasePath "C:\Daten\Keksrezepte.mdb" `
       -OutputDirectory "C:\Daten\Keksrezepte-Bridge-2003" `
       -AccessProgId "Access.Application.11" `
       -RequiredAccessMajorVersion 11
   ```

   `Access.Application.11` bezeichnet den COM-Eintrag von Access 2003;
   in Mehrfachinstallationen kann die Registrierung dennoch anders liegen.
   **Darum prüft die Bridge zusätzlich die tatsächliche Versionsnummer und
   bricht bei Abweichung ab, bevor sie die MDB öffnet.** Falls das COM-Objekt
   nicht registriert ist, kann das Skript Access 2003 nicht erzwingen.
   Ein neuer, leerer Ausgabeordner ist für jeden Versuch notwendig.

4. Im Ausgabeordner `manifest.json` prüfen. Jedes Modul hat einen Status:
   `source_only`, `failed` oder `unavailable`. Ein Export kann wegen
   Office-Sicherheitseinstellungen oder projektgeschützter VBA-Quellen scheitern;
   das Skript ändert diese Einstellungen **nicht**.
5. Den gesamten Exportordner auf Linux übertragen, LinAccess-Projekt öffnen,
   `Bridge-Export öffnen` wählen und den Ordner importieren.
6. Importierte VBA-Module sind im Explorer als Quellobjekte sichtbar.
   Doppelklick, Prozedur auswählen, zur **eingeschränkten** Ausführung bestätigen.
   Für `Me!Feld` und `DoCmd` ein eigenes natives LinAccess-Formular erstellen
   und dessen Designerfunktion `VBA aus Bridge übernehmen` benutzen.

Der Import prüft SHA-256 jedes Artefakts, erhält Originalbytes und führt beim
Import **keinen** VBA-Code aus. Vollständige VBA-/Access-/DAO-/COM-Kompatibilität
ist auch nach erfolgreichem Export **nicht** gegeben. `.mdb` selbst und
Formulare/Berichte werden nicht 1:1 in ausführbare Qt-Layouts übersetzt.

Ein vom Nutzer mit Microsoft Access 365 exportiertes Access-2003-Projekt
mit 26 Quellobjekten wurde erfolgreich in LinAccess importiert. Das beweist
**keine** Access-2003-COM-Exportfunktion und keine allgemeine VBA-Kompatibilität.
Für neue Exporte stehen Version und ProgID im Manifest. Maßgeblich bleiben
die tatsächlichen `manifest.json`-Statusangaben.
