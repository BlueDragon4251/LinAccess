# Native application options and menus

The main menu now follows Datei/Bearbeiten/Ansicht/Einfügen/Extras/Fenster/?.
Extras > Optionen opens an actual native settings dialog, not an inert Access
menu replica. It edits the implemented classic interface toggle, installed font
family and point size6..72. Apply changes the current UI and writes QSettings;
OK applies and closes; Cancel discards changes since the last Apply. Persistence
errors show an actual warning while preserving the applied session settings.

The configured application font is restored by the same startup initializer
used in main, including when classic style is disabled, and survives style
toggling. Imported controls with explicitly assigned source fonts continue to
use their own Qt fonts. There is no bundled Microsoft font, complete Access
options dialog, exact Windows rendering or pixel oracle claim. The ? menu opens
real information about LinAccess and the incomplete compatibility status.

Existing main-window cases use an isolated temporary Ini settings location and
explicit synthetic organization/application names. They check Cancel before
Apply, actual application font and disk settings after Apply, Cancel after
Apply, startup initializer restoration, reopened controls and OK persistence.
The first attempt exposed the test application's missing QSettings organization
name and timed out in the genuine persistence warning; the fixture was corrected.
The real main application already has an organization name.

Validation: final main-window1/1 (2.71 s), existing layouts1/1 (10.70 s), Debian
package/build ID/pinned SQLite/four original fixture hashes passed. Last full
suite before these focused changes48/48 (82.86 s); latest headless core1/1
(1.54 s), unchanged by this GUI-only change. Catalog456/384/72 remains fixed;
full menu/options/Windows presentation parity is open under existing F-IDs.
