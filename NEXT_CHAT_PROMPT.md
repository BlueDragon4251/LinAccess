Neueste Ergänzung nach093fcb8: atomarer ALTER-COLUMN-COUNTER-Reseed für
bestehende LEERE verwaltete32-Bit-Zähler; Namen kanonisch, Nichtcounter/nonempty/
rowid-Alias explizit verweigert, Rollback stellt bisherigen Stand/Schritt wieder
her. Native Generatorsteuerungen sind schon im Formular readonly; normale Felder
weiter editierbar. Bestehende Ziele2/2 (1,35s) plus betroffene Grenzen1/1 (0,24s),
MainWindow/Layout2/2 (14,89s), App gebaut. Details Handoff-Ende. Keine vollständige
ALTER-/Mode-/Random-/Replikationsparität; F171 und456/384/72 unverändert.
Nutzer verlangt ausdrücklich KEINE Zwischenabschlüsse; kontinuierlich arbeiten.

Neuester Stand nachfa9c41a: Anwendungsschreibwege erfassen Generatoraufrufe
in geschachtelten AccessCounterSqlScope-Operationen. Bei SQL-/Batchfehlern werden
verbrauchte Nummern per Compare-and-Set erhalten (kein Zurücksetzen fremder
Fortschritte); aktive Workspaces übernehmen sie für Rollback. Eingebunden sind
DAO/QueryDef/SQL-Editor, Formularsave, Datenblattinsert/batch und Recordappend.
Gebündelte bestehende Integration5/5 (15,81s), App gebaut; Handoff-Ende beachten.
Externe/raw-Qt-Queries, Crash-/Concurrent-App-/vollständige Event-Parität sind damit
nicht belegt. F171/456/384/72 bleiben offen/unverändert. Nicht erneut breit testen,
sondern verbleibende vollständige Anforderungen weiter implementieren.

Neueste Ergänzung: normales lokales CREATE TABLE mit COUNTER und
COUNTER(seed,increment) ist implementiert: tatsächlicher Spaltentyp-Token,
signed-Long-Literalseed/positive Long-Schrittweite, SQL+DAO und Reopen, optionale
_la_counter32_steps, atomare Initialisierung/Drop. Access-Referenz100/107/114 und
Cancel/Rollback/NoPK/ungültige Schrittweite im vorhandenen Import-/DAO-Ziel2/2
(1,30s) geprüft. Das ist KEINE vollständige binäre MDB-Modus-/Schrittweitenerkennung,
keine Random/Replikations-/ALTER-/Property-Gleichheit. Handoff-Ende lesen.

Aktuelle Ergänzung 2026-11-08: Nutzer will größere zusammenhängende Implementierung,
danach eine gebündelte gezielte Prüfung; keine Zwischenabschluss-/Testflut. Allgemeiner
Access-Ersatz, keine Kassa-Spezialisierung. Access2003/DAO3.6-v3/v4-Referenzen sind
vorhanden; Details und Grenzen vollständig am Ende von CODEX_HANDOFF.md lesen.
F171 bleibt offen, aber der neue Stand umfasst gespeicherte Quellstände,
DAO-Reservierung vor AddNew-Rückgabe, verbrauchte Abbruch-/Updatefehler-/Workspace-
Rollback-Nummern, signed32-Übergang und verwaltete Vergabe für alle21 Original-
Zähler sowie neue Designertabellen. Long-Typweitergabe, NULL-/Änderungsschutz und
atomare Lifecycle-Grenzen sind implementiert. Kein Legacy-Autoumbau. Noch offen:
vollständige Quellmodi, Random/Replikation, ALTER/Reseed/allgemeine
Properties, SQL-Statementfehler-/Formevent-Allokation und weitere Lifecycle-Grenzen.
Technische Access-Counter-Oracles: tests/access2003_reference/.
Private v4-SQL-/DAO-Setupfehler nicht als Ergebnisparität interpretieren. Keine
erneute pauschale Referenzanforderung und kein Abschlussanruf für diese Teiletappe.
Tatsächlichen Remote/HEAD sowie nachfolgende Prüfergebnisse zuerst prüfen.

Setze die Entwicklung von BlueDragon4251/LinAccess auf development/access-compatibility fort.

Aktueller Fortsetzungspunkt (08.11.2026; alle historischen Nachweise unten sind snapshotbezogen)
- Aktuell implementiert: echter Datenblattdruck/PDF/Vorschau/Seiteneinrichtung, DAO-Metadaten/Chunks/Beziehungen/Originalindexkatalog, echte Jet-Rohbinärdaten und bitgenaue endliche Single-/Double-Übernahme. Bei Fortsetzung nicht beim bereits implementierten Datenblattdruck neu anfangen; tatsächliches Git-HEAD prüfen.
- Aktuelle vollständige lokale Desktop-Prüfung:48/48 (143,10 s), aktuelles Paket/Build-IDs/Helfer/SQLite3.54.1 und vier Originalfixture-/drei Uploadhashes. Beiden Kassa-Kopien entsprechen57 Double-Spalten/69405 Nicht-NULLwerte/514 NULLs bitgenau zum unabhängigen Jackcess-Auszug;78/100 Originalabfragen ausgeführt, kein Windows-Ergebnisorakel. Negative Null/Nonfinite, Currency/Decimal-Vollsemantik und viele andere Vollanforderungen bleiben offen.
- Lokaler GUI-Build benötigt zusätzlich zum bestehenden Aktivierungsskript: PKG_CONFIG_SYSROOT_DIR=/workspace/work/mdb-native-probe/sysroot und PKG_CONFIG_LIBDIR=/workspace/work/mdb-native-probe/sysroot/usr/lib/x86_64-linux-gnu/pkgconfig (lokal extrahierte mdbtools-dev/GLib-Header, keine Systemänderung). Build-CMake-Prefix bleibt /workspace/linaccess-onboarding/sysroot/usr. Der gepinnte SQLite-Librarypfad muss bei Tests vor dem SDK-Librarypfad stehen.
- Private unabhängige Double-Oracles: /workspace/work/double-oracles/<source_sha256>, Source-Marker und .rows-Dateien; LINACCESS_DOUBLE_ORACLE_ROOT aktiviert die bestehenden Prüfungen. Originalwerte und Protokolle niemals committen. Optionale numerische Jet-Referenzen liegen in /workspace/work/jet-single-precision-reference und /workspace/work/jet-double-precision-reference.
- Codex Call: letzte Anforderung und einmaliges lesendes Resume sind unresolved; Teilantwort „Genau“ liefert keine technische Windows-Referenz oder bestätigte Freigabe. Keine neue Anfrage aufgrund des alten Zustands erzeugen. GitHub Actions ist vor Build wegen Konto-Zahlungs-/Ausgabenlimit blockiert (Run42497872453); kein GitHub-Erfolg behaupten. Unabhängig an den72 offenen Vollanforderungen weiterprogrammieren.

Ziel und Arbeitsweise
- Ziel bleibt vollständige funktionale, optische und Tastatur-Parität mit Access 2003 als native C++20-/Qt6-App für Debian/Ubuntu, mit tatsächlicher Access-/VBA-/SQL-Ausführung. Keine Mockups, erfundenen Ergebnisse oder pauschalen Kompatibilitätsbehauptungen.
- Arbeite kontinuierlich bis zum Abschluss der vollständigen Parität. Keine freiwilligen Zwischenabschlüsse, kein „Soll ich weitermachen?“ nach einzelnen Etappen. Nach geprüften Segmenten selbstständig weiterprogrammieren. Nur ausdrücklich angeforderter Chatwechsel/Abbruch oder eine echte, nicht selbst lösbare technische Blockade rechtfertigen eine Unterbrechung. Bei einer Blockade unabhängige Arbeit fortsetzen und die konkrete Grenze nennen.
- Programmierung hat Vorrang. Bestehende gezielte Tests nach Risiko verwenden; keine Testflut, keine neuen Testziele oder Kleinstpunkte zur Fortschrittssteigerung. Notwendige Integritäts-, Build-, Paket- und Regressionsprüfungen bleiben erforderlich. Keine Quelldateien während laufender Builds/Tests ändern.
- Regelmäßig kurze deutsche Statusmeldungen mit Balken je Segment. Zusätzlich „Paritätskatalog (ungewichtet): ████████░░ 384/456“. Das ist theoretische Katalogabdeckung bestätigter Teilumfänge, kein Nachweis von 84 Prozent oder vollständiger Access-Parität. Keine erfundenen Gesamtprozentsätze. Nur nachgewiesene Funktionen als fertig markieren.

Privatsphäre
- Das Repository MUSS privat bleiben. Vor Veröffentlichungen den tatsächlichen Sichtbarkeits- und Remote-Stand prüfen. Keine sensiblen Daten, Original-VBA-Quellen, Zugangsdaten oder Originaldaten in öffentliche Logs, PRs, Releases oder Artefakte übernehmen.
- Mit Testkopien arbeiten. Alle vier Originalfixture-SHA-256-Werte erhalten. Debian-Pakete dürfen keine Fixtures/privaten Quellen enthalten. REA ist optionales externes Analysewerkzeug, keine App-Abhängigkeit und kein Entwicklungsblocker.

Zuerst lesen und prüfen
1. Tatsächlichen Remote-Stand, Git-Status und laufende Builds/Tests prüfen; nicht von alten Chat-SHAs ausgehen.
2. CODEX_HANDOFF.md vollständig lesen und als verbindliche technische Spezifikation behandeln.
3. Danach README.md, docs/COMPATIBILITY.md, docs/DEVELOPMENT_CHECKLIST.md, docs/FIXED_WORK_CATALOG.md und die vorhandenen Fortschritts-/Laufzeitdokumente lesen.
4. Der feste Katalog hat genau 456 IDs F001–F456 in 14 Unterkategorien: aktuell 384 bestätigte Teilumfänge und 72 offene Vollanforderungen. IDs/Gesamtzahl nicht erhöhen, keine Punkte durch Zerlegung vermehren. Neue Details bestehenden offenen Anforderungen zuordnen. Alle offenen Punkte stehen im Katalog; auch breite COM-/VBA-/Jet-/Designer-/Berichts-/UI-/Tastaturanforderungen bleiben verbindlich.

Historischer Ausgangsstand vor dieser Fortsetzung
- Letzter vollständig geprüfter Laufzeitcommit: 6f77ac842e169ba34c144372dde6b682f82e50c0; identischer privater Sicherungscommit: 1a43d9c03c314bb551e16147450f4d9a5379c867; Laufzeit-Tree: bcc7d8b6122b8f2abf7df9558e4ac30a88bd20d4.
- Beide Verläufe wurden danach ohne Inhaltsänderung mit 11751d4dc270e3c237514802426db8beb33d4b1f zusammengeführt. Nachfolgende Übergabedokumentation ist zusätzlich gepusht. Der Branch wurde ohne Force aktualisiert; trotzdem den tatsächlichen Remote-HEAD zuerst prüfen.
- Aktueller Gesamtlauf: 48/48 CTests bestanden (86,11 s). Debian-Paket, Build-ID, SQLite 3.54.1 und alle vier Originalfixture-Hashes bestanden. GUI-freier Kern zuletzt 1/1 (1,92 s); die anschließenden Druckänderungen betreffen ausschließlich GUI-Code. Originalabfragen: 78/100 tatsächlich ausgeführt; kein Windows-Access-Ergebnisorakel.
- Die ursprünglichen Formular-/Report-Lebensdauer-, aktives-Dokument-, Quit-/Unload-Abbruch- und dynamischen ComboBox-Arbeiten sind inzwischen implementiert und geprüft; nicht erneut als verlorene uncommittete Arbeit behandeln.
- Zusätzlich vorhanden: native Dateioperationen und typisierte Scripting.FileSystemObject/File/Folder/Files/Folders/TextStream-Objekte; reale Verzeichnisiteration, Metadaten, Copy/Move/Delete/Name; ASCII- und explizite UTF-16LE-BOM-Streams mit Lesen, Schreiben, Anhängen, Skip/SkipLine und begrenzter Zeilenende-Abfrage. Kodierungs-, Pfad-, EOF-, Sharing- und COM-Grenzen stehen ausdrücklich in docs/NATIVE_FILE_SYSTEM_OBJECT.md und docs/NATIVE_TEXT_STREAM.md. Keine vollständige FSO-/COM-Parität behaupten.
- Strg+P im aktiven Hauptformular respektiert KeyPreview/KeyDown-Abbruch und öffnet den echten Druckdialog. Formulare/Berichte teilen Lebensdauer-/Reentry-Schutz; Formularausgabe bleibt aktuelles Layout. Details: docs/FORM_PRINT_SHORTCUT_RUNTIME.md. Keine physische Druckausgabe wurde für Tests ausgelöst.

Aktuelle Ausgabefortsetzung (2026-11-06)
- Tabellen-/Abfragedatenblatt-Druck und atomare PDF-Ausgabe sind jetzt begrenzt implementiert: docs/DATASHEET_PRINT_RUNTIME.md. Reale sichtbare Spalten/Zeilen, Filter/Sortierung, vollständiges begrenztes Nachladen und eigener Snapshot vor Dialogen; Entwürfe/Binär-/unrepräsentierbare Inhalte werden ausdrücklich abgewiesen. PDF-Dateien einschließlich PDF aus dem Druckdialog werden atomar ersetzt. Lebensdauer-/Projekt-/Reentry-Schutz und echte mehrseitige Textextraktion sind geprüft; keine physischen Druckjobs.
- Aktueller Vollstand: Desktop48/48 (92,62 s), frisch GUI-freier Kern1/1 (1,90 s), aktuelles Debian-Paket/Build-ID/SQLite3.54.1 und vier Originalfixture-Hashes bestanden. Zusätzlich beide Original-Rezepttabellen/39 Zeilen aus erneut hochgeladener MDB als PDF auf isolierter Kopie geprüft. Keine neue Katalog-ID/kein neues Testziel;456/384/72 unverändert.
- Neu geliefert: Keks-MDB und Kassa-Bridge-ZIP sind bytegleich zum bisherigen Fixture. Die neue Kassa-MDB hat SHA-256 d9ff0f9d002b3ab59b3336a34e6c09d9c058d5b9b944b754259a78e277e77102 und31227904Bytes; altes Fixture unverändert. Gleiches exportiertes Schema,28/29 Tabellen CSV-identisch; tblAnmeldung ist im Upload leer gegenüber einer bisherigen Zeile. Nichts ergänzen/erfinden. Neue Kopie unter /workspace/work/reference-20261106/kassa-upload.mdb bestand die bestehenden Import-/Zellwert- und Abfrageinventar-Prüfungen:78/100 ausgeführt. Keine Windows-Ergebnis-/Ereignis-/Druckreferenz daraus ableiten.
- GitHub Actions startet wegen Konto-Zahlungs-/Ausgabenlimitfehlern aktuell keine Jobs (Baseline c0ed5d1, Run42221275422). Keine frische erfolgreiche Distributionsmatrix behaupten.
- Echte paginierte Datenblatt-Seitenansicht mit Navigation/Zoom sowie DoCmd/Application RunCommand acCmdPrint340/acCmdPrintPreview54 sind anschließend implementiert und gezielt geprüft: Kern/MainWindow/Layout3/3 (18,86 s), frisch GUI-freier Kern1/1 (1,77 s), aktuelles Paket/Build-ID/SQLite/vier Hashes bestanden. Letzter einzelner Vollsuite-Lauf48/48 gehört zu40c6b13 vor diesem Inkrement. Wahre9-Punkt-Metrik im PDF per Glyphhöhe geprüft. Keine neuen Katalogpunkte. Weitere feste offene Anforderungen selbstständig bearbeiten; tatsächlichen Git-Stand prüfen.

Weitere geprüfte Fortsetzung
- Unveränderliche Quellkopien von e2d64c9 bestehen lokal Debian12/Qt6.4 48/48 (101,29 s) und Ubuntu24.04/Qt6.4 48/48 (93,65 s), beide mit Paket-/Build-ID-/SQLite-/Installations-/Hashprüfung. Das sind lokale Containerprüfungen; GitHub-CI bleibt gesperrt.
- Danach native Recordset.Type-Metadaten ergänzt: aufgezeichneter Cursor2/4 als readonly Integer, getrennt von SQL-Readonly, Clone-Erhalt und echte Type-Feld-Bang-Auflösung. Unbewiesene gefrorene Formularadapter3251; kein Table-/ADO-/Locking-Vollnachweis. DAO-Ziel1/1 (1,02 s), aktuelles Paket/Build-ID/SQLite/vier Hashes bestanden. Noch keine neue Vollmatrix dieses letzten Inkrements behaupten. F054/F055 bleiben offen;456/384/72 unverändert.

Aktuellste Nachweise und Moduskorrektur
-418719d bestand anschließend Debian13 Vollsuite48/48 (94,15 s); unveränderliche418719d-Folgekopien bestanden auf Debian12/Ubuntu24.04 aktualisierte App/DAO1/1 (1,55 s /1,06 s), Paketinstallation/SQLite/vier Hashes.
- Danach gemeinsame geprüfte Long-Konvertierung der nativen OpenRecordset-Typargumente ergänzt: banker rounding, Missing/ausgelassen nativer Default2, Empty0/andere Modi3251, Null94/Text13; keine1/Table-/ADO-/Locking-Parität. DAO1/1 (0,97 s), aktuelle App/Paket/Build-ID/SQLite/vier Originalhashes und drei Uploadhashes bestanden. Breite Vollsuite/Matrix dieses letzten kleinen Inkrements nicht neu behaupten. F055/F122 bleiben offen.
- Codex Call ist für notwendige Fragen und tatsächlichen Gesamtabschluss nach aktueller Nutzeranweisung zu verwenden. Skill-/Client-Checkout: /workspace/work/codex-call, feat/ask-via-call fd2cea7; vorhandene Schnittstellen lesen, keine Secrets ausgeben. Erste Referenzanfrage endete timeout ohne API-Antwort; Nutzer lieferte Dateien im Chat. Die weitere gebündelte Original-Login-/Windows-Referenzfrage endete failed: "The recipient ended the call before agent finalization." Keine Antworten/Entscheidungen erhalten; Ergebnis unter /workspace/work/login-reference-call-result.json. Nichts aus Schweigen ableiten, nicht blind neu anlegen. Kein Gesamtabschluss oder Abschlussanruf:72 Vollanforderungen offen.

Unbedingt offen lassen, bis korrekt gelöst
- Original-Login: ColumnCount 5 gegenüber Column(5). Keine sechste Spalte, kein Passwort, kein Null-Fallback erfinden; Access-Referenz fehlt.
- Verkaufsinitialisierung: Fehler3075 wegen fehlender aktiver Verkaufs-ID; keinen Geschäftszustand erfinden.
- Verbleibende Originalabfragen haben echte fehlende Abhängigkeiten/Spalten/Bindungskontexte und zwei geschützte Joined-Update-Grenzen. Schutzprüfungen nicht entfernen, um Zahlen zu erhöhen.
- Vollständiges Jet/ACE-Schreiben, OLE-/Attachment-Roundtrip, Decimal/Kollation/Metadaten, vollständiger VBA-Compiler/Debugger/Referenzbinder, COM/ActiveX/WinAPI, vollständige Designer/Berichte/Druckereignisse, Access-2003-Optik/Tastatur/IME und Windows-Referenzvalidierung bleiben im festen Katalog offen.

Falls der bisherige Workspace erhalten ist
- Repository: /workspace/LinAccess. Geschützte lokale Logs/Checkpoints: /workspace/linaccess-onboarding, insbesondere private-api-checkpoint.json.
- Aktivierung: . /workspace/linaccess-onboarding/activate.sh. Build: build; GUI-freier Kern: build-core. Für GUI-Tests LD_LIBRARY_PATH auf build/third_party/sqlite ergänzen, QT_QPA_PLATFORM=offscreen und QT_STYLE_OVERRIDE=Windows verwenden.
- Alte WIP-Patches nicht blind anwenden: ihre Arbeiten sind inzwischen in späteren geprüften Commits enthalten. Aktuellen Git-Status und Handoff vor jeder Wiederherstellung abgleichen.
- GitHub-API-Ref-Updates scheiterten wiederholt serverseitig. Direkter Git-Zugang funktionierte beim Chatwechsel; damit wurden die verifizierten privaten Commits ohne Force gepusht und die Historien sauber zusammengeführt. API-Fehler nicht mit fehlgeschlagenem aktuellen Git-Push verwechseln.

Nach jeder sinnvollen Programmieretappe angemessen prüfen, Dokumentation aktualisieren, committen und privat pushen; anschließend selbstständig weiterarbeiten. Den Gesamtauftrag erst als abgeschlossen bezeichnen, wenn die vollständige Parität tatsächlich implementiert und nachgewiesen ist.

DAO-Abfragefeld-Fortsetzung: QueryDef.Fields für vorbereitbare SELECTs implementiert; Metadaten ohne Zeilenausführung/UDF-Aufruf, konservative Type/Size, Lebensdauer-/Schema-/Rollback-Invalidierung. Grenzen und Nachweise in docs/DAO_RUNTIME.md;456/384/72 unverändert.

Nachweis dieser QueryDef.Fields-Etappe: bestehendes DAO-Ziel1/1 (0,96 s), anschließend vollständige Debian13-Desktop-Suite48/48 (98,29 s). Aktuelles Debian-Paket/Build-ID/SQLite3.54.1 und vier Originalfixture-Hashes bestanden. Keine neue Distributionsmatrix/GitHub-CI behauptet.

Metadatenanbindung: vorhandenes DAO-Ziel1/1 (0,99 s) einschließlich tatsächlichem typisiertem VBA, benannten Zugriffen, Refresh und fremdem Thread bestanden; aktuelles Paket/Build-ID/SQLite3.54.1 und vier Originalhashes bestanden. Die Vollsuite48/48 (98,29 s) gehört zum vorherigen a00eb29; kein neuer Vollsuite-/Matrixlauf für dieses kleine Inkrement behauptet.

Parameterfortsetzung: bestehendes DAO-Ziel1/1 (1,09 s), aktuelle App und Debian-Paket/Build-ID/SQLite3.54.1/vier Originalhashes bestanden. Named Index/Refresh und Werterhalt durch echtes VBA geprüft; letzter Vollsuite-Lauf bleibt a00eb29 mit48/48 (98,29 s). Kein neuer Katalogpunkt, keine vollständige Parameter-/Jet-Parität.

TableDef.RecordCount: DAO/originales Kassa-Ziel2/2 (53,71 s) bestanden, einschließlich unabhängiger Zählung aller29 Originaltabellen. Aktuelle App/Paket/Build-ID/SQLite3.54.1 und vier Originalfixture-Hashes bestanden. Bestehende unveränderte Syntaxinventur separat187/187 akzeptiert (1,12 s), keine Ausführungsaussage. Kein neuer Vollsuite-/Matrixlauf dieses Inkrements.456/384/72 unverändert.

Binär-GetChunk: bestehendes DAO-Ziel1/1 (1,23 s), aktuelle App/Paket/Build-ID/SQLite3.54.1 und vier Originalfixture-Hashes bestanden. Nur gespeicherte positive native Binärbereiche; AppendChunk/OLE/Memo/Null/Randsemantik bleiben offen. Letzte Vollsuite a00eb29:48/48 (98,29 s), nicht diesem Inkrement zuschreiben.456/384/72 unverändert.

Native binäre Field.AppendChunk-Fortsetzung implementiert: Byte()-Chunks im Edit/AddNew-Puffer, erster Chunk ersetzt, folgende hängen an; je Feld getrennt, 16MiB, Update/CancelUpdate/Session-Reset und geschützte Mischverwendung. Genauer begrenzter Umfang in docs/DAO_BINARY_FIELDS.md. F016/F055/F121/F122 bleiben offen,456/384/72 unverändert.

Binär-AppendChunk-Nachweis: finales bestehendes DAO-Ziel1/1 (1,38 s), danach vollständiger aktueller Debian13-Desktop48/48 (98,95 s), aktuelles Paket/Build-ID/SQLite3.54.1 und vier Originalhashes bestanden. Byte-Value-Zuweisung bleibt separat geprüft; Mischverwendungsfehler bewahren den Puffer. Kein neuer Test/Katalogpunkt.
Unveränderlicher629b3c9-Folgecheck: Debian12 DAO1/1 (1,49 s), Ubuntu24.04 DAO1/1 (1,27 s), jeweils aktuelle App/DAO neu gebaut, passende Pakete erzeugt/tatsächlich installiert, SQLite3.54.1 und vier Originalhashes bestanden. Diese629-Nachweise enthalten noch nicht die anschließenden TableDef.RecordCount-/Binärchunk-Inkremente; keine neue Vollmatrix/GitHub-CI behaupten. Letzte früheren vollständigen lokalen Distributionssuiten bleiben e2d64c9 mit48/48 je Plattform.456/384/72 unverändert; kein Gesamtabschluss.

Recordset.Field.Name prüft jetzt wie die übrigen Feldmetadaten die tatsächliche Besitzer-/Verbindungs-/Rollback-Lebensdauer. Der Name bleibt ohne aktuellen Datensatz am EOF lesbar, nach Close aber3420. Vorhandenes DAO-Ziel1/1 (1,05 s), aktuelles Paket/Build-ID/SQLite3.54.1 und vier Originalhashes bestanden. Letzte Vollsuite41b193c:48/48 (98,95 s), kein neuer Gesamtlauf für diese Lebensdauerkorrektur.456/384/72 unverändert.

Recordset.Fields-Fortsetzung: benannte/geprüfte Indizes, kanonische Ergebnisnamen, gehaltene Collection-/Field-Identität auch bei Enumeration, Null94/fehlend3265/Arity450 und explizite Doppelspalten-Verweigerung3251. Bestehendes DAO-Ziel1/1 (1,01 s), aktuelles Paket/Build-ID/SQLite3.54.1/vier Originalhashes bestanden; Details docs/DAO_RUNTIME.md. Kein neuer Vollsuite-Lauf dieses Inkrements.
Unveränderliche41b193c-Folgeläufe bestanden Debian12 DAO1/1 (1,61 s) und Ubuntu24.04 DAO1/1 (1,21 s), jeweils aktuelle App/Paket/Installation/SQLite3.54.1/vier Originalhashes. Diese Belege enthalten TableDef.RecordCount/GetChunk/AppendChunk, nicht die danach folgenden Field.Name-/Recordset.Fields-Inkremente. Letzte komplette Host-Suite41b193c:48/48 (98,95 s).456/384/72 unverändert; GitHub-CI weiter gesperrt.

Datenblatt-Seiteneinrichtung: funktionierender gemeinsamer Qt-Geometriedialog, Sitzungslayout und explizite versionierte Projektpersistenz, native RunCommand PageSetup32, Papier/Ränder für Vorschau/Druck/PDF. Conditional Savepoint-Speicherung bewahrt konkurrierende Änderungen und Aufrufertransaktionen; ungültige gespeicherte Geometrie blockiert PDF bis zur expliziten Reparatur. Details docs/DATASHEET_PRINT_RUNTIME.md. Ziel1/1 (3,64 s), vollständiger aktueller Desktop48/48 (89,86 s) inklusive Originalrezepte2 Tabellen/39 Zeilen, Paket/Build-ID/SQLite3.54.1/vier Originalhashes bestanden. Keine neue Distributionsmatrix/GitHub-CI;456/384/72 unverändert, F080/F087/F148 bleiben offen.

Druckexemplare: reale Datenblatt-PDFs mit1..100 Kopien, sortierter/unsortierter und vorwärts/rückwärts Seitenfolge; native Qt-Treiberkopien getrennt von manuellen Kopien. Vorhandenes Ziel1/1 (3,99 s), aktuelle App/Paket/Build-ID/SQLite/vier Originalhashes bestanden. Letzte Vollsuitee8375e8:48/48 (89,86 s), dessen GUI-freier Kern1/1 (2,20 s). Unveränderlichee837-Folgeprüfungen Debian12/Ubuntu24.04:2/2 (6,65 s /5,65 s), passende Paketinstallation/Backend/vier Hashes; noch ohne Kopien-Inkrement. Details docs/DATASHEET_PRINT_RUNTIME.md.456/384/72 unverändert, keine physische Drucker-/Windows-Parität behaupten.

Auswahldruck: vollständige markierte sichtbare Datensätze als eingefrorener eigener Text-Snapshot, reale Tabellen-/Readonly-Abfrage-PDFs mit unabhängig geprüfter Ein-/Ausschlussmenge. Leere/teilweise Zellmarkierungen verweigern3251 ohne Ganzdatenblatt-Fallback. Ziel1/1 (4,15 s), aktuelle App/Paket/Build-ID/SQLite3.54.1/vier Originalhashes und drei Uploadhashes bestanden. Details docs/DATASHEET_PRINT_RUNTIME.md. Kein neuer Vollsuite-/Matrixlauf dieses GUI-Inkrements;456/384/72 unverändert.

DAO-DDL-Identität: beobachtete eigene Database/QueryDef.Execute CREATE/DROP und SELECT INTO führen monotone Main-Tabellengenerationen; alte TableDefs/Fields/physische Recordsets/Clones verweigern3420 statt neu erstellte gleichnamige Daten zu ändern. Gleiches SQL/Rootpage/Rowid, alter Schreibpuffer, frische Wrapper, No-op/Fehl-DDL, ALTER und Workspace-Rollback unabhängig geprüft. Externes/unbeobachtetes SQL, TEMP-only und volle Jet-Locks/Lebensdauer bleiben offen. Vollständiger aktueller Desktop48/48 (89,70 s), App/Paket/Build-ID/SQLite/vier Originalhashes bestanden; schließt Kopien/Auswahldruck und Originalrezepte ein. Details docs/DAO_RUNTIME.md.456/384/72 unverändert.

Jet-Rohbinärimport: ausdrücklich gewählter MDBTools-Hexcodec (kein Raten aus CSV), exakte BLOB-/NULL-Werte und Byteanzahl-Vorschau im realen Teil-/Gesamtimport. Laufend begrenzte Child-Ausgabe20MiB/Diagnostik64KiB,30s; großer Export bricht atomar ohne Quelldateiänderung ab. Technische echte Jet4-MDBs mit Jackcess5.0.4 unabhängig geschrieben/zurückgelesen, keine Geschäftsdaten und keine neue App-Java-Abhängigkeit. Decoder-/GUI-/DAO-Bytes/Typ/Chunks/Hash/Rollback geprüft. Ziel2/2 (12,26 s), vollständiger aktueller Desktop48/48 (91,47 s), neue Kassa-Kopie Import/Zellwerte und78/100 Abfragen, App/Paket/Build-ID/SQLite/vier Originalhashes bestanden. Details docs/JET_BINARY_IMPORT.md; keine OLE-Anzeige/Anhänge/Originalformat-Schreib- oder Windows-Parität behaupten.456/384/72 unverändert; kein neuer CTest.

DAO-Katalog-Count liefert nun für Parameters/QueryDefs/TableDefs/Workspaces/Databases nach öffentlicher statisch geprüfter PIA-Int16-Signatur200006 einen kontrollierten VBA Integer, keine generischen Double/Long. Echtes VBA VarType/NativeSubtype geprüft; Ziel1/1 (1,06 s), aktuelle App/Paket/Build-ID/SQLite/vier Originalhashes bestanden. Details/PIA-Hashes docs/DAO_COUNTS_RUNTIME.md; keine COM-Ausführung/Access2003-Oraclebehauptung. Letzte Vollsuite6cd40d1:48/48 (91,47 s).456/384/72 unverändert.

DAO.Relations: lesender tatsächlicher Importkatalog mit Name/Table/ForeignTable/Attributes und geordneten Fields.Name/ForeignName; Integer Count, benannte Indexargumente, echte VBA-Typen/TypeOf/Enumeration, gehaltene Identität und kontrollierte Invalidierung. Erzwungene Beziehungen werden mit allen tatsächlichen Main-Fremdschlüsseln und CASCADE-Aktionen abgeglichen; unbekannte Schlüssel/uneindeutige Metadaten verweigern3251 statt Teilkatalog. Bestehende Ziele2/2 (11,87 s), zusätzlich neue Kassa-Uploadkopie mit acht Originaldefinitionen/acht Feldpaaren bestanden. Details docs/DAO_RELATIONS_RUNTIME.md; Append/Delete/Indexes/Linked und vollständige Jet-/Windows-Parität bleiben offen.456/384/72 unverändert, kein neuer Testkatalogpunkt.

Aktuelle App/Debian-Paketprüfung einschließlich Build-ID/SQLite3.54.1 und vier Originalfixture-Hashes bestanden. Letzte Vollsuite bleibt6cd40d1:48/48 (91,47 s); kein neuer Vollsuite-/Distributionsnachweis für Relations.

Nicht-Primärindexrichtungen: ursprünglicher Parserverlust von DESC behoben; vollständige geordnete ASC/DESC-Feldlisten für CREATE/ALTER INDEX, ungültige Zusätze verweigern statt Normalisierung. Native Indexanlage erhält Richtung, Reihenfolge und Unique in derselben Importtransaktion. Optionale echte Jet4-Referenz ohne Datensätze mit zwei unabhängig zurückgelesenen Indizes, tatsächlicher Gesamtimport und PRAGMA index_xinfo geprüft; keine App-Java-Abhängigkeit. Details docs/SOURCE_INDEX_DIRECTIONS.md. Kein echter Primärindexname aus generierten MDBTools-Namen behauptet; DAO.Indexes/Primärindexrichtungen/Locale/Originalformat weiterhin offen.456/384/72 unverändert.

Indexrichtungsnachweis: vollständiger aktueller Debian13-Desktop48/48 (89,51 s) bestanden, einschließlich tatsächlicher Jet-Binär-/Indexreferenzen, Originalrezepte für Datenblatt-PDF, Originalkassenhandler und aller originalen Kassa-Indexfeldreihenfolgen/-richtungen. Schließt DAO.Relations und Integer-Count ein. Keine neue Distributionsmatrix/GitHub-CI für diese Inkremente;456/384/72 unverändert.

Aktuelles Debian-Paket/Build-ID/SQLite3.54.1, vier Originalfixture- und drei Uploadhashes bestanden.

DAO-Single-Werte: bestätigte physische Felder mit DAO-Typ6 liefern nun bei Value/Bang/Clone und im Edit/AddNew-Puffer echte Single-Werte; Schreiben verengt vor Puffermutation und persistiert binary32-Wert, Null/Empty und explizite Konvertierungsfehler bleiben geprüft. Echte technische Jet4-Float-Referenz mit unabhängig zurückgelesenen Werten, tatsächlicher Import/Metadaten/Update/Requery/Quellhash und VBA VarType/CSng im bestehenden Ziel2/2 (1,94 s) bestanden. Keine automatische Query-/Formularlineage und keine verlustfreie allgemeine Single-Importbehauptung: MDBTools1.0.0 %.8g kann einzelne Bitmuster verändern, belegtes Gegenbeispiel in docs/DAO_SINGLE_FIELDS.md. F014/F041/F055/F105/F122 bleiben offen,456/384/72 unverändert. Letzte Vollsuite4548ce7:48/48 (89,51 s), nicht diesem Typinkrement zuschreiben.

Aktuelle App/Paket/Build-ID/SQLite3.54.1 und vier Originalfixture-/drei Uploadhashes bestanden.

Originale Indexmetadaten: separater lesender nativer MDBTools-Helfer statt generierter CLI-Primärnamen. Strikt versionierte Namen/Kind/Flags/geordnete Felder/Richtungen, gegen tatsächliche Tabellenspalten und ursprünglichen Primärschlüssel validiert; atomare main._la_access_indexes-Persistenz. Neue Importe verwenden echte Nichtprimärnamen. Zusammengesetzte und einzelne DESC-Primärschlüssel bleiben echte native DESC-Indizes; DESC-AutoNumber und unbewiesene IgnoreNulls/Required-Konfigurationen verweigern explizit. Bestehende GUI/Import/Kassa-Ziele3/3 (15,27 s), abschließende erweiterte technische Jet-Referenz1/1 (0,96 s);20 echte Primär-/48 andere Originalnamen in29 Tabellen nachgewiesen. Details docs/SOURCE_INDEX_METADATA.md. Noch kein vollständiger DAO.Indexes-/Windows-Nachweis;456/384/72 unverändert. Build benötigt mdbtools-dev/pkg-config, Paket enthält geprüften lesenden Helfer.

Indexmetadaten-Grundlage vollständig verifiziert: aktueller Debian13-Desktop48/48 (85,91 s), neue echte Jet-Referenzen einschließlich einzelner und zusammengesetzter DESC-Primärindizes, Originalrezepte, Originalkassenhandler und Originalindexnamen bestanden. Aktuelles Paket/Build-ID für App und Metadatenhelfer, tatsächliches Laden des entpackten Helfers, SQLite3.54.1 sowie vier Originalfixture-/drei Uploadhashes bestanden. Kein neuer Katalogpunkt und keine vollständige DAO.Indexes-/Windows-Parität;456/384/72 unverändert. Lokale Debian12-/Ubuntu24.04-Folgeprüfung der neuen Build-Abhängigkeit folgt separat.

DAO.Indexes: lesender originaler Indexkatalog mit nativen Typen/TypeOf, Integer Count/Item/named Index/Refresh/Enumeration, echten Namen und Boolean Primary/Unique/Required/IgnoreNulls sowie geordneten Fields.Name/Long Attributes (dbDescending). Jeder Zugriff vergleicht den vollständigen tatsächlichen Main-Indexkatalog, Primärschlüssel/Rowid-Alias, Unique, Felder/Richtungen und BINARY-Collation. Fehlende Quellen, logische kind2/IgnoreNulls/unbekannte oder zusätzliche native Indizes verweigern3251 ohne Teilkatalog. Gehaltene Definitionen/Fields invalidieren kontrolliert nach Schema/Metadaten/Rollback, TEMP bleibt wirkungslos. Bestehende Ziele3/3 (11,69 s), zusätzlich neue Kassa-Kopie mit20 Primär-/48 anderen DAO-Indexidentitäten, aktuelle App/Paket/Build-IDs/Helfer/SQLite3.54.1/vier Originalfixture-/drei Uploadhashes bestanden. Details docs/DAO_INDEXES_RUNTIME.md. Keine Mutation/Foreign/Clustered/DistinctCount/Windows-Parität;456/384/72 unverändert. Letzte Vollsuite bleibt74d75c7:48/48 (85,91 s).

Unveränderliche74d75c7-Folgeprüfung abgeschlossen: Debian12/GCC12/Qt6.4 und Ubuntu24.04/GCC13/Qt6.4 jeweils5/5 gezielte Ziele (30,64 s /27,39 s), neues mdbtools-dev/pkg-config, Paketinstallation, Backend3.54.1, Originalhashes und tatsächlich installierter Helfer mit originalen Jet-Namen/Richtungen unter MDBTools1.0.0 bestanden. Diese Snapshots enthalten die Metadaten-Grundlage, noch nicht die folgende DAO.Indexes-Anbindung; keine GitHub-CI- oder Vollsuitebehauptung für die Matrix.

Exakter Single-Import: realer MDB_FLOAT-Rohwertzugriff im vorhandenen lesenden Helfer (--single-values), IEEE754-Bitmuster statt verlustbehafteter %.8g-Speicherung. Source-SHA256 vor/nach Scan und atomarem Import, echte SINGLE-Spalten/REAL-Affinität, CSV-Zeilen/NULL/8g-Abgleich,1M-Zellen/16MiB-Grenze. Rohwerte bleiben außerhalb des angezeigten SQL-Schemas. Gegenfallbits509022793, minimale Subnormale, maximaler endlicher Single, Brüche/+0/NULL nach unabhängiger Jackcess-Jet-Erzeugung im tatsächlichen Gesamt-/Teilimport und DAO bitgenau nachgewiesen. NaN/Infinity/negatives Null oder inkonsistente CSV-Daten verweigern atomar. Bestehende relevante Ziele5/5 (27,76 s), finaler Affinitätsguard im Importziel1/1 (1,05 s); aktuelle App/Paket/Build-IDs/SQLite3.54.1/vier Originalfixture-/drei Uploadhashes und tatsächlich entpackter Helfer mit bekannten Originalbits/Quellhash bestanden. Details docs/SOURCE_SINGLE_PRECISION.md. Legacy-CSV, Double/Decimal/Currency, generische SQL-/Default-/Query-Semantik und Windows bleiben offen;456/384/72 unverändert. Letzte Vollsuite74d75c7:48/48 (85,91 s), neue Matrix noch nicht behaupten.

Codex-Call-Rückfrage vom08.11.2026 zu echter Windows/Access2003-Referenz und GitHub-Actions-Kontofreigabe: Request3ea2a1b8-75c6-411c-8f85-3cb76c4f3571 wurde mit gespeicherter Call-ID angenommen; lokales Abfragefenster endete timeout, answers/confirmed_decisions leer. Keine Zustellung/Gesprächsannahme oder Freigabe nachgewiesen, keine Antwort erfinden, nicht erneut anlegen. Privater Zustand /workspace/work/reference-blocker-call-20261108/state.json; Ergebnis result.json. GitHub-Run42497872453 (a11b95e) hat keine gestarteten Build-Schritte; Annotation bestätigt fehlgeschlagene Kontozahlungen oder Ausgabenlimit. Keine Abrechnungseinstellungen geändert; unabhängige Entwicklung läuft weiter.

Unveränderliche a11b95e-Folgeprüfung: Debian12/GCC12/Qt6.4 und Ubuntu24.04/GCC13/Qt6.4 jeweils5/5 (28,31 s /27,15 s), Paketinstallation, Originalhashes/Backend und tatsächlich installierter --single-values-Helfer mit unabhängig bekannten Originalbits/SHA bestanden. Diese Snapshots schließen DAO.Indexes und Single-Rohpräzision ein, noch keine spätere Double-Korrektur. Keine GitHub-CI-/Vollsuitebehauptung für diese Matrix.

Double-Import bitgenau: vorhandener lesender Helfer --double-values, binäre64-Bit-Quelle statt %.16g-Rückwandlung. Gemeinsamer streng validierter access_floating_metadata-Codec für Single/Double, Classic-Locale-Formatprüfung ohne Double-Overflow bei der Dezimaldarstellung von MAX_VALUE. Original-Rohwerte aus der bereits geschützten Gesamtimport-Snapshotdatei; strukturleere Anlage scannt Daten nicht doppelt. Low-level API unterscheidet ausdrücklich caller-owned Csv (kein Dateizugriff aufgrund eines Loglabels) von NativeMdbBits; reale Gesamt-/Inspector-Kopien und Originalprüfungen wählen NativeMdbBits, fehlender Nachweis fällt niemals still auf CSV zurück. Dadurch bleiben technische Tests mit echten Schemas und eigens bereitgestellten CSV-Zeilen korrekt getrennt.

Unabhängiger Jackcess-Auszug beider Kassa-Testkopien:13 Tabellen/57 Double-Spalten,69405 Nicht-NULLwerte/514 NULLs, kein negatives Null/NaN/Infinity. Der bisherige Dezimalexport reproduzierte2404 numerische Zeilentupel in zwei Tabellen nicht exakt; neue native Kopien stimmen bitgenau mit allen unabhängigen Tupeln überein. Echte Jet-Extremwertreferenz, Gesamt-/Teilimport, DAO Type7/Size8, Herkunfts-/CSV-Guards und atomare Negativnull-Verweigerung geprüft. Aktueller korrigierter Debian13-Desktop48/48 (143,10 s), aktuelle neue Uploadkopie erneut29/29 und unabhängiger Auszug exakt; dort weiterhin78/100 Originalabfragen ohne Windows-Ergebnisorakel. App/Paket/Build-IDs/SQLite3.54.1/vier Originalfixture-/drei Uploadhashes und tatsächlich entpackter --double-values-Helfer mit bekannten Originalbits/SHA bestanden. Details docs/SOURCE_DOUBLE_PRECISION.md. Zuvor entdeckte technische CSV-Herkunftsfehler behoben, keine Schutzprüfung für Originale entfernt.456/384/72 unverändert; keine neue Matrix für Double behaupten.

Nach einmaligem lesendem Resume derselben Codex-Call-Anfrage3ea2a1b8-75c6-411c-8f85-3cb76c4f3571 wurde status=unresolved und eine normale finale LiveKit-Teilantwort „Genau“ zur Windows-Referenzfrage geliefert; keine konkrete Umgebung/Protokolle/Zugriffsinformation und keine confirmed_decisions. CI-Frage bleibt unbeantwortet. Beide externen Blocker bleiben offen, keine Freigabe erfinden oder neue Anrufanforderung daraus ableiten. Privates Ergebnis resumed-result.json im bestehenden Call-Verzeichnis.

Neu: allgemeiner nativer DAO-Schemaaufbau CreateTableDef/CreateField/Fields.Append/
TableDefs.Append + Delete implementiert, gemeinsame atomare native Typ-/Counter-
Metadaten; DAO/Import2/2 (1,25s), App gebaut. Details/Herkunftsgrenzen in
DAO_TABLEDEF_FIELDS.md und Handoff.456/384/72 unverändert. Allgemeine Indexbuilder
als nächste zusammenhängende Ergänzung, keine erneute Referenzanforderung.

Detached CreateIndex/Index.CreateField/Indexes.Append für eigene neue Tabellen
jetzt implementiert einschließlich echter Primary/Unique/ASC/DESC-Constraints
und atomarem Kollisions-Rollback/Retry. DAO/Import2/2 (1,22s), App gebaut.
456/384/72 unverändert. Fortsetzung: sichere Nichtprimär-Indexmutation an
vorhandenen Tabellen, bestehende vollständige Katalogprüfung beibehalten.

CreateIndex/Indexes.Append/Delete auf vorhandenen vollständig geprüften Tabellen
für Nichtprimär-Indizes umgesetzt, einschließlich Unique-Fehleratomizität,
ASC/DESC, Workspace-Rollback und immutable gehaltenen Buildern. DAO/Import2/2
(1,31s), App gebaut; Details DAO_INDEXES_RUNTIME.md.456/384/72 unverändert.
Danach grafischen nativen Tabellenentwurf an gemeinsamen DAO-Schemaweg anbinden.

Grafischer Tabellenentwurf nutzt jetzt denselben atomaren DAO-Schemaweg. Neue
TEXT255/MEMO/Byte/16-/32-Bit-Felder, tatsächliche Typ-/Indexidentitäten; aus UI
gebaute Tabelle anschließend über DAO indexiert. Mainwindow/DAO2/2 (4,95s), App
gebaut. Keine Bestandsmigration.456/384/72 gleich. Fortsetzung: grafische normale
Indexverwaltung auf Basis der vorhandenen geprüften DAO-Operationen.

Grafische Indexverwaltung über Ansicht → Tabellenindizes bearbeiten implementiert.
Realer GUI-Create/Delete-Durchlauf über gemeinsamen DAO-Weg, unveränderte Zeilen;
Mainwindow/DAO2/2 (5,11s), App gebaut. DAO_INDEXES_RUNTIME.md/README.456/384/72 gleich.
Grundlegende native Felddefinitionen/Eigenschaften/Speichergrenzen weiterführen.

Native primitive Speichergrenzen und explizite Feld.Required/Attributes ergänzt;
INSERT SELECT kann Byte/Integer/Long nicht überlaufen. Tatsächliche Double-Date-
Memo-BLOB-Werte/Speicherarten geprüft; letztes DAO1/1 (1,21s), davor GUI/DAO2/2
(5,00s), App gebaut. Alle vier Fixture-/drei Uploadhashes unverändert, Repo privat.
DAO_TABLEDEF_FIELDS.md,456/384/72 gleich. Fortsetzung: grafische Textgröße1..255.

Grafische Textgröße1..255 umgesetzt, ungültige Größe lässt gesamten Entwurf stehen,
korrigierte Größe12 wird tatsächlich gespeichert. Mainwindow1/1 (3,91s), App
gebaut. DAO_TABLEDEF_FIELDS.md;456/384/72 gleich. Fortsetzung: sichere normale
Feldanlage/-löschung auf bestehenden Tabellen über DAO.

Normale Felder vorhandener Tabellen können über CreateField/Fields.Append/Delete
atomar angelegt/gelöscht werden, inklusive Typ-/Indexmetadaten und geprüften
Abhängigkeitsgrenzen. Reales VBA/named args/Bestandswerte/Fehleratomizität/Rollback,
Mainwindow/DAO2/2 (5,38s), App gebaut. DAO_TABLEDEF_FIELDS.md;456/384/72 gleich.
Fortsetzung: echte grafische Feldanlage/-löschung über denselben DAO-Weg.

Grafische Feldanlage/-löschung über Ansicht → Tabellenfelder bearbeiten umgesetzt;
echter GUI-Durchlauf/Bestandswerte und DAO-Invalidierung alter Cursor/Clones bis
Requery im bestehenden Mainwindow/DAO2/2 (5,11s), App gebaut. Index-/Felddialog
melden gespeicherte Änderungen getrennt von Refreshfehlern. README/DAO_TABLEDEF_FIELDS.md.
456/384/72 gleich. Rückgabe-Reader der Schemaoperationen vor Savepoint-Release
binden; anschließend andere offene Anforderungen fortlaufend umsetzen.

Tabellen-/Feld-/Index-Rückgabe-Reader jetzt vor Savepoint-Release gebunden;
anschließend nur Veröffentlichung, kein nachträglicher DB-Lesefehler als
Mutationsfehler. DAO1/1 (1,25s), App gebaut; letzte GUI/DAO2/2 (5,11s).
456/384/72 gleich. Als Nächstes kontrolliertes Fields/Indexes.Refresh nach
eigener Schemaänderung, alte Items weiterhin immutable/invalid.

Kontrolliertes TableDef.Fields/Indexes.Refresh nach eigener DDL/Rollback umgesetzt;
alte Items/gelöschte oder ersetzte Tabellen bleiben invalid. DAO1/1 (1,23s),
App gebaut. DAO_TABLEDEF_FIELDS.md;456/384/72 unverändert. Allgemeine Relation-
Builder/Katalogmutation fortsetzen, keine vorgetäuschte referenzielle Integrität.

Detached CreateRelation/CreateField/ForeignName/Fields.Append/Relations.Append/
Delete für explizite dbRelationDontEnforce2-Beziehungen implementiert. Persistenz,
Generationsschutz/Abhängigkeiten/Rollback: DAO/Import2/2 (1,33s), App gebaut.
DAO_RELATIONS_RUNTIME.md;456/384/72 gleich. Reale FK-/Cascade-Migration noch nicht
implementiert; nächste tatsächliche Aufgabe erzwungene native Beziehungen.

Reale native numerische FK-/Cascade-Anlage und -Entfernung implementiert, nicht
bloß Flags. Streng bekannter nativer Kind-Schemaumbau, Digestprüfung aller Werte/
rowids, Index-/Trigger-/View-/Countererhalt, echte Fremdschlüssel und verifizierte
PRAGMA-Wiederherstellung. Final DAO1/1 (1,33s), App gebaut; vorher GUI/Import grün,
keinen frischen3/3-Nachweis erfinden. Scope/Grenzen DAO_RELATIONS_RUNTIME.md und
vollständiger jüngster Handoffabschnitt;456/384/72 gleich. Fortsetzung: grafische
Beziehungsverwaltung (echte RI/Cascades vs explizites DontEnforce), dann weitere
offene Anforderungen autonom weiterprogrammieren, keine Zwischen-Finals/Anrufe.

Grafische Beziehungen (echte RI/Cascades vs explizites DontEnforce) samt bestätigter
Entfernung und korrektem gespeicherten Fehlerstatus implementiert. Reales modales
Create/Delete/physischer FK-/Werte-Nachweis Mainwindow/DAO2/2 (5,33s), App gebaut.
DAO_RELATIONS_RUNTIME.md/README;456/384/72 unverändert. Fortsetzung: native 1:1-
Relationen über dbRelationUnique und tatsächlichen eindeutigen Kindindex.

Native explizite1:1-Beziehungen inkl. echter Unique-Kindprüfung, FK/Cascades oder
DontEnforce3, Indexabhängigkeitsschutz und GUI umgesetzt. Mainwindow/DAO2/2
(5,64s), App gebaut. DAO_RELATIONS_RUNTIME.md;456/384/72 gleich. Fortsetzung:
Required/DefaultValue/Text.Size vorhandener nativer Felder atomar ändern, Schema-
und Bestandswertschutz beibehalten; anschließend weiter ohne Zwischen-Final.

Vorhandene native Feldproperties Required/DefaultValue/Text.Size jetzt über DAO
und grafischen atomaren Batch-Save änderbar. Geprüfter Schema-/Bestandswert-/
Countererhalt, Fehlerrollback und Cursorinvalidierung bis Requery; Original-/
Counter-/unbekannte Schemaänderungen bleiben begrenzt. Unabhängige unbenannte
Fremdschlüssel blockieren den Feldumbau nicht; betroffene unbekannte Schlüssel
bleiben geschützt. Bestehendes GUI/DAO2/2 (5,92s), App gebaut. Details/Grenzen
DAO_TABLEDEF_FIELDS.md.456/384/72 unverändert. Weitere offene Anforderungen
fortlaufend implementieren; kein Zwischenabschluss oder Abschlussanruf.

Native Nicht-Counter-Felder sind jetzt per DAO Field.Name und grafisch atomar
umbenennbar, einschließlich Index-/Beziehungsmetadaten, echten FK-Verweisen und
Digestprüfung der Werte/rowids. Access-Abfrageartefakte/VBA/Formularbindungen
werden nicht refaktoriert. Echtes GUI/DAO2/2 (5,66s), App gebaut. Umfang/Grenzen
DAO_TABLEDEF_FIELDS.md.456/384/72 unverändert. Weiter native Feldvalidierung
und weitere offene Anforderungen implementieren, keine Zwischen-Finals.

Native AllowZeroLength für Text/Memo ist bei DAO-Anlage und vorhandenen Feldern
sowie in Neuer-Tabelle-/Feld-Dialogen implementiert, echte SQL-Prüfung/3315,
Null separat Required. Physische DAO-/Clonefelder lesen explizite Required-/
AllowZeroLength-Metadaten schreibgeschützt. Atomarer Umbau/Fehlerrücknahme;
Original-/unbewiesene Eigenschaften nicht erfunden. GUI/DAO2/2 (5,99s), App
gebaut; Details DAO_TABLEDEF_FIELDS.md.456/384/72 unverändert. Nächste Arbeit:
Primärindizes vorhandener nativer Tabellen wirklich anlegen/entfernen, samt
Daten-/Counter-/Abhängigkeitsschutz und grafischer Bedienung; dann weiter.

Primärindizes vorhandener nativer Tabellen sind jetzt per DAO und grafisch
wirklich anleg-/entfernbar: echter PK/NOT NULL, atomarer Daten-/Counter-/
Regelerhalt, Abhängigkeits-/Cursor-/Transaktionsschutz. GUI1/1 (4,30s), finales
DAO1/1 (1,64s), App gebaut; keinen frischen2/2-/Gesamtlauf behaupten. Details
DAO_INDEXES_RUNTIME.md.456/384/72 unverändert. Als Nächstes Required-Indizes
vorhandener nativer Tabellen mit echten Nullregeln auch auf bisher nullable
Feldern sowie korrekter Regelentfernung umsetzen; danach fortlaufend weiter.

Required-Indizes vorhandener nativer Tabellen erzwingen jetzt echte NOT-NULL-
Regeln auch auf bisher nullable Feldern. Vorhandene Nulls verweigern atomar3314;
Löschen entfernt nur implizite Indexregeln, andere Pflichtregeln bleiben erhalten.
Entfernte physische Indizes werden nicht versehentlich wiederhergestellt. GUI/
DAO2/2 (5,73s), App gebaut; DAO_INDEXES_RUNTIME.md.456/384/72 unverändert.
Nächste Arbeit: sichere physische DAO-Zeilenidentität bei normalen Nutzerfeldern
namens rowid/_rowid_/oid; danach andere offene Anforderungen ohne Zwischenfinal.

Physische DAO-Recordsets verwenden jetzt einen nachgewiesenen unverschatteten
rowid/_rowid_/oid-Alias für Lesen, Update/RETURNING und Löschen; normale gleich
benannte Nutzerfelder können keine falschen Zeilen mehr adressieren. Familie/
Clone/Requery übernehmen den Alias. Alle drei Aliasnamen oder WITHOUT ROWID
verweigern3251. SQLite3.54.1-Alias-ALTER-Grenze wird ebenfalls ausdrücklich3251,
keine vollständige Schemaunterstützung behaupten. Final DAO1/1 (1,58s), App
gebaut, DAO_TABLEDEF_FIELDS.md.456/384/72 unverändert. Nächste größere native
Katalogaufgabe: TableDef.Name für vorhandene eigene Tabellen samt tatsächlicher
Counterdefault-/Index-/Beziehungsmetadaten-Umbindung; dann weiter ohne Zwischenfinal.

TableDef.Name und grafische Tabellenumbenennung für eigene native Tabellen
implementiert: tatsächliche Counterdefault-/Frontier-/Schritt-/Index-/
Beziehungs-/Seiteneinstellungs-Umbindung, physische FK/View/Trigger-Verweise,
Digest-/Generations-/Rollbackschutz. Gespeicherte Access-Abfragen/VBA/Formulare
werden nicht refaktoriert; weitere Grenzen DAO_TABLEDEF_FIELDS.md. GUI/DAO2/2
(5,82s), App gebaut. Originaluploads3/3 und Fixtures4/4 Hashes unverändert.
456/384/72 unverändert. Nächste Aufgabe: echte Verbindungs-Schreibfähigkeit und
readonly DAO-Fehler/Metadaten sowie nachgewiesene lokale TableDef-Metadaten;
dann weitere offene Anforderungen fortlaufend, kein Zwischenabschluss.

DAO-Schreibfähigkeit prüft jetzt echtes SQLite-readonly und query_only statt
lediglich Recordsettyp. Database/TableDef/Recordset.Updatable, readonly Counter-
Metadaten und3027-Verweigerung vor Daten-/Countermutation implementiert; Cancel
bleibt möglich. Explizit eigene lokale TableDefs haben Attributes0/Connect=""/
SourceTableName=""; keine erfundenen importierten Linkdaten oder externen Provider.
Final DAO1/1 (1,66s), App gebaut, technischer Readonly-Dateivergleich unverändert.
DAO_TABLEDEF_FIELDS.md;456/384/72 gleich. Weitere offene DAO-/Schemaanforderungen
fortlaufend implementieren; keine Zwischen-Finals/Abschlussanrufe.

Native Single-Felder6/Size4 jetzt im DAO-/Tabellen-/Feldbau verfügbar, mit echter
quantisierter REAL-Speicherung/Prüfung. Gemeinsame Single-Konvertierung für DAO,
Datenblatt und begrenzte direkte SQL-Schreibpfade/physische Append-Selects,
Literaldefaults/ältere Defaultinserts und direkte Query-Typweitergabe. Negative
Null/Unterlauf bleiben3251 statt stiller Normalisierung. Final GUI/DAO2/2
(6,12s), App gebaut; vorher Import1/1 (0,06s), kein frischer3/3-Lauf. Grenzen
DAO_TABLEDEF_FIELDS.md/SOURCE_SINGLE_PRECISION.md; insb. komplexe/joined Writes,
Arithmetik/Locale/Access-Oracle und erzwungene Single-FKs noch offen.456/384/72
unverändert. Andere offene Anforderungen fortlaufend weiterprogrammieren.

Echte Single-Fremdschlüssel/Cascades jetzt für nachgewiesene eigene SINGLE_REAL-
Felder beider Seiten, einschließlich GUI-Auswahl. Nicht nachgewiesene Legacy-/
importierte REAL-Single-Schlüssel bleiben verweigert. GUI/DAO2/2 (6,26s), App
gebaut; DAO_RELATIONS_RUNTIME.md. Nur der Titelguard des GUI-Fehlerhelfers wurde
danach eingeengt, keine Produktionsänderung nach Prüfung.456/384/72 unverändert.
Fortsetzung: native Date-/Now-/Time-Defaults und Datumsnormalisierung prüfen/
implementieren, danach weitere offene Anforderungen ohne Zwischen-Final.

Native Datumsdefaults Date()/Now()/Time() jetzt für eigene Date-Felder im DAO-,
Tabellen-/Feldentwurf, manuellen Datenblatt und direkten Tabellenformular.
Neue Puffer speichern den initialisierten Uhrwert; explizit Null überschreibt,
ungebundene Datumsdefaults werden mitgespeichert. Neue Date-Felder erzwingen
kanonische ISO-Speicherung; alte Definitionen bleiben bei Migration erhalten.
Nur begrenzte eigene Uhrfunktionen, keine beliebigen SQL-/VBA-Defaults. App
gebaut, GUI/DAO2/2 (6,24s). SOURCE_LITERAL_DEFAULTS.md/DATE_RUNTIME.md;
F014/F015/DAO-/Formular-Vollanforderungen bleiben offen,456/384/72 unverändert.
Fortsetzung: eigene kanonische Date-Beziehungsschlüssel samt GUI und tatsächlichen
FK-/Kaskadenwirkungen; danach weitere offene Anforderungen, kein Zwischen-Final.

Eigene kanonische Date-Schlüssel jetzt für tatsächliche referenzielle Integrität,
Änderungs-/Löschkaskaden und Beziehungsdialog freigegeben. Beide Seiten benötigen
nachgewiesene eigene Date-Metadaten und DATETIME-Speicherung; ältere/importierte
Date-Schlüssel bleiben3251. Bestehende Uhrdefaults und binäre Nutzwerte bleiben
bei der Migration erhalten. GUI/DAO2/2 (6,11s), App gebaut. Keine vollständige
Jet-Datums-/Locale-/Untersekundenparität;456/384/72 unverändert. Fortsetzung:
native numerische Gültigkeitsregeln im DAO/Feldentwurf mit tatsächlicher
Speicherprüfung und atomarer Bearbeitung, danach weitere offene Anforderungen.

Numerische native Field.ValidationRule im DAO/Feldentwurf mit echten CHECKs
implementiert: explizite Vergleiche mit And/Or oder Between, Null separat von
Required,3317 bei Verletzung. Atomare Regeländerung prüft vorhandene Werte;
Rollback, fehlerhafte Mehrzeilenänderung, erhaltene DAO-Puffer, Löschen der Regel,
Feld-/Tabellenrename und Byte/Integer/Long/Single/Double/Boolean geprüft.
GUI/DAO2/2 (7,08s), App gebaut; DAO_TABLEDEF_FIELDS.md. Allgemeine Expressions,
ValidationText und volle Referenzsemantik bleiben offen;456/384/72 unverändert.
Fortsetzung: eigene Validierungstexte und physische Recordset-Feldmetadaten,
danach übrige offene Anforderungen ohne Zwischenabschluss.

Native ValidationText jetzt im DAO/Feldentwurf für unterstützte numerische Regeln,
mit wirklichem3317-SQL-Fehlertext; UTF-8/Umlaute, Großschreibung und Quotes bleiben
erhalten. Leeren/Bearbeiten/Migration unterstützt; Begrenzung255 UTF-16-Einheiten,
kein NUL. Physische Recordset.Fields lesen eigene ValidationRule/ValidationText
schreibgeschützt, mit bestehenden Clone/Requery-/Generationsguards. GUI/DAO2/2
(7,32s), App gebaut. Allgemeine Ausdrücke und genaue Windows-Formular-/Fehler-
Ereignissemantik weiterhin offen.456/384/72 unverändert. Fortsetzung: tatsächliche
Feldreihenfolge/OrdinalPosition im DAO und Feldentwurf mit atomarem Werterhalt.

Native TableDef.Field.OrdinalPosition verschiebt jetzt echte physische Spalten,
auch Counter, atomar mit Feld-/Indexkatalog und GUI-Positionssteuerung. Vorhandene
Werte/Rowids, Counterfrontier, Single-/Uhrdefaults, Regeln/Texte, Indizes, FK,
Trigger und Views bleiben erhalten; physische DAO-Puffer/Clones3420 bis Requery.
Unveränderte Position ohne Schemaänderung, außerhalb vorhandener Positionen3251;
keine ursprünglichen/unklaren Schemas oder aktiven Aufrufertransaktionen umbauen.
GUI/DAO2/2 (7,15s), App gebaut. DAO_TABLEDEF_FIELDS.md;456/384/72 unverändert.
Fortsetzung: Gültigkeitsregeln für eigene kanonische Date-Felder einschließlich
Uhrgrenzen im DAO/Feldentwurf, dann übrige offene Anforderungen fortlaufend.

Gültigkeitsregeln jetzt auch für eigene kanonische Date-Felder: ISO-#Datum#-
Konstanten (ganze Sekunden) und Date()/Now()/Time() mit Vergleichen, And/Or,
Between. Tatsächliche CHECK-Durchsetzung/3317 und ValidationText, atomare
Bestandsprüfung, Rename/Migration/GUI; alte/importierte Date-Speicherung bleibt
verweigert. Regionale Konstanten, Untersekunden/allgemeine Ausdrücke offen.
GUI/DAO2/2 (7,08s), App gebaut;456/384/72 unverändert. Fortsetzung: verlustfreie
numerische Datentypänderungen im Feldentwurf mit Bestands-/Defaultprüfung;
DAO Field.Type nach Append bleibt readonly. Danach offene Anforderungen weiter.

Feldentwurf unterstützt jetzt verlustfreie numerische Typänderung für eigene
Nicht-Counter: Boolean/Byte/Integer/Long/Single/Double mit Bestands-/Default-
Konvertierungsprüfung vor Umbau. Verlust/Rundung verweigert3251, Überlauf6;
Null, Indizes, Regeln/Texte und Counterfrontier geschützt. Numerische Werte
werden bei beabsichtigtem Speicherklassenwechsel bitgenau verglichen; andere
Werte/Rowids bleiben typgenau. Relationkeys vorher lösen; DAO Field.Type nach
Append weiter383. GUI/DAO2/2 (6,94s), App gebaut; DAO_TABLEDEF_FIELDS.md.
456/384/72 unverändert. Fortsetzung: doppelte Hauptfenster-Kompilierung von App
und GUI-Ziel über gemeinsame statische Desktopimplementierung beseitigen, dann
weitere offene Entwurfs-/Access-Anforderungen ohne Zwischenabschluss.

App und Hauptfensterintegration teilen jetzt dieselbe statische Desktopbibliothek
(MainWindow/Inspector/BridgeImporter und Qt-Metaobjekte), ohne zusätzliche Runtime-
DSO. Hauptfenster wird im gemeinsamen Build nur einmal kompiliert. GUI1/1 (5,25s),
GUI-freier Kern1/1 (1,93s), App mit SQLite3.54.1/JOIN-/Null-/Counter-Backendcheck
bestanden; kein frischer Vollsuite-/Distributionslauf.456/384/72 unverändert.
Fortsetzung: ISO-#Datum#-Defaults und führendes= bei begrenzten Literaldefaults
im Tabellen-/Feldentwurf und DAO; danach weitere offene Anforderungen fortlaufend.

Native Datum-Standardwerte akzeptieren jetzt eindeutige ISO-#Datum#-Literale
(optional Uhrzeit mit ganzen Sekunden) und führendes=. Begrenzte allgemeine
Literaldefaults akzeptieren ebenfalls= und äußeren Leerraum; Textinhalt bleibt
unverändert. DAO-Anlage/Bearbeitung, Tabellen-/Feldentwurf und Datumsregeln nutzen
gemeinsamen Date-Parser; ungültige Tage/regionale Formen verweigern3251. GUI/DAO2/2
(7,98s), App gebaut; erster Compilefehler war fehlender Prüfcode-Include und ist
behoben. Kein Importparser-/Windowslexik-/Vollsuiteclaim;456/384/72 unverändert.
Fortsetzung: Table-Recordsets/Index/Seek für belegte native numerische Indizes,
mit tatsächlicher Indexreihenfolge und klaren unbewiesenen Grenzen.

Native physische Table-Recordsets1 jetzt explizit/dbOpenTable sowie als Default
für Main-Tabellen/TableDefs; SQL-/QueryDef-Dynasets2/Snapshots4 bleiben erhalten.
Geprüfte native nichtnullable numerische Indizes liefern tatsächliche Reihenfolge,
Index/Seek mit kompletten Schlüsseln, ASC-Bereichen, DESC-Gleichheit, NoMatch/
Bookmarks, benannten Argumenten und kontrollierten Clonegrenzen. NativeSnapshot-
Suche, kein B-tree-/Live-Jet-Claim; Date/Text/Null/Partial/DescendingRange und
unabhängiger Clone-Indexwechsel noch offen. DAO_RECORDSET_INDEX_SEEK.md mit
aktuellen Microsoft-API-Referenzen, kein2003-Oracle. Core/GUI/DAO/originalesKassa-
Ereignis4/4 (122,64s; davon Originalimport113,62s), App gebaut. Originaluploads3/3,
Fixtures4/4 Hashes unverändert. Keine erneute breite Originalprüfung ohne neuen
betroffenen Umfang.456/384/72 gleich. Fortsetzung: eigene kanonische Date-Index-
Schlüssel und typed Seek, danach offene Anforderungen ohne Zwischenabschluss.

Index/Seek jetzt auch für eigene kanonische nichtnullable Date-Felder und gemischte
Date-/Zahlenschlüssel. Gebundene Date-Variants und VBA-#Datum# bleiben typisiert;
String-Datekeys verweigern13. Datebereiche vergleichen kanonische ganze Sekunden
entsprechend tatsächlicher Indexreihenfolge. DAO1/1 (2,20s), App gebaut; keine
Wiederholung des unverändert betroffenen Originalimports. Voller Jet-/OLE-/Locale-
Oracle bleibt offen,456/384/72 gleich. Fortsetzung: dokumentierte zusammengesetzte
Teilschlüssel-Suche mit fehlenden hinteren Schlüsseln als Null für die bereits
strukturell nichtnullen unterstützten Indizes; danach weitere offene Anforderungen.

Seek unterstützt jetzt ausgelassene hintere Schlüssel zusammengesetzter
nichtnuller eigener Indizes: Suffix wird gemäß dokumentierter Access-Suche als
Null behandelt; Gleichheit NoMatch, aufsteigende Präfixbereiche passend. Numerik/
Date, explizit Missing am Suffix, Bookmark-Rettung geprüft. Explizite Nullwerte,
Lücken und absteigende Bereiche bleiben3251. DAO1/1 (2,25s), App gebaut;
456/384/72 gleich. Fortsetzung: eigene Feld-Caption/Description im Entwurf mit
wirklichen Tabellenköpfen und Formular-/Berichtslabels, dann übrige Anforderungen.

Eigene Feld-Caption/Description jetzt tatsächlich im Projekt, Entwurf (auchCounter),
Datenblattkopf/Hinweisen und kopierten Formular-/Berichtslabels. Physische Namen/
Werte/Counter bleiben gleich. DAO.CreateProperty Text10 sowie begrenztes eigenes
Field.Properties Item/Append/Delete/Refresh, Value/defaultValue, echte native
Property/Properties-Typbindung und readonly Recordset-Properties implementiert.
Persistente Property-IDs verhindern Wiederbelebung nach Delete/Recreate; keine
vollständige Count-/Ordinal-/Enum-Liste erfunden. Schema-/Rollbackguards bleiben.
Details NATIVE_FIELD_DISPLAY_PROPERTIES.md. GUI1/1 (5,65s); danach fehlende
Typregistrierung/TypeOf-Registrierung korrigiert und zu gemeinsamer Registry
vereinheitlicht: Core/DAO2/2 (4,66s), App gebaut. Kein frischer3/3-/Vollsuiteclaim.
456/384/72 unverändert. Fortsetzung: weitere tatsächlich vorhandene native Feld-
Property-Metadaten über Properties statt nur direkte Member anbinden; vollständige
originale Properties-/Windowsparität bleibt offen. Kontinuierlich weiterarbeiten.

Native Field.Properties stellt jetzt zusätzlich vorhandene intrinsische Metadaten
über echte Property-Objekte bereit: Name/Type/Size/Attributes/DataUpdatable/
OrdinalPosition/Required/AllowZeroLength/DefaultValue/ValidationRule/-Text, jeweils
nur soweit direkter Getter belegt. Kontext bleibt erhalten (Counter TableDef17,
Recordset49); Value-Schreiben nutzt vorhandene atomare Setter, readonly bleibt383.
Eigene Property-Familie folgt zulässigen Größen-/Pflicht-/Namensänderungen; kein
Count/Enum/Vollkatalog erfunden. DAO1/1 (2,46s), App gebaut;456/384/72 gleich.
Fortsetzung: tatsächliche Caption/Description-Spaltenproperties aus MDBTools
lesend übernehmen; vorhandene mdb_col_get_prop-API geprüft, keine VBA-Extraktion.

Echte Jet4-Caption/Description-Spaltenproperties werden jetzt aus vorhandener
lesender mdb_col_get_prop-API im selben Metadatenlauf übernommen (keine extra
Prozesse proTabelle). Versionierter begrenzter Abschnitt, vollständige geordnete
Spaltenbindung, Format-/SHA-/Textprüfung; fehlende Komponente bleibt unbekannt.
Import speichert belegte Werte/Abwesenheit, leere Strings und eigene Lease-IDs;
Datenblatt/Form-/Berichtserzeugung und readonly DAO-Properties lesen sie.
Andere Formate/alte Projekte/Vererbungsflag/Vollkatalog nicht geraten. Echte
unabhängige Jackcess-Jet4-Referenz ohne Datensätze bestätigt Unicode/Quotes,
Description-Zeilenumbruch, leer/fehlend, MDB→SQLite/DAO/Kopf und Quellhash.
Import/DAO2/2 (2,57s), App/lesender Helfer gebaut. Java nur Prüfdaten-Erzeugung,
keine Runtime-Abhängigkeit. SOURCE_FIELD_DISPLAY_PROPERTIES.md;456/384/72 gleich.
Fortsetzung: weitere belegbare Eigenschaften und allgemeine Entwurfsanforderungen
fortlaufend umsetzen; keine Zwischenabschlüsse oder Fertiganrufe.

Jet4 Required/AllowZeroLength werden jetzt im vorhandenen lesenden Metadatenlauf
als ausdrücklich vorhandene boolesche Properties übernommen. Tatsächliche
MDBTools-Repräsentation yes/no gegen unabhängig wieder geöffneten technischen
Jet4-Behälter geprüft; fehlende Properties bleiben unbekannt. Arbeitskopien
setzen NOT NULL und Text-/Memo-Leerstring-CHECK tatsächlich durch; DAO liest
Originalwerte readonly. GUI verwendet ausdrückliches nativeOwned statt aus
bekannten Getterwerten auf Bearbeitbarkeit zu schließen; eine Metadatenabfrage
pro Feldliste, importierte Entwurfsänderungen deaktiviert. Keine Originalmutation.
Technische Importprüfung1/1 (0,16s), anschließend bestehende DAO/GUI/echte
Kassa-Tabellenkopie3/3 (95,41s; davon Quellkopie87,28s); App/Helfer gebaut.
SOURCE_FIELD_DISPLAY_PROPERTIES.md. Keine neue Testregistrierung oder F-ID;
456/384/72 bleiben unverändert. Fortsetzung: eigene DAO-QueryDefs dauerhaft
umbenennen, Katalog-/View-/Generationseigenschaften mit berücksichtigen.

Eigene persistente DAO-QueryDef.Name lässt sich jetzt atomar umbenennen, auch bei
reiner Groß-/Kleinschreibungsänderung. Gesamte gespeicherte Definition bleibt
bytegleich; SQL, Parameterbindungen und aufrufereigene Field-/Parameterreferenzen
bleiben erhalten. Katalogcache folgt dem neuen Namen, verwaiste alte Einträge
geben nach Neuanlage nicht versehentlich die umbenannte Abfrage zurück. Andere
alte Namenshandles werden über Namespace-Generation3420 ungültig. Vorhandener
View-Installer läuft im selben Savepoint; Fehler stellt SQL und fünf Registrywerte
zurück. Kollision/ungültiger Name/readonly/aktive äußere Transaktion verweigert;
importierte/temporäre Namen und automatische Referenzkorrektur bleiben offen.
DAO1/1 (2,52s), App gebaut; keine Quellkopie-/Gesamtsuite unnötig wiederholt.
DAO_QUERYDEFS_CATALOG.md; F055/F121 offen,456/384/72 unverändert. Weiterarbeiten.

Datenbankfenster bietet jetzt Ansicht→Abfrage umbenennen für eigene Abfragen.
Offene Abfragedatenblätter vorher schließen; importierte Abfragen bleiben
unverändert/geschützt. Live-QueryDef wird über Namensdialog gehalten, damit
zwischenzeitliches Löschen/Neuanlegen nicht auf ein anderes Objekt umleitet.
Explorer/Selektion folgen tatsächlichem Namen. Tabellen-/Abfrageumbenennung
nutzen gemeinsamen Core-Helfer für vorhandene Druckseiteneinstellungen:
Schema geprüft, Layout unverändert verschoben, Zielkollision bricht gesamten
Savepoint ab. Kein QtGui im DAO-Backend. Bestehende DAO/GUI2/2 (7,84s) prüfen
auch tatsächliches Öffnen/Ausführen, Seitenbindung und atomare Teilfehler-Rücknahme.
App gebaut; langsame Originalkopie nicht ohne Importänderung wiederholt.
DAO_QUERYDEFS_CATALOG.md; F055/F080/F121 offen,456/384/72 unverändert. Weiterarbeiten.

DoCmd.Rename führt jetzt über tatsächliche eigene TableDef-/QueryDef.Name-Setter
Tabellen-/Abfrageumbenennung aus, einschließlich neu geordneter benannter
Argumente. ObjectType und OldName derzeit ausdrücklich erforderlich; Missing
bleibt auch bei Namen/Kommalücken Missing statt versehentlich acTable(Empty).
Null94, unbewiesener Kontext/Typ3251; geschlossene/beschäftigte/modale Projekte
und passende offene Dokumente werden geschützt. Aufrufer über skalare Auswertung
und Dispatch abgesichert. Reale VBA-Integrationsprüfung bestätigt beide Objekttypen,
resultierendes SQL, Seitenbindung und Fehlerfälle. DAO/GUI2/2 (7,77s), App gebaut.
Große DAO-/GUI-Prüfprogramme verwenden unter GNU/Clang Release jetzt -O1, um
Wiederübersetzungen zu verkürzen; Produkt-/Desktop-/Datenbibliotheken bleiben
nach tatsächlich geprüften flags.make bei -O3. Keine neue Testregistrierung.
DOCMD_RENAME_RUNTIME.md; F052/F055/F080/F121 offen,456/384/72 unverändert.
Fortsetzung: tatsächliches Entfernen eigener Datenbankobjekte samt abgeleiteten
Ansichten und Seitenmetadaten prüfen/umsetzen; kontinuierlich weiterarbeiten.

Eigene Tabellen/Abfragen können jetzt über Ansicht→Ausgewähltes Objekt löschen
(mit Nein als Vorgabe und über Dialog gehaltenem DAO-Handle) oder ausdrückliches
DoCmd.DeleteObject entfernt werden. Benannte Argumente/Missing/Null und passende
offene Dokumente geschützt; importierte Objekte im Fensterpfad deaktiviert.
QueryDefs.Delete entfernt Definition, vorhandene Seitenmetadaten und verwaltete
Views im selben Savepoint; Fehler stellt fünf Registrywerte wieder her, eigene
Workspace-Rücknahme bleibt erhalten. Native DROP/TableDefs.Delete räumt ebenso
Seitenmetadaten auf, damit Neuanlage keine Altzustände erbt. Unbekannte Katalogform
bricht gesamte Mutation ab; Roh-SQL-Rename mit Seitenzustand verweigert ohne
Rebinding. Keine automatische Löschung/Umschreibung abhängiger gespeicherter
Bindings; Relationenschutz bleibt. Bestehende DAO/GUI2/2 (8,06s), App gebaut:
echter VBA-Aufruf, Offen-/Missing-/Null-Schutz, GUI Nein/Ja, Views/Seiten/Objekte
wirklich entfernt, Namenswiederverwendung, Teilfehlerrücknahme, andere Daten gleich.
DELETE_OBJECT_RUNTIME.md; F052/F055/F080/F121 offen,456/384/72 unverändert.
Fortsetzung: vorhandene echte Access-Eval-Beobachtungen für allgemeine typisierte
Ausdrücke nutzen; VBA-/Eval-Kontext unterscheiden, keine pauschale Oracle-Parität.

Native Decimal-Variants (VarType14/CDec) implementiert: geprüfter96-Bit-Koeffizient,
Skala0..28, exakte Dezimalarithmetik/-Vergleiche, Vorzeichen/Abs/Fix/Int/Round,
Ganzzahl- und Currencykonvertierung ohne Double-Zwischenrunde; Variant-Arrays,
einmalige Skalarauflösung, homogene Berichtssummen/-Extrema/-Selektoren sowie
verlustfreie ungebundene Steuerelemente. Boost.Multiprecision nur Header>=1.74;
Buildabhängigkeit README/Workflow ergänzt, keine neue Shared-Library/Java.
Normales VBA, berechnete Controls und tatsächliches Application/Host.Eval bleiben
getrennte Kontexte. Eval übernimmt beobachtete Long-/Integer-/Decimal-Literale,
Typen, Präzedenz/Nz und Fehlerresultate. Acht Fehlerausdrücke über exaktenSHA256-
Fallabgleich mit sämtlichen112 technischen Collector-Ausdrücken wiedergewonnen;
keine Eingaben erfunden. CDec in beobachtetem Eval-Kontext verweigert2439,
Switch/Error-Variant6; gewöhnliche VBA-CDec-Konvertierung bleibt verfügbar.
Bereinigte echte Access11.0/DAO3.6/de-AT-Beobachtungen:104Werte plus8zugeordnete
Fehler, ohne private Quellmetadaten/Pfade/VBA. Bestehender DAO-Prüflauf bestätigt
56Zahlen-/Typbeobachtungen und8Fehler;48andere Werte bleiben ausdrücklich offen.
Core/DAO/GUI3/3 (9,49s), App/Backend3.54.1/JOIN/NULL/Counter geprüft; GUI-freie
Konfiguration erfolgreich, gelaufener Core-Binärpfad lautldd ohneQt/Boost-DSO.
Kein neuer CTest-Punkt/Katalogpunkt, keine neue CI-/Matrix-/Vollparitätsbehauptung.
NATIVE_DECIMAL_RUNTIME.md; exakte dbDecimal-Felder/Jet-SQL-Speicherung, vollständige
Locale/IEEE-/Automation-Identität/AST weiteroffen,456/384/72 unverändert.
Fortsetzung: die48weiteren echten Ausdrucksbeobachtungen und allgemeine
Sprach-/Datenbankanforderungen fortlaufend umsetzen; keine Zwischenabschlüsse.

Allgemeine Formatfunktionen erweitert: kombinierte Datum-/Zeitmasken, Monats-/
Wochentagsnamen, AM/PM, Literale/Datum-Escapes, feste/optionale Dezimalstellen,
Gruppierung/Prozent sowie FormatNumber/FormatCurrency mit benannten/ausgelassenen
Optionen und TriState. AktuelleQLocale liefert Zahlenkultur/regionale Erkennung;
Currency-Affixabstände werden als Legacy-Leerzeichen normalisiert, numerische
Gruppierungszeichen bleiben erhalten. MathematischeZahlenformatierung nutzt
Decimal, SQLiteINTEGER wird ausdrücklich ohne Double-Umweg gelesen (2^53+1 geprüft).
NativeSQL und Desktop teilen Formatter; reineCore-Hosts bekommen keinenQt-Zwang.
VBA.CDec unterstützt jetzt benanntesExpression-Argument. Eval-FormatNull liefert
beobachteten Leerstring, andere bestehendeVariant-/SQL-Nullregeln bleiben getrennt.
Gesamte vorhandene Ausdruckssuite abgeglichen:56Zahlen-/Typbeobachtungen+8Fehler
imDAO-Ziel sowie48weitereString-/Datums-/Locale-Beobachtungen über öffentliche
Application.Eval-Schnittstelle des tatsächlichenForm-/App-Objektmodells unterde-AT.
Insgesamt104beobachteteWerte+8Fehler; Doublefinanzwerte weitermitrelativerToleranz,
KalenderwertemitunabhängigerQt-Tages-/Sekundenberechnung. KeinepauschaleAST-/
OriginalVBA-/Locale-/Gesamtparität. Core/DAO/GUI/Formatter4/4 (9,49s), Appgebaut.
NATIVE_FORMAT_RUNTIME.md beschreibt Grenzen: komplexe/skalierende/wissenschaftliche/
Stringmasken, volleWindowsRegionaleinstellungen, negativeCurrencyprofile,
Originalfeld-Formatproperties und exakteNativeDecimal-/Currency-Datenbankfelder
weiteroffen. F014/F041/F068/F071/F118/F164;456/384/72 unverändert.
Kontinuierlich an allgemeinenDatenbankanforderungen weiterarbeiten, keinZwischenfinal.

Präzisionslücke in Str/Str$ geschlossen: Decimal-/Currency-Skalare werden einmal
aufgelöst und ohne Double-Zwischentext ausgegeben, einschließlich maximalem
96-Bit-Decimalwert und kleinster28erNachkommastelle. Currency-Konvertierung aus
Double wertet nun53-Bit-Mantisse/Exponent über ganzzahlige625/2er-Skalierung aus;
keine Abhängigkeit von plattformspezifischerlong-double-Präzision. Geprüft sind
exakte Halbwerte, direkt benachbarte IEEE-Werte, Vorzeichen, Subnormale sowie
signed64-Grenze/Nichtfinite/Überlauf. Die nachgewiesene REAL-Speichergrenze bleibt
unverändert; nativevolleCurrency-Datenbankfelder werden damit nicht vorgetäuscht.
Core/DAO/GUI/CurrencySQL/-Felder5/5 (10,03s), Appgebaut; bestehende112Ausdrucksfälle
weiterimDAO-/GUI-Prüfverbund. KeineOriginalkopie ohnebetroffeneImportänderung
wiederholt. Core-Prüfharness jetztwieDAO/GUI unterGNU/ClangRelease -O1; Runtime
lautflags.make weiter-O3. Keine neueTestregistrierung/F-ID;456/384/72 gleich.
Fortsetzung: Currency-Parameterwerte von begrenzterSQL-REAL-Kodierung trennen und
an allgemeinenexaktenDatenbanktypen fortlaufendweiterarbeiten.

Currency-Parameterwerte sind jetzt von der begrenzten SQL-REAL-Kodierung getrennt.
Value/defaultValue behält exakteCurrency-Werte einschließlich beider signed64-
Extrema undNull; ungültige Zuweisung ersetztvorigenWert nicht. ErstBind/Execute
prüft weiterdiebelegteREAL-Grenze undverweigertunsupportedWerte3251 ohnesieaus
demParameterzulöschen. Skalarstandardwerte werden einmalaufgelöst, danachwird
Generation/Lebensdauererneutgeprüft: einSQL-ändernderRückruf kannnichtdurchalte
ParameterinsneueSQLschreiben. VorhandeneParameter-/DAO2/2 (2,40s), Appgebaut;
Eingabe/Lesen/Überlauf-/SQL-Verweigerung/Null/Rückruf-/CloseundechteDatensätzegeprüft.
CURRENCY_PARAMETERS_RUNTIME.md. KeinexakterDatenbank-Currency-Speicherbehauptet;
456/384/72gleich. Fortsetzung: sichere reineCurrency-Parameterprojektionenohne
REAL-Umweg undallgemeineexakteDatenbanktypenweiterumsetzen.

Reine tabellenfreie Currency/Money-Parameterprojektionen liefern jetzt beide
signed64-Grenzwerte ohne REAL-Umweg: ganzzahlige Zehntausendstel mit eindeutigem
privatem Ergebnisdeskriptor. DAO behält Currency/Type5/Size8 und Null; Datenblatt
und Zwischenablage zeigen den exakten Betrag. Explizite Ausgabenamen verhindern
Verwechslungen mit SQLite-Parameternamen, auch bei Alias=Parametername.
SQL-Ausdrücke/FROM/WHERE/mehrdeutige Spaltennamen bleiben an der bisherigen
REAL-Grenze; veränderte Pläne sowie textuelle Filter/DAO-Find auf Tick-Ergebnissen
werden vor falschen Einheitenvergleichen verweigert. Bestehende Parameter-/DAO-
Prüfziele gemeinsam 2/2 bestanden (2,69s), Anwendung gebaut. Keine neue CTest-
Registrierung, kein neuer Katalogpunkt, keine native Currency-Felderstellung
oder allgemeine exakte Currency-SQL-Parität behauptet. 456/384/72 unverändert.
Details: CURRENCY_PARAMETERS_RUNTIME.md. Fortlaufend allgemeine native
Datenbank-/VBA-Anforderungen umsetzen; kein freiwilliger Zwischenabschluss.


Declared Currency For/Next counters now use exact signed64 ticks for the start,
frozen end/step and checked increment. Positive/negative four-decimal steps,
full-range boundaries, zero iterations, explicit zero step with Exit For,
body-written counters and the final post-loop value retain Currency kind.
Increment overflow raises6 before replacing the prior counter. Existing structured
loop stacks, Exit For and error recovery remain shared; no Double conversion is
used for this counter family. The existing execution budget still bounds zero-
step loops. Core/DAO checked together2/2 (3.99s), application built. Actual Access
2003 loop-oracle coverage, mixed Variant counter promotion and complete numeric
For/Next semantics remain open; no full parity assertion. Existing F041/F043/F045/
F105 only, fixed456/384/72 unchanged. Continue implementation without a final.

Native Zahlenmasken unterstützen jetzt bis zu drei Abschnitte für positive/
negative Zahlen und Zahlenwert0, leere Ersatzabschnitte, zitierte/escaped
Präfixe/Suffixe, negative Klammern und Prozent an beiden Seiten. Tausender-
Skalierung durch abschließende Kommata prüft exakte Rückrechnung und verweigert
Präzisionsverlust3251, statt still zu runden. .00 bleibt unterstützt. Gemeinsamer
Code für VBA/SQL/Formausdrücke; Masken-/SQL-Fälle und vorhandene GUI-Ausdrucks-
referenzen gemeinsam2/2 (6,01s), App gebaut. Keine zusätzlichen CTest-Ziele.
Wissenschaftliche/bedingte/Farb-/vierteNull-/interleavte Masken und vollständige
Windows-Formatgleichheit bleiben offen. Neue technische Prüfungen sind kein
zusätzlicher Windows-Access-Nachweis. NATIVE_FORMAT_RUNTIME.md; bestehende
F041/F068/F071/F118/F164,456/384/72 unverändert. Direkt weiterentwickeln.

FormatPercent ist jetzt als echte native VBA-/SQL-Funktion mit allen fünf
benannten/ausgelassenen Standardargumenten verfügbar. Die gemeinsame geprüfte
Maskenstrecke berücksichtigt Nachkommastellen, führendeNull, negative Klammern,
Gruppierung und lokale Prozentzeichen. Die benannte Format-Maske Percent liefert
zwei Nachkommastellen ohne Gruppierung. SQL-Integer bleiben vor der Skalierung
exakte Decimal-Werte; 9007199254740993 wird ohne Double-Verlust formatiert.
Null94, ungültige Optionen5 und Arity450 sind geprüft. Bestehender Formatter und
GUI gemeinsam2/2 (5,83s), App gebaut; native VBA-Positional-/Named-/Missing-Fälle
und vorhandene48 GUI-Referenzwerte enthalten. Keine neue Windows-Beobachtung
oder zusätzliche CTest-Registrierung. Decimal96/28-Grenze und vollständige
regionale Prozentprofile bleiben offen;456/384/72 unverändert.
NATIVE_FORMAT_RUNTIME.md; kontinuierlich allgemeine Datenbankfunktionen umsetzen.

Eigene native Tabellenfelder besitzen jetzt eine persistente DAO-Format-Text-
Eigenschaft mit derselben Detached-/Append-/Value-/Delete-/Identitätslogik wie
Caption/Description. Feldentwurf kann Masken setzen/löschen. Datenblatt zeigt
Formatwerte; SQL und EditRole bleiben unverändert. Zwischenablage behält den
bestehenden rohen typisierten Text. Druck/PDF nutzt eigene geprüfte formatierte
Werte: ungültige Masken zeigen #Fehler im Raster und brechen die Ausgabe vor
Erzeugung ab. Native Datum-/Zahlenmasken sind über echte DAO-/Designer-/Raster-
und PDF-Textextraktion geprüft, einschließlich beibehaltener SQL-Werte, Property-
Identität/Delete/Recreate und Rename/Schemaumbau. Fehlendes Include im GUI-
Prüfziel gezielt behoben; danach vorhandene DAO/GUI gemeinsam2/2 (8,00s), App
gebaut. Originale Format-Eigenschaften werden noch nicht als bekannt behandelt
(3251 statt erfundener Abwesenheit); Query-Lineage und Formular-/Berichtserbung
bleiben offen. NATIVE_FIELD_DISPLAY_PROPERTIES.md;456/384/72 unverändert.
Direkt als Nächstes proveniente Jet4-Formatübernahme mit technischer MDB prüfen.

Jet4-Format-Eigenschaften werden jetzt durch die bestehende schreibgeschützte
MDBTools-Grenze gelesen und mit eigener expliziter Format-Provenienz importiert.
DAO liest vorhandene Werte readonly383 bzw. meldet nachgewiesene Abwesenheit3270;
alte Metadaten ohne neue Format-Provenienz bleiben unbekannt3251. Datenblatt-
Formatierung verwendet diese tatsächlich gelesenen Masken, ohne Bearbeitungs-
oder SQL-Werte zu verändern. Die neue reine Entwicklungs-Javaquelle erzeugt eine
unabhängig wieder geöffnete technische Jet4-MDB mit null Zeilen/ohne VBA; nur die
importierte SQLite-Testkopie erhält Zahlen-/Datumsprüfwerte. Bestehende Import-/
DAO-Ziele gemeinsam2/2 (2,70s), zusätzlich frühere Caption-/Required-/AllowZeroLength-
Referenzdateien im selben Importziel; Herkunft, Typen, readonly, bekannte Abwesenheit,
Legacy unbekannt, ungültige Provenienz, Anzeige/Edit/Copy und unveränderter MDB-Hash
geprüft. App/Helper gebaut; vier Fixture-Hashes und drei Upload-Originalhashes OK.
Keine MDB oder Java-Abhängigkeit ins Produkt/Git übernommen. Keine Windows-
Render-/Vollparität;456/384/72 unverändert. SOURCE_FIELD_DISPLAY_PROPERTIES.md.
Kontinuierlich weitere allgemeine Datenbank-/VBA-Anforderungen umsetzen.

FormatDateTime unterstützt jetzt native VBA-/SQL-Aufrufe mit Date und optionalem
NamedFormat sowie vbGeneralDate/vbLongDate/vbShortDate/vbLongTime/vbShortTime.
Default und explizit ausgelassener NamedFormat wählen General Date; reine Zeit
und reines Datum werden getrennt dargestellt, sonst kombiniert. Lange Zeit nutzt
das tatsächliche QLocale-Profil (einschließlich US-PM), ohne Zeitzonentext; kurze
Zeit bleibt24-stündig. No-host-Syntaxprüfung darf geschachteltes CDate nicht
vorzeitig als Null konvertieren: gezielt behoben. CLDR-Schmal-/NBSP vor AM/PM
werden im nativen Legacy-Profil als Leerzeichen ausgegeben. Vorhandene Formatter/
GUI gemeinsam2/2 (5,57s), App gebaut; echte VBA Named/Missing/qualified-Fälle,
SQL1..2Argumente, Null94/ungültigeOption5/Arity450/ungültigesDatum13 und deutsche/
US-Zeitdarstellung geprüft. Keine zusätzliche CTest-Registrierung oder Windows-
Beobachtung. Volle regionale Datums-/Subsekunden-/Windows-Gleichheit weiter offen.
NATIVE_FORMAT_RUNTIME.md;456/384/72 unverändert. Kontinuierlich weiterentwickeln.

TextBox.Format wird jetzt aus SaveAsText übernommen und für berechnete Formular-
Textfelder sowie Berichtstextfelder tatsächlich gerendert. Getrennte typisierte
Werte bleiben maßgeblich: Anzeige-Fehler ändern Value nicht; Currency/Datum und
Long-Untertypen bleiben auch beim Löschen einer Maske erhalten. Native Berichte
erben proveniente direkte Feldmasken aus dem Berichtsentwurf und drucken sie
wirklich ins PDF. Bestehende gespeicherte Überschreibungen bleiben unabhängig.
Runtime-Formatlesen/-setzen/-löschen unterstützt diese Kontexte; ungültige Masken
lassen vorherige Eigenschaft stehen. Ausgabe verweigert Formatfehler vor falschem
PDF-Erfolg. Bearbeitbare Formularfelder behalten Masken nur mit explizitem3251-
Hinweis; deren Fokus-/Eingabepipeline ist nicht vorgetäuscht. CONTROL_FORMAT_RUNTIME.md.
GUI-Ziel bestanden (5,70s); anschließend Layoutziel einzeln bestanden (11,56s).
Der breitere Layoutverbund enthielt noch die alte Currency-For-Verweigerung:
gezielt durch echten Schleifenlauf mit zwei SQL-Schreibwirkungen und exaktem
Currency-Endzähler ersetzt. Keine neue CTest-Registrierung, keine frische gesamte
Suite/Windows-Gleichheit. App gebaut;456/384/72 unverändert. Ohne Zwischenfinal
weiter an den allgemeinen offenen Anforderungen programmieren.

General Number und Scientific sind jetzt native benannte Zahlenformate. General
Number behält Currency-/Decimal-Text exakt, trimmt überflüssige Nachkommastellen
und gruppiert nicht. Double/Single verwenden native15/7-signifikante Stellen.
Wissenschaftliche Einzelmasken 0.00E+00 bzw. e/E und +/- unterstützen optionale
Nachkommastellen, Exponentenpolsterung und exakte halbe Decimal-Rundung samt
Mantissenübertrag. Auch endliche Double-Werte außerhalb Decimal96 (1e100 und
Subnormale) können so formatiert werden. Vorhandene Formatter/GUI gemeinsam2/2
(5,79s), App gebaut; Max-Currency ohne Double-Umweg, größter Decimal, sehr kleine/
große Double, Rundung, Locale und echte SQL-Integer geprüft. Bisherige gezielte
Ausgabefehlerfälle verwenden weiterhin unsupported Farbmasken statt nun gültiger
wissenschaftlicher Masken. Keine neue Windows-Beobachtung/CTest-ID; kombinierte
wissenschaftliche Abschnitts-/Literal-/Engineering-Masken, vollständige regionale
Coercion und Access-Gleichheit bleiben offen. NATIVE_FORMAT_RUNTIME.md;456/384/72
unverändert. Ohne Zwischenabschluss direkt weiterimplementieren.

Textmasken unterstützen jetzt @ (mit Leerstellen), & (ohne Leerstellen), ! für
Linksfüllung, native Groß-/Kleinschreibung mit >/<, zitierte/escaped Literale und
einen zweiten Abschnitt für leere Strings/Empty. Null-Konventionen bleiben
getrennt. Tabellenanzeige und geerbte Berichtsformate verwenden diese Strecke;
Value und rohe Zwischenablage behalten den gespeicherten Originaltext. Native
Formatter/GUI gemeinsam2/2 (5,72s), App gebaut; echte SQL-Fälle, Füllrichtungen,
überzählige Zeichen, Leerwerte, Tabellen-/Berichtsmaske bei unveränderten Daten,
Unicode und vor unzulässigem Surrogat-Split verweigerte Ausgabe geprüft. NUL/
überlange Ausgabe sowie längenändernde Case-Konvertierung mit Platzhaltern werden
explizit verweigert; vollständige Windows-/Unicode-/Maskengleichheit bleibt offen.
Keine neue Windows-Beobachtung oder CTest-ID. NATIVE_FORMAT_RUNTIME.md;
456/384/72 unverändert. Nächste Umsetzung: Fokus-/Eingabestrecke formatierter
bearbeitbarer Formulartextfelder, ohne Anzeigen als Daten zurückzuparsen.

Bearbeitbare Formular-TextBoxen nutzen jetzt echte Formatdarstellung außerhalb
des Fokus und kanonische, ungerundete Eingabetexte im Fokus. Maßgeblich bleiben
typisierte Werte; nur tatsächliche Bearbeitung verwirft die Wertkopie. Formate
werden nicht als Daten zurückgeparst. Reine Darstellungswechsel erzeugen keine
Change-/Dirty-/SQL-Wirkungen. Doppelte Bindungen erhalten unabhängig ihre Masken;
verschiedene Werte mit identischer gerundeter Anzeige werden dennoch gespeichert.
Native Formularerstellung übernimmt proveniente Feldmasken. Bestehende Validierung,
Value/Text-Puffer, Undo und AfterUpdate bleiben wirksam; Maskenänderungen während
laufender Eingabe/BeforeUpdate verweigern3251. Skalarstandardwerte werden einmal
aufgelöst und Owner/Beschreibbarkeit erneut geprüft, geschlossener Owner2450.
Vorhandene GUI/Layout gemeinsam2/2 (17,22s), App gebaut. Tatsächliche Qt-Eingabe,
negative Abbruchvalidierung ohne Datenverlust, Undo, präzise SQL-Änderungen,
Text-Großdarstellung ohne Wertänderung und vollständige Uhrzeit hinter Datumsmaske
geprüft. Zunächst fehlende explizite Textgröße im technischen DAO-Aufbau gezielt
korrigiert; keine zusätzliche CTest-Registrierung. Regionale Eingabemasken,
beliebiges reentrantes Fokus-/IME-/Windows-Verhalten und vollständiger Access-
Formulardruck bleiben offen. CONTROL_FORMAT_RUNTIME.md;456/384/72 unverändert.
Ohne Zwischenfinal weitere allgemeine offene Funktionen umsetzen.

Formatierte native Suche ist in Tabellen-, Abfrage- und Formularansichten sowie
DoCmd.FindRecord integriert. Wörtliche String-Suche verwendet tatsächliche
Delegate-/TextBox-Darstellung, getrennten Wiederholungszustand und bei doppelten
Bindungen die gewählte Steuerelementmaske. Formatierte Suche über alle sichtbaren
gebundenen Textfelder ist möglich; Passwortdarstellung liefert keine Klartext-
Treffer. Rohwertsuche bleibt getrennt; formatierte Ersetzung verweigert3251.
Zell-/Zeitgrenzen, Null-/Binärausschluss und vorhandene Validierung gelten.
Vorhandene GUI/Layout gemeinsam2/2 (16,95s), App gebaut; tatsächliche Navigation,
Datumsmasken, doppelte Bindungen, Allfeldsuche und Passwortanzeige geprüft.
Ein veralteter geschlossener Dialogzeiger im Prüffall wurde gezielt korrigiert.
Keine neue CTest-ID oder Windows-Beobachtung; volle Jet-/Windows-Suchgleichheit
bleibt offen. FORMATTED_FIND_RUNTIME.md;456/384/72 unverändert.
Ohne Zwischenabschluss nächste Umsetzung: Bereichseinfügen mit neuen Datensätzen.

Native Tabellen-Bereichseinfügung kann jetzt über das Ende vorhandener Datensätze
hinaus neue Entwurfszeilen anlegen. Vorher werden noch nicht geladene vorhandene
Zeilen abgeholt und alle Zellen/Typen vorgeprüft; AutoWert-/Binärziele bleiben
geschützt. Neue Zeilen nutzen vorhandene Defaults, Datensatz-Undo, Verwerfen und
transaktionales Speichern mit echter Schlüsselerfassung. SQL-Constraintfehler
rollen alle SQL-Änderungen zurück; Entwürfe bleiben korrigierbar. Bestehender
Currency-/Feld-/Clipboard-Prüflauf1/1 (0,44s), App gebaut; Typfehler ohne Teiländerung,
zusätzliche Zeilen, Undo/Verwerfen, SQL-Defaults, erzeugte Schlüssel und Unique-
Rollback geprüft. Technische Schutzprüfung auf explizites AUTOINCREMENT korrigiert;
manuelle INTEGER-Primärschlüssel sind keine behaupteten AutoWerte. Kein neuer
CTest. Noch kein Einfügen in leere Tabellen/gesonderter Tabellen-Anfügebefehl;
weitere F340-Anforderungen bleiben offen. CLIPBOARD_APPEND_RUNTIME.md;
456/384/72 unverändert. Direkt weiterimplementieren.

Leere native Tabellen akzeptieren jetzt Texteingabe aus der Zwischenablage ab der
ersten sichtbaren beschreibbaren Spalte. Ein ausdrücklicher Kontextmenübefehl
und RunCommand acCmdPasteAppend(38) legen neue Datensatzentwürfe nach dem vollständig
geladenen Tabellenende an. Normales Einfügen behält die aktuelle Zelle als Ziel;
AutoWerte werden nicht überschrieben. Typisierte Formular-MIME wird ausdrücklich
verweigert, statt ihren Textkopf als Datensatz einzufügen. Vorhandene GUI/Feld-
Prüfungen gemeinsam2/2 (6,35s), App gebaut: leere Tabelle, Anfügen ohne Überschreiben,
SQL-Persistenz und tatsächlicher Hauptfenster-Dispatcher geprüft. F340 bleibt wegen
weiterer MIME-/Regional-/Undo-/Entwurfsszenarien offen. CLIPBOARD_APPEND_RUNTIME.md;
456/384/72 unverändert. Direkt weitere Clipboard-Teilanforderungen umsetzen.

Native Tabellen-Zwischenablage überträgt jetzt zusätzlich zum TSV eine typisierte,
begrenzte Rechteck-MIME. NULL, leerer Text, leere Binärwerte, Binärbytes, volle
SQLite-64-Bit-Ganzzahlen und vorhandene Currency-/Datumsdarstellung bleiben getrennt.
Binärdaten werden nur über validierte native Nutzdaten in Binärspalten geschrieben;
Text-/AutoWert-/Readonly-/verdeckte Zielgrenzen bleiben geschützt. Fehlerhafte MIME
wird nicht still als TSV interpretiert. Tabellen ohne SQL-Defaults benötigen keine
unnötige DAO-Rekonstruktion ihrer BIGINT-/BLOB-Felder. Für einen unveränderten
ungespeicherten Einfügebereich bestehen jetzt gemeinsame Strg+Z/Y-, Bearbeiten-Menü-
und Undo/Redo(292/199)-Wege; gespeicherte Date-Defaults werden beim Redo übernommen.
Spätere Fremdänderung, Requery oder erfolgreicher Save verwirft den Bereichsverlauf.
Vorhandene GUI/Feld/DAO gemeinsam3/3 (8,59s), App gebaut. Echte SQL-/Qt-Nachweise für
Bytes inklusive leerem BLOB, NULL/leer, große 64-Bit-Ganzzahlen, datumhaltige
Bereiche, Anfügen, leere Tabelle, Undo/Redo und Verlaufverfall; ungültige MIME,
AutoWert und verdeckte Ziele verweigert. Zwei veraltete Binär-Copy-Negativannahmen
angepasst; Qt-interner Schreibschutz während geprüfter Binärübernahme begrenzt
freigegeben, außerhalb weiterhin wirksam. Keine neue CTest-ID oder Windows-
Beobachtung. Kein VBA-Binary-Variant, keine allgemeine SQL-Rücknahme; Einfügen über
fremde Entwürfe, volle Regional-/Access-/MIME-Gleichheit bleiben offen. F340/F168/
F016 unverändert; TYPED_GRID_CLIPBOARD.md;456/384/72 unverändert. Direkt weiter.

Kreuztabellen mit TRANSFORM Count(Feld)/Count(*) und festen Textüberschriften in
PIVOT ... IN(...) laufen jetzt über denselben nativen Compiler, direkte DAO-
Recordsets, parametrisierte QueryDefs und gespeicherte Abfragen. Zählungen lassen
NULL-Feldwerte aus, Count(*) zählt Zeilen; fehlende Kategorien ergeben0. Reihenfolge,
Aliases und Fallkollisionsschutz der Ausgabespalten werden erhalten, Count-Spalten
sind provenient INTEGER4/DAO Long, QueryDef.Type ist dbQCrosstab16. Simple gruppierte
Zeilenfelder, einzelner Quellname mit optionalem Alias sowie WHERE/HAVING/ORDER
bleiben begrenzt; numerische Aggregate, dynamische/numerische Überschriften und
allgemeine Join-/Ausdrucks-Kreuztabellen verweigern3251. Vorhandene Query-Resultat/
Schema-/DAO-Prüfungen gemeinsam3/3 (2,44s), App gebaut. Zusätzlich vorhandene SQL-
Dokument-/DAO-Prüfungen2/2 (2,37s); keine neue CTest-ID. Klammer-/Nichtklammernamen
werden semantisch gleich verglichen. Kein Windows-Ergebnis-/Kollationsorakel und
keine volle Kreuztabellenparität behauptet. Siehe COUNT_CROSSTAB_RUNTIME.md.
F068/F067 und456/384/72 bleiben unverändert. Vier Repository-Originale und drei
Upload-Originale aktuell per SHA unverändert. Ohne Zwischenabschluss fortsetzen.

Feste Count-Kreuztabellen unterstützen jetzt zusätzlich vorzeichenbehaftete
Long-Überschriften und reine Year(Feld)/Month(Feld)/Day(Feld)-Pivot-Ausdrücke.
Integer-Spaltennamen werden kanonisch gebildet;01/1 und Text-/Zahl-Namenskonflikte
werden vor Ausführung verweigert. Nichtintegrale, außerhalb Long liegende und
beliebige andere Pivot-Ausdrücke bleiben3251. Vorhandene Query-Resultat-/Schema-
Prüfungen gemeinsam2/2 (0,12s), App gebaut: echte Datumszeilen mit Uhrzeit,
Schalttag und NULL, vorgegebene Monats-/Jahresreihenfolge, fehlende Kategorie,
Signum und Negativfälle geprüft. Keine neue CTest-ID/Windows-Beobachtung;
COUNT_CROSSTAB_RUNTIME.md; F068/F067 und456/384/72 unverändert. Direkt weiter.

Vollständige native Tabellenzeilen können jetzt über ausdrückliches Anfügen mit
neuen AutoWerten übernommen werden. Die Grid-MIME enthält physische Feldnamen;
Schemaform, Namen und Reihenfolge müssen vollständig passen. Nur provenient
generierte Spalten werden ausgelassen, manuelle Primärschlüssel bleiben echte
Werte mit ihren SQL-Constraints. Die neue Schlüsselvergabe erfolgt ausschließlich
beim normalen transaktionalen Save. Die Regel bleibt bei Bereichs-Undo/Redo erhalten.
Vorhandene GUI/Feld-Prüfungen gemeinsam2/2 (7,05s), App gebaut: vollständiger echter
Datensatz mit Binärbytes und großer Ganzzahl angefügt, Undo/Redo, tatsächlicher
neuer SQL-Schlüssel und verweigerte fremde Feldform geprüft. Positionsbereiche und
ältere Grid-MIME bleiben getrennt; anders geordnete/partielle Namenszuordnung und
Formular-MIME-Interop bleiben offen. TYPED_GRID_CLIPBOARD.md; F340/F168 und456/384/72
unverändert. Direkt weitere Bearbeitungsanforderungen umsetzen.

Saubere native Tabellen-Datenblätter unterstützen jetzt Ausschneiden sichtbarer
zusammenhängender Zellbereiche per Strg+X/Umschalt+Entf, Bearbeiten-/Kontextmenü und
RunCommand Cut189. Die typisierte Zwischenablage erhält die geprüften Quellwerte;
danach werden nur Cache-Zellen auf SQL NULL gesetzt. SQL bleibt bis zum normalen
transaktionalen Save unverändert. Dieselbe begrenzte Bereichs-Undo/Redo-Strecke
stellt die Werte wieder her. AutoWert-/Readonly-/verdeckte Ziele und fremde
Entwürfe bleiben geschützt. Vorhandene GUI/Feld-Prüfungen gemeinsam2/2 (6,91s),
App gebaut: echtes Menü/Dispatcher/Qt-Tastatur, Binär-Cut ohne frühe SQL-Änderung,
Undo/Redo, persistentes NULL statt leerem Text und verweigerter AutoWert-Cut bei
unveränderter Zwischenablage geprüft. Keine neue CTest-ID oder Windows-Beobachtung.
Dies ist Zellbereichsleeren mit Undo, keine Datensatzlöschung oder volle Access-
Cut-/Ereignisparität. TYPED_GRID_CLIPBOARD.md; F340/F168 und456/384/72 unverändert.
Ohne Zwischenabschluss weiterimplementieren.

Bereichseinfügen kann jetzt vorhandene ungespeicherte Zelländerungen physischer
Tabellenzeilen erhalten. Vorprüfung erfasst rohe Werte und geänderte Feldpositionen;
Undo stellt diese vorherigen Entwürfe wieder her, Redo wendet nur den neuen Bereich
an. Fehlgeschlagene Bereichsänderungen nutzen dieselbe Cache-Rücknahme, ohne SQL-
Schreibvorgang oder automatisches Speichern. Vorhandene GUI/Feld-Prüfungen gemeinsam
2/2 (6,28s), App gebaut: aufeinanderfolgende echte VBA-/UI-Pastes, ungültige Eingabe
bei erhaltenem vorherigem Entwurf/Undo, Currencywerte, separater Textentwurf und
Binäränderung mit leerem BLOB vor/nach Undo/Redo geprüft. Veraltete globale Dirty-
Negativannahmen der bestehenden Prüfungen angepasst. Neue/gelöschte Datensatzentwürfe
und mehr als1M geladene Zellen bleiben3251; Cut benötigt weiter einen sauberen Cache.
Keine neue CTest-ID/Windows-Beobachtung. TYPED_GRID_CLIPBOARD.md; F340/F168 und
456/384/72 unverändert. Ohne Zwischenabschluss weiterimplementieren.

Bereichseinfügen erhält jetzt auch bestehende neue Datensatzentwürfe und vorgemerkte
Löschungen. Cache-Rücknahme rekonstruiert eingefügte Zeilen an ihren Positionen,
übernimmt ihre tatsächlich gesetzten Felder, lässt SQL-Defaults ungesetzt und
stellt Löschmarker wieder her. Direkte Ziele in gelöschten Zeilen bleiben2448.
Bereichs-Undo/Redo stellt nach dieser Rekonstruktion eine gültige aktuelle Zelle
wieder her. Vorhandene GUI/Feld-Prüfungen gemeinsam2/2 (6,19s), App gebaut: früherer
Insert plus nachfolgende Bereichseingabe, Qt-Ctrl+Z/Y bei erhaltenem Zellziel,
tatsächliche spätere SQL-Neuanlage einschließlich Default, vorgemerkte Löschung,
verweigerter gelöschter Zielbereich und gemeinsamer SQL-Save geprüft. Erste Strecke
2/2 (6,36s), danach Zellziel-/Defaultnachweis gezielt ergänzt; keine neue CTest-ID.
Kein vorzeitiger SQL-Write, keine neue Windows-Beobachtung; beliebige externe
Index-/Reentranz-/Konkurrenz-/Access-Ereignisgleichheit bleibt offen.
TYPED_GRID_CLIPBOARD.md; F340/F168 und456/384/72 unverändert. Direkt weiter.

Native einfache numerische For-Zähler unterstützen jetzt Byte und Single zusätzlich
zu Integer/Long/Double. Grenzen/Schrittweite werden einmal ausgewertet und nicht
auf den Zählertyp eingeengt; Start/Next werden regulär Let-zugewiesen, danach wird
der tatsächlich zugewiesene (ggf. Single-gerundete) Zähler verglichen. Integer-/Long-
Bruchgrenzen, Byte-Rückwärtsschritte und Step0 mit Exit For funktionieren; Overflow6
ersetzt den letzten angenommenen Byte-/Single-Wert nicht. Deklariertes Date ist
laut MS-VBAL kein numerischer Zählertyp und bleibt vor SQL-Präfix verweigert3251.
Quelle: veröffentlichte MS-VBAL v20260818 §5.4.2.3/5.5.1.2.1, PDF-SHA dokumentiert;
keine neue Windows-/Access2003-Beobachtung. Vorhandene Core/DAO/Layout gemeinsam3/3
(17,41s), App gebaut: Rundung vor Anfangsvergleich, Endtypen Single/Double,
Bruchgrenzen, eingefrorene Grenzen, Body-Zähleränderung, Zero-Step, Overflow und
tatsächliche SQL-Schleifenaufrufe. Bestehende globale Bruchgrenzen-/Single-/Byte-
Negativannahmen gezielt in reale positive Fälle geändert; keine neue CTest-ID.
Volle gemischte Variant-/Fixedpoint-/Fehler-/Compilerparität bleibt offen.
TYPED_FOR_RUNTIME.md; F045/F041/F122 und456/384/72 unverändert. Direkt weiter.

Tabellennavigation kann jetzt von einer sauberen aktuellen Zeile wechseln, ohne
unbeteiligte Cache-Entwürfe zu speichern oder zu verwerfen. Native Pfeiltasten und
GoToRecord/RunCommand nutzen denselben Pfad. Ein einzelner befüllter neuer Datensatz
kann vor dem Verlassen über den normalen transaktionalen Save mit echter
Schlüsselerfassung gespeichert werden; vollständig leere neue Entwürfe bleiben
begrenzt, neue Zeilen werden bereits für Record-Undo als pending erkannt.
Der Shared-Find-Guard erlaubt ebenfalls eine saubere aktuelle Zeile bei anderen
Entwürfen. Andere Guard-Aufrufer behalten ihre bisherigen Insert-Grenzen.
Vorhandene GUI/Feld gemeinsam2/2 (6,20s), App gebaut: tatsächliche Qt-/VBA-Navigation,
erhaltene mehrere Entwürfe, unverändertes SQL für unbeteiligte Änderungen,
verweigerter leerer Insert, echter einzelner Insert beim Wechsel mit neuem
SQL-Schlüssel; Mehrfach-Save-/Sort-/Requery-Sperren weiterhin geprüft. Keine neue
CTest-ID oder Windows-Beobachtung; TABLE_NAVIGATION_DRAFTS.md. Aktueller gelöschter
Datensatz, nötige Teil-Saves mehrerer Entwürfe, allgemeine Addition-Row-/Windows-
Ereignisse bleiben offen. F368/F168/F080 und456/384/72 unverändert. Direkt weiter.

Native Find-Dialoge von Tabellen, Abfragen und Formularen unterstützen jetzt einen
expliziten Platzhaltermodus, getrennt von wörtlicher Suche. DoCmd.FindRecord
String-Muster aktivieren ihn bei * ? # oder [. Syntax/Matching teilen den bestehenden
VBA-Like-Parser und UTF16-DP-Matcher, einschließlich Zeichenklassen/Negation,
Whole/Prefix/Substring-Grenzen, Case-Modus und gespeichertem FindNext-Zustand.
Formatierte Suche behält reale Delegate-/Kontrollmaske und Passwortdarstellung;
typisierte Gleichheit bleibt getrennt, Musterersetzung verweigert3251. Musterfehler93
und absolute Patternlänge werden vor Navigation/Save geprüft; pro Kandidat gelten
1M Input×Pattern-Units,16384 Pattern-Units und Scan-Zell-/Zeitgrenzen. Bestehende
Navigation kann beim Matchscan den angenommenen aktuellen Datensatz speichern.
Vorhandene Core/GUI/Layout gemeinsam3/3 (19,13s), App gebaut: tatsächliche VBA-/Qt-
Muster, Ziffern/Klassen/Negation, Text-/Datumsmasken, Literal-Umschaltung, Query,
Fehler-/Größengrenzen und unveränderte SQL-Daten geprüft. Zunächst kollidierenden
lokalen Testvariablennamen korrigiert; veraltete globale Wildcard-Negativannahme
gezielt ersetzt. Keine neue CTest-ID/Windows-Beobachtung. Nicht-ASCII-Textranges/
Surrogatvergleiche und volle Jet-/Locale-/Allfeld-/Error/Event-Gleichheit bleiben
offen. WILDCARD_FIND_RUNTIME.md; F373/F071/F130 und456/384/72 unverändert. Direkt weiter.

Unformatierte DoCmd.FindRecord-Suche unterstützt jetzt OnlyCurrentField:=acAll0
über die jeweils passenden unterstützten Feldtypen. Text sucht Textfelder,
native Number/Currency/Date/Boolean verwenden ihre vorhandene geprüfte Gleichheit;
typisierte Werte benötigen weiter Whole-Field-Match. In direkten Tabellenformularen
ist der gerade fokussierte Steuerelementtyp im Allfeldmodus nicht mehr maßgeblich.
Tabellen/Readonly-Query-Resultate verwenden den bestehenden Allspalten-Scanner,
Formulare ihren begrenzten Datensatzbestand; formatierte Allfeld-Strings bleiben
getrennt. Vorhandene GUI-Prüfung1/1 (5,90s), App gebaut: tatsächliche VBA-/Qt-
Textsuche vom fokussierten Zahlenfeld, numerische Suche aus Textfokus, vollständiges
Datum mit Uhrzeit sowie Tabelle mit gefundenem Zelltyp/Spaltenwechsel geprüft.
Keine neuen CTest-/Windows-Beobachtungen; keine unbekannten Typen aus Zelltext
erraten. Volle Cross-Type-Coercion, Allfeld-/Spalten-/Fokus-/Ereignis-/Locale-
Gleichheit bleibt offen. ALL_FIELD_FIND_RUNTIME.md; F373/F071/F130/F050 und456/384/72
unverändert. Ohne Zwischenabschluss weitere allgemeine Funktionen umsetzen.

Wörtliches Tabellen-Ersetzen unterstützt jetzt Whole-/Teiltext-/Präfixmodus.
Current ersetzt das erste Vorkommen im Entwurf, Alle ersetzt alle nichtüberlappenden
Vorkommen jedes passenden Feldes bzw. einen Präfix/Whole-Wert. UI übergibt den
gewählten Vergleichsmodus tatsächlich; formatierte/wildcard/typisierte Rückersetzung
bleibt verweigert. Leerer Ersetzungstext ist echte Leerzeichenfolge, kein SQL NULL.
Batch-Ausgaben werden vor Cache-Änderung vollständig vorbereitet,16Mi UTF16-Units/
1M Zeilen/5s-Scan begrenzt und gemeinsam transaktional geschrieben; SQL-Fehler rollen
alle Änderungen zurück und hinterlassen keine Batch-Entwürfe. Vorhandene GUI/Feld
2/2 (6,26s), App gebaut: echte Dialoge, Case-insensitive Mehrfachvorkommen,
Prefix/current-erstes-Vorkommen, getrennte SQL-Save-Grenze, leere Textwerte,
pro Zeile verschiedene Ausgaben, vollständiger Check-Constraint-Rollback und NUL-
Verweigerung geprüft. Keine neue CTest-ID oder Windows-Beobachtung.
LITERAL_REPLACE_RUNTIME.md; volle Match-Offset-/Highlight-/Wildcard-/Locale-/Event-
Gleichheit bleibt offen; F373/F340/F168 und456/384/72 unverändert. Direkt weiter.

Explizites natives Datensatz-Anfügen ordnet Grid-MIME-Felder jetzt nach eindeutigen
Namen zu, auch bei anderer Reihenfolge und Teilmengen. Ausgelassene Felder bleiben
ungesetzt und erhalten SQL-Defaults; ausdrücklich geliefertes NULL überschreibt
Defaults. Nur proveniente AutoWerte werden ausgelassen. Name-/Formprüfung und
Expandierungsgrenze erfolgen vor Cacheänderung; die Supplied-/Default-Regel bleibt
bei Undo/Redo erhalten. Normales Paste/alte namenlose MIME bleiben positionsbezogen.
Vorhandene GUI/Feld gemeinsam2/2 (6,21s), App gebaut; zusätzlich Default/NULL-Nachweis
gezielt1/1 (0,48s): anders geordnete Zieltabelle mit zusätzlichem Defaultfeld,
AutoWert an anderer Position und echter Folge-ID11/12, Binärbytes, voller SQLite-
Integer, vollständiges Datum, Teilrecord mit ausgelassenen NULL-Feldern, explizites
NULL bei existierendem Textdefault, Undo/Redo und verweigerte doppelte Namen.
Keine neue CTest-ID oder Windows-Beobachtung. TYPED_GRID_CLIPBOARD.md; externe/
Formular-MIME-/Access-Ereignisgleichheit bleibt offen. F340/F168/F173 und456/384/72
unverändert. Direkt weiterimplementieren.


Readonly-Abfragedatenblätter veröffentlichen jetzt neben TSV das vorhandene typisierte
Grid-MIME mit Feldnamen. Tatsächliche rohe Querywerte erhalten Binärbytes, NULL,
Leertext/leeres BLOB und volle SQLite-Ganzzahlen; TSV verwendet Hex-Binärtext und
exakte Ganzzahldarstellung. Proveniente CURRENCY_TICKS-Ergebnisse erhalten ein
getrenntes Currency-Tag mit kanonischen 64-Bit-Ticks, keine Umdeutung als Ganzzahl.
Paste wandelt Currency vor Cacheänderung in exakten Text oder verlustfrei begrenzte
bestehende Currency/Double-Speicherung um; andere Zieltypen werden verweigert.
Keine physische CURRENCY_TICKS-Unterstützung oder externe Access-MIME-Parität.
Vorhandene GUI/Feld-Prüfungen gemeinsam2/2 (6,42s), App gebaut: tatsächliches
Query→anders geordnetes Tabellenziel, SQL-Default, volle Integer/Binär-/NULL-
Unterscheidung, beide Currency-64-Bit-Grenzen nach Text, kleine Currency nach REAL,
Undo/Redo/Save und vollständige Verweigerung großer Currency→REAL ohne Entwurf/
SQL-Änderung. Fehlendes QMimeData-Include beim Build gezielt korrigiert. Keine neue
CTest-ID/Windows-Beobachtung. TYPED_GRID_CLIPBOARD.md; F340/F168/F173/F014 und
456/384/72 unverändert. Direkt weiterimplementieren.


Bestehende aktuelle Tabellenzeilen können bei Navigation/Find jetzt einzeln
transaktional gespeichert werden, während andere Update-/Insert-/Delete-Entwürfe
im Qt-Cache unverändert bleiben. Bei mehreren Entwürfen sind geänderte/NULL-
Primärschlüssel und neue/gelöschte aktuelle Zeilen weiter verweigert; alleinige
populierte Inserts behalten den bisherigen Save-/Key-Pfad. Nach Commit aktualisiert
Qt selectRow nur die gespeicherte Zeile; Cache-Undo wird an der SQL-Grenze ungültig.
Vor Commit verschwundene/trigger-umgeschlüsselte Schlüssel führen zum Rollback.
GUI/Feld gemeinsam2/2 (6,21s), App gebaut: echte VBA-Navigation/Find, andere Entwürfe
und unveränderte Fremdzeilen, gemischte Update/Insert/Delete-Caches, Constraint-
Rollback und Trigger-Rekey-Rollback geprüft. Veraltete Find-Negativannahme und
Testversuch, ein geschütztes AutoWertfeld zu bearbeiten, gezielt korrigiert;
manuelles Text-PK separat geprüft. Uploads3/3/Fixtures4/4 hashgleich. Keine neue
CTest-ID/Windows-Beobachtung. TABLE_NAVIGATION_DRAFTS.md; vollständige aktuelle
Insert/Delete-/Fokus-/Ereignis-/Addition-Row-Parität bleibt offen. F368/F168/F080 und
456/384/72 unverändert. Direkt weiter.


„Neuer Datensatz“/acNewRec spart jetzt das pauschale Speichern fremder Entwürfe:
sauberer aktueller Datensatz darf bei weiteren Entwürfen eine neue Zeile öffnen;
bestehender geänderter aktueller Datensatz nutzt den gezielten Save-Pfad. Ein
alleiniger gefüllter Insert wird regulär gespeichert, bevor die nächste neue Zeile
entsteht. Leere aktuelle Inserts oder selektive Inserts mit weiteren Entwürfen
bleiben verweigert. Bestehende GUI-Prüfung1/1 (6,42s), App gebaut: tatsächliche
RunCommand/VBA-Aufrufe, SQL-Abgrenzung, anderer Entwurf, neuer Zeilenfokus und Escape
ohne fremden Datenverlust. Keine neue CTest-ID/Windows-Beobachtung. F368/F173/F168
und456/384/72 unverändert. Direkt weiter.


F088 hat jetzt einen ersten ausführbaren nativen Makroumfang: echter +Makro-
Aktionseditor (anlegen/entfernen/umsortieren/speichern/wiederöffnen), Explorergruppe,
DoCmd.RunMacro mit benannten Parametern und begrenzter Wiederholung. Acht Aktionen
verwenden vorhandene DoCmd-Pfade; StopMacro beendet einschließlich Wiederholungen.
Gesamte Struktur/Arity/Typen/Attribute werden vor erster Aktion geprüft, unbekannte
Bedingungen verweigert. Aktionen behalten ihre jeweiligen Transaktionsgrenzen;
späterer Laufzeitfehler rollt frühere erfolgreiche Aktionen nicht zurück. Diagnostik
bewahrt Fehlernummer/-metadaten mit Makroname/Aktionsnummer. AllMacros bietet den
bewachten vorhandenen Projektkatalog; SysCmd unterstützt nativen acMacro-Zustand.
GUI/Layout gemeinsam2/2 (17,16s), App gebaut: echter Editor/Reopen, gespeicherte
Definition, tatsächlicher benannter VBA-Aufruf mit dreifacher SQL-Wirkung,
StopMacro, präventiv verweigerte spätere unbekannte/bedingte/falsch typisierte
Aktionen, fehlendes Makro2485, bestehende SQL-Fehlernummer3078 mit Aktion2 und
bereits erfolgreich geschriebenen früheren Zeilen, Wiederholungsgrenze und
Katalog Count/Item/Enumeration/Name/IsLoaded geprüft. Anfänglich zu generische
SQL-Fehlernummer3146 im Test auf den tatsächlichen bestehenden3078-Pfad korrigiert.
Keine neue CTest-ID/Windows-Beobachtung. NATIVE_MACROS_RUNTIME.md beschreibt eigene
Format-/Budgetgrenzen; kein Access-Makroimport. Gruppen, Bedingungen/Ausdrücke,
RepeatExpression, Eventbindung, volle Aktionssprache und Oracle bleiben offen.
F088/F050/F057/F148 und456/384/72 unverändert. Direkt weiter programmieren.


Gespeicherte native Makronamen funktionieren jetzt auch als Ereignisbindungen:
vorhandene generische Steuerelementpfade (u.a. Click/DblClick/Change), Kontroll-
BeforeUpdate/AfterUpdate sowie Form-/Berichtsereignisse und Timer verwenden einen
geteilten Dispatcher für [Event Procedure], =Funktion() und benannte Makros.
OnActivate verschiebt Current bei jeder vorhandenen Bindung wie bisher bei VBA.
Makrofehler werden als echte Ereignisergebnisse/Statusdiagnose gespeichert;
fehlendes Kontroll-BeforeUpdate-Makro verhindert Save mit unveränderten SQL-Werten.
GUI/Layout gemeinsam2/2 (17,33s), App gebaut: tatsächliches Load/Current, Qt-Klick,
Doppelklick, Texttastatur/Change, Before/AfterUpdate mit SQL-Persistenz, tatsächliches
Timer-Timeoutsignal, Close sowie fehlender Click/BeforeUpdate-Name2485 nachgewiesen.
Timerprüfung prüft Signaldelivery, keinen neuen Wandzeit-/Access-Timing-Oracle.
Makronamen mit führendem =/[ sind zur Abgrenzung von Ausdrücken/Embedded-Makros
verweigert. Keine neue CTest-ID/Windows-Beobachtung. CancelEvent, Tastatur-
Makroargumente, Gruppen/Bedingungen/volle Ereignisgleichheit bleiben offen.
NATIVE_MACROS_RUNTIME.md; F088/F050/F057/F148 und456/384/72 unverändert. Direkt weiter.


Native Makros/DoCmd unterstützen jetzt CancelEvent ohne Argumente in unterstützten
abbrechbaren Bindungen (Open/Unload/NoData/Dirty/BeforeInsert/BeforeUpdate/DblClick).
Der verschachtelt restaurierte Ereigniszustand wird in VbaExecutionResult.cancelled
übernommen; Folgeaktionen laufen normal weiter. Außerhalb unterstützter Ereignisse
bleibt3251, zusätzliche Argumente450. GUI/Layout2/2 (17,08s) bestätigen echte
Kontroll-BeforeUpdate-Verweigerung ohne SQL-Save, Open-/Unload-Abbruch und keinen
über den Ereignisrahmen hinaus verbleibenden Cancel-Zustand. App gebaut; die neue
Designerzuordnung liegt separat noch als WIP vor.
Zusätzlich bestehender Original-Kassa-Lauf gezielt1/1 (114,46s) bestanden. Dabei
alte Rohtextannahmen für tatsächlich formatierte Original-Form-/Berichtsfelder und
Summen korrigiert: gemeinsame Prüfhilfe vergleicht reale skalare Werte mit SQL und
Anzeige separat über den bestehenden Formatter. SQLite-Ausdrucksmetadaten sind
kein voller Access-Varianttyp-Oracle. Eigenen anfänglich falschen Objektzugriff in
der Prüfhilfe behoben; keine Produktionsformatierung zur Erzwingung alter Erwartungen
abgeschaltet. Ungültige explizite WhereCondition meldet den tatsächlichen
Datenquellenfehler vor nachfolgenden Initialisierungsfolgefehlern; die bekannten
beiden fehlende-Verkaufs-ID-Fälle liefern wieder3075/Zeile1023 statt überlagertem3146.
Das ist Fehlerpriorität nach Konstruktion, keine Zusicherung ausbleibender früherer
Handlerwirkungen. Originalverkaufsworkflow bleibt offen. Uploads3/3 und Fixtures4/4
erneut hashgleich. Keine neue CTest-ID oder Windows-Beobachtung; kein frischer
Gesamtsuite-/Plattform-/CI-Nachweis. NATIVE_MACROS_RUNTIME.md; F088/F057/F114/F129/
F148 und456/384/72 unverändert. Direkt weiter.


Native Formulare können Makros jetzt über „Ereignisse zuordnen...“ an unterstützte
Form-/Feldereignisse binden. Automatisch/Makro/Funktion/[Event Procedure]/Keine
Bindung sind echte gespeicherte Alternativen; „Keine“ unterdrückt automatische
VBA-Erkennung. Unterdialog arbeitet auf Kopien, Cancel verwirft, OK validiert,
Eltern-Save persistiert. NativeStoredForm prüft/emittiert events/field_events/
timer_interval, unbekannte/doppelte Ereignisse/Felder und ungültige Werte werden
verweigert; keine Ausführung beim Save. Makronamenauflösung und Save-Konfliktprüfung
verwenden begrenzten Qt-Case-insensitive-Katalogscan (1M Namen/5s), danach exakten
Definitionslookup; Ä/ä und mehrdeutige Namen sind tatsächlich geprüft, keine
Jet-/Locale-Gleichheit behauptet. Eigenständiges VbaValue-Include korrigiert.
GUI/Layout zusammen2/2 (18,42s), App gebaut: echte verschachtelte Dialoge,
Cancel/Reopen/OK, gespeicherte Definition/Layout, Load-SQL-Wirkung, per Designer
gebundener BeforeUpdate-Abbruch mit unverändertem SQL, explizit deaktivierte
vorhandene Form_Load-Prozedur, unbekanntes Ereignis/Feld und negativer Timer.
Keine neue CTest-ID/Windows-Beobachtung. Originallauf1/1 (114,46s) ist der zuvor
geprüfte Stand, kein neuer Gesamtsuite-/Plattformnachweis. NATIVE_FORM_EVENT_DESIGNER.md;
F080/F088/F050/F057/F130/F148 und456/384/72 unverändert. Direkt weiter.


Native Makroaktionen unterstützen jetzt optionale Bedingungsausdrücke und
verschachtelte RunMacro-Aktionen. Der echte Aktionseditor speichert/verschiebt die
Bedingung mit der Aktion. Syntax/Typ/UTF16-/4096-Unit-Grenzen werden für die ganze
Definition vor ihren Aktionen geprüft; anschließend genau eine Auswertung pro
besuchter Aktion/Iteration im vorhandenen Eval-Kontext. False/Null überspringen.
Öffentliche Standardmodulfunktionen bleiben vom CodeBehind getrennt; anfangs im
falschen Kontext angelegte technische Prüffunktion in echten Standardmodul-Artefakt
verschoben, keine privaten/öffentlichen Zugriffsgrenzen aufgeweicht. Nested RunMacro
teilt16 aktive Aufrufe/10000 besuchte Aktionen, restauriert Budgets bei Fehlern;
StopMacro beendet nur den aktuellen Aufruf samt Wiederholungen, Eltern können
weiterlaufen. Vorherige SQL-Wirkungen werden bei späterem Laufzeitfehler nicht
pauschal zurückgerollt. RepeatExpression/Gruppen/volle Macro-Expr-/Event-Parität offen.
App gebaut; Layout1/1 (11,21s) im initialen Bündel; nach korrigiertem Modulkontext
nur betroffene GUI-Prüfung wiederholt1/1 (6,43s): echte Editorbedingung,
True/False/Null, exakt dreifache öffentliche Function-SQL-Wirkung, ungültige spätere
Bedingung ohne Prefix-SQL, Nested-Wiederholung, Child-Stop/Parent-Fortsetzung,
Rekursionsschutz/anschließende Recovery sowie echte bedingte Formularvalidierung
(blocked verhindert Save, allowed wird gespeichert). Keine neue CTest-ID/Windows-
Beobachtung, kein neuer Vollsuite-/Plattform-/CI-Nachweis.
Aktuelles .deb unter /workspace/work/native-conditional-macros/package/linaccess-
0.7.0-Linux.deb gebaut; vorhandene Paketprüfung bestätigt App-/MDB-Helfer-Build-IDs,
bytegleichen SQLite3.54.1-Backend, tatsächlichen Backend-Selbstcheck und erlaubten
Payload ohne Referenzdaten/Quellen. CPack-Preinstall-Build einmal mit Parallelität4
abgeschlossen, keine zusätzliche Gesamttestsuite ausgeführt. NATIVE_MACROS_RUNTIME.md;
F088/F041/F050/F057/F148 und456/384/72 unverändert. Direkt weiter programmieren.


Eigene Standardmodule sind jetzt echte native Projektobjekte: +Modul/Texteditor,
Ctrl+S ohne Schließen, Wiederöffnen/Bearbeiten, Syntaxinventar ohne Ausführungsclaim,
AllModules-/SysCmd-Metadaten und bestätigte Löschung nur eigener Module. Namen
werden nach Save gesperrt; Umbenennen/Classes/VBE/volle Sprache bleiben offen.
Version1/vba-Definitionen liegen getrennt in _la_objects, importierte VBA-Zeilen
werden nicht geändert. Gemeinsamer Sources-Leser versorgt native Aufrufe/Eval,
öffentliche Variablen und bestehende SQL-Registry, erhält Archiv-Latest-/Shadow-
Prüfungen und Quellenrevision/ROLLBACK-Kohärenz und verweigert Kollisionen statt
Auswahl. Bounds256 Module/16MiB Quellen/1M Katalogzeilen/5s Scan, eigene Quelle1MiB.
Löschen prüft Projektdentität/-generation nach Modalinteraktion und gespeicherte
Revision. Fehlende direkte Includes und gemischte auto-Typdeklaration korrigiert.
DAO1/1 (2,55s), Layout1/1 (11,36s) im Bündel bestanden. Nach tatsächlicher
Qt-Fokussetzung für Ctrl+S und normalem LinAccess-Abfragepfad zur initialen UDF-
Registrierung nur betroffene GUI-Prüfung erneut1/1 (6,23s), App gebaut: echte
Editor-/VBA-/SQL-Aufrufe, gemeinsamer Counter/Private-Helper, SQL-Private-Sperre,
vorbereitete Quellenänderung/Savepoint-Rücknahme, native/importierte Namens- und
TEMP-Schattenfehler, Löschung/alte Prepared-Verweigerung sowie unveränderte
vorhandene Archivbytes. Keine neue CTest-ID/Windows-Beobachtung, kein neuer
Gesamtsuite-/Plattformnachweis. Paketprüfung gehört zum vorherigen64abd69-Stand,
kein frisch paketierter Modulstand behauptet. NATIVE_STANDARD_MODULES_RUNTIME.md;
F042/F050/F068/F118/F130/F155 und456/384/72 unverändert. Direkt weiter.


Application.Run für unqualifizierte öffentliche Standardmodul-Functions/Subs
implementiert: Procedure/Arg1..Arg30 mit bestehender benannter Bindung,
typisierte Rückgabe bzw. Empty, optionale Parameter und geteilter Modulzustand.
Private/fehlende Ziele2517, Projektqualifikation3251, Arity450; Fehlernummer,
Source und Help bleiben erhalten. Eigener Run-Grenzwert32 liefert28 vor der
allgemeinen Interpreter-Stapelgrenze. Ausführungstiefe schützt laufenden Kontext.
Kein DoCmd.Run, keine neue COM-Automation oder ByRef-Aufrufer-Rückschreibung.
App gebaut; bestehendes mainwindow_documents1/1 (6,48s) bestanden. Nach falscher
neuer Arity-Testannahme450 statt449 und konkret gefundenem Zusammenspiel der
Rekursionsgrenzen nur betroffene GUI-Prüfung wiederholt. Keine neue CTest-ID,
Gesamtsuite/Windows-Beobachtung/Paketprüfung behauptet. NATIVE_APPLICATION_RUN.md;
F042/F050/F118/F148 und456/384/72 unverändert. Direkt weiter mit Makro-RunCode.


Native Makroaktion RunCode implementiert, keine DoCmd.RunCode-Methode:
ein einzelner unqualifizierter Function-Aufruf mit Klammern, Rückgabewert ignoriert;
Quellsyntax/äußere Aufrufform vor erster Aktion geprüft. Geschachtelte Argumente,
doppelte Stringquotes, Datums-/Bracket-Literale im Scanner; ASCII-Prozedurnamen,
64Ki UTF16 Ausdruckslimit. Im gebundenen Formular-/Berichtsereignis CodeBehind
vor Standardmodulen, danach restaurierter Ereigniskontext; Standard-Private/
fehlende Funktionen2517, Subs/Namenskonflikte3251, originale Fehlerdaten plus
Makroaktionskontext. GUI/Layout zusammen2/2 (18,17s) ohne Korrekturlauf, App gebaut:
wirkliche Funktionen/Seiteneffekte, verschachtelte Makros, Private-CodeBehind,
Kontextrestauration, verschachtelte/quotierte Argumente, Sub-/Private-Sperren,
struktureller Abbruch ohne vorangestellten SQL-Effekt und Fehler7001/Source.
Aktuelles DEB gebaut; erste Paketprüfung hatte im aktivierten SDK dessen älteres
SQLite im LD_LIBRARY_PATH und scheiterte3251. Nur Paketprüfung mit wie bei den
gezielten Tests vorangestelltem gebündeltem SQLite wiederholt:5-Datei-Payload,
App-/Helper-BuildIDs, MDBTools1.0.1, bytegleiches SQLite3.54.1 samt tatsächlichem
JOIN/NULL-/Counter-Backendcheck bestanden. Keine Wiederholung der GUI-Prüfungen.
Originalupload-Hashes3/3 und Fixtures4/4 unverändert. Keine neue CTest-ID,
Gesamtsuite/Windows-Beobachtung. NATIVE_MACROS_RUNTIME.md; F088/F042/F050/F057/F148
und456/384/72 unverändert. Weiterprogrammieren, kein Gesamtabschluss.


Eigene gespeicherte Standardmodule können im realen Editor mit Prozedurauswahl,
Ausführen/F5 direkt laufen. Neuer NativeStandardModuleHost nutzt tatsächliches
CurrentDb/DBEngine, gemeinsamen Modul-/Staticzustand und öffentliche Modulaufrufe,
bestehende Qt-Vergleichs-/Format-/MsgBox-/InputBox-Pfade, keinen Formularattrappen-
kontext. Me/current-record und fehlende Application-Objekte bleiben abgegrenzt.
Projekt-/Generationsprüfung bei Modulentry/DB/Dialog; gespeicherte Definition
gegen gemeinsamen Katalog geprüft, ungespeicherter Text verweigert. Auswahl ist
transient und markiert Quelle nicht dirty; Edit deaktiviert Run bis Analyze/Save.
Neue gemeinsame MsgBox-Brücke korrigiert zunächst gefundenen Linkagefehler der
anonymen Hilfsfunktion; fehlenden QTextBrowser-Testinclude ergänzt. App gebaut,
bestehendes GUI1/1 (7,20s) bestanden: tatsächliches F5, DAO-Effekt und zweites Modul,
mit Formular geteilter Counter44, Rückgabewert, MsgBox-Ja6, Arity450, Me438,
GoSub-Preflight3251 ohne Präfix-SQL und Unsaved-Sperre. Keine neue CTest-ID oder
Gesamtsuite/Windows-Beobachtung; aktuelles DEB gehört vorherigem d36ca9f-Stand.
NATIVE_STANDARD_MODULES_RUNTIME.md; F042/F043/F050/F060/F118/F148 und456/384/72
unverändert. Direkt weiter mit formularfreiem Application.Run/Eval-Kontext.


Formularfreier Modul-Anwendungskontext erweitert: echte projekt-/generations-
gebundene Application für Run/Eval/DBEngine; Domänenfunktionen lesen tatsächliche
DB. Gemeinsamer Qt-freier Run-Helfer erhält in beiden Hosts Arity, öffentliche
Auflösung, typisierte Rückgaben/Fehlerdaten und32-Level-Grenze. Variant-Argumente
bewahren Objektverweise (DAO.Field); skalare Zielparameter behalten native
Coercion-Grenzen. Schwacher Projektcache hält keine Editor-/Host-Referenz und
prüft Projektgeneration/GUI-Thread; gehaltene Application bleibt nach Editorclose
verwendbar, Generationswechsel3420. Volles Application-Modell/identische Objekt-
identität zu Formularhosts/implizite Objekt-Skalarkoerzierung bleiben offen.
App gebaut; vorhandene GUI/Layout2/2 (18,74s) bestanden: dynamischer zweiter
Modulaufruf mit DAO-Schreiben, Eval/echtes DCount, DAO.Field-Objektargument,
rekursives Run28, gehaltene Application nach Editorclose, Identität über neuen
Host, Fehler7001/Source und Generationssperre. Testzugriff auf private Eval-Methode
vor dem Lauf auf öffentliche Run-Rückgabe eines Access.Application-Funktionstyps
korrigiert. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung.
NATIVE_APPLICATION_RUN.md; F042/F043/F050/F118/F148 und456/384/72 unverändert.
Direkt weiter mit formularfreier nativer Makroausführung.


Gespeicherte native Makros können über Ansicht/Ausgewähltes Makro ausführen direkt
vom Datenbankfenster laufen. Reeller Modul-/DAO-Anwendungshost mit DoCmd-Proxy,
kein verstecktes Formular. Makro-Budget/Repeat/Bedingungen/Stop/Fehlerkontext in
gemeinsamen Runner extrahiert;16 aktive/10000 besuchte Aktionen gelten auch über
Hostwechsel, VbaExecutionScope hält Stop über den gesamten Aufruf. RunSQL hat
geprüfte Savepoint-/UseTransaction-Pfade; RunCode öffentliche Functions. Tatsächliche
Tabellen/Abfragen über bestehende Dokument-APIs; eigene gespeicherte Normalformulare
und Berichtsvorschauen über bestehende unabhängige showDocument-Lebensdauer.
Aktives reales Formular erhält RunCommand/Requery; sonst tatsächliches Datenblatt.
Projektgeneration/Verbindung/GUI-Thread und Modalgrenzen bleiben geprüft; Menülauf
schützt laufende Projektoperationen. Formular-/Berichtsfilter/weitere Modi und
Archiv-only-Dokumente sind in diesem Host noch nicht implementiert. ReportView0
wird3251 abgewiesen; Vorschau benötigt2. CancelEvent ohne Ereignis3251.
App gebaut, GUI/Layout zusammen2/2 (19,15s) bestanden: tatsächliche Menüstarter-
Kette mit SQL/Function/verschachteltem StopMacro/Bedingungen, Table/Query/Form/
Report/Requery/RunCommand, OnLoad-SQL und Form-RecordCount3, Fehler in Aktion2 mit
persistiertem Vorläufereffekt, Rekursion3251 mit folgendem unabhängigen Erfolg,
SELECT-RunSQL2342, nichtgebundenes CancelEvent und echtes abgebrochenes Form_Open.
Fehlenden QStatusBar-Testinclude vor Lauf korrigiert; keine neue CTest-ID.
Aktuelles DEB mit4 parallelen Preinstall-Jobs gebaut; Payload/Berechtigungen,
App-/Helper-BuildIDs, MDBTools1.0.1 und bytegleiches SQLite3.54.1 samt tatsächlichem
JOIN/NULL-/Counter-Backendcheck bestanden. Keine Gesamtsuite/Windows-Beobachtung.
NATIVE_MACROS_RUNTIME.md; F088/F042/F050/F057/F114/F148 und456/384/72 unverändert.
Direkt weiter, keine Gesamtabschlussmeldung.


Native AutoExec beim tatsächlichen Projektöffnen implementiert: nach Schema/
QueryViews/Explorer-Initialisierung einmal pro neuem Öffnen, CI-Katalogauflösung;
nur gespeicherte native _la_objects-Makros, keine Archiv-only-Access-Makros.
Gehaltener Shift-Zustand überspringt Startmakro. Menü/Startup nutzen gemeinsamen
Projektstarter mit Generations-/Verbindungs-/Modal-/Ausführungsschutz und Refresh
auch nach Teilfehlern; neue Generation leert schwachen Modul-Application-Cache.
AutoExec-Fehler bleiben sichtbar, Projekt offen, bereits abgeschlossene Effekte
werden nicht erfunden zurückgerollt; Shift erlaubt nächsten Start ohne Wiederholung.
AllowBypassKey=False/Startup-Formular/Trust-/Access-Makroimport bleiben offen.
App gebaut; nur bestehende GUI1/1 (7,30s) bestanden. Erste Shift-Prüfung hatte
synthetische Taste ohne Modifier-Zustand; korrigiert auf echte Qt-Shift-KeyPress
mit ShiftModifier im geöffneten Dateidialog, nur betroffene GUI erneut. Reale
Dateiprojekte: SQL+RunCode+Table beim ersten/erneuten Öffnen je einmal, Shift ohne
Effekte, unveränderte Archiv-only-Makrobytes ohne Ausführung, Fehler2517 mit
persistiertem Vorläufer und erfolgreichem Shift-Neustart, gehaltene Application
nach tatsächlichem Projektwechsel3420. Keine neue CTest-ID/Windows-Beobachtung/
Gesamtsuite; Paket gehört vorherigem db892cc-Stand. NATIVE_AUTOEXEC_RUNTIME.md;
F088/F042/F050/F121/F148 und456/384/72 unverändert. Direkt weiter.


RunMacro RepeatExpression implementiert, in API und nativem Aktionseditor:
Bedingung vor jedem Versuch; Anzahl allein begrenzt, Bedingung allein bisFalse,
beide bis erster Grenze, beide ausgelassen einmal. Bool-/Zahlkonstanten über
native String-/Expression-Konvertierung, Textsyntax bis4096 UTF16 vor Effekten.
Ausgelassene Count-Position im nativen JSON ausschließlich für RunMacro als null
(Missing); Editor-Save/Reopen erhält leere Anzahl und Argument3. Ohne Count max
1000 abgeschlossene Wiederholungen; weiterer wahrer Versuch3251, abschließendes
False normal. Gemeinsame16-/10000-Aktionsbudgets und StopMacro unverändert.
Layout1/1 (11,44s) bestand im Bündel, GUI fand verlorene Fehlerquelle im alten
VbaHost::eval-Pfad. Erklärung/Source/HelpFile/HelpContext dort korrekt weitergereicht,
gezielt bestehende Kern-/GUI2/2 (9,32s) bestanden, App gebaut: reale DAO-Predicate-
Effekte3/Prüfungen4, Count+Expr, False0/TrueCount, Fehler7002 samt Wiederholungs-
kontext und separater Kern-Source/Help42-Nachweis, ungültige spätere Expr ohne
Präfixeffekt, nichtendender verschachtelter Noop an1000er-Grenze, danach UI-
Erstellung/CtrlS/Wiederöffnen und erneuter erfolgreicher Ablauf. Alte explizite
False-RepeatExpression-Verweigerung auf tatsächlichen erfolgreichen Null-Durchlauf
umgestellt. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung;
aktuelles DEB gehört db892cc. NATIVE_MACROS_RUNTIME.md; F088/F041/F051/F118/F148
und456/384/72 unverändert. Direkt weiter mit nativen Makrogruppen.


Native Makrogruppen implementiert: eigenes Formatversion2 mit geordneten groups,
1..64 eindeutig CI benannten Teilmakros, insgesamt max256 Aktionen/1MiB; Version1
bleibt kompatibel. Gruppe startet erstes Teilmakro, Gruppe.Teilmakro gezielt;
Namens-/Callvalidierung und Formular-/Feldereignisdesigner akzeptieren diese Form.
Alle Teilmakros strukturell geprüft, bevor ausgewählter Teil Effekte ausführt;
fehlerhafte inaktive Gruppe blockiert, doppelte Unicode-CI-Namen3251, fehlender
Teil2485. Bestehender Runner hält Wiederholung/Stop/übergreifende Budgets.
Realer Editor ergänzt Gruppenmodus, Name/Auswahl/Hinzufügen/Entfernen/Umordnen,
behält rohe auch ungültige Zell-Drafts beim Wechsel; Auswahl/Render markieren
Quelle nicht dirty. Mehrteilige Gruppe wird nicht still in Einzelmakro verwandelt.
Projektgeneration vor Save geprüft, aktive Zellbearbeitung vor CtrlS-Serialisierung
committed; neue Draft-Serialisierung bewahrt RepeatCount-Missing-Hole. Keine
automatische Verweisrefaktorierung oder Access-SaveAsText-/Binär-Makroimport.
App gebaut; bestehende GUI/Layout2/2 (20,10s) ohne Korrekturlauf bestanden: echte
Gruppenerstellung/CtrlS/Wechsel/Wiederöffnen/Umordnung mit geänderter Default-
Ausführung, Unicode-Teilaufruf, verschachtelter Parent/StopMacro/Repeat, reales
Form_Load, ungültige inaktive Gruppe und doppelte Labels ohne Effekte. Aktuelles
DEB gebaut/geprüft:5-Datei-Payload, Rechte, App-/Helper-BuildIDs, MDBTools1.0.1,
bytegleiches SQLite3.54.1 mit tatsächlichem JOIN/NULL-/Counter-Backendcheck.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung. NATIVE_MACROS_RUNTIME.md;
F088/F050/F057/F118/F148/F155 und456/384/72 unverändert. Direkt weiter mit
Speicher-/Löschschutz für eigene Makroobjekte.


Eigene Makroobjekte erhalten transaktionalen Speicher-/Löschschutz: neue Namen
kollidieren CI mit vorhandenen nativen/importierten Makronamen3010 statt Replace;
gespeicherte Objektnamen sind im Editor readonly. Save und Delete vergleichen
exakte alte Definition/Name unter Savepoint, staleRevision3251 statt Lost Update.
Main-Katalog physisch/ohne TEMP-Schatten geprüft, zusätzliche main/temp Trigger
verweigert; Größenlimit1MiB vor Materialisierung. Lesen/Ändern importierter
Quellarchive bleibt ausgeschlossen. Eigene Makros löschbar mit Bestätigung im
Editor/Explorer; auch unlesbare eigene Definitionen über Explorer entfernbar.
Projektgeneration nach Modalfrage und Definition nach Antwort erneut geprüft.
DoCmd.DeleteObject acMacro4 für eigene gespeicherte Objekte über bestehende
Projekt-API ergänzt. Gemeinsame laufende MacroFrames sperren Ersetzen/Löschen
auch für Gruppen und über Formular-/Modulhostgrenzen; kein globaler Jet-Lockclaim.
Layout1/1 (11,74s) im Bündel bestanden. GUI fand alte vorgelagerte Case-only-
Namensprüfung; entfernt zugunsten zentralem3010-Pfad. Nur betroffene GUI erneut
1/1 (7,41s), App gebaut: echte zweite Verbindung und Revision während bestätigter
Löschfrage, Native-/Archiv-Kollision, No/Yes, TEMP-/Trigger-Verweigerung, Löschen
ungültiger eigener Definition, aktive3251-Sperre in beiden Hosts mit folgenden
SQL-Effekten, anschließendes DeleteObject und unveränderte Archivbytes. Original-
Uploads3/3 sowie Fixtures4/4 erneut hashgleich. Keine neue CTest-ID/Gesamtsuite/
Windows-Beobachtung/Paketprüfung; DEB gehört9c80f00. NATIVE_MACRO_STORAGE.md;
F088/F050/F121/F148/F155 und456/384/72 unverändert. Direkt weiter mit allgemeinen
Makro-Projektbefehlen und getrenntem Projekt-Ausführungsschutz.


Native Makroaktionen Close/DeleteObject/Rename ergänzt über reale bestehende
Projekt-APIs: eigene geschlossene Tabellen/Abfragen umbenennen/löschen, eigene
Makroobjekte löschen, offene Datenblätter sowie reale Form-/Berichtsdokumente
schließen. Keine Save=acSaveNo-Verwechslung mit Datensatz-Undo; bestehende
validierte Close-Pfade committen Recorddrafts und geben SQL-Modelle frei.
Macro-Projektstarter erhält eigene Ausführungstiefe statt SQL-Tiefe; legitime
Objektbefehle bleiben nutzbar. Projekt-/Anwendungsschließen und neue UI-Makro-
Reentry bleiben bei aktiven Frames blockiert, auch im direkten Formularhost.
Gemeinsamer Preview-Dispatch hält NativeExecutionDepth bis Befehlsende, damit
Close/Requery/RunCommand keine laufenden Dokumentaufrufe vorzeitig entsorgen.
Wiederverwendete gespeicherte Fenster wiederholen keine historischen Initfehler;
neue fehlgeschlagene Initialisierungen bleiben echte Fehler mit neuerDoc-Bereinigung.
App gebaut; bestehende GUI/Layout2/2 (18,92s) ohne Korrekturlauf bestanden:
Menü-Makrokette mit realem Table/Query-Rename, Open/Close/Drop und Makrolöschung,
aktuellem Query-DefaultClose, Pending-Draft43 trotz acSaveNo, echtem Form_Unload/
Report_Close und anschließendem SQL. Während MsgBox ist tatsächliches Window.close
abgewiesen bei gleicher Projektgeneration/offener DB; danach Weiterlauf. Bestehende
Form mit altem Load7003 nur aktiviert, neue Init7003 abgewiesen/bereinigt. Grenzen
bleiben explizite Objekttypen/Modi, native Ownership, aktive Model-/Makro-Leases,
keine automatische Verweisrefaktorierung oder vollständige Access-Close-Parität.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; DEB gehört9c80f00.
NATIVE_MACRO_PROJECT_COMMANDS.md; F088/F050/F057/F114/F121/F135/F140/F148/F155
und456/384/72 unverändert. Kontinuierlich weiter.


SelectObject als native Makroaktion und formularfreier DoCmd-Member ergänzt:
wirklich offene Tabellen/Abfragen über bestehende Aktivierung, echte Form-/Report-
Instanzen über lebensdauergeschützten Preview-Dispatch, eigene Makro-/Modulobjekte
und Ordner im Datenbankfenster. Gleiche4/5-InDatabase-Erweiterung im Formularhost;
kein erfundenes Modulfenster, kein Öffnen eines geschlossenen Datenobjekts.
Explorer behält tatsächliche logische Auswahl/Archiv-ID bei Refresh derselben
Projektgeneration; neues Projekt übernimmt keine alte Auswahl. Damit bleibt die
von einem Makro gesetzte Auswahl nach dessen abschließendem Refresh erhalten.
Gewachsene NativeStandardModuleHost/DoCmd/Application-Implementierung aus Header
in eigene Forms-C++-Einheit verschoben (öffentliche schlanke Deklaration); zwei
mechanische Qualifizierungs-/Duplikatfehler vor Testlauf korrigiert. Keine Änderung
an Kern-Qt-Freiheit oder Python/Java-App-Abhängigkeiten. App gebaut, bestehende
GUI/Layout2/2 (19,03s) bestanden: reale Table/Query/Form/Report-Aktivierung,
Macro/Module-/Ordnerauswahl, Refresh und tatsächlicher Menü-Makro-Endzustand,
Formularhost4/5, fehlende/geschlossene/unsupported Ziele. Vorige Modul-/Makro-/
Speicher-/Startprüfungen im gleichen Ziel weiter enthalten. Aktuelles DEB gebaut
und Payload/Rechte/App-/Helper-BuildIDs/MDBTools1.0.1/bytegleiches SQLite3.54.1
samt realem JOIN/NULL-/Countercheck verifiziert. Keine neue CTest-ID/Gesamtsuite/
Windows-Beobachtung. NATIVE_MACRO_PROJECT_COMMANDS.md; F088/F050/F135/F140/F148
und456/384/72 unverändert. Direkt weiter mit StopAllMacros/Echo-Wiederherstellung.


StopAllMacros als echte parameterlose Makroaktion implementiert, kein erfundener
DoCmd-Member. Gemeinsamer Stop-Zustand beendet den aktuellen verschachtelten
Aufrufbaum, Wiederholungen und bedingungsgesteuerte Fortsetzungen; nach äußerstem
Unwind sind Zustand/Budgets/Frames wieder frei. Bereits laufender nativer VBA-
Functioncode wird nicht global abgebrochen; folgende Makroaktionen bleiben aber
beendet. Neue unabhängige Rootaufrufe funktionieren. Echo-Aktion und Echo-/
Application.Echo im formularfreien Host nutzen bestehende reale UI-Suppression;
StopAllMacros stellt implementiertes Echo wieder her. Vollständiger SetWarnings-
Zustand/Access-Stop-/Reentry-Scheduling bleiben offen, keine Rücksetzung eines
nicht implementierten Warnungszustands behauptet. App gebaut, vorhandene GUI/
Layout2/2 (19,55s) ohne Korrekturlauf bestanden: tatsächliche MsgBox bei ausgeschal-
teten Updates, verschachtelter SQL/RunCode/Stop beider Hosts, native VBA-Fortsetzung,
fehlende nachfolgende Aktionen/Wiederholungen, unabhängiger Folgestart, Stop aus
RepeatExpression, False-Stop ohne vorzeitige Echo-Restauration, benanntes
Application.Echo, ungültiger Status ohne UI-Mutation, kein DoCmd.StopAllMacros.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; DEB gehört cb8931b.
NATIVE_MACROS_RUNTIME.md; F088/F050/F051/F059/F135/F148 und456/384/72 unverändert.
Direkt weiter mit vollständigem Aktionsargument-Editor/weiterer Aktionsausführung.


## Fortschreibung 2026-11-09: vollständige native Aktionsargumente / MessageBox
Argumente… bearbeitet sämtliche deklarierten Aktionsargumente; zusätzliche Werte
bleiben bei Teilmakrowechseln und Zeilenbewegungen erhalten, Abbrechen verändert
keinen Entwurf. MessageBox ist eine echte vierargumentige Makroaktion in beiden
Hosts: kontextgebundene =Auswertung, Signalton, benannter Typ, Titel, escaped
@-Abschnitte; kein erfundenes DoCmd.MessageBox. Native JSON-Typnamen sind kein
Access-Binärimport. Quellen-/Ergebnisgrenzen und bekannte Typen werden geprüft.
App gebaut; vorhandene GUI/Layout2/2 bestanden (19,30s), einschließlich echtem
Argumentdialog, vierter Titel, Abbruch, Teilmakrowechsel/Speichern, realen Dialogen
in beiden Hosts und unbekanntem Typ. Build-Include für vollständigen VbaValue
korrigiert, kein Testfehler/Korrekturlauf. Originaluploads3/3 und Fixturehashes4/4
unverändert. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung.
NATIVE_MACROS_RUNTIME.md; F088/F050/F051/F059/F135/F148,456/384/72 unverändert.
Weiter: GoToRecord aus tatsächlichen Standardmodulen/Makros mit vier Argumenten
an vorhandene Daten-/Formularnavigation anbinden; mehrzeilige Argumente erhalten.


## Fortschreibung 2026-11-09: GoToRecord aus Modul/Makro, mehrzeilige Argumente
Standardmodule Application.DoCmd.GoToRecord und native Makroaktion unterstützen
vier optionale/namentliche Argumente, tatsächliche offene Tabellen/Abfragen und
sichtbare Projektformulare. Bestehende Save-/Ereignis-/Reihenfolge-/Grenzlogik wird
wiederverwendet; keine erzeugten Ersatzdokumente. Native JSON-null ist hier eine
Auslassung, direktes VBA-Null bleibt ein Laufzeitwert. Detailargumente bearbeiten
Text mehrzeilig ohne SQL-Zeilenverlust; Bedingung bleibt separat.
App gebaut; Layout1/1 (11,40s), nach Korrektur nur des neuen Prüffall-Abfragemodus
auf Nur-Lesen GUI1/1 (7,61s) bestanden. Geprüft: tatsächliche Cursorpositionen,
benannter VBA-Modulaufruf, vierargumentiges Makro beider Hosts, optionale Löcher,
Tabellen-/Abfrage-/Formularwechsel, Offsetfehler ohne Cursoränderung, geschlossene
Form verweigert, mehrzeiliger Argumentdialog. Erste GUI verweigerte korrekt die
nicht unterstützte bearbeitbare sortierte Projektionsabfrage; keine Lockerung.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung. Modale Designer
und allgemeine versteckte Formularsuche bleiben explizite Grenzen.
NATIVE_MACROS_RUNTIME.md; F178/F088/F052 und456/384/72 unverändert.
Direkt weiter: vollständige FindRecord-Anbindung aus Standardmodulen/Makros.


## Fortschreibung 2026-11-09: FindRecord aus Standardmodul/Makro
Application.DoCmd.FindRecord unterstützt sieben/namentliche Argumente und echte
aktive Projektformulare, Tabellen und Abfragen. Gemeinsame Argumentprüfung mit
bestehendem Formularhost; typisierte Werte/Platzhalter/Formatgrenzen bleiben
unverändert. Native Makroaktion hat literaltextiges FindWhat und optionale Löcher;
keine Behauptung beliebiger Access-Makroausdrücke/Binärdekodierung.
Tatsächliche Tabellenöffnung wählt nun eine erste Zelle, damit FindRecord ohne
vorherigen manuellen Klick ein Suchfeld hat. Datensatzmarkierung selektiert die
aktuelle ganze Zeile ausdrücklich unabhängig vom initialen Qt-Auswahlanker und
bewahrt die aktuelle Spalte. Vorhandene GUI deckte diesen Folgefehler auf.
App gebaut; Layout1/1 (11,23s), nach gezielten Korrekturen GUI1/1 (7,91s) bestanden.
Der neue Suchprüffall setzt seine Sortierung ausdrücklich: die vorherige Annahme
aufsteigender Standardsortierung war falsch, der Zahlen-Treffer bereits korrekt.
Geprüft: sieben tatsächliche Editorargumente/Speicherung, Zahlen/Currency,
Platzhalter/Weitersuche, benanntes Modul, Nur-Lesen-Abfrage, Formular/Makro beider
Hosts, Null-/Mustervalidierung ohne Cursoränderung und vorhandene Zeilenauswahl.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; vollständige
No-Match-/Zeit-/Untertyp-Parität bleibt offen. NATIVE_MACROS_RUNTIME.md;
F088/F052/F178/F363 und456/384/72 unverändert. Direkt weiter mit tatsächlichen
Fokus-/Cursorbefehlen und gemeinsamer UI-Ausführungslebensdauer im Modulhost.


## Fortschreibung 2026-11-09: Fokus/Cursor und gemeinsame GUI-Ausführungsrahmen
GoToControl und Hourglass stehen nativen Makros und dem formularfreien DoCmd-Host
mit namentlichen Argumenten zur Verfügung. Fokus geht an echte Controls eines
aktiven Projektformulars; ohne Formular2046, kein erfundener Me-Kontext.
Modul-Eintritt/-Austritt und beide Makrohosts verwenden denselben nativen Qt-
Ausführungszähler. Wartezeiger bleibt über verschachtelte Funktionen/RunCode bis
zu späteren Aktionen erhalten und wird bei äußerstem normalem/fehlerhaftem Ende
zurückgesetzt; fremde Cursor bleiben geschützt. Echo bleibt separat persistent.
Vorhandene GUI/Layout2/2 (20,12s) bestanden, kein Korrekturlauf: benannter Fokus mit
Textauswahl, tatsächlicher Cursor in verschachtelten und nachfolgenden Dialogen
beider Makrohosts, StopMacro-Ende, Fehler7004, direkte Steuerung/fremder CrossCursor,
fehlender aktiver Formularkontext. App nach finalem Formularhost-Nachlauf über
CPack neu gelinkt. Frisches DEB native-module-ui/package/linaccess-0.7.0-Linux.deb:
Payload/Rechte/Build-ID/bytegleicheSQLite und tatsächlicher Backendcheck mit
MDBTools1.0.1/SQLite3.54.1 JOIN/NULL/Zähler bestanden. Paket ohne MDB/VBA/Fixtures.
Originaluploads3/3 und Fixturehashes4/4 unverändert. Keine neue CTest-ID/Gesamtsuite/
Windows-Beobachtung. NATIVE_MACROS_RUNTIME.md/NATIVE_HOURGLASS.md;
F088/F052/F068/F135/F148 und456/384/72 unverändert. Direkt weiter mit gemeinsamem
CurrentProject/CurrentData aus echten Standardmodulen und Projektlebensdauer.


## Fortschreibung 2026-11-09: gemeinsame CurrentProject-/CurrentData-Objekte
Standardmodule können Application.CurrentProject/CurrentData und die unqualifi-
zierten VBA-Bezeichner verwenden. Formular- und Modulhost nutzen dieselbe echte
Projektkatalogimplementierung und Objektidentität am tatsächlichen Projektfenster;
kein kopierter Katalog/kein erfundenes Formular. Cache zusätzlich generations-
gebunden. Vorhandene Datei-/Readonly-/Deskriptor-/IsLoaded-Grenzen bleiben bestehen.
App gebaut; vorhandene GUI/Layout2/2 (19,49s) ohne Korrekturlauf bestanden:
wirklicher VBA-Katalogzugriff/Counts, gemeinsame Identität in beiden Hosts,
Fortbestand nach Formularschluss, live Modul-Löschung, Argument/Readonly-Fehler,
3420 für gehaltene Projekt-/Datenkatalogobjekte nach tatsächlichem Projektwechsel
im vorhandenen AutoExec-Ablauf. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/
Paketprüfung; zuletzt geprüftes DEB gehört69ac1b9. NATIVE_STANDARD_MODULES_RUNTIME.md;
F052/F042/F059/F135 und456/384/72 unverändert. Direkt weiter: echte Forms-/Reports-
Sammlungen und Screen aus dem formularfreien Host, ohne Platzhalterformular.


## Fortschreibung 2026-11-09: tatsächliche Forms/Reports/Screen aus Standardmodulen
Formularfreier Host unterstützt Forms/Reports und Application.Forms/Reports mit
realen Dokumentobjekten, Count/Name/Index/ForEach, einschließlich nach Öffnen
versteckter Formulare. Application.Screen/bare Screen liefern tatsächliches
ActiveForm/ActiveControl/MousePointer; vorhandene Cursor-/Fokusgrenzen bleiben.
Gemeinsame Implementierung statt Platzhalterformular. Neue Projektobjekte sind
via echtem CurrentProject-Lebenszyklus geschützt; alte formulargebundene Handles
behalten ihren bestehenden Caller-Lebenszyklus. Screen-Suche zusätzlich an DB-
Identität gebunden und nach skalarer Wertauswertung erneut geprüft.
App gebaut; vorhandene GUI/Layout2/2 (20,16s) ohne Korrekturlauf bestanden:
echte sichtbare/versteckte Formulare und Bericht, index-/namensbasierte identische
Dokumentreferenzen, native VBA-ForEach über beide Sammlungen, reale aktive Form/
Control, VBA-MousePointer mit Rücksetzung, live Schließen/alte Dokumentreferenz,
3420 für gehaltene Sammlung/Screen nach tatsächlichem Projektwechsel.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; DEB69ac1b9.
NATIVE_STANDARD_MODULES_RUNTIME.md; F052/F042/F059/F135 und456/384/72 unverändert.
Direkt weiter: globale projektgebundene Sammlungsidentität und tatsächlich
geöffnete/versteckte Formulare als benannte Navigations-/Schließziele.


## Fortschreibung 2026-11-09: globale Projekt-Sammlungsidentität/versteckte Ziele
Native Forms/Reports/Screen teilen einen schwachen, verbindungs-/datei-/generations-
gebundenen Cache am echten Projektfenster, auch über kurzlebige Application-
Wrapper hinweg. Projektwechsel leert den Cache; gehaltene Altobjekte bleiben
3420-geschützt. Tatsächlich geöffnete lebende Dokumente inkl. versteckter/modaler
Instanzen werden nach nativer Öffnungsreihenfolge erfasst; Schließen/Abbruch wird
berücksichtigt. Benanntes GoToRecord und Close erreichen nun echte versteckte
geöffnete Formulare, ohne sie zur Navigation sichtbar zu machen.
App gebaut; vorhandene Layout1/1 und korrigierte GUI1/1 (7,96s) bestanden. Der neue
Prüffall wurde berichtigt: Konstruktor dispatcht bereits Open/Load vor show();
eine erfolgreich geladene noch nicht angezeigte Instanz zählt deshalb richtig
als geöffnet. Kein Produktionsfehler verdeckt. Geprüft: direkte native VBA-Is-
Identität ohne gehaltenen Application-Wrapper, reale unangezeigte/ausgeblendete
Instanzen, versteckte Navigation ohne Anzeigen, benanntes Close und ungültig
gewordene Dokumentreferenz. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/
Paketprüfung; DEB69ac1b9. NATIVE_STANDARD_MODULES_RUNTIME.md;
F052/F178/F042/F135 und456/384/72 unverändert. Direkt weiter mit echtem FileDialog
für Standardmodule und zugehöriger Projekt-/Dialogzustandslebensdauer.


## Fortschreibung 2026-11-09: echte FileDialog-Picker aus Standardmodulen
Application.FileDialog(DialogType) verwendet den vorhandenen tatsächlichen Datei-/
Ordnerpicker (3/4), Office.FileDialog-Typ, Filter/Show/SelectedItems. Keine
Ersatzform, kein injiziertes Ergebnis, Auswahl öffnet/führt keine Datei aus.
Projektfenster hält den Pickerzustand; tatsächlicher Projektwechsel leert den
Cache. Alte Picker bleiben3420-geschützt, neuer Projektaufruf erzeugt funktionie-
renden neuen Picker. Application besitzt jetzt zusätzlich eine eigene Datei-/
Generationslease, unabhängig von einer optionalen Prüfung des Aufrufers.
App gebaut; bestehende Layout1/1 (11,31s) und abschließende GUI1/1 (8,15s) bestanden:
reale temporäre Datei/Ordner, Abbruch/geleerte Auswahl, Filter, typisierter
namentlicher VBA-Aufruf, Open1 verweigert/Null94 sowie tatsächlicher Projektwechsel
mit frisch erzeugtem Application-/Pickerobjekt bei ausdrücklich leerer Zusatz-
prüfung des Hosts. Finale App/GUI auf demselben Quellstand neu gebaut.
Frisches DEB native-module-file-dialog/package/linaccess-0.7.0-Linux.deb:
Payload/Rechte/Build-ID/bytegleicheSQLite und tatsächliches MDBTools1.0.1/
SQLite3.54.1-Backend (JOIN/NULL/Zähler) bestanden; keine MDB/VBA/Fixtures enthalten.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung. NATIVE_FILE_DIALOG.md;
F052/F042/F059/F135 und456/384/72 unverändert. Direkt weiter mit SysCmd am echten
Projekt-Statusbalken/GetObjectState und weiteren vorhandenen Application-Diensten.


## Fortschreibung 2026-11-09: wirkliche Application-Dienste/gemeinsames SysCmd
Standardmodule unterstützen Application.CurrentDb (vorhandener echter DAO-Adapter)
und Application-Domänenfunktionen mit Expr/Domain/Criteria über die bestehenden
Implementierungen. SysCmd1..5 verwendet realen Projektstatus/Meter,10 den echten
CurrentProject-Objektzustand; bare und namentliche Application-Aufrufe unterstützt.
Beide Hosts teilen Argumentprüfung/einmalige Skalargetter-Auswertung mit erneuter
Kontextprüfung. Native Statusoberfläche erkennt jetzt Hauptfenster selbst und
verwendet dessen echten Statusbalken; korrigiert auch Echo-Statustextposition.
App gebaut; bestehende GUI/Layout2/2 (19,72s) ohne Korrekturlauf bestanden: echte
VBA-DAO-Lese-/Domänenaufrufe, benannte Parameter, Status/Meter/Entfernen, tatsäch-
licher Tabellen-Ladezustand, einmaliger Standardwertgetter beider Hosts und Echo
am Projektbalken. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung;
DEBf4e9ff8. NATIVE_STANDARD_MODULES_RUNTIME.md; F052/F053/F135 und456/384/72 bleiben
unverändert. Direkt weiter mit vollständiger tatsächlicher OpenForm-/OpenReport-
Anbindung (Where/DataMode/OpenArgs/WindowMode) über gemeinsame Dokumentöffnung.


## Fortschreibung 2026-11-09: gemeinsame tatsächliche Dokumentöffnung
Standardmodul-DoCmd.OpenForm hat sieben, OpenReport sechs Argumente und nutzt nun
dieselbe tatsächliche Öffnung wie der Formularhost: gespeicherte/importierte
Objekte, WhereCondition, frische Datenmodi, OpenArgs vor Load, native Fenster-
modi/acDialog-Warten, unabhängige Projektzuordnung und verzögerte Freigabe. Keine
Platzhalterform. Importquellen main-qualifiziert; explizite Where-Quellfehler
behalten Vorrang und Initialisierungsfehler jetzt Source-/Help-Metadaten.
Normalformular0/Berichtsvorschau2 bleiben unterstützt; andere Ansichten/benannte
Filter und unerlaubte Änderung bestehender Instanzen weiterhin3251. Native Makro-
aktionen haben bis zum nächsten Editorumfang weiterhin nur Name/Ansicht.
App gebaut; vorhandene GUI/Layout/Original-Kassa3/3 (137,86s;9,24/12,39/116,22s)
bestanden. Neuer Prüffall-Cast korrigiert, importiertes Testlayout mit tatsächlicher
OnLoad-Bindung ergänzt und alte Report-Window-Verweigerung durch echtes Aus-/
Einblenden ersetzt. Geprüft: Filter/Readonly/OpenArgs/Load, Wiederverwendung ohne
Mutation bei verweigertem Modus, echtes Add-Insert, modales Warten, versteckter
Bericht, Fehler7006 inkl.Source/Help und tatsächlicher importierter Handler.
Original-F129/3075 bleibt bekannte Grenze. Keine neue CTest-ID/Gesamtsuite/Windows-
Beobachtung/Paketprüfung; DEBf4e9ff8. NATIVE_STANDARD_MODULES_RUNTIME.md;
F052/F178/F042/F135 und456/384/72 unverändert. Direkt weiter: optionale Makro-
argumente einschließlich ausdrücklich leerem OpenArgs verlustfrei bearbeiten und
vollständige OpenForm-/OpenReport-Makroargumente ausführen.


## Fortschreibung 2026-11-09: verlustfreie optionale Makroargumente
Argumente… hat Übergeben-Schalter und bewahrt ausdrücklich leeren Text getrennt
vom Standardwert. Präsenz über Laden/Teilmakrowechsel/Aktionsreihenfolge/Speichern
und Zusatzargumente erhalten; Übersichtsänderungen inferieren neu, Pflichtwerte
bleiben erforderlich. Vollständige native OpenForm7-/OpenReport6-Makroargumente.
Optionales JSON-null bezeichnet einheitlich Missing, kein Datenbank-/VBA-Null;
Pflicht-null bleibt ungültig. Alte Formularbefehle nutzen ihre bewährten Defaults,
MessageBox verarbeitet ausgelassene Werte explizit. Kein erfundenes DoCmd-API.
Vorhandene GUI/Layout2/2 (20,59s) ohne Korrekturlauf bestanden: echte Editor-
Schalter/Move/Gruppe/Save/Reopen, leeres OpenArgs bleibt String und wird in beiden
Hosts tatsächlich von Null unterschieden; reale alte Befehle mit Default-Löchern
und MessageBox-Icon-Default. Anschließend nur irreführende optionale Übersichts-
spaltenbeschriftungen entfernt, keine Logikänderung. Frische App/DEB native-macro-
presence/package/linaccess-0.7.0-Linux.deb mit Payload/Rechten/Build-ID/bytegleicher
SQLite und echtem MDBTools1.0.1/SQLite3.54.1-Backendcheck (JOIN/NULL/Zähler) geprüft.
Uploads3/3 und Fixtures4/4 hashgleich. Keine neue CTest-ID/Gesamtsuite/Windows-
Beobachtung. Originaltest zuletzt f86d8eb:3/3 inkl.Original116,22s; F129/3075 offen.
NATIVE_MACROS_RUNTIME.md; F088/F052/F178/F135 und456/384/72 unverändert.
Direkt weiter mit echter SetValue-Makrozuweisung über vorhandene native Formular-
/Steuerelementobjekte und bestehende Ausdrucksauswertung; dokumentierte Grenzen.


## Fortschreibung 2026-11-09: tatsächliche SetValue-Makrozuweisung
SetValue(Item,Expression) greift auf echte native Controls/ausdrückliche Formular-/
Berichts-/Control-Eigenschaften und vorhandene Eval-/Setterwege zu, keine erzeugte
VBA-Zuweisung/keine Ersatzform/kein erfundenes DoCmd.SetValue. Bare Controls brauchen
echten Aufrufer/aktives Dokument, explizite Forms-Ziele funktionieren ohne Me.
Syntax/4096-Grenzen vorab, dokumentierte führende=-Doppelauswertung ausdrücklich
3251, Ziel-/Kontextprüfung vor/nach Auswertung. Vorhandene Readonly-/Puffer-/Save-
und Validierungsregeln gelten; Feld-ohne-Control/Funktionsziel-/Vollreport-/Ereignis-
parität bleibt offen. App gebaut; Layout1/1 (11,34s), abschließende GUI1/1 (8,49s)
bestanden. Gezielter Prüffall nach Einblenden wählt tatsächliche Form ausdrücklich;
vorher438 war mangels aktivem Dokument korrekt. Reale Zuweisung/Speicherung beider
Hosts, explizites verstecktes Ziel/Visible, Null, Fehler11/2448 ohne Wertänderung,
Schließen während RHS2450, ungültiges Rohmakro vor SQL-Präfix, DoCmd.SetValue438.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; DEB22a9284.
NATIVE_MACROS_RUNTIME.md; F088/F052/F030/F135 und456/384/72 unverändert.
Direkt weiter mit tatsächlichem RepaintObject für geöffnete native Dokumente.


## Fortschreibung 2026-11-09: tatsächliches RepaintObject
RepaintObject(ObjectType,ObjectName) in Standardmodulen/Formularen und nativen
Makros zeichnet echte geöffnete Tabellen/Abfragen/Modul-/Makroentwürfe sowie
Formulare/Berichte neu. Form.Repaint verwendet vorhandene berechnete Controls,
Berichte rendern ihren vorhandenen Datensatzbestand erneut; keine Requery oder
implizite Speicherung. Projekt-/Generation-/Lebensdauer- und Kontextprüfungen
bleiben wirksam. Tabellen-Cache, laufende Zelleneingabe und Entwurfsänderungen
bleiben erhalten. GUI1/1 (9,43s), Layout1/1 (11,34s) bestanden; App gebaut.
Der letzte gezielte Korrekturlauf betraf ausschließlich den Prüffall: Speichern
per Schaltfläche schließt den Editor bestimmungsgemäß; Strg+S hält ihn offen.
GUI-Erweiterung in kleinere Quelldatei desselben bestehenden CTest-Ziels getrennt,
keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; DEB22a9284.
Vollständige Access-Neuzeichnungs-/Ereignis-/Berichtparität bleibt offen;
F088/F052/F050/F135 und456/384/72 unverändert.


## Fortschreibung 2026-11-09: SetValue-Ausdruck mit führendem Gleichheitszeichen
Expression mit führendem= wertet erst den Rest einmal aus und danach dessen
Skalarwert als Ausdruckstext einmal. Reale Referenzen/Funktionen beider Hosts
verwenden vorhandene Eval-/Setterwege; keine bloße Entfernung des=, keine
unbegrenzte Rekursion. Null94/Errornummer, Syntax-/4096-Textgrenzen, erneutes=3251
und Kontext-/Lebensdauerprüfungen vor Zuweisung. Dynamisches Item bleibt offen.
App gebaut, vorhandene GUI/Layout2/2 (20,35s) ohne Korrekturlauf bestanden:
Zahl, referenzierender Text, tatsächlicher öffentlicher Funktionsaufruf exakt
einmal je Host (SQL-Audit), sechs Fehlerfälle ohne Wert-/Dirty-Änderung und reale
Speicherung. Ungültiger Rohmakro-Präfixtest verwendet nun tatsächlich ungültige
Syntax=1+; vollständiger Vorabcheck vor SQL bleibt belegt. Keine neue CTest-ID/
Gesamtsuite/Windows-Beobachtung/Paketprüfung; bisheriges DEB22a9284.
NATIVE_MACROS_RUNTIME.md; F088/F052/F118/F164 und456/384/72 unverändert.
Direkt weiter mit echten Formularfeldern ohne gebundenes Steuerelement.


## Fortschreibung 2026-11-09: echte Formularfelder ohne Steuerelement
Native FormField-Referenzen lesen/ändern tatsächliche RecordSource-Spalten ohne
Ersatz-Controls; Name/Value/Parent und Lebensdauer geprüft, Controls haben Vorrang,
bereits gebundene Feldnamen verwenden tatsächliches Control. Eigener typisierter
Datensatzpuffer ist in Save/Undo/ControlUndo/Neuanlage/Navigation und berechnete
Controls integriert. Parameterisierte gemeinsame Speicherung mit anderen Controls,
CHECK-Fehler behalten Entwurf, AutoWert/Readonly/geschlossene Form verweigern.
VBA und SetValue beider Hosts verwenden denselben Weg; keine DAO-/Berichts-/
Join-/Vollereignisparität behauptet. FORM_RECORD_FIELDS_RUNTIME.md.
App gebaut; abschließende GUI1/1 (8,50s), Layout1/1 (11,32s), Original1/1 (124,97s)
bestanden. Korrekturlauf änderte ausschließlich den ControlUndo-Prüffall auf
echte Tastatureingabe statt programmatischer Zuweisung ohne Eingabepuffer.
Original-F129/3075 bleibt erwartete Grenze. Neues DEB native-form-record-fields/
package/linaccess-0.7.0-Linux.deb mit tatsächlichem MDBTools1.0.1/SQLite3.54.1,
JOIN/NULL/Zähler, Payload/Rechten/Build-ID/fehlenden Fixtures geprüft. Uploads3/3
und Fixtures4/4 hashgleich. Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung.
F056/F050/F088/F052 und456/384/72 unverändert.
Direkt weiter mit echten ApplyFilter-/ShowAllRecords-Dokumentfiltern.


## Fortschreibung 2026-11-09: tatsächliche Dokumentfilter
Gemeinsamer nativer Filterweg für echte Datenblatt-Dialoge und DoCmd/Makro
ApplyFilter/ShowAllRecords beider Hosts. Jet-Prädikat ohne zusätzliche Klauseln/
Subqueries; VBA32768/Makro256 UTF-16-Grenzen, Null94, benannte Filter3251. Tabellen
behalten Entwürfe (Filterwechsel2501), bearbeitbare Abfragen behalten Basis-WHERE.
Readonly-Ergebnisse verwenden tatsächliche Quelle/Parameter und werden erst nach
Erfolg ausgetauscht; Requery behält Filter. Formulare nutzen vorhandenen atomaren
Filter-/Validierungsweg; Berichte nur tatsächliches gebundenes Open-Ereignis,
Kontext nach Verschachtelung wiederhergestellt. DATASHEET_FILTER_RUNTIME.md.
App gebaut; abschließende GUI1/1 (8,62s), Layout1/1 (11,77s) bestanden. Prüffall
sortiert technische Tabellen ausdrücklich statt Qt-Standardreihenfolge zu raten.
Wiederherstellung berücksichtigt Qt: nach fehlgeschlagener Selektion ist Query
inaktiv und setFilter allein selektiert nicht erneut; explizite Auswahl ohne
doppelte erfolgreiche SQL-Ausführung. Reale UI/Cancel/SQL/Dirty/LIKE/Datum/
Parameter-/Abfrage-/Makro-/Form-/Report-Open-Prüfungen, kein neues CTest-Ziel.
Keine Gesamtsuite/Windows-Beobachtung/neue Paketprüfung; DEB073c20a, Original
zuletzt073c20a1/1 (124,97s), erwartetes F129/3075 unverändert offen.
F080/F056/F052/F088/F087 und456/384/72 unverändert.


## Fortschreibung 2026-11-09: echte Fensterbefehle und kürzere Änderungswege
Standardmodule/Makros unterstützen Maximize/Minimize/Restore ohne Ersatzform/Me.
Tatsächliche aktive Form/Report delegiert bestehende native Befehle, Tabellen/
Abfragen/Explorer adressieren echtes Projektfenster. Spezialisierter VBA-
Maximize-Statementpfad überschreibt tatsächliche Host-API; Minimize/Restore
laufen über DoCmd, dafür existieren keine virtuellen Hostmethoden. Modulfenster
3251, andere modale Entwürfe2046; keine implizite Speicherung/Entwurfsverwerfung.
Makroparser/Katalog/Loader einmalig als native_macro.cpp in bestehendem Daten-
Staticlib statt Headerdefinitionen; kleinere GUI-Erweiterungsdatei erhält festen
Einstiegspunkt. Keine neue DSO/GUI-Abhängigkeit des Datenlayers/CTest-ID.
App gebaut; abschließende GUI1/1 (8,69s), Layout1/1 (11,29s) bestanden. Gezielte
Prüffallkorrektur berücksichtigt Qt-Kombination minimiert+zuvor maximiert statt
fälschlich exklusive Statusflags zu verlangen. Reale Fensterzustandsereignisse,
Statement/Makros, unveränderte SQL-/Cache-/Formwerte, extra450, tatsächlicher
Moduleditor mit erhaltenem Entwurf. NATIVE_WINDOW_STATE_RUNTIME.md.
Volle Access-MDI-/Popup-/Geometrie-/Windows-Ereignisparität bleibt offen. Keine
Gesamtsuite/Windows-Beobachtung/neue Paketprüfung; DEB073c20a, Originalzuletzt
073c20a1/1 (124,97s), erwartetes F129/3075 offen. F052/F088/F135 und456/384/72
unverändert. Direkt weiter mit zusammenhängenden Makro-Bedingungsblöcken.


## Fortschreibung 2026-11-09: zusammenhängende Makro-Bedingungsblöcke
Condition=... übernimmt einmal ausgewertetes Ergebnis des unmittelbar vorigen
Bedingungsblocks. Leere Zeile beendet Block, neue Bedingung startet neuen;
Wiederholungen/Unteraufrufe haben eigenen Zustand. Verwaiste Fortsetzungen und
Gruppenanfänge werden im vollständigen Vorabcheck3251 vor SQL-Präfix abgelehnt.
Bestehende Aktions-/Tiefenbudgets zählen übersprungene Einträge weiter. Tatsächlicher
Makroeditor speichert Fortsetzungstext, Standard-/Formhost führen echte SQL-Audits
aus. Runner/Frame-/Stop-Helfer nach native_macro_runner.cpp im bestehenden
Daten-Staticlib ausgelagert, UI bleibt Callback, keine neue GUI-Abhängigkeit/DSO.
App gebaut; bestehende GUI/Layout2/2 (20,25s; 8,89/11,36s) ohne Korrekturlauf
bestanden: echter Editor, Funktion exakt einmal pro Block, True/False-Repeats,
Child isoliert Parent, defekte andere Gruppe/Präfix ohne Ausführung, Folgelauf
sauber. MACRO_CONDITION_BLOCKS_RUNTIME.md; keine neue CTest-ID/Windows-Referenz/
Gesamtsuite. Frisches DEB native-macro-continuation/package/linaccess-0.7.0-Linux.deb
mit Payload/Rechten/Build-ID/echtem MDBTools1.0.1 und SQLite3.54.1-JOIN/NULL/Zählern
geprüft; umfasst auch Filter/Fensterbefehle. Uploads3/3 und Fixtures4/4 hashgleich.
Originalzuletzt073c20a1/1 (124,97s), erwartetes F129/3075 bleibt offen.
Vollständige Access-Makro-/Ereignisparität bleibt offen. F088/F052/F135 und
456/384/72 unverändert. Direkt weiter mit SQL-Zugriff auf tatsächliche skalare
Formular-/Berichts-/Steuerelementeigenschaften ohne Methodenaufrufe.


## Fortschreibung 2026-11-09: echte skalare Objekteigenschaften in SQL
Jet-Tokenadapter unterstützt Forms!Dokument.Eigenschaft und Control.Eigenschaft,
auch Reports/lokalisierte Namen/Bracket-/Punktnamen. Bestehende Werte/.Value
bleiben auf alter3er-API, neue4er-Property-API ist driver-/verbindungsgebunden.
Explizite Getter-Allowlist verwendet tatsächliche Form-/Report-/Control-/
Datensatzfeldobjekte; Undo/Requery/Objektgetter/Methoden werden nicht ausgeführt
(438), tiefere/parametrisierte Pfade bleiben offen. Bisherige skalare SQL-
Konvertierung wiederverwendet, keine neue Typ-/Currency-/Decimal-Vollbehauptung.
Gemeinsamer Lookup prüft GUI-Thread, Connection/Pfad/Kind/echte Lebensdauer,
mehrdeutige Namen5 und nach Getter erneut tatsächliche Dokumentgültigkeit.
SQL_OBJECT_PROPERTIES_RUNTIME.md. App und Ziele gebaut; gebündelt4/4 (115,92s)
ohne Korrekturlauf bestanden: Layout11,68s, GUI9,01s, SQL-Objekte0,02s, Original
115,91s. Reale Caption/Tag/Visible/Dirty/OldValue/ListCount, native versteckte
Feldentwürfe, Live-View nach Änderung, fehlender Kontext, API-Reopen, echte
Verbindungsisolation/Mehrdeutigkeit, Getterverweigerung ohne Undo, Form-/
Reportfehler2450/2451. Original-F129/3075 bleibt erwartete Grenze.
Keine neue CTest-ID/Gesamtsuite/Windows-Beobachtung/Paketprüfung; DEB72ed466.
F067/F068/F118/F164/F050/F052 und456/384/72 unverändert.


## 2026-11-09: nativer VBA-Debugger, F060

Implementiert: tatsächliche gespeicherte Standardmodul-Ausführung mit F9/F5/F8,
StepInto/Over/Out, lebenden Aufrufrahmen/ByRef-Locals, paginierter Array- und
vollständiger Stringinspektion ohne Objektgetter, unbehandelten Fehlern auch
bei normalem F5, Resume-Next-Überspringen und explizitem Abbruch. Keine erneute
Ausführung bereits erfolgter SQL-Aktionen zum Fehlerhalt. Quelle/Projektwechsel
sind während des Halts geschützt; beendete Rahmen werden nicht weiter gelesen.
Details und Grenzen: docs/NATIVE_VBA_DEBUGGER_RUNTIME.md.

Nachweise in bestehenden Zielen: Kern1/1 (2,18s), Layout1/1 (11,80s), originales
Kassa-Ziel1/1 (118,52s), abschließender GUI-Ablauf1/1 (9,06s) sowie separat frisch
GUI-freier Kern1/1 (1,89s). Abschließendes DEB/Build-ID/MDBTools1.0.1/SQLite3.54.1
und vier Fixture-/drei Uploadhashes bestanden. Kein neuer Vollsuite-, Windows-,
Distributionsmatrix- oder GitHub-CI-Nachweis. Original-F129/3075 bleibt offen.
F060 bestätigt diesen beschriebenen nativen Teilumfang; vollständige IDE-,
Makrodebugger-, Sprach-, COM- und Ereignisparität bleibt in den übrigen offenen
Anforderungen. Fester Katalog456:385 bestätigte Teilumfänge/71 offene
Vollanforderungen, ausschließlich theoretisch und ungewichtet. Kein Gesamtabschluss.


## 2026-11-09: projektbezogene Einstufung, F090

Extras-Bericht mit deutschem Statusbaum und explizitem JSON-Export für den
wirklichen Import-/Objekt-/Tabellenkatalog. Getrennte statische Syntaxannahme,
Kontextgrenzen und tatsächlich beobachtete VBA-/Makro-/SELECT-Aufrufe samt
Fehlern und Abbruch. Verschachtelte Aufrufe, Quellrevisionen, nicht gestartete
Prozeduren, ungültige Definitionen und Rollback-Beobachtungen nachgewiesen.
Keine automatische Geschäftsprozedurausführung, keine VBA-/Argument-/
Datensatzwerte im Export. Projektwechsel leert die maximal1024 Sitzungsbelege;
gleichlautende Quellen belegen ausdrücklich keine konkrete Objektinstanz.
Andere SQL-/Tabellenmodellpfade und komplettes Dokumentrendering gelten als
nicht beobachtet, niemals als erfolgreich. Scope: docs/PROJECT_CAPABILITY_REPORT.md.

Gebündelter Lauf: Kern1/1 (2,04s), Layout1/1 (12,09s), Original1/1 (120,17s);
GUI-Testreihenfolge nach absichtlich defektem Modul korrigiert. Abschließende
gezielte GUI/Layout-Nachprüfung2/2 (13,93s; GUI13,93s/Layout12,54s), frisch
GUI-freier Kern1/1 (2,02s). Tatsächliche Makrorevision und gültiger Projektkontext
für Beobachtungen geschützt. Abschließendes DEB/Build-ID/MDBTools1.0.1/
SQLite3.54.1 und vier Fixture-/drei Uploadhashes bestanden. Keine neue CTest-ID,
keine aktuelle Vollsuite/Distributionsmatrix/GitHub-CI behaupten. Original-F129/
3075 bleibt offene Ausführungsgrenze. F090 bestätigt den beschriebenen
Klassifizierungsumfang, keine vollständige Ausführungs- oder Access-Parität.
Fester Katalog456:386 bestätigte Teilumfänge/70 offene Vollanforderungen,
ausschließlich theoretisch und ungewichtet. Kein Gesamtabschluss.


## 2026-11-09: Boolean-Bedingungen und If-Fehlerfortsetzung

Gemeinsame native Let-Boolean-Konvertierung für If/ElseIf, Do-/While-Köpfe und
Fußbedingungen, IIf sowie Makrobedingungen/Wiederholung: skalare Standardwerte
einmal, Null als False, True/False-Strings und case-sensitive #TRUE#/#FALSE#.
CBool(Null) bleibt94; Fehlerwerte/Arrays/Nothing werden nicht als True versteckt.
Belegte Tokenformen mit angrenzenden Klammern, EndIf/End If/End With mit normalen
Tokenabständen und leere einzeilige Then-/Else-Teile. EndIf vor Doppelpunkt ist
kein Label; leeres Case wird vor Effekten verweigert. Fehlgeschlagene If/ElseIf-
Köpfe behalten ausstehende Rahmen für Resume, Resume Next und Debugger-Skip;
keine irrtümlichen Then-/Else-Effekte und keine beschädigten äußeren If-Rahmen.
Scope/Sprachspezifikationsbeleg/Grenzen: docs/VBA_BOOLEAN_CONDITIONS.md.

Gezielte Nachweise nach vollständiger Implementierung/Korrektur der technischen
Prüfszenarien: Kern1/1 (2,22s), GUI1/1 (13,00s), abschließend aktuelles originales
Kassa-Ziel1/1 (113,95s), frisch GUI-freier Kern1/1 (2,00s). Tatsächliche Makro-
SQL-Effekte, Null-Standardwerte/Getterzahl, alle Bedingungspositionen, Resume und
Debugger-Skip geprüft. Finales DEB/Build-ID/MDBTools1.0.1/SQLite3.54.1 sowie
vier Fixture-/drei Uploadhashes bestanden. Kein neuer CTest, kein aktueller
Gesamtsuite-/Windows-/Distributionsmatrix-Nachweis. F129/3075 bleibt erwartet.
F045/F041/F051/F088 und456/386/70 bleiben offen/unverändert; kein Gesamtabschluss.


## 2026-11-09: fehlerhafte Schleifen-/With-Einstiege

Native Resume-Next-/Debugger-Skip-Fortsetzung überspringt fehlgeschlagene
While/Do-/For-/For-Each-Einstiege bis zur tatsächlichen Blockgrenze. Kein
Rumpflauf ohne gültigen Schleifenrahmen; äußere Schleifen/If bleiben erhalten.
With hält die äußere Aktivierung getrennt vom ausstehenden Objektkopf, damit
sein fehlgeschlagener Rumpf inaktiv bleibt und End With den äußeren Zustand
wiederherstellt. Wiederholendes Resume versucht den wirklichen Kopf erneut.
Scope und Grenzen in docs/VBA_BOOLEAN_CONDITIONS.md.

Ein gebündelter bestehender Kern-/GUI-Lauf2/2 (13,20s; Kern2,13s/GUI13,20s)
ohne Korrekturlauf bestanden. Aktuelle App und DEB/Build-ID/MDBTools1.0.1/
SQLite3.54.1 geprüft. Letzter Original- und GUI-freier Kernnachweis bleibt der
unveränderliche vorherige dae6681 (Original113,95s/Kern2,00s), kein neuer
Original-/Qt-frei-/Gesamtsuite-/Windows-/Distributionslauf behaupten.
Keine neue CTest-ID oder Katalog-ID; F045/F051 und456/386/70 bleiben offen/
unverändert. Kein Gesamtabschluss; weitere Arbeit läuft fort.


## 2026-11-09: Datensatzcursor und übergeordnete Requery-Berechnung

Schreibgeschützte Abfrageformulare wählen nach leer→gefüllt den ersten echten
Datensatz. Nur zulässige direkte Tabellenformulare behalten eine neue Leerzeile.
Geänderte tatsächliche Kind-Snapshots aktualisieren berechnete Elternwerte, ohne
Elterneingaben zu speichern; unveränderte Snapshots bleiben zwischengespeichert.
Details und Grenzen: docs/READONLY_FORM_REQUERY_RUNTIME.md.

Gebündelt Layout/GUI2/2 (14,90s), originales Kassa-Ziel1/1 (118,10s), finales
DEB/Build-ID/MDBTools1.0.1/SQLite3.54.1 und vier Fixture-/drei Uploadhashes
bestanden. Originale Keyboard-Grenze jetzt3146/1023; echter getaggter Klick
reicht bis2485 (fehlende exportierte Makroquelle). Beide MDB-Kataloge enthalten
vier Makros, bisheriges Brückenarchiv null. Kein neuer Gesamtsuite-/Windowslauf,
keine vollständige F129-Erfüllung. Katalog456/386/70 bleibt unverändert.


## 2026-11-09: fehlende Makroquellen gezielt exportierbar

Vorhandene Windows-Brücke exportiert nun AllMacros/SaveAsText acMacro4. Optional
-MacrosOnly schließt MDB/VBA-/Form-/Bericht-/SQL-Quellenexport aus;
-IncludeQueryMetadata liest vorhandene lokale QueryDef-Feld-/Parametermetadaten
ohne Recordsets/Aktionsabfragen. Externe Quellen werden ausgelassen, Fehler
einzeln ausgewiesen. START.cmd bietet Access2003-Dateiauswahl, Testkopie,
Originalhashvergleich und Ergebnis-ZIP auch bei Fehlern. Makrodefinitionen bleiben
private Projektquellen außerhalb Git. Beide Skripte syntaktisch mit offiziellem
PowerShell7.4.6-Parser geprüft; tatsächlicher Windows/Access-COM-Lauf ausstehend.
Kein Makrotext-Laufzeitparitätsnachweis, Katalog456/386/70 unverändert.


## 2026-11-09: natives Form.RecordsetClone

Eigene DAO-Snapshotcursor über tatsächliche Tabellen-/Abfrageformularergebnisse,
mit unabhängigem Suchen/Navigieren, Bookmark-Rücknavigation, leerem Ergebnis,
Requery-Familienwechsel und weiterlebenden gehaltenen Clones. Eigenschaft ohne
Argumente/schreibgeschützt; keine editierbaren Abfragen behauptet. Bestehende
Layout-/GUI-Prüfung gebündelt2/2 (12,40/14,29s) bestanden, kein neuer Testtarget
oder Windows-/Gesamtsuite-/Installerlauf. Details FORM_RECORDSET_CLONE_RUNTIME.md;
F054/F160 und Katalog456/386/70 bleiben offen/unverändert.


## 2026-11-09: tatsächliche Dateisammlungen und Unterordner

Native Scripting.Files/Folders unterstützen Item(Key), Standardzugriff und
indizierte Folder.Files/SubFolders-Eigenschaften über echte direkte Linux-Kinder.
SubFolders.Add(Name) erzeugt einen wirklichen Unterordner; Duplikate/Null/fehlende
Einträge/ungeeignete Pfade bleiben explizite Fehler. Native Namens-/Snapshotgrenzen
siehe NATIVE_FILE_SYSTEM_OBJECT.md; keine Windows-Gesamtparität behauptet.

Bestehender Kernlauf nach Implementierung gebündelt: Weiterleitung indizierter
Elterneigenschaften korrigiert, neue technische Unterordner vor dem vorhandenen
Symlink-Schutzfall entfernt; derselbe abschließende Kern1/1 (2,17s) bestanden.
App frisch gebaut. Kein neuer CTest, kein neuer vollständiger Desktop-/Windows-/
Qt-frei-/Installernachweis. F091 und Katalog456/386/70 bleiben offen/unverändert.
