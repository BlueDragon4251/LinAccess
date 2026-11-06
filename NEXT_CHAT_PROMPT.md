Setze die Entwicklung von BlueDragon4251/LinAccess auf development/access-compatibility fort.

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

Gesicherter Ausgangsstand
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
- Nächste native Programmierung: echte Datenblatt-Seitenansicht und Anbindung der Druckbefehle an vorhandenes VBA RunCommand, dann weitere feste offene Anforderungen. AcCommand aus Microsoft Learn tatsächlich geprüft: acCmdPrint340, acCmdPrintPreview54; noch NICHT angebunden. Keine neuen Katalogpunkte. Stets tatsächlichen Git-Stand prüfen.

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
