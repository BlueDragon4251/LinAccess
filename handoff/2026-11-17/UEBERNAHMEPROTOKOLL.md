# LinAccess – Übernahmeprotokoll für lokalen Codex

Stand der Übergabe: **17. November 2026**, Europe/Vienna.

**Die Entwicklung in diesem Chat ist auf Benutzerwunsch gestoppt.** Dieses Paket übergibt den Auftrag; es meldet keinen Gesamtabschluss und keine Access-2003-Parität.

## 1. Gesicherter Ausgangspunkt

Am Übergabetag tatsächlich über die GitHub-Verbindung geprüft:

| Gegenstand | Ergebnis |
|---|---|
| Repository | `BlueDragon4251/LinAccess` |
| Sichtbarkeit | **private** – muss privat bleiben |
| Vorgesehener Branch | `development/access-compatibility` |
| Remote-Commit | `5798ec8c4662891c732e8a9fe7a9f5fb131ee6af` |
| Remote-Baum | `fe0bb25a24c75ded5a69d99eaf1a2b72aede5552` |
| Elterncommit | `a00b4ffdc48b547a65224a273ce43d8bcf68fd80` |
| Fester Katalog | **456 Punkte** |
| Markierungen im Remote-Katalog | **386 bestätigte Teilumfänge, 70 offene Vollanforderungen** |

Die 386/70 wurden für diese Übergabe direkt aus den 456 nummerierten Einträgen in `docs/FIXED_WORK_CATALOG.md` am genannten Commit gezählt. Das ist ausdrücklich eine **theoretische, ungewichtete Katalogabdeckung**, kein Gesamtfortschritt oder Kompatibilitätsprozentsatz. Die früher genannten 384/72 sind überholt.

Die neu gestartete Cloud-Umgebung enthält dagegen nur einen sauberen, **alten** Checkout auf Branch `work`, Commit `70c83efd170c6b6f402f15b09aabfd31093c2c00`. Dort fehlen unter anderem das aktuelle `NEXT_CHAT_PROMPT.md` und die früheren Arbeits-/Prüfverzeichnisse. Dieser alte Checkout ist **nicht** der Fortsetzungsstand.

`remote-specs/` enthält vollständige Textkopien der sechs verpflichtenden Dokumente vom oben festgehaltenen Remote-Commit. Diese Kopien sind Referenzmaterial; nicht über einen neueren lokalen Checkout kopieren.

## 2. Kritisch: nicht gesicherte Cloud-Arbeit

Nach `5798ec8` wurde eine größere MONEY-/Currency-Etappe entwickelt, lokal committed und mehrfach korrigiert. **Sie wurde nicht auf den Remote-Branch gepusht.**

Letzter im vorherigen Arbeitsraum beobachteter lokaler Commit:

- Commit: `efe72551ae94d3cd836e0ac237fbe8ab78b3879a`
- Baum: `66ebdf8b4a4299b2b68166203a14608dd3ea00b4`
- Titel: `Import proven Jet4 and ACE2010 MONEY with exact Currency transport`
- Elternstand war der veröffentlichte `5798ec8`.

Dieser Commit, seine vollständigen Quelldateien, die Prüflogs und eine getrennt vorbereitete Decimal-Arbeitskopie sind in der neu gestarteten Umgebung **nicht vorhanden**. Das Übergabepaket enthält deshalb **keinen vollständigen Patch, kein Git-Bundle und kein gerettetes Build-Artefakt** dieser Etappe. Ein Commitname im Protokoll bedeutet nicht, dass das Objekt über GitHub abrufbar ist.

Lokal zuerst nach einer tatsächlich vorhandenen alten Arbeitskopie, einem Bundle oder Export suchen. Falls nichts vorliegt, die unten beschriebene Etappe vom sicheren Remote-Stand aus nachvollziehbar neu implementieren. Niemals behaupten, die verlorenen Änderungen seien bereits übernommen.

## 3. Produktauftrag und verbindliche Regeln

LinAccess soll ein **allgemeiner Linux-Ersatz für Microsoft Access** werden: vorhandene Datenbanken nutzen und eigene Datenbanken, Tabellen, Abfragen, Formulare und Berichte erstellen. Die Kassa- und Rezeptdatenbanken sind Referenzen, keine Beschränkung des Produkts.

- C++20, Qt 6, Linux; Debian 12/13 und Ubuntu 24.04+. Windows/Access bei Bedarf als tatsächliche Referenzumgebung verwenden.
- Access 2003/JET4 zuerst; ACE/ACCDB im dokumentierten Umfang. Keine pauschale vollständige Kompatibilität behaupten.
- `CODEX_HANDOFF.md` vollständig lesen, danach `NEXT_CHAT_PROMPT.md` vollständig; dann README, Kompatibilitätsdokumentation, Entwicklungscheckliste und festen Katalog vollständig lesen.
- Benutzerwunsch: Programmierung priorisieren, einen zusammenhängenden Umfang fertigstellen, anschließend bestehende gezielte Prüfungen gebündelt ausführen. Bei Fehlern das konkrete Problem beheben und betroffene Prüfungen wiederholen. Keine Testflut und keine neuen Testziele für jedes Detail.
- Während der Entwicklung kontinuierlich weiterarbeiten, keine freiwilligen Zwischenabschlüsse und keine Frage „Soll ich weitermachen?“.
- Regelmäßige deutsche Statusmeldungen mit Segmentbalken und ausdrücklich theoretischer, ungewichteter Katalogzählung. Keine erfundenen Gesamtprozentsätze.
- Katalog auf 456 IDs festhalten. Neue Details bestehenden Anforderungen zuordnen; offene Vollanforderungen nicht wegen eines weiteren Teilumfangs schließen.
- Auf Testkopien arbeiten. Original-MDBs, Bridge-Archive und private VBA-Quellen schützen. Keine privaten SQL-Texte, VBA-Literale, Zugangsdaten oder Geschäftsdaten in öffentliche Logs, Issues oder Releases übertragen.
- Keine Mockups, erfundenen Geschäftsdaten oder Ersatzdaten zur Herstellung scheinbar grüner Originalprüfungen. Technische Grenzwert-Fixtures sind zulässig.
- Nach geprüften Etappen sauber committen und auf den vorgesehenen **privaten** Branch pushen. Keine lokalen Benutzeränderungen überschreiben, kein Force-Push, kein unaufgefordertes destruktives Reset.

## 4. Was am veröffentlichten Stand bereits vorhanden ist

Der Remote enthält zahlreiche vorherige Implementierungen; diese nicht anhand der alten Cloud-Kopie erneut beginnen. Zu den zuletzt veröffentlichten Etappen gehören:

| Commitanfang | Umfang |
|---|---|
| `4ee2daf` | eingangstypabhängige Currency-SQL-Ziele |
| `4ce1758` | exaktes Tick-Datenblatt, Bearbeitung und Clipboard/History |
| `dbd87ce` | exakte Currency-Selektoraggregate |
| `53ff703` | passende exakte Decimal-Literaldefaults |
| `6a8da87` | automatische Primärschlüsselwahl beim Tabellenanlegen |
| `fad7b1a` | belegte ganzzahlige Currency-Operanden |
| `a00b4ff` | Currency-Typen in bedingten SQL-Ergebnissen |
| `5798ec8` | begrenztes typisiertes UNION/UNION ALL |

Tabellen-/Abfragedruck und PDF waren vor der letzten MONEY-Arbeit bereits bearbeitet worden. Den aktuellen Umfang im Repository prüfen; die ursprüngliche frühe Planung „Druck als Nächstes“ nicht ungeprüft übernehmen.

Bestehender Decimal-Unterbau: exakte native Decimal-Werte, sortierbare 24-Byte-Schlüssel, belegte physische Präzision/Skala sowie Jet-/ACE-Importteilumfänge. Bestehende REAL-Projekte und neue exakte Speicherkontrakte nicht stillschweigend vermischen.

## 5. Rekonstruktionsprotokoll der nicht gepushten MONEY-Etappe

Die folgenden Punkte beschreiben den früheren Arbeitsstand, **nicht im Remote nachgewiesenen oder hier mitgelieferten Code**.

### 5.1 Rohimport und Speichervertrag

- Neue, durch Quellformat und Spaltentyp belegte Jet4-/ACE2010-MONEY-Importe sollten vollständig auf signed64-`CURRENCY_TICKS` umgestellt werden, nicht anhand ausgewählter Zellwerte.
- MDBTools-Helfer um `--currency-values FILE TABLE` erweitern: rohe acht Byte pro MONEY-Zelle, Little Endian, Hextransport, NULL, Spaltennamen, Quellformat und SHA-256; Größen-/Zeilen-/Zellgrenzen prüfen.
- Schemahash, Rohquellenhash und tatsächlichen Quelldateihash abgleichen. CSV-Wert, Rohwert, NULL und Zeilenzahl exakt vergleichen, ohne MONEY durch Double zu führen.
- SQLite-Werte als Integer-Ticks mit entsprechendem CHECK speichern. Passende vierstellige Literaldefaults in Ticks umsetzen; Original-Defaultmetadaten erhalten.
- Bestehende REAL-Projekte und reine CSV-Aufrufer behalten ihren bisherigen Vertrag. Andere Formate, beliebige Defaults und Originalformat-Schreiben nicht als erledigt ausgeben.
- Unabhängige technische Jackcess-Fixtures für V2000 und V2010: MAX/MIN signed64-Ticks, `9007199254740993` Ticks, ±1 Tick, 0, NULL, Default 1,25; nach Erstellung unabhängig erneut öffnen. Keine MDB/JAR/Class-Dateien oder privaten Quellen neu committen.

### 5.2 Durchgängiger Currency-Transport

- DAO Type 5/Size 8, Datenblatt, Formulare, Defaults, SQL-Ziele, feste Kreuztabellen und begrenztes SELECT INTO mussten den Tickvertrag behalten.
- Belegte Joins/Aliase und gespeicherte Abhängigkeiten typisieren; unbekannte oder nicht unterstützte Exact-Kontexte explizit abweisen, bevor rohe Ticks als Dollarwerte gerechnet werden.
- Typnachweise an tatsächliche View-SQL, Tabellen-DDL, Quellmetadaten und Abhängigkeiten binden. Veränderungen und neue temporäre Schattenquellen müssen erkannt werden; Requery muss tatsächlich neu aufbauen.
- View-Installationscache nicht nur auf Vorhandensein prüfen: auch tatsächlich installierte Definitionen abgleichen.
- Bei rangbegrenzten Views nicht nachträglich den ursprünglichen Jet-Text erneut mit anderer Ursprungszuordnung übersetzen. Den tatsächlich kompilierten Plan und LIMIT-0-Ausgabespalten verwenden. Unbelegte Exact-Kontexte erhalten keinen erfundenen Nachweis. Installations-Savepoint auch bei Ausnahmen zurückrollen.
- Reine Projektionsalias-Erweiterung um die bekannten seiteneffektfreien Currency-Rechen-/Summenadapter und `linaccess_jet_field_store` ergänzen. Keine beliebigen VBA-Aufrufe dadurch mehrfach auswerten.
- Ganzzahlige bekannte Werteverzweigungen passend in Currency umsetzen; Gleitkommamischungen als Double ausdrücklich kenntlich machen. Kein neuer Windows-Jet-Promotion-Oracle behauptet.
- Syntaktische Literalbindung nur bei vollständig bekannten Werteverzweigungen: die bestehende rohe int64-Verzweigung `IIf(True,9007199254740993,0)` darf nicht durch partielle Integer4-Bindung kaputtgehen.

### 5.3 VBA-SQL-Aufrufe und Typnachweise

- Currency-Argumente wurden über einen kurzlebigen, subtypmarkierten acht-Byte-BLOB transportiert (interner Subtyp `0xc5`). Normale gespeicherte BLOBs dürfen dadurch nicht nachträglich als Currency gelten.
- Deklarierte Currency-SQL-Ergebnisse als Ticks nur über einen kontrollierten Compileradapter ausgeben. Öffentliche, unaufgelöste Aufrufe dürfen keine rohen Ticks an beliebige SQLite-Arithmetik liefern.
- Im früheren Stand wurden reservierte Aufrufe `linaccess_vba_currency_<name>` mit Registrierung für typisierte Currency-Aufrufe und einer aktuellen Deklarationsprüfung verwendet; `linaccess_vba_result` prüfte den Rückgabevertrag. Änderungen der Funktionssignatur mussten vor falscher Interpretation abgewiesen werden.
- Das bloße Vorhandensein einer VBA-Funktionsregistry ist **kein** Typnachweis für eine physische Quelle oder Unterabfrage.
- Geändertes SQL eines Plans verliert den alten Kontext-/Ergebnistypnachweis.
- Unaufgelöste CASE-/Unterabfragekontexte mit exakten Werten nicht stillschweigend laufen lassen. Access-Jet-Unterstützung und SQLite-Erweiterungen unterscheiden.

### 5.4 Quotierte Quellen und GUI

- Nachgewiesene Lücke: Bei einer technischen Tickspalte mit Wert 1 ergab `amount+1` aus einer quotierten Tabelle zuvor rohe 2 statt 10001 Ticks. String-/Doppelquote-Relationen müssen in Quellenposition erkannt und typisiert werden.
- FROM/JOIN, Quellen-Kommalisten und geprüfte UPDATE-/INTO-Ziele positionsabhängig behandeln; normale Stringliterale und Funktionsargumente nicht in Bezeichner umwandeln. Auch Exact-Abhängigkeitsguards müssen diese Quellen erkennen.
- Nicht unterstützte Unterabfragen müssen auch bei vorhandenen VBA-Funktionen geschlossen bleiben.
- Formular-/Report-RecordSource-Views und Tabellen-RowSources nicht als rohe `plan.sql` ohne Quellnachweis aufbauen, sondern durch den Jet-Compiler führen. Bezeichner passend mit Jet-Klammern escapen.
- Combo-/List-Spalten müssen belegte RowSource-Typen mitführen. `Column`, Auswahlwert, Anzeige, Zuweisung und Requery dürfen Tickwerte nicht als gewöhnliche int64-Variants oder Double lesen.
- Geprüfter technischer Fall war MAX-1 Ticks in gespeichertem Query-Formular und ComboBox, inklusive `Column`, Currency-Zuweisung und Requery.

### 5.5 Originalprüfungen richtig weiterführen

- Der Inventarprüfer baute für einen ungefilterten Vergleich noch einen rohen nativen SQL-Plan. Stattdessen den ursprünglichen Jet-SELECT ohne Parameterfilter erneut kompilieren; keine Typguards umgehen.
- Inventarbericht vor den abschließenden Zählasserts schreiben und flushen, damit ein Fehler auswertbar bleibt. Nur technische Kategorien/Hashes ausgeben, keine privaten SQL-Inhalte.
- Die unabhängige ursprüngliche DSum-Prüfung summierte zuletzt rohe Ticks als Dollar-Double. Korrektur: bei belegten Tickspalten unabhängig mit geprüften signed64-Überlaufgrenzen summieren und Currency-Einheiten exakt vergleichen. Nicht den Produkt-Summenadapter als eigene Referenz wiederverwenden.
- Vor einem erneuten langen Originallauf zuerst das konkrete Problem mit bestehenden gezielten Prüfungen bzw. einer kurzen Diagnose auf einer geschützten Arbeitskopie eingrenzen.

### 5.6 Betroffene Dateien im verlorenen Arbeitsstand

`src/access_catalog.{cpp,hpp}`, `access_floating_metadata.{cpp,hpp}`, `access_query.cpp`, `access_table_import.cpp`, `access_table_model.cpp`, `dao_database.cpp`, `jet_sql_functions.cpp`, `mdb_metadata_main.cpp`, `native_vba_sql.hpp`, `vba_qt_value.{cpp,hpp}`, `access_form_preview.cpp` sowie bestehende Import-, Currency-, DAO-, Inventar- und Berichtprüfer. Neu vorgesehen waren `docs/NATIVE_MONEY_SOURCE_IMPORT.md` und `tests/jet_binary_reference/JetCurrencyReference.java`.

Die Namen sind Orientierung für die lokale Suche/Rekonstruktion, kein Nachweis vorhandener Dateien oder übertragener Änderungen.

## 6. Letztbeobachtete Prüfergebnisse – keine neue Verifikation

Diese Ergebnisse stammen aus dem vorherigen Arbeitsraum. Seine Logs sind jetzt nicht verfügbar.

- Für den veröffentlichten `5798ec8` war ein Paket-/Originallauf mit 2/2 in 186,89s und sieben unveränderten Schutzhashes protokolliert worden.
- MONEY-Etappe: gezielte Schema-/Parameter-/Import- und Formularprüfungen wurden nach konkreten Fehlern erfolgreich wiederholt. Zuletzt wurden Parameter 0,29s, Import 1,07s, der gezielte Formularlauf 0,68s und nach der int64-IIf-Korrektur DAO 7,71s beobachtet.
- Vier zuvor blockierte Originalquellen ließen sich im gemeinsamen direkten Lauf auf einer schreibgeschützten Arbeitskopie ausführen. Das war kein Access-Wertvergleich.
- Letzter Paketlauf: `/workspace/work/money-source-package-v4/`, Source `efe7255…`, Baum `66ebdf8…`; Build, DEB-Erzeugung und Payloadcheck meldeten Erfolg.
- Sein `kassa_query_inventory` bestand zuletzt in **87,49s**. Die Klassifikation blieb 73 gewöhnlich + fünf parameterisiert ausführbar; 3 ursprüngliche fehlende Spalten, 9 fehlende Abhängigkeiten und 10 andere explizit ungelöste Fälle.
- **`real_kassa_dao_event` lief zuletzt noch. Sein Endergebnis und die abschließenden sieben Hashprüfungen liegen für diese Übergabe nicht vor.** Die MONEY-Etappe war daher nicht abschließend verifiziert und nicht veröffentlicht.

Die alten `/workspace/work/...`-Pfade sind historische Fundstellen, keine hier nutzbaren Dateien. Ergebnisse nicht allein aufgrund dieser Liste grün markieren.

## 7. Danach vorbereitete, ebenfalls nicht übertragene Etappe

Eine getrennte Arbeitskopie unter `/workspace/work/decimal-maketable-source` war für **Decimal-Felder in SELECT INTO** vorbereitet. Kein getesteter oder gepushter Abschluss.

Rekonstruktionsauftrag:

1. Für direkt projizierte physische Decimal-Felder den tatsächlichen Ursprung aus `accessQueryColumns(..., fieldOrigins)` verwenden.
2. Präzision/Skala über den bestehenden geprüften Quellvertrag `nativeTableDecimalTypes` holen, nicht aus Zellwerten oder frei geratenen Ausdrücken.
3. Ziel vor dem Einfügen explizit als `DECIMAL_EXACT` mit passendem `linaccess_jet_decimal_field_valid`-CHECK deklarieren und Metadaten speichern.
4. Tabelle, Metadaten und INSERT unter einem gemeinsamen Schema-Savepoint atomar behandeln.
5. Umbenannte direkte Felder und NULL erhalten; anschließende geprüfte Schreiboperationen müssen die Skala weiter erzwingen. DAO Type 20 prüfen.
6. Berechnete/zusammengesetzte Ergebnisse ohne belegte physische Präzision/Skala abweisen. Kein allgemeines Übertragen von Counter, Indizes, Required, Defaults oder Lookups behaupten.
7. Vorhandene Import-/DAO-Prüfer gebündelt nutzen: technische tatsächliche Jet-/ACE-Numeric-Dateien mit (28,4)/(28,28), große Werte, NULL, Alias, anschließende Skala-Verweigerung und kein Ziel bei unbelegtem Ausdruck.

F014/F055/F067/F130 bleiben damit weiterhin breite offene Vollanforderungen. Neue Dokumentation war als `docs/DECIMAL_SELECT_INTO.md` vorgesehen.

## 8. Referenzen, Daten und bekannte Grenzen

Vom Benutzer im Chat bereitgestellt wurden unter anderem:

- `Keksrezepte.mdb`
- `Kassa_CSE_V2.0.mdb`
- `Kassa_CSE_V2.0-Bridge-2003-Retry.zip`
- `Ergebnisse.zip`
- `LinAccess-Referenz-20261108-175907-2a00ae88.zip`
- `LinAccess-Referenz-20261108-180556-31981c90.zip`
- `LinAccess-Gesamtreferenz-20261108-181858-f70a7afd.zip`

Diese Anhänge werden durch dieses Übergabepaket **nicht** erneut mitgeliefert und sind auf dem lokalen Rechner nicht automatisch vorhanden. Vorhandene private lokale Kopien bzw. die ursprünglichen Chat-Downloads nutzen, nicht eine Verfügbarkeit behaupten.

Die Rezept-/Kassa-Datenbanken allein ersetzen keinen Access-Ausführungsoracle. Der v4-Referenzlauf meldete `partial_modules_failed`: nutzbare Ausdrucks-/Counterbeobachtungen, aber fehlgeschlagene SQL-/DAO-Einrichtung. Früher trat außerdem ein Access-Interop-`CreateControl`-Absturz auf. Kein pauschales „alle Referenzen vollständig“ daraus ableiten. Ein SQL-/DAO-Referenznachlauf v5 war noch nicht als erfolgreicher vollständiger Oracle belegt.

Die hochgeladene Kassa-Datei und die Git-Fixture waren nicht identisch:

- Benutzerdatei: `d9ff0f9d002b3ab59b3336a34e6c09d9c058d5b9b944b754259a78e277e77102`
- Git-Fixture: `fd3cc0c4613348152d6d8f694bcd51f0fe941129180dbbcffcb42fda4f2a3f5b`

Keine Datei allein wegen des gleichen Dateinamens ersetzen. Bridge-Quellen: 47 Formulare, 9 Berichte, 100 Abfragen, 32 VBA-Dateien, vier Standardmodule. Vier ursprüngliche MDB-Makrodefinitionen waren weiterhin nicht durch den Bridge-Export belegt. Original-Kassa-Gesamtbetrieb bleibt offen.

## 9. Lokale Umgebung und Prüfdisziplin

- Die früheren Cloud-Sysroots, Runtimepfade, Generatordateien und Logordner sind nicht verfügbar. Keine alten absoluten Pfade ungeprüft übernehmen.
- Lokale Linux-/WSL2-Umgebung anhand des aktuellen Repositorys einrichten; CMake/README/CI als Quelle für Abhängigkeiten und SQLite-/MDBTools-Vorgaben nutzen.
- Historisch verwendet: GCC 14, Qt 6.8, private SQLite 3.54.1, libmdb 1.0.1 und MDBTools-CLI 1.0.0. Das ersetzt keine Prüfung der tatsächlich lokal verfügbaren Versionen.
- Relevante bestehende Ziele: `table_import_contracts`, `dao_execution`, `currency_query_parameters`, `currency_schema_queries`, `currency_field_documents`, `access_query_model`, `kassa_query_inventory`, `real_kassa_dao_event`. Tatsächliche Verfügbarkeit per CMake/CTest prüfen.
- Keine Source-/Testdateien während laufendem Build/Check verändern; keine Paketdokumentation während CPack ändern. Keine parallelen CPU-Builds während GUI-/Originalprüfungen.
- Private Originallogs nicht roh ausgeben. Technische Fehlernummern, Prüferdatei/-zeile, Statuszählungen und Hashes verwenden.

## 10. Codex Call und Abschluss

Der Benutzer hat für wirklich erforderliche Entscheidungen und den tatsächlichen Gesamtabschluss Codex Call autorisiert:

- Repository `BlueDragon4251/codex-call`, vorgesehener Stand `feat/ask-via-call`, früher verifiziert einschließlich `fd2cea7`.
- `.agents/skills/codex-call/SKILL.md`
- `.agents/skills/codex-call/scripts/ask_via_call.py`
- `.agents/skills/codex-call/scripts/notify_completion.py`

Vor Nutzung tatsächliche Skill, Schnittstellen, Parameter und Authentifizierung prüfen. Keine Endpunkte, Parameter oder Anrufergebnisse erfinden. Keine Infrastruktur neu einrichten, solange die vorhandene funktioniert; keine Geheimnisse loggen oder committen.

Notwendige zusammenhängende Fragen deutsch bündeln und automatisch anrufen. Keine Routine-/Testanrufe. Nicht beantwortete Fragen nicht als Zustimmung werten; unabhängige Arbeit fortsetzen, nicht endlos erneut anrufen. Frühere SQL-/DAO- und Makro-Rückfragen blieben unbeantwortet.

Diese Übergabe ist **kein Gesamtabschluss**. Deshalb kein Fertiganruf für angeblich abgeschlossene Access-Parität. Erst nach tatsächlicher Erfüllung und Nachweis des Gesamtauftrags entsprechend informieren; verbleibende Grenzen ehrlich nennen.
