# Reale Access-2003-Regressionsdaten

Dieses Verzeichnis enthält die vom Projektauftraggeber bereitgestellten beiden
Microsoft-Access-2003-`.mdb`-Dateien und die zugehörigen Windows-Bridge-Exporte
**verlustfrei** als zusammengesetztes, Base64-kodiertes `tar.xz`-Archiv.
Die Aufteilung in Teilstücke dient ausschließlich der Übertragung ins private
GitHub-Repository mit der verwendeten GitHub-Schnittstelle.

**Diese Dateien können personenbezogene Daten, Geschäftsdetails und VBA-Code
mit hinterlegten Passwörtern enthalten. Das Repository muss privat bleiben.**
Keines der Archive gehört in öffentliche Releases oder öffentliche CI-Artefakte.

Wiederherstellen (Debian/Ubuntu):

```bash
./scripts/restore-fixtures.sh
```

Wiederhergestellte Originaldateien:

- `fixtures/raw/Keksrezepte.mdb`
- `fixtures/raw/Kassa_CSE_V2.0.mdb`
- `fixtures/raw/Keksrezepte-Bridge.zip`
- `fixtures/raw/Kassa_CSE_V2.0-Bridge-2003.zip`

Die Prüfung `sha256sum -c SHA256SUMS` erfolgt automatisch. Anschließend
stehen die entpackten Bridge-Verzeichnisse unter `fixtures/unpacked/` bereit.
Die Originale werden bei normalem App-Betrieb nie direkt verändert; Tests und
VBA-Aktionen verwenden eine eigene Datenbankkopie.

Reale Objektreferenzwerte aus den bisherigen Tests:

| Datenbank | Tabellen | Bridge-Objekte | Formulare | Berichte | Abfragen | VBA-Dateien |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Keksrezepte | 2 | 26 | 4 | 1 | 19 (inkl. Systemabfragen) | 2 |
| Kassa CSE V2.0 | 29 | 192 | 47 | 9 | 100 | 32 |

`manifest.json` in jedem Bridge-Paket nennt fehlgeschlagene oder nicht
verfügbare Exporte ausdrücklich. Ein erfolgreicher Quellenimport beweist
**noch keine vollständige VBA-Ausführung**.
