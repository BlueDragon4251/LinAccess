# Reale Access-2003-Datenbank: Windows-Bridge-Export und VBA-Laufzeittest

Stand: LinAccess Native 0.6.0

## Testquelle

- Original: vom Anwender bereitgestellte `Keksrezepte.mdb`, Jet 4 (Access 2003).
- Export: vom Anwender unter **Microsoft Access 365** erzeugter Ordner
  `Keksrezepte-Bridge` via `Export-LinAccessBridge.ps1`.
- Eingabe für LinAccess: gezippter Exportordner mit `manifest.json`,
  Quelltextartefakten und SHA-256-Prüfwerten.
- Die originale MDB und der persönliche Export werden nicht in der
  veröffentlichten Quellcode-/Debian-Paketversion mitgeliefert.

## Verifizierte Daten

| Art | Anzahl | Zustand |
|---|---:|---|
| Formulare | 4 | `SaveAsText`-Originalquelle in Bridge importiert |
| Berichte | 1 | Originalquelle importiert, keine native Layoutkonvertierung |
| Abfragen | 19 | Original-Jet-SQL/Abfragetexte importiert, nicht automatisch ausgeführt |
| VBA-Klassenmodule | 2 | Quellbytes erhalten, statisch analysiert |
| **Gesamt** | **26** | Alle Manifestobjekte `source_only`, keine Exportfehler gemeldet |

Im Qt-Offscreen-Regressionstest wird die Bridge-GUI mit echten Artefakten
aufgerufen, der Import in eine temporäre SQLite-Datenbank abgeschlossen,
der Objektkatalog gezählt sowie der gespeicherte Quellcode geprüft.

Insgesamt **8 VBA-Ereignisprozeduren**:

- Modul `Form_3_frm_Rezepte`: 6 Prozeduren. Davon enthält nur
  `Form_Open(Cancel As Integer)` ausführbare Anweisungen, nämlich
  `DoCmd.Maximize` und `DoCmd.GoToRecord , , acLast`.
- Modul `Form_33_Ufo_frm_ZutatenRezept`: 2 Prozeduren. `Menge__g__AfterUpdate`
  enthält Access-Formreferenzen, `SetFocus`, `DoCmd.Close`, `DoCmd.OpenForm`,
  `DoCmd.GoToRecord` und eine formulargebundene Feldzuweisung.
- Die übrigen 6 Ereignisprozeduren enthalten nur Kommentare.

**Tatsächlich ausgeführt:** Der Test extrahiert den unveränderten Text des
echten Moduls `Form_3_frm_Rezepte`, ruft darin `Form_Open` mit dem eigenen
C++-Interpreter auf und prüft, dass er die beiden `DoCmd`-Hostaktionen je
genau einmal ausführt und das Ereignis nicht abbricht. Die reale `Form_Open`-
Prozedur ist daher nicht nur erkannt, sondern innerhalb der dokumentierten
Teilmenge ausgeführt.

**Nicht ausführbar:** `Menge__g__AfterUpdate` benutzt das Access-Forms-
Objektmodell samt geöffnetem anderen Formular, Feldreferenzen und Fokus.
Diese Objekte sind in der derzeitigen nativen Qt-Laufzeit nicht implementiert;
die Prozedur wird **abgelehnt**. Dies ist ausdrücklich **keine vollständige
Access-2003-/VBA-Kompatibilität**.

## Formular-Teilkonvertierung

Der Parser liest aus dem echten Export von `3_frm_Rezepte` die Datenquelle
`3_tbl_Rezepte`, gebundene Feldnamen (u.a. `Rezeptname`) und das Ereignis
`OnOpen`. Das Unterformular `33_Ufo_frm_ZutatenRezept` referenziert die
Abfrage `33_qry_ZutatenRezept`.

Nach separatem Tabellenimport kann in LinAccess der Quellformular-Eintrag
geöffnet und die Funktion **Als natives Formular übernehmen
(Teilkonvertierung)** verwendet werden. Dabei werden nur Felder übernommen,
die in der tatsächlich importierten SQLite-Tabelle existieren; das originale
VBA-Klassenmodul wird dem nativen Formular zugeordnet. Der Import überschreibt
kein bereits bestehendes natives Formular und verändert niemals die MDB.

Die importierte MDB enthält nur 2 physische Tabellen mit zusammen 39
Datensätzen; verschiedene Feld-/Tabellenreferenzen im Access-Formular und in
Access-Abfragen sind nicht darin enthalten (u.a. `2_tbl_Nährwerte`). Deshalb
ist auch mit einem vollständigen VBA-Interpreter eine unveränderte
Ausführung dieses Projekts **nicht garantiert**: Es fehlen Datenabhängigkeiten.

## Lokaler Reproduktionstest

```bash
cmake -S . -B build -DCMAKE_BUILD_TYPE=Release \
  -DLINACCESS_REAL_MDB_FILE=/absoluter/pfad/Keksrezepte.mdb \
  -DLINACCESS_REAL_BRIDGE_DIR=/absoluter/pfad/Keksrezepte-Bridge
cmake --build build -j2
ctest --test-dir build --output-on-failure
```

Der `real_access_bridge`-Test wird nur registriert, wenn der angegebene
Bridge-Ordner ein `manifest.json` enthält. Keine Nutzerdateien müssen dazu
in das Repository eingecheckt werden.
