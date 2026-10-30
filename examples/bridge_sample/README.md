# Synthetisches Bridge-Testpaket

Dieses Paket stammt **nicht aus einer echten Access-Datenbank**. Es dient nur
zum Testen des Linux-Dialogs „Bridge-Export öffnen“, seiner JSON-Verarbeitung,
Hashprüfung und Speicherung von Quellbytes.

Nach dem Öffnen eines neuen LinAccess-Projekts diesen Ordner auswählen. Er enthält
ein VBA-Testmodul und einen SQL-Text. Beides wird als **source_only** gespeichert,
nicht ausgeführt. Die Fingerabdrücke in `manifest.json` dürfen nicht verändert
werden, ohne die Artefakte neu zu hashen.
