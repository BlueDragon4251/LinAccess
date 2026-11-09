# Projektbezogene Einstufung und Ausführungsbeobachtungen

Unter Extras → Kompatibilität und Ausführungsbeobachtungen lässt sich der aktuelle
Projektkatalog untersuchen und explizit als JSON speichern. Das Lesen führt keine
Geschäftsprozeduren, SQL-Aktionen oder Property-Getter aus. Die Ergebnisdatei
enthält keine Modulquellen, Argumentwerte oder Datensatzwerte. Objektnamen und
technische Metadaten sind dennoch private Projektinformationen.

Die Einstufung trennt gespeicherten Importstatus, statische Fähigkeit und
beobachteten Lauf. Bei VBA werden die tatsächlich inventarisierten Prozeduren
mit dem vorhandenen Interpreterparser geprüft. „syntax_accepted_runtime_unproven“
bedeutet ausschließlich Parserannahme, keine erfolgreiche Ausführung. Nicht
verfügbare, fehlerhafte oder zu große Definitionen werden ausdrücklich benannt.
SQL-Kompilierung ist ebenfalls kein Ergebnis- oder Schreibnachweis. Für Dokumente
und Makros bleibt notwendiger Laufzeitkontext sichtbar; vorhandene Tabellenschemata
beweisen keine vollständige Datentyp-/Daten-/Schreibkompatibilität.

Die laufende Sitzung beobachtet tatsächlich versuchte VBA-Aufrufe aus Standardmodulen,
importierten Modulansichten, Dokumenten und SQL-Funktionen einschließlich
verschachtelter Aufrufe, Fehlernummern und Abbruch. Beobachtungen enthalten den
SHA-256 der tatsächlichen Quellrevision, Prozedurnamen, Ausführungskontext,
Zeitpunkt, Ergebnisstatus und Fehlerzeile/-nummer. Diagnose-Ausdrücke werden
nicht exportiert, damit private Quellfragmente nicht in den Bericht geraten. Es werden höchstens1024
Beobachtungen behalten; fehlende Beobachtung bedeutet ausdrücklich keinen
Ausführungsnachweis. Projektwechsel setzt diese Sitzung zurück. Beobachtung
schreibt keine zusätzliche SQL-Transaktion und beeinflusst keine VBA-Fehlerhandler.

Identische Quellen in mehreren Objekten werden nicht als bewiesene Objektinstanz
behandelt: Eine Übereinstimmung von Quelle und Prozedur heißt ausdrücklich
„matching_source_observed“. Geänderte Quellrevisionen übernehmen frühere
Laufergebnisse nicht. Eine erfolgreiche Prozedur beweist nur den tatsächlich
durchlaufenen Kontext und Pfad, keine allgemeine Access-Parität.

Native Makrodefinitionen werden zusätzlich ohne Ausführung mit dem vorhandenen
Aktionsparser geprüft. Tatsächlich gestartete native Makros protokollieren
Beenden/StopMacro/StopAllMacros und Fehler für die zu Beginn gelesene Revision.
Der SELECT-Ausführungspfad protokolliert echte Ergebnisse oder Ausführungsfehler
für die verwendete SQL-Quelle; Parameterwerte und Ergebnisdaten fehlen bewusst.
Dies ist kein vollständiger SQL-/Tabellenzugriffs-Audit: andere SQL-Pfade, direkte
Treiber-/Tabellenmodellzugriffe und Rendering eines gesamten Dokuments gelten
weiterhin als nicht beobachtet. Fehlende Instrumentierung wird niemals als
Erfolg eingestuft. Ein SQL-Rollback entfernt Datenänderungen, aber nicht die
Beobachtung, dass der Aufruf tatsächlich versucht wurde.
