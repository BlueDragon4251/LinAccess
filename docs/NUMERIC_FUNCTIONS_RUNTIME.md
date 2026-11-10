# Native Abs, Int, Fix and Sgn

Native VBA/Access expressions now execute Fix (truncate toward zero) and Sgn
(Integer sign result). Abs and Int now propagate Null as documented, alongside
Fix. Sgn Null rejects with 94. All four functions validate one argument (450),
resolve object scalar defaults once and use checked numeric input conversion.

Known Integer/Long/Single/Double values retain their supported numeric subtype
through Abs/Int/Fix. Abs of the signed Integer/Long minimum rejects with 6 rather
than widening silently. Currency retains exact fixed-point operations: Abs and
Int reuse checked Currency functions; Fix truncates scaled units without passing
through Double, including the lower Currency boundary. Boolean/Empty inputs use
bounded Integer conversion; numeric text uses the existing checked Double path.
Full Variant/Byte/Date/locale/formatting/promotion/error-timing identity remains
open. SQL function rewriting is not added.

Core tests verify Null behavior, negative floor versus truncation, sign/subtype
results, Currency boundary precision, arity/range/error handling and one-time
object evaluation. A second actual Qt form in the fixture-independent arithmetic
CTest calculates Abs/Int/Fix/Sgn for -3.75 and writes the native values
3.75/-4/-3/-1 through VBA into SQL. A Long Abs overflow prevents the next write;
reading calculated values leaves the record clean. Targeted tests: 2/2 (1.02 s).
Full validation is recorded in PROGRESS.md. Original business execution coverage
is unchanged.

References:
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/abs-function
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/int-fix-functions
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sgn-function


Typisierte native SQL-Normalisierung (2026-11-10)
----------------------------------------------
Abs/Int/Fix, Sgn, Round und CCur erhalten nachgewiesene Date/Boolean/Single/
Currency/integrale Argumente über den gemeinsamen Scalar-Adapter. Int/Fix/Sgn
sind zusätzliche native SQL-Einstiegspunkte. Die polymorphen Ergebnisprojektionen
von Abs/Int/Fix/Round folgen dem nachgewiesenen ersten Argument: Byte/Integer/
Long/Single und begrenzte Currency bleiben erhalten, Boolean verwendet den
bestehenden Integer-Pfad, Date den bestehenden Double-Pfad. Sgn liefert Integer,
CCur begrenzte Currency; tatsächliche DAO-Subtypen bleiben aus der Feldbindung.

Die vorhandenen nativen Berechnungs-/Overflow-/Null-Regeln bleiben unverändert.
Round über nicht-NULL Single bleibt ausdrücklich3251; unbelegte Genauigkeit wird
nicht ergänzt. Round/Currency arbeiten mit den vorhandenen Integer-Ticks, bevor
die beschränkte REAL-Ausgabe geprüft wird. Unproven ABS behält den bisherigen
SQLite-int64-Pfad; dies ist Bestandserhalt, keine Access-Int64-Literal-Parität.
Andere unproven Eingaben behalten ihre dokumentierten Transportgrenzen.
Unterstützte einfache Verschachtelungen/Klauseln bleiben typisiert; die übrigen
Date/Currency-Operator-/Quellgrenzen werden nicht entfernt. Kein neuer Windows-
Oracle oder vollständiger Variant-/Promotion-/Rundungsnachweis wird behauptet.

Die Familienprüfung deckte zusätzlich die fehlende native Date-zu-Currency-
Konvertierung auf (13). Sie ist anhand MS-VBAL5.5.1.2.3 und6.1.2.3.1.3 ergänzt:
Der standardmäßige Double-Serienwert wird durch den vorhandenen geprüften
Currency-Konverter geführt. Positive und negative exakte Bruchteile werden
nativ und über die tatsächliche SQL-Date-Bindung geprüft. Andere fehlende
Coercion-/Locale-/Variant-Grenzen werden dadurch nicht geschlossen.
