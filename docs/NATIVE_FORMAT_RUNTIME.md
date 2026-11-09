# Native number, currency and compound date formats

Format/Format$ now share a native date/numeric formatter between VBA desktop
hosts and SQL. Date masks combine yyyy/yy/y, m/mm/mmm/mmmm, d/dd/ddd/dddd,
h/hh, n/nn, s/ss, quarters, Sunday-based weekday/week fields, AM/PM and A/P.
Quoted date literals and backslash escapes are parsed; an mm immediately after
h/hh denotes minutes. Long/Short Date and Long/Short Time use the current native
locale/date policy. ISO text and bounded OA serials keep wall-clock date/time
components separate, including DST-gap times and the integral day of negative
OA dates. This is not a full localized calendar, user date-separator, millisecond,
custom week-rule or arbitrary Access mask implementation.

Numeric masks support0/# integer/fraction placeholders, optional trailing
fraction digits, three-digit grouping and percentage suffixes. Rounding uses the
native Decimal kernel; large integers remain exact. Fixed/Standard/Currency named
formats and English/German Yes/No, True/False and On/Off are supported. Numeric
sections/colors, scientific notation, scaling commas, numeric quoted suffixes and
string @/&/case masks refuse rather than being silently interpreted. Values and
precision remain within the bounded Decimal domain (up to28fraction places).
Full IEEE/extreme-Double formatting and original field Format properties remain
separate open work; these function paths do not infer missing source properties.

FormatNumber and FormatCurrency accept their five standard VBA arguments,
reordered names, omitted optional positions and vbTrue/vbFalse/vbUseDefault.
NumDigitsAfterDecimal -1 uses the default; TriState -2 uses the default. Invalid
flags/precision report5 and Null numeric expressions94. Defaults use the current
QLocale profile. Currency precision/symbol/placement come from its actual positive
currency pattern; legacy currency-affix nonbreaking spaces become ordinary spaces
while numeric grouping separators remain unchanged. Other custom Windows regional
settings, currency negative patterns and independent UI-language settings are not
claimed. Current desktop IsNumeric recognizes text through that native locale;
headless hosts retain their invariant parser unless they provide a locale policy.

SQL uses the same formatter, including explicit positional FormatNumber/Currency.
SQLite INTEGER arguments are converted from their actual signed integer text,
not Double, preserving9007199254740993 and other supported large values. Locale
format callbacks are not marked deterministic. The pure core host interface adds
no Qt dependency. A host without a numeric-format policy refuses3251. Unbound
control/bound SQL value conversion and query-source updateability retain their
existing independent guards. Access Eval's observed Format(Null,mask) empty
String behavior is separate from ordinary/SQL Variant Format's retained Null
behavior; Format$ retains its strict Null94 rule.

The real Access2003 archive has104 observed expressions and8 recoverable errors.
The existing DAO target checks56 numeric values/types plus8 errors. The window
target now checks the remaining48 string/date/locale observations through the real
native application/form Eval host under the recorded de-AT locale, including
Unicode comparison, dates, boolean formats and exact separators/spacing. Calendar
expectations use the independent Qt day/second representation of the recorded
Windows ISO timestamps; binary Double simple cases retain their recorded bits.
No original MDB/VBA, business record or fabricated expectation is included.
Additional existing formatter checks cover real SQL, exact large integers, masks,
flags and escaping. No CTest target/catalog point is added. Passing these finite
cases does not establish complete Eval/SQL/UI/locale equivalence; full
F014/F041/F068/F071/F118/F164 and fixed456/384/72 remain open.

Validation: core/DAO/window/formatter4/4 (9.49s), application built. The full
finite104-value/8-error reference suite passed in the stated split contexts.
No fresh GitHub CI, distro matrix, original workflow or full catalog parity claim.


Custom numeric masks now support up to three positive/negative/zero sections,
quoted and escaped literal prefixes/suffixes (including escaped semicolons),
explicit negative parentheses, percent prefixes/suffixes and exact trailing-comma
thousand scaling. Empty negative/zero sections reuse the first section. Explicit
negative sections format the absolute magnitude; fallback automatic minus signs
precede the complete output. Decimal/Currency values retain their exact text.
Scaling checks an exact inverse before rendering, refuses a lossy scale with3251
and is bounded to eight comma factors. Leading fractional masks such as .00
remain supported. Unique sections are chosen from the unrounded input sign.

Scientific notation, conditional/color directives, fourth Null sections,
interleaved digit/literal masks and full Windows regional/custom-mask equivalence
remain explicitly unsupported. Existing Format/Format$ Null rules are preserved.
The code is shared by native VBA/SQL/form formatting; technical regression cases
are implementation checks, not newly supplied Windows Access observations.

Existing formatter/window checks2/2 (6.01s); application built.

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
