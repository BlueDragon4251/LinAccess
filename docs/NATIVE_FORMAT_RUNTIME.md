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
