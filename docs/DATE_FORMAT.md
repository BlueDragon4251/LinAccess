# Native date Format / Format$ subset

SQL and VBA expression lexers recognize the trailing dollar on function names.
Format and Format$ dispatch to a shared C++/Qt date formatter, also exposed through
native Qt form/report VBA hosts. This is execution, not procedure recognition.

Supported masks are yyyy, mm, dd and Long Date. The first three produce fixed
Gregorian year/month/day strings; Long Date uses the current Qt default locale.
Accepted inputs are canonical ISO date/timestamp text, Empty (OA day zero) and
numeric OLE Automation day values in years 100–9999. Negative fractional OA days
retain their integral calendar day. Numeric formatting, time masks, custom masks,
alternative calendars, locale-specific input parsing and a full Date Variant
representation remain unsupported.

The date parser validates wall-clock time without applying a time zone. A valid
Jet timestamp in a host daylight-saving gap therefore keeps its calendar date.
The dedicated test runs with TZ=Europe/Berlin and covers that boundary.

Format$ rejects Null with VBA error 94. The Variant-returning date Format path
retains Null. Unsupported masks, invalid inputs and date overflow fail explicitly.
Native SQLite callbacks copy UTF-8 output and propagate the bounded formatter's
error numbers back to DAO/query execution. A Qt-free host without this adapter
fails explicitly instead of claiming it executed Format$.

## Evidence and remaining original-query blockers

The date-format regression suite covers year/month/day output, leap days, negative OA fractions,
invalid clocks/dates, NULL, overflow, German/US Long Date locale output, stored
query dependencies, SQL and real Qt VBA button execution. The original 29-table
MDB regression additionally executes Format$ on actual non-NULL values in ten
original date columns and verifies year/month/day against the stored ISO values.
Original fixture hashes remain unchanged.

All 100 original Kassa queries are still classified. Subsequent concatenation
support now executes 49 of them and binds six of nine original report sources to
54 non-NULL controls (SQL_CONCATENATION.md). Three report sources still require
localized SQL form references; fixing syntax alone does not establish complete
original report execution.

Qt locale output is not proof of the original Windows/Access locale. Database
locale metadata and Microsoft Access reference results are still needed for
cross-engine formatting equivalence.
