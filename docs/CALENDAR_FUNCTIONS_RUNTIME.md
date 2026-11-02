# Bounded native Gregorian date functions

The C++20 runtime and Jet SQL adapter share native DateDiff, Year, Month, Day,
Hour, Minute, Second and Weekday calculations. Input is unambiguous Gregorian
ISO text: YYYY-MM-DD, YYYY-MM-DD HH:MM:SS (also T separator), or HH:MM:SS. The
supported year range is 0100–9999. Time-only/Empty input uses the OLE epoch date
1899-12-30; date-only input uses midnight. Invalid date/time text is rejected,
including non-leap February 29. No system timezone/DST conversion is involved.

The part functions and Weekday return VBA Integer; Null input returns Null.
Weekday defaults to Sunday and supports explicit weekday constants 1–7.
DateDiff returns a checked Long and counts interval boundaries, not elapsed
fractional durations: yyyy, q, m, y, d, h, n and s. The ww interval counts calendar
week boundaries with default/explicit first weekday and the default first-week
setting. Negative directions and dates before the Unix epoch are handled with
integer calendar arithmetic. Long overflow reports 6. DateDiff Null reports 94.
Optional argument omissions and VBA weekday/first-week constants are recognized.
Objects are scalarized once per argument. SQL uses the same calculations and
preserves Null/integer results; existing query alias processing recognizes the
new pure scalar functions. SQL date literals use the existing Jet literal parser.
DAO query-column subtype inference is not extended by this batch.

Focused coverage extends existing core and query-model tests; no new test target
is added. It exercises actual VBA procedure execution, scalar getter count,
Integer/Long VarType, leap days, boundary differences, negative results, week
starts, omitted defaults, Null, invalid dates/arguments/options and overflow.
Actual SQL covers all eight functions, Jet date literals and result aliases.

This does not implement a Date Variant, CDate, locale-dependent text parsing,
numeric date coercion, fractional seconds, Hijri calendars, system regional
options or DateAdd/DatePart/DateSerial/TimeSerial. Unsupported numeric/calendar
options report 3251; w and non-default first-week ww remain explicitly unsupported.
Malformed text/argument counts/options are diagnosed rather than guessed. Windows
Access differential validation is still unavailable. Public contract references:

- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/datediff-function
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/year-function
- https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/weekday-function
