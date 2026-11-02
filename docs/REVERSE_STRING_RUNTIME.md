# Native reverse string functions

InStrRev and StrReverse now execute in the C++ VBA expression runtime and native
Access SQL queries. SQL arguments are bound values passed to the same native
runtime; database text is never interpolated as executable code.

InStrRev returns Long positions in UTF-16 code units, with optional start and
comparison settings. Default comparison is binary; explicit vbUseCompareOption
uses the current VBA module setting. SQL has no module Option Compare context.
An oversized start returns zero; zero or values below -1 fail 5. Null arguments
fail 94. Named scalar projections can be referenced by later query expressions.

StrReverse reverses UTF-16 code units. A result containing isolated surrogates
fails explicitly with 3251 because the runtime boundary uses valid UTF-8. It does
not silently substitute Unicode replacement characters or claim full BSTR parity.
SQL rejects invalid UTF-8 and BLOB input with 13. SQL errors retain native error
numbers. General Windows locale, database collation and DAO projection subtype
inference remain open; the SQL adapter currently has the base host's bounded
comparison behavior. No Windows Access result oracle was available.
