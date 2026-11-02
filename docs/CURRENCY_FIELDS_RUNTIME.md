# Bounded native Currency fields over REAL backing

Verified CURRENCY source metadata with REAL backing now exposes actual direct-
table Currency fields as fixed int64 units (four decimals), through DAO and
bound native text controls. Field.Type returns Integer dbCurrency (5), including
EOF; VarType returns vbCurrency (6). Null remains Null. Verified source metadata
is required: ambiguous old FLOAT labels do not acquire an inferred Currency type.

The supported storage interval is strictly -2^39 < amount < 2^39
(-549755813888 < amount < 549755813888). Within it, adjacent four-decimal Currency
ticks have distinct IEEE binary64 encodings. The maximum spacing is 2^-14,
smaller than one 0.0001 tick. Native decoding extracts the 53-bit mantissa and
computes value*10000 as an integer mantissa*625 divided by a power of two. The
product fits 63 bits. Integer rounding recovers the exact tick, and an independent
canonical re-encoding must equal the stored Double. This avoids floating-point
multiply-and-round guesses at large values.

Writes convert through the existing checked Currency conversion and validate the
same round trip before field/buffer mutation. Valid Currency outside this backing
range diagnoses 3251; noncanonical/nonfinite stored numbers diagnose 13. Exact
full-range unbound Currency arithmetic and existing TEXT-result paths retain
their previous behavior. [Currency query parameters](CURRENCY_PARAMETERS_RUNTIME.md) subsequently add
bounded typed/numeric binding. Exact formatted regression results in this field
test remain explicitly stored via Text parameters.

Text controls format these fields from recovered integer units, avoiding Double
string shortening during load, assignment or Ctrl+apostrophe. Pending Text/Value,
OldValue, validation/save and Undo retain existing native semantics. Copying into
a new record leaves accepted Value unchanged until control validation and does
not implicitly insert a record. Other control presentation, regional currency
symbols/formatting and complete Windows editor behavior remain open.

Core tests compare boundary/power-of-two neighborhoods and 10000 deterministic
values against independently parsed exact decimal text, including negative
values, overflow/invalid backing, Null/Error/Nothing and scalar-once conversion.
Actual Qt/VBA/DAO/SQL tests cover fixed arithmetic, calculated controls, type
codes, pending text/rounding/save, programmatic values/OldValue/Undo, checked DAO
Edit/Update, range rejection before mutation, large value display/navigation,
copy/new-row/no-insert and invalid raw storage without hidden repair.

All 121613 non-NULL Currency values across the original 43 Currency columns are
read through native DAO and compared as exact integer-unit row/NULL multisets
against independent original CSV decimal parsing. Four-decimal export syntax is
checked; all existing CSV, Long/Boolean/date/counter checks remain active. Original
layout and Kassa-DAO regressions also pass. The corrected native test passes 1/1
(0.09 s); core/original regressions passed 4/4 in the preceding targeted run.
GUI suite now has 38 CTests; full validation is recorded in PROGRESS.md.

Full-range exact field storage, raw SQL Currency coercion/operators/aggregates/
sorting, full-range/SQL parameter semantics, saved-query/computed-column subtype propagation,
Single/Decimal/Date storage and full Jet/Windows parity remain open. Raw SQLite
writes can still violate the four-decimal/range contract; native reads diagnose
that storage rather than guessing a value. This bounded native field adapter
adds no complete original business workflow or blanket money compatibility.
