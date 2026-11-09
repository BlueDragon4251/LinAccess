# Native CVErr value conversion

MS-VBAL6.1.2.3.1.14 requires an existing Error argument to be returned as a value,
without trying to convert it to a number. CVErr now preserves this actual Error
value. Other arguments use the existing Long Let-coercion after normal scalar
object-default resolution, retaining typed Decimal/Currency rounding and errors.
Codes outside0..65535 fail5 after successful Long conversion; Long overflow remains
6, Null94 and Nothing91. Existing arity and named/Missing transport boundaries are
unchanged; complete SCODE/BSTR/Missing equivalence remains open.

Existing shared core/native project source checks nested CVErr(Error2015), type10,
upper valid code65535, Empty/Error0, Currency2.5→2, Decimal3.5→4, invalid rounded
code5, Long overflow6, Null94 and Nothing91. Actual stored pure standard-module
SQL functions additionally return the shared project Err.Number907 and the text
of a preserved CVErr(Error2015). The Error value creates no new raised error;
the actual project error object still contains907 after both SQL functions.
SQL object/array-state and result guards are retained. The temporary technical
module and error state are cleaned up after the SQL assertions.

App built; first existing core/GUI bundle2/2 passed22.83s (4.65/18.17).
No new CTest ID, Original rerun, Windows/full-suite/package refresh. The Original
path last passed118.10s and seven hashes were unchanged before this isolated
CVErr change. Project error environment725b4bd was privately published/verified.
Latest checked DEB55096a7 precedes project Err/CVErr changes. Catalog456/386/70
unchanged; complete F041/F051 remain open, bounded F049 is extended in place.
