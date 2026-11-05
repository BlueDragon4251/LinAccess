# Native FV, Pmt, PV, NPer, SLN and SYD

FV(Rate,NPer,Pmt,[Pv],[Type]) and Pmt(Rate,NPer,Pv,[Fv],[Type]) now execute
native C++ floating-point calculations. The VBA binder supports the documented
named arguments, including Type (not the differently named .NET Due argument),
and omitted optional arguments default to zero. The SQLite adapter calls the
same expression-engine implementation for real Access SQL calls.

Supported scope: finite Double-coercible inputs, whole payment periods0..32767,
rate greater than-1, and Type0/end or1/beginning. Pmt with zero periods raises5.
Fractional/negative/larger periods, rate<=-1 and other timing values explicitly
raise3251 rather than invent compatibility outside the proven scope. Null raises94;
missing required/wrong arity raises450; nonfinite arguments/results raise6.
The result retains Double identity. There is no complete financial library claim:
NPV/Rate/IRR/MIRR/DDB and other financial functions and Windows error/edge-case oracle remain open.

The computation uses long-double intermediates and log1p/expm1 to avoid the
usual near-zero-rate loss from subtracting1 from a rounded power. Rate0 uses
the linear formula. Payment timing multiplies the annuity factor by1+rate.
The signs follow cash-flow direction, rather than absolute-value correction.

Existing core cases compare FV with independently accumulated monthly deposits,
and Pmt with an independently amortized balance including a future cash balance.
They check beginning/end timing, omitted/named arguments, zero/tiny rate and
zero-period rejection. Existing DAO cases execute both functions over actual
SQLite financial rows and independently check the balances. No extra CTest target.

Validation: targeted core/DAO/query3/3 (2.38 s), fresh GUI-free core1/1 (1.54 s),
Debian package/build ID/pinned SQLite and four original fixture hashes passed.
Last full baseline immediately before this change48/48 (82.86 s); no full-suite
rerun is represented as current. F068 remains open; fixed catalog456/384/72.

References: [VBA FV](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/fv-function),
[VBA Pmt](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/pmt-function).
No execution against Windows Access2003 was available.

PV(Rate,NPer,Pmt,[Fv],[Type]) now shares the annuity calculation and its bounded
period/rate/timing rules, returning the actual discounted value. SLN(Cost,Salvage,
Life) and SYD(Cost,Salvage,Life,Period) execute straight-line and sum-of-years
asset depreciation with finite Double inputs, nonnegative cost/salvage, positive
life and (for SYD)0<Period<=Life. Negative asset inputs are explicitly3251;
invalid lifetime/period5, Null94, arity450 and nonfinite result6. These remain
bounded native numeric paths, not the complete Windows financial implementation.

Existing core/DAO cases independently sum discounted payments, execute both
asset functions against five actual SQLite period rows, and check the complete
9000-unit depreciation with first-period3000 and constant straight-line1800.
New validation for this extension: targeted3/3 (2.36 s), fresh headless core1/1
(1.66 s), Debian/build ID/SQLite and four original fixture hashes passed. Last
full baseline remains48/48 (82.86 s), not rerun for this bounded extension.
Additional references: [VBA PV](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/pv-function),
[VBA SLN](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/sln-function),
[VBA SYD](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/syd-function).

## NPer

NPer(Rate,Pmt,Pv,[Fv],[Type]) now computes the actual number of payment periods
in the native expression engine and SQL adapter. Optional Fv/Type default to0;
VBA named arguments and VBA.NPer are supported. Finite Double-coercible inputs,
rate greater than-1 and timing0/end or1/beginning are supported. Results may be
fractional and retain Double identity; the earlier FV/PV/Pmt whole-period bound
does not constrain this inverse result.

Zero rate uses the linear balance. Nonzero rate uses long-double logarithms,
log1p for a ratio near1 and separate magnitude logs otherwise. No zero/negative
logarithm or indeterminate linear solution is returned as a successful value.
No finite unique real solution raises5; unsupported rate/timing3251, Null94,
arity450 and nonfinite argument/intermediate/result6 use the native error path.
Complete Windows edge-case/error/precision comparison remains open.

Core cases check zero/tiny rate, a one-payment balance at both timing choices,
capital doubling and named qualified calls. Actual SQL financial rows invert
their computed payments and match independently amortized twelve-period balances.
Existing core/DAO/query3/3 (2.76 s), fresh GUI-free core1/1 (1.77 s), updated
Debian/build-ID/SQLite3.54.1 and all four original hashes pass. Last broader full
48/48 (87.98 s); no new target/ID. F068 remains open, fixed456/384/72.
