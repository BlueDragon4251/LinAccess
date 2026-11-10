# Native financial calculations

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
Unrestricted array/period semantics and the Windows error/edge-case oracle remain open.

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

## NPV and MIRR over actual Double arrays

NPV(Rate,Values) and MIRR(Values,Finance_rate,Reinvest_rate) now execute native
C++ calculations over allocated, one-dimensional Double arrays. The normal VBA
named-argument binder uses these documented names, including Values rather than
the differently named .NET ValueArray. Arbitrary lower bounds are supported;
ascending array position determines chronology, including explicit zero flows.
The native limit is1,000,000 entries. Variant/other-element/rank and empty arrays
are outside this bounded stage; they are not silently converted or flattened.

NPV discounts the first flow at the end of period1 using reverse Horner evaluation
with long-double intermediates. The first initial investment must be included or
added separately according to the intended dates; no automatic date correction
is invented. This bounded scope follows the documented positive/negative-flow
requirement and refuses single-sign arrays3251 pending an Access execution oracle.

MIRR discounts negative flows at the financing rate and accumulates positive
flows at the reinvestment rate. The first flow is at period0. The final annualized
per-period return uses log/expm1, with the actual array length minus1 determining
periods. At least two flows, both signs and rates greater than-1 are required.
Missing signs raise5, unsupported rate range3251, Null94, arity450, scalar-array
mismatch13 and nonfinite input/intermediate/result6 follow the native error path.
Complete sign/underflow/error precedence and coercion comparison remain open.

There is no SQLite array transport for direct SQL NPV/MIRR arguments. Public native
VBA functions may use supported local arrays through the existing native runtime,
subject to the SQL host's documented context/state restrictions. No SQL array
callback or complete financial library compatibility is claimed.

Actual VBA checks negative-lower-bound arrays, reversed named arguments, a mixed
cash-flow NPV, zero discount and separate financing/reinvestment across multiple
negative and positive periods.

## DDB

DDB(Cost,Salvage,Life,Period,[Factor]) executes native declining-balance depreciation
in VBA and real SQL rows. Factor defaults to2. The period's opening book value
comes from the declining-balance factor; depreciation is capped at the remaining
amount above salvage. Fraction at least1 consumes that amount in period1 and
returns0 for later periods. There is no invented switch to straight-line.

Supported scope: nonnegative finite asset values, salvage no greater than cost,
whole positive life/period up to32767, period no greater than life and positive
factor. Fractional/larger periods, salvage above cost and negative asset values
remain explicit3251 boundaries. Invalid lifetime/period/factor5, Null94, arity450
and nonfinite values6 use the native error path. Full Windows numeric/error
equivalence remains open. Actual SQL depreciation is checked against independently
updated balances, including the final salvage cap.

Combined current milestone: targeted core/DAO/query3/3 (2.54 s); full desktop48/48
(85.94 s), fresh GUI-free core1/1 (1.67 s), updated Debian payload/build-ID/pinned
SQLite3.54.1 and all four original fixture hashes pass. Original query coverage
stays78/100 without an Access result oracle. No new CTest target or catalog ID;
F068 stays open, fixed456/384/72.

References: [NPV](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/npv-function),
[MIRR](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/mirr-function),
[DDB](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/ddb-function).

## Iterative Rate and IRR

Rate(NPer,Pmt,Pv,[Fv],[Type],[Guess]) and IRR(Values,[Guess]) now execute the
native financial root solver, with default guess0.1 and at most20 iterations.
Rate is available in actual SQL expressions; direct SQL array transport for IRR
is not implemented. Native VBA calls support qualified and documented named
arguments. Results retain Double identity.

Each evaluation advances the actual balance and its analytic derivative together.
The solver uses long-double intermediates, a relative cash-flow residual threshold
of1e-12, Newton updates and backtracking inside rate greater than-1. No derivative,
no advance or no convergence returns native error5, not a fabricated0/guess.
Nonfinite arguments/intermediates/results raise6; Null94 and arity450 remain
native errors. Root existence, unique selection and exact Windows iteration/error/
rounding equivalence are not claimed. Multiple-root behavior remains open even
when the native solver finds a mathematically valid root near its starting guess.

Rate requires whole positive periods up to32767 and timing0/end or1/beginning;
other period/timing/initial-guess domains refuse3251. All-zero annuities raise5.
IRR uses the existing allocated one-dimensional Double-array scope and1,000,000
entry limit. Both cash-flow signs are required, and the first flow is at period0;
chronology follows ascending array position regardless of the lower bound.
The solver reads a scalar copy of each flow and does not mutate the input array.

Actual core/DAO/query3/3 (2.52 s) verify nondefault guesses, qualified named calls,
zero-rate annuities, inversion of actual SQL loan payments and IRR balances with
multiple positive/negative periods. Fresh GUI-free core1/1 (1.73 s), updated
Debian/build-ID/SQLite3.54.1/four fixture hashes pass. Last broader full suite is
the previous financial milestone48/48 (85.94 s), not a new full run after roots.
No new target or catalog ID; F068 stays open,456/384/72.

References: [Rate](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/rate-function),
[IRR](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/irr-function).

## IPmt and PPmt

IPmt(Rate,Per,NPer,Pv,[Fv],[Type]) and PPmt with the same signature now compute
the real interest/principal components of a payment in native VBA and SQL. The
normal binder supports documented named arguments. Fv defaults to0, Type to0.
Shared stable annuity factors supply the payment; the actual balance is advanced
through prior periods. Interest is0 for the first beginning-of-period payment;
later beginning payments account for interest from the preceding period. Principal
is payment minus interest. Cash-flow signs are preserved.

The bounded domain is finite Double-coercible inputs, rate greater than-1, timing
0/end or1/beginning and whole periods1..NPer up to32767. Invalid period range5,
unsupported fractional/larger/rate/timing3251, Null94, arity450 and nonfinite
argument/payment/balance/component6 use the native error path. Exact Windows
rounding, error precedence and unrestricted domain behavior remain open.

Existing core cases validate all payments of an independently advanced credit
balance and their sum against Pmt, beginning/end timing and zero-rate principal.
Actual SQL financial rows verify interest plus principal equals payment. Final
core/DAO/query3/3 (2.68 s), fresh GUI-free core1/1 (1.79 s), current Debian/
build-ID/SQLite3.54.1 and all four original hashes pass. Last broader full remains
fb63120 financial milestone48/48 (85.94 s), preceding roots/payment components.
No new target or catalog ID; F068 remains open,456/384/72.

The native entry points now include FV, PV, Pmt, IPmt, PPmt, NPer, Rate, NPV,
IRR, MIRR, SLN, SYD and DDB within their explicit domains above. This inventories
implemented entry points, not complete financial-library or Access2003 parity.
References: [IPmt](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/ipmt-function),
[PPmt](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/ppmt-function).


## Proven native SQL argument types (2026-11-10)

FV/PV/Pmt/NPer/Rate, SLN/SYD/DDB and IPmt/PPmt preserve proven native Date,
Boolean, Single, bounded Currency and integral/Double arguments through the
shared scalar binder. Date values reach the numeric calculation as OLE Date
Variants; True remains -1 even for physical BOOLEAN1; Single input is rounded
before calculation. ISO TEXT does not become a Date declaration by its contents.
Nested supported expressions and simple WHERE/HAVING/GROUP-BY/ORDER-BY paths
retain these bindings and declared Double result metadata, including real DAO
Fields.Type=dbDouble7 and Double values.

The numerical algorithms, iteration count, supported domains, timing0/1 rules,
Null94, unsupported-domain3251, type13 and nonfinite6 behavior are unchanged.
In particular a Boolean True timing argument is -1 and does not acquire the
semantics of physical integer1. This adds no array transport for NPV/IRR/MIRR,
no new solver, no Windows precision/error oracle and no full financial-library
or Jet equivalence claim. The raw unproven-input callbacks keep their limits.
