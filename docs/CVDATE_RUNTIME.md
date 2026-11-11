# Native CVDate compatibility conversion

MS-VBAL6.1.2.3.1.4 defines CVDate(Expression As Variant) As Variant: its conversion
is identical to CDate, with a Date data value inside the declared Variant result.
The native intrinsic reuses the existing checked Date conversion and marks the
VBA result's arithmetic binding as Variant. CDate retains its declared Date
binding. Qualified VBA.CVDate and the Expression named argument share this path.
Null, Error values and invalid arity retain the existing conversion errors.

The SQL bridge accepts CVDate through the same bounded raw/proven-argument
adapters as CDate. Proven source kinds are decoded before conversion; results
use the existing canonical Date transport and DATETIME/dbDate metadata. No type
is guessed from sample result rows. Date/Currency/Boolean input checks and SQL
Date representation limits are inherited from the checked Date kernel.

This adds the legacy intrinsic; it does not add full regional string recognition,
Hijri/time-zone/subsecond SQL transport, referenced libraries or complete mixed
Variant arithmetic equivalence. No new Windows/Access oracle is claimed.
F041/F067/F068/F071 remain open;456 fixed IDs/386 confirmed partial scopes/70 open
full requirements remain unchanged.

App built. Existing core passed0.52s; after correcting a missing public SQL-name
recognition entry (3085), only the affected DAO target was repeated and passed
5.29s. Cases include named/qualified VBA, actual Variant-versus-Date arithmetic
binding, Date values/type inspection, proven SQL Date/Boolean/Currency sources,
negative fractions, DAO date fields and Null/Error/arity guards. No new CTest ID.
