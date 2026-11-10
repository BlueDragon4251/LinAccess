# Native SQL CStr

The native Jet SQL adapter registers CStr through the shared VBA evaluator.
Scalar text (including embedded NUL), exactly representable SQLite integers and
finite floating values use the existing native String conversion. Null raises
94, binary input raises13 and invalid arity raises450. Integers outside the
exact Double-backed native Variant range are rejected rather than rounded.

The existing token-based typed projection compiler preserves proven Boolean,
Byte/Integer/Long, Single, Double, Currency and Date inputs using the common
checked field decoder before CStr. This includes declared parameters, supported
physical source fields/aliases and declared native standard-function results.
Boolean literals and nested checked conversions retain their kinds. No source
text substitution, inferred business values or new result cache is used.

This is native conversion support, not proof of Access locale parity. It shares
the existing native date, numeric and Currency bounds. The current typed-query
resolver does not prove arbitrary joined/subquery/view expression kinds; those
remain outside this typed-field claim.

App/DAO built; the existing dao_execution bundle passed1/1 in2.43s. The actual
DAO path covers typed fields/aliases, a Boolean query parameter, nested
conversions, Currency, Single, Date, Boolean literals, embedded NUL and checked
Null/binary/range/arity failures. Fixture corrections preserve the prior
NativeScale update (first input is2) and use a real BLOB field instead of a
SQLite-only hex literal in Jet SQL. No new CTest or catalog identity.
Latest checked package is application62d00929 with the packaging correction
published5ec76beb, before this function; original116.96s, inventory49.42s, seven
protected hashes unchanged. No Windows or full-suite claim.
