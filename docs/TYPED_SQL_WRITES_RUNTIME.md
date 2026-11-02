# Checked target types in bounded SQL writes

The existing physical-table UPDATE/INSERT VALUES/simple INSERT SELECT path now
also converts verified source Byte, Integer, Long, Boolean and canonical date
storage targets. Money retains its existing separate exact Currency conversion
and proven fixed-point RHS arithmetic. Target types come from live schema/source
metadata; ordinary unmarked SQLite INTEGER/REAL storage is not guessed as Jet
Byte/Integer/Long/Currency.

Byte/Integer/Long use the shared checked nearest-even assignment conversion and
signed/source ranges. Boolean stores -1/0. Date assignment uses the existing
unambiguous date/time text conversion; canonical stored text includes midnight
when needed. The shared date reader/converter rejects years outside 0100–9999,
including QDateTime input, rather than accepting values outside the VBA domain.
Null remains Null; database required/unique constraints remain in force.

Callbacks execute before each target write. A conversion error aborts the whole
statement, including earlier rows. Original parameter order, simultaneous update
semantics and materialized positional append source remain intact. Existing
Currency behavior is retained. No Date Variant, complete numeric-date coercion,
locale coercion or general SQL expression subtype inference is claimed.

The existing DAO test now exercises actual SQL UPDATE/INSERT/append, Byte/short/
Long bounds, ties, canonical Boolean/date storage, Null, declared Double parameter
conversion, late-row failure rollback and unchanged ordinary SQLite INTEGER
behavior. Existing Currency insert/update tests remain active. No new CTest
target was added. Joined/complex/view/TOP/nested/HAVING/set-operation action
conversion, comprehensive constraints/AutoNumber permissions, mixed arithmetic
and complete Jet/Windows behavior remain open.
