# Native field DLookup

Global DLookup and Application.DLookup accept positional/named Expr, Domain and
optional Criteria in a live project context. The current subset binds Expr to
one actual existing field (including bracketed names) and Domain to a physical
table or installed non-parameter saved query. Canonical domain/field lookup is
case-aware and rejects ambiguity. Calculated Expr, explicit NULL criteria and
unproven non-NULL result types remain explicit unsupported cases. Missing/Empty
optional criteria acts as omission; ordinary string predicates use the existing
bounded Access SQL compiler. Full expressions, parameter-domain/reference binding,
collations and exact Access error identity remain open.

The implementation prepares actual source schema, then executes a read-only TOP1
field selection with the supplied predicate. It returns NULL for no match or a
NULL matching field; an empty string remains String. It never turns a missing
domain, unsupported field or failed query into a fabricated empty lookup. Proven
result-position types feed the existing native converter, retaining supported
Date/Boolean/Currency rather than guessing from cell text. No arbitrary ordering
guarantee is asserted when criteria do not uniquely select a record.

Each call reads the real database. Unsaved form buffers are excluded; after an
actual record save the new value is visible. No persistent lookup cache or host/
cursor retention is introduced. Only SELECT/non-ranked field projections are
allowed; original MDB files are not edited by lookup.

The existing native form case executes actual VBA and a calculated textbox over
a real nullable table and saved-query alias. It covers named/global/Application
calls, apostrophe criteria, empty/NULL/no-match distinctions, Date/Boolean/Currency
identity/value, and a dirty label before/after actual Save. Missing domains and
computed Expr refuse explicitly. No additional CTest target or catalog ID.

Reference: [Access 2003 DLookup](https://documentation.help/MS-Office-Access-2003/file9o.htm)
describes table/non-parameter query domains, no-match NULL, arbitrary selection
without unique criteria, and exclusion of unsaved records. This is a bounded
field implementation, not complete expression/domain compatibility.
