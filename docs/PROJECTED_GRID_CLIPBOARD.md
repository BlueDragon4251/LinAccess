# Visible grid-column clipboard mapping

Native grid copy/cut/paste now uses the real visible column order rather than
physical neighbours. A projected query row selection exports only its visible
fields; hidden keys/omitted data are not sent to the clipboard. Ordinary hidden
selection guards remain. Nonrectangular visible selections and native text/MIME/
cell limits still refuse explicitly; output is built before replacing clipboard.
Native typed fields identify their actual source fields, not invented alias keys.

The table model accepts a validated unique physical column vector for a block.
All target values are converted before cache mutation. Existing read-only/generated/
BLOB/type/row/range rules and atomic failure restoration remain. Unselected fields
and unrelated cached drafts retain their values. Undo/Redo retains the exact
column vector together with converted cells, prior cache and insertion defaults.
Redo does not replay into physical neighbours. Mapped insertion preserves actual
omitted SQL defaults and key generation. Named native record append keeps its
existing source-field/protected-hidden-field checks.

Cut clears precisely the copied visible columns. Paste maps its visible starting
field and following visible fields once, then selects actual mapped cells.
This replaces the preceding projected-query multi-column3251 refusal. General
Windows/OLE formats, all locale/mask semantics and complete Access range history
remain open; no full F340/Access clipboard parity is claimed.

App and existing targets built. First combined GUI/currency-field/model run3/3
passed18.69s (18.11/0.48/0.11). The actual stored projected query proves TSV paste
in reordered nonadjacent columns, Save by the hidden real key, omitted SQL values
unchanged, visible-only typed copy, Cut and Undo. Model cases prove multirow maps,
Undo/Redo, late conversion/duplicate-map failure without changes, mapped append,
real omitted SQL defaults and SQL results. Existing exact int64/Currency/BLOB/
NULL/Date/Boolean native MIME and record append cases also passed. No new CTest,
Original/Windows/full-suite/DEB refresh. Catalog456/386/70 remains unchanged.
