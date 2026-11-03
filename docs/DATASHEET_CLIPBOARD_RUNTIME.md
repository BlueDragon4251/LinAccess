# Native datasheet selection and text copy

Shift+Space selects the current row, Ctrl+Space the current column and Ctrl+A
all cells. Column/all selection fetches the remaining Qt query batches first;
it refuses an unsupported size instead of silently selecting only the initial
fetch batch. The current limit is one million cells.

Ctrl+C and the main-window Copy menu command copy a rectangular native table
selection to the actual Qt text clipboard. Rows use CRLF, columns use tabs;
tabs/newlines/quotes inside a value use doubled-quote TSV quoting. No header row
is added. The native table model provides checked values, retaining Currency
precision, finite numbers with round-trip digits, ISO dates and Boolean text.
Null and empty text both appear empty in this text format.

An active editor must pass normal cell validation into the manual-submit cache
before copying; copying does not submit database changes. Binary/non-NULL BLOB,
NUL-containing text, hidden cells, disjoint selections and oversized text are
refused before replacing the clipboard. Text is capped at four million UTF-16
units. Cut/Paste/Delete for whole ranges remain disabled in the editing menu.

The existing main-window/Currency targets cover actual selection/copy keys and
menu dispatch, SQL-backed cells, Currency/date/Boolean values, invalid editor
preservation, quote handling, full selection beyond the first fetch batch and
binary/disjoint refusal. Typed clipboard MIME, range paste, headers, regional
format masks and exact Access/Windows clipboard representation remain open.
