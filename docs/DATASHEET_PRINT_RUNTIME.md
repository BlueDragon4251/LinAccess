# Native datasheet print and PDF

The File menu now prints an active table or saved-query document through a real
Qt print dialog, and offers **Datenblatt als PDF**. Ctrl+P bubbles from the actual
datasheet/editor to the same main-window controller; it installs no global
shortcut competing with form KeyPreview/KeyDown. Native form/report print paths
remain separate. This is bounded native output under existing F080/F148, with
F087's complete output/reference requirement still open; no catalog IDs are added.

The controller validates an active table editor into its existing manual-submit
cache. Any unsaved table/editable-query changes refuse 2501 before the output
dialog. There is no implicit Save, discarded draft or fabricated printed value.
Editable saved-query source guards remain active. Readonly query results use their
actual materialized model; expressions are not executed again to print them.

A synchronous bounded snapshot owns only strings and column widths. It fetches
all remaining table batches, takes visible columns in horizontal header order,
visible rows in vertical header order, and preserves the model's actual filter
and sort. Header labels and the existing checked cell formatter provide values:
Null is blank, supported Currency/Date/Boolean/number/text conversions remain as
documented for native clipboard values. Non-null binary values, NUL, malformed
Unicode, failed SQL fetching/model changes, more than one million candidate cells,
and more than 16 Mi UTF-16 text units refuse explicitly. Hidden cells are omitted;
this prints the whole supported visible datasheet, not just selected records.

The snapshot precedes either modal dialog. Subsequent SQL writes do not change
its values. A scoped output/reentry flag and native project-close participant
prevent project replacement while preparing or displaying output dialogs.
Weak owners, actual tab membership, project generation and model identity are
checked after acceptance; closing a tab cannot render/resurrect it afterwards.
PDF/print requests from other modal contexts or executing query callbacks refuse.

Qt's text layout draws actual paginated vector output with wrapped plain text,
proportional visible column widths and repeated header rows. A4/15-mm margins are
the PDF action defaults, with native 9-point text. The print dialog supplies actual
page layout, page range and forward/reverse page order. Page/painter failures
propagate. Column geometry which exceeds the selected printable width is refused,
not silently clipped. Empty results retain their real headers. This does not
implement Access field-format masks, row-height/font parity, horizontal multi-sheet
pagination, selection printing, multiple copies, full page setup, print events or
an Access 2003 printer/rendering oracle.

Dedicated PDF output and PDF selected in the print dialog both use QPdfWriter
with QSaveFile, direct-write fallback disabled, and commit only after rendering
and writer finalization. Existing targets survive conversion/render/range/I/O
failure. A .pdf extension is required, symbolic-link targets and the current
project path are refused. Other native print-to-file formats refuse 3251.
This is not a guarantee against arbitrary external filesystem races.

The existing mainwindow_documents target exercises actual SQL/view values,
311 filtered/sorted rows beyond Qt's first fetch batch, moved/hidden columns,
multiple PDF pages, independently extracted text and repeated headers, readonly
query output, rejected dirty/binary/symlink targets, late rendering failure,
frozen dialog data and closed-tab refusal. A real print dialog is rejected for
Ctrl+P, and its shared printer is separately configured for a one-page PDF job;
no physical printer is accepted. Optional LINACCESS_DATASHEET_REFERENCE points to
the recipe MDB: both real tables/39 rows/text/Null values reach PDFs on an isolated
import, with unchanged original SHA-256. Temporary PDFs and originals never enter
the Debian package. No new CTest target or checkbox count.

Validation: final current desktop48/48 (92.62 s), fresh GUI-free core1/1 (1.90 s),
current Debian payload/Build-ID/pinned SQLite and four original fixture hashes pass.
Original recipe proof uses the supplied byte-identical MDB. Remote Debian12/13/
Ubuntu24.04 CI is blocked before jobs by account payment/spending limits.
