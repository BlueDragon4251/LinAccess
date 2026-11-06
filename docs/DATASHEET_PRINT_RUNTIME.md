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
the default prints the whole supported visible datasheet; bounded full-row selection is described below.

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
pagination, arbitrary cell-only output, full Access page setup, print events or
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


## Datasheet preview and VBA print commands (2026-11-06)

The same checked snapshot now drives a genuine modal QPrintPreviewWidget page
view with actual page count, page navigation and zoom. Its dialog owns the widget
and is destroyed before its printer. It offers no uncontrolled print-to-file
shortcut bypassing the atomic output controller. Failed rendering disables the
page view and preserves an explicit failure; source/project closure cannot be
reported as successful output. The original project/reentry/draft guards remain.

DoCmd.RunCommand acCmdPrint and Application.RunCommand acCmdPrintPreview address
the active table/query through the existing dispatcher, preserving caller Me.
Microsoft's actual AcCommand enumeration was read: Print340, PrintPreview54.
Unsupported form/report RunCommand coverage remains unchanged. The renderer now
uses a 72-dpi measurement device for true 9-point document metrics, independent
of screen DPI; an extracted PDF glyph-height assertion checks the actual result.

Existing core/mainwindow/layout3/3 (18.86 s), fresh GUI-free core1/1 (1.77 s),
current Debian payload/Build-ID/SQLite3.54.1 and four original hashes pass.
Actual VBA Print writes a private PDF, then continues in its caller; named
Application preview navigates real pages and continues with its own Tag intact.
The last single full desktop48/48 (92.62 s) is the preceding40c6b13 output commit,
not a new full-suite claim for this increment. No new target or ID;456/384/72 fixed.


## Datasheet page setup and persisted geometry (2026-11-06)

File > Datenblatt: Seite einrichten and native RunCommand acCmdPageSetup32 open
the existing functional Qt document-geometry dialog. It edits actual QPageLayout
paper size/custom dimensions, portrait/landscape and four millimeter margins.
The report dialog implementation is reused; its save choice is labeled for the
project. No fake printer or print job is created to configure page geometry.
Accepted geometry belongs to the open table/query document and drives the same
preview, PDF writer and print-layout transfer. Native print-dialog selection can
still supply actual device settings. Initial defaults remain A4/15mm; Cancel
preserves the prior layout. Existing snapshot/draft/content/size/modal/project/
reentry/weak-owner checks also gate setup; this does not expand binary printing.

The explicit Im Projekt speichern choice stores version1 mm geometry in the
checked main._la_datasheet_page_setup catalog, keyed by table/query kind and
case-folded object name. It writes only the native project, never the MDB or
archived VBA. A savepoint plus conditional existing-row update/new-row insert
rejects concurrent edits and does not commit a caller's transaction. Actual catalog
shape and resulting settings are verified; TEMP shadows cannot redirect reads/
writes. Declining persistence keeps session-only geometry. Name-based settings
are not proof of immutable drop/recreate/rename identity or imported printer data.

Malformed stored geometry blocks fresh-document output3251 without touching an
existing PDF. Setup explicitly offers repair geometry; persistence still requires
the save choice. Oversized/wrong-schema catalogs are refused rather than rewritten.
Project/table closure after a dialog cannot save/apply settings to another target.
Physical printer capabilities, original PrtMip/PrtDevMode/PrtDevNames, print events,
full Access format/layout/error equivalence and horizontal multi-sheet output
remain open under F080/F087/F148. No new catalog ID/checkbox/test target.

Existing mainwindow_documents checks real table/query setup, exact PDF MediaBox
orientation and independent25mm glyph x-position, close/reopen persistence,
cancel, TEMP shadow immunity, concurrent-update preservation, outer-transaction
rollback, malformed-setting refusal/repair and actual named Application.RunCommand
with caller Me/continuation preserved. Microsoft's actual AcCommand table confirms
PageSetup32; this is an API source, not an Access2003 execution oracle.
Targeted1/1 (3.64 s), then full current Debian13 desktop48/48 (89.86 s), including
both original recipe tables/39 rows from a protected copy, pass. Current package/
Build-ID/pinned SQLite3.54.1 and four original fixture hashes pass. Latest earlier
local Debian12/Ubuntu24.04 follow-ups cover41b193c DAO/payload/install/hash checks;
no new distribution matrix or GitHub CI claim for this page-setup increment.
456/384/72 unchanged.


## Checked multiple copies and page collation (2026-11-06)

The real print dialog now accepts1..100 copies. PDF jobs explicitly render each
requested copy in the chosen collation and forward/reverse page order, preserving
selected page-range semantics and atomic replacement. Manual multi-copy output
is capped at10000 pages. Unsupported counts fail3251 and retain an existing PDF.

For actual native devices, Qt's documented supportsMultipleCopies determines
whether the driver handles the requested count or the application renders copies
with the printer count reset to1. This prevents application/driver duplication.
Collation follows the actual printer setting. Physical spooler/hardware behavior
has not been tested and is not an Access2003 printer oracle.

Existing mainwindow_documents produces real four-page PDFs from two selected
pages/two copies: collated forward and uncollated reverse. Independently extracted
page text verifies exact order and equality of repeated pages; the invalid101-copy
request preserves a prior destination. Original recipe PDF proof is rerun in the
same target. Current target1/1 (3.99 s), package/Build-ID/SQLite3.54.1 and four
original hashes pass. Latest full host suite remains e8375e8 with48/48 (89.86 s).
Its updated GUI-free core1/1 (2.20 s) also passes. Immutable e8375e8 Debian12 and
Ubuntu24.04 app/DAO/mainwindow follow-ups pass2/2 (6.65 s /5.65 s), matching package
installation/backend and all original hashes. These platform checks precede the
copy increment; no fresh complete matrix or GitHub CI success is claimed.
F080/F087/F148 remain open;456/384/72 unchanged. No new CTest target or ID.
Reference: https://doc.qt.io/qt-6/qprinter.html#supportsMultipleCopies


## Frozen selected-record printing (2026-11-06)

The print dialog now enables Selection when fully selected visible records exist.
Actual selected model-row ordinals are captured before remaining batches load;
the owned string snapshot retains matching row ordinals, with no QModelIndex,
SQL cursor or host pointer kept across the modal dialog. Selected output preserves
visible column/row order, filter/sort, formatting and hidden-cell rules, and uses
the existing page geometry, atomic PDF writer and copy/collation path.

Only complete record selections are eligible; individual cell rectangles are not
expanded. An unavailable/empty Selection request fails3251 before output and never
falls back to the entire table. Dedicated PDF/preview defaults still address the
whole datasheet; selected PDF output is available through the real print dialog.
The complete bounded snapshot is still required, so selection does not bypass
whole-source cell/text/draft/binary limits or introduce unbounded streaming.

Existing mainwindow_documents prints two nonadjacent table records and one actual
readonly query record to PDFs, then independently extracts text to check inclusion
and exclusion. Changing the selection inside the dialog cannot change the frozen
output. Empty and partial-cell Selection requests leave an existing PDF untouched.
Target1/1 (4.15 s), matching app/package/Build-ID/SQLite3.54.1 and four original
hashes pass; the three uploaded original hashes are also unchanged. No new target,
ID or checkbox. Last full host suite e8375e8:48/48 (89.86 s), with its earlier
platform/headless evidence distinguished from these later GUI increments.
Full Access selection/error/printer/rendering equivalence remains open under
F080/F087/F148;456/384/72 unchanged.
