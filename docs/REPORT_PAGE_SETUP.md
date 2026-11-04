# Native report page setup

The report toolbar and preview S key open a functional Qt page-geometry dialog.
It edits a real QPageLayout: Qt standard paper sizes, positive custom dimensions
up to 5000 mm, portrait/landscape and four margins in millimeters. Invalid paper
or margins leaving no printable area are refused. Existing minimum margins are
retained; UI precision is 0.01 mm. Cancel leaves document geometry unchanged.

Accepted geometry updates the native paginated report view. Report PDF export,
the separate report preview and the print dialog copy that same document layout
into their actual QPrinter. Unsupported printer-layout transfer fails explicitly.
The one-page current-layout tools remain separate. Rendering still applies its
width, band/control and source-availability guards to the chosen printable area.

The first implementation used QPageSetupDialog, but Qt refused geometry changes
for the PDF printer used when no physical printer is installed. The application
now provides its own native Qt dialog backed by QPageLayout, without a fabricated
printer or dummy print job. It works with the actual PDF backend. Physical printer
capabilities/output remain unvalidated; the print dialog remains the user's job
selection and confirmation path.

These settings belong to the open report session. Full Access page/printer object
properties, original PrtMip/PrtDevMode/PrtDevNames decoding, persisted design setup,
column layouts and Windows print equivalence remain open. No source MDB is edited.

The existing report case opens setup via an actual S key event, accepts landscape
and a 25-mm left margin through real controls, and independently checks PDF
MediaBox orientation and extracted glyph x-position. Reopening exposes the saved
session values; changing orientation then cancelling retains landscape in another
actual PDF. No additional CTest target or catalog point.
