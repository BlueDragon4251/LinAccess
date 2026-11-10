# Native DAO Currency field creation

TableDef.CreateField(Name,dbCurrency) now builds actual Currency columns in new
native tables and through Fields.Append on existing tables. The existing checked
CURRENCY_REAL storage protocol is reused: lossless four-decimal values strictly
inside +/-2^39. Native source metadata retains dbCurrency5, Size8, fixed flags and
Required; drafts, SQL actions, alias cursors, fields and formatting use the same
already supported Currency paths. A physical CHECK prevents raw SQL from storing
non-normalized values or values beyond this protocol. This does not expand native
storage to the complete Access Currency Int64-tick range; F014/F105 remain open.

Creation defaults support Null or signed decimal literals with at most four
places, through exact decimal-to-Currency conversion before checked real storage.
Unmeasured excess precision, exponent/call expressions and unsupported ranges
refuse3251 before replacing the detached default. Reported native DefaultValue
normalizes the actual stored literal back to its Currency text instead of leaking
binary64 serialization digits. Existing Required/default SQL enforcement, append
savepoints and field revision guards are retained. Full Currency schema rebuild/
type conversion, numeric ValidationRule/relations and every default-expression/
Windows provider behavior remain separately bounded; no general parity claim.

App/DAO/query-model/Currency targets built. A missing fixture include was corrected
before checks. The existing five-check bundle passed four checks and exposed an actual physical
type mismatch in DAO: the metadata protocol requires physical REAL. The declaration
was corrected to REAL, retaining logical CURRENCY_REAL from source metadata. Only
affected DAO was retried, passing1/1 in3.50s. Actual new-table/
existing-table field creation, source Type/Size/Required/default metadata, draft
and physical SQL defaults, compiled SQL Currency rounding, raw SQL refusal,
unchanged rejected defaults, and actual typed VBA dbCurrency creation/insert/
function return are covered. The four passed Currency/model checks from the first bundle were not repeated. No new CTest/catalog ID;456/386/70 unchanged; F055/F121 remain broad.

Latest private published predecessor2af79959 retains the mechanically identical
native type vocabulary bodies. Latest checked DEB/Original remainsa23c7009,
Original1/1 in119.54s and seven protected hashes unchanged; that package predates
later field/container/report/creation changes. No fresh Original/package or
Windows Currency creation oracle is claimed here.
