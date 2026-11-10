# Native Currency relationship keys

Native DAO relations now accept matching dbCurrency keys for enforced and
one-to-one relationships. Both fields must have explicit native fixed-field
metadata, logical CURRENCY_REAL semantics and physical REAL storage. Imported,
unknown, mixed-type and legacy MONEY representations are not promoted into this
contract. The existing bounded exact four-place amount range remains unchanged;
this does not provide full Jet Currency range or a Windows parity oracle.

The existing checked unique/required parent-key proof, atomic schema migration,
foreign-key validation, row identity/value digest and custom-schema rejection
remain active. Cascade update/delete uses real SQLite foreign keys. Native field
reordering and relation removal reconstruct the same checked Currency storage.
An orphan prevents relation creation without replacing the existing rows or
installing a partial foreign key. NULL child keys remain valid.

App and DAO targets built; the existing dao_execution bundle passed 1/1 in
2.44s. It exercises four-place positive/negative keys, orphan rollback and retry,
raw non-quantized write refusal, update/delete cascades, retained binary payload,
NULL child rows, field reorder, relation deletion and legacy metadata refusal. No
new CTest or catalog identity is introduced; F014/F055/F121 remain broader open
requirements.
