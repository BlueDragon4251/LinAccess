# Native Currency index navigation and relationship authoring

The actual relationship designer offers native Currency fields while referential
integrity or one-to-one mode is enabled. A shared exact-native-storage check is
used by that UI, DAO relation enforcement and table-type index navigation. It
requires explicit native metadata and the existing physical REAL/logical
CURRENCY_REAL protocol; imported/legacy declarations remain excluded.

Non-null Currency index keys now support the existing Index and Seek operations.
Actual Currency values and supported numeric keys pass through the field storage
codec; textual keys remain rejected. Existing ascending range, composite-key,
index identity, pending-edit, clone, transaction and cursor safeguards remain.
Descending range semantics and nullable keys remain outside this scope. This is
bounded native storage, not full Jet Currency range or a Windows Seek oracle.

App/GUI/DAO built. One existing bundle passed2/2 in14.37s: GUI11.91s and DAO2.46s.
GUI creates a real Currency foreign key, updates the parent and verifies the child
cascade, then deletes and verifies the cascade. DAO covers typed Currency equality
and range keys, numeric keys, invalid text with retained current value, NoMatch
and bookmark recovery, plus the previous checked integrity/migration cases.
No new CTest/catalog identity; F014/F055/F121/F160 remain broader requirements.
