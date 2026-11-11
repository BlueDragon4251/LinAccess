# Access single-target DELETE star syntax

DELETE * FROM a physical table uses the existing DELETE action path. The compiler
removes only the single-target star directly before FROM; checked action parsing
accepts the same source spelling and retains known field kinds, aliases and
criteria. Ordinary DELETE FROM remains supported. Joined/multi-target delete
variants are outside this bounded scope.

Existing DAO scope covers a typed aliased Date/Boolean/Single predicate and an
ordinary text-table criterion. App/helper built; existing DAO/model/import bundle3/3 in6.86s
(5.96/0.43/0.46) passed, including typed alias and text criteria. Fixed456/386/70, F067 open.
