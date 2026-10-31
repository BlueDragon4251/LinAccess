#!/usr/bin/env bash
# Dependency metadata only: never accepts or opens a database.
set -euo pipefail
for tool in mdb-schema mdb-export mdb-sql; do
    executable=$(command -v "$tool")
    "$executable" --version
    sha256sum "$executable"
    while IFS= read -r library; do
        [[ -r "$library" ]] || { echo 'Resolved MDB library is unreadable' >&2; exit 1; }
        sha256sum "$library"
    done < <(LC_ALL=C ldd "$executable" | awk '$1 ~ /^libmdb(sql)?\.so\./ && $2 == "=>" { print $3 }')
done
dpkg-query -W -f='${binary:Package}\t${Version}\n' | awk '$1 ~ /^(mdbtools|libmdb3|libmdb3t64|libmdbsql3|libmdbsql3t64)(:|$)/'
