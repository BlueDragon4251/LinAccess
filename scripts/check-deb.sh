#!/usr/bin/env bash
# Check package permissions and ensure no test corpus or source files are shipped.
set -euo pipefail
if (( $# < 1 || $# > 2 )); then
    echo 'Usage: check-deb.sh path/to/linaccess.deb [reference-build-binary]' >&2
    exit 2
fi
listing=$(mktemp)
staging=$(mktemp -d)
trap 'rm -f "$listing"; rm -rf "$staging"' EXIT
dpkg-deb --contents "$1" > "$listing"
awk '
    $1 ~ /^d/ {
        if ($1 != "drwxr-xr-x") exit 1
        if ($6 != "./usr/" && $6 != "./usr/bin/" && $6 != "./usr/lib/" && $6 != "./usr/lib/linaccess/" && $6 != "./usr/share/" && $6 != "./usr/share/applications/" && $6 != "./usr/share/doc/" && $6 != "./usr/share/doc/linaccess/") exit 1
        next
    }
    $6 == "./usr/bin/linaccess" { if ($1 != "-rwxr-xr-x" || $3 == 0) exit 1; app++; next }
    $6 == "./usr/lib/linaccess/linaccess-mdb-metadata" { if ($1 != "-rwxr-xr-x" || $3 == 0) exit 1; metadata++; next }
    $6 == "./usr/share/applications/linaccess.desktop" { if ($1 != "-rw-r--r--" || $3 == 0) exit 1; desktop++; next }
    $6 == "./usr/lib/linaccess/libsqlite3.so.0" { if ($1 != "-rw-r--r--" && $1 != "-rwxr-xr-x") exit 1; if ($3 == 0) exit 1; sqlite++; next }
    $6 == "./usr/share/doc/linaccess/sqlite.txt" { if ($1 != "-rw-r--r--" || $3 == 0) exit 1; notice++; next }
    { exit 1 }
    END { if (app != 1 || desktop != 1 || sqlite != 1 || notice != 1 || metadata != 1) exit 1 }
' "$listing" || { echo 'Package content or permissions check failed' >&2; exit 1; }
if (( $# == 2 )); then
    dpkg-deb --extract "$1" "$staging"
    reference_id=$(LC_ALL=C readelf -n "$2" | awk '$1 == "Build" && $2 == "ID:" { print $3 }')
    package_id=$(LC_ALL=C readelf -n "$staging/usr/bin/linaccess" | awk '$1 == "Build" && $2 == "ID:" { print $3 }')
    if [[ -z "$reference_id" || "$package_id" != "$reference_id" ]]; then
        echo 'Package binary does not match the reference build ID' >&2
        exit 1
    fi
    reference_metadata="$(dirname "$2")/linaccess-mdb-metadata"
    reference_metadata_id=$(LC_ALL=C readelf -n "$reference_metadata" | awk '$1 == "Build" && $2 == "ID:" { print $3 }')
    package_metadata_id=$(LC_ALL=C readelf -n "$staging/usr/lib/linaccess/linaccess-mdb-metadata" | awk '$1 == "Build" && $2 == "ID:" { print $3 }')
    if [[ -z "$reference_metadata_id" || "$reference_metadata_id" != "$package_metadata_id" ]]; then
        echo 'Package metadata helper does not match the reference build ID' >&2
        exit 1
    fi
    "$staging/usr/lib/linaccess/linaccess-mdb-metadata" --version
    reference_sqlite="$(dirname "$2")/third_party/sqlite/libsqlite3.so.0"
    if ! cmp -s "$reference_sqlite" "$staging/usr/lib/linaccess/libsqlite3.so.0"; then
        echo 'Package SQLite backend does not match the reference build' >&2
        exit 1
    fi
    "$staging/usr/bin/linaccess" --check-sql-backend
    echo 'Package binary matches the reference build ID'
fi
echo 'Package checked: executable, desktop entry, public directory permissions, no fixtures'
