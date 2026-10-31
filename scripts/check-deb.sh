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
        if ($6 != "./usr/" && $6 != "./usr/bin/" && $6 != "./usr/share/" && $6 != "./usr/share/applications/") exit 1
        next
    }
    $6 == "./usr/bin/linaccess" { if ($1 != "-rwxr-xr-x" || $3 == 0) exit 1; app++; next }
    $6 == "./usr/share/applications/linaccess.desktop" { if ($1 != "-rw-r--r--" || $3 == 0) exit 1; desktop++; next }
    { exit 1 }
    END { if (app != 1 || desktop != 1) exit 1 }
' "$listing" || { echo 'Package content or permissions check failed' >&2; exit 1; }
if (( $# == 2 )); then
    dpkg-deb --extract "$1" "$staging"
    reference_id=$(LC_ALL=C readelf -n "$2" | awk '$1 == "Build" && $2 == "ID:" { print $3 }')
    package_id=$(LC_ALL=C readelf -n "$staging/usr/bin/linaccess" | awk '$1 == "Build" && $2 == "ID:" { print $3 }')
    if [[ -z "$reference_id" || "$package_id" != "$reference_id" ]]; then
        echo 'Package binary does not match the reference build ID' >&2
        exit 1
    fi
    echo 'Package binary matches the reference build ID'
fi
echo 'Package checked: executable, desktop entry, public directory permissions, no fixtures'
