#!/usr/bin/env bash
# Check package permissions and ensure no test corpus or source files are shipped.
set -euo pipefail
if (( $# != 1 )); then
    echo 'Usage: check-deb.sh path/to/linaccess.deb' >&2
    exit 2
fi
listing=$(mktemp)
trap 'rm -f "$listing"' EXIT
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
echo 'Package checked: executable, desktop entry, public directory permissions, no fixtures'
