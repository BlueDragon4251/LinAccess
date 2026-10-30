#!/usr/bin/env bash
# Reconstruct the exact two MDB files and two Bridge ZIPs from the private fixture bundle.
set -euo pipefail

project_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$project_root"

shopt -s nullglob
parts=(fixtures/archive/part-*.b64)
if (( ${#parts[@]} == 0 )); then
    echo 'Missing fixtures/archive/part-*.b64' >&2
    exit 1
fi

mkdir -p fixtures/raw fixtures/unpacked
# Archive extraction never touches source files elsewhere in the repository.
cat "${parts[@]}" | base64 --decode | tar -xJf - -C fixtures
(cd fixtures && sha256sum -c SHA256SUMS)

unzip -oq fixtures/raw/Keksrezepte-Bridge.zip -d fixtures/unpacked
unzip -oq fixtures/raw/Kassa_CSE_V2.0-Bridge-2003.zip -d fixtures/unpacked
echo 'Restored both MDB files and both Bridge packages with verified SHA-256.'
