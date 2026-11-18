#!/usr/bin/env bash
# Generate public technical fixtures outside the checkout, run tests, then remove them.
set -euo pipefail

if (($# == 0)); then
    echo "Usage: $0 command [arguments...]" >&2
    exit 2
fi

script_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
scratch=$(mktemp -d /tmp/linaccess-currency.XXXXXXXX)
trap 'rm -rf -- "$scratch"' EXIT
trap 'exit 130' INT
trap 'exit 143' TERM

# Jackcess 5.0.4 has no mandatory runtime dependencies for these MONEY fixtures.
# Pin both the Maven Central URL and the downloaded bytes before executing Java.
jar="$scratch/jackcess-5.0.4.jar"
curl --fail --silent --show-error --location --retry 3 --connect-timeout 30 --max-time 180 \
    --proto '=https' --proto-redir '=https' \
    https://repo.maven.apache.org/maven2/com/healthmarketscience/jackcess/jackcess/5.0.4/jackcess-5.0.4.jar \
    --output "$jar"
if ! printf '%s  %s\n' '19a7494a56f439a57fda85c059b019b665802085cec5ae8186d1e028cc1bdfbc' "$jar" | sha256sum --check --status; then
    echo 'Jackcess SHA256 verification failed; refusing to generate fixtures.' >&2
    exit 1
fi

mkdir "$scratch/classes" "$scratch/fixtures"
javac --release 11 -classpath "$jar" -d "$scratch/classes" \
    "$script_dir/../tests/jet_binary_reference/JetCurrencyReference.java"
java -Duser.timezone=UTC -classpath "$scratch/classes:$jar" JetCurrencyReference "$scratch/fixtures"

export LINACCESS_JET_CURRENCY_REFERENCE="$scratch/fixtures/currency-probe.mdb"
export LINACCESS_ACE_CURRENCY_REFERENCE="$scratch/fixtures/currency-probe.accdb"
test -s "$LINACCESS_JET_CURRENCY_REFERENCE"
test -s "$LINACCESS_ACE_CURRENCY_REFERENCE"
echo 'Running tests with independently verified Jet4 and ACE2010 MONEY fixtures.'
"$@"
