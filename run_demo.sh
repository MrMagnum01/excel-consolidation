#!/usr/bin/env bash
# One-command demo run: generate the messy synthetic regional workbooks,
# then consolidate them into master.xlsx. Assumes the venv is already
# created and activated (see README "Setup"), or falls back to the system
# python3 if not.
set -euo pipefail

cd "$(dirname "${BASH_SOURCE[0]}")"
export PYTHONPATH="src${PYTHONPATH:+:$PYTHONPATH}"

OUT_WORKBOOKS="${1:-data/workbooks}"
OUT_MASTER="${2:-data/output}"

python3 -m excel_consolidation generate --out "$OUT_WORKBOOKS" --seed 7
python3 -m excel_consolidation consolidate --input "$OUT_WORKBOOKS" --out "$OUT_MASTER"

echo
echo "Done. master.xlsx written to $OUT_MASTER/ (Data / Validation / Change log sheets)."
