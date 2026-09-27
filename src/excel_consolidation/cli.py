"""Command-line entry point.

    python -m excel_consolidation generate    --out data/workbooks [--seed 7]
    python -m excel_consolidation consolidate --input data/workbooks --out data/output
"""

from __future__ import annotations

import argparse
from pathlib import Path

from .consolidator import process_directory
from .generator import generate_all


def _cmd_generate(args: argparse.Namespace) -> None:
    manifests = generate_all(Path(args.out), seed=args.seed)
    print(f"Generated {len(manifests)} synthetic regional workbooks in {args.out} (seed={args.seed}).")


def _cmd_consolidate(args: argparse.Namespace) -> None:
    result = process_directory(Path(args.input), Path(args.out))
    t = result.totals
    print(
        f"Processed {len(result.file_results)} workbooks from {args.input}: "
        f"{t['rows_in']} rows in, {t['rows_out']} in Data, "
        f"{t['rejected']} rejected, {t['duplicates']} duplicates removed."
    )
    print(f"Reconciled: {result.reconciles()}")
    print(f"Wrote master.xlsx (Data / Validation / Change log) to {args.out}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel_consolidation",
        description="Synthetic regional-workbook generator/consolidator demo.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    gen = sub.add_parser("generate", help="Generate the synthetic regional workbook corpus.")
    gen.add_argument("--out", required=True, help="Output directory for generated workbooks.")
    gen.add_argument("--seed", type=int, default=7, help="Random seed (default: 7).")
    gen.set_defaults(func=_cmd_generate)

    con = sub.add_parser("consolidate", help="Consolidate a directory of regional workbooks into master.xlsx.")
    con.add_argument("--input", required=True, help="Directory containing regional workbooks.")
    con.add_argument("--out", required=True, help="Output directory for master.xlsx.")
    con.set_defaults(func=_cmd_consolidate)

    return parser


def main(argv: list[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    main()
