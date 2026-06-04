#!/usr/bin/env python3
"""Run IMF-based debt sustainability workbook generation."""

from __future__ import annotations

import argparse
import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
SRC_PATH = os.path.join(ROOT, "src")
if SRC_PATH not in sys.path:
    sys.path.insert(0, SRC_PATH)

from pbs_imf.pipeline import PipelineInputs, run_pipeline


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate debt sustainability workbook for EZ/FR/DE/IT/ES/US/UK using IMF API."
    )
    parser.add_argument(
        "--start-year",
        type=int,
        default=2020,
        help="First visible year in report (default: 2020)",
    )
    parser.add_argument(
        "--end-year",
        type=int,
        default=2030,
        help="Last year in report (default: 2030)",
    )
    parser.add_argument(
        "--output",
        default="pbs_imf_output.xlsx",
        help="Output workbook path (default: pbs_imf_output.xlsx)",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    out_path = run_pipeline(
        PipelineInputs(start_year=args.start_year, end_year=args.end_year, output_path=args.output)
    )
    print(f"Workbook generated: {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
