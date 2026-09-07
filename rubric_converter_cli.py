"""Command-line interface for rubric conversion."""

from __future__ import annotations

import argparse
from pathlib import Path

from rubric_converter import RubricConversionError, convert_rubric


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert Turnitin .rbc rubric files.")
    parser.add_argument("input", type=Path, help="Input .rbc file, or directory with --batch")
    parser.add_argument("output", type=Path, help="Output .csv/.xlsx file, or directory with --batch")
    parser.add_argument("--format", choices=("csv", "excel"), help="Output format (inferred by default)")
    parser.add_argument("--batch", action="store_true", help="Convert every .rbc file in the input directory")
    parser.add_argument(
        "--use-name-and-value",
        action="store_true",
        help="Use criterion name and point value instead of its description in the matrix.",
    )
    args = parser.parse_args()
    try:
        if args.batch:
            if not args.input.is_dir():
                raise RubricConversionError(f"Batch input directory does not exist: {args.input}")
            args.output.mkdir(parents=True, exist_ok=True)
            inputs = sorted(args.input.glob("*.rbc"))
            if not inputs:
                raise RubricConversionError(f"No .rbc files found in {args.input}")
            for source in inputs:
                output_format = args.format or "csv"
                suffix = ".xlsx" if output_format == "excel" else ".csv"
                convert_rubric(
                    source, args.output / f"{source.stem}{suffix}", output_format, args.use_name_and_value
                )
            print(f"Converted {len(inputs)} rubrics to {args.output}")
            return 0
        output_format = args.format or ("excel" if args.output.suffix.lower() == ".xlsx" else "csv")
        convert_rubric(args.input, args.output, output_format, args.use_name_and_value)
    except RubricConversionError as exc:
        parser.error(str(exc))
    print(f"Converted rubric saved to {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
