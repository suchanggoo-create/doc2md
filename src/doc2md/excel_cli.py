import argparse
import sys
from pathlib import Path
from typing import List, Optional

from .excel_to_md import write_xlsx_as_markdown
from .logging_utils import configure_logging


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="xlsx2md",
        description="Convert .xlsx to Markdown (HTML <table> with merged-cell support).",
    )
    p.add_argument("input", help="Path to input .xlsx")
    p.add_argument("--out", required=True, help="Output .md file path")
    p.add_argument(
        "--preview-rows",
        type=int,
        default=0,
        help="Rows to render per sheet (0: all rows)",
    )
    p.add_argument("--log-level", default="info", choices=["error", "warn", "info", "debug"], help="Log level")
    return p


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    configure_logging(args.log_level)

    in_path = Path(args.input)
    out_path = Path(args.out)
    if not in_path.exists():
        print(f"xlsx2md: error: input not found: {in_path}", file=sys.stderr)
        return 1
    if in_path.suffix.lower() != ".xlsx":
        print("xlsx2md: error: input must be a .xlsx file", file=sys.stderr)
        return 1

    try:
        write_xlsx_as_markdown(in_path, out_path, preview_rows=int(args.preview_rows))
    except Exception as e:
        print(f"xlsx2md: error: {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

