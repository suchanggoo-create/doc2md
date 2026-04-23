from __future__ import annotations

from pathlib import Path

from .excel_to_md import write_xlsx_as_markdown


def render_xlsx_to_markdown(xlsx_path: Path, md_path: Path, *, preview_rows: int = 40) -> None:
    # Prefer HTML <table> rendering with merged-cell support.
    write_xlsx_as_markdown(xlsx_path, md_path, preview_rows=preview_rows)
