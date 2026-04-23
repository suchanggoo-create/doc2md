from __future__ import annotations

import logging
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Dict, Any

from .docx_package import DocxPackage, ZipReadLimits
from .renderer import render_docx_to_markdown
from .security import ensure_dir

log = logging.getLogger(__name__)


@dataclass(frozen=True)
class ConvertOptions:
    force: bool = False
    render_embedded: bool = True
    max_embed_depth: int = 2
    max_embed_bytes: int = 50 * 1024 * 1024
    excel_preview_rows: int = 0
    keep_image_size: bool = False
    toc: bool = False


def _normalize_options(options: Optional[Dict[str, Any]]) -> ConvertOptions:
    if not options:
        return ConvertOptions()
    return ConvertOptions(
        force=bool(options.get("force", False)),
        render_embedded=bool(options.get("render_embedded", True)),
        max_embed_depth=int(options.get("max_embed_depth", 2)),
        max_embed_bytes=int(options.get("max_embed_bytes", 50 * 1024 * 1024)),
        excel_preview_rows=int(options.get("excel_preview_rows", 0)),
        keep_image_size=bool(options.get("keep_image_size", False)),
        toc=bool(options.get("toc", False)),
    )


def convertDocxToMd(inputPath: str, outDir: str, options: Optional[Dict[str, Any]] = None) -> None:
    opts = _normalize_options(options)
    out_dir = Path(outDir)
 
    if out_dir.exists():
        if not opts.force:
            raise ValueError(f"output directory exists (use --force to overwrite): {out_dir}")
        shutil.rmtree(out_dir)
 
    ensure_dir(out_dir)
    ensure_dir(out_dir / "assets" / "images")
    ensure_dir(out_dir / "assets" / "attachments")
 
    pkg = DocxPackage(inputPath, limits=ZipReadLimits())
    try:
        render_docx_to_markdown(pkg, out_dir, opts, depth=0)
    finally:
        pkg.close()
