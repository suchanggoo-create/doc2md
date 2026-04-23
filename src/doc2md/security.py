from __future__ import annotations

import os
import re
from pathlib import Path


_BAD_CHARS_RE = re.compile(r"[^A-Za-z0-9._-]+")


def sanitize_filename(name: str, *, default: str = "file", max_len: int = 80) -> str:
    base = (name or "").strip()
    if not base:
        base = default
    base = base.replace(os.sep, "_")
    if os.altsep:
        base = base.replace(os.altsep, "_")
    base = base.replace("..", "_")
    base = _BAD_CHARS_RE.sub("_", base)
    base = base.strip("._-") or default
    if len(base) > max_len:
        base = base[:max_len].rstrip("._-") or default
    return base


def safe_join(root: Path, relative: str) -> Path:
    """
    Join and resolve a relative path under root; prevent Zip Slip / traversal.
    """
    if not relative:
        raise ValueError("empty relative path")
    # Reject absolute paths and any traversal attempts.
    if relative.startswith(("/", "\\")) or (os.path.isabs(relative)):
        raise ValueError(f"unsafe absolute path: {relative!r}")
    rel_path = Path(relative)
    if any(part == ".." for part in rel_path.parts):
        raise ValueError(f"unsafe path traversal: {relative!r}")
    rel = str(rel_path)
    out = (root / rel).resolve()
    root_resolved = root.resolve()
    if out == root_resolved or root_resolved in out.parents:
        return out
    raise ValueError(f"unsafe path traversal: {relative!r}")


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)
