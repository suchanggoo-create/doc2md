from pathlib import Path

import pytest

from doc2md.security import safe_join, sanitize_filename


def test_sanitize_filename_basic():
    assert sanitize_filename("a b/c.txt") == "a_b_c.txt"


def test_sanitize_filename_default():
    assert sanitize_filename("") == "file"
    assert sanitize_filename("   ", default="x") == "x"


def test_safe_join_allows_child(tmp_path: Path):
    out = safe_join(tmp_path, "a/b.txt")
    assert out.parent.name == "a"
    assert str(out).startswith(str(tmp_path.resolve()))


def test_safe_join_blocks_traversal(tmp_path: Path):
    with pytest.raises(ValueError):
        safe_join(tmp_path, "../evil.txt")

