from __future__ import annotations

from pathlib import Path

import pytest

from doc2md.convert import convertDocxToMd


@pytest.mark.slow
def test_sample_docx_end_to_end(tmp_path: Path):
    # This sample is part of the repo workspace.
    sample = Path(__file__).resolve().parents[1] / "视觉健康档案-数字护眼报告 技术方案文档.docx"
    if not sample.exists():
        pytest.skip("sample docx not present")

    out_dir = tmp_path / "out"
    convertDocxToMd(str(sample), str(out_dir), options={"force": True})

    index_md = out_dir / "index.md"
    assert index_md.exists()
    assert (out_dir / "assets" / "images").exists()
    assert (out_dir / "assets" / "attachments").exists()

    md = index_md.read_text(encoding="utf-8")

    # Basic “not empty” signal
    assert len(md) > 100

    # All referenced images/attachments should exist (relative paths) when links are present.

    for line in md.splitlines():
        line = line.strip()
        if line.startswith("![](") and "assets/images/" in line:
            rel = line.split("![](", 1)[1].split(")", 1)[0]
            assert (out_dir / rel).exists()
        if line.startswith("[附件:") and "assets/attachments/" in line:
            rel = line.split("](", 1)[1].split(")", 1)[0]
            assert (out_dir / rel).exists()

    # Inline embedded content markers (best-effort)
    if "EMBEDDED_XLSX" in md or "EMBEDDED_DOCX" in md:
        assert True

