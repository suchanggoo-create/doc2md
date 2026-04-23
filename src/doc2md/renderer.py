from __future__ import annotations

import html
import logging
import re
import io
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from typing import Optional, Tuple

from lxml import etree

from .docx_package import DocxPackage, parse_rels
from .security import ensure_dir, safe_join, sanitize_filename
from .xlsx_render import render_xlsx_to_markdown

log = logging.getLogger(__name__)

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "o": "urn:schemas-microsoft-com:office:office",
}


@dataclass
class _Ctx:
    out_dir: Path
    images_dir: Path
    attachments_dir: Path
    opts: object
    rels: dict[str, dict[str, str]]
    image_seq: int = 0
    att_seq: int = 0


def _md_escape(text: str) -> str:
    text = text.replace("\\", "\\\\")
    text = text.replace("|", "\\|")
    return text


def _apply_styles(text: str, *, bold: bool, italic: bool, strike: bool, underline: bool) -> str:
    out = text
    # Markdown has no underline; HTML fallback.
    if underline:
        out = f"<u>{html.escape(out)}</u>" if ("<" in out or "&" in out) else f"<u>{out}</u>"

    # Bold/italic first, then strike as outer wrapper (keeps consistent)
    if bold and italic:
        out = f"***{out}***"
    elif bold:
        out = f"**{out}**"
    elif italic:
        out = f"*{out}*"

    if strike:
        out = f"~~{out}~~"
    return out


def _p_style(p: etree._Element) -> Optional[str]:
    ppr = p.find("w:pPr", namespaces=NS)
    if ppr is None:
        return None
    ps = ppr.find("w:pStyle", namespaces=NS)
    if ps is None:
        return None
    return ps.get(f"{{{NS['w']}}}val") or ps.get("val")


def _p_num(p: etree._Element) -> Optional[Tuple[int, str]]:
    ppr = p.find("w:pPr", namespaces=NS)
    if ppr is None:
        return None
    numpr = ppr.find("w:numPr", namespaces=NS)
    if numpr is None:
        return None
    ilvl = numpr.find("w:ilvl", namespaces=NS)
    level = int((ilvl.get(f"{{{NS['w']}}}val") or ilvl.get("val") or "0")) if ilvl is not None else 0
    # Best-effort: without numbering.xml we can't reliably tell ol vs ul.
    return (level, "ul")


def _iter_body_blocks(body: etree._Element) -> Iterable[etree._Element]:
    for child in body:
        if not isinstance(child.tag, str):
            continue
        local = etree.QName(child).localname
        # keep order for unknown blocks too
        yield child


def _read_rels(pkg: DocxPackage) -> dict[str, dict[str, str]]:
    if not pkg.exists("word/_rels/document.xml.rels"):
        return {}
    rels_xml = pkg.read_xml("word/_rels/document.xml.rels")
    return parse_rels(rels_xml)


def _find_image_rid(drawing: etree._Element) -> Optional[str]:
    blip = drawing.find(".//a:blip", namespaces=NS)
    if blip is None:
        return None
    return blip.get(f"{{{NS['r']}}}embed")


def _get_drawing_extent_px(drawing: etree._Element) -> Optional[Tuple[int, int]]:
    extent = drawing.find(".//wp:extent", namespaces=NS)
    if extent is None:
        return None
    cx = extent.get("cx")
    cy = extent.get("cy")
    if not cx or not cy:
        return None
    try:
        emu_x = int(cx)
        emu_y = int(cy)
    except ValueError:
        return None
    # 1px ~ 9525 EMU at 96dpi
    return (max(1, emu_x // 9525), max(1, emu_y // 9525))


def _extract_image(pkg: DocxPackage, ctx: _Ctx, rid: str) -> str:
    rel = ctx.rels.get(rid)
    if not rel:
        raise ValueError(f"missing image relationship: {rid}")
    target = rel.get("Target", "")
    if not target:
        raise ValueError(f"empty image relationship target: {rid}")

    part = DocxPackage.normalize_target("word/document.xml", target)
    data = pkg.read_bytes(part)
    ext = (Path(part).suffix or ".bin").lstrip(".")

    ctx.image_seq += 1
    out_name = f"img_{ctx.image_seq:04d}.{ext}"
    out_path = safe_join(ctx.images_dir, out_name)
    out_path.write_bytes(data)
    return f"assets/images/{out_name}"


def _extract_attachment(pkg: DocxPackage, ctx: _Ctx, rid: str, *, hint_name: Optional[str]) -> tuple[str, str, bytes]:
    rel = ctx.rels.get(rid)
    if not rel:
        raise ValueError(f"missing attachment relationship: {rid}")
    target = rel.get("Target", "")
    if not target:
        raise ValueError(f"empty attachment relationship target: {rid}")

    part = DocxPackage.normalize_target("word/document.xml", target)
    data = pkg.read_bytes(part, max_bytes=int(getattr(ctx.opts, "max_embed_bytes", 50 * 1024 * 1024)) * 2)

    # IMPORTANT: hint_name is often a ProgID like "Excel.Sheet.12" (not a real filename).
    # Prefer the embedded part's name and infer type by magic bytes/zip contents.
    part_name = sanitize_filename(Path(part).name, default="attachment")
    suffix = Path(part_name).suffix.lower()
    ext: Optional[str] = suffix.lstrip(".") if suffix else None
    safe_base = part_name[: -len(suffix)] if suffix else part_name

    # If ext is missing or suspicious (ProgID-ish numeric suffix like ".12"), infer.
    if ext is None or ext.isdigit() or len(ext) > 8:
        ext = None

    if ext is None:
        if data.startswith(b"PK\x03\x04"):
            # Determine docx vs xlsx by inspecting internal parts.
            try:
                with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
                    names = set(zf.namelist())
                if "xl/workbook.xml" in names:
                    ext = "xlsx"
                elif "word/document.xml" in names:
                    ext = "docx"
                else:
                    ext = "zip"
            except Exception:
                ext = "zip"
        elif data.startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
            # OLE2 container (legacy doc/xls/ppt) or other; keep as bin for MVP.
            ext = "bin"
        else:
            ext = "bin"

    ctx.att_seq += 1
    out_file = f"att_{ctx.att_seq:04d}_{sanitize_filename(safe_base)}.{ext}"
    out_path = safe_join(ctx.attachments_dir, out_file)
    out_path.write_bytes(data)
    return f"assets/attachments/{out_file}", out_file, data


def _maybe_render_embedded_inline(ctx: _Ctx, attachment_path: Path, attachment_rel_md: str, data: bytes, *, depth: int) -> str | None:
    if not getattr(ctx.opts, "render_embedded", True):
        return None
    if depth + 1 > int(getattr(ctx.opts, "max_embed_depth", 2)):
        return None
    max_bytes = int(getattr(ctx.opts, "max_embed_bytes", 50 * 1024 * 1024))
    if len(data) > max_bytes:
        return f"> [EMBED_RENDER_SKIPPED: size {len(data)} > max_embed_bytes]"

    if not data.startswith(b"PK\x03\x04"):
        return None

    try:
        with zipfile.ZipFile(attachment_path, "r") as zf:
            names = set(zf.namelist())
    except Exception as e:
        return f"> [EMBED_RENDER_FAILED: zip inspect failed: {e}]"

    is_docx = "word/document.xml" in names
    is_xlsx = "xl/workbook.xml" in names

    if is_docx:
        subdir = ctx.attachments_dir / f"{Path(attachment_rel_md).stem}_rendered"
        ensure_dir(subdir)
        try:
            pkg2 = DocxPackage(str(attachment_path))
            try:
                render_docx_to_markdown(pkg2, subdir, ctx.opts, depth=depth + 1)
            finally:
                pkg2.close()
            embedded_md = (subdir / "index.md").read_text(encoding="utf-8")
            title = f"> [EMBEDDED_DOCX: {Path(attachment_rel_md).name}]"
            return "\n".join([title, "", embedded_md.strip()])
        except Exception as e:
            return f"> [EMBED_RENDER_FAILED: docx: {e}]"

    if is_xlsx:
        md_name = f"{Path(attachment_rel_md).stem}.md"
        md_path = safe_join(ctx.attachments_dir, md_name)
        try:
            render_xlsx_to_markdown(
                attachment_path,
                md_path,
                preview_rows=int(getattr(ctx.opts, "excel_preview_rows", 20)),
            )
            embedded_md = md_path.read_text(encoding="utf-8")
            title = f"> [EMBEDDED_XLSX: {Path(attachment_rel_md).name}]"
            return "\n".join([title, "", embedded_md.strip()])
        except Exception as e:
            return f"> [EMBED_RENDER_FAILED: xlsx: {e}]"

    return None


def _run_text(run: etree._Element) -> str:
    parts: list[str] = []
    for node in run:
        if not isinstance(node.tag, str):
            continue
        local = etree.QName(node).localname
        if local == "t":
            parts.append(node.text or "")
        elif local == "tab":
            parts.append("\t")
        elif local in ("br", "cr"):
            parts.append("\n")
    return "".join(parts)


def _render_runs_inline(pkg: DocxPackage, ctx: _Ctx, node: etree._Element, *, depth: int) -> str:
    chunks: list[str] = []

    def emit(text: str) -> None:
        if text:
            chunks.append(text)

    for child in node:
        if not isinstance(child.tag, str):
            continue
        local = etree.QName(child).localname

        if local == "hyperlink":
            rid = child.get(f"{{{NS['r']}}}id")
            url = None
            if rid and rid in ctx.rels and ctx.rels[rid].get("TargetMode") == "External":
                url = ctx.rels[rid].get("Target")
            inner = _render_runs_inline(pkg, ctx, child, depth=depth)
            if url:
                emit(f"[{inner}]({url})")
            else:
                emit(inner)
            continue

        if local != "r":
            continue

        rpr = child.find("w:rPr", namespaces=NS)
        bold = italic = strike = underline = False
        if rpr is not None:
            bold = rpr.find("w:b", namespaces=NS) is not None
            italic = rpr.find("w:i", namespaces=NS) is not None
            underline = rpr.find("w:u", namespaces=NS) is not None
            strike = (rpr.find("w:strike", namespaces=NS) is not None) or (rpr.find("w:dstrike", namespaces=NS) is not None)

        drawing = child.find("w:drawing", namespaces=NS)
        if drawing is not None:
            try:
                rid = _find_image_rid(drawing)
                if not rid:
                    emit("> [UNSUPPORTED: drawing_without_rId]")
                    continue
                img_rel = _extract_image(pkg, ctx, rid)
                if getattr(ctx.opts, "keep_image_size", False):
                    size = _get_drawing_extent_px(drawing)
                    if size:
                        wpx, hpx = size
                        emit(f'<img src="{img_rel}" width="{wpx}" height="{hpx}" />')
                    else:
                        emit(f"![]({img_rel})")
                else:
                    emit(f"![]({img_rel})")
            except Exception as e:
                log.warning("image render failed: %s", e)
                emit(f"> [UNSUPPORTED: image_render_failed: {e}]")
            continue

        # OLE attachments (best-effort): find o:OLEObject and read r:id or r:embed
        ole = child.find(".//o:OLEObject", namespaces=NS)
        embed_rid = None
        hint_name = None
        if ole is not None:
            embed_rid = ole.get(f"{{{NS['r']}}}id") or ole.get(f"{{{NS['r']}}}embed")
            hint_name = ole.get("ProgID") or ole.get("ObjectID")
        if embed_rid:
            try:
                att_rel, att_file, data = _extract_attachment(pkg, ctx, embed_rid, hint_name=hint_name)
                inline = _maybe_render_embedded_inline(ctx, ctx.attachments_dir / att_file, att_rel, data, depth=depth)
                if inline:
                    emit("\n\n" + inline + "\n\n")
                else:
                    # Only fall back to a link when we can't render it as text.
                    emit(f"[附件: {att_file}]({att_rel})")
            except Exception as e:
                log.warning("attachment extract failed: %s", e)
                emit(f"> [ATTACHMENT_EXTRACT_FAILED: {e}]")
            continue

        text = _run_text(child)
        if text:
            emit(_apply_styles(text, bold=bold, italic=italic, strike=strike, underline=underline))

    return "".join(chunks)


def _render_table_cell_text(pkg: DocxPackage, ctx: _Ctx, tc: etree._Element, *, depth: int) -> str:
    parts: list[str] = []
    for p in tc.findall("./w:p", namespaces=NS):
        parts.append(_render_runs_inline(pkg, ctx, p, depth=depth).strip())
    return _md_escape(" ".join([p for p in parts if p]))


def _table_has_merge(tbl: etree._Element) -> bool:
    for tcpr in tbl.findall(".//w:tcPr", namespaces=NS):
        if tcpr.find("w:gridSpan", namespaces=NS) is not None:
            return True
        if tcpr.find("w:vMerge", namespaces=NS) is not None:
            return True
    return False


def _render_tbl(pkg: DocxPackage, ctx: _Ctx, tbl: etree._Element, *, depth: int) -> list[str]:
    if _table_has_merge(tbl):
        lines = ["<table>"]
        for tr in tbl.findall("./w:tr", namespaces=NS):
            lines.append("  <tr>")
            for tc in tr.findall("./w:tc", namespaces=NS):
                text = html.escape(_render_table_cell_text(pkg, ctx, tc, depth=depth))
                lines.append(f"    <td>{text}</td>")
            lines.append("  </tr>")
        lines.append("</table>")
        return lines

    rows: list[list[str]] = []
    for tr in tbl.findall("./w:tr", namespaces=NS):
        row: list[str] = []
        for tc in tr.findall("./w:tc", namespaces=NS):
            row.append(_render_table_cell_text(pkg, ctx, tc, depth=depth))
        rows.append(row)

    if not rows:
        return ["> [UNSUPPORTED: empty_table]"]

    width = max(len(r) for r in rows)
    for r in rows:
        while len(r) < width:
            r.append("")

    header = rows[0]
    lines = ["| " + " | ".join(header) + " |", "| " + " | ".join(["---"] * width) + " |"]
    for r in rows[1:]:
        lines.append("| " + " | ".join(r) + " |")
    return lines


def _render_paragraph(pkg: DocxPackage, ctx: _Ctx, p: etree._Element, *, depth: int) -> list[str]:
    style = _p_style(p) or ""
    text = _render_runs_inline(pkg, ctx, p, depth=depth).rstrip()

    num = _p_num(p)
    if num:
        level, kind = num
        indent = "  " * level
        marker = "1." if kind == "ol" else "-"
        return [f"{indent}{marker} {text}".rstrip()]

    m = re.fullmatch(r"Heading([1-6])", style or "")
    if m:
        n = int(m.group(1))
        return [("#" * n) + " " + text]

    return [text] if text else [""]


def render_docx_to_markdown(pkg: DocxPackage, out_dir: Path, opts: object, *, depth: int) -> None:
    ensure_dir(out_dir)
    ensure_dir(out_dir / "assets" / "images")
    ensure_dir(out_dir / "assets" / "attachments")

    doc = pkg.read_xml("word/document.xml")
    body = doc.find(".//w:body", namespaces=NS)
    if body is None:
        raise ValueError("missing w:body in word/document.xml")

    rels = _read_rels(pkg)
    ctx = _Ctx(
        out_dir=out_dir,
        images_dir=out_dir / "assets" / "images",
        attachments_dir=out_dir / "assets" / "attachments",
        opts=opts,
        rels=rels,
    )

    lines_out: list[str] = []
    if getattr(opts, "toc", False):
        lines_out.append("> [TOC_NOT_IMPLEMENTED_MVP]")
        lines_out.append("")

    for block in _iter_body_blocks(body):
        try:
            local = etree.QName(block).localname if isinstance(block.tag, str) else "unknown"
            if local == "p":
                lines_out.extend(_render_paragraph(pkg, ctx, block, depth=depth))
                lines_out.append("")
            elif local == "tbl":
                lines_out.extend(_render_tbl(pkg, ctx, block, depth=depth))
                lines_out.append("")
            else:
                lines_out.append(f"> [UNSUPPORTED: {local}]")
                lines_out.append("")
        except Exception as e:
            log.warning("block render failed (%s): %s", getattr(block, "tag", "unknown"), e)
            lines_out.append(f"> [UNSUPPORTED: block_render_failed: {e}]")
            lines_out.append("")

    (out_dir / "index.md").write_text("\n".join(lines_out).rstrip() + "\n", encoding="utf-8")
