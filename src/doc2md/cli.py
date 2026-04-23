import argparse
import sys
from typing import List, Optional

from .convert import convertDocxToMd
from .logging_utils import configure_logging


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="doc2md", description="Convert DOCX to Markdown (images + embedded attachments).")
    p.add_argument("input", help="Path to input .docx")
    p.add_argument("--out", required=True, help="Output directory")
    p.add_argument("--force", action="store_true", help="Allow clearing output directory if it exists")

    p.add_argument("--render-embedded", default="true", choices=["true", "false"], help="Render embedded docx/xlsx if possible")
    p.add_argument("--max-embed-depth", type=int, default=2, help="Max recursive embed render depth")
    p.add_argument("--max-embed-bytes", type=int, default=50 * 1024 * 1024, help="Max bytes to attempt rendering embeds")
    p.add_argument("--excel-preview-rows", type=int, default=0, help="Rows to render for xlsx embeds (0: all rows)")
    p.add_argument("--keep-image-size", action="store_true", help="Keep image size using HTML <img>")
    p.add_argument("--toc", action="store_true", help="Generate a simple TOC at top")
    p.add_argument("--log-level", default="info", choices=["error", "warn", "info", "debug"], help="Log level")
    return p


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    configure_logging(args.log_level)

    render_embedded = args.render_embedded.lower() == "true"
    try:
        convertDocxToMd(
            inputPath=args.input,
            outDir=args.out,
            options={
                "force": bool(args.force),
                "render_embedded": render_embedded,
                "max_embed_depth": int(args.max_embed_depth),
                "max_embed_bytes": int(args.max_embed_bytes),
                "excel_preview_rows": int(args.excel_preview_rows),
                "keep_image_size": bool(args.keep_image_size),
                "toc": bool(args.toc),
            },
        )
    except Exception as e:
        # CLI should be user-friendly; details are in logs at debug level.
        print(f"doc2md: error: {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
