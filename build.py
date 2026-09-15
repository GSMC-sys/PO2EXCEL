#!/usr/bin/env python3
"""Builds index.html from src/app_template.html + src/parsers.js + vendored
libraries. Never hand-edit index.html -- edit src/* and re-run this script.

Usage: python3 build.py
"""
import base64
import pathlib
import sys

ROOT = pathlib.Path(__file__).parent
SRC = ROOT / "src"

TEMPLATE = SRC / "app_template.html"
PARSERS = SRC / "parsers.js"
XLSX_LIB = ROOT / "node_modules" / "xlsx" / "dist" / "xlsx.full.min.js"
PDFJS_LIB = ROOT / "node_modules" / "pdfjs-dist" / "build" / "pdf.min.js"
PDFJS_WORKER = ROOT / "node_modules" / "pdfjs-dist" / "build" / "pdf.worker.min.js"
OUT = ROOT / "index.html"


def read(path):
    if not path.exists():
        sys.exit(f"ERROR: missing required file: {path}")
    return path.read_text(encoding="utf-8")


def strip_sourcemap_comment(text):
    # Avoids a harmless-but-confusing 404 in DevTools for a .map file that
    # doesn't ship next to the built index.html.
    import re
    return re.sub(r"//[#@]\s*sourceMappingURL=.*", "", text)


def assert_safe_for_script_tag(name, text):
    if "</script" in text.lower():
        sys.exit(
            f"ERROR: {name} contains a literal '</script' sequence, which would "
            "break out of its <script> tag when inlined. Refusing to build."
        )


def main():
    template = read(TEMPLATE)
    parsers_js = read(PARSERS)
    xlsx_js = strip_sourcemap_comment(read(XLSX_LIB))
    pdfjs_js = strip_sourcemap_comment(read(PDFJS_LIB))
    worker_bytes = PDFJS_WORKER.read_bytes() if PDFJS_WORKER.exists() else sys.exit(
        f"ERROR: missing required file: {PDFJS_WORKER}"
    )
    worker_b64 = base64.b64encode(worker_bytes).decode("ascii")

    for name, text in [("parsers.js", parsers_js), ("xlsx.full.min.js", xlsx_js), ("pdf.min.js", pdfjs_js)]:
        assert_safe_for_script_tag(name, text)

    out = template
    out = out.replace("//__XLSX_LIB__//", xlsx_js)
    out = out.replace("//__PDFJS_LIB__//", pdfjs_js)
    out = out.replace("__PDFJS_WORKER_B64__", worker_b64)
    out = out.replace("//__PARSERS__//", parsers_js)

    remaining_markers = [m for m in ("__XLSX_LIB__", "__PDFJS_LIB__", "__PDFJS_WORKER_B64__", "__PARSERS__") if m in out]
    if remaining_markers:
        sys.exit(f"ERROR: build markers not replaced: {remaining_markers}")

    OUT.write_text(out, encoding="utf-8")
    size_mb = OUT.stat().st_size / (1024 * 1024)
    print(f"Built {OUT} ({size_mb:.1f} MB)")


if __name__ == "__main__":
    main()
