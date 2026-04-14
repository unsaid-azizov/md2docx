#!/usr/bin/env python3
"""
md2docx — convert a Markdown file (with Mermaid diagrams) to .docx

Usage:
    md2docx <input.md> [output_dir]

    input.md    — path to the source Markdown file
    output_dir  — where to write results (default: <input_stem>_output
                  next to the input file)

Output layout:
    <output_dir>/
        <stem>.docx        — final document
        diagrams/          — rendered Mermaid PNG files

Pipeline:
    1. mmdc   — render Mermaid code blocks to PNG (theme neutral, width 900)
    2. merge  — move standalone *Рисунок N — ...* captions into image alt text
    3. pandoc — convert processed markdown to docx
    4. docx   — Times New Roman 12pt, images downscaled only, table borders
"""

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

# Max dimensions — images are ONLY SHRUNK, never enlarged
MAX_WIDTH_CM = 14.0
MAX_WIDTH_EMU = int(MAX_WIDTH_CM / 2.54 * 914400)

MAX_HEIGHT_CM = 20.0   # safely below one A4 page
MAX_HEIGHT_EMU = int(MAX_HEIGHT_CM / 2.54 * 914400)

PUPPETEER_CFG_NAME = "_puppeteer.json"


# ---------------------------------------------------------------------------
# Step 1: render Mermaid diagrams
# ---------------------------------------------------------------------------

def run_mmdc(src_md: Path, rendered_md: Path, diagrams_dir: Path) -> None:
    print("▶ Rendering Mermaid diagrams (theme neutral, width 900)...")
    diagrams_dir.mkdir(parents=True, exist_ok=True)

    # Write puppeteer config for headless Chromium in Docker (no-sandbox)
    cfg_path = diagrams_dir / PUPPETEER_CFG_NAME
    cfg_path.write_text(json.dumps({"args": ["--no-sandbox", "--disable-setuid-sandbox"]}))

    result = subprocess.run(
        [
            "mmdc",
            "-i", str(src_md),
            "-o", str(rendered_md),
            "--outputFormat", "png",
            "--theme", "neutral",
            "--width", "900",
            "--backgroundColor", "white",
            "--puppeteerConfigFile", str(cfg_path),
        ],
        capture_output=True,
        text=True,
    )
    cfg_path.unlink(missing_ok=True)

    if result.stderr.strip():
        print(f"mmdc: {result.stderr.strip()}", file=sys.stderr)
    if not rendered_md.exists():
        print("ERROR: mmdc produced no output", file=sys.stderr)
        sys.exit(1)


# ---------------------------------------------------------------------------
# Step 2: merge captions
# ---------------------------------------------------------------------------

def process_markdown(text: str) -> str:
    """
    Move standalone *Рисунок N — caption* / *Figure N — caption* lines
    that precede an image into the image's alt text so pandoc renders
    them as proper figure captions.
    """
    return re.sub(
        r"\*([^\*\n]+)\*[ \t]*\n+[ \t]*!\[[^\]]*\]\(([^)]+)\)",
        lambda m: f"![{m.group(1).strip()}]({m.group(2)})",
        text,
    )


# ---------------------------------------------------------------------------
# Step 3: pandoc conversion
# ---------------------------------------------------------------------------

def run_pandoc(processed_md: Path, output_docx: Path,
               src_dir: Path, diagrams_dir: Path) -> None:
    print("▶ Converting to docx (pandoc)...")
    resource_path = f"{diagrams_dir}:{src_dir}"
    result = subprocess.run(
        [
            "pandoc",
            str(processed_md),
            "--from", "markdown",
            "--to", "docx",
            "--output", str(output_docx),
            "--resource-path", resource_path,
            "--standalone",
        ],
        capture_output=True,
        text=True,
        cwd=str(src_dir),
    )
    if result.stderr.strip():
        print(result.stderr.strip())
    if result.returncode != 0:
        sys.exit(1)


# ---------------------------------------------------------------------------
# Step 4: python-docx post-processing
# ---------------------------------------------------------------------------

def _set_font(para, name: str = "Times New Roman", size_pt: int = 12) -> None:
    for run in para.runs:
        run.font.name = name
        run.font.size = Pt(size_pt)
        rPr = run._element.get_or_add_rPr()
        rf = rPr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rPr.insert(0, rf)
        for attr in ("w:ascii", "w:hAnsi", "w:cs"):
            rf.set(qn(attr), name)

    pPr = para._element.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rf.set(qn(attr), name)
    half = str(size_pt * 2)
    for tag in ("w:sz", "w:szCs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn("w:val"), half)


def _has_drawing(para) -> bool:
    return bool(para._element.findall(".//" + qn("w:drawing")))


def _cap_images(doc: Document) -> None:
    """Downscale images to fit within max dimensions. Never enlarge."""
    for shape in doc.inline_shapes:
        w, h = shape.width, shape.height
        if not w or not h:
            continue
        scale = min(1.0, MAX_WIDTH_EMU / w, MAX_HEIGHT_EMU / h)
        if scale < 1.0:
            shape.width = int(w * scale)
            shape.height = int(h * scale)


def _add_borders(table) -> None:
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    ex = tblPr.find(qn("w:tblBorders"))
    if ex is not None:
        tblPr.remove(ex)
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tblPr.append(borders)


def postprocess(output_docx: Path) -> None:
    print("▶ Post-processing: font, images, tables...")
    doc = Document(str(output_docx))

    try:
        doc.styles["Normal"].font.name = "Times New Roman"
        doc.styles["Normal"].font.size = Pt(12)
    except Exception:
        pass

    prev_image = False
    for para in doc.paragraphs:
        _set_font(para)
        has_image = _has_drawing(para)
        if has_image or prev_image:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if para.style and "caption" in para.style.name.lower():
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        prev_image = has_image

    _cap_images(doc)

    for table in doc.tables:
        _add_borders(table)
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _set_font(para)

    doc.save(str(output_docx))


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

DOCTOR_TEXT = """
md2docx — Markdown + Mermaid → .docx

USAGE
  md2docx <input.md> [output_dir]

  input.md    path to source Markdown file
  output_dir  where to write results
              (default: <stem>_output/ next to the input file)

EXAMPLES
  md2docx thesis.md
  md2docx docs/report.md /tmp/report_out

OUTPUT
  <output_dir>/
    <stem>.docx      — final document
    diagrams/        — rendered Mermaid PNG files

SYSTEM DEPENDENCIES  (must be installed separately)
  pandoc      brew install pandoc
  mmdc        npm install -g @mermaid-js/mermaid-cli

UPDATE THIS TOOL
  cd ~/Documents/Projects/md2docx
  uv tool install . --reinstall

UNINSTALL
  uv tool uninstall md2docx
"""


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert Markdown + Mermaid to .docx",
        add_help=True,
    )
    parser.add_argument("input", nargs="?", help="Path to source .md file")
    parser.add_argument(
        "output_dir",
        nargs="?",
        default=None,
        help="Output directory (default: <input_stem>_output next to input file)",
    )
    parser.add_argument(
        "--info", "--doctor",
        action="store_true",
        help="Show usage, dependencies, and update instructions",
    )
    args = parser.parse_args()

    if args.info or not args.input:
        print(DOCTOR_TEXT)
        sys.exit(0)

    src_md = Path(args.input).resolve()
    if not src_md.exists():
        print(f"ERROR: file not found: {src_md}", file=sys.stderr)
        sys.exit(1)

    src_dir = src_md.parent
    stem = src_md.stem

    if args.output_dir:
        out_dir = Path(args.output_dir).resolve()
    else:
        out_dir = src_dir / f"{stem}_output"

    diagrams_dir = out_dir / "diagrams"
    rendered_md = diagrams_dir / "_rendered.md"
    processed_md = diagrams_dir / "_processed.md"
    output_docx = out_dir / f"{stem}.docx"

    out_dir.mkdir(parents=True, exist_ok=True)

    run_mmdc(src_md, rendered_md, diagrams_dir)

    print("▶ Processing captions...")
    processed = process_markdown(rendered_md.read_text(encoding="utf-8"))
    processed_md.write_text(processed, encoding="utf-8")

    run_pandoc(processed_md, output_docx, src_dir, diagrams_dir)
    postprocess(output_docx)

    # Clean up temp markdown files; keep PNGs in diagrams/
    rendered_md.unlink(missing_ok=True)
    processed_md.unlink(missing_ok=True)

    print(f"✓ Done: {output_docx}")
    print(f"  Diagrams: {diagrams_dir}/")


if __name__ == "__main__":
    main()
