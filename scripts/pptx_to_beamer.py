#!/usr/bin/env python3
"""
pptx_to_beamer.py  —  Convert a PowerPoint (.pptx) to a LaTeX Beamer document.

Usage:
    python pptx_to_beamer.py <input.pptx> [output.tex] [--theme THEME]

Options:
    --theme THEME   Override the inferred Beamer theme (e.g. Metropolis, Madrid).
                    Use "auto" (default) to let the script infer a theme from the
                    presentation's color palette.

Dependencies:
    pip install python-pptx pillow
"""

import sys
import os
import re
import argparse
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.util import Pt, Emu
    from pptx.enum.text import PP_ALIGN
    from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
    from pptx.dml.color import RGBColor
    import pptx.shapes.picture
except ImportError:
    print("ERROR: python-pptx not installed. Run: pip install python-pptx pillow")
    sys.exit(1)


# ──────────────────────────────────────────────────────────────────────────────
# LaTeX character escaping
# ──────────────────────────────────────────────────────────────────────────────

_ESCAPE_TABLE = [
    ("\\", r"\textbackslash{}"),
    ("&",  r"\&"),
    ("%",  r"\%"),
    ("$",  r"\$"),
    ("#",  r"\#"),
    ("_",  r"\_"),
    ("{",  r"\{"),
    ("}",  r"\}"),
    ("~",  r"\textasciitilde{}"),
    ("^",  r"\textasciicircum{}"),
    ("<",  r"\textless{}"),
    (">",  r"\textgreater{}"),
]

def escape_latex(text: str) -> str:
    """Escape special LaTeX characters in plain text."""
    for char, replacement in _ESCAPE_TABLE:
        text = text.replace(char, replacement)
    return text


# ──────────────────────────────────────────────────────────────────────────────
# Rich-text run formatting
# ──────────────────────────────────────────────────────────────────────────────

def format_run(run) -> str:
    """Return a single run's text wrapped in LaTeX formatting commands."""
    text = escape_latex(run.text)
    if not text.strip():
        return text

    bold      = run.font.bold
    italic    = run.font.italic
    underline = run.font.underline

    if bold and italic:
        text = rf"\textbf{{\textit{{{text}}}}}"
    elif bold:
        text = rf"\textbf{{{text}}}"
    elif italic:
        text = rf"\textit{{{text}}}"

    if underline:
        text = rf"\underline{{{text}}}"

    return text


def paragraph_to_latex(para) -> str:
    """Convert all runs in a paragraph to LaTeX, preserving rich formatting."""
    return "".join(format_run(r) for r in para.runs)


# ──────────────────────────────────────────────────────────────────────────────
# Shape-type helpers
# ──────────────────────────────────────────────────────────────────────────────

_TITLE_PH_TYPES = {
    PP_PLACEHOLDER.TITLE,
    PP_PLACEHOLDER.CENTER_TITLE,
    PP_PLACEHOLDER.SUBTITLE,
}

def _ph_type(shape):
    """Return the placeholder type for a shape, or None."""
    try:
        ph = shape.placeholder_format
        return ph.type if ph else None
    except (ValueError, AttributeError):
        return None


def is_title(shape) -> bool:
    return shape.has_text_frame and _ph_type(shape) in _TITLE_PH_TYPES


def is_body(shape) -> bool:
    """True for body/content placeholders (bullet areas)."""
    t = _ph_type(shape)
    return t is not None and t not in _TITLE_PH_TYPES


# ──────────────────────────────────────────────────────────────────────────────
# Table conversion
# ──────────────────────────────────────────────────────────────────────────────

def table_to_latex(table) -> str:
    """Convert a pptx Table object to a LaTeX tabular environment."""
    rows = table.rows
    if not rows:
        return ""

    num_cols = len(rows[0].cells)
    col_spec = "|" + "|".join(["l"] * num_cols) + "|"

    lines = [
        r"\begin{center}",
        rf"\begin{{tabular}}{{{col_spec}}}",
        r"\hline",
    ]

    for row_idx, row in enumerate(rows):
        cells = []
        for cell in row.cells:
            raw = cell.text_frame.text.strip() if cell.text_frame else ""
            cells.append(escape_latex(raw))
        row_str = " & ".join(cells) + r" \\"
        lines.append(row_str)
        lines.append(r"\hline")

    lines += [r"\end{tabular}", r"\end{center}"]
    return "\n".join(lines)


# ──────────────────────────────────────────────────────────────────────────────
# Bullet-list builder (handles nested levels)
# ──────────────────────────────────────────────────────────────────────────────

def bullets_to_latex(items: list[tuple[int, str]]) -> list[str]:
    """
    Convert a list of (indent_level, latex_text) tuples into nested
    \\begin{itemize} ... \\end{itemize} blocks.
    """
    lines = []
    depth = -1
    open_envs: list[int] = []

    for level, text in items:
        # Close deeper envs
        while depth > level:
            lines.append("  " * depth + r"\end{itemize}")
            open_envs.pop()
            depth -= 1
        # Open new env
        while depth < level:
            depth += 1
            lines.append("  " * depth + r"\begin{itemize}")
            open_envs.append(depth)
        lines.append("  " * (depth + 1) + rf"\item {text}")

    # Close remaining envs
    while open_envs:
        d = open_envs.pop()
        lines.append("  " * d + r"\end{itemize}")

    return lines


# ──────────────────────────────────────────────────────────────────────────────
# Single-slide conversion
# ──────────────────────────────────────────────────────────────────────────────

def extract_slide(slide, images_dir: str, slide_num: int) -> str:
    """Convert one slide to a \\begin{frame}...\\end{frame} block."""
    title_text = ""
    bullet_items: list[tuple[int, str]] = []
    table_blocks: list[str] = []
    freetext_blocks: list[str] = []
    image_names: list[str] = []

    for shape in slide.shapes:

        # ── Images ──────────────────────────────────────────────────────────
        if isinstance(shape, pptx.shapes.picture.Picture):
            try:
                img_name = f"slide{slide_num}_img{len(image_names)+1}.png"
                img_path = os.path.join(images_dir, img_name)
                with open(img_path, "wb") as fh:
                    fh.write(shape.image.blob)
                image_names.append(img_name)
            except Exception as e:
                print(f"  [warn] Could not extract image on slide {slide_num}: {e}")
            continue

        # ── Tables ──────────────────────────────────────────────────────────
        if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
            try:
                table_blocks.append(table_to_latex(shape.table))
            except Exception as e:
                print(f"  [warn] Could not convert table on slide {slide_num}: {e}")
            continue

        # ── Text shapes ─────────────────────────────────────────────────────
        if not shape.has_text_frame:
            continue

        if is_title(shape):
            title_text = escape_latex(shape.text_frame.text.strip())
            continue

        if is_body(shape):
            for para in shape.text_frame.paragraphs:
                text = paragraph_to_latex(para).strip()
                if text:
                    bullet_items.append((para.level, text))
        else:
            # Free-floating text box
            lines = []
            for para in shape.text_frame.paragraphs:
                text = paragraph_to_latex(para).strip()
                if text:
                    lines.append(text)
            if lines:
                freetext_blocks.append("\n".join(lines))

    # ── Build frame LaTeX ────────────────────────────────────────────────────
    frame_lines: list[str] = []
    title_arg = f"{{{title_text}}}" if title_text else "{}"
    frame_lines.append(rf"\begin{{frame}}{title_arg}")

    if bullet_items:
        frame_lines.extend(bullets_to_latex(bullet_items))

    for tbl in table_blocks:
        frame_lines.append("")
        frame_lines.append(tbl)

    for block in freetext_blocks:
        frame_lines.append("")
        frame_lines.append(r"\begin{block}{}")
        frame_lines.append(block)
        frame_lines.append(r"\end{block}")

    for img in image_names:
        frame_lines.append(rf"  \includegraphics[width=0.8\textwidth]{{beamer_images/{img}}}")

    frame_lines.append(r"\end{frame}")
    return "\n".join(frame_lines)


# ──────────────────────────────────────────────────────────────────────────────
# Theme inference
# ──────────────────────────────────────────────────────────────────────────────

# Maps lowercase keywords found in a pptx theme name to Beamer theme choices
_THEME_NAME_MAP = {
    "metropolitan": "Metropolis",
    "metro":        "Metropolis",
    "office":       "Madrid",
    "circuit":      "Berlin",
    "facet":        "AnnArbor",
    "ion":          "Warsaw",
    "wood":         "Hannover",
    "slate":        "CambridgeUS",
    "retrospect":   "Boadilla",
    "wisp":         "Luebeck",
    "organic":      "Bergen",
}

def _get_theme_name_from_pptx(prs: Presentation) -> str | None:
    """Try to read the pptx theme name from XML."""
    try:
        theme_elem = prs.slide_masters[0].theme.element
        name = theme_elem.get("name", "")
        return name.lower() if name else None
    except Exception:
        return None


def _dominant_color(prs: Presentation) -> tuple[int, int, int] | None:
    """Return the (R,G,B) of the first slide's background fill, if solid."""
    try:
        bg_fill = prs.slides[0].background.fill
        if bg_fill.type is not None:
            clr = bg_fill.fore_color.rgb
            return int(clr[0:2], 16), int(clr[2:4], 16), int(clr[4:6], 16)
    except Exception:
        pass
    try:
        bg_fill = prs.slide_masters[0].background.fill
        if bg_fill.type is not None:
            clr = bg_fill.fore_color.rgb
            return int(clr[0:2], 16), int(clr[2:4], 16), int(clr[4:6], 16)
    except Exception:
        pass
    return None


def infer_beamer_theme(prs: Presentation) -> str:
    """
    Heuristically map a pptx presentation to a Beamer theme:
    1. Try to match the embedded pptx theme name.
    2. Fall back to background-color analysis.
    3. Default to Madrid.
    """
    pptx_name = _get_theme_name_from_pptx(prs)
    if pptx_name:
        for keyword, beamer in _THEME_NAME_MAP.items():
            if keyword in pptx_name:
                return beamer

    color = _dominant_color(prs)
    if color:
        r, g, b = color
        brightness = (r * 299 + g * 587 + b * 114) // 1000
        if brightness < 60:
            return "Metropolis"   # dark background → modern dark theme
        if r < 80 and b > 120:
            return "Warsaw"       # blue-dominant
        if r > 140 and g < 80:
            return "AnnArbor"     # red/maroon dominant
        if g > 120 and r < 100:
            return "Hannover"     # green dominant

    return "Madrid"               # professional default


# ──────────────────────────────────────────────────────────────────────────────
# Title extraction
# ──────────────────────────────────────────────────────────────────────────────

def get_presentation_title(prs: Presentation) -> str:
    if not prs.slides:
        return "Presentation"
    for shape in prs.slides[0].shapes:
        if is_title(shape) and shape.text_frame.text.strip():
            return escape_latex(shape.text_frame.text.strip())
    return "Presentation"


# ──────────────────────────────────────────────────────────────────────────────
# Document assembly
# ──────────────────────────────────────────────────────────────────────────────

def build_preamble(title: str, theme: str) -> str:
    return rf"""% Generated by pptx_to_beamer.py
\documentclass{{beamer}}
\usetheme{{{theme}}}
\usepackage{{graphicx}}
\usepackage{{hyperref}}
\usepackage{{amsmath}}
\usepackage{{booktabs}}
\usepackage[utf8]{{inputenc}}

\title{{{title}}}
\date{{\today}}

\begin{{document}}

\begin{{frame}}
  \titlepage
\end{{frame}}

"""


def convert(input_path: str, output_path: str, theme_override: str = "auto") -> None:
    prs = Presentation(input_path)

    # Image output directory (sibling of .tex file)
    out_dir = os.path.dirname(os.path.abspath(output_path))
    images_dir = os.path.join(out_dir, "beamer_images")
    os.makedirs(images_dir, exist_ok=True)

    doc_title = get_presentation_title(prs)

    if theme_override.lower() == "auto":
        theme = infer_beamer_theme(prs)
        print(f"  Inferred Beamer theme: {theme}")
    else:
        theme = theme_override
        print(f"  Using Beamer theme: {theme}")

    frames = []
    total = len(prs.slides)
    for i, slide in enumerate(prs.slides, 1):
        print(f"  Converting slide {i}/{total} …", end="\r")
        frames.append(extract_slide(slide, images_dir, i))

    print()  # newline after progress

    with open(output_path, "w", encoding="utf-8") as fh:
        fh.write(build_preamble(doc_title, theme))
        fh.write("\n\n".join(frames))
        fh.write("\n\n\\end{document}\n")

    print(f"Done: {output_path}")
    extracted = os.listdir(images_dir)
    if extracted:
        print(f"Images ({len(extracted)}): {images_dir}/")
    else:
        os.rmdir(images_dir)

    print(f"\nTo compile:\n  pdflatex \"{os.path.basename(output_path)}\"")


# ──────────────────────────────────────────────────────────────────────────────
# Entry point
# ──────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Convert a .pptx file to a LaTeX Beamer presentation."
    )
    parser.add_argument("input", help="Input .pptx file")
    parser.add_argument("output", nargs="?", help="Output .tex file (optional)")
    parser.add_argument(
        "--theme",
        default="auto",
        help='Beamer theme name, or "auto" to infer from the presentation (default: auto)',
    )
    args = parser.parse_args()

    inp = args.input
    out = args.output or (Path(inp).stem + "_beamer.tex")
    convert(inp, out, theme_override=args.theme)


if __name__ == "__main__":
    main()
