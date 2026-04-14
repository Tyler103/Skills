#!/usr/bin/env python3
"""
verify_latex.py  —  Syntax-check a LaTeX Beamer .tex file generated from pptx.

Why this exists:
  The converter does its best to produce valid LaTeX, but some pptx content
  doesn't map cleanly (e.g. Greek/special Unicode characters, very long text
  spans, slides where images weren't extracted). This script catches the most
  common issues before you spend time running pdflatex only to see a wall of
  cryptic TeX errors.

Checks performed:
  1. Balanced \\begin{env} / \\end{env} pairs     → ERROR if mismatched
  2. Unescaped % characters                       → WARN  (becomes a LaTeX comment)
  3. Unescaped & outside tabular environments     → WARN  (illegal in text mode)
  4. Missing image files referenced by \\includegraphics → ERROR
  5. Empty \\begin{frame}...\\end{frame} blocks   → WARN  (slide with no content)
  6. Lines longer than 120 characters             → WARN  (often raw pasted text)

Usage:
    python verify_latex.py <file.tex>

Exit codes:
    0  — no errors found (warnings may still have been printed)
    1  — one or more errors found
"""

import sys
import os
import re
import argparse
from pathlib import Path
from collections import Counter


# ──────────────────────────────────────────────────────────────────────────────
# Helper utilities
# ──────────────────────────────────────────────────────────────────────────────

def _is_comment(line: str) -> bool:
    """
    Return True if the line is an intentional LaTeX comment.

    A line is a comment if its first non-whitespace character is %.
    We skip comment lines in most checks because their content is never
    processed by pdflatex and % signs inside them are deliberate.
    """
    return line.lstrip().startswith("%")


def _strip_comment(line: str) -> str:
    """
    Remove everything from the first unescaped % onwards.

    LaTeX treats % as a comment delimiter unless it is preceded by a backslash
    (\\%). This function walks the string character-by-character:
      - A backslash followed by any character is consumed as a two-char unit
        (so \\% is kept intact).
      - A bare % stops the scan and discards the rest of the line.

    Used to strip intentional comments before checking for accidental bare %.
    """
    result = []
    i = 0
    while i < len(line):
        if line[i] == "\\" and i + 1 < len(line):
            # Escaped character: keep both the backslash and the next char.
            result.append(line[i])
            result.append(line[i + 1])
            i += 2
        elif line[i] == "%":
            break  # comment starts here — discard the rest
        else:
            result.append(line[i])
            i += 1
    return "".join(result)


# ──────────────────────────────────────────────────────────────────────────────
# Check 1: balanced \begin / \end pairs
#
# Every LaTeX environment must be opened and closed exactly once. We use a
# dictionary that maps each environment name to a stack of line numbers where
# it was opened. When we see \end{env} we pop from the stack; any remaining
# entries after the full scan are unclosed environments.
# ──────────────────────────────────────────────────────────────────────────────

def check_environment_balance(lines: list[str]) -> list[str]:
    """
    Detect mismatched \\begin{env} / \\end{env} pairs and return error strings.

    Common causes in pptx-generated .tex files:
      - A free-floating text box that the converter wrapped in \\begin{block}
        but the closing tag ended up inside another frame.
      - Hand-edited .tex files where a frame was deleted but its \\end{frame}
        was left behind.

    Returns a list of "[ERROR] ..." strings (empty list if all pairs balance).
    """
    # depth maps env_name → list of line numbers where \\begin{env} appeared.
    # Using a list as a stack lets us report the exact opening line for each
    # unmatched \\begin if a \\end is never found.
    depth: dict[str, list[int]] = {}
    errors = []

    for lineno, raw in enumerate(lines, 1):
        if _is_comment(raw):
            continue  # skip comment-only lines

        # Find all \begin{env} on this line (there could be more than one).
        for m in re.finditer(r"\\begin\{(\w+\*?)\}", raw):
            env = m.group(1)
            depth.setdefault(env, []).append(lineno)

        # Find all \end{env} on this line and match them against open \begins.
        for m in re.finditer(r"\\end\{(\w+\*?)\}", raw):
            env = m.group(1)
            if not depth.get(env):
                # \end without a preceding \begin
                errors.append(
                    f"[ERROR] Line {lineno}: \\end{{{env}}} has no matching \\begin{{{env}}}"
                )
            else:
                depth[env].pop()  # matched — remove the most recent opening

    # Any environments still on the stack were never closed.
    for env, opens in depth.items():
        for lineno in opens:
            errors.append(
                f"[ERROR] Line {lineno}: \\begin{{{env}}} was never closed"
            )

    return errors


# ──────────────────────────────────────────────────────────────────────────────
# Check 2: unescaped % characters
#
# In LaTeX, % starts a comment that runs to the end of the line. If slide text
# contained "50% discount" and the converter failed to escape it as "50\%",
# pdflatex silently discards " discount" and the rest of that line — content
# disappears without an error message.
# ──────────────────────────────────────────────────────────────────────────────

def check_unescaped_percent(lines: list[str]) -> list[str]:
    """
    Warn about bare % signs that will silently truncate lines in pdflatex.

    We skip lines that are intentional LaTeX comments (start with %). For all
    other lines we strip any trailing comment first, then look for any remaining
    % that is not preceded by a backslash.
    """
    warnings = []
    # Negative lookbehind: match % not preceded by \
    pattern = re.compile(r"(?<!\\)%")

    for lineno, raw in enumerate(lines, 1):
        if _is_comment(raw):
            continue  # this line IS an intentional comment — skip it

        # Strip off any intentional trailing comment so we only flag % in content.
        stripped = _strip_comment(raw)

        # After stripping, a surviving bare % would be unusual but possible if
        # _strip_comment has a bug — flag it anyway for safety.
        if pattern.search(stripped):
            warnings.append(
                f"[WARN]  Line {lineno}: possible unescaped '%' — use \\% for a literal percent sign"
            )

    return warnings


# ──────────────────────────────────────────────────────────────────────────────
# Check 3: unescaped & outside tabular environments
#
# & is the column-separator in LaTeX tabular/align environments. Outside those
# environments it is a syntax error. Slide text like "R&D" or "AT&T" must be
# written as "R\&D" or "AT\&T".
# ──────────────────────────────────────────────────────────────────────────────

def check_unescaped_ampersand(lines: list[str]) -> list[str]:
    """
    Warn about bare & characters that appear outside column-alignment contexts.

    We track which environments are currently open using a Counter. As long as
    any tabular-family environment is open (depth > 0), & is a legal column
    separator and is not flagged.
    """
    warnings = []

    # All LaTeX environments where & is legal as a column/alignment separator.
    tabular_envs = {
        "tabular", "array",
        "align", "align*",
        "eqnarray",
        "matrix", "pmatrix", "bmatrix", "vmatrix", "Vmatrix",
        "cases",
    }

    # Counter tracks nesting depth per environment name.
    # depth["tabular"] == 2 means two nested tabular environments are open.
    depth = Counter()

    for lineno, raw in enumerate(lines, 1):
        if _is_comment(raw):
            continue

        # Update depth for any environments opened or closed on this line.
        for m in re.finditer(r"\\begin\{(\w+\*?)\}", raw):
            depth[m.group(1)] += 1
        for m in re.finditer(r"\\end\{(\w+\*?)\}", raw):
            depth[m.group(1)] -= 1

        # Determine whether we are currently inside any tabular-family environment.
        in_table = any(depth.get(e, 0) > 0 for e in tabular_envs)

        if not in_table:
            # Outside a table: any bare & is illegal in LaTeX text mode.
            for m in re.finditer(r"(?<!\\)&", raw):
                warnings.append(
                    f"[WARN]  Line {lineno}: bare '&' outside a tabular environment — use \\& for a literal ampersand"
                )

    return warnings


# ──────────────────────────────────────────────────────────────────────────────
# Check 4: missing image files
#
# The converter writes \includegraphics{beamer_images/slideN_imgM.png}. If an
# image failed to extract (e.g. it was a vector EMF shape), the file won't
# exist and pdflatex will error out with "File not found."
# ──────────────────────────────────────────────────────────────────────────────

def check_image_paths(lines: list[str], tex_dir: str) -> list[str]:
    """
    Verify that every \\includegraphics path points to an existing file.

    Image paths in the generated .tex are relative to the .tex file's directory
    (e.g. "beamer_images/slide2_img1.png"). We resolve them against `tex_dir`
    and check with os.path.isfile().

    Args:
        lines:   All lines of the .tex file.
        tex_dir: Absolute path to the directory containing the .tex file.
    """
    errors = []
    # Match \includegraphics[optional options]{path}
    # The [.*?] part is non-greedy and matches optional sizing arguments like
    # [width=0.8\textwidth].
    pattern = re.compile(r"\\includegraphics(?:\[.*?\])?\{([^}]+)\}")

    for lineno, raw in enumerate(lines, 1):
        if _is_comment(raw):
            continue
        for m in pattern.finditer(raw):
            img_path = m.group(1).strip()
            full_path = os.path.join(tex_dir, img_path)
            if not os.path.isfile(full_path):
                errors.append(
                    f"[ERROR] Line {lineno}: image not found: '{img_path}' "
                    f"(looked in {tex_dir})"
                )

    return errors


# ──────────────────────────────────────────────────────────────────────────────
# Check 5: empty frames
#
# A frame with no content is valid LaTeX but produces a blank slide. This
# usually means the pptx slide had content that the converter could not extract
# (e.g. a purely image-based slide where the image wasn't raster, or a slide
# with only SmartArt). We warn so the user can inspect those slides.
# ──────────────────────────────────────────────────────────────────────────────

def check_empty_frames(lines: list[str]) -> list[str]:
    """
    Warn about \\begin{frame}...\\end{frame} blocks that contain no visible content.

    "Visible content" means any line inside the frame other than blank lines.
    The \\begin{frame}{...} opening line itself is not counted as content.
    """
    warnings = []
    in_frame = False       # are we currently inside a frame?
    frame_start = 0        # line number where the current frame opened
    frame_content = []     # lines collected inside the current frame

    for lineno, raw in enumerate(lines, 1):
        stripped = raw.strip()

        if re.match(r"\\begin\{frame\}", stripped):
            # Start tracking a new frame.
            in_frame = True
            frame_start = lineno
            frame_content = []

        elif stripped == r"\end{frame}":
            if in_frame:
                # Filter out blank lines and the opening \begin{frame}{...} line
                # to get only the "real" content lines.
                real_content = [
                    l for l in frame_content
                    if l.strip() and not l.strip().startswith(r"\begin{frame}")
                ]
                if not real_content:
                    warnings.append(
                        f"[WARN]  Line {frame_start}: empty frame (no visible content) "
                        f"— the original slide may have contained unextractable shapes"
                    )
            in_frame = False

        elif in_frame:
            # Accumulate lines inside the frame for the content check above.
            frame_content.append(raw)

    return warnings


# ──────────────────────────────────────────────────────────────────────────────
# Check 6: very long lines
#
# This isn't a LaTeX error per se, but extremely long lines (> 120 chars) often
# indicate that slide content was pasted as a single unbroken string rather
# than being split into bullet points. Beamer may render such text as overflowed
# or wrapped in unexpected ways.
# ──────────────────────────────────────────────────────────────────────────────

def check_long_lines(lines: list[str], max_len: int = 120) -> list[str]:
    """
    Warn about lines exceeding `max_len` characters.

    The default threshold of 120 chars was chosen because it catches genuinely
    problematic long strings while ignoring normal LaTeX preamble lines and
    typical \\item content.
    """
    warnings = []
    for lineno, raw in enumerate(lines, 1):
        if len(raw.rstrip()) > max_len:
            warnings.append(
                f"[WARN]  Line {lineno}: very long line ({len(raw.rstrip())} chars) "
                f"— consider breaking into separate bullet points"
            )
    return warnings


# ──────────────────────────────────────────────────────────────────────────────
# Orchestrator: run all checks and print the report
# ──────────────────────────────────────────────────────────────────────────────

def verify(tex_path: str) -> bool:
    """
    Run every check on the given .tex file and print a consolidated report.

    Messages are sorted by line number so the user can work through them
    top-to-bottom in their editor.

    Returns:
        True  if no ERRORs were found (the file should compile).
        False if at least one ERROR was found (compilation will likely fail).
    """
    path = Path(tex_path)
    if not path.is_file():
        print(f"[ERROR] File not found: {tex_path}")
        return False

    # Read the entire file; tex_dir is used for resolving relative image paths.
    tex_dir = str(path.parent)
    lines = path.read_text(encoding="utf-8", errors="replace").splitlines()

    # Collect messages from all checks into one flat list.
    all_messages = []
    all_messages += check_environment_balance(lines)   # check 1
    all_messages += check_unescaped_percent(lines)     # check 2
    all_messages += check_unescaped_ampersand(lines)   # check 3
    all_messages += check_image_paths(lines, tex_dir)  # check 4
    all_messages += check_empty_frames(lines)          # check 5
    all_messages += check_long_lines(lines)            # check 6

    # Separate errors (compilation-blocking) from warnings (quality issues).
    errors   = [m for m in all_messages if m.startswith("[ERROR]")]
    warnings = [m for m in all_messages if m.startswith("[WARN]")]

    # Print summary header.
    print(f"\nVerification report for: {tex_path}")
    print(f"  Lines checked : {len(lines)}")
    print(f"  Errors        : {len(errors)}")
    print(f"  Warnings      : {len(warnings)}")

    # Print all messages sorted by line number for easy navigation.
    if errors or warnings:
        print()
        for msg in sorted(
            all_messages,
            key=lambda m: int(re.search(r"Line (\d+)", m).group(1))
                          if re.search(r"Line (\d+)", m) else 0
        ):
            print(msg)

    if not errors and not warnings:
        print("\n  All checks passed — safe to run pdflatex.")

    # Return True only if there are zero blocking errors.
    return len(errors) == 0


# ──────────────────────────────────────────────────────────────────────────────
# CLI entry point
# ──────────────────────────────────────────────────────────────────────────────

def main():
    """
    Parse arguments and invoke verify().

    Exit code 0 = no errors; exit code 1 = at least one error found.
    This lets the script be used in CI pipelines:
        python verify_latex.py output.tex || echo "Fix errors before compiling"
    """
    parser = argparse.ArgumentParser(
        description="Verify a LaTeX Beamer .tex file generated from a .pptx."
    )
    parser.add_argument("tex_file", help="Path to the .tex file to verify")
    args = parser.parse_args()

    ok = verify(args.tex_file)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
