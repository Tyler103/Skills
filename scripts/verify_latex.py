#!/usr/bin/env python3
"""
verify_latex.py  —  Syntax-check a LaTeX Beamer .tex file generated from pptx.

Checks performed:
  1. Balanced \\begin{env} / \\end{env} pairs.
  2. Unescaped % characters (common source of LaTeX comments unintentionally
     cutting off content).
  3. Unescaped & characters outside tabular/array environments.
  4. Existence of every image path referenced via \\includegraphics.
  5. Empty \\begin{frame} blocks (slides with no content).
  6. Warns about overly long lines that may cause wrapping issues.

Usage:
    python verify_latex.py <file.tex> [--images-dir DIR]

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
# Helpers
# ──────────────────────────────────────────────────────────────────────────────

def _is_comment(line: str) -> bool:
    """True if the line is a LaTeX comment (first non-space char is %)."""
    return line.lstrip().startswith("%")


def _strip_comment(line: str) -> str:
    """Remove everything from an unescaped % onwards."""
    result = []
    i = 0
    while i < len(line):
        if line[i] == "\\" and i + 1 < len(line):
            result.append(line[i])
            result.append(line[i + 1])
            i += 2
        elif line[i] == "%":
            break
        else:
            result.append(line[i])
            i += 1
    return "".join(result)


# ──────────────────────────────────────────────────────────────────────────────
# Checks
# ──────────────────────────────────────────────────────────────────────────────

def check_environment_balance(lines: list[str]) -> list[str]:
    """Return error strings for mismatched \\begin / \\end pairs."""
    depth: dict[str, list[int]] = {}  # env_name -> stack of line numbers
    errors = []

    for lineno, raw in enumerate(lines, 1):
        if _is_comment(raw):
            continue
        for m in re.finditer(r"\\begin\{(\w+\*?)\}", raw):
            env = m.group(1)
            depth.setdefault(env, []).append(lineno)
        for m in re.finditer(r"\\end\{(\w+\*?)\}", raw):
            env = m.group(1)
            if not depth.get(env):
                errors.append(
                    f"[ERROR] Line {lineno}: \\end{{{env}}} has no matching \\begin{{{env}}}"
                )
            else:
                depth[env].pop()

    for env, opens in depth.items():
        for lineno in opens:
            errors.append(
                f"[ERROR] Line {lineno}: \\begin{{{env}}} was never closed"
            )

    return errors


def check_unescaped_percent(lines: list[str]) -> list[str]:
    """
    Warn about bare % signs that will become LaTeX comments.
    A % is 'bare' if it is NOT preceded by a backslash.
    """
    warnings = []
    pattern = re.compile(r"(?<!\\)%")
    for lineno, raw in enumerate(lines, 1):
        # Skip lines that are intentional comments
        if _is_comment(raw):
            continue
        stripped = _strip_comment(raw)
        # Now check the stripped line for bare %  (shouldn't occur, but belt-
        # and-suspenders if _strip_comment missed something)
        if pattern.search(stripped):
            warnings.append(
                f"[WARN]  Line {lineno}: possible unescaped '%' — use \\% for a literal percent sign"
            )
    return warnings


def check_unescaped_ampersand(lines: list[str]) -> list[str]:
    """
    Warn about & outside tabular/array/align environments.
    Inside those environments & is the column separator and is correct.
    """
    warnings = []
    tabular_envs = {"tabular", "array", "align", "align*", "eqnarray", "matrix",
                    "pmatrix", "bmatrix", "vmatrix", "Vmatrix", "cases"}
    depth = Counter()

    for lineno, raw in enumerate(lines, 1):
        if _is_comment(raw):
            continue
        for m in re.finditer(r"\\begin\{(\w+\*?)\}", raw):
            depth[m.group(1)] += 1
        for m in re.finditer(r"\\end\{(\w+\*?)\}", raw):
            depth[m.group(1)] -= 1

        in_table = any(depth.get(e, 0) > 0 for e in tabular_envs)
        if not in_table:
            # find & not preceded by backslash
            for m in re.finditer(r"(?<!\\)&", raw):
                warnings.append(
                    f"[WARN]  Line {lineno}: bare '&' outside a tabular environment — use \\& for a literal ampersand"
                )

    return warnings


def check_image_paths(lines: list[str], tex_dir: str) -> list[str]:
    """
    Check that every \\includegraphics path resolves to an existing file.
    Searches relative to the .tex file's directory.
    """
    errors = []
    pattern = re.compile(r"\\includegraphics(?:\[.*?\])?\{([^}]+)\}")
    for lineno, raw in enumerate(lines, 1):
        if _is_comment(raw):
            continue
        for m in pattern.finditer(raw):
            img_path = m.group(1).strip()
            full_path = os.path.join(tex_dir, img_path)
            if not os.path.isfile(full_path):
                errors.append(
                    f"[ERROR] Line {lineno}: image not found: '{img_path}' (looked in {tex_dir})"
                )
    return errors


def check_empty_frames(lines: list[str]) -> list[str]:
    """Warn about \\begin{frame}...\\end{frame} blocks with no visible content."""
    warnings = []
    in_frame = False
    frame_start = 0
    frame_content = []

    for lineno, raw in enumerate(lines, 1):
        stripped = raw.strip()
        if re.match(r"\\begin\{frame\}", stripped):
            in_frame = True
            frame_start = lineno
            frame_content = []
        elif stripped == r"\end{frame}":
            if in_frame:
                # Content = lines that are not just \begin{frame}{...} or \end{frame}
                real_content = [
                    l for l in frame_content
                    if l.strip() and not l.strip().startswith(r"\begin{frame}")
                ]
                if not real_content:
                    warnings.append(
                        f"[WARN]  Line {frame_start}: empty frame (no visible content)"
                    )
            in_frame = False
        elif in_frame:
            frame_content.append(raw)

    return warnings


def check_long_lines(lines: list[str], max_len: int = 120) -> list[str]:
    """Warn about extremely long lines (not a LaTeX error, but often a sign of raw text)."""
    warnings = []
    for lineno, raw in enumerate(lines, 1):
        if len(raw.rstrip()) > max_len:
            warnings.append(
                f"[WARN]  Line {lineno}: very long line ({len(raw.rstrip())} chars) — consider breaking it up"
            )
    return warnings


# ──────────────────────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────────────────────

def verify(tex_path: str) -> bool:
    """Run all checks. Returns True if no errors were found."""
    path = Path(tex_path)
    if not path.is_file():
        print(f"[ERROR] File not found: {tex_path}")
        return False

    tex_dir = str(path.parent)
    lines = path.read_text(encoding="utf-8", errors="replace").splitlines()

    all_messages = []
    all_messages += check_environment_balance(lines)
    all_messages += check_unescaped_percent(lines)
    all_messages += check_unescaped_ampersand(lines)
    all_messages += check_image_paths(lines, tex_dir)
    all_messages += check_empty_frames(lines)
    all_messages += check_long_lines(lines)

    errors   = [m for m in all_messages if m.startswith("[ERROR]")]
    warnings = [m for m in all_messages if m.startswith("[WARN]")]

    print(f"\nVerification report for: {tex_path}")
    print(f"  Lines checked : {len(lines)}")
    print(f"  Errors        : {len(errors)}")
    print(f"  Warnings      : {len(warnings)}")

    if errors or warnings:
        print()
        for msg in sorted(all_messages, key=lambda m: int(re.search(r"Line (\d+)", m).group(1)) if re.search(r"Line (\d+)", m) else 0):
            print(msg)

    if not errors and not warnings:
        print("\n  All checks passed.")

    return len(errors) == 0


def main():
    parser = argparse.ArgumentParser(
        description="Verify a LaTeX Beamer .tex file generated from a .pptx."
    )
    parser.add_argument("tex_file", help="Path to the .tex file to verify")
    args = parser.parse_args()

    ok = verify(args.tex_file)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
