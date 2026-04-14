---
name: pptx-to-beamer
description: >
  Converts a PowerPoint (.pptx) file into a compilable LaTeX Beamer presentation.
  Extracts slide titles, bullet lists, images, tables, and rich text formatting
  (bold, italic, underline). Infers or accepts a Beamer theme, then runs a
  verification pass on the generated .tex file. Use when the user provides a
  .pptx file and asks for a Beamer/LaTeX presentation, PDF slides, or a LaTeX
  conversion of their slides.
license: MIT
compatibility: Requires Python 3.10+, python-pptx, and Pillow (pip install python-pptx pillow). pdflatex must be available to compile the output.
metadata:
  author: tyler-ton
  version: "2.0"
allowed-tools: Bash(python3:*) Bash(pip:*) Bash(pdflatex:*) Read Write Edit Glob
---

# pptx-to-beamer

Convert a PowerPoint presentation to a compilable LaTeX Beamer document.
The converter parses every slide, maps its structure to LaTeX environments,
extracts embedded images, and writes a clean `.tex` file ready to compile
with `pdflatex`.

## Checklist

Work through these steps in order. Mark each ✓ complete before moving on.

- [ ] Step 1 — Parse arguments and validate inputs
- [ ] Step 2 — Install dependencies
- [ ] Step 3 — Decide on a Beamer theme
- [ ] Step 4 — Run the converter script
- [ ] Step 5 — Verify the generated .tex file
- [ ] Step 6 — Fix any errors or warnings
- [ ] Step 7 — Report results to the user

---

## Step 1 — Parse arguments and validate inputs

The user provides a `.pptx` path, and optionally an output `.tex` path and/or a theme name.

- Extract `INPUT_PPTX` from the user's message (e.g. `slides.pptx` or an absolute path).
- If no output path is given, default to: `<stem>_beamer.tex` in the same directory.
- If the user mentions a theme (e.g. "use Metropolis", "Madrid style"), capture it.
  Otherwise set theme to `auto` to let the script infer it.
- Confirm the `.pptx` file exists before continuing. If it does not, stop and ask the user
  for the correct path.

## Step 2 — Install dependencies

```bash
python3 -c "import pptx; import PIL" 2>&1
```

If the import fails, install:

```bash
pip install python-pptx pillow
```

## Step 3 — Decide on a Beamer theme

### User-specified theme
If the user named a theme, pass it via `--theme <NAME>`.

### Agent-inferred theme
If no theme was specified, pass `--theme auto`. The script will:
1. Read the embedded pptx theme name and map it to a Beamer equivalent.
2. Fall back to background-color analysis (dark → Metropolis; blue → Warsaw; red → AnnArbor; default → Madrid).

Consult [references/beamer-themes.md](references/beamer-themes.md) for the full
mapping table and color-theme options.

## Step 4 — Run the converter script

```bash
python3 scripts/pptx_to_beamer.py "<INPUT_PPTX>" "<OUTPUT_TEX>" --theme <THEME>
```

The script will:
- Map each slide to `\begin{frame}{Title} ... \end{frame}`.
- Translate bullet hierarchies into nested `\begin{itemize}` environments.
- Detect numbered lists and use `\begin{enumerate}` when appropriate.
- Preserve **bold** → `\textbf{}`, *italic* → `\textit{}`, underline → `\underline{}`.
- Detect tables and render them as `tabular` environments.
- Extract embedded images to `beamer_images/` next to the `.tex` file and generate
  `\includegraphics[width=0.8\textwidth]{beamer_images/<name>.png}`.
- Wrap free-floating text boxes in `\begin{block}{} ... \end{block}`.
- Escape all LaTeX special characters: `& % $ # _ { } ~ ^ \`.

## Step 5 — Verify the generated .tex file

```bash
python3 scripts/verify_latex.py "<OUTPUT_TEX>"
```

The verifier checks:
| Check                          | Severity |
|-------------------------------|----------|
| Unmatched `\begin` / `\end`   | ERROR    |
| Image file not found on disk  | ERROR    |
| Unescaped `%` sign            | WARN     |
| Unescaped `&` outside tabular | WARN     |
| Empty frame (no content)      | WARN     |
| Lines > 120 chars             | WARN     |

## Step 6 — Fix errors and warnings

### Errors (must fix before the file will compile)

**Unmatched environments**
Read the frame around the flagged line number. A common cause is a pptx text box
that opened an environment but the script closed it prematurely.
Fix by editing the `.tex` file to add the missing `\end{<env>}` or remove the
extra `\begin{<env>}`.

**Missing image**
The image file was not extracted (usually a vector/EMF shape, not a raster picture).
Replace the `\includegraphics` line with a placeholder:
```latex
% [IMAGE PLACEHOLDER: shape description here]
\fbox{\parbox{0.8\textwidth}{\centering [Image: shape description]}}
```

### Warnings (fix where practical)

**Unescaped `%`**
Replace each bare `%` with `\%`. Example: `100% complete` → `100\% complete`.

**Unescaped `&`**
Replace each bare `&` with `\&`. Example: `R&D` → `R\&D`.

**Empty frame**
Either add content or remove the frame entirely.

**Overly long line**
Usually raw text that should be wrapped into bullet items or a `\begin{block}`.

## Step 7 — Report results to the user

Tell the user:
1. The full path to the generated `.tex` file.
2. The Beamer theme chosen (and how to change it).
3. How many slides were converted.
4. Whether any slides may need manual review (empty frames, placeholders).
5. The compile command:
   ```bash
   cd "<output directory>"
   pdflatex "<output filename>.tex"
   ```
6. Point them to [references/beamer-themes.md](references/beamer-themes.md) if they
   want to explore other themes.
7. Offer to compile the PDF directly if `pdflatex` is available.

---

## Handling complex elements

### Tables
The converter renders all pptx tables as `tabular` with vertical and horizontal rules.
For publication-quality output, suggest replacing `|l|l|l|` with `booktabs` style:
```latex
\begin{tabular}{lll}
\toprule
\textbf{A} & \textbf{B} & \textbf{C} \\
\midrule
...
\bottomrule
\end{tabular}
```

### Shapes that don't map 1:1
Elements the converter cannot fully translate (diagrams, SmartArt, charts, connectors)
are silently skipped. After conversion, read the `.tex` output and compare it against
the original. For each missing element:
- Simple shapes → describe them in a `\begin{block}{}` or use TikZ.
- Charts → export from PowerPoint as PNG first, then re-run the converter.
- SmartArt → recreate as a `tikzpicture` or as a nested `itemize`.

### Speaker notes
Speaker notes are not extracted by default. If needed, they can be accessed via
`shape.notes_slide` and added as `\note{}` commands inside each frame.

---

## Gotchas

- `\includegraphics` paths are **relative to where you run `pdflatex`**. Always
  `cd` into the output directory before compiling; do not compile from another directory.
- The Metropolis theme requires the `beamertheme-metropolis` TeX package and
  Fira Sans font. If it is not installed, fall back to `Madrid` or `CambridgeUS`.
- pptx `level` attributes start at 0 (top bullet = level 0). The converter maps
  level 0 to the first `\begin{itemize}` depth, level 1 to the second, etc.
- Some pptx files have corrupt or missing placeholder format data. The converter
  catches these silently with `try/except`; check empty frames in the verifier output.
- LaTeX underline (`\underline`) does not line-break automatically. For long underlined
  spans, use the `soul` package: `\usepackage{soul}` and `\ul{long text here}`.
