# md2docx

Convert a Markdown file (with [Mermaid](https://mermaid.js.org/) diagrams) to a properly formatted `.docx`.

## What it does

1. **Renders Mermaid blocks** to PNG via `mmdc` (theme neutral, white background, 2400px wide for print-quality output)
2. **Merges captions** — a `*Figure N — caption*` / `*Рисунок N — ...*` line immediately before an image is moved into the image's alt text so pandoc renders it as a proper figure caption
3. **Converts to docx** via `pandoc`
4. **Post-processes** the docx: font and size applied to all text, images centered and downscaled to fit page margins (never enlarged), table borders added

Image pixel data is never resampled — only the display dimensions in the document are adjusted.

## System dependencies

```bash
brew install pandoc
npm install -g @mermaid-js/mermaid-cli   # provides mmdc
```

## Install

```bash
uv tool install git+https://github.com/unsaid-azizov/md2docx.git
```

Or clone and install locally:

```bash
git clone https://github.com/unsaid-azizov/md2docx.git
cd md2docx
uv tool install .
```

## Usage

```
md2docx <input.md> [output_dir] [options]
```

Output goes to `<stem>_output/` next to the input file by default:

```bash
md2docx thesis.md
# → thesis_output/thesis.docx
# → thesis_output/diagrams/*.png
```

With an explicit output directory:

```bash
md2docx docs/report.md /tmp/report_out
```

### Options

| Flag | Default | Description |
|---|---|---|
| `--width PX` | `2400` | Mermaid render width in pixels (~366 DPI at 14 cm) |
| `--font NAME` | `Times New Roman` | Body font name |
| `--font-size PT` | `12` | Body font size in points |
| `--max-width CM` | `14.0` | Max image display width in the document |
| `--max-height CM` | `20.0` | Max image display height in the document |
| `--info` | — | Print usage reference and exit |

### Examples

```bash
# Default (Times New Roman 12pt, 2400px Mermaid render)
md2docx thesis.md

# Different font and size
md2docx report.md --font "Arial" --font-size 11

# Wider diagrams and wider image area
md2docx arch.md --width 3600 --max-width 16

# Everything explicit
md2docx doc.md /tmp/out --font "Calibri" --font-size 11 --max-width 15 --max-height 18
```

### Figure captions

A line of the form `*Caption text*` immediately before an image is automatically merged into the image's alt text and rendered as a caption by pandoc:

```markdown
*Рисунок 1 — System architecture*
![](diagrams/arch.png)

*Figure 2 — Data flow*
![](diagrams/flow.png)
```

### Output layout

```
<output_dir>/
    <stem>.docx      — final document
    diagrams/        — rendered Mermaid PNG files
```

## Update

```bash
uv tool install git+https://github.com/unsaid-azizov/md2docx.git --reinstall
```

## Uninstall

```bash
uv tool uninstall md2docx
```
