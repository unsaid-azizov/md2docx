# md2docx

Convert a Markdown file (with [Mermaid](https://mermaid.js.org/) diagrams) to a properly formatted `.docx`.

## What it does

1. **Renders Mermaid blocks** to PNG via `mmdc` (theme neutral, width 900, white background)
2. **Merges captions** — standalone `*Figure N — caption*` / `*Рисунок N — ...* ` lines before an image are moved into the image's alt text so pandoc renders them as figure captions
3. **Converts to docx** via `pandoc`
4. **Post-processes** the docx: Times New Roman 12 pt, images centered and downscaled to fit the page, table borders added

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
md2docx <input.md> [output_dir]
```

| Argument | Description |
|---|---|
| `input.md` | Path to the source Markdown file |
| `output_dir` | Where to write results (default: `<stem>_output/` next to the input file) |

### Output layout

```
<output_dir>/
    <stem>.docx        — final document
    diagrams/          — rendered Mermaid PNG files
```

### Examples

```bash
md2docx thesis.md
md2docx docs/report.md /tmp/report_out
```

Run `md2docx --info` for a quick reference including update and uninstall commands.

## Update

```bash
uv tool install git+https://github.com/unsaid-azizov/md2docx.git --reinstall
```

## Uninstall

```bash
uv tool uninstall md2docx
```
