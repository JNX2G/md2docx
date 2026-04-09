# MD → DOCX Converter

A FastAPI web application that converts Markdown documents to styled `.docx` files, with support for Korean text, Mermaid diagrams, tables, code blocks, and fully configurable styling.

## Features

- Full GFM Markdown support (headings, bold/italic/strikethrough, tables, code blocks, lists, blockquotes, links, images, horizontal rules)
- Mermaid diagram rendering via headless Chromium (Playwright), with PNG export
- Configurable styling (fonts, colors, spacing, page size) — persisted in the browser
- Korean text support (Malgun Gothic by default)
- Single-page web UI with live preview and style editor
- Standalone distributable executable (PyInstaller)

---

## Dev Setup (All Platforms)

**Prerequisites:** Python 3.11+ on PATH. On Windows, install from [python.org](https://python.org) with "Add to PATH" checked.

### macOS / Linux

```bash
git clone <repo-url>
cd md2docx
chmod +x setup.sh build.sh
./setup.sh
```

### Windows

```bat
git clone <repo-url>
cd md2docx
setup.bat
```

### Start the dev server

```bash
# macOS / Linux — activate venv first
source .venv/bin/activate
uvicorn main:app --reload --port 8765

# Windows — activate venv first
.venv\Scripts\activate.bat
uvicorn main:app --reload --port 8765
```

Open **http://localhost:8765** in your browser.

---

## Build a Standalone Executable

The executable bundles the app + Chromium and needs no Python on the target machine.

### macOS / Linux

```bash
source .venv/bin/activate
./build.sh
# Output: dist/md2docx
```

### Windows

```bat
.venv\Scripts\activate.bat
build.bat
REM Output: dist\md2docx.exe
```

Run the executable — it starts uvicorn on `localhost:8765` and opens the browser automatically.

---

## Project Structure

```
md2docx/
├── main.py               # FastAPI app & endpoints
├── converter.py          # Markdown → DOCX conversion logic
├── mermaid_renderer.py   # Mermaid → PNG via Playwright
├── static/
│   └── index.html        # Single-page frontend (vanilla JS, no build tools)
├── requirements.txt      # Python dependencies
├── setup.sh              # venv setup (macOS/Linux)
├── setup.bat             # venv setup (Windows)
├── build.sh              # PyInstaller build (macOS/Linux)
├── build.bat             # PyInstaller build (Windows)
└── mermaid_output/       # PNG files saved from last conversion (auto-created)
```

---

## API Reference

| Method | Path | Description |
|--------|------|-------------|
| `GET`  | `/` | Serve the web UI |
| `POST` | `/preview` | Parse markdown, return AST JSON |
| `POST` | `/convert` | Convert to DOCX, return file download |
| `POST` | `/render-mermaid` | Render mermaid code to PNG |
| `GET`  | `/download-mermaid-images` | ZIP of PNGs from last conversion |

### `/convert` form fields

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `markdown` | `string` | required | Markdown source text |
| `style_config` | `string` (JSON) | `{}` | Style overrides (see below) |
| `include_mermaid` | `bool` | `true` | Render mermaid as image vs. code block |

---

## Style Config Schema

```json
{
  "page_size": "A4",
  "font_name": "Malgun Gothic",
  "font_size_body": 11,
  "heading1": { "font_size": 20, "bold": true, "color": "1F3864", "space_before": 240, "space_after": 120 },
  "heading2": { "font_size": 16, "bold": true, "color": "2E75B6", "space_before": 200, "space_after": 80  },
  "heading3": { "font_size": 13, "bold": true, "color": "44546A", "space_before": 160, "space_after": 60  },
  "paragraph": { "space_after": 120, "line_spacing": 1.15 },
  "code_block": { "font_name": "Consolas", "font_size": 9, "background_color": "F2F2F2" },
  "table_header": { "background_color": "2E75B6", "font_color": "FFFFFF", "bold": true },
  "table_row_alt": { "background_color": "EBF3FB" },
  "margin": { "top": 1440, "bottom": 1440, "left": 1440, "right": 1440 }
}
```

Spacing values are in **twips** (1 inch = 1440 twips). Colors are 6-digit hex without `#`.

---

## Notes

- The `.venv` directory is excluded from version control
- Mermaid rendering requires an internet connection on first use (mermaid.js CDN). If Playwright is unavailable, a placeholder image is generated instead.
- On Windows, Playwright stores Chromium under `%USERPROFILE%\AppData\Local\ms-playwright`
- The textarea content is **never mutated** by the UI — your original markdown is always preserved
