# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

### Setup (first time)
```bash
# Windows
setup.bat

# macOS / Linux
chmod +x setup.sh && ./setup.sh
```

### Run dev server
```bash
# Windows (activate venv first)
.venv\Scripts\activate.bat
uvicorn main:app --reload --port 8765

# macOS / Linux
source .venv/bin/activate
uvicorn main:app --reload --port 8765
```

Open http://localhost:8765 in the browser.

### Build standalone executable
```bash
# Windows
.venv\Scripts\activate.bat && build.bat    # → dist\md2docx.exe

# macOS / Linux
source .venv/bin/activate && ./build.sh    # → dist/md2docx
```

The build uses PyInstaller `--onefile` and bundles `static/` and all uvicorn hidden imports. No Python needed on target machines.

## Architecture

The app is a FastAPI server (`main.py`) with a single vanilla-JS frontend (`static/index.html`). All conversion work happens in two backend modules:

### `converter.py` — Markdown → DOCX
- Public API: `convert_markdown_to_docx(markdown, style_config, include_mermaid)` → `(bytes, dict[str, bytes])` and `parse_markdown_structure(markdown)` → token list.
- Uses **mistune ≥ 3.0** with its AST renderer. Token shape: `{"type": ..., "attrs": {...}, "children": [...], "raw": ...}`. The `attrs` and `children` keys exist only when the token type has them — do not assume their presence.
- `style_config` is a dict of overrides merged on top of defaults. Spacing values are **twips** (1440 = 1 inch). Colors are **6-digit hex without `#`**.
- Korean text uses Malgun Gothic by default.

### `mermaid_renderer.py` — Mermaid → PNG
- Public API: `render_mermaid(code)` → `bytes` (PNG).
- Rendering priority: (1) Playwright headless Chromium with locally-cached `mermaid.min.js`, (2) mermaid.ink public API fallback, (3) PIL placeholder image.
- mermaid.js is cached in `.cache/mermaid.min.js` on first use.

### `main.py` — FastAPI routes
- `BASE_DIR` switches between script directory (dev) and `sys._MEIPASS` (PyInstaller frozen). Always use `BASE_DIR` when resolving bundled assets like `static/`.
- `_last_mermaid` is an in-memory `dict[str, bytes]` holding PNGs from the most recent `/convert` call; it feeds `/download-mermaid-images`.

### `static/index.html`
- 3-column layout: markdown input | live preview | style editor.
- Style config is persisted in `localStorage` and POSTed as JSON string to `/convert`.
- The textarea is never mutated by JS — the user's raw markdown is always preserved.
- Preview calls `/preview` (returns AST JSON) with debouncing; the panel renders a simplified HTML view of the AST, not the actual DOCX output.
