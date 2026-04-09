"""
FastAPI backend for the Markdown → DOCX converter.

Endpoints
---------
GET  /                       → serve static/index.html
POST /preview                → parse markdown, return AST JSON
POST /convert                → convert to .docx, return file download
POST /render-mermaid         → render mermaid code, return PNG
GET  /download-mermaid-images → return ZIP of mermaid PNGs from last conversion
"""

import io
import json
import sys
import unicodedata
import zipfile
from pathlib import Path
from typing import Dict

from fastapi import FastAPI, Form, HTTPException
from fastapi.responses import FileResponse, JSONResponse, Response, StreamingResponse

# ─── Paths ────────────────────────────────────────────────────────────────────

# Support both dev (script dir) and PyInstaller frozen (_MEIPASS)
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys._MEIPASS)  # type: ignore[attr-defined]
else:
    BASE_DIR = Path(__file__).parent

STATIC_DIR       = BASE_DIR / "static"
MERMAID_OUT_DIR  = Path(__file__).parent / "mermaid_output"
MERMAID_OUT_DIR.mkdir(exist_ok=True)

# ─── App ──────────────────────────────────────────────────────────────────────

app = FastAPI(title="MD → DOCX Converter", docs_url=None, redoc_url=None)

# Simple in-memory store for the last conversion's mermaid images
_last_mermaid: Dict[str, bytes] = {}

# ─── Routes ───────────────────────────────────────────────────────────────────


@app.get("/")
async def index():
    html = STATIC_DIR / "index.html"
    if not html.exists():
        raise HTTPException(status_code=404, detail="index.html not found")
    return FileResponse(str(html), media_type="text/html")


@app.post("/preview")
async def preview(markdown: str = Form(...)):
    """Parse markdown and return the mistune AST as JSON for the preview panel."""
    from converter import parse_markdown_structure
    markdown = unicodedata.normalize("NFC", markdown)
    try:
        tokens = parse_markdown_structure(markdown)
        return JSONResponse({"tokens": tokens})
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/convert")
async def convert(
    markdown:      str  = Form(...),
    style_config:  str  = Form(default="{}"),
    include_mermaid: bool = Form(default=True),
):
    """Convert markdown to DOCX and return as a file download."""
    global _last_mermaid

    from converter import convert_markdown_to_docx

    markdown = unicodedata.normalize("NFC", markdown)

    try:
        config = json.loads(style_config)
    except Exception:
        config = {}

    try:
        docx_bytes, mermaid_images = convert_markdown_to_docx(
            markdown, config, include_mermaid
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Conversion error: {exc}")

    _last_mermaid = mermaid_images

    # Persist mermaid images to disk for inspection
    for fname, data in mermaid_images.items():
        (MERMAID_OUT_DIR / fname).write_bytes(data)

    return Response(
        content=docx_bytes,
        media_type=(
            "application/vnd.openxmlformats-officedocument"
            ".wordprocessingml.document"
        ),
        headers={
            "Content-Disposition": 'attachment; filename="converted.docx"',
            "X-Mermaid-Count": str(len(mermaid_images)),
        },
    )


@app.post("/render-mermaid")
async def render_mermaid_endpoint(mermaid_code: str = Form(...)):
    """Render a mermaid diagram to PNG and return the image bytes."""
    from mermaid_renderer import render_mermaid
    try:
        png = render_mermaid(mermaid_code)
        return Response(content=png, media_type="image/png")
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.get("/download-mermaid-images")
async def download_mermaid_images():
    """Return a ZIP archive of all mermaid PNGs from the last conversion."""
    if not _last_mermaid:
        raise HTTPException(
            status_code=404,
            detail="No mermaid images available. Run a conversion first.",
        )
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, data in _last_mermaid.items():
            zf.writestr(fname, data)
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/zip",
        headers={"Content-Disposition": 'attachment; filename="mermaid_images.zip"'},
    )


# ─── Dev / standalone entry point ─────────────────────────────────────────────

if __name__ == "__main__":
    import webbrowser
    import uvicorn

    PORT = 8765
    webbrowser.open(f"http://localhost:{PORT}")
    uvicorn.run(app, host="127.0.0.1", port=PORT, reload=False)
