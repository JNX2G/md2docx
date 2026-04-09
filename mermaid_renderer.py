"""
Mermaid → PNG 렌더러

렌더링 우선순위
--------------
1. Playwright (headless Chromium) — 오프라인 고품질, mermaid.js 로컬 캐시 사용
2. mermaid.ink 공개 API        — Playwright 불가 시 인터넷 fallback
3. PIL 플레이스홀더             — 완전 오프라인 최후 수단 (코드 텍스트 이미지)
"""

from __future__ import annotations

import asyncio
import base64
import io
import urllib.request
import urllib.error
from pathlib import Path

# ─── mermaid.js 로컬 캐시 ────────────────────────────────────────────────────

_MERMAID_CDN      = "https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"
_CDN_ROUTE_GLOB   = "**/mermaid.min.js"
_CACHE_DIR        = Path(__file__).parent / ".cache"
_CACHE_FILE       = _CACHE_DIR / "mermaid.min.js"


def _ensure_mermaid_js() -> bool:
    """캐시가 없으면 CDN 에서 내려받는다. 성공 여부를 반환."""
    if _CACHE_FILE.exists() and _CACHE_FILE.stat().st_size > 100_000:
        return True
    try:
        print("[mermaid] mermaid.min.js 다운로드 중 …")
        _CACHE_DIR.mkdir(parents=True, exist_ok=True)
        req = urllib.request.Request(_MERMAID_CDN,
                                     headers={"User-Agent": "md2docx/1.0"})
        with urllib.request.urlopen(req, timeout=20) as resp:
            data = resp.read()
        _CACHE_FILE.write_bytes(data)
        print(f"[mermaid] 캐시 저장 완료 ({len(data):,} bytes): {_CACHE_FILE}")
        return True
    except Exception as exc:
        print(f"[mermaid] mermaid.js 다운로드 실패: {exc!r}")
        return False


# ─── HTML 템플릿 ──────────────────────────────────────────────────────────────

_HTML_TPL = """\
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  html, body {{ margin:0; padding:0; background:#fff; }}
  #_wrap {{
    display: inline-block;
    padding: 24px;
    background: #fff;
    min-width: 200px;
  }}
  #_wrap svg {{ max-width:none !important; height:auto !important; }}
  #_done {{ display:none; }}
  #_err  {{ display:none; color:red; font-size:12px; padding:8px; }}
</style>
</head>
<body>
<div id="_wrap">
  <div id="_graph" class="mermaid">{code}</div>
</div>
<span id="_done"></span>
<pre  id="_err"></pre>
<script src="{js_url}"></script>
<script>
  mermaid.initialize({{
    startOnLoad: false,
    theme: 'default',
    securityLevel: 'loose',
    flowchart: {{ useMaxWidth: false, htmlLabels: true }},
    sequence:  {{ useMaxWidth: false }},
    gantt:     {{ useMaxWidth: false }}
  }});
  mermaid.run({{ querySelector: '#_graph' }})
    .then(function() {{
      document.getElementById('_done').style.display = 'inline';
    }})
    .catch(function(e) {{
      document.getElementById('_err').textContent = String(e);
      document.getElementById('_err').style.display = 'block';
      document.getElementById('_done').style.display = 'inline';  /* signal anyway */
    }});
</script>
</body>
</html>
"""


# ─── 방법 1: Playwright ───────────────────────────────────────────────────────

async def _playwright_render(mermaid_code: str) -> bytes:
    from playwright.async_api import async_playwright

    has_cache = _ensure_mermaid_js()

    # HTML 안의 코드 — 브라우저가 HTML 엔티티를 디코딩하므로 mermaid 는 원본 텍스트를 받음
    safe = (mermaid_code
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;"))

    html = _HTML_TPL.format(code=safe, js_url=_MERMAID_CDN)

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)
        context = await browser.new_context(viewport={"width": 1800, "height": 1200})
        page    = await context.new_page()

        # 콘솔 오류 기록
        page.on("console", lambda m: (
            print(f"[mermaid:browser] {m.type}: {m.text}")
            if m.type in ("error", "warning") else None
        ))

        # CDN 요청을 로컬 캐시로 대체 (캐시 있을 때)
        if has_cache:
            cached_bytes = _CACHE_FILE.read_bytes()
            async def _serve_cache(route, _req):
                await route.fulfill(
                    content_type="application/javascript; charset=utf-8",
                    body=cached_bytes,
                )
            await page.route(_CDN_ROUTE_GLOB, _serve_cache)

        # load — 외부 스크립트(mermaid.js)까지 모두 로드된 뒤 반환
        await page.set_content(html, wait_until="load", timeout=30_000)

        # mermaid.run() 완료 대기 (최대 25초)
        try:
            await page.wait_for_function(
                "document.getElementById('_done').style.display !== 'none'",
                timeout=25_000,
            )
        except Exception as te:
            print(f"[mermaid] 렌더링 대기 타임아웃: {te!r}")

        # 오류 메시지 확인
        err_text = await page.eval_on_selector("#_err", "el => el.textContent")
        if err_text and err_text.strip():
            raise RuntimeError(f"mermaid 렌더링 오류: {err_text.strip()}")

        # SVG 생성 여부 확인
        svg = await page.query_selector("#_wrap svg")
        if svg is None:
            raise RuntimeError("SVG 요소가 생성되지 않았습니다.")

        # 실제 콘텐츠 영역 스크린샷
        wrap = await page.query_selector("#_wrap")
        box  = await wrap.bounding_box()
        png  = await page.screenshot(
            type="png",
            clip={"x": box["x"], "y": box["y"],
                  "width": box["width"], "height": box["height"]},
        )
        await browser.close()
        return png


def _run_playwright(mermaid_code: str) -> bytes:
    """
    FastAPI/uvicorn 은 이미 asyncio 루프를 실행 중이므로,
    같은 스레드에서 loop.run_until_complete() 를 호출하면
    'Cannot run the event loop while another loop is running' 오류가 발생한다.
    별도 스레드에서 새 루프를 생성해 실행함으로써 이를 회피한다.
    """
    import concurrent.futures

    def _in_thread() -> bytes:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            return loop.run_until_complete(_playwright_render(mermaid_code))
        finally:
            loop.close()

    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as pool:
        future = pool.submit(_in_thread)
        return future.result(timeout=60)  # 최대 60초 대기


# ─── 방법 2: mermaid.ink 공개 API ────────────────────────────────────────────

def _render_mermaid_ink(mermaid_code: str) -> bytes:
    """
    https://mermaid.ink 공개 렌더링 API 사용.
    인터넷 연결이 필요하지만 Playwright 없이도 동작한다.
    """
    # mermaid.ink 는 mermaid 코드를 base64url 인코딩해 URL 에 포함
    encoded = base64.urlsafe_b64encode(
        mermaid_code.encode("utf-8")
    ).decode("ascii").rstrip("=")

    url = f"https://mermaid.ink/img/{encoded}?bgColor=ffffff&width=900"
    req = urllib.request.Request(url, headers={"User-Agent": "md2docx/1.0"})

    with urllib.request.urlopen(req, timeout=20) as resp:
        ct = resp.headers.get("Content-Type", "")
        data = resp.read()

    if "image" not in ct and len(data) < 500:
        raise RuntimeError(f"mermaid.ink 응답 오류 (Content-Type: {ct})")

    return data


# ─── 방법 3: PIL 플레이스홀더 ────────────────────────────────────────────────

def _render_placeholder(mermaid_code: str) -> bytes:
    """Playwright 와 API 모두 실패했을 때 코드 텍스트 이미지를 반환."""
    try:
        from PIL import Image, ImageDraw

        lines  = mermaid_code.strip().splitlines()
        width  = 780
        height = max(180, 56 + 20 * len(lines))
        img    = Image.new("RGB", (width, height), (248, 249, 252))
        draw   = ImageDraw.Draw(img)

        draw.rectangle([0, 0, width - 1, height - 1], outline=(180, 185, 200), width=1)
        draw.rectangle([0, 0, width, 36], fill=(228, 233, 245))
        draw.text((10, 10), "[ Mermaid 다이어그램 — 렌더링 불가 (Playwright / 인터넷 연결 확인) ]",
                  fill=(80, 90, 120))

        y = 50
        for line in lines:
            if y > height - 22:
                draw.text((12, y), "…", fill=(130, 130, 140))
                break
            draw.text((12, y), line[:100], fill=(40, 50, 70))
            y += 20

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()

    except Exception:
        return base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk"
            "+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
        )


# ─── 공개 API ─────────────────────────────────────────────────────────────────

def render_mermaid(mermaid_code: str) -> bytes:
    """
    Mermaid 코드를 PNG 바이트로 변환한다.

    시도 순서:
      1. Playwright (headless Chromium + 로컬 mermaid.js 캐시)
      2. mermaid.ink 공개 API (인터넷 필요)
      3. PIL 플레이스홀더 (코드 텍스트 이미지)
    """
    # ── 1. Playwright ────────────────────────────────────────────────────────
    try:
        result = _run_playwright(mermaid_code)
        print("[mermaid] Playwright 렌더링 성공")
        return result
    except Exception as exc:
        print(f"[mermaid] Playwright 실패: {exc!r}")

    # ── 2. mermaid.ink API ───────────────────────────────────────────────────
    try:
        result = _render_mermaid_ink(mermaid_code)
        print("[mermaid] mermaid.ink API 렌더링 성공")
        return result
    except Exception as exc:
        print(f"[mermaid] mermaid.ink 실패: {exc!r}")

    # ── 3. PIL 플레이스홀더 ──────────────────────────────────────────────────
    print("[mermaid] 플레이스홀더 이미지 사용")
    return _render_placeholder(mermaid_code)


# ─── 다이어그램 타입 감지 ─────────────────────────────────────────────────────

_DIAGRAM_TYPES: list[tuple[str, str]] = [
    ("sequencediagram",    "Sequence Diagram"),
    ("classDiagram",       "Class Diagram"),
    ("statediagram-v2",    "State Diagram"),
    ("statediagram",       "State Diagram"),
    ("erdiagram",          "ER Diagram"),
    ("flowchart",          "Flow Diagram"),
    ("graph",              "Flow Diagram"),
    ("gantt",              "Gantt Chart"),
    ("pie",                "Pie Chart"),
    ("gitgraph",           "Git Graph"),
    ("mindmap",            "Mind Map"),
    ("timeline",           "Timeline"),
    ("journey",            "User Journey"),
    ("quadrantchart",      "Quadrant Chart"),
    ("requirementdiagram", "Requirement Diagram"),
    ("c4context",          "C4 Context Diagram"),
    ("c4container",        "C4 Container Diagram"),
    ("c4component",        "C4 Component Diagram"),
    ("xychart-beta",       "XY Chart"),
    ("sankey-beta",        "Sankey Diagram"),
]


def detect_diagram_type(mermaid_code: str) -> str:
    first = (mermaid_code.strip().lower()
             .splitlines()[0].strip()
             .replace(" ", "")
             .replace("-", ""))
    for key, label in _DIAGRAM_TYPES:
        if first.startswith(key.lower().replace("-", "")):
            return label
    return "Diagram"
