#!/usr/bin/env python3
"""
backdrops.py - Generate decorative background PNGs for document pages.

Renders lightweight HTML pages with Playwright and captures them as
high-resolution PNGs that are later embedded as floating images behind
document content in the C# generation step.

Three variants are produced:

* **title_backdrop** - full-bleed visual for the title page
* **content_watermark** - subtle decoration for content pages
* **closing_flourish** - lighter echo for the closing page

The visual language uses radial gradients and dot patterns
for a modern, subtle aesthetic.

Usage::

    python backdrops.py [output_dir] [palette_name]
"""

import os
import sys

# ---------------------------------------------------------------------------
# Palette resolution
# ---------------------------------------------------------------------------

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def _resolve_palette(name: str | None) -> dict[str, str]:
    """Load palette and return a flat dict of CSS colours."""
    from .palettes import fetch_palette

    pal = fetch_palette(name or "Verdant")
    orn = pal["Ornamental"]
    typo = pal["Typography"]
    ui = pal["Interface"]
    return {
        "p1": f"#{orn['Subtle']['Primary']}",
        "p2": f"#{orn['Subtle']['Secondary']}",
        "p3": f"#{orn['Subtle']['Highlight']}",
        "dark": f"#{typo['Heading']}",
        "mid": f"#{typo['Caption']}",
        "divider": f"#{ui['Divider']}",
        "surface": f"#{ui['Surface']}",
    }


# A4 at 96 dpi - derived from mm
_W = round(210 * 96 / 25.4) - 1   # 793
_H = round(297 * 96 / 25.4) - 1   # 1122


# ---------------------------------------------------------------------------
# HTML builders - each returns a complete HTML string
# ---------------------------------------------------------------------------

def _title_backdrop_html(c: dict[str, str]) -> str:
    """Title page: radial gradient with dot accents."""
    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{
  width:{_W}px;height:{_H}px;
  background:{c['surface']};
  position:relative;overflow:hidden
}}

/* Large radial gradient top-right */
.glow-a{{
  position:absolute;top:-150px;right:-150px;
  width:500px;height:500px;
  background:radial-gradient(circle,{c['p1']}30 0%,transparent 70%);
  border-radius:50%;
}}

/* Smaller glow bottom-left */
.glow-b{{
  position:absolute;bottom:-80px;left:-80px;
  width:300px;height:300px;
  background:radial-gradient(circle,{c['p2']}25 0%,transparent 70%);
  border-radius:50%;
}}

/* Dot pattern overlay */
.dots{{
  position:absolute;top:60px;left:40px;
  display:grid;
  grid-template-columns:repeat(5,12px);
  grid-gap:8px;
}}
.dot{{
  width:4px;height:4px;
  background:{c['p1']}40;
  border-radius:50%;
}}

/* Vertical accent line */
.vline{{
  position:absolute;top:0;right:120px;
  width:1px;height:30%;
  background:linear-gradient(180deg,{c['p1']}35,transparent);
}}

/* Horizontal rule near bottom */
.hrule{{
  position:absolute;bottom:140px;left:80px;
  width:180px;height:1px;
  background:linear-gradient(90deg,{c['p1']}40,transparent);
}}
</style></head>
<body>
  <div class="glow-a"></div>
  <div class="glow-b"></div>
  <div class="dots">
    <div class="dot"></div><div class="dot"></div><div class="dot"></div>
    <div class="dot"></div><div class="dot"></div><div class="dot"></div>
    <div class="dot"></div><div class="dot"></div><div class="dot"></div>
    <div class="dot"></div><div class="dot"></div><div class="dot"></div>
    <div class="dot"></div><div class="dot"></div><div class="dot"></div>
  </div>
  <div class="vline"></div>
  <div class="hrule"></div>
</body></html>"""


def _closing_flourish_html(c: dict[str, str]) -> str:
    """Closing page: subtle glow and minimal accents."""
    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{
  width:{_W}px;height:{_H}px;
  background:{c['surface']};
  position:relative;overflow:hidden
}}

/* Radial glow bottom-right */
.glow{{
  position:absolute;bottom:-100px;right:-100px;
  width:350px;height:350px;
  background:radial-gradient(circle,{c['p2']}22 0%,transparent 70%);
  border-radius:50%;
}}

/* Top corner accent */
.corner{{
  position:absolute;top:50px;right:50px;
  width:60px;height:1px;
  background:{c['p1']}30;
}}

/* Bottom-left dot cluster */
.dots{{
  position:absolute;bottom:60px;left:60px;
  display:flex;gap:6px;
}}
.dot{{
  width:3px;height:3px;
  background:{c['p1']}35;
  border-radius:50%;
}}
</style></head>
<body>
  <div class="glow"></div>
  <div class="corner"></div>
  <div class="dots">
    <div class="dot"></div><div class="dot"></div><div class="dot"></div>
  </div>
</body></html>"""


def _content_watermark_html(c: dict[str, str]) -> str:
    """Content pages: minimal, near-invisible watermark."""
    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{
  width:{_W}px;height:{_H}px;
  background:#FDFDFD;
  position:relative;overflow:hidden
}}

/* Faint left-edge gradient */
.edge{{
  position:absolute;top:0;left:0;
  width:2px;height:100%;
  background:linear-gradient(180deg,{c['p1']}18,{c['p1']}05 50%,transparent);
}}

/* Hairline top rule */
.top-rule{{
  position:absolute;top:0;left:0;right:0;
  height:1px;
  background:linear-gradient(90deg,{c['divider']}80,transparent 50%);
}}

/* Subtle corner glow */
.corner-glow{{
  position:absolute;bottom:-40px;right:-40px;
  width:120px;height:120px;
  background:radial-gradient(circle,{c['p2']}08 0%,transparent 70%);
  border-radius:50%;
}}
</style></head>
<body>
  <div class="edge"></div>
  <div class="top-rule"></div>
  <div class="corner-glow"></div>
</body></html>"""


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------

# Background asset naming (semantic, distinct from competitors)
BACKGROUND_ASSETS = {
    'title_page': 'title_backdrop.png',
    'content_page': 'content_watermark.png',
    'closing_page': 'closing_flourish.png',
}


def render_backdrops(
    output_dir: str,
    palette_name: str | None = None,
) -> list[str]:
    """Render title, content, and closing backgrounds. Returns output paths."""
    from playwright.sync_api import sync_playwright

    colours = _resolve_palette(palette_name)
    os.makedirs(output_dir, exist_ok=True)

    pages_spec = [
        (BACKGROUND_ASSETS['title_page'], _title_backdrop_html(colours)),
        (BACKGROUND_ASSETS['content_page'], _content_watermark_html(colours)),
        (BACKGROUND_ASSETS['closing_page'], _closing_flourish_html(colours)),
    ]

    outputs: list[str] = []
    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page(
            viewport={"width": _W, "height": _H},
            device_scale_factor=3,
        )
        for filename, html in pages_spec:
            dest = os.path.join(output_dir, filename)
            page.set_content(html)
            page.screenshot(path=dest, type="png")
            outputs.append(dest)
            print(f"  -> {filename}")
        browser.close()

    return outputs


if __name__ == "__main__":
    out_dir = sys.argv[1] if len(sys.argv) > 1 else _SCRIPT_DIR
    pal = sys.argv[2] if len(sys.argv) > 2 else None
    render_backdrops(out_dir, pal)
    print("Done")
