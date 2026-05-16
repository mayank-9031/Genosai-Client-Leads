from pathlib import Path
from playwright.sync_api import sync_playwright
from pathlib import Path
import time, sys, io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

BASE = "https://swimmingpoolpro.in"

PAGES = [
    "/",
    "/about",
    "/about-us",
    "/services",
    "/contact",
    "/contact-us",
    "/projects",
    "/gallery",
    "/team",
]

ROOT = Path(__file__).parent.parent.parent
OUT = str(ROOT / "data" / "scraped" / "swimmingpoolpro_scraped.txt")

lines = []

def log(s):
    lines.append(s)
    print(s)

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    ctx = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
    for path in PAGES:
        full = BASE + path
        try:
            page = ctx.new_page()
            resp = page.goto(full, timeout=20000, wait_until="networkidle")
            time.sleep(3)
            # scroll to trigger lazy loads
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            time.sleep(1)
            # get text via JS evaluate to bypass encoding issues
            text = page.evaluate("() => document.body.innerText") or ""
            html = page.content()
            status = resp.status if resp else "?"
            log(f"\n{'='*60}")
            log(f"URL: {full}  [HTTP {status}]")
            log(f"TEXT_LENGTH: {len(text)}  HTML_LENGTH: {len(html)}")
            log(text[:6000] if text.strip() else "[NO TEXT — SHOWING HTML SNIPPET]")
            if not text.strip():
                log(html[:3000])
            page.close()
        except Exception as e:
            log(f"\n[ERROR] {full} -- {str(e)[:200]}")
    browser.close()

with open(OUT, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))

print(f"\nSaved to {OUT}")
