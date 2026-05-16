from pathlib import Path
from playwright.sync_api import sync_playwright
from pathlib import Path
import time, sys, io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

BASE = "https://swimmingpoolpro.in"

PAGES = [
    "/",
    "/commercial",
    "/residential",
    "/pricing",
    "/financing",
    "/blog",
    "/contact",
]

ROOT = Path(__file__).parent.parent.parent
OUT = str(ROOT / "data" / "scraped" / "swimmingpoolpro_scraped2.txt")
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
            page.goto(full, timeout=30000, wait_until="domcontentloaded")
            time.sleep(4)
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            time.sleep(2)
            text = page.evaluate("() => document.body.innerText") or ""
            status = "200"
            log(f"\n{'='*60}")
            log(f"URL: {full}")
            log(f"TEXT_LEN: {len(text)}")
            log(text[:7000])
            page.close()
        except Exception as e:
            log(f"\n[ERROR] {full} -- {str(e)[:300]}")
    browser.close()

with open(OUT, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))

print(f"\nSaved to {OUT}")
