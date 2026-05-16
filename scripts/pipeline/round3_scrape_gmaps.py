"""
Round 3 Google Maps scraper — restaurants + dentists (USA).
Scrapes Google Maps search results for each query, extracts business details,
then scrapes each website for contact email addresses.
Writes: scripts/_round3_raw.json
"""
import json, re, time, random
from pathlib import Path
from playwright.sync_api import sync_playwright

ROOT  = Path(r"c:/Users/91638/Desktop/GenosApollp clients")
OUT   = ROOT / "data" / "pipeline" / "_round3_raw.json"

QUERIES = [
    {"category": "restaurant", "query": "restaurants in Houston TX",   "target": 25},
    {"category": "restaurant", "query": "restaurants in Austin TX",    "target": 25},
    {"category": "dentist",    "query": "dentists in Houston TX",      "target": 25},
    {"category": "dentist",    "query": "dentists in Dallas TX",       "target": 25},
]

EMAIL_RE = re.compile(
    r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+",
    re.IGNORECASE
)
SKIP_EMAIL_DOMAINS = {
    "sentry.io","wixpress.com","example.com","your-domain","domain.com",
    "email.com","youremail","privacy","placeholder","png","jpg","jpeg",
    "wix.com","squarespace.com","wordpress.com","godaddy.com","webflow.io",
}

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)


def clean_email(email: str) -> str | None:
    email = email.strip().lower()
    domain = email.split("@")[-1] if "@" in email else ""
    if any(s in domain for s in SKIP_EMAIL_DOMAINS):
        return None
    if any(s in email for s in ["@2x", ".png", ".jpg", "example", "placeholder"]):
        return None
    if len(email) < 6 or len(email) > 80:
        return None
    return email


def find_email_on_page(text: str) -> str | None:
    for m in EMAIL_RE.finditer(text):
        e = clean_email(m.group())
        if e:
            return e
    return None


def scrape_website_email(page, url: str) -> str | None:
    """Try homepage + /contact page for an email address."""
    for path in ["", "/contact", "/contact-us", "/about"]:
        target = url.rstrip("/") + path
        try:
            page.goto(target, timeout=12000, wait_until="domcontentloaded")
            text = page.evaluate("() => document.body.innerText")
            e = find_email_on_page(text)
            if e:
                return e
        except Exception:
            pass
    return None


def scroll_and_collect(page, target: int) -> list[dict]:
    """Scroll the results panel until we have >= target entries or no more load."""
    results = []
    seen_names = set()
    max_scroll = 30

    for _ in range(max_scroll):
        # grab all result cards currently visible
        cards = page.query_selector_all('a[href*="maps/place"]')
        for card in cards:
            try:
                name = card.get_attribute("aria-label") or ""
                if not name or name in seen_names:
                    continue
                seen_names.add(name)
                results.append({"name": name, "href": card.get_attribute("href") or ""})
            except Exception:
                pass

        if len(results) >= target:
            break

        # scroll the left panel
        try:
            panel = page.query_selector('div[role="feed"]')
            if panel:
                panel.evaluate("el => el.scrollBy(0, 1200)")
            else:
                page.mouse.wheel(0, 1200)
        except Exception:
            pass
        time.sleep(1.2)

    return results[:target]


def get_place_details(page, href: str) -> dict:
    """Navigate to a place URL and extract phone, website, address, rating."""
    details: dict = {"phone": "", "website": "", "address": "", "rating": ""}
    try:
        page.goto(href, timeout=15000, wait_until="domcontentloaded")
        time.sleep(1.5)

        # website button
        try:
            wb = page.query_selector('a[data-item-id="authority"]')
            if wb:
                details["website"] = wb.get_attribute("href") or ""
        except Exception:
            pass

        # phone
        try:
            phones = page.query_selector_all('button[data-tooltip="Copy phone number"]')
            if phones:
                details["phone"] = (phones[0].get_attribute("aria-label") or "").replace("Phone:", "").strip()
        except Exception:
            pass

        # address
        try:
            addrs = page.query_selector_all('button[data-item-id^="address"]')
            if addrs:
                details["address"] = (addrs[0].get_attribute("aria-label") or "").replace("Address:", "").strip()
        except Exception:
            pass

        # rating
        try:
            rating_el = page.query_selector('div[jsaction*="pane.ratingStars"]')
            if rating_el:
                details["rating"] = rating_el.inner_text().split()[0]
        except Exception:
            pass

    except Exception:
        pass
    return details


def parse_city_state(address: str) -> tuple[str, str]:
    parts = [p.strip() for p in address.split(",")]
    city, state = "", ""
    if len(parts) >= 3:
        city = parts[-3] if len(parts) > 3 else parts[-2]
        state_zip = parts[-2]
        state = state_zip.split()[0] if state_zip else ""
    elif len(parts) == 2:
        city = parts[0]
        state = parts[1].split()[0]
    return city, state


def main():
    all_leads = []
    BATCH_KEY = set()

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context(user_agent=UA, locale="en-US")
        ctx.set_default_timeout(20000)

        # One page for Maps navigation, one for website email scraping
        maps_page = ctx.new_page()
        site_page = ctx.new_page()

        for q in QUERIES:
            category = q["category"]
            query    = q["query"]
            target   = q["target"]
            print(f"\n=== {query} (target {target}) ===")

            maps_url = f"https://www.google.com/maps/search/{query.replace(' ', '+')}/"
            maps_page.goto(maps_url, timeout=20000, wait_until="domcontentloaded")
            time.sleep(2)

            cards = scroll_and_collect(maps_page, target)
            print(f"  Collected {len(cards)} place cards")

            for i, card in enumerate(cards):
                name = card["name"]
                href = card["href"]
                if not href.startswith("http"):
                    href = "https://www.google.com" + href

                print(f"  [{i+1}/{len(cards)}] {name[:50]}", end=" ", flush=True)
                details = get_place_details(maps_page, href)

                website = details["website"]
                # clean website
                website_clean = re.sub(r"^https?://", "", website, flags=re.I).rstrip("/")

                # find email
                email = ""
                if website:
                    print(f"-> {website_clean[:30]}", end=" ", flush=True)
                    email = scrape_website_email(site_page, website) or ""
                    if email:
                        print(f"  email: {email}", flush=True)
                    else:
                        print("  (no email)", flush=True)
                else:
                    print("  (no website)", flush=True)

                city, state = parse_city_state(details["address"])

                key = (name.lower(), website_clean.lower())
                if key in BATCH_KEY:
                    continue
                BATCH_KEY.add(key)

                all_leads.append({
                    "category":     category,
                    "name":         name,
                    "website_raw":  website,
                    "website":      website_clean,
                    "website_url":  website if website.startswith("http") else (f"https://{website_clean}" if website_clean else ""),
                    "phone":        details["phone"],
                    "address":      details["address"],
                    "city":         city,
                    "state":        state,
                    "country":      "United States",
                    "rating":       details["rating"],
                    "email":        email,
                    "query_source": query,
                })
                time.sleep(random.uniform(0.4, 0.9))

        browser.close()

    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(json.dumps(all_leads, indent=2), encoding="utf-8")
    total = len(all_leads)
    with_web  = sum(1 for l in all_leads if l["website"])
    with_email = sum(1 for l in all_leads if l["email"])
    restaurants = sum(1 for l in all_leads if l["category"] == "restaurant")
    dentists    = sum(1 for l in all_leads if l["category"] == "dentist")
    print(f"\nDone. {total} leads ({restaurants} restaurants, {dentists} dentists)")
    print(f"  With website: {with_web}  |  With email: {with_email}")
    print(f"  Saved -> {OUT}")


if __name__ == "__main__":
    main()
