"""
Round 3 website audit — restaurants & dentists.
Reads:  scripts/_round3_leads.json
Writes: scripts/_round3_audits.json
        scripts/_round3_issues.csv   (failed fetches)
"""
import json, re, csv, time
from pathlib import Path
import requests
from bs4 import BeautifulSoup

ROOT   = Path(r"c:/Users/91638/Desktop/GenosApollp clients")
LEADS  = ROOT / "data" / "pipeline" / "_round3_leads.json"
OUT    = ROOT / "data" / "pipeline" / "_round3_audits.json"
ISSUES = ROOT / "data" / "pipeline" / "_round3_issues.csv"

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)
HEADERS = {
    "User-Agent": UA,
    "Accept": "text/html,application/xhtml+xml,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

PLATFORM_HINTS = [
    ("WordPress",   [r"wp-content", r"wp-includes", r"/wp-json/", r'name="generator"\s+content="WordPress'], 0),
    ("Wix",         [r"static\.wixstatic", r"\.wix\.com"], 0),
    ("Squarespace", [r"static1\.squarespace", r"squarespace\.com"], 0),
    ("Webflow",     [r"webflow\.com", r"data-wf-page"], 1),
    ("Shopify",     [r"cdn\.shopify", r"myshopify\.com"], 0),
    ("Next.js",     [r"__NEXT_DATA__", r"/_next/"], 1),
    ("React",       [r"react", r"data-reactroot"], 1),
    ("GoDaddy",     [r"img1\.wsimg\.com", r"godaddysites"], 0),
    ("Custom",      [], 0),
]

ANALYTICS   = [r"gtag\(", r"google-analytics", r"googletagmanager", r"GTM-", r"fbq\(", r"plausible"]
BOOKING     = [r"opentable", r"resy\.com", r"yelp\.com/reservations", r"sevenrooms",
               r"toasttab\.com", r"scheduling", r"book.*appointment", r"online.*booking",
               r"calendly", r"acuityscheduling", r"healthgrades", r"zocdoc",
               r"patientpop", r"demandforce", r"nxtstep", r"carestack"]
ONLINE_ORDER= [r"doordash", r"ubereats", r"grubhub", r"toasttab", r"order.*online",
               r"online.*order", r"seamless", r"postmates"]
LOYALTY     = [r"loyalty", r"rewards", r"points.*program", r"punch.*card", r"frequent.*diner"]
CHAT        = [r"intercom\.io", r"drift\.com", r"tidio", r"tawk\.to", r"livechat",
               r"crisp\.chat", r"chatra", r"freshchat", r"zdassets"]
REVIEWS     = [r"trustpilot", r"birdeye", r"podium", r"grade\.us", r"reviewtrackers",
               r"reputation\.com", r"yotpo", r"google.*review", r"healthgrades"]
EMAIL_MKT   = [r"mailchimp", r"klaviyo", r"activecampaign", r"convertkit",
               r"constantcontact", r"sendgrid", r"mailerlite"]
CDN         = [r"cloudflare", r"fastly", r"akamaihd", r"jsdelivr\.net"]
SSL_PATTERN = [r"https://"]
REMARKET    = [r"fbq\(.*track", r"gtag\(.*conversion", r"adroll", r"pinterest"]


def signal(text, pats):
    return any(re.search(p, text, re.I) for p in pats)


def detect_platform(text):
    for name, pats, bonus in PLATFORM_HINTS:
        if pats and signal(text, pats):
            return name, bonus
    return "Custom", 0


def fetch(url, timeout=12):
    try:
        return requests.get(url, headers=HEADERS, timeout=timeout, allow_redirects=True)
    except Exception:
        return None


def audit_site(lead: dict) -> dict | None:
    url = lead["website_url"]
    if not url:
        return None
    r = fetch(url)
    if r is None or r.status_code >= 400:
        return None

    text = r.text
    soup = BeautifulSoup(text, "html.parser")

    title = (soup.title.string or "").strip() if soup.title and soup.title.string else ""
    meta  = ""
    md = soup.find("meta", attrs={"name": "description"})
    if md and md.get("content"):
        meta = md["content"].strip()
    h1 = soup.find("h1")
    h1_text = h1.get_text(strip=True) if h1 else ""

    platform, stack_bonus = detect_platform(text)
    mobile      = bool(soup.find("meta", attrs={"name": "viewport"}))
    has_analytics  = signal(text, ANALYTICS)
    has_booking    = signal(text, BOOKING)
    has_order      = signal(text, ONLINE_ORDER)
    has_loyalty    = signal(text, LOYALTY)
    has_chat       = signal(text, CHAT)
    has_reviews    = signal(text, REVIEWS)
    has_email_mkt  = signal(text, EMAIL_MKT)
    has_cdn        = signal(text, CDN) or "cloudflare" in (r.headers.get("Server","").lower())
    has_remarket   = signal(text, REMARKET)

    category = lead.get("category", "restaurant")

    # Score weights differ slightly by category
    if category == "restaurant":
        score_bits = [mobile, has_analytics, has_booking, has_order, has_chat,
                      has_reviews, has_email_mkt, has_remarket]
    else:  # dentist
        score_bits = [mobile, has_analytics, has_booking, has_chat,
                      has_reviews, has_email_mkt, has_cdn, has_remarket]

    raw   = sum(1 for b in score_bits if b)
    score = raw + stack_bonus
    if mobile and has_analytics and has_booking:
        score += 1
    score = min(round(score * 2) / 2, 10)

    if score < 4:
        priority = "HIGH"
    elif score < 6.5:
        priority = "MED"
    else:
        priority = "LOW"

    # Issues — industry-specific
    issues = []
    if category == "restaurant":
        if not has_booking:  issues.append("No online reservation system — customers call or leave")
        if not has_order:    issues.append("No online ordering — missing DoorDash/UberEats integration or native ordering")
        if not has_chat:     issues.append("No chatbot/live chat — after-hours questions go unanswered")
        if not has_analytics: issues.append("No analytics — no visibility into which pages drive covers")
        if not has_reviews:  issues.append("No review management tool — Google/Yelp reputation is unmonitored")
        if not has_email_mkt: issues.append("No email marketing — no way to bring diners back automatically")
        if platform in ("WordPress","GoDaddy","Wix"):
            issues.append(f"Running on {platform} — slow loads hurt mobile diner experience")
        if not mobile:       issues.append("No mobile viewport meta — most diners search on phones")
    else:
        if not has_booking:  issues.append("No online appointment booking — patients call or go elsewhere")
        if not has_chat:     issues.append("No chatbot/live chat — after-hours inquiries lost")
        if not has_analytics: issues.append("No analytics — no data on which services attract patients")
        if not has_reviews:  issues.append("No review management — patients choose dentists by star rating")
        if not has_email_mkt: issues.append("No email marketing — no automated recall or follow-up")
        if not has_remarket: issues.append("No remarketing — warm traffic that didn't book is gone")
        if platform in ("WordPress","GoDaddy","Wix"):
            issues.append(f"Running on {platform} — older template hurts trust vs modern practices")
        if not mobile:       issues.append("No mobile viewport — most patients search on phones")

    issues = issues[:3] or ["Site is in reasonable shape; biggest wins are conversion-side"]

    # Gaps (opportunities)
    gaps = []
    if category == "restaurant":
        if not has_booking:  gaps.append("Online reservation + waitlist system (OpenTable, Resy, or native)")
        if not has_order:    gaps.append("Direct online ordering to reduce third-party commission fees")
        if not has_chat:     gaps.append("AI chatbot for menu/hours/reservation questions after hours")
        if not has_email_mkt: gaps.append("Email loyalty & re-engagement campaigns to bring diners back")
        if not has_reviews:  gaps.append("Automated Google/Yelp review request after each visit")
    else:
        if not has_booking:  gaps.append("Online booking integration (Zocdoc, PatientPop, or native calendar)")
        if not has_chat:     gaps.append("AI chatbot for insurance/hours/appointment questions 24/7")
        if not has_email_mkt: gaps.append("Automated recall emails & appointment reminders")
        if not has_reviews:  gaps.append("Automated review requests after each appointment")
        if not has_remarket: gaps.append("Retargeting campaigns for patients who visited but didn't book")

    gaps = gaps[:3]

    breaker = (h1_text or meta or title or "").strip()
    breaker = re.sub(r"\s+", " ", breaker)[:200]

    return {
        "ok":               True,
        "url":              url,
        "final_url":        r.url,
        "platform":         platform,
        "mobile":           "Yes" if mobile else "No",
        "analytics":        "Yes" if has_analytics else "No",
        "booking":          "Yes" if has_booking else "No",
        "online_order":     "Yes" if has_order else "No",
        "loyalty":          "Yes" if has_loyalty else "No",
        "chat":             "Yes" if has_chat else "No",
        "reviews":          "Yes" if has_reviews else "No",
        "email_mkt":        "Yes" if has_email_mkt else "No",
        "cdn":              "Yes" if has_cdn else "No",
        "remarketing":      "Yes" if has_remarket else "No",
        "score":            score,
        "priority":         priority,
        "issues":           issues,
        "gaps":             gaps,
        "title":            title[:200],
        "meta_description": meta[:300],
        "h1":               h1_text[:200],
        "breaker_source":   breaker,
    }


def main():
    leads = json.loads(LEADS.read_text(encoding="utf-8"))
    audits = {}
    issues_rows = []

    for i, lead in enumerate(leads, 1):
        if not lead["website_url"]:
            print(f"[{i}/{len(leads)}] {lead['company'][:40]} -> NO WEBSITE, skipping")
            audits[lead["num"]] = {"lead_num": lead["num"], "company": lead["company"], "audit": None}
            continue

        print(f"[{i}/{len(leads)}] {lead['company'][:40]} -> {lead['website'][:35]}", flush=True)
        a = audit_site(lead)
        audits[lead["num"]] = {"lead_num": lead["num"], "company": lead["company"], "audit": a}
        if a is None:
            issues_rows.append([lead["num"], lead["company"], lead["website_url"], "fetch failed"])
        time.sleep(0.35)

    OUT.write_text(json.dumps(audits, indent=2), encoding="utf-8")

    if issues_rows:
        with open(ISSUES, "w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["#", "Company", "URL", "Issue"])
            w.writerows(issues_rows)

    ok       = sum(1 for v in audits.values() if v["audit"])
    no_web   = sum(1 for v in audits.values() if not v["audit"] and not leads[list(audits.keys()).index(v["lead_num"])]["website_url"])
    print(f"\nDone. Audited ok: {ok}  |  No website: {no_web}  |  Fetch failed: {len(issues_rows)}")
    print(f"Wrote -> {OUT}")


if __name__ == "__main__":
    main()
