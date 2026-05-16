"""
Round 2 website audit: fetch each lead's homepage, detect tech stack and
conversion-related signals, score 0-10. Also try the LinkedIn URL for an
ice-breaker hook (best-effort; LinkedIn typically blocks).

Reads: scripts/_round2_leads.json
Writes: scripts/_round2_audits.json   (one entry per lead, by email)
        scripts/_round2_issues.csv    (only leads where the homepage fetch failed)
"""
import json, re, csv, time, sys
from pathlib import Path
import requests
from bs4 import BeautifulSoup

ROOT = Path(__file__).parent.parent.parent
LEADS = ROOT / "data" / "pipeline" / "_round2_leads.json"
OUT = ROOT / "data" / "pipeline" / "_round2_audits.json"
ISSUES = ROOT / "data" / "pipeline" / "_round2_issues.csv"

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)
HEADERS = {
    "User-Agent": UA,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

PLATFORM_HINTS = [
    ("WordPress", [r"wp-content", r"wp-includes", r"/wp-json/", r'name="generator"\s+content="WordPress'], 0),
    ("Wix",       [r"static\.wixstatic", r"\.wix\.com", r"X-Wix-"], 0),
    ("Squarespace", [r"static1\.squarespace", r"squarespace\.com"], 0),
    ("Webflow",   [r"webflow\.com", r"data-wf-page"], 1),
    ("Shopify",   [r"cdn\.shopify", r"myshopify\.com", r"Shopify\.theme"], 0),
    ("Next.js",   [r"__NEXT_DATA__", r"/_next/"], 1),
    ("React",     [r"react", r"data-reactroot"], 1),
    ("HubSpot CMS", [r"\.hubspot\.com", r"hs-scripts\.com"], 1),
    ("GoDaddy",   [r"img1\.wsimg\.com", r"godaddysites"], 0),
    ("Custom",    [], 0),
]

ANALYTICS = [r"gtag\(", r"google-analytics", r"googletagmanager", r"GTM-", r"fbq\(", r"plausible", r"matomo"]
CRM       = [r"hubspot", r"salesforce", r"hsforms", r"_hsq", r"pipedrive", r"zoho\.com"]
CHAT      = [r"intercom\.io", r"drift\.com", r"tidio", r"tawk\.to", r"livechat", r"zdassets", r"crisp\.chat", r"hubspot.*chatflow", r"chatra", r"freshchat"]
AI        = [r"openai", r"anthropic", r"chatgpt", r"\bai chat", r"\bai assistant", r"copilot"]
CDN       = [r"cloudflare", r"fastly", r"akamaihd", r"jsdelivr\.net", r"cdn77", r"netdna"]
REMARKET  = [r"fbq\(.*track", r"gtag\(.*conversion", r"linkedin\.com/li\.lms-analytics", r"twq\(", r"adroll", r"pinterest"]
EMAIL_MKT = [r"mailchimp", r"klaviyo", r"activecampaign", r"convertkit", r"hubspot.*forms", r"sendgrid", r"mailerlite", r"constantcontact"]

def signal(text, patterns):
    for p in patterns:
        if re.search(p, text, re.I):
            return True
    return False

def detect_platform(text):
    for name, pats, bonus in PLATFORM_HINTS:
        if pats and signal(text, pats):
            return name, bonus
    return "Custom", 0

def fetch(url, timeout=12):
    try:
        r = requests.get(url, headers=HEADERS, timeout=timeout, allow_redirects=True)
        return r
    except Exception as e:
        return None

def audit_site(lead):
    url = lead["website_url"]
    r = fetch(url)
    if r is None or r.status_code >= 400:
        return None
    text = r.text
    soup = BeautifulSoup(text, "html.parser")
    title = (soup.title.string or "").strip() if soup.title and soup.title.string else ""
    meta = ""
    md = soup.find("meta", attrs={"name": "description"})
    if md and md.get("content"):
        meta = md["content"].strip()
    h1 = soup.find("h1")
    h1_text = h1.get_text(strip=True) if h1 else ""

    platform, stack_bonus = detect_platform(text)
    mobile = bool(soup.find("meta", attrs={"name": "viewport"}))
    has_analytics = signal(text, ANALYTICS)
    has_crm = signal(text, CRM)
    has_chat = signal(text, CHAT)
    has_ai = signal(text, AI)
    cdn_in_headers = any(
        signal(v, CDN) or signal(k, CDN)
        for k, v in r.headers.items()
    ) or "cloudflare" in (r.headers.get("Server", "").lower())
    has_cdn = signal(text, CDN) or cdn_in_headers
    has_remarket = signal(text, REMARKET)
    has_email_mkt = signal(text, EMAIL_MKT)

    bins = [mobile, has_analytics, has_crm, has_chat, has_ai, has_cdn, has_remarket, has_email_mkt]
    raw = sum(1 for b in bins if b)
    score = raw + stack_bonus
    if mobile and has_analytics and has_crm:
        score += 1
    score = min(round(score * 2) / 2, 10)  # half-point granularity, cap 10

    if score < 4:
        priority = "HIGH"
    elif score < 6.5:
        priority = "MED"
    else:
        priority = "LOW"

    issues = []
    if not has_chat:    issues.append("No chatbot/live chat - visitors leave without contacting")
    if not has_crm:     issues.append("No CRM detected - leads not captured or nurtured automatically")
    if not has_ai:      issues.append("No AI layer - every prospect question routes to a human")
    if not has_analytics: issues.append("No analytics - flying blind on what converts")
    if not has_remarket: issues.append("No remarketing pixels - warm traffic lost on exit")
    if not has_email_mkt: issues.append("No email marketing wiring - nurture is manual")
    if platform in ("WordPress", "GoDaddy", "Wix"):
        issues.append(f"Running on {platform} - plugins age fast, page speed lags Next.js/Webflow")
    if not mobile:      issues.append("No mobile viewport meta - mobile experience is degraded")
    issues = issues[:3] or ["Site is in decent shape; remaining wins are conversion-side"]

    gaps = []
    if not has_chat: gaps.append("AI chatbot for 24/7 lead qualification")
    if not has_crm:  gaps.append("AI lead capture & CRM automation")
    if not has_ai:   gaps.append("AI-powered property matching and answer engine")
    if not has_email_mkt: gaps.append("Automated follow-up & nurture sequences")
    if not has_remarket:  gaps.append("AI-driven retargeting based on visitor intent")
    gaps = gaps[:3]

    # ice-breaker source: title + meta + h1
    breaker = (h1_text or meta or title or "").strip()
    breaker = re.sub(r"\s+", " ", breaker)[:200]

    return {
        "ok": True,
        "url": url,
        "final_url": r.url,
        "platform": platform,
        "mobile": "Yes" if mobile else "No",
        "analytics": "Yes" if has_analytics else "No",
        "crm": "Yes" if has_crm else "No",
        "chat": "Yes" if has_chat else "No",
        "ai_tools": "Yes" if has_ai else "No",
        "cdn": "Yes" if has_cdn else "No",
        "remarketing": "Yes" if has_remarket else "No",
        "email_mkt": "Yes" if has_email_mkt else "No",
        "score": score,
        "priority": priority,
        "issues": issues,
        "gaps": gaps,
        "title": title[:200],
        "meta_description": meta[:300],
        "h1": h1_text[:200],
        "breaker_source": breaker,
    }

def try_linkedin(url):
    if not url:
        return None
    if not url.startswith("http"):
        url = "https://" + url
    r = fetch(url, timeout=8)
    if r is None:
        return None
    if r.status_code != 200:
        return {"ok": False, "status": r.status_code}
    text = r.text
    # public profiles partly render headline in og:description / og:title
    soup = BeautifulSoup(text, "html.parser")
    og_title = ""
    og_desc = ""
    for m in soup.find_all("meta"):
        prop = (m.get("property") or m.get("name") or "").lower()
        if prop in ("og:title",) and m.get("content"):
            og_title = m["content"].strip()
        if prop in ("og:description",) and m.get("content"):
            og_desc = m["content"].strip()
    if og_title or og_desc:
        return {"ok": True, "og_title": og_title[:200], "og_desc": og_desc[:400]}
    return {"ok": False, "status": r.status_code, "blocked": True}

def main():
    leads = json.loads(LEADS.read_text())
    audits = {}
    issues_rows = []
    for i, lead in enumerate(leads, 1):
        print(f"[{i}/{len(leads)}] {lead['company'][:50]} -> {lead['website']}", flush=True)
        a = audit_site(lead)
        li = try_linkedin(lead.get("linkedin"))
        entry = {"lead_num": lead["num"], "email": lead["email"], "audit": a, "linkedin": li}
        audits[lead["email"]] = entry
        if a is None:
            issues_rows.append([lead["num"], lead["company"], lead["website_url"], "homepage fetch failed"])
        time.sleep(0.4)  # be polite
    OUT.write_text(json.dumps(audits, indent=2))
    if issues_rows:
        with open(ISSUES, "w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["#", "Company", "URL", "Issue"])
            w.writerows(issues_rows)
    ok = sum(1 for v in audits.values() if v["audit"])
    li_ok = sum(1 for v in audits.values() if v["linkedin"] and v["linkedin"].get("ok"))
    print(f"Done. Site audits ok: {ok}/{len(audits)}. LinkedIn ok: {li_ok}/{len(audits)}. Wrote {OUT}")

if __name__ == "__main__":
    main()
