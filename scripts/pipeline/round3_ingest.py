"""
Round 3 ingest: convert _round3_raw.json (Google Maps scrape) into _round3_leads.json.
Lead numbers start at 90 (rows 2-90 already used in existing workbook).
"""
import json, re
from pathlib import Path

ROOT    = Path(r"c:/Users/91638/Desktop/GenosApollp clients")
RAW     = ROOT / "data" / "pipeline" / "_round3_raw.json"
OUT     = ROOT / "data" / "pipeline" / "_round3_leads.json"

START_NUM = 90  # continue numbering after existing workbook's last lead


def slugify(s: str) -> str:
    s = re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")
    return s[:40]


def short_name(company: str) -> str:
    return company.split(",")[0].split("|")[0].strip()[:30]


def main():
    raw = json.loads(RAW.read_text(encoding="utf-8"))

    leads = []
    for i, r in enumerate(raw):
        num = START_NUM + i
        company = r["name"]
        website = r["website"]
        website_url = r["website_url"]

        leads.append({
            "num":          num,
            "category":     r["category"],      # "restaurant" | "dentist"
            "first_name":   "",                  # no owner name from Google Maps
            "last_name":    "",
            "full_name":    "",
            "title":        "Owner",
            "company":      company,
            "email":        r["email"],
            "email_status": "Found" if r["email"] else "Not Found",
            "website_raw":  r["website_raw"],
            "website":      website,
            "website_url":  website_url,
            "phone":        r["phone"],
            "address":      r["address"],
            "city":         r["city"],
            "state":        r["state"],
            "country":      "United States",
            "rating":       r["rating"],
            "employees":    "",
            "industry":     r["category"],
            "linkedin":     "",
            "company_linkedin": "",
            "technologies": "",
            "keywords":     "",
            "annual_revenue": "",
            "slug":         slugify(company),
            "query_source": r["query_source"],
        })

    OUT.write_text(json.dumps(leads, indent=2), encoding="utf-8")

    restaurants = sum(1 for l in leads if l["category"] == "restaurant")
    dentists    = sum(1 for l in leads if l["category"] == "dentist")
    with_email  = sum(1 for l in leads if l["email"])
    with_web    = sum(1 for l in leads if l["website"])
    no_web      = sum(1 for l in leads if not l["website"])

    print(f"Wrote {len(leads)} leads -> {OUT}")
    print(f"  Restaurants: {restaurants}  |  Dentists: {dentists}")
    print(f"  With website: {with_web}  |  No website: {no_web}")
    print(f"  With email:   {with_email}")
    print(f"  Lead numbers: #{leads[0]['num']} - #{leads[-1]['num']}")


if __name__ == "__main__":
    main()
