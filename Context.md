# GenosAI — Client Outreach Pipeline: Project Context

## What this project does

GenosAI runs a LinkedIn-sourced lead generation pipeline. Each lead is a company found via LinkedIn (or Apollo exports). We visit their website, score its quality, and choose one of two tracks:

| Track | Trigger | Deliverable |
|---|---|---|
| **Audit track** | Website is weak / has clear gaps | Free website audit PPT + AI automation opportunities |
| **Pitch track** | Website is already solid | Automation-only proposal (no audit deck) |

After the deliverable is produced, the lead — along with the cold email — is pushed into the master Excel workbook.

---

## Lead sources

- **Primary:** LinkedIn company search → export via Apollo or manual scrape
- **Input fields we need per lead:**
  - `first_name`, `last_name`, `title`, `company`
  - `email` + `email_status`
  - `website` (URL)
  - `linkedin` (founder's personal LinkedIn URL) ← used for ice-breakers
  - `company_linkedin` (company page URL)
  - `phone`, `city`, `state`, `country`, `employees`, `industry`

---

## Two-track decision logic

### How to score a website
Run `scripts/pipeline/round3_audit.py` logic (or equivalent) against the URL:
- Checks: mobile viewport, analytics, booking/scheduling, live chat, review tools, email marketing, CDN, remarketing pixels
- Scores 0–10; **< 6.5 = Audit track**, **≥ 6.5 = Pitch-only track**
- Priority: HIGH (< 4), MED (4–6.5), LOW (≥ 6.5)

### Audit track
- Generate a **website audit PPT** using `python-pptx`
- **Design reference:** `clients/AquaTechnic/AquaTechnic_Audit_GenosAI [Repaired].pptx`
- **Brand system:** `docs/BRAND_ASSETS.md` (canvas `#0A0A0F`, Instrument Serif headlines, Satoshi body, violet eyebrows)
- Slide structure follows: Eyebrow → Headline → Subhead → Visual (see BRAND_ASSETS §5.3)
- Save output deck to: `output/decks/<CompanySlug>_Audit_GenosAI.pptx`
- Email includes the deck as an attachment

### Pitch-only track
- No deck generated
- Email focuses purely on AI automation services relevant to the company's industry
- Still personalized with an ice-breaker from the founder's LinkedIn

---

## Email personalization — mandatory approach

Every cold email **must** include an ice-breaker derived from the **founder's personal LinkedIn profile** (not the company page). Steps:

1. Visit the founder's personal LinkedIn URL (from the `linkedin` field)
2. Extract a specific, genuine detail: a recent post topic, a milestone they shared, a career achievement, a school, a location, a skill endorsement — anything concrete
3. Open the email with 1–2 sentences referencing that detail (never generic)
4. Do **not** fabricate details; if the profile is inaccessible, fall back to a website-derived hook

Ice-breaker tone: calm, observational, zero flattery. Reference the fact, not praise for the fact.

Example (good): *"Saw the post on your Coast Guard background — 14 years of service before real estate is a story that sells itself."*
Example (bad): *"I love what you're doing with your incredible career!"*

Follow-up emails (FU1, FU2) are shorter, reference the first email, and drop the ice-breaker — just a nudge.

---

## Output workbook

**Path:** `C:\Users\deevansho\Desktop\Genos-clients\output\workbooks\GenosAI_OrganicLeads.xlsx`

The workbook has five sheets. Every new lead must be added to **all five**:

| Sheet | Key columns |
|---|---|
| **Leads** | #, First, Last, Full Name, Title, Company, Email, Email Status, Website, Phone, City, State, Country, Employees, Industry, LinkedIn, PPT filename |
| **Website Audits** | #, Company, Website, Platform, Mobile, Analytics, Chat, Booking, Reviews, Email Mkt, CDN, Remarketing, Score (0–10), Priority, Issues (text), AI Opportunities (text) |
| **Cold Emails** | #, First Name, Company, Email, Subject, Body, Follow-up 1, Follow-up 2 |
| **Audit PPT Content** | #, Company, Contact, Email, Website, Score, Priority, Deck Title, Company Background, Strengths, Weaknesses, AI Opportunities, Pain Points, Proposed Solutions, 90-day Roadmap, CTA |
| **Organic Leads** | Same as Leads + Source, Date Added, PPT filename, Status |

Use `copy_style()` from prior row to maintain formatting consistency (see `scripts/leads/lead2_rich_digiovanna.py` for the pattern).

---

## Folder structure

```
Genos-clients/
├── clients/                     # One subfolder per retained client
│   └── AquaTechnic/             # Reference client — PPT design template here
├── data/
│   ├── exports/                 # Raw Apollo CSV exports
│   ├── pipeline/                # JSON pipeline state files (_roundN_leads.json, _audits, _content)
│   └── scraped/                 # Raw HTML/text scraped from lead websites
├── docs/
│   ├── BRAND_ASSETS.md          # Single source of truth for all design decisions
│   └── team-claude.md           # Claude Code agent config
├── output/
│   ├── decks/                   # Generated audit PPTs (<Slug>_Audit_GenosAI.pptx)
│   └── workbooks/               # Master Excel workbooks
│       └── GenosAI_OrganicLeads.xlsx   ← PRIMARY OUTPUT
├── scripts/
│   ├── leads/                   # One-off per-lead scripts (lead2_*.py … )
│   ├── pipeline/                # Batch pipeline scripts (ingest → audit → content → workbook → decks)
│   ├── scraping/                # Site scrapers
│   └── utils/                   # Workbook update helpers
├── Context.md                   # ← This file
├── docs/BRAND_ASSETS.md
└── requirements.txt
```

---

## Pipeline stages (batch mode)

For processing a batch of LinkedIn leads, run these in order:

1. **Ingest** (`scripts/pipeline/roundN_ingest.py`) — parse raw Apollo CSV / Google Maps JSON into `_roundN_leads.json`
2. **Audit** (`scripts/pipeline/roundN_audit.py`) — fetch each website, score it, write `_roundN_audits.json`
3. **Content** (`scripts/pipeline/roundN_content.py`) — generate cold email body + audit slide content per lead, write `_roundN_content.json`
4. **Workbook** (`scripts/pipeline/roundN_workbook.py`) — push all leads into the Excel workbook
5. **Decks** (`scripts/pipeline/roundN_decks.py`) — generate PPTX audit decks for HIGH/MED priority leads

For individual (one-off) leads use the pattern in `scripts/leads/lead*.py`.

---

## PPT deck design rules (non-negotiable)

Derived from `docs/BRAND_ASSETS.md` and the AquaTechnic reference deck:

- Canvas: `#0A0A0F` (not pure black)
- Headlines: Instrument Serif 400 — never bold
- Body / UI text: Satoshi
- Eyebrows: Satoshi 600, UPPERCASE, `#A78BFA` @ 70%, letter-spacing 0.2em
- Primary CTA button: white pill, `#0A0A0F` text, rounded-full
- Cards: `rounded-2xl`, 1 px `rgba(255,255,255,0.06)` border
- Violet accent (`#6D28D9` / `#A78BFA`) only for eyebrows and ambient glows — never body text
- No emojis, no exclamation marks, no buzzwords
- Logo `GenosAI` (Instrument Serif, white) appears once per deck, bottom of cover slide

---

## Sender identity (in all emails)

```
Rohan Malik
CEO, Genos AI
+91 638-714-2699  |  hello@genosai.tech  |  www.genosai.tech
```

---

## Python dependencies

```
openpyxl>=3.1.2
requests>=2.31.0
beautifulsoup4>=4.12.2
python-pptx>=0.6.21
playwright>=1.40.0
lxml>=4.9.3
```

ROOT path in scripts: `C:\Users\deevansho\Desktop\Genos-clients`
