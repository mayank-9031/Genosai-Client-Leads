# GenosAI — Client Outreach Pipeline: Project Context

## What this project does

GenosAI runs a LinkedIn-sourced lead generation pipeline. Each lead is a company found via LinkedIn (or Apollo exports). We visit their website, score its quality, and choose one of two content tracks:

| Track | Trigger | Deliverable |
|---|---|---|
| **Audit track** | Website is weak / has clear gaps | AI automation proposal PPT focusing on website gaps + automation opportunities |
| **Pitch track** | Website is already solid | AI automation proposal PPT focusing purely on automation services relevant to the industry |

**Both tracks produce a PPT deck and a cold email.** The deck content differs — audit track highlights website weaknesses; pitch track focuses entirely on automation. After the deliverable is produced, the lead and cold email are pushed into the master Excel workbook.

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
- Generate an **AI automation proposal PPT** using `python-pptx` — **build from scratch**, do not clone any existing file
- **Design system:** `docs/BRAND_ASSETS.md` — single source of truth for all design decisions
- Slide structure follows: Eyebrow → Headline → Subhead → Visual (see BRAND_ASSETS §5.3)
- Save output deck to: `output/decks/<CompanySlug>_Audit_GenosAI.pptx`
- Email includes the deck as an attachment

### Pitch-only track
- Same PPT structure, but no website gap analysis — deck focuses entirely on AI automation
- Email body is automation-focused without referencing site weaknesses
- Still personalized with an ice-breaker from the founder's LinkedIn

---

## AI automation content — company-specific (critical rule)

**Every company operates differently.** The automation opportunities, pain points, and proposed solutions in both the email and the deck must be derived specifically from:

1. **Their website** — what they offer, how their business operates, what's missing or broken
2. **Their industry** — sector-specific workflows, tools, and customer journeys
3. **Their scale** — number of listings / clients / agents / locations / products visible on site
4. **Their digital gaps** — what's manual, slow, or absent

**Never use generic AI automation copy.** A real estate platform needs WhatsApp buyer AI + FICA automation. A swimming pool builder needs WhatsApp cost qualification + AMC renewal flows. A restaurant needs reservation AI + review automation. Research the company first, then write content specific to how they run their business.

Automation categories to consider (pick only what is genuinely relevant):
- WhatsApp AI agent (buyer / client inquiry response)
- Lead qualification and routing flows
- Booking / appointment / viewing scheduling
- Document collection and compliance workflows
- CRM / follow-up automation
- Review and testimonial collection
- Recruitment intake and onboarding
- Drip / nurture sequences (email or WhatsApp)
- Invoice, renewal, and payment reminder automation
- Social media / content automation
- Market reports and client update sequences

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

This is the **only** workbook to write to. Never write to `GenosAI_RealEstate_Leads*.xlsx` or any other file.

The workbook has **three sheets**. Every new lead must be added to **all three**:

| Sheet | Key columns |
|---|---|
| **Organic Leads** | #, Company, Contact Person, Title, Email, Phone, Website, LinkedIn, City, State/Country, Industry, Source, Status, Audit Score, Deck Path, Automation Opportunities, Notes |
| **Cold Emails** | #, Company, Contact, Email, Subject, Body, Follow-Up 1, Follow-Up 2 |
| **Pipeline** | #, Company, Contact, Email, Industry, Audit Score, Status, Last Touchpoint, Next Action, Est. Deal Value |

**Row numbering:** Before writing, always check `ws.max_row` on each sheet to find the next available row. Do not hardcode row numbers.

Use `copy_style()` from prior row to maintain formatting consistency (see `scripts/leads/lead2_rich_digiovanna.py` for the pattern).

---

## Per-lead script pattern

Each new lead produces **two scripts** in `scripts/leads/`:

| Script | Purpose |
|---|---|
| `lead{N}_{firstname}_{lastname}.py` | Writes all lead data to `GenosAI_OrganicLeads.xlsx` (all 3 sheets) |
| `lead{N}_{company_slug}_deck.py` | Builds the PPT deck from scratch using BRAND_ASSETS.md specs |

Run both scripts after writing them to verify output.

---

## Folder structure

```
Genos-clients/
├── clients/                     # One subfolder per retained client (reference only)
│   └── AquaTechnic/             # Example client work — do not clone for new decks
├── data/
│   ├── exports/                 # Raw Apollo CSV exports
│   ├── pipeline/                # JSON pipeline state files (_roundN_leads.json, _audits, _content)
│   └── scraped/                 # Raw HTML/text scraped from lead websites
├── docs/
│   ├── BRAND_ASSETS.md          # Single source of truth for all design decisions
│   └── team-claude.md           # Claude Code agent config
├── output/
│   ├── decks/                   # Generated audit PPTs (<Slug>_Audit_GenosAI.pptx)
│   └── workbooks/
│       └── GenosAI_OrganicLeads.xlsx   ← PRIMARY OUTPUT (only workbook to write to)
├── scripts/
│   ├── leads/                   # Per-lead scripts: lead{N}_{name}.py + lead{N}_{slug}_deck.py
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

For individual (one-off) leads use the two-script pattern in `scripts/leads/`.

---

## PPT deck design rules (non-negotiable)

All specs from `docs/BRAND_ASSETS.md`. Build every deck from scratch with python-pptx — never clone a client file.

- Canvas: `#0A0A0F` (not pure black)
- Headlines: Instrument Serif 400 — never bold
- Body / UI text: Satoshi (fallback: Calibri)
- Eyebrows: Satoshi 600, UPPERCASE, `#A78BFA`, letter-spacing 0.2em
- Cards: dark `#111118` fill, `#22222A` border, rounded corners
- Violet accent (`#6D28D9` / `#A78BFA`) only for eyebrows and accent elements — never body text
- No emojis, no exclamation marks, no buzzwords
- Logo `GenosAI` (Instrument Serif, white) appears once per deck, bottom of cover slide
- Slide dimensions: 13.33in × 7.5in (16:9 widescreen)

### Standard 12-slide structure (adapt sections to company context)

| # | Section | Content |
|---|---|---|
| 1 | Cover | Company name, location, "AI Automation & Growth Proposal", GenosAI wordmark |
| 2 | Audit | Current digital presence — strengths + gaps (derived from site research) |
| 3 | Vision | The opportunity — what changes with AI automation |
| 4 | Systems overview | 4–6 specific AI systems relevant to this company |
| 5–9 | Module slides | One per automation system, with feature bullets and context |
| 10 | Lead / workflow flow | End-to-end automated journey (customer or operational) |
| 11 | Implementation | 3 phases with deliverables and timeline |
| 12 | Closing | About GenosAI, next steps, contact |

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
