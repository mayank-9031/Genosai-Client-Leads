"""
Round 2 content generator: produces per-lead cold email + audit deck content
from leads + audits JSON. Writes scripts/_round2_content.json.
"""
import json, re
from pathlib import Path

ROOT = Path(__file__).parent.parent.parent
LEADS = ROOT / "data" / "pipeline" / "_round2_leads.json"
AUDITS = ROOT / "data" / "pipeline" / "_round2_audits.json"
OUT = ROOT / "data" / "pipeline" / "_round2_content.json"

SUBJECTS = [
    "{first}, the {site} chat gap",
    "{site} - quick note",
    "{first} - 90 seconds on {company_short}",
]

def shortname(c):
    c = re.sub(r"[\s,].*", "", c).strip()  # first token
    return c[:30] or c

def short_company(c):
    return c.split(",")[0].split("|")[0].strip().rstrip()[:40].strip() or c

def clean_text(s):
    """Strip mojibake and normalize unicode artifacts from scraped text."""
    if not s: return s
    # common mojibake from utf-8 read as latin-1
    repl = {
        "’": "'", "‘": "'", "“": '"', "”": '"',
        "–": "-", "—": "-", "…": "...",
        "�": "", " ": " ",
        "·": "-",  # middle dot
        # raw mojibake bytes-as-latin1
        "â": "'", "â": "'",
        "â": '"', "â": '"',
        "â": "-", "â": "-",
    }
    for k, v in repl.items():
        s = s.replace(k, v)
    s = re.sub(r"[�]+", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def is_boilerplate(text):
    """LinkedIn OG descriptions are often pure profile-card boilerplate; reject."""
    t = text.lower()
    boilerplate_markers = [
        "experience:", "education:", "location:", "connections on linkedin",
        "view profile", "join linkedin", "create your free",
        "posts about", "find out more about",
    ]
    hits = sum(1 for m in boilerplate_markers if m in t)
    return hits >= 2  # 2+ markers = it's a profile card, not real content

def ice_breaker(lead, audit, li):
    """Choose the strongest hook available. Avoids LinkedIn boilerplate."""
    company = short_company(lead["company"])
    city = lead.get("city") or ""
    state = lead.get("state") or ""
    title = clean_text(lead.get("title") or "")
    site = lead["website"]
    primary_title = title.split(",")[0].split("|")[0].strip()

    # 1. LinkedIn OG description - only if it's REAL content (not profile card boilerplate)
    if li and li.get("ok"):
        og = clean_text(li.get("og_desc") or "")
        if og and not is_boilerplate(og) and len(og) > 30:
            # take first sentence, sanity-trim, keep original casing
            sent = og.split(".")[0].split("|")[0].split("·")[0].strip()
            sent = sent.rstrip(",;: ").strip()
            # discard if it starts with "Experience" / "Education" / etc
            if sent and len(sent) > 25 and len(sent) < 180:
                if not re.match(r"^(experience|education|location|view|join)\b", sent, re.I):
                    return f'Pulled up your LinkedIn before reaching out - the "{sent}" line stood out.'

    # 2. Site-derived hook - prefer the H1 (real positioning copy on most real-estate sites)
    if audit:
        hook_src = clean_text(audit.get("h1") or "")
        if not hook_src or len(hook_src) < 20:
            hook_src = clean_text(audit.get("meta_description") or "")
        if hook_src and len(hook_src) > 20 and len(hook_src) < 180:
            sent = hook_src.split(".")[0].split("|")[0].strip().rstrip(",;: ")
            if len(sent) > 20:
                return f'The "{sent}" line on {site} caught my eye when I pulled it up.'

    # 3. Geographic / role fallback (round-1 style)
    geo = ", ".join(x for x in (city, state) if x)
    if geo and primary_title:
        return f"Saw {company}'s footprint in {geo} and your role as {primary_title} - went looking at the site."
    if geo:
        return f"Saw {company}'s footprint in {geo} and went looking at the site itself."
    if primary_title:
        return f"Saw your role as {primary_title} at {company} and went looking at the site."
    return f"Was looking at {site} this week and the lead-capture stack stood out."

def specific_gap(audit, lead):
    if not audit:
        return f"From the outside, the lead-capture stack on {lead['website']} looks light - no obvious chatbot or CRM hook."
    if audit["chat"] == "No":
        return f"There's no live chat or AI assistant on {lead['website']}, so anyone landing after 6pm has nothing to do but close the tab."
    if audit["crm"] == "No":
        return f"Form fills on {lead['website']} don't appear to route into a CRM - they sit in an inbox and quietly age."
    if audit["analytics"] == "No":
        return f"There's no analytics on {lead['website']}, so it's impossible to tell which pages actually convert."
    if audit["ai_tools"] == "No":
        return f"Every prospect question on {lead['website']} routes to a human - works at low volume, breaks the second traffic scales."
    if audit["remarketing"] == "No":
        return f"No remarketing pixels on {lead['website']} - warm traffic is being lost on exit."
    return f"The biggest leak on {lead['website']} right now is the gap between visitor intent and a human picking up."

def top_gap_phrase(audit):
    if not audit: return "the chat / CRM gap"
    if audit["chat"] == "No": return "the after-hours chat gap"
    if audit["crm"] == "No": return "the CRM gap"
    if audit["analytics"] == "No": return "the analytics blind spot"
    if audit["ai_tools"] == "No": return "the missing AI layer"
    return "the conversion-side gap"

def score_line(audit, company):
    if not audit:
        return ""
    score = audit["score"]
    return f"{company} came in at {score}/10 on the audit; "

def build_email(lead, audit, li, subject_idx):
    first = lead["first_name"]
    company = short_company(lead["company"])
    site = lead["website"]
    breaker = ice_breaker(lead, audit, li)
    gap = specific_gap(audit, lead)
    sl = score_line(audit, company)
    top_gap = top_gap_phrase(audit)
    score_sentence = ""
    if sl:
        # tie back to the gap
        if audit:
            if audit["chat"] == "No":
                score_sentence = f"{sl}the chat + CRM gap is the biggest lever."
            elif audit["crm"] == "No":
                score_sentence = f"{sl}the CRM gap is the biggest lever."
            else:
                score_sentence = f"{sl}the conversion-side gaps are the biggest lever."

    subject = SUBJECTS[subject_idx % 3].format(
        first=first, site=site, company_short=shortname(company)
    )

    # tighter score line
    if audit:
        score = audit["score"]
        body_score = f"Score: {score}/10. Top lever: {top_gap}."
    else:
        body_score = f"Top lever: {top_gap}."

    body = (
        f"Hi {first},\n\n"
        f"{breaker}\n\n"
        f"{gap}\n\n"
        f"GenosAI builds the AI layer that closes it - 24/7 chatbot, CRM + follow-up automation, "
        f"and a high-end Next.js or 3D rebuild when the site needs it (no WordPress lookalikes).\n\n"
        f"Free 9-slide audit of {site} attached. {body_score}\n\n"
        f"Worth 15 minutes next week if any of it lands?\n\n"
        f"Best,\n"
        f"Rohan Malik\n"
        f"GenosAI - genosai.tech"
    )

    fu1 = (
        f"Hi {first},\n\n"
        f"Bumping the {company} audit deck - still attached if it got buried.\n\n"
        f"If reading isn't the easy option this week, I can drop a 3-min Loom walking through the top finding "
        f"({top_gap}) instead.\n\n"
        f"Either way, no pressure.\n\n"
        f"Best,\n"
        f"Rohan"
    )

    fu2 = (
        f"Hi {first},\n\n"
        f"Last note on this. The {company} audit is still in your inbox - yours to keep, no reply needed.\n\n"
        f"If the timing's wrong, totally fair. If {top_gap} keeps showing up in your numbers later, you know where to find me.\n\n"
        f"Best,\n"
        f"Rohan Malik\n"
        f"GenosAI · genosai.tech"
    )
    return subject, body, fu1, fu2

# ---- audit deck slide content (matches round-1 schema) ----

def slide_content(lead, audit):
    company = short_company(lead["company"])
    full = lead["full_name"]
    title = lead["title"]
    site = lead["website"]
    if audit:
        platform = audit["platform"]
        score = f"{audit['score']}"
        priority = audit["priority"]
        issues = audit["issues"]
        gaps = audit["gaps"]
        chat = audit["chat"]; crm = audit["crm"]; ai = audit["ai_tools"]
        analytics = audit["analytics"]; cdn = audit["cdn"]
        remarket = audit["remarketing"]; mailing = audit["email_mkt"]; mobile = audit["mobile"]
    else:
        platform = "Custom"
        score = "-"
        priority = "MED"
        issues = ["Site fetch failed during audit - manual review needed",
                  "Likely missing chat / CRM / AI layer based on category norms",
                  "Recommend a quick walkthrough before the call"]
        gaps = ["AI chatbot for 24/7 lead qualification",
                "AI lead capture & CRM automation",
                "Automated follow-up & nurture sequences"]
        chat = crm = ai = analytics = cdn = remarket = mailing = mobile = "?"

    issue_block = "\n".join(f"- {i}" for i in issues)
    gap_block = "\n".join(f"- {g}" for g in gaps)

    s1 = (
        f"{company}\n"
        f"Website and AI Audit\n\n"
        f"Prepared for {full}, {title}\n"
        f"By GenosAI (genosai.tech)\n\n"
        f"What's inside: 9 slides. 3 minutes. Yours either way."
    )
    s2 = (
        f"At a glance\n\n"
        f"Site: {site}\n"
        f"Tech stack: {platform}\n"
        f"Audit score: {score} / 10\n"
        f"Priority: {priority}\n\n"
        f"Top finding: {issues[0] if issues else 'Conversion-side gaps are the biggest lever.'}"
    )
    working = []
    if mobile == "Yes": working.append("Mobile layout works on phones, which a surprising number of competitors still get wrong.")
    if analytics == "Yes": working.append("Analytics is firing, so we can see what visitors actually do once they land.")
    if crm == "Yes": working.append("CRM is connected, so leads that do convert end up somewhere usable.")
    if cdn == "Yes": working.append("CDN is configured, so global page loads aren't dragging.")
    if remarket == "Yes": working.append("Remarketing pixels are set, so warm traffic isn't being lost on exit.")
    if mailing == "Yes": working.append("Email marketing is wired in, which is half the nurture battle solved.")
    if ai == "Yes": working.append("Some AI tooling already detected - good foundation to build on.")
    if not working:
        working = ["Domain and brand are in place - the infrastructure to start improving is ready to use."]
    s3 = "What's working\n\n" + "\n".join(f"- {w}" for w in working[:6]) + \
         "\n\nNot everything on the site is broken. These are real strengths to keep when we rebuild around the gaps."

    leaks = []
    if chat == "No":
        leaks.append("No live chat or chatbot. A visitor who lands at 9pm with a question has nothing to do except close the tab.")
    if ai == "No":
        leaks.append("No AI layer. Every prospect question routes to a human. Works at low volume; breaks the second traffic scales.")
    if crm == "No":
        leaks.append("No CRM connected. Form fills get sent to an inbox and quietly age. There's no scoring, no follow-up logic, no audit trail.")
    if analytics == "No":
        leaks.append("No analytics. You're guessing about what's working - which pages convert, where people drop off.")
    if remarket == "No":
        leaks.append("No remarketing pixels. Warm traffic that doesn't convert on first visit is gone.")
    if mailing == "No":
        leaks.append("No email marketing wiring. Follow-up depends on someone remembering to send it.")
    if platform in ("WordPress", "Wix", "GoDaddy"):
        leaks.append(f"Running on {platform}. It works, but plugins age fast, security patches pile up, and page speed suffers vs Next.js or Webflow.")
    if not leaks:
        leaks = ["Conversion-side gaps - the site captures some leads, but the after-hours and follow-up paths are doing the most damage."]
    s4 = "Where leads are leaking\n\n" + "\n".join(f"- {l}" for l in leaks[:4])

    s5 = ("AI automation opportunities\n\n"
          "- AI chatbot that qualifies leads after hours. Handles the 'do you cover [city]?', 'what's the price range?', 'can I see X?' questions, and books calls into the calendar.\n"
          "- Automated follow-up & booking. Missed call - auto-text. No reply in 48h - nudge. Hot lead - calendar link in the next message.\n"
          "- AI-powered property/listing matching. Visitor tells the bot what they want; the system surfaces the right options without anyone digging.")

    s6 = ("What this is costing you\n\n"
          "The site captures some leads, but the after-hours and follow-up gaps are doing the most damage.\n\n"
          "Most operators in this range see 2-3x more captured leads once chat and follow-up automation are in place.\n\n"
          "Two ways to think about it:\n"
          "- Every after-hours visitor who doesn't get a response is a competitor's problem to solve.\n"
          "- Every prospect your team manually emails for the third time is an hour they're not closing deals.")

    s7 = ("What we'd build\n\n"
          "Phase 1 (weeks 1-3): rebuild on Next.js or Webflow (or a 3D / high-end front end where the brand can carry it). Mobile-first. Clear primary CTA on every page.\n\n"
          "Phase 2 (weeks 2-4): AI chatbot trained on your services, regions, and FAQs. Books calls into your calendar. CRM + follow-up automations wired in.\n\n"
          "Total timeline: usually 4-6 weeks end to end. One project lead, one engineering lead, weekly check-in.")

    s8 = ("What changes in the first 90 days\n\n"
          "Numbers our real estate clients typically see after launch:\n\n"
          "- 40-60% more inbound leads captured, mostly from after-hours chat\n"
          "- 2x faster lead-to-call conversion (the bot replies in seconds)\n"
          "- 10+ hours per week back for the team that used to do manual follow-up\n"
          "- Better quality conversations: by the time a human picks up, the lead is qualified\n\n"
          f"Numbers depend on traffic and offer. Happy to share the case study closest to {company}'s model on the call.")

    s9 = ("Next steps\n\n"
          "A 15-minute call. Bring your top two or three questions. I'll bring the pricing model, a relevant case study, "
          f"and a rough timeline tailored to {company}.\n\n"
          "If now isn't the right window, the deck is yours either way.")

    return [s1, s2, s3, s4, s5, s6, s7, s8, s9]

def main():
    leads = json.loads(LEADS.read_text())
    audits = json.loads(AUDITS.read_text())
    out = {}
    for i, lead in enumerate(leads):
        a_entry = audits.get(lead["email"], {})
        audit = a_entry.get("audit")
        li = a_entry.get("linkedin")
        subject, body, fu1, fu2 = build_email(lead, audit, li, i)
        slides = slide_content(lead, audit)
        out[lead["email"]] = {
            "lead_num": lead["num"],
            "subject": subject,
            "body": body,
            "fu1": fu1,
            "fu2": fu2,
            "slides": slides,
            "audit_summary": audit,
        }
    OUT.write_text(json.dumps(out, indent=2))
    print(f"Wrote {len(out)} content entries -> {OUT}")
    # show one sample
    sample = list(out.values())[0]
    print("--- SAMPLE ---")
    print("subject:", sample["subject"])
    print("body words:", len(sample["body"].split()))
    print(sample["body"][:600])

if __name__ == "__main__":
    main()
