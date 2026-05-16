"""
Round 3 content generator - restaurants & dentists.
Produces per-lead cold email + audit deck slide content.
Only generates emails for leads that have an email address.
Writes: scripts/_round3_content.json
"""
import json, re
from pathlib import Path

ROOT   = Path(r"c:/Users/91638/Desktop/GenosApollp clients")
LEADS  = ROOT / "data" / "pipeline" / "_round3_leads.json"
AUDITS = ROOT / "data" / "pipeline" / "_round3_audits.json"
OUT    = ROOT / "data" / "pipeline" / "_round3_content.json"


def short_company(c: str) -> str:
    return c.split(",")[0].split("|")[0].strip()[:40].strip()


def shortname(c: str) -> str:
    return re.sub(r"[\s,].*", "", c).strip()[:25] or c


def clean(s: str) -> str:
    if not s:
        return s
    for a, b in [("'","'"),("'","'"),("’","'"),("‘","'"),
                 ("“",'"'),("”",'"'),("-","-"),("…","..."),
                 ("â","'"),("â","'")]:
        s = s.replace(a, b)
    s = re.sub(r"[^\x00-\x7F]+", "", s)
    return re.sub(r"\s+", " ", s).strip()


# ─── Ice-breaker ─────────────────────────────────────────────────────────────

def ice_breaker(lead: dict, audit: dict | None) -> str:
    company = short_company(lead["company"])
    site    = lead["website"] or company
    city    = lead.get("city") or ""
    state   = lead.get("state") or ""
    cat     = lead.get("category", "restaurant")

    if audit:
        hook = clean(audit.get("h1") or audit.get("meta_description") or audit.get("title") or "")
        if hook and 20 < len(hook) < 180:
            sent = hook.split(".")[0].split("|")[0].strip().rstrip(",;:")
            if len(sent) > 20:
                return f'The "{sent}" line on {site} caught my eye when I pulled it up.'

    geo = ", ".join(x for x in (city, state) if x)
    if geo:
        if cat == "restaurant":
            return f"Was browsing restaurants in {geo} and {company} stood out - went looking at the site."
        else:
            return f"Was looking at dental practices in {geo} and {company} came up - pulled up the site."
    if cat == "restaurant":
        return f"Was looking at {site} this week and the online presence stood out."
    return f"Was looking at {site} this week and the patient booking experience stood out."


# ─── Specific gap ────────────────────────────────────────────────────────────

def specific_gap(lead: dict, audit: dict | None) -> str:
    site = lead["website"] or lead["company"]
    cat  = lead.get("category", "restaurant")

    if not audit:
        if cat == "restaurant":
            return (f"From the outside, {site} doesn't appear to have online reservations or "
                    "an ordering integration - anyone landing after hours has no clear next step.")
        return (f"From the outside, {site} doesn't appear to have an online booking system - "
                "patients who land after hours have no way to schedule.")

    if cat == "restaurant":
        if audit["booking"] == "No":
            return (f"There's no online reservation system on {site} - "
                    "a diner who wants to book a table at 10pm has to call or hope for a walk-in.")
        if audit["online_order"] == "No":
            return (f"There's no direct online ordering on {site} - "
                    "every order goes through a third-party that keeps 15-30% of the ticket.")
        if audit["chat"] == "No":
            return (f"There's no chatbot on {site} - "
                    "questions about the menu, hours, or specials after hours go completely unanswered.")
        if audit["reviews"] == "No":
            return (f"No review management on {site} - "
                    "new diners are deciding based on whatever shows up on Google with no curation.")
        if audit["analytics"] == "No":
            return (f"No analytics on {site} - "
                    "there's no way to tell which pages drive covers or where diners drop off.")
        return (f"The biggest gap on {site} is the space between a diner's intent and a "
                "confirmed reservation or order.")
    else:
        if audit["booking"] == "No":
            return (f"There's no online booking on {site} - "
                    "a patient who finds the practice at 11pm has to wait until the office opens to schedule.")
        if audit["chat"] == "No":
            return (f"There's no chatbot on {site} - "
                    "after-hours questions about insurance, availability, or procedures go unanswered.")
        if audit["reviews"] == "No":
            return (f"No review management on {site} - "
                    "patients choose their dentist by star rating, and there's no system to solicit them.")
        if audit["email_mkt"] == "No":
            return (f"No email marketing on {site} - "
                    "recall and follow-up depend on someone in the office manually reaching out.")
        return (f"The biggest gap on {site} is the space between a patient's search and a booked appointment.")


def top_gap_phrase(lead: dict, audit: dict | None) -> str:
    cat = lead.get("category", "restaurant")
    if not audit:
        return "the booking gap" if cat == "restaurant" else "the appointment booking gap"
    if cat == "restaurant":
        if audit["booking"] == "No":   return "the online reservation gap"
        if audit["online_order"] == "No": return "the direct ordering gap"
        if audit["chat"] == "No":      return "the after-hours chat gap"
        if audit["reviews"] == "No":   return "the review management gap"
        return "the conversion-side gap"
    else:
        if audit["booking"] == "No":   return "the online booking gap"
        if audit["chat"] == "No":      return "the after-hours chat gap"
        if audit["reviews"] == "No":   return "the review management gap"
        if audit["email_mkt"] == "No": return "the recall / follow-up gap"
        return "the patient conversion gap"


# ─── Email builder ───────────────────────────────────────────────────────────

RESTAURANT_SUBJECTS = [
    "{company_short} - quick note on {top_gap}",
    "Found something on {site}",
    "{company_short} - 2 min read",
]
DENTIST_SUBJECTS = [
    "{company_short} - quick note on {top_gap}",
    "Found something on {site}",
    "{company_short} - 2 min read",
]


def build_email(lead: dict, audit: dict | None, idx: int):
    cat     = lead.get("category", "restaurant")
    company = short_company(lead["company"])
    site    = lead["website"] or company
    top_gap = top_gap_phrase(lead, audit)
    breaker = ice_breaker(lead, audit)
    gap     = specific_gap(lead, audit)

    subj_pool = RESTAURANT_SUBJECTS if cat == "restaurant" else DENTIST_SUBJECTS
    subject = subj_pool[idx % 3].format(
        company_short=shortname(company), top_gap=top_gap, site=site
    )

    if audit:
        score = audit["score"]
        body_score = f"Score: {score}/10. Top lever: {top_gap}."
    else:
        body_score = f"Top lever: {top_gap}."

    if cat == "restaurant":
        service_line = (
            "GenosAI builds the AI layer that closes it - online reservation + ordering integration, "
            "AI chatbot for after-hours inquiries, automated review requests, and a high-end Next.js or "
            "3D rebuild when the site needs it."
        )
        call_to_action = "Worth 15 minutes next week if any of it lands?"
    else:
        service_line = (
            "GenosAI builds the AI layer that closes it - online booking integration, AI chatbot for "
            "24/7 patient questions, automated recall & review requests, and a modern Next.js or 3D "
            "rebuild when the site needs it."
        )
        call_to_action = "Worth 15 minutes next week to walk through the findings?"

    body = (
        f"Hi there,\n\n"
        f"{breaker}\n\n"
        f"{gap}\n\n"
        f"{service_line}\n\n"
        f"Free 9-slide audit of {site} attached. {body_score}\n\n"
        f"{call_to_action}\n\n"
        f"Best,\n"
        f"Rohan Malik\n"
        f"GenosAI - genosai.tech"
    )

    fu1 = (
        f"Hi,\n\n"
        f"Bumping the {company} audit deck - still attached if it got buried.\n\n"
        f"If reading isn't easy this week, I can drop a 3-min Loom walking through the top finding "
        f"({top_gap}) instead.\n\n"
        f"Either way, no pressure.\n\n"
        f"Best,\nRohan"
    )

    fu2 = (
        f"Hi,\n\n"
        f"Last note on this. The {company} audit is still in your inbox - yours to keep, no reply needed.\n\n"
        f"If the timing's wrong, totally fair. If {top_gap} keeps showing up later, you know where to find me.\n\n"
        f"Best,\nRohan Malik\nGenosAI · genosai.tech"
    )

    return subject, body, fu1, fu2


# ─── Slide content ───────────────────────────────────────────────────────────

def slide_content(lead: dict, audit: dict | None) -> list[str]:
    company = short_company(lead["company"])
    site    = lead["website"] or company
    cat     = lead.get("category", "restaurant")

    if audit:
        platform = audit["platform"]
        score    = f"{audit['score']}"
        priority = audit["priority"]
        issues   = audit["issues"]
        gaps     = audit["gaps"]
        booking  = audit["booking"]; chat = audit["chat"]
        analytics= audit["analytics"]; mobile = audit["mobile"]
        reviews  = audit["reviews"]; email_mkt = audit["email_mkt"]
        order    = audit.get("online_order","?")
        remarket = audit["remarketing"]
    else:
        platform = "Custom"
        score = "-"; priority = "MED"
        if cat == "restaurant":
            issues = ["No online reservation system detected",
                      "No direct ordering integration visible",
                      "No review management tool found"]
            gaps   = ["Online reservation + ordering integration",
                      "AI chatbot for after-hours inquiries",
                      "Automated review request system"]
        else:
            issues = ["No online booking system detected",
                      "No after-hours chatbot found",
                      "No automated recall / review system found"]
            gaps   = ["Online appointment booking integration",
                      "AI chatbot for 24/7 patient questions",
                      "Automated recall & review request system"]
        booking=chat=analytics=mobile=reviews=email_mkt=order=remarket="?"

    issue_block = "\n".join(f"- {i}" for i in issues)

    s1 = (
        f"{company}\n"
        f"Website & AI Audit\n\n"
        f"Prepared by GenosAI (genosai.tech)\n\n"
        f"What's inside: 9 slides. 3 minutes. Yours either way."
    )

    s2 = (
        f"At a glance\n\n"
        f"Site: {site}\n"
        f"Tech stack: {platform}\n"
        f"Audit score: {score} / 10\n"
        f"Priority: {priority}\n\n"
        f"Top finding: {issues[0]}"
    )

    working = []
    if mobile   == "Yes": working.append("Mobile layout renders correctly - most visitors are on phones.")
    if analytics== "Yes": working.append("Analytics is firing - visitor behaviour is being tracked.")
    if booking  == "Yes": working.append("Online booking / reservation system is in place - great foundation.")
    if reviews  == "Yes": working.append("Review management tool detected - reputation is being monitored.")
    if email_mkt== "Yes": working.append("Email marketing is connected - the nurture channel is open.")
    if remarket == "Yes": working.append("Remarketing pixels set - warm traffic is being recaptured.")
    if not working:
        working = ["Domain and brand are established - the infrastructure to improve on is in place."]
    s3 = ("What's working\n\n"
          + "\n".join(f"- {w}" for w in working[:5])
          + "\n\nNot everything is broken. These are real strengths to build on.")

    leaks = []
    if cat == "restaurant":
        if booking  == "No": leaks.append("No online reservations. A diner who wants to book at 10pm calls or leaves.")
        if order    == "No": leaks.append("No direct ordering. Third-party commissions (15-30%) eat the margin on every ticket.")
        if chat     == "No": leaks.append("No chatbot. After-hours menu, hours, and event questions go completely dark.")
        if reviews  == "No": leaks.append("No review management. New diners decide based on whatever Google shows - unmanaged.")
        if analytics== "No": leaks.append("No analytics. No visibility into what pages drive bookings or where visitors drop.")
        if email_mkt== "No": leaks.append("No email marketing. There's no automated way to bring regulars back.")
    else:
        if booking  == "No": leaks.append("No online booking. A patient who finds you at 11pm has to wait until the office opens.")
        if chat     == "No": leaks.append("No chatbot. After-hours insurance/availability questions go unanswered.")
        if reviews  == "No": leaks.append("No review management. Patients choose by star rating - no system to earn more.")
        if email_mkt== "No": leaks.append("No recall emails. Follow-up is fully manual and easy to forget.")
        if analytics== "No": leaks.append("No analytics. No way to tell which services attract patients or where they drop off.")
        if remarket == "No": leaks.append("No remarketing. Patients who visited but didn't book are gone.")
    if not leaks:
        leaks = ["Conversion-side gaps - the site is solid but the after-hours and follow-up paths need work."]
    s4 = "Where leads are leaking\n\n" + "\n".join(f"- {l}" for l in leaks[:4])

    if cat == "restaurant":
        s5 = (
            "AI automation opportunities\n\n"
            "- Online reservation + waitlist AI: handles bookings, special requests, and capacity management automatically.\n"
            "- AI chatbot: answers menu, hours, allergen, and event questions 24/7. Reduces calls by 60-80%.\n"
            "- Direct ordering integration: native ordering removes the third-party middleman and keeps the margin.\n"
            "- Automated review requests: sends a personalised Google/Yelp request after each visit."
        )
        s6 = (
            "What this is costing you\n\n"
            "Every unanswered after-hours inquiry is a table another restaurant fills.\n"
            "Every third-party order is 15-30% of revenue leaving the business.\n\n"
            "Two ways to think about it:\n"
            "- A table that can't be booked at 10pm is a table that books somewhere else.\n"
            "- A 30% commission on $50 is $15 per order, every order, forever."
        )
        s7 = (
            "What we'd build\n\n"
            "Phase 1 (weeks 1-2): Online reservation + ordering integration or direct native system.\n"
            "Phase 2 (weeks 2-4): AI chatbot trained on your menu, hours, and events. "
            "Automated review requests after each visit.\n"
            "Optional: Full Next.js or 3D visual rebuild for a flagship web presence.\n\n"
            "Total timeline: 3-6 weeks end to end."
        )
        s8 = (
            "What changes in the first 90 days\n\n"
            "Numbers our restaurant clients typically see:\n\n"
            "- 3-5x more online reservations (no friction, no phone tag)\n"
            "- 20-40% reduction in third-party commission spend via direct ordering\n"
            "- 2x more Google reviews (automated ask at the right moment)\n"
            "- 60-80% fewer after-hours calls (chatbot handles the common questions)\n\n"
            f"Happy to share the case study closest to {company}'s model on the call."
        )
    else:
        s5 = (
            "AI automation opportunities\n\n"
            "- Online booking integration: patients book 24/7 without calling. Connects to your existing calendar.\n"
            "- AI chatbot: handles insurance, hours, and new-patient questions around the clock. "
            "Qualifies and books without staff involvement.\n"
            "- Automated recall & reminders: 6-month recall emails, appointment reminders, "
            "and post-visit review requests - all set-and-forget.\n"
            "- Reputation management: automated review request after each visit builds your Google/Healthgrades rating."
        )
        s6 = (
            "What this is costing you\n\n"
            "Every patient who searches at 11pm and can't book moves to the next practice in the list.\n"
            "Every recall that relies on a staff member remembering to send it gets missed.\n\n"
            "Two ways to think about it:\n"
            "- A missed recall = a patient who walks into a competitor's office next quarter.\n"
            "- An unanswered chatbot inquiry = a new patient who booked somewhere else by morning."
        )
        s7 = (
            "What we'd build\n\n"
            "Phase 1 (weeks 1-2): Online booking integration (or native calendar). "
            "AI chatbot trained on your services, insurance, and FAQs.\n"
            "Phase 2 (weeks 2-4): Automated recall emails, post-visit review requests, "
            "and remarketing campaigns.\n"
            "Optional: Full Next.js or 3D visual rebuild to match the standard of premium practices.\n\n"
            "Total timeline: 3-5 weeks end to end."
        )
        s8 = (
            "What changes in the first 90 days\n\n"
            "Numbers our dental clients typically see:\n\n"
            "- 40-60% more new patients booked online (zero phone friction)\n"
            "- 3x more Google reviews (automated ask after checkout)\n"
            "- 30-50% reduction in no-shows (automated reminder sequences)\n"
            "- 5-10 hrs/week back for front-desk staff handling calls the chatbot now handles\n\n"
            f"Happy to share the case study closest to {company}'s model on the call."
        )

    s9 = (
        "Next steps\n\n"
        "A 15-minute call. Bring your top two or three questions. I'll bring pricing, "
        f"a relevant case study, and a rough timeline tailored to {company}.\n\n"
        "If now isn't the right window, the deck is yours either way."
    )

    return [s1, s2, s3, s4, s5, s6, s7, s8, s9]


# ─── Main ────────────────────────────────────────────────────────────────────

def main():
    leads  = json.loads(LEADS.read_text(encoding="utf-8"))
    audits = json.loads(AUDITS.read_text(encoding="utf-8"))
    out    = {}

    for i, lead in enumerate(leads):
        a_entry = audits.get(str(lead["num"]), audits.get(lead["num"], {}))
        audit   = a_entry.get("audit")
        subject, body, fu1, fu2 = build_email(lead, audit, i)
        slides  = slide_content(lead, audit)

        out[lead["num"]] = {
            "lead_num":      lead["num"],
            "company":       lead["company"],
            "category":      lead["category"],
            "email":         lead["email"],
            "has_email":     bool(lead["email"]),
            "subject":       subject,
            "body":          body,
            "fu1":           fu1,
            "fu2":           fu2,
            "slides":        slides,
            "audit_summary": audit,
        }

    OUT.write_text(json.dumps(out, indent=2), encoding="utf-8")
    with_email = sum(1 for v in out.values() if v["has_email"])
    print(f"Wrote {len(out)} content entries ({with_email} with email) -> {OUT}")
    sample = next(v for v in out.values() if v["has_email"])
    print("--- SAMPLE ---")
    print("subject:", sample["subject"])
    print(sample["body"][:400])


if __name__ == "__main__":
    main()
