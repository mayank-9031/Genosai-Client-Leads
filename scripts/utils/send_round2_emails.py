"""
Send Round 2 cold emails via SMTP.
Reads Cold Emails sheet rows 45-90, attaches per-lead PPT, sends via GoDaddy SMTP.
"""
import json, smtplib, ssl, time, uuid, openpyxl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formataddr, formatdate
from pathlib import Path

# ── Config ────────────────────────────────────────────────────────────────────
SMTP_HOST  = "smtpout.secureserver.net"
SMTP_PORT  = 465
FROM_EMAIL = "hello@genosai.tech"
FROM_NAME  = "Rohan Malik | GenosAI"
SMTP_PASS  = "rohmay@147"

ROOT       = Path(__file__).parent.parent.parent
DECKS_DIR  = ROOT / "output" / "decks"
XLSX       = ROOT / "output" / "workbooks" / "GenosAI_RealEstate_Leads (2).xlsx"
PATHS_JSON = ROOT / "data" / "pipeline" / "_round2_deck_paths.json"
LOG_FILE   = ROOT / "data" / "pipeline" / "_round2_send_log.json"

LINKEDIN   = "https://www.linkedin.com/in/rohanxmalik/"
WEBSITE    = "https://www.genosai.tech"

DELAY_SECS = 4   # seconds between sends (GoDaddy rate limit buffer)

# ── Signature ─────────────────────────────────────────────────────────────────
SIG_HTML = (
    "<br><br>"
    "Best,<br>"
    "Rohan<br>"
    "Co-Founder &amp; CEO<br>"
    f'<a href="{WEBSITE}" style="color:#7B2FBE;text-decoration:none;font-weight:bold;">GenosAI</a>'
    ' &nbsp;|&nbsp; '
    f'<a href="{LINKEDIN}" style="color:#7B2FBE;text-decoration:none;font-weight:bold;">LinkedIn</a>'
)

SIG_PLAIN = (
    "\r\n\r\nBest,\r\nRohan\r\nCo-Founder & CEO\r\n"
    f"GenosAI — {WEBSITE}\r\nLinkedIn — {LINKEDIN}"
)

# ── Helpers ───────────────────────────────────────────────────────────────────
def strip_old_sig(body: str) -> str:
    idx = body.rfind("\n\nBest,")
    return body[:idx].strip() if idx != -1 else body.strip()

def body_to_html(plain: str) -> str:
    return (
        plain
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\r\n", "<br>\n")
        .replace("\n", "<br>\n")
    )

def make_message(to_email: str, subject: str, body_plain: str, ppt_path) -> MIMEMultipart:
    clean = strip_old_sig(body_plain)
    clean_crlf = clean.replace("\r\n", "\n").replace("\n", "\r\n")

    msg = MIMEMultipart("mixed")
    msg["From"]       = formataddr((FROM_NAME, FROM_EMAIL))
    msg["To"]         = to_email
    msg["Subject"]    = subject
    msg["Date"]       = formatdate(localtime=True)
    msg["Message-ID"] = f"<{uuid.uuid4()}@genosai.tech>"
    msg["Reply-To"]   = formataddr((FROM_NAME, FROM_EMAIL))

    # Multipart/alternative: plain + HTML
    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(clean_crlf + SIG_PLAIN, "plain", "utf-8"))
    html_body = (
        "<html><body style='font-family:Arial,sans-serif;font-size:14px;"
        f"color:#222;line-height:1.6;'>{body_to_html(clean)}{SIG_HTML}</body></html>"
    )
    alt.attach(MIMEText(html_body, "html", "utf-8"))
    msg.attach(alt)

    # PPT attachment
    ppt = Path(ppt_path) if ppt_path else None
    if ppt and ppt.exists():
        with open(ppt, "rb") as f:
            att = MIMEBase(
                "application",
                "vnd.openxmlformats-officedocument.presentationml.presentation"
            )
            att.set_payload(f.read())
        encoders.encode_base64(att)
        att.add_header("Content-Disposition", "attachment", filename=ppt.name)
        msg.attach(att)
    else:
        print(f"  WARNING: PPT not found at {ppt_path}")

    return msg

def main():
    deck_paths = json.loads(PATHS_JSON.read_text())
    wb = openpyxl.load_workbook(XLSX)
    ws = wb["Cold Emails"]

    leads = []
    for row in range(2, ws.max_row + 1):
        num = ws.cell(row, 1).value
        if num is not None and num >= 44:
            leads.append({
                "num":     num,
                "company": ws.cell(row, 3).value or "",
                "email":   ws.cell(row, 4).value or "",
                "subject": ws.cell(row, 5).value or "",
                "body":    ws.cell(row, 6).value or "",
            })

    print(f"Found {len(leads)} round-2 leads.\n")

    # Resume: load existing log so already-sent emails are skipped
    log = json.loads(LOG_FILE.read_text()) if LOG_FILE.exists() else {}
    already = sum(1 for v in log.values() if v == "sent")
    if already:
        print(f"Resuming — {already} already sent, skipping those.\n")

    ctx = ssl.create_default_context()

    BATCH_SIZE = 20   # reconnect every N emails to avoid GoDaddy session limit
    smtp = None

    def new_conn():
        s = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ctx)
        s.login(FROM_EMAIL, SMTP_PASS)
        print("SMTP session opened.")
        return s

    batch_count = 0
    smtp = new_conn()

    for i, lead in enumerate(leads, 1):
        email = lead["email"]
        rel   = deck_paths.get(email)
        ppt   = (DECKS_DIR / rel) if rel else None

        if log.get(email) == "sent":
            print(f"[{i:02d}/{len(leads)}] SKIP  -> {email} (already sent)")
            continue

        # Reconnect every BATCH_SIZE sends
        if batch_count > 0 and batch_count % BATCH_SIZE == 0:
            try:
                smtp.quit()
            except Exception:
                pass
            time.sleep(10)   # wait before new session
            smtp = new_conn()

        try:
            msg = make_message(email, lead["subject"], lead["body"], ppt)
            smtp.send_message(msg)
            log[email] = "sent"
            LOG_FILE.write_text(json.dumps(log, indent=2))
            print(f"[{i:02d}/{len(leads)}] SENT  -> {email}  ({lead['company']})")
            batch_count += 1
        except Exception as e:
            log[email] = f"ERROR: {e}"
            LOG_FILE.write_text(json.dumps(log, indent=2))
            print(f"[{i:02d}/{len(leads)}] FAIL  -> {email}: {e}")
            # Reconnect on SMTP errors
            try:
                smtp.quit()
            except Exception:
                pass
            time.sleep(5)
            smtp = new_conn()
            batch_count = 0

        time.sleep(DELAY_SECS)

    try:
        smtp.quit()
    except Exception:
        pass

    LOG_FILE.write_text(json.dumps(log, indent=2))
    sent  = sum(1 for v in log.values() if v == "sent")
    fails = sum(1 for v in log.values() if v != "sent")
    print(f"\nDone. {sent} sent, {fails} failed. Log -> scripts/_round2_send_log.json")

if __name__ == "__main__":
    main()
