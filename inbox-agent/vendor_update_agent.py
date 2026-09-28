"""Vendor/landlord weekly campaign update agent for IB Property.

Drafts weekly campaign update emails for each of Edward Ghattas's own active
listings, saves them to Outlook Drafts (and Gmail Drafts, if that auth is
up) for his review. Never auto-sends.

"What's actually on market right now" comes from EagleAgent (the CRM, see
eagle_client.py / campaign_source.py), filtered to Eddie's own book by agent
email -- NOT a doc Eddie has to hand-maintain. That was the previous design
(the Campaign Enquiry Reply Templates Google Doc's ACTIVE CAMPAIGNS section)
and its whole problem: the doc sat empty for the entire 2+ months it was in
use, so the agent silently drafted nothing, every single week, the entire
time. Eddie was explicit this rebuild must not repeat that shape: he
shouldn't be doing data entry for the agent to read back.

The weekly narrative -- who's enquired, what's changed, what's next -- comes
from two things merged together:
  1. vendor_update_ledger.py: a small per-campaign continuity record (known
     leads + their status), carried forward week to week so the draft can
     say "as mentioned last week" honestly instead of resetting. Seeded on
     a campaign's first run by backfilling from Eddie's own most recent past
     update email for that address (see backfill_ledger_entry), not blank.
  2. This week's email activity (gather_weekly_activity, Sent + Inbox,
     address-anchored -- NOT Eagle's `reaId`, which was checked live and
     confirmed to not hold portals' real Property IDs for real listings).

Email-mining can only see correspondence Eddie was actually on -- it cannot
see phone calls or other CRM-only activity, so it will always be a partial
picture next to Eagle's own enquiry/inspection counts. claude_draft_vendor_update
is told to flag that gap honestly rather than imply completeness.

Usage:
    python vendor_update_agent.py             # creates real Outlook/Gmail drafts
    python vendor_update_agent.py --dry-run    # preview only, saves nothing
"""

import base64
import html as _html
import json
import logging
import os
import re
import sys
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Dict, List, Optional, Tuple

import anthropic
from dotenv import load_dotenv

from campaign_source import Campaign, EagleCampaignSource
from eagle_client import EagleClient, EagleAuthRequired
from gmail_client import GmailClient, GmailAuthRequired
from outlook_client import OutlookClient, OutlookAuthRequired
from vendor_update_ledger import Lead, LedgerEntry, load_ledger, save_ledger

# ─── Bootstrap ───────────────────────────────────────────────────────────────

load_dotenv()

LOG_PATH = "/tmp/vendor-update-agent.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOG_PATH, encoding="utf-8"),
    ],
)
logger = logging.getLogger(__name__)

# Eddie's identity in Eagle -- email is primary (see campaign_source.py),
# name is only a fallback for a property whose agent record has no email.
_AGENT_NAME_FALLBACK = "Edward Ghattas"

# Matches the real range of subject lines Eddie's own past weekly updates
# have actually used -- "Campaign Update", "Leasing Update", "Week N" alone
# (e.g. "1/2-4 Atchison Street, St Leonards - Sale Campaign Week 1") -- a
# literal "campaign update" check alone (the original version of this file)
# would silently miss real history like "176 Victoria Street, Potts Point -
# Week 7 Leasing Update".
_PAST_UPDATE_SUBJECT_RE = re.compile(r"campaign update|leasing update|week\s*\(?\d+", re.IGNORECASE)

# ─── Outlook helpers ──────────────────────────────────────────────────────────


def _body_to_plain(msg: Dict) -> str:
    """Strip HTML from a Graph message dict and return plain text."""
    raw = msg.get("body", {}).get("content", "") if isinstance(msg.get("body"), dict) else ""
    decoded = _html.unescape(raw).replace("\xa0", " ")
    plain = re.sub(r"<[^>]+>", " ", decoded)
    return re.sub(r"\s{2,}", " ", plain).strip()


def _search_sent_items(
    outlook: OutlookClient, query: str, top: int = 10
) -> List[Dict]:
    """Search Sent Items by keyword, returning messages with body + recipients."""
    result = outlook._get(
        "/me/mailFolders/sentitems/messages",
        params={
            "$search": f'"{query}"',
            "$top": top,
            "$select": "id,subject,body,sentDateTime,toRecipients",
        },
    )
    return result.get("value", []) if result else []


def _search_inbox(outlook: OutlookClient, query: str, top: int = 10) -> List[Dict]:
    """Search whole mailbox (inbox + sent) by keyword."""
    result = outlook._get(
        "/me/messages",
        params={
            "$search": f'"{query}"',
            "$top": top,
            "$select": "id,subject,body,receivedDateTime,from",
        },
    )
    return result.get("value", []) if result else []


def _short_address(address: str) -> str:
    """Return a short search-friendly form of the address (number + first street word)."""
    m = re.search(r"(\d+[\w/]*\s+\w+)", address)
    return m.group(1) if m else address[:30]


def find_landlord_email(outlook: OutlookClient, address: str) -> Optional[str]:
    """Find the landlord email by searching Sent Items for past campaign updates."""
    short_addr = _short_address(address)
    msgs = _search_sent_items(outlook, short_addr, top=20)
    for msg in msgs:
        subj = msg.get("subject", "")
        if _PAST_UPDATE_SUBJECT_RE.search(subj):
            for recipient in msg.get("toRecipients", []):
                email_addr = recipient.get("emailAddress", {}).get("address", "")
                name = recipient.get("emailAddress", {}).get("name", "")
                # Skip internal IB Property addresses
                if email_addr and "ibproperty.com.au" not in email_addr.lower():
                    logger.info(
                        "  Found landlord email for %s: %s <%s>",
                        address, name, email_addr,
                    )
                    return email_addr
    return None


def find_landlord_first_name(outlook: OutlookClient, address: str) -> Optional[str]:
    """Find the landlord's first name from the greeting in a past campaign update."""
    short_addr = _short_address(address)
    msgs = _search_sent_items(outlook, short_addr, top=20)
    for msg in msgs:
        subj = msg.get("subject", "")
        if _PAST_UPDATE_SUBJECT_RE.search(subj):
            plain = _body_to_plain(msg)
            m = re.match(r"Hi\s+(\w+)", plain.strip())
            if m:
                return m.group(1)
    return None


def resolve_landlord_contact(outlook: OutlookClient, campaign: Campaign) -> Tuple[Optional[str], Optional[str]]:
    """Return (email, first_name). Sent Items search is primary -- proven
    already to work (correctly found Sandra Odorisio for Sailors Bay Road).
    Eagle's own vendor contact data (campaign.landlord_contacts) is fallback
    only: confirmed live that 4 of 5 sampled active properties come back
    with an empty vendors list, so it can't be the primary source."""
    email = find_landlord_email(outlook, campaign.address)
    first_name = find_landlord_first_name(outlook, campaign.address)

    if not email and campaign.landlord_contacts:
        fallback_email = campaign.landlord_contacts[0].email
        if fallback_email:
            email = fallback_email
            logger.info("  No landlord email in Sent Items -- using Eagle vendor contact fallback: %s", email)

    if not first_name and campaign.landlord_contacts:
        fallback_name = campaign.landlord_contacts[0].first_name
        if fallback_name:
            first_name = fallback_name

    return email, first_name


def gather_weekly_activity(
    outlook: OutlookClient, address: str, days: int = 7
) -> Tuple[List[str], List[str]]:
    """
    Return (inbox_snippets, sent_snippets) for emails mentioning address in the last N days.

    Each snippet is "subject | date | snippet".
    """
    since = datetime.utcnow() - timedelta(days=days)
    short_addr = _short_address(address)

    inbox_snippets: List[str] = []
    sent_snippets: List[str] = []

    # Inbox / all messages
    for msg in _search_inbox(outlook, short_addr, top=20):
        dt_raw = msg.get("receivedDateTime", "")
        try:
            dt = datetime.strptime(dt_raw[:19], "%Y-%m-%dT%H:%M:%S")
            if dt < since:
                continue
        except ValueError:
            pass
        subj = msg.get("subject", "")
        plain = _body_to_plain(msg)[:300]
        inbox_snippets.append(f"Subject: {subj} | Date: {dt_raw[:10]} | {plain}")

    # Sent Items
    for msg in _search_sent_items(outlook, short_addr, top=20):
        dt_raw = msg.get("sentDateTime", "")
        try:
            dt = datetime.strptime(dt_raw[:19], "%Y-%m-%dT%H:%M:%S")
            if dt < since:
                continue
        except ValueError:
            pass
        subj = msg.get("subject", "")
        # Skip the previous campaign update itself
        if _PAST_UPDATE_SUBJECT_RE.search(subj):
            continue
        plain = _body_to_plain(msg)[:300]
        sent_snippets.append(f"Subject: {subj} | Date: {dt_raw[:10]} | {plain}")

    return inbox_snippets, sent_snippets


def fetch_style_examples(outlook: OutlookClient) -> List[str]:
    """Return plain-text excerpts from Eddie's recent campaign update emails (style ref)."""
    msgs = _search_sent_items(outlook, "campaign update", top=10)
    examples = []
    for msg in msgs:
        subj = msg.get("subject", "")
        if not _PAST_UPDATE_SUBJECT_RE.search(subj):
            continue
        plain = _body_to_plain(msg)
        # Trim to just the outgoing portion (before any quoted reply)
        plain = plain[:1500]
        examples.append(f"Subject: {subj}\n{plain}")
        if len(examples) >= 3:
            break
    return examples


# ─── Continuity ledger ─────────────────────────────────────────────────────────


def backfill_ledger_entry(outlook: OutlookClient, ai: anthropic.Anthropic, address: str) -> LedgerEntry:
    """Seed a new campaign's ledger entry from Eddie's own most recent past
    weekly update email for this address, per his explicit instruction not
    to start blank for a campaign with real history. Returns a fresh, empty
    entry if no past update is found -- a genuinely new campaign has nothing
    to backfill, which is correct, not a failure."""
    short_addr = _short_address(address)
    msgs = _search_sent_items(outlook, short_addr, top=20)
    candidates = [
        m for m in msgs
        if _PAST_UPDATE_SUBJECT_RE.search(m.get("subject", ""))
        and not m.get("subject", "").lower().startswith("re:")
    ]
    if not candidates:
        return LedgerEntry(address=address)

    latest = max(candidates, key=lambda m: m.get("sentDateTime", ""))
    body = _body_to_plain(latest)[:3000]

    prompt = (
        "This is a real weekly campaign update email Edward Ghattas already sent "
        "to a landlord about one of his own listings. Extract exactly what it says, "
        "inventing nothing:\n"
        '- "week_number": the week number mentioned (integer), or 0 if none is stated\n'
        '- "leads": a list of {"name": "...", "status": "..."} for every named '
        "prospect/enquirer/lead mentioned, with their status described as closely "
        "to the email's own wording as possible\n\n"
        f"Subject: {latest.get('subject', '')}\n"
        f"Body:\n{body}\n\n"
        'Reply with ONLY valid JSON: {"week_number": <int>, "leads": [{"name": "...", "status": "..."}]}'
    )
    try:
        response = ai.messages.create(
            model="claude-sonnet-4-6", max_tokens=600,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = response.content[0].text.strip() if response.content else ""
        raw = re.sub(r"^```(?:json)?\s*", "", raw, flags=re.IGNORECASE)
        raw = re.sub(r"\s*```$", "", raw).strip()
        data = json.loads(raw)
        sent_date = latest.get("sentDateTime", "")[:10]
        leads = [
            Lead(name=lead["name"], status=lead.get("status", ""), last_updated=sent_date)
            for lead in data.get("leads", []) if lead.get("name")
        ]
        week_number = int(data.get("week_number") or 0)
        logger.info(
            "  Backfilled ledger from %r (sent %s): week %d, %d known lead(s)",
            latest.get("subject"), sent_date, week_number, len(leads),
        )
        return LedgerEntry(address=address, week_number=week_number, known_leads=leads)
    except Exception as exc:
        logger.warning("  Ledger backfill failed for %s (%s) -- starting blank", address, exc)
        return LedgerEntry(address=address)


# ─── Outlook draft creation ───────────────────────────────────────────────────


def outlook_create_new_draft(
    outlook: OutlookClient,
    to_address: str,
    subject: str,
    html_body: str,
) -> Optional[str]:
    """Create a standalone new draft in Outlook Drafts. Returns draft message ID."""
    payload = {
        "subject": subject,
        "body": {"contentType": "HTML", "content": html_body},
        "toRecipients": [{"emailAddress": {"address": to_address}}],
    }
    result = outlook._post("/me/messages", payload)
    if result and "id" in result:
        draft_id = result["id"]
        logger.info("  Outlook draft saved: id=%s", draft_id)
        return draft_id
    logger.error("  Outlook draft creation failed for %s", subject)
    return None


# ─── Gmail draft creation ─────────────────────────────────────────────────────


def gmail_create_new_draft(
    gmail: GmailClient,
    to_address: str,
    subject: str,
    html_body: str,
    plain_body: str = "",
) -> Optional[str]:
    """Create a standalone new draft in Gmail Drafts. Returns draft ID."""
    try:
        msg = MIMEMultipart("alternative")
        msg["To"] = to_address
        msg["Subject"] = subject
        if plain_body:
            msg.attach(MIMEText(plain_body, "plain"))
        msg.attach(MIMEText(html_body, "html"))

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        result = (
            gmail.service.users()
            .drafts()
            .create(userId="me", body={"message": {"raw": raw}})
            .execute()
        )
        draft_id = result.get("id")
        logger.info("  Gmail draft saved: id=%s", draft_id)
        return draft_id
    except Exception as exc:
        logger.error("  Gmail draft failed: %s", exc)
        return None


# ─── Claude ───────────────────────────────────────────────────────────────────


def claude_draft_vendor_update(
    ai: anthropic.Anthropic,
    campaign: Campaign,
    ledger_entry: LedgerEntry,
    landlord_first_name: Optional[str],
    inbox_snippets: List[str],
    sent_snippets: List[str],
    style_examples: List[str],
) -> Tuple[str, str, LedgerEntry]:
    """Return (subject, html_body, updated_ledger_entry).

    Numbers (enquiries/inspections/offers/days on market) come straight from
    Eagle -- authoritative, never re-derived from email. The narrative --
    who's who and why -- comes from the ledger (continuity from previous
    weeks), folded together with whatever's new in this week's email
    activity. Email can only see portal/inbox correspondence Eddie was
    actually on, never phone calls or other CRM-only activity, so the model
    is told to flag that gap honestly rather than imply the leads list below
    is the complete picture behind Eagle's enquiry count.

    One Claude call returns both the drafted email AND the updated leads
    list together (as JSON), so the ledger update is grounded in the same
    pass that wrote the email, not a separate re-derivation.
    """
    address = campaign.address
    first_name = landlord_first_name or "[LANDLORD FIRST NAME - PLEASE UPDATE]"
    week_number = (
        campaign.days_on_market // 7 + 1
        if campaign.days_on_market is not None
        else ledger_entry.week_number + 1
    )

    known_leads_block = (
        "\n".join(f"  - {lead.name}: {lead.status}" for lead in ledger_entry.known_leads)
        if ledger_entry.known_leads else "  (none known yet)"
    )

    activity_block = ""
    if inbox_snippets:
        activity_block += "Emails RECEIVED about this property since the last update:\n"
        activity_block += "\n".join(f"  - {s}" for s in inbox_snippets[:8]) + "\n\n"
    if sent_snippets:
        activity_block += "Emails Eddie SENT about this property since the last update:\n"
        activity_block += "\n".join(f"  - {s}" for s in sent_snippets[:8]) + "\n\n"
    if not inbox_snippets and not sent_snippets:
        activity_block = "No new email activity found since the last update.\n\n"

    style_text = ""
    if style_examples:
        style_text = (
            "\n\n--- STYLE REFERENCE (Eddie's own real campaign updates -- match this tone, "
            "not a single fixed template: his own format genuinely varies property to property) ---\n"
            + "\n\n---\n".join(style_examples)
            + "\n--- END STYLE REFERENCE ---"
        )

    def _fmt(n: Optional[int]) -> str:
        return str(n) if n is not None else "not recorded in the CRM"

    prompt = (
        "You are drafting a vendor/landlord campaign update email on behalf of Edward Ghattas, "
        "commercial real estate agent at IB Property Sydney.\n\n"
        f"Property: {address}\n"
        f"Listing type: {campaign.sale_or_lease or 'unknown'}\n"
        f"Week {week_number} of the campaign (from days on market, per the CRM -- ground truth)\n"
        "CONFIRMED FROM THE CRM (ground truth -- state these accurately, never contradict or invent "
        "a different figure): "
        f"{_fmt(campaign.num_enquiries)} enquiries to date, "
        f"{_fmt(campaign.num_inspection_attendances)} inspections, "
        f"{_fmt(campaign.num_offers)} offers.\n\n"
        "KNOWN LEADS AS OF LAST UPDATE (carry these forward by name; update a lead's status only if "
        "this week's activity below actually says more about them, e.g. \"as mentioned last week\"):\n"
        f"{known_leads_block}\n\n"
        f"WHAT'S NEW SINCE THE LAST UPDATE:\n{activity_block}"
        f"{style_text}\n\n"
        "--- INSTRUCTIONS ---\n"
        f"- The landlord's first name is: {first_name}\n"
        "- Write a real weekly update in Edward's own voice and structure (see style reference) -- "
        "not a fixed template. Cover: what's happened, what the CRM numbers above mean, and what "
        "he'll do next.\n"
        "- For named leads: reuse the known-leads list where still relevant; add a new named lead "
        "only if this week's activity actually introduces one. Do NOT invent leads, outcomes, or "
        "reasons that aren't in the activity or known-leads list above.\n"
        "- Email can only show correspondence Eddie was actually on -- if the CRM's enquiry count is "
        "higher than the leads you can name, add one honest line noting the gap (e.g. some enquiries "
        "came through calls/the database, not reflected in the list here) rather than implying the "
        "list above is the complete picture.\n"
        "- Match Edward's direct, honest, professional tone -- he does not soften bad news.\n"
        "- Do NOT use a formal signature block -- end with 'Thank you.' only, no name/company/email.\n\n"
        "Return ONLY valid JSON, no markdown fences, no other text:\n"
        '{"html_body": "<p>...</p>...", "leads": [{"name": "...", "status": "..."}]}\n'
        "\"leads\" must be the FULL current list (known leads carried forward, with any status updates, "
        "plus any genuinely new ones this week) -- not just what changed."
    )

    try:
        response = ai.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1400,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = response.content[0].text.strip() if response.content else ""
        raw = re.sub(r"^```(?:json)?\s*", "", raw, flags=re.IGNORECASE)
        raw = re.sub(r"\s*```$", "", raw).strip()
        data = json.loads(raw)
        html_body = (data.get("html_body") or "").strip()
        if not html_body:
            raise ValueError("Empty html_body from Claude")
        today = datetime.now().strftime("%Y-%m-%d")
        leads = [
            Lead(name=lead["name"], status=lead.get("status", ""), last_updated=today)
            for lead in data.get("leads", []) if lead.get("name")
        ]
    except Exception as exc:
        logger.error("Claude draft failed for %s: %s", address, exc)
        html_body = (
            f"<p>Hi {first_name},</p>"
            "<p>Hope you're well.</p>"
            f"<p>I want to give you a brief update on the campaign for {address}. "
            "The campaign is progressing and we continue to market the property "
            "across all major platforms. I will be in touch with further updates as activity develops.</p>"
            "<p>Thank you.</p>"
        )
        leads = ledger_entry.known_leads  # don't lose known state just because this draft call failed

    subject = f"{address} — Week {week_number} Campaign Update"
    updated_entry = LedgerEntry(
        address=address,
        week_number=week_number,
        known_leads=leads,
        last_drafted_at=datetime.now().isoformat(),
    )
    return subject, html_body, updated_entry


# ─── Main ─────────────────────────────────────────────────────────────────────


def main(dry_run: bool = False) -> Dict:
    logger.info(
        "=== Vendor Update Agent starting — %s%s ===",
        datetime.now().strftime("%Y-%m-%d %H:%M"),
        " [DRY RUN -- nothing will be saved]" if dry_run else "",
    )

    missing = [
        k for k in ("ANTHROPIC_API_KEY", "AZURE_CLIENT_ID", "AZURE_TENANT_ID",
                    "EAGLE_CLIENT_ID", "EAGLE_CLIENT_SECRET")
        if not os.getenv(k)
    ]
    if missing:
        logger.error("Missing required env vars: %s", ", ".join(missing))
        sys.exit(1)

    msal_cache = os.getenv("MSAL_TOKEN_CACHE_PATH", "msal_token_cache.bin")
    gmail_creds = os.getenv("GMAIL_CREDENTIALS_PATH", "gmail_credentials.json")
    gmail_token = os.getenv("GMAIL_TOKEN_PATH", "gmail_token.json")
    user_gmail = os.getenv("USER_GMAIL", "edwardenag@gmail.com")
    user_outlook = os.getenv("USER_OUTLOOK", "edward@ibproperty.com.au")

    # Outlook isn't optional here -- landlord lookup, weekly activity and style
    # examples all come from it -- but a launchd run has nobody to complete a
    # device-code/MFA prompt, so exit cleanly with the fix instead of a traceback.
    logger.info("Connecting to Outlook...")
    try:
        outlook = OutlookClient(
            client_id=os.getenv("AZURE_CLIENT_ID"),
            tenant_id=os.getenv("AZURE_TENANT_ID"),
            token_cache_path=msal_cache,
        )
    except OutlookAuthRequired as exc:
        logger.error(
            "Outlook needs interactive re-consent and this unattended run can't "
            "provide it -- nothing to draft this run. Fix by running `python "
            "agent.py --auth` yourself: %s", exc,
        )
        sys.exit(1)

    # Gmail is optional -- an unattended run must not crash just because the
    # cached refresh token died. Same rule as agent.py's build_gmail_client:
    # degrade to Outlook-only drafts rather than raise.
    logger.info("Connecting to Gmail...")
    try:
        gmail = GmailClient(credentials_path=gmail_creds, token_path=gmail_token)
    except GmailAuthRequired as exc:
        logger.error(
            "Gmail needs interactive re-consent and this unattended run can't "
            "provide it -- Gmail drafts will be skipped this run (Outlook "
            "drafts still get created). Fix by running `python agent.py "
            "--auth` yourself: %s", exc,
        )
        gmail = None

    # Eagle (the CRM) is the live source of "what's actually on market" --
    # required, fails closed: log clearly and stop rather than fall back to
    # a doc Eddie would have to hand-maintain (that was the entire problem
    # with the previous design).
    logger.info("Connecting to EagleAgent (CRM, live campaign source)...")
    try:
        eagle = EagleClient(
            os.getenv("EAGLE_CLIENT_ID"),
            os.getenv("EAGLE_CLIENT_SECRET"),
            token_cache_path=os.getenv("EAGLE_TOKEN_CACHE_PATH", "eagle_token_cache.json"),
        )
        campaign_source = EagleCampaignSource(
            eagle, agent_email=user_outlook, agent_name=_AGENT_NAME_FALLBACK,
        )
        campaigns = campaign_source.get_active_campaigns()
    except EagleAuthRequired as exc:
        logger.error(
            "EagleAgent authentication failed -- nothing to draft this run. "
            "Check EAGLE_CLIENT_ID / EAGLE_CLIENT_SECRET in .env: %s", exc,
        )
        sys.exit(1)

    if not campaigns:
        logger.warning(
            "No active campaigns found in Eagle for %s. Nothing to draft this run "
            "(this is a real result, not an error -- check the CRM if it's unexpected).",
            user_outlook,
        )
        return {"drafted": 0, "failed": 0, "campaigns": 0}

    logger.info("Loaded %d active campaign(s) from Eagle for %s", len(campaigns), user_outlook)

    ai = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

    logger.info("Fetching style examples from Sent Items...")
    style_examples = fetch_style_examples(outlook)
    logger.info("Found %d style example(s)", len(style_examples))

    ledger = load_ledger()
    drafted = 0
    failed = 0
    previews: List[Dict] = []

    for campaign in campaigns:
        address = campaign.address
        logger.info("Processing: %s", address)
        try:
            landlord_email, landlord_first_name = resolve_landlord_contact(outlook, campaign)
            if not landlord_email:
                landlord_email = "[LANDLORD EMAIL - PLEASE ADD]"
                logger.warning("  No landlord email found for %s", address)
            if not landlord_first_name:
                logger.warning("  No landlord first name found for %s", address)

            entry = ledger.get(campaign.property_id)
            if entry is None:
                entry = backfill_ledger_entry(outlook, ai, address)
                logger.info(
                    "  New campaign in ledger -- backfilled %d known lead(s), week %d",
                    len(entry.known_leads), entry.week_number,
                )

            inbox_snippets, sent_snippets = gather_weekly_activity(outlook, address, days=7)
            logger.info(
                "  Activity: %d inbox email(s), %d sent email(s)",
                len(inbox_snippets), len(sent_snippets),
            )

            subject, html_body, updated_entry = claude_draft_vendor_update(
                ai, campaign, entry, landlord_first_name, inbox_snippets, sent_snippets, style_examples,
            )
            ledger[campaign.property_id] = updated_entry
            logger.info("  Drafted: %s → To: %s", subject, landlord_email)

            if dry_run:
                previews.append({
                    "address": address, "to": landlord_email, "subject": subject,
                    "num_enquiries": campaign.num_enquiries, "week": updated_entry.week_number,
                    "known_leads": len(updated_entry.known_leads),
                })
                drafted += 1
                continue

            plain_body = re.sub(r"<[^>]+>", " ", html_body)
            plain_body = re.sub(r"\s{2,}", " ", plain_body).strip()

            outlook_id = outlook_create_new_draft(outlook, landlord_email, subject, html_body)

            if gmail is not None:
                gmail_to = landlord_email if "@" in landlord_email and "[" not in landlord_email else user_gmail
                gmail_id = gmail_create_new_draft(gmail, gmail_to, subject, html_body, plain_body)
            else:
                gmail_id = None

            if outlook_id or gmail_id:
                drafted += 1
                logger.info(
                    "  Saved drafts — outlook=%s gmail=%s",
                    outlook_id or "FAILED",
                    gmail_id or "FAILED",
                )
            else:
                failed += 1
                logger.warning("  Both draft saves failed for: %s", address)
        except Exception as exc:
            # One campaign's Outlook hiccup or malformed data must not lose
            # the other 16 -- log and keep going, same isolation pattern as
            # draft_agent.py's per-email processing loop.
            failed += 1
            logger.error("  Failed processing %s -- skipping it, continuing: %s", address, exc)

    if not dry_run:
        save_ledger(ledger)

    logger.info(
        "=== Done%s: %d draft(s) %s, %d failed === Log: %s",
        " (DRY RUN)" if dry_run else "",
        drafted, "previewed" if dry_run else "created", failed, LOG_PATH,
    )
    return {"drafted": drafted, "failed": failed, "campaigns": len(campaigns), "previews": previews}


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="IB Property vendor/landlord weekly update agent")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Preview only -- reads Eagle + Outlook and drafts the update text, saves/sends nothing",
    )
    args = parser.parse_args()
    result = main(dry_run=args.dry_run)
    if args.dry_run:
        print(json.dumps(result, indent=2))
