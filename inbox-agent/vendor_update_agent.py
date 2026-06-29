"""Vendor/landlord weekly campaign update agent for IB Property.

Drafts Monday-morning campaign update emails for each active listing in
listings_db.json, saves them to both Outlook Drafts and Gmail Drafts.
Never auto-sends.

Usage:
    python vendor_update_agent.py
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

from gmail_client import GmailClient
from outlook_client import OutlookClient

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

_LISTINGS_DB_PATH = os.path.join(os.path.dirname(__file__), "listings_db.json")

# Listings whose notes contain these markers are skipped (not active Eddie campaigns)
_SKIP_NOTES = re.compile(
    r"property has been sold|competing listing|marketed by sutton anderson",
    re.IGNORECASE,
)

# ─── Listing helpers ──────────────────────────────────────────────────────────


def load_active_listings() -> List[Dict]:
    """Load listings_db.json and return unique, active listings only."""
    try:
        with open(_LISTINGS_DB_PATH) as fh:
            data = json.load(fh)
    except (FileNotFoundError, json.JSONDecodeError) as exc:
        logger.error("Could not load listings_db.json: %s", exc)
        return []

    seen_addresses: set = set()
    active = []
    for listing in data.get("listings", []):
        notes = listing.get("notes", "")
        if _SKIP_NOTES.search(notes):
            logger.info("Skipping listing (sold/competing): %s", listing.get("address"))
            continue
        addr = listing.get("address", "").strip().lower()
        if not addr or addr in seen_addresses:
            continue
        seen_addresses.add(addr)
        active.append(listing)

    logger.info("Loaded %d active unique listings", len(active))
    return active


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
        if re.search(r"campaign update", subj, re.IGNORECASE):
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
        if re.search(r"campaign update", subj, re.IGNORECASE):
            plain = _body_to_plain(msg)
            m = re.match(r"Hi\s+(\w+)", plain.strip())
            if m:
                return m.group(1)
    return None


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
        if re.search(r"campaign update", subj, re.IGNORECASE):
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
        if not re.search(r"campaign update", subj, re.IGNORECASE):
            continue
        plain = _body_to_plain(msg)
        # Trim to just the outgoing portion (before any quoted reply)
        plain = plain[:1500]
        examples.append(f"Subject: {subj}\n{plain}")
        if len(examples) >= 3:
            break
    return examples


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

_FORMAT_GUIDE = """Eddie's campaign update email format (reproduce exactly):

Hi [FIRST NAME],

Hope you're well.

I want to give you a [thorough/detailed] update on [address or campaign name, brief context].

On the activity side, [PARAGRAPH: specific things that happened this week — enquiries received with names/dates, inspections conducted, EDMs sent, offers, feedback from parties, agreements signed, any issues. Be specific and concrete. If there was no activity this week, describe the ongoing marketing efforts across platforms and any pipeline being worked.]

The market [response/feedback] [PARAGRAPH: interpret what the activity means — is demand strong or soft? What are buyers/tenants saying? What does the level of interest tell us about the property's positioning? 2–4 sentences.]

My recommendation is [PARAGRAPH: specific next steps Eddie will take — follow up with named parties, arrange inspections, review pricing, push a particular channel. Direct, decisive, professional. 2–4 sentences.]

[Optional: list marketing platform links if relevant]

Thank you."""


def claude_draft_vendor_update(
    ai: anthropic.Anthropic,
    listing: Dict,
    landlord_first_name: Optional[str],
    inbox_snippets: List[str],
    sent_snippets: List[str],
    style_examples: List[str],
) -> Tuple[str, str]:
    """Return (subject, html_body) for the vendor update email."""

    address = listing.get("address", "[ADDRESS]")
    listing_type = listing.get("type", "sale").lower()
    notes = listing.get("notes", "")

    first_name = landlord_first_name or "[LANDLORD FIRST NAME - PLEASE UPDATE]"

    activity_block = ""
    if inbox_snippets:
        activity_block += "Emails RECEIVED about this property this week:\n"
        activity_block += "\n".join(f"  - {s}" for s in inbox_snippets[:8])
        activity_block += "\n\n"
    if sent_snippets:
        activity_block += "Emails Eddie SENT about this property this week:\n"
        activity_block += "\n".join(f"  - {s}" for s in sent_snippets[:8])
        activity_block += "\n\n"
    if not inbox_snippets and not sent_snippets:
        activity_block = (
            "No email activity found for this property in the last 7 days. "
            "Write a brief ongoing-campaign update referencing the campaign notes below."
        )

    style_text = ""
    if style_examples:
        style_text = (
            "\n\n--- STYLE REFERENCE (Eddie's real campaign updates — match this tone and structure) ---\n"
            + "\n\n---\n".join(style_examples)
            + "\n--- END STYLE REFERENCE ---"
        )

    prompt = (
        "You are drafting a vendor/landlord campaign update email on behalf of Edward Ghattas, "
        "commercial real estate agent at IB Property Sydney.\n\n"
        f"Property: {address}\n"
        f"Type: {listing_type}\n"
        f"Campaign notes (authoritative background):\n{notes}\n\n"
        f"Weekly email activity:\n{activity_block}"
        f"{style_text}\n\n"
        "--- FORMAT TO FOLLOW (reproduce this structure exactly) ---\n"
        f"{_FORMAT_GUIDE}\n\n"
        "--- INSTRUCTIONS ---\n"
        f"- The landlord's first name is: {first_name}\n"
        "- Write three substantive paragraphs: Activity | Market reading | Recommendation\n"
        "- Use specific names, dates, and details from the campaign notes and email activity\n"
        "- If no activity this week, describe ongoing platform marketing and pipeline work\n"
        "- Do NOT use a formal signature block — end with 'Thank you.' only\n"
        "- Do NOT include 'Edward Ghattas', 'IB Property Sydney', or any sign-off name\n"
        "- Match Eddie's direct, honest, professional tone — he does not soften bad news\n"
        "- Return ONLY the HTML body using <p> and <br> tags (no subject line in body)\n"
        "- Do NOT wrap in markdown code fences"
    )

    try:
        response = ai.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1000,
            messages=[{"role": "user", "content": prompt}],
        )
        html_body = response.content[0].text.strip() if response.content else ""
        # Strip any accidental markdown fences
        html_body = re.sub(r"^```(?:html)?\s*", "", html_body, flags=re.IGNORECASE)
        html_body = re.sub(r"\s*```$", "", html_body)
        html_body = html_body.strip()
        if not html_body:
            raise ValueError("Empty response from Claude")
    except Exception as exc:
        logger.error("Claude draft failed for %s: %s", address, exc)
        html_body = (
            f"<p>Hi {first_name},</p>"
            "<p>Hope you're well.</p>"
            "<p>I want to give you a brief update on the campaign. "
            "The campaign is progressing and we continue to market the property "
            "across all major platforms including RealCommercial, Commercial Real Estate, "
            "Instagram, LinkedIn, and the IB Property website. "
            "I will be in touch with further updates as activity develops.</p>"
            "<p>Thank you.</p>"
        )

    subject = f"Campaign Update - {address}"
    return subject, html_body


# ─── Main ─────────────────────────────────────────────────────────────────────


def main() -> None:
    logger.info("=== Vendor Update Agent starting — %s ===", datetime.now().strftime("%Y-%m-%d %H:%M"))

    missing = [
        k for k in ("ANTHROPIC_API_KEY", "AZURE_CLIENT_ID", "AZURE_TENANT_ID")
        if not os.getenv(k)
    ]
    if missing:
        logger.error("Missing required env vars: %s", ", ".join(missing))
        sys.exit(1)

    msal_cache = os.getenv("MSAL_TOKEN_CACHE_PATH", "msal_token_cache.bin")
    gmail_creds = os.getenv("GMAIL_CREDENTIALS_PATH", "gmail_credentials.json")
    gmail_token = os.getenv("GMAIL_TOKEN_PATH", "gmail_token.json")
    user_gmail = os.getenv("USER_GMAIL", "edwardenag@gmail.com")

    logger.info("Connecting to Outlook...")
    outlook = OutlookClient(
        client_id=os.getenv("AZURE_CLIENT_ID"),
        tenant_id=os.getenv("AZURE_TENANT_ID"),
        token_cache_path=msal_cache,
    )

    logger.info("Connecting to Gmail...")
    gmail = GmailClient(credentials_path=gmail_creds, token_path=gmail_token)

    ai = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

    listings = load_active_listings()
    if not listings:
        logger.error("No active listings found. Exiting.")
        sys.exit(1)

    logger.info("Fetching style examples from Sent Items...")
    style_examples = fetch_style_examples(outlook)
    logger.info("Found %d style example(s)", len(style_examples))

    drafted = 0
    failed = 0

    for listing in listings:
        address = listing.get("address", "")
        logger.info("Processing: %s", address)

        # 1. Find landlord contact
        landlord_email = find_landlord_email(outlook, address)
        landlord_first_name = find_landlord_first_name(outlook, address)
        if not landlord_email:
            landlord_email = "[LANDLORD EMAIL - PLEASE ADD]"
            logger.warning("  No landlord email found for %s", address)
        if not landlord_first_name:
            logger.warning("  No landlord first name found for %s", address)

        # 2. Gather weekly activity
        inbox_snippets, sent_snippets = gather_weekly_activity(outlook, address, days=7)
        logger.info(
            "  Activity: %d inbox email(s), %d sent email(s)",
            len(inbox_snippets), len(sent_snippets),
        )

        # 3. Draft with Claude
        subject, html_body = claude_draft_vendor_update(
            ai, listing, landlord_first_name,
            inbox_snippets, sent_snippets, style_examples,
        )
        logger.info("  Drafted: %s → To: %s", subject, landlord_email)

        plain_body = re.sub(r"<[^>]+>", " ", html_body)
        plain_body = re.sub(r"\s{2,}", " ", plain_body).strip()

        # 4. Save to Outlook Drafts
        outlook_id = outlook_create_new_draft(outlook, landlord_email, subject, html_body)

        # 5. Save to Gmail Drafts — if no real email found, park draft in Eddie's own inbox
        gmail_to = landlord_email if "@" in landlord_email and "[" not in landlord_email else user_gmail
        gmail_id = gmail_create_new_draft(gmail, gmail_to, subject, html_body, plain_body)

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

    logger.info(
        "=== Done: %d draft(s) created, %d failed === Log: %s",
        drafted, failed, LOG_PATH,
    )


if __name__ == "__main__":
    main()
