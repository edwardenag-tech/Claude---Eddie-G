"""Draft reply agent for IB Property inbox.

Reads Outlook inbox (last 48 hrs), classifies each email via Claude,
searches for related property emails, then saves reply drafts to both
Outlook Drafts and Gmail Drafts — nothing is sent automatically.

Usage:
    python draft_agent.py
"""

import os
import sys
import base64
import logging
import re
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Dict, List, Optional, Tuple

from dotenv import load_dotenv
import anthropic

from outlook_client import OutlookClient
from gmail_client import GmailClient

# ─── Bootstrap ───────────────────────────────────────────────────────────────

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

USER_OUTLOOK = os.getenv("USER_OUTLOOK", "edward@ibproperty.com.au").lower()
USER_GMAIL = os.getenv("USER_GMAIL", "edwardenag@gmail.com").lower()
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")

CATEGORIES = ["lease_enquiry", "sale_enquiry", "vendor_update", "landlord_query", "general"]

# ─── Newsletter / Automated detection ────────────────────────────────────────

_FROM_NOISE = re.compile(
    r"no[_\-.]?reply|donotreply|mailer.daemon|postmaster|"
    r"notifications?@|updates?@|alerts?@|marketing@|"
    r"bounce|sendgrid|mailchimp|constantcontact|hubspot|"
    r"salesforce|marketo|campaign\.monitor",
    re.IGNORECASE,
)

_SUBJECT_NOISE = re.compile(
    r"unsubscribe|newsletter|weekly digest|monthly (update|report)|"
    r"promotional|special offer|deal of the",
    re.IGNORECASE,
)


def is_automated(email: Dict) -> bool:
    """True if the email looks like a newsletter or system notification."""
    return bool(
        _FROM_NOISE.search(email.get("from", ""))
        or _SUBJECT_NOISE.search(email.get("subject", ""))
    )


def is_self_sent(email: Dict) -> bool:
    """True if the sender is one of the user's own addresses."""
    addr = email.get("from", "").lower()
    return USER_OUTLOOK in addr or USER_GMAIL in addr


# ─── Outlook helpers ──────────────────────────────────────────────────────────

def outlook_already_replied(outlook: OutlookClient, conversation_id: str) -> bool:
    """True if the Sent Items folder contains any message in this conversation."""
    if not conversation_id:
        return False
    safe_id = conversation_id.replace("'", "''")
    params = {
        "$filter": f"conversationId eq '{safe_id}'",
        "$top": 1,
        "$select": "id",
    }
    result = outlook._get("/me/mailFolders/sentitems/messages", params=params)
    return bool(result and result.get("value"))


def outlook_search_related(
    outlook: OutlookClient, hint: str, exclude_id: str
) -> List[Dict]:
    """Return inbox messages whose subject/body contain the property hint."""
    params = {
        "$search": f'"{hint}"',
        "$top": 20,
        "$select": "id,subject,from,receivedDateTime,hasAttachments",
    }
    result = outlook._get("/me/messages", params=params)
    messages = result.get("value", []) if result else []
    return [m for m in messages if m.get("id") != exclude_id]


def outlook_attachment_names(outlook: OutlookClient, msg_id: str) -> List[str]:
    """Return filenames of attachments on an Outlook message."""
    result = outlook._get(f"/me/messages/{msg_id}/attachments", params={"$select": "name"})
    if not result:
        return []
    return [a["name"] for a in result.get("value", []) if a.get("name")]


def outlook_create_draft(
    outlook: OutlookClient,
    email_id: str,
    html_body: str,
) -> Optional[str]:
    """Create a threaded reply draft in Outlook. Returns the draft message ID."""
    # Step 1: create a reply shell (sets To, subject Re:..., thread references)
    shell = outlook._post(f"/me/messages/{email_id}/createReply", {"comment": ""})
    if not shell or "id" not in shell:
        logger.error("  Outlook createReply failed for email_id=%s", email_id)
        return None

    draft_id = shell["id"]

    # Step 2: patch in Claude-generated body
    ok = outlook._patch(
        f"/me/messages/{draft_id}",
        {"body": {"contentType": "HTML", "content": html_body}},
    )
    if not ok:
        logger.error("  Outlook PATCH draft body failed for draft_id=%s", draft_id)
        return None

    logger.info("  Outlook reply draft saved: id=%s", draft_id)
    return draft_id


# ─── Gmail helpers ────────────────────────────────────────────────────────────

def gmail_create_draft(
    gmail: GmailClient,
    to_address: str,
    subject: str,
    html_body: str,
    plain_body: str = "",
    in_reply_to: str = "",
) -> Optional[str]:
    """Create a reply draft in Gmail Drafts. Returns the draft ID."""
    try:
        msg = MIMEMultipart("alternative")
        msg["To"] = to_address
        msg["Subject"] = subject
        if in_reply_to:
            msg["In-Reply-To"] = in_reply_to
            msg["References"] = in_reply_to
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


# ─── Property hint extraction ─────────────────────────────────────────────────

def extract_property_hint(subject: str) -> Optional[str]:
    """Strip Re:/Fwd: prefixes and return the core subject as a search hint."""
    cleaned = re.sub(r"^(re|fwd?):\s*", "", subject.strip(), flags=re.IGNORECASE).strip()
    return cleaned[:60] if len(cleaned) > 5 else None


# ─── Claude helpers ───────────────────────────────────────────────────────────

def claude_classify(ai: anthropic.Anthropic, email: Dict) -> str:
    """Return one of the five CRE category strings."""
    prompt = (
        "You are an assistant for IB Property Sydney, a commercial real estate agency.\n\n"
        "Classify the following email into EXACTLY ONE category:\n"
        "  lease_enquiry   — enquiry about leasing a property\n"
        "  sale_enquiry    — enquiry about buying or selling a property\n"
        "  vendor_update   — update from a vendor, supplier, or tradesperson\n"
        "  landlord_query  — query or request from a landlord or property owner\n"
        "  general         — anything else\n\n"
        f"From: {email.get('from_name', '')} <{email.get('from', '')}>\n"
        f"Subject: {email.get('subject', '')}\n"
        f"Body: {email.get('body', '')[:600]}\n\n"
        "Reply with only the category name, nothing else."
    )
    try:
        response = ai.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=20,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = response.content[0].text.strip().lower()
        return raw if raw in CATEGORIES else "general"
    except Exception as exc:
        logger.error("Claude classify failed: %s", exc)
        return "general"


def claude_draft_reply(
    ai: anthropic.Anthropic,
    email: Dict,
    category: str,
    related_attachments: List[str],
) -> Tuple[str, str]:
    """Return (reply_subject, html_body) for a professional CRE reply."""
    category_context = {
        "lease_enquiry": (
            "The sender is enquiring about leasing. Acknowledge their interest, "
            "provide helpful information about availability, and suggest an inspection time."
        ),
        "sale_enquiry": (
            "The sender is enquiring about buying or selling. Acknowledge their interest "
            "and offer to discuss further, suggesting a call or meeting."
        ),
        "vendor_update": (
            "This is from a vendor or supplier. Acknowledge receipt professionally "
            "and confirm any required next steps."
        ),
        "landlord_query": (
            "This is from a landlord or property owner. Address their query "
            "professionally and provide clear next steps or reassurance."
        ),
        "general": "This is a general enquiry. Reply professionally and helpfully.",
    }

    attachments_note = ""
    if related_attachments:
        names = ", ".join(related_attachments[:10])
        attachments_note = f"\n\nRelated documents found in the thread: {names}"

    prompt = (
        "You are drafting a professional reply on behalf of Edward Ghattas, "
        "commercial real estate agent at IB Property Sydney.\n\n"
        "--- Original email ---\n"
        f"From: {email.get('from_name', 'the sender')} <{email.get('from', '')}>\n"
        f"Subject: {email.get('subject', '')}\n"
        f"Body:\n{email.get('body', '')[:1500]}"
        f"{attachments_note}\n\n"
        "--- Instructions ---\n"
        f"Context: {category_context.get(category, category_context['general'])}\n"
        "- Greet by first name where possible\n"
        "- Keep the reply under 180 words\n"
        "- Use a professional, warm tone appropriate for commercial real estate\n"
        "- Sign off as:\n"
        "  Edward Ghattas\n"
        "  IB Property Sydney\n"
        "  edward@ibproperty.com.au\n\n"
        "Return ONLY the HTML body content using <p> and <br> tags. "
        "Do not include a subject line inside the body."
    )

    try:
        response = ai.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=700,
            messages=[{"role": "user", "content": prompt}],
        )
        html_body = response.content[0].text.strip()
    except Exception as exc:
        logger.error("Claude draft failed: %s", exc)
        html_body = (
            "<p>Thank you for your email. I will review this and get back to you shortly.</p>"
            "<p>Kind regards,<br>Edward Ghattas<br>IB Property Sydney<br>"
            "edward@ibproperty.com.au</p>"
        )

    original_subject = email.get("subject", "")
    reply_subject = (
        original_subject
        if original_subject.lower().startswith("re:")
        else f"Re: {original_subject}"
    )
    return reply_subject, html_body


# ─── Main ─────────────────────────────────────────────────────────────────────

def main(max_emails: Optional[int] = None) -> None:
    logger.info("=== Draft Agent starting — %s ===", datetime.now().strftime("%Y-%m-%d %H:%M"))

    # Validate required env vars
    missing = [k for k in ("ANTHROPIC_API_KEY", "AZURE_CLIENT_ID", "AZURE_TENANT_ID") if not os.getenv(k)]
    if missing:
        logger.error("Missing required env vars: %s", ", ".join(missing))
        sys.exit(1)

    msal_cache = os.getenv("MSAL_TOKEN_CACHE_PATH", "msal_token_cache.bin")
    gmail_creds = os.getenv("GMAIL_CREDENTIALS_PATH", "gmail_credentials.json")
    gmail_token = os.getenv("GMAIL_TOKEN_PATH", "gmail_token.json")

    # Initialise clients
    logger.info("Connecting to Outlook (Microsoft Graph)...")
    outlook = OutlookClient(
        client_id=os.getenv("AZURE_CLIENT_ID"),
        tenant_id=os.getenv("AZURE_TENANT_ID"),
        token_cache_path=msal_cache,
    )

    logger.info("Connecting to Gmail...")
    gmail = GmailClient(credentials_path=gmail_creds, token_path=gmail_token)

    ai = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # Fetch last 48 hours from Inbox + "Front of Mind"
    logger.info("Fetching Outlook inbox + 'Front of Mind' (last 48 hours)...")
    raw_messages = outlook.get_recent_emails(since_days=2, extra_folders=["Front of Mind"])
    logger.info("Retrieved %d messages", len(raw_messages))

    drafted = 0
    skipped = 0

    if max_emails is not None:
        raw_messages = raw_messages[:max_emails]

    for raw in raw_messages:
        email = OutlookClient.extract_email_data(raw)
        subject = email.get("subject", "(no subject)")
        from_addr = email.get("from", "")
        msg_id = email.get("id", "")
        conversation_id = raw.get("conversationId", "")

        # ── Skip conditions ──────────────────────────────────────────────────
        if is_self_sent(email):
            logger.debug("Skip (self-sent): %s", subject)
            skipped += 1
            continue

        if is_automated(email):
            logger.info("Skip (automated): %s | from=%s", subject, from_addr)
            skipped += 1
            continue

        if outlook_already_replied(outlook, conversation_id):
            logger.info("Skip (already replied): %s", subject)
            skipped += 1
            continue

        # ── Process ──────────────────────────────────────────────────────────
        logger.info("Processing: %s | from=%s", subject, from_addr)

        # 1. Classify
        category = claude_classify(ai, email)
        logger.info("  Category → %s", category)

        # 2. Find related emails by property hint; collect attachment names
        related_attachments: List[str] = []
        hint = extract_property_hint(subject)
        if hint:
            related = outlook_search_related(outlook, hint, exclude_id=msg_id)
            logger.info("  Related emails found: %d (hint=%r)", len(related), hint)
            for rel in related[:5]:
                if rel.get("hasAttachments"):
                    names = outlook_attachment_names(outlook, rel["id"])
                    related_attachments.extend(names)
            if related_attachments:
                logger.info("  Related attachments: %s", related_attachments)

        # 3. Draft reply via Claude
        reply_subject, html_body = claude_draft_reply(ai, email, category, related_attachments)

        # 4. Save to Outlook Drafts (threaded reply via createReply + PATCH)
        outlook_id = outlook_create_draft(
            outlook,
            email_id=msg_id,
            html_body=html_body,
        )

        # 5. Save to Gmail Drafts (reply headers from Outlook's internetMessageId)
        internet_message_id = raw.get("internetMessageId", "")
        plain_body = re.sub(r"<[^>]+>", " ", html_body)
        plain_body = re.sub(r"\s{2,}", " ", plain_body).strip()
        gmail_id = gmail_create_draft(
            gmail,
            to_address=from_addr,
            subject=reply_subject,
            html_body=html_body,
            plain_body=plain_body,
            in_reply_to=internet_message_id,
        )

        # 6. Log outcome
        if outlook_id or gmail_id:
            drafted += 1
            logger.info(
                "  Drafted '%s' → outlook=%s gmail=%s",
                reply_subject,
                outlook_id or "FAILED",
                gmail_id or "FAILED",
            )
        else:
            logger.warning("  Both draft saves failed for: %s", subject)

    logger.info(
        "=== Done: %d draft(s) created, %d email(s) skipped ===",
        drafted,
        skipped,
    )


if __name__ == "__main__":
    main()
