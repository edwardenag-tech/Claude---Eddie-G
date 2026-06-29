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
    r"notifications?@|updates?@|alerts?@|marketing@|newsletter@|promo@|"
    r"bounce|sendgrid|mailchimp|constantcontact|hubspot|"
    r"salesforce|marketo|campaign\.monitor|grammarly|"
    r"designline",
    re.IGNORECASE,
)

_SUBJECT_NOISE = re.compile(
    r"unsubscribe|newsletter|weekly digest|monthly (update|report)|"
    r"promotional|special offer|deal of the|"
    r"\d+%\s*off|discount|sale ends|promotion",
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


_LISTING_MARKERS = [
    "Property Highlights",
    "Asking Rent",
    "Floor Area",
    "IB Property is pleased",
]


def fetch_sent_enquiry_examples(outlook: OutlookClient) -> List[str]:
    """Return up to 5 plain-text excerpts of Eddie's listing reply emails from Sent Items."""
    params = {
        "$top": 50,
        "$orderby": "sentDateTime desc",
        "$select": "subject,body,sentDateTime",
    }
    result = outlook._get("/me/mailFolders/sentitems/messages", params=params)
    messages = result.get("value", []) if result else []

    examples = []
    for msg in messages:
        body_obj = msg.get("body", {})
        raw_body = body_obj.get("content", "") if isinstance(body_obj, dict) else ""
        plain = re.sub(r"<[^>]+>", " ", raw_body)
        plain = re.sub(r"\s{2,}", " ", plain).strip()

        if any(marker in raw_body or marker in plain for marker in _LISTING_MARKERS):
            subj = msg.get("subject", "")
            examples.append(f"Subject: {subj}\n{plain[:800]}")
            if len(examples) >= 5:
                break

    return examples


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


# ─── Listing detail extraction ────────────────────────────────────────────────

_STREET_TYPES = (
    r'Street|St|Road|Rd|Avenue|Ave|Drive|Dr|Lane|Ln|Place|Pl|Way|'
    r'Highway|Hwy|Crescent|Cres|Boulevard|Blvd|Parade|Pde|Court|Ct|'
    r'Close|Circuit|Cct|Terrace|Tce'
)


def _extract_address(subject: str, body: str) -> Optional[str]:
    """Pull property address out of subject or body text."""
    combined = f"{subject}\n{body}"
    # realcommercial.com.au: "Property ID: 12345, ADDRESS, Contacted"
    m = re.search(r'Property ID:\s*\d+,\s*(.+?),\s*Contacted', combined, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    # commercialrealestate.com.au: "New Enquiry - ADDRESS Contacted: NAME"
    m = re.search(r'New\s+Enquiry\s*[-–]\s*(.+?)\s+Contacted:', combined, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    # Generic: "for/at/about NUMBER ... STREET_TYPE"
    m = re.search(
        rf'(?:for|at|about|regarding)\s+(\d+[^,\n]+?(?:{_STREET_TYPES})[^,\n]*(?:,\s*[A-Z][A-Za-z\s]+)?)',
        combined, re.IGNORECASE,
    )
    if m:
        return m.group(1).strip()
    # Subject line contains a street address
    subj_clean = re.sub(r'^(re|fwd?|enquiry):\s*', '', subject, flags=re.IGNORECASE).strip()
    if re.search(rf'\d+.*(?:{_STREET_TYPES})', subj_clean, re.IGNORECASE):
        return subj_clean[:120]
    return None


def _parse_property_fields(text: str, category: str) -> Dict[str, Optional[str]]:
    """Run regex patterns on text to extract commercial property data fields."""
    found: Dict[str, Optional[str]] = {}
    _DOLLAR = r'\$([\d,]+(?:\.\d+)?(?:k|K|m|M)?)'
    _AREA_VAL = r'([\d,]+(?:\.\d+)?)'
    _AREA_UNIT = r'\s*(?:sqm|m²|m2|square\s+metres?)'

    # ── Rent (lease) ───────────────────────────────────────────────────────────
    if category == "lease_enquiry":
        # Labelled: "Asking Rent: $70,000" or "Rent: $70,000"
        m = re.search(
            r'(?:asking\s+rent|rental|rent(?:al)?)\s*[:\-]?\s*' + _DOLLAR,
            text, re.IGNORECASE,
        )
        if not m:
            # Amount + period suffix: "$70,000 p.a." / "$70,000 pa" / "per annum"
            m = re.search(_DOLLAR + r'\s*(?:p\.?a\.?|per\s+annum|pa\b)', text, re.IGNORECASE)
        if not m:
            # Amount + gross/net qualifier: "$70,000 gross"
            m = re.search(_DOLLAR + r'[^.\n]{0,40}?(?:gross|net)', text, re.IGNORECASE)
        if m:
            raw_val = m.group(1)
            if re.search(r'p\.?c\.?m\.?|per\s+month', m.group(0), re.IGNORECASE):
                try:
                    monthly = float(raw_val.replace(",", "").rstrip("kKmM"))
                    if raw_val.lower().endswith("k"):
                        monthly *= 1000
                    found["asking_rent"] = f"${int(monthly * 12):,} p.a. gross + GST"
                except ValueError:
                    found["asking_rent"] = f"${raw_val} p.a. gross + GST"
            else:
                found["asking_rent"] = f"${raw_val} p.a. gross + GST"

    # ── Sale price ─────────────────────────────────────────────────────────────
    if category == "sale_enquiry":
        m = re.search(
            r'(?:asking\s+price|sale\s+price|price)\s*[:\-]?\s*' + _DOLLAR,
            text, re.IGNORECASE,
        )
        if not m:
            m = re.search(r'offers?\s+(?:over|above|from)\s+' + _DOLLAR, text, re.IGNORECASE)
        if m:
            found["asking_price"] = f"${m.group(1)}"
        elif re.search(r'\bEOI\b|expressions?\s+of\s+interest', text, re.IGNORECASE):
            found["asking_price"] = "EOI — Expressions of Interest"

    # ── Internal floor area ────────────────────────────────────────────────────
    m = re.search(
        r'(?:nla|gfa|internal\s+(?:floor\s+)?area|floor\s+area|'
        r'lettable\s+area|net\s+lettable|gross\s+floor|area|size)\s*[:\-]?\s*'
        + _AREA_VAL + _AREA_UNIT,
        text, re.IGNORECASE,
    )
    if m:
        found["internal_area"] = f"{m.group(1)} sqm*"

    # ── External area ──────────────────────────────────────────────────────────
    m = re.search(
        r'(?:external|outdoor|alfresco|terrace|balcony)\s*(?:area|space)?\s*[:\-]?\s*'
        + _AREA_VAL + _AREA_UNIT,
        text, re.IGNORECASE,
    )
    if m:
        found["external_area"] = f"{m.group(1)} sqm*"

    # Fallback: any sqm values in document order → first = internal, second = external
    if not found.get("internal_area"):
        all_areas = re.findall(_AREA_VAL + _AREA_UNIT, text, re.IGNORECASE)
        if all_areas:
            found["internal_area"] = f"{all_areas[0]} sqm*"
            if len(all_areas) > 1 and not found.get("external_area"):
                found["external_area"] = f"{all_areas[1]} sqm*"

    # ── Building name ──────────────────────────────────────────────────────────
    for pat in [
        r'\b(Sirius[^,.\n]*)',
        r'award[- ]winning\s+([A-Z][A-Za-z0-9\s]+?)(?:\s+development|\s+building|\s+complex|[,.\n])',
        r'(?:located\s+(?:in|within)|nestled\s+(?:in|within)|part\s+of)\s+the\s+([A-Z][A-Za-z0-9\s]+?)(?:\s+development|\s+building|\s+complex|[,.\n])',
        r'\bthe\s+([A-Z][A-Za-z0-9\s]+(?:Centre|Center|Tower|Building|Plaza|House|Court|Arcade|Mall))',
    ]:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            found["building_name"] = m.group(1).strip()
            break

    # ── Outgoings ──────────────────────────────────────────────────────────────
    m = re.search(r'outgoings?\s*[:\-]?\s*' + _DOLLAR + r'[^.\n]{0,20}', text, re.IGNORECASE)
    if m:
        found["outgoings"] = f"${m.group(1)} p.a."

    # ── Car spaces ─────────────────────────────────────────────────────────────
    m = re.search(
        r'(\d+)\s*(?:car\s+(?:space|park|bay)s?|parking\s+(?:space|bay)s?)',
        text, re.IGNORECASE,
    )
    if m:
        n = int(m.group(1))
        found["car_spaces"] = f"{n} car space{'s' if n > 1 else ''}"

    # ── Zoning ─────────────────────────────────────────────────────────────────
    m = re.search(
        r'(?:zone[d]?|zoning)\s*[:\-]?\s*([A-Z][A-Za-z0-9\s/]+?)(?:\s*\(|\s*$|[,.\n])',
        text, re.IGNORECASE,
    )
    if m:
        found["zoning"] = m.group(1).strip()

    return found


def _merge_into(base: Dict, update: Dict) -> None:
    """Fill None values in base from update without overwriting existing values."""
    for k, v in update.items():
        if v and not base.get(k):
            base[k] = v


def _shorten_address(address: Optional[str]) -> Optional[str]:
    """Extract a short search-friendly address string (number + street name + type)."""
    if not address:
        return None
    m = re.search(rf'(\d+\s+\w+\s+(?:{_STREET_TYPES}))', address, re.IGNORECASE)
    if m:
        return m.group(1)
    return address[:40]


def _search_folder_body(
    outlook: OutlookClient, endpoint: str, query: str, top: int = 5
) -> List[Dict]:
    """Run a $search query on a Graph messages endpoint and return messages with body."""
    params = {
        "$search": f'"{query}"',
        "$top": top,
        "$select": "id,subject,body,sentDateTime,receivedDateTime",
    }
    result = outlook._get(endpoint, params=params)
    return result.get("value", []) if result else []


def _body_to_plain(msg: Dict) -> str:
    """Extract and flatten HTML body from a Graph message dict."""
    raw = msg.get("body", {}).get("content", "") if isinstance(msg.get("body"), dict) else ""
    plain = re.sub(r"<[^>]+>", " ", raw)
    return re.sub(r"\s{2,}", " ", plain).strip()


def collect_property_data(
    outlook: OutlookClient,
    raw_message: Dict,
    email: Dict,
    category: str,
) -> Dict[str, Optional[str]]:
    """
    Gather property data from three sources in priority order:
      A) The enquiry email — raw HTML body (richer than stripped plain text)
      B) Sent Items search by address (Eddie's own prior emails about the property)
      C) Full inbox search by address (listing info sheets, brochures, agent emails)
    """
    details: Dict[str, Optional[str]] = {
        "address": None,
        "asking_rent": None,
        "asking_price": None,
        "internal_area": None,
        "external_area": None,
        "building_name": None,
        "outgoings": None,
        "car_spaces": None,
        "zoning": None,
    }

    subject = email.get("subject", "")

    # ── Address first — needed for subsequent searches ─────────────────────────
    details["address"] = _extract_address(subject, email.get("body", ""))

    # ── Source A: enquiry email, raw HTML ─────────────────────────────────────
    raw_html = (
        raw_message.get("body", {}).get("content", "")
        if isinstance(raw_message.get("body"), dict)
        else ""
    )
    html_plain = re.sub(r"\s{2,}", " ", re.sub(r"<[^>]+>", " ", raw_html)).strip()
    _merge_into(details, _parse_property_fields(f"{subject}\n{html_plain}", category))
    logger.debug("  [A] After email body: %s", {k: v for k, v in details.items() if v})

    # ── Source B: Sent Items search by street address ──────────────────────────
    short_addr = _shorten_address(details.get("address"))
    if short_addr:
        sent_msgs = _search_folder_body(
            outlook, "/me/mailFolders/sentitems/messages", short_addr, top=5
        )
        logger.info("  [B] Sent Items hits for %r: %d", short_addr, len(sent_msgs))
        for msg in sent_msgs:
            _merge_into(details, _parse_property_fields(
                f"{msg.get('subject', '')}\n{_body_to_plain(msg)}", category
            ))
        logger.debug("  [B] After Sent Items: %s", {k: v for k, v in details.items() if v})

    # ── Source C: full inbox search by street address ─────────────────────────
    if short_addr:
        inbox_msgs = _search_folder_body(outlook, "/me/messages", short_addr, top=5)
        logger.info("  [C] Inbox hits for %r: %d", short_addr, len(inbox_msgs))
        for msg in inbox_msgs:
            if msg.get("id") == email.get("id"):
                continue
            _merge_into(details, _parse_property_fields(
                f"{msg.get('subject', '')}\n{_body_to_plain(msg)}", category
            ))
        logger.debug("  [C] After inbox search: %s", {k: v for k, v in details.items() if v})

    # ── Fallback placeholders for still-missing key fields ─────────────────────
    if not details.get("address"):
        details["address"] = "[PLEASE ADD - PROPERTY ADDRESS]"
    if category == "lease_enquiry" and not details.get("asking_rent"):
        details["asking_rent"] = "[PLEASE ADD - RENT/PRICE]"
    if category == "sale_enquiry" and not details.get("asking_price"):
        details["asking_price"] = "[PLEASE ADD - SALE PRICE]"
    if not details.get("internal_area"):
        details["internal_area"] = "[PLEASE ADD - SIZE]"

    return details


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
    style_examples: Optional[List[str]] = None,
    listing_details: Optional[Dict] = None,
) -> Tuple[str, str]:
    """Return (reply_subject, html_body) for a professional CRE reply."""
    attachments_note = ""
    if related_attachments:
        names = ", ".join(related_attachments[:10])
        attachments_note = f"\n\nRelated documents found in the thread: {names}"

    is_enquiry = category in ("lease_enquiry", "sale_enquiry")

    if is_enquiry:
        ld = listing_details or {}
        address = ld.get("address") or "[PLEASE ADD - PROPERTY ADDRESS]"
        building_name = ld.get("building_name")

        body_lower = email.get("body", "").lower()
        prop_type = (
            "retail" if "retail" in body_lower
            else "office" if "office" in body_lower
            else "industrial" if "industrial" in body_lower
            else "commercial"
        )
        action_type = "for lease" if category == "lease_enquiry" else "for sale"

        # Build dynamic highlights block (lease vs sale)
        highlights: List[str] = []
        if category == "lease_enquiry":
            highlights.append(
                f"• Asking Rent: {ld.get('asking_rent') or '[PLEASE ADD - RENT/PRICE]'}"
            )
            highlights.append(
                f"• Internal Floor Area: {ld.get('internal_area') or '[PLEASE ADD - SIZE]'}"
            )
            if ld.get("external_area"):
                highlights.append(f"• External Area: {ld['external_area']}")
            if ld.get("outgoings"):
                highlights.append(f"• Outgoings: {ld['outgoings']}")
            if ld.get("car_spaces"):
                highlights.append(f"• Car Spaces: {ld['car_spaces']}")
        else:  # sale_enquiry
            highlights.append(
                f"• Asking Price: {ld.get('asking_price') or '[PLEASE ADD - SALE PRICE]'}"
            )
            highlights.append(
                f"• Internal Floor Area: {ld.get('internal_area') or '[PLEASE ADD - SIZE]'}"
            )
            if ld.get("external_area"):
                highlights.append(f"• External Area: {ld['external_area']}")
            if ld.get("zoning"):
                highlights.append(f"• Zoning: {ld['zoning']}")
            if ld.get("car_spaces"):
                highlights.append(f"• Car Spaces: {ld['car_spaces']}")
        if building_name:
            highlights.append(f"• Part of the award-winning {building_name} development")
        highlights_str = "\n".join(highlights)

        # "Nestled within..." paragraph only when building name is known
        nestled_para = (
            f"Nestled within the {building_name}, this is a rare chance to secure a "
            "premium position in one of Sydney's most iconic precincts.\n\n"
        ) if building_name else ""

        style_block = ""
        if style_examples:
            examples_text = "\n\n---\n".join(style_examples)
            style_block = (
                "\n\n--- Edward's past listing replies (style reference) ---\n"
                f"{examples_text}\n"
                "--- End style reference ---"
            )

        prompt = (
            "You are drafting a reply on behalf of Edward Ghattas, "
            "commercial real estate agent at IB Property Sydney.\n\n"
            "--- Original email ---\n"
            f"From: {email.get('from_name', 'the sender')} <{email.get('from', '')}>\n"
            f"Subject: {email.get('subject', '')}\n"
            f"Body:\n{email.get('body', '')[:1200]}"
            f"{attachments_note}"
            f"{style_block}\n\n"
            "--- EXACT TEMPLATE TO USE ---\n"
            "Reproduce this structure EXACTLY — do not add, remove, or reorder any section:\n\n"
            "Hi [sender's first name],\n\n"
            "Hope all is well.\n\n"
            f"IB Property is pleased to bring to market {address}, an exceptional "
            f"{prop_type} opportunity available {action_type}.\n\n"
            f"{nestled_para}"
            "**Property Highlights**\n"
            f"{highlights_str}\n\n"
            "For further information or to arrange an inspection, please don't hesitate "
            "to reach out to our exclusive listing agents.\n\n"
            "We look forward to hearing from you.\n\n"
            "Edward Ghattas\n"
            "IB Property Sydney\n"
            "edward@ibproperty.com.au\n\n"
            "--- RENDERING INSTRUCTIONS ---\n"
            "- Replace [sender's first name] with the actual first name from the From field\n"
            "- Any value shown as [PLEASE ADD - X] must remain exactly as-is in the output\n"
            "- Return ONLY the HTML body using <p>, <strong>, and <br> tags\n"
            "- Render '**Property Highlights**' as <strong>Property Highlights</strong>\n"
            "- Each bullet point on its own line with a • character\n"
            "- Do not include a subject line"
        )
    else:
        category_context = {
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
            max_tokens=800,
            messages=[{"role": "user", "content": prompt}],
        )
        html_body = response.content[0].text.strip()
        # Strip markdown code fences if Claude wrapped the output
        html_body = re.sub(r'^```(?:html)?\s*', '', html_body, flags=re.IGNORECASE)
        html_body = re.sub(r'\s*```$', '', html_body)
        html_body = html_body.strip()
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

    # Fetch style examples from Sent Items (used for enquiry drafts)
    logger.info("Fetching sent enquiry reply examples for style matching...")
    sent_enquiry_examples = fetch_sent_enquiry_examples(outlook)
    logger.info("Found %d past enquiry reply example(s)", len(sent_enquiry_examples))

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
            logger.info("SKIPPED (promotional): %s | from=%s", subject, from_addr)
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

        # 3. Collect property data and draft reply via Claude
        is_enq = category in ("lease_enquiry", "sale_enquiry")
        examples = sent_enquiry_examples if is_enq else None
        listing_details = collect_property_data(outlook, raw, email, category) if is_enq else None
        if listing_details:
            logger.info("  Property data: %s", {k: v for k, v in listing_details.items() if v})
        reply_subject, html_body = claude_draft_reply(
            ai, email, category, related_attachments,
            style_examples=examples, listing_details=listing_details,
        )

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


def test_draft(max_scan: int = 20, subject_filter: str = "") -> None:
    """Fetch a matching enquiry email from inbox and print the draft — nothing is saved."""
    filter_desc = f" matching {subject_filter!r}" if subject_filter else ""
    logger.info("=== TEST MODE — printing draft, not saving%s ===", filter_desc)

    missing = [k for k in ("ANTHROPIC_API_KEY", "AZURE_CLIENT_ID", "AZURE_TENANT_ID") if not os.getenv(k)]
    if missing:
        logger.error("Missing required env vars: %s", ", ".join(missing))
        sys.exit(1)

    msal_cache = os.getenv("MSAL_TOKEN_CACHE_PATH", "msal_token_cache.bin")
    outlook = OutlookClient(
        client_id=os.getenv("AZURE_CLIENT_ID"),
        tenant_id=os.getenv("AZURE_TENANT_ID"),
        token_cache_path=msal_cache,
    )
    ai = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    logger.info("Fetching style examples from Sent Items...")
    sent_examples = fetch_sent_enquiry_examples(outlook)
    logger.info("Found %d style example(s)", len(sent_examples))

    logger.info("Fetching recent inbox emails (last 30 days)...")
    raw_messages = outlook.get_recent_emails(since_days=30, extra_folders=["Front of Mind"])
    logger.info("Total messages fetched: %d", len(raw_messages))

    for raw in raw_messages[:max_scan]:
        email = OutlookClient.extract_email_data(raw)
        if is_self_sent(email) or is_automated(email):
            continue

        # Subject filter (case-insensitive substring match)
        if subject_filter and subject_filter.lower() not in email.get("subject", "").lower():
            continue

        category = claude_classify(ai, email)
        if category not in ("lease_enquiry", "sale_enquiry"):
            logger.info("  Skipping (category=%s): %s", category, email.get("subject"))
            continue

        logger.info("Found enquiry: %s | category=%s", email.get("subject"), category)
        listing_details = collect_property_data(outlook, raw, email, category)
        logger.info(
            "Property data collected: %s",
            {k: v for k, v in listing_details.items() if v},
        )

        _, html_body = claude_draft_reply(
            ai, email, category, [],
            style_examples=sent_examples,
            listing_details=listing_details,
        )
        plain = re.sub(r"<[^>]+>", " ", html_body)
        plain = re.sub(r"\s{2,}", " ", plain).strip()

        print("\n" + "=" * 70)
        print(f"TO:      {email.get('from_name')} <{email.get('from')}>")
        print(f"SUBJECT: Re: {email.get('subject')}")
        print("=" * 70)
        print(plain)
        print("=" * 70)
        return

    logger.warning(
        "No enquiry email found in the first %d messages scanned%s.",
        max_scan, filter_desc,
    )


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="IB Property draft reply agent")
    parser.add_argument(
        "--test",
        action="store_true",
        help="Print a draft for the first enquiry found, do not save to Drafts",
    )
    parser.add_argument(
        "--subject",
        default="",
        help="Filter: only test against emails whose subject contains this string",
    )
    args = parser.parse_args()

    if args.test:
        test_draft(subject_filter=args.subject)
    else:
        main()
