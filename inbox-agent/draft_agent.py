"""Draft reply agent for IB Property inbox.

Reads Outlook inbox (last 48 hrs), classifies each email via Claude,
searches for related property emails, then saves reply drafts to both
Outlook Drafts and Gmail Drafts — nothing is sent automatically.

Usage:
    python draft_agent.py
"""

import html as _html
import json
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
    "Price Guide",
    "private treaty",
    "Information Memorandum",
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


_ENQUIRY_REPLY_MARKERS = [
    "pleased to present",
    "pleased to bring to market",
    "ib property is pleased",
    "property highlights",
    "floor area",
    "asking rent",
    "price guide",
    "information memorandum",
    "private treaty",
]


def fetch_sent_reply_for_address(
    outlook: OutlookClient, address: Optional[str]
) -> Optional[str]:
    """Search Sent Items for the most recent enquiry reply Eddie sent about this property.

    Returns the plain-text body of the best match, or None if not found.
    """
    if not address:
        return None

    short = _shorten_address(address)
    if not short:
        return None

    # Search by street name only — more robust than full address
    street_m = re.search(r'\d+\w?\s+(\w+)', short)
    search_term = street_m.group(1) if street_m else short

    logger.info("  [Sent] Searching for past reply to: %r", search_term)
    msgs = _search_folder_body(
        outlook, "/me/mailFolders/sentitems/messages", search_term, top=10
    )
    logger.info("  [Sent] Found %d sent messages", len(msgs))

    for msg in msgs:
        plain = _body_to_plain(msg)
        raw_html = msg.get("body", {}).get("content", "") if isinstance(msg.get("body"), dict) else ""
        combined = (plain + " " + raw_html).lower()
        marker_hits = sum(1 for m in _ENQUIRY_REPLY_MARKERS if m in combined)
        if marker_hits >= 2:
            logger.info(
                "  [Sent] Matched reply: %s (markers=%d)",
                msg.get("subject", ""), marker_hits,
            )
            return plain

    logger.info("  [Sent] No matching enquiry reply found for %r", search_term)
    return None


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


# ─── Listings database ───────────────────────────────────────────────────────

_LISTINGS_DB_PATH = os.path.join(os.path.dirname(__file__), "listings_db.json")


def load_listings_db() -> Dict:
    """Load listings_db.json. Returns empty DB if file doesn't exist yet."""
    try:
        with open(_LISTINGS_DB_PATH) as fh:
            return json.load(fh)
    except (FileNotFoundError, json.JSONDecodeError):
        return {"listings": []}


def find_listing_in_db(db: Dict, address: str) -> Optional[Dict]:
    """Find the best matching listing by street number + first street-name word."""
    if not address:
        return None
    addr_lower = address.lower()

    num_m = re.search(r'\b(\d+\w?)\b', addr_lower)
    street_m = re.search(r'\d+\w?\s+(\w+)', addr_lower)
    if not num_m or not street_m:
        return None

    street_num = num_m.group(1)
    street_word = street_m.group(1)

    for listing in db.get("listings", []):
        db_addr = listing.get("address", "").lower()
        if street_num in db_addr and street_word in db_addr:
            return listing

    return None


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

    # ── Rent (lease only) ─────────────────────────────────────────────────────
    if category == "lease_enquiry":
        m = re.search(
            r'(?:asking\s+rent|rental|rent(?:al)?)\s*[:\-]?\s*' + _DOLLAR,
            text, re.IGNORECASE,
        )
        if not m:
            m = re.search(_DOLLAR + r'\s*(?:p\.?a\.?|per\s+annum|pa\b)', text, re.IGNORECASE)
        if not m:
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

    # ── Sale-specific fields ───────────────────────────────────────────────────
    if category == "sale_enquiry":
        # Current net rent
        m = re.search(
            r'(?:current\s+net|net\s+rent|net\s+income|net\s+return)\s*[:\-]?\s*' + _DOLLAR,
            text, re.IGNORECASE,
        )
        if m:
            found["net_rent"] = f"${m.group(1)} + GST"

        # Estimated fully leased rent
        m = re.search(
            r'(?:(?:estimated\s+)?fully\s+leas(?:ed|t)|fully\s+let|'
            r'potential\s+(?:net\s+)?(?:rent|income))(?:[^$\n]{0,60})?' + _DOLLAR,
            text, re.IGNORECASE,
        )
        if m:
            found["fully_leased_rent"] = f"${m.group(1)} + GST"

        # Land area — explicit label, extracted BEFORE floor area to avoid confusion
        m = re.search(
            r'(?:land\s+area|site\s+area|land)\s*[:\-]?\s*' + _AREA_VAL + _AREA_UNIT,
            text, re.IGNORECASE,
        )
        if m:
            found["land_area"] = f"{m.group(1)} sqm*"

        # Sale price / price guide — find ALL labelled matches, take the highest
        # (listing price > contract price; avoids picking up contract-of-sale figures)
        labelled_prices = re.findall(
            r'(?:price\s+guide|asking\s+price|sale\s+price|for\s+sale|offers?\s+(?:over|above|from))\s*[:\-]?\s*'
            + _DOLLAR,
            text, re.IGNORECASE,
        )
        if labelled_prices:
            try:
                best = max(labelled_prices, key=lambda x: float(x.replace(",", "")))
                found["asking_price"] = f"${best}"
            except ValueError:
                found["asking_price"] = f"${labelled_prices[0]}"
        if not found.get("asking_price"):
            for candidate in re.finditer(r'\$([\d]{1,3}(?:,[\d]{3})+(?:\.\d+)?)', text):
                try:
                    if float(candidate.group(1).replace(",", "")) >= 500_000:
                        found["asking_price"] = f"${candidate.group(1)}"
                        break
                except ValueError:
                    pass
        if not found.get("asking_price") and re.search(r'\bEOI\b|expressions?\s+of\s+interest', text, re.IGNORECASE):
            found["asking_price"] = "EOI — Expressions of Interest"

        # IM URL: heyzine first, then any flipbook URL, then any URL near IM keyword
        im_m = re.search(r'(https?://heyzine\.com[^\s"<>]*)', text, re.IGNORECASE)
        if not im_m:
            im_m = re.search(r'(https?://[^\s"<>]*flipbook[^\s"<>]*)', text, re.IGNORECASE)
        if not im_m:
            im_m = re.search(
                r'(?:information\s+memorandum|(?<!\w)IM(?!\w))[^\n]{0,300}?(https?://[^\s"<>]+)',
                text, re.IGNORECASE | re.DOTALL,
            )
        if im_m:
            found["im_url"] = im_m.group(1)

        # Property type
        m = re.search(
            r'\b(freehold\s+investment|freehold\s+(?:commercial|retail|office|property)|'
            r'commercial\s+(?:property|investment)|retail\s+investment|'
            r'mixed[\s-]use(?:\s+investment)?|strata\s+title|industrial\s+(?:property|investment))\b',
            text, re.IGNORECASE,
        )
        if m:
            found["property_type"] = m.group(1).title()

        # Lease terms — grab relevant sentences for Claude to reformat
        lease_sents = re.findall(
            r'[^.!?\n]*(?:(?:lease|tenancy)\s+(?:term|variation|extension|expires?|commenc)|'
            r'tenant\s+(?:has|recently|signed)|option(?:s)?\s+(?:to\s+)?(?:renew|extend|purchase)|'
            r'further\s+\d+\s*[×x]\s*\d+)[^.!?\n]*[.!?]?',
            text, re.IGNORECASE,
        )
        if lease_sents:
            found["lease_summary"] = " ".join(s.strip() for s in lease_sents[:3])

        # Vacancy / value-add sentences (skip listing spec contamination)
        _SPEC_MARKERS = re.compile(
            r'land\s+area|floor\s+area|asking\s+rent|net\s+rent|price\s+guide|sqm|m²', re.IGNORECASE
        )
        vacancy_sents = [
            s for s in re.findall(
                r'[^.!?\n]{0,200}(?:vacant|vacancy|value[\s-]add|upstairs|level\s+1\s+(?:is\s+)?vacant|'
                r'additional\s+income|income\s+(?:opportunity|potential)|currently\s+unoccupied)[^.!?\n]{0,200}[.!?]?',
                text, re.IGNORECASE,
            )
            if not _SPEC_MARKERS.search(s)
        ]
        if vacancy_sents:
            found["vacancy_note"] = " ".join(s.strip() for s in vacancy_sents[:2])

    # ── Internal floor area ────────────────────────────────────────────────────
    # "area" and "size" deliberately excluded — too broad; would match "land area"
    m = re.search(
        r'(?:nla|gfa|internal\s+(?:floor\s+)?area|floor\s+area|'
        r'(?:total\s+)?floor\s+(?:space|plate)|lettable\s+area|net\s+lettable|gross\s+floor)\s*[:\-]?\s*'
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

    # Fallback: unlabelled sqm values; skip the land_area value if already found
    if not found.get("internal_area"):
        land_val = found.get("land_area", "").replace(" sqm*", "") if found.get("land_area") else None
        all_areas = re.findall(_AREA_VAL + _AREA_UNIT, text, re.IGNORECASE)
        remaining = [a for a in all_areas if a != land_val] if land_val else all_areas
        if remaining:
            found["internal_area"] = f"{remaining[0]} sqm*"
            if len(remaining) > 1 and not found.get("external_area"):
                found["external_area"] = f"{remaining[1]} sqm*"

    # ── Building name ──────────────────────────────────────────────────────────
    _GENERIC_WORDS = {'the', 'this', 'that', 'here', 'there', 'conversation',
                      'development', 'building', 'complex', 'property', 'address'}
    for pat in [
        r'\b(Sirius[^,.\n]*)',
        r'award[- ]winning\s+([A-Z][A-Za-z0-9\s]+?)(?:\s+development|\s+building|\s+complex|[,.\n])',
        r'(?:located\s+(?:in|within)|nestled\s+(?:in|within)|part\s+of)\s+the\s+([A-Z][A-Za-z0-9\s]+?)(?:\s+development|\s+building|\s+complex|[,.\n])',
        r'\bthe\s+([A-Z][A-Za-z0-9\s]+(?:Centre|Center|Tower|Building|Plaza|House|Court|Arcade|Mall))',
    ]:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            name = m.group(1).strip()
            # Reject if first character isn't actually uppercase (IGNORECASE made [A-Z] match lower)
            if name and name[0].isupper() and name.lower() not in _GENERIC_WORDS:
                found["building_name"] = name
                break

    # ── Outgoings ──────────────────────────────────────────────────────────────
    m = re.search(r'outgoings?\s*[:\-]?\s*' + _DOLLAR + r'[^.\n]{0,20}', text, re.IGNORECASE)
    if m:
        found["outgoings"] = f"${m.group(1)} p.a."

    # ── Car spaces — store raw number; template decides formatting ─────────────
    m = re.search(
        r'(?:secure\s+)?(?:parking\s+for\s+(\d+)|(\d+)\s*(?:car\s+(?:space|park|bay)s?|'
        r'parking\s+(?:space|bay)s?))',
        text, re.IGNORECASE,
    )
    if m:
        found["car_spaces"] = str(int(m.group(1) or m.group(2)))

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
    """Extract and flatten HTML body from a Graph message dict, decoding HTML entities."""
    raw = msg.get("body", {}).get("content", "") if isinstance(msg.get("body"), dict) else ""
    decoded = _html.unescape(raw).replace("\xa0", " ")  # &nbsp; → space
    plain = re.sub(r"<[^>]+>", " ", decoded)
    return re.sub(r"\s{2,}", " ", plain).strip()


def collect_property_data(
    outlook: OutlookClient,
    raw_message: Dict,
    email: Dict,
    category: str,
    listings_db: Optional[Dict] = None,
) -> Dict[str, Optional[str]]:
    """
    Gather property data in priority order:
      DB) listings_db.json — pre-scraped from Eddie's own vendor update emails (primary)
      A)  The enquiry email — raw HTML body
      B)  Sent Items search by address (only when no DB match)
      C)  Full inbox search by address (only when no DB match)
    """
    details: Dict[str, Optional[str]] = {
        "address": None,
        "suburb": None,
        # lease
        "asking_rent": None,
        # sale
        "asking_price": None,
        "net_rent": None,
        "fully_leased_rent": None,
        "land_area": None,
        "property_type": None,
        "im_url": None,
        "lease_summary": None,
        "vacancy_note": None,
        # shared
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

    # Extract suburb from address (e.g. "14A Hannah Street, Beecroft NSW 2119" → "Beecroft")
    if details["address"]:
        sm = re.search(
            r',\s*([A-Z][A-Za-z\s]+?)\s+(?:NSW|VIC|QLD|SA|WA|TAS|ACT|NT)\b',
            details["address"], re.IGNORECASE,
        )
        if not sm:
            # Address without state: "14a Hannah St, Beecroft" → last comma-separated word(s)
            sm = re.search(
                r',\s*([A-Z][A-Za-z]+(?:\s+[A-Z][A-Za-z]+)?)\s*(?:\d{4})?$',
                details["address"].strip(), re.IGNORECASE,
            )
        if sm:
            details["suburb"] = sm.group(1).strip()

    # ── Source A: enquiry email, raw HTML ─────────────────────────────────────
    raw_html = (
        raw_message.get("body", {}).get("content", "")
        if isinstance(raw_message.get("body"), dict)
        else ""
    )
    html_plain = re.sub(r"\s{2,}", " ", re.sub(r"<[^>]+>", " ", _html.unescape(raw_html).replace("\xa0", " "))).strip()
    _merge_into(details, _parse_property_fields(f"{subject}\n{html_plain}", category))
    logger.debug("  [A] After email body: %s", {k: v for k, v in details.items() if v})

    # ── Source DB: listings_db.json (pre-scraped from vendor update emails) ────
    db_matched = False
    if listings_db and details.get("address"):
        db_match = find_listing_in_db(listings_db, details["address"])
        if db_match:
            logger.info("  [DB] Match in listings_db: %s", db_match.get("address"))
            # DB data is authoritative — overwrite anything Source A may have found
            if db_match.get("asking_rent"):
                details["asking_rent"] = db_match["asking_rent"]
            if db_match.get("price"):
                details["asking_price"] = db_match["price"]
            if db_match.get("net_rent"):
                details["net_rent"] = db_match["net_rent"]
            if db_match.get("fully_leased_rent"):
                details["fully_leased_rent"] = db_match["fully_leased_rent"]
            if db_match.get("floor_area"):
                details["internal_area"] = db_match["floor_area"]
            if db_match.get("land_area"):
                details["land_area"] = db_match["land_area"]
            if db_match.get("building"):
                details["building_name"] = db_match["building"]
            if db_match.get("im_url"):
                details["im_url"] = db_match["im_url"]
            if db_match.get("lease_terms"):
                details["lease_summary"] = db_match["lease_terms"]
            if db_match.get("notes"):
                details["vacancy_note"] = db_match["notes"]
            if db_match.get("parking"):
                car_m = re.search(r'\d+', db_match["parking"])
                if car_m:
                    details["car_spaces"] = car_m.group()
            db_matched = True
            logger.debug("  [DB] After DB merge: %s", {k: v for k, v in details.items() if v})

    if not db_matched:
        # ── Source B: Sent Items search by street address ──────────────────────
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

        # ── Source C: full inbox search by street address ──────────────────────
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
    if not details.get("internal_area"):
        details["internal_area"] = "[PLEASE ADD - SIZE]"

    if category == "lease_enquiry":
        if not details.get("asking_rent"):
            details["asking_rent"] = "[PLEASE ADD - RENT/PRICE]"
    elif category == "sale_enquiry":
        if not details.get("asking_price"):
            details["asking_price"] = "[PLEASE ADD - SALE PRICE]"
        if not details.get("net_rent"):
            details["net_rent"] = "[PLEASE ADD - NET RENT]"
        if not details.get("im_url"):
            details["im_url"] = "[PLEASE ADD - IM LINK]"
        if not details.get("property_type"):
            details["property_type"] = "Freehold Investment"
        if not details.get("suburb"):
            details["suburb"] = "[SUBURB]"

    return details


# ─── Claude helpers ───────────────────────────────────────────────────────────

def claude_classify(ai: anthropic.Anthropic, email: Dict) -> Tuple[str, bool]:
    """Return (category, is_lion). is_lion=True means a reply draft should be created."""
    prompt = (
        "You are an assistant for IB Property Sydney, a commercial real estate agency.\n\n"
        "Classify the following email and decide if it is urgent enough to require a reply draft.\n\n"
        "Categories:\n"
        "  lease_enquiry   — enquiry about leasing a property\n"
        "  sale_enquiry    — enquiry about buying or selling a property\n"
        "  vendor_update   — update from a vendor, supplier, or tradesperson\n"
        "  landlord_query  — query or request from a landlord or property owner\n"
        "  general         — anything else\n\n"
        "An email IS a lion (is_lion: true) if:\n"
        "- It is a direct enquiry from a potential buyer or tenant\n"
        "- It has a direct question that needs answering from a client, landlord, or vendor\n"
        "- It involves an offer, contract, or negotiation\n"
        "- It is from a known contact asking something specific\n"
        "- Not replying would cause a missed deal, upset a client, or create a problem\n\n"
        "An email is NOT a lion (is_lion: false) if:\n"
        "- It is a listing performance report or campaign stats from a portal (realcommercial, commercialrealestate)\n"
        "- It is a FYI or announcement (LEASED / SOLD notices that need no reply)\n"
        "- It is a weekly or monthly stats digest or purely informational portal notification\n"
        "- It is from Grammarly, a newsletter, or a marketing list\n"
        "- It is an internal IB Property broadcast that requires no reply\n"
        "- It is just an update with no question or action required\n\n"
        "Ask yourself: Is this email urgent enough that not replying would cause a problem? "
        "If it is just an informational update, report, or announcement, answer false.\n\n"
        f"From: {email.get('from_name', '')} <{email.get('from', '')}>\n"
        f"Subject: {email.get('subject', '')}\n"
        f"Body: {email.get('body', '')[:600]}\n\n"
        'Reply with ONLY valid JSON in this exact format: {"category": "<category>", "is_lion": <true|false>}'
    )
    try:
        response = ai.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=60,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = response.content[0].text.strip()
        # Strip markdown fences if present
        raw = re.sub(r'^```(?:json)?\s*', '', raw, flags=re.IGNORECASE)
        raw = re.sub(r'\s*```$', '', raw).strip()
        data = json.loads(raw)
        category = data.get("category", "general").lower().strip()
        is_lion = bool(data.get("is_lion", False))
        if category not in CATEGORIES:
            category = "general"
        return category, is_lion
    except Exception as exc:
        logger.error("Claude classify failed: %s", exc)
        return "general", False


def claude_draft_reply(
    ai: anthropic.Anthropic,
    email: Dict,
    category: str,
    related_attachments: List[str],
    style_examples: Optional[List[str]] = None,
    listing_details: Optional[Dict] = None,
    sent_template: Optional[str] = None,
) -> Tuple[str, str]:
    """Return (reply_subject, html_body) for a professional CRE reply."""
    attachments_note = ""
    if related_attachments:
        names = ", ".join(related_attachments[:10])
        attachments_note = f"\n\nRelated documents found in the thread: {names}"

    # ── Sent Items template path — replicate Eddie's previous reply almost verbatim ──
    if sent_template and category in ("lease_enquiry", "sale_enquiry"):
        from_name = email.get("from_name", "").strip()
        enquirer_first = from_name.split()[0] if from_name else "there"
        prompt = (
            "You are drafting a reply on behalf of Edward Ghattas, "
            "commercial real estate agent at IB Property Sydney.\n\n"
            "--- Original enquiry ---\n"
            f"From: {email.get('from_name', 'the sender')} <{email.get('from', '')}>\n"
            f"Subject: {email.get('subject', '')}\n"
            f"Body:\n{email.get('body', '')[:800]}"
            f"{attachments_note}\n\n"
            "--- Previous reply Edward sent about this property ---\n"
            f"{sent_template[:2500]}\n\n"
            "--- Instructions ---\n"
            "This is the exact email Edward sent previously about this property. "
            "Replicate it almost word for word. Only change:\n"
            f"(a) the recipient's first name to: {enquirer_first}\n"
            "(b) any direct reference to the previous enquirer's name elsewhere in the body\n"
            "Keep everything else identical — the property details, the tone, "
            "the structure, the sign-off.\n\n"
            "Return ONLY the HTML body content using <p>, <strong>, and <br> tags. "
            "Do not include a subject line inside the body."
        )
        try:
            response = ai.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=1200,
                messages=[{"role": "user", "content": prompt}],
            )
            html_body = response.content[0].text.strip()
            html_body = re.sub(r'^```(?:html)?\s*', '', html_body, flags=re.IGNORECASE)
            html_body = re.sub(r'\s*```$', '', html_body).strip()
        except Exception as exc:
            logger.error("Claude sent-template draft failed: %s", exc)
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

    if category == "lease_enquiry":
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

        highlights: List[str] = [
            f"• Asking Rent: {ld.get('asking_rent') or '[PLEASE ADD - RENT/PRICE]'}",
            f"• Internal Floor Area: {ld.get('internal_area') or '[PLEASE ADD - SIZE]'}",
        ]
        if ld.get("external_area"):
            highlights.append(f"• External Area: {ld['external_area']}")
        if ld.get("outgoings"):
            highlights.append(f"• Outgoings: {ld['outgoings']}")
        if ld.get("car_spaces"):
            n = ld["car_spaces"]
            highlights.append(f"• Car Spaces: {n} car space{'s' if n != '1' else ''}")
        if building_name:
            highlights.append(f"• Part of the award-winning {building_name} development")
        highlights_str = "\n".join(highlights)

        nestled_para = (
            f"Nestled within the {building_name}, this is a rare chance to secure a "
            "premium position in one of Sydney's most iconic precincts.\n\n"
        ) if building_name else ""

        style_block = ""
        if style_examples:
            examples_text = "\n\n---\n".join(style_examples)
            style_block = (
                "\n\n--- Edward's past listing replies (style reference) ---\n"
                f"{examples_text}\n--- End style reference ---"
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
            "Reproduce this structure EXACTLY. Fill in [sender's first name]. "
            "Leave any [PLEASE ADD - X] untouched:\n\n"
            "Hi [sender's first name],\n\n"
            "Hope all is well.\n\n"
            f"IB Property is pleased to bring to market {address}, an exceptional "
            f"{prop_type} opportunity available for lease.\n\n"
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
            "- Return ONLY the HTML body using <p>, <strong>, and <br> tags\n"
            "- Render '**Property Highlights**' as <strong>Property Highlights</strong>\n"
            "- Each bullet on its own line with a • character\n"
            "- Do not include a subject line"
        )

    elif category == "sale_enquiry":
        ld = listing_details or {}
        address = ld.get("address") or "[PLEASE ADD - PROPERTY ADDRESS]"
        property_type = ld.get("property_type") or "Freehold Investment"
        suburb = ld.get("suburb") or "[SUBURB]"

        # Build the bullet list
        bullets: List[str] = []
        if ld.get("net_rent"):
            bullets.append(f"• Current net rent: {ld['net_rent']}")
        if ld.get("fully_leased_rent"):
            bullets.append(f"• Estimated Fully Leased net rent: {ld['fully_leased_rent']}")
        if ld.get("land_area"):
            bullets.append(f"• Land area: {ld['land_area']}")
        bullets.append(f"• Floor area: {ld.get('internal_area') or '[PLEASE ADD - SIZE]'}")
        if ld.get("car_spaces"):
            n = ld["car_spaces"]
            bullets.append(f"• Secure parking for {n} car{'s' if n != '1' else ''}")
        bullets.append(
            f"• Prime location near {suburb} village, train station, cafés, restaurants, and amenities"
        )
        bullets_str = "\n".join(bullets)

        # Lease paragraph (extracted sentences, Claude reformats to final paragraph)
        if ld.get("lease_summary"):
            lease_para = f"{ld['lease_summary']}\n\n"
        else:
            lease_para = ""

        # Vacancy paragraph (use extracted text verbatim)
        if ld.get("vacancy_note"):
            vacancy_para = f"{ld['vacancy_note']}\n\n"
        else:
            vacancy_para = ""

        # Price and IM
        price_val = ld.get("asking_price") or "[PLEASE ADD - SALE PRICE]"
        if price_val.startswith("["):
            price_str = price_val
        else:
            price_str = f"{price_val} + GST"
        im_url = ld.get("im_url") or "[PLEASE ADD - IM LINK]"

        style_block = ""
        if style_examples:
            examples_text = "\n\n---\n".join(style_examples)
            style_block = (
                "\n\n--- Edward's past sale replies (style reference) ---\n"
                f"{examples_text}\n--- End style reference ---"
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
            "Reproduce this structure EXACTLY. Fill in [sender's first name]. "
            "Leave any [PLEASE ADD - X] untouched:\n\n"
            "Hi [sender's first name],\n\n"
            "Hope you are well.\n\n"
            f"IB Property is pleased to present {address}, to the market for sale via "
            f"private treaty, an exceptional {property_type}.\n\n"
            f"{bullets_str}\n\n"
            f"{lease_para}"
            f"{vacancy_para}"
            "**Price Guide**\n\n"
            f"For sale {price_str}\n\n"
            "**Information Memorandum:**\n\n"
            f"{im_url}\n\n"
            "We look forward to hearing from you.\n\n"
            "Edward Ghattas\n"
            "IB Property Sydney\n"
            "edward@ibproperty.com.au\n\n"
            "--- RENDERING INSTRUCTIONS ---\n"
            "- Return ONLY the HTML body using <p>, <strong>, and <br> tags\n"
            "- Render '**Price Guide**' as <strong>Price Guide</strong>\n"
            "- Render '**Information Memorandum:**' as <strong>Information Memorandum:</strong>\n"
            "- Each bullet on its own line with a • character\n"
            "- Reproduce the lease and vacancy paragraphs exactly as given, do not rephrase\n"
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
            max_tokens=1200,
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

    # Load property listings database
    listings_db = load_listings_db()
    n_listings = len(listings_db.get("listings", []))
    if n_listings:
        logger.info("Loaded listings_db.json: %d listing(s)", n_listings)
    else:
        logger.info("listings_db.json not found or empty — run refresh_listings_db.py to build it")

    # Fetch style examples from Sent Items (used for enquiry drafts)
    logger.info("Fetching sent enquiry reply examples for style matching...")
    sent_enquiry_examples = fetch_sent_enquiry_examples(outlook)
    logger.info("Found %d past enquiry reply example(s)", len(sent_enquiry_examples))

    # Fetch last 48 hours from Inbox + "Front Of Mind"
    logger.info("Fetching Outlook inbox + 'Front Of Mind' (last 48 hours)...")
    raw_messages = outlook.get_recent_emails(since_days=2, extra_folders=["Front Of Mind"])
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

        # 1. Classify + lion check
        category, is_lion = claude_classify(ai, email)
        logger.info("  Category → %s | is_lion=%s", category, is_lion)

        if not is_lion:
            logger.info("SKIPPED (not urgent): %s", subject)
            skipped += 1
            continue

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

        # 3. Collect property data / find sent template, then draft reply via Claude
        is_enq = category in ("lease_enquiry", "sale_enquiry")
        sent_template: Optional[str] = None
        listing_details = None

        if is_enq:
            # Try to find a past sent reply for the same property first
            address = _extract_address(email.get("subject", ""), email.get("body", "")[:500])
            sent_template = fetch_sent_reply_for_address(outlook, address)
            if sent_template:
                logger.info("  Using Sent Items template for reply")
            else:
                # Fall back to listings_db / multi-source data collection
                listing_details = collect_property_data(
                    outlook, raw, email, category, listings_db=listings_db
                )
                logger.info("  Property data: %s", {k: v for k, v in listing_details.items() if v})

        examples = sent_enquiry_examples if is_enq and not sent_template else None
        reply_subject, html_body = claude_draft_reply(
            ai, email, category, related_attachments,
            style_examples=examples,
            listing_details=listing_details,
            sent_template=sent_template,
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

    listings_db = load_listings_db()
    n_listings = len(listings_db.get("listings", []))
    logger.info(
        "Loaded listings_db.json: %d listing(s)" if n_listings else
        "listings_db.json empty — run refresh_listings_db.py",
        n_listings,
    )

    logger.info("Fetching style examples from Sent Items...")
    sent_examples = fetch_sent_enquiry_examples(outlook)
    logger.info("Found %d style example(s)", len(sent_examples))

    logger.info("Fetching recent inbox emails (last 30 days)...")
    raw_messages = outlook.get_recent_emails(since_days=30, extra_folders=["Front Of Mind"])
    logger.info("Total messages fetched: %d", len(raw_messages))

    scan_limit = len(raw_messages) if subject_filter else max_scan
    for raw in raw_messages[:scan_limit]:
        email = OutlookClient.extract_email_data(raw)
        if is_self_sent(email) or is_automated(email):
            continue

        # Subject filter (case-insensitive substring match)
        if subject_filter and subject_filter.lower() not in email.get("subject", "").lower():
            continue

        category, is_lion = claude_classify(ai, email)
        if category not in ("lease_enquiry", "sale_enquiry"):
            logger.info("  Skipping (category=%s): %s", category, email.get("subject"))
            continue
        if not is_lion:
            logger.info("  Skipping (not urgent, is_lion=False): %s", email.get("subject"))
            continue

        logger.info("Found enquiry: %s | category=%s", email.get("subject"), category)

        # Try Sent Items template first
        address = _extract_address(email.get("subject", ""), email.get("body", "")[:500])
        sent_template = fetch_sent_reply_for_address(outlook, address)
        listing_details = None
        if sent_template:
            logger.info("Using Sent Items template for reply")
        else:
            listing_details = collect_property_data(outlook, raw, email, category, listings_db=listings_db)
            logger.info(
                "Property data collected: %s",
                {k: v for k, v in listing_details.items() if v},
            )

        examples = sent_examples if not sent_template else None
        _, html_body = claude_draft_reply(
            ai, email, category, [],
            style_examples=examples,
            listing_details=listing_details,
            sent_template=sent_template,
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
