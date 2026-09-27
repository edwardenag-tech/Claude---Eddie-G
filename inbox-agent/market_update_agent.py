"""Market-update mail-merge agent for IB Property.

Reads owner/vendor contact rows from the live Dropbox "DATA BASE.xlsx"
workbook, segments them by region (one workbook tab per region) then by
suburb (column C, "Suburb"), and creates ONE personalised draft per owner
PER SUBURB in Outlook Drafts (and, only if DRAFT_GMAIL_COPIES is on, Gmail
Drafts too) -- nothing is ever sent automatically, in this trial or
otherwise, without a separate change to explicitly allow it.

Eddie's real letters are per-SUBURB, not per-region: a region like Lower
North Shore covers dozens of suburbs, and each week Eddie supplies a
polished, branded PDF (photos, FOR SALE/FOR LEASE/LEASED sections) for
whichever suburbs he's actually written that week -- never all of them at
once. This script does NOT write the letter. For each suburb Eddie wants to
send, he drops that suburb's designed PDF under market_update_letters/ (see
market_update_letters/README.md for the exact filename and optional
subject/greeting override). A suburb with no PDF supplied this week is
skipped entirely -- never sent to, never invented on Eddie's behalf. Every
recipient's email embeds the actual designed PDF pages as images directly
in the body (not a plain attachment).

This script only handles: reading the live recipient list, segmenting,
grouping by suburb, deduping, rendering Eddie's PDF pages as inline images,
merging text fields, and saving one draft per owner per suburb.

Usage:
    python market_update_agent.py --dry-run                       # preview every segment, nothing saved
    python market_update_agent.py --dry-run --segment upper_north_shore
    python market_update_agent.py --dry-run --limit 5              # preview only the first 5 per segment
    python market_update_agent.py                                  # creates real Outlook (+ optional Gmail) drafts

Status (see chat / commit message for the full report): the workbook's tabs
today only cover Lower North Shore ("Master Database") and Upper North Shore
("Upper North") -- there is no Northern Beaches or City Fringe tab or data in
this file yet. SEGMENTS below reflects that; the other two are commented out
until Eddie confirms where that data should come from.
"""

import argparse
import base64
import json
import logging
import os
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import fitz  # PyMuPDF -- rasterizes Eddie's PDF letters to inline email images
import openpyxl
from dotenv import load_dotenv

from gmail_client import GmailClient, GmailAuthRequired
from outlook_client import OutlookClient, OutlookAuthRequired

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

_HERE = Path(__file__).parent
_LETTERS_DIR = _HERE / "market_update_letters"

# Same policy and env var as draft_agent.py's DRAFT_GMAIL_COPIES: Outlook only
# by default, since a Gmail copy would send from Eddie's personal address
# rather than his IB Property one.
GMAIL_COPIES_ENABLED = os.getenv("DRAFT_GMAIL_COPIES", "").strip().lower() in ("1", "true", "yes")

DATA_BASE_XLSX_PATH = os.getenv(
    "MARKET_UPDATE_DB_PATH",
    str(Path.home() / "Dropbox" / "Data Base" / "Main Data Base" / "DATA BASE.xlsx"),
)

_STATE_PATH = os.getenv(
    "MARKET_UPDATE_STATE_PATH", str(_HERE / "market_update_state.json")
)

# One entry per workbook tab this agent knows how to read. "sheet" must match
# the tab name in DATA BASE.xlsx exactly (case-sensitive, per openpyxl).
#
# Northern Beaches and City Fringe are NOT here: DATA BASE.xlsx (the file
# confirmed as source of truth) has no tab and no Suburb-column data for
# either today. Older backup copies in the same Dropbox folder use a "Fringe"
# tab name (not "City Fringe") and still have no "Northern Beaches" tab
# anywhere. See the status report for what's needed to add them.
SEGMENTS: Dict[str, Dict[str, str]] = {
    "lower_north_shore": {"label": "Lower North Shore", "sheet": "Master Database"},
    "upper_north_shore": {"label": "Upper North Shore", "sheet": "Upper North"},
}

# Column headers this agent requires, matched case-insensitively against the
# workbook's own header row rather than hardcoded positions -- Eddie edits
# this file by hand, and a silently-wrong column offset would be worse than a
# loud failure. "Suburb" is expected in column C (index 2) per the confirmed
# source-of-truth description; that's checked explicitly, not just found by
# name, since a mismatch there could mean the whole workbook's shape changed.
_REQUIRED_COLUMNS = ("Street", "Address", "Suburb", "State", "Postcode",
                     "Vendor First Name", "Vendor Surname Name", "Email")
_SUBURB_COLUMN_LETTER = "C"


class WorkbookShapeError(Exception):
    """The workbook's columns don't match what this agent expects. Raised
    instead of silently reading the wrong field into a live mail-merge."""


@dataclass
class Recipient:
    segment: str
    segment_label: str
    suburb: str
    street: str
    address: str
    state: str
    postcode: str
    first_name: str
    last_name: str
    email: str
    property_type: str = ""
    company_name: str = ""
    source_row: int = 0  # 1-based row number in the sheet, for logs only

    @property
    def full_name(self) -> str:
        name = f"{self.first_name} {self.last_name}".strip()
        return name or self.company_name or "there"

    @property
    def street_address(self) -> str:
        return f"{self.street} {self.address}".strip()

    def merge_fields(self) -> Dict[str, str]:
        return {
            "first_name": self.first_name or "there",
            "last_name": self.last_name,
            "full_name": self.full_name,
            "suburb": self.suburb.title(),
            "street_address": self.street_address.title(),
            "state": self.state,
            "postcode": str(self.postcode),
            "property_type": self.property_type,
        }


# ─── Workbook reading ───────────────────────────────────────────────────────

def _header_index_map(header_row: tuple) -> Dict[str, int]:
    return {
        str(v).strip().lower(): i
        for i, v in enumerate(header_row)
        if v is not None and str(v).strip()
    }


def _find_header_row(ws) -> tuple:
    """DATA BASE.xlsx's real column headers are on row 2 (row 1 is a set of
    merged group titles for the deal-stage columns further right) -- but that
    is discovered here by content, not assumed, so a reordered sheet fails
    loudly instead of reading garbage into a live mail-merge."""
    for row in ws.iter_rows(min_row=1, max_row=5, values_only=True):
        values = {str(v).strip().lower() for v in row if v}
        if "suburb" in values and "email" in values:
            return row
    raise WorkbookShapeError(
        f"Could not find a header row with 'Suburb' and 'Email' columns in the "
        f"first 5 rows of sheet {ws.title!r}. The workbook's layout may have changed."
    )


def read_recipients(
    xlsx_path: str, segment_key: str, sheet_name: str, segment_label: str
) -> List[Recipient]:
    """Read every usable row from one workbook tab. Never writes to the file."""
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise WorkbookShapeError(
            f"Sheet {sheet_name!r} (segment {segment_key!r}) not found in "
            f"{xlsx_path!r}. Sheets present: {wb.sheetnames}"
        )
    ws = wb[sheet_name]
    header = _find_header_row(ws)
    cols = _header_index_map(header)

    missing = [c for c in _REQUIRED_COLUMNS if c.lower() not in cols]
    if missing:
        raise WorkbookShapeError(
            f"Sheet {sheet_name!r} is missing required column(s) {missing} -- "
            f"found columns: {list(cols)}"
        )
    suburb_idx = cols["suburb"]
    actual_letter = openpyxl.utils.get_column_letter(suburb_idx + 1)
    if actual_letter != _SUBURB_COLUMN_LETTER:
        logger.warning(
            "Sheet %r: 'Suburb' is in column %s, not %s as originally confirmed -- "
            "reading it by header name regardless, but flagging the mismatch.",
            sheet_name, actual_letter, _SUBURB_COLUMN_LETTER,
        )

    def cell(row, name):
        v = row[cols[name.lower()]]
        if v is None or (isinstance(v, str) and v.strip() in ("", "\xa0")):
            return ""
        return str(v).strip()

    header_row_num = next(
        i for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), 1)
        if row == header
    )

    recipients: List[Recipient] = []
    for row_num, row in enumerate(ws.iter_rows(min_row=header_row_num + 1, values_only=True), header_row_num + 1):
        if not any(row):
            continue
        suburb = cell(row, "Suburb")
        email = cell(row, "Email")
        if not suburb and not email:
            continue  # fully blank data row
        if not email or "@" not in email:
            continue  # can't draft an email without one
        recipients.append(Recipient(
            segment=segment_key,
            segment_label=segment_label,
            suburb=suburb,
            street=cell(row, "Street"),
            address=cell(row, "Address"),
            state=cell(row, "State"),
            postcode=cell(row, "Postcode"),
            first_name=cell(row, "Vendor First Name"),
            last_name=cell(row, "Vendor Surname Name"),
            email=email.lower(),
            property_type=cell(row, "Property Type") if "property type" in cols else "",
            company_name=cell(row, "Company Name") if "company name" in cols else "",
            source_row=row_num,
        ))
    return recipients


def dedupe_by_owner_and_suburb(recipients: List[Recipient]) -> "tuple[List[Recipient], Dict[tuple, int]]":
    """Keep one recipient per (email, suburb) pair within the segment.

    Per Eddie: one email per owner PER SUBURB, not one per owner overall. An
    owner with several properties in the SAME suburb still gets only one
    draft for that suburb (first row seen wins); an owner with properties in
    two different suburbs gets one draft per suburb, each addressed with
    that suburb's own details. Suburb is matched case-insensitively.
    """
    seen: Dict[tuple, Recipient] = {}
    extra_properties: Dict[tuple, int] = {}
    for r in recipients:
        key = (r.email, r.suburb.strip().lower())
        if key in seen:
            extra_properties[key] = extra_properties.get(key, 0) + 1
            continue
        seen[key] = r
    return list(seen.values()), extra_properties


# ─── Letter template ────────────────────────────────────────────────────────

_NOT_WRITTEN_MARKER = "[EDDIE: WRITE THIS LETTER]"
_PDF_RENDER_DPI = 150  # legible in an email body without bloating draft size

_DEFAULT_SUBJECT_TEMPLATE = "Market update for {{suburb}}"
_DEFAULT_GREETING_TEMPLATE = (
    "<p>Hi {{first_name}},</p>\n"
    "<p>Here's this week's market update for {{suburb}}.</p>"
)


def suburb_slug(suburb: str) -> str:
    """Normalize a suburb name to the filename Eddie uses under
    market_update_letters/ (lowercase, non-alphanumerics collapsed to '_').
    Matches DATA BASE.xlsx's ALL CAPS suburbs and however Eddie happens to
    type the filename equally well."""
    return re.sub(r"[^a-z0-9]+", "_", suburb.strip().lower()).strip("_")


@dataclass
class Letter:
    subject_template: str
    greeting_template: str
    pdf_path: Path


def load_letter(suburb: str) -> Optional[Letter]:
    """Load Eddie's letter for one suburb. Returns None unless Eddie has
    actually supplied a designed PDF for this suburb this week
    (market_update_letters/<slug>.pdf) -- callers must treat that as "skip
    this suburb", never draft a generic stand-in on Eddie's behalf. Eddie
    writes the letter itself as that PDF (photos, branding, listings) --
    this script only embeds its pages as images. An optional
    market_update_letters/<slug>.html file can override the default
    subject line and the greeting text shown above the embedded pages;
    format is unchanged from before (a 'Subject: ...' first line, then an
    HTML body with merge fields)."""
    slug = suburb_slug(suburb)
    pdf_path = _LETTERS_DIR / f"{slug}.pdf"
    if not pdf_path.exists():
        return None

    subject_template = _DEFAULT_SUBJECT_TEMPLATE
    greeting_template = _DEFAULT_GREETING_TEMPLATE
    html_path = _LETTERS_DIR / f"{slug}.html"
    if html_path.exists():
        text = html_path.read_text(encoding="utf-8")
        if _NOT_WRITTEN_MARKER not in text:
            lines = text.splitlines()
            if not lines or not lines[0].lower().startswith("subject:"):
                raise WorkbookShapeError(
                    f"{html_path} must start with a 'Subject: ...' line -- see "
                    f"market_update_letters/README.md"
                )
            subject_template = lines[0].split(":", 1)[1].strip()
            greeting_template = "\n".join(lines[1:]).lstrip("\n")

    return Letter(subject_template=subject_template, greeting_template=greeting_template, pdf_path=pdf_path)


def render_pdf_pages(pdf_path: Path) -> List[Tuple[str, bytes, str]]:
    """Rasterize every page of Eddie's PDF letter to a PNG, ready to embed
    inline in the email body via Content-ID -- so the email shows the actual
    designed page itself, not a plain attachment with a short note. Returns
    (content_id, png_bytes, content_type) triplets in page order. Same
    result for every recipient of a suburb, so callers render this once per
    suburb and reuse it, not once per recipient."""
    images: List[Tuple[str, bytes, str]] = []
    doc = fitz.open(pdf_path)
    try:
        for i, page in enumerate(doc):
            pixmap = page.get_pixmap(dpi=_PDF_RENDER_DPI)
            cid = f"letterpage{i}.{pdf_path.stem}@marketupdate"
            images.append((cid, pixmap.tobytes("png"), "image/png"))
    finally:
        doc.close()
    return images


def render_letter(
    letter: Letter, recipient: Recipient, images: List[Tuple[str, bytes, str]]
) -> "tuple[str, str]":
    """Merge-fill the subject and greeting for one recipient, then append
    Eddie's PDF pages (already rasterized once per suburb by
    render_pdf_pages) as inline images referenced by cid:."""
    fields = recipient.merge_fields()
    subject = letter.subject_template
    greeting = letter.greeting_template
    for key, value in fields.items():
        token = "{{%s}}" % key
        subject = subject.replace(token, value)
        greeting = greeting.replace(token, value)

    img_tags = "\n".join(
        f'<img src="cid:{cid}" alt="Market update page {i + 1}" '
        f'style="max-width:100%; display:block; margin:0 0 8px 0;">'
        for i, (cid, _data, _content_type) in enumerate(images)
    )
    body = f"{greeting}\n{img_tags}" if img_tags else greeting
    return subject, body


# ─── Draft creation (mirrors vendor_update_agent.py's new-message drafts) ───

def outlook_create_new_draft(
    outlook: OutlookClient, to_address: str, subject: str, html_body: str,
    images: Optional[List[Tuple[str, bytes, str]]] = None,
) -> Optional[str]:
    payload = {
        "subject": subject,
        "body": {"contentType": "HTML", "content": html_body},
        "toRecipients": [{"emailAddress": {"address": to_address}}],
    }
    if images:
        payload["attachments"] = [
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": f"page{i + 1}.png",
                "contentType": content_type,
                "contentBytes": base64.b64encode(data).decode(),
                "isInline": True,
                "contentId": cid,
            }
            for i, (cid, data, content_type) in enumerate(images)
        ]
    result = outlook._post("/me/messages", payload)
    if result and "id" in result:
        return result["id"]
    logger.error("  Outlook draft creation failed for %s <%s>", subject, to_address)
    return None


def gmail_create_new_draft(
    gmail: GmailClient, to_address: str, subject: str, html_body: str, plain_body: str = "",
    images: Optional[List[Tuple[str, bytes, str]]] = None,
) -> Optional[str]:
    try:
        msg = MIMEMultipart("related")
        msg["To"] = to_address
        msg["Subject"] = subject
        alt = MIMEMultipart("alternative")
        if plain_body:
            alt.attach(MIMEText(plain_body, "plain"))
        alt.attach(MIMEText(html_body, "html"))
        msg.attach(alt)
        for cid, data, content_type in (images or []):
            _maintype, _sep, subtype = content_type.partition("/")
            img = MIMEImage(data, _subtype=subtype or "png")
            img.add_header("Content-ID", f"<{cid}>")
            img.add_header("Content-Disposition", "inline")
            msg.attach(img)
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        result = gmail.service.users().drafts().create(userId="me", body={"message": {"raw": raw}}).execute()
        return result.get("id")
    except Exception as exc:
        logger.error("  Gmail draft failed: %s", exc)
        return None


# ─── Dedupe state across runs ────────────────────────────────────────────────

def load_state(path: Optional[str] = None) -> Dict[str, str]:
    try:
        with open(path or _STATE_PATH) as fh:
            return json.load(fh)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_state(state: Dict[str, str], path: Optional[str] = None) -> None:
    with open(path or _STATE_PATH, "w") as fh:
        json.dump(state, fh, indent=2, sort_keys=True)


# ─── Main ─────────────────────────────────────────────────────────────────────

def run(
    dry_run: bool = True,
    only_segment: Optional[str] = None,
    limit: Optional[int] = None,
    xlsx_path: Optional[str] = None,
    gmail_copies: Optional[bool] = None,
) -> Dict:
    """Read, segment, and draft. dry_run=True (the default) reads the workbook
    and renders every letter but saves and sends nothing, and records no
    state -- safe to run against the real workbook and real Drafts folder at
    any time. Set dry_run=False only once Eddie has confirmed real drafts
    should be created.
    """
    xlsx_path = xlsx_path or DATA_BASE_XLSX_PATH
    want_gmail = GMAIL_COPIES_ENABLED if gmail_copies is None else gmail_copies
    segments = {only_segment: SEGMENTS[only_segment]} if only_segment else SEGMENTS

    logger.info(
        "=== Market Update Agent starting — %s%s ===",
        datetime.now().strftime("%Y-%m-%d %H:%M"),
        " [DRY RUN -- nothing will be saved]" if dry_run else "",
    )
    logger.info("Reading recipients from: %s", xlsx_path)

    outlook = None
    gmail = None
    if not dry_run:
        outlook = OutlookClient(
            client_id=os.getenv("AZURE_CLIENT_ID"),
            tenant_id=os.getenv("AZURE_TENANT_ID"),
            token_cache_path=os.getenv("MSAL_TOKEN_CACHE_PATH", "msal_token_cache.json"),
        )
        if want_gmail:
            try:
                gmail = GmailClient(
                    credentials_path=os.getenv("GMAIL_CREDENTIALS_PATH", "gmail_credentials.json"),
                    token_path=os.getenv("GMAIL_TOKEN_PATH", "gmail_token.json"),
                )
            except GmailAuthRequired as exc:
                logger.error("Gmail unavailable this run (Outlook drafts still created): %s", exc)

    state = {} if dry_run else load_state()
    summary: Dict[str, Dict] = {}

    for segment_key, cfg in segments.items():
        label, sheet = cfg["label"], cfg["sheet"]
        logger.info("--- Segment: %s (sheet %r) ---", label, sheet)

        try:
            recipients = read_recipients(xlsx_path, segment_key, sheet, label)
        except WorkbookShapeError as exc:
            logger.error("Skipping segment %s: %s", label, exc)
            summary[segment_key] = {"error": str(exc)}
            continue

        recipients, extra_properties = dedupe_by_owner_and_suburb(recipients)
        recipients = sorted(recipients, key=lambda r: (r.suburb.lower(), r.email))
        if limit is not None:
            recipients = recipients[:limit]

        suburbs_all = sorted({r.suburb for r in recipients if r.suburb})
        logger.info(
            "  %d owner-in-suburb draft(s) across %d suburb(s) (%d owner+suburb pair(s) "
            "had more than one property in that same suburb, collapsed to one draft each; "
            "an owner with properties in several suburbs still gets one draft per suburb)",
            len(recipients), len(suburbs_all), len(extra_properties),
        )

        by_suburb: Dict[str, List[Recipient]] = {}
        for r in recipients:
            if r.suburb:
                by_suburb.setdefault(r.suburb, []).append(r)

        drafted = 0
        skipped_state = 0
        skipped_no_letter = 0
        suburbs_drafted: List[str] = []
        suburbs_skipped_no_letter: List[str] = []
        previews: List[Dict] = []

        for suburb in sorted(by_suburb):
            sub_recipients = by_suburb[suburb]
            letter = load_letter(suburb)
            if letter is None:
                suburbs_skipped_no_letter.append(suburb)
                skipped_no_letter += len(sub_recipients)
                continue
            suburbs_drafted.append(suburb)
            images = render_pdf_pages(letter.pdf_path)  # once per suburb -- identical for every owner in it

            for r in sub_recipients:
                # Suburb is part of the key: dedupe now keeps one recipient per
                # (owner, suburb), so an owner in two suburbs needs two
                # independent "already drafted" records, not one shared by both.
                state_key = f"{segment_key}:{r.email}:{suburb_slug(r.suburb)}"
                if not dry_run and state_key in state:
                    skipped_state += 1
                    continue

                subject, html_body = render_letter(letter, r, images)
                plain_body = ""  # real content is the embedded PDF pages; no separate plain-text source

                if dry_run:
                    previews.append({
                        "to": r.email, "suburb": r.suburb, "subject": subject,
                        "row": r.source_row, "pages": len(images),
                    })
                    drafted += 1
                    continue

                outlook_id = outlook_create_new_draft(outlook, r.email, subject, html_body, images)
                gmail_id = gmail_create_new_draft(gmail, r.email, subject, html_body, plain_body, images) if gmail else None
                if outlook_id or gmail_id:
                    drafted += 1
                    state[state_key] = datetime.now().strftime("%Y-%m-%d")
                else:
                    logger.warning("  Draft failed for %s (row %d)", r.email, r.source_row)

        if suburbs_skipped_no_letter:
            logger.warning(
                "  No letter supplied this week for %d suburb(s) in %s -- skipped, drafting "
                "nothing for them (market_update_letters/<suburb>.pdf not found): %s",
                len(suburbs_skipped_no_letter), label, ", ".join(suburbs_skipped_no_letter),
            )

        if not dry_run:
            save_state(state)

        summary[segment_key] = {
            "recipients": len(recipients),
            "suburbs_total": len(suburbs_all),
            "suburbs_drafted": suburbs_drafted,
            "suburbs_skipped_no_letter": suburbs_skipped_no_letter,
            "drafted": drafted,
            "skipped_already_drafted": skipped_state,
            "skipped_no_letter": skipped_no_letter,
            "previews": previews,
        }
        logger.info(
            "  %s %d draft(s) across %d suburb(s)%s", "Would create" if dry_run else "Created",
            drafted, len(suburbs_drafted),
            f", {skipped_state} already drafted in a previous run" if skipped_state else "",
        )

    missing = [k for k in ("northern_beaches", "city_fringe") if k not in SEGMENTS]
    if missing:
        logger.warning(
            "Segments not run (no source data in DATA BASE.xlsx yet): %s -- see status report",
            missing,
        )

    logger.info("=== Market Update Agent done%s ===", " (DRY RUN)" if dry_run else "")
    return summary


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="IB Property market-update mail-merge agent")
    parser.add_argument("--dry-run", action="store_true", help="Preview only -- reads the workbook and renders letters, saves/sends nothing")
    parser.add_argument("--segment", choices=list(SEGMENTS), default=None, help="Only process this segment")
    parser.add_argument("--limit", type=int, default=None, help="Only process the first N recipients per segment (for testing)")
    args = parser.parse_args()

    if not args.dry_run:
        print(
            "Refusing to run for real: this agent hasn't been approved to create live "
            "drafts yet. Use --dry-run, or edit market_update_agent.py once Eddie has "
            "confirmed the segment mapping and at least one letter is written.",
            file=sys.stderr,
        )
        sys.exit(1)

    result = run(dry_run=True, only_segment=args.segment, limit=args.limit)
    print(json.dumps(
        {k: {kk: vv for kk, vv in v.items() if kk != "previews"} for k, v in result.items()},
        indent=2,
    ))
