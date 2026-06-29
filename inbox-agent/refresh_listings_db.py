"""Scrape vendor update emails from Sent Items and rebuild listings_db.json.

Searches Eddie's Sent Items for "vendor update" and "campaign update" emails,
extracts structured property data via Claude, and writes listings_db.json.

Usage:
    python refresh_listings_db.py
"""

import html as _html
import json
import logging
import os
import re
import sys
from datetime import date
from typing import Dict, List, Optional

from dotenv import load_dotenv
import anthropic

from outlook_client import OutlookClient

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

LISTINGS_DB_PATH = os.path.join(os.path.dirname(__file__), "listings_db.json")


def fetch_vendor_update_emails(outlook: OutlookClient) -> List[Dict]:
    """Search Sent Items for vendor/campaign update emails."""
    seen_ids: set = set()
    results: List[Dict] = []

    for query in ["vendor update", "campaign update"]:
        params = {
            "$search": f'"{query}"',
            "$top": 20,
            "$select": "id,subject,body,sentDateTime,toRecipients",
        }
        resp = outlook._get("/me/mailFolders/sentitems/messages", params=params)
        for msg in (resp.get("value", []) if resp else []):
            if msg["id"] not in seen_ids:
                seen_ids.add(msg["id"])
                results.append(msg)

    return results


def _body_to_plain(msg: Dict) -> str:
    """Convert Graph API message body to plain text."""
    raw = msg.get("body", {}).get("content", "") if isinstance(msg.get("body"), dict) else ""
    decoded = _html.unescape(raw).replace("\xa0", " ")
    plain = re.sub(r"<[^>]+>", " ", decoded)
    return re.sub(r"\s{2,}", " ", plain).strip()


def extract_listings_with_claude(
    ai: anthropic.Anthropic, subject: str, body: str
) -> List[Dict]:
    """Use Claude to extract structured property listings from a vendor update email."""
    today = date.today().isoformat()

    prompt = (
        "You are extracting commercial property listing data from a vendor update email "
        "sent by a Sydney commercial real estate agent (IB Property).\n\n"
        f"Email Subject: {subject}\n"
        f"Email Body:\n{body[:4000]}\n\n"
        "Extract ALL properties mentioned in this email. For each property return:\n"
        "- address: full street address including suburb and state\n"
        "- type: \"sale\" or \"lease\"\n"
        "- price: asking price string (for sale), e.g. \"$3,275,000\", or null\n"
        "- asking_rent: asking rent per annum (for lease), e.g. \"$55,000 p.a. gross + GST\", or null\n"
        "- net_rent: current net rent (for sale investments), e.g. \"$181,495.96 + GST\", or null\n"
        "- fully_leased_rent: estimated fully leased net rent, or null\n"
        "- floor_area: internal floor area, e.g. \"205 sqm\", or null\n"
        "- land_area: land/site area, e.g. \"221 sqm\", or null\n"
        "- building: building name if the property is part of a named complex, or null\n"
        "- im_url: URL to the information memorandum (heyzine.com or similar flipbook), or null\n"
        "- lease_terms: description of current lease, options, and expiry date, or null\n"
        "- parking: parking/car spaces description, e.g. \"2 cars\", or null\n"
        "- notes: any other notable info (vacancy, value-add opportunity, zoning, etc.), or null\n\n"
        "Return ONLY a valid JSON array of property objects. If no properties are found, return [].\n\n"
        "Example output:\n"
        "[\n"
        "  {\n"
        "    \"address\": \"14A Hannah Street, Beecroft NSW 2119\",\n"
        "    \"type\": \"sale\",\n"
        "    \"price\": \"$3,275,000\",\n"
        "    \"asking_rent\": null,\n"
        "    \"net_rent\": \"$181,495.96 + GST\",\n"
        "    \"fully_leased_rent\": \"$226,495.00 + GST\",\n"
        "    \"floor_area\": \"205 sqm\",\n"
        "    \"land_area\": \"221 sqm\",\n"
        "    \"building\": null,\n"
        "    \"im_url\": \"https://heyzine.com/flip-book/97160c3d71.html\",\n"
        "    \"lease_terms\": \"5-year variation since Nov 2025, 2x5yr options\",\n"
        "    \"parking\": \"2 cars\",\n"
        "    \"notes\": \"Level 1 vacant - value-add potential\"\n"
        "  }\n"
        "]"
    )

    try:
        response = ai.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=2000,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = response.content[0].text.strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw, flags=re.IGNORECASE)
        raw = re.sub(r"\s*```$", "", raw)
        if not raw:
            return []
        listings = json.loads(raw)
        for listing in listings:
            listing["last_updated"] = today
        return listings
    except Exception as exc:
        logger.error("Claude extraction failed: %s", exc)
        return []


def _normalize_address(address: str) -> str:
    """Lowercase and collapse whitespace for deduplication keying."""
    return re.sub(r"\s+", " ", address.lower().strip())


def deduplicate_listings(listings: List[Dict]) -> List[Dict]:
    """Keep the most recently updated entry for each unique address."""
    seen: Dict[str, Dict] = {}
    for listing in listings:
        key = _normalize_address(listing.get("address", ""))
        if key and (
            key not in seen
            or listing.get("last_updated", "") > seen[key].get("last_updated", "")
        ):
            seen[key] = listing
    return list(seen.values())


def main() -> None:
    missing = [k for k in ("ANTHROPIC_API_KEY", "AZURE_CLIENT_ID", "AZURE_TENANT_ID") if not os.getenv(k)]
    if missing:
        logger.error("Missing required env vars: %s", ", ".join(missing))
        sys.exit(1)

    msal_cache = os.getenv("MSAL_TOKEN_CACHE_PATH", "msal_token_cache.bin")

    logger.info("Connecting to Outlook (Microsoft Graph)...")
    outlook = OutlookClient(
        client_id=os.getenv("AZURE_CLIENT_ID"),
        tenant_id=os.getenv("AZURE_TENANT_ID"),
        token_cache_path=msal_cache,
    )
    ai = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

    logger.info("Searching Sent Items for vendor/campaign update emails...")
    vendor_emails = fetch_vendor_update_emails(outlook)

    if not vendor_emails:
        logger.warning(
            'No vendor update emails found in Sent Items.\n'
            'Try searching manually for keywords like "vendor update" or "campaign update".'
        )
        sys.exit(0)

    print(f"\nFound {len(vendor_emails)} vendor update email(s) in Sent Items:\n")
    for msg in vendor_emails:
        sent = msg.get("sentDateTime", "unknown")[:10]
        subj = msg.get("subject", "(no subject)")
        print(f"  {sent} | {subj}")
    print()

    all_listings: List[Dict] = []
    for msg in vendor_emails:
        subject = msg.get("subject", "")
        body = _body_to_plain(msg)
        sent_date = msg.get("sentDateTime", "")[:10]
        logger.info("Extracting properties from: %s (%s)", subject, sent_date)
        listings = extract_listings_with_claude(ai, subject, body)
        logger.info("  → %d propert%s found", len(listings), "y" if len(listings) == 1 else "ies")
        all_listings.extend(listings)

    if not all_listings:
        logger.warning(
            "No properties could be extracted from any vendor update email. "
            "Check that the emails contain property listing details."
        )
        sys.exit(0)

    deduped = deduplicate_listings(all_listings)
    db = {"listings": deduped}

    with open(LISTINGS_DB_PATH, "w") as fh:
        json.dump(db, fh, indent=2)

    print(f"Written {len(deduped)} listing(s) to {LISTINGS_DB_PATH}:\n")
    for listing in deduped:
        listing_type = listing.get("type", "?").upper()
        address = listing.get("address", "?")
        print(f"  [{listing_type}] {address}")
        if listing.get("price"):
            print(f"         Price:  {listing['price']}")
        if listing.get("asking_rent"):
            print(f"          Rent:  {listing['asking_rent']}")
        if listing.get("net_rent"):
            print(f"      Net rent:  {listing['net_rent']}")
        if listing.get("floor_area"):
            print(f"    Floor area:  {listing['floor_area']}")
        if listing.get("im_url"):
            print(f"            IM:  {listing['im_url']}")
        print()

    logger.info("listings_db.json updated successfully.")


if __name__ == "__main__":
    main()
