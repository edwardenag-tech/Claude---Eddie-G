"""Parse the "Campaign Enquiry Reply Templates" Google Doc into structured
campaign dicts.

The doc's ACTIVE CAMPAIGNS section holds one or more blocks delimited by a
line containing only "===", each following the TEMPLATE BLOCK layout:

    ===
    Property: [address]
    Status (from Campaign Checklist): [...]
    Positioning line: [...]
    Key facts:
    - [fact]
    - [fact]
    Price guide: [...]
    IM link: [...]
    Inspections: [...]
    Special notes for the agent: [...]
    ===

This parser is intentionally forgiving about exact whitespace/formatting
since the doc is hand-edited by Eddie, not machine-generated. Any block
missing a "Property:" line is skipped (can't match it to an enquiry or
listing without an address).
"""

import re
from typing import Dict, List

_FIELD_PATTERNS = {
    "address":         re.compile(r"^\s*Property\s*:\s*(.+)$", re.IGNORECASE),
    "status":          re.compile(r"^\s*Status\s*\(from Campaign Checklist\)\s*:\s*(.+)$", re.IGNORECASE),
    "positioning":     re.compile(r"^\s*Positioning line\s*:\s*(.+)$", re.IGNORECASE),
    "price_guide":     re.compile(r"^\s*Price guide\s*:\s*(.+)$", re.IGNORECASE),
    "im_link":         re.compile(r"^\s*IM link\s*:\s*(.+)$", re.IGNORECASE),
    "inspections":     re.compile(r"^\s*Inspections\s*:\s*(.+)$", re.IGNORECASE),
    "special_notes":   re.compile(r"^\s*Special notes for the agent\s*:\s*(.+)$", re.IGNORECASE),
}
_KEY_FACTS_HEADER = re.compile(r"^\s*Key facts\s*:\s*$", re.IGNORECASE)
_BLOCK_DELIM = re.compile(r"^\s*={3,}\s*$")
_ACTIVE_CAMPAIGNS_HEADER = re.compile(r"^\s*ACTIVE CAMPAIGNS\s*$", re.IGNORECASE)
_BULLET_PREFIX = re.compile(r"^\s*[-*•]\s*")

# Free-text markers Eddie uses in the doc to flag a field as needing his
# attention (e.g. "Status: NEEDS EDDIE TO CONFIRM inspection time").
_ATTENTION_MARKERS = re.compile(
    r"needs?\s+eddie|eddie\s+to\s+(confirm|review|check|fill|update)|attention needed",
    re.IGNORECASE,
)


def _clean_placeholder(value: str) -> str:
    """Strip [square-bracket placeholder] wrapping so unfilled fields read as empty."""
    v = value.strip()
    if v.startswith("[") and v.endswith("]"):
        return ""
    return v


def _parse_block(lines: List[str]) -> Dict:
    campaign: Dict = {
        "address": "", "status": "", "positioning": "", "price_guide": "",
        "im_link": "", "inspections": "", "special_notes": "", "key_facts": [],
    }
    in_key_facts = False

    for raw_line in lines:
        line = raw_line.rstrip()
        if not line.strip():
            continue

        if _KEY_FACTS_HEADER.match(line):
            in_key_facts = True
            continue

        matched_new_field = False
        for field, pattern in _FIELD_PATTERNS.items():
            m = pattern.match(line)
            if m:
                campaign[field] = _clean_placeholder(m.group(1))
                in_key_facts = False
                matched_new_field = True
                break
        if matched_new_field:
            continue

        if in_key_facts:
            fact = _BULLET_PREFIX.sub("", line).strip()
            fact = _clean_placeholder(fact)
            if fact:
                campaign["key_facts"].append(fact)

    return campaign


def parse_active_campaigns(doc_text: str) -> List[Dict]:
    """Return a list of campaign dicts parsed from the doc's ACTIVE CAMPAIGNS section.

    Each dict has: address, status, positioning, key_facts (list), price_guide,
    im_link, inspections, special_notes. Blocks without a non-empty "address"
    are dropped.
    """
    lines = doc_text.splitlines()

    # Find the start of the ACTIVE CAMPAIGNS section; if not found, scan the
    # whole doc (still delimiter-bounded, so the template block itself --
    # which has placeholder [brackets] that get stripped to "" -- won't
    # produce a false-positive campaign since it'll have no address).
    start_idx = 0
    for i, line in enumerate(lines):
        if _ACTIVE_CAMPAIGNS_HEADER.match(line):
            start_idx = i + 1
            break

    section_lines = lines[start_idx:]

    blocks: List[List[str]] = []
    current: List[str] = []
    inside = False
    for line in section_lines:
        if _BLOCK_DELIM.match(line):
            if inside:
                blocks.append(current)
                current = []
                inside = False
            else:
                inside = True
                current = []
            continue
        if inside:
            current.append(line)
    # Tolerate a trailing block with no closing "===" (e.g. doc still being edited)
    if inside and current:
        blocks.append(current)

    campaigns = []
    for block_lines in blocks:
        campaign = _parse_block(block_lines)
        if campaign["address"]:
            campaigns.append(campaign)

    return campaigns


def summarize_campaign_highlights(campaigns: List[Dict]) -> List[Dict]:
    """Produce a short per-campaign highlight for the morning briefing.

    Each returned dict has: address, status, needs_attention (bool), and
    attention_note (the field text that triggered it, or "" if none).
    Not a full doc dump -- just enough to scan in a briefing email.
    """
    highlights = []
    for campaign in campaigns:
        address = campaign.get("address", "")
        if not address:
            continue

        attention_note = ""
        for field in ("status", "positioning", "special_notes"):
            value = campaign.get(field, "")
            if value and _ATTENTION_MARKERS.search(value):
                attention_note = value.strip()
                break
        if not attention_note:
            for fact in campaign.get("key_facts", []):
                if _ATTENTION_MARKERS.search(fact):
                    attention_note = fact.strip()
                    break

        highlights.append({
            "address": address,
            "status": (campaign.get("status", "") or "(status not set)").strip(),
            "needs_attention": bool(attention_note),
            "attention_note": attention_note,
        })

    return highlights
