"""Persistent per-campaign continuity ledger for vendor_update_agent.py.

Eagle gives authoritative numbers (enquiries/inspections/offers/days on
market) and email-mining gives this week's raw activity, but neither
carries forward WHO was already being followed up on and WHY -- which is
what makes Eddie's real updates read as a continuing conversation ("As
mentioned last week, this group remains interested") instead of a fresh
summary every time. This is that missing piece: a small JSON file, one
entry per campaign, holding the known leads and their last-known status.
Each week's draft reads the entry, folds in whatever's new, and writes the
result back -- so state accumulates instead of resetting.

Seeded on first sight of a campaign by backfilling from Eddie's own
already-sent weekly update emails (see vendor_update_agent.backfill_ledger_entry),
per his explicit instruction not to start blank for a campaign with real
history.
"""

import json
import os
from dataclasses import asdict, dataclass, field
from typing import Dict, List, Optional

_LEDGER_PATH = os.getenv("VENDOR_UPDATE_LEDGER_PATH", "vendor_update_ledger.json")


@dataclass
class Lead:
    name: str
    status: str
    last_updated: str = ""  # YYYY-MM-DD, the date this status was last confirmed


@dataclass
class LedgerEntry:
    address: str
    week_number: int = 0
    known_leads: List[Lead] = field(default_factory=list)
    last_drafted_at: str = ""  # ISO8601

    def to_dict(self) -> Dict:
        return {
            "address": self.address,
            "week_number": self.week_number,
            "known_leads": [asdict(lead) for lead in self.known_leads],
            "last_drafted_at": self.last_drafted_at,
        }

    @staticmethod
    def from_dict(d: Dict) -> "LedgerEntry":
        return LedgerEntry(
            address=d.get("address", ""),
            week_number=d.get("week_number", 0),
            known_leads=[Lead(**lead) for lead in d.get("known_leads", [])],
            last_drafted_at=d.get("last_drafted_at", ""),
        )


def load_ledger(path: Optional[str] = None) -> Dict[str, LedgerEntry]:
    """Keyed by campaign property_id (the CRM's own stable id -- not
    address, which can be formatted slightly differently run to run).
    Returns an empty ledger if the file doesn't exist yet -- every campaign
    looks "new" on the very first run, which is correct."""
    try:
        with open(path or _LEDGER_PATH) as fh:
            raw = json.load(fh)
        return {k: LedgerEntry.from_dict(v) for k, v in raw.items()}
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_ledger(ledger: Dict[str, LedgerEntry], path: Optional[str] = None) -> None:
    with open(path or _LEDGER_PATH, "w") as fh:
        json.dump({k: v.to_dict() for k, v in ledger.items()}, fh, indent=2, sort_keys=True)
