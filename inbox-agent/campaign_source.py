"""Pluggable "what's actually on market right now" data source for
vendor_update_agent.py.

Replaces the old Campaign Enquiry Reply Templates Google Doc entirely --
Eddie does not want to hand-maintain a status doc; that was the whole
complaint that started this rework. The real source of truth is his CRM
(EagleAgent, see eagle_client.py). This module defines that dependency as a
small Protocol so a different CRM could be swapped in later without
touching vendor_update_agent.py's own logic, and so tests can use a fake
without hitting a real API.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Protocol

from eagle_client import EagleClient


@dataclass
class CampaignContact:
    first_name: str = ""
    last_name: str = ""
    email: str = ""
    phone: str = ""
    company: str = ""

    @property
    def full_name(self) -> str:
        name = f"{self.first_name} {self.last_name}".strip()
        return name or self.company


@dataclass
class Campaign:
    """One active listing, as reported by the CRM -- ground truth for what
    Eddie has on market right now. Never inferred from email, never
    hand-entered by Eddie."""
    property_id: str  # the CRM's own internal ID
    address: str  # formatted address
    rea_id: Optional[str] = None
    # Eagle's own field, kept for reference/display only -- NOT a reliable
    # join key to a portal-enquiry email's "Property ID: NNNNNN". Verified
    # live: Suite C/154-156 Sailors Bay Road, Northbridge is a real, current,
    # ACTIVE listing whose actual realcommercial.com.au Property ID (505195504,
    # confirmed from a real enquiry email) does not resolve at all via
    # `property(reaId: "505195504")` ("Property Not Found") -- Eagle stores
    # "eagle_1817949" (its own internal id, prefixed) for that exact listing
    # instead. Other properties in the same account carry plain numeric or
    # "1P"-prefixed values in this field that don't match realcommercial's
    # current 9-digit format either, so they're not a substitute source of a
    # real portal ID. Match a portal-enquiry email to a Campaign by address
    # instead (street number + street name, same approach as
    # draft_agent.py's _match_by_address) -- structured, reliable, and
    # Eagle's `street`/`streetNo`/`unit`/`postcode` fields (not yet pulled
    # into this query) support it precisely if formattedAddress alone proves
    # too loose in practice.
    sale_or_lease: str = ""  # "SALE" | "LEASE" | "SALE_AND_LEASE", from the CRM's own field
    active_at: Optional[str] = None  # ISO8601 -- when the listing went live
    days_on_market: Optional[int] = None
    num_enquiries: Optional[int] = None
    num_inspection_attendances: Optional[int] = None
    num_offers: Optional[int] = None
    landlord_contacts: List[CampaignContact] = field(default_factory=list)
    # Secondary/fallback only -- confirmed live that most properties (4 of 5
    # in an initial ACTIVE sample) come back with an empty vendors list. The
    # primary landlord-contact source stays vendor_update_agent.py's existing
    # find_landlord_email/find_landlord_first_name (Sent Items search, already
    # working -- it correctly found Sandra Odorisio for Sailors Bay Road).
    # Only fall back to this field when that search comes up empty.


class CampaignSource(Protocol):
    """What vendor_update_agent.py needs from wherever "active campaigns"
    actually lives. Read-only by design -- nothing implementing this may
    write back to the CRM."""

    def get_active_campaigns(self) -> List[Campaign]:
        ...


class EagleCampaignSource:
    """Real implementation, backed by EagleAgent's `properties(status: [ACTIVE])`
    query. Every field referenced below was confirmed against the live
    GraphQL schema at api.eaglesoftware.com.au/v3 (Property, Vendor, Contact,
    PriorityEmail, PriorityPhoneNumber, PageInfo) before being written here --
    none of this is guessed."""

    _QUERY = """
    query ActiveCampaigns($after: String) {
      properties(status: [ACTIVE], first: 50, after: $after) {
        nodes {
          id
          formattedAddress
          reaId
          saleOrLease
          activeAt
          daysOnMarket
          numEnquiries
          numInspectionAttendances
          numOffers
          agents { id name email }
          vendors {
            contact {
              firstName
              lastName
              company
              emails { email }
              phoneNumbers { phoneNumber }
            }
          }
        }
        pageInfo { hasNextPage endCursor }
      }
    }
    """

    def __init__(self, client: EagleClient, agent_email: str, agent_name: Optional[str] = None):
        """agent_email is the primary filter -- there's no server-side
        `agentId`/`agentEmail` argument on the `properties` query (checked
        the full arg list against the live schema; it isn't there), so this
        fetches every ACTIVE property agency-wide and filters client-side on
        `Property.agents`, confirmed live: 17 of 252 for
        edward@ibproperty.com.au, including both addresses Eddie's mentioned
        this session (490 Pacific Highway, Sailors Bay Road). agent_name is
        an exact-match (case-insensitive, whitespace-trimmed -- Eagle data
        has at least one agent name with a stray trailing space) fallback
        for a property whose agent record has no email populated; email is
        preferred since names can vary (middle names, nicknames)."""
        self.client = client
        self.agent_email = agent_email.strip().lower()
        self.agent_name = agent_name.strip().lower() if agent_name else None

    def get_active_campaigns(self) -> List[Campaign]:
        campaigns: List[Campaign] = []
        after = None
        while True:
            data = self.client.graphql(self._QUERY, {"after": after})
            connection = data.get("properties") or {}
            for node in connection.get("nodes") or []:
                if self._is_agents_match(node.get("agents") or []):
                    campaigns.append(self._to_campaign(node))
            page_info = connection.get("pageInfo") or {}
            if not page_info.get("hasNextPage"):
                break
            after = page_info.get("endCursor")
        return campaigns

    def _is_agents_match(self, agents: List[Dict]) -> bool:
        for a in agents:
            email = (a.get("email") or "").strip().lower()
            if email and email == self.agent_email:
                return True
        if self.agent_name:
            for a in agents:
                name = (a.get("name") or "").strip().lower()
                if name and name == self.agent_name:
                    return True
        return False

    @staticmethod
    def _to_campaign(node: Dict) -> Campaign:
        contacts: List[CampaignContact] = []
        for vendor in node.get("vendors") or []:
            c = vendor.get("contact") or {}
            emails = c.get("emails") or []
            phones = c.get("phoneNumbers") or []
            contacts.append(CampaignContact(
                first_name=c.get("firstName") or "",
                last_name=c.get("lastName") or "",
                email=(emails[0].get("email") or "") if emails else "",
                phone=(phones[0].get("phoneNumber") or "") if phones else "",
                company=c.get("company") or "",
            ))
        return Campaign(
            property_id=node["id"],
            address=node.get("formattedAddress") or "",
            rea_id=node.get("reaId"),
            sale_or_lease=node.get("saleOrLease") or "",
            active_at=node.get("activeAt"),
            days_on_market=node.get("daysOnMarket"),
            num_enquiries=node.get("numEnquiries"),
            num_inspection_attendances=node.get("numInspectionAttendances"),
            num_offers=node.get("numOffers"),
            landlord_contacts=contacts,
        )
