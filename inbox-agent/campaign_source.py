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
    rea_id: Optional[str] = None  # realcommercial.com.au / commercialrealestate.com.au
    # portal Property ID, when the listing is pushed there -- this is the
    # exact join key for matching a portal-enquiry email to this campaign,
    # see draft_agent.py's _extract_address for the "Property ID: NNNN" /
    # "Contacted:" subject patterns those emails use.
    sale_or_lease: str = ""  # "SALE" | "LEASE", from the CRM's own field
    active_at: Optional[str] = None  # ISO8601 -- when the listing went live
    days_on_market: Optional[int] = None
    num_enquiries: Optional[int] = None
    num_inspection_attendances: Optional[int] = None
    num_offers: Optional[int] = None
    landlord_contacts: List[CampaignContact] = field(default_factory=list)


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

    def __init__(self, client: EagleClient):
        self.client = client

    def get_active_campaigns(self) -> List[Campaign]:
        campaigns: List[Campaign] = []
        after = None
        while True:
            data = self.client.graphql(self._QUERY, {"after": after})
            connection = data.get("properties") or {}
            for node in connection.get("nodes") or []:
                campaigns.append(self._to_campaign(node))
            page_info = connection.get("pageInfo") or {}
            if not page_info.get("hasNextPage"):
                break
            after = page_info.get("endCursor")
        return campaigns

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
