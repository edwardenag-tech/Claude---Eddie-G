"""Tests for campaign_source.py: EagleCampaignSource's mapping from the raw
GraphQL response shape (confirmed against the live schema, see
campaign_source.py's docstring) into Campaign/CampaignContact. The Eagle
client itself is mocked -- these tests only exercise the mapping and
pagination logic, not the transport (see test_eagle_client.py for that).
"""

import os
import sys
import unittest
from unittest.mock import MagicMock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import campaign_source as m


def _node(**overrides):
    node = {
        "id": "123",
        "formattedAddress": "490 Pacific Highway, St Leonards NSW 2065",
        "reaId": "505213104",
        "saleOrLease": "LEASE",
        "activeAt": "2026-08-01T00:00:00Z",
        "daysOnMarket": 49,
        "numEnquiries": 5,
        "numInspectionAttendances": 4,
        "numOffers": 0,
        "vendors": [{"contact": {
            "firstName": "Sandra", "lastName": "Odorisio", "company": "",
            "emails": [{"email": "sandra@example.com"}],
            "phoneNumbers": [{"phoneNumber": "0418698187"}],
        }}],
    }
    node.update(overrides)
    return node


class TestEagleCampaignSourceMapping(unittest.TestCase):
    def _client_returning(self, *pages):
        client = MagicMock()
        client.graphql.side_effect = pages
        return client

    def test_maps_a_single_page(self):
        client = self._client_returning(
            {"properties": {"nodes": [_node()], "pageInfo": {"hasNextPage": False, "endCursor": None}}},
        )
        campaigns = m.EagleCampaignSource(client).get_active_campaigns()
        self.assertEqual(len(campaigns), 1)
        c = campaigns[0]
        self.assertEqual(c.property_id, "123")
        self.assertEqual(c.address, "490 Pacific Highway, St Leonards NSW 2065")
        self.assertEqual(c.rea_id, "505213104")
        self.assertEqual(c.sale_or_lease, "LEASE")
        self.assertEqual(c.days_on_market, 49)
        self.assertEqual(c.num_enquiries, 5)
        self.assertEqual(c.num_inspection_attendances, 4)
        self.assertEqual(c.num_offers, 0)
        self.assertEqual(len(c.landlord_contacts), 1)
        contact = c.landlord_contacts[0]
        self.assertEqual(contact.first_name, "Sandra")
        self.assertEqual(contact.email, "sandra@example.com")
        self.assertEqual(contact.phone, "0418698187")
        self.assertEqual(contact.full_name, "Sandra Odorisio")

    def test_paginates_until_hasNextPage_false(self):
        client = self._client_returning(
            {"properties": {"nodes": [_node(id="1")], "pageInfo": {"hasNextPage": True, "endCursor": "cursor1"}}},
            {"properties": {"nodes": [_node(id="2")], "pageInfo": {"hasNextPage": False, "endCursor": None}}},
        )
        campaigns = m.EagleCampaignSource(client).get_active_campaigns()
        self.assertEqual([c.property_id for c in campaigns], ["1", "2"])
        self.assertEqual(client.graphql.call_count, 2)
        # second call must carry the cursor from the first page's pageInfo
        second_call_variables = client.graphql.call_args_list[1].args[1]
        self.assertEqual(second_call_variables, {"after": "cursor1"})

    def test_contact_with_no_email_or_phone_yields_blank_not_error(self):
        node = _node(vendors=[{"contact": {
            "firstName": "Bob", "lastName": "Buyer", "company": "",
            "emails": [], "phoneNumbers": [],
        }}])
        client = self._client_returning(
            {"properties": {"nodes": [node], "pageInfo": {"hasNextPage": False, "endCursor": None}}},
        )
        campaigns = m.EagleCampaignSource(client).get_active_campaigns()
        contact = campaigns[0].landlord_contacts[0]
        self.assertEqual(contact.email, "")
        self.assertEqual(contact.phone, "")

    def test_no_vendors_yields_empty_contact_list(self):
        node = _node(vendors=[])
        client = self._client_returning(
            {"properties": {"nodes": [node], "pageInfo": {"hasNextPage": False, "endCursor": None}}},
        )
        campaigns = m.EagleCampaignSource(client).get_active_campaigns()
        self.assertEqual(campaigns[0].landlord_contacts, [])

    def test_company_only_contact_falls_back_full_name_to_company(self):
        node = _node(vendors=[{"contact": {
            "firstName": "", "lastName": "", "company": "Acme Pty Ltd",
            "emails": [], "phoneNumbers": [],
        }}])
        client = self._client_returning(
            {"properties": {"nodes": [node], "pageInfo": {"hasNextPage": False, "endCursor": None}}},
        )
        campaigns = m.EagleCampaignSource(client).get_active_campaigns()
        self.assertEqual(campaigns[0].landlord_contacts[0].full_name, "Acme Pty Ltd")

    def test_no_active_campaigns_returns_empty_list(self):
        client = self._client_returning(
            {"properties": {"nodes": [], "pageInfo": {"hasNextPage": False, "endCursor": None}}},
        )
        self.assertEqual(m.EagleCampaignSource(client).get_active_campaigns(), [])


if __name__ == "__main__":
    unittest.main()
