"""Tests for campaign_source.py: EagleCampaignSource's mapping from the raw
GraphQL response shape (confirmed against the live schema, see
campaign_source.py's docstring) into Campaign/CampaignContact, and its
agent-email filtering (confirmed live: 17 of 252 active properties for
edward@ibproperty.com.au). The Eagle client itself is mocked -- these tests
only exercise the mapping/filtering/pagination logic, not the transport
(see test_eagle_client.py for that).
"""

import os
import sys
import unittest
from unittest.mock import MagicMock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import campaign_source as m

EDDIE_EMAIL = "edward@ibproperty.com.au"


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
        "agents": [{"id": "1", "name": "Edward Ghattas", "email": EDDIE_EMAIL}],
        "vendors": [{"contact": {
            "firstName": "Sandra", "lastName": "Odorisio", "company": "",
            "emails": [{"email": "sandra@example.com"}],
            "phoneNumbers": [{"phoneNumber": "0418698187"}],
        }}],
    }
    node.update(overrides)
    return node


def _page(*nodes, has_next=False, cursor=None):
    return {"properties": {"nodes": list(nodes), "pageInfo": {"hasNextPage": has_next, "endCursor": cursor}}}


class TestEagleCampaignSourceMapping(unittest.TestCase):
    def _client_returning(self, *pages):
        client = MagicMock()
        client.graphql.side_effect = pages
        return client

    def _source(self, client, agent_name=None):
        return m.EagleCampaignSource(client, agent_email=EDDIE_EMAIL, agent_name=agent_name)

    def test_maps_a_single_page(self):
        client = self._client_returning(_page(_node()))
        campaigns = self._source(client).get_active_campaigns()
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
            _page(_node(id="1"), has_next=True, cursor="cursor1"),
            _page(_node(id="2")),
        )
        campaigns = self._source(client).get_active_campaigns()
        self.assertEqual([c.property_id for c in campaigns], ["1", "2"])
        self.assertEqual(client.graphql.call_count, 2)
        second_call_variables = client.graphql.call_args_list[1].args[1]
        self.assertEqual(second_call_variables, {"after": "cursor1"})

    def test_contact_with_no_email_or_phone_yields_blank_not_error(self):
        node = _node(vendors=[{"contact": {
            "firstName": "Bob", "lastName": "Buyer", "company": "",
            "emails": [], "phoneNumbers": [],
        }}])
        client = self._client_returning(_page(node))
        campaigns = self._source(client).get_active_campaigns()
        contact = campaigns[0].landlord_contacts[0]
        self.assertEqual(contact.email, "")
        self.assertEqual(contact.phone, "")

    def test_no_vendors_yields_empty_contact_list(self):
        client = self._client_returning(_page(_node(vendors=[])))
        campaigns = self._source(client).get_active_campaigns()
        self.assertEqual(campaigns[0].landlord_contacts, [])

    def test_company_only_contact_falls_back_full_name_to_company(self):
        node = _node(vendors=[{"contact": {
            "firstName": "", "lastName": "", "company": "Acme Pty Ltd",
            "emails": [], "phoneNumbers": [],
        }}])
        client = self._client_returning(_page(node))
        campaigns = self._source(client).get_active_campaigns()
        self.assertEqual(campaigns[0].landlord_contacts[0].full_name, "Acme Pty Ltd")

    def test_no_active_campaigns_returns_empty_list(self):
        client = self._client_returning(_page())
        self.assertEqual(self._source(client).get_active_campaigns(), [])


class TestAgentFiltering(unittest.TestCase):
    """Property.agents-based filtering -- there's no server-side agent filter
    on the properties() query (checked the full arg list), so this is done
    client-side, confirmed against the real 252-property/17-Eddie dataset."""

    def _client_returning(self, *pages):
        client = MagicMock()
        client.graphql.side_effect = pages
        return client

    def test_matches_by_email(self):
        node = _node(agents=[{"id": "1", "name": "Edward Ghattas", "email": EDDIE_EMAIL}])
        client = self._client_returning(_page(node))
        campaigns = m.EagleCampaignSource(client, agent_email=EDDIE_EMAIL).get_active_campaigns()
        self.assertEqual(len(campaigns), 1)

    def test_email_match_is_case_insensitive(self):
        node = _node(agents=[{"id": "1", "name": "Edward Ghattas", "email": "Edward@IBProperty.com.au"}])
        client = self._client_returning(_page(node))
        campaigns = m.EagleCampaignSource(client, agent_email=EDDIE_EMAIL).get_active_campaigns()
        self.assertEqual(len(campaigns), 1)

    def test_other_agents_email_excluded(self):
        node = _node(agents=[{"id": "2", "name": "Steffan Ippolito", "email": "steffan@ibproperty.com.au"}])
        client = self._client_returning(_page(node))
        campaigns = m.EagleCampaignSource(client, agent_email=EDDIE_EMAIL).get_active_campaigns()
        self.assertEqual(campaigns, [])

    def test_co_listed_property_included(self):
        # Real shape: most of Eddie's 17 are co-listed with another agent.
        node = _node(agents=[
            {"id": "2", "name": "Steffan Ippolito", "email": "steffan@ibproperty.com.au"},
            {"id": "1", "name": "Edward Ghattas", "email": EDDIE_EMAIL},
        ])
        client = self._client_returning(_page(node))
        campaigns = m.EagleCampaignSource(client, agent_email=EDDIE_EMAIL).get_active_campaigns()
        self.assertEqual(len(campaigns), 1)

    def test_name_fallback_used_when_email_blank(self):
        node = _node(agents=[{"id": "1", "name": "Edward Ghattas", "email": ""}])
        client = self._client_returning(_page(node))
        campaigns = m.EagleCampaignSource(
            client, agent_email=EDDIE_EMAIL, agent_name="Edward Ghattas"
        ).get_active_campaigns()
        self.assertEqual(len(campaigns), 1)

    def test_name_fallback_not_used_when_not_provided(self):
        node = _node(agents=[{"id": "1", "name": "Edward Ghattas", "email": ""}])
        client = self._client_returning(_page(node))
        campaigns = m.EagleCampaignSource(client, agent_email=EDDIE_EMAIL).get_active_campaigns()
        self.assertEqual(campaigns, [])

    def test_name_fallback_matches_despite_trailing_whitespace(self):
        # Real data has at least one agent name with a stray trailing space
        # ("Isaac Jackson "), so the fallback must tolerate that.
        node = _node(agents=[{"id": "1", "name": "Edward Ghattas ", "email": ""}])
        client = self._client_returning(_page(node))
        campaigns = m.EagleCampaignSource(
            client, agent_email=EDDIE_EMAIL, agent_name="Edward Ghattas"
        ).get_active_campaigns()
        self.assertEqual(len(campaigns), 1)

    def test_email_takes_precedence_over_a_non_matching_name(self):
        node = _node(agents=[{"id": "1", "name": "Ed G", "email": EDDIE_EMAIL}])
        client = self._client_returning(_page(node))
        campaigns = m.EagleCampaignSource(
            client, agent_email=EDDIE_EMAIL, agent_name="Edward Ghattas"
        ).get_active_campaigns()
        self.assertEqual(len(campaigns), 1)

    def test_no_agents_on_property_excluded(self):
        node = _node(agents=[])
        client = self._client_returning(_page(node))
        campaigns = m.EagleCampaignSource(client, agent_email=EDDIE_EMAIL).get_active_campaigns()
        self.assertEqual(campaigns, [])

    def test_filters_across_paginated_results(self):
        mine = _node(id="mine", agents=[{"id": "1", "name": "Edward Ghattas", "email": EDDIE_EMAIL}])
        not_mine = _node(id="not-mine", agents=[{"id": "2", "name": "Steffan Ippolito", "email": "steffan@ibproperty.com.au"}])
        client = self._client_returning(
            _page(mine, has_next=True, cursor="c1"),
            _page(not_mine),
        )
        campaigns = m.EagleCampaignSource(client, agent_email=EDDIE_EMAIL).get_active_campaigns()
        self.assertEqual([c.property_id for c in campaigns], ["mine"])


if __name__ == "__main__":
    unittest.main()
