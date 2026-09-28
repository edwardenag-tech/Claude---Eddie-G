"""Tests for the branded ibproperty.com.au listing-link feature: resolving a
known listing URL for an address (campaign doc > listings_db.json >
listing_links.json, in that priority order, never invented), and making sure
claude_draft_reply actually includes it in the drafted reply -- on the
lease/sale template paths and on the Sent Items "replicate a past reply" path.
"""

import os
import sys
import unittest
from unittest.mock import MagicMock

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import draft_agent as m
import campaign_doc_parser


class TestMatchByAddress(unittest.TestCase):
    """find_listing_in_db and find_campaign_for_address both delegate to the
    shared _match_by_address matcher -- exercised through the public
    functions so a refactor of the internals can't silently change behavior."""

    def test_find_listing_in_db_matches_by_number_and_street_word(self):
        db = {"listings": [{"address": "490 Pacific Highway, St Leonards NSW 2065", "price": "$1m"}]}
        match = m.find_listing_in_db(db, "Suite 207, 490 Pacific Highway, St Leonards")
        self.assertEqual(match["price"], "$1m")

    def test_find_listing_in_db_no_match_returns_none(self):
        db = {"listings": [{"address": "12 Smith Street, Sydney"}]}
        self.assertIsNone(m.find_listing_in_db(db, "490 Pacific Highway, St Leonards"))

    def test_find_listing_in_db_blank_address_returns_none(self):
        self.assertIsNone(m.find_listing_in_db({"listings": []}, ""))

    def test_find_campaign_for_address_matches_by_number_and_street_word(self):
        campaigns = [{"address": "196 Pacific Highway, St Leonards", "status": "Live"}]
        match = m.find_campaign_for_address(campaigns, "196 Pacific Highway, St Leonards NSW")
        self.assertEqual(match["status"], "Live")


class TestListingLinksFile(unittest.TestCase):
    def test_load_listing_links_missing_file_returns_empty(self):
        links = m.load_listing_links()
        self.assertIsInstance(links, dict)

    def test_find_listing_link_matches_by_address(self):
        links_db = {"listings": [
            {"address": "490 Pacific Highway, St Leonards NSW 2065",
             "url": "https://ibproperty.com.au/commercial/properties-for-lease/490-stleonards-nsw-2065/6186"},
        ]}
        url = m.find_listing_link(links_db, "Suite 207, 490 Pacific Highway, St Leonards NSW 2065")
        self.assertEqual(url, "https://ibproperty.com.au/commercial/properties-for-lease/490-stleonards-nsw-2065/6186")

    def test_find_listing_link_no_match_returns_none(self):
        links_db = {"listings": [{"address": "12 Smith Street", "url": "https://example.com/x"}]}
        self.assertIsNone(m.find_listing_link(links_db, "490 Pacific Highway"))

    def test_the_committed_listing_links_json_resolves_the_real_st_leonards_listing(self):
        # Guards against the seeded real entry being edited/removed by accident.
        links = m.load_listing_links()
        url = m.find_listing_link(links, "Suite 207, 490 Pacific Highway, St Leonards NSW 2065")
        self.assertEqual(
            url, "https://ibproperty.com.au/commercial/properties-for-lease/490-stleonards-nsw-2065/6186"
        )


class TestFindListingUrlPriority(unittest.TestCase):
    ADDRESS = "490 Pacific Highway, St Leonards NSW 2065"

    def test_no_source_has_it_returns_none(self):
        self.assertIsNone(m.find_listing_url(self.ADDRESS, [], {"listings": []}, {"listings": []}))

    def test_blank_address_returns_none_without_checking_sources(self):
        campaigns = [{"address": self.ADDRESS, "listing_link": "https://doc.example/x"}]
        self.assertIsNone(m.find_listing_url("", campaigns, {"listings": []}, {"listings": []}))

    def test_listing_links_json_used_when_nothing_else_matches(self):
        links_db = {"listings": [{"address": self.ADDRESS, "url": "https://cache.example/x"}]}
        url = m.find_listing_url(self.ADDRESS, [], {"listings": []}, links_db)
        self.assertEqual(url, "https://cache.example/x")

    def test_listings_db_outranks_listing_links_json(self):
        listings_db = {"listings": [{"address": self.ADDRESS, "listing_url": "https://db.example/x"}]}
        links_db = {"listings": [{"address": self.ADDRESS, "url": "https://cache.example/x"}]}
        url = m.find_listing_url(self.ADDRESS, [], listings_db, links_db)
        self.assertEqual(url, "https://db.example/x")

    def test_campaign_doc_outranks_everything(self):
        campaigns = [{"address": self.ADDRESS, "listing_link": "https://doc.example/x"}]
        listings_db = {"listings": [{"address": self.ADDRESS, "listing_url": "https://db.example/x"}]}
        links_db = {"listings": [{"address": self.ADDRESS, "url": "https://cache.example/x"}]}
        url = m.find_listing_url(self.ADDRESS, campaigns, listings_db, links_db)
        self.assertEqual(url, "https://doc.example/x")

    def test_campaign_doc_match_without_listing_link_falls_through_to_next_source(self):
        # Address matches a campaign block, but that block has no listing_link
        # filled in -- must not treat "matched, blank field" as "use nothing".
        campaigns = [{"address": self.ADDRESS, "listing_link": ""}]
        listings_db = {"listings": [{"address": self.ADDRESS, "listing_url": "https://db.example/x"}]}
        url = m.find_listing_url(self.ADDRESS, campaigns, listings_db, {"listings": []})
        self.assertEqual(url, "https://db.example/x")


class TestCampaignDocParsesListingLink(unittest.TestCase):
    def test_listing_link_field_parsed(self):
        doc_text = (
            "ACTIVE CAMPAIGNS\n"
            "===\n"
            "Property: 490 Pacific Highway, St Leonards\n"
            "Listing link: https://ibproperty.com.au/commercial/properties-for-lease/490-stleonards-nsw-2065/6186\n"
            "===\n"
        )
        campaigns = campaign_doc_parser.parse_active_campaigns(doc_text)
        self.assertEqual(len(campaigns), 1)
        self.assertEqual(
            campaigns[0]["listing_link"],
            "https://ibproperty.com.au/commercial/properties-for-lease/490-stleonards-nsw-2065/6186",
        )

    def test_placeholder_listing_link_is_blank(self):
        doc_text = "===\nProperty: 1 Test St\nListing link: [paste the ibproperty.com.au URL]\n===\n"
        campaigns = campaign_doc_parser.parse_active_campaigns(doc_text)
        self.assertEqual(campaigns[0]["listing_link"], "")


class TestClaudeDraftReplyIncludesListingLink(unittest.TestCase):
    def _ai(self, text="<p>ok</p>"):
        ai = MagicMock()
        ai.messages.create.return_value.content = [MagicMock(text=text)]
        return ai

    def test_lease_enquiry_prompt_hyperlinks_address_when_listing_url_known(self):
        ai = self._ai()
        listing_details = {
            "address": "490 Pacific Highway, St Leonards", "asking_rent": "$45,000 p.a.",
            "internal_area": "81 sqm", "listing_url": "https://ibproperty.com.au/x/6186",
        }
        m.claude_draft_reply(ai, {"subject": "s", "body": "b", "from": "a@b.com"}, "lease_enquiry", [],
                              listing_details=listing_details)
        prompt = ai.messages.create.call_args.kwargs["messages"][0]["content"]
        self.assertIn('<a href="https://ibproperty.com.au/x/6186">490 Pacific Highway, St Leonards</a>', prompt)

    def test_lease_enquiry_prompt_has_plain_address_when_no_listing_url(self):
        ai = self._ai()
        listing_details = {"address": "490 Pacific Highway, St Leonards", "asking_rent": "$45,000 p.a.",
                            "internal_area": "81 sqm"}
        m.claude_draft_reply(ai, {"subject": "s", "body": "b", "from": "a@b.com"}, "lease_enquiry", [],
                              listing_details=listing_details)
        prompt = ai.messages.create.call_args.kwargs["messages"][0]["content"]
        self.assertNotIn("<a href", prompt)
        self.assertIn("bring to market 490 Pacific Highway, St Leonards,", prompt)

    def test_sale_enquiry_prompt_hyperlinks_address_when_listing_url_known(self):
        ai = self._ai()
        listing_details = {"address": "1 Test St, Sydney", "internal_area": "100 sqm",
                            "listing_url": "https://ibproperty.com.au/y/1"}
        m.claude_draft_reply(ai, {"subject": "s", "body": "b", "from": "a@b.com"}, "sale_enquiry", [],
                              listing_details=listing_details)
        prompt = ai.messages.create.call_args.kwargs["messages"][0]["content"]
        self.assertIn('<a href="https://ibproperty.com.au/y/1">1 Test St, Sydney</a>', prompt)

    def test_sent_template_path_gets_listing_link_instruction_when_known(self):
        ai = self._ai()
        m.claude_draft_reply(
            ai, {"subject": "s", "body": "b", "from": "a@b.com", "from_name": "Bob Buyer"},
            "lease_enquiry", [], sent_template="Hi there, re 490 Pacific Highway...",
            listing_url="https://ibproperty.com.au/x/6186",
        )
        prompt = ai.messages.create.call_args.kwargs["messages"][0]["content"]
        self.assertIn("https://ibproperty.com.au/x/6186", prompt)
        self.assertIn("<a href=", prompt)

    def test_sent_template_path_has_no_listing_link_mention_when_unknown(self):
        ai = self._ai()
        m.claude_draft_reply(
            ai, {"subject": "s", "body": "b", "from": "a@b.com", "from_name": "Bob Buyer"},
            "lease_enquiry", [], sent_template="Hi there, re 490 Pacific Highway...",
        )
        prompt = ai.messages.create.call_args.kwargs["messages"][0]["content"]
        self.assertNotIn("Listing link", prompt)
        self.assertNotIn("<a href", prompt)


if __name__ == "__main__":
    unittest.main()
