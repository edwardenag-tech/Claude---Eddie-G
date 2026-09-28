"""Tests for vendor_update_agent.py: landlord-contact precedence, ledger
backfill, the Claude drafting call's continuity handling, and main()'s
orchestration -- Eagle fail-closed, dry-run safety, per-campaign isolation,
ledger persistence. Everything external (Outlook, Gmail, Eagle, Anthropic)
is mocked -- no network.
"""

import json
import os
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import vendor_update_agent as m
from campaign_source import Campaign, CampaignContact
from eagle_client import EagleAuthRequired
from gmail_client import GmailAuthRequired
from vendor_update_ledger import Lead, LedgerEntry

ENV = {
    "ANTHROPIC_API_KEY": "k", "AZURE_CLIENT_ID": "c", "AZURE_TENANT_ID": "t",
    "EAGLE_CLIENT_ID": "eid", "EAGLE_CLIENT_SECRET": "esecret",
}


def _campaign(**overrides) -> Campaign:
    defaults = dict(
        property_id="1825106",
        address="207/490 Pacific Highway, St Leonards",
        rea_id="eagle_1825106",
        sale_or_lease="LEASE",
        active_at="2026-08-01T00:00:00Z",
        days_on_market=49,
        num_enquiries=5,
        num_inspection_attendances=4,
        num_offers=0,
        landlord_contacts=[],
    )
    defaults.update(overrides)
    return Campaign(**defaults)


class TestExtractJsonObject(unittest.TestCase):
    """Confirmed live: asking the model to judge whether correspondence is
    genuinely about the right property made it reason out loud in prose
    before the JSON for one real draft (Pymble), despite being told not to
    -- this is the parser fix so that doesn't silently produce a fallback
    generic email."""

    def test_clean_json(self):
        self.assertEqual(m._extract_json_object('{"a": 1}'), {"a": 1})

    def test_markdown_fenced_json(self):
        self.assertEqual(m._extract_json_object('```json\n{"a": 1}\n```'), {"a": 1})

    def test_prose_preamble_before_json(self):
        raw = 'I need to carefully analyze this.\n\nHere is my answer:\n{"a": 1}'
        self.assertEqual(m._extract_json_object(raw), {"a": 1})

    def test_prose_preamble_with_fenced_json(self):
        raw = 'Let me think through this step by step...\n\n```json\n{"a": 1, "b": "text"}\n```'
        self.assertEqual(m._extract_json_object(raw), {"a": 1, "b": "text"})


class TestResolveLandlordContact(unittest.TestCase):
    @patch("vendor_update_agent.find_landlord_first_name", return_value="Sandra")
    @patch("vendor_update_agent.find_landlord_email", return_value="sandra@example.com")
    def test_sent_items_result_used_when_present(self, _email, _name):
        outlook = MagicMock()
        campaign = _campaign(landlord_contacts=[CampaignContact(first_name="EagleName", email="eagle@example.com")])
        email, name = m.resolve_landlord_contact(outlook, campaign)
        self.assertEqual(email, "sandra@example.com")
        self.assertEqual(name, "Sandra")

    @patch("vendor_update_agent.find_landlord_first_name", return_value=None)
    @patch("vendor_update_agent.find_landlord_email", return_value=None)
    def test_falls_back_to_eagle_vendor_when_sent_items_empty(self, _email, _name):
        outlook = MagicMock()
        campaign = _campaign(landlord_contacts=[CampaignContact(first_name="Will", email="will@example.com")])
        email, name = m.resolve_landlord_contact(outlook, campaign)
        self.assertEqual(email, "will@example.com")
        self.assertEqual(name, "Will")

    @patch("vendor_update_agent.find_landlord_first_name", return_value=None)
    @patch("vendor_update_agent.find_landlord_email", return_value=None)
    def test_no_source_has_contact_returns_none(self, _email, _name):
        outlook = MagicMock()
        campaign = _campaign(landlord_contacts=[])
        email, name = m.resolve_landlord_contact(outlook, campaign)
        self.assertIsNone(email)
        self.assertIsNone(name)


class TestBackfillLedgerEntry(unittest.TestCase):
    def _ai(self, text):
        ai = MagicMock()
        ai.messages.create.return_value.content = [MagicMock(text=text)]
        return ai

    @patch("vendor_update_agent._search_sent_items")
    def test_no_past_update_returns_blank_entry(self, search):
        search.return_value = [{"subject": "Re: Random unrelated thread", "sentDateTime": "2026-09-01T00:00:00Z", "body": {"content": "hi"}}]
        entry = m.backfill_ledger_entry(MagicMock(), self._ai("{}"), "1 Test St")
        self.assertEqual(entry.address, "1 Test St")
        self.assertEqual(entry.week_number, 0)
        self.assertEqual(entry.known_leads, [])

    @patch("vendor_update_agent._search_sent_items")
    def test_extracts_leads_from_most_recent_past_update(self, search):
        search.return_value = [
            {"subject": "1 Test St - Week 5 Campaign Update", "sentDateTime": "2026-09-01T00:00:00Z",
             "body": {"content": "<p>old</p>"}},
            {"subject": "1 Test St - Week 7 Campaign Update", "sentDateTime": "2026-09-15T00:00:00Z",
             "body": {"content": "<p>new</p>"}},
        ]
        ai = self._ai(json.dumps({
            "week_number": 7,
            "leads": [{"name": "The Art School", "status": "Following up"}],
        }))
        entry = m.backfill_ledger_entry(MagicMock(), ai, "1 Test St")
        self.assertEqual(entry.week_number, 7)
        self.assertEqual(len(entry.known_leads), 1)
        self.assertEqual(entry.known_leads[0].name, "The Art School")
        self.assertEqual(entry.known_leads[0].last_updated, "2026-09-15")
        # picked the WEEK 7 email (most recent), not week 5
        prompt = ai.messages.create.call_args.kwargs["messages"][0]["content"]
        self.assertIn("Week 7", prompt)

    @patch("vendor_update_agent._search_sent_items")
    def test_reply_subjects_excluded_from_candidates(self, search):
        search.return_value = [
            {"subject": "Re: 1 Test St - Week 5 Campaign Update", "sentDateTime": "2026-09-20T00:00:00Z", "body": {"content": "reply"}},
        ]
        entry = m.backfill_ledger_entry(MagicMock(), self._ai("{}"), "1 Test St")
        self.assertEqual(entry.known_leads, [])
        self.assertEqual(entry.week_number, 0)

    @patch("vendor_update_agent._search_sent_items")
    def test_claude_failure_starts_blank_not_raises(self, search):
        search.return_value = [
            {"subject": "1 Test St - Week 5 Campaign Update", "sentDateTime": "2026-09-01T00:00:00Z", "body": {"content": "x"}},
        ]
        ai = MagicMock()
        ai.messages.create.side_effect = Exception("boom")
        entry = m.backfill_ledger_entry(MagicMock(), ai, "1 Test St")
        self.assertEqual(entry.known_leads, [])


class TestStyleExamples(unittest.TestCase):
    @patch("vendor_update_agent._search_sent_items")
    def test_fetch_style_examples_merges_multiple_queries_and_dedupes(self, search):
        # Same message surfaces under both the "campaign update" and
        # "leasing update" queries (Graph's fuzzy $search would do this) --
        # must not appear twice.
        shared = {"id": "1", "subject": "1 Test St - Week 5 Campaign Update",
                  "sentDateTime": "2026-09-01T00:00:00Z", "body": {"content": "<p>a</p>"}}
        victoria = {"id": "2", "subject": "176 Victoria Street, Potts Point - Week 7 Leasing Update",
                    "sentDateTime": "2026-09-14T00:00:00Z", "body": {"content": "<p>b</p>"}}
        search.side_effect = lambda outlook, query, top=10: {
            "campaign update": [shared],
            "leasing update": [shared, victoria],
            "weekly update": [],
        }[query]
        examples = m.fetch_style_examples(MagicMock())
        self.assertEqual(len(examples), 2)
        subjects = "\n".join(examples)
        self.assertIn("Week 5 Campaign Update", subjects)
        self.assertIn("Week 7 Leasing Update", subjects)  # would be missed by a literal "campaign update" search alone

    @patch("vendor_update_agent._search_sent_items")
    def test_fetch_style_examples_excludes_replies_and_caps_at_five(self, search):
        msgs = [
            {"id": str(i), "subject": f"Property {i} - Week {i} Campaign Update",
             "sentDateTime": f"2026-09-{i:02d}T00:00:00Z", "body": {"content": f"<p>{i}</p>"}}
            for i in range(1, 8)
        ]
        msgs.append({"id": "reply", "subject": "Re: Property 1 - Week 1 Campaign Update",
                     "sentDateTime": "2026-09-10T00:00:00Z", "body": {"content": "<p>reply</p>"}})
        search.side_effect = lambda outlook, query, top=10: msgs if query == "campaign update" else []
        examples = m.fetch_style_examples(MagicMock())
        self.assertEqual(len(examples), 5)
        self.assertFalse(any("Re:" in e for e in examples))

    @patch("vendor_update_agent._search_sent_items")
    def test_property_specific_style_example_none_when_no_history(self, search):
        search.return_value = []
        self.assertIsNone(m.fetch_property_specific_style_example(MagicMock(), "1 New St"))

    @patch("vendor_update_agent._search_sent_items")
    def test_property_specific_style_example_returns_most_recent_for_that_address(self, search):
        search.return_value = [
            {"subject": "1 Test St - Week 5 Campaign Update", "sentDateTime": "2026-09-01T00:00:00Z", "body": {"content": "<p>old</p>"}},
            {"subject": "1 Test St - Week 7 Campaign Update", "sentDateTime": "2026-09-15T00:00:00Z", "body": {"content": "<p>Hi Sandra, new content here</p>"}},
        ]
        example = m.fetch_property_specific_style_example(MagicMock(), "1 Test St")
        self.assertIn("Week 7 Campaign Update", example)
        self.assertIn("new content here", example)
        self.assertNotIn("old", example)


class TestGatherInspectionAndOfferHistory(unittest.TestCase):
    """gather_weekly_activity alone was confirmed (by reading the code) to do
    a generic address+7-day search with no targeting of inspection/offer
    language at all -- this is the dedicated function that fixes that, per
    Eddie's review: real inspections/offers were happening in correspondence
    but not surfacing in drafts. Checked live against his real 17 campaigns:
    Eagle's own numInspectionAttendances/numOffers were 0 for all 17 despite
    real, confirmed inspection/HOA activity for several (249 Pitt Street,
    1/67 Burns Bay Road, 6/895 Pacific Highway, 488 Pacific Highway) -- this
    function is what makes that real correspondence reachable.
    """

    def test_keyword_match_required_not_just_address_match(self):
        outlook = MagicMock()
        outlook._get.return_value = {"value": [
            {"id": "1", "subject": "1 Test St enquiry", "receivedDateTime": "2026-09-01T00:00:00Z",
             "body": {"content": "<p>just asking about parking, nothing else</p>"}},
        ]}
        result = m.gather_inspection_and_offer_history(outlook, "1 Test St")
        self.assertEqual(result, [])  # matches address but no inspection/offer language

    def test_inspection_language_matched(self):
        outlook = MagicMock()
        outlook._get.side_effect = lambda endpoint, params=None: (
            {"value": [
                {"id": "1", "subject": "Heads of agreement - 1 Test St", "receivedDateTime": "2026-08-01T00:00:00Z",
                 "body": {"content": "<p>Please find attached the HOA for 1 Test St, offer of $500,000</p>"}},
            ]} if "messages" in endpoint and "sentitems" not in endpoint else {"value": []}
        )
        result = m.gather_inspection_and_offer_history(outlook, "1 Test St")
        self.assertEqual(len(result), 1)
        self.assertIn("$500,000", result[0])

    def test_not_limited_to_recent_dates_unlike_gather_weekly_activity(self):
        # A real gap this session: 249 Pitt Street's HOA was from ~2 weeks
        # before the run and gather_weekly_activity's 7-day window missed it
        # entirely. This function must not apply any such cutoff.
        outlook = MagicMock()
        old_date = "2026-01-01T00:00:00Z"  # far outside any 7-day window
        outlook._get.side_effect = lambda endpoint, params=None: (
            {"value": [
                {"id": "1", "subject": "Offer - 1 Test St", "receivedDateTime": old_date,
                 "body": {"content": "<p>Offer received: $200,000 from Jane Smith</p>"}},
            ]} if "messages" in endpoint and "sentitems" not in endpoint else {"value": []}
        )
        result = m.gather_inspection_and_offer_history(outlook, "1 Test St")
        self.assertEqual(len(result), 1)
        self.assertIn("Jane Smith", result[0])

    def test_past_weekly_update_subjects_excluded(self):
        outlook = MagicMock()
        outlook._get.return_value = {"value": [
            {"id": "1", "subject": "1 Test St - Week 5 Campaign Update", "receivedDateTime": "2026-09-01T00:00:00Z",
             "body": {"content": "<p>we conducted an inspection and received an offer</p>"}},
        ]}
        result = m.gather_inspection_and_offer_history(outlook, "1 Test St")
        self.assertEqual(result, [])  # Eddie's own past update, not raw correspondence

    def test_dedupes_message_seen_in_both_inbox_and_sent_search(self):
        outlook = MagicMock()
        shared_msg = {"id": "dup", "subject": "Offer on 1 Test St", "receivedDateTime": "2026-09-01T00:00:00Z",
                      "sentDateTime": "2026-09-01T00:00:00Z", "body": {"content": "<p>offer of $1</p>"}}
        outlook._get.return_value = {"value": [shared_msg]}
        result = m.gather_inspection_and_offer_history(outlook, "1 Test St")
        self.assertEqual(len(result), 1)

    def test_snippet_length_is_longer_than_general_activity(self):
        # gather_weekly_activity truncates to 300 chars -- key offer/inspection
        # details are often further in, so this needs more room.
        outlook = MagicMock()
        long_body = "Offer details: " + ("x" * 500) + " final amount $750,000"
        outlook._get.side_effect = lambda endpoint, params=None: (
            {"value": [
                {"id": "1", "subject": "Offer - 1 Test St", "receivedDateTime": "2026-09-01T00:00:00Z",
                 "body": {"content": f"<p>{long_body}</p>"}},
            ]} if "messages" in endpoint and "sentitems" not in endpoint else {"value": []}
        )
        result = m.gather_inspection_and_offer_history(outlook, "1 Test St")
        self.assertIn("$750,000", result[0])


class TestClaudeDraftVendorUpdate(unittest.TestCase):
    def _ai(self, text):
        ai = MagicMock()
        ai.messages.create.return_value.content = [MagicMock(text=text)]
        return ai

    def test_uses_eagle_numbers_and_ledger_continuity(self):
        ai = self._ai(json.dumps({
            "html_body": "<p>Hi Sandra,</p><p>Update...</p><p>Thank you.</p>",
            "leads": [{"name": "The Art School", "status": "Still interested"}],
        }))
        campaign = _campaign(num_enquiries=5, num_inspection_attendances=4, days_on_market=49)
        ledger_entry = LedgerEntry(
            address=campaign.address, week_number=6,
            known_leads=[Lead(name="The Art School", status="Was due to come back", last_updated="2026-09-07")],
        )
        subject, html_body, updated = m.claude_draft_vendor_update(
            ai, campaign, ledger_entry, "Sandra", [], [], [],
        )
        prompt = ai.messages.create.call_args.kwargs["messages"][0]["content"]
        self.assertIn("5 enquiries to date", prompt)
        self.assertIn("4 inspections", prompt)
        self.assertIn("The Art School: Was due to come back", prompt)
        self.assertEqual(subject, f"{campaign.address} — Week 8 Campaign Update")  # 49//7+1
        self.assertIn("Hi Sandra", html_body)
        self.assertEqual(updated.known_leads[0].name, "The Art School")
        self.assertEqual(updated.known_leads[0].status, "Still interested")
        self.assertEqual(updated.week_number, 8)

    def test_week_number_falls_back_to_ledger_plus_one_when_days_on_market_unknown(self):
        ai = self._ai(json.dumps({"html_body": "<p>x</p>", "leads": []}))
        campaign = _campaign(days_on_market=None)
        ledger_entry = LedgerEntry(address=campaign.address, week_number=3)
        subject, _html, updated = m.claude_draft_vendor_update(
            ai, campaign, ledger_entry, "Sandra", [], [], [],
        )
        self.assertIn("Week 4", subject)
        self.assertEqual(updated.week_number, 4)

    def test_claude_failure_falls_back_to_generic_body_and_keeps_known_leads(self):
        ai = MagicMock()
        ai.messages.create.side_effect = Exception("rate limited")
        campaign = _campaign()
        existing_lead = Lead(name="Cigar Lounge", status="Not pursued", last_updated="2026-09-07")
        ledger_entry = LedgerEntry(address=campaign.address, week_number=6, known_leads=[existing_lead])
        subject, html_body, updated = m.claude_draft_vendor_update(
            ai, campaign, ledger_entry, "Sandra", [], [], [],
        )
        self.assertIn("Hi Sandra", html_body)
        self.assertIn("Thank you.", html_body)
        self.assertEqual(updated.known_leads, [existing_lead])  # not lost on failure

    def test_model_reasoning_preamble_before_json_still_parses(self):
        # Regression test for the real Pymble failure: the model reasoned
        # out loud about property relevance before emitting the JSON.
        raw = (
            'I need to carefully analyze the inspection/offer correspondence to determine '
            'what is genuinely about this property.\n\nKey observations:\n- discard unrelated one\n\n'
            '```json\n{"html_body": "<p>Hi Sandra,</p><p>Real content.</p><p>Thank you.</p>", '
            '"leads": [{"name": "David Tebbutt", "status": "Reviewing"}]}\n```'
        )
        ai = self._ai(raw)
        campaign = _campaign()
        subject, html_body, updated = m.claude_draft_vendor_update(
            ai, campaign, LedgerEntry(address=campaign.address), "Sandra", [], [], [],
        )
        self.assertIn("Real content", html_body)
        self.assertEqual(updated.known_leads[0].name, "David Tebbutt")

    def test_malformed_json_response_falls_back_gracefully(self):
        ai = self._ai("not json at all")
        campaign = _campaign()
        ledger_entry = LedgerEntry(address=campaign.address)
        subject, html_body, updated = m.claude_draft_vendor_update(
            ai, campaign, ledger_entry, None, [], [], [],
        )
        self.assertIn("[LANDLORD FIRST NAME - PLEASE UPDATE]", html_body)

    def test_activity_gap_instruction_present_in_prompt(self):
        ai = self._ai(json.dumps({"html_body": "<p>x</p>", "leads": []}))
        campaign = _campaign()
        m.claude_draft_vendor_update(ai, campaign, LedgerEntry(address=campaign.address), "Sandra", [], [], [])
        prompt = ai.messages.create.call_args.kwargs["messages"][0]["content"]
        self.assertIn("meaningfully higher", prompt.lower())

    def test_no_corporate_filler_instruction_present(self):
        ai = self._ai(json.dumps({"html_body": "<p>x</p>", "leads": []}))
        campaign = _campaign()
        m.claude_draft_vendor_update(ai, campaign, LedgerEntry(address=campaign.address), "Sandra", [], [], [])
        prompt = ai.messages.create.call_args.kwargs["messages"][0]["content"]
        self.assertIn("please do not hesitate to reach out", prompt.lower())
        self.assertIn("robotic", prompt.lower())

    def test_property_specific_style_example_weighted_above_general_examples(self):
        ai = self._ai(json.dumps({"html_body": "<p>x</p>", "leads": []}))
        campaign = _campaign()
        m.claude_draft_vendor_update(
            ai, campaign, LedgerEntry(address=campaign.address), "Sandra", [], [],
            style_examples=["Subject: Other St Update\nGeneric body"],
            property_style_example="Subject: 207/490 Pacific Highway - Week 5 Campaign Update\nHi Kay, real past email",
        )
        prompt = ai.messages.create.call_args.kwargs["messages"][0]["content"]
        self.assertIn("MOST IMPORTANT STYLE REFERENCE", prompt)
        self.assertIn("Hi Kay, real past email", prompt)
        self.assertIn("SECONDARY STYLE REFERENCE", prompt)
        # the property-specific example must appear before the general ones
        self.assertLess(prompt.index("Hi Kay, real past email"), prompt.index("Generic body"))

    def test_crm_inspection_offer_count_framed_as_unreliable(self):
        ai = self._ai(json.dumps({"html_body": "<p>x</p>", "leads": []}))
        campaign = _campaign(num_inspection_attendances=0, num_offers=0)
        m.claude_draft_vendor_update(ai, campaign, LedgerEntry(address=campaign.address), "Sandra", [], [], [])
        prompt = ai.messages.create.call_args.kwargs["messages"][0]["content"]
        self.assertIn("UNRELIABLE", prompt)
        self.assertIn("Do NOT state this count", prompt)

    def test_inspection_offer_snippets_included_and_instructed_to_name_specifics(self):
        ai = self._ai(json.dumps({"html_body": "<p>x</p>", "leads": []}))
        campaign = _campaign()
        m.claude_draft_vendor_update(
            ai, campaign, LedgerEntry(address=campaign.address), "Sandra", [], [], [],
            inspection_offer_snippets=["Subject: Offer - 1 Test St | Date: 2026-08-01 | Offer of $500,000 from Jane Smith"],
        )
        prompt = ai.messages.create.call_args.kwargs["messages"][0]["content"]
        self.assertIn("POSSIBLE INSPECTION / OFFER CORRESPONDENCE", prompt)
        self.assertIn("Jane Smith", prompt)
        self.assertIn("MUST name the specific inspection", prompt)
        self.assertIn("may be about a DIFFERENT property", prompt)

    def test_no_inspection_offer_snippets_omits_that_section(self):
        # The CRM-numbers framing always references the section by name (so
        # the model knows where to look if it's there) -- what must NOT
        # appear with no snippets is the actual populated block content.
        ai = self._ai(json.dumps({"html_body": "<p>x</p>", "leads": []}))
        campaign = _campaign()
        m.claude_draft_vendor_update(
            ai, campaign, LedgerEntry(address=campaign.address), "Sandra", [], [], [],
            inspection_offer_snippets=None,
        )
        prompt = ai.messages.create.call_args.kwargs["messages"][0]["content"]
        # The standing instruction bullet ("if present above...") is fine to
        # always be there; the populated block itself must not be.
        self.assertNotIn("may be about a DIFFERENT property", prompt)
        self.assertNotIn("Date: 2026", prompt)

    def test_no_property_style_example_omits_that_section(self):
        ai = self._ai(json.dumps({"html_body": "<p>x</p>", "leads": []}))
        campaign = _campaign()
        m.claude_draft_vendor_update(
            ai, campaign, LedgerEntry(address=campaign.address), "Sandra", [], [], [],
            property_style_example=None,
        )
        prompt = ai.messages.create.call_args.kwargs["messages"][0]["content"]
        self.assertNotIn("MOST IMPORTANT STYLE REFERENCE", prompt)


class TestMainOrchestration(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.ledger_path = os.path.join(self.tmp, "ledger.json")
        self.env_patch = patch.dict(os.environ, ENV, clear=False)
        self.env_patch.start()
        self.addCleanup(self.env_patch.stop)

    def _patches(self, campaigns, dry_run=False):
        return {
            "outlook_cls": patch("vendor_update_agent.OutlookClient"),
            "gmail_cls": patch("vendor_update_agent.GmailClient"),
            "eagle_cls": patch("vendor_update_agent.EagleClient"),
            "source_cls": patch("vendor_update_agent.EagleCampaignSource"),
            "anthropic": patch("vendor_update_agent.anthropic.Anthropic"),
            "style": patch("vendor_update_agent.fetch_style_examples", return_value=[]),
            "property_style": patch("vendor_update_agent.fetch_property_specific_style_example", return_value=None),
            "inspection_offer": patch("vendor_update_agent.gather_inspection_and_offer_history", return_value=[]),
            "resolve_contact": patch("vendor_update_agent.resolve_landlord_contact", return_value=("owner@example.com", "Sandra")),
            "backfill": patch("vendor_update_agent.backfill_ledger_entry", return_value=LedgerEntry(address="x")),
            "activity": patch("vendor_update_agent.gather_weekly_activity", return_value=([], [])),
            "draft": patch(
                "vendor_update_agent.claude_draft_vendor_update",
                side_effect=lambda ai, campaign, entry, *a, **k: (
                    f"{campaign.address} — Week 1 Campaign Update", "<p>body</p>",
                    LedgerEntry(address=campaign.address, week_number=1),
                ),
            ),
            "outlook_draft": patch("vendor_update_agent.outlook_create_new_draft", return_value="OD1"),
            "gmail_draft": patch("vendor_update_agent.gmail_create_new_draft", return_value="GD1"),
            "ledger_path": patch("vendor_update_agent.load_ledger", side_effect=lambda: {}),
            "save_ledger": patch("vendor_update_agent.save_ledger"),
        }

    def _run_with(self, campaigns, dry_run=False, extra_patches=None):
        patches = self._patches(campaigns, dry_run=dry_run)
        started = {name: p.start() for name, p in patches.items()}
        for p in patches.values():
            self.addCleanup(p.stop)
        started["source_cls"].return_value.get_active_campaigns.return_value = campaigns
        if extra_patches:
            for fn in extra_patches:
                fn(started)
        return m.main(dry_run=dry_run), started

    def test_no_active_campaigns_returns_zero_without_error(self):
        result, _ = self._run_with([])
        self.assertEqual(result, {"drafted": 0, "failed": 0, "campaigns": 0})

    def test_eagle_auth_failure_is_fail_closed_not_crash(self):
        patches = self._patches([])
        started = {name: p.start() for name, p in patches.items()}
        for p in patches.values():
            self.addCleanup(p.stop)
        started["eagle_cls"].side_effect = EagleAuthRequired("bad creds")
        with self.assertRaises(SystemExit) as ctx:
            m.main(dry_run=False)
        self.assertEqual(ctx.exception.code, 1)

    def test_dry_run_drafts_nothing_real(self):
        campaign = _campaign()
        result, started = self._run_with([campaign], dry_run=True)
        self.assertEqual(result["drafted"], 1)
        started["outlook_draft"].assert_not_called()
        started["gmail_draft"].assert_not_called()
        started["save_ledger"].assert_not_called()
        self.assertEqual(len(result["previews"]), 1)
        self.assertEqual(result["previews"][0]["to"], "owner@example.com")

    def test_real_run_creates_drafts_and_saves_ledger(self):
        campaign = _campaign()
        result, started = self._run_with([campaign], dry_run=False)
        self.assertEqual(result["drafted"], 1)
        started["outlook_draft"].assert_called_once()
        started["gmail_draft"].assert_called_once()
        started["save_ledger"].assert_called_once()

    def test_gmail_auth_down_still_creates_outlook_draft(self):
        campaign = _campaign()
        patches = self._patches([campaign])
        started = {name: p.start() for name, p in patches.items()}
        for p in patches.values():
            self.addCleanup(p.stop)
        started["source_cls"].return_value.get_active_campaigns.return_value = [campaign]
        started["gmail_cls"].side_effect = GmailAuthRequired("bad refresh token")
        result = m.main(dry_run=False)
        self.assertEqual(result["drafted"], 1)
        started["outlook_draft"].assert_called_once()
        started["gmail_draft"].assert_not_called()

    def test_one_campaign_failing_does_not_stop_the_others(self):
        good = _campaign(property_id="1", address="1 Good St")
        bad = _campaign(property_id="2", address="2 Bad St")
        patches = self._patches([good, bad])
        started = {name: p.start() for name, p in patches.items()}
        for p in patches.values():
            self.addCleanup(p.stop)
        started["source_cls"].return_value.get_active_campaigns.return_value = [bad, good]

        def flaky_resolve(outlook, campaign):
            if campaign.address == "2 Bad St":
                raise RuntimeError("Outlook hiccup")
            return "owner@example.com", "Sandra"
        started["resolve_contact"].side_effect = flaky_resolve

        result = m.main(dry_run=False)
        self.assertEqual(result["drafted"], 1)
        self.assertEqual(result["failed"], 1)

    def test_missing_env_vars_exits_before_any_connection(self):
        with patch.dict(os.environ, {"EAGLE_CLIENT_ID": ""}, clear=False):
            with self.assertRaises(SystemExit):
                m.main(dry_run=True)


if __name__ == "__main__":
    unittest.main()
