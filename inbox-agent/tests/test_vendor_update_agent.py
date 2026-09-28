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
        self.assertIn("calls", prompt.lower())


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
