"""Tests for draft_agent's scheduled-run wiring: dry-run safety, dedupe, live-doc
precedence, per-email isolation, and agent.run_draft_step's never-raises contract.

Everything external (Outlook, Gmail, Docs, Anthropic) is mocked -- no network.
"""

import os
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import anthropic

import agent
import draft_agent
from outlook_client import OutlookClient, OutlookAuthRequired

ENV = {"ANTHROPIC_API_KEY": "k", "AZURE_CLIENT_ID": "c", "AZURE_TENANT_ID": "t"}


def _raw(n, subject="Enquiry 12 Smith Street Sydney", sender="buyer@example.com"):
    return {
        "id": f"id{n}", "internetMessageId": f"<msg{n}@x>", "conversationId": f"conv{n}",
        "subject": subject, "bodyPreview": "hi",
        "from": {"emailAddress": {"address": sender, "name": "Bob Buyer"}},
        "body": {"content": "<p>Interested in 12 Smith Street Sydney</p>"},
    }


class DraftMainCase(unittest.TestCase):
    """Runs draft_agent.main() with every collaborator patched."""

    def setUp(self):
        tmp = tempfile.mkdtemp()
        self.state_path = os.path.join(tmp, "state.json")
        self.outlook = MagicMock()
        self.outlook.get_recent_emails.return_value = [_raw(1)]
        self.gmail = MagicMock()
        patches = {
            "OutlookClient": patch("draft_agent.OutlookClient"),
            "GmailClient": patch("draft_agent.GmailClient", return_value=self.gmail),
            "DocsClient": patch("draft_agent.DocsClient"),
            "Anthropic": patch("draft_agent.anthropic.Anthropic"),
            "examples": patch("draft_agent.fetch_sent_enquiry_examples", return_value=[]),
            "replied": patch("draft_agent.outlook_already_replied", return_value=False),
            "related": patch("draft_agent.outlook_search_related", return_value=[]),
            "sent_reply": patch("draft_agent.fetch_sent_reply_for_address", return_value=None),
            "collect": patch("draft_agent.collect_property_data", return_value={"address": "12 Smith Street"}),
            "classify": patch("draft_agent.claude_classify", return_value=("sale_enquiry", True)),
            "draft": patch("draft_agent.claude_draft_reply", return_value=("Re: x", "<p>Hi Bob</p>")),
            "o_draft": patch("draft_agent.outlook_create_draft", return_value="OD1"),
            "g_draft": patch("draft_agent.gmail_create_draft", return_value="GD1"),
            "listings": patch("draft_agent.load_listings_db", return_value={"listings": []}),
            "campaigns": patch("draft_agent.load_campaigns_from_doc", return_value=[]),
            "state": patch("draft_agent._STATE_PATH", self.state_path),
            "env": patch.dict(os.environ, ENV),
        }
        self.m = {k: p.start() for k, p in patches.items()}
        for p in patches.values():
            self.addCleanup(p.stop)
        self.m["OutlookClient"].return_value = self.outlook
        self.m["OutlookClient"].extract_email_data = OutlookClient.extract_email_data

    def test_dry_run_saves_nothing_and_records_nothing(self):
        summary = draft_agent.main(dry_run=True)
        self.assertEqual(summary["drafted"], 1)
        self.m["o_draft"].assert_not_called()
        self.m["g_draft"].assert_not_called()
        self.outlook.send_reply.assert_not_called()
        self.assertFalse(os.path.exists(self.state_path))

    def test_real_run_creates_both_drafts_and_never_sends(self):
        summary = draft_agent.main(allow_auto_send=False)
        self.assertEqual(summary["drafted"], 1)
        self.m["o_draft"].assert_called_once()
        self.m["g_draft"].assert_called_once()
        self.outlook.send_reply.assert_not_called()

    def test_forced_draft_only_beats_auto_send_env(self):
        with patch("draft_agent.AUTO_SEND_ENABLED", True):
            draft_agent.main(allow_auto_send=False)
        self.outlook.send_reply.assert_not_called()

    def test_same_email_is_not_drafted_twice_across_runs(self):
        draft_agent.main(allow_auto_send=False)
        second = draft_agent.main(allow_auto_send=False)
        self.assertEqual(second["drafted"], 0)
        self.assertEqual(second["skipped"], 1)
        self.assertEqual(self.m["o_draft"].call_count, 1)

    def test_dedupe_survives_the_email_moving_folders(self):
        draft_agent.main(allow_auto_send=False)
        moved = _raw(1)
        moved["id"] = "different-id-after-move"
        self.outlook.get_recent_emails.return_value = [moved]
        self.assertEqual(draft_agent.main(allow_auto_send=False)["drafted"], 0)

    def test_campaign_doc_entry_outranks_past_sent_reply(self):
        self.m["campaigns"].return_value = [{"address": "12 Smith Street, Sydney"}]
        with patch("draft_agent._extract_address", return_value="12 Smith Street, Sydney"):
            draft_agent.main(allow_auto_send=False)
        self.m["sent_reply"].assert_not_called()
        self.m["collect"].assert_called_once()

    def test_past_sent_reply_still_used_when_doc_has_no_entry(self):
        self.m["sent_reply"].return_value = "old reply"
        with patch("draft_agent._extract_address", return_value="12 Smith Street, Sydney"):
            draft_agent.main(allow_auto_send=False)
        self.m["sent_reply"].assert_called_once()
        self.m["collect"].assert_not_called()

    def test_one_failing_email_does_not_stop_the_rest(self):
        self.outlook.get_recent_emails.return_value = [_raw(1), _raw(2)]
        self.m["draft"].side_effect = [draft_agent.DraftGenerationError("boom"), ("Re: y", "<p>ok</p>")]
        summary = draft_agent.main(allow_auto_send=False)
        self.assertEqual((summary["failed"], summary["drafted"]), (1, 1))
        # the failed one must stay eligible for the next run
        self.assertNotIn("<msg1@x>", draft_agent.load_drafted_state())
        self.assertIn("<msg2@x>", draft_agent.load_drafted_state())

    def test_rejected_api_key_stops_the_run_instead_of_silently_drafting_nothing(self):
        err = anthropic.AuthenticationError("bad key", response=MagicMock(status_code=401), body=None)
        self.m["classify"].side_effect = err
        with self.assertRaises(anthropic.AuthenticationError):
            draft_agent.main(allow_auto_send=False)

    def test_gmail_down_still_creates_outlook_draft(self):
        self.m["GmailClient"].side_effect = draft_agent.GmailAuthRequired("dead")
        summary = draft_agent.main(allow_auto_send=False)
        self.assertEqual(summary["drafted"], 1)
        self.m["o_draft"].assert_called_once()
        self.m["g_draft"].assert_not_called()


class TestClassifyAuthPassthrough(unittest.TestCase):
    def test_auth_error_propagates_other_errors_are_swallowed(self):
        ai = MagicMock()
        ai.messages.create.side_effect = anthropic.AuthenticationError(
            "bad", response=MagicMock(status_code=401), body=None)
        with self.assertRaises(anthropic.AuthenticationError):
            draft_agent.claude_classify(ai, {})
        ai.messages.create.side_effect = RuntimeError("timeout")
        self.assertEqual(draft_agent.claude_classify(ai, {}), ("general", False))


class TestDraftedState(unittest.TestCase):
    def test_old_entries_are_pruned(self):
        path = os.path.join(tempfile.mkdtemp(), "s.json")
        draft_agent.save_drafted_state({"old": "2000-01-01", "new": "2999-01-01"}, path)
        self.assertEqual(draft_agent.load_drafted_state(path), {"new": "2999-01-01"})

    def test_missing_or_corrupt_file_is_empty(self):
        path = os.path.join(tempfile.mkdtemp(), "s.json")
        self.assertEqual(draft_agent.load_drafted_state(path), {})
        open(path, "w").write("{not json")
        self.assertEqual(draft_agent.load_drafted_state(path), {})


class TestRunDraftStep(unittest.TestCase):
    def test_skipped_without_outlook(self):
        alerts = []
        with patch("draft_agent.main") as m:
            self.assertIsNone(agent.run_draft_step(None, alerts))
        m.assert_not_called()
        self.assertEqual(alerts, [])

    def test_forces_draft_only_and_reports_waiting_drafts(self):
        alerts = []
        with patch("draft_agent.main", return_value={"drafted": 2, "failed": 1}) as m:
            agent.run_draft_step(MagicMock(), alerts, dry_run=False)
        m.assert_called_once_with(dry_run=False, allow_auto_send=False)
        self.assertTrue(any("2 enquiry reply drafts waiting" in a for a in alerts))
        self.assertTrue(any("1 email(s) could not be drafted" in a for a in alerts))

    def test_dry_run_adds_no_alerts(self):
        alerts = []
        with patch("draft_agent.main", return_value={"drafted": 3, "failed": 0}):
            agent.run_draft_step(MagicMock(), alerts, dry_run=True)
        self.assertEqual(alerts, [])

    def test_system_exit_from_draft_agent_does_not_abort_the_run(self):
        alerts = []
        with patch("draft_agent.main", side_effect=SystemExit(1)):
            self.assertIsNone(agent.run_draft_step(MagicMock(), alerts))
        self.assertEqual(len(alerts), 1)

    def test_failures_become_alerts_not_exceptions(self):
        alerts = []
        with patch("draft_agent.main", side_effect=OutlookAuthRequired("dead")):
            agent.run_draft_step(MagicMock(), alerts)
        self.assertIn("dead", alerts[0])

    def test_rejected_key_gets_a_readable_alert(self):
        alerts = []
        err = anthropic.AuthenticationError("x", response=MagicMock(status_code=401), body=None)
        with patch("draft_agent.main", side_effect=err):
            agent.run_draft_step(MagicMock(), alerts)
        self.assertIn("ANTHROPIC_API_KEY", alerts[0])


if __name__ == "__main__":
    unittest.main()
