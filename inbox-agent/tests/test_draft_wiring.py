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


ME = "edward@ibproperty.com.au"


def _addrs(*addresses):
    return [{"emailAddress": {"address": a}} for a in addresses]


def _raw(n, subject="Enquiry 12 Smith Street Sydney", sender="buyer@example.com", to=(ME,), cc=()):
    return {
        "toRecipients": _addrs(*to), "ccRecipients": _addrs(*cc),
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
            "classify": patch("draft_agent.claude_classify", return_value=("sale_enquiry", True, "What is the price guide?")),
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

    def test_real_run_creates_outlook_draft_only_by_default_and_never_sends(self):
        summary = draft_agent.main(allow_auto_send=False)
        self.assertEqual(summary["drafted"], 1)
        self.m["o_draft"].assert_called_once()
        self.m["GmailClient"].assert_not_called()
        self.m["g_draft"].assert_not_called()
        self.outlook.send_reply.assert_not_called()

    def test_gmail_copy_created_only_when_explicitly_enabled(self):
        summary = draft_agent.main(allow_auto_send=False, gmail_copies=True)
        self.assertEqual(summary["drafted"], 1)
        self.m["o_draft"].assert_called_once()
        self.m["g_draft"].assert_called_once()

    def test_env_var_can_enable_gmail_copies(self):
        with patch("draft_agent.GMAIL_COPIES_ENABLED", True):
            draft_agent.main(allow_auto_send=False)
        self.m["g_draft"].assert_called_once()

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

    def test_gmail_down_still_creates_outlook_draft_when_copies_enabled(self):
        self.m["GmailClient"].side_effect = draft_agent.GmailAuthRequired("dead")
        summary = draft_agent.main(allow_auto_send=False, gmail_copies=True)
        self.assertEqual(summary["drafted"], 1)
        self.m["o_draft"].assert_called_once()
        self.m["g_draft"].assert_not_called()

    def test_cc_only_email_is_skipped_before_any_ai_call(self):
        self.outlook.get_recent_emails.return_value = [_raw(1, to=("someone@else.com",), cc=(ME,))]
        summary = draft_agent.main(allow_auto_send=False)
        self.assertEqual((summary["drafted"], summary["skipped"]), (0, 1))
        self.m["classify"].assert_not_called()
        self.m["draft"].assert_not_called()

    def test_direct_recipient_among_several_is_still_considered(self):
        self.outlook.get_recent_emails.return_value = [_raw(1, to=(ME, "sarah@ibproperty.com.au"), cc=("boss@x.com",))]
        self.assertEqual(draft_agent.main(allow_auto_send=False)["drafted"], 1)

    def test_email_where_eddie_is_in_neither_field_goes_to_the_classifier(self):
        # alias / Bcc delivery -- must not be dropped by the cheap recipient filter
        self.outlook.get_recent_emails.return_value = [_raw(1, to=("enquiries@ibproperty.com.au",))]
        draft_agent.main(allow_auto_send=False)
        self.m["classify"].assert_called_once()

    def test_no_direct_question_means_no_draft_whatever_the_category(self):
        for category in ("sale_enquiry", "lease_enquiry", "landlord_query", "general"):
            self.m["classify"].return_value = (category, False, "")
            summary = draft_agent.main(allow_auto_send=False)
            self.assertEqual(summary["drafted"], 0, category)
        self.m["draft"].assert_not_called()

    def test_non_enquiry_email_with_a_direct_question_does_get_a_draft(self):
        self.m["classify"].return_value = ("landlord_query", True, "Can you confirm Tuesday's inspection?")
        summary = draft_agent.main(allow_auto_send=False)
        self.assertEqual(summary["drafted"], 1)
        self.assertEqual(self.m["draft"].call_args.kwargs["question"], "Can you confirm Tuesday's inspection?")

    def test_dry_run_summary_carries_the_detected_question(self):
        summary = draft_agent.main(dry_run=True)
        self.assertEqual(summary["drafts"][0]["question"], "What is the price guide?")


class TestClassifyAuthPassthrough(unittest.TestCase):
    def test_auth_error_propagates_other_errors_are_swallowed(self):
        ai = MagicMock()
        ai.messages.create.side_effect = anthropic.AuthenticationError(
            "bad", response=MagicMock(status_code=401), body=None)
        with self.assertRaises(anthropic.AuthenticationError):
            draft_agent.claude_classify(ai, {})
        ai.messages.create.side_effect = RuntimeError("timeout")
        self.assertEqual(draft_agent.claude_classify(ai, {}), ("general", False, ""))


class TestClassifierContract(unittest.TestCase):
    def _classify(self, reply_text, email=None):
        ai = MagicMock()
        ai.messages.create.return_value.content = [MagicMock(text=reply_text)]
        return draft_agent.claude_classify(ai, email or {"subject": "s", "body": "b"}), ai

    def test_yes_with_a_question_is_accepted(self):
        result, _ = self._classify(
            '{"category": "lease_enquiry", "asks_directly": true, "question": "What is the rent?"}')
        self.assertEqual(result, ("lease_enquiry", True, "What is the rent?"))

    def test_yes_without_a_question_is_downgraded_to_no(self):
        result, _ = self._classify('{"category": "general", "asks_directly": true, "question": ""}')
        self.assertEqual(result, ("general", False, ""))

    def test_no_drops_any_stray_question_text(self):
        result, _ = self._classify('{"category": "general", "asks_directly": false, "question": "x?"}')
        self.assertEqual(result, ("general", False, ""))

    def test_garbage_or_fenced_output_fails_closed_or_parses(self):
        self.assertEqual(self._classify("not json")[0], ("general", False, ""))
        fenced = '```json\n{"category": "sale_enquiry", "asks_directly": true, "question": "Price?"}\n```'
        self.assertEqual(self._classify(fenced)[0], ("sale_enquiry", True, "Price?"))

    def test_unknown_category_falls_back_to_general(self):
        result, _ = self._classify('{"category": "weird", "asks_directly": true, "question": "Q?"}')
        self.assertEqual(result[0], "general")

    def test_prompt_shows_recipients_and_uses_the_strong_model(self):
        _, ai = self._classify(
            '{"category": "general", "asks_directly": false, "question": ""}',
            {"to": "a@x.com, b@x.com", "cc": "c@x.com", "subject": "s", "body": "b"},
        )
        kwargs = ai.messages.create.call_args.kwargs
        prompt = kwargs["messages"][0]["content"]
        self.assertIn("To: a@x.com, b@x.com", prompt)
        self.assertIn("Cc: c@x.com", prompt)
        self.assertEqual(kwargs["model"], draft_agent.CLASSIFY_MODEL)


class TestIsCcOnly(unittest.TestCase):
    def test_cases(self):
        f = draft_agent.is_cc_only
        self.assertTrue(f({"to": "x@y.com", "cc": ME}))
        self.assertFalse(f({"to": ME, "cc": ""}))
        self.assertFalse(f({"to": f"x@y.com, {ME}", "cc": "z@y.com"}))
        self.assertFalse(f({"to": ME, "cc": ME}))            # in both -> direct
        self.assertFalse(f({"to": "alias@ibproperty.com.au", "cc": ""}))  # neither
        self.assertFalse(f({}))


class TestGenericDraftPrompt(unittest.TestCase):
    def test_generic_reply_gets_the_question_and_no_invention_rule(self):
        ai = MagicMock()
        ai.messages.create.return_value.content = [MagicMock(text="<p>ok</p>")]
        draft_agent.claude_draft_reply(
            ai, {"subject": "Q", "body": "b", "from": "a@b.com"}, "landlord_query", [],
            question="Can you confirm Tuesday?",
        )
        prompt = ai.messages.create.call_args.kwargs["messages"][0]["content"]
        self.assertIn("Can you confirm Tuesday?", prompt)
        self.assertIn("Do NOT invent", prompt)
        self.assertIn("[EDDIE TO CONFIRM", prompt)

    def test_confirm_placeholder_counts_as_unresolved(self):
        self.assertTrue(draft_agent._PLACEHOLDER_RE.findall("<p>Rent is [EDDIE TO CONFIRM: rent]</p>"))


class TestQuotedThreadStripping(unittest.TestCase):
    def test_outlook_original_message_marker_is_cut(self):
        text = ("Hi Toby, we sold for just under 1 mil.\n\n"
                "From: Someone Else Sent: 20 August 2026 11:34 AM To: Edward\n"
                "Subject: Enquiry ... User Details: Name: Previous Enquirer Email: prev@x.com")
        self.assertEqual(
            draft_agent._strip_quoted_thread(text),
            "Hi Toby, we sold for just under 1 mil.",
        )

    def test_gmail_style_on_wrote_marker_is_cut(self):
        text = "Sure, happy to help.\n\nOn Mon, 1 Sep 2026 at 10:00, Prev Person <prev@x.com> wrote:\n> old question"
        self.assertEqual(draft_agent._strip_quoted_thread(text), "Sure, happy to help.")

    def test_no_marker_returns_the_whole_text(self):
        self.assertEqual(draft_agent._strip_quoted_thread("Just a plain reply."), "Just a plain reply.")

    def test_style_examples_never_carry_a_previous_enquirers_details(self):
        outlook = MagicMock()
        outlook._get.return_value = {"value": [{
            "subject": "Re: Shop 1, 38 Cumberland Street",
            "sentDateTime": "2026-08-20T00:00:00Z",
            "body": {"content": (
                "<p>Hi Toby, IB Property is pleased to bring to market this Property Highlights "
                "Asking Rent listing.</p>"
                "<p>From: Prev Enquirer &lt;prev@example.com&gt; Sent: 1 Jan 2026 To: Edward "
                "Subject: old enquiry User Details: Phone: 0400000000</p>"
            )},
        }]}
        examples = draft_agent.fetch_sent_enquiry_examples(outlook)
        self.assertEqual(len(examples), 1)
        self.assertNotIn("prev@example.com", examples[0])
        self.assertNotIn("0400000000", examples[0])

    def test_sent_reply_template_reuse_never_carries_a_previous_enquirers_details(self):
        outlook = MagicMock()
        outlook._get.return_value = {"value": [{
            "subject": "Re: Enquiry", "sentDateTime": "2026-08-20T00:00:00Z",
            "body": {"content": (
                "<p>Hi Toby, we sold for just under 1 mil, price guide and floor area as discussed.</p>"
                "<p>From: Prev Enquirer &lt;prev@example.com&gt; Sent: 1 Jan 2026 To: Edward</p>"
            )},
        }]}
        template = draft_agent.fetch_sent_reply_for_address(outlook, "12 Smith Street")
        self.assertIsNotNone(template)
        self.assertNotIn("prev@example.com", template)


class TestPortalLeadPrompt(unittest.TestCase):
    def test_prompt_treats_portal_lead_comments_as_a_real_enquiry_channel(self):
        ai = MagicMock()
        ai.messages.create.return_value.content = [MagicMock(text='{"category":"general","asks_directly":false,"question":""}')]
        draft_agent.claude_classify(ai, {"subject": "s", "body": "b"})
        prompt = ai.messages.create.call_args.kwargs["messages"][0]["content"]
        self.assertIn("new lead", prompt.lower())
        self.assertIn("counts as asking Edward directly", prompt)


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
        self.assertTrue(any("2 reply drafts waiting" in a for a in alerts))
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
