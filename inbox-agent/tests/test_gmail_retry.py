"""Tests for gmail_client's rate-limit retry/backoff and partial-fetch behaviour.

Covers the Sep 2026 failure where one 403 rateLimitExceeded mid-fetch made
get_messages() return [] (discarding everything already fetched) and the
briefing send died on the same per-minute quota. No real network calls; the
Gmail service and time.sleep are mocked.
"""

import os
import sys
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from googleapiclient.errors import HttpError

import gmail_client
from gmail_client import GmailClient


def _http_error(status, reason="rateLimitExceeded"):
    resp = MagicMock(status=status, reason="err")
    content = ('{"error": {"errors": [{"reason": "%s"}]}}' % reason).encode()
    return HttpError(resp, content)


def _client():
    client = GmailClient.__new__(GmailClient)  # skip OAuth in __init__
    client.service = MagicMock()
    client._label_cache = {}
    return client


def _request(*outcomes):
    """A request whose execute() yields each outcome in turn (raising exceptions)."""
    req = MagicMock()
    req.execute.side_effect = list(outcomes)
    return req


@patch("gmail_client.time.sleep")
class TestExecuteRetry(unittest.TestCase):
    def test_retries_rate_limit_then_succeeds(self, sleep):
        req = _request(_http_error(403), _http_error(403), {"ok": True})
        self.assertEqual(GmailClient._execute(req), {"ok": True})
        self.assertEqual(req.execute.call_count, 3)
        self.assertEqual([c.args[0] for c in sleep.call_args_list], list(gmail_client._RETRY_DELAYS[:2]))

    def test_retries_429_and_5xx(self, sleep):
        for status in (429, 503):
            req = _request(_http_error(status, reason="backendError"), {"ok": True})
            self.assertEqual(GmailClient._execute(req), {"ok": True})

    def test_permanent_403_is_not_retried(self, sleep):
        req = _request(_http_error(403, reason="insufficientPermissions"))
        with self.assertRaises(HttpError):
            GmailClient._execute(req)
        self.assertEqual(req.execute.call_count, 1)
        sleep.assert_not_called()

    def test_gives_up_and_raises_after_all_delays(self, sleep):
        n = len(gmail_client._RETRY_DELAYS) + 1
        req = _request(*[_http_error(403)] * n)
        with self.assertRaises(HttpError):
            GmailClient._execute(req)
        self.assertEqual(req.execute.call_count, n)
        self.assertEqual(sleep.call_count, len(gmail_client._RETRY_DELAYS))


@patch("gmail_client.time.sleep")
class TestGetMessages(unittest.TestCase):
    def _wire(self, client, stubs, get_requests):
        messages = client.service.users.return_value.messages.return_value
        messages.list.return_value = _request({"messages": [{"id": s} for s in stubs]})
        messages.get.side_effect = get_requests

    def test_unfetchable_message_is_skipped_not_fatal(self, sleep):
        client = _client()
        permanent = _request(_http_error(404, reason="notFound"))
        self._wire(client, ["a", "b", "c"], [_request({"id": "a"}), permanent, _request({"id": "c"})])
        self.assertEqual([m["id"] for m in client.get_messages("q")], ["a", "c"])

    def test_list_failure_returns_empty(self, sleep):
        client = _client()
        messages = client.service.users.return_value.messages.return_value
        messages.list.return_value = _request(_http_error(403, reason="insufficientPermissions"))
        self.assertEqual(client.get_messages("q"), [])


@patch("gmail_client.time.sleep")
class TestSendEmail(unittest.TestCase):
    def test_send_survives_transient_rate_limit(self, sleep):
        client = _client()
        messages = client.service.users.return_value.messages.return_value
        messages.send.return_value = _request(_http_error(403), {"id": "sent"})
        self.assertTrue(client.send_email(["a@example.com"], "subj", "<p>hi</p>"))

    def test_send_returns_false_on_permanent_failure(self, sleep):
        client = _client()
        messages = client.service.users.return_value.messages.return_value
        messages.send.return_value = _request(_http_error(403, reason="insufficientPermissions"))
        self.assertFalse(client.send_email(["a@example.com"], "subj", "<p>hi</p>"))


if __name__ == "__main__":
    unittest.main()
