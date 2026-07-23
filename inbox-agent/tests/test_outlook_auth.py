"""Tests for outlook_client.OutlookClient's silent-auth / interactive-fallback logic.

Covers the scope-split fix (full scopes -> core mail scopes -> interactive-only-
if-allowed) using a fully mocked msal module -- no real network calls, no real
device-code flow is ever started.
"""

import os
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from outlook_client import CORE_SCOPES, SCOPES, OutlookAuthRequired, OutlookClient

ACCOUNT = {"username": "edward@ibproperty.com.au"}


def _mock_app(accounts, silent_results, device_flow_init=None, device_flow_result=None):
    app = MagicMock()
    app.get_accounts.return_value = accounts
    app.acquire_token_silent.side_effect = silent_results
    if device_flow_init is not None:
        app.initiate_device_flow.return_value = device_flow_init
    if device_flow_result is not None:
        app.acquire_token_by_device_flow.return_value = device_flow_result
    return app


class TestSilentAuthScopeSplit(unittest.TestCase):
    @patch("outlook_client.msal")
    def test_full_scope_silent_success_skips_fallback_and_device_flow(self, mock_msal):
        cache = MagicMock(has_state_changed=False)
        mock_msal.SerializableTokenCache.return_value = cache
        app = _mock_app(accounts=[ACCOUNT], silent_results=[{"access_token": "FULL_TOKEN"}])
        mock_msal.PublicClientApplication.return_value = app

        client = OutlookClient(
            "cid", "tid", token_cache_path="/nonexistent/cache.json", allow_interactive=False
        )

        self.assertEqual(client._access_token, "FULL_TOKEN")
        app.acquire_token_silent.assert_called_once_with(SCOPES, account=ACCOUNT)
        app.initiate_device_flow.assert_not_called()

    @patch("outlook_client.msal")
    def test_full_scope_fails_core_scope_silent_succeeds(self, mock_msal):
        """The exact scenario from today's regression: a token cache granted
        before Calendars.Read existed must still refresh mail silently."""
        cache = MagicMock(has_state_changed=False)
        mock_msal.SerializableTokenCache.return_value = cache
        app = _mock_app(
            accounts=[ACCOUNT],
            silent_results=[None, {"access_token": "CORE_TOKEN"}],
        )
        mock_msal.PublicClientApplication.return_value = app

        client = OutlookClient(
            "cid", "tid", token_cache_path="/nonexistent/cache.json", allow_interactive=False
        )

        self.assertEqual(client._access_token, "CORE_TOKEN")
        self.assertEqual(app.acquire_token_silent.call_count, 2)
        first_call, second_call = app.acquire_token_silent.call_args_list
        self.assertEqual(first_call.args[0], SCOPES)
        self.assertEqual(second_call.args[0], CORE_SCOPES)
        app.initiate_device_flow.assert_not_called()

    @patch("outlook_client.msal")
    def test_both_silent_fail_non_interactive_raises_without_touching_device_flow(self, mock_msal):
        cache = MagicMock()
        mock_msal.SerializableTokenCache.return_value = cache
        app = _mock_app(accounts=[ACCOUNT], silent_results=[None, None])
        mock_msal.PublicClientApplication.return_value = app

        with self.assertRaises(OutlookAuthRequired):
            OutlookClient(
                "cid", "tid", token_cache_path="/nonexistent/cache.json", allow_interactive=False
            )

        app.initiate_device_flow.assert_not_called()
        app.acquire_token_by_device_flow.assert_not_called()

    @patch("outlook_client.msal")
    def test_no_cached_accounts_non_interactive_raises_without_device_flow(self, mock_msal):
        """Empty/missing cache (e.g. first run) must behave the same as an
        exhausted cache -- no silent calls attempted, no device flow, clean
        OutlookAuthRequired instead of a hang."""
        cache = MagicMock()
        mock_msal.SerializableTokenCache.return_value = cache
        app = _mock_app(accounts=[], silent_results=[])
        mock_msal.PublicClientApplication.return_value = app

        with self.assertRaises(OutlookAuthRequired):
            OutlookClient(
                "cid", "tid", token_cache_path="/nonexistent/cache.json", allow_interactive=False
            )

        app.acquire_token_silent.assert_not_called()
        app.initiate_device_flow.assert_not_called()

    @patch("outlook_client.msal")
    def test_both_silent_fail_interactive_allowed_uses_device_flow(self, mock_msal):
        """Only the explicit --auth path (allow_interactive=True) may reach
        the device-code flow."""
        cache = MagicMock(has_state_changed=False)
        mock_msal.SerializableTokenCache.return_value = cache
        app = _mock_app(
            accounts=[ACCOUNT],
            silent_results=[None, None],
            device_flow_init={"user_code": "ABC-123", "message": "go to https://microsoft.com/devicelogin"},
            device_flow_result={"access_token": "DEVICE_TOKEN"},
        )
        mock_msal.PublicClientApplication.return_value = app

        client = OutlookClient(
            "cid", "tid", token_cache_path="/nonexistent/cache.json", allow_interactive=True
        )

        self.assertEqual(client._access_token, "DEVICE_TOKEN")
        app.initiate_device_flow.assert_called_once_with(scopes=SCOPES)
        app.acquire_token_by_device_flow.assert_called_once()

    @patch("outlook_client.msal")
    def test_device_flow_init_failure_raises_runtime_error(self, mock_msal):
        cache = MagicMock(has_state_changed=False)
        mock_msal.SerializableTokenCache.return_value = cache
        app = _mock_app(
            accounts=[ACCOUNT],
            silent_results=[None, None],
            device_flow_init={"error_description": "app registration misconfigured"},
        )
        mock_msal.PublicClientApplication.return_value = app

        with self.assertRaises(RuntimeError):
            OutlookClient(
                "cid", "tid", token_cache_path="/nonexistent/cache.json", allow_interactive=True
            )

    @patch("outlook_client.msal")
    def test_cache_persisted_when_state_changed(self, mock_msal):
        cache = MagicMock(has_state_changed=True)
        cache.serialize.return_value = '{"fake":"cache"}'
        mock_msal.SerializableTokenCache.return_value = cache
        app = _mock_app(accounts=[ACCOUNT], silent_results=[{"access_token": "FULL_TOKEN"}])
        mock_msal.PublicClientApplication.return_value = app

        with tempfile.TemporaryDirectory() as tmp_dir:
            cache_path = os.path.join(tmp_dir, "cache.json")
            OutlookClient("cid", "tid", token_cache_path=cache_path, allow_interactive=False)
            with open(cache_path) as fh:
                self.assertEqual(fh.read(), '{"fake":"cache"}')


if __name__ == "__main__":
    unittest.main()
