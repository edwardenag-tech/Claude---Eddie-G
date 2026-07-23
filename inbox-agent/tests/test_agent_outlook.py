"""Tests for agent.py's Outlook wiring: the scheduled/unattended path must
never be able to reach the interactive device-code flow, while the explicit
`--auth` path still can.

Imports the real agent module (harmless at import time -- just logging/config
setup, no network calls) and mocks agent.OutlookClient directly so no real
MSAL/Graph calls happen.
"""

import os
import sys
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import agent
from outlook_client import OutlookAuthRequired

CFG = {
    "azure_client_id": "cid",
    "azure_tenant_id": "tid",
    "msal_token_cache_path": "/nonexistent/cache.json",
}


class TestBuildOutlookClient(unittest.TestCase):
    @patch("agent.OutlookClient")
    def test_default_call_is_non_interactive(self, mock_cls):
        """The scheduled path calls build_outlook_client(cfg) with no extra
        args -- this must resolve to allow_interactive=False."""
        mock_cls.return_value = MagicMock()

        agent.build_outlook_client(CFG)

        mock_cls.assert_called_once_with(
            client_id="cid",
            tenant_id="tid",
            token_cache_path="/nonexistent/cache.json",
            allow_interactive=False,
        )

    @patch("agent.OutlookClient")
    def test_explicit_auth_mode_is_interactive(self, mock_cls):
        mock_cls.return_value = MagicMock()

        agent.build_outlook_client(CFG, allow_interactive=True)

        mock_cls.assert_called_once_with(
            client_id="cid",
            tenant_id="tid",
            token_cache_path="/nonexistent/cache.json",
            allow_interactive=True,
        )

    @patch("agent.OutlookClient")
    def test_auth_required_in_scheduled_context_returns_none_not_raise(self, mock_cls):
        """This is the core regression fix: when silent auth can't satisfy any
        scope set and interactive isn't allowed, the scheduled run must
        degrade to Gmail-only instead of crashing or hanging."""
        mock_cls.side_effect = OutlookAuthRequired("no silent token available")

        result = agent.build_outlook_client(CFG)  # default allow_interactive=False

        self.assertIsNone(result)

    @patch("agent.OutlookClient")
    def test_other_exceptions_still_return_none(self, mock_cls):
        mock_cls.side_effect = RuntimeError("network blip")

        result = agent.build_outlook_client(CFG)

        self.assertIsNone(result)

    @patch("agent.OutlookClient")
    def test_success_returns_client(self, mock_cls):
        instance = MagicMock()
        mock_cls.return_value = instance

        result = agent.build_outlook_client(CFG)

        self.assertIs(result, instance)


if __name__ == "__main__":
    unittest.main()
