"""Tests for the 'Needs attention' banner in briefing.build_briefing_html."""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from briefing import build_briefing_html


def _build(**kwargs):
    return build_briefing_html(
        todo_md="- item", gmail_new=[], outlook_new=[], cleaning_report={"actions": []},
        date_str="Sunday, 20 September 2026", **kwargs,
    )


class TestAlertsBanner(unittest.TestCase):
    def test_no_banner_without_alerts(self):
        self.assertNotIn("Needs attention", _build())
        self.assertNotIn("Needs attention", _build(alerts=[]))

    def test_banner_lists_alerts_and_escapes_html(self):
        html_out = _build(alerts=["Outlook is not connected", "<b>bad</b> & worse"])
        self.assertIn("Needs attention", html_out)
        self.assertIn("Outlook is not connected", html_out)
        self.assertIn("&lt;b&gt;bad&lt;/b&gt; &amp; worse", html_out)


if __name__ == "__main__":
    unittest.main()
