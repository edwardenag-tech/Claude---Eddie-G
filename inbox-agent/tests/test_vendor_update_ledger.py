"""Tests for vendor_update_ledger.py: load/save round-tripping and the
"missing file looks like an empty ledger" contract new campaigns rely on.
"""

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import vendor_update_ledger as m


class TestLedgerRoundTrip(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.path = os.path.join(self.tmp, "ledger.json")

    def test_missing_file_returns_empty_ledger(self):
        self.assertEqual(m.load_ledger(self.path), {})

    def test_corrupt_file_returns_empty_ledger(self):
        with open(self.path, "w") as fh:
            fh.write("not json")
        self.assertEqual(m.load_ledger(self.path), {})

    def test_save_then_load_round_trips(self):
        ledger = {
            "1825106": m.LedgerEntry(
                address="207/490 Pacific Highway, St Leonards",
                week_number=7,
                known_leads=[
                    m.Lead(name="The Marketing Agency", status="No further interest", last_updated="2026-09-14"),
                    m.Lead(name="The Art School", status="Following up", last_updated="2026-09-14"),
                ],
                last_drafted_at="2026-09-14T02:54:16",
            ),
        }
        m.save_ledger(ledger, self.path)
        loaded = m.load_ledger(self.path)
        self.assertEqual(set(loaded), {"1825106"})
        entry = loaded["1825106"]
        self.assertEqual(entry.address, "207/490 Pacific Highway, St Leonards")
        self.assertEqual(entry.week_number, 7)
        self.assertEqual(len(entry.known_leads), 2)
        self.assertEqual(entry.known_leads[0].name, "The Marketing Agency")
        self.assertEqual(entry.known_leads[0].status, "No further interest")
        self.assertEqual(entry.last_drafted_at, "2026-09-14T02:54:16")

    def test_entry_with_no_leads_round_trips(self):
        ledger = {"1": m.LedgerEntry(address="1 Test St")}
        m.save_ledger(ledger, self.path)
        loaded = m.load_ledger(self.path)
        self.assertEqual(loaded["1"].known_leads, [])
        self.assertEqual(loaded["1"].week_number, 0)


if __name__ == "__main__":
    unittest.main()
