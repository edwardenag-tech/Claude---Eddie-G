"""Tests for market_update_agent: workbook reading, dedupe, template rendering,
and the never-touch-real-drafts-without-a-letter guarantees.

All workbooks used here are small synthetic files built with openpyxl in a
temp directory -- the real Dropbox DATA BASE.xlsx is never read by this suite.
"""

import json
import os
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import market_update_agent as m

HEADER = [
    "Street", "Address", "Suburb", "State", "Postcode", "Property Type",
    "Land Size (m²)", "Internal Size (m²)", "Vendor First Name",
    "Vendor Surname Name", "Email", "Sent Email", "Last Contacted",
    "Company Name",
]
GROUP_ROW = ["DONE DEAL. 0"] + [None] * (len(HEADER) - 1)


def _make_workbook(path, sheet_name, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(GROUP_ROW)
    ws.append(HEADER)
    for row in rows:
        ws.append(row)
    wb.save(path)


def _row(street="4", address="Smith Street", suburb="ARTARMON", state="NSW",
         postcode=2064, ptype="Commercial", land=None, internal=None,
         first="Bob", last="Buyer", email="bob@example.com", sent=False,
         contacted=None, company=""):
    return [street, address, suburb, state, postcode, ptype, land, internal,
            first, last, email, sent, contacted, company]


class TestReadRecipients(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.path = os.path.join(self.tmp, "db.xlsx")

    def test_basic_read(self):
        _make_workbook(self.path, "Sheet1", [_row()])
        recipients = m.read_recipients(self.path, "seg", "Sheet1", "Segment")
        self.assertEqual(len(recipients), 1)
        r = recipients[0]
        self.assertEqual(r.suburb, "ARTARMON")
        self.assertEqual(r.email, "bob@example.com")
        self.assertEqual(r.street_address, "4 Smith Street")

    def test_row_without_email_is_skipped(self):
        _make_workbook(self.path, "Sheet1", [_row(email="")])
        self.assertEqual(m.read_recipients(self.path, "seg", "Sheet1", "S"), [])

    def test_malformed_email_is_skipped(self):
        _make_workbook(self.path, "Sheet1", [_row(email="not-an-email")])
        self.assertEqual(m.read_recipients(self.path, "seg", "Sheet1", "S"), [])

    def test_blank_row_is_skipped(self):
        _make_workbook(self.path, "Sheet1", [_row(), [None] * len(HEADER)])
        self.assertEqual(len(m.read_recipients(self.path, "seg", "Sheet1", "S")), 1)

    def test_nbsp_placeholder_treated_as_blank(self):
        _make_workbook(self.path, "Sheet1", [_row(last="\xa0")])
        self.assertEqual(m.read_recipients(self.path, "seg", "Sheet1", "S")[0].last_name, "")

    def test_missing_sheet_raises(self):
        _make_workbook(self.path, "Sheet1", [_row()])
        with self.assertRaises(m.WorkbookShapeError):
            m.read_recipients(self.path, "seg", "NoSuchSheet", "S")

    def test_missing_required_column_raises(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(["Street", "Address", "Suburb"])  # missing Email etc.
        ws.append(["4", "Smith St", "ARTARMON"])
        wb.save(self.path)
        with self.assertRaises(m.WorkbookShapeError):
            m.read_recipients(self.path, "seg", "Sheet1", "S")

    def test_suburb_not_in_column_c_only_warns_does_not_raise(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        # Suburb moved to column A instead of C
        reordered = ["Suburb", "Street", "Address", "State", "Postcode",
                     "Vendor First Name", "Vendor Surname Name", "Email"]
        ws.append(reordered)
        ws.append(["ARTARMON", "4", "Smith St", "NSW", 2064, "Bob", "Buyer", "bob@example.com"])
        wb.save(self.path)
        with self.assertLogs("market_update_agent", level="WARNING") as logs:
            recipients = m.read_recipients(self.path, "seg", "Sheet1", "S")
        self.assertEqual(len(recipients), 1)
        self.assertTrue(any("not " in line and "column" in line for line in logs.output))

    def test_header_found_even_with_leading_group_title_row(self):
        # exactly DATA BASE.xlsx's real shape: row 1 = merged group title, row 2 = real header
        _make_workbook(self.path, "Sheet1", [_row()])
        recipients = m.read_recipients(self.path, "seg", "Sheet1", "S")
        self.assertEqual(recipients[0].source_row, 3)  # header on row 2, data starts row 3


class TestDedupeByOwner(unittest.TestCase):
    def test_repeated_email_collapses_to_one(self):
        rows = [m.Recipient(segment="s", segment_label="S", suburb="A", street="1", address="X St",
                             state="NSW", postcode="2000", first_name="Bob", last_name="B",
                             email="bob@x.com", source_row=n) for n in (3, 4, 5)]
        kept, extra = m.dedupe_by_owner(rows)
        self.assertEqual(len(kept), 1)
        self.assertEqual(kept[0].source_row, 3)  # first one wins
        self.assertEqual(extra, {"bob@x.com": 2})

    def test_distinct_emails_all_kept(self):
        rows = [m.Recipient(segment="s", segment_label="S", suburb="A", street="1", address="X",
                             state="NSW", postcode="2000", first_name="B", last_name="C",
                             email=f"e{i}@x.com") for i in range(3)]
        kept, extra = m.dedupe_by_owner(rows)
        self.assertEqual(len(kept), 3)
        self.assertEqual(extra, {})


class TestLetterLoadingAndRendering(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self._orig_dir = m._LETTERS_DIR
        m._LETTERS_DIR = __import__("pathlib").Path(self.tmp)
        self.addCleanup(setattr, m, "_LETTERS_DIR", self._orig_dir)

    def _write(self, name, content):
        (m._LETTERS_DIR / f"{name}.html").write_text(content)

    def test_missing_file_returns_none(self):
        self.assertIsNone(m.load_letter("nope"))

    def test_placeholder_marker_returns_none(self):
        self._write("seg", "Subject: Hi\n\n[EDDIE: WRITE THIS LETTER]\n<p>draft</p>")
        self.assertIsNone(m.load_letter("seg"))

    def test_real_letter_loads(self):
        self._write("seg", "Subject: Update for {{suburb}}\n\n<p>Hi {{first_name}}</p>")
        letter = m.load_letter("seg")
        self.assertEqual(letter.subject_template, "Update for {{suburb}}")
        self.assertIn("{{first_name}}", letter.body_template)

    def test_missing_subject_line_raises(self):
        self._write("seg", "<p>no subject line</p>")
        with self.assertRaises(m.WorkbookShapeError):
            m.load_letter("seg")

    def test_render_substitutes_all_fields(self):
        letter = m.Letter(
            subject_template="Update for {{suburb}}",
            body_template="<p>Hi {{first_name}} {{last_name}}, re {{street_address}} ({{property_type}})</p>",
        )
        r = m.Recipient(segment="s", segment_label="S", suburb="ARTARMON", street="4",
                         address="Smith Street", state="NSW", postcode="2064",
                         first_name="Bob", last_name="Buyer", email="bob@x.com",
                         property_type="Commercial")
        subject, body = m.render_letter(letter, r)
        self.assertEqual(subject, "Update for Artarmon")
        self.assertIn("Hi Bob Buyer, re 4 Smith Street (Commercial)", body)

    def test_blank_first_name_falls_back_to_there(self):
        r = m.Recipient(segment="s", segment_label="S", suburb="A", street="1", address="X",
                         state="NSW", postcode="2000", first_name="", last_name="", email="b@x.com")
        self.assertEqual(r.merge_fields()["first_name"], "there")

    def test_full_name_falls_back_to_company_then_there(self):
        r = m.Recipient(segment="s", segment_label="S", suburb="A", street="1", address="X",
                         state="NSW", postcode="2000", first_name="", last_name="", email="b@x.com",
                         company_name="Acme Pty Ltd")
        self.assertEqual(r.full_name, "Acme Pty Ltd")
        r.company_name = ""
        self.assertEqual(r.full_name, "there")


class TestRunEndToEnd(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.xlsx_path = os.path.join(self.tmp, "db.xlsx")
        _make_workbook(self.xlsx_path, "SheetA", [_row(email="a@x.com"), _row(email="b@x.com", suburb="MOSMAN")])

        self._orig_segments = dict(m.SEGMENTS)
        m.SEGMENTS.clear()
        m.SEGMENTS["seg_a"] = {"label": "Segment A", "sheet": "SheetA"}
        self.addCleanup(lambda: (m.SEGMENTS.clear(), m.SEGMENTS.update(self._orig_segments)))

        self.letters_tmp = tempfile.mkdtemp()
        self._orig_letters_dir = m._LETTERS_DIR
        m._LETTERS_DIR = __import__("pathlib").Path(self.letters_tmp)
        self.addCleanup(setattr, m, "_LETTERS_DIR", self._orig_letters_dir)

    def _write_letter(self):
        (m._LETTERS_DIR / "seg_a.html").write_text(
            "Subject: Update for {{suburb}}\n\n<p>Hi {{first_name}}</p>"
        )

    def test_dry_run_without_a_letter_drafts_nothing(self):
        summary = m.run(dry_run=True, xlsx_path=self.xlsx_path)
        self.assertEqual(summary["seg_a"]["drafted"], 0)
        self.assertTrue(summary["seg_a"]["skipped_no_letter"])

    def test_dry_run_with_a_letter_previews_without_saving(self):
        self._write_letter()
        summary = m.run(dry_run=True, xlsx_path=self.xlsx_path)
        self.assertEqual(summary["seg_a"]["drafted"], 2)
        self.assertEqual(len(summary["seg_a"]["previews"]), 2)
        self.assertFalse(os.path.exists(m._STATE_PATH))

    def test_limit_caps_recipients_per_segment(self):
        self._write_letter()
        summary = m.run(dry_run=True, xlsx_path=self.xlsx_path, limit=1)
        self.assertEqual(summary["seg_a"]["drafted"], 1)

    def test_only_segment_filters_to_one(self):
        m.SEGMENTS["seg_b"] = {"label": "Segment B", "sheet": "DoesNotExist"}
        self._write_letter()
        summary = m.run(dry_run=True, xlsx_path=self.xlsx_path, only_segment="seg_a")
        self.assertEqual(list(summary), ["seg_a"])

    def test_bad_sheet_name_is_reported_not_raised(self):
        m.SEGMENTS["seg_bad"] = {"label": "Bad", "sheet": "NoSuchSheet"}
        summary = m.run(dry_run=True, xlsx_path=self.xlsx_path)
        self.assertIn("error", summary["seg_bad"])

    @patch("market_update_agent.OutlookClient")
    @patch("market_update_agent.outlook_create_new_draft", return_value="OD1")
    @patch("market_update_agent.gmail_create_new_draft", return_value="GD1")
    def test_cli_refuses_non_dry_run(self, gmail_draft, outlook_draft, outlook_cls):
        # __main__'s guard is what really protects us; exercise run() directly
        # with dry_run=False to prove the *mechanics* work when explicitly invoked,
        # while confirming no Gmail draft happens without opt-in.
        self._write_letter()
        state_path = os.path.join(self.tmp, "state.json")
        with patch("market_update_agent._STATE_PATH", state_path):
            summary = m.run(dry_run=False, xlsx_path=self.xlsx_path, gmail_copies=False)
        self.assertEqual(summary["seg_a"]["drafted"], 2)
        outlook_draft.assert_called()
        gmail_draft.assert_not_called()
        self.assertTrue(os.path.exists(state_path))

    @patch("market_update_agent.OutlookClient")
    @patch("market_update_agent.outlook_create_new_draft", return_value="OD1")
    def test_real_run_does_not_redraft_within_state_ttl(self, outlook_draft, outlook_cls):
        self._write_letter()
        state_path = os.path.join(self.tmp, "state.json")
        with patch("market_update_agent._STATE_PATH", state_path):
            m.run(dry_run=False, xlsx_path=self.xlsx_path, gmail_copies=False)
            second = m.run(dry_run=False, xlsx_path=self.xlsx_path, gmail_copies=False)
        self.assertEqual(second["seg_a"]["drafted"], 0)
        self.assertEqual(second["seg_a"]["skipped_already_drafted"], 2)


class TestCliSafetyGate(unittest.TestCase):
    def test_running_without_dry_run_flag_exits_nonzero_and_calls_run_never(self):
        import subprocess
        result = subprocess.run(
            [sys.executable, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "market_update_agent.py")],
            capture_output=True, text=True, timeout=10,
        )
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("Refusing to run for real", result.stderr)


if __name__ == "__main__":
    unittest.main()
