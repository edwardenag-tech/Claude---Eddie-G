"""Tests for market_update_agent: workbook reading, dedupe, per-suburb letter
loading and PDF-page embedding, and the never-touch-real-drafts-without-a-
letter guarantees.

All workbooks used here are small synthetic files built with openpyxl in a
temp directory -- the real Dropbox DATA BASE.xlsx is never read by this
suite. Letters are small synthetic PDFs built with fitz (PyMuPDF) -- Eddie's
real branded PDFs are never read by this suite either.
"""

import json
import os
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch

import fitz
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


def _make_pdf(path, pages=1):
    """A minimal real PDF -- exercises actual PyMuPDF rasterization rather
    than a mock, since that's the part most likely to break in practice."""
    doc = fitz.open()
    for _ in range(pages):
        page = doc.new_page()
        page.insert_text((72, 72), "Market update")
    doc.save(str(path))
    doc.close()


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


class TestDedupeByOwnerAndSuburb(unittest.TestCase):
    def test_same_owner_same_suburb_collapses_to_one(self):
        rows = [m.Recipient(segment="s", segment_label="S", suburb="ARTARMON", street="1", address="X St",
                             state="NSW", postcode="2000", first_name="Bob", last_name="B",
                             email="bob@x.com", source_row=n) for n in (3, 4, 5)]
        kept, extra = m.dedupe_by_owner_and_suburb(rows)
        self.assertEqual(len(kept), 1)
        self.assertEqual(kept[0].source_row, 3)  # first one wins
        self.assertEqual(extra, {("bob@x.com", "artarmon"): 2})

    def test_same_owner_different_suburbs_both_kept(self):
        rows = [
            m.Recipient(segment="s", segment_label="S", suburb="ARTARMON", street="1", address="X St",
                        state="NSW", postcode="2000", first_name="Bob", last_name="B",
                        email="bob@x.com", source_row=3),
            m.Recipient(segment="s", segment_label="S", suburb="MOSMAN", street="2", address="Y St",
                        state="NSW", postcode="2088", first_name="Bob", last_name="B",
                        email="bob@x.com", source_row=4),
        ]
        kept, extra = m.dedupe_by_owner_and_suburb(rows)
        self.assertEqual({r.suburb for r in kept}, {"ARTARMON", "MOSMAN"})
        self.assertEqual(extra, {})

    def test_same_owner_two_properties_in_one_suburb_and_one_in_another(self):
        rows = [
            m.Recipient(segment="s", segment_label="S", suburb="ARTARMON", street="1", address="X St",
                        state="NSW", postcode="2000", first_name="Bob", last_name="B",
                        email="bob@x.com", source_row=3),
            m.Recipient(segment="s", segment_label="S", suburb="ARTARMON", street="2", address="Y St",
                        state="NSW", postcode="2000", first_name="Bob", last_name="B",
                        email="bob@x.com", source_row=4),
            m.Recipient(segment="s", segment_label="S", suburb="MOSMAN", street="3", address="Z St",
                        state="NSW", postcode="2088", first_name="Bob", last_name="B",
                        email="bob@x.com", source_row=5),
        ]
        kept, extra = m.dedupe_by_owner_and_suburb(rows)
        self.assertEqual(len(kept), 2)  # one per suburb
        self.assertEqual(extra, {("bob@x.com", "artarmon"): 1})

    def test_suburb_matched_case_insensitively(self):
        rows = [
            m.Recipient(segment="s", segment_label="S", suburb="Artarmon", street="1", address="X",
                        state="NSW", postcode="2000", first_name="B", last_name="C", email="bob@x.com", source_row=3),
            m.Recipient(segment="s", segment_label="S", suburb="ARTARMON", street="2", address="Y",
                        state="NSW", postcode="2000", first_name="B", last_name="C", email="bob@x.com", source_row=4),
        ]
        kept, extra = m.dedupe_by_owner_and_suburb(rows)
        self.assertEqual(len(kept), 1)

    def test_distinct_emails_all_kept(self):
        rows = [m.Recipient(segment="s", segment_label="S", suburb="A", street="1", address="X",
                             state="NSW", postcode="2000", first_name="B", last_name="C",
                             email=f"e{i}@x.com") for i in range(3)]
        kept, extra = m.dedupe_by_owner_and_suburb(rows)
        self.assertEqual(len(kept), 3)
        self.assertEqual(extra, {})


class TestSuburbSlug(unittest.TestCase):
    def test_lowercases(self):
        self.assertEqual(m.suburb_slug("ARTARMON"), "artarmon")

    def test_spaces_become_underscores(self):
        self.assertEqual(m.suburb_slug("St Ives"), "st_ives")

    def test_punctuation_collapsed(self):
        self.assertEqual(m.suburb_slug("Wahroonga -- North"), "wahroonga_north")

    def test_leading_trailing_whitespace_stripped(self):
        self.assertEqual(m.suburb_slug("  Mosman  "), "mosman")


class TestLetterLoadingAndRendering(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self._orig_dir = m._LETTERS_DIR
        m._LETTERS_DIR = __import__("pathlib").Path(self.tmp)
        self.addCleanup(setattr, m, "_LETTERS_DIR", self._orig_dir)

    def _write_pdf(self, slug, pages=1):
        _make_pdf(m._LETTERS_DIR / f"{slug}.pdf", pages=pages)

    def _write_html(self, slug, content):
        (m._LETTERS_DIR / f"{slug}.html").write_text(content)

    def test_no_pdf_returns_none(self):
        self.assertIsNone(m.load_letter("nope"))

    def test_html_without_pdf_still_returns_none(self):
        # The PDF is the letter -- an override file with no PDF is not enough.
        self._write_html("northbridge", "Subject: Hi\n\n<p>Hi {{first_name}}</p>")
        self.assertIsNone(m.load_letter("northbridge"))

    def test_pdf_only_uses_default_subject_and_greeting(self):
        self._write_pdf("northbridge")
        letter = m.load_letter("Northbridge")  # exercises slug matching too
        self.assertEqual(letter.subject_template, m._DEFAULT_SUBJECT_TEMPLATE)
        self.assertEqual(letter.greeting_template, m._DEFAULT_GREETING_TEMPLATE)
        self.assertTrue(str(letter.pdf_path).endswith("northbridge.pdf"))

    def test_html_override_replaces_subject_and_greeting(self):
        self._write_pdf("mosman")
        self._write_html("mosman", "Subject: Update for {{suburb}}\n\n<p>Hi {{first_name}}</p>")
        letter = m.load_letter("MOSMAN")
        self.assertEqual(letter.subject_template, "Update for {{suburb}}")
        self.assertIn("{{first_name}}", letter.greeting_template)

    def test_placeholder_marker_in_html_falls_back_to_defaults(self):
        self._write_pdf("cammeray")
        self._write_html("cammeray", "Subject: Hi\n\n[EDDIE: WRITE THIS LETTER]\n<p>draft</p>")
        letter = m.load_letter("cammeray")
        self.assertIsNotNone(letter)  # PDF present -> still a real letter
        self.assertEqual(letter.subject_template, m._DEFAULT_SUBJECT_TEMPLATE)

    def test_html_missing_subject_line_raises(self):
        self._write_pdf("killara")
        self._write_html("killara", "<p>no subject line</p>")
        with self.assertRaises(m.WorkbookShapeError):
            m.load_letter("killara")

    def test_render_pdf_pages_returns_one_image_per_page(self):
        path = m._LETTERS_DIR / "roseville.pdf"
        _make_pdf(path, pages=3)
        images = m.render_pdf_pages(path)
        self.assertEqual(len(images), 3)
        cids = [cid for cid, _data, _ct in images]
        self.assertEqual(len(set(cids)), 3)  # each page gets a unique cid
        for _cid, data, content_type in images:
            self.assertEqual(content_type, "image/png")
            self.assertTrue(data.startswith(b"\x89PNG"))

    def test_render_letter_substitutes_fields_and_embeds_images(self):
        letter = m.Letter(
            subject_template="Update for {{suburb}}",
            greeting_template="<p>Hi {{first_name}} {{last_name}}, re {{street_address}} ({{property_type}})</p>",
            pdf_path=m._LETTERS_DIR / "unused.pdf",
        )
        r = m.Recipient(segment="s", segment_label="S", suburb="ARTARMON", street="4",
                         address="Smith Street", state="NSW", postcode="2064",
                         first_name="Bob", last_name="Buyer", email="bob@x.com",
                         property_type="Commercial")
        images = [("cid1", b"pngbytes", "image/png"), ("cid2", b"pngbytes2", "image/png")]
        subject, body = m.render_letter(letter, r, images)
        self.assertEqual(subject, "Update for Artarmon")
        self.assertIn("Hi Bob Buyer, re 4 Smith Street (Commercial)", body)
        self.assertIn('src="cid:cid1"', body)
        self.assertIn('src="cid:cid2"', body)

    def test_render_letter_with_no_images_omits_img_tags(self):
        letter = m.Letter(subject_template="S", greeting_template="<p>Hi</p>", pdf_path=m._LETTERS_DIR / "x.pdf")
        r = m.Recipient(segment="s", segment_label="S", suburb="A", street="1", address="X",
                         state="NSW", postcode="2000", first_name="B", last_name="C", email="b@x.com")
        subject, body = m.render_letter(letter, r, [])
        self.assertNotIn("<img", body)

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
        _make_workbook(self.xlsx_path, "SheetA", [
            _row(email="a@x.com", suburb="ARTARMON"),
            _row(email="b@x.com", suburb="MOSMAN"),
        ])

        self._orig_segments = dict(m.SEGMENTS)
        m.SEGMENTS.clear()
        m.SEGMENTS["seg_a"] = {"label": "Segment A", "sheet": "SheetA"}
        self.addCleanup(lambda: (m.SEGMENTS.clear(), m.SEGMENTS.update(self._orig_segments)))

        self.letters_tmp = tempfile.mkdtemp()
        self._orig_letters_dir = m._LETTERS_DIR
        m._LETTERS_DIR = __import__("pathlib").Path(self.letters_tmp)
        self.addCleanup(setattr, m, "_LETTERS_DIR", self._orig_letters_dir)

    def _write_letter(self, suburb):
        _make_pdf(m._LETTERS_DIR / f"{m.suburb_slug(suburb)}.pdf")

    def test_dry_run_with_no_letters_drafts_nothing(self):
        summary = m.run(dry_run=True, xlsx_path=self.xlsx_path)
        self.assertEqual(summary["seg_a"]["drafted"], 0)
        self.assertEqual(summary["seg_a"]["skipped_no_letter"], 2)
        self.assertEqual(sorted(summary["seg_a"]["suburbs_skipped_no_letter"]), ["ARTARMON", "MOSMAN"])
        self.assertEqual(summary["seg_a"]["suburbs_drafted"], [])

    def test_only_suburbs_with_a_letter_are_drafted(self):
        self._write_letter("ARTARMON")  # MOSMAN gets no letter this week
        summary = m.run(dry_run=True, xlsx_path=self.xlsx_path)
        self.assertEqual(summary["seg_a"]["drafted"], 1)
        self.assertEqual(summary["seg_a"]["suburbs_drafted"], ["ARTARMON"])
        self.assertEqual(summary["seg_a"]["suburbs_skipped_no_letter"], ["MOSMAN"])
        self.assertEqual(summary["seg_a"]["skipped_no_letter"], 1)
        self.assertEqual(summary["seg_a"]["previews"][0]["to"], "a@x.com")

    def test_dry_run_with_letters_previews_without_saving(self):
        self._write_letter("ARTARMON")
        self._write_letter("MOSMAN")
        summary = m.run(dry_run=True, xlsx_path=self.xlsx_path)
        self.assertEqual(summary["seg_a"]["drafted"], 2)
        self.assertEqual(len(summary["seg_a"]["previews"]), 2)
        self.assertFalse(os.path.exists(m._STATE_PATH))

    def test_limit_caps_recipients_per_segment(self):
        self._write_letter("ARTARMON")
        self._write_letter("MOSMAN")
        summary = m.run(dry_run=True, xlsx_path=self.xlsx_path, limit=1)
        self.assertEqual(summary["seg_a"]["drafted"], 1)

    def test_only_segment_filters_to_one(self):
        m.SEGMENTS["seg_b"] = {"label": "Segment B", "sheet": "DoesNotExist"}
        self._write_letter("ARTARMON")
        self._write_letter("MOSMAN")
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
        self._write_letter("ARTARMON")
        self._write_letter("MOSMAN")
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
        self._write_letter("ARTARMON")
        self._write_letter("MOSMAN")
        state_path = os.path.join(self.tmp, "state.json")
        with patch("market_update_agent._STATE_PATH", state_path):
            m.run(dry_run=False, xlsx_path=self.xlsx_path, gmail_copies=False)
            second = m.run(dry_run=False, xlsx_path=self.xlsx_path, gmail_copies=False)
        self.assertEqual(second["seg_a"]["drafted"], 0)
        self.assertEqual(second["seg_a"]["skipped_already_drafted"], 2)

    @patch("market_update_agent.OutlookClient")
    @patch("market_update_agent.outlook_create_new_draft", return_value="OD1")
    def test_owner_in_two_suburbs_gets_two_independent_drafts_not_one_blocking_the_other(self, outlook_draft, outlook_cls):
        # Same owner, two suburbs, both with letters -- dedupe already keeps
        # both rows (see TestDedupeByOwnerAndSuburb); this checks the run-time
        # state-key bug that would silently make the second suburb's draft
        # look "already sent" because it shared a state key with the first.
        xlsx_path = os.path.join(self.tmp, "two_suburb.xlsx")
        _make_workbook(xlsx_path, "SheetA", [
            _row(email="same@x.com", suburb="ARTARMON"),
            _row(email="same@x.com", suburb="MOSMAN"),
        ])
        self._write_letter("ARTARMON")
        self._write_letter("MOSMAN")
        state_path = os.path.join(self.tmp, "state2.json")
        with patch("market_update_agent._STATE_PATH", state_path):
            summary = m.run(dry_run=False, xlsx_path=xlsx_path, gmail_copies=False)
        self.assertEqual(summary["seg_a"]["drafted"], 2)
        self.assertEqual(summary["seg_a"]["skipped_already_drafted"], 0)
        with open(state_path) as fh:
            state = json.load(fh)
        self.assertEqual(len(state), 2)  # two distinct (email, suburb) state entries

    @patch("market_update_agent.OutlookClient")
    def test_images_are_passed_through_to_outlook_draft_call(self, outlook_cls):
        _make_pdf(m._LETTERS_DIR / "artarmon.pdf", pages=2)  # a known 2-page PDF
        state_path = os.path.join(self.tmp, "state3.json")
        captured = {}

        def fake_outlook_post(endpoint, body):
            captured["body"] = body
            return {"id": "OD1"}

        outlook_cls.return_value._post.side_effect = fake_outlook_post
        with patch("market_update_agent._STATE_PATH", state_path), \
             patch("market_update_agent.OutlookClient", outlook_cls):
            m.run(dry_run=False, xlsx_path=self.xlsx_path, only_segment="seg_a", gmail_copies=False)
        attachments = captured["body"].get("attachments", [])
        self.assertEqual(len(attachments), 2)  # one per PDF page
        self.assertTrue(all(a["isInline"] for a in attachments))
        self.assertIn('src="cid:', captured["body"]["body"]["content"])


class TestDraftCreationEmbedsInlineImages(unittest.TestCase):
    def test_outlook_payload_includes_inline_attachments(self):
        outlook = MagicMock()
        outlook._post.return_value = {"id": "OD1"}
        images = [("cid1", b"\x89PNGdata", "image/png")]
        result = m.outlook_create_new_draft(outlook, "to@x.com", "Subj", "<p>Hi</p><img src=\"cid:cid1\">", images)
        self.assertEqual(result, "OD1")
        payload = outlook._post.call_args[0][1]
        self.assertEqual(len(payload["attachments"]), 1)
        att = payload["attachments"][0]
        self.assertTrue(att["isInline"])
        self.assertEqual(att["contentId"], "cid1")
        self.assertEqual(att["contentType"], "image/png")

    def test_outlook_payload_has_no_attachments_key_without_images(self):
        outlook = MagicMock()
        outlook._post.return_value = {"id": "OD1"}
        m.outlook_create_new_draft(outlook, "to@x.com", "Subj", "<p>Hi</p>")
        payload = outlook._post.call_args[0][1]
        self.assertNotIn("attachments", payload)

    def test_gmail_draft_embeds_inline_images_with_content_id(self):
        gmail = MagicMock()
        gmail.service.users.return_value.drafts.return_value.create.return_value.execute.return_value = {"id": "GD1"}
        images = [("cid1", b"\x89PNGdata", "image/png")]
        result = m.gmail_create_new_draft(gmail, "to@x.com", "Subj", '<p>Hi</p><img src="cid:cid1">', "", images)
        self.assertEqual(result, "GD1")
        raw = gmail.service.users.return_value.drafts.return_value.create.call_args[1]["body"]["message"]["raw"]
        import base64
        decoded = base64.urlsafe_b64decode(raw)
        self.assertIn(b"Content-ID: <cid1>", decoded)
        self.assertIn(b"Content-Disposition: inline", decoded)


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
