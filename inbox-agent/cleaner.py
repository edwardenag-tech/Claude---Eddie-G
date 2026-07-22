"""Inbox cleaning logic.

Cleaning rules:
  - Archive (not delete) read emails older than 7 days that aren't deal-related
  - Move promotional / newsletter / marketing emails to "Promotions" folder
  - Trash obvious spam
  - NEVER auto-action emails that look personal, are from known contacts,
    or mention real estate deal keywords

Optional aggressive-delete pass (added 2026-07-22, OFF by default):
  Per Eddie's explicit call, two categories should be deleted outright rather
  than archived or left as "keep": (1) other agents' "just listed" / market-
  update spam, and (2) emails where Eddie is only CC'd and the content isn't
  relevant to him. Both are judgment calls that plain regex over-includes on
  (the broad _DEAL_KEYWORDS list currently marks almost anything mentioning
  "listing" or "commercial" as untouchable "keep") -- so per Eddie's own
  instruction this is approximated via Claude classification, not rigid
  keyword rules. This pass only runs if an anthropic_api_key is passed to
  InboxCleaner AND is_aggressive_delete_enabled() is True (env var
  AGGRESSIVE_DELETE_ENABLED=true) -- it is deliberately double-gated so
  pulling this code does not silently change live behaviour. Do not enable
  it against a real inbox without first reviewing real flagged examples --
  see [[feedback-inbox-agent-approval-process]] in Claude's memory.
"""

import os
import re
import logging
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)


def is_aggressive_delete_enabled() -> bool:
    return os.getenv("AGGRESSIVE_DELETE_ENABLED", "").strip().lower() in ("1", "true", "yes")

# ─── Classification patterns ─────────────────────────────────────────────────

# Sender patterns that almost always mean bulk/marketing email
_PROMO_SENDER_RE = re.compile(
    r"(no[_-]?reply|noreply|donotreply|newsletter|marketing|promo|offers?"
    r"|notifications?|updates?|info@(?!ibproperty)|hello@|hi@|team@)@",
    re.IGNORECASE,
)

# Body/subject indicators of a promotional email
_PROMO_BODY_RE = re.compile(
    r"unsubscribe|manage your (preferences|subscription)|"
    r"you('re| are) receiving this (email|because)|"
    r"view (in|this|online) (browser|email)|"
    r"privacy policy|terms (of service|& conditions)",
    re.IGNORECASE,
)

# High-confidence spam patterns
_SPAM_RE = re.compile(
    r"(win(ner)?|won).{0,40}(prize|lottery|million|reward|cash)"
    r"|(click here|act now|limited.{0,10}time).{0,30}(offer|deal)"
    r"|(bitcoin|crypto|forex|investment).{0,30}(opportunit|guaranteed profit|return)"
    r"|(free|complimentary).{0,20}(iphone|macbook|gift.?card|vacation)"
    r"|guaranteed (income|return|profit|earning)"
    r"|nigerian?.{0,20}(prince|royalt|fund|inheritance)"
    r"|enlarg(e|ement)|v[1i]agra|c[1i]al[1i]s",
    re.IGNORECASE,
)

# Keywords that mark an email as deal/client/work-related — never auto-touch
_DEAL_KEYWORDS = [
    "lease", "tenant", "landlord", "inspection", "rent", "contract",
    "offer", "listing", "vendor", "buyer", "settlement", "auction",
    "appraisal", "strata", "council", "zoning", "development",
    "commercial", "retail", "office", "warehouse", "property",
    "ibproperty", "edward@", "pm ", "property management",
    "due diligence", "exchange", "deposit", "valuation",
]

# Domains we treat as trusted business contacts — never archive aggressively
_SAFE_DOMAINS = {
    "ibproperty.com.au",
    "nsw.gov.au",
    "gov.au",
    "rea.com.au",
    "domain.com.au",
    "rpdata.com",
    "corelogic.com",
    "raywhite.com",
    "ljhooker.com",
    "mcgrath.com.au",
}


# ─── Classification helpers ──────────────────────────────────────────────────

def _combined_text(email: Dict) -> str:
    return f"{email.get('subject', '')} {email.get('snippet', '')} {email.get('body', '')}".lower()


def _sender(email: Dict) -> str:
    return email.get("from", "").lower()


def is_deal_related(email: Dict) -> bool:
    """Return True if the email mentions commercial real estate topics."""
    text = _combined_text(email)
    sender = _sender(email)
    return any(kw in text or kw in sender for kw in _DEAL_KEYWORDS)


def is_from_safe_domain(email: Dict) -> bool:
    sender = _sender(email)
    return any(domain in sender for domain in _SAFE_DOMAINS)


def is_spam(email: Dict) -> bool:
    text = _combined_text(email)
    return bool(_SPAM_RE.search(text))


def is_promotional(email: Dict) -> bool:
    sender = _sender(email)
    text = _combined_text(email)
    return bool(_PROMO_SENDER_RE.search(sender) or _PROMO_BODY_RE.search(text))


def is_safe_to_auto_action(email: Dict) -> bool:
    """
    Return True only if it's safe to archive/move/delete this email.
    Emails that look personal, are from safe domains, or mention deal keywords
    are never auto-actioned.
    """
    if is_deal_related(email):
        return False
    if is_from_safe_domain(email):
        return False
    return True


def classify(email: Dict) -> str:
    """
    Classify an email into one of: 'spam', 'promo', 'archive', 'keep'.
    """
    if not is_safe_to_auto_action(email):
        return "keep"
    if is_spam(email):
        return "spam"
    if is_promotional(email):
        return "promo"
    return "archive"


# ─── Optional Claude-judged aggressive delete (opt-in, see module docstring) ──

_AGGRESSIVE_DELETE_PROMPT = """You are helping a commercial real estate agent (Eddie, IB Property Sydney) triage his inbox. Below is one email that a rule-based filter marked "keep" because it mentioned a real-estate-related keyword -- but the keyword filter is too broad and lets junk through.

Eddie's explicit instruction: delete outright (don't just archive) if EITHER of these is true:
1. It's another agent/agency announcing a property they just listed, a market update, or similar promotional "just hit the market" style content -- not something addressed to Eddie personally about his own deals.
2. Eddie is only CC'd (not a direct recipient) and the content isn't actually relevant or actionable for him -- e.g. an internal thread between other people he doesn't need to track.

If NEITHER applies -- it's a real client/deal/personal email, or you're not confident -- say KEEP. When in doubt, KEEP; false deletes are far worse than a missed archive.

Email:
From: {sender}
To/Cc note: {recipient_note}
Subject: {subject}
Body preview: {body}

Respond with exactly one word: DELETE or KEEP."""


def claude_judge_aggressive_delete(
    email: Dict,
    ai_client,
    is_cc_only: bool = False,
) -> bool:
    """Ask Claude whether this 'keep'-classified email should actually be deleted.

    Returns True only for a confident DELETE. Any error, ambiguous response,
    or KEEP verdict returns False (never delete on uncertainty).
    """
    recipient_note = (
        "Eddie is CC'd, not a direct recipient" if is_cc_only else "Eddie is a direct recipient"
    )
    prompt = _AGGRESSIVE_DELETE_PROMPT.format(
        sender=email.get("from", ""),
        recipient_note=recipient_note,
        subject=email.get("subject", ""),
        body=(email.get("body") or email.get("snippet") or "")[:800],
    )
    try:
        response = ai_client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=10,
            messages=[{"role": "user", "content": prompt}],
        )
        verdict = response.content[0].text.strip().upper() if response.content else ""
        return verdict.startswith("DELETE")
    except Exception as exc:
        logger.error("Aggressive-delete Claude judgment failed for '%s': %s",
                     email.get("subject", "")[:60], exc)
        return False


# ─── Cleaner class ───────────────────────────────────────────────────────────

def _is_cc_only(email: Dict, user_addresses: List[str]) -> bool:
    """Best-effort check: is the user only in Cc, not To? Defaults to False
    (assume direct recipient) if To/Cc data isn't available."""
    to_field = (email.get("to") or "").lower()
    cc_field = (email.get("cc") or "").lower()
    if not to_field and not cc_field:
        return False
    for addr in user_addresses:
        addr = addr.lower()
        if addr in cc_field and addr not in to_field:
            return True
    return False


class InboxCleaner:
    """Orchestrates cleaning for both Gmail and Outlook inboxes."""

    def __init__(
        self,
        gmail_client=None,
        outlook_client=None,
        anthropic_api_key: Optional[str] = None,
        user_addresses: Optional[List[str]] = None,
    ):
        self.gmail = gmail_client
        self.outlook = outlook_client
        self.user_addresses = user_addresses or []
        self._ai = None
        if anthropic_api_key and is_aggressive_delete_enabled():
            import anthropic
            self._ai = anthropic.Anthropic(api_key=anthropic_api_key)
            logger.info("Aggressive-delete pass ENABLED (AGGRESSIVE_DELETE_ENABLED=true)")
        self.report: Dict = {
            "gmail_archived": 0,
            "gmail_promoted": 0,
            "gmail_trashed": 0,
            "gmail_aggressive_deleted": 0,
            "outlook_archived": 0,
            "outlook_promoted": 0,
            "outlook_deleted": 0,
            "outlook_aggressive_deleted": 0,
            "actions": [],
        }

    def _log_action(self, action: str):
        self.report["actions"].append(action)
        logger.info(action)

    def _aggressive_delete_check(self, email: Dict) -> bool:
        """For a 'keep'-classified email, ask Claude if it should be deleted
        anyway per Eddie's two aggressive-delete categories. No-op (returns
        False) unless the feature is enabled and an AI client is configured."""
        if not self._ai:
            return False
        is_cc_only = _is_cc_only(email, self.user_addresses)
        return claude_judge_aggressive_delete(email, self._ai, is_cc_only=is_cc_only)

    # ─── Gmail ────────────────────────────────────────────────────────────────

    def clean_gmail(self):
        if not self.gmail:
            return

        from gmail_client import GmailClient

        old_msgs = self.gmail.get_old_read_emails(older_than_days=7)
        logger.info("Gmail: found %d old read messages to evaluate", len(old_msgs))

        for raw in old_msgs:
            email = GmailClient.extract_email_data(raw)
            decision = classify(email)
            subj = email["subject"][:70]

            if decision == "spam":
                if self.gmail.trash_message(email["id"]):
                    self.report["gmail_trashed"] += 1
                    self._log_action(f"[Gmail] TRASHED (spam): {subj}")

            elif decision == "promo":
                if self.gmail.move_to_label(email["id"], "Promotions"):
                    self.report["gmail_promoted"] += 1
                    self._log_action(f"[Gmail] → Promotions: {subj}")

            elif decision == "archive":
                if self.gmail.archive_message(email["id"]):
                    self.report["gmail_archived"] += 1
                    self._log_action(f"[Gmail] ARCHIVED: {subj}")

            elif decision == "keep" and self._aggressive_delete_check(email):
                if self.gmail.trash_message(email["id"]):
                    self.report["gmail_aggressive_deleted"] += 1
                    self._log_action(f"[Gmail] TRASHED (AI-judged, competitor/irrelevant-cc): {subj}")

    # ─── Outlook ──────────────────────────────────────────────────────────────

    def clean_outlook(self):
        if not self.outlook:
            return

        from outlook_client import OutlookClient

        old_msgs = self.outlook.get_old_read_emails(older_than_days=7)
        logger.info("Outlook: found %d old read messages to evaluate", len(old_msgs))

        for raw in old_msgs:
            email = OutlookClient.extract_email_data(raw)
            decision = classify(email)
            subj = email["subject"][:70]

            if decision == "spam":
                if self.outlook.delete_message(email["id"]):
                    self.report["outlook_deleted"] += 1
                    self._log_action(f"[Outlook] DELETED (spam): {subj}")

            elif decision == "promo":
                if self.outlook.move_to_folder(email["id"], "Promotions"):
                    self.report["outlook_promoted"] += 1
                    self._log_action(f"[Outlook] → Promotions: {subj}")

            elif decision == "archive":
                if self.outlook.archive_message(email["id"]):
                    self.report["outlook_archived"] += 1
                    self._log_action(f"[Outlook] ARCHIVED: {subj}")

            elif decision == "keep" and self._aggressive_delete_check(email):
                if self.outlook.delete_message(email["id"]):
                    self.report["outlook_aggressive_deleted"] += 1
                    self._log_action(f"[Outlook] DELETED (AI-judged, competitor/irrelevant-cc): {subj}")

    # ─── Combined ─────────────────────────────────────────────────────────────

    def run_all(self) -> Dict:
        """Run cleaning on both inboxes and return the consolidated report."""
        self.clean_gmail()
        self.clean_outlook()

        total = sum(
            self.report[k]
            for k in (
                "gmail_archived", "gmail_promoted", "gmail_trashed", "gmail_aggressive_deleted",
                "outlook_archived", "outlook_promoted", "outlook_deleted", "outlook_aggressive_deleted",
            )
        )
        logger.info("Inbox cleaning complete — %d total actions", total)
        return self.report
