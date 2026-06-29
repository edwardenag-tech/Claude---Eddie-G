"""Inbox cleaning logic.

Cleaning rules:
  - Archive (not delete) read emails older than 7 days that aren't deal-related
  - Move promotional / newsletter / marketing emails to "Promotions" folder
  - Trash obvious spam
  - NEVER auto-action emails that look personal, are from known contacts,
    or mention real estate deal keywords
"""

import re
import logging
from typing import Dict, List

logger = logging.getLogger(__name__)

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


# ─── Cleaner class ───────────────────────────────────────────────────────────

class InboxCleaner:
    """Orchestrates cleaning for both Gmail and Outlook inboxes."""

    def __init__(self, gmail_client=None, outlook_client=None):
        self.gmail = gmail_client
        self.outlook = outlook_client
        self.report: Dict = {
            "gmail_archived": 0,
            "gmail_promoted": 0,
            "gmail_trashed": 0,
            "outlook_archived": 0,
            "outlook_promoted": 0,
            "outlook_deleted": 0,
            "actions": [],
        }

    def _log_action(self, action: str):
        self.report["actions"].append(action)
        logger.info(action)

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

            # decision == "keep" → do nothing

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

    # ─── Combined ─────────────────────────────────────────────────────────────

    def run_all(self) -> Dict:
        """Run cleaning on both inboxes and return the consolidated report."""
        self.clean_gmail()
        self.clean_outlook()

        total = sum(
            self.report[k]
            for k in (
                "gmail_archived", "gmail_promoted", "gmail_trashed",
                "outlook_archived", "outlook_promoted", "outlook_deleted",
            )
        )
        logger.info("Inbox cleaning complete — %d total actions", total)
        return self.report
