"""Gmail API wrapper — search, label, archive, delete, send."""

import os
import base64
import logging
import time
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import List, Optional, Dict

from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

logger = logging.getLogger(__name__)


class GmailAuthRequired(Exception):
    """Raised when the cached token can't be silently refreshed (expired or
    revoked refresh token) and interactive auth isn't allowed in this context.

    Callers running unattended (cron/launchd) should catch this and skip
    Gmail for the run rather than let the OAuth flow try to open a browser
    that has nobody to complete it. Only the explicit, human-run
    `python agent.py --auth` command should pass allow_interactive=True.
    """

# Minimum scopes needed: read, modify (archive/label), and send.
SCOPES = [
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.send",
    # Added so the same token cache also covers DocsClient (docs_client.py),
    # used to read live campaign data from the shared Google Doc.
    "https://www.googleapis.com/auth/documents.readonly",
]

# Gmail enforces a per-user quota per *minute*, so a rate-limit 403 clears if we
# wait it out. Delays (seconds) before each retry; the last one covers a full
# quota window. Exhausting them re-raises the original HttpError.
_RETRY_DELAYS = (5, 15, 30, 60)


def _is_retryable(exc: HttpError) -> bool:
    """True for transient failures: rate limits (429, or 403 rateLimitExceeded /
    userRateLimitExceeded) and 5xx server errors. Other 403s (e.g. revoked
    access) are permanent and must not be retried."""
    status = getattr(exc.resp, "status", None)
    if status == 429 or (status is not None and 500 <= status < 600):
        return True
    if status != 403:
        return False
    body = exc.content.decode("utf-8", "replace") if isinstance(exc.content, bytes) else str(exc.content)
    return "ratelimitexceeded" in f"{body} {exc}".lower()


# Gmail's own ML categorisation for bulk/automated mail. A message carrying
# any of these never warrants a reply, so has_replied() shouldn't bother
# checking for one (empirically, ~100% of promo/notification threads are
# single-message with no SENT follow-up, which made every one of them look
# "awaiting reply").
_AUTOMATED_CATEGORIES = {
    "CATEGORY_PROMOTIONS", "CATEGORY_SOCIAL", "CATEGORY_UPDATES", "CATEGORY_FORUMS",
}
# Belt-and-suspenders for automated senders that land in the primary category
# without a CATEGORY_* label.
_AUTOMATED_SENDER_MARKERS = (
    "no-reply@", "noreply@", "donotreply@", "do-not-reply@", "notifications@",
)


class GmailClient:
    def __init__(
        self,
        credentials_path: str,
        token_path: str = "gmail_token.json",
        allow_interactive: bool = False,
    ):
        self.credentials_path = credentials_path
        self.token_path = token_path
        self.allow_interactive = allow_interactive
        self.service = None
        self._label_cache: Dict[str, str] = {}  # name → id
        self._authenticate()

    # ─── Auth ────────────────────────────────────────────────────────────────

    def _authenticate(self):
        """Run OAuth2 flow on first call; refresh silently on subsequent calls.

        A refresh token can die (expired or revoked, e.g. the 7-day expiry
        Google applies to OAuth consent screens still in "Testing" mode) even
        though a cached token file exists. That must not silently propagate
        as a raw RefreshError -- it needs to fall back to the interactive
        flow, but only when self.allow_interactive is True (see
        GmailAuthRequired's docstring), so an unattended run degrades
        gracefully instead of hanging on a browser prompt nobody can answer.
        """
        creds = None
        need_interactive = False

        if os.path.exists(self.token_path):
            creds = Credentials.from_authorized_user_file(self.token_path, SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except RefreshError as exc:
                    logger.warning(
                        "Gmail: cached refresh token is dead (%s) -- needs "
                        "interactive re-consent", exc,
                    )
                    creds = None
                    need_interactive = True
            else:
                need_interactive = True

            if need_interactive:
                if not self.allow_interactive:
                    raise GmailAuthRequired(
                        "No cached Gmail token could be silently refreshed, and "
                        "interactive auth is not allowed in this context. Run "
                        "`python agent.py --auth` interactively to (re-)consent."
                    )
                if not os.path.exists(self.credentials_path):
                    raise FileNotFoundError(
                        f"Gmail credentials not found at '{self.credentials_path}'. "
                        "Download credentials.json from Google Cloud Console and set "
                        "GMAIL_CREDENTIALS_PATH in your .env file."
                    )
                flow = InstalledAppFlow.from_client_secrets_file(self.credentials_path, SCOPES)
                creds = flow.run_local_server(port=0)

            with open(self.token_path, "w") as fh:
                fh.write(creds.to_json())
            logger.info("Gmail token saved to %s", self.token_path)

        self.service = build("gmail", "v1", credentials=creds)
        logger.info("Gmail authenticated successfully")

    # ─── Request helper ──────────────────────────────────────────────────────

    @staticmethod
    def _execute(request):
        """Run a Gmail API request, retrying transient failures with backoff.

        Non-retryable errors, and retryable ones that outlast _RETRY_DELAYS,
        propagate as the original HttpError for the caller's own handling.
        """
        for delay in _RETRY_DELAYS:
            try:
                return request.execute()
            except HttpError as exc:
                if not _is_retryable(exc):
                    raise
                logger.warning("Gmail request throttled/failed (%s) -- retrying in %ds", exc.resp.status, delay)
                time.sleep(delay)
        return request.execute()

    # ─── Fetch ───────────────────────────────────────────────────────────────

    def get_messages(self, query: str = "", max_results: int = 100) -> List[Dict]:
        """Return full message objects matching a Gmail search query.

        A message that still can't be fetched after retries is skipped rather
        than discarding everything fetched so far.
        """
        try:
            result = self._execute(
                self.service.users().messages().list(userId="me", q=query, maxResults=max_results)
            )
        except HttpError as exc:
            logger.error("Gmail fetch error: %s", exc)
            return []

        full_messages = []
        for stub in result.get("messages", []):
            try:
                full_messages.append(
                    self._execute(
                        self.service.users().messages().get(userId="me", id=stub["id"], format="full")
                    )
                )
            except HttpError as exc:
                logger.error("Gmail fetch error for message %s (skipping): %s", stub["id"], exc)
        return full_messages

    def get_recent_emails(self, since_days: int = 1) -> List[Dict]:
        """Emails received in the last N days (inbox + all mail)."""
        after = (datetime.now() - timedelta(days=since_days)).strftime("%Y/%m/%d")
        return self.get_messages(query=f"after:{after}", max_results=100)

    def get_old_read_emails(self, older_than_days: int = 7) -> List[Dict]:
        """Read emails still sitting in INBOX that are older than N days."""
        before = (datetime.now() - timedelta(days=older_than_days)).strftime("%Y/%m/%d")
        return self.get_messages(query=f"in:inbox is:read before:{before}", max_results=200)

    def has_replied(self, thread_id: str, message_id: str) -> Optional[bool]:
        """Whether this account has sent a message in the thread after message_id.

        True/False when known; None if the lookup fails, message_id can't be
        found in the thread, or the message never warranted a reply in the
        first place (Gmail-categorised promo/social/updates/forums mail, or
        an obvious no-reply/notifications sender) -- callers should treat
        None as "don't know / not applicable" rather than "not replied".
        Uses threads().get(format='metadata') -- cheap, no message bodies --
        rather than scanning the mailbox.
        """
        try:
            thread = self._execute(
                self.service.users()
                .threads()
                .get(userId="me", id=thread_id, format="metadata", metadataHeaders=["From"])
            )
        except HttpError as exc:
            logger.warning("Thread lookup failed for %s: %s", thread_id, exc)
            return None

        messages = thread.get("messages", [])
        incoming = next((m for m in messages if m.get("id") == message_id), None)
        if incoming is None:
            return None

        incoming_labels = set(incoming.get("labelIds", []))
        if incoming_labels & _AUTOMATED_CATEGORIES:
            return None

        from_header = next(
            (h["value"] for h in incoming.get("payload", {}).get("headers", []) if h["name"].lower() == "from"),
            "",
        ).lower()
        if any(marker in from_header for marker in _AUTOMATED_SENDER_MARKERS):
            return None

        incoming_date = int(incoming.get("internalDate", 0))
        for m in messages:
            if "SENT" in m.get("labelIds", []) and int(m.get("internalDate", 0)) > incoming_date:
                return True
        return False

    # ─── Actions ─────────────────────────────────────────────────────────────

    def archive_message(self, msg_id: str) -> bool:
        """Archive by removing INBOX label (email stays in All Mail)."""
        try:
            self._execute(self.service.users().messages().modify(
                userId="me", id=msg_id, body={"removeLabelIds": ["INBOX"]}
            ))
            return True
        except HttpError as exc:
            logger.error("Archive failed for %s: %s", msg_id, exc)
            return False

    def move_to_label(self, msg_id: str, label_name: str) -> bool:
        """Apply a label and remove from INBOX (creates label if it doesn't exist)."""
        label_id = self._get_or_create_label(label_name)
        if not label_id:
            return False
        try:
            self._execute(self.service.users().messages().modify(
                userId="me",
                id=msg_id,
                body={"addLabelIds": [label_id], "removeLabelIds": ["INBOX"]},
            ))
            return True
        except HttpError as exc:
            logger.error("Move-to-label failed for %s → %s: %s", msg_id, label_name, exc)
            return False

    def trash_message(self, msg_id: str) -> bool:
        """Move to Trash (recoverable for 30 days)."""
        try:
            self._execute(self.service.users().messages().trash(userId="me", id=msg_id))
            return True
        except HttpError as exc:
            logger.error("Trash failed for %s: %s", msg_id, exc)
            return False

    def send_email(self, to: List[str], subject: str, body_html: str, body_text: str = "") -> bool:
        """Send an email from the authenticated account."""
        try:
            msg = MIMEMultipart("alternative")
            msg["to"] = ", ".join(to)
            msg["subject"] = subject

            if body_text:
                msg.attach(MIMEText(body_text, "plain"))
            msg.attach(MIMEText(body_html, "html"))

            raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
            self._execute(self.service.users().messages().send(
                userId="me", body={"raw": raw}
            ))
            logger.info("Gmail: sent to %s | subject: %s", to, subject)
            return True
        except HttpError as exc:
            logger.error("Gmail send failed: %s", exc)
            return False

    def create_draft(self, to: List[str], subject: str, body_html: str, body_text: str = "") -> Optional[str]:
        """Create a Gmail draft (not sent). Returns the draft ID, or None on failure."""
        try:
            msg = MIMEMultipart("alternative")
            msg["to"] = ", ".join(to)
            msg["subject"] = subject

            if body_text:
                msg.attach(MIMEText(body_text, "plain"))
            msg.attach(MIMEText(body_html, "html"))

            raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
            draft = self._execute(self.service.users().drafts().create(
                userId="me", body={"message": {"raw": raw}}
            ))
            draft_id = draft.get("id")
            logger.info("Gmail: draft created (id=%s) to %s | subject: %s", draft_id, to, subject)
            return draft_id
        except HttpError as exc:
            logger.error("Gmail draft creation failed: %s", exc)
            return None

    # ─── Labels ──────────────────────────────────────────────────────────────

    def _get_or_create_label(self, label_name: str) -> Optional[str]:
        """Return the label ID, creating the label if it doesn't exist yet."""
        if label_name in self._label_cache:
            return self._label_cache[label_name]

        try:
            result = self._execute(self.service.users().labels().list(userId="me"))
            for label in result.get("labels", []):
                if label["name"].lower() == label_name.lower():
                    self._label_cache[label_name] = label["id"]
                    return label["id"]

            # Label doesn't exist — create it
            new_label = self._execute(
                self.service.users()
                .labels()
                .create(
                    userId="me",
                    body={
                        "name": label_name,
                        "labelListVisibility": "labelShow",
                        "messageListVisibility": "show",
                    },
                )
            )
            logger.info("Created Gmail label: %s", label_name)
            label_id = new_label["id"]
            self._label_cache[label_name] = label_id
            return label_id

        except HttpError as exc:
            logger.error("Label lookup/create failed for '%s': %s", label_name, exc)
            return None

    # ─── Data extraction ─────────────────────────────────────────────────────

    @staticmethod
    def extract_email_data(message: Dict) -> Dict:
        """Flatten a raw Gmail API message object into a simple dict."""
        payload = message.get("payload", {})
        headers = {h["name"]: h["value"] for h in payload.get("headers", [])}

        # Prefer plain-text part; fall back to the full body
        body = ""
        if "parts" in payload:
            for part in payload["parts"]:
                if part.get("mimeType") == "text/plain":
                    data = part.get("body", {}).get("data", "")
                    if data:
                        body = base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")
                        break
        else:
            data = payload.get("body", {}).get("data", "")
            if data:
                body = base64.urlsafe_b64decode(data).decode("utf-8", errors="replace")

        labels = message.get("labelIds", [])

        return {
            "id": message["id"],
            "thread_id": message.get("threadId", ""),
            "subject": headers.get("Subject", "(no subject)"),
            "from": headers.get("From", ""),
            "to": headers.get("To", ""),
            "cc": headers.get("Cc", ""),
            "date": headers.get("Date", ""),
            "snippet": message.get("snippet", ""),
            "body": body[:3000],
            "labels": labels,
            "is_unread": "UNREAD" in labels,
            "is_inbox": "INBOX" in labels,
            "source": "gmail",
        }
