"""Gmail API wrapper — search, label, archive, delete, send."""

import os
import base64
import logging
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import List, Optional, Dict

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

logger = logging.getLogger(__name__)

# Minimum scopes needed: read, modify (archive/label), and send.
SCOPES = [
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.send",
    # Added so the same token cache also covers DocsClient (docs_client.py),
    # used to read live campaign data from the shared Google Doc.
    "https://www.googleapis.com/auth/documents.readonly",
]


class GmailClient:
    def __init__(self, credentials_path: str, token_path: str = "gmail_token.json"):
        self.credentials_path = credentials_path
        self.token_path = token_path
        self.service = None
        self._label_cache: Dict[str, str] = {}  # name → id
        self._authenticate()

    # ─── Auth ────────────────────────────────────────────────────────────────

    def _authenticate(self):
        """Run OAuth2 flow on first call; refresh silently on subsequent calls."""
        creds = None

        if os.path.exists(self.token_path):
            creds = Credentials.from_authorized_user_file(self.token_path, SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
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

    # ─── Fetch ───────────────────────────────────────────────────────────────

    def get_messages(self, query: str = "", max_results: int = 100) -> List[Dict]:
        """Return full message objects matching a Gmail search query."""
        try:
            result = (
                self.service.users()
                .messages()
                .list(userId="me", q=query, maxResults=max_results)
                .execute()
            )
            message_stubs = result.get("messages", [])

            full_messages = []
            for stub in message_stubs:
                msg = (
                    self.service.users()
                    .messages()
                    .get(userId="me", id=stub["id"], format="full")
                    .execute()
                )
                full_messages.append(msg)

            return full_messages
        except HttpError as exc:
            logger.error("Gmail fetch error: %s", exc)
            return []

    def get_recent_emails(self, since_days: int = 1) -> List[Dict]:
        """Emails received in the last N days (inbox + all mail)."""
        after = (datetime.now() - timedelta(days=since_days)).strftime("%Y/%m/%d")
        return self.get_messages(query=f"after:{after}", max_results=100)

    def get_old_read_emails(self, older_than_days: int = 7) -> List[Dict]:
        """Read emails still sitting in INBOX that are older than N days."""
        before = (datetime.now() - timedelta(days=older_than_days)).strftime("%Y/%m/%d")
        return self.get_messages(query=f"in:inbox is:read before:{before}", max_results=200)

    # ─── Actions ─────────────────────────────────────────────────────────────

    def archive_message(self, msg_id: str) -> bool:
        """Archive by removing INBOX label (email stays in All Mail)."""
        try:
            self.service.users().messages().modify(
                userId="me", id=msg_id, body={"removeLabelIds": ["INBOX"]}
            ).execute()
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
            self.service.users().messages().modify(
                userId="me",
                id=msg_id,
                body={"addLabelIds": [label_id], "removeLabelIds": ["INBOX"]},
            ).execute()
            return True
        except HttpError as exc:
            logger.error("Move-to-label failed for %s → %s: %s", msg_id, label_name, exc)
            return False

    def trash_message(self, msg_id: str) -> bool:
        """Move to Trash (recoverable for 30 days)."""
        try:
            self.service.users().messages().trash(userId="me", id=msg_id).execute()
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
            self.service.users().messages().send(
                userId="me", body={"raw": raw}
            ).execute()
            logger.info("Gmail: sent to %s | subject: %s", to, subject)
            return True
        except HttpError as exc:
            logger.error("Gmail send failed: %s", exc)
            return False

    # ─── Labels ──────────────────────────────────────────────────────────────

    def _get_or_create_label(self, label_name: str) -> Optional[str]:
        """Return the label ID, creating the label if it doesn't exist yet."""
        if label_name in self._label_cache:
            return self._label_cache[label_name]

        try:
            result = self.service.users().labels().list(userId="me").execute()
            for label in result.get("labels", []):
                if label["name"].lower() == label_name.lower():
                    self._label_cache[label_name] = label["id"]
                    return label["id"]

            # Label doesn't exist — create it
            new_label = (
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
                .execute()
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
