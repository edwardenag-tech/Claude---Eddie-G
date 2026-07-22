"""Microsoft Graph API wrapper using MSAL device code flow.

First-run: prints a URL + code to the terminal so you can authenticate in a browser.
After that the token is cached in msal_token_cache.json and refreshed silently.
"""

import os
import re
import logging
from datetime import datetime, timedelta
from typing import List, Dict, Optional

import msal
import pytz
import requests

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# offline_access is required to get a refresh token so the cache stays valid.
# Calendars.Read added for the morning briefing's "Today's Calendar" section --
# a token cached before this was added won't have consent for it (Graph calls
# will fail with a permission error until Eddie re-runs `python agent.py --auth`).
SCOPES = [
    "Mail.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "User.Read",
    "Calendars.Read",
]


class OutlookClient:
    def __init__(
        self,
        client_id: str,
        tenant_id: str,
        token_cache_path: str = "msal_token_cache.json",
    ):
        self.client_id = client_id
        self.tenant_id = tenant_id
        self.token_cache_path = token_cache_path
        self._access_token: Optional[str] = None
        self._folder_cache: Dict[str, str] = {}  # display_name → id
        self._authenticate()

    # ─── Auth ────────────────────────────────────────────────────────────────

    def _authenticate(self):
        """Authenticate via MSAL. On first run, prints device-code instructions."""
        cache = msal.SerializableTokenCache()
        if os.path.exists(self.token_cache_path):
            with open(self.token_cache_path) as fh:
                cache.deserialize(fh.read())

        app = msal.PublicClientApplication(
            self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
            token_cache=cache,
        )

        result = None
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])
            if result:
                logger.info("Outlook: silent token refresh succeeded")

        if not result:
            logger.info("Outlook: no cached token — starting device code flow")
            flow = app.initiate_device_flow(scopes=SCOPES)
            if "user_code" not in flow:
                raise RuntimeError(
                    f"MSAL device flow failed: {flow.get('error_description', 'unknown')}"
                )

            print("\n" + "=" * 65)
            print("  OUTLOOK AUTHENTICATION REQUIRED")
            print("=" * 65)
            print(flow["message"])
            print("=" * 65 + "\n")

            result = app.acquire_token_by_device_flow(flow)

        if "access_token" not in result:
            raise RuntimeError(
                f"Outlook auth failed: {result.get('error_description', result.get('error', 'unknown'))}"
            )

        self._access_token = result["access_token"]

        if cache.has_state_changed:
            with open(self.token_cache_path, "w") as fh:
                fh.write(cache.serialize())
            logger.info("Outlook: token cache saved to %s", self.token_cache_path)

        logger.info("Outlook authenticated successfully")

    # ─── HTTP helpers ─────────────────────────────────────────────────────────

    @property
    def _headers(self) -> Dict:
        return {
            "Authorization": f"Bearer {self._access_token}",
            "Content-Type": "application/json",
        }

    def _get(self, endpoint: str, params: Dict = None, extra_headers: Dict = None) -> Optional[Dict]:
        headers = self._headers
        if extra_headers:
            headers = {**headers, **extra_headers}
        resp = requests.get(f"{GRAPH_BASE}{endpoint}", headers=headers, params=params)
        if resp.status_code == 200:
            return resp.json()
        logger.error("Graph GET %s → %s: %s", endpoint, resp.status_code, resp.text[:300])
        return None

    def _post(self, endpoint: str, body: Dict) -> Optional[Dict]:
        resp = requests.post(f"{GRAPH_BASE}{endpoint}", headers=self._headers, json=body)
        if resp.status_code in (200, 201, 202):
            try:
                return resp.json()
            except ValueError:
                return {}  # 202 Accepted (sendMail) returns no body
        logger.error("Graph POST %s → %s: %s", endpoint, resp.status_code, resp.text[:300])
        return None

    def _patch(self, endpoint: str, body: Dict) -> bool:
        resp = requests.patch(f"{GRAPH_BASE}{endpoint}", headers=self._headers, json=body)
        return resp.status_code in (200, 204)

    def _delete(self, endpoint: str) -> bool:
        resp = requests.delete(f"{GRAPH_BASE}{endpoint}", headers=self._headers)
        return resp.status_code in (200, 204)

    # ─── Fetch ───────────────────────────────────────────────────────────────

    def get_messages(
        self,
        folder: str = "inbox",
        filter_query: str = None,
        top: int = 50,
    ) -> List[Dict]:
        """Return messages from a well-known folder or a folder ID."""
        params: Dict = {
            "$top": top,
            "$orderby": "receivedDateTime desc",
            "$select": (
                "id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,"
                "importance,bodyPreview,body,hasAttachments,conversationId,internetMessageId"
            ),
        }
        if filter_query:
            params["$filter"] = filter_query

        result = self._get(f"/me/mailFolders/{folder}/messages", params=params)
        return result.get("value", []) if result else []

    def _find_folder_id(self, display_name: str) -> Optional[str]:
        """Return folder ID by display name (case-insensitive). None if not found."""
        result = self._get("/me/mailFolders", params={"$top": 100, "$select": "id,displayName"})
        if not result:
            return None
        for folder in result.get("value", []):
            if folder.get("displayName", "").lower() == display_name.lower():
                return folder["id"]
        return None

    def _find_subfolder_id(self, parent_folder: str, child_name: str) -> Optional[str]:
        """Return the ID of a named child folder under parent_folder.

        parent_folder may be a well-known name (e.g. 'inbox') or a folder ID.
        Uses GET /me/mailFolders/{parent}/childFolders.
        """
        result = self._get(
            f"/me/mailFolders/{parent_folder}/childFolders",
            params={"$top": 100, "$select": "id,displayName"},
        )
        if not result:
            return None
        for folder in result.get("value", []):
            if folder.get("displayName", "").lower() == child_name.lower():
                return folder["id"]
        return None

    def get_recent_emails(self, since_days: int = 1, extra_folders: List[str] = None) -> List[Dict]:
        """Emails from Inbox (and any named extra folders) in the last N days.

        extra_folders: list of top-level folder display names to include alongside Inbox.
        Folders that don't exist are skipped with a warning.
        """
        after_dt = (datetime.utcnow() - timedelta(days=since_days)).strftime(
            "%Y-%m-%dT%H:%M:%SZ"
        )
        filter_q = f"receivedDateTime ge {after_dt}"

        messages = self.get_messages(folder="inbox", filter_query=filter_q, top=100)
        seen_ids = {m["id"] for m in messages}

        fom_id = self._find_subfolder_id("inbox", "Front Of Mind")
        if fom_id:
            logger.info("Fetching from 'Front Of Mind' subfolder of Inbox...")
            for m in self.get_messages(folder=fom_id, filter_query=filter_q, top=100):
                if m["id"] not in seen_ids:
                    seen_ids.add(m["id"])
                    messages.append(m)
        else:
            logger.warning("'Front Of Mind' subfolder not found under Inbox")

        for folder_name in (extra_folders or []):
            folder_id = self._find_folder_id(folder_name)
            if not folder_id:
                logger.warning("Outlook folder not found, skipping: %r", folder_name)
                continue
            logger.info("Fetching from folder %r...", folder_name)
            extras = self.get_messages(folder=folder_id, filter_query=filter_q, top=100)
            for m in extras:
                if m["id"] not in seen_ids:
                    seen_ids.add(m["id"])
                    messages.append(m)

        return messages

    def get_old_read_emails(self, older_than_days: int = 7) -> List[Dict]:
        """Read emails older than N days still in the inbox."""
        before_dt = (datetime.utcnow() - timedelta(days=older_than_days)).strftime(
            "%Y-%m-%dT%H:%M:%SZ"
        )
        return self.get_messages(
            folder="inbox",
            filter_query=f"isRead eq true and receivedDateTime le {before_dt}",
            top=200,
        )

    def get_todays_events(self, tz_name: str = "Australia/Sydney") -> List[Dict]:
        """Today's calendar events (local calendar day), sorted by start time.

        Uses GET /me/calendarView with a Prefer: outlook.timezone header so
        Graph returns start/end times already localized to tz_name, rather
        than UTC. Raises RuntimeError if the request fails -- most likely
        because the token cache predates the Calendars.Read scope (see
        SCOPES); callers should catch this and degrade gracefully.
        """
        tz = pytz.timezone(tz_name)
        start_local = datetime.now(tz).replace(hour=0, minute=0, second=0, microsecond=0)
        end_local = start_local + timedelta(days=1)

        params = {
            "startDateTime": start_local.strftime("%Y-%m-%dT%H:%M:%S"),
            "endDateTime": end_local.strftime("%Y-%m-%dT%H:%M:%S"),
            "$orderby": "start/dateTime",
            "$select": "subject,start,end,location,organizer,isAllDay",
            "$top": 50,
        }
        result = self._get(
            "/me/calendarView",
            params=params,
            extra_headers={"Prefer": f'outlook.timezone="{tz_name}"'},
        )
        if result is None:
            raise RuntimeError(
                "Graph calendarView request failed -- check the Calendars.Read "
                "scope has been consented (re-run `python agent.py --auth`)"
            )

        return [self._extract_event_data(e) for e in result.get("value", [])]

    @staticmethod
    def _extract_event_data(event: Dict) -> Dict:
        """Flatten a raw Graph calendarView event into a simple dict."""
        organizer = event.get("organizer", {}).get("emailAddress", {})
        return {
            "subject": event.get("subject", "(no subject)"),
            "start": event.get("start", {}).get("dateTime", ""),
            "end": event.get("end", {}).get("dateTime", ""),
            "location": event.get("location", {}).get("displayName", ""),
            "organizer": organizer.get("name") or organizer.get("address", ""),
            "is_all_day": event.get("isAllDay", False),
        }

    def has_replied(self, conversation_id: str, received_at: str) -> Optional[bool]:
        """Whether a Sent Items message exists in this conversation after received_at.

        received_at should be a Graph datetime string (e.g. receivedDateTime) --
        sentDateTime and receivedDateTime are both UTC 'Z' timestamps from Graph,
        so a plain string comparison is sufficient without parsing.

        True/False when known; None if conversation_id is missing or the lookup
        fails. Queries only Sent Items filtered by conversationId ($top=1) --
        cheap, not a full conversation or mailbox scan.
        """
        if not conversation_id:
            return None

        result = self._get(
            "/me/mailFolders/SentItems/messages",
            params={
                "$filter": f"conversationId eq '{conversation_id}'",
                "$select": "id,sentDateTime",
                "$orderby": "sentDateTime desc",
                "$top": 1,
            },
        )
        if result is None:
            return None

        sent_items = result.get("value", [])
        if not sent_items:
            return False
        return sent_items[0].get("sentDateTime", "") > received_at

    # ─── Actions ─────────────────────────────────────────────────────────────

    def move_message(self, msg_id: str, destination_folder_id: str) -> bool:
        """Move a message to a folder by folder ID."""
        result = self._post(
            f"/me/messages/{msg_id}/move",
            {"destinationId": destination_folder_id},
        )
        return result is not None

    def archive_message(self, msg_id: str) -> bool:
        """Move to the Archive folder (creates it if missing)."""
        folder_id = self._get_or_create_folder("Archive")
        return bool(folder_id) and self.move_message(msg_id, folder_id)

    def move_to_folder(self, msg_id: str, folder_name: str) -> bool:
        """Move to a named top-level folder (creates it if missing)."""
        folder_id = self._get_or_create_folder(folder_name)
        return bool(folder_id) and self.move_message(msg_id, folder_id)

    def delete_message(self, msg_id: str) -> bool:
        """Move to Deleted Items (soft delete — recoverable)."""
        return self._delete(f"/me/messages/{msg_id}")

    def send_email(self, to: List[str], subject: str, body_html: str) -> bool:
        """Send an email via Graph API."""
        recipients = [{"emailAddress": {"address": addr}} for addr in to]
        payload = {
            "message": {
                "subject": subject,
                "body": {"contentType": "HTML", "content": body_html},
                "toRecipients": recipients,
            },
            "saveToSentItems": True,
        }
        result = self._post("/me/sendMail", payload)
        if result is not None:
            logger.info("Outlook: sent to %s | subject: %s", to, subject)
            return True
        return False

    def send_reply(self, msg_id: str, body_html: str) -> bool:
        """Send a reply to an existing message immediately (NOT a draft).

        Uses POST /me/messages/{id}/reply, which sends to the original
        sender and saves a copy to Sent Items automatically. This is
        irreversible the moment it returns True -- only call it from a path
        that has already been explicitly gated (see draft_agent.py's
        AUTO_SEND_ENABLED + quality-bar check). There is no dry-run mode for
        this specific call; test with send_email/drafts first.
        """
        payload = {"comment": body_html}
        resp = self._post(f"/me/messages/{msg_id}/reply", payload)
        # /reply returns 202 Accepted with an empty body -- _post treats that as {}
        if resp is not None:
            logger.info("Outlook: SENT reply to message id=%s", msg_id)
            return True
        logger.error("Outlook: send_reply failed for message id=%s", msg_id)
        return False

    # ─── Folders ─────────────────────────────────────────────────────────────

    def _get_or_create_folder(self, folder_name: str) -> Optional[str]:
        """Return folder ID by display name, creating it if it doesn't exist."""
        if folder_name in self._folder_cache:
            return self._folder_cache[folder_name]

        result = self._get("/me/mailFolders", params={"$top": 100})
        if result:
            for folder in result.get("value", []):
                if folder["displayName"].lower() == folder_name.lower():
                    self._folder_cache[folder_name] = folder["id"]
                    return folder["id"]

        new_folder = self._post("/me/mailFolders", {"displayName": folder_name})
        if new_folder and "id" in new_folder:
            logger.info("Created Outlook folder: %s", folder_name)
            self._folder_cache[folder_name] = new_folder["id"]
            return new_folder["id"]

        logger.error("Could not get or create folder: %s", folder_name)
        return None

    # ─── Data extraction ─────────────────────────────────────────────────────

    @staticmethod
    def _recipients_to_str(recipients: List[Dict]) -> str:
        return ", ".join(
            r.get("emailAddress", {}).get("address", "")
            for r in (recipients or [])
            if r.get("emailAddress", {}).get("address")
        )

    @staticmethod
    def extract_email_data(message: Dict) -> Dict:
        """Flatten a raw Graph API message object into a simple dict."""
        sender = message.get("from", {}).get("emailAddress", {})
        raw_body = message.get("body", {}).get("content", "")
        # Strip HTML tags for plain-text preview
        body_text = re.sub(r"<[^>]+>", " ", raw_body)
        body_text = re.sub(r"\s{2,}", " ", body_text).strip()[:3000]

        return {
            "id": message.get("id", ""),
            "conversation_id": message.get("conversationId", ""),
            "subject": message.get("subject", "(no subject)"),
            "from": sender.get("address", ""),
            "from_name": sender.get("name", ""),
            "to": OutlookClient._recipients_to_str(message.get("toRecipients", [])),
            "cc": OutlookClient._recipients_to_str(message.get("ccRecipients", [])),
            "date": message.get("receivedDateTime", ""),
            "snippet": message.get("bodyPreview", ""),
            "body": body_text,
            "is_unread": not message.get("isRead", True),
            "is_inbox": True,
            "importance": message.get("importance", "normal"),
            "has_attachments": message.get("hasAttachments", False),
            "source": "outlook",
        }
