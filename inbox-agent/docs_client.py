"""Google Docs API wrapper -- read-only fetch of a live Google Doc's text content.

Reuses the same Google OAuth2 app/credentials as GmailClient. The token cache
must include the 'documents.readonly' scope (see gmail_client.SCOPES) -- if the
cached token predates this scope being added, delete the token file once and
re-run any script that authenticates (e.g. `python agent.py --auth`) to go
through consent again.

This is the module vendor_update_agent.py and draft_agent.py use to always
pull the CURRENT content of a shared Google Doc (e.g. the "Campaign Enquiry
Reply Templates" doc) at run time -- never a cached/stale copy.
"""

import logging
import os
from typing import Optional

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

logger = logging.getLogger(__name__)

# Same Google Cloud project as Gmail. documents.readonly is additive to
# gmail_client.SCOPES -- request both together so one token cache covers both.
DOCS_SCOPES = [
    "https://www.googleapis.com/auth/documents.readonly",
]


class DocsClient:
    def __init__(self, credentials_path: str, token_path: str = "gmail_token.json"):
        """
        credentials_path / token_path: pass the SAME paths used for GmailClient
        so both clients share one OAuth consent + token cache. If the cached
        token was issued before documents.readonly was added to gmail_client's
        SCOPES, this will raise -- delete the token file and re-auth once.
        """
        self.credentials_path = credentials_path
        self.token_path = token_path
        self.service = None
        self._authenticate()

    def _authenticate(self):
        creds = None
        if os.path.exists(self.token_path):
            creds = Credentials.from_authorized_user_file(self.token_path)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if not os.path.exists(self.credentials_path):
                    raise FileNotFoundError(
                        f"Google credentials not found at '{self.credentials_path}'."
                    )
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials_path, DOCS_SCOPES
                )
                creds = flow.run_local_server(port=0)
            with open(self.token_path, "w") as fh:
                fh.write(creds.to_json())

        if not creds.has_scopes(DOCS_SCOPES):
            raise RuntimeError(
                "Cached Google token is missing the 'documents.readonly' scope. "
                "Delete gmail_token.json and re-run `python agent.py --auth` to "
                "re-consent with the updated scope list (see gmail_client.SCOPES)."
            )

        self.service = build("docs", "v1", credentials=creds)
        logger.info("Google Docs client authenticated successfully")

    def fetch_doc_text(self, doc_id: str) -> Optional[str]:
        """Return the plain-text content of a Google Doc, or None on failure."""
        try:
            doc = self.service.documents().get(documentId=doc_id).execute()
        except HttpError as exc:
            logger.error("Docs API fetch failed for %s: %s", doc_id, exc)
            return None

        text_parts = []
        for element in doc.get("body", {}).get("content", []):
            paragraph = element.get("paragraph")
            if not paragraph:
                continue
            for el in paragraph.get("elements", []):
                run = el.get("textRun")
                if run and "content" in run:
                    text_parts.append(run["content"])

        return "".join(text_parts)
