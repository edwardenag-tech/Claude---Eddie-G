"""EagleAgent (eagleagent.com.au) GraphQL API client.

EagleAgent is the CRM IB Property runs on (confirmed by a live
`property_owner_report` link with an embedded property_id found in one of
Eddie's real sent emails, plus the `eaglesoftware.com.au` addresses CC'd on
real portal-enquiry emails). This client is read-only in this codebase --
nothing here issues a mutation.

Auth is a two-step bearer flow, not OAuth2 client_credentials despite the
client_id/client_secret naming:
  1. POST /api/v3/token with `Authorization: Bearer {client_id}:{client_secret}`
     returns a session token valid 24h.
  2. Every GraphQL request uses that session token as `Authorization: Bearer {token}`.

Base URL is https://www.eagleagent.com.au -- NOT api.eaglesoftware.com.au.
The latter only hosts the static GraphQL schema docs: confirmed by a live
request -- it returns HTTP 405 with AWS x-amz-* headers (an S3 bucket, can't
process POST at all), while www.eagleagent.com.au returns a real 401/"Invalid
credentials" JSON error for the same bad-credential test. The docs there
(https://api.eaglesoftware.com.au/v3/query.doc.html etc.) are real and worth
reading when adding new queries -- it just isn't itself a working endpoint.
"""

import json
import logging
import time
from datetime import datetime, timezone
from typing import Any, Dict, Optional

import requests

logger = logging.getLogger(__name__)

_BASE_URL = "https://www.eagleagent.com.au"
_TOKEN_ENDPOINT = f"{_BASE_URL}/api/v3/token"
_GRAPHQL_ENDPOINT = f"{_BASE_URL}/api/v3/graphql"

# Same rationale as gmail_client.py's _RETRY_DELAYS: a transient 429/5xx is
# worth a short wait-and-retry, a 4xx that isn't a stale token is not.
_RETRY_DELAYS = (5, 15, 30, 60)

# The session token is only ever used with this much life left in it, so a
# request started right before expiry can't die mid-flight.
_EXPIRY_SAFETY_MARGIN_SECONDS = 60


class EagleAuthRequired(Exception):
    """Raised when EAGLE_CLIENT_ID/EAGLE_CLIENT_SECRET are missing or the API
    rejects them. Unlike Outlook/Gmail there's no interactive consent flow to
    fall back to here -- this is always a "fix the credentials" problem, so
    callers should surface it and skip Eagle for the run, same pattern as
    OutlookAuthRequired/GmailAuthRequired elsewhere in this codebase."""


class EagleClient:
    def __init__(
        self,
        client_id: Optional[str],
        client_secret: Optional[str],
        token_cache_path: str = "eagle_token_cache.json",
    ):
        if not client_id or not client_secret:
            raise EagleAuthRequired(
                "EAGLE_CLIENT_ID / EAGLE_CLIENT_SECRET not set -- generate an API "
                "credential in the EagleAgent dashboard (agent.eagleagent.com.au -- "
                "Settings > API Credentials) and add both to .env."
            )
        self.client_id = client_id
        self.client_secret = client_secret
        self.token_cache_path = token_cache_path
        self._token: Optional[str] = None
        self._expires_at: Optional[float] = None
        self._load_cached_token()
        if not self._token_valid():
            self._authenticate()

    # ─── Auth ────────────────────────────────────────────────────────────────

    def _load_cached_token(self) -> None:
        try:
            with open(self.token_cache_path) as fh:
                data = json.load(fh)
            self._token = data.get("token")
            self._expires_at = data.get("expiresAt")
        except (FileNotFoundError, json.JSONDecodeError):
            pass

    def _save_cached_token(self) -> None:
        with open(self.token_cache_path, "w") as fh:
            json.dump({"token": self._token, "expiresAt": self._expires_at}, fh)

    def _token_valid(self) -> bool:
        return bool(
            self._token and self._expires_at
            and self._expires_at > time.time() + _EXPIRY_SAFETY_MARGIN_SECONDS
        )

    def _authenticate(self) -> None:
        resp = requests.post(
            _TOKEN_ENDPOINT,
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.client_id}:{self.client_secret}",
            },
        )
        try:
            data = resp.json()
        except ValueError:
            raise EagleAuthRequired(
                f"Eagle token endpoint returned non-JSON (HTTP {resp.status_code}): "
                f"{resp.text[:300]}"
            )

        if resp.status_code != 200 or data.get("errors"):
            reason = (
                data["errors"][0].get("message", "")
                if data.get("errors") else resp.text[:200]
            )
            raise EagleAuthRequired(
                f"Eagle authentication failed (HTTP {resp.status_code}): {reason}"
            )

        token_data = (data.get("data") or {}).get("token") or {}
        self._token = token_data.get("token")
        self._expires_at = token_data.get("expiresAt")
        if not self._token:
            raise EagleAuthRequired(f"Eagle token response had no token: {data}")
        self._save_cached_token()
        expiry_str = (
            datetime.fromtimestamp(self._expires_at, tz=timezone.utc).isoformat()
            if self._expires_at else "unknown"
        )
        logger.info("Eagle: authenticated, session token valid until %s", expiry_str)

    # ─── GraphQL ─────────────────────────────────────────────────────────────

    @property
    def _headers(self) -> Dict[str, str]:
        return {"Content-Type": "application/json", "Authorization": f"Bearer {self._token}"}

    def graphql(self, query: str, variables: Optional[Dict[str, Any]] = None) -> Dict:
        """POST a GraphQL query/mutation, return its `data` dict.

        Refreshes the session token once (not counted against the retry
        budget below) if the API rejects it as expired/invalid -- a cached
        token from a previous run is the common case, not a credentials
        problem, so it shouldn't need a human to intervene. Retries 429/5xx
        with backoff; any other error raises immediately since retrying
        won't fix a malformed query."""
        if not self._token_valid():
            self._authenticate()

        body: Dict[str, Any] = {"query": query}
        if variables:
            body["variables"] = variables

        reauthed = False
        last_error: Optional[str] = None
        for delay in (0,) + _RETRY_DELAYS:
            if delay:
                logger.warning("Eagle GraphQL request failed (%s) -- retrying in %ds", last_error, delay)
                time.sleep(delay)

            resp = requests.post(_GRAPHQL_ENDPOINT, headers=self._headers, json=body)

            if resp.status_code == 401 and not reauthed:
                reauthed = True
                self._authenticate()
                resp = requests.post(_GRAPHQL_ENDPOINT, headers=self._headers, json=body)

            if resp.status_code == 200:
                data = resp.json()
                if data.get("errors"):
                    raise RuntimeError(f"Eagle GraphQL error: {data['errors']}")
                return data.get("data") or {}

            if resp.status_code == 429 or 500 <= resp.status_code < 600:
                last_error = f"HTTP {resp.status_code}"
                continue

            raise RuntimeError(f"Eagle GraphQL HTTP {resp.status_code}: {resp.text[:300]}")

        raise RuntimeError(f"Eagle GraphQL request failed after retries: {last_error}")
