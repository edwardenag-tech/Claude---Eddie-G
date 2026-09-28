"""Tests for eagle_client.py: token auth/caching and the GraphQL transport's
error handling. All HTTP is mocked -- no real network calls, no real
credentials needed to run this suite.
"""

import json
import os
import sys
import tempfile
import time
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import eagle_client as m


def _resp(status_code=200, json_data=None, text=""):
    r = MagicMock()
    r.status_code = status_code
    r.text = text or json.dumps(json_data or {})
    if json_data is None:
        r.json.side_effect = ValueError("no json")
    else:
        r.json.return_value = json_data
    return r


def _token_resp(token="tok123", expires_at=None, status_code=200):
    expires_at = expires_at if expires_at is not None else time.time() + 86400
    return _resp(status_code, {"data": {"token": {"token": token, "expiresAt": expires_at}}})


class TestAuth(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.cache_path = os.path.join(self.tmp, "eagle_token_cache.json")

    def test_missing_client_id_raises(self):
        with self.assertRaises(m.EagleAuthRequired):
            m.EagleClient(None, "secret", token_cache_path=self.cache_path)

    def test_missing_client_secret_raises(self):
        with self.assertRaises(m.EagleAuthRequired):
            m.EagleClient("id", "", token_cache_path=self.cache_path)

    @patch("eagle_client.requests.post")
    def test_successful_auth_caches_token(self, post):
        post.return_value = _token_resp(token="abc", expires_at=time.time() + 86400)
        client = m.EagleClient("id", "secret", token_cache_path=self.cache_path)
        self.assertEqual(client._token, "abc")
        with open(self.cache_path) as fh:
            cached = json.load(fh)
        self.assertEqual(cached["token"], "abc")

        # Correct auth header shape: "Bearer client_id:client_secret"
        headers = post.call_args.kwargs["headers"]
        self.assertEqual(headers["Authorization"], "Bearer id:secret")
        self.assertEqual(post.call_args.args[0], "https://www.eagleagent.com.au/api/v3/token")

    @patch("eagle_client.requests.post")
    def test_invalid_credentials_raises_auth_required(self, post):
        post.return_value = _resp(200, {"errors": [{"message": "Invalid credentials"}], "data": {"token": None}})
        with self.assertRaises(m.EagleAuthRequired) as ctx:
            m.EagleClient("id", "wrong-secret", token_cache_path=self.cache_path)
        self.assertIn("Invalid credentials", str(ctx.exception))

    @patch("eagle_client.requests.post")
    def test_non_200_status_raises_auth_required(self, post):
        post.return_value = _resp(500, text="Internal Server Error")
        with self.assertRaises(m.EagleAuthRequired):
            m.EagleClient("id", "secret", token_cache_path=self.cache_path)

    @patch("eagle_client.requests.post")
    def test_non_json_response_raises_auth_required(self, post):
        post.return_value = _resp(405, json_data=None, text="<html>405 Method Not Allowed</html>")
        with self.assertRaises(m.EagleAuthRequired):
            m.EagleClient("id", "secret", token_cache_path=self.cache_path)

    @patch("eagle_client.requests.post")
    def test_cached_valid_token_skips_reauth(self, post):
        with open(self.cache_path, "w") as fh:
            json.dump({"token": "cached-tok", "expiresAt": time.time() + 86400}, fh)
        client = m.EagleClient("id", "secret", token_cache_path=self.cache_path)
        post.assert_not_called()
        self.assertEqual(client._token, "cached-tok")

    @patch("eagle_client.requests.post")
    def test_cached_expired_token_triggers_reauth(self, post):
        with open(self.cache_path, "w") as fh:
            json.dump({"token": "stale-tok", "expiresAt": time.time() - 10}, fh)
        post.return_value = _token_resp(token="fresh-tok")
        client = m.EagleClient("id", "secret", token_cache_path=self.cache_path)
        post.assert_called_once()
        self.assertEqual(client._token, "fresh-tok")


class TestGraphQL(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.cache_path = os.path.join(self.tmp, "eagle_token_cache.json")
        with open(self.cache_path, "w") as fh:
            json.dump({"token": "tok", "expiresAt": time.time() + 86400}, fh)
        self.client = m.EagleClient("id", "secret", token_cache_path=self.cache_path)

    @patch("eagle_client.requests.post")
    def test_successful_query_returns_data(self, post):
        post.return_value = _resp(200, {"data": {"properties": {"nodes": []}}})
        result = self.client.graphql("query { properties { nodes { id } } }")
        self.assertEqual(result, {"properties": {"nodes": []}})
        self.assertEqual(post.call_args.args[0], "https://www.eagleagent.com.au/api/v3/graphql")
        self.assertEqual(post.call_args.kwargs["headers"]["Authorization"], "Bearer tok")

    @patch("eagle_client.requests.post")
    def test_variables_included_in_body_when_given(self, post):
        post.return_value = _resp(200, {"data": {"ok": True}})
        self.client.graphql("query Q($x: String) { ok }", {"x": "y"})
        self.assertEqual(post.call_args.kwargs["json"]["variables"], {"x": "y"})

    @patch("eagle_client.requests.post")
    def test_graphql_level_error_raises_runtime_error(self, post):
        post.return_value = _resp(200, {"errors": [{"message": "field does not exist"}]})
        with self.assertRaises(RuntimeError) as ctx:
            self.client.graphql("query { nonsense }")
        self.assertIn("field does not exist", str(ctx.exception))

    @patch("eagle_client.requests.post")
    def test_401_triggers_one_reauth_and_retries(self, post):
        post.side_effect = [
            _resp(401, text="token expired"),
            _token_resp(token="new-tok"),
            _resp(200, {"data": {"ok": True}}),
        ]
        result = self.client.graphql("query { ok }")
        self.assertEqual(result, {"ok": True})
        self.assertEqual(self.client._token, "new-tok")
        self.assertEqual(post.call_count, 3)

    @patch("eagle_client.time.sleep", return_value=None)
    @patch("eagle_client.requests.post")
    def test_429_is_retried_with_backoff_then_succeeds(self, post, _sleep):
        post.side_effect = [_resp(429, text="rate limited"), _resp(200, {"data": {"ok": True}})]
        result = self.client.graphql("query { ok }")
        self.assertEqual(result, {"ok": True})

    @patch("eagle_client.time.sleep", return_value=None)
    @patch("eagle_client.requests.post")
    def test_persistent_5xx_exhausts_retries_and_raises(self, post, _sleep):
        post.return_value = _resp(503, text="unavailable")
        with self.assertRaises(RuntimeError):
            self.client.graphql("query { ok }")
        # one initial attempt + one per _RETRY_DELAYS entry
        self.assertEqual(post.call_count, 1 + len(m._RETRY_DELAYS))

    @patch("eagle_client.requests.post")
    def test_non_retryable_4xx_raises_immediately_without_retrying(self, post):
        post.return_value = _resp(400, text="bad query")
        with self.assertRaises(RuntimeError):
            self.client.graphql("query { ok }")
        post.assert_called_once()


if __name__ == "__main__":
    unittest.main()
