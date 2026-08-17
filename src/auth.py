import http.server
import json
import os
import threading
import time
import urllib.parse
import webbrowser

import requests

import config

SCOPES = "https://www.googleapis.com/auth/gmail.send https://www.googleapis.com/auth/gmail.readonly"

AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth"
TOKEN_URL = "https://oauth2.googleapis.com/token"
TIMEOUT = 30
LOGIN_TIMEOUT_SECONDS = 300


def _token_cache_path() -> str:
    return config.path("token_cache_path", "token_cache.json")


def _client_id() -> str:
    client_id = config.get("google", "client_id")
    if not client_id or client_id.upper().startswith("YOUR_"):
        raise RuntimeError(
            "google.client_id is not set in config.ini. "
            "Paste the OAuth client ID from your Google Cloud Console project "
            "(APIs & Services > Credentials)."
        )
    return client_id


def _client_secret() -> str:
    secret = config.get("google", "client_secret")
    if not secret or secret.upper().startswith("YOUR_"):
        raise RuntimeError(
            "google.client_secret is not set in config.ini. "
            "Paste the OAuth client secret from your Google Cloud Console project."
        )
    return secret


def _load_cache() -> dict:
    path = _token_cache_path()
    if not os.path.exists(path):
        return {}
    try:
        return json.load(open(path, encoding="utf-8"))
    except Exception:
        return {}


def _save_cache(data: dict):
    path = _token_cache_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    os.chmod(path, 0o600)


# ── Browser + local-loopback login ────────────────────────────────────────────
#
# Gmail's scopes are Google-classified as sensitive/restricted, which Google's
# device-code (TV/limited-input) grant rejects outright with `invalid_scope` —
# confirmed by testing against the live endpoint, not just documentation. The
# sanctioned flow for a restricted-scope CLI/installed app is this one: open a
# browser for consent, catch the redirect on a short-lived localhost server,
# exchange the code for tokens. Requires a "Desktop app" OAuth client, not
# "TVs and Limited Input devices".

class _CallbackHandler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        params = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        self.server.auth_result = params
        self.send_response(200)
        self.send_header("Content-Type", "text/html")
        self.end_headers()
        if "code" in params:
            self.wfile.write(b"<html><body>Login successful. You can close this tab.</body></html>")
        else:
            self.wfile.write(b"<html><body>Login failed. Close this tab and check the terminal.</body></html>")

    def log_message(self, *args):
        pass  # silence default per-request stderr logging


def _run_local_server() -> tuple[http.server.HTTPServer, int]:
    server = http.server.HTTPServer(("127.0.0.1", 0), _CallbackHandler)
    server.auth_result = None
    port = server.server_address[1]
    threading.Thread(target=server.handle_request, daemon=True).start()
    return server, port


def _browser_login() -> dict:
    server, port = _run_local_server()
    redirect_uri = f"http://127.0.0.1:{port}/"

    params = {
        "client_id": _client_id(),
        "redirect_uri": redirect_uri,
        "response_type": "code",
        "scope": SCOPES,
        "access_type": "offline",
        "prompt": "consent",  # force a fresh refresh_token on every login
    }
    auth_url = f"{AUTH_URL}?{urllib.parse.urlencode(params)}"

    print("\n" + "=" * 60)
    print("ONE-TIME LOGIN REQUIRED")
    print("=" * 60)
    print("Opening your browser. If it doesn't open, go to:")
    print(auth_url)
    print("\nSign in as the Gmail account you want to send from.")
    print("Google will warn 'Google hasn't verified this app' — click Advanced")
    print("→ Go to (app name) — expected for a personal-use app you own.")
    print("=" * 60 + "\n")

    webbrowser.open(auth_url)

    deadline = time.time() + LOGIN_TIMEOUT_SECONDS
    while server.auth_result is None and time.time() < deadline:
        time.sleep(0.25)

    if server.auth_result is None:
        raise RuntimeError(f"Timed out waiting for browser login ({LOGIN_TIMEOUT_SECONDS}s).")

    result = server.auth_result
    if "error" in result:
        raise RuntimeError(f"Login failed: {result['error'][0]}")
    if "code" not in result:
        raise RuntimeError("Login failed: no authorization code received.")

    resp = requests.post(
        TOKEN_URL,
        data={
            "client_id": _client_id(),
            "client_secret": _client_secret(),
            "code": result["code"][0],
            "grant_type": "authorization_code",
            "redirect_uri": redirect_uri,
        },
        timeout=TIMEOUT,
    )
    if not resp.ok:
        raise RuntimeError(f"Token exchange failed [{resp.status_code}]: {resp.text}")

    print("Login successful. Token cached.\n")
    return resp.json()


def _refresh_token(refresh_token: str) -> dict:
    resp = requests.post(
        TOKEN_URL,
        data={
            "client_id": _client_id(),
            "client_secret": _client_secret(),
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
        },
        timeout=TIMEOUT,
    )
    return resp.json() if resp.ok else {}


# ── Token acquisition ─────────────────────────────────────────────────────────

def get_access_token() -> str:
    cache = _load_cache()

    if cache.get("access_token") and time.time() < cache.get("expires_at", 0):
        return cache["access_token"]

    if cache.get("refresh_token"):
        result = _refresh_token(cache["refresh_token"])
        if result.get("access_token"):
            cache["access_token"] = result["access_token"]
            cache["expires_at"] = time.time() + result.get("expires_in", 3600) - 60
            _save_cache(cache)
            return cache["access_token"]
        # Refresh token dead — expired (Testing-mode grants expire after 7 days),
        # revoked, or scopes changed. Fall through to a fresh login instead of
        # raising, so a stale cache degrades gracefully rather than crashing.

    result = _browser_login()
    new_cache = {
        "access_token": result["access_token"],
        "refresh_token": result.get("refresh_token", cache.get("refresh_token", "")),
        "expires_at": time.time() + result.get("expires_in", 3600) - 60,
    }
    _save_cache(new_cache)
    return new_cache["access_token"]


def sign_out():
    """Drop the cached token so the next call forces a fresh browser login."""
    path = _token_cache_path()
    if os.path.exists(path):
        os.remove(path)
