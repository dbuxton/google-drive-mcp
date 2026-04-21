#!/usr/bin/env python3
"""
auth_setup.py — Standalone OAuth setup for google-docs-mcp
============================================================
Works in three modes:
  1. Local browser  — opens a browser on this machine (default)
  2. Headless       -- prints a URL, you visit on any device, paste back the redirect URL
  3. Code exchange  — you already have an auth code, just exchange it

Usage:
    python3 auth_setup.py --credentials ~/credentials.json
    python3 auth_setup.py --credentials ~/credentials.json --headless
    python3 auth_setup.py --credentials ~/credentials.json --code "4/0Afr..."
    GOOGLE_DOCS_MCP_CLIENT_ID=... GOOGLE_DOCS_MCP_CLIENT_SECRET=... python3 auth_setup.py

Get credentials.json:
    https://console.cloud.google.com/
    → APIs & Services → Credentials
    → Create OAuth 2.0 Client ID → Desktop App → Download JSON

Required APIs to enable in your project:
    - Google Docs API
    - Google Drive API
    - Google Apps Script API
"""

import argparse
import json
import os
import sys
import threading
import urllib.parse
import urllib.request
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

DEFAULT_TOKEN_PATH = Path.home() / ".google-docs-mcp" / "token.json"

SCOPES = [
    # Core
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.readonly",
    # Comments
    "https://www.googleapis.com/auth/drive.file",
    # Apps Script (required for inline-anchored comments)
    "https://www.googleapis.com/auth/script.projects",
    "https://www.googleapis.com/auth/script.deployments",
    "https://www.googleapis.com/auth/script.processes",
    # Identity
    "openid",
    "email",
    "profile",
]

REDIRECT_PORT = 14399
REDIRECT_URI = f"http://127.0.0.1:{REDIRECT_PORT}/oauth2/callback"
CLIENT_ID_ENV_VAR = "GOOGLE_DOCS_MCP_CLIENT_ID"
CLIENT_SECRET_ENV_VAR = "GOOGLE_DOCS_MCP_CLIENT_SECRET"
CLIENT_ID_ENV_ALIASES = (CLIENT_ID_ENV_VAR, "GOOGLE_DRIVE_MCP_CLIENT_ID")
CLIENT_SECRET_ENV_ALIASES = (CLIENT_SECRET_ENV_VAR, "GOOGLE_DRIVE_MCP_CLIENT_SECRET")


def _first_env(*names: str) -> str | None:
    for name in names:
        value = os.environ.get(name)
        if value:
            return value
    return None


def load_client_config(creds_path: Path) -> dict:
    raw = json.loads(creds_path.read_text())
    # Google downloads credentials.json in {installed: {...}} or {web: {...}} wrapper
    if "installed" in raw:
        return raw["installed"]
    if "web" in raw:
        return raw["web"]
    # Bare format (client_id / client_secret at root)
    if "client_id" in raw:
        return raw
    raise ValueError(
        "credentials.json format not recognised. Expected {installed: ...} or {web: ...} "
        "from Google Cloud Console → Credentials → Download JSON."
    )


def build_auth_url(client_id: str, state: str = "auth") -> str:
    params = {
        "client_id": client_id,
        "redirect_uri": REDIRECT_URI,
        "response_type": "code",
        "scope": " ".join(SCOPES),
        "access_type": "offline",
        "prompt": "consent",
        "state": state,
    }
    return "https://accounts.google.com/o/oauth2/v2/auth?" + urllib.parse.urlencode(params)


def exchange_code(code: str, client_id: str, client_secret: str) -> dict:
    """Exchange an auth code for tokens."""
    data = urllib.parse.urlencode({
        "code": code,
        "client_id": client_id,
        "client_secret": client_secret,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code",
    }).encode()
    req = urllib.request.Request(
        "https://oauth2.googleapis.com/token",
        data=data,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read())


def get_email(access_token: str) -> str:
    req = urllib.request.Request(
        "https://www.googleapis.com/oauth2/v2/userinfo",
        headers={"Authorization": f"Bearer {access_token}"},
    )
    try:
        with urllib.request.urlopen(req) as resp:
            return json.loads(resp.read()).get("email", "")
    except Exception:
        return ""


def save_token(token_response: dict, client_id: str, client_secret: str, out_path: Path):
    email = get_email(token_response.get("access_token", ""))
    token_data = {
        "email": email,
        "client_id": client_id,
        "client_secret": client_secret,
        "refresh_token": token_response["refresh_token"],
        "scopes": SCOPES,
    }
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(token_data, indent=2))
    out_path.chmod(0o600)
    return email


class _CallbackHandler(BaseHTTPRequestHandler):
    code = None
    error = None

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        params = dict(urllib.parse.parse_qsl(parsed.query))
        if "code" in params:
            _CallbackHandler.code = params["code"]
            self.send_response(200)
            self.send_header("Content-Type", "text/html")
            self.end_headers()
            self.wfile.write(b"""
                <html><body style="font-family:sans-serif;padding:40px">
                <h2>&#10003; Authorised</h2>
                <p>You can close this tab and return to the terminal.</p>
                </body></html>
            """)
        else:
            _CallbackHandler.error = params.get("error", "unknown error")
            self.send_response(400)
            self.send_header("Content-Type", "text/html")
            self.end_headers()
            self.wfile.write(f"<html><body>Error: {_CallbackHandler.error}</body></html>".encode())

    def log_message(self, *args):
        pass  # silence request logs


def run_local_flow(client_id: str, client_secret: str, out_path: Path):
    """Start local server, open browser, wait for callback."""
    import webbrowser
    server = HTTPServer(("127.0.0.1", REDIRECT_PORT), _CallbackHandler)
    auth_url = build_auth_url(client_id)

    print(f"\nOpening browser for Google authorisation...")
    print(f"If it doesn't open automatically, visit:\n\n  {auth_url}\n")
    webbrowser.open(auth_url)

    server.handle_request()  # blocks until one request arrives

    if _CallbackHandler.error:
        print(f"\nAuthorisation failed: {_CallbackHandler.error}", file=sys.stderr)
        sys.exit(1)
    if not _CallbackHandler.code:
        print("\nNo auth code received.", file=sys.stderr)
        sys.exit(1)

    tokens = exchange_code(_CallbackHandler.code, client_id, client_secret)
    email = save_token(tokens, client_id, client_secret, out_path)
    return email


def run_headless_flow(client_id: str, client_secret: str, out_path: Path):
    """Print URL, wait for user to paste the redirect URL back."""
    auth_url = build_auth_url(client_id)

    print("\n" + "=" * 60)
    print("STEP 1: Open this URL in any browser (phone, laptop, etc.):")
    print("=" * 60)
    print(f"\n  {auth_url}\n")
    print("=" * 60)
    print("\nSTEP 2: After authorising, Google will redirect to a URL")
    print("         starting with http://127.0.0.1:14399/oauth2/callback?...")
    print("         Copy the FULL redirect URL and paste it here:\n")

    redirect_url = input("Paste redirect URL: ").strip()

    parsed = urllib.parse.urlparse(redirect_url)
    params = dict(urllib.parse.parse_qsl(parsed.query))
    code = params.get("code")
    if not code:
        # Maybe they pasted just the code
        if redirect_url.startswith("4/"):
            code = redirect_url
        else:
            print("\nCould not extract auth code from URL.", file=sys.stderr)
            sys.exit(1)

    tokens = exchange_code(code, client_id, client_secret)
    email = save_token(tokens, client_id, client_secret, out_path)
    return email


def run_code_exchange(code: str, client_id: str, client_secret: str, out_path: Path):
    """Exchange a code that was already obtained."""
    tokens = exchange_code(code, client_id, client_secret)
    email = save_token(tokens, client_id, client_secret, out_path)
    return email


def main():
    parser = argparse.ArgumentParser(
        description="Set up Google OAuth credentials for google-docs-mcp",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Normal (local browser opens automatically):
  python3 auth_setup.py --credentials ~/credentials.json

  # Or use env vars instead of a credentials file:
  GOOGLE_DOCS_MCP_CLIENT_ID=... GOOGLE_DOCS_MCP_CLIENT_SECRET=... python3 auth_setup.py

  # Headless / remote server (prints URL, paste back redirect):
  python3 auth_setup.py --credentials ~/credentials.json --headless

  # Already have an auth code:
  python3 auth_setup.py --credentials ~/credentials.json --code "4/0Afr..."

  # Custom token output path:
  python3 auth_setup.py --credentials ~/credentials.json --out ~/my-token.json
        """,
    )
    parser.add_argument(
        "--credentials", "-c",
        help="Path to credentials.json from Google Cloud Console (OAuth 2.0 Client ID → Desktop App)",
    )
    parser.add_argument(
        "--client-id",
        help=f"OAuth client ID (alternative to --credentials file; defaults to ${CLIENT_ID_ENV_VAR} if set)",
    )
    parser.add_argument(
        "--client-secret",
        help=f"OAuth client secret (alternative to --credentials file; defaults to ${CLIENT_SECRET_ENV_VAR} if set)",
    )
    parser.add_argument(
        "--out", "-o",
        default=str(DEFAULT_TOKEN_PATH),
        help=f"Output path for token file (default: {DEFAULT_TOKEN_PATH})",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Headless mode: print URL, paste back the redirect URL",
    )
    parser.add_argument(
        "--code",
        help="Exchange an existing auth code directly (skips browser flow)",
    )
    args = parser.parse_args()

    env_client_id = _first_env(*CLIENT_ID_ENV_ALIASES)
    env_client_secret = _first_env(*CLIENT_SECRET_ENV_ALIASES)
    client_id = args.client_id or env_client_id
    client_secret = args.client_secret or env_client_secret

    # Resolve credentials
    if args.credentials:
        creds_path = Path(args.credentials).expanduser()
        if not creds_path.exists():
            print(f"Error: credentials file not found: {creds_path}", file=sys.stderr)
            sys.exit(1)
        cfg = load_client_config(creds_path)
        client_id = cfg["client_id"]
        client_secret = cfg["client_secret"]
    elif client_id and client_secret:
        pass
    else:
        print(
            "Error: provide --credentials, both --client-id and --client-secret, "
            f"or set both {CLIENT_ID_ENV_VAR} and {CLIENT_SECRET_ENV_VAR} "
            "(legacy GOOGLE_DRIVE_MCP_* aliases also work)",
            file=sys.stderr,
        )
        sys.exit(1)

    out_path = Path(args.out).expanduser()

    print(f"Token will be saved to: {out_path}")
    print(f"Scopes: docs, drive, appscript (for inline comments)")

    # Run the appropriate flow
    if args.code:
        email = run_code_exchange(args.code, client_id, client_secret, out_path)
    elif args.headless:
        email = run_headless_flow(client_id, client_secret, out_path)
    else:
        email = run_local_flow(client_id, client_secret, out_path)

    print(f"\n✓ Authenticated as: {email or '(email unknown)'}")
    print(f"✓ Token saved: {out_path}")
    print()
    print("Set this env var to use the token:")
    print(f"  export GOOGLE_DOCS_MCP_TOKEN={out_path}")
    print("  # Legacy alias still works: export GOOGLE_DRIVE_MCP_TOKEN=...")
    print()
    print("Or add to your MCP config:")
    print(f"""  "env": {{
    "GOOGLE_DOCS_MCP_TOKEN": "{out_path}"
  }}""")


if __name__ == "__main__":
    main()
