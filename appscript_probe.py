#!/usr/bin/env python3
"""
appscript_probe.py
==================

Stdlib-only helper for probing whether Apps Script offers a viable route for
proper inline Google Docs comments.

Why this exists:
- The current Drive comment + named-range path renders as "Original content
  deleted" in the Docs UI.
- The next promising route is Apps Script automation.
- Before wiring that into google-drive-mcp, we want a concrete probe that can:
  1. create an Apps Script project,
  2. upload a tiny script,
  3. execute it remotely,
  4. report whether DocumentApp exposes any comment-related methods.

Prerequisites:
- An existing google-drive-mcp token with script.* scopes.
- The Google account must have the Apps Script API enabled at:
  https://script.google.com/home/usersettings
- For scripts.run to execute successfully, the temporary script must use the
  same standard Google Cloud project as the OAuth client. A default Apps Script
  project is not enough for remote execution.

Usage:
    python3 appscript_probe.py inspect-comment-api --doc-id <DOC_ID>
"""

from __future__ import annotations

import argparse
import json
import sys
import urllib.error
import urllib.parse
import urllib.request

import docs_edit

APPS_SCRIPT_SETTINGS_URL = "https://script.google.com/home/usersettings"
CLOUD_PROJECTS_GUIDE_URL = "https://developers.google.com/apps-script/guides/cloud-platform-projects"
TOKEN_URL = "https://oauth2.googleapis.com/token"
SCRIPT_API_BASE = "https://script.googleapis.com/v1"


class AppsScriptProbeError(RuntimeError):
    pass


class AppsScriptApiDisabledError(AppsScriptProbeError):
    pass


def _refresh_access_token() -> str:
    token_data = docs_edit._load_token()
    payload = urllib.parse.urlencode(
        {
            "client_id": token_data["client_id"],
            "client_secret": token_data["client_secret"],
            "refresh_token": token_data["refresh_token"],
            "grant_type": "refresh_token",
        }
    ).encode("utf-8")
    req = urllib.request.Request(TOKEN_URL, data=payload, method="POST")
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read().decode("utf-8"))["access_token"]


def _api_request(access_token: str, method: str, path: str, body: dict | None = None) -> dict:
    req = urllib.request.Request(f"{SCRIPT_API_BASE}{path}", method=method)
    req.add_header("Authorization", f"Bearer {access_token}")
    req.add_header("Content-Type", "application/json; charset=utf-8")
    if body is not None:
        req.data = json.dumps(body).encode("utf-8")
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            raw = resp.read().decode("utf-8")
            return json.loads(raw) if raw else {}
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8")
        try:
            payload = json.loads(raw)
        except json.JSONDecodeError:
            payload = {"error": {"message": raw}}
        message = payload.get("error", {}).get("message", raw)
        if e.code == 403 and "User has not enabled the Apps Script API" in message:
            raise AppsScriptApiDisabledError(
                "Apps Script API is not enabled for this Google account. "
                f"Open {APPS_SCRIPT_SETTINGS_URL} and enable it, then rerun this probe."
            ) from None
        if e.code == 403 and "The caller does not have permission" in message:
            raise AppsScriptProbeError(
                "Apps Script execution is now getting past account-level enablement, but the API executable "
                "still cannot run because the script project is not using the same standard Google Cloud "
                "project as the OAuth client. A default Apps Script project is not enough for scripts.run. "
                f"See {CLOUD_PROJECTS_GUIDE_URL}."
            ) from None
        raise AppsScriptProbeError(f"Apps Script API request failed ({e.code}): {message}") from None


def _build_probe_files() -> list[dict]:
    code = r'''
function inspectDocumentCommentApi(docId) {
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();
  var text = body.editAsText();
  var rangeBuilder = doc.newRange();

  function commentMembers(obj) {
    return Object
      .getOwnPropertyNames(Object.getPrototypeOf(obj))
      .filter(function(name) { return /comment/i.test(name); })
      .sort();
  }

  return {
    documentCommentMembers: commentMembers(doc),
    bodyCommentMembers: commentMembers(body),
    textCommentMembers: commentMembers(text),
    rangeBuilderCommentMembers: commentMembers(rangeBuilder),
    documentHasAddComment: typeof doc.addComment,
    bodyHasAddComment: typeof body.addComment,
    textHasAddComment: typeof text.addComment,
    rangeBuilderHasAddComment: typeof rangeBuilder.addComment
  };
}
'''.strip()

    manifest = {
        "timeZone": "Europe/London",
        "exceptionLogging": "STACKDRIVER",
        "runtimeVersion": "V8",
        "oauthScopes": ["https://www.googleapis.com/auth/documents"],
        "executionApi": {"access": "MYSELF"},
    }

    return [
        {"name": "Code", "type": "SERVER_JS", "source": code},
        {"name": "appsscript", "type": "JSON", "source": json.dumps(manifest)},
    ]


def inspect_comment_api(
    doc_id: str,
    *,
    title: str = "google-drive-mcp comment probe",
    script_id: str | None = None,
) -> dict:
    access_token = _refresh_access_token()

    if script_id is None:
        project = _api_request(
            access_token,
            "POST",
            "/projects",
            {"title": title},
        )
        script_id = project["scriptId"]

    _api_request(
        access_token,
        "PUT",
        f"/projects/{script_id}/content",
        {"files": _build_probe_files()},
    )

    version = _api_request(
        access_token,
        "POST",
        f"/projects/{script_id}/versions",
        {"description": "Comment API probe"},
    )

    deployment = _api_request(
        access_token,
        "POST",
        f"/projects/{script_id}/deployments",
        {
            "versionNumber": version["versionNumber"],
            "manifestFileName": "appsscript",
            "description": "API executable for comment probing",
        },
    )
    deployment_id = deployment["deploymentId"]

    result = _api_request(
        access_token,
        "POST",
        f"/scripts/{deployment_id}:run",
        {
            "function": "inspectDocumentCommentApi",
            "parameters": [doc_id],
            "devMode": True,
        },
    )

    return {
        "script_id": script_id,
        "deployment_id": deployment_id,
        "version_number": version["versionNumber"],
        "execution": result,
    }


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Probe Apps Script comment support for Google Docs.")
    sub = parser.add_subparsers(dest="command", required=True)

    inspect = sub.add_parser("inspect-comment-api", help="Check whether DocumentApp exposes comment methods.")
    inspect.add_argument("--doc-id", required=True, help="Google Doc ID to open during the probe")
    inspect.add_argument("--title", default="google-drive-mcp comment probe", help="Temporary Apps Script project title")
    inspect.add_argument("--script-id", help="Existing Apps Script project ID to update and execute instead of creating a temporary one")
    return parser


def main() -> int:
    parser = _build_parser()
    args = parser.parse_args()

    try:
        if args.command == "inspect-comment-api":
            result = inspect_comment_api(
                args.doc_id,
                title=args.title,
                script_id=args.script_id,
            )
            print(json.dumps(result, indent=2))
            return 0
    except AppsScriptApiDisabledError as e:
        print(str(e), file=sys.stderr)
        return 2
    except AppsScriptProbeError as e:
        print(str(e), file=sys.stderr)
        return 1

    parser.error(f"Unknown command: {args.command}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
