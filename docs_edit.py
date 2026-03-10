#!/usr/bin/env python3
"""
docs_edit.py — Surgical Google Docs editor
==========================================
Edits Google Docs in-place using real batchUpdate requests, preserving
document history, comments, and suggestions. LLMs never touch indices.

Usage (CLI):
    docs_edit.py get <docId>
    docs_edit.py search_replace <docId> --find "old" --replace "new" [--occurrence 1] [--regex]
    docs_edit.py insert_after <docId> --anchor "paragraph text" --text "new text"
    docs_edit.py insert_before <docId> --anchor "paragraph text" --text "new text"
    docs_edit.py delete_paragraph <docId> --anchor "paragraph text"
    docs_edit.py append <docId> --text "text to append"
    docs_edit.py batch_replace <docId> --replacements '[{"find":"a","replace":"b"}]'

Auth (in priority order):
    1. GOOGLE_DOCS_TOKEN_FILE  env var → path to gog-exported token JSON
    2. GOG_KEYRING_PASSWORD env var   → auto-export from gog
    3. Default gog token path         → ~/.config/gog/token_export.json

Python API:
    from docs_edit import get, search_replace, insert_after, insert_before,
                          delete_paragraph, append, batch_replace
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import re
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

log = logging.getLogger("docs_edit")

# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------

GOG_CREDENTIALS_PATH = Path("/config/gogcli/credentials.json")
GOG_TOKEN_CACHE = Path("/tmp/docs_edit_token_cache.json")


def _export_gog_token(email: str = "david@harriethq.com") -> dict:
    """Export token from gog keyring using GOG_KEYRING_PASSWORD env var."""
    password = os.environ.get("GOG_KEYRING_PASSWORD")
    if not password:
        raise RuntimeError(
            "GOG_KEYRING_PASSWORD not set. Cannot export gog token. "
            "Set GOOGLE_DOCS_TOKEN_FILE to a pre-exported token path instead."
        )
    env = {**os.environ, "GOG_KEYRING_PASSWORD": password}
    with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as f:
        tmp = f.name
    try:
        result = subprocess.run(
            ["gog", "auth", "tokens", "export", email, "--out", tmp, "--overwrite", "-y"],
            env=env,
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            raise RuntimeError(f"gog token export failed: {result.stderr}")
        return json.loads(Path(tmp).read_text())
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass


def _load_token() -> dict:
    """Load gog-exported token from env, auto-export, or cache."""
    token_file = os.environ.get("GOOGLE_DOCS_TOKEN_FILE")
    if token_file:
        return json.loads(Path(token_file).read_text())

    # Try auto-export (uses GOG_KEYRING_PASSWORD)
    try:
        token = _export_gog_token()
        # Cache it
        GOG_TOKEN_CACHE.write_text(json.dumps(token))
        return token
    except RuntimeError as e:
        log.warning("Auto-export failed: %s", e)

    # Try cache
    if GOG_TOKEN_CACHE.exists():
        log.info("Using cached token from %s", GOG_TOKEN_CACHE)
        return json.loads(GOG_TOKEN_CACHE.read_text())

    raise RuntimeError(
        "No Google credentials found. Set GOOGLE_DOCS_TOKEN_FILE to a gog-exported "
        "token JSON, or set GOG_KEYRING_PASSWORD for auto-export."
    )


def _load_creds():
    """Return refreshed google.oauth2.credentials.Credentials."""
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request

    token_data = _load_token()
    creds_data = json.loads(GOG_CREDENTIALS_PATH.read_text())

    creds = Credentials(
        token=None,
        refresh_token=token_data["refresh_token"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=creds_data["client_id"],
        client_secret=creds_data["client_secret"],
        scopes=[s for s in token_data.get("scopes", []) if s != "email"],  # email scope not always granted
    )
    creds.refresh(Request())
    return creds


def _get_service(api: str = "docs", version: str = "v1"):
    """Build a Google API service client."""
    from googleapiclient.discovery import build
    return build(api, version, credentials=_load_creds())


# ---------------------------------------------------------------------------
# Document structure helpers
# ---------------------------------------------------------------------------

@dataclass
class TextRun:
    text: str
    start: int  # Document index (inclusive)
    end: int    # Document index (exclusive)


@dataclass
class Paragraph:
    text: str
    style: str
    start: int   # Start of paragraph including leading newline-like element
    end: int     # End of paragraph (exclusive)
    runs: list[TextRun]


def _extract_paragraphs(doc: dict) -> list[Paragraph]:
    """Extract paragraphs with full structure from a Docs API response."""
    paragraphs = []
    for elem in doc.get("body", {}).get("content", []):
        if "paragraph" not in elem:
            continue
        para = elem["paragraph"]
        style = para.get("paragraphStyle", {}).get("namedStyleType", "NORMAL_TEXT")
        runs = []
        full_text = ""
        for pe in para.get("elements", []):
            if "textRun" in pe:
                content = pe["textRun"]["content"]
                runs.append(TextRun(
                    text=content,
                    start=pe["startIndex"],
                    end=pe["endIndex"],
                ))
                full_text += content
        # Strip trailing newline for display (Google Docs always ends paragraphs with \n)
        display_text = full_text.rstrip("\n")
        paragraphs.append(Paragraph(
            text=display_text,
            style=style,
            start=elem["startIndex"],
            end=elem["endIndex"],
            runs=runs,
        ))
    return paragraphs


def _build_full_text_map(paragraphs: list[Paragraph]) -> tuple[str, list[tuple[int, int, int]]]:
    """
    Returns (full_text, text_map) where:
    - full_text is all characters concatenated from all text runs
    - text_map is list of (offset_in_full_text, doc_start_index, length)
      allowing mapping from full_text position → document index
    """
    parts = []
    text_map = []
    offset = 0
    for para in paragraphs:
        for run in para.runs:
            text_map.append((offset, run.start, len(run.text)))
            parts.append(run.text)
            offset += len(run.text)
    return "".join(parts), text_map


def _full_text_pos_to_doc_index(pos: int, text_map: list[tuple[int, int, int]]) -> int:
    """Map a position in the concatenated full_text to a document character index."""
    for (ft_offset, doc_start, length) in text_map:
        if ft_offset <= pos < ft_offset + length:
            return doc_start + (pos - ft_offset)
    raise ValueError(f"Position {pos} is not within any text run in the document.")


# ---------------------------------------------------------------------------
# Core operations
# ---------------------------------------------------------------------------

def get(doc_id: str) -> dict:
    """
    Fetch a Google Doc and return structured representation.

    Returns:
        {
          "title": "Document Title",
          "paragraphs": [
            {"text": "...", "style": "HEADING_1", "start": 0, "end": 45}
          ],
          "plain_text": "full document text..."
        }
    """
    service = _get_service("docs", "v1")
    doc = service.documents().get(documentId=doc_id).execute()
    paragraphs = _extract_paragraphs(doc)
    return {
        "title": doc.get("title", ""),
        "paragraphs": [
            {
                "text": p.text,
                "style": p.style,
                "start": p.start,
                "end": p.end,
            }
            for p in paragraphs
        ],
        "plain_text": "\n".join(p.text for p in paragraphs),
    }


def search_replace(
    doc_id: str,
    find: str,
    replace: str,
    occurrence: int = 1,
    regex: bool = False,
) -> dict:
    """
    Find text in a document and replace a specific occurrence.

    Args:
        doc_id:     Google Doc ID
        find:       Text to find (or regex pattern if regex=True)
        replace:    Replacement text
        occurrence: Which occurrence to replace (1-based). 0 = replace all.
        regex:      Treat `find` as a regular expression

    Returns:
        {"ok": True, "replaced": "old text", "at_index": 45, "occurrences_found": 3}
    """
    service = _get_service("docs", "v1")

    # Replace-all: use the native replaceAllText API (fast, atomic)
    if occurrence == 0 and not regex:
        result = service.documents().batchUpdate(
            documentId=doc_id,
            body={
                "requests": [{
                    "replaceAllText": {
                        "containsText": {"text": find, "matchCase": True},
                        "replaceText": replace,
                    }
                }]
            },
        ).execute()
        count = (
            result.get("replies", [{}])[0]
            .get("replaceAllText", {})
            .get("occurrencesChanged", 0)
        )
        return {"ok": True, "replaced": find, "occurrences_changed": count}

    # Targeted occurrence: find index manually
    doc = service.documents().get(documentId=doc_id).execute()
    paragraphs = _extract_paragraphs(doc)
    full_text, text_map = _build_full_text_map(paragraphs)

    # Build list of (ft_start, ft_end) tuples
    if regex:
        matches = [(m.start(), m.end()) for m in re.finditer(find, full_text)]
    else:
        matches = []
        search_from = 0
        while True:
            pos = full_text.find(find, search_from)
            if pos == -1:
                break
            matches.append((pos, pos + len(find)))
            search_from = pos + 1

    if not matches:
        raise ValueError(f"Text not found in document: {find!r}")

    target_idx = occurrence - 1
    if target_idx >= len(matches):
        raise ValueError(
            f"Occurrence {occurrence} not found. Document has {len(matches)} occurrence(s) of {find!r}"
        )

    ft_start, ft_end = matches[target_idx]

    doc_start = _full_text_pos_to_doc_index(ft_start, text_map)
    doc_end = _full_text_pos_to_doc_index(ft_end - 1, text_map) + 1

    old_text = full_text[ft_start:ft_end]

    # Apply: delete old text, then insert replacement at same position
    requests = []
    if replace:
        requests.append({
            "insertText": {
                "location": {"index": doc_start},
                "text": replace,
            }
        })
        requests.append({
            "deleteContentRange": {
                "range": {
                    "startIndex": doc_start + len(replace),
                    "endIndex": doc_end + len(replace),
                }
            }
        })
    else:
        # Replacing with empty string = pure delete
        requests.append({
            "deleteContentRange": {
                "range": {"startIndex": doc_start, "endIndex": doc_end}
            }
        })

    service.documents().batchUpdate(
        documentId=doc_id,
        body={"requests": requests},
    ).execute()

    return {
        "ok": True,
        "replaced": old_text,
        "at_index": doc_start,
        "occurrences_found": len(matches),
    }


def insert_after(doc_id: str, anchor: str, text: str) -> dict:
    """
    Insert text as a new paragraph after the paragraph containing `anchor`.

    The inserted text becomes a separate paragraph (newline appended automatically).
    """
    service = _get_service("docs", "v1")
    doc = service.documents().get(documentId=doc_id).execute()
    paragraphs = _extract_paragraphs(doc)

    target = None
    for p in paragraphs:
        if anchor.lower() in p.text.lower():
            target = p
            break

    if target is None:
        raise ValueError(f"No paragraph containing anchor: {anchor!r}")

    # Insert after the end of the paragraph (doc end index includes the \n)
    insert_index = target.end - 1  # position of the terminating \n

    service.documents().batchUpdate(
        documentId=doc_id,
        body={
            "requests": [{
                "insertText": {
                    "location": {"index": insert_index},
                    "text": "\n" + text,
                }
            }]
        },
    ).execute()

    return {"ok": True, "inserted_after": target.text[:80], "at_index": insert_index}


def insert_before(doc_id: str, anchor: str, text: str) -> dict:
    """
    Insert text as a new paragraph before the paragraph containing `anchor`.
    """
    service = _get_service("docs", "v1")
    doc = service.documents().get(documentId=doc_id).execute()
    paragraphs = _extract_paragraphs(doc)

    target = None
    for p in paragraphs:
        if anchor.lower() in p.text.lower():
            target = p
            break

    if target is None:
        raise ValueError(f"No paragraph containing anchor: {anchor!r}")

    # Insert at the start of the paragraph
    insert_index = target.start

    service.documents().batchUpdate(
        documentId=doc_id,
        body={
            "requests": [{
                "insertText": {
                    "location": {"index": insert_index},
                    "text": text + "\n",
                }
            }]
        },
    ).execute()

    return {"ok": True, "inserted_before": target.text[:80], "at_index": insert_index}


def delete_paragraph(doc_id: str, anchor: str) -> dict:
    """
    Delete the paragraph(s) containing `anchor` text.

    Deletes ALL paragraphs matching the anchor (case-insensitive substring).
    Returns count of deleted paragraphs.
    """
    service = _get_service("docs", "v1")
    doc = service.documents().get(documentId=doc_id).execute()
    paragraphs = _extract_paragraphs(doc)

    targets = [p for p in paragraphs if anchor.lower() in p.text.lower()]
    if not targets:
        raise ValueError(f"No paragraph containing anchor: {anchor!r}")

    # Sort by start index descending so deleting earlier content doesn't shift later indices
    targets.sort(key=lambda p: p.start, reverse=True)

    requests = []
    for t in targets:
        requests.append({
            "deleteContentRange": {
                "range": {
                    "startIndex": t.start,
                    "endIndex": t.end,
                }
            }
        })

    service.documents().batchUpdate(
        documentId=doc_id,
        body={"requests": requests},
    ).execute()

    return {
        "ok": True,
        "deleted_count": len(targets),
        "deleted": [t.text[:80] for t in targets],
    }


def append(doc_id: str, text: str) -> dict:
    """
    Append text as a new paragraph at the end of the document.
    """
    service = _get_service("docs", "v1")
    doc = service.documents().get(documentId=doc_id).execute()

    # Find the last content index (end of the body, before the body's closing)
    body_content = doc.get("body", {}).get("content", [])
    if not body_content:
        raise ValueError("Document body is empty.")

    # The last element's endIndex is the body end (exclusive).
    # We insert just before that.
    last_elem = body_content[-1]
    insert_index = last_elem["endIndex"] - 1

    service.documents().batchUpdate(
        documentId=doc_id,
        body={
            "requests": [{
                "insertText": {
                    "location": {"index": insert_index},
                    "text": "\n" + text,
                }
            }]
        },
    ).execute()

    return {"ok": True, "appended": text[:80], "at_index": insert_index}


def batch_replace(doc_id: str, replacements: list[dict]) -> dict:
    """
    Apply multiple find→replace operations atomically (all or nothing).

    Replacements are sorted by position (end→start) automatically, so earlier
    indices remain valid throughout the batch.

    Args:
        doc_id:       Google Doc ID
        replacements: List of {"find": "...", "replace": "...", "occurrence": 1}
                      `occurrence` defaults to 1. Use 0 for replace-all.

    Returns:
        {"ok": True, "applied": N, "changes": [...]}
    """
    service = _get_service("docs", "v1")
    doc = service.documents().get(documentId=doc_id).execute()
    paragraphs = _extract_paragraphs(doc)
    full_text, text_map = _build_full_text_map(paragraphs)

    # Collect all changes with their document indices
    changes = []  # (doc_start, doc_end, replace_text, find_text)
    for rep in replacements:
        find = rep["find"]
        replace_text = rep.get("replace", "")
        occurrence = rep.get("occurrence", 1)
        is_regex = rep.get("regex", False)

        # Build list of (ft_start, ft_end) tuples
        if is_regex:
            matches = [(m.start(), m.end()) for m in re.finditer(find, full_text)]
        else:
            matches = []
            search_from = 0
            while True:
                pos = full_text.find(find, search_from)
                if pos == -1:
                    break
                matches.append((pos, pos + len(find)))
                search_from = pos + 1

        if not matches:
            raise ValueError(f"Text not found: {find!r}")

        if occurrence == 0:
            # Replace all
            for ft_start, ft_end in matches:
                ds = _full_text_pos_to_doc_index(ft_start, text_map)
                de = _full_text_pos_to_doc_index(ft_end - 1, text_map) + 1
                changes.append((ds, de, replace_text, full_text[ft_start:ft_end]))
        else:
            idx = occurrence - 1
            if idx >= len(matches):
                raise ValueError(f"Occurrence {occurrence} not found for {find!r}")
            ft_start, ft_end = matches[idx]
            ds = _full_text_pos_to_doc_index(ft_start, text_map)
            de = _full_text_pos_to_doc_index(ft_end - 1, text_map) + 1
            changes.append((ds, de, replace_text, full_text[ft_start:ft_end]))

    # Sort by doc_start DESCENDING (apply end-of-doc first so indices stay valid)
    changes.sort(key=lambda c: c[0], reverse=True)

    # Build batchUpdate requests
    requests = []
    for (ds, de, replace_text, old_text) in changes:
        if replace_text:
            requests.append({
                "insertText": {
                    "location": {"index": ds},
                    "text": replace_text,
                }
            })
            requests.append({
                "deleteContentRange": {
                    "range": {
                        "startIndex": ds + len(replace_text),
                        "endIndex": de + len(replace_text),
                    }
                }
            })
        else:
            requests.append({
                "deleteContentRange": {
                    "range": {"startIndex": ds, "endIndex": de}
                }
            })

    service.documents().batchUpdate(
        documentId=doc_id,
        body={"requests": requests},
    ).execute()

    return {
        "ok": True,
        "applied": len(changes),
        "changes": [
            {"replaced": old, "with": new, "at_index": ds}
            for (ds, de, new, old) in sorted(changes, key=lambda c: c[0])
        ],
    }


def add_comment(
    doc_id: str,
    comment: str,
    anchor_text: str,
    occurrence: int = 1,
) -> dict:
    """
    Add a comment anchored to specific text in a Google Doc.

    Unlike the Drive API's quotedFileContent (which shows as "original content
    deleted"), this creates a real named range in the document and attaches the
    comment to it — so it appears as highlighted text with a sidebar comment,
    exactly like a human adding a comment via Ctrl+Alt+M.

    Steps:
      1. Find anchor_text in the document → get character indices
      2. CreateNamedRangeRequest via batchUpdate → get named_range_id
      3. Drive API comments.create with anchor JSON → attached comment

    Args:
        doc_id:      Google Doc ID
        comment:     Comment text to post
        anchor_text: Text in the document to attach the comment to
        occurrence:  Which occurrence to anchor to (default 1 = first)

    Returns:
        {"ok": True, "comment_id": "...", "anchored_to": "...", "at_index": N,
         "named_range_id": "..."}
    """
    import uuid

    docs_service = _get_service("docs", "v1")
    doc = docs_service.documents().get(documentId=doc_id).execute()
    paragraphs = _extract_paragraphs(doc)
    full_text, text_map = _build_full_text_map(paragraphs)

    # Find all occurrences of anchor_text
    matches = []
    search_from = 0
    while True:
        pos = full_text.find(anchor_text, search_from)
        if pos == -1:
            break
        matches.append((pos, pos + len(anchor_text)))
        search_from = pos + 1

    if not matches:
        raise ValueError(f"Anchor text not found in document: {anchor_text!r}")

    target_idx = occurrence - 1
    if target_idx >= len(matches):
        raise ValueError(
            f"Occurrence {occurrence} not found — document has {len(matches)} "
            f"occurrence(s) of {anchor_text!r}"
        )

    ft_start, ft_end = matches[target_idx]
    doc_start = _full_text_pos_to_doc_index(ft_start, text_map)
    doc_end = _full_text_pos_to_doc_index(ft_end - 1, text_map) + 1

    # Create a named range at the anchor position
    range_name = f"comment-anchor-{uuid.uuid4().hex[:12]}"
    result = docs_service.documents().batchUpdate(
        documentId=doc_id,
        body={
            "requests": [{
                "createNamedRange": {
                    "name": range_name,
                    "range": {
                        "startIndex": doc_start,
                        "endIndex": doc_end,
                        "segmentId": "",  # empty = main document body
                    },
                }
            }]
        },
    ).execute()

    named_range_id = (
        result.get("replies", [{}])[0]
        .get("createNamedRange", {})
        .get("namedRangeId")
    )
    if not named_range_id:
        raise RuntimeError("Failed to create named range — no ID returned by Docs API")

    # Build the Drive API anchor JSON for a Google Docs named range
    anchor_json = json.dumps({
        "r": "head",
        "a": [
            {"t": "g", "v": doc_id},
            {"t": "r", "v": named_range_id},
        ],
    })

    from googleapiclient.discovery import build
    drive = build("drive", "v3", credentials=_load_creds())
    comment_result = drive.comments().create(
        fileId=doc_id,
        body={
            "content": comment,
            "anchor": anchor_json,
        },
        fields="id,content,anchor",
    ).execute()

    return {
        "ok": True,
        "comment_id": comment_result.get("id"),
        "anchored_to": full_text[ft_start:ft_end],
        "at_index": doc_start,
        "named_range_id": named_range_id,
    }


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Surgical Google Docs editor — LLMs never touch indices.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--json", action="store_true", help="Output JSON (default for scripting)")

    sub = p.add_subparsers(dest="command", required=True)

    # get
    g = sub.add_parser("get", help="Get document structure")
    g.add_argument("doc_id")

    # search_replace
    sr = sub.add_parser("search_replace", help="Search and replace text")
    sr.add_argument("doc_id")
    sr.add_argument("--find", required=True)
    sr.add_argument("--replace", required=True)
    sr.add_argument("--occurrence", type=int, default=1, help="Which occurrence (1=first, 0=all)")
    sr.add_argument("--regex", action="store_true")

    # insert_after
    ia = sub.add_parser("insert_after", help="Insert paragraph after anchor")
    ia.add_argument("doc_id")
    ia.add_argument("--anchor", required=True)
    ia.add_argument("--text", required=True)

    # insert_before
    ib = sub.add_parser("insert_before", help="Insert paragraph before anchor")
    ib.add_argument("doc_id")
    ib.add_argument("--anchor", required=True)
    ib.add_argument("--text", required=True)

    # delete_paragraph
    dp = sub.add_parser("delete_paragraph", help="Delete paragraph(s) matching anchor")
    dp.add_argument("doc_id")
    dp.add_argument("--anchor", required=True)

    # append
    ap = sub.add_parser("append", help="Append text to end of document")
    ap.add_argument("doc_id")
    ap.add_argument("--text", required=True)

    # batch_replace
    br = sub.add_parser("batch_replace", help="Multiple replacements (atomic)")
    br.add_argument("doc_id")
    br.add_argument(
        "--replacements",
        required=True,
        help='JSON array: [{"find":"...","replace":"...","occurrence":1}]',
    )

    # add_comment
    ac = sub.add_parser("add_comment", help="Add a comment anchored to specific text")
    ac.add_argument("doc_id")
    ac.add_argument("--anchor", required=True, help="Text in the document to anchor the comment to")
    ac.add_argument("--comment", required=True, help="Comment text to post")
    ac.add_argument("--occurrence", type=int, default=1, help="Which occurrence to anchor to (default 1)")

    return p


def main():
    logging.basicConfig(level=logging.WARNING)
    parser = _build_parser()
    args = parser.parse_args()

    try:
        if args.command == "get":
            result = get(args.doc_id)
        elif args.command == "search_replace":
            result = search_replace(
                args.doc_id, args.find, args.replace, args.occurrence, args.regex
            )
        elif args.command == "insert_after":
            result = insert_after(args.doc_id, args.anchor, args.text)
        elif args.command == "insert_before":
            result = insert_before(args.doc_id, args.anchor, args.text)
        elif args.command == "delete_paragraph":
            result = delete_paragraph(args.doc_id, args.anchor)
        elif args.command == "append":
            result = append(args.doc_id, args.text)
        elif args.command == "batch_replace":
            replacements = json.loads(args.replacements)
            result = batch_replace(args.doc_id, replacements)
        elif args.command == "add_comment":
            result = add_comment(args.doc_id, args.comment, args.anchor, args.occurrence)
        else:
            parser.print_help()
            sys.exit(1)

        print(json.dumps(result, indent=2))

    except Exception as e:
        print(json.dumps({"ok": False, "error": str(e)}), file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
