#!/usr/bin/env python3
"""
google-docs-mcp — MCP server for surgical Google Docs editing
=============================================================
Exposes Google Docs operations as MCP tools. Unlike most Google Docs
integrations that only support read/create, this server provides
surgical in-place editing that preserves document history, comments,
and suggestions.

The core abstraction: search-by-text, not by character index.
LLMs describe WHAT they want to change; the server figures out WHERE.

Auth (in priority order):
  1. GOOGLE_DOCS_TOKEN_FILE env var — path to gog-exported token JSON
  2. GOG_KEYRING_PASSWORD env var   — auto-export from gog CLI
  3. GOOGLE_DOCS_CREDENTIALS_JSON   — service account JSON (Workspace admin only)

Transport:
  stdio (for Claude Desktop / OpenClaw MCP config)
"""

from __future__ import annotations

import json
import logging
import os
import sys
from typing import Optional

# Add the docs-edit script to path (if running from the repo alongside that skill)
# Or copy docs_edit.py here — we bundle a copy for standalone use.
sys.path.insert(0, os.path.dirname(__file__))

from fastmcp import FastMCP
import docs_edit

logging.basicConfig(level=logging.WARNING, stream=sys.stderr)

mcp = FastMCP(
    name="google-docs-mcp",
    instructions="""
Google Docs surgical editing tools.

These tools let you read and modify Google Docs in-place, preserving
version history, comments, and suggestions. All operations use text
anchors — you describe what text to find, never deal with character indices.

Typical workflow:
1. docs_get — read the document structure to understand content
2. docs_search_replace / docs_insert_after / etc. — make targeted edits
3. docs_get again — verify changes look right

For listing or creating documents, use gog CLI or Google Drive API.
""".strip(),
)


@mcp.tool
def docs_get(doc_id: str) -> str:
    """
    Read a Google Doc and return its structure as JSON.

    Returns title, list of paragraphs (with text, style, start/end indices),
    and the full plain text. Use this before editing to understand the document.

    Args:
        doc_id: Google Doc ID (from the URL: /document/d/{DOC_ID}/edit)
    """
    result = docs_edit.get(doc_id)
    return json.dumps(result, indent=2)


@mcp.tool
def docs_search_replace(
    doc_id: str,
    find: str,
    replace: str,
    occurrence: int = 1,
    regex: bool = False,
) -> str:
    """
    Find text in a Google Doc and replace a specific occurrence.

    Preserves document history — uses real batchUpdate, not delete-and-rewrite.

    Args:
        doc_id:     Google Doc ID
        find:       Text to search for (or regex pattern if regex=True)
        replace:    Text to replace it with
        occurrence: Which occurrence to replace. 1 = first (default), 2 = second,
                    0 = replace ALL occurrences.
        regex:      If True, treat `find` as a Python regular expression

    Returns:
        JSON with: ok, replaced (original text), at_index, occurrences_found
    """
    result = docs_edit.search_replace(doc_id, find, replace, occurrence, regex)
    return json.dumps(result, indent=2)


@mcp.tool
def docs_insert_after(doc_id: str, anchor: str, text: str) -> str:
    """
    Insert a new paragraph immediately after the paragraph containing `anchor`.

    The anchor is matched case-insensitively as a substring of the paragraph.
    The inserted text becomes a new paragraph with default (Normal Text) style.

    Args:
        doc_id: Google Doc ID
        anchor: Text to search for to find the target paragraph
        text:   Text to insert as the new paragraph

    Returns:
        JSON with: ok, inserted_after (matched paragraph preview), at_index
    """
    result = docs_edit.insert_after(doc_id, anchor, text)
    return json.dumps(result, indent=2)


@mcp.tool
def docs_insert_before(doc_id: str, anchor: str, text: str) -> str:
    """
    Insert a new paragraph immediately before the paragraph containing `anchor`.

    The anchor is matched case-insensitively as a substring of the paragraph.

    Args:
        doc_id: Google Doc ID
        anchor: Text to search for to find the target paragraph
        text:   Text to insert as the new paragraph

    Returns:
        JSON with: ok, inserted_before (matched paragraph preview), at_index
    """
    result = docs_edit.insert_before(doc_id, anchor, text)
    return json.dumps(result, indent=2)


@mcp.tool
def docs_delete_paragraph(doc_id: str, anchor: str) -> str:
    """
    Delete the paragraph(s) containing `anchor` text.

    Deletes ALL paragraphs that contain the anchor string (case-insensitive).
    If the anchor matches only one paragraph, only that paragraph is deleted.

    Args:
        doc_id: Google Doc ID
        anchor: Text to search for in paragraphs to delete

    Returns:
        JSON with: ok, deleted_count, deleted (list of deleted paragraph previews)
    """
    result = docs_edit.delete_paragraph(doc_id, anchor)
    return json.dumps(result, indent=2)


@mcp.tool
def docs_append(doc_id: str, text: str) -> str:
    """
    Append a new paragraph at the end of a Google Doc.

    Args:
        doc_id: Google Doc ID
        text:   Text to append as the final paragraph

    Returns:
        JSON with: ok, appended (text preview), at_index
    """
    result = docs_edit.append(doc_id, text)
    return json.dumps(result, indent=2)


@mcp.tool
def docs_batch_replace(doc_id: str, replacements_json: str) -> str:
    """
    Apply multiple find→replace operations atomically in a single batchUpdate.

    All replacements are applied in one API call (end-of-document first to
    preserve index validity). Either ALL changes succeed, or none do.

    Args:
        doc_id:           Google Doc ID
        replacements_json: JSON array of replacements, e.g.:
                          '[{"find": "Q1", "replace": "Q2"},
                            {"find": "draft", "replace": "final", "occurrence": 0}]'

                          Each item: {"find": str, "replace": str,
                                      "occurrence": int (default 1, 0=all),
                                      "regex": bool (default false)}

    Returns:
        JSON with: ok, applied (count), changes (list of what changed)
    """
    replacements = json.loads(replacements_json)
    result = docs_edit.batch_replace(doc_id, replacements)
    return json.dumps(result, indent=2)


@mcp.tool
def docs_add_comment(
    doc_id: str,
    comment: str,
    anchor_text: str,
    occurrence: int = 1,
) -> str:
    """
    Add a comment anchored to specific text in a Google Doc.

    The comment appears as highlighted text with a sidebar comment — exactly
    like Ctrl+Alt+M in Docs. Unlike quotedFileContent (which shows as
    "original content deleted"), this creates a real named range and attaches
    the comment to it.

    Args:
        doc_id:      Google Doc ID (from the URL: /document/d/{DOC_ID}/edit)
        comment:     The comment text to post
        anchor_text: Exact text in the document to attach the comment to.
                     Use a short, unique phrase (a few words) for reliable matching.
        occurrence:  Which occurrence of anchor_text to use (default 1 = first).
                     Use 2, 3, etc. if the text appears multiple times.

    Returns:
        JSON with: ok, comment_id, anchored_to (matched text), at_index, named_range_id
    """
    result = docs_edit.add_comment(doc_id, comment, anchor_text, occurrence)
    return json.dumps(result, indent=2)


@mcp.tool
def docs_read_comments(doc_id: str, include_resolved: bool = False) -> str:
    """
    Read all comments on a Google Doc.

    Returns each comment with its content, author, anchor status, and whether
    it's resolved or deleted. Useful for auditing what comments exist and
    whether they are properly anchored to text.

    Args:
        doc_id:           Google Doc ID
        include_resolved: Include resolved/deleted comments (default False)

    Returns:
        JSON array of comments with id, content, author, anchored, resolved, deleted,
        anchored_to (the named range id if anchored), created_time
    """
    from googleapiclient.discovery import build
    import json as _json

    creds = docs_edit._load_creds()
    drive = build("drive", "v3", credentials=creds)

    resp = drive.comments().list(
        fileId=doc_id,
        fields="comments(id,content,anchor,resolved,deleted,author,createdTime,quotedFileContent)",
        pageSize=100,
    ).execute()

    results = []
    for c in resp.get("comments", []):
        if not include_resolved and (c.get("deleted") or c.get("resolved")):
            continue
        anchor = c.get("anchor") or ""
        named_range_id = None
        if anchor:
            try:
                a = _json.loads(anchor)
                for part in a.get("a", []):
                    if part.get("t") == "r":
                        named_range_id = part.get("v")
            except Exception:
                pass
        results.append({
            "id": c["id"],
            "content": c.get("content", ""),
            "author": c.get("author", {}).get("displayName", "?"),
            "anchored": bool(anchor),
            "named_range_id": named_range_id,
            "quoted_text": (c.get("quotedFileContent") or {}).get("value", ""),
            "resolved": c.get("resolved", False),
            "deleted": c.get("deleted", False),
            "created": c.get("createdTime", ""),
        })

    return json.dumps({"count": len(results), "comments": results}, indent=2)


@mcp.tool
def docs_list(query: str = "", limit: int = 20) -> str:
    """
    List Google Docs from Drive, optionally filtered by a search query.

    Args:
        query: Optional search terms (searches title and content)
        limit: Maximum number of results (default 20)

    Returns:
        JSON array of {id, name, modifiedTime, webViewLink}
    """
    from googleapiclient.discovery import build
    creds = docs_edit._load_creds()
    drive = build("drive", "v3", credentials=creds)

    q = 'mimeType="application/vnd.google-apps.document" and trashed=false'
    if query:
        q += f' and fullText contains "{query}"'

    results = drive.files().list(
        q=q,
        pageSize=min(limit, 100),
        fields="files(id, name, modifiedTime, webViewLink)",
        orderBy="modifiedTime desc",
    ).execute()

    files = results.get("files", [])
    return json.dumps(files, indent=2)


@mcp.tool
def docs_create(title: str, initial_text: str = "") -> str:
    """
    Create a new Google Doc with an optional initial paragraph of text.

    Args:
        title:        Document title
        initial_text: Optional first paragraph content

    Returns:
        JSON with: id, title, webViewLink
    """
    creds = docs_edit._load_creds()
    from googleapiclient.discovery import build

    service = build("docs", "v1", credentials=creds)
    doc = service.documents().create(body={"title": title}).execute()
    doc_id = doc["documentId"]

    if initial_text:
        service.documents().batchUpdate(
            documentId=doc_id,
            body={
                "requests": [{
                    "insertText": {
                        "location": {"index": 1},
                        "text": initial_text,
                    }
                }]
            },
        ).execute()

    return json.dumps({
        "id": doc_id,
        "title": title,
        "webViewLink": f"https://docs.google.com/document/d/{doc_id}/edit",
    }, indent=2)


if __name__ == "__main__":
    mcp.run(transport="stdio")
