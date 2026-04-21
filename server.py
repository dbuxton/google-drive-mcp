#!/usr/bin/env python3
"""
google-docs-mcp — Surgical Google Docs editing for AI agents
=============================================================
Exposes Google Docs and Drive operations as MCP tools.

Core abstraction: search-by-text, not by character index.
LLMs describe WHAT they want to change; the server figures out WHERE.
All edits use real batchUpdate — version history, comments, and
suggestions are preserved.

Tools:
  Editing:  docs_get, docs_search_replace, docs_insert_after,
            docs_insert_before, docs_delete_paragraph, docs_append,
            docs_batch_replace
  Comments: docs_add_comment, docs_read_comments, docs_reply_to_comment,
            docs_resolve_comment, docs_delete_comment
  Manage:   docs_list, docs_create

Auth (in priority order):
  1. GOOGLE_DOCS_MCP_TOKEN env var   — standalone token (auth_setup.py)
  2. GOOGLE_DRIVE_MCP_TOKEN env var  — backward-compatible alias
  3. GOOGLE_DOCS_TOKEN_FILE env var  — legacy gog alias
  4. ~/.google-docs-mcp/token.json   — default standalone location
  5. ~/.google-drive-mcp/token.json  — backward-compatible fallback
  6. GOG_KEYRING_PASSWORD env var    — auto-export from gog CLI

Setup:
  python3 auth_setup.py --credentials ~/credentials.json
  python3 auth_setup.py --credentials ~/credentials.json --headless  # no browser
  GOOGLE_DOCS_MCP_CLIENT_ID=... GOOGLE_DOCS_MCP_CLIENT_SECRET=... python3 auth_setup.py

Transport: stdio (Claude Desktop / OpenClaw MCP config)
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
Surgical Google Docs editing — search by text, never by character index.

EDITING TOOLS:
  docs_get               Read document structure (paragraphs + plain text)
  docs_search_replace    Find and replace (occurrence-targeted or all)
  docs_insert_after      Insert paragraph after anchor text
  docs_insert_before     Insert paragraph before anchor text
  docs_delete_paragraph  Delete paragraphs matching anchor text
  docs_append            Append paragraph at end of document
  docs_batch_replace     Multiple replacements atomically (one API call)

COMMENT TOOLS:
  docs_add_comment       Add comment anchored to specific text
  docs_read_comments     List all comments with anchor/resolve status
  docs_reply_to_comment  Reply to an existing comment
  docs_resolve_comment   Resolve (close) a comment, optionally with reply
  docs_delete_comment    Delete a comment

DOCUMENT MANAGEMENT:
  docs_list              List recent docs (optional search query)
  docs_create            Create a new document

TYPICAL WORKFLOW:
  1. docs_get — read document to understand structure
  2. docs_search_replace / docs_insert_after / etc. — make targeted edits
  3. docs_get — verify changes
  For review: docs_add_comment → docs_read_comments → docs_resolve_comment

All edits preserve version history. Use short, distinctive anchor_text
(a few unique words) for reliable text matching.
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
def docs_insert_after(doc_id: str, anchor: str, text: str, rich: bool = True) -> str:
    """
    Insert a new paragraph immediately after the paragraph containing `anchor`.

    The anchor is matched case-insensitively as a substring of the paragraph.

    Rich formatting is ON by default. Set rich=False to insert literal text.

    Supported rich subset:
      - # / ## / ### headings
      - - item / * item bullet lists
      - 1. item / 1) item numbered lists
      - **bold**, *italic*, ***bold italic***

    Args:
        doc_id: Google Doc ID
        anchor: Text to search for to find the target paragraph
        text:   Text to insert as the new paragraph
        rich:   If True (default), interpret simple markdown-like formatting natively

    Returns:
        JSON with: ok, inserted_after (matched paragraph preview), at_index
    """
    result = docs_edit.insert_after(doc_id, anchor, text, rich=rich)
    return json.dumps(result, indent=2)


@mcp.tool
def docs_insert_before(doc_id: str, anchor: str, text: str, rich: bool = True) -> str:
    """
    Insert a new paragraph immediately before the paragraph containing `anchor`.

    The anchor is matched case-insensitively as a substring of the paragraph.

    Rich formatting is ON by default. Set rich=False to insert literal text.

    Supported rich subset:
      - # / ## / ### headings
      - - item / * item bullet lists
      - 1. item / 1) item numbered lists
      - **bold**, *italic*, ***bold italic***

    Args:
        doc_id: Google Doc ID
        anchor: Text to search for to find the target paragraph
        text:   Text to insert as the new paragraph
        rich:   If True (default), interpret simple markdown-like formatting natively

    Returns:
        JSON with: ok, inserted_before (matched paragraph preview), at_index
    """
    result = docs_edit.insert_before(doc_id, anchor, text, rich=rich)
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
def docs_append(doc_id: str, text: str, rich: bool = True) -> str:
    """
    Append a new paragraph at the end of a Google Doc.

    Rich formatting is ON by default. Set rich=False to append literal text.

    Supported rich subset:
      - # / ## / ### headings
      - - item / * item bullet lists
      - 1. item / 1) item numbered lists
      - **bold**, *italic*, ***bold italic***

    Args:
        doc_id: Google Doc ID
        text:   Text to append as the final paragraph
        rich:   If True (default), interpret simple markdown-like formatting natively

    Returns:
        JSON with: ok, appended (text preview), at_index
    """
    result = docs_edit.append(doc_id, text, rich=rich)
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
    include_anchor_text: bool = True,
    bookmark_jump: bool = False,
) -> str:
    """
    Add a comment anchored to specific text in a Google Doc.

    NOTE: The current implementation uses the Drive comments API plus a Docs
    named range. In the Docs UI this still shows as "Original content deleted"
    rather than a proper inline comment highlight. The comments ARE readable via
    docs_read_comments and the Docs 💬 panel.

    The next path to investigate is Apps Script automation. The OAuth token must
    include the script.* scopes, and the Google account must also have the Apps
    Script API enabled at https://script.google.com/home/usersettings before
    remote Apps Script project creation/execution will work.

    By default the tool also appends the matched anchor text into the comment
    body so the user can understand what was targeted even though the Docs UI
    does not render a proper inline highlight with the current API path.

    When bookmark_jump=True, the tool also uses a persistent Apps Script bridge
    (configured via GOOGLE_DOCS_MCP_APPS_SCRIPT_ID) to create a bookmark at the
    target text, then appends a #bookmark jump URL into the comment body.

    Args:
        doc_id:      Google Doc ID (from the URL: /document/d/{DOC_ID}/edit)
        comment:     The comment text to post
        anchor_text: Exact text in the document to attach the comment to.
                     Use a short, unique phrase (a few words) for reliable matching.
        occurrence:  Which occurrence of anchor_text to use (default 1 = first).
                     Use 2, 3, etc. if the text appears multiple times.
        include_anchor_text: If True (default), append the matched anchor text
                     excerpt into the comment body for human readability.
        bookmark_jump: If True, create an Apps Script bookmark and append a
                     jump URL into the comment body. Requires the Apps Script
                     bridge env var setup documented in the README.

    Returns:
        JSON with: ok, comment_id, anchored_to (matched text), at_index, named_range_id
    """
    result = docs_edit.add_comment(
        doc_id,
        comment,
        anchor_text,
        occurrence,
        include_anchor_text=include_anchor_text,
        bookmark_jump=bookmark_jump,
    )
    return json.dumps(result, indent=2)


@mcp.tool
def docs_reply_to_comment(doc_id: str, comment_id: str, reply: str) -> str:
    """
    Reply to an existing comment on a Google Doc.

    Args:
        doc_id:     Google Doc ID
        comment_id: ID of the comment to reply to (from docs_read_comments)
        reply:      Reply text to post

    Returns:
        JSON with: ok, reply_id, comment_id, content
    """
    from googleapiclient.discovery import build
    creds = docs_edit._load_creds()
    drive = build("drive", "v3", credentials=creds)
    result = drive.replies().create(
        fileId=doc_id,
        commentId=comment_id,
        body={"content": reply},
        fields="id,content,createdTime",
    ).execute()
    return json.dumps({
        "ok": True,
        "reply_id": result.get("id"),
        "comment_id": comment_id,
        "content": result.get("content"),
    }, indent=2)


@mcp.tool
def docs_resolve_comment(doc_id: str, comment_id: str, reply: str = "") -> str:
    """
    Resolve (close) a comment on a Google Doc, optionally with a final reply.

    Args:
        doc_id:     Google Doc ID
        comment_id: ID of the comment to resolve (from docs_read_comments)
        reply:      Optional reply text to post before resolving

    Returns:
        JSON with: ok, comment_id, resolved
    """
    from googleapiclient.discovery import build
    creds = docs_edit._load_creds()
    drive = build("drive", "v3", credentials=creds)

    if reply:
        drive.replies().create(
            fileId=doc_id,
            commentId=comment_id,
            body={"content": reply},
            fields="id",
        ).execute()

    drive.comments().update(
        fileId=doc_id,
        commentId=comment_id,
        body={"resolved": True},
        fields="id,resolved",
    ).execute()
    return json.dumps({"ok": True, "comment_id": comment_id, "resolved": True}, indent=2)


@mcp.tool
def docs_delete_comment(doc_id: str, comment_id: str) -> str:
    """
    Delete a comment from a Google Doc.

    Args:
        doc_id:     Google Doc ID
        comment_id: ID of the comment to delete

    Returns:
        JSON with: ok, comment_id
    """
    from googleapiclient.discovery import build
    creds = docs_edit._load_creds()
    drive = build("drive", "v3", credentials=creds)
    drive.comments().delete(fileId=doc_id, commentId=comment_id).execute()
    return json.dumps({"ok": True, "comment_id": comment_id}, indent=2)


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


def main():
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
