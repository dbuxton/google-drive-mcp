# google-drive-mcp

> Surgical Google Docs editing for AI agents — preserves history, never touches character indices.

An MCP server that makes Google Docs actually usable for LLMs. **Standalone** — no other tools required beyond a Google Cloud OAuth app.

## Why

The Google Docs API uses **character indices** for every edit. LLMs are bad at counting characters. Everyone ends up deleting and rewriting entire documents, which destroys version history, comments, and collaborator attribution.

This server uses the same abstraction as code editors: **search by text, not by position**. You describe *what* to change; the server finds *where* it is and handles the index arithmetic.

---

## Quick start

### 1. Run directly from GitHub

```bash
uvx --from git+https://github.com/dbuxton/google-drive-mcp google-drive-mcp --help
```

That command downloads the package, creates an isolated environment, installs dependencies, and runs the `google-drive-mcp` entry point.

No PyPI release is required. `uvx` can execute the tool straight from the GitHub repository.

To pin to a branch, tag, or commit, add a ref to the URL:

```bash
uvx --from git+https://github.com/dbuxton/google-drive-mcp@main google-drive-mcp --help
```

If you prefer a persistent local install instead of `uvx`, use:

```bash
uv tool install --from git+https://github.com/dbuxton/google-drive-mcp google-drive-mcp
uv tool install --from git+https://github.com/dbuxton/google-drive-mcp google-drive-mcp-auth
```

### 2. Create a Google Cloud OAuth app

1. Go to [console.cloud.google.com](https://console.cloud.google.com/)
2. Create a project (or select an existing one)
3. Enable these APIs (**APIs & Services → Library**):
   - **Google Docs API**
   - **Google Drive API**
   - **Google Apps Script API** *(for inline-anchored comments)*
4. Go to **APIs & Services → Credentials**
5. Click **Create Credentials → OAuth 2.0 Client ID**
6. Application type: **Desktop App**
7. Download the JSON file

### 3. Authenticate

**Normal — browser opens automatically:**
```bash
uvx --from git+https://github.com/dbuxton/google-drive-mcp \
  google-drive-mcp-auth --credentials ~/credentials.json
```

**Headless / remote server — no browser on device:**
```bash
uvx --from git+https://github.com/dbuxton/google-drive-mcp \
  google-drive-mcp-auth --credentials ~/credentials.json --headless
# Prints a URL → open on any device (phone, laptop, etc.)
# Paste the full redirect URL back into the terminal
```

**Already have an auth code:**
```bash
uvx --from git+https://github.com/dbuxton/google-drive-mcp \
  google-drive-mcp-auth --credentials ~/credentials.json --code "4/0Afr..."
```

Token is saved to `~/.google-drive-mcp/token.json` by default. Override with `--out /path/to/token.json`.

### 4. Configure your MCP client

**Claude Desktop** (`~/Library/Application Support/Claude/claude_desktop_config.json`):
```json
{
  "mcpServers": {
    "google-drive": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/dbuxton/google-drive-mcp", "google-drive-mcp"],
      "env": {
        "GOOGLE_DRIVE_MCP_TOKEN": "/Users/you/.google-drive-mcp/token.json"
      }
    }
  }
}
```

**OpenClaw** (gateway config):
```json
{
  "mcp": {
    "servers": {
      "google-drive": {
        "command": "uvx",
        "args": ["--from", "git+https://github.com/dbuxton/google-drive-mcp", "google-drive-mcp"],
        "env": {
          "GOOGLE_DRIVE_MCP_TOKEN": "~/.google-drive-mcp/token.json"
        }
      }
    }
  }
}
```

**Local checkout during development:**
```json
{
  "mcpServers": {
    "google-drive": {
      "command": "uvx",
      "args": ["--from", "/absolute/path/to/google-drive-mcp", "google-drive-mcp"],
      "env": {
        "GOOGLE_DRIVE_MCP_TOKEN": "/Users/you/.google-drive-mcp/token.json"
      }
    }
  }
}
```

---

## Tools reference

### Document reading

#### `docs_get(doc_id)`
Read a Google Doc and return its full structure.

Returns the document title, a list of paragraphs (with text, heading style, and character indices), and the full plain text. Use this first to understand the document before making edits.

```
docs_get(doc_id="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms")
```

Returns:
```json
{
  "title": "My Document",
  "paragraphs": [
    {"text": "Introduction", "style": "HEADING_1", "start": 0, "end": 13},
    {"text": "This is the body.", "style": "NORMAL_TEXT", "start": 13, "end": 31}
  ],
  "full_text": "Introduction\nThis is the body."
}
```

---

#### `docs_list(query, limit)`
List Google Docs from Drive, optionally filtered by a search query.

| Param | Type | Default | Description |
|-------|------|---------|-------------|
| `query` | string | `""` | Search terms (searches title and content) |
| `limit` | int | `20` | Maximum results |

```
docs_list(query="board deck 2026", limit=5)
```

---

### Document editing

All editing tools use **text anchors**, never character indices. The server finds the text and handles the indices internally.

#### `docs_search_replace(doc_id, find, replace, occurrence, regex)`
Find text in a document and replace a specific occurrence.

| Param | Type | Default | Description |
|-------|------|---------|-------------|
| `doc_id` | string | required | Google Doc ID |
| `find` | string | required | Text to find |
| `replace` | string | required | Replacement text |
| `occurrence` | int | `1` | Which occurrence: `1` = first, `2` = second, `0` = **all** |
| `regex` | bool | `false` | Treat `find` as a Python regex |

```
# Replace first occurrence
docs_search_replace(doc_id="...", find="Q1 2024", replace="Q2 2024")

# Replace all occurrences
docs_search_replace(doc_id="...", find="ACME Corp", replace="Initech", occurrence=0)

# Regex replace
docs_search_replace(doc_id="...", find=r"\bDraft\b", replace="Final", regex=true)
```

---

#### `docs_insert_after(doc_id, anchor, text)`
Insert a new paragraph immediately after the paragraph containing `anchor`.

```
docs_insert_after(
  doc_id="...",
  anchor="Executive Summary",
  text="Updated as of March 2026 following board review."
)
```

---

#### `docs_insert_before(doc_id, anchor, text)`
Insert a new paragraph immediately before the paragraph containing `anchor`.

```
docs_insert_before(
  doc_id="...",
  anchor="Appendix A",
  text="See the following appendix for supporting data."
)
```

---

#### `docs_delete_paragraph(doc_id, anchor)`
Delete all paragraphs containing `anchor` text (case-insensitive).

```
docs_delete_paragraph(doc_id="...", anchor="[PLACEHOLDER — DELETE ME]")
```

---

#### `docs_append(doc_id, text)`
Append a new paragraph at the end of the document.

```
docs_append(doc_id="...", text="Document last updated: March 2026.")
```

---

#### `docs_batch_replace(doc_id, replacements_json)`
Apply multiple find→replace operations **atomically** in a single API call. Either all changes succeed or none do.

```
docs_batch_replace(
  doc_id="...",
  replacements_json='[
    {"find": "[CLIENT]", "replace": "Acme Corp", "occurrence": 0},
    {"find": "[DATE]", "replace": "10 March 2026", "occurrence": 0},
    {"find": "DRAFT", "replace": "FINAL"}
  ]'
)
```

Each item in the array:

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `find` | string | required | Text to find |
| `replace` | string | required | Replacement text |
| `occurrence` | int | `1` | `1` = first, `0` = all |
| `regex` | bool | `false` | Regex mode |

---

#### `docs_create(title, initial_text)`
Create a new Google Doc.

```
docs_create(title="Q2 Board Deck", initial_text="Confidential — not for distribution.")
```

Returns `{id, title, webViewLink}`.

---

### Comments

#### `docs_add_comment(doc_id, comment, anchor_text, occurrence)`
Add a comment anchored to specific text in the document.

| Param | Type | Default | Description |
|-------|------|---------|-------------|
| `doc_id` | string | required | Google Doc ID |
| `comment` | string | required | Comment text |
| `anchor_text` | string | required | Text in the document to attach the comment to |
| `occurrence` | int | `1` | Which occurrence of `anchor_text` to use |

Use a short, distinctive phrase for `anchor_text` — a few words that are unique enough to match exactly one location.

```
docs_add_comment(
  doc_id="...",
  anchor_text="unable to perform the Employee's duties",
  comment="Legal risk: 3-month absence threshold may not satisfy Equality Act 2010 duty to make reasonable adjustments before terminating."
)
```

> **Note:** Due to a Drive API limitation, comments created via the API currently show as *"Original content deleted"* in the Docs UI rather than as inline highlights. The comments are fully readable via `docs_read_comments` and the Docs 💬 panel. True inline anchoring requires Apps Script — run auth with `--headless` to include the `script.projects` scope, which will enable an upgraded implementation.

---

#### `docs_read_comments(doc_id, include_resolved)`
List all comments on a document.

| Param | Type | Default | Description |
|-------|------|---------|-------------|
| `doc_id` | string | required | Google Doc ID |
| `include_resolved` | bool | `false` | Include resolved/deleted comments |

Returns an array of comments with `id`, `content`, `author`, `anchored` (bool), `named_range_id`, `quoted_text`, `resolved`, `deleted`, `created`.

```
docs_read_comments(doc_id="...")
```

---

#### `docs_reply_to_comment(doc_id, comment_id, reply)`
Post a reply to an existing comment.

| Param | Type | Default | Description |
|-------|------|---------|-------------|
| `doc_id` | string | required | Google Doc ID |
| `comment_id` | string | required | Comment ID (from `docs_read_comments`) |
| `reply` | string | required | Reply text |

```
docs_reply_to_comment(
  doc_id="...",
  comment_id="AAAB1iPyaUY",
  reply="Agreed — adding Carer's Leave clause before we sign."
)
```

---

#### `docs_resolve_comment(doc_id, comment_id, reply)`
Resolve (close) a comment, optionally posting a final reply first.

| Param | Type | Default | Description |
|-------|------|---------|-------------|
| `doc_id` | string | required | Google Doc ID |
| `comment_id` | string | required | Comment ID |
| `reply` | string | `""` | Optional reply to post before resolving |

```
docs_resolve_comment(
  doc_id="...",
  comment_id="AAAB1iPyaUY",
  reply="Fixed in v2 — carer's leave clause added at 15.1."
)
```

---

#### `docs_delete_comment(doc_id, comment_id)`
Permanently delete a comment.

```
docs_delete_comment(doc_id="...", comment_id="AAAB1iPyaUY")
```

---

## Typical workflows

### Contract review
```
1. docs_get          — read the document
2. docs_add_comment  — flag issues with anchor_text pointing to specific clauses
3. docs_read_comments — audit what's been flagged
4. docs_search_replace — fix straightforward issues directly
5. docs_resolve_comment — close comments as they're addressed
```

### Bulk document update
```
1. docs_list         — find all relevant documents
2. docs_batch_replace — apply changes atomically (e.g. rebrand, date update)
3. docs_get          — verify the result
```

### Collaborative review
```
1. docs_add_comment  — add review notes
2. docs_reply_to_comment — respond to collaborator comments
3. docs_resolve_comment  — close resolved threads
```

---

## Auth environment variables

| Variable | Description |
|----------|-------------|
| `GOOGLE_DRIVE_MCP_TOKEN` | Path to token file (preferred for standalone use) |
| `GOOGLE_DOCS_TOKEN_FILE` | Legacy alias |
| `GOG_KEYRING_PASSWORD` | Auto-export from gog CLI (for personal/OpenClaw use) |

---

## Scopes

The auth setup requests these scopes:

| Scope | Purpose |
|-------|---------|
| `https://www.googleapis.com/auth/documents` | Read and write Google Docs |
| `https://www.googleapis.com/auth/drive` | Access Drive files and comments |
| `https://www.googleapis.com/auth/drive.readonly` | Read Drive file metadata |
| `https://www.googleapis.com/auth/drive.file` | Per-file Drive access |
| `https://www.googleapis.com/auth/script.projects` | Create Apps Script projects (inline comments) |
| `https://www.googleapis.com/auth/script.deployments` | Deploy Apps Script functions |
| `https://www.googleapis.com/auth/script.processes` | View script execution |
| `openid`, `email`, `profile` | Identity |

---

## License

MIT
