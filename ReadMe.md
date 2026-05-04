# OneNote MCP Server (SharePoint-aware fork)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A Model Context Protocol (MCP) server that lets Claude (and other MCP-capable
assistants) read, write, search, and edit Microsoft OneNote — including notebooks
that live on SharePoint sites, not just personal OneDrive.

This is a fork of [eshlon/onenotemcp](https://github.com/eshlon/onenotemcp) with
fixes and feature work that were necessary to make it work in real
work-account / SharePoint scenarios. See [What's new in this fork](#whats-new-in-this-fork)
below.

## Features

* **Authentication** — secure device-code flow against Microsoft Graph, with
  configurable tenant.
* **Read** — list notebooks/sections/pages, search by title, fetch page content
  as readable text, raw HTML, or short summary.
* **Write & edit** — create new pages, replace or append page content, rename
  pages, find/replace text, add formatted notes (note / todo / important /
  question), insert tables from CSV.
* **SharePoint sites** — every read/write tool optionally targets a SharePoint
  site instead of `/me`, so shared team notebooks (e.g. Vorstand, Sales,
  Engineering) are first-class.
* **Session context** — set a default site once with `useSite`, and every
  subsequent call uses it. The setting persists across restarts.
* **Robust HTML processing** — JSDOM-based extraction, markdown-style input
  conversion.
* **Zod schemas** for tool input validation.

## What's new in this fork

Compared to the upstream [eshlon/onenotemcp](https://github.com/eshlon/onenotemcp):

| Area | Change |
|---|---|
| **Auth** | Configurable `tenantId` so the server works against single-tenant work apps (the upstream defaults to `organizations` and fails with `invalid_grant` for personal accounts and certain tenant configurations). |
| **Auth crash** | Hardens the `authenticate` flow against unhandled rejections that previously killed the MCP process whenever Azure rejected the device-code grant. |
| **SharePoint** | New tools `searchSites`, `getSiteByUrl`, `listSiteNotebooks`, `listSections` — discover and enumerate notebooks living on SharePoint sites. |
| **Site-scoped tools** | Every existing read/edit/create tool now accepts an optional `siteId` parameter. Path resolution flips between `/me/onenote/...` and `/sites/{id}/onenote/...` automatically. |
| **Session default** | `useSite(siteUrl)`, `useMyOneNote()`, `getCurrentSite()` — set a working site once, and every subsequent call uses it. Persists to `.default-site.json` so it survives Claude Desktop restarts. |
| **Scopes** | Adds `Notes.Read.All`, `Notes.ReadWrite.All`, `Sites.Read.All` so the server can actually read shared/site notebooks. |
| **Misc** | Fixed broken `bugs`/`homepage` URLs in `package.json`; added `.default-site.json` to `.gitignore`. |

A standalone `test-auth.mjs` is included to validate Azure App Registration
configuration outside Claude Desktop — useful for debugging `invalid_grant`
errors.

## Prerequisites

* **Node.js 18+** ([nodejs.org](https://nodejs.org/))
* **Git**
* **Microsoft account** with access to OneNote (personal, work, or school)
* **Azure App Registration** — strongly recommended (the upstream "use the Graph
  Explorer client ID" shortcut works for some flows but reliably fails for
  single-tenant work accounts and any flow that needs `*.All` permissions).

## Installation

```bash
git clone https://github.com/djlagaffe/onenotemcp.git
cd onenotemcp
npm install
```

## Azure App Registration setup

This is the part that bites everyone. Do all of it:

1. **Portal → App registrations → + New registration**
   * Name: anything (e.g. `OneNote MCP`)
   * **Supported account types:**
     * Personal Microsoft account → "Accounts in any organizational directory and personal Microsoft accounts"
     * Single-tenant work/school → "Accounts in this organizational directory only"
     * Multi-tenant + personal → "Accounts in any organizational directory and personal Microsoft accounts"
   * Redirect URI: leave blank
   * Copy the **Application (client) ID** and **Directory (tenant) ID** from the Overview page after registration

2. **Authentication blade**
   * Scroll to **Advanced settings** → **Allow public client flows** → **Yes** → Save

3. **API permissions blade** → **+ Add a permission** → **Microsoft Graph** → **Delegated permissions** — add:
   * `User.Read`
   * `Notes.Read`
   * `Notes.ReadWrite`
   * `Notes.Create`
   * `Notes.Read.All` (required for SharePoint site notebooks)
   * `Notes.ReadWrite.All` (required for editing site notebooks)
   * `Sites.Read.All` (required to look up sites by URL/name)
   * Click **Grant admin consent for &lt;your tenant&gt;**. The `*.All` scopes typically require admin rights — if the button is greyed out, ask whoever administers your M365 tenant to do this step.

4. **Decide which tenant value to use**

   The server reads the tenant from the `AZURE_TENANT_ID` environment variable
   (configured in your Claude Desktop config — see next section). Pick:

   * **Single-tenant work/school app:** your Directory (tenant) ID GUID
     (visible on the Overview page of the app registration)
   * **Multi-tenant or mixed (work + personal):** `common`
   * **Personal-only:** `consumers`

   If you don't set `AZURE_TENANT_ID`, the server defaults to `common`. That's
   fine for personal Microsoft accounts and multi-tenant apps, but
   single-tenant work apps will fail with `invalid_grant` until you set the
   correct tenant GUID.

## Configuring the Claude Desktop MCP entry

### 1. Find the config file

Claude Desktop reads its MCP servers from a JSON file. Open it in any text
editor — if it doesn't exist yet, create it.

| OS | Path |
|---|---|
| **Windows** | `%APPDATA%\Claude\claude_desktop_config.json` |
| **macOS** | `~/Library/Application Support/Claude/claude_desktop_config.json` |
| **Linux** | `~/.config/Claude/claude_desktop_config.json` |

On Windows you can paste `%APPDATA%\Claude\` into File Explorer's address bar
to jump straight there. The folder is normally
`C:\Users\<you>\AppData\Roaming\Claude\`.

### 2. Add the `onenote` entry

If the file is empty, paste this whole block:

```json
{
  "mcpServers": {
    "onenote": {
      "command": "node",
      "args": ["C:\\full\\path\\to\\onenotemcp\\onenote-mcp.mjs"],
      "env": {
        "AZURE_CLIENT_ID": "<your-application-client-id>",
        "AZURE_TENANT_ID": "<your-tenant-id-or-common-or-consumers>"
      }
    }
  }
}
```

If you already have other MCP servers configured, just add the `onenote` key
inside the existing `mcpServers` object — don't nest a second `mcpServers`:

```json
{
  "mcpServers": {
    "filesystem": { "command": "...", "args": ["..."] },
    "onenote": {
      "command": "node",
      "args": ["C:\\full\\path\\to\\onenotemcp\\onenote-mcp.mjs"],
      "env": {
        "AZURE_CLIENT_ID": "<your-application-client-id>",
        "AZURE_TENANT_ID": "<your-tenant-id-or-common-or-consumers>"
      }
    }
  }
}
```

### 3. What each field means

| Field | Value |
|---|---|
| `"onenote"` | The display name shown in Claude Desktop. You can rename this; the rest of the docs just assume it's `onenote`. |
| `command` | `"node"` — the Node.js executable. Must be on your PATH, or use the absolute path (e.g. `"C:\\Program Files\\nodejs\\node.exe"`). |
| `args` | A single-element array containing the absolute path to `onenote-mcp.mjs` from this repo. **On Windows, escape every backslash as `\\`** in JSON. |
| `env.AZURE_CLIENT_ID` | The "Application (client) ID" GUID from your Azure App Registration. Required. |
| `env.AZURE_TENANT_ID` | Your Directory (tenant) ID GUID for single-tenant work apps, `common` for multi-tenant or mixed personal+work, or `consumers` for personal-only. Defaults to `common` if omitted. |

Optional extra `env` entries you may want:

```json
"env": {
  "AZURE_CLIENT_ID": "...",
  "AZURE_TENANT_ID": "...",
  "PATH": "C:\\Program Files\\nodejs;C:\\Windows\\System32"
}
```

The explicit `PATH` entry only matters if `node` isn't being resolved
automatically (you'd see a `spawn node ENOENT` error in the log).

### 4. Worked examples

Replace the GUIDs with your own values from the app registration's Overview
page.

**Windows (using forward slashes — also valid in JSON):**

```json
{
  "mcpServers": {
    "onenote": {
      "command": "node",
      "args": ["D:/path/to/onenotemcp/onenote-mcp.mjs"],
      "env": {
        "AZURE_CLIENT_ID": "00000000-0000-0000-0000-000000000000",
        "AZURE_TENANT_ID": "00000000-0000-0000-0000-000000000000"
      }
    }
  }
}
```

**macOS:**

```json
{
  "mcpServers": {
    "onenote": {
      "command": "node",
      "args": ["/Users/yourname/code/onenotemcp/onenote-mcp.mjs"],
      "env": {
        "AZURE_CLIENT_ID": "00000000-0000-0000-0000-000000000000",
        "AZURE_TENANT_ID": "common"
      }
    }
  }
}
```

### 5. Restart Claude Desktop

Save the file, then **fully quit Claude Desktop** — system tray icon → Quit on
Windows, or ⌘Q on macOS. Closing the window is not enough; the MCP config is
only re-read when the app actually exits and restarts.

When you relaunch and open a chat, you should see OneNote's tools listed
(authenticate, listNotebooks, useSite, …). If they're not there, check the
log:

* **Windows:** `%APPDATA%\Claude\logs\mcp-server-onenote.log`
* **macOS:** `~/Library/Logs/Claude/mcp-server-onenote.log`

The log is the single most useful debugging tool — it shows the spawn command,
the JSON-RPC handshake, and any stderr output from the server.

## First-run authentication

In a Claude Desktop chat:

1. Run the `authenticate` tool. It prints a URL (`https://microsoft.com/devicelogin`) and a 9-character code.
2. Open the URL, enter the code, sign in, and approve the permissions prompt.
3. The token is saved to `.access-token.txt` in the project directory and reused on subsequent runs.

If `authenticate` fails with `invalid_grant`, run `node test-auth.mjs` from the
project directory after setting `AZURE_CLIENT_ID` — it prints the full Azure
error response and is much easier to debug than reading Claude Desktop logs.

## Working with SharePoint site notebooks

```
useSite https://contoso.sharepoint.com/sites/marketing
```

After that, every OneNote tool call (`listNotebooks`, `searchPages`,
`getPageContent`, `appendToPage`, `createPage`, …) targets the marketing site
notebooks by default. The setting persists to `.default-site.json` and is
restored automatically on every server restart.

To look up a site you don't yet have a URL for:

```
searchSites query=Marketing
```

To switch:

```
useSite https://contoso.sharepoint.com/sites/sales
```

To revert to your personal OneNote:

```
useMyOneNote
```

To check which scope is currently active:

```
getCurrentSite
```

You can also pass `siteId=<id>` on any tool call to override the default for
just that call without changing the default.

## Tool reference

### Auth
| Tool | Args |
|---|---|
| `authenticate` | — |
| `saveAccessToken` | — |

### Context
| Tool | Args |
|---|---|
| `useSite` | `siteUrl: string` |
| `useMyOneNote` | — |
| `getCurrentSite` | — |

### Site discovery
| Tool | Args |
|---|---|
| `searchSites` | `query: string` |
| `getSiteByUrl` | `siteUrl: string` |
| `listSiteNotebooks` | `siteId: string` |
| `listSections` | `siteId?: string`, `notebookId?: string` |

### Read
| Tool | Args |
|---|---|
| `listNotebooks` | `siteId?: string` |
| `searchPages` | `query?: string`, `siteId?: string` |
| `getPageContent` | `pageId: string`, `format?: 'text'\|'html'\|'summary'`, `siteId?: string` |
| `getPageByTitle` | `title: string`, `format?: 'text'\|'html'\|'summary'`, `siteId?: string` |

### Edit
| Tool | Args |
|---|---|
| `updatePageContent` | `pageId`, `content`, `preserveTitle?`, `siteId?` |
| `appendToPage` | `pageId`, `content`, `addTimestamp?`, `addSeparator?`, `siteId?` |
| `updatePageTitle` | `pageId`, `newTitle`, `siteId?` |
| `replaceTextInPage` | `pageId`, `findText`, `replaceText`, `caseSensitive?`, `siteId?` |
| `addNoteToPage` | `pageId`, `note`, `noteType?`, `position?`, `siteId?` |
| `addTableToPage` | `pageId`, `tableData (CSV)`, `title?`, `position?`, `siteId?` |

### Create
| Tool | Args |
|---|---|
| `createPage` | `title`, `content`, `siteId?`, `sectionId?` |

For every tool that takes `siteId`, omitting it falls back to (in order): the
session default set by `useSite`, otherwise `/me/onenote`.

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| `invalid_grant` during `authenticate` | Tenant mismatch (single-tenant app + wrong tenantId), or missing "Allow public client flows" |
| `AADSTS65001: needs admin consent` | Click "Grant admin consent" in the API permissions blade — admin only |
| `listNotebooks` returns nothing for a known site notebook | The notebook is on a SharePoint site — use `useSite` first, or pass `siteId=` |
| Tools work in `test-auth.mjs` but not Claude Desktop | `AZURE_CLIENT_ID` not set in `claude_desktop_config.json`, or Claude Desktop wasn't fully quit |
| Server crashes on `authenticate` with no error to client | Run `node onenote-mcp.mjs` directly in PowerShell to see the stderr trace |

## Security

* `.access-token.txt` and `.default-site.json` are both in `.gitignore`. Don't
  commit them.
* `Notes.ReadWrite.All` lets the signed-in user (and therefore this MCP) read
  and write any OneNote notebook the user can access. Treat the access token
  as you would your password.

## Acknowledgements

This fork builds on a chain of community work. Credit where due:

* **[eshlon/onenotemcp](https://github.com/eshlon/onenotemcp)** by Ehsan Shahabi
  — the direct upstream this fork is based on. Brought together the rich set
  of editing tools, JSDOM-based HTML processing, Zod schemas, and the overall
  structure.
* **[ZubeidHendricks/azure-onenote-mcp-server](https://github.com/ZubeidHendricks/azure-onenote-mcp-server)**
  — provided the device-code authentication pattern, token-cache strategy, and
  foundational Graph-API wrapping conventions.
* **[danosb/onenote-mcp](https://github.com/danosb/onenote-mcp)** — earlier
  reference for structuring a OneNote MCP server.

The SharePoint site support, session-default context, tenant-aware auth, and
crash hardening in this fork are new contributions on top of that lineage.

Development of this fork was assisted by Anthropic's Claude.

## License

MIT — see [LICENSE](LICENSE).
