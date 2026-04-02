# Outlook EML Exporter (Local Archive)

Export Outlook Web mail to local `.eml` files with the original MIME payload intact, including attachments. Runs entirely in the browser. No backend. No PST. No desktop Outlook.

Works in two modes:
- as a userscript installed from [GreasyFork](https://greasyfork.org/en/scripts/572163-outlook-eml-exporter-local-archive) or any standard userscript manager (persists via the manager and is injected on matching pages each load)
- as a one-off script pasted directly into the DevTools console on Outlook Web (temporary — will be removed on reload)

> Requires a Chromium-based browser. The script depends on the File System Access API (`showDirectoryPicker()`).

## Table of Contents

- [Requirements](#requirements)
- [Supported pages](#supported-pages)
- [Installation](#installation)
  - [Userscript](#userscript)
  - [DevTools console](#devtools-console)
- [Usage](#usage)
- [Getting a token](#getting-a-token)
- [Backend selection](#backend-selection)
- [Output and resume behavior](#output-and-resume-behavior)
- [Configuration](#configuration)
- [Limitations](#limitations)
- [Security](#security)
- [Troubleshooting](#troubleshooting)
- [License](#license)

## Requirements

| Requirement | Why it matters |
| --- | --- |
| Chromium-based browser (File System Access API required; Firefox not supported) | Needed for `showDirectoryPicker()` and local folder writes |
| Outlook Web session | The script exports from the mailbox visible in the current web session |
| Valid bearer token | Used to call Microsoft Graph or Outlook REST mail endpoints |
| Userscript manager or DevTools console | Execution method |

## Supported pages

| Host | Notes |
| --- | --- |
| `https://outlook.office.com/*` | Microsoft 365 Outlook Web |
| `https://outlook.office365.com/*` | Microsoft 365 Outlook Web |
| `https://outlook.live.com/*` | Outlook.com / consumer mail |

## Installation

### Userscript

Install the script from GreasyFork ([see here](https://greasyfork.org/en/scripts/572163-outlook-eml-exporter-local-archive)), or load the `.user.js` file into any userscript manager. The script uses `@grant none`, so it does not depend on manager-specific APIs. The userscript manager will inject the script on matching pages each time the page loads; disable/uninstall it in the manager to remove persistence.

### DevTools console

1. Open a supported Outlook Web page.
2. Open DevTools.
3. Paste the full script into the **Console**.
4. Press Enter.

This mode is temporary. Reloading the page removes it.

## Usage

The panel starts collapsed.

1. Click the panel header to expand it.
2. Click **Pick export folder** and choose a local directory.
3. Obtain a mailbox token from DevTools Network.
4. Paste the token, `Authorization` header, `Copy as fetch`, or `Copy as cURL` into the input box.
5. Click **Test token**.
6. Click **Start export**.
7. Keep the tab open until the run completes.
8. Re-run later to resume. Existing files are skipped.

UI notes: the header is the drag handle; the panel cannot be dragged off the viewport (movement is clamped). Clicking the header toggles collapse unless you just dragged the header.

### Controls

| Control | Function |
| --- | --- |
| Header | Click to collapse/expand; drag to move (viewport-clamped) |
| Pick export folder | Opens the File System Access directory picker |
| Token box | Accepts raw bearer token, `Authorization` header, `Copy as fetch`, or `Copy as cURL` |
| Test token | Extracts token, decodes JWT if possible, resolves backend, probes identity and mail access |
| Start export | Recursively exports folders and messages |
| Stop | Cooperative stop; finishes the current request first |
| Log | Selectable run log |

## Getting a token

Use a token issued to Outlook Web.

1. Open Outlook Web.
2. Press `F12` and go to **Network**.
3. Trigger mailbox traffic by opening a folder.
4. Filter for `mailfolders` or `messages`.
5. Pick a successful `200` request to either:
   - `graph.microsoft.com`
   - `outlook.office.com`
6. Right-click the request and copy one of:
   - **Copy as fetch**
   - **Copy as cURL**
7. Paste the result into the script and click **Test token**.

### Accepted input formats

- raw bearer token
- `Authorization: Bearer ...`
- DevTools `Copy as fetch`
- DevTools `Copy as cURL`

## Backend selection

The script chooses the API from the JWT audience (`aud`) claim.

| JWT `aud` | Backend used |
| --- | --- |
| `graph.microsoft.com` or Graph resource GUID `00000003-0000-0000-c000-000000000000` | Microsoft Graph `https://graph.microsoft.com/v1.0` |
| `https://outlook.office.com` | Outlook REST `https://outlook.office.com/api/v2.0` (probes `/api/beta` if needed) |

If the audience does not match a supported mail API the script will fail and show hints — copy a token from a request to the same host as the API you intend to call.

## Output and resume behavior

- Exports one `.eml` file per message using the message MIME endpoint (`/$value`), so attachments are embedded.
- Mirrors the Outlook folder tree under the selected local directory.
- Traverses top-level folders recursively (configurable).
- Sanitizes and truncates filenames/dirnames for filesystem safety; avoids Windows reserved names.
- If the OS rejects a directory name, the script falls back to a hashed folder name.
- Handles `429` throttling with `Retry-After` backoff.

### Filename shape

Files are written approximately as:

```
<timestamp>__<sender>__<subject>__m-<hash>.eml
```

The `m-<hash>` suffix is stable per message ID and used for resume detection.

### Resume semantics

Re-running the export against the same target folder will skip messages already present in that folder if the script finds:
- the exact exported filename, or
- the stable message key suffix, or
- the legacy short ID suffix

Resume is folder-local. Use the same export root for reliable skipping.

## Configuration

Edit these constants near the top of the script before installing or pasting it:

| Constant | Default | Effect |
| --- | ---: | --- |
| `ONLY_TOP_LEVEL_FOLDERS` | `null` | Export all top-level folders recursively. Set an array like `['Inbox','Sent Items']` to restrict scope. |
| `REQUEST_DELAY_MS` | `120` | Delay between message downloads (ms) |
| `MAX_FILENAME_BYTES` | `180` | Maximum filename byte length |
| `MAX_DIRNAME_BYTES` | `80` | Maximum directory name byte length |

Example:

```js
const ONLY_TOP_LEVEL_FOLDERS = ['Inbox', 'Sent Items'];
```

## Limitations

- Chromium-only in practice (File System Access API required). Firefox support is not guaranteed.
- No token refresh flow. When the token expires, copy a fresh one and re-run.
- Export is serial; large mailboxes take time.
- Targets `/me/...` endpoints; not designed for arbitrary mailbox ID migrations.
- Active runs stop if the page is closed or reloaded.

## Security

- The token you paste is a bearer credential. Treat it as a secret.
- The script runs locally in the browser and writes only to the directory you choose.
- Network calls go to Microsoft mail APIs, not to a third-party service.
- Review the source before installing from any distribution channel.
- Clear the token box when you are done.
- Exported `.eml` files may contain sensitive headers, body content, and attachments.

## Troubleshooting

<details>
<summary>Common failures</summary>

| Symptom | Likely cause | Fix |
| --- | --- | --- |
| `showDirectoryPicker not available` | Unsupported browser | Use Chrome, Edge, or another Chromium browser |
| `Could not extract a bearer token` | Incomplete paste | Paste the full raw token, `Authorization` header, `Copy as fetch`, or `Copy as cURL` |
| Token looks truncated | Copied partial header or trimmed fetch/cURL | Paste the full `Copy as fetch` / `Copy as cURL` blob (token can be very long) |
| `Unsupported token audience` | Wrong token audience | Copy a token from a request to the same host as the API being called (`graph.microsoft.com` or `outlook.office.com`) |
| `Invalid audience` / `401` / `403` | Token expired, wrong audience, or missing scopes | Refresh Outlook, copy a fresh mailbox request, then click **Test token** |
| Export is slow | Sequential MIME download and server throttling | Narrow scope with `ONLY_TOP_LEVEL_FOLDERS` or increase patience |
| `Rate-limited. Waiting ...` | Microsoft API throttling | Script will back off and continue; be patient |
| Stop does not interrupt immediately | Stop is cooperative | It takes effect after the current request finishes |
| Start button stays disabled | Missing folder or token | Pick a folder and paste a token first |
| Panel moved off-screen | UI repositioned or page resized | Remove the panel via console (document.getElementById('oeml-exporter-panel')?.remove()) and reload the page |

</details>

## License

MIT
