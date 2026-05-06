# Outlook (M365) Email Exporter

A Tampermonkey userscript that exports every email in mainstream (Inbox, Sent, etc) Outlook folders as `.eml` files, bundled into ZIPs. Works on Outlook Web (for tenants/orgs).

[![Install with Tampermonkey](https://img.shields.io/badge/Install%20with-Tampermonkey-00485B?style=for-the-badge&logo=tampermonkey&logoColor=white)](https://raw.githubusercontent.com/nicholasyoannou/outlook-web-email-exporter/main/outlook_web_email_exporter.user.js)

Created and tested via Microsoft 365 Outlook (outlook.cloud.microsoft) to export emails in .EML format for the purpose of migration to Google Workspace due to personal recurring authentication problems with M365; this script is just a general exporter, though, and is general purpose. This script
requires you to be logged in on Outlook for Web, and uses the built-in export feature (via Right click individual mail > Download as > .EML), scaled to the entire inbox. This _hasn't_ been tested with ordinary Outlook (outlook.live.com), only MS365 Outlook, so no promises in that regard.

After successful export, importing the .EML files to your new inbox can be done via Thunderbird's ImportExportTools NG extension, creating and importing the .EML files to a local folder, and copying it over to your mail provider. For validation reasons, you may want to filter the exported files to ensure they fit under the file limit of your new email provider. Alternatively, if migrating to Google Workspace, you can use [GYB (Got Your Back)](https://github.com/GAM-team/got-your-back) to validate and migrate emails across. GYB will promptly outline if there are problems with specific emails, and will skip emails you've already migrated if re-running on a bulk of emails which can be handy.

---

## Features

- **Exports `.eml` files** using the Outlook REST API (`/api/v2.0/me/messages/{id}/$value`) — full MIME with headers and attachments intact
- **Supports modern folders** — Inbox, Sent Items, Archive, Drafts, Deleted Items, Junk
- **Paginated listing** — handles mailboxes of any size, fetches up to 200 message IDs per page
- **Concurrent downloads** — configurable concurrency (1–8 workers, default 4) for fast exports
- **Automatic throttle handling** — detects `429 Too Many Requests` and backs off with exponential cooldown shared across all workers
- **Retry on failure** — retries on `429` and `5xx` errors up to 6 times before skipping
- **Atomic ZIP flushing** — concurrent workers can't trigger duplicate ZIP downloads; each ZIP is flushed safely
- **Configurable batch size** — control how many emails go into each ZIP (50–5000, default 500)
- **Smart filenames** — each `.eml` is named `YYYY-MM-DD__sender@domain__Subject__ID.eml` using decoded MIME headers (RFC 2047 support)
- **Session-safe auth capture** — intercepts `fetch` and `XHR` to silently capture the bearer token and anchor mailbox from Outlook's own traffic — no API keys or OAuth setup required
- **Pause / Resume / Stop** — full playback controls in the UI panel

---

## Requirements

- [Tampermonkey](https://www.tampermonkey.net/) browser extension (Chrome, Firefox, Edge, Safari)
- An active Outlook Web session (Microsoft 365 / Office 365)

---

## Installation

### One-click install

Click the button at the top of this page, or click below:

[![Install with Tampermonkey](https://img.shields.io/badge/Install%20with-Tampermonkey-00485B?style=for-the-badge&logo=tampermonkey&logoColor=white)](https://raw.githubusercontent.com/nicholasyoannou/outlook-web-email-exporter/main/outlook_web_email_exporter.user.js)

Tampermonkey will open an install confirmation dialog. Click **Install**.

### Manual install

1. Open Tampermonkey → **Dashboard** → **+** (Create new script)
2. Paste the contents of [`outlook_web_email_exporter.user.js`](./outlook_web_email_exporter.user.js)
3. Press **Ctrl+S** to save

---

## Usage

1. Open [Outlook Web](https://outlook.office.com) and sign in
2. Click any email to trigger Outlook's API traffic — the script will capture your session token automatically
3. The **📥 Outlook Email Exporter** panel will appear in the top-right corner
4. Select a **folder** (Inbox, Sent Items, etc.)
5. Optionally adjust **Emails per ZIP** and **Concurrency**
6. Wait for the status to show **✓ Ready**, then click **▶ Start**
7. ZIP files will download automatically as each batch completes

> **Tip:** If the status stays on "missing: bearer, mailbox", click any email in Outlook to trigger a fresh API request and let the script capture the token.

---

## Settings

| Setting | Default | Range | Notes |
|---|---|---|---|
| Emails per ZIP | 500 | 50–5000 | Number of `.eml` files bundled per ZIP download |
| Concurrency | 4 | 1–8 | Parallel download workers. Values above 4 may trigger throttling |

---

## Supported Outlook Hosts

- `https://outlook.cloud.microsoft/*`
- `https://outlook.office.com/*`
- `https://outlook.office365.com/*`

---

## How It Works

The script intercepts `fetch` and `XMLHttpRequest` calls made by Outlook Web and silently captures the **Bearer token** and **X-AnchorMailbox** header from existing API traffic. It then uses the Outlook REST API directly:

- **Listing:** `GET /api/v2.0/me/mailFolders/{folder}/messages?$top=200&$count=true`
- **Download:** `GET /api/v2.0/me/messages/{id}/$value`

No credentials are stored or transmitted anywhere other than back to Microsoft's own servers. The token expires with your Outlook session.

---

## Troubleshooting

**Status stuck on "missing: bearer, mailbox"**
→ Click any email in Outlook to trigger an API request. The script captures the token passively.

**`401` error mid-export**
→ Your session expired. Refresh the Outlook tab and restart the export.

**Large emails fail to upload to Gmail**
→ Gmail rejects messages over 25 MB via IMAP. Move oversized EMLs out before importing.

---

## License & disclaimer

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

This software is an independent, unofficial tool and is not affiliated with,
endorsed by, or in any way associated with Microsoft Corporation, Microsoft
365, Outlook, or any related Microsoft products or services.

Use of this software is entirely at your own risk. The author(s) make no
representations or warranties of any kind, express or implied, regarding the
software's fitness for a particular purpose, accuracy, reliability, or
continued availability.
