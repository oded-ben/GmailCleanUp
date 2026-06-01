# GmailCleanup

Automated Gmail triage for unread inbox mail. Uses rule-based filtering first, then AI for everything else.

**Delete** does not permanently remove mail — it adds a `CleanupQueue` label and archives the thread so you can review or bulk-delete later in Gmail.

## How it works

| Pass | What it does |
|------|----------------|
| **1 — Rules** | Matches known senders, domains, and subject keywords → queue or keep |
| **2 — AI** | Classifies remaining unread threads (Gemini by default) |

### Apps Script (recommended for scheduling)

- Pass 2 assigns **your existing Gmail labels**, or `Keep` / `Delete`
- Logs every decision to a **Google Sheet** dashboard

### Node.js (local CLI)

- Pass 2 returns **keep** or **delete** only
- Logs to `decisions.log` in the project folder

---

## Option A — Google Apps Script

Best if you want to run on a schedule from Google's servers with Sheet logging.

### Setup

1. Copy [`apps-script/Code.gs`](apps-script/Code.gs) into a project at [script.google.com](https://script.google.com) (or bind it to a Google Sheet via **Extensions → Apps Script**).

2. **Project settings → Script properties** — add:

   | Property | Required | Description |
   |----------|----------|-------------|
   | `GEMINI_API_KEY` | Yes | From [Google AI Studio](https://aistudio.google.com/apikey) |
   | `GEMINI_MODEL` | No | Default: `gemini-2.5-flash` |
   | `DASHBOARD_SPREADSHEET_ID` | No | Auto-created on first log if unset |

3. Authorize Gmail + external URL access when prompted.

### Run

| Function | Purpose |
|----------|---------|
| `runFullCleanup` | Rules → AI → dashboard log |
| `classifyInbox` | Rule pass only |
| `classifyUnclassified` | AI pass only (uses your Gmail labels) |
| `setupLabel` | Create `CleanupQueue` label |
| `listExistingLabels` | Show labels the AI can assign |
| `testGeminiKey` | Verify Gemini API key |
| `testLogToDashboard` | Write sample rows to the sheet (no Gmail changes) |
| `runSetupDashboardOnly` | Create/bind dashboard spreadsheet |

### Dashboard

- Tab: **Classification Log**
- Columns: `Timestamp`, `Source`, `Sender`, `Subject`, `Action`, `Reason`
- On first run, creates **Gmail Cleanup Dashboard** in Drive and saves its ID to Script Properties

---

## Option B — Node.js (local)

Run from your machine with OAuth to Gmail.

### Prerequisites

- Node.js 20+
- Google Cloud OAuth **Desktop** client → download as `credentials.json`
- Redirect URI: `http://localhost:3000`

### Setup

```bash
npm install
```

Create `.env` in the project root:

```env
GEMINI_API_KEY=your_key_here
AI_PROVIDER=gemini
GEMINI_MODEL=gemini-2.5-flash
```

Optional — use Anthropic instead:

```env
AI_PROVIDER=anthropic
ANTHROPIC_API_KEY=your_key_here
```

On first run, a browser opens for Google sign-in. Token is saved to `token.json`.

If you see `invalid_grant`, delete `token.json` and run again.

### Run

```bash
node main.js
```

List remaining unread threads (read-only):

```bash
node audit.js
```

---

## Project layout

```
GmailCleanup/
├── apps-script/Code.gs   # Google Apps Script version (multi-label + Sheet log)
├── main.js               # Local CLI entry point
├── classify.js           # Rules + AI batch classification
├── auth.js               # Gmail OAuth (local)
├── audit.js              # Read-only inbox listing
├── package.json
└── .gitignore            # Excludes secrets and node_modules
```

---

## Security

**Never commit these files** (already in `.gitignore`):

- `.env`
- `credentials.json`
- `token.json`
- `decisions.log`

---

## License

ISC
