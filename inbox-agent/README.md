# Inbox Agent

A Python agent that runs every morning at 7:30am Sydney time to:

1. Scan both Gmail (`edwardenag@gmail.com`) and Outlook (`edward@ibproperty.com.au`)
2. Generate a prioritised to-do list using Claude AI
3. Clean the inboxes (archive old read mail, sort promotions, trash spam)
4. Send a structured briefing to both addresses

---

## Quick start

### 1. Install dependencies

```bash
cd inbox-agent
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 2. Create your `.env` file

```bash
cp .env.example .env
```

Open `.env` and fill in your `ANTHROPIC_API_KEY`. The Azure credentials are already pre-filled.

### 3. Set up Gmail OAuth2

**Step 1 — Create a Google Cloud project (if you haven't already)**

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create a new project (or select an existing one)
3. Enable the **Gmail API**: APIs & Services → Enable APIs → search "Gmail API" → Enable

**Step 2 — Create OAuth2 credentials**

1. APIs & Services → Credentials → Create Credentials → **OAuth client ID**
2. Application type: **Desktop app**
3. Name: "Inbox Agent" (or anything)
4. Download the JSON file and save it as `inbox-agent/gmail_credentials.json`

**Step 3 — Add your Google account as a test user** (while the app is in "Testing" mode)

1. OAuth consent screen → Test users → Add `edwardenag@gmail.com`

**Step 4 — Run the auth flow** (opens a browser window once, then caches the token)

```bash
python agent.py --auth
```

The token is saved to `gmail_token.json` and refreshed automatically on future runs.

---

### 4. Authenticate with Outlook (MSAL device code flow)

On first run, the agent prints a URL and a short code:

```
==================================================================
  OUTLOOK AUTHENTICATION REQUIRED
  To sign in, use a web browser to open the page
  https://microsoft.com/devicelogin and enter the code XXXXXXXX
==================================================================
```

Open that URL in any browser, enter the code, and sign in with `edward@ibproperty.com.au`.

The token is cached in `msal_token_cache.json` and refreshed automatically. You will only need to do this once (or when the refresh token expires — usually after 90 days of inactivity).

To trigger auth without running the full agent:

```bash
python agent.py --auth
```

---

## Running the agent

```bash
# Full run (clean + to-do + briefing)
python agent.py

# Dry run — analyse but don't clean or send anything
python agent.py --dry-run

# Auth check only
python agent.py --auth
```

Logs are written to `logs/inbox_agent_YYYY-MM-DD.log`.

---

## Scheduling at 7:30am Sydney time

### Option A — cron

```bash
crontab -e
```

Add:

```cron
30 7 * * * cd /Users/edghattas/Documents/GitHub/Claude\ -\ Eddie\ G/inbox-agent && /Users/edghattas/Documents/GitHub/Claude\ -\ Eddie\ G/inbox-agent/.venv/bin/python agent.py >> logs/cron.log 2>&1
```

> **Note:** cron runs in UTC. Sydney is UTC+10 (AEST) or UTC+11 (AEDT). Adjust accordingly, or use `TZ=Australia/Sydney` in your crontab.

### Option B — launchd (macOS, recommended)

Create `~/Library/LaunchAgents/com.eddieg.inboxagent.plist`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>com.eddieg.inboxagent</string>

  <key>ProgramArguments</key>
  <array>
    <string>/Users/edghattas/Documents/GitHub/Claude - Eddie G/inbox-agent/.venv/bin/python</string>
    <string>/Users/edghattas/Documents/GitHub/Claude - Eddie G/inbox-agent/agent.py</string>
  </array>

  <key>WorkingDirectory</key>
  <string>/Users/edghattas/Documents/GitHub/Claude - Eddie G/inbox-agent</string>

  <key>StartCalendarInterval</key>
  <dict>
    <key>Hour</key>
    <integer>7</integer>
    <key>Minute</key>
    <integer>30</integer>
  </dict>

  <key>StandardOutPath</key>
  <string>/Users/edghattas/Documents/GitHub/Claude - Eddie G/inbox-agent/logs/launchd.log</string>
  <key>StandardErrorPath</key>
  <string>/Users/edghattas/Documents/GitHub/Claude - Eddie G/inbox-agent/logs/launchd.log</string>

  <key>EnvironmentVariables</key>
  <dict>
    <key>PATH</key>
    <string>/usr/local/bin:/usr/bin:/bin</string>
  </dict>
</dict>
</plist>
```

Load it:

```bash
launchctl load ~/Library/LaunchAgents/com.eddieg.inboxagent.plist
```

launchd respects the local system clock (Sydney time), so no UTC conversion needed.

---

## Environment variables reference

| Variable | Description | Default |
|---|---|---|
| `ANTHROPIC_API_KEY` | Claude API key | *(required)* |
| `AZURE_CLIENT_ID` | Azure app client ID | Pre-filled |
| `AZURE_TENANT_ID` | Azure tenant ID | Pre-filled |
| `USER_GMAIL` | Gmail address | `edwardenag@gmail.com` |
| `USER_OUTLOOK` | Outlook address | `edward@ibproperty.com.au` |
| `NOTIFY_EMAILS` | Comma-separated briefing recipients | Both addresses |
| `GMAIL_CREDENTIALS_PATH` | Path to `credentials.json` from Google | `gmail_credentials.json` |
| `GMAIL_TOKEN_PATH` | Where to cache the Gmail token | `gmail_token.json` |
| `MSAL_TOKEN_CACHE_PATH` | Where to cache the MSAL token | `msal_token_cache.json` |

---

## Cleaning rules

| Rule | Action |
|---|---|
| Email is read + older than 7 days + not deal-related | Archive (removed from inbox, kept in All Mail / Archive) |
| Sender matches bulk/newsletter patterns OR body has unsubscribe link | Move to Promotions folder |
| Matches known spam patterns (prize scams, crypto offers, etc.) | Trash / Deleted Items (recoverable) |
| Contains real estate deal keywords OR from a known business domain | **Never touched** |

Deal keywords checked: lease, tenant, landlord, inspection, rent, contract, offer, listing, vendor, buyer, settlement, auction, appraisal, strata, council, zoning, development, commercial, retail, office, warehouse, property, ibproperty, PM, property management, due diligence, exchange, deposit, valuation.

---

## File structure

```
inbox-agent/
├── agent.py              # Main orchestrator — start here
├── gmail_client.py       # Gmail API wrapper
├── outlook_client.py     # Microsoft Graph / MSAL wrapper
├── cleaner.py            # Inbox cleaning logic & classification
├── todo_generator.py     # Claude AI to-do list generation
├── briefing.py           # HTML briefing builder & sender
├── requirements.txt
├── .env.example          # Copy to .env and fill in ANTHROPIC_API_KEY
├── README.md
└── logs/                 # Daily log files (auto-created)
```

---

## Troubleshooting

**"Gmail credentials not found"** — Download `credentials.json` from Google Cloud Console and put it in the `inbox-agent/` folder (or set `GMAIL_CREDENTIALS_PATH`).

**"MSAL device flow failed"** — Check that your Azure app has the correct permissions (`Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `User.Read`) and that admin consent has been granted.

**"Claude API error"** — Verify `ANTHROPIC_API_KEY` in your `.env`. The agent still sends the briefing (without the AI to-do list) even if Claude is unavailable.

**Briefing not arriving** — Check `logs/inbox_agent_YYYY-MM-DD.log` for send errors. If Gmail send fails the agent automatically retries via Outlook.
