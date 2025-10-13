Here’s a complete **README.md** you can drop into the repo.

---

# Webex PPTX → JSON Bot (Websocket + Python)

A self‑hosted Webex messaging bot that runs on Ubuntu, listens for **1:1** messages over **websockets**, accepts an uploaded **.pptx** file, extracts slide **text** with a Python helper, and replies in-chat with the extracted **JSON** (chunked automatically to fit message size limits).

* **No public webhook required** — uses the Webex JS SDK websocket listener (`webex.messages.listen()` + `messages.on('created')`). ([Webex for Developers][1])
* **Direct spaces**: the bot can read all messages; in **group spaces** you must `@mention` it. ([Webex for Developers][2])
* **Files** are downloaded via the Webex contents URL using the bot token, with built‑in retry for anti‑malware scanning (HTTP `423` + `Retry-After`) and handling for unscannable/blocked cases. ([Webex for Developers][3])
* Replies are sent as Markdown code blocks, respecting the **7439‑byte** per‑message limit (auto-chunked with margin). ([Webex for Developers][4])

---

## Contents

* [Architecture](#architecture)
* [Prerequisites](#prerequisites)
* [Quick Start (Ubuntu)](#quick-start-ubuntu)
* [Configuration](#configuration)
* [Usage](#usage)
* [Sample Output](#sample-output)
* [How It Works](#how-it-works)
* [Troubleshooting](#troubleshooting)
* [Security Notes](#security-notes)
* [Run as a systemd Service (optional)](#run-as-a-systemd-service-optional)
* [Roadmap / Enhancements](#roadmap--enhancements)
* [License](#license)

---

## Architecture

```
Webex (user DM)
     │
     │ 1) User uploads .pptx to 1:1 space
     ▼
Webex Cloud ─(websocket event: messages/created)─► Node (server.js)
                                                  │
                                                  │ 2) Get message details; find `files[]`
                                                  │ 3) HEAD/GET contents URL with bot token
                                                  │ 4) Save to /tmp
                                                  │ 5) exec: python3 pptx_to_json.py <file>
                                                  │ 6) Read JSON from stdout
                                                  │ 7) Send Markdown JSON back (chunked)
                                                  ▼
                                                Webex (bot reply)
```

**Directories**

```
webex-pptx-json-bot/
├── .env.example
├── .gitignore
├── package.json
├── requirements.txt
├── server.js
└── pptx_to_json.py
```

---

## Prerequisites

* **Ubuntu** (22.04+ recommended)
* **Node.js** 18+ and **npm**
* **Python 3.8+**, `pip`, optional `venv`
* A **Webex Bot** and **bot access token** (create in Developer Portal)

  * Websocket listening is supported by the Webex JS SDK. ([Webex for Developers][1])
  * In direct 1:1 spaces, the bot receives all user messages; in group spaces you must `@mention` the bot. ([Webex for Developers][2])

---

## Quick Start (Ubuntu)

```bash
# 0) System packages
sudo apt-get update
sudo apt-get install -y python3 python3-venv python3-pip git

# 1) Clone your project
git clone <your repo url> webex-pptx-json-bot
cd webex-pptx-json-bot

# 2) Node deps
npm install

# 3) Python deps (venv optional but recommended)
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# 4) Configure environment
cp .env.example .env
# edit .env and paste WEBEX_ACCESS_TOKEN=<your bot token>

# 5) Run
npm start
```

**Test:** DM your bot in Webex and upload a `.pptx`. It replies with JSON (split across multiple messages if long).

---

## Configuration

Copy `.env.example` → `.env` and set:

```bash
WEBEX_ACCESS_TOKEN=xxxxx_your_bot_token_xxxxx
# Optional overrides
# PYTHON_BIN=/usr/bin/python3
# PPTX_EXTRACTOR_SCRIPT=./pptx_to_json.py
```

Other behavior (already implemented in code):

* Saves incoming files to a temp folder under your system `/tmp`.
* Auto‑retries downloads while anti‑malware scanning is pending (`423 Locked` + `Retry-After`) and handles:

  * `410 Gone` → infected/blocked
  * `428 Precondition Required` → appends `?allow=unscannable` and retries (user assumes risk) ([Webex for Developers][3])
* Splits outgoing Markdown into ~7 KB chunks to stay below the **7439‑byte** API limit. ([Webex for Developers][4])

---

## Usage

1. In Webex, start a **direct (1:1)** chat with your bot.
2. Upload a **.pptx** file.
3. The bot replies with a Markdown‑formatted JSON block representing slide text (plus notes, when present).

> **Note:** In **group spaces** you must `@mention` the bot for it to receive messages. For 1:1 spaces, mention is not required. ([Webex for Developers][2])

---

## Sample Output

```json
{
  "file_name": "Quarterly_Update.pptx",
  "slide_count": 3,
  "slides": [
    {
      "index": 1,
      "title": "Q3 Highlights",
      "text": [
        "Revenue up 12% YoY",
        "Launched APAC pilot in August",
        "Improved NPS by 4.2 points"
      ],
      "notes": ["Call out pilot success stories."]
    },
    {
      "index": 2,
      "title": "Pipeline",
      "text": [
        "Enterprise: $3.1M (12 opps)",
        "Mid-market: $1.4M (27 opps)"
      ]
    },
    {
      "index": 3,
      "title": "Next Steps",
      "text": [
        "Finalize FY roadmap",
        "Staff 2 FTEs for APAC",
        "Kickoff Q4 incentive program"
      ]
    }
  ]
}
```

---

## How It Works

* **Websocket listener**: `server.js` authenticates with your bot token and starts a websocket listener:
  `await webex.messages.listen(); webex.messages.on('created', handler)` — no public webhook needed. ([Webex for Developers][1])
* **Message handling**:

  1. On `messages:created`, fetch **message details** to access `files[]`.
  2. For each file URL, issue `HEAD` to infer filename/content‑type and select the first **.pptx**.
  3. Download the file with `GET` using the bot token (`Authorization: Bearer <token>`).

     * If **anti‑malware scanning** is in progress, the API returns `423 Locked` with `Retry-After`; the bot waits and retries.
     * If infected, the API returns `410 Gone`.
     * If unscannable (e.g., encrypted), the API returns `428 Precondition Required`; the bot retries with `?allow=unscannable`. ([Webex for Developers][3])
* **Python extraction**: Saves the PPTX to `/tmp` and runs:

  ```
  python3 pptx_to_json.py /tmp/<file>.pptx
  ```

  The script uses `python-pptx` to iterate slides and shapes (`has_text_frame`) and prints a compact JSON to **stdout**.
* **Reply**: JSON is wrapped in a Markdown code fence and **chunked** to fit the **7439‑byte** message limit; multiple messages are sent as needed. ([Webex for Developers][4])

---

## Troubleshooting

* **Bot not responding**

  * Verify `WEBEX_ACCESS_TOKEN` is valid and the process logs show `Authenticated as ...`.
  * Ensure you are messaging the bot in a **direct** space. In group spaces, **`@mention`** the bot. ([Webex for Developers][2])
* **No file detected**

  * The message didn’t include attachments, or the attachment isn’t `.pptx`.
* **Stuck on download**

  * Your org may have **anti‑malware scanning** enabled; the contents URL will return `423 Locked` until the scan completes. The bot auto‑retries using the `Retry-After` header. ([Webex for Developers][3])
* **Download blocked**

  * `410 Gone` → file flagged as infected; Webex prevents download.
  * `428 Precondition Required` → file unscannable; the bot retries with `?allow=unscannable` and downloads at your own risk. ([Webex for Developers][3])
* **Reply truncated or missing lines**

  * Large JSON is **split into multiple messages**; scroll up in the 1:1 to see all parts. The bot caps each chunk well under the **7439‑byte** limit. ([Webex for Developers][4])

---

## Security Notes

* The bot uses the **bot token** only on your server (no third‑party services).
* Temporary files are written under the OS temp directory and unlinked after processing.
* If you do not want to allow unscannable files, set a policy or remove the `allow=unscannable` logic. (Unscannable downloads shift risk to you.) ([Webex for Developers][3])

---

## Run as a systemd Service (optional)

Create `/etc/systemd/system/webex-pptx-json-bot.service`:

```ini
[Unit]
Description=Webex PPTX → JSON Bot
After=network.target

[Service]
WorkingDirectory=/opt/webex-pptx-json-bot
Environment=NODE_ENV=production
EnvironmentFile=/opt/webex-pptx-json-bot/.env
ExecStart=/usr/bin/node /opt/webex-pptx-json-bot/server.js
Restart=always
RestartSec=5
# If using a Python venv:
Environment=PATH=/opt/webex-pptx-json-bot/venv/bin:/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin
User=ubuntu
Group=ubuntu
StandardOutput=journal
StandardError=journal

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now webex-pptx-json-bot
sudo systemctl status webex-pptx-json-bot
journalctl -u webex-pptx-json-bot -f
```

---

## Roadmap / Enhancements

* Optional **Adaptive Card** summary with “Download JSON” action (keep raw JSON reply as default). ([Webex for Developers][5])
* Replace `execFile` with a **FastAPI** microservice for long‑running Python.
* Add **image/shape** extraction or layout metadata (beyond text).
* Persist results to disk/object storage (currently stateless).

---

### References

* **Websockets with Webex JS SDK** — listen for `messages:created` events without public webhooks. ([Webex for Developers][1])
* **Bot access in spaces** — 1:1 vs. group `@mention` requirement. ([Webex for Developers][2])
* **Anti‑malware scanning & contents download status codes** (`423` + `Retry‑After`, `410`, `428` + `allow=unscannable`). ([Webex for Developers][3])
* **Create Message API** — Markdown and **7439‑byte** message limit. ([Webex for Developers][4])

---

[1]: https://developer.webex.com/blog/using-websockets-with-the-webex-javascript-sdk "Using Websockets with the Webex JavaScript SDK | Webex Developers Blog"
[2]: https://developer.webex.com/docs/bots?utm_source=chatgpt.com "Webex Messaging"
[3]: https://developer.webex.com/docs/basics?utm_source=chatgpt.com "REST API Basics"
[4]: https://developer.webex.com/docs/api/v1/messages/create-a-message?utm_source=chatgpt.com "Create a Message"
[5]: https://developer.webex.com/docs/buttons-and-cards?utm_source=chatgpt.com "Buttons and Cards"
