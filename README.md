# Lapse Pre-Lapse Automation

Automates the daily NASU lapse workflow: download PDF workflow files from **Tango**, package them into zip archives, upload to **SharePoint**, and optionally send **Slack** alerts when files are missing.

Built with Python and [Playwright](https://playwright.dev/python/).

## What it does

### 1. Tango download

Logs into [Tango PDF Workflow Approval](https://tango.se2.com/web/tango/pdf-workflow-approval), searches for files using the **previous US business day** date, and downloads all output PDFs for each file type.

| Category | File prefixes searched |
|----------|------------------------|
| Billing | `P{date}NASULFGRCLETTER`, `P{date}NASULFFLPSLETTER` |
| Letters | `P{date}NASUANCMDLETTER`, `P{date}ANNUITYMLTLRLETTER`, `P{date}NASUSCHLTRLETTER` |

Output zips:

- `output/P{date}NASU_BILLING.zip`
- `output/P{date}NASU_LETTERS.zip`

### 2. SharePoint upload

Uploads zip files to the NassauRe SharePoint library, sets document metadata, and saves a screenshot of the library.

After upload, files start in **Unassigned**. Once **Item Type = Test Results** is set, they move to the **Test Results** section.

### 3. Slack alerts (optional)

Sends a notification when expected file types were not found during a Tango run.

## Business day date logic

The `{date}` in file names is the **previous US business day** (Monday–Friday), formatted as `MMDDYY`.

| Run day | Date used | Example prefix |
|---------|-----------|------------------|
| Tuesday | Monday | `P052526NASULFGRCLETTER` |
| Monday | Friday | `P052226NASULFGRCLETTER` |
| Saturday / Sunday | Friday | `P052226NASULFGRCLETTER` |

Weekends are skipped; the script walks back to the last weekday.

By default, Tango polling waits until **4:30 PM IST** on the day after the business date, then checks every 30 minutes until **11:59 PM IST**.

## Prerequisites

- Windows 10/11
- Python 3.10+
- Microsoft Edge (used for SharePoint with a saved login profile)

## Setup

```powershell
cd "C:\Demo\Lapse Pre-Lapse Automation"

python -m venv .venv
.\.venv\Scripts\Activate.ps1

pip install -r requirements.txt
playwright install chromium
```

Copy the environment template and add your Tango credentials:

```powershell
copy .env.example .env
```

Edit `.env`:

```env
TANGO_USERNAME=your_username
TANGO_PASSWORD=your_password
```

**Do not commit `.env`** — it is listed in `.gitignore`.

### SharePoint one-time login

SharePoint uses a saved Edge browser profile (no password in `.env` required):

```powershell
py sharepoint_upload.py --save-login
```

Sign in with your work account in the Edge window that opens. The session is stored in `edge_sharepoint_profile/` and reused on future runs.

## Usage

### Download billing files only (quick test)

```powershell
py tango_billing_download.py --billing-only --single-pass --no-wait --no-sharepoint --no-slack
```

### Full Tango download (billing + letters, with polling)

```powershell
py tango_billing_download.py
```

### Download with visible browser (debugging)

```powershell
py tango_billing_download.py --headed --billing-only --single-pass --no-wait
```

### Upload existing zips to SharePoint

```powershell
py sharepoint_upload.py --headed output\P052226NASU_BILLING.zip output\P052226NASU_LETTERS.zip
```

### Tango download + SharePoint upload in one run

```powershell
py tango_billing_download.py --no-slack
```

### Upload only (skip Tango download)

```powershell
py tango_billing_download.py --sharepoint-only --headed
```

## Command reference

### `tango_billing_download.py`

| Flag | Description |
|------|-------------|
| `--headed` | Show the browser window |
| `--billing-only` | Download billing files only; create billing zip |
| `--single-pass` | Check once; no polling |
| `--no-wait` | Do not wait until 4:30 PM IST |
| `--run-date YYYY-MM-DD` | Override today for business-day calculation |
| `--poll-interval N` | Minutes between checks (default: 30) |
| `--no-sharepoint` | Skip SharePoint upload |
| `--no-slack` | Skip Slack notifications |
| `--sharepoint-only` | Upload existing zips from `output/` |
| `--output-dir PATH` | Output directory (default: `output`) |

### `sharepoint_upload.py`

| Flag | Description |
|------|-------------|
| `--headed` | Show the browser window |
| `--save-login` | Save SharePoint / Edge login session |
| `--screenshot-only` | Capture library screenshot without uploading |
| `--no-screenshot` | Skip screenshot after upload |
| `--properties-only` | Set metadata on files already in the library |
| `--inspect-properties` | Debug: dump properties panel HTML |

### `slack_notify.py`

```powershell
py slack_notify.py --test
```

## Optional configuration (`.env`)

| Variable | Purpose |
|----------|---------|
| `TANGO_USERNAME` / `TANGO_PASSWORD` | Tango login (required for download) |
| `SHAREPOINT_URL` | Override default SharePoint library URL |
| `SLACK_WEBHOOK_URL` | Slack incoming webhook |
| `SLACK_BOT_TOKEN` / `SLACK_CHANNEL_ID` | Slack bot alternative to webhook |
| `MS365_EMAIL` / `MS365_PASSWORD` | Optional; prefer `--save-login` for SharePoint |

## Output layout

```
output/
├── downloads/          # Individual PDFs from Tango
├── P052226NASU_BILLING.zip
├── P052226NASU_LETTERS.zip
└── screenshots/        # SharePoint upload screenshots
```

## Project structure

```
├── tango_billing_download.py   # Entry point: Tango download + zip
├── sharepoint_upload.py        # Entry point: SharePoint upload
├── slack_notify.py             # Entry point: Slack test / alerts
├── lapse/
│   ├── tango.py                # Tango automation logic
│   ├── sharepoint.py           # SharePoint upload logic
│   └── slack.py                # Slack notification logic
├── .env.example
├── requirements.txt
└── edge_sharepoint_profile/    # Saved Edge session (local, not committed)
```

## Troubleshooting

**Cancel a running script**

Press `Ctrl+C` in PowerShell, or:

```powershell
Get-Process python -ErrorAction SilentlyContinue | Stop-Process -Force
```

**"Please sign in" dialog during SharePoint upload**

The script automatically clicks **Not Now** when that prompt appears.

**Upload button blocked / dialog backdrop**

Usually caused by a leftover search filter or modal. Re-run with `--headed` to watch the browser; the script clears search and dismisses overlays before uploading.

**File not in Unassigned after metadata is set**

That is expected. After **Item Type = Test Results** is applied, the file appears under **Test Results**, not Unassigned. The script searches both sections.

**Missing Tango files**

Files typically appear after 4:30 PM IST the day after the business date. Use default polling, or run again later with `--single-pass --no-wait`.

## Security notes

- Never commit `.env` or `edge_sharepoint_profile/` to source control.
- Rotate credentials if they were ever exposed in a commit or chat.
