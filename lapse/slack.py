"""
Slack notifications for missing Tango PDF workflow files.

Posts directly to a channel via Incoming Webhook (.env):

  SLACK_WEBHOOK_URL=https://hooks.slack.com/services/T.../B.../...

Create the URL: Slack → Apps → Incoming Webhooks → Add to channel → copy URL.

Do NOT use hooks.slack.com/triggers/... (Workflow) — that runs a workflow and
cannot post script text to the channel unless you edit the workflow in Slack.
"""

from __future__ import annotations

import json
import logging
import os
import sys
import urllib.error
import urllib.request
from typing import TYPE_CHECKING

from dotenv import load_dotenv

load_dotenv()

if TYPE_CHECKING:
    from tango_billing_download import RunResult

log = logging.getLogger(__name__)

DEFAULT_SLACK_CHANNEL = "D07RQNA6XRV"
DEFAULT_SLACK_DM_URL = (
    "https://se2.enterprise.slack.com/archives/D07RQNA6XRV"
)


def _valid_bot_token() -> bool:
    token = _env("SLACK_BOT_TOKEN")
    return bool(
        token
        and token not in {".", "-"}
        and not token.startswith("paste-")
        and token.startswith("xoxb-")
    )


def is_slack_configured() -> bool:
    url = _env("SLACK_WEBHOOK_URL")
    if url and is_incoming_webhook_url(url):
        return True
    return _valid_bot_token()

def _env(name: str, default: str = "") -> str:
    return os.environ.get(name, default).strip()


def is_incoming_webhook_url(url: str) -> bool:
    return "/services/" in url


def is_workflow_trigger_url(url: str) -> bool:
    return "/triggers/" in url


def build_missing_files_message(
    *,
    date_code: str,
    missing_suffixes: list[str],
    downloaded_suffixes: list[str] | None = None,
    total_expected: int | None = None,
) -> str | None:
    """Build Slack message text. Returns None if nothing is missing."""
    if not missing_suffixes:
        return None

    lines = [
        "*Tango Pre-Lapse — Missing Files*",
        f"Business date code: `{date_code}`",
        "",
        "@here The following file types were *not found* in Tango PDF Workflow Approval "
        "(status: Under Review):",
    ]

    for suffix in missing_suffixes:
        lines.append(f"• `P{date_code}{suffix}`")

    if downloaded_suffixes is not None and total_expected is not None:
        lines.extend(
            [
                "",
                f"Downloaded: {len(downloaded_suffixes)}/{total_expected} file type(s).",
                "Zip archives were created with whatever files were available.",
            ]


        )

    lines.append("")
    lines.append(f"<{DEFAULT_SLACK_DM_URL}|Open Slack conversation>")
    return "\n".join(lines)


def send_slack_message(text: str) -> bool:
    """Post script text directly to a Slack channel (webhook or bot token)."""
    webhook_url = _env("SLACK_WEBHOOK_URL")

    if webhook_url and is_incoming_webhook_url(webhook_url):
        return _send_via_incoming_webhook(webhook_url, text)

    if webhook_url and is_workflow_trigger_url(webhook_url):
        log.error(
            "SLACK_WEBHOOK_URL is a Workflow trigger — it cannot post script text. "
            "Use https://api.slack.com/apps to create an Incoming Webhook, "
            "or set SLACK_BOT_TOKEN (see README / py slack_notify.py --help)."
        )
        return False

    if webhook_url and not is_incoming_webhook_url(webhook_url):
        log.error("SLACK_WEBHOOK_URL must start with https://hooks.slack.com/services/...")
        return False

    if _valid_bot_token():
        channel = _env("SLACK_CHANNEL_ID", DEFAULT_SLACK_CHANNEL)
        return _send_via_bot_token(_env("SLACK_BOT_TOKEN"), channel, text)

    log.warning(
        "Slack not configured. Set SLACK_WEBHOOK_URL (services/...) or SLACK_BOT_TOKEN in .env."
    )
    return False


def _send_via_bot_token(bot_token: str, channel: str, text: str) -> bool:
    """Post directly to channel via chat.postMessage (works in most se2 workspaces)."""
    payload = json.dumps(
        {"channel": channel, "text": text},
        ensure_ascii=False,
    ).encode("utf-8")
    log.info("Posting message to Slack channel %s via bot token...", channel)
    request = urllib.request.Request(
        "https://slack.com/api/chat.postMessage",
        data=payload,
        headers={
            "Authorization": f"Bearer {bot_token}",
            "Content-Type": "application/json; charset=utf-8",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            result = json.loads(response.read().decode("utf-8"))
            if not result.get("ok"):
                log.error("Slack API error: %s", result.get("error", "unknown"))
                return False
            log.info("Message posted to Slack channel.")
            return True
    except urllib.error.HTTPError as exc:
        log.error("Slack API HTTP error: %s", exc.read().decode("utf-8", errors="replace"))
        return False
    except urllib.error.URLError as exc:
        log.error("Slack API connection error: %s", exc.reason)
        return False


def _send_via_incoming_webhook(webhook_url: str, text: str) -> bool:
    """POST {"text": "..."} — Slack posts this directly to the channel."""
    payload = json.dumps({"text": text}, ensure_ascii=False).encode("utf-8")
    log.info("Posting message to Slack channel via Incoming Webhook...")

    request = urllib.request.Request(
        webhook_url.strip(),
        data=payload,
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            body = response.read().decode("utf-8")
            if response.status != 200 or body.strip().lower() != "ok":
                log.error("Slack webhook failed (%s): %s", response.status, body)
                return False
            log.info("Message posted to Slack channel.")
            return True
    except urllib.error.HTTPError as exc:
        log.error(
            "Slack webhook HTTP error: %s",
            exc.read().decode("utf-8", errors="replace"),
        )
        return False
    except urllib.error.URLError as exc:
        log.error("Slack webhook connection error: %s", exc.reason)
        return False


def notify_missing_files(result: RunResult) -> bool:
    """Send Slack alert for file types that were not downloaded."""
    expected = getattr(result, "expected_suffixes", None)
    if expected is None:
        from tango_billing_download import ALL_SUFFIXES

        expected = ALL_SUFFIXES

    missing = result.state.missing(expected)
    downloaded = [suffix for suffix in expected if result.state.has_suffix(suffix)]

    message = build_missing_files_message(
        date_code=result.date_code,
        missing_suffixes=missing,
        downloaded_suffixes=downloaded,
        total_expected=len(expected),
    )
    if message is None:
        log.info("All file types downloaded — no Slack notification needed.")
        return False

    if not is_slack_configured():
        log.warning(
            "Missing files detected but Slack is not configured. "
            "Set SLACK_WEBHOOK_URL to an Incoming Webhook (hooks.slack.com/services/...)."
        )
        return False

    log.info("Sending Slack alert (%d missing file type(s)).", len(missing))
    log.info("Message from script:\n%s", message)
    return send_slack_message(message)


def check_slack_config() -> int:
    """Validate Slack configuration."""
    url = _env("SLACK_WEBHOOK_URL")
    if url and is_incoming_webhook_url(url):
        print("OK: Incoming Webhook — script posts directly to channel.")
        return 0
    if _valid_bot_token():
        print("OK: Bot token — script posts directly to channel.")
        print(f"     Channel ID: {_env('SLACK_CHANNEL_ID', DEFAULT_SLACK_CHANNEL)}")
        return 0
    if url and is_workflow_trigger_url(url):
        print("WRONG: hooks.slack.com/triggers/... cannot post script text.")
        print("  Fix: use api.slack.com/apps → Incoming Webhooks, OR use SLACK_BOT_TOKEN.")
        return 1
    if url:
        print("WRONG: URL must be https://hooks.slack.com/services/...")
        return 1
    print("MISSING: Set SLACK_WEBHOOK_URL or SLACK_BOT_TOKEN in .env")
    return 1


def main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%H:%M:%S",
    )

    if len(sys.argv) > 1 and sys.argv[1] == "--check-config":
        return check_slack_config()

    if len(sys.argv) > 1 and sys.argv[1] in ("--test-webhook", "--test"):
        if check_slack_config() != 0:
            return 1
        text = (
            "*Tango Pre-Lapse — Webhook Test*\n"
            "This message was sent directly from the Python script to your channel.\n"
            f"<{DEFAULT_SLACK_DM_URL}|Open Slack>"
        )
        if send_slack_message(text):
            print("Test message posted to channel.")
            return 0
        return 1

    if len(sys.argv) > 1 and sys.argv[1] == "--test-missing":
        if check_slack_config() != 0:
            return 1
        sample = build_missing_files_message(
            date_code="052226",
            missing_suffixes=["NASULFGRCLETTER", "NASULFFLPSLETTER"],
            downloaded_suffixes=["NASUSCHLTRLETTER"],
            total_expected=3,
        )
        if not sample:
            print("Sample message is empty.")
            return 1
        if send_slack_message(sample):
            print("Missing-files alert posted to channel.")
            return 0
        return 1

    if len(sys.argv) > 1 and sys.argv[1] == "--preview-missing":
        sample = build_missing_files_message(
            date_code="052226",
            missing_suffixes=["NASULFGRCLETTER", "NASULFFLPSLETTER"],
            downloaded_suffixes=["NASUSCHLTRLETTER"],
            total_expected=3,
        )
        print(sample or "(no message — nothing missing)")
        return 0

    print("Usage:")
    print("  py slack_notify.py --check-config    Validate Slack setup")
    print("  py slack_notify.py --test-webhook     Post test message to channel")
    print("  py slack_notify.py --test-missing      Post sample missing-files alert")
    print("")
    print("Option A — Incoming Webhook (.env):")
    print("  https://api.slack.com/apps → Create app → Incoming Webhooks → ON")
    print("  SLACK_WEBHOOK_URL=https://hooks.slack.com/services/...")
    print("")
    print("Option B — Bot token (easier in se2 enterprise):")
    print("  https://api.slack.com/apps → OAuth & Permissions → chat:write → Install")
    print("  SLACK_BOT_TOKEN=xoxb-...")
    print("  SLACK_CHANNEL_ID=C... or D...")
    return 0


if __name__ == "__main__":
    sys.exit(main())
