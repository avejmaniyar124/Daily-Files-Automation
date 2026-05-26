"""
Tango PDF Workflow Approval automation.

Logs into Tango, searches for lapse letter files using the previous US business
day date, polls until files appear (typically after 4:30 PM IST the next day),
downloads output PDFs, and creates billing and letters zip archives.
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
import time
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime, time as dt_time, timedelta, timezone
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import Page, sync_playwright

load_dotenv()

BASE_URL = "https://tango.se2.com"
WORKFLOW_URL = f"{BASE_URL}/web/tango/pdf-workflow-approval"
IST = timezone(timedelta(hours=5, minutes=30))

BILLING_SUFFIXES = ("NASULFGRCLETTER", "NASULFFLPSLETTER")
LETTERS_SUFFIXES = ("NASUANCMDLETTER", "ANNUITYMLTLRLETTER", "NASUSCHLTRLETTER", "NASULIFEMLTLRLETTER")
ALL_SUFFIXES = BILLING_SUFFIXES + LETTERS_SUFFIXES

DEFAULT_POLL_INTERVAL_MINUTES = 30
DEFAULT_READY_HOUR = 16
DEFAULT_READY_MINUTE = 30
DEFAULT_POLL_END_HOUR = 23
DEFAULT_POLL_END_MINUTE = 59

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


@dataclass
class DownloadState:
    """Tracks downloaded files keyed by search suffix."""

    files_by_suffix: dict[str, list[Path]] = field(default_factory=dict)

    def record(self, suffix: str, paths: list[Path]) -> None:
        if paths:
            self.files_by_suffix[suffix] = paths

    def has_suffix(self, suffix: str) -> bool:
        return suffix in self.files_by_suffix and bool(self.files_by_suffix[suffix])

    def missing(self, suffixes: tuple[str, ...]) -> list[str]:
        return [s for s in suffixes if not self.has_suffix(s)]

    def files_for(self, suffixes: tuple[str, ...]) -> list[Path]:
        files: list[Path] = []
        for suffix in suffixes:
            files.extend(self.files_by_suffix.get(suffix, []))
        return files

    def all_found(self, suffixes: tuple[str, ...]) -> bool:
        return not self.missing(suffixes)


def previous_us_business_day(from_date: date | None = None) -> tuple[date, str]:
    """Return the previous US business day (Mon-Fri) and its MMDDYY code."""
    current = from_date or date.today()
    candidate = current - timedelta(days=1)
    while candidate.weekday() >= 5:
        candidate -= timedelta(days=1)
    return candidate, candidate.strftime("%m%d%y")


def generation_ready_at(business_day: date) -> datetime:
    """
    Files for a business date (e.g. 052226) typically appear after 4:30 PM IST
    on the following calendar day (e.g. 052326).
    """
    generation_day = business_day + timedelta(days=1)
    return datetime.combine(
        generation_day,
        dt_time(DEFAULT_READY_HOUR, DEFAULT_READY_MINUTE),
        tzinfo=IST,
    )


def poll_deadline_at(business_day: date) -> datetime:
    """Stop polling at end of the generation day (11:59 PM IST)."""
    generation_day = business_day + timedelta(days=1)
    return datetime.combine(
        generation_day,
        dt_time(DEFAULT_POLL_END_HOUR, DEFAULT_POLL_END_MINUTE),
        tzinfo=IST,
    )


def now_ist() -> datetime:
    return datetime.now(IST)


def wait_until_ready(ready_at: datetime, *, skip: bool = False) -> None:
    if skip:
        log.info("Skipping wait for generation window (--no-wait).")
        return

    current = now_ist()
    if current >= ready_at:
        log.info(
            "Generation window already open (ready since %s IST).",
            ready_at.strftime("%Y-%m-%d %H:%M"),
        )
        return

    wait_seconds = (ready_at - current).total_seconds()
    log.info(
        "Waiting until %s IST (%.0f minutes)...",
        ready_at.strftime("%Y-%m-%d %H:%M"),
        wait_seconds / 60,
    )
    time.sleep(wait_seconds)


def sleep_until_next_poll(interval_minutes: int, deadline: datetime) -> bool:
    """
    Sleep for poll interval unless that would pass the deadline.
    Returns False if polling should stop.
    """
    current = now_ist()
    if current >= deadline:
        return False

    next_check = current + timedelta(minutes=interval_minutes)
    if next_check >= deadline:
        remaining = (deadline - current).total_seconds()
        if remaining > 0:
            log.info(
                "Final wait of %.0f minutes until poll deadline %s IST.",
                remaining / 60,
                deadline.strftime("%H:%M"),
            )
            time.sleep(remaining)
        return False

    log.info("Next check in %d minutes.", interval_minutes)
    time.sleep(interval_minutes * 60)
    return True


def get_credentials() -> tuple[str, str]:
    username = os.environ.get("TANGO_USERNAME", "").strip()
    password = os.environ.get("TANGO_PASSWORD", "").strip()
    if not username or not password:
        raise ValueError(
            "Set TANGO_USERNAME and TANGO_PASSWORD in .env or environment variables."
        )
    return username, password


def login(page: Page, username: str, password: str) -> None:
    log.info("Signing in to Tango...")
    page.goto(WORKFLOW_URL, wait_until="networkidle", timeout=120_000)

    if page.locator("#_58_login").count():
        page.fill("#_58_login", username)
        page.fill("#_58_password", password)
        page.get_by_role("button", name="Sign In").click()
        page.wait_for_load_state("networkidle", timeout=120_000)

    if "release" in page.url.lower():
        log.info("Dismissing release notice...")
        page.locator("#btnok, button:has-text('OK')").first.click()
        page.wait_for_timeout(1500)

    page.goto(WORKFLOW_URL, wait_until="networkidle", timeout=120_000)
    page.wait_for_selector("img.navigationBtnCursor[src*='search-28px']", timeout=60_000)
    log.info("PDF Workflow Approval page ready.")


def wait_for_main_grid(page: Page) -> None:
    page.wait_for_selector("img.navigationBtnCursor[src*='search-28px']", timeout=60_000)
    page.wait_for_selector("#jqGrid", timeout=60_000)


def open_search_panel(page: Page) -> None:
    wait_for_main_grid(page)
    page.locator("img.navigationBtnCursor[src*='search-28px']").click()
    page.wait_for_selector("#btnSearch", state="visible", timeout=30_000)


def set_status_under_review(page: Page) -> None:
    page.evaluate(
        """() => {
            const select = document.querySelector('#STATUS select');
            if (!select) throw new Error('Status filter not found');
            select.value = '2';
            select.dispatchEvent(new Event('change', { bubbles: true }));
            const input = document.querySelector('#STATUS .custom-combobox-input');
            if (input) input.value = 'Under Review';
        }"""
    )


def wait_for_grid_load(page: Page) -> None:
    loader = page.locator("#load_jqGrid")
    if loader.count():
        try:
            loader.wait_for(state="hidden", timeout=120_000)
        except Exception:
            page.wait_for_timeout(5000)


def wait_for_detail_frame(page: Page, timeout_ms: int = 60_000):
    for _ in range(timeout_ms // 1000):
        for frame in page.frames:
            if "ApprovalDetail.jsp" in frame.url:
                if frame.locator("img[src*='download-23px'], #jqGrid tbody tr.jqgrow").count():
                    return frame
        page.wait_for_timeout(1000)
    raise RuntimeError("Detail view did not load in time.")


def download_outputs(page: Page, detail_frame, download_dir: Path) -> list[Path]:
    saved: list[Path] = []
    download_icons = detail_frame.locator("img[src*='download-23px']")
    count = download_icons.count()
    log.info("Found %d output file(s) to download.", count)

    for index in range(count):
        with page.expect_download(timeout=120_000) as download_info:
            download_icons.nth(index).click()
        download = download_info.value
        filename = download.suggested_filename or f"output_{index + 1}.pdf"
        target = download_dir / filename
        download.save_as(target)
        saved.append(target)
        log.info("Saved %s", target.name)

    return saved


def return_to_main_grid(page: Page) -> None:
    page.goto(WORKFLOW_URL, wait_until="networkidle", timeout=120_000)
    wait_for_main_grid(page)


def process_file_search(page: Page, file_prefix: str, download_dir: Path) -> list[Path]:
    log.info("Searching for %s with status UNDER REVIEW...", file_prefix)
    open_search_panel(page)
    set_status_under_review(page)
    page.locator("#FILE_NAME input").fill(file_prefix)
    page.locator("#btnSearch").click()
    wait_for_grid_load(page)
    page.wait_for_timeout(2000)

    row_count = page.locator("#jqGrid tbody tr.jqgrow").count()
    if row_count == 0:
        log.warning("No records found for %s.", file_prefix)
        return_to_main_grid(page)
        return []

    log.info("Found %d record(s). Opening detail view...", row_count)
    page.locator("img[title='Details']").first.click()
    detail_frame = wait_for_detail_frame(page)
    saved = download_outputs(page, detail_frame, download_dir)
    return_to_main_grid(page)
    return saved


def try_download_suffix(
    page: Page,
    date_code: str,
    suffix: str,
    download_dir: Path,
    state: DownloadState,
) -> bool:
    if state.has_suffix(suffix):
        return True

    file_prefix = f"P{date_code}{suffix}"
    try:
        files = process_file_search(page, file_prefix, download_dir)
    except Exception:
        log.exception("Failed while processing %s.", file_prefix)
        try:
            return_to_main_grid(page)
        except Exception:
            log.exception("Failed to recover main grid after error.")
        return False

    if files:
        state.record(suffix, files)
        return True

    return False


def poll_for_files(
    page: Page,
    date_code: str,
    download_dir: Path,
    state: DownloadState,
    *,
    target_suffixes: tuple[str, ...] = ALL_SUFFIXES,
    poll_interval_minutes: int,
    ready_at: datetime,
    deadline: datetime,
    skip_wait: bool,
) -> None:
    wait_until_ready(ready_at, skip=skip_wait)

    attempt = 0
    while True:
        attempt += 1
        current = now_ist()
        missing = state.missing(target_suffixes)
        log.info(
            "Poll attempt %d at %s IST — %d/%d file type(s) still missing.",
            attempt,
            current.strftime("%H:%M"),
            len(missing),
            len(target_suffixes),
        )

        for suffix in missing:
            try_download_suffix(page, date_code, suffix, download_dir, state)

        if state.all_found(target_suffixes):
            log.info("All file types downloaded.")
            break

        current = now_ist()
        if current >= deadline:
            log.warning(
                "Poll deadline reached (%s IST). Zipping available files.",
                deadline.strftime("%Y-%m-%d %H:%M"),
            )
            break

        if not sleep_until_next_poll(poll_interval_minutes, deadline):
            log.warning(
                "Poll window ended. Zipping available files (may be partial)."
            )
            break


def create_zip(files: list[Path], zip_path: Path) -> Path | None:
    if not files:
        return None

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for file_path in files:
            archive.write(file_path, arcname=file_path.name)
    return zip_path


@dataclass
class RunResult:
    billing_zip: Path | None
    letters_zip: Path | None
    state: DownloadState
    date_code: str
    expected_suffixes: tuple[str, ...] = ALL_SUFFIXES


def run(
    *,
    headless: bool = True,
    output_dir: Path | None = None,
    run_date: date | None = None,
    poll_interval_minutes: int = DEFAULT_POLL_INTERVAL_MINUTES,
    skip_wait: bool = False,
    single_pass: bool = False,
    billing_only: bool = False,
) -> RunResult:
    business_day, date_code = previous_us_business_day(run_date)
    target_suffixes = BILLING_SUFFIXES if billing_only else ALL_SUFFIXES
    output_dir = output_dir or Path("output")
    download_dir = output_dir / "downloads"
    download_dir.mkdir(parents=True, exist_ok=True)

    username, password = get_credentials()
    ready_at = generation_ready_at(business_day)
    deadline = poll_deadline_at(business_day)

    log.info(
        "Business date %s (%s) — files expected after %s IST, poll until %s IST.",
        business_day.strftime("%Y-%m-%d"),
        date_code,
        ready_at.strftime("%Y-%m-%d %H:%M"),
        deadline.strftime("%Y-%m-%d %H:%M"),
    )

    state = DownloadState()

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=headless)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            login(page, username, password)

            if single_pass:
                for suffix in target_suffixes:
                    try_download_suffix(page, date_code, suffix, download_dir, state)
            else:
                poll_for_files(
                    page,
                    date_code,
                    download_dir,
                    state,
                    target_suffixes=target_suffixes,
                    poll_interval_minutes=poll_interval_minutes,
                    ready_at=ready_at,
                    deadline=deadline,
                    skip_wait=skip_wait,
                )
        finally:
            browser.close()

    billing_files = state.files_for(BILLING_SUFFIXES)
    letters_files = state.files_for(LETTERS_SUFFIXES) if not billing_only else []

    billing_zip = create_zip(
        billing_files, output_dir / f"P{date_code}NASU_BILLING.zip"
    )
    letters_zip = None
    if not billing_only:
        letters_zip = create_zip(
            letters_files, output_dir / f"P{date_code}NASU_LETTERS.zip"
        )

    if billing_zip:
        log.info(
            "Created billing archive: %s (%d file(s))",
            billing_zip.name,
            len(billing_files),
        )
    else:
        log.warning("No billing files downloaded — billing zip not created.")

    if letters_zip:
        log.info(
            "Created letters archive: %s (%d file(s))",
            letters_zip.name,
            len(letters_files),
        )
    elif not billing_only:
        log.warning("No letters files downloaded — letters zip not created.")

    for suffix in state.missing(target_suffixes):
        log.warning("Missing file type: P%s%s", date_code, suffix)

    return RunResult(
        billing_zip=billing_zip,
        letters_zip=letters_zip,
        state=state,
        date_code=date_code,
        expected_suffixes=target_suffixes,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Download Tango PDF workflow files and create billing/letters zip archives. "
            "Polls every 30 minutes after 4:30 PM IST until files appear or deadline."
        )
    )
    parser.add_argument(
        "--headed",
        action="store_true",
        help="Run the browser with a visible window (useful for debugging).",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("output"),
        help="Directory for downloaded files and zip archives (default: output).",
    )
    parser.add_argument(
        "--run-date",
        type=lambda value: date.fromisoformat(value),
        default=None,
        help="Override today's date for business-day calculation (YYYY-MM-DD).",
    )
    parser.add_argument(
        "--poll-interval",
        type=int,
        default=DEFAULT_POLL_INTERVAL_MINUTES,
        help=f"Minutes between availability checks (default: {DEFAULT_POLL_INTERVAL_MINUTES}).",
    )
    parser.add_argument(
        "--no-wait",
        action="store_true",
        help="Do not wait until 4:30 PM IST before the first check.",
    )
    parser.add_argument(
        "--single-pass",
        action="store_true",
        help="Check once with no polling (useful for testing).",
    )
    parser.add_argument(
        "--billing-only",
        action="store_true",
        help=(
            "Download only billing files (NASULFGRCLETTER, NASULFFLPSLETTER) "
            "and create P{date}NASU_BILLING.zip."
        ),
    )
    parser.add_argument(
        "--no-slack",
        action="store_true",
        help="Do not send Slack alerts for missing files.",
    )
    parser.add_argument(
        "--no-sharepoint",
        action="store_true",
        help="Do not upload zip files to SharePoint.",
    )
    parser.add_argument(
        "--sharepoint-only",
        action="store_true",
        help="Skip Tango download; upload existing zips from output folder.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    if args.sharepoint_only:
        return _run_sharepoint_only(headless=not args.headed)

    try:
        result = run(
            headless=not args.headed,
            output_dir=args.output_dir,
            run_date=args.run_date,
            poll_interval_minutes=args.poll_interval,
            skip_wait=args.no_wait,
            single_pass=args.single_pass,
            billing_only=args.billing_only,
        )
    except Exception:
        log.exception("Automation failed.")
        return 1

    if result.billing_zip is None and result.letters_zip is None:
        log.error("No files were downloaded. No zip archives were created.")
        if not args.no_slack:
            _send_slack_alerts(result)
        return 2

    if not args.no_slack:
        _send_slack_alerts(result)

    if not args.no_sharepoint:
        _upload_to_sharepoint(result, headless=not args.headed)

    print("\nResults:")
    if result.billing_zip:
        print(f"  Billing: {result.billing_zip.resolve()}")
    if result.letters_zip:
        print(f"  Letters: {result.letters_zip.resolve()}")
    return 0


def _run_sharepoint_only(*, headless: bool) -> int:
    from dataclasses import dataclass

    @dataclass
    class _MinimalResult:
        billing_zip: Path | None
        letters_zip: Path | None

    output = Path("output")
    date_codes = sorted(
        {p.name[1:7] for p in output.glob("P*NASU_*.zip") if len(p.name) > 7},
        reverse=True,
    )
    if not date_codes:
        log.error("No zip files found in output/. Run the download first.")
        return 2

    code = date_codes[0]
    result = _MinimalResult(
        billing_zip=output / f"P{code}NASU_BILLING.zip",
        letters_zip=output / f"P{code}NASU_LETTERS.zip",
    )
    if not result.billing_zip.exists():
        result.billing_zip = None
    if not result.letters_zip.exists():
        result.letters_zip = None

    if not result.billing_zip and not result.letters_zip:
        log.error("No billing or letters zip found for date %s.", code)
        return 2

    if not _upload_to_sharepoint(result, headless=headless):
        return 2
    return 0


def _upload_to_sharepoint(result: RunResult, *, headless: bool) -> bool:
    try:
        from sharepoint_upload import (
            is_sharepoint_configured,
            print_sharepoint_setup_error,
            print_upload_summary,
            upload_from_run_result,
        )

        if not is_sharepoint_configured():
            print_sharepoint_setup_error()
            return False

        ok, screenshot = upload_from_run_result(result, headless=headless)
        if ok:
            files = [
                f for f in (result.billing_zip, result.letters_zip) if f and f.exists()
            ]
            print_upload_summary(files, screenshot=screenshot)
        return ok
    except Exception:
        log.exception("Failed to upload files to SharePoint.")
        return False


def _send_slack_alerts(result: RunResult) -> None:
    try:
        from slack_notify import notify_missing_files

        notify_missing_files(result)
    except Exception:
        log.exception("Failed to send Slack notification.")


if __name__ == "__main__":
    sys.exit(main())
