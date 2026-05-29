"""
SharePoint upload automation for NASU billing and letters zip files.

Uses Microsoft Edge with a saved browser profile (edge_sharepoint_profile/).
On first upload, save-login runs automatically — sign in once with your normal
work SSO when the browser opens. All later runs reuse the saved session.
"""

from __future__ import annotations

import argparse
import logging
import os
import re
import sys
import time
from pathlib import Path
from typing import TYPE_CHECKING

from dotenv import load_dotenv
from playwright.sync_api import BrowserContext, Page, sync_playwright

load_dotenv()

if TYPE_CHECKING:
    from tango_billing_download import RunResult

log = logging.getLogger(__name__)

DEFAULT_SHAREPOINT_URL = (
    "https://se2office365.sharepoint.com/Clients/NassauRe/"
    "Phase%201%20Implementation%20Library/Forms/"
    "Latest%201%20Year.aspx?sortField=Modified&isAscending=false"
    "&viewid=e0966e99%2D17e5%2D4209%2D8336%2D280e8b332176"
)

SHAREPOINT_LIBRARY_PATH = (
    "se2office365.sharepoint.com / Clients / NassauRe / "
    "Phase 1 Implementation Library / Latest 1 Year"
)

EDGE_PROFILE_DIR = Path(
    os.environ.get("SHAREPOINT_EDGE_PROFILE", "edge_sharepoint_profile")
)

FILE_METADATA = {
    "Item Type": "Test Results",
    "Service": "Correspondence",
    "Phase": "Solution Launch",
    "Product": "Both",
    "Workstream": "QAS",
    "Phase 1 to Done": "Phase 1 to Done",
}

# UI may show "Service" or "Services"
FIELD_LABEL_ALIASES: dict[str, list[str]] = {
    "Item Type": ["Item Type"],
    "Service": ["Service", "Services"],
    "Phase": ["Phase"],
    "Product": ["Product"],
    "Workstream": ["Workstream"],
    "Phase 1 to Done": ["Phase 1 to Done"],
}

# SharePoint properties panel — Fluent UI field title labels
FIELD_LABEL_CLASS = "ReactFieldEditor-fieldTitle"
PROPERTIES_DIALOG_SELECTOR = ".sp-itemDialog"
# Multi-select pill field — must confirm with sticky checkmark, not header blur.
SERVICE_FIELD_NAME = "Service"
SERVICE_SEARCH_PREFIX = "corr"

# Library is grouped by Item Type. New uploads sit under Unassigned until metadata is set.
ITEM_TYPE_GROUP = FILE_METADATA["Item Type"]
LIBRARY_GROUP_ALIASES: dict[str, tuple[str, ...]] = {
    "Unassigned": ("Unassigned",),
    "Test Results": ("Test Results", "Test Result"),
}


def _iter_page_roots(page: Page):
    """Yield the main page and child frames for SharePoint list locators."""
    yield page
    for frame in page.frames:
        if frame != page.main_frame:
            yield frame


def show_library_group(page: Page, group_name: str, *, log_miss: bool = True) -> bool:
    """
    Expand/select an Item Type group (e.g. Unassigned, Test Results).

    After Item Type is set to Test Results, files leave Unassigned and appear
    under the Test Results section instead.
    """
    if not _page_alive(page):
        return False

    labels = LIBRARY_GROUP_ALIASES.get(group_name, (group_name,))
    log.info("Opening library group: %s...", group_name)

    for root in _iter_page_roots(page):
        for label in labels:
            selectors = [
                f"[data-automationid='GroupedListGroupHeader']:has-text('{label}')",
                f"[data-automationid*='GroupHeader']:has-text('{label}')",
                f"[data-automationid*='groupHeader']:has-text('{label}')",
                f"[class*='groupHeader']:has-text('{label}')",
                f"[role='button']:has-text('{label}')",
                f"[role='link']:has-text('{label}')",
                f"a:has-text('{label}')",
                f"button:has-text('{label}')",
            ]
            for selector in selectors:
                loc = root.locator(selector).first
                try:
                    if loc.count() and loc.is_visible():
                        loc.click()
                        page.wait_for_timeout(2500)
                        log.info("Opened %s via: %s", group_name, selector)
                        return True
                except Exception:
                    continue

            # Match headers with counts, e.g. "Unassigned (3)" or "Test Results (1)"
            pattern = re.compile(rf"^{re.escape(label)}(\s*\(\d+\))?$", re.I)
            try:
                matches = root.get_by_text(pattern)
                for i in range(min(matches.count(), 8)):
                    el = matches.nth(i)
                    if el.is_visible():
                        el.click()
                        page.wait_for_timeout(2500)
                        log.info("Opened %s (group header text match).", group_name)
                        return True
            except Exception:
                continue

    if log_miss:
        log.info(
            "%s group header not visible — will use search or scan the full grid.",
            group_name,
        )
    return False


def show_unassigned_files(page: Page) -> bool:
    """Show newly uploaded files (Item Type not yet set)."""
    return show_library_group(page, "Unassigned", log_miss=False)


def show_test_results_files(page: Page) -> bool:
    """Show files after Item Type has been set to Test Results."""
    return show_library_group(page, ITEM_TYPE_GROUP, log_miss=False)


def _find_row_in_current_view(page: Page, pattern: re.Pattern[str]):
    """Return the first grid row matching file name in any frame."""
    for root in _iter_page_roots(page):
        for selector in (
            "[data-automationid='DetailsRow']",
            "[role='row']",
            "[data-list-item-id]",
        ):
            row = root.locator(selector).filter(has_text=pattern)
            if row.count():
                return row.first
    return None


def _find_file_row_via_search(page: Page, file_name: str, pattern: re.Pattern[str]):
    """Locate a file using the library search box."""
    search_library_for_file(page, file_name)
    page.wait_for_timeout(2000)
    handle_upload_dialogs(page, wait_ms=2_000)
    return _find_row_in_current_view(page, pattern)


def _find_file_row_in_group(
    page: Page, file_name: str, pattern: re.Pattern[str], group_name: str
):
    """Locate a file under a specific Item Type group."""
    show_library_group(page, group_name, log_miss=False)
    page.wait_for_timeout(1500)
    return _find_row_in_current_view(page, pattern)


def _optional_ms365_credentials() -> tuple[str, str] | tuple[None, None]:
    """Email/password from .env when both are set."""
    email = os.environ.get("MS365_EMAIL", "").strip()
    password = os.environ.get("MS365_PASSWORD", "").strip()
    if email and password:
        return email, password
    return None, None


def has_ms365_credentials() -> bool:
    email, password = _optional_ms365_credentials()
    return bool(email and password)


def _saved_edge_profile_exists() -> bool:
    if not EDGE_PROFILE_DIR.exists():
        return False
    return any(EDGE_PROFILE_DIR.iterdir())


def is_sharepoint_configured() -> bool:
    """True when a saved Edge browser profile exists."""
    return _saved_edge_profile_exists()


def print_sharepoint_setup_error() -> None:
    print(
        "\nERROR: SharePoint login could not be saved.\n"
        "When the Edge browser opens, sign in with your normal work account.\n"
        "The session is saved locally and reused on all future runs.\n",
        file=sys.stderr,
    )


def ensure_sharepoint_session(*, headless: bool | None = None) -> None:
    """
    Run save-login automatically when no Edge profile exists yet.

    Opens Edge so you can sign in with your work SSO once. No credentials
    are stored in .env — the session lives in edge_sharepoint_profile/.
    """
    if _saved_edge_profile_exists():
        return

    log.info(
        "No saved SharePoint session — running save-login automatically. "
        "Sign in with your work account when the browser opens."
    )
    save_sharepoint_session(headless=False)


def get_sharepoint_url() -> str:
    return os.environ.get("SHAREPOINT_URL", DEFAULT_SHAREPOINT_URL).strip()


def print_upload_summary(
    uploaded_files: list[Path],
    *,
    screenshot: Path | None = None,
) -> None:
    """Print SharePoint destination and local file paths."""
    print("\nSharePoint upload successful.")
    print(f"\nUploaded to library:\n  {SHAREPOINT_LIBRARY_PATH}")
    print(f"\nSharePoint URL:\n  {get_sharepoint_url()}")
    print("\nFiles uploaded:")
    for file_path in uploaded_files:
        print(f"  {file_path.resolve()}")
    if screenshot:
        print(f"\nScreenshot saved:\n  {screenshot.resolve()}")


def archive_uploaded_zips(zip_paths: list[Path]) -> list[Path]:
    """Move uploaded zip files into an Archive folder next to the zips."""
    existing = [path for path in zip_paths if path and path.exists()]
    if not existing:
        return []

    archive_dir = existing[0].parent / "Archive"
    archive_dir.mkdir(parents=True, exist_ok=True)

    archived: list[Path] = []
    for zip_path in existing:
        destination = archive_dir / zip_path.name
        if destination.exists():
            destination.unlink()
        zip_path.rename(destination)
        archived.append(destination)
        log.info("Archived %s to %s", zip_path.name, archive_dir.resolve())

    return archived


def _launch_edge_context(playwright, *, headless: bool) -> BrowserContext:
    EDGE_PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    kwargs = dict(
        user_data_dir=str(EDGE_PROFILE_DIR.resolve()),
        headless=headless,
        accept_downloads=True,
    )
    try:
        return playwright.chromium.launch_persistent_context(channel="msedge", **kwargs)
    except Exception:
        log.warning("Microsoft Edge not available — falling back to Chromium.")
        return playwright.chromium.launch_persistent_context(**kwargs)


def dismiss_stay_signed_in(page: Page) -> bool:
    """Click Yes/No on 'Stay signed in?' — prefer Yes to keep the session."""
    for label in ("Yes", "No"):
        btn = page.get_by_role("button", name=label)
        if btn.count() and btn.first.is_visible():
            btn.first.click()
            page.wait_for_timeout(1500)
            log.info("Answered 'Stay signed in?' with %s.", label)
            return True
    return False


def _detect_mfa_or_blocked_login(page: Page) -> None:
    """Raise when Microsoft requires MFA or manual approval."""
    patterns = (
        r"approve\s+sign[\s-]?in",
        r"enter\s+code",
        r"verify\s+your\s+identity",
        r"microsoft\s+authenticator",
        r"use\s+your\s+password\s+or\s+approved\s+device",
        r"we\s+sent\s+a\s+code",
        r"phone\s+sign[\s-]?in",
        r"enter\s+the\s+code",
    )
    for root in _iter_page_roots(page):
        for pattern in patterns:
            try:
                matches = root.get_by_text(re.compile(pattern, re.I))
                for i in range(min(matches.count(), 5)):
                    if matches.nth(i).is_visible():
                        raise RuntimeError(
                            "Microsoft login requires MFA or manual approval, which "
                            "cannot be automated. Use an app password in MS365_PASSWORD, "
                            "or complete py sharepoint_upload.py --save-login once."
                        )
            except RuntimeError:
                raise
            except Exception:
                continue


def _try_select_account(page: Page, email: str) -> bool:
    """Pick the work account on Microsoft's account chooser screen."""
    local_part = email.split("@", 1)[0]

    for root in _iter_page_roots(page):
        for pattern in (email, local_part):
            try:
                matches = root.get_by_text(re.compile(re.escape(pattern), re.I))
                for i in range(min(matches.count(), 8)):
                    candidate = matches.nth(i)
                    if candidate.is_visible():
                        candidate.click()
                        page.wait_for_timeout(1500)
                        log.info("Selected Microsoft account for %s.", email)
                        return True
            except Exception:
                continue

        for label in ("Use another account", "Sign in with another account"):
            try:
                link = root.get_by_text(label, exact=False)
                if link.count() and link.first.is_visible():
                    link.first.click()
                    page.wait_for_timeout(1500)
                    log.info("Opened '%s' on account picker.", label)
                    return True
            except Exception:
                continue

    return False


def _submit_login_form(page: Page) -> bool:
    for selector in ("#idSIButton9", "input[type='submit']", "button[type='submit']"):
        btn = page.locator(selector).first
        try:
            if btn.count() and btn.is_visible():
                btn.click()
                page.wait_for_timeout(2000)
                return True
        except Exception:
            continue
    return False


def _complete_microsoft_login(
    page: Page,
    email: str,
    password: str,
    *,
    timeout_ms: int = 120_000,
) -> None:
    """Walk through Microsoft login screens until authenticated or blocked."""
    log.info("Signing in to Microsoft 365 as %s...", email)
    deadline = time.monotonic() + timeout_ms / 1000
    password_entered = False

    while time.monotonic() < deadline:
        if not _needs_login(page):
            log.info("Microsoft sign-in complete.")
            return

        _detect_mfa_or_blocked_login(page)

        if _try_select_account(page, email):
            continue

        email_input = page.locator(
            "#i0116, input[name='loginfmt'], input[type='email']"
        ).first
        if email_input.count():
            try:
                if email_input.is_visible():
                    current = (email_input.input_value() or "").strip()
                    if current.lower() != email.lower():
                        email_input.fill(email)
                    if _submit_login_form(page):
                        continue
            except Exception:
                pass

        pwd_input = page.locator(
            "#i0118, input[name='passwd'], input[type='password']"
        ).first
        if pwd_input.count() and not password_entered:
            try:
                if pwd_input.is_visible():
                    pwd_input.fill(password)
                    password_entered = True
                    if _submit_login_form(page):
                        continue
            except Exception:
                pass

        if dismiss_stay_signed_in(page):
            continue

        page.wait_for_timeout(1000)

    if _needs_login(page):
        raise RuntimeError(
            "Microsoft login did not complete in time. "
            "Sign in with your work account when the browser opens."
        )


def microsoft_login(page: Page, email: str, password: str) -> None:
    _complete_microsoft_login(page, email, password)


def _needs_login(page: Page) -> bool:
    url = page.url.lower()
    if any(
        host in url
        for host in (
            "login.microsoftonline.com",
            "login.live.com",
            "login.microsoft.com",
            "accounts.microsoft.com",
        )
    ):
        return True

    for selector in ("#i0116", "input[name='loginfmt']", "input[type='email']"):
        loc = page.locator(selector).first
        try:
            if loc.count() and loc.is_visible():
                return True
        except Exception:
            continue

    return False


def open_sharepoint_library(
    page: Page,
    *,
    email: str | None = None,
    password: str | None = None,
    allow_manual_login: bool = False,
    login_timeout_ms: int = 300_000,
) -> None:
    if email is None or password is None:
        email, password = _optional_ms365_credentials()

    url = get_sharepoint_url()
    log.info("Opening SharePoint library...")
    page.goto(url, wait_until="domcontentloaded", timeout=120_000)
    page.wait_for_timeout(3000)

    if _needs_login(page):
        if email and password:
            if not _saved_edge_profile_exists():
                log.info(
                    "No saved Edge profile yet — signing in automatically and "
                    "saving session to %s.",
                    EDGE_PROFILE_DIR.resolve(),
                )
            microsoft_login(page, email, password)
            page.goto(url, wait_until="networkidle", timeout=120_000)
        elif allow_manual_login:
            log.info(
                "Waiting for manual sign-in in the browser (up to %d seconds)...",
                login_timeout_ms // 1000,
            )
            for _ in range(login_timeout_ms // 2000):
                if not _needs_login(page):
                    break
                page.wait_for_timeout(2000)
            if _needs_login(page):
                raise RuntimeError(
                    "SharePoint sign-in timed out. Sign in with your work account "
                    "when the browser opens."
                )
            page.goto(url, wait_until="networkidle", timeout=120_000)
        else:
            raise RuntimeError(
                "SharePoint sign-in required. Complete sign-in when the browser opens."
            )

    page.wait_for_timeout(3000)
    _wait_for_library_ready(page)
    handle_upload_dialogs(page)
    show_unassigned_files(page)


def refresh_library_page(page: Page) -> None:
    """Reload the document library so newly uploaded files appear in the grid."""
    log.info("Refreshing SharePoint library view...")
    page.goto(get_sharepoint_url(), wait_until="domcontentloaded", timeout=120_000)
    page.wait_for_timeout(4000)
    _wait_for_library_ready(page)
    handle_upload_dialogs(page)
    show_unassigned_files(page)


def _wait_for_library_ready(page: Page, timeout_ms: int = 120_000) -> None:
    selectors = [
        "[data-automationid='uploadCommand']",
        "button:has-text('Upload')",
        "span:has-text('Upload')",
        "[role='grid']",
        "[data-automationid='DetailsRow']",
    ]
    for _ in range(timeout_ms // 1000):
        dismiss_sign_in_prompt(page)
        for selector in selectors:
            if page.locator(selector).count():
                log.info("SharePoint library loaded.")
                return
        page.wait_for_timeout(1000)
    raise RuntimeError("SharePoint library did not load in time.")


def _page_alive(page: Page) -> bool:
    try:
        return not page.is_closed()
    except Exception:
        return False


def clear_library_search(page: Page) -> None:
    """Clear the library search box so the command bar is clickable again."""
    if not _page_alive(page):
        return

    search_selectors = [
        "input[aria-label*='Search' i]",
        "input[placeholder*='Search' i]",
        "[data-automationid='searchBox'] input",
        "[data-automationid='SearchBox'] input",
        "input[type='search']",
    ]

    cleared = False
    for root in _iter_page_roots(page):
        for selector in search_selectors:
            box = root.locator(selector).first
            try:
                if box.count() and box.is_visible():
                    current = (box.input_value() or "").strip()
                    if not current:
                        continue
                    box.click()
                    box.fill("")
                    box.press("Enter")
                    page.wait_for_timeout(1000)
                    cleared = True
            except Exception:
                continue

        for selector in (
            "button[aria-label*='Clear search' i]",
            "button[title*='Clear' i]",
            "[data-automationid='clearSearch']",
        ):
            btn = root.locator(selector).first
            try:
                if btn.count() and btn.is_visible():
                    btn.click(force=True)
                    page.wait_for_timeout(1000)
                    cleared = True
            except Exception:
                continue

    if cleared:
        log.debug("Cleared library search filter.")


def dismiss_sign_in_prompt(page: Page) -> bool:
    """
    Dismiss the Office/SharePoint 'Please sign in' prompt via Not Now.

    This dialog can appear mid-run even when the saved Edge session is valid.
    """
    if not _page_alive(page):
        return False

    sign_in_pattern = re.compile(r"please\s+sign\s+in", re.I)
    not_now_pattern = re.compile(r"^Not\s+[Nn]ow$")

    for root in _iter_page_roots(page):
        try:
            if root.get_by_text(sign_in_pattern).count() == 0:
                continue

            prompt_visible = False
            matches = root.get_by_text(sign_in_pattern)
            for i in range(min(matches.count(), 8)):
                if matches.nth(i).is_visible():
                    prompt_visible = True
                    break
            if not prompt_visible:
                continue
        except Exception:
            continue

        for scope in (root.locator("[role='dialog']"), root.locator("[role='alertdialog']"), root):
            try:
                btn = scope.get_by_role("button", name=not_now_pattern)
                for i in range(min(btn.count(), 5)):
                    candidate = btn.nth(i)
                    if candidate.is_visible():
                        candidate.click(force=True)
                        page.wait_for_timeout(1500)
                        log.info("Dismissed 'Please sign in' dialog via Not Now.")
                        return True
            except Exception:
                pass

            for selector in (
                "button:has-text('Not now')",
                "button:has-text('Not Now')",
                "[role='button']:has-text('Not now')",
            ):
                btn = scope.locator(selector).first
                try:
                    if btn.count() and btn.is_visible():
                        btn.click(force=True)
                        page.wait_for_timeout(1500)
                        log.info("Dismissed 'Please sign in' dialog via Not Now.")
                        return True
                except Exception:
                    continue

    return False


def dismiss_blocking_overlays(page: Page) -> bool:
    """
    Close modals/backdrops that block the upload command bar.

    Library search and Fluent UI dialogs leave a DialogSurface backdrop that
    intercepts clicks on Create or upload.
    """
    if not _page_alive(page):
        return False

    acted = False

    if dismiss_sign_in_prompt(page):
        acted = True

    if dismiss_sharepoint_dialogs(page):
        acted = True

    clear_library_search(page)

    if _properties_panel_open(page):
        try:
            click_properties_close(page)
            acted = True
        except Exception:
            try:
                page.keyboard.press("Escape")
                page.wait_for_timeout(800)
                acted = True
            except Exception:
                pass

    for _ in range(4):
        try:
            backdrop = page.locator(
                ".fui-DialogSurface__backdrop, [class*='DialogSurface__backdrop']"
            ).first
            if not backdrop.count() or not backdrop.is_visible():
                break
        except Exception:
            break

        page.keyboard.press("Escape")
        page.wait_for_timeout(800)
        acted = True
        clear_library_search(page)

    for root in _iter_page_roots(page):
        try:
            dialogs = root.locator("[role='alertdialog'], [role='dialog']")
            for i in range(min(dialogs.count(), 6)):
                dlg = dialogs.nth(i)
                if not dlg.is_visible():
                    continue
                if dlg.locator(PROPERTIES_DIALOG_SELECTOR).count():
                    continue
                for close_sel in (
                    "button[aria-label='Close']",
                    "button[title='Close']",
                    "button:has-text('Got it')",
                    "button:has-text('Dismiss')",
                    "button:has-text('Not now')",
                    "button:has-text('Not Now')",
                ):
                    close_btn = dlg.locator(close_sel).first
                    if close_btn.count() and close_btn.is_visible():
                        close_btn.click(force=True)
                        page.wait_for_timeout(1000)
                        acted = True
                        break
        except Exception:
            continue

    return acted


def dismiss_sharepoint_dialogs(page: Page) -> bool:
    """
    Dismiss blocking SharePoint dialogs (e.g. duplicate file → Keep).

    Returns True if a dialog button was clicked.
    """
    if not _page_alive(page):
        return False

    if dismiss_sign_in_prompt(page):
        return True

    # Never auto-click "Close" — it closes the properties panel (.sp-itemDialog-closeBtn).
    button_labels = ["Keep", "Keep both", "Replace", "OK", "Got it", "Dismiss"]

    try:
        for label in button_labels:
            try:
                buttons = page.get_by_role(
                    "button", name=re.compile(rf"^{re.escape(label)}$", re.I)
                )
                count = buttons.count()
            except Exception:
                continue

            for i in range(min(count, 10)):
                btn = buttons.nth(i)
                try:
                    if not btn.is_visible():
                        continue
                    if btn.evaluate(
                        "el => !!el.closest('.sp-itemDialog')"
                    ):
                        continue
                    btn.click(force=True)
                    page.wait_for_timeout(1500)
                    log.info("Dismissed SharePoint dialog via '%s' button.", label)
                    return True
                except Exception:
                    continue

            try:
                for btn in page.locator("button").all()[:30]:
                    text = (btn.inner_text() or "").strip()
                    if text.lower() != label.lower() or not btn.is_visible():
                        continue
                    if btn.evaluate(
                        "el => !!el.closest('.sp-itemDialog')"
                    ):
                        continue
                    btn.click(force=True)
                    page.wait_for_timeout(1500)
                    log.info("Dismissed SharePoint dialog via '%s' (fallback).", label)
                    return True
            except Exception:
                continue
    except Exception as exc:
        log.debug("Dialog dismiss skipped: %s", exc)

    return False


def handle_upload_dialogs(page: Page, *, wait_ms: int = 15_000) -> None:
    """Repeatedly dismiss blocking dialogs (sign-in, Keep, etc.) during uploads."""
    if not _page_alive(page):
        return
    attempts = max(1, wait_ms // 2000)
    for _ in range(attempts):
        if not dismiss_sign_in_prompt(page) and not dismiss_sharepoint_dialogs(page):
            break
        page.wait_for_timeout(1000)


def _file_stem(file_name: str) -> str:
    stem = file_name
    if stem.lower().endswith(".zip"):
        stem = stem[:-4]
    return stem


def search_library_for_file(page: Page, file_name: str) -> None:
    """Use the library search box to locate a file by name."""
    if not _page_alive(page):
        return

    stem = _file_stem(file_name)
    log.info("Searching library for %s...", stem)

    search_selectors = [
        "input[aria-label*='Search' i]",
        "input[placeholder*='Search' i]",
        "[data-automationid='searchBox'] input",
        "[data-automationid='SearchBox'] input",
        "input[type='search']",
    ]

    for root in _iter_page_roots(page):
        for selector in search_selectors:
            box = root.locator(selector).first
            try:
                if box.count() and box.is_visible():
                    box.click()
                    box.fill("")
                    box.fill(stem)
                    box.press("Enter")
                    page.wait_for_timeout(4000)
                    handle_upload_dialogs(page, wait_ms=3_000)
                    return
            except Exception:
                continue


def _quick_file_visible(page: Page, file_name: str) -> bool:
    stem = _file_stem(file_name)
    pattern = re.compile(re.escape(stem), re.I)
    return _find_row_in_current_view(page, pattern) is not None


def file_visible_in_library(page: Page, file_name: str) -> bool:
    """
    Check if a file is visible — Unassigned (new upload), Test Results
    (after metadata), or via search.
    """
    stem = _file_stem(file_name)
    pattern = re.compile(re.escape(stem), re.I)

    found = False
    for finder in (
        lambda: _find_file_row_via_search(page, file_name, pattern),
        lambda: _find_file_row_in_group(page, file_name, pattern, "Unassigned"),
        lambda: _find_file_row_in_group(page, file_name, pattern, ITEM_TYPE_GROUP),
        lambda: _find_row_in_current_view(page, pattern),
    ):
        if finder():
            found = True
            break

    clear_library_search(page)
    dismiss_blocking_overlays(page)
    return found


def file_visible_in_unassigned(page: Page, file_name: str) -> bool:
    """Backward-compatible alias — checks Unassigned first, then search."""
    return file_visible_in_library(page, file_name)


def wait_for_file_in_unassigned(
    page: Page, file_name: str, *, timeout_ms: int = 120_000
) -> None:
    """Wait until an uploaded file appears (Unassigned group or library search)."""
    log.info("Waiting for %s to appear after upload...", file_name)
    for attempt in range(max(1, timeout_ms // 4000)):
        handle_upload_dialogs(page, wait_ms=2_000)
        if file_visible_in_library(page, file_name):
            log.info("%s is visible in the library.", file_name)
            return
        if attempt % 2 == 1:
            refresh_library_page(page)
        else:
            show_unassigned_files(page)
            page.wait_for_timeout(4000)
    raise RuntimeError(
        f"{file_name} did not appear after upload. "
        "Check for a Keep/duplicate dialog in the browser."
    )


def upload_once(page: Page, file_path: Path) -> None:
    """
    Upload exactly one file: Create or upload → Files.

    Uses a single file-chooser intercept (no fallback — avoids double upload).
    """
    log.info("Uploading %s (once)...", file_path.name)
    resolved = str(file_path.resolve())

    dismiss_blocking_overlays(page)
    handle_upload_dialogs(page)
    open_upload_menu(page)

    with page.expect_file_chooser(timeout=30_000) as chooser_info:
        click_upload_files_menu(page)
    chooser_info.value.set_files(resolved)
    log.info("File attached via upload menu.")

    handle_upload_dialogs(page, wait_ms=25_000)

    try:
        page.keyboard.press("Escape")
    except Exception:
        pass
    page.wait_for_timeout(1000)


def configure_from_unassigned(page: Page, file_name: str) -> None:
    """
    More Actions → More → Properties → set metadata (Item Type = Test Results).

    File may be in Unassigned before properties; it moves to Test Results after.
    """
    set_file_properties(page, file_name)


def upload_and_configure_file(page: Page, file_path: Path) -> None:
    """
    Per-file flow:
      1. Upload once (skip if file already in library)
      2. Find file (Unassigned / search) → set properties
      3. After Item Type is set, file appears under Test Results
    """
    handle_upload_dialogs(page)

    if file_visible_in_library(page, file_path.name):
        log.info(
            "%s already in library — skipping upload, configuring properties only.",
            file_path.name,
        )
    else:
        upload_once(page, file_path)
        wait_for_file_in_unassigned(page, file_path.name)

    configure_from_unassigned(page, file_path.name)


def _find_file_row(page: Page, file_name: str, *, timeout_ms: int = 60_000):
    """
    Find a file row — tries search, Unassigned, then Test Results.

    New uploads sit in Unassigned; after Item Type is set they move to
    Test Results and are no longer under Unassigned.
    """
    stem = _file_stem(file_name)
    pattern = re.compile(re.escape(stem), re.I)

    finders = (
        ("search", lambda: _find_file_row_via_search(page, file_name, pattern)),
        ("Unassigned", lambda: _find_file_row_in_group(page, file_name, pattern, "Unassigned")),
        (
            ITEM_TYPE_GROUP,
            lambda: _find_file_row_in_group(page, file_name, pattern, ITEM_TYPE_GROUP),
        ),
        ("grid", lambda: _find_row_in_current_view(page, pattern)),
    )

    for attempt in range(max(1, timeout_ms // 3000)):
        if not _page_alive(page):
            raise RuntimeError("Browser page closed unexpectedly.")

        handle_upload_dialogs(page, wait_ms=2_000)

        for label, finder in finders:
            row = finder()
            if row:
                log.info("Found %s via %s.", file_name, label)
                return row

        page.wait_for_timeout(3000)

    raise RuntimeError(
        f"Could not find {file_name} in the library. "
        f"Look under Unassigned (new upload) or {ITEM_TYPE_GROUP} (after metadata)."
    )


def click_upload_files_menu(page: Page) -> None:
    """Click Files in the upload dropdown (use with expect_file_chooser)."""
    files_selectors = [
        "[data-automationid='upload-file']",
        "[role='menuitem']:has-text('Files')",
        "button:has-text('Files')",
        "span:has-text('Files')",
    ]
    for selector in files_selectors:
        loc = page.locator(selector).first
        if loc.count() and loc.is_visible():
            loc.click()
            log.debug("Clicked Files menu: %s", selector)
            return
    page.get_by_text("Files", exact=True).first.click()


def open_upload_menu(page: Page) -> None:
    """Open Create or upload / Upload dropdown."""
    dismiss_blocking_overlays(page)
    handle_upload_dialogs(page)

    menu_triggers = [
        "[data-automationid='newCommand']",
        "[data-automationid='uploadCommand']",
        "button:has-text('Create or upload')",
        "button:has-text('Upload')",
        "span:has-text('Create or upload')",
        "span:has-text('Upload')",
        "button[name='Upload']",
    ]
    for selector in menu_triggers:
        loc = page.locator(selector).first
        try:
            if loc.count() and loc.is_visible():
                dismiss_blocking_overlays(page)
                loc.click(force=True)
                log.debug("Opened upload menu via: %s", selector)
                page.wait_for_timeout(1500)
                handle_upload_dialogs(page)
                return
        except Exception:
            continue

    dismiss_blocking_overlays(page)
    page.get_by_role("button", name=re.compile(r"upload|create", re.I)).first.click(
        force=True
    )
    page.wait_for_timeout(1500)
    handle_upload_dialogs(page)


def open_row_more_actions(page: Page, file_name: str) -> None:
    """Click More Actions on the file row (Unassigned view must already be open)."""
    log.info("Opening More Actions for %s...", file_name)
    row = _find_file_row(page, file_name)
    row.scroll_into_view_if_needed()
    page.wait_for_timeout(1000)

    more_selectors = [
        "button[aria-label='More Actions']",
        "button[title='More Actions']",
        "[aria-label='More Actions']",
        "[title='More Actions']",
        "[data-automationid='moreActions']",
        "button[aria-label*='More actions' i]",
        "[data-icon-name='More']",
    ]
    for selector in more_selectors:
        btn = row.locator(selector).first
        if btn.count():
            try:
                if btn.is_visible():
                    btn.click()
                    page.wait_for_timeout(1500)
                    log.info("Opened row menu via: %s", selector)
                    return
            except Exception:
                continue

    # Icon may be nested — click parent button of More icon
    icon = row.locator("[data-icon-name='More']").first
    if icon.count():
        icon.evaluate("el => (el.closest('button') || el).click()")
        page.wait_for_timeout(1500)
        return

    row.get_by_role("button", name=re.compile(r"more actions", re.I)).first.click()
    page.wait_for_timeout(1500)


def _wait_for_context_menu(page: Page, timeout_ms: int = 10_000) -> None:
    for _ in range(timeout_ms // 500):
        if page.locator("[role='menu'], [role='menuitem'], .ms-ContextualMenu").count():
            return
        page.wait_for_timeout(500)
    log.warning("Context menu not detected; continuing anyway.")


def _click_menu_item(page: Page, name: str) -> None:
    """Click a visible context-menu item (e.g. More, Properties)."""
    log.info("Clicking menu item: %s", name)
    pattern = re.compile(rf"^{re.escape(name)}$", re.I)

    candidates = [
        page.get_by_role("menuitem", name=pattern),
        page.locator("[role='menuitem']").filter(has_text=pattern),
        page.locator(".ms-ContextualMenu-link").filter(has_text=pattern),
        page.locator("button").filter(has_text=pattern),
        page.locator("span").filter(has_text=pattern),
        page.get_by_text(name, exact=True),
    ]

    for loc in candidates:
        try:
            count = loc.count()
        except Exception:
            continue
        for i in range(min(count, 15)):
            item = loc.nth(i)
            try:
                if item.is_visible():
                    item.click(force=True)
                    page.wait_for_timeout(2000)
                    log.info("Clicked menu item '%s'.", name)
                    return
            except Exception:
                continue

    raise RuntimeError(f"Menu item not found: {name}")


def _click_menu_item_with_retry(page: Page, name: str, *, retries: int = 3) -> None:
    last_error: Exception | None = None
    for attempt in range(retries):
        try:
            _wait_for_context_menu(page)
            _click_menu_item(page, name)
            return
        except RuntimeError as exc:
            last_error = exc
            log.warning("Menu item %r not found (attempt %d/%d).", name, attempt + 1, retries)
            page.wait_for_timeout(1000)
    if last_error:
        raise last_error


def click_nested_more(page: Page) -> None:
    """Step 5: Click More (nested submenu after More Actions)."""
    _click_menu_item_with_retry(page, "More")


def click_properties(page: Page) -> None:
    """Step 6: Click Properties in the submenu."""
    _click_menu_item_with_retry(page, "Properties")


def click_properties_close(page: Page) -> None:
    """Close the properties panel (saves inline field edits)."""
    log.info("Closing properties panel...")
    dialog = page.locator(PROPERTIES_DIALOG_SELECTOR).first

    for loc in (
        dialog.locator(".sp-itemDialog-closeBtn"),
        dialog.get_by_role("button", name=re.compile(r"^Close$", re.I)),
        dialog.locator("button[title='Close']"),
        page.locator(".sp-itemDialog-closeBtn"),
        page.get_by_role("button", name=re.compile(r"^Close$", re.I)),
    ):
        try:
            if loc.count():
                btn = loc.first
                if btn.is_visible():
                    btn.click(force=True)
                    page.wait_for_timeout(2000)
                    log.info("Properties panel closed (Close).")
                    return
        except Exception:
            continue

    # Legacy name from original workflow notes — do not press Escape (discards edits).
    pattern = re.compile(r"^Cancel$", re.I)
    for scope in (dialog, page.locator("[role='dialog']"), page):
        try:
            btn = scope.get_by_role("button", name=pattern)
            if btn.count():
                for i in range(btn.count()):
                    candidate = btn.nth(i)
                    if candidate.is_visible():
                        candidate.click(force=True)
                        page.wait_for_timeout(2000)
                        log.info("Properties panel closed (Cancel).")
                        return
        except Exception:
            continue

    if not _page_alive(page):
        log.warning("Properties panel close skipped — browser page already closed.")
        return

    raise RuntimeError("Could not find Close button on properties panel.")


def click_properties_cancel(page: Page) -> None:
    """Backward-compatible alias — closes via Close, not Escape."""
    click_properties_close(page)


def _wait_for_properties_panel(page: Page, timeout_ms: int = 30_000) -> None:
    """Wait until the SharePoint metadata properties panel is open."""
    for _ in range(timeout_ms // 1000):
        if page.locator(f".{FIELD_LABEL_CLASS}").count():
            log.info("Properties panel loaded.")
            return
        if page.get_by_text("Item Type", exact=True).count():
            log.info("Properties panel loaded (Item Type visible).")
            return
        page.wait_for_timeout(1000)
    raise RuntimeError("Properties panel did not open in time.")


def _find_field_label(page: Page, label: str):
    """Locate a SharePoint Fluent UI field label by visible title text."""
    aliases = FIELD_LABEL_ALIASES.get(label, [label])

    for alias in aliases:
        pattern = re.compile(rf"^{re.escape(alias)}$", re.I)
        selectors = [
            page.locator(f"label.{FIELD_LABEL_CLASS}").filter(has_text=pattern),
            page.locator(f".{FIELD_LABEL_CLASS}").filter(has_text=pattern),
            page.locator("label.ms-Label").filter(has_text=pattern),
            page.get_by_text(alias, exact=True),
        ]

        for loc in selectors:
            if loc.count():
                candidate = loc.first
                try:
                    if candidate.is_visible():
                        return candidate
                except Exception:
                    return candidate

    raise RuntimeError(f"Property field label not found: {label} (tried {aliases})")


def _get_properties_root(page: Page):
    """Scope to the SharePoint item properties dialog when open."""
    dialog = page.locator(PROPERTIES_DIALOG_SELECTOR).first
    try:
        if dialog.count() and dialog.is_visible():
            return dialog
    except Exception:
        pass
    return page


def _find_field_editor(root, label: str):
    """Locate the ReactFieldEditor block for a metadata field label."""
    label_el = _find_field_label(root, label)
    editor = label_el.locator(
        "xpath=ancestor::div[contains(@class,'ReactFieldEditor')]"
        "[contains(@data-automationtype,'clientFormField')][1]"
    )
    if not editor.count():
        editor = label_el.locator(
            "xpath=ancestor::div[contains(@class,'ReactFieldEditor')][1]"
        )
    if not editor.count():
        raise RuntimeError(f"Field editor not found for {label}")
    return editor


def _scroll_field_into_view(page: Page, editor) -> None:
    """Scroll the properties form so a field is centered and clickable."""
    editor.scroll_into_view_if_needed()
    try:
        editor.evaluate(
            """el => {
                const node = el.closest('.ReactFieldEditor') || el;
                node.scrollIntoView({ block: 'center', inline: 'nearest' });
            }"""
        )
        page.wait_for_timeout(300)
    except Exception:
        pass


def _field_edit_active(editor, page: Page) -> bool:
    edit_core = editor.locator(".ReactFieldEditor-core--edit").first
    if not edit_core.count():
        return False
    try:
        return edit_core.is_visible()
    except Exception:
        return False


def _close_property_pickers(page: Page) -> None:
    """Close any open inline-edit callouts before editing the next field."""
    if not _page_alive(page) or not _properties_panel_open(page):
        return

    dialog = page.locator(PROPERTIES_DIALOG_SELECTOR).first
    header = page.locator("#reactClientFormHeader").first

    for _ in range(5):
        if header.count() and header.is_visible():
            try:
                header.click(force=True)
                page.wait_for_timeout(400)
            except Exception:
                pass

        still_open = False
        try:
            edits = dialog.locator(".ReactFieldEditor-core--edit")
            for i in range(min(edits.count(), 40)):
                if edits.nth(i).is_visible():
                    still_open = True
                    break
        except Exception:
            pass

        if not still_open:
            try:
                callout = page.locator(".ms-Callout").first
                if callout.count() and callout.is_visible():
                    still_open = True
            except Exception:
                pass

        if not still_open:
            break
        page.wait_for_timeout(250)


def _activate_field_edit(editor, page: Page) -> None:
    """
    SharePoint shows fields in display mode first — click to enter edit mode.
    Control is ReactFieldEditor-core--display with role=button, not a dropdown.
    """
    if _field_edit_active(editor, page):
        return

    display = editor.locator(".ReactFieldEditor-core--display[role='button']").first
    if not display.count():
        display = editor.locator(".ReactFieldEditor-core--display").first
    if not display.count():
        raise RuntimeError("Field display control not found (cannot enter edit mode).")

    last_error: Exception | None = None
    for attempt in range(3):
        try:
            _close_property_pickers(page)
            _scroll_field_into_view(page, editor)
            display.scroll_into_view_if_needed()
            display.click(force=True)
            page.wait_for_timeout(900)
            if _field_edit_active(editor, page):
                return
            # Single-select choice fields open a pill list without edit core.
            if page.locator("[role='option']").count() > 0:
                return
            if attempt == 2:
                return
        except Exception as exc:
            last_error = exc
            page.wait_for_timeout(400)

    if last_error:
        raise last_error


def _find_edit_control(editor, page: Page):
    """Find combobox/dropdown/input after entering field edit mode."""
    scopes = [editor, _get_properties_root(page), page]
    selectors = (
        "[role='combobox']",
        "button[aria-haspopup='listbox']",
        "button[aria-haspopup='true']",
        ".ms-Dropdown-title",
        ".ms-Dropdown",
        "input[type='text']",
        "input",
    )
    for scope in scopes:
        for sel in selectors:
            ctrl = scope.locator(sel).first
            if ctrl.count():
                try:
                    if ctrl.is_visible():
                        return ctrl
                except Exception:
                    return ctrl
    return None


def _normalize_sharepoint_field_value(raw: str) -> str:
    """Decode SharePoint multi-select encoding, e.g. ';#Correspondence;#' -> 'Correspondence'."""
    raw = raw.strip()
    if ";#" not in raw:
        return raw
    parts = [part.strip() for part in re.findall(r";#([^;#]+)", raw) if part.strip()]
    if parts:
        return ", ".join(dict.fromkeys(parts))
    return raw.replace(";#", "").strip()


def _field_display_value(editor) -> str:
    """Read the current committed value shown on a metadata field."""
    display = editor.locator(".ReactFieldEditor-core--display").first
    if not display.count():
        return ""

    aria = display.get_attribute("aria-label") or ""
    match = re.search(r",\s*(.+?),\s*press enter to edit", aria, re.I)
    if match:
        value = _normalize_sharepoint_field_value(match.group(1).strip())
        if value.lower() not in {"empty", ""}:
            return value

    chip_texts: list[str] = []
    for sel in (
        ".od-FieldRenderer-text",
        "[class*='pill']",
        "[class*='tag']",
        "[class*='Token']",
    ):
        chips = display.locator(sel)
        for i in range(chips.count()):
            text = chips.nth(i).inner_text(timeout=2000).strip()
            lowered = text.lower()
            if not text:
                continue
            if "select an option" in lowered or "select options" in lowered:
                continue
            if "enter value here" in lowered:
                continue
            chip_texts.append(text)
    if chip_texts:
        return ", ".join(dict.fromkeys(chip_texts))

    text = display.inner_text(timeout=3000).strip()
    lowered = text.lower()
    if text and "select an option" not in lowered and "select options" not in lowered:
        if "enter value here" not in lowered:
            return text
    return ""


def _wait_for_field_value(
    editor, expected: str, page: Page, *, timeout_ms: int = 8000
) -> str:
    """Poll until the field leaves edit mode and shows the expected value."""
    expected_lower = expected.lower()
    for _ in range(max(1, timeout_ms // 300)):
        try:
            edit_core = editor.locator(".ReactFieldEditor-core--edit").first
            if edit_core.count() and edit_core.is_visible():
                page.wait_for_timeout(300)
                continue
        except Exception:
            pass

        actual = _field_display_value(editor)
        if expected_lower in actual.lower():
            return actual
        page.wait_for_timeout(300)

    return _field_display_value(editor)


def _verify_field_value(
    editor, label: str, expected: str, page: Page | None = None
) -> None:
    if page is not None:
        actual = _wait_for_field_value(editor, expected, page)
    else:
        actual = _field_display_value(editor)
    if expected.lower() in actual.lower():
        return
    raise RuntimeError(
        f"Field {label!r} not set. Expected {expected!r}, display shows {actual!r}."
    )


def _properties_panel_open(page: Page) -> bool:
    if not _page_alive(page):
        return False
    try:
        dialog = page.locator(PROPERTIES_DIALOG_SELECTOR).first
        return dialog.count() > 0 and dialog.is_visible()
    except Exception:
        return False


def _dismiss_open_option_picker(page: Page) -> None:
    """Close an open pill/option list without closing the whole properties dialog."""
    if not _page_alive(page):
        return
    if (
        page.locator("[role='option']").count() == 0
        and page.locator("[role='checkbox']").count() == 0
    ):
        return
    header = page.locator("#reactClientFormHeader").first
    if header.count() and header.is_visible():
        header.click(force=True)
        page.wait_for_timeout(500)
        return
    if not _properties_panel_open(page):
        page.keyboard.press("Escape")
        page.wait_for_timeout(500)


def _click_sticky_checkmark(page: Page) -> bool:
    """Confirm inline edit for multi-select fields (Service)."""
    if not _page_alive(page):
        return False

    dialog = page.locator(PROPERTIES_DIALOG_SELECTOR).first
    scopes = [dialog, page.locator(PROPERTIES_DIALOG_SELECTOR)]
    icon_names = ("CheckMark", "Accept")

    for scope in scopes:
        if not scope.count():
            continue
        try:
            sticky = scope.locator(".ReactClientForm.isStickyEditButtons").first
            if sticky.count():
                sticky.scroll_into_view_if_needed()
        except Exception:
            pass

        for icon in icon_names:
            btns = scope.locator(f"button:has([data-icon-name='{icon}'])")
            try:
                for i in range(min(btns.count(), 5)):
                    btn = btns.nth(i)
                    if btn.is_visible():
                        btn.click(force=True)
                        return True
            except Exception:
                continue

    for name in ("Save", "Apply", "Confirm"):
        try:
            btn = dialog.get_by_role(
                "button", name=re.compile(rf"^{re.escape(name)}$", re.I)
            ).first
            if btn.count() and btn.is_visible():
                btn.click(force=True)
                return True
        except Exception:
            continue

    return False


def _commit_field_edit(page: Page, editor, label: str = "") -> None:
    """Leave edit mode so SharePoint saves the inline change."""
    if label == SERVICE_FIELD_NAME:
        if _click_sticky_checkmark(page):
            page.wait_for_timeout(900)
            return
        page.keyboard.press("Enter")
        page.wait_for_timeout(700)
        return

    header = page.locator("#reactClientFormHeader").first
    if header.count():
        try:
            if header.is_visible():
                header.click(force=True)
                page.wait_for_timeout(700)
                return
        except Exception:
            pass

    title = editor.locator(".ReactFieldEditor-titleContainer").first
    if title.count():
        title.click(force=True)
        page.wait_for_timeout(700)


def _wait_for_option(
    page: Page, value: str, timeout_ms: int = 8000, *, allow_checkbox: bool = False
) -> None:
    """Wait until a pill/dropdown option appears (same UI as Item Type)."""
    pattern = re.compile(rf"^{re.escape(value)}$", re.I)
    for _ in range(timeout_ms // 300):
        if page.locator("[role='option']").filter(has_text=pattern).count():
            return
        if allow_checkbox:
            if page.get_by_role("checkbox", name=value).count():
                return
            if page.locator("[role='checkbox']").filter(has_text=pattern).count():
                return
        page.wait_for_timeout(300)


def _get_service_edit_core(editor):
    return editor.locator(".ReactFieldEditor-core--edit").first


def _get_service_picker_scopes(page: Page, editor):
    """Limit Service clicks to the field editor and its open callout/listbox."""
    dialog = page.locator(PROPERTIES_DIALOG_SELECTOR).first
    edit = _get_service_edit_core(editor)
    scopes = []
    if edit.count():
        scopes.extend(
            [
                edit.locator("[role='listbox']"),
                edit,
            ]
        )
    if dialog.count():
        scopes.extend(
            [
                dialog.locator("[role='listbox']"),
                dialog.locator(".ms-Callout"),
                dialog.locator(".spDeThemeContentPanel"),
            ]
        )
    scopes.append(page.locator(".ms-Callout"))
    return scopes


def _wait_for_service_option(
    page: Page, editor, value: str, timeout_ms: int = 10_000
) -> None:
    """Wait for Correspondence inside the Service picker only (not page-wide)."""
    pattern = re.compile(rf"^{re.escape(value)}$", re.I)
    for _ in range(max(1, timeout_ms // 300)):
        for scope in _get_service_picker_scopes(page, editor):
            try:
                if not scope.count():
                    continue
                if scope.get_by_role("checkbox", name=value).first.is_visible():
                    return
                if scope.get_by_role("option", name=value).first.is_visible():
                    return
                if scope.locator("[role='option']").filter(has_text=pattern).first.is_visible():
                    return
            except Exception:
                continue
        page.wait_for_timeout(300)


def _clear_service_chips(editor, page: Page) -> None:
    """Remove pills already selected in Service edit mode."""
    edit = _get_service_edit_core(editor)
    if not edit.count():
        return

    for _ in range(15):
        removed = False
        for sel in (
            "button[title='Remove']",
            "button[aria-label*='Remove']",
            "button:has([data-icon-name='ChromeClose'])",
            ".ms-TagItem-close",
        ):
            btn = edit.locator(sel).first
            try:
                if btn.count() and btn.is_visible():
                    btn.click(force=True)
                    page.wait_for_timeout(300)
                    removed = True
                    break
            except Exception:
                continue
        if not removed:
            break


def _is_picker_item_selected(item) -> bool:
    try:
        if item.get_attribute("aria-selected") == "true":
            return True
        if item.get_attribute("aria-checked") == "true":
            return True
        return bool(
            item.evaluate(
                """el => {
                    const cls = el.className || '';
                    return cls.includes('is-selected')
                        || cls.includes('is-checked')
                        || el.getAttribute('aria-checked') === 'true'
                        || el.getAttribute('aria-selected') === 'true';
                }"""
            )
        )
    except Exception:
        return False


def _click_exact_picker_item(item, page: Page) -> None:
    item.scroll_into_view_if_needed()
    item.click(force=True)
    page.wait_for_timeout(500)


def _focus_service_search_input(page: Page, editor) -> bool:
    """Focus the Service multi-select search/combobox input."""
    edit = _get_service_edit_core(editor)
    if not edit.count():
        return False

    for sel in (
        "[role='combobox']",
        "input[type='text']",
        "input",
        "[contenteditable='true']",
    ):
        inp = edit.locator(sel).first
        try:
            if inp.count() and inp.is_visible():
                inp.click(force=True)
                page.wait_for_timeout(300)
                return True
        except Exception:
            continue

    try:
        edit.click(force=True)
        page.wait_for_timeout(300)
        return True
    except Exception:
        return False


def _filter_service_options(page: Page, editor, search_prefix: str) -> None:
    """Type into Service search to narrow the list (e.g. 'corr' -> Correspondence)."""
    if not _focus_service_search_input(page, editor):
        raise RuntimeError("Service search input not found.")

    page.keyboard.press("Control+A")
    page.wait_for_timeout(100)
    page.keyboard.press("Backspace")
    page.wait_for_timeout(200)
    page.keyboard.type(search_prefix, delay=50)
    page.wait_for_timeout(700)
    log.info("Filtered Service options with %r", search_prefix)


def _pick_service_option(page: Page, editor, value: str) -> None:
    """Click the filtered Service option (exact name match, scoped to picker)."""
    exact = re.compile(rf"^{re.escape(value)}$", re.I)

    for scope in _get_service_picker_scopes(page, editor):
        try:
            if not scope.count():
                continue

            checkbox = scope.get_by_role("checkbox", name=value)
            for i in range(min(checkbox.count(), 3)):
                item = checkbox.nth(i)
                if item.is_visible():
                    if not _is_picker_item_selected(item):
                        _click_exact_picker_item(item, page)
                    return

            option = scope.get_by_role("option", name=value)
            for i in range(min(option.count(), 3)):
                item = option.nth(i)
                if item.is_visible():
                    if not _is_picker_item_selected(item):
                        _click_exact_picker_item(item, page)
                    return

            for item in scope.locator("[role='option']").filter(has_text=exact).all()[:3]:
                if item.is_visible():
                    if not _is_picker_item_selected(item):
                        _click_exact_picker_item(item, page)
                    return
        except Exception:
            continue

    raise RuntimeError(f"Service option not found after filter: {value}")


def _select_service_value(
    page: Page,
    editor,
    value: str,
    *,
    search_prefix: str = SERVICE_SEARCH_PREFIX,
) -> None:
    """Type to filter Service list, then pick the one matching value."""
    _filter_service_options(page, editor, search_prefix)
    _wait_for_service_option(page, editor, value)
    _pick_service_option(page, editor, value)


def _click_service_save(page: Page, editor) -> bool:
    """Confirm Service multi-select — checkmark in callout or sticky footer."""
    dialog = page.locator(PROPERTIES_DIALOG_SELECTOR).first

    for loc in (
        page.locator(".calloutButtonsContainer button:has([data-icon-name='CheckMark'])"),
        editor.locator(".calloutButtonsContainer button:has([data-icon-name='CheckMark'])"),
        editor.locator("button:has([data-icon-name='CheckMark'])"),
        dialog.locator(".ReactClientForm.isStickyEditButtons button:has([data-icon-name='CheckMark'])"),
        dialog.locator("button:has([data-icon-name='CheckMark'])"),
    ):
        try:
            for i in range(min(loc.count(), 5)):
                btn = loc.nth(i)
                if not btn.is_visible():
                    continue
                if btn.evaluate(
                    "el => !!el.closest('.sp-itemDialog-closeBtn, .sp-itemDialog-commandBar')"
                ):
                    continue
                btn.scroll_into_view_if_needed()
                btn.click(force=True)
                return True
        except Exception:
            continue

    return False


def _commit_service_field(page: Page, editor) -> None:
    """Save Service and fully close its picker before the next field."""
    if _click_service_save(page, editor):
        page.wait_for_timeout(900)
    else:
        log.info("Service checkmark not found — blurring Service field to save.")
        for target in (
            editor.locator(f".{FIELD_LABEL_CLASS}").first,
            editor.locator(".ReactFieldEditor-core--display").first,
            page.locator("#reactClientFormHeader").first,
        ):
            try:
                if target.count() and target.is_visible():
                    target.click(force=True)
                    page.wait_for_timeout(700)
                    break
            except Exception:
                continue

    _close_property_pickers(page)


def _set_service_field(page: Page, label: str, value: str) -> None:
    """Service is multi-select: clear chips, pick only Correspondence, save."""
    log.info("Setting property %s = %s (multi-select)", label, value)
    root = _get_properties_root(page)
    _dismiss_open_option_picker(page)
    editor = _find_field_editor(root, label)
    editor.scroll_into_view_if_needed()
    page.wait_for_timeout(300)

    _activate_field_edit(editor, page)
    page.wait_for_timeout(800)
    _clear_service_chips(editor, page)
    _select_service_value(page, editor, value)
    page.wait_for_timeout(400)
    _commit_service_field(page, editor)
    page.wait_for_timeout(500)
    _verify_field_value(editor, label, value, page=page)
    _close_property_pickers(page)
    page.wait_for_timeout(300)


def _pick_choice_option(page: Page, root, editor, value: str) -> None:
    """Open a single-select choice picker and choose value (with one retry)."""
    last_error: Exception | None = None
    for attempt in range(2):
        try:
            if attempt:
                log.info("Retrying choice field picker for %r (attempt 2)", value)
                _close_property_pickers(page)
                _scroll_field_into_view(page, editor)
            _activate_field_edit(editor, page)
            page.wait_for_timeout(600)
            _wait_for_option(page, value, timeout_ms=6000)

            if _options_visible(page, value):
                _select_dropdown_value(root, page, value)
                return

            control = _find_edit_control(editor, page)
            if control is not None:
                tag = control.evaluate("el => el.tagName.toLowerCase()")
                if tag == "select":
                    control.select_option(label=value)
                    return
                control.click(force=True)
                page.wait_for_timeout(700)
                _select_dropdown_value(root, page, value)
                return

            _select_dropdown_value(root, page, value)
            return
        except RuntimeError as exc:
            last_error = exc
            page.wait_for_timeout(500)

    if last_error:
        raise last_error
    raise RuntimeError(f"Dropdown option not found: {value}")


def _set_dropdown_field(page: Page, label: str, value: str) -> None:
    if not _page_alive(page):
        raise RuntimeError("Browser page closed unexpectedly.")

    if label == SERVICE_FIELD_NAME:
        _set_service_field(page, label, value)
        return

    log.info("Setting property %s = %s", label, value)
    root = _get_properties_root(page)
    _close_property_pickers(page)
    editor = _find_field_editor(root, label)
    _scroll_field_into_view(page, editor)
    page.wait_for_timeout(300)

    _pick_choice_option(page, root, editor, value)

    _commit_field_edit(page, editor, label=label)
    page.wait_for_timeout(500)
    _verify_field_value(editor, label, value, page=page)
    page.wait_for_timeout(300)


def _get_properties_frame(page: Page):
    """Properties panel may render in the main page or an iframe."""
    for frame in page.frames:
        try:
            if frame.locator(f".{FIELD_LABEL_CLASS}").count() > 0:
                return frame
            if frame.get_by_text("Item Type", exact=True).count() > 0:
                return frame
        except Exception:
            continue
    return page


def dump_properties_panel(page: Page) -> list[dict]:
    """Return JSON-friendly structure of all property fields (for debugging)."""
    root = _get_properties_frame(page)
    return root.evaluate(
        """() => {
        const fields = [];
        const labels = document.querySelectorAll(
            '.ReactFieldEditor-fieldTitle, label.ms-Label.fui-Label, label.ms-Label'
        );
        for (const label of labels) {
            const text = (label.textContent || '').trim();
            if (!text) continue;
            const editor = label.closest('[class*="ReactFieldEditor"]') || label.parentElement;
            const controls = editor
                ? [...editor.querySelectorAll(
                    '[role="combobox"], button[aria-haspopup], select, button, input, [class*="Dropdown"]'
                  )].map(el => ({
                    tag: el.tagName,
                    role: el.getAttribute('role'),
                    ariaLabel: el.getAttribute('aria-label'),
                    className: (el.className || '').toString().slice(0, 120),
                    text: (el.textContent || '').trim().slice(0, 80),
                }))
                : [];
            fields.push({ label: text, controls });
        }
        return fields;
    }"""
    )


def inspect_properties_for_file(page: Page, file_name: str, output_dir: Path) -> Path:
    """
    Open Properties on a file and save panel HTML + screenshot + field dump.
    Run: py sharepoint_upload.py --inspect-properties --headed
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    show_unassigned_files(page)
    open_row_more_actions(page, file_name)
    _wait_for_context_menu(page)
    click_nested_more(page)
    click_properties(page)
    _wait_for_properties_panel(page)

    root = _get_properties_frame(page)
    html_path = output_dir / "properties_panel.html"
    png_path = output_dir / "properties_panel.png"
    json_path = output_dir / "properties_fields.json"

    html_path.write_text(root.content(), encoding="utf-8")
    page.screenshot(path=str(png_path), full_page=True)

    import json

    field_dump = dump_properties_panel(page)
    json_path.write_text(json.dumps(field_dump, indent=2), encoding="utf-8")

    log.info("Saved properties inspect files to %s", output_dir.resolve())
    print(f"\nInspect output:\n  {html_path}\n  {png_path}\n  {json_path}\n")
    print("Field labels found:")
    for field in field_dump:
        print(f"  - {field.get('label')!r} ({len(field.get('controls', []))} control(s))")

    return output_dir


def _options_visible(page: Page, value: str, *, allow_checkbox: bool = False) -> bool:
    pattern = re.compile(rf"^{re.escape(value)}$", re.I)
    if page.locator("[role='option']").filter(has_text=pattern).count() > 0:
        return True
    if allow_checkbox:
        if page.get_by_role("checkbox", name=value).count() > 0:
            return True
        if page.locator("[role='checkbox']").filter(has_text=pattern).count() > 0:
            return True
    return False


def _select_dropdown_value(
    root, page: Page, value: str, *, allow_checkbox: bool = False
) -> None:
    """Pick an option from SharePoint pill list or Fluent UI dropdown."""
    option_pattern = re.compile(rf"^{re.escape(value)}$", re.I)

    option_locators = [
        page.locator("[role='option']").filter(has_text=option_pattern),
        page.get_by_role("option", name=value),
        root.locator("[role='option']").filter(has_text=option_pattern),
        root.get_by_role("option", name=value),
        page.locator("[role='listbox'] [role='option']").filter(has_text=option_pattern),
        root.locator("[role='listbox'] [role='option']").filter(has_text=option_pattern),
        page.locator(".ms-Dropdown-item").filter(has_text=option_pattern),
        root.locator(".ms-Dropdown-item").filter(has_text=option_pattern),
    ]
    if allow_checkbox:
        option_locators = [
            page.get_by_role("checkbox", name=value),
            page.locator("[role='checkbox']").filter(has_text=option_pattern),
            root.get_by_role("checkbox", name=value),
            root.locator("[role='checkbox']").filter(has_text=option_pattern),
            *option_locators,
        ]

    for opt in option_locators:
        try:
            count = opt.count()
            for i in range(min(count, 15)):
                item = opt.nth(i)
                if item.is_visible():
                    item.scroll_into_view_if_needed()
                    item.click(force=True)
                    page.wait_for_timeout(800)
                    return
        except Exception:
            continue

    raise RuntimeError(f"Dropdown option not found: {value}")


def _open_properties_panel(page: Page, file_name: str) -> None:
    """Open More Actions → More → Properties with retries."""
    last_error: Exception | None = None
    for attempt in range(3):
        try:
            if attempt:
                page.keyboard.press("Escape")
                page.wait_for_timeout(1000)
                show_test_results_files(page)
                show_unassigned_files(page)
            open_row_more_actions(page, file_name)
            page.wait_for_timeout(500)
            click_nested_more(page)
            click_properties(page)
            _wait_for_properties_panel(page)
            return
        except Exception as exc:
            last_error = exc
            log.warning(
                "Open properties failed for %s (attempt %d/3): %s",
                file_name,
                attempt + 1,
                exc,
            )
    if last_error:
        raise last_error


def set_file_properties(page: Page, file_name: str, metadata: dict[str, str] | None = None) -> None:
    """
    Steps 4–13 for each file in Unassigned view:
      More Actions → More → Properties → set fields → Close
    """
    metadata = metadata or FILE_METADATA

    _open_properties_panel(page, file_name)

    # Steps 7–12: Set each metadata field
    failed_fields: list[str] = []
    for label, value in metadata.items():
        if not _page_alive(page):
            raise RuntimeError("Browser page closed while setting properties.")
        try:
            _set_dropdown_field(page, label, value)
            log.info("Set %s = %s", label, value)
        except Exception:
            log.exception("Failed to set field %s.", label)
            failed_fields.append(label)
            if not _page_alive(page):
                break

    # Step 13: Close panel (inline edits auto-save; Escape would discard them)
    click_properties_close(page)

    if failed_fields:
        raise RuntimeError(
            f"Properties incomplete for {file_name}. Failed fields: {', '.join(failed_fields)}"
        )

    log.info("Properties configured for %s.", file_name)


def capture_upload_screenshot(
    page: Page,
    uploaded_files: list[Path],
    output_path: Path,
) -> Path:
    """Capture a full-page screenshot of the SharePoint library showing uploaded files."""
    log.info("Capturing SharePoint screenshot...")
    page.goto(get_sharepoint_url(), wait_until="networkidle", timeout=120_000)
    page.wait_for_timeout(5000)
    _wait_for_library_ready(page)

    # After metadata, files live under Test Results — not Unassigned.
    show_test_results_files(page)
    for file_path in uploaded_files:
        try:
            row = _find_file_row(page, file_path.name)
            row.scroll_into_view_if_needed()
            page.wait_for_timeout(1000)
        except Exception:
            log.debug("Could not scroll to %s for screenshot.", file_path.name)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        page.screenshot(path=str(output_path), full_page=True)
    except Exception as exc:
        log.warning("Screenshot capture failed: %s", exc)
        raise
    log.info("Saved SharePoint screenshot: %s", output_path.resolve())
    return output_path


def default_screenshot_path(uploaded_files: list[Path]) -> Path:
    """Build output/screenshots/sharepoint_upload_{date}_{timestamp}.png"""
    from datetime import datetime

    date_part = "unknown"
    for f in uploaded_files:
        match = re.search(r"P(\d{6})NASU", f.name)
        if match:
            date_part = match.group(1)
            break

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return Path("output") / "screenshots" / f"sharepoint_upload_{date_part}_{stamp}.png"


def screenshot_sharepoint_library(
    *,
    headless: bool = True,
    output_path: Path | None = None,
    highlight_files: list[Path] | None = None,
) -> Path | None:
    """Open SharePoint and save a full-page screenshot (no upload)."""
    try:
        ensure_sharepoint_session(headless=headless)
    except Exception:
        log.exception("Automatic SharePoint save-login failed.")
        print_sharepoint_setup_error()
        return None

    email, password = _optional_ms365_credentials()
    output_path = output_path or default_screenshot_path(highlight_files or [])

    with sync_playwright() as playwright:
        context = _launch_edge_context(playwright, headless=headless)
        page = context.pages[0] if context.pages else context.new_page()
        try:
            open_sharepoint_library(page, email=email, password=password)
            capture_upload_screenshot(page, highlight_files or [], output_path)
        finally:
            context.close()

    return output_path


def save_sharepoint_session(*, headless: bool | None = None) -> None:
    """
    Open SharePoint in Edge and save the login profile locally.

    Uses your normal work SSO sign-in in the browser — nothing is stored in .env.
    """
    email, password = _optional_ms365_credentials()
    if headless is None:
        headless = bool(email and password)

    with sync_playwright() as playwright:
        context = _launch_edge_context(playwright, headless=headless)
        page = context.pages[0] if context.pages else context.new_page()
        try:
            open_sharepoint_library(
                page,
                email=email,
                password=password,
                allow_manual_login=True,
                login_timeout_ms=300_000,
            )
            log.info("SharePoint session saved in %s", EDGE_PROFILE_DIR.resolve())
        finally:
            context.close()


def upload_files_to_sharepoint(
    files: list[Path],
    *,
    headless: bool = True,
    screenshot: bool = True,
    screenshot_path: Path | None = None,
) -> tuple[bool, Path | None]:
    """Upload zip files to SharePoint and set metadata on each."""
    existing = [f for f in files if f and f.exists()]
    if not existing:
        log.warning("No files to upload to SharePoint.")
        return False, None

    try:
        ensure_sharepoint_session(headless=headless)
    except Exception:
        log.exception("Automatic SharePoint save-login failed.")
        print_sharepoint_setup_error()
        return False, None

    email, password = _optional_ms365_credentials()
    saved_screenshot: Path | None = None

    with sync_playwright() as playwright:
        context = _launch_edge_context(playwright, headless=headless)
        page = context.pages[0] if context.pages else context.new_page()

        try:
            open_sharepoint_library(
                page,
                email=email,
                password=password,
                allow_manual_login=True,
            )

            for file_path in existing:
                upload_and_configure_file(page, file_path)

            if screenshot:
                try:
                    if page.is_closed():
                        log.warning("Skipping screenshot — browser page already closed.")
                    else:
                        path = screenshot_path or default_screenshot_path(existing)
                        saved_screenshot = capture_upload_screenshot(page, existing, path)
                except Exception:
                    log.exception("Screenshot failed (upload may still have succeeded).")

        finally:
            context.close()

    log.info("SharePoint upload complete (%d file(s)).", len(existing))
    return True, saved_screenshot


def upload_from_run_result(
    result: RunResult, *, headless: bool = True
) -> tuple[bool, Path | None]:
    """Upload billing and letters zips from a Tango run result."""
    files: list[Path] = []
    if result.billing_zip and result.billing_zip.exists():
        files.append(result.billing_zip)
    if result.letters_zip and result.letters_zip.exists():
        files.append(result.letters_zip)

    if not files:
        log.warning("No zip files available for SharePoint upload.")
        return False, None

    return upload_files_to_sharepoint(files, headless=headless)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Upload NASU zip files to SharePoint (uses saved Edge login, no .env password needed)."
    )
    parser.add_argument(
        "files",
        nargs="*",
        type=Path,
        help="Zip files to upload (billing first, then letters).",
    )
    parser.add_argument(
        "--headed",
        action="store_true",
        help="Show browser window.",
    )
    parser.add_argument(
        "--save-login",
        action="store_true",
        help="Sign in via Edge and save session to edge_sharepoint_profile/ (also runs automatically on first upload).",
    )
    parser.add_argument(
        "--screenshot-only",
        action="store_true",
        help="Open SharePoint and save a full-page screenshot (no upload).",
    )
    parser.add_argument(
        "--no-screenshot",
        action="store_true",
        help="Skip saving a screenshot after upload.",
    )
    parser.add_argument(
        "--properties-only",
        action="store_true",
        help="Skip upload; only set metadata on files already in Unassigned.",
    )
    parser.add_argument(
        "--inspect-properties",
        action="store_true",
        help="Open Properties on first zip in Unassigned and save HTML/screenshot/field dump.",
    )
    parser.add_argument(
        "--file",
        type=str,
        default="",
        help="File name for --inspect-properties (default: first P*NASU_*.zip in output/).",
    )
    return parser.parse_args()


def main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        datefmt="%H:%M:%S",
    )

    args = parse_args()

    if args.save_login:
        try:
            save_sharepoint_session(headless=False)
        except Exception:
            log.exception("Failed to save SharePoint login session.")
            return 1
        print(f"\nSession saved in {EDGE_PROFILE_DIR.resolve()}")
        print("Future runs will reuse this saved browser session — no .env credentials needed.")
        return 0

    if args.screenshot_only:
        output = Path("output")
        highlight = sorted(output.glob("P*NASU_*.zip"))
        try:
            shot = screenshot_sharepoint_library(
                headless=not args.headed,
                highlight_files=highlight,
            )
        except Exception:
            log.exception("SharePoint screenshot failed.")
            return 1
        if shot is None:
            return 2
        print(f"\nScreenshot saved: {shot.resolve()}")
        return 0

    if args.inspect_properties:
        try:
            ensure_sharepoint_session(headless=not args.headed)
        except Exception:
            log.exception("Automatic SharePoint save-login failed.")
            print_sharepoint_setup_error()
            return 2
        output = Path("output")
        file_name = args.file
        if not file_name:
            zips = sorted(output.glob("P*NASU_*.zip"))
            if not zips:
                print("No zip files in output/. Pass --file P052226NASU_BILLING.zip")
                return 2
            file_name = zips[0].name
        email, password = _optional_ms365_credentials()
        inspect_dir = output / "properties_inspect"
        try:
            with sync_playwright() as playwright:
                context = _launch_edge_context(playwright, headless=not args.headed)
                page = context.pages[0] if context.pages else context.new_page()
                try:
                    open_sharepoint_library(page, email=email, password=password)
                    inspect_properties_for_file(page, file_name, inspect_dir)
                finally:
                    context.close()
        except Exception:
            log.exception("Properties inspect failed.")
            return 1
        print("Share the files in output/properties_inspect/ if fields still fail.")
        return 0

    if args.properties_only:
        try:
            ensure_sharepoint_session(headless=not args.headed)
        except Exception:
            log.exception("Automatic SharePoint save-login failed.")
            print_sharepoint_setup_error()
            return 2
        output = Path("output")
        names: list[str] = []
        if args.file:
            names = [args.file]
        elif args.files:
            names = [f.name for f in args.files]
        else:
            names = [f.name for f in sorted(output.glob("P*NASU_*.zip"))]
        if not names:
            print("No files specified. Pass --file P052226NASU_BILLING.zip")
            return 2
        email, password = _optional_ms365_credentials()
        try:
            with sync_playwright() as playwright:
                context = _launch_edge_context(playwright, headless=not args.headed)
                page = context.pages[0] if context.pages else context.new_page()
                try:
                    open_sharepoint_library(page, email=email, password=password)
                    show_unassigned_files(page)
                    for name in names:
                        set_file_properties(page, name)
                finally:
                    context.close()
        except Exception:
            log.exception("Properties-only run failed.")
            return 1
        print(f"\nMetadata set on {len(names)} file(s): {', '.join(names)}")
        return 0

    files = args.files
    if not files:
        output = Path("output")
        files = sorted(output.glob("P*NASU_*.zip"))

    if not files:
        print("No zip files found. Pass file paths or run tango_billing_download.py first.")
        return 2

    try:
        ok, screenshot = upload_files_to_sharepoint(
            files,
            headless=not args.headed,
            screenshot=not args.no_screenshot,
        )
    except Exception:
        log.exception("SharePoint upload failed.")
        return 1

    if not ok:
        return 2

    print_upload_summary(files, screenshot=screenshot)
    archive_uploaded_zips(files)
    return 0


if __name__ == "__main__":
    sys.exit(main())
