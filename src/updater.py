"""GitHub Releases updater for SIOMAY's portable Windows distribution."""

from __future__ import annotations

import asyncio
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
from typing import Any

import flet as ft

from src.app_config import APP_NAME, APP_VERSION, GITHUB_REPOSITORY, IS_BETA_BUILD

API_URL = f"https://api.github.com/repos/{GITHUB_REPOSITORY}/releases"
USER_AGENT = f"{APP_NAME}/{APP_VERSION}"
VERSION_RE = re.compile(r"^v?(\d+(?:\.\d+)*)(?:-([0-9A-Za-z.-]+))?$")


def _version_key(version: str) -> tuple:
    """Return a sortable SemVer-like value; a stable release ranks above a beta."""
    match = VERSION_RE.match(version.strip())
    if not match:
        return ((0,), 0, ())
    numeric = tuple(int(part) for part in match.group(1).split("."))
    prerelease = match.group(2)
    if prerelease is None:
        return (numeric, 1, ())
    tokens = tuple((0, int(part)) if part.isdigit() else (1, part.lower()) for part in prerelease.split("."))
    return (numeric, 0, tokens)


def _get_json(url: str) -> Any:
    request = urllib.request.Request(url, headers={"Accept": "application/vnd.github+json", "User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=8) as response:
        return json.loads(response.read().decode("utf-8"))


def find_available_update() -> dict[str, Any] | None:
    """Find the newest release and matching ZIP for this stable/beta build."""
    asset_prefix = "SIOMAY-Beta-" if IS_BETA_BUILD else "SIOMAY-"
    candidates = []
    for release in _get_json(API_URL):
        if release.get("draft") or (release.get("prerelease") and not IS_BETA_BUILD):
            continue
        asset = next((
            item for item in release.get("assets", [])
            if item.get("name", "").startswith(asset_prefix)
            and item["name"].endswith("-windows.zip")
            and (IS_BETA_BUILD or not item["name"].startswith("SIOMAY-Beta-"))
        ), None)
        if asset:
            candidates.append({"version": release["tag_name"].lstrip("v"), "release": release, "asset": asset})
    if not candidates:
        return None
    newest = max(candidates, key=lambda item: _version_key(item["version"]))
    return newest if _version_key(newest["version"]) > _version_key(APP_VERSION) else None


def _download(url: str, destination: Path) -> None:
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=30) as response, destination.open("wb") as output:
        while chunk := response.read(1024 * 1024):
            output.write(chunk)


def _verify_sha256(path: Path, expected: str) -> bool:
    if not expected:
        return True
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    return digest.lower() == expected.removeprefix("sha256:").lower()


def _install_root() -> Path | None:
    return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else None


def _start_replacement(zip_path: Path, install_root: Path) -> None:
    script = Path(tempfile.gettempdir()) / f"siomay-update-{os.getpid()}.ps1"
    executable = Path(sys.executable).name
    script.write_text(
        "param([int]$ProcessId,[string]$ZipPath,[string]$InstallRoot,[string]$ExecutableName)\n"
        "$ErrorActionPreference='Stop'\n"
        "Wait-Process -Id $ProcessId -ErrorAction SilentlyContinue\n"
        "$parent=Split-Path -Parent $InstallRoot\n"
        "$staging=Join-Path $parent ('.SIOMAY-update-'+[guid]::NewGuid())\n"
        "$backup=Join-Path $parent ('.SIOMAY-backup-'+[guid]::NewGuid())\n"
        "Expand-Archive -LiteralPath $ZipPath -DestinationPath $staging -Force\n"
        "$payload=Get-ChildItem -LiteralPath $staging -Directory | Select-Object -First 1\n"
        "if ($null -eq $payload -or -not (Test-Path (Join-Path $payload.FullName $ExecutableName))) { throw 'Invalid SIOMAY update archive.' }\n"
        "Move-Item -LiteralPath $InstallRoot -Destination $backup\n"
        "Move-Item -LiteralPath $payload.FullName -Destination $InstallRoot\n"
        "Remove-Item -LiteralPath $staging,$backup,$ZipPath -Recurse -Force -ErrorAction SilentlyContinue\n"
        "Start-Process -FilePath (Join-Path $InstallRoot $ExecutableName)\n",
        encoding="utf-8",
    )
    subprocess.Popen(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script), "-ProcessId", str(os.getpid()), "-ZipPath", str(zip_path), "-InstallRoot", str(install_root), "-ExecutableName", executable], creationflags=subprocess.CREATE_NO_WINDOW)


async def install_update(page: ft.Page, update: dict[str, Any]) -> None:
    zip_path = Path(tempfile.gettempdir()) / update["asset"]["name"]
    try:
        await asyncio.to_thread(_download, update["asset"]["browser_download_url"], zip_path)
        if not _verify_sha256(zip_path, update["asset"].get("digest", "")):
            raise ValueError("Downloaded file checksum does not match GitHub Release.")
        install_root = _install_root()
        if install_root is None:
            raise RuntimeError("Updates are available only from the packaged portable application.")
        _start_replacement(zip_path, install_root)
        await page.window.close()
    except Exception as error:
        zip_path.unlink(missing_ok=True)
        page.show_dialog(ft.SnackBar(content=ft.Text(f"Update could not be installed: {error}"), bgcolor=ft.Colors.RED_700))


async def check_for_updates(page: ft.Page) -> None:
    """Silently ignore unavailable internet; prompt only when an update exists."""
    if _install_root() is None:
        return
    try:
        update = await asyncio.to_thread(find_available_update)
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError, OSError):
        return
    if update is None:
        return

    async def later(_: ft.ControlEvent) -> None:
        page.pop_dialog()

    async def update_now(_: ft.ControlEvent) -> None:
        page.pop_dialog()
        await install_update(page, update)

    notes = update["release"].get("body") or "No release notes were provided."
    page.show_dialog(ft.AlertDialog(
        modal=True,
        title=ft.Text(f"SIOMAY {update['version']} is available"),
        content=ft.Container(ft.Column([ft.Text("The update will be downloaded from the official SIOMAY GitHub Release."), ft.Text(notes, selectable=True, size=12)], tight=True, scroll=ft.ScrollMode.AUTO), width=500, height=220),
        actions=[ft.TextButton("Later", on_click=later), ft.FilledButton("Update now", on_click=update_now)],
    ))