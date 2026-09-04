"""Pemeriksaan metadata pembaruan tanpa mengunduh atau menjalankan installer."""

from __future__ import annotations

import json
from dataclasses import dataclass
from urllib.parse import quote, unquote, urlparse

import requests

from src.release import REPOSITORY_URL, UPDATE_MANIFEST_URL, is_newer_package_version


@dataclass(frozen=True)
class UpdateInfo:
    """Informasi pembaruan yang aman untuk ditampilkan kepada pengguna."""

    display_version: str
    package_version: str
    release_notes_url: str
    download_url: str
    mandatory: bool
    changelog: str | None = None


def _is_official_release_url(url: str) -> bool:
    """Allow only HTTPS links belonging to this repository's release pages."""
    parsed = urlparse(url)
    release_path = urlparse(f"{REPOSITORY_URL}/releases").path
    return (
        parsed.scheme == "https"
        and parsed.netloc == "github.com"
        and (parsed.path == release_path or parsed.path.startswith(f"{release_path}/"))
    )


def _release_tag_from_url(url: str) -> str:
    """Extract a tag only from this repository's canonical release-tag URL."""
    if not _is_official_release_url(url):
        raise ValueError("Tautan catatan rilis tidak resmi.")

    parsed = urlparse(url)
    tag_prefix = f"{urlparse(REPOSITORY_URL).path}/releases/tag/"
    if not parsed.path.startswith(tag_prefix):
        raise ValueError("Tautan catatan rilis tidak memuat tag rilis.")

    tag = unquote(parsed.path[len(tag_prefix):])
    if not tag or "/" in tag:
        raise ValueError("Tag rilis tidak valid.")
    return tag


def fetch_release_changelog(release_notes_url: str, timeout: float = 5.0) -> str | None:
    """Fetch Markdown release notes for an official GitHub release tag."""
    tag = _release_tag_from_url(release_notes_url)
    repository = urlparse(REPOSITORY_URL).path.strip("/")
    api_url = (
        f"https://api.github.com/repos/{repository}/releases/tags/"
        f"{quote(tag, safe='')}"
    )
    response = requests.get(
        api_url,
        headers={"Accept": "application/vnd.github+json"},
        timeout=timeout,
    )
    response.raise_for_status()
    release = response.json()
    if not isinstance(release, dict) or release.get("tag_name") != tag:
        raise ValueError("Respons catatan rilis tidak sesuai dengan tag pembaruan.")

    body = release.get("body")
    if not isinstance(body, str) or not body.strip():
        return None
    return body.strip()


def parse_update_manifest(manifest: dict) -> UpdateInfo | None:
    """Validate and parse a release manifest, returning a newer update if any."""
    required = ("display_version", "package_version", "release_notes_url", "download_url")
    if not isinstance(manifest, dict) or any(not manifest.get(key) for key in required):
        raise ValueError("Manifest pembaruan tidak lengkap atau tidak valid.")

    release_notes_url = str(manifest["release_notes_url"])
    download_url = str(manifest["download_url"])
    if not (_is_official_release_url(release_notes_url) and _is_official_release_url(download_url)):
        raise ValueError("Manifest pembaruan memiliki tautan rilis yang tidak resmi.")

    package_version = str(manifest["package_version"])
    if not is_newer_package_version(package_version):
        return None

    return UpdateInfo(
        display_version=str(manifest["display_version"]),
        package_version=package_version,
        release_notes_url=release_notes_url,
        download_url=download_url,
        mandatory=bool(manifest.get("mandatory", False)),
    )


def check_for_update(timeout: float = 5.0) -> UpdateInfo | None:
    """Return a newer update from GitHub's public manifest, if one exists.

    The application deliberately does not download or execute installers here.
    The user is directed to the official GitHub Release page until installer
    signing and a checksum-verified updater are introduced.
    """
    response = requests.get(UPDATE_MANIFEST_URL, timeout=timeout)
    response.raise_for_status()
    # Decode using utf-8-sig to cleanly strip any UTF-8 BOM if present
    data = json.loads(response.content.decode("utf-8-sig"))
    update = parse_update_manifest(data)
    if update is None:
        return None

    # Release notes improve the decision prompt, but must never make update
    # detection fail when GitHub's API is unavailable or rate-limited.
    try:
        changelog = fetch_release_changelog(update.release_notes_url, timeout=timeout)
    except (requests.RequestException, ValueError, json.JSONDecodeError):
        changelog = None
    return UpdateInfo(
        display_version=update.display_version,
        package_version=update.package_version,
        release_notes_url=update.release_notes_url,
        download_url=update.download_url,
        mandatory=update.mandatory,
        changelog=changelog,
    )