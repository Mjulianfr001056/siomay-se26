"""Identitas rilis dan konfigurasi kanal pembaruan SIOMAY."""

from __future__ import annotations

APP_NAME = "SIOMAY"
APP_TITLE = "SIOMAY — Sistem Otomasi Massal dan Terpercaya"
APP_FULL_NAME = "SIOMAY: Sistem Otomasi Massal dan Terpercaya"
PUBLISHER = "6304 - Muhammad Julian Firdaus, S.Tr.Stat."
DISPLAY_VERSION = "v2026.1-beta.1"
PACKAGE_VERSION = "2026.1.0.1"
RELEASE_CHANNEL = "beta"
APPLICATION_IDENTIFIER = "id.go.bps.siomay"
REPOSITORY_URL = "https://github.com/Mjulianfr001056/siomay-se26"
RELEASES_URL = f"{REPOSITORY_URL}/releases"
UPDATE_MANIFEST_URL = (
    f"https://raw.githubusercontent.com/Mjulianfr001056/siomay-se26/"
    f"master/updates/{RELEASE_CHANNEL}.json"
)


def version_key(value: str) -> tuple[int, ...]:
    """Return a comparable numeric key for Windows package versions.

    Tuples are padded to four segments so that ``"2026.1"`` is equivalent to
    ``"2026.1.0.0"`` while ``"2026.1.0.1"`` is correctly treated as newer.
    """
    try:
        parts = tuple(int(part) for part in value.split("."))
    except (AttributeError, ValueError):
        return ()
    return parts + (0,) * max(0, 4 - len(parts))


def is_newer_package_version(candidate: str) -> bool:
    """Return whether a valid candidate package version is newer than this app."""
    candidate_key = version_key(candidate)
    current_key = version_key(PACKAGE_VERSION)
    return bool(candidate_key and current_key and candidate_key > current_key)