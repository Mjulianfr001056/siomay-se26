"""Build-time configuration for SIOMAY releases."""

from __future__ import annotations

import os

APP_NAME = "SIOMAY"
APP_TITLE = "SIOMAY — Sistem Otomasi Massal dan Terpercaya"
APP_VERSION = "2026.1.0"
GITHUB_REPOSITORY = "Mjulianfr001056/siomay-se26"

# GitHub Actions sets SIOMAY_UPDATE_CHANNEL to "beta" only while building
# the separately distributed SIOMAY Beta portable application.
UPDATE_CHANNEL = os.environ.get("SIOMAY_UPDATE_CHANNEL", "stable").lower()
IS_BETA_BUILD = UPDATE_CHANNEL == "beta"