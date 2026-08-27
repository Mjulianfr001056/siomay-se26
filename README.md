# SIOMAY

**SIOMAY — Sistem Otomasi Massal dan Terpercaya** is a Windows desktop application for generating SE2026 administrative documents in bulk.

## Requirements

- Windows 10 or later.
- Microsoft Word desktop application, required for DOCX-to-PDF conversion.
- Internet access only when checking or downloading an application update.

## Install the portable application

1. Open the [SIOMAY Releases page](https://github.com/Mjulianfr001056/siomay-se26/releases).
2. Normal users download `SIOMAY-<version>-windows.zip` from the newest **stable** release.
3. Beta testers download `SIOMAY-Beta-<version>-windows.zip` from the newest **pre-release**.
4. Extract the ZIP completely to a writable folder, such as `Documents\SIOMAY`. Do not run SIOMAY inside the ZIP or from `Program Files`.
5. Run `SIOMAY.exe` from the extracted folder. You may create a desktop shortcut to that file.

SIOMAY checks for updates every time it opens. Choose **Update now** to download, replace, and restart the portable application automatically, or choose **Later** to be asked again on the next launch. Stable builds receive stable releases only. The separate beta build also receives beta releases.

> Windows SmartScreen can warn for a newly published unsigned application. Download only from the Releases page linked above; if you trust the publisher, select **More info** and then **Run anyway**.

## Publishing a release

Push a version tag to `master`. GitHub Actions builds both portable channels, calculates SHA-256 checksums, and creates a draft GitHub Release with the ZIP files attached.

```powershell
# Stable release
git tag v2026.1.0
git push origin v2026.1.0

# Beta release
git tag v2026.1.1-beta.1
git push origin v2026.1.1-beta.1
```

Review and publish the resulting draft release in GitHub. Tags containing a hyphen become GitHub prereleases. The workflow embeds the pushed tag's version in each compiled build; update `version` in `pyproject.toml` before creating the matching tag to keep project metadata consistent.

## Local development

```powershell
py -m pip install -r requirements.txt
py app.py
```