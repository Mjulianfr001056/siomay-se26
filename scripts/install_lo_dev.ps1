<#
.SYNOPSIS
    Download and install LibreOffice (dev-only, not committed to git).

.DESCRIPTION
    Fetches the exact same LibreOffice MSI used by the GitHub Actions release
    workflow and installs it into <repo_root>/LibreOffice/ so that the app can
    convert DOCX to PDF when run locally with `python app.py`.

    The /LibreOffice/ folder is listed in .gitignore and will never be committed.

.NOTES
    Run once from the repo root:
        powershell -ExecutionPolicy Bypass -File scripts\install_lo_dev.ps1

    Re-run any time you want to upgrade or reinstall.
    Expected download size: ~340 MB.  Install time: ~2 minutes.
#>
[CmdletBinding()]
param(
    [string]$Version = "25.8.6"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Paths ──────────────────────────────────────────────────────────────────
$repoRoot    = Split-Path -Parent $PSScriptRoot
$installDest = Join-Path $repoRoot "LibreOffice"
$msiName     = "LibreOffice_${Version}_Win_x86-64.msi"
$downloadUrl = "https://download.documentfoundation.org/libreoffice/stable/$Version/win/x86_64/$msiName"
$msiPath     = Join-Path $env:TEMP $msiName
$logPath     = Join-Path $env:TEMP "libreoffice-msi.log"

Write-Host ""
Write-Host "=== LibreOffice dev-bundle installer ===" -ForegroundColor Cyan
Write-Host "Version    : $Version"
Write-Host "Source URL : $downloadUrl"
Write-Host "Install to : $installDest"
Write-Host ""

# ── Already installed? ─────────────────────────────────────────────────────
$soffice = Join-Path $installDest "program\soffice.com"
if (Test-Path $soffice -PathType Leaf) {
    Write-Host "[OK] LibreOffice already present at $installDest" -ForegroundColor Green
    Write-Host "     Delete the folder and re-run this script to reinstall."
    exit 0
}

# ── Download ───────────────────────────────────────────────────────────────
if (-not (Test-Path $msiPath -PathType Leaf)) {
    Write-Host "[1/3] Downloading $msiName (~340 MB) ..." -ForegroundColor Yellow
    & curl.exe -L --retry 5 --retry-delay 10 --retry-all-errors `
        --connect-timeout 30 --max-time 900 `
        -o $msiPath $downloadUrl
    if ($LASTEXITCODE -ne 0) {
        throw "curl.exe failed with exit code $LASTEXITCODE while downloading $downloadUrl"
    }
    $sizeMB = [math]::Round((Get-Item $msiPath).Length / 1MB, 1)
    Write-Host "    Downloaded: $msiPath ($sizeMB MB)" -ForegroundColor Green
} else {
    $sizeMB = [math]::Round((Get-Item $msiPath).Length / 1MB, 1)
    Write-Host "[1/3] MSI already cached at $msiPath ($sizeMB MB) — skipping download."
}

# ── Install ────────────────────────────────────────────────────────────────
Write-Host "[2/3] Installing into $installDest ..." -ForegroundColor Yellow
Write-Host "      (This may take 1-2 minutes. A UAC prompt may appear.)"

# msiexec requires an absolute path for INSTALLLOCATION
$absoluteDest = [System.IO.Path]::GetFullPath($installDest)
New-Item -ItemType Directory -Path $absoluteDest -Force | Out-Null

$arguments = "/i `"$msiPath`" /qn /norestart /L*v `"$logPath`" INSTALLLOCATION=`"$absoluteDest`""
$proc = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -Wait -PassThru

if ($proc.ExitCode -ne 0) {
    Write-Host "ERROR: msiexec.exe exited with code $($proc.ExitCode)." -ForegroundColor Red
    if (Test-Path $logPath) {
        Write-Host "--- Last 40 lines of MSI log ---" -ForegroundColor Red
        Get-Content $logPath -Tail 40
    }
    throw "LibreOffice installation failed."
}

# ── Verify ─────────────────────────────────────────────────────────────────
Write-Host "[3/3] Verifying installation ..." -ForegroundColor Yellow
if (-not (Test-Path $soffice -PathType Leaf)) {
    if (Test-Path $logPath) { Get-Content $logPath -Tail 40 }
    throw "Installation appeared to succeed but soffice.com is missing at $soffice"
}

Write-Host ""
Write-Host "=== Done! ===" -ForegroundColor Green
Write-Host "LibreOffice $Version installed to: $installDest"
Write-Host "soffice.com: $soffice"
Write-Host ""
Write-Host "The /LibreOffice/ folder is already in .gitignore — it will NOT be committed."
Write-Host "Run `python app.py` and PDF conversion should now be available."
