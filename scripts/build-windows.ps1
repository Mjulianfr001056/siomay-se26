[CmdletBinding()]
param(
    [switch]$Installer
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

if (-not (Get-Command flet -ErrorAction SilentlyContinue)) {
    throw "Flet CLI tidak ditemukan. Instal dependensi proyek dengan Python 3.13 terlebih dahulu."
}

flet build windows . `
    --arch x64 `
    --python-version 3.13 `
    --project siomay `
    --artifact siomay `
    --product SIOMAY `
    --org id.go.bps `
    --company "6304 - Muhammad Julian Firdaus, S.Tr.Stat." `
    --copyright "Copyright (c) 2026 6304 - Muhammad Julian Firdaus, S.Tr.Stat." `
    --build-version 2026.1.0.1 `
    --exclude data db generator __pycache__ .git .github tests docs installer scripts updates

if ($LASTEXITCODE -ne 0) {
    throw "Build Windows Flet gagal."
}

if ($Installer) {
    $iscc = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if (-not $iscc) {
        throw "Inno Setup Compiler (ISCC.exe) tidak ditemukan. Instal Inno Setup 6, lalu ulangi dengan -Installer."
    }
    & $iscc.Source "$root\installer\siomay.iss"
    if ($LASTEXITCODE -ne 0) {
        throw "Pembuatan installer Inno Setup gagal."
    }
}