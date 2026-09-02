[CmdletBinding()]
param(
    [switch]$Installer
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

if (-not (Get-Command flet -ErrorAction SilentlyContinue)) {
    throw "Flet CLI tidak ditemukan. Instal dependensi proyek dengan Python 3.14 terlebih dahulu."
}

# Fail sebelum build mahal dimulai bila wheel HEIF atau DLL native-nya tidak
# tersedia pada environment build.
py -3.14 -c "from utils.images import HAS_HEIF; assert HAS_HEIF, 'pillow-heif/native HEIF codec tidak tersedia'"
if ($LASTEXITCODE -ne 0) {
    throw "Decoder HEIC/HEIF tidak tersedia. Jalankan: py -3.14 -m pip install ."
}

# Python 3.14 digunakan untuk pengembangan dan build tooling. Runtime Python
# 3.13 tetap dibundel sampai build Flet dengan runtime 3.14 tervalidasi.
flet build windows . `
    --yes `
    --arch x64 `
    --no-compile-packages `
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

$exe = Get-ChildItem -Path build -Recurse -File -Filter "siomay.exe" | Select-Object -First 1
if ($null -eq $exe) { throw "siomay.exe tidak ditemukan pada hasil build." }
$appRoot = $exe.Directory.FullName
$heifModule = Get-ChildItem -Path $appRoot -Recurse -File -Filter "_pillow_heif*.pyd" | Select-Object -First 1
$heifDll = Get-ChildItem -Path $appRoot -Recurse -File -Filter "libheif*.dll" | Select-Object -First 1
$decoderDll = Get-ChildItem -Path $appRoot -Recurse -File -Filter "libde265*.dll" | Select-Object -First 1
if ($null -eq $heifModule -or $null -eq $heifDll -or $null -eq $decoderDll) {
    throw "Build tidak lengkap: modul pillow-heif atau DLL decoder HEIC tidak terbundel."
}

if ($Installer) {
    $iscc = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if (-not $iscc) {
        $defaultIscc = Join-Path ${env:ProgramFiles(x86)} "Inno Setup 6\ISCC.exe"
        if (Test-Path $defaultIscc) {
            $isccPath = $defaultIscc
        } else {
            throw "Inno Setup Compiler (ISCC.exe) tidak ditemukan. Instal Inno Setup 6, lalu ulangi dengan -Installer."
        }
    } else {
        $isccPath = $iscc.Source
    }
    & $isccPath "$root\installer\siomay.iss"
    if ($LASTEXITCODE -ne 0) {
        throw "Pembuatan installer Inno Setup gagal."
    }
}
