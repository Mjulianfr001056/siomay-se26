# Panduan Rilis SIOMAY

Dokumen ini menjelaskan proses rilis portable SIOMAY melalui GitHub Releases.

## Identitas rilis

| Item | Nilai |
|---|---|
| Display version | `v2026.1-beta.1` |
| Windows package version | `2026.1.0.1` |
| Channel | `beta` atau `stable`, ditentukan oleh tag |
| Application identifier | `id.go.bps.siomay` |
| Minimum Windows | Windows 10 versi 1809, x64 |
| Python pengembangan/CI | Python 3.14 |
| Runtime terbundel | Python 3.13 (sementara) |

Gunakan versi display untuk Git tag dan judul GitHub Release. Gunakan package version numerik untuk metadata Windows.

## Prasyarat build Windows

- Windows 10/11 x64.
- Python 3.14 untuk pengembangan dan menjalankan Flet CLI.
- Visual Studio 2022 atau 2026 dengan workload **Desktop development with C++**.
- Windows Developer Mode bila build memberi pesan bahwa symlink diperlukan.
- Dependensi proyek terinstal dari `pyproject.toml`.

Flet membundel Python pada hasil build; pengguna akhir tidak perlu memasang Python.

## Build aplikasi

Jalankan dari root repositori:

```powershell
py -3.14 -m pip install --upgrade pip
py -3.14 -m pip install .
.\scripts\build-windows.ps1
```

Gunakan `flet build`, bukan `python -m flet`; pada Flet 0.86 CLI tersedia sebagai executable terpisah. Saat ini skrip build memakai runtime Python 3.13 yang dibundel Flet, sementara Python 3.14 dipakai untuk pengembangan dan build tooling. Jangan mengganti `--python-version 3.13` menjadi 3.14 sebelum build Windows Flet dengan runtime 3.14 berhasil diuji. Build Windows menggunakan `--no-compile-packages` agar modul Python murni milik dependensi biner seperti Pandas tetap tersedia dalam runtime terbundel. `certifi` juga dicantumkan sebagai dependensi langsung, walaupun dipakai oleh `requests`, agar selalu disertakan oleh Flet. Runner Windows Flet menanamkan Python di dalam proses aplikasi dan tidak menyertakan `python.exe` pada hasil akhir. Karena itu, workflow rilis memastikan hasil paket memuat berkas impor penting milik `certifi`, `requests`, dan `pandas`, serta `pandas.util`; rilis gagal sebelum ZIP dibuat bila pemeriksaan ini gagal. Sebelum membagikan build, pastikan hasilnya memuat `build/windows/siomay.exe` serta folder `build/windows/app/template/` dan `build/windows/app/input/`. Folder `data/`, `db/`, `generator/`, notebook, cache Python, dan hasil dokumen tidak boleh dimasukkan ke dalam paket rilis.

## Publikasi GitHub Release

1. Pastikan branch `master` sudah berisi perubahan yang diuji.
2. Buat dan dorong satu tag untuk rilis: `v2026.1.2-beta.1` untuk beta atau `v2026.1.2` untuk stable.
3. Workflow GitHub Actions membuat satu draft release dengan satu aset `SIOMAY-<tag>-windows.zip` dan checksum SHA-256 pada catatan rilis.
4. Tag dengan suffix setelah tanda hubung otomatis ditandai sebagai **pre-release**; tag tanpa suffix menjadi rilis normal.
5. Uji ZIP dengan mengekstraknya di direktori baru yang dapat ditulis, lalu jalankan `SIOMAY\SIOMAY.exe`.
6. Publikasikan draft release setelah pengujian berhasil.
7. Perbarui `updates/beta.json` atau `updates/stable.json` dengan tag dan package version yang baru agar pemeriksaan pembaruan menunjuk ke rilis tersebut.

## Keamanan

Sebelum distribusi luas, beli/siapkan sertifikat code-signing Windows dan tandatangani executable aplikasi. Jangan menandai executable sebagai terverifikasi sebelum proses tersebut tersedia.

Pemeriksaan pembaruan saat ini hanya memberi informasi dan membuka halaman GitHub Release. Pembaruan otomatis yang mengganti file aplikasi baru boleh ditambahkan setelah code signing dan verifikasi SHA-256 diterapkan.