# Panduan Rilis SIOMAY

Dokumen ini menjelaskan proses rilis pilot SIOMAY melalui GitHub Releases.

## Identitas rilis

| Item | Nilai |
|---|---|
| Display version | `v2026.1-beta.1` |
| Windows package version | `2026.1.0.1` |
| Channel awal | `pilot` |
| Application identifier | `id.go.bps.siomay` |
| Minimum Windows | Windows 10 versi 1809, x64 |
| Python pengembangan/CI | Python 3.14 |
| Runtime terbundel | Python 3.13 (sementara) |

Gunakan versi display untuk Git tag dan judul GitHub Release. Gunakan package version numerik untuk nama installer serta metadata Windows.

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

Gunakan `flet build`, bukan `python -m flet`; pada Flet 0.86 CLI tersedia sebagai executable terpisah. Saat ini skrip build memakai runtime Python 3.13 yang dibundel Flet, sementara Python 3.14 dipakai untuk pengembangan dan build tooling. Jangan mengganti `--python-version 3.13` menjadi 3.14 sebelum build Windows Flet dengan runtime 3.14 berhasil diuji. Sebelum membagikan build, pastikan hasilnya memuat `build/windows/siomay.exe` serta folder `build/windows/app/template/` dan `build/windows/app/input/`. Folder `data/`, `db/`, `generator/`, notebook, cache Python, dan hasil dokumen tidak boleh dimasukkan ke dalam paket rilis.

## Membuat installer

Konfigurasi Inno Setup yang dapat direproduksi tersedia di `installer/siomay.iss`. Setelah Flet build berhasil dan Inno Setup 6 terpasang, buat installer dengan:

```powershell
.\scripts\build-windows.ps1 -Installer
```

Installer pilot harus:

- memakai nama `SIOMAY-Setup-2026.1.0.1.exe`;
- memasang aplikasi dan template bawaan;
- menyediakan shortcut Start Menu dan uninstaller;
- memperbarui versi sebelumnya tanpa menghapus dokumen hasil pengguna;
- tidak memerlukan Python dari pengguna akhir.

## Publikasi GitHub Release

1. Pastikan branch `master` sudah berisi perubahan yang diuji.
2. Buat tag `v2026.1-beta.1` pada commit rilis.
3. Di GitHub, buat Release dari tag tersebut.
4. Judul rilis: `SIOMAY v2026.1-beta.1 (Pilot)`.
5. Centang **Set as a pre-release**.
6. Unggah installer dan berkas checksum SHA-256.
7. Masukkan ringkasan perubahan, persyaratan Microsoft Word untuk PDF, dan cara mengirim umpan balik.
8. Setelah aset rilis tersedia, perbarui `updates/pilot.json` dengan URL installer langsung dan SHA-256 aktual.

`updates/stable.json` hanya boleh menunjuk ke rilis yang telah lolos pengujian pilot.

## Keamanan

Sebelum distribusi luas, beli/siapkan sertifikat code-signing Windows dan tandatangani executable aplikasi maupun installer. Jangan menandai installer sebagai terverifikasi sebelum proses tersebut tersedia.

Pemeriksaan pembaruan saat ini hanya memberi informasi dan membuka halaman GitHub Release. Pembaruan otomatis yang mengunduh serta menjalankan installer baru boleh ditambahkan setelah code signing dan verifikasi SHA-256 diterapkan.