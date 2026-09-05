# SIOMAY

## Sistem Otomasi Massal dan Terpercaya

**SIOMAY** adalah aplikasi desktop Windows untuk membuat dokumen administrasi Sensus Ekonomi 2026 (SE2026) secara massal dari data Microsoft Excel dan template Microsoft Word. Aplikasi memandu pengguna dari pemilihan dokumen hingga penyimpanan hasil, sekaligus memvalidasi input untuk mengurangi kesalahan dan menjaga konsistensi dokumen.

> **Status:** Stable<br>
> **Versi saat ini:** `v2026.1.8`<br>
> **Platform rilis:** Windows x64

## Daftar Isi

- [Fitur Utama](#fitur-utama)
- [Dokumen yang Didukung](#dokumen-yang-didukung)
- [Persyaratan Sistem](#persyaratan-sistem)
- [Instalasi](#instalasi)
- [Cara Menggunakan](#cara-menggunakan)
- [Gambar dan Bukti Dukung](#gambar-dan-bukti-dukung)
- [Format Keluaran](#format-keluaran)
- [Pembaruan Aplikasi](#pembaruan-aplikasi)
- [Privasi dan Keamanan Data](#privasi-dan-keamanan-data)
- [Pemecahan Masalah](#pemecahan-masalah)
- [Untuk Pengembang](#untuk-pengembang)
- [Rilis, Publisher, dan Lisensi](#rilis-publisher-dan-lisensi)

## Fitur Utama

- Alur kerja terpandu dalam lima langkah: pilih dokumen, siapkan template Word, unggah data Excel, generate, dan simpan hasil.
- 13 jenis/varian dokumen untuk PPL dan PML, termasuk alur terpisah untuk Termin 1 dan Termin 2.
- Template DOCX dan format input XLSX bawaan yang dapat diunduh langsung dari aplikasi.
- Validasi struktur workbook, sheet, kolom, relasi data, dan nilai yang diperlukan sesuai jenis dokumen.
- Validasi template Word sebelum generate, termasuk placeholder yang hilang atau tidak dikenal.
- Penggantian placeholder `{{nama_kolom}}` pada paragraf, tabel, header, dan footer, termasuk placeholder yang terpecah menjadi beberapa *run* Word.
- Placeholder kustom tanpa batas pada template Word; kolom pasangannya ditambahkan otomatis ke template Excel yang diunduh.
- Nilai placeholder kustom dapat berupa teks, tautan gambar, atau tautan PDF dari Google Drive; gambar juga dapat berasal dari URL HTTP(S) langsung.
- Pembuatan DOCX massal dengan log proses, progres, timer aktif, serta ringkasan durasi.
- Pengunduhan dan penyisipan bukti dukung dari Google Drive dalam format JPEG, PNG, HEIC, HEIF, atau PDF.
- Koreksi orientasi foto berdasarkan metadata EXIF dan konversi HEIC/HEIF otomatis agar dapat dimasukkan ke DOCX.
- Konversi seluruh DOCX ke PDF secara batch melalui LibreOffice *headless* yang dibundel dalam rilis Windows.
- Estimasi waktu maksimum dan sisa waktu selama konversi PDF berdasarkan jenis dan jumlah dokumen.
- Keluaran berupa ZIP DOCX asli, ZIP PDF per petugas, atau satu PDF gabungan dengan urutan dokumen tetap terjaga.
- Pemeriksaan pembaruan berdasarkan kanal stable/beta dan akses langsung ke halaman GitHub Releases resmi.

## Dokumen yang Didukung

| Kelompok | Dokumen | Catatan utama |
|---|---|---|
| Lampiran SPK | Lampiran SPK PPL | Dokumen massal untuk Petugas Lapangan |
| Lampiran SPK | Lampiran SPK PML | Dokumen massal untuk Pemeriksa Lapangan |
| BAPP Termin 1 | BAPP PPL Termin 1 | Mendukung bukti dukung Google Drive dalam grid adaptif |
| BAPP Termin 1 | BAPP PML Termin 1 | Mendukung bukti dukung Google Drive dalam grid adaptif |
| SPP Termin 1 | SPP PPL Termin 1 | Alur dan format input khusus Termin 1 |
| SPP Termin 1 | SPP PML Termin 1 | Alur dan format input khusus Termin 1 |
| BAPP Termin 2 | BAPP PPL Termin 2 | Pilihan grid adaptif atau satu gambar per halaman |
| BAPP Termin 2 | BAPP PML Termin 2 | Pilihan grid adaptif atau satu gambar per halaman |
| SPP Termin 2 | SPP PPL Termin 2 | Alur, template, dan validasi khusus Termin 2 |
| SPP Termin 2 | SPP PML Termin 2 | Alur, template, dan validasi khusus Termin 2 |
| BAST | BAST PPL | Mengolah data mitra, supervisi, alokasi tugas, dan bukti dukung |
| BAST | BAST PML | Mengolah data mitra, supervisi, alokasi tugas, dan bukti dukung |
| Bukti Terima | Bukti Terima Paket Internet | Satu dokumen multi-halaman dengan grid foto 2×2 per halaman A4; tanpa template Word |

Nomor urut BAPP Termin 2 dan BAST yang bersifat numerik diformat menjadi tiga digit, misalnya `1` menjadi `001` dan `21` menjadi `021`. Nilai alfanumerik tetap dipertahankan.

## Persyaratan Sistem

| Komponen | Persyaratan |
|---|---|
| Sistem operasi | Windows 10 64-bit versi 1809 atau lebih baru; Windows 11 didukung |
| Arsitektur | x64 |
| Ruang penyimpanan | Sediakan ruang untuk aplikasi, LibreOffice terbundel, file sementara, dan hasil dokumen |
| Microsoft Excel | Disarankan untuk mengisi dan memeriksa berkas input `.xlsx` |
| Microsoft Word desktop | Diperlukan bila ingin menyunting template DOCX; tidak diperlukan untuk konversi PDF |
| Koneksi internet | Diperlukan untuk bukti dukung Google Drive dan pemeriksaan pembaruan |
| Python | Tidak perlu diinstal oleh pengguna akhir |

> Rilis portable menyertakan LibreOffice. Folder `LibreOffice` harus tetap berada di dalam folder aplikasi, di samping `SIOMAY.exe`, agar pilihan keluaran PDF tersedia.

## Instalasi

1. Buka halaman [GitHub Releases SIOMAY](https://github.com/Mjulianfr001056/siomay-se26/releases).
2. Pilih rilis yang diinginkan dan unduh `SIOMAY-<tag>-windows.zip`.
3. Cocokkan checksum SHA-256 arsip dengan nilai yang tercantum pada catatan rilis.
4. Ekstrak **seluruh** isi ZIP ke folder yang dapat ditulis, misalnya `Documents\SIOMAY`.
5. Jalankan `SIOMAY\SIOMAY.exe` dari folder hasil ekstraksi.

Jangan menjalankan aplikasi langsung dari dalam ZIP dan jangan memindahkan `SIOMAY.exe` tanpa folder pendampingnya. Tag tanpa suffix, misalnya `v2026.1.7`, merupakan rilis stabil; tag dengan suffix, misalnya `v2026.1.7-beta.1`, merupakan prerelease/beta.

Jika Windows menampilkan peringatan keamanan, pastikan paket berasal dari halaman Releases resmi dan checksum-nya sesuai. Executable proyek saat ini belum dinyatakan memiliki tanda tangan kode Windows.

## Cara Menggunakan

### 1. Pilih Dokumen

Pilih satu jenis dokumen untuk sesi saat ini. Setiap pilihan menentukan template Word, format Excel, validator, generator, nama keluaran, dan varian PPL/PML yang digunakan.

### 2. Siapkan Template Dokumen

Untuk dokumen berbasis template:

1. Klik **Download Template Word**.
2. Sunting salinannya di Microsoft Word tanpa menghapus atau mengganti placeholder bawaan `{{nama_kolom}}`.
3. Jika diperlukan, tambahkan placeholder kustom baru menggunakan format `{{nama_kolom_baru}}`. Nama hanya boleh berisi huruf, angka, dan garis bawah, misalnya `{{nomor_surat_tambahan}}` atau `{{foto_kegiatan}}`.
4. Simpan sebagai `.docx`.
5. Klik **Pilih Template Word (.docx)** untuk mengunggah kembali template yang telah disunting.

SIOMAY memeriksa placeholder di isi dokumen, tabel, header, dan footer. Template ditolak bila ada placeholder bawaan wajib yang hilang. Jumlah placeholder kustom tidak dibatasi. Setelah template yang valid diunggah, placeholder bawaan ditampilkan dengan warna indigo dan placeholder kustom dengan warna hijau pada pratinjau. Bukti Terima Paket Internet dibuat dari dokumen kosong sehingga langkah ini tidak memerlukan template Word.

### 3. Unggah dan Validasi Data

1. Setelah template Word diterima, periksa pemberitahuan hijau di Langkah 3. Jika terdapat placeholder kustom, pemberitahuan ini menampilkan seluruh nama kolom tambahan yang akan dibuat.
2. Klik **Download Template Excel** untuk memperoleh format yang sesuai dengan dokumen terpilih. Untuk BAPP, SPP, dan BAST, setiap placeholder kustom otomatis menjadi kolom tambahan dengan nama yang sama tanpa kurung kurawal.
3. Isi workbook tanpa mengganti nama sheet atau kolom yang diwajibkan. Isi kolom kustom dengan teks biasa untuk substitusi teks, atau dengan satu URL HTTP(S) lengkap untuk menyisipkan gambar/PDF pada posisi placeholder.
4. Simpan dalam format `.xlsx`, lalu unggah ke SIOMAY.
5. Periksa ringkasan hasil validasi dan perbaiki kesalahan yang ditampilkan sebelum melanjutkan.

Jangan menggunakan format Excel dari kelompok atau termin lain. SPP Termin 1 dan Termin 2 memiliki alur input dan validasi yang terpisah.

### 4. Generate Dokumen

Periksa ringkasan dokumen, template, dan sumber data, lalu klik **Mulai Generate Dokumen**. SIOMAY menampilkan progres, log aktivitas per baris, peringatan yang tidak menghentikan batch, serta ringkasan waktu pembuatan.

Khusus BAPP Termin 2 dan BAST, pilih salah satu tata letak bukti dukung:

- **Grid adaptif:** maksimal lima gambar per halaman, dengan ukuran menyesuaikan halaman.
- **Halaman khusus:** seluruh gambar ditempatkan satu per halaman.

Setiap halaman bukti berformat PDF selalu ditempatkan pada halaman khusus,
terlepas dari pilihan tata letak. Gambar setelah PDF kembali mengikuti pilihan
grid atau halaman khusus pengguna. Tautan pada placeholder kustom mengikuti
pilihan tata letak dan orientasi bukti yang sama, tetapi tidak menambahkan judul
`BUKTI DUKUNG` pada dokumen hasil.

### 5. Simpan Hasil

Pilih format keluaran yang tersedia, tentukan lokasi penyimpanan, lalu tunggu proses selesai. Untuk keluaran PDF, dialog menampilkan timer, estimasi durasi maksimum, dan perkiraan sisa waktu. Setelah selesai, aplikasi menampilkan lokasi, ukuran berkas, ringkasan durasi, tombol **Buka Folder**, dan pilihan untuk memulai sesi baru.

## Gambar dan Bukti Dukung

Alur BAPP, BAST, dan Bukti Terima dapat menggunakan tautan Google Drive untuk mengambil foto atau tangkapan layar. Placeholder kustom pada BAPP, SPP, dan BAST juga dapat diisi dengan URL HTTP(S) gambar langsung atau URL Google Drive berisi gambar/PDF. Agar bukti dapat diunduh:

- Atur akses file menjadi **Anyone with the link / Siapa saja yang memiliki tautan**.
- Jika file berasal dari folder unggahan Google Forms, pastikan folder tersebut juga dapat diakses melalui tautan.
- Gunakan tautan file Google Drive yang valid; beberapa tautan dapat dipisahkan dengan koma pada kolom yang mendukung banyak gambar.
- Format JPEG, PNG, HEIC, HEIF, dan PDF didukung pada BAPP Termin 2 dan BAST.
  Setiap halaman PDF dirender dan disisipkan sebagai halaman khusus. Orientasi
  EXIF pada gambar diterapkan otomatis.
- Pada kolom placeholder kustom, gunakan satu URL lengkap per sel. PDF kustom didukung melalui tautan Google Drive, sedangkan URL web selain Google Drive harus mengarah ke gambar yang valid.
- Jika nilai kustom bukan URL, nilainya dimasukkan sebagai teks. Jika pengunduhan atau validasi URL gagal, URL asli tetap dimasukkan sebagai teks agar informasi tidak hilang.
- Respons HTML, file kosong, dan gambar rusak/tidak dikenal dilaporkan sebagai peringatan tanpa harus menggagalkan seluruh batch.

Perhatikan bahwa penggunaan tautan yang dapat diakses siapa saja memiliki implikasi privasi. Batasi isi gambar pada data yang memang diperlukan, dan cabut akses tautan setelah proses selesai bila kebijakan kerja mengharuskannya.

## Format Keluaran

| Pilihan | Hasil | Keterangan |
|---|---|---|
| ZIP DOCX | Arsip `.zip` berisi DOCX asli | Paling cepat karena tidak melakukan konversi PDF |
| ZIP PDF | Arsip `.zip` berisi PDF per petugas/dokumen | Seluruh DOCX dikonversi dalam satu proses batch LibreOffice |
| PDF gabungan | Satu `.pdf` multi-halaman | Disusun mengikuti urutan hasil generate dan siap dicetak dari satu berkas |

Bukti Terima Paket Internet selalu dibuat sebagai satu dokumen multi-halaman yang memuat seluruh petugas. Bila LibreOffice tidak ditemukan, aplikasi tetap dapat menyimpan ZIP DOCX, tetapi pilihan PDF dinonaktifkan.

## Pembaruan Aplikasi

SIOMAY memeriksa manifest pembaruan sesuai kanal build:

- Build stabil memeriksa kanal **stable**.
- Build beta memeriksa kanal **beta**.
- Manifest dan tautan divalidasi agar hanya mengarah ke area Releases repositori GitHub resmi.
- Aplikasi hanya menampilkan versi dan catatan rilis lalu membuka halaman unduhan resmi; aplikasi tidak mengunduh atau menjalankan installer secara otomatis.

Setelah mengunduh versi baru, ekstrak paket ke folder baru dan pertahankan seluruh struktur paket, termasuk folder `LibreOffice`.

## Privasi dan Keamanan Data

- Workbook Excel, template Word, dan dokumen hasil diproses secara lokal di komputer pengguna.
- Aplikasi tidak mengunggah data input atau hasil generate sebagai bagian dari proses pembuatan dokumen.
- Koneksi keluar digunakan untuk mengambil gambar dari tautan yang dimasukkan pengguna dan untuk mengambil metadata pembaruan.
- Hasil disimpan hanya ke lokasi yang dipilih pengguna.
- Jangan membagikan workbook, hasil dokumen, log, atau tangkapan layar yang mengandung NIK, nomor telepon, maupun data pribadi lain melalui kanal publik.

## Pemecahan Masalah

### Pilihan PDF tidak tersedia

Pastikan folder `LibreOffice` masih utuh dan berada di samping `SIOMAY.exe`. Ekstrak ulang ZIP resmi bila folder hilang atau tidak lengkap. ZIP DOCX tetap dapat digunakan tanpa LibreOffice.

### Template Word ditolak

Unduh kembali template untuk dokumen terpilih. Pertahankan semua placeholder bawaan, pastikan placeholder kustom mengikuti format `{{nama_kolom}}` dengan huruf, angka, atau garis bawah saja, lalu unggah berkas `.docx`, bukan `.doc`. Placeholder kustom boleh ditambahkan sebanyak yang diperlukan.

### Data Excel tidak valid

Gunakan template Excel yang diunduh setelah memilih dokumen. Jangan mengubah nama sheet/kolom, pastikan kolom wajib tersedia, dan gunakan format untuk termin serta peran yang benar.

### Gambar Google Drive tidak muncul

Periksa akses **Siapa saja yang memiliki tautan**, validitas URL, koneksi internet, dan isi file. Halaman login/HTML bukan file gambar dan akan ditolak.


### 💬 Feedback & Laporan Masalah

Punya kendala atau saran pengembangan? Kirimkan umpan balik Anda melalui form berikut:
🔗 http://s.bps.go.id/FeedbackSIOMAY


## Untuk Pengembang

### Teknologi

- Python 3.13–3.14 (Python 3.14 untuk pengembangan/CI; runtime Windows Flet saat ini menggunakan Python 3.13)
- [Flet](https://flet.dev/)
- pandas dan openpyxl
- python-docx
- Pillow dan pillow-heif
- pypdf
- PyMuPDF
- requests dan certifi
- LibreOffice sebagai runtime native untuk konversi PDF

### Menjalankan dari kode sumber

Prasyarat: Windows, Python 3.14, Git, dan PowerShell.

```powershell
git clone https://github.com/Mjulianfr001056/siomay-se26.git
cd siomay-se26
py -3.14 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install .
python app.py
```

PDF membutuhkan LibreOffice. Aplikasi akan mencari bundle lokal, instalasi Windows standar, lalu `soffice` pada `PATH`. Untuk memasang versi pengembangan yang sama dengan proses rilis ke folder `LibreOffice/` lokal:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\install_lo_dev.ps1
```

Unduhan LibreOffice berukuran sekitar 340 MB. Skrip memeriksa kegagalan HTTP, ukuran minimum, dan header MSI sebelum instalasi. Folder lokal `LibreOffice/` diabaikan oleh Git.

### Menjalankan validasi dan pengujian

```powershell
python -m compileall -q app.py src utils
python -m unittest discover -s tests -v
```

Suite pengujian mencakup generator dokumen, pemisahan routing workflow, placeholder DOCX bawaan dan kustom, penyisipan gambar/PDF dari URL, gambar JPEG/PNG/HEIC, tata letak serta orientasi bukti, konversi PDF batch, estimasi konversi, pembaruan/rilis, dan helper UI.

### Build Windows

Build membutuhkan Visual Studio dengan workload **Desktop development with C++** dan, pada kondisi tertentu, Windows Developer Mode. Jalankan:

```powershell
py -3.14 -m pip install .
.\scripts\build-windows.ps1
```

Untuk build installer opsional, instal Inno Setup 6 lalu gunakan:

```powershell
.\scripts\build-windows.ps1 -Installer
```

Detail proses publikasi, runtime Flet, isi paket, checksum, dan kanal pembaruan tersedia di [`docs/RELEASING.md`](docs/RELEASING.md).

### Struktur proyek

```text
app.py          Entrypoint dan antarmuka wizard Flet
src/            Katalog workflow, validator, dan generator dokumen
utils/          Utilitas file, gambar, PDF, estimasi, dan UI
template/       Template DOCX bawaan
input/          Template/formats input XLSX bawaan
tests/          Unit test dan regression test
scripts/        Skrip setup LibreOffice dan build Windows
installer/      Konfigurasi installer Windows opsional
updates/        Manifest kanal stable dan beta
docs/           Dokumentasi teknis dan panduan rilis
```

## Rilis, Publisher, dan Lisensi

- **Rilis:** <https://github.com/Mjulianfr001056/siomay-se26/releases>
- **Changelog:** [`CHANGELOG.md`](CHANGELOG.md)
- **Publisher:** 6304 - Muhammad Julian Firdaus, S.Tr.Stat.
- **Application ID:** `id.go.bps.siomay`

Repositori ini belum menyertakan berkas lisensi. Hak untuk menggunakan, memodifikasi, atau mendistribusikan kode tidak boleh diasumsikan sampai lisensi proyek ditetapkan.
