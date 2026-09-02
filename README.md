# SIOMAY
## Sistem Otomasi Massal dan Terpercaya

**SIOMAY** adalah aplikasi desktop Windows untuk membantu otomatisasi pembuatan dokumen administrasi SE2026 secara massal dari data Microsoft Excel dan template Microsoft Word.

Aplikasi ini dirancang untuk mempercepat proses pembuatan dokumen, mengurangi pekerjaan berulang, dan menjaga konsistensi hasil dokumen.

> Status: **Stable**
> Versi saat ini: **v2026.1.4**

## Fitur

SIOMAY mendukung pembuatan dokumen berikut:

- Lampiran SPK PPL
- Lampiran SPK PML
- BAPP PPL Termin 1
- BAPP PML Termin 1
- SPP PPL Termin 1
- SPP PML Termin 1
- BAPP PPL Termin 2
- BAPP PML Termin 2
- SPP PPL Termin 2
- SPP PML Termin 2
- BAST PPL
- BAST PML

Kemampuan utama:

- Memilih jenis dokumen melalui alur kerja bertahap.
- Menggunakan template dokumen bawaan atau template `.docx` dari pengguna.
- Mengunduh format/template data Excel yang diperlukan.
- Memvalidasi struktur data Excel sebelum proses generate.
- Menghasilkan dokumen `.docx` secara massal dari setiap baris data.
- Mengonversi dokumen ke PDF dengan LibreOffice yang dibundel di rilis Windows.
- Menggabungkan PDF dan/atau membuat arsip ZIP hasil dokumen.
- Membuka lokasi hasil dokumen langsung melalui Windows Explorer.

## Persyaratan Sistem

| Komponen | Persyaratan |
|---|---|
| Sistem operasi | Windows 10 64-bit versi 1809 atau lebih baru; Windows 11 didukung |
| Arsitektur | x64 |
| Microsoft Excel | Disarankan untuk mengisi atau memeriksa file input `.xlsx` |
| Microsoft Word desktop | Tidak diperlukan untuk konversi DOCX ke PDF |
| Koneksi internet | Diperlukan untuk memeriksa dan mengunduh pembaruan aplikasi |
| Python | Tidak perlu diinstal oleh pengguna akhir |

> Rilis Windows menyertakan LibreOffice untuk konversi PDF. Jangan menghapus folder `LibreOffice` yang berada di samping `SIOMAY.exe` setelah mengekstrak ZIP.

## Instalasi

1. Buka halaman [Releases](https://github.com/Mjulianfr001056/siomay-se26/releases).
2. Unduh arsip `SIOMAY-<tag>-windows.zip` dari rilis yang sesuai.
3. Ekstrak seluruh isi ZIP ke folder yang dapat ditulis, misalnya `Documents\SIOMAY`.
4. Jalankan `SIOMAY\SIOMAY.exe`.

Gunakan tag tanpa tanda hubung (misalnya `v2026.1.1`) untuk rilis stabil. Tag yang memiliki suffix setelah tanda hubung (misalnya `v2026.1.2-beta.1`) adalah rilis beta/prerelease.

Jika Windows menampilkan peringatan keamanan, pastikan ZIP diperoleh langsung dari halaman Releases resmi proyek ini dan periksa nilai SHA-256 yang disertakan pada rilis.

## Cara Menggunakan

1. **Pilih Dokumen**  
   Pilih jenis dokumen administrasi yang ingin dibuat.

2. **Template Dokumen**  
   Gunakan template bawaan atau unggah template Microsoft Word (`.docx`) yang sesuai.

3. **Upload Data**  
   Unduh format Excel bila diperlukan, isi data sesuai kolom yang tersedia, lalu unggah file `.xlsx`.

4. **Generate**  
   Periksa hasil validasi data dan mulai proses pembuatan dokumen.

5. **Simpan & Selesai**  
   Simpan hasil dalam format DOCX, PDF, PDF gabungan, atau ZIP sesuai opsi yang tersedia.

## Pembaruan Aplikasi

SIOMAY menyediakan pemeriksaan pembaruan untuk memperoleh perbaikan dan fitur baru.

- Pembaruan patch dapat memperbaiki masalah tanpa mengubah alur utama aplikasi.
- Pembaruan versi dapat menambah jenis dokumen, template, validasi, atau fitur baru.
- Saat pembaruan tersedia, SIOMAY akan menampilkan informasi versi dan catatan perubahan.
- Saat pembaruan tersedia, pengguna diarahkan ke halaman GitHub Releases resmi untuk mengunduh ZIP secara manual.

Build beta memeriksa manifest kanal **beta**, sedangkan build stabil memeriksa kanal **stable**.

## Umpan Balik dan Pelaporan Masalah

Masukan pengguna tetap penting untuk membantu peningkatan kualitas SIOMAY.

Saat melaporkan masalah, sertakan:

- Versi aplikasi yang digunakan.
- Jenis dokumen yang diproses.
- Langkah-langkah sebelum masalah terjadi.
- Pesan kesalahan atau tangkapan layar, bila aman untuk dibagikan.
- Informasi apakah folder `LibreOffice` masih berada di samping `SIOMAY.exe`.

**Jangan mengirim file Excel, dokumen hasil, atau tangkapan layar yang mengandung data pribadi/sensitif melalui kanal publik.**

Saluran umpan balik dicantumkan pada halaman rilis aplikasi.

## Privasi Data

SIOMAY memproses file yang dipilih pengguna secara lokal di komputer pengguna.

- Data Excel dan dokumen tidak diunggah oleh aplikasi sebagai bagian dari proses generate.
- Hasil dokumen disimpan di lokasi yang dipilih pengguna.
- Pemeriksaan pembaruan hanya mengambil informasi versi rilis.
- Pengiriman log diagnostik atau umpan balik, bila tersedia, harus dilakukan dengan persetujuan pengguna.

## Untuk Pengembang

### Teknologi

- Python
- [Flet](https://flet.dev/)
- pandas
- openpyxl
- python-docx
- pypdf
- LibreOffice (runtime native yang dibundel saat rilis Windows)

### Menjalankan dari kode sumber

1. Instal Python 3.14.
2. Buat dan aktifkan virtual environment.
3. Instal dependensi:

   ```powershell
   py -3.14 -m pip install .
   ```

4. Jalankan aplikasi:

   ```powershell
   python app.py
   ```

### Struktur penting proyek

```text
app.py       Entrypoint aplikasi Flet
src/         Logika workflow, validasi, dan generator dokumen
utils/       Utilitas UI, file, dan PDF
template/    Template DOCX bawaan
input/       Template Excel bawaan
```

## Rilis

Rilis aplikasi tersedia di:

<https://github.com/Mjulianfr001056/siomay-se26/releases>

Versi awal pilot:

- `v2026.1-beta.1`

## Publisher

**6304 - Muhammad Julian Firdaus, S.Tr.Stat.**

## Lisensi

Lisensi proyek akan ditentukan.
