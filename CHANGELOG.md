# Changelog

Semua perubahan penting pada proyek SIOMAY didokumentasikan di berkas ini. Format changelog ini mengacu pada [Keep a Changelog](https://keepachangelog.com/id/1.0.0/).

---

## [v2026.1.1-beta.3] - 2026-08-31

### Otomatisasi & CI/CD
- **Sinkronisasi Otomatis Manifest Pembaruan**: Workflow rilis GitHub Actions (`.github/workflows/release.yml`) kini secara otomatis memperbarui file manifest (`updates/beta.json` atau `updates/stable.json`) dengan versi paket, tautan unduhan, dan checksum SHA-256 saat tag rilis baru dipublikasikan, lalu melakukan push pembaruan tersebut langsung ke branch `master`.
- **Pembersihan Encoding UTF-8 (No-BOM)**: Memastikan manifest JSON ditulis dalam format UTF-8 murni tanpa BOM agar kompatibel dengan seluruh parser JSON Python.
- **Pembaruan Panduan Rilis**: Memperbarui dokumentasi alur rilis pada `docs/RELEASING.md` agar mencerminkan otomatisasi manifest pembaruan aplikasi.

---

## [v2026.1.1-beta.2] - 2026-08-31

### Peningkatan UI/UX
- **Tombol Generate Responsif & Greyed-out**: Tombol *"Mulai Generate Dokumen"* di Langkah 4 secara visual dinonaktifkan (berubah warna menjadi abu-abu), label berubah menjadi *"Sedang Memproses Dokumen…"*, dan ikon berubah menjadi *hourglass* saat proses berlangsung guna mencegah klik ganda.
- **Siklus State Generator Terpusat**: Penambahan handler `_gen_ui_start()` dan `_gen_ui_finish()` untuk memastikan pemulihan status UI tetap berjalan mulus meskipun terjadi error selama proses pembuatan dokumen.

### Perbaikan Konversi PDF
- **Konversi PDF Headless & Senyap**: Menambahkan bendera `CREATE_NO_WINDOW` dan `SW_HIDE` pada pemanggilan subprocess LibreOffice agar tidak memunculkan popup jendela Command Prompt (CMD) hitam saat konversi DOCX ke PDF.
- **Prioritas Binary LibreOffice**: Memprioritaskan `soffice.exe` dibanding wrapper console `soffice.com` di seluruh jalur pencarian executable.

---

## [v2026.1.1-beta.1] - 2026-08-30

### Fitur Baru
- **Generator Bukti Terima Paket Internet**: Menambahkan dukungan pembuatan dokumen Bukti Terima Paket Internet dengan tata letak grid foto bukti 2x2 per halaman A4.
- **Validasi Bukti Dukung**: Pengecekan otomatis tautan dan ketersediaan berkas tangkapan layar/bukti pembelian kuota internet per petugas.

### Perbaikan & Penyesuaian
- **Pencegahan IndexError Tabel & Paragraf**: Penanganan defensif pada manipulasi paragraf dan tabel template dokumen Word (`python-docx`).
- **Dukungan Lingkungan Pengembangan**: Deteksi otomatis folder instalasi LibreOffice di root repositori untuk eksekusi langsung via `python app.py`.
- **Skrip Instalasi Lokal**: Penyediaan skrip PowerShell `scripts/install_lo_dev.ps1` untuk mempermudah pemasangan LibreOffice lokal bagi developer.
