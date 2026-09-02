# Changelog

Semua perubahan penting pada proyek SIOMAY didokumentasikan di berkas ini. Format changelog ini mengacu pada [Keep a Changelog](https://keepachangelog.com/id/1.0.0/).

---

## [v2026.1.5] - 2026-09-02

### Dokumen SPP Termin 1 dan Termin 2
- **Dukungan SPP Termin 2**: Menambahkan generator, template DOCX, dan berkas input Excel khusus SPP Termin 2 untuk petugas PPL maupun PML.
- **Pemisahan Alur SPP**: SPP yang sudah ada diperjelas sebagai SPP Termin 1, dengan pemilihan dokumen, validasi data, dan perutean generator yang terpisah untuk setiap termin.
- **Penyesuaian Urutan Dokumen**: Aset BAST dan Bukti Terima dipindahkan mengikuti penambahan SPP Termin 2 agar urutan template dan input tetap konsisten.

### Peningkatan Proses dan Validasi
- **Estimasi Konversi PDF**: Menampilkan estimasi waktu maksimum dan perkiraan sisa waktu berdasarkan jenis dokumen serta jumlah berkas yang akan dikonversi.
- **Validasi Template DOCX**: Memeriksa kelengkapan dan kesesuaian placeholder template sebelum dokumen dibuat sehingga kesalahan template dapat dilaporkan lebih awal dan lebih jelas.
- **Nomor Urut Tiga Digit**: Nomor urut BAPP Termin 2 dan BAST dari Excel kini dipertahankan dalam format tiga digit, misalnya `001` dan `021`, tanpa mengubah nomor alfanumerik.

### Distribusi dan Lingkungan Pengembangan
- **Instalasi LibreOffice Andal**: Menggunakan arsip resmi LibreOffice `25.8.7.2` yang immutable serta menambahkan pemeriksaan kegagalan HTTP, ukuran minimum berkas, dan signature MSI sebelum instalasi.
- Menambahkan dan memperbarui pengujian regresi untuk SPP Termin 1/2, perutean alur kerja, validasi template, estimasi konversi, dan format nomor urut.

---

## [v2026.1.4] - 2026-09-01

### Dukungan Gambar
- **Dukungan HEIC/HEIF**: Foto HEIC/HEIF dari Google Drive kini didekode dan dikonversi otomatis menjadi PNG sebelum disisipkan ke dokumen Word.
- **Orientasi Foto**: Metadata EXIF orientation diterapkan agar foto dari perangkat seluler tidak tampil terbalik atau menyamping.
- **Validasi Unduhan**: Respons HTML Google Drive, file kosong, dan format gambar yang tidak valid kini menghasilkan peringatan yang lebih jelas tanpa menggagalkan keseluruhan batch.

### Pengujian dan Distribusi
- Menambahkan pengujian regresi decoding HEIC, kompatibilitas JPEG/PNG, penyisipan gambar ke DOCX, dan penolakan respons HTML.
- Memastikan codec native `pillow-heif` tersedia di paket portable Windows sebelum aset rilis dibuat.

---

## [v2026.1.3] - 2026-09-01

### Performa Konversi dan Keluaran
- **Konversi PDF Batch**: Mengonversi seluruh dokumen DOCX melalui satu proses LibreOffice, sehingga biaya startup LibreOffice tidak lagi berulang untuk setiap dokumen.
- **Keluaran DOCX Tanpa Konversi**: Menambahkan pilihan arsip ZIP berisi DOCX asli agar hasil generator dapat langsung disimpan tanpa menunggu konversi PDF.
- **Penanganan Kegagalan Parsial**: Dokumen yang berhasil dikonversi tetap disimpan dan kegagalan individual tetap dilaporkan tanpa mengubah urutan hasil.

### Pengujian
- Menambahkan pengujian konversi batch yang memastikan beberapa DOCX hanya menjalankan satu subprocess LibreOffice serta memverifikasi urutan hasil dan kegagalan parsial.

---

## [v2026.1.2.1-beta.1] - 2026-09-01

### Fitur Baru & Peningkatan UX
- **Live Counter & Rekap Durasi Langkah 5**:
  - Penambahan timer aktif (*stopwatch*) berformat `mm:ss` saat proses konversi DOCX ke PDF, pengemasan ZIP, dan penggabungan PDF.
  - Kartu rekap durasi interaktif beraksen ungu (*lilac*) yang merangkum rincian format keluaran, jumlah berkas, dan waktu proses simpan secara presisi.
- **Dukungan Pembukaan Tautan Rilis yang Andal**:
  - Implementasi fungsi `open_external_url` dengan strategi fallback bertingkat (`webbrowser.open` → `page.launch_url` → `os.startfile`) untuk mengatasi isu tombol tautan rilis tidak merespons di lingkungan desktop Windows.
  - Penutupan otomatis dialog pembaruan/tentang saat pengguna membuka tautan GitHub Releases di peramban web.

### Perbaikan Bug & Stabilitas
- **Reset State Alur Kerja**: Pembersihan durasi dan komponen kartu ringkasan saat pengguna memilih jenis dokumen baru atau mengulang alur pembuatan dokumen (`restart_workflow`).
- **Penguatan Unit Test Suite**: Penambahan pengujian otomatis untuk helper durasi waktu (`format_duration`, `format_timer_clock`), komponen kartu UI, dan pengujian fallback pembuka URL browser.

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
