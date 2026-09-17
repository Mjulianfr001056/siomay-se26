# Layanan Folder Google Drive

Worker di `worker/` menyimpan `GOOGLE_DRIVE_API_KEY` sebagai secret Cloudflare
dan menyediakan hanya `GET /v1/drive/folders/{folderId}/images`. Aplikasi desktop
menggunakan hasil metadata tersebut, lalu mengunduh file publik langsung dari
Google Drive. Workbook, dokumen, dan isi gambar tidak melewati Worker.

## Bootstrap Google Cloud

1. Buat/pilih project khusus, aktifkan **Google Drive API v3**, lalu buat API key.
2. Pada **API restrictions**, pilih **Restrict key** dan izinkan hanya Google
   Drive API. Jangan memakai application restriction berbasis IP: egress
   Cloudflare tidak memiliki satu IP statis. Kuota, alert, dan billing alert
   tetap harus dikonfigurasi di Google Cloud.
3. Folder yang akan digunakan harus berakses **Anyone with the link**. Worker
   tidak memperoleh akses ke folder privat.

## Bootstrap Cloudflare

Prasyarat: Node.js 22+ dan akses ke account Cloudflare tujuan.

```powershell
cd worker
npm ci
npx wrangler login
npx wrangler deploy
npx wrangler secret put GOOGLE_DRIVE_API_KEY
npx wrangler secret list
```

Secret diunggah langsung dari terminal, tidak disimpan di `.env`, `.dev.vars`,
GitHub, atau source code. Deployment Wrangler berikutnya mempertahankan secret
yang sudah ada. Untuk pengembangan Worker lokal saja, `worker/.dev.vars` boleh
berisi key uji; file tersebut diabaikan Git dan tidak boleh didistribusikan.

Setelah deployment pertama, salin `.env.example` menjadi `.env`, lalu isi
`SIOMAY_DRIVE_FOLDER_WORKER_URL` dengan hostname `workers.dev` yang ditampilkan
Wrangler atau hostname custom production. Aplikasi mengutamakan environment
variable proses dengan nama yang sama, kemudian membaca `.env`. Nilai ini bukan
secret, tetapi harus berupa URL HTTPS tanpa query atau fragment.

Jika memakai custom domain, tambahkan domain melalui Workers **Settings >
Domains & Routes**, aktifkan TLS, kemudian gunakan hostname tersebut dalam `.env`
desktop. Jangan menambahkan proxy publik lain yang mencatat query Google.

## CI/CD GitHub

Workflow `deploy-worker.yml` memvalidasi pull request tanpa secret. Deployment
hanya berjalan setelah push ke `master` dan persetujuan environment GitHub
`production-worker`. Konfigurasikan environment tersebut dengan:

- `CLOUDFLARE_ACCOUNT_ID`
- `CLOUDFLARE_API_TOKEN` — token account-scoped dengan izin minimum **Workers
  Scripts: Edit** (tambahkan **Workers Routes: Edit** hanya bila workflow kelak
  mengelola route).

Untuk workflow rilis Windows, konfigurasikan repository/environment variable
non-rahasia `SIOMAY_DRIVE_FOLDER_WORKER_URL`. Workflow membuat `.env` khusus
rilis dari variable tersebut dan menyertakannya di samping executable.

Jangan tambahkan Google key ke GitHub. Pastikan branch protection mewajibkan job
validasi dan environment protection mewajibkan reviewer deployment.

## Verifikasi dan operasi

Smoke test dengan ID folder publik khusus pengujian:

```powershell
curl.exe -i "https://HOST/v1/drive/folders/FOLDER_ID/images"
```

Respons sukses adalah JSON `{"files":[...]}` tanpa API key. Verifikasi juga ID
invalid (400), folder privat/tidak ada (404/502 tergantung respons Google), method
POST (405), dan burst di atas 60 request/menit (429). Jangan memakai folder berisi
data pribadi untuk smoke test.

Pantau Workers Logs/Traces untuk status 429 dan 5xx, Workers Analytics untuk
latency/error, serta Google Cloud API metrics untuk kuota dan penolakan key.
Rate limiter bersifat per lokasi Cloudflare dan eventually consistent; ini
perlindungan penyalahgunaan, bukan sistem accounting presisi.

### Rotasi

1. Buat key Google baru dengan pembatasan Drive API yang sama.
2. Jalankan `npx wrangler secret put GOOGLE_DRIVE_API_KEY`, lalu smoke-test.
3. Nonaktifkan dan hapus key lama setelah observasi berhasil.
4. Rotasi token GitHub secara terpisah dengan membuat token Cloudflare baru,
   mengganti secret environment, menjalankan deployment, lalu mencabut token lama.

### Rollback dan insiden

- Rollback kode melalui **Workers > Deployments** ke version terakhir yang sehat,
  atau revert commit dan jalankan ulang workflow. Secret yang ada tidak perlu
  diunggah ulang.
- Jika key Google bocor, nonaktifkan key segera, buat pengganti, unggah dengan
  `wrangler secret put`, dan periksa metric/audit log Google.
- Jika terjadi abuse, turunkan limit binding, blokir sumber melalui Cloudflare
  WAF bila tersedia, atau nonaktifkan route sementara. Jangan menampilkan body
  error Google kepada klien; Worker sengaja mengembalikan pesan tersanitasi.