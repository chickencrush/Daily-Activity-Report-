# DAILY REPORT & ATTENDANCE SYSTEM 🐔📊
**Sistem Presensi Biometrik Multi-Gerai, Laporan Harian (Kanban), & Analisis KPI Berbasis Progressive Web App (PWA)**

Aplikasi Web App internal perusahaan berbasis **PWA (Progressive Web App)** dengan arsitektur **Serverless**:
- **Database & Penyimpanan**: Google Sheets & Google Drive via Google Apps Script (GAS) Web App Endpoint.
- **Biometrik & Keamanan**: Kamera Langsung (Strict Live-Only, tanpa unggah berkas galeri) + Formula Haversine Geo-fencing multi-titik gerai.
- **Produktivitas & Kinerja**: Papan Kanban Laporan Harian (Rencana Kerja, Sedang Dikerjakan, Tuntas Selesai), Analisis KPI 4 Pilar, Laporan Presensi Per Divisi, dan Kutipan Motivasi Harian.

---

## 📁 Struktur File Repositori

```text
├── index.html            # Frontend Utama (Single-File PWA, Tailwind CSS, Lucide Icons, Engine)
├── Code.gs               # Backend Script Google Apps Script (REST API Endpoint)
├── manifest.json         # Konfigurasi PWA (Web App Manifest agar bisa di-install)
├── sw.js                 # Service Worker (Offline Cache & Kecepatan Akses PWA)
├── chicken_mascot.png    # Aset Logo & Animasi Memuat (Maskot Ayam Penyelidik)
├── README.md             # Dokumentasi Sistem & Panduan Pemasangan
└── .gitignore            # File yang diabaikan Git
```

---

## 🚀 Panduan Pemasangan (Setup Guide)

### 1. Setup Backend Google Sheets & Apps Script
1. Buat Google Spreadsheet baru di Google Drive Anda.
2. Buka menu **Extensions** > **Apps Script**.
3. Salin dan tempelkan seluruh isi berkas `Code.gs` ke editor Google Apps Script.
4. Simpan (`Ctrl + S`), lalu pilih fungsi `initDatabase` dan klik **Run** sekali untuk membuat seluruh tab tabel secara otomatis.
5. Klik **Deploy** > **New deployment** > Pilih tipe **Web app**:
   - **Execute as**: *Me (email akun Anda)*
   - **Who has access**: *Anyone*
6. Salin URL Web App yang dihasilkan.

### 2. Jalankan Frontend PWA
1. Buka berkas `index.html` langsung di peramban (Google Chrome / Microsoft Edge / Safari).
2. Hubungkan URL Web App Apps Script pada formulir pengaturan atau konstanta API.
3. Untuk meng-host secara gratis:
   - Aktifkan **GitHub Pages** pada repositori ini (Settings > Pages > Deploy from branch `main` / root).
   - Atau deploy ke **Vercel / Netlify / Cloudflare Pages** hanya dengan menghubungkan repositori GitHub ini.

---

## 📱 Fitur-Fitur Utama
1. **Presensi Biometrik Swafoto**: Kamera langsung tanpa unggah berkas dengan enkripsi Base64 ke Google Drive.
2. **Multi-Gerai Geo-Fencing**: Otomatis mendeteksi gerai terdekat dan mengunci tombol presensi jika berada di luar radius $\le 100\text{m}$.
3. **Laporan Presensi Per Divisi**: Filter mendalam per divisi (HRD, BM, Finance, QC, DOS, Logistik, Produksi, Driver, SM) beserta ekspor data ke format Excel.
4. **Analisis KPI 4 Pilar**: Evaluasi otomatis berdasarkan Kehadiran (30%), Laporan Kerja (30%), Target Divisi (30%), dan Disiplin (10%).
5. **Kutipan Motivasi Harian**: Tampilan motivasi kerja harian yang berganti secara otomatis.
6. **Multi-Platform**: Dapat dipasang (*Add to Home Screen*) di Android, iPhone (iOS), dan Desktop PC.
