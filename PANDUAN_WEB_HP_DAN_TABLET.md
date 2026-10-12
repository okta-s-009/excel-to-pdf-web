# 📱 Panduan Menggunakan Excel2PDF Studio di Web, HP & Tablet

Aplikasi **Excel2PDF Studio** kini sepenuhnya berjalan berbasis **Web Modern (Responsive Web App & PWA)**. Anda tidak perlu lagi repot merakit atau menginstal file APK Android. Aplikasi ini dapat diakses langsung dari browser di Komputer, Laptop, HP Android, iPhone, iPad, maupun Tablet.

---

## 🚀 Cara Menjalankan Aplikasi

### 1. Di Komputer / Laptop (Host Server)
1. Cukup klik ganda (double-click) file:
   ```
   run.bat
   ```
2. Jendela terminal akan terbuka dan menampilkan informasi alamat akses:
   - **Akses Komputer:** `http://localhost:8000`
   - **Akses HP / Tablet:** `http://192.168.xxx.xxx:8000` (contoh: `http://192.168.100.246:8000`)
3. Browser komputer akan otomatis terbuka menampilkan aplikasi.

> **Catatan Windows Firewall:**  
> Jika saat pertama kali dijalankan muncul jendela Windows Defender Firewall, centang kedua pilihan (Private & Public) lalu klik **Allow access** (Izinkan akses) agar HP/Tablet Anda dapat terhubung.

---

## 📲 Cara Membuka dari HP atau Tablet

### Metode A: Pindai QR Code (Paling Cepat & Praktis)
1. Pastikan **HP / Tablet** terhubung ke **jaringan Wi-Fi / Hotspot yang sama** dengan komputer Anda.
2. Di layar browser komputer, klik tombol hijau **"📱 Buka di HP / Tab"** di bagian atas (header).
3. Arahkan kamera HP Anda ke layar monitor untuk **memindai QR Code** yang muncul.
4. Ketuk tautan yang muncul di kamera HP Anda — aplikasi langsung terbuka!

---

### Metode B: Ketik Alamat IP di Browser HP
1. Pastikan HP terhubung ke Wi-Fi yang sama dengan komputer.
2. Buka browser di HP (Google Chrome, Safari, atau browser lainnya).
3. Ketik alamat IP lokal komputer Anda yang tertera di terminal atau di jendela QR Code, contoh:
   ```
   http://192.168.100.246:8000
   ```
4. Tekan Masuk / Enter.

---

## ⭐ Cara Memasang sebagai Aplikasi di HP (PWA Tanpa APK)

Aplikasi ini sudah dilengkapi teknologi **Progressive Web App (PWA)**, sehingga bisa dipasang di Layar Utama HP layaknya aplikasi bawaan tanpa memakan memori:

### Untuk HP Android (Google Chrome):
1. Buka website di Google Chrome HP.
2. Klik tombol **"📥 Pasang App"** yang muncul di bagian atas aplikasi, **atau**:
3. Ketuk ikon **titik tiga (⋮)** di pojok kanan atas Chrome.
4. Pilih **"Tambahkan ke Layar Utama"** *(Add to Home screen)* atau **"Pasang Aplikasi"**.
5. Ikon **Excel2PDF Studio** akan langsung muncul di halaman menu/layar utama HP Anda!

### Untuk iPhone / iPad (Apple Safari):
1. Buka website di Safari.
2. Ketuk tombol **Bagikan (Share)** (ikon kotak dengan panah ke atas di bagian bawah).
3. Gulir ke bawah dan pilih **"Tambahkan ke Layar Utama"** *(Add to Home Screen)*.
4. Ketuk **Tambah** *(Add)*. Aplikasi siap dibuka kapan saja secara layar penuh!

---

## ✨ Fitur & Keunggulan di HP & Tablet

- 📁 **Unggah File Mudah:** Ketuk area unggah untuk memilih file Excel dari memori HP, Google Drive, WhatsApp Document, atau File Manager.
- 📐 **Kontrol Margin Presisi:** Atur margin Kiri, Kanan, Atas, Bawah, Header, dan Footer dalam satuan **cm** atau **inci** dengan keyboard angka yang nyaman.
- 📄 **Tata Letak Lengkap:** Pilihan orientasi Portrait/Landscape, ukuran kertas (A4, Letter, Legal, A3), serta penskalaan tabel agar muat dalam 1 halaman lebar (*Fit Width*).
- 👁️ **Simulasi Margin & Pratinjau Real-Time:** Melihat pratinjau halaman dan batas margin yang responsif sesuai ukuran layar HP Anda.
- 📤 **Langsung Bagikan:** Setelah konversi selesai, ketuk tombol **"Bagikan"** untuk langsung mengirimkan file PDF ke WhatsApp, Email, Telegram, atau menyimpannya ke Google Drive.
- ⚡ **Dual Engine:** Mendukung konversi presisi via Microsoft Excel COM di server PC, serta didukung engine offline bawaan jika dibutuhkan.
