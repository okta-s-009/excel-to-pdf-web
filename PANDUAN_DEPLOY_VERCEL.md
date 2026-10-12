# ☁️ Panduan Deploy Excel2PDF Studio ke Vercel (Online 24/7)

Dengan menaruh aplikasi ini di **Vercel Cloud**, Anda dapat membuka dan menggunakannya kapan saja dari **HP, Tablet, maupun Komputer** tanpa perlu menyalakan laptop Anda terus-menerus.

---

## 🌟 Keuntungan Deploy di Vercel:
1. **Laptop Bisa Dimatikan:** Sistem berjalan di cloud Vercel secara mandiri 24 jam nonstop.
2. **Akses dari Mana Saja:** Dapat dibuka lewat link publik HTTPS (misal: `https://excel2pdf-studio.vercel.app`) dari jaringan apa pun (tanpa harus 1 Wi-Fi).
3. **PWA Layar Utama di HP:** Bisa langsung di-*"Tambahkan ke Layar Utama"* di HP sehingga menjadi aplikasi mobile instan tanpa APK.
4. **100% Gratis:** Paket gratis (Hobby Plan) dari Vercel sudah lebih dari cukup untuk penggunaan sehari-hari.

---

## 🚀 Cara 1: Menggunakan Script Otomatis `deploy_vercel.bat` (Paling Cepat)

Jika Anda ingin deploy langsung dari laptop dalam 1 menit:

1. Klik ganda (double-click) file:
   ```
   deploy_vercel.bat
   ```
2. Terminal akan menjalankan `npx vercel`:
   - Jika Anda belum login ke Vercel, terminal akan meminta konfirmasi login via browser. Buka link tersebut dan klik **Login with GitHub / Google / Email**.
3. Di terminal, jawab beberapa pertanyaan sederhana (cukup tekan **Enter** untuk memilih default):
   - *Set up and deploy?* Ketik `y` lalu Enter.
   - *Which scope?* Pilih akun Vercel Anda.
   - *Link to existing project?* Ketik `n` (No).
   - *What's your project's name?* Beri nama, misal `excel2pdf-studio`.
   - *In which directory is your code located?* Tekan Enter (`./`).
   - *Want to modify settings?* Ketik `n` lalu Enter.
4. Vercel akan otomatis mengunggah dan memproses proyek.
5. Selesai! Anda akan mendapatkan link publik seperti:
   ```
   https://excel2pdf-studio.vercel.app
   ```
6. Buka link tersebut di HP atau Tablet Anda!

---

## 🐙 Cara 2: Menggunakan GitHub + Vercel (Rekomendasi Terbaik)

Jika Anda ingin proyek Anda tersimpan rapi di GitHub dan otomatis update setiap ada perubahan:

### Langkah 1: Unggah ke GitHub
1. Buka [GitHub](https://github.com) dan buat repository baru (misal: `excel-to-pdf-web`).
2. Di folder proyek ini, jalankan perintah Git di terminal / command prompt:
   ```bash
   git init
   git add .
   git commit -m "Excel2PDF Studio Web Ready for Vercel"
   git branch -M main
   git remote add origin https://github.com/USERNAME_ANDA/excel-to-pdf-web.git
   git push -u origin main
   ```

### Langkah 2: Hubungkan ke Vercel
1. Buka website [vercel.com](https://vercel.com) dan login (bisa login menggunakan akun GitHub).
2. Di Dashboard Vercel, klik tombol **"Add New..."** > pilih **"Project"**.
3. Cari repository `excel-to-pdf-web` yang baru saja Anda buat, lalu klik **"Import"**.
4. Biarkan semua pengaturan default (Vercel sudah membaca konfigurasi `vercel.json` dan `requirements.txt` yang telah kami siapkan).
5. Klik tombol **"Deploy"**.
6. Tunggu sekitar 30 detik hingga muncul pesan selamat. Tautan website Anda siap digunakan!

---

## 📱 Membuka dan Menggunakan di HP / Tablet

Setelah link Vercel Anda aktif (misal `https://excel2pdf-studio.vercel.app`):
1. Buka browser di HP / Tablet (Google Chrome atau Safari).
2. Masukkan alamat URL Vercel Anda.
3. Klik tombol **"Pasang App"** atau menu browser **"Tambahkan ke Layar Utama" (Add to Home screen)**.
4. Ikon aplikasi **Excel2PDF Studio** kini terpasang di layar HP Anda dan siap digunakan kapan pun di mana pun, **meskipun laptop Anda dalam kondisi mati total**.
