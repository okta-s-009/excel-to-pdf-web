# Panduan Lengkap Pembuatan Installer APK Android (Excel2PDF Studio)

Project Android untuk aplikasi **Excel2PDF Studio** telah berhasil dibuat dan dikonfigurasi. Aplikasi ini dirancang untuk berjalan secara **100% Standalone & Offline** langsung di perangkat HP Android tanpa memerlukan koneksi internet atau laptop/komputer menyala.

---

## 📱 Fitur Aplikasi Android
- **100% Offline Engine**: Mengonversi file Excel (.xlsx, .xls, .csv) ke PDF langsung di dalam perangkat Android.
- **Kontrol Margin Lengkap**: Mengatur margin Atas, Bawah, Kiri, Kanan (cm & inci), orientasi (Portrait/Landscape), dan ukuran kertas (A4, Letter, Legal, A3).
- **Native Android File Picker**: Membuka file Excel langsung dari penyimpanan internal HP, Google Drive, atau WhatsApp.
- **Simpan & Bagikan Instan**:
  - Tombol **Unduh**: Menyimpan PDF ke folder `Download/Excel2PDF` di HP.
  - Tombol **Bagikan**: Membuka share sheet Android untuk langsung mengirim PDF ke WhatsApp, Telegram, Gmail, atau mencetak ke printer.

---

## 🛠️ Struktur Project Android

Folder project Android berada di:
`d:\ASISTEN VIRTUAL\EXCEL TO PDF\android\`

```
android/
├── app/
│   ├── build.gradle                 <- Konfigurasi SDK & dependencies
│   └── src/main/
│       ├── AndroidManifest.xml      <- Izin storage & konfigurasi app
│       ├── java/com/excel2pdf/studio/
│       │   ├── MainActivity.java    <- Native WebView & SAF File Picker
│       │   └── AndroidBridge.java   <- Bridge penyimpanan & share PDF
│       ├── res/                     <- Ikon app, tema gelap, & layout
│       └── assets/www/              <- Seluruh aset web offline (HTML, CSS, JS, libs)
├── build.gradle                     <- Gradle root
├── settings.gradle
├── gradlew.bat                      <- Gradle wrapper untuk Windows
└── gradle/wrapper/
```

---

## 🚀 3 Cara Membuat File APK

### Cara 1: Menggunakan Android Studio (Paling Mudah & Visual)
1. Buka aplikasi **Android Studio** di komputer Anda.
2. Di layar utama, pilih **Open** (atau klik menu **File > Open**).
3. Arahkan dan pilih folder:
   ```
   d:\ASISTEN VIRTUAL\EXCEL TO PDF\android
   ```
4. Tunggu beberapa saat sampai Android Studio selesai melakukan sinkronisasi Gradle (indikator di pojok bawah selesai).
5. Pada menu atas, klik:
   **Build** > **Build Bundle(s) / APK(s)** > **Build APK(s)**
6. Setelah selesai, akan muncul balon notifikasi di pojok kanan bawah:
   *`APK(s) generated successfully for 1 module`*
7. Klik tautan bertuliskan **locate** di dalam notifikasi tersebut. File installer `app-debug.apk` akan langsung terbuka di Windows Explorer!

---

### Cara 2: Satu Klik via Terminal / File `build_apk.bat`
Jika komputer Anda sudah terpasang Java / JDK 17:
1. Cukup klik ganda (double-click) file:
   [build_apk.bat](file:///d:/ASISTEN%20VIRTUAL/EXCEL%20TO%20PDF/build_apk.bat)
2. Script akan secara otomatis:
   - Menyinkronkan aset HTML/JS terbaru ke folder Android.
   - Mengompilasi file APK melalui Gradle.
   - Menyalin file hasil ke direktori utama sebagai:
     `d:\ASISTEN VIRTUAL\EXCEL TO PDF\Excel2PDF-Studio-v1.0.apk`
   - Membuka folder lokasi file APK secara otomatis.

> **Catatan jika Java belum terpasang:**
> Anda dapat memasang OpenJDK 17 dengan cepat lewat PowerShell Administrator:
> ```powershell
> winget install Microsoft.OpenJDK.17
> ```

---

### Cara 3: Langsung Pasang di HP Tanpa Compile (Mode PWA WebAPK)
Aplikasi ini sudah dilengkapi dengan dukungan **Progressive Web App (PWA)**:
1. Jalankan server lokal di komputer dengan double-click [run.bat](file:///d:/ASISTEN%20VIRTUAL/EXCEL%20TO%20PDF/run.bat).
2. Hubungkan HP Android ke jaringan WiFi yang sama dengan laptop.
3. Buka browser **Google Chrome** di HP Android, ketik alamat IP laptop Anda (contoh: `http://192.168.1.5:8000`).
4. Tekan tombol menu Chrome (ikon titik tiga di pojok kanan atas) > pilih **"Instal aplikasi"** atau **"Tambahkan ke Layar Utama"**.
5. Aplikasi akan langsung terpasang di HP Anda dengan ikon dan pengalaman layar penuh tanpa bilah browser.

---

## 📲 Cara Memasang (Install) APK di HP Android

1. Pindahkan file `.apk` yang sudah dibuat ke HP Anda (bisa lewat kabel data USB, kirim via WhatsApp Web ke chat sendiri, atau unggah ke Google Drive).
2. Di HP Android, buka file `.apk` tersebut melalui aplikasi **File Manager** / **Unduhan**.
3. Jika Android menampilkan peringatan keamanan:
   - Ketuk **Setelan (Settings)**.
   - Aktifkan opsi **"Izinkan dari sumber ini" (Allow from this source)**.
4. Ketuk **Instal (Install)** dan tunggu beberapa detik.
5. Selesai! Buka aplikasi **Excel2PDF Studio** dari daftar aplikasi HP Anda.
