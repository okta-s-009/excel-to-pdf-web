@echo off
title Upload Proyek ke GitHub
cls
echo ======================================================================
echo                 UPLOAD EXCEL2PDF STUDIO KE GITHUB
echo ======================================================================
echo.
echo Langkah awal:
echo 1. Buka browser dan login ke https://github.com
echo 2. Buat repository baru di https://github.com/new (misal: excel-to-pdf)
echo 3. Salin (copy) link repository HTTPS yang diberikan GitHub.
echo.
set /p REPO_URL="Paste / Tempelkan Link Repository GitHub di sini: "

if "%REPO_URL%"=="" (
    echo.
    echo [ERROR] Link repository tidak boleh kosong! Silakan jalankan ulang file ini.
    echo.
    pause
    exit /b
)

echo.
echo [*] Memulai proses upload ke GitHub...
echo.

echo [1/5] Inisialisasi Git...
git init

echo [2/5] Menyiapkan seluruh file...
git add .

echo [3/5] Menyimpan perubahan (commit)...
git commit -m "Excel2PDF Studio Ready for Vercel"

echo [4/5] Mengatur branch ke 'main'...
git branch -M main
git remote remove origin 2>nul
git remote add origin %REPO_URL%

echo [5/5] Mengunggah (push) ke GitHub...
git push -u origin main

echo.
echo ======================================================================
echo SELESAI! Seluruh file berhasil diunggah ke GitHub.
echo.
echo Langkah selanjutnya untuk Vercel:
echo 1. Buka https://vercel.com
echo 2. Klik "Add New..." > "Project"
echo 3. Pilih repository GitHub Anda tadi, lalu klik "Deploy"!
echo ======================================================================
echo.
pause
