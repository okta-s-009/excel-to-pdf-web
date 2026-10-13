@echo off
title Update Perubahan ke GitHub
cls
echo ======================================================================
echo             UPDATE OTOMATIS EXCEL2PDF STUDIO KE GITHUB
echo ======================================================================
echo.
echo [*] Menyiapkan seluruh perubahan berkas...
git add .

echo.
set /p COMMIT_MSG="Masukkan pesan update (kosongkan untuk default): "
if "%COMMIT_MSG%"=="" (
    set COMMIT_MSG=Update fitur dan perbaikan antarmuka Excel2PDF Studio
)

echo.
echo [*] Menyimpan commit: "%COMMIT_MSG%"...
git commit -m "%COMMIT_MSG%"

echo.
echo [*] Mengirim (push) pembaruan ke GitHub (branch main)...
git push origin main

echo.
if %ERRORLEVEL% EQU 0 (
    echo ======================================================================
    echo [BERHASIL] Repository GitHub Anda telah berhasil diperbarui!
    echo Jika terhubung ke Vercel, website Anda akan otomatis terupdate.
    echo ======================================================================
) else (
    echo ======================================================================
    echo [GAGAL] Terjadi kendala saat melakukan push ke GitHub.
    echo Silakan periksa koneksi internet Anda.
    echo ======================================================================
)
echo.
pause
