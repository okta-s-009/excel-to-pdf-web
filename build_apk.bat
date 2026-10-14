@echo off
chcp 65001 >nul
title Excel2PDF Studio - APK Builder

echo =======================================================
echo         EXCEL2PDF STUDIO - ANDROID APK BUILDER         
echo =======================================================
echo.

set "SCRIPT_DIR=%~dp0"
set "ANDROID_DIR=%SCRIPT_DIR%android"

:: 1. Sync latest web assets to Android project
echo [1/3] Menyinkronkan aset web terbaru ke project Android...
python "%SCRIPT_DIR%setup_android_assets.py"
echo.

:: 2. Check Java / JDK
echo [2/3] Memeriksa Java Development Kit (JDK)...
where java >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    :: Periksa lokasi JDK / JBR umum secara otomatis
    if exist "C:\Program Files\Android\Android Studio\jbr\bin\java.exe" (
        set "JAVA_HOME=C:\Program Files\Android\Android Studio\jbr"
        set "PATH=C:\Program Files\Android\Android Studio\jbr\bin;%PATH%"
    ) else if exist "C:\Program Files\Microsoft\jdk-17\bin\java.exe" (
        set "JAVA_HOME=C:\Program Files\Microsoft\jdk-17"
        set "PATH=C:\Program Files\Microsoft\jdk-17\bin;%PATH%"
    ) else if exist "C:\Program Files\Eclipse Adoptium\jdk-17\bin\java.exe" (
        set "JAVA_HOME=C:\Program Files\Eclipse Adoptium\jdk-17"
        set "PATH=C:\Program Files\Eclipse Adoptium\jdk-17\bin;%PATH%"
    ) else if exist "C:\Program Files\Java\jdk-17\bin\java.exe" (
        set "JAVA_HOME=C:\Program Files\Java\jdk-17"
        set "PATH=C:\Program Files\Java\jdk-17\bin;%PATH%"
    )
)
where java >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [!] PERINGATAN: Java / JDK belum terdeteksi di sistem PATH.
    echo.
    echo Anda memiliki 2 opsi mudah untuk membuat file APK:
    echo.
    echo   OPSI A (Rekomendasi Utama & Paling Mudah):
    echo   1. Buka aplikasi "Android Studio".
    echo   2. Pilih menu "Open" dan pilih folder:
    echo      "%ANDROID_DIR%"
    echo   3. Klik menu: Build ^> Build Bundle(s) / APK(s) ^> Build APK(s)
    echo   4. Selesai! File APK langsung jadi dalam hitungan detik.
    echo.
    echo   OPSI B (Install JDK 17 otomatis via winget):
    echo   Ketik perintah berikut di PowerShell / CMD Administrator:
    echo      winget install Microsoft.OpenJDK.17
    echo.
    pause
    exit /b 1
)

:: 3. Build APK with Gradle
echo [3/3] Mengompilasi APK Android via Gradle Wrapper...
cd /d "%ANDROID_DIR%"
call gradlew.bat assembleDebug

if %ERRORLEVEL% EQU 0 (
    echo.
    echo =======================================================
    echo [SUKSES] File APK berhasil dikompilasi!
    echo =======================================================
    set "SOURCE_APK=%ANDROID_DIR%\app\build\outputs\apk\debug\app-debug.apk"
    set "DEST_APK=%SCRIPT_DIR%Excel2PDF-Studio-v1.0.apk"
    if exist "%SOURCE_APK%" (
        copy /y "%SOURCE_APK%" "%DEST_APK%" >nul
        echo File installer APK siap digunakan:
        echo   ^> "%DEST_APK%"
        echo.
        explorer.exe /select,"%DEST_APK%"
    )
) else (
    echo.
    echo [!] Terjadi kendala saat proses compile Gradle.
    echo Silakan buka folder 'android' langsung di Android Studio untuk kompilasi visual.
)

pause
