@echo off
setlocal enabledelayedexpansion

:: =======================================================
:: SETUP AWAL — Clone SPT ke Server (JALANKAN SEKALI SAJA)
:: Repository  : https://github.com/adhito-nugroho/spt.git
:: =======================================================

:: --- SSH ke server ---
set SERVER_USER=adit
set SERVER_IP=127.0.0.1
set SERVER_PORT=2222

:: --- Tempat clone di server ---
set "REMOTE_PARENT=C:\laragon\www"
set "REMOTE_DIR=C:\laragon\www\spt-php"
set BRANCH=main

:: --- Path binary di server ---
set "REMOTE_GIT=C:\laragon\bin\git\cmd\git.exe"
set REPO_URL=https://github.com/adhito-nugroho/spt.git

echo ======================================================
echo  SPT — SETUP AWAL SERVER (Clone Repository)
echo  Repo   : %REPO_URL%
echo  Target : %REMOTE_DIR%
echo ======================================================
echo.
echo [PERINGATAN] Script ini dijalankan SEKALI SAJA untuk setup awal.
echo Setelah ini, gunakan deploy-git.bat untuk update rutin.
echo.
pause

echo.
echo [1/2] Mengecek apakah direktori sudah ada di server...
ssh -p %SERVER_PORT% %SERVER_USER%@%SERVER_IP% "if exist \"%REMOTE_DIR%\" echo ALREADY_EXISTS"

echo.
echo [2/2] Meng-clone repository ke server...
echo *(Anda mungkin diminta login GitHub: masukkan token PAT sebagai password)*
echo.

ssh -p %SERVER_PORT% %SERVER_USER%@%SERVER_IP% ^
    "%REMOTE_GIT% clone %REPO_URL% %REMOTE_DIR%"

if errorlevel 1 (
    echo.
    echo [ERROR] Clone gagal!
    echo Kemungkinan penyebab:
    echo   - Direktori %REMOTE_DIR% sudah ada (hapus dulu atau gunakan deploy-git.bat)
    echo   - GitHub credentials tidak valid (gunakan Personal Access Token)
    echo   - SSH tunnel tidak aktif
    pause
    exit /b 1
)

echo.
echo Clone berhasil! Menyiapkan konfigurasi...

echo.
echo [INFO] Langkah manual setelah ini:
echo   1. Salin .env.example ke .env di server:
echo      %REMOTE_DIR%\.env
echo   2. Sesuaikan isi .env dengan konfigurasi database server
echo   3. Jalankan deploy-db.bat untuk migrasi database
echo.
echo ======================================================
echo  SETUP AWAL SELESAI!
echo ======================================================
pause
