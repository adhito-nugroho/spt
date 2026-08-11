@echo off
setlocal enabledelayedexpansion

:: =======================================================
:: KONFIGURASI — SPT (Surat Perintah Tugas)
:: Repository  : https://github.com/adhito-nugroho/spt.git
:: =======================================================

:: --- SSH ke server (Cloudflare Tunnel) ---
set SERVER_USER=adit
set SERVER_IP=127.0.0.1
set SERVER_PORT=2222

:: --- Direktori aplikasi di server ---
set "REMOTE_DIR=C:\laragon\www\spt-php"

:: --- Branch yang di-deploy ---
set BRANCH=main

:: --- Path binary di server (Windows Laragon) ---
set "REMOTE_GIT=C:\laragon\bin\git\cmd\git.exe"
set "REMOTE_PHP=C:\laragon\bin\php\php-8.1.10-Win32-vs16-x64\php.exe"

echo ======================================================
echo  SPT ^| Surat Perintah Tugas — DEPLOY GIT PULL
echo  Repo : https://github.com/adhito-nugroho/spt.git
echo  Branch : %BRANCH%
echo ======================================================
echo.

:: 1. Commit lokal (jika ada perubahan) lalu push ke GitHub
echo [1/3] Mendorong perubahan lokal ke GitHub...
echo.

git add .

set "NEED_COMMIT=0"
for /f %%i in ('git diff --cached --name-only') do set "NEED_COMMIT=1"

if "!NEED_COMMIT!"=="0" (
    echo Tidak ada perubahan lokal yang perlu di-commit.
    echo Melanjutkan push/pull...
) else (
    set /p msg="Masukkan pesan commit (Enter = 'update aplikasi SPT'): "
    if "!msg!"=="" set "msg=update aplikasi SPT"

    git commit -m "!msg!"
    if errorlevel 1 (
        echo.
        echo [ERROR] Gagal melakukan git commit!
        pause
        exit /b 1
    )
)

git push origin %BRANCH%
if errorlevel 1 (
    echo.
    echo [ERROR] Gagal git push ke GitHub!
    echo Pastikan Anda sudah login ke GitHub dan remote sudah dikonfigurasi.
    pause
    exit /b 1
)

echo Push berhasil!

:: 2. Git Pull di Server via SSH
echo.
echo [2/3] Menjalankan 'git pull' di server via SSH...
echo *(Jika diminta password SSH, masukkan password akun server)*
echo.

set "REPO_URL=https://github.com/adhito-nugroho/spt.git"
ssh -p %SERVER_PORT% %SERVER_USER%@%SERVER_IP% "%REMOTE_GIT% config --global --add safe.directory %REMOTE_DIR:\=/% && cd /d %REMOTE_DIR% && (%REMOTE_GIT% remote add origin %REPO_URL% 2>nul || %REMOTE_GIT% remote set-url origin %REPO_URL%) && %REMOTE_GIT% pull origin %BRANCH%"

if errorlevel 1 (
    echo.
    echo [ERROR] Git Pull di server gagal.
    pause
    exit /b 1
)

echo Git Pull di server berhasil!

:: 3. Jalankan PHP Migration di Server via SSH
echo.
echo [3/3] Menjalankan migrasi database di server...

ssh -p %SERVER_PORT% %SERVER_USER%@%SERVER_IP% "cd /d %REMOTE_DIR% && %REMOTE_PHP% migrate.php"

if errorlevel 1 (
    echo.
    echo [PERINGATAN] Migrasi CLI server gagal. Anda juga bisa membuka migrate.php di browser di server.
)

echo.
echo ======================================================
echo  DEPLOYMENT BERHASIL!
echo  URL Server : http://%SERVER_IP%/spt (sesuaikan)
echo ======================================================
pause
