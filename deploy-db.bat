@echo off
setlocal enabledelayedexpansion

:: =======================================================
:: KONFIGURASI — SPT Database Migration
:: Repository  : https://github.com/adhito-nugroho/spt.git
:: =======================================================

:: --- Database lokal (sumber) ---
set DB_NAME=db_surat_tugas
set DB_USER_LOCAL=root
set DB_PASS_LOCAL=

:: --- SSH ke server (Cloudflare Tunnel) ---
set SERVER_USER=adit
set SERVER_IP=127.0.0.1
set SERVER_PORT=2222

:: --- Database di server (tujuan) ---
set DB_USER_REMOTE=root
set DB_PASS_REMOTE=

:: --- File dump sementara ---
set DUMP_FILE=db_export_temp.sql

:: --- Path MySQL binary di server (Windows Laragon) ---
set "REMOTE_MYSQL=C:\laragon\bin\mysql\mysql-8.0.30-winx64\bin\mysql.exe"

echo ======================================================
echo  SPT ^| Surat Perintah Tugas — MIGRASI DATABASE
echo  DB Lokal : %DB_NAME%
echo  DB Server: %DB_NAME% (host: %SERVER_IP%)
echo ======================================================
echo.

:: 1. Export database dari MySQL Lokal
echo [1/3] Meng-export database lokal (%DB_NAME%)...

powershell -NoProfile -Command ^
    "$candidates = @(Get-ChildItem -Path 'C:\laragon\bin\mysql\*\bin\mysqldump.exe','D:\laragon\bin\mysql\*\bin\mysqldump.exe','C:\Program Files\MySQL\*\bin\mysqldump.exe' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName); if (-not $candidates) { Write-Error 'mysqldump tidak ditemukan di path Laragon/MySQL!'; exit 1 }; $d = $candidates[0]; Write-Host ('Menggunakan: ' + $d); if ('%DB_PASS_LOCAL%' -eq '') { & $d -u %DB_USER_LOCAL% --databases %DB_NAME% --routines --events --single-transaction --result-file='%DUMP_FILE%' } else { & $d -u %DB_USER_LOCAL% -p%DB_PASS_LOCAL% --databases %DB_NAME% --routines --events --single-transaction --result-file='%DUMP_FILE%' }; if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }"

if errorlevel 1 (
    echo.
    echo [ERROR] Gagal export database lokal!
    echo Pastikan Laragon lokal aktif dan nama database benar: %DB_NAME%
    pause
    exit /b 1
)

if not exist %DUMP_FILE% (
    echo [ERROR] File dump tidak terbuat! Cek apakah MySQL Laragon lokal aktif.
    pause
    exit /b 1
)

for %%A in (%DUMP_FILE%) do if %%~zA lss 100 (
    echo [ERROR] File dump terlalu kecil / kosong. Export gagal.
    del %DUMP_FILE%
    pause
    exit /b 1
)

echo Export database OK (file: %DUMP_FILE%)

:: 2. Upload file SQL ke Server via SCP
echo.
echo [2/3] Mengirim file dump SQL ke server via SCP...
echo *(Jika diminta password SSH, masukkan password akun server)*

scp -P %SERVER_PORT% %DUMP_FILE% %SERVER_USER%@%SERVER_IP%:C:/Windows/Temp/%DUMP_FILE%

if errorlevel 1 (
    echo.
    echo [ERROR] Gagal mengunggah file SQL ke server!
    echo Pastikan SSH tunnel Cloudflare aktif dan server dapat dijangkau.
    del %DUMP_FILE%
    pause
    exit /b 1
)

echo Upload SQL OK.

:: 3. Import SQL ke MySQL Server via SSH
echo.
echo [3/3] Meng-import database di MySQL Server (%SERVER_IP%)...

if "%DB_PASS_REMOTE%"=="" (
    ssh -p %SERVER_PORT% %SERVER_USER%@%SERVER_IP% "%REMOTE_MYSQL% -u %DB_USER_REMOTE% < C:\Windows\Temp\%DUMP_FILE% && del C:\Windows\Temp\%DUMP_FILE%"
) else (
    ssh -p %SERVER_PORT% %SERVER_USER%@%SERVER_IP% "%REMOTE_MYSQL% -u %DB_USER_REMOTE% -p%DB_PASS_REMOTE% < C:\Windows\Temp\%DUMP_FILE% && del C:\Windows\Temp\%DUMP_FILE%"
)

if errorlevel 1 (
    echo.
    echo [ERROR] Gagal import database di server!
    if exist %DUMP_FILE% del %DUMP_FILE%
    pause
    exit /b 1
)

if exist %DUMP_FILE% del %DUMP_FILE%

echo.
echo ======================================================
echo  MIGRASI DATABASE MYSQL BERHASIL!
echo  Database '%DB_NAME%' sudah tersinkron ke server.
echo ======================================================
pause
