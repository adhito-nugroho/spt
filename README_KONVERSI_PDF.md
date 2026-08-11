# Konversi DOCX ke PDF

Sistem ini mendukung tiga metode konversi (prioritas):

1. **Microsoft Word COM** – Hasil PDF 100% sama dengan tampilan Word. Hanya di Windows + Word terinstall.
2. **LibreOffice** – Hasil mendekati Word. Cross-platform.
3. **DomPDF** – Fallback terakhir (tampilan bisa berbeda). Via Composer.

---

## Mengaktifkan Microsoft Word COM (untuk hasil persis Word)

### Syarat
- OS: **Windows Server** (atau Windows dengan Microsoft Office/Word terinstall).
- PHP **7.4+** dengan extension **COM** aktif.
- User yang menjalankan PHP punya akses ke Microsoft Word (Desktop/COM).

### 1. Aktifkan extension COM di PHP
- Buka **php.ini** (path bisa dari `php --ini` atau info dari Laragon).
- Pastikan baris berikut **tidak** dikomentari:
  ```ini
  extension=php_com_dotnet
  ```
  Di Windows sering berupa:
  ```ini
  extension=php_com_dotnet.dll
  ```
- Simpan lalu **restart web server** (Apache/IIS/Laragon).

### 2. User yang menjalankan PHP
- PHP harus jalan sebagai user yang bisa membuka Word (bukan SYSTEM tanpa profil).
- **IIS**: Application Pool → Identity = account yang punya Word; atau set **Load User Profile = true**.
- **Apache/Laragon**: Jalankan service sebagai user yang sama dengan yang dipakai sehari-hari (bisa akses Word).
- **Laragon**: Biasanya sudah benar jika Anda pakai “Run as current user”.

### 3. (Opsional) DCOM / Word permission
- Jika Word tidak bisa dijalankan dari PHP:
  - Buka **Component Services** → **DCOM Config**.
  - Cari **Microsoft Word** → Properties → **Security**.
  - Pastikan account yang menjalankan PHP punya **Launch and Activation** permission.

### 4. Cek dengan script diagnostik
- Buka di browser: `http://your-server/spt-php/modules/debug_word_com.php`
- Pastikan tampil **BERHASIL** dan **Word Version** terbaca.

### 5. Path Word custom
- Jika Word terinstall di path tidak standar, di `config/database.php` (atau file config yang di-include paling awal) bisa ditambah:
  ```php
  // Tidak wajib; COM pakai ProgID "Word.Application", biasanya tidak perlu path.
  ```

---

## LibreOffice (fallback)
- Install LibreOffice di server.
- Script akan mencari `soffice.exe` di path standar (Program Files, Laragon, dll).
- Untuk path custom: definisikan `PATH_SOFFICE` di config.

## DomPDF (fallback terakhir)
- Jalankan: `composer install` atau `composer update`.
- Paket `dompdf/dompdf` harus terpasang di `vendor/`.

---

## Troubleshooting
- **"class COM tidak ada"** → Aktifkan `php_com_dotnet` di php.ini dan restart web server.
- **Word tidak terbuka / timeout** → Pastikan `set_time_limit(120)` cukup; cek user dan DCOM.
- **File PDF tidak muncul** → Cek log PHP (`error_log`); jalankan `debug_word_com.php` untuk diagnosa.
