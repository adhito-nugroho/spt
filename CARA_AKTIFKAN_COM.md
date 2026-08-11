# Cara Mengaktifkan Extension COM di PHP (untuk Word COM)

## Masalah
File `php_com_dotnet.dll` **tidak ditemukan** di folder `ext` PHP Anda. Extension COM tidak tersedia di build PHP 8.3.6 NTS ini.

## Solusi

### Opsi 1: Download DLL dari PECL (Disarankan)
1. Kunjungi: https://pecl.php.net/package/com_dotnet
2. Download versi yang sesuai dengan PHP 8.3 NTS (Non Thread Safe)
3. Extract file `php_com_dotnet.dll`
4. Copy ke folder: `D:\laragon\bin\php\php-8.3.6-nts-Win32-vs16-x64\ext\`
5. Di `php.ini`, uncomment baris: `extension=com_dotnet` (hapus `;` di depan)
6. Restart Laragon/Apache

### Opsi 2: Gunakan PHP versi lain yang sudah include COM
- Download PHP 8.1 atau 8.2 **TS (Thread Safe)** dari php.net
- Build TS biasanya sudah include `php_com_dotnet.dll`
- Atau gunakan PHP dari XAMPP/WAMP yang biasanya sudah include COM

### Opsi 3: Build dari source (Advanced)
- Clone PHP source code
- Build dengan flag `--enable-com-dotnet`
- Copy DLL yang dihasilkan ke folder ext

## Setelah DLL tersedia

1. **Uncomment di php.ini:**
   ```ini
   extension=com_dotnet
   ```

2. **Restart web server** (Laragon: Stop → Start)

3. **Cek dengan script diagnostik:**
   ```
   http://localhost/spt-php/modules/debug_word_com.php
   ```
   Harus tampil: "BERHASIL: Word COM siap dipakai"

## Catatan
- Extension COM hanya tersedia di **Windows**
- PHP harus jalan sebagai **user yang punya akses ke Microsoft Word**
- Jika masih error setelah DLL ada, cek DCOM permission (lihat README_KONVERSI_PDF.md)
