# 🚨 INSTRUKSI PERBAIKAN UNTUK SERVER

## Error yang Terjadi:
```
Error: Template untuk kategori "Banyak Pegawai (4+ pegawai)" tidak ditemukan
```

**URL Error**: `tugas.cdkbojonegoro.site/modules/generate-surat.php?id=35`

---

## ✅ LANGKAH PERBAIKAN (Pilih Salah Satu)

### 🔥 METODE 1: Auto-Fix (TERCEPAT - 2 Menit)

1. **Upload file ini ke server:**
   - `modules/debug_templates.php`
   - `modules/check_and_fix_templates.php`

2. **Buka di browser:**
   ```
   https://tugas.cdkbojonegoro.site/modules/debug_templates.php
   ```
   Script ini akan menampilkan:
   - ✅ Apakah template ada di database
   - ✅ Apakah file fisik ada
   - ✅ Kategori mana yang hilang

3. **Kemudian buka:**
   ```
   https://tugas.cdkbojonegoro.site/modules/check_and_fix_templates.php
   ```

4. **Klik tombol "Perbaiki Sekarang"**

5. **✅ SELESAI!** Test generate surat lagi.

---

### 📝 METODE 2: Manual via phpMyAdmin (5 Menit)

1. **Login ke cPanel/phpMyAdmin server**

2. **Pilih database aplikasi** (lihat di `config/database.php`)

3. **Klik tab "SQL"**

4. **Jalankan query ini satu per satu:**

```sql
-- Cek apakah ada template
SELECT * FROM templates;
```

**Jika hasilnya KOSONG atau tidak ada kategori `banyak_pegawai`:**

```sql
-- Tambahkan kolom kategori_template (jika belum ada)
ALTER TABLE templates 
ADD COLUMN IF NOT EXISTS kategori_template 
ENUM('1_pegawai', '2_pegawai', '3_pegawai', 'banyak_pegawai') 
NOT NULL DEFAULT 'banyak_pegawai' 
AFTER nama;

-- Ganti 'nama_file_template_anda.docx' dengan nama file yang ADA di folder assets/templates/
-- Untuk cek file apa yang ada, buka via FTP folder: assets/templates/

-- Template untuk Banyak Pegawai (4+)
INSERT INTO templates 
(nama, kategori_template, nama_file, path_file, ukuran_file, is_default, deskripsi) 
VALUES 
('Template Surat Tugas - Banyak Pegawai', 'banyak_pegawai', 'nama_file_template_anda.docx', 'assets/templates/nama_file_template_anda.docx', 25600, 1, 'Template untuk surat tugas dengan 4 atau lebih pegawai');

-- Template untuk 1 Pegawai
INSERT INTO templates 
(nama, kategori_template, nama_file, path_file, ukuran_file, is_default, deskripsi) 
VALUES 
('Template Surat Tugas - 1 Pegawai', '1_pegawai', 'nama_file_template_anda.docx', 'assets/templates/nama_file_template_anda.docx', 25600, 1, 'Template untuk surat tugas dengan 1 pegawai');

-- Template untuk 2 Pegawai
INSERT INTO templates 
(nama, kategori_template, nama_file, path_file, ukuran_file, is_default, deskripsi) 
VALUES 
('Template Surat Tugas - 2 Pegawai', '2_pegawai', 'nama_file_template_anda.docx', 'assets/templates/nama_file_template_anda.docx', 25600, 1, 'Template untuk surat tugas dengan 2 pegawai');

-- Template untuk 3 Pegawai
INSERT INTO templates 
(nama, kategori_template, nama_file, path_file, ukuran_file, is_default, deskripsi) 
VALUES 
('Template Surat Tugas - 3 Pegawai', '3_pegawai', 'nama_file_template_anda.docx', 'assets/templates/nama_file_template_anda.docx', 25600, 1, 'Template untuk surat tugas dengan 3 pegawai');
```

5. **Verifikasi:**
```sql
SELECT id, nama, kategori_template, nama_file, is_default 
FROM templates 
ORDER BY kategori_template;
```

Harus ada 4 baris (1_pegawai, 2_pegawai, 3_pegawai, banyak_pegawai)

6. **✅ SELESAI!** Test generate surat lagi.

---

### 💾 METODE 3: Import Database Lengkap (10 Menit)

**Jika Anda punya backup database lokal yang lengkap:**

1. **Di Lokal - Export Database:**
   - Buka phpMyAdmin lokal
   - Pilih database
   - Tab "Export"
   - Pilih "Custom"
   - Centang tabel `templates` (atau semua tabel)
   - Klik "Go" → Download file `.sql`

2. **Di Server - Import Database:**
   - Login cPanel → phpMyAdmin
   - Pilih database
   - Tab "Import"
   - Choose file → Pilih file `.sql` tadi
   - Klik "Go"

3. **✅ SELESAI!** Test generate surat lagi.

---

## 🔍 CARA CEK NAMA FILE TEMPLATE DI SERVER

**Via FTP/File Manager:**
1. Login cPanel
2. Buka File Manager
3. Navigate ke: `public_html/assets/templates/`
4. Lihat file `.docx` yang ada
5. Catat nama file lengkapnya (case-sensitive!)

**Via Script:**
1. Upload `debug_templates.php` ke folder `modules/`
2. Akses: `https://tugas.cdkbojonegoro.site/modules/debug_templates.php`
3. Lihat section "File Template di Server"

---

## ⚠️ CATATAN PENTING

1. **Nama file harus PERSIS SAMA** antara database dan file fisik (case-sensitive)
2. **File harus ada** di folder `assets/templates/`
3. **Satu file bisa digunakan** untuk semua kategori (tidak perlu 4 file berbeda)
4. **Ukuran file** (25600) adalah contoh, bisa diganti dengan ukuran sebenarnya

---

## 🆘 Jika Masih Error

1. **Cek error log:**
   - cPanel → Error Log
   - Atau file `error_log` di root folder

2. **Cek permission folder:**
   - Folder `assets/templates/` harus readable (755)
   - File `.docx` harus readable (644)

3. **Cek database connection:**
   - File `config/database.php`
   - Pastikan DB_HOST, DB_NAME, DB_USER, DB_PASS benar

4. **Contact support** dengan info:
   - Screenshot error
   - Hasil dari `debug_templates.php`
   - Error log

---

## 📞 Quick Reference

| Script | URL | Fungsi |
|--------|-----|--------|
| Debug | `https://tugas.cdkbojonegoro.site/modules/debug_templates.php` | Diagnosa masalah |
| Auto-Fix | `https://tugas.cdkbojonegoro.site/modules/check_and_fix_templates.php` | Perbaikan otomatis |
| Template Manager | `https://tugas.cdkbojonegoro.site/modules/template.php` | Kelola template |

---

**Dibuat**: 22 Januari 2026  
**Untuk**: tugas.cdkbojonegoro.site
