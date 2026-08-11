# Dokumentasi Perbaikan Error "Template Tidak Ditemukan"

## Masalah yang Ditemukan

Pada modul generate surat, terjadi error **"template tidak ditemukan"** meskipun file template sudah ada di folder `assets/templates/`.

### Analisis Masalah

Setelah melakukan investigasi mendalam, ditemukan bahwa:

1. ✅ Tabel `templates` sudah memiliki kolom `kategori_template`
2. ✅ File template fisik sudah ada di folder `assets/templates/`
3. ❌ **Database hanya memiliki 2 kategori template**: `1_pegawai` dan `banyak_pegawai`
4. ❌ **Kategori `2_pegawai` dan `3_pegawai` TIDAK ADA di database**

### Penyebab Error

Sistem generate surat menggunakan logika pemilihan template berdasarkan jumlah pegawai:
- 1 pegawai → kategori `1_pegawai`
- 2 pegawai → kategori `2_pegawai` ❌ (tidak ada di database)
- 3 pegawai → kategori `3_pegawai` ❌ (tidak ada di database)
- 4+ pegawai → kategori `banyak_pegawai`

Ketika user mencoba generate surat dengan 2 atau 3 pegawai, sistem mencari template dengan kategori `2_pegawai` atau `3_pegawai` di database, tetapi tidak menemukannya, sehingga muncul error.

## Solusi yang Diterapkan

### 1. Script Diagnostic (`check_templates.php`)

Dibuat script untuk mengecek:
- Struktur tabel templates
- Data template di database
- File template di folder
- Mapping antara database dan file
- Simulasi pencarian template untuk berbagai jumlah pegawai

### 2. Script Fix (`fix_missing_templates.php`)

Dibuat script untuk menambahkan template yang hilang:
- Mendeteksi kategori yang hilang
- Menggunakan file template yang sudah ada sebagai referensi
- Menambahkan entry baru ke database untuk kategori `2_pegawai` dan `3_pegawai`

### 3. Hasil Perbaikan

Setelah menjalankan script fix:
- ✅ Template untuk kategori `2_pegawai` berhasil ditambahkan (ID: 11)
- ✅ Template untuk kategori `3_pegawai` berhasil ditambahkan (ID: 12)
- ✅ Total template di database: **4 template** (semua kategori lengkap)
- ✅ Simulasi pencarian template berhasil untuk semua jumlah pegawai (1, 2, 3, 4+)

## Struktur Database Template

```sql
CREATE TABLE templates (
    id INT AUTO_INCREMENT PRIMARY KEY,
    nama VARCHAR(255) NOT NULL,
    kategori_template ENUM('1_pegawai', '2_pegawai', '3_pegawai', 'banyak_pegawai') NOT NULL DEFAULT 'banyak_pegawai',
    nama_file VARCHAR(255) NOT NULL,
    path_file VARCHAR(500) NOT NULL,
    ukuran_file BIGINT NOT NULL,
    is_default TINYINT(1) DEFAULT 0,
    deskripsi TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX idx_kategori (kategori_template, is_default)
);
```

## Template yang Terdaftar

| ID | Nama | Kategori | File | Default |
|----|------|----------|------|---------|
| 6 | Template satu orang | 1_pegawai | 1766636680_template_satu_orang.docx | YES |
| 11 | Template Surat Tugas - 2 Pegawai | 2_pegawai | 1768880139_template_banyak_orang.docx | YES |
| 12 | Template Surat Tugas - 3 Pegawai | 3_pegawai | 1768880139_template_banyak_orang.docx | YES |
| 10 | template banyak orang | banyak_pegawai | 1768880139_template_banyak_orang.docx | YES |

## Cara Kerja Sistem Template

1. **User membuat surat tugas** dengan memilih pegawai
2. **Sistem menghitung jumlah pegawai** yang dipilih
3. **Sistem menentukan kategori template** berdasarkan jumlah:
   - 1 pegawai → `1_pegawai`
   - 2 pegawai → `2_pegawai`
   - 3 pegawai → `3_pegawai`
   - 4+ pegawai → `banyak_pegawai`
4. **Sistem mencari template default** untuk kategori tersebut di database
5. **Jika tidak ada default**, sistem mencari template pertama dalam kategori tersebut
6. **Jika kategori tidak ada**, sistem fallback ke `banyak_pegawai`
7. **Sistem memvalidasi file** template (exists, readable, valid .docx)
8. **Sistem generate surat** menggunakan template yang dipilih

## File Template yang Digunakan

Saat ini, sistem menggunakan 2 file template fisik:
1. `1766636680_template_satu_orang.docx` - untuk 1 pegawai
2. `1768880139_template_banyak_orang.docx` - untuk 2, 3, dan 4+ pegawai

**Catatan**: Satu file template yang sama bisa digunakan untuk beberapa kategori. Sistem akan otomatis memilih kategori yang sesuai berdasarkan jumlah pegawai.

## Rekomendasi

### Upload Template Khusus (Opsional)

Jika diperlukan template khusus untuk setiap kategori:
1. Buka halaman **Template Word** (`template.php`)
2. Klik **Upload Template**
3. Pilih kategori yang sesuai
4. Upload file template .docx
5. Set sebagai default jika diperlukan

### Maintenance

Untuk memastikan sistem berjalan lancar:
1. Pastikan setiap kategori memiliki minimal 1 template
2. Pastikan minimal 1 template per kategori di-set sebagai default
3. Pastikan file template tidak corrupt dan bisa dibaca
4. Backup file template secara berkala

## Troubleshooting

Jika masih muncul error "template tidak ditemukan":

1. **Cek database**: Jalankan `check_templates.php` untuk melihat template yang terdaftar
2. **Cek file**: Pastikan file template ada di folder `assets/templates/`
3. **Cek kategori**: Pastikan kategori template sesuai dengan jumlah pegawai
4. **Cek default**: Pastikan ada template default untuk setiap kategori
5. **Cek permission**: Pastikan file template bisa dibaca (readable)

## Log Error

Error log bisa dilihat di:
- PHP error log (biasanya di `laragon/logs/php_error.log`)
- Apache error log (biasanya di `laragon/logs/apache_error.log`)

Untuk debugging, aktifkan error logging di `generate-surat.php` dengan melihat fungsi `error_log()`.

## Kesimpulan

Masalah **"template tidak ditemukan"** sudah berhasil diperbaiki dengan menambahkan entry database untuk kategori `2_pegawai` dan `3_pegawai`. Sistem sekarang bisa generate surat untuk semua jumlah pegawai (1, 2, 3, 4+) tanpa error.

---

**Tanggal Perbaikan**: 22 Januari 2026  
**Status**: ✅ **SELESAI**
