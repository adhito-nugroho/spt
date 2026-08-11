# 🔧 Troubleshooting: Template Tidak Ditemukan

## ❌ Error yang Terjadi

```
Error: Template untuk kategori "Banyak Pegawai (4+ pegawai)" tidak ditemukan. 
Silakan upload template yang sesuai.
```

---

## 🔍 Penyebab Masalah

### Kenapa Error Ini Terjadi di Server tapi Tidak di Lokal?

Ketika aplikasi dipindahkan dari **lokal ke server**, yang terjadi adalah:

1. **✅ File aplikasi (PHP, CSS, JS)** → Berhasil dipindahkan
2. **✅ File template (.docx)** di folder `assets/templates/` → Berhasil dipindahkan
3. **❌ Data database (tabel `templates`)** → **TIDAK dipindahkan** atau tidak lengkap

### Penjelasan Detail:

```
LOKAL (Berfungsi Normal):
├── File aplikasi ✓
├── File template .docx ✓
└── Database:
    └── Tabel templates:
        ├── ID 1: Template 1 Pegawai ✓
        ├── ID 2: Template 2 Pegawai ✓
        ├── ID 3: Template 3 Pegawai ✓
        └── ID 4: Template Banyak Pegawai ✓

SERVER (Error):
├── File aplikasi ✓
├── File template .docx ✓
└── Database:
    └── Tabel templates:
        └── (KOSONG atau tidak lengkap) ❌
```

### Mengapa Database Tidak Ikut Pindah?

- Saat upload file ke server via FTP/cPanel, hanya **file fisik** yang dipindahkan
- **Data database** harus di-export dari lokal dan di-import ke server secara terpisah
- Jika lupa export/import database, tabel `templates` akan kosong meskipun file .docx sudah ada

---

## ✅ Solusi

### **Solusi 1: Gunakan Script Auto-Fix (TERCEPAT)** ⚡

1. **Buka browser** dan akses:
   ```
   http://domain-anda.com/modules/check_and_fix_templates.php
   ```
   atau di lokal:
   ```
   http://localhost/spt-php/modules/check_and_fix_templates.php
   ```

2. Script akan menampilkan:
   - ✅ Template yang ada di database
   - 📁 File template yang ada di folder
   - ⚠️ Kategori yang hilang
   - 🔧 Tombol "Perbaiki Sekarang"

3. **Klik tombol "Perbaiki Sekarang"**

4. ✅ Selesai! Template akan otomatis ditambahkan ke database

---

### **Solusi 2: Upload Ulang Template** 📤

1. Buka halaman **Template Word**:
   ```
   http://domain-anda.com/modules/template.php
   ```

2. Klik **"Upload Template Baru"**

3. Upload file template .docx Anda

4. Pilih kategori: **"Banyak Pegawai (4+ pegawai)"**

5. Klik **"Upload"**

6. ✅ Selesai!

---

### **Solusi 3: Export/Import Database (LENGKAP)** 💾

Jika Anda ingin memindahkan semua data dari lokal ke server:

#### Di Lokal (Export):

1. Buka **phpMyAdmin** lokal
2. Pilih database aplikasi
3. Klik tab **"Export"**
4. Pilih **"Custom"**
5. Centang tabel yang ingin di-export (atau pilih semua)
6. Klik **"Go"** untuk download file `.sql`

#### Di Server (Import):

1. Buka **phpMyAdmin** di server
2. Pilih database aplikasi
3. Klik tab **"Import"**
4. Pilih file `.sql` yang sudah di-download
5. Klik **"Go"**
6. ✅ Selesai!

---

## 🎯 Cara Mencegah Masalah Ini

### Saat Migrasi ke Server:

1. **Export database** dari lokal
2. **Upload file aplikasi** ke server
3. **Import database** ke server
4. **Update file `config/database.php`** dengan kredensial database server
5. **Test aplikasi** di server

### Checklist Migrasi:

```
☐ Export database lokal (.sql)
☐ Upload semua file aplikasi
☐ Upload file template (.docx) ke folder assets/templates/
☐ Buat database baru di server
☐ Import file .sql ke database server
☐ Update config/database.php:
   - DB_HOST (biasanya 'localhost')
   - DB_NAME (nama database di server)
   - DB_USER (username database server)
   - DB_PASS (password database server)
☐ Test akses aplikasi
☐ Test generate surat tugas
```

---

## 📋 Informasi Teknis

### Struktur Tabel `templates`:

```sql
CREATE TABLE templates (
    id INT AUTO_INCREMENT PRIMARY KEY,
    nama VARCHAR(255) NOT NULL,
    kategori_template ENUM('1_pegawai', '2_pegawai', '3_pegawai', 'banyak_pegawai') NOT NULL DEFAULT 'banyak_pegawai',
    nama_file VARCHAR(255) NOT NULL,
    path_file VARCHAR(255) NOT NULL,
    ukuran_file INT NOT NULL,
    is_default TINYINT(1) DEFAULT 0,
    deskripsi TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX idx_kategori (kategori_template, is_default)
);
```

### Kategori Template yang Diperlukan:

| Kategori | Label | Digunakan Untuk |
|----------|-------|-----------------|
| `1_pegawai` | 1 Pegawai | Surat tugas dengan 1 pegawai |
| `2_pegawai` | 2 Pegawai | Surat tugas dengan 2 pegawai |
| `3_pegawai` | 3 Pegawai | Surat tugas dengan 3 pegawai |
| `banyak_pegawai` | Banyak Pegawai (4+ pegawai) | Surat tugas dengan 4 atau lebih pegawai |

### Logika Pemilihan Template:

```php
// Di file generate-surat.php
if ($jumlah_pegawai == 1) {
    $kategori_template = '1_pegawai';
} elseif ($jumlah_pegawai == 2) {
    $kategori_template = '2_pegawai';
} elseif ($jumlah_pegawai == 3) {
    $kategori_template = '3_pegawai';
} else {
    $kategori_template = 'banyak_pegawai'; // 4+ pegawai
}
```

### Fallback Mechanism:

Jika template untuk kategori spesifik tidak ditemukan, sistem akan:

1. Coba cari template default untuk kategori tersebut
2. Jika tidak ada, coba cari template non-default untuk kategori tersebut
3. Jika masih tidak ada, fallback ke template `banyak_pegawai`
4. Jika `banyak_pegawai` juga tidak ada → **ERROR**

---

## 🆘 Bantuan Lebih Lanjut

### File-file Terkait:

- `modules/generate-surat.php` - Logika pemilihan template
- `modules/template.php` - Halaman kelola template
- `modules/check_and_fix_templates.php` - Script diagnosa & perbaikan
- `modules/fix_missing_templates.php` - Script perbaikan alternatif
- `config/database.php` - Konfigurasi database

### Log Error:

Cek error log di:
- **Server**: `error_log` di root atau folder `logs/`
- **PHP Error Log**: Lihat di phpinfo() untuk lokasi error_log

### Kontak Support:

Jika masalah masih berlanjut, hubungi administrator sistem dengan informasi:
- Screenshot error
- Hasil dari `check_and_fix_templates.php`
- Versi PHP server
- Versi MySQL/MariaDB

---

## 📝 Catatan Penting

⚠️ **Satu file template bisa digunakan untuk beberapa kategori**

Anda tidak perlu membuat 4 file template berbeda. Satu file .docx yang sama bisa digunakan untuk semua kategori. Sistem akan otomatis mengisi data sesuai jumlah pegawai.

✅ **Template yang baik harus memiliki placeholder:**

```
${nomor_surat}
${tanggal_surat}
${dasar_surat}
${untuk}
${waktu_pelaksanaan}

Untuk pegawai (pilih salah satu metode):
1. Individual: ${nama1}, ${nip1}, ${pangkat1}, ${jabatan1}
2. Blok teks: ${daftar_pegawai}
3. Tabel dinamis: ${no}, ${nama}, ${nip}, ${pangkat}, ${jabatan}
```

---

**Terakhir diperbarui**: 22 Januari 2026
