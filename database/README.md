# 📁 Database Scripts

Folder ini berisi script SQL untuk setup dan troubleshooting database.

## 📄 File yang Tersedia:

### `add_templates.sql`
Script untuk menambahkan template ke database secara manual.

**Kapan digunakan:**
- Saat migrasi ke server dan template tidak ikut
- Saat ingin menambahkan template via SQL langsung
- Untuk troubleshooting masalah template

**Cara menggunakan:**
1. Buka phpMyAdmin
2. Pilih database aplikasi
3. Klik tab "SQL"
4. Copy-paste isi file `add_templates.sql`
5. **PENTING**: Ganti `template_surat_tugas.docx` dengan nama file template Anda
6. Klik "Go"

---

## 🔧 Script PHP untuk Auto-Fix:

Jika Anda lebih suka menggunakan interface web, gunakan:

### `modules/check_and_fix_templates.php`
Script PHP dengan interface web yang bisa:
- ✅ Mendeteksi template yang hilang
- ✅ Menampilkan file fisik vs database
- ✅ Auto-fix dengan 1 klik
- ✅ Verifikasi hasil perbaikan

**Cara menggunakan:**
```
http://domain-anda.com/modules/check_and_fix_templates.php
```

---

## 📚 Dokumentasi:

- **QUICK_FIX_TEMPLATE.md** - Panduan cepat (3 langkah)
- **TROUBLESHOOTING_TEMPLATE.md** - Dokumentasi lengkap

---

## ⚠️ Catatan Penting:

1. **Backup database** sebelum menjalankan script SQL
2. Pastikan file template (.docx) sudah ada di `assets/templates/`
3. Satu file template bisa digunakan untuk semua kategori
4. Ukuran file dalam bytes (contoh: 25600 = 25KB)

---

## 🆘 Troubleshooting:

### Error: "Table 'templates' doesn't exist"
Jalankan script pembuatan tabel terlebih dahulu (ada di migration files)

### Error: "Column 'kategori_template' doesn't exist"
Script akan otomatis menambahkan kolom ini

### Template masih tidak ditemukan setelah insert
1. Cek apakah file .docx benar-benar ada di folder `assets/templates/`
2. Cek nama file di database sama dengan nama file fisik (case-sensitive)
3. Cek permission folder (harus readable)

---

**Terakhir diperbarui**: 22 Januari 2026
