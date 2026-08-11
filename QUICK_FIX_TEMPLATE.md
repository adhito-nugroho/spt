# 🚀 Panduan Cepat: Mengatasi Error Template di Server

## ❌ Error:
```
Template untuk kategori "Banyak Pegawai (4+ pegawai)" tidak ditemukan
```

## ✅ Solusi Tercepat (3 Langkah):

### 1️⃣ Buka Script Perbaikan
Akses di browser:
```
http://domain-anda.com/modules/check_and_fix_templates.php
```

### 2️⃣ Klik "Perbaiki Sekarang"
Script akan otomatis menambahkan template yang hilang ke database

### 3️⃣ Selesai!
Coba generate surat tugas lagi

---

## 🔍 Kenapa Terjadi?

**Di Lokal**: File template ✅ + Database ✅ = Berfungsi ✅

**Di Server**: File template ✅ + Database ❌ = Error ❌

Saat upload ke server, Anda hanya memindahkan **file**, tapi **database tidak ikut**.

---

## 📋 Solusi Alternatif:

### Opsi A: Upload Ulang Template
1. Buka `modules/template.php`
2. Klik "Upload Template Baru"
3. Upload file .docx
4. Pilih kategori "Banyak Pegawai (4+ pegawai)"
5. Selesai!

### Opsi B: Import Database
1. Export database dari lokal (phpMyAdmin → Export)
2. Import ke server (phpMyAdmin → Import)
3. Update `config/database.php` dengan kredensial server
4. Selesai!

### Opsi C: Manual via SQL
1. Buka phpMyAdmin di server
2. Pilih database aplikasi
3. Klik tab "SQL"
4. Copy-paste isi file `database/add_templates.sql`
5. Ganti nama file template sesuai file Anda
6. Klik "Go"
7. Selesai!

---

## 📞 Butuh Bantuan?

Lihat dokumentasi lengkap: `TROUBLESHOOTING_TEMPLATE.md`

---

**Tips**: Satu file template bisa digunakan untuk semua kategori (1, 2, 3, atau banyak pegawai). Tidak perlu 4 file berbeda!
