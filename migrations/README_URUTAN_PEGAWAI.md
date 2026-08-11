# Update: Urutan Pegawai di Surat Tugas

## Masalah yang Diperbaiki
Sebelumnya, urutan pegawai yang ditampilkan di file Word yang di-generate tidak sesuai dengan urutan ketika nama pegawai di-input. Pegawai yang lebih dulu di-input tidak selalu dicetak lebih dulu.

## Solusi
Menambahkan kolom `urutan` pada tabel `pegawai_tugas` untuk menyimpan urutan input pegawai, sehingga urutan pegawai di dokumen Word akan sesuai dengan urutan input.

## Cara Menjalankan Migration

### Opsi 1: Via Browser (Recommended)
1. Buka browser
2. Akses URL: `http://localhost/spt-php/modules/run_migration.php`
3. Migration akan berjalan otomatis
4. Tunggu sampai muncul pesan "Migration completed successfully!"

### Opsi 2: Via SQL (Manual)
1. Buka phpMyAdmin atau MySQL client
2. Pilih database yang digunakan
3. Jalankan query berikut:

```sql
-- Add urutan column
ALTER TABLE pegawai_tugas ADD COLUMN urutan INT NOT NULL DEFAULT 0 AFTER nip;
```

4. Kemudian jalankan script untuk update data existing (opsional, untuk data lama):

```sql
-- Update existing records
SET @row_num = 0;
UPDATE pegawai_tugas 
SET urutan = (@row_num := @row_num + 1)
ORDER BY id_surat_tugas, id;
```

## Perubahan yang Dilakukan

### 1. Database
- **Tabel**: `pegawai_tugas`
- **Kolom baru**: `urutan` (INT, NOT NULL, DEFAULT 0)
- **Posisi**: Setelah kolom `nip`

### 2. File yang Dimodifikasi

#### `modules/surat-tugas.php`
- **Fungsi Create**: Menambahkan penyimpanan urutan saat insert pegawai baru
- **Fungsi Update**: Menambahkan penyimpanan urutan saat update pegawai

#### `modules/generate-surat.php`
- **Query SELECT**: Menambahkan `ORDER BY pt.urutan` pada semua GROUP_CONCAT
- Ini memastikan data pegawai diambil sesuai urutan input

## Cara Kerja

1. **Saat Input Surat Tugas Baru**:
   - Pegawai pertama yang dipilih mendapat urutan = 1
   - Pegawai kedua mendapat urutan = 2
   - Dan seterusnya...

2. **Saat Generate Surat**:
   - Query mengambil data pegawai dengan `ORDER BY pt.urutan`
   - Data ditampilkan di Word sesuai urutan 1, 2, 3, dst
   - Pegawai yang lebih dulu di-input akan dicetak lebih dulu

## Testing

Setelah migration, lakukan testing:

1. **Test dengan Surat Baru**:
   - Buat surat tugas baru
   - Pilih beberapa pegawai dengan urutan tertentu
   - Generate surat
   - Periksa apakah urutan di Word sesuai dengan urutan input

2. **Test dengan Surat Lama** (setelah migration):
   - Edit surat tugas yang sudah ada
   - Ubah urutan pegawai jika perlu
   - Generate ulang surat
   - Periksa apakah urutan sudah benar

## Catatan Penting

- Migration hanya perlu dijalankan **SATU KALI**
- Untuk data surat tugas yang sudah ada sebelum migration:
  - Urutan akan diset berdasarkan ID (urutan insert ke database)
  - Jika ingin mengubah urutan, edit surat tugas dan save ulang
- Untuk surat tugas baru setelah migration:
  - Urutan otomatis tersimpan sesuai urutan input

## Troubleshooting

### Error: Column 'urutan' already exists
- Migration sudah pernah dijalankan
- Tidak perlu menjalankan lagi

### Error: Unknown column 'pt.urutan'
- Migration belum dijalankan
- Jalankan migration terlebih dahulu

### Urutan masih tidak sesuai
- Pastikan migration sudah berhasil
- Edit surat tugas dan save ulang untuk update urutan
- Clear cache browser jika perlu

## Support
Jika ada masalah, hubungi developer atau cek log error di:
- Browser console (F12)
- PHP error log
- MySQL error log
