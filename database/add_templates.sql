-- ============================================
-- Script SQL untuk Menambahkan Template
-- ============================================
-- Gunakan script ini jika Anda ingin menambahkan
-- template secara manual melalui phpMyAdmin
-- ============================================

-- 1. Pastikan kolom kategori_template ada
ALTER TABLE `templates` 
ADD COLUMN IF NOT EXISTS `kategori_template` 
ENUM('1_pegawai', '2_pegawai', '3_pegawai', 'banyak_pegawai') 
NOT NULL DEFAULT 'banyak_pegawai' 
AFTER `nama`;

-- 2. Tambahkan index untuk performa
ALTER TABLE `templates` 
ADD INDEX IF NOT EXISTS `idx_kategori` (`kategori_template`, `is_default`);

-- ============================================
-- CONTOH: Menambahkan Template
-- ============================================
-- PENTING: Ganti 'template_surat_tugas.docx' dengan nama file template Anda yang sebenarnya
-- File harus sudah ada di folder: assets/templates/

-- Template untuk 1 Pegawai
INSERT INTO `templates` 
(`nama`, `kategori_template`, `nama_file`, `path_file`, `ukuran_file`, `is_default`, `deskripsi`) 
VALUES 
('Template Surat Tugas - 1 Pegawai', '1_pegawai', 'template_surat_tugas.docx', 'assets/templates/template_surat_tugas.docx', 25600, 1, 'Template untuk surat tugas dengan 1 pegawai');

-- Template untuk 2 Pegawai
INSERT INTO `templates` 
(`nama`, `kategori_template`, `nama_file`, `path_file`, `ukuran_file`, `is_default`, `deskripsi`) 
VALUES 
('Template Surat Tugas - 2 Pegawai', '2_pegawai', 'template_surat_tugas.docx', 'assets/templates/template_surat_tugas.docx', 25600, 1, 'Template untuk surat tugas dengan 2 pegawai');

-- Template untuk 3 Pegawai
INSERT INTO `templates` 
(`nama`, `kategori_template`, `nama_file`, `path_file`, `ukuran_file`, `is_default`, `deskripsi`) 
VALUES 
('Template Surat Tugas - 3 Pegawai', '3_pegawai', 'template_surat_tugas.docx', 'assets/templates/template_surat_tugas.docx', 25600, 1, 'Template untuk surat tugas dengan 3 pegawai');

-- Template untuk Banyak Pegawai (4+)
INSERT INTO `templates` 
(`nama`, `kategori_template`, `nama_file`, `path_file`, `ukuran_file`, `is_default`, `deskripsi`) 
VALUES 
('Template Surat Tugas - Banyak Pegawai', 'banyak_pegawai', 'template_surat_tugas.docx', 'assets/templates/template_surat_tugas.docx', 25600, 1, 'Template untuk surat tugas dengan 4 atau lebih pegawai');

-- ============================================
-- Verifikasi: Cek template yang sudah ditambahkan
-- ============================================
SELECT 
    id,
    nama,
    kategori_template,
    nama_file,
    ROUND(ukuran_file/1024, 2) as 'Ukuran (KB)',
    is_default as 'Default',
    created_at as 'Tanggal Dibuat'
FROM templates 
ORDER BY kategori_template, is_default DESC;

-- ============================================
-- CATATAN PENTING:
-- ============================================
-- 1. Ganti 'template_surat_tugas.docx' dengan nama file Anda
-- 2. Pastikan file sudah ada di folder assets/templates/
-- 3. Ukuran file (25600) adalah contoh, sesuaikan dengan ukuran file Anda
--    Untuk mendapatkan ukuran file dalam bytes:
--    - Windows: Klik kanan file > Properties > Size (bytes)
--    - Linux: ls -l nama_file.docx
-- 4. Satu file yang sama bisa digunakan untuk semua kategori
-- 5. is_default = 1 artinya template default untuk kategori tersebut
-- ============================================

-- ============================================
-- TROUBLESHOOTING
-- ============================================

-- Cek apakah ada template di database
SELECT COUNT(*) as 'Total Template' FROM templates;

-- Cek template per kategori
SELECT 
    kategori_template,
    COUNT(*) as 'Jumlah'
FROM templates 
GROUP BY kategori_template;

-- Cek kategori yang hilang
SELECT 'Kategori yang ada:' as 'Info';
SELECT DISTINCT kategori_template FROM templates;

-- Hapus semua template (HATI-HATI!)
-- DELETE FROM templates;

-- Hapus template tertentu berdasarkan ID
-- DELETE FROM templates WHERE id = 1;

-- Update template menjadi default
-- UPDATE templates SET is_default = 0 WHERE kategori_template = 'banyak_pegawai';
-- UPDATE templates SET is_default = 1 WHERE id = 1;

-- ============================================
-- QUICK FIX: Jika sudah ada 1 template, duplikat untuk kategori lain
-- ============================================
-- Contoh: Jika sudah ada template ID 1 dengan kategori 'banyak_pegawai',
-- dan ingin menggunakan file yang sama untuk kategori lain:

/*
INSERT INTO templates (nama, kategori_template, nama_file, path_file, ukuran_file, is_default, deskripsi)
SELECT 
    CONCAT('Template Surat Tugas - 1 Pegawai'),
    '1_pegawai',
    nama_file,
    path_file,
    ukuran_file,
    1,
    'Template untuk surat tugas dengan 1 pegawai (menggunakan file yang sama)'
FROM templates 
WHERE id = 1;

INSERT INTO templates (nama, kategori_template, nama_file, path_file, ukuran_file, is_default, deskripsi)
SELECT 
    CONCAT('Template Surat Tugas - 2 Pegawai'),
    '2_pegawai',
    nama_file,
    path_file,
    ukuran_file,
    1,
    'Template untuk surat tugas dengan 2 pegawai (menggunakan file yang sama)'
FROM templates 
WHERE id = 1;

INSERT INTO templates (nama, kategori_template, nama_file, path_file, ukuran_file, is_default, deskripsi)
SELECT 
    CONCAT('Template Surat Tugas - 3 Pegawai'),
    '3_pegawai',
    nama_file,
    path_file,
    ukuran_file,
    1,
    'Template untuk surat tugas dengan 3 pegawai (menggunakan file yang sama)'
FROM templates 
WHERE id = 1;
*/
