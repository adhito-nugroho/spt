-- Update tabel templates untuk menambahkan kolom kategori_template
ALTER TABLE templates 
ADD COLUMN kategori_template ENUM('1_pegawai', '2_pegawai', '3_pegawai', 'banyak_pegawai') NOT NULL DEFAULT 'banyak_pegawai' AFTER nama;

-- Update index untuk memudahkan query berdasarkan kategori
ALTER TABLE templates 
ADD INDEX idx_kategori (kategori_template, is_default);

