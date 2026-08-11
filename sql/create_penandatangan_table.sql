-- ============================================
-- Tabel Penanda Tangan Surat Tugas
-- ============================================
-- Menyimpan data pejabat yang berhak menandatangani surat tugas

CREATE TABLE IF NOT EXISTS penandatangan (
    id INT AUTO_INCREMENT PRIMARY KEY,
    nama VARCHAR(255) NOT NULL,
    nip VARCHAR(30) NOT NULL,
    pangkat VARCHAR(100) DEFAULT NULL,
    jabatan VARCHAR(255) NOT NULL,
    is_kepala TINYINT(1) DEFAULT 1,
    jabatan_atasan VARCHAR(255) DEFAULT NULL,
    tanda_tangan VARCHAR(255) DEFAULT NULL,
    is_default TINYINT(1) DEFAULT 0,
    aktif TINYINT(1) DEFAULT 1,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX idx_aktif (aktif),
    INDEX idx_default (is_default)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- ============================================
-- Tambahkan kolom id_penandatangan ke surat_tugas
-- ============================================
ALTER TABLE surat_tugas 
ADD COLUMN id_penandatangan INT DEFAULT NULL AFTER tanggal_selesai,
ADD CONSTRAINT fk_penandatangan FOREIGN KEY (id_penandatangan) REFERENCES penandatangan(id) ON DELETE SET NULL;
