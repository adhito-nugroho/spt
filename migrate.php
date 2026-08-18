<?php
/**
 * SPT Database Migration Runner
 * Dapat dijalankan via CLI: php migrate.php
 * Atau via Browser: http://localhost/spt-php/migrate.php
 */

if (php_sapi_name() !== 'cli') {
    header('Content-Type: text/plain; charset=utf-8');
}

echo "======================================================\n";
echo " SPT | Surat Perintah Tugas — Database Migration\n";
echo "======================================================\n\n";

try {
    $pdo = require_once __DIR__ . '/config/database.php';
    echo "✓ Terhubung ke database dengan sukses.\n\n";

    // 1. Migration: Kolom urutan pada tabel pegawai_tugas
    echo "[1/4] Memeriksa tabel 'pegawai_tugas'...\n";
    $stmt = $pdo->query("SHOW COLUMNS FROM pegawai_tugas LIKE 'urutan'");
    $colUrutan = $stmt->fetch();
    if (!$colUrutan) {
        $pdo->exec("ALTER TABLE pegawai_tugas ADD COLUMN urutan INT NOT NULL DEFAULT 0 AFTER nip");
        echo "  ✓ Kolom 'urutan' berhasil ditambahkan ke tabel pegawai_tugas.\n";
        
        $stmt = $pdo->query("SELECT DISTINCT id_surat_tugas FROM pegawai_tugas ORDER BY id_surat_tugas");
        $suratTugasIds = $stmt->fetchAll(PDO::FETCH_COLUMN);
        foreach ($suratTugasIds as $idSurat) {
            $stmtP = $pdo->prepare("SELECT id FROM pegawai_tugas WHERE id_surat_tugas = ? ORDER BY id");
            $stmtP->execute([$idSurat]);
            $pegawaiIds = $stmtP->fetchAll(PDO::FETCH_COLUMN);
            $urutan = 1;
            $updateStmt = $pdo->prepare("UPDATE pegawai_tugas SET urutan = ? WHERE id = ?");
            foreach ($pegawaiIds as $pegawaiId) {
                $updateStmt->execute([$urutan, $pegawaiId]);
                $urutan++;
            }
        }
        echo "  ✓ Urutan pegawai lama berhasil diinisialisasi.\n";
    } else {
        echo "  ℹ Kolom 'urutan' sudah ada.\n";
    }

    // 2. Migration: Kolom file_pdf pada tabel surat_keluar
    echo "\n[2/4] Memeriksa tabel 'surat_keluar'...\n";
    $stmt = $pdo->query("SHOW TABLES LIKE 'surat_keluar'");
    if ($stmt->fetch()) {
        $stmtCol = $pdo->query("SHOW COLUMNS FROM surat_keluar LIKE 'file_pdf'");
        if (!$stmtCol->fetch()) {
            $pdo->exec("ALTER TABLE surat_keluar ADD COLUMN file_pdf VARCHAR(255) DEFAULT NULL AFTER keterangan");
            echo "  ✓ Kolom 'file_pdf' berhasil ditambahkan ke tabel surat_keluar.\n";
        } else {
            echo "  ℹ Kolom 'file_pdf' sudah ada pada tabel surat_keluar.\n";
        }
    } else {
        echo "  ℹ Tabel 'surat_keluar' belum ada / dilewati.\n";
    }

    // 3. Migration: Tabel buku_nomor_surat_tugas
    echo "\n[3/4] Memeriksa tabel 'buku_nomor_surat_tugas'...\n";
    $pdo->exec("CREATE TABLE IF NOT EXISTS buku_nomor_surat_tugas (
        id INT AUTO_INCREMENT PRIMARY KEY,
        tahun INT NOT NULL,
        nomor_surat VARCHAR(50) NOT NULL,
        tanggal_surat DATE DEFAULT NULL,
        status ENUM('kosong','cadangan','terisi','batal') NOT NULL DEFAULT 'kosong',
        id_surat_tugas INT DEFAULT NULL,
        keterangan VARCHAR(255) DEFAULT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        UNIQUE KEY uniq_tahun_nomor (tahun, nomor_surat),
        UNIQUE KEY uniq_id_surat_tugas (id_surat_tugas),
        INDEX idx_status_tanggal (status, tanggal_surat)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci");

    // Indexes
    $indexStatements = [
        "ALTER TABLE buku_nomor_surat_tugas ADD INDEX idx_buku_tahun_status_nomor (tahun, status, nomor_surat)",
        "ALTER TABLE buku_nomor_surat_tugas ADD INDEX idx_buku_tanggal (tanggal_surat)",
    ];
    foreach ($indexStatements as $sql) {
        try {
            $pdo->exec($sql);
        } catch (PDOException $e) {}
    }
    echo "  ✓ Tabel 'buku_nomor_surat_tugas' siap.\n";

    // 4. Migration: Tabel buku_nomor_surat_keluar
    echo "\n[4/4] Memeriksa tabel 'buku_nomor_surat_keluar'...\n";
    $pdo->exec("CREATE TABLE IF NOT EXISTS buku_nomor_surat_keluar (
        id INT AUTO_INCREMENT PRIMARY KEY,
        tahun INT NOT NULL,
        nomor_urut INT NOT NULL,
        tanggal_surat DATE DEFAULT NULL,
        status ENUM('kosong','terisi') NOT NULL DEFAULT 'kosong',
        id_surat_keluar INT DEFAULT NULL,
        keterangan VARCHAR(255) DEFAULT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        UNIQUE KEY uniq_tahun_nomor_urut (tahun, nomor_urut),
        UNIQUE KEY uniq_id_surat_keluar (id_surat_keluar),
        INDEX idx_status_tanggal (status, tanggal_surat)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci");
    echo "  ✓ Tabel 'buku_nomor_surat_keluar' siap.\n";

    echo "\n======================================================\n";
    echo " SEMUA MIGRASI DATABASE BERHASIL DIJALANKAN (100% OK)\n";
    echo "======================================================\n";

} catch (Throwable $e) {
    echo "\n[ERROR] Migrasi gagal: " . $e->getMessage() . "\n";
    exit(1);
}
