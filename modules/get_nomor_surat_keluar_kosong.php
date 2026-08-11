<?php
header('Content-Type: application/json; charset=utf-8');

$pdo = require_once '../config/database.php';

try {
    $tanggal = isset($_GET['tanggal_surat']) ? trim((string)$_GET['tanggal_surat']) : '';
    if ($tanggal === '') {
        throw new Exception('Tanggal surat wajib diisi.');
    }

    $ts = strtotime($tanggal);
    if ($ts === false) {
        throw new Exception('Format tanggal tidak valid.');
    }

    $tahun = (int)date('Y', $ts);
    $tanggalYmd = date('Y-m-d', $ts);

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

    $stmtLastTanggal = $pdo->prepare("SELECT COALESCE(MAX(nomor_urut), 0)
        FROM buku_nomor_surat_keluar
        WHERE tahun = :tahun AND tanggal_surat = :tanggal");
    $stmtLastTanggal->execute([
        ':tahun' => $tahun,
        ':tanggal' => $tanggalYmd,
    ]);
    $lastNomorTanggal = (int)$stmtLastTanggal->fetchColumn();

    $stmtLastYear = $pdo->prepare("SELECT COALESCE(MAX(nomor_urut), 0)
        FROM buku_nomor_surat_keluar
        WHERE tahun = :tahun");
    $stmtLastYear->execute([':tahun' => $tahun]);
    $lastNomorTahun = (int)$stmtLastYear->fetchColumn();

    $stmt = $pdo->prepare("SELECT nomor_urut, status, tanggal_surat
        FROM buku_nomor_surat_keluar
        WHERE tahun = :tahun
          AND tanggal_surat = :tanggal
          AND status = 'kosong'
          AND id_surat_keluar IS NULL
        ORDER BY nomor_urut ASC");
    $stmt->execute([
        ':tahun' => $tahun,
        ':tanggal' => $tanggalYmd,
    ]);

    echo json_encode([
        'success' => true,
        'tanggal_surat' => $tanggalYmd,
        'tahun' => $tahun,
        'last_nomor_tanggal' => $lastNomorTanggal,
        'last_nomor_tahun' => $lastNomorTahun,
        'nomor' => $stmt->fetchAll(PDO::FETCH_ASSOC),
    ]);
} catch (Throwable $e) {
    echo json_encode([
        'success' => false,
        'message' => $e->getMessage(),
        'nomor' => [],
    ]);
}
