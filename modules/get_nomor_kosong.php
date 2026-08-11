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

    // Informasi nomor terakhir pada tanggal bersangkutan
    $stmtLast = $pdo->prepare("SELECT COALESCE(MAX(CAST(nomor_surat AS UNSIGNED)), 0) AS last_no
        FROM buku_nomor_surat_tugas
        WHERE tahun = :tahun
          AND tanggal_surat = :tanggal
          AND nomor_surat REGEXP '^[0-9]+$'");
    $stmtLast->execute([
        ':tahun' => $tahun,
        ':tanggal' => $tanggalYmd,
    ]);
    $lastNomorTanggal = (int)$stmtLast->fetchColumn();

    $stmtLastYear = $pdo->prepare("SELECT COALESCE(MAX(CAST(nomor_surat AS UNSIGNED)), 0) AS last_no
        FROM buku_nomor_surat_tugas
        WHERE tahun = :tahun
          AND nomor_surat REGEXP '^[0-9]+$'");
    $stmtLastYear->execute([':tahun' => $tahun]);
    $lastNomorTahun = (int)$stmtLastYear->fetchColumn();

    // Hanya slot KOSONG untuk tanggal yang sama
    $stmtTanggal = $pdo->prepare("SELECT nomor_surat, status, tanggal_surat
        FROM buku_nomor_surat_tugas
        WHERE tahun = :tahun
          AND tanggal_surat = :tanggal
          AND status = 'kosong'
          AND id_surat_tugas IS NULL
        ORDER BY CAST(nomor_surat AS UNSIGNED) ASC, nomor_surat ASC");
    $stmtTanggal->execute([
        ':tahun' => $tahun,
        ':tanggal' => $tanggalYmd,
    ]);
    $rowsTanggal = $stmtTanggal->fetchAll(PDO::FETCH_ASSOC);

    $nomor = [];
    foreach ($rowsTanggal as $r) {
        $nomor[] = [
            'nomor_surat' => $r['nomor_surat'],
            'tipe' => 'tanggal',
            'status' => $r['status'],
            'tanggal_slot' => $r['tanggal_surat'],
        ];
    }

    echo json_encode([
        'success' => true,
        'tanggal_surat' => $tanggalYmd,
        'tahun' => $tahun,
        'last_nomor_tanggal' => $lastNomorTanggal,
        'last_nomor_tahun' => $lastNomorTahun,
        'nomor' => $nomor,
    ]);
} catch (Throwable $e) {
    echo json_encode([
        'success' => false,
        'message' => $e->getMessage(),
        'nomor' => [],
    ]);
}

