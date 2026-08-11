<?php
$pdo = require_once '../config/database.php';

if (isset($_GET['id'])) {
    try {
        $stmt = $pdo->prepare("SELECT nip FROM pegawai_tugas WHERE id_surat_tugas = ? ORDER BY urutan");
        $stmt->execute([$_GET['id']]);
        $result = $stmt->fetchAll(PDO::FETCH_COLUMN);

        header('Content-Type: application/json');
        echo json_encode($result);
    } catch (PDOException $e) {
        header('HTTP/1.1 500 Internal Server Error');
        echo json_encode(['error' => $e->getMessage()]);
    }
}