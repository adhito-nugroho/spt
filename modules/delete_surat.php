<?php
$pdo = require_once '../config/database.php';

$response = ['success' => false, 'message' => ''];

if (isset($_POST['id'])) {
    try {
        $pdo->beginTransaction();

        // Lepaskan nomor dari buku nomor (jadikan slot kosong kembali)
        try {
            $stmt = $pdo->prepare("UPDATE buku_nomor_surat_tugas
                SET id_surat_tugas = NULL,
                    status = 'kosong',
                    keterangan = 'Dikosongkan karena surat dihapus'
                WHERE id_surat_tugas = ?");
            $stmt->execute([$_POST['id']]);
        } catch (PDOException $e) {
            // Abaikan bila tabel buku nomor belum ada
        }

        // Hapus data dari pegawai_tugas
        $stmt = $pdo->prepare("DELETE FROM pegawai_tugas WHERE id_surat_tugas = ?");
        $stmt->execute([$_POST['id']]);

        // Hapus data dari surat_tugas 
        $stmt = $pdo->prepare("DELETE FROM surat_tugas WHERE id = ?");
        $stmt->execute([$_POST['id']]);

        $pdo->commit();

        $response['success'] = true;
        $response['message'] = 'Data berhasil dihapus';

    } catch (PDOException $e) {
        $pdo->rollBack();
        $response['message'] = $e->getMessage();
    }
}

header('Content-Type: application/json');
echo json_encode($response);