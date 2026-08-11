<?php
function checkPegawaiAvailability($pdo, $nip, $tanggal_mulai, $tanggal_selesai, $id_surat_tugas = null)
{
    try {
        $params = [
            ':nip' => $nip,
            ':tanggal_mulai' => $tanggal_mulai,
            ':tanggal_selesai' => $tanggal_selesai
        ];

        $sql = "SELECT st.* 
                FROM surat_tugas st 
                JOIN pegawai_tugas pt ON st.id = pt.id_surat_tugas 
                WHERE pt.nip = :nip 
                AND (
                    (st.tanggal_mulai BETWEEN :tanggal_mulai AND :tanggal_selesai)
                    OR (st.tanggal_selesai BETWEEN :tanggal_mulai AND :tanggal_selesai)
                    OR (:tanggal_mulai BETWEEN st.tanggal_mulai AND st.tanggal_selesai)
                )";

        // Jika update, exclude surat tugas yang sedang diedit
        if ($id_surat_tugas) {
            $sql .= " AND st.id != :id_surat_tugas";
            $params[':id_surat_tugas'] = $id_surat_tugas;
        }

        $stmt = $pdo->prepare($sql);
        $stmt->execute($params);

        $conflicts = $stmt->fetchAll();

        if (count($conflicts) > 0) {
            // Return detail konflik
            return [
                'available' => false,
                'conflicts' => array_map(function ($conflict) {
                    return [
                        'nomor_surat' => $conflict['nomor_surat'],
                        'tanggal_mulai' => $conflict['tanggal_mulai'],
                        'tanggal_selesai' => $conflict['tanggal_selesai']
                    ];
                }, $conflicts)
            ];
        }

        return ['available' => true];

    } catch (PDOException $e) {
        return [
            'available' => false,
            'error' => $e->getMessage()
        ];
    }
}