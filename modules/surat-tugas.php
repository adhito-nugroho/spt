<?php
error_reporting(E_ALL);
ini_set('display_errors', 1);

$pdo = require_once '../config/database.php';
require_once __DIR__ . '/check_pegawai_availability.php';

function ensureBukuNomorSuratTugasTable(PDO $pdo): void
{
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
}

function normalizeBukuNomorStatus(PDO $pdo): void
{
    // Menyederhanakan status operasional menjadi: kosong / terisi
    $pdo->exec("UPDATE buku_nomor_surat_tugas
        SET status = 'kosong'
        WHERE status IN ('cadangan','batal') AND id_surat_tugas IS NULL");
}

function ensurePerformanceIndexes(PDO $pdo): void
{
    // Index untuk query besar (safe, jika sudah ada akan diabaikan via catch)
    $indexStatements = [
        "ALTER TABLE surat_tugas ADD INDEX idx_surat_tanggal (tanggal_surat)",
        "ALTER TABLE surat_tugas ADD INDEX idx_surat_nomor_tanggal (nomor_surat, tanggal_surat)",
        "ALTER TABLE pegawai_tugas ADD INDEX idx_pt_surat_urutan (id_surat_tugas, urutan)",
        "ALTER TABLE pegawai_tugas ADD INDEX idx_pt_nip (nip)",
    ];

    foreach ($indexStatements as $sql) {
        try {
            $pdo->exec($sql);
        } catch (PDOException $e) {
            // Ignore jika index sudah ada / belum bisa dibuat
        }
    }
}

function syncSuratTugasTerpakaiKeBukuNomor(PDO $pdo): void
{
    // Incremental sync: hanya data surat_tugas yang belum tercatat via id_surat_tugas
    $pdo->exec("INSERT INTO buku_nomor_surat_tugas (
            tahun, nomor_surat, tanggal_surat, status, id_surat_tugas, keterangan
        )
        SELECT
            YEAR(st.tanggal_surat) AS tahun,
            st.nomor_surat,
            st.tanggal_surat,
            'terisi' AS status,
            st.id AS id_surat_tugas,
            'Sinkron dari surat_tugas' AS keterangan
        FROM surat_tugas st
        LEFT JOIN buku_nomor_surat_tugas b ON b.id_surat_tugas = st.id
        WHERE st.nomor_surat IS NOT NULL AND st.nomor_surat <> ''
          AND b.id IS NULL
        ON DUPLICATE KEY UPDATE
            tanggal_surat = VALUES(tanggal_surat),
            status = 'terisi',
            id_surat_tugas = VALUES(id_surat_tugas),
            keterangan = VALUES(keterangan)");
}

function bindNomorKeSuratTugas(PDO $pdo, int $idSuratTugas, string $nomorSurat, string $tanggalSurat): void
{
    $nomorSurat = trim($nomorSurat);
    if ($nomorSurat === '') {
        throw new Exception('Nomor surat kosong, tidak bisa dicatat ke buku nomor.');
    }

    $tahun = (int)date('Y', strtotime($tanggalSurat));
    if ($tahun < 1900) {
        throw new Exception('Tanggal surat tidak valid untuk buku nomor.');
    }

    // Lepas binding lama nomor yang sebelumnya terkait surat ini (jika ganti nomor)
    $stmtRelease = $pdo->prepare("UPDATE buku_nomor_surat_tugas
        SET id_surat_tugas = NULL, status = 'kosong', keterangan = 'Nomor dikosongkan setelah update'
        WHERE id_surat_tugas = :id_surat_tugas
          AND NOT (tahun = :tahun AND nomor_surat = :nomor_surat)");
    $stmtRelease->execute([
        ':id_surat_tugas' => $idSuratTugas,
        ':tahun' => $tahun,
        ':nomor_surat' => $nomorSurat,
    ]);

    $tanggalSurat = date('Y-m-d', strtotime($tanggalSurat));

    // Cek slot nomor pada buku nomor
    $stmtCek = $pdo->prepare("SELECT id_surat_tugas, status
        , tanggal_surat
        FROM buku_nomor_surat_tugas
        WHERE tahun = :tahun AND nomor_surat = :nomor_surat
        LIMIT 1");
    $stmtCek->execute([
        ':tahun' => $tahun,
        ':nomor_surat' => $nomorSurat,
    ]);
    $row = $stmtCek->fetch(PDO::FETCH_ASSOC);

    if (!$row) {
        throw new Exception('Nomor belum dibuat di Buku Nomor untuk tanggal ini.');
    }

    if (!empty($row['tanggal_surat']) && $row['tanggal_surat'] !== $tanggalSurat) {
        throw new Exception('Nomor ini diperuntukkan untuk tanggal surat lain dan tidak bisa dipakai lintas tanggal.');
    }

    if (!empty($row['id_surat_tugas']) && (int)$row['id_surat_tugas'] !== $idSuratTugas) {
        throw new Exception('Nomor surat sudah terpakai oleh surat tugas lain pada tahun yang sama.');
    }

    if ($row['status'] !== 'kosong' && (int)($row['id_surat_tugas'] ?? 0) !== $idSuratTugas) {
        throw new Exception('Status nomor bukan kosong, tidak dapat dipakai.');
    }

    // Isi/replace slot buku nomor sebagai terisi
    $stmtBind = $pdo->prepare("UPDATE buku_nomor_surat_tugas
        SET status = 'terisi',
            id_surat_tugas = :id_surat_tugas,
            keterangan = 'Terbit dari aplikasi'
        WHERE tahun = :tahun
          AND nomor_surat = :nomor_surat");
    $stmtBind->execute([
        ':tahun' => $tahun,
        ':nomor_surat' => $nomorSurat,
        ':id_surat_tugas' => $idSuratTugas,
    ]);
}

ensureBukuNomorSuratTugasTable($pdo);
normalizeBukuNomorStatus($pdo);
ensurePerformanceIndexes($pdo);
syncSuratTugasTerpakaiKeBukuNomor($pdo);

// Debug query pegawai
$stmt = $pdo->query("SELECT * FROM pegawai ORDER BY nama");
$pegawai = $stmt->fetchAll();

// Pastikan tabel penandatangan ada
try {
    $pdo->query("SELECT 1 FROM penandatangan LIMIT 1");
}
catch (PDOException $e) {
    $sql = "CREATE TABLE IF NOT EXISTS penandatangan (
        id INT AUTO_INCREMENT PRIMARY KEY,
        nama VARCHAR(255) NOT NULL,
        nip VARCHAR(30) NOT NULL,
        pangkat VARCHAR(100) DEFAULT NULL,
        jabatan VARCHAR(255) NOT NULL,
        is_default TINYINT(1) DEFAULT 0,
        aktif TINYINT(1) DEFAULT 1,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        INDEX idx_aktif (aktif),
        INDEX idx_default (is_default)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci";
    $pdo->exec($sql);
}

// Pastikan kolom id_penandatangan ada di surat_tugas
try {
    $pdo->query("SELECT id_penandatangan FROM surat_tugas LIMIT 1");
}
catch (PDOException $e) {
    try {
        $pdo->exec("ALTER TABLE surat_tugas ADD COLUMN id_penandatangan INT DEFAULT NULL AFTER tanggal_selesai");
    }
    catch (PDOException $e2) {
    }
}

// Pastikan kolom tipe_surat ada di surat_tugas
try {
    $pdo->query("SELECT tipe_surat FROM surat_tugas LIMIT 1");
}
catch (PDOException $e) {
    try {
        $pdo->exec("ALTER TABLE surat_tugas ADD COLUMN tipe_surat VARCHAR(20) NOT NULL DEFAULT 'umum' AFTER id_penandatangan");
    }
    catch (PDOException $e2) {
    }
}

// Ambil data penandatangan yang aktif
$stmt_pt = $pdo->query("SELECT * FROM penandatangan WHERE aktif = 1 ORDER BY is_default DESC, nama ASC");
$penandatangan_list = $stmt_pt->fetchAll();

// Konfigurasi Pencarian
$search = isset($_GET['search']) ? $_GET['search'] : '';
$search_condition = '';

// Konfigurasi Pagination
$limit = 10;
$page = isset($_GET['page']) ? $_GET['page'] : 1;
$start = ($page - 1) * $limit;

// Persiapkan query sesuai apakah ada pencarian atau tidak
if (!empty($search)) {
    $search_condition = "WHERE st.nomor_surat LIKE :search OR st.untuk LIKE :search OR p.nama LIKE :search";
}

// Hitung total data dengan kondisi pencarian
$query_count = "SELECT COUNT(DISTINCT st.id) AS total 
                FROM surat_tugas st 
                LEFT JOIN pegawai_tugas pt ON st.id = pt.id_surat_tugas 
                LEFT JOIN pegawai p ON pt.nip = p.nip 
                $search_condition";
$stmt_count = $pdo->prepare($query_count);
if (!empty($search)) {
    $stmt_count->bindValue(':search', "%$search%", PDO::PARAM_STR);
}
$stmt_count->execute();
$total_records = $stmt_count->fetch()['total'];
$total_pages = ceil($total_records / $limit);

// Ambil data surat tugas dengan pagination dan detail pegawai
$query = "SELECT 
    st.*,
    GROUP_CONCAT(DISTINCT p.nama ORDER BY pt.urutan SEPARATOR ', ') as pegawai_names,
    GROUP_CONCAT(DISTINCT CONCAT(p.nip, ':::', p.nama, ':::', p.jabatan) ORDER BY pt.urutan SEPARATOR '|||') as pegawai_details 
    FROM surat_tugas st 
    LEFT JOIN pegawai_tugas pt ON st.id = pt.id_surat_tugas 
    LEFT JOIN pegawai p ON pt.nip = p.nip 
    $search_condition
    GROUP BY st.id 
    ORDER BY st.tanggal_surat DESC
    LIMIT :start, :limit";
$stmt = $pdo->prepare($query);
$stmt->bindValue(':start', $start, PDO::PARAM_INT);
$stmt->bindValue(':limit', $limit, PDO::PARAM_INT);
if (!empty($search)) {
    $stmt->bindValue(':search', "%$search%", PDO::PARAM_STR);
}
$stmt->execute();
$surat_tugas = $stmt->fetchAll();

// Proses Form Submit
if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_POST['action'])) {
    if ($_POST['action'] == 'create') {
        try {
            // Validasi input wajib
            if (empty($_POST['nomor_surat'])) {
                throw new Exception("Nomor surat harus diisi.");
            }
            if (empty($_POST['tanggal_surat'])) {
                throw new Exception("Tanggal surat harus diisi.");
            }
            if (empty($_POST['tanggal_mulai'])) {
                throw new Exception("Tanggal mulai harus diisi.");
            }
            if (empty($_POST['tanggal_selesai'])) {
                throw new Exception("Tanggal selesai harus diisi.");
            }
            if (empty($_POST['untuk'])) {
                throw new Exception("Tujuan/maksud surat harus diisi.");
            }

            // Validasi pegawai
            if (!isset($_POST['pegawai']) || !is_array($_POST['pegawai']) || empty($_POST['pegawai'])) {
                throw new Exception("Minimal harus memilih satu pegawai.");
            }

            // Cek apakah nomor surat dan tanggal surat sudah ada
            $stmt = $pdo->prepare("SELECT COUNT(*) FROM surat_tugas WHERE nomor_surat = ? AND tanggal_surat = ?");
            $stmt->execute([$_POST['nomor_surat'], $_POST['tanggal_surat']]);
            $count = $stmt->fetchColumn();

            if ($count > 0) {
                throw new Exception("Nomor surat dan tanggal surat sudah digunakan. Silakan gunakan nomor surat lain.");
            }

            // Check availability untuk setiap pegawai
            $unavailable_pegawai = [];
            $check_errors = [];

            foreach ($_POST['pegawai'] as $nip) {
                if (empty($nip)) {
                    continue; // Skip jika NIP kosong
                }

                $check = checkPegawaiAvailability(
                    $pdo,
                    $nip,
                    $_POST['tanggal_mulai'],
                    $_POST['tanggal_selesai']
                );

                // Cek jika ada error dari checkPegawaiAvailability
                if (isset($check['error'])) {
                    $check_errors[] = "Error saat mengecek ketersediaan pegawai: " . $check['error'];
                    continue;
                }

                if (!$check['available']) {
                    $stmt = $pdo->prepare("SELECT nama FROM pegawai WHERE nip = ?");
                    $stmt->execute([$nip]);
                    $pegawai_data = $stmt->fetch();

                    if ($pegawai_data) {
                        $unavailable_pegawai[] = [
                            'nama' => $pegawai_data['nama'],
                            'conflicts' => isset($check['conflicts']) ? $check['conflicts'] : []
                        ];
                    }
                }
            }

            // Jika ada error saat check availability
            if (!empty($check_errors)) {
                throw new Exception("Terjadi kesalahan saat mengecek ketersediaan pegawai:\n" . implode("\n", $check_errors));
            }

            if (!empty($unavailable_pegawai)) {
                $message = "Beberapa pegawai sudah memiliki tugas di tanggal yang sama:\\n\\n";
                foreach ($unavailable_pegawai as $p) {
                    $message .= "- {$p['nama']}:\\n";
                    if (isset($p['conflicts']) && is_array($p['conflicts'])) {
                        foreach ($p['conflicts'] as $conflict) {
                            $message .= "  • Surat {$conflict['nomor_surat']}: " .
                                date('d/m/Y', strtotime($conflict['tanggal_mulai'])) . " s/d " .
                                date('d/m/Y', strtotime($conflict['tanggal_selesai'])) . "\\n";
                        }
                    }
                }

                throw new Exception($message);
            }

            $pdo->beginTransaction();

            $id_penandatangan = !empty($_POST['id_penandatangan']) ? $_POST['id_penandatangan'] : null;

            $stmt = $pdo->prepare("INSERT INTO surat_tugas (nomor_surat, tanggal_surat, dasar_surat, untuk, tanggal_mulai, tanggal_selesai, id_penandatangan, tipe_surat) 
                                    VALUES (:nomor_surat, :tanggal_surat, :dasar_surat, :untuk, :tanggal_mulai, :tanggal_selesai, :id_penandatangan, :tipe_surat)");

            $stmt->execute([
                ':nomor_surat' => $_POST['nomor_surat'],
                ':tanggal_surat' => $_POST['tanggal_surat'],
                ':dasar_surat' => $_POST['dasar_surat'],
                ':untuk' => $_POST['untuk'],
                ':tanggal_mulai' => $_POST['tanggal_mulai'],
                ':tanggal_selesai' => $_POST['tanggal_selesai'],
                ':id_penandatangan' => $id_penandatangan,
                ':tipe_surat' => isset($_POST['tipe_surat']) ? $_POST['tipe_surat'] : 'umum'
            ]);

            $id_surat_tugas = $pdo->lastInsertId();
            bindNomorKeSuratTugas($pdo, (int)$id_surat_tugas, (string)$_POST['nomor_surat'], (string)$_POST['tanggal_surat']);

            // Insert pegawai ke pegawai_tugas
            if (isset($_POST['pegawai']) && is_array($_POST['pegawai']) && !empty($_POST['pegawai'])) {
                $stmt = $pdo->prepare("INSERT INTO pegawai_tugas (id_surat_tugas, nip, urutan) VALUES (:id_surat_tugas, :nip, :urutan)");
                $inserted_count = 0;
                $urutan = 1; // Start order from 1

                foreach ($_POST['pegawai'] as $nip) {
                    if (!empty($nip)) {
                        try {
                            $stmt->execute([
                                ':id_surat_tugas' => $id_surat_tugas,
                                ':nip' => $nip,
                                ':urutan' => $urutan
                            ]);
                            $inserted_count++;
                            $urutan++; // Increment order for next employee
                        }
                        catch (PDOException $e) {
                            // Jika error duplicate atau constraint, skip
                            if ($e->getCode() != '23000') {
                                throw new Exception("Error saat menyimpan data pegawai: " . $e->getMessage());
                            }
                        }
                    }
                }

                if ($inserted_count == 0) {
                    throw new Exception("Tidak ada pegawai yang berhasil disimpan. Pastikan pegawai yang dipilih valid.");
                }
            }
            else {
                throw new Exception("Data pegawai tidak valid atau kosong.");
            }

            $pdo->commit();

            include '../includes/header.php';
            echo "<script>
                    Swal.fire({
                        title: 'Berhasil!',
                        text: 'Surat tugas berhasil dibuat',
                        icon: 'success',
                        confirmButtonText: 'OK'
                    }).then((result) => {
                        if (result.isConfirmed) {
                            window.location.href = 'surat-tugas.php';
                        }
                    });
                </script>";
            include '../includes/footer.php';
            exit;

        }
        catch (PDOException $e) {
            if ($pdo->inTransaction()) {
                $pdo->rollBack();
            }
            $error_message = "Terjadi kesalahan database: " . $e->getMessage();
            include '../includes/header.php';
            echo "<script>
                    Swal.fire({
                        title: 'Gagal!',
                        text: " . json_encode($error_message) . ",
                        icon: 'error',
                        confirmButtonText: 'OK',
                        width: '600px'
                    }).then((result) => {
                        if (result.isConfirmed) {
                            window.location.href = 'surat-tugas.php';
                        }
                    });
                </script>";
            include '../includes/footer.php';
            exit;
        }
        catch (Exception $e) {
            if ($pdo->inTransaction()) {
                $pdo->rollBack();
            }
            $error_message = $e->getMessage();
            include '../includes/header.php';
            echo "<script>
                    Swal.fire({
                        title: 'Gagal!',
                        text: " . json_encode($error_message) . ",
                        icon: 'error',
                        confirmButtonText: 'OK',
                        width: '600px'
                    }).then((result) => {
                        if (result.isConfirmed) {
                            window.location.href = 'surat-tugas.php';
                        }
                    });
                </script>";
            include '../includes/footer.php';
            exit;
        }
    }
    else if ($_POST['action'] == 'update') {
        try {
            $stmt = $pdo->prepare("SELECT COUNT(*) FROM surat_tugas WHERE nomor_surat = ? AND tanggal_surat = ? AND id != ?");
            $stmt->execute([$_POST['nomor_surat'], $_POST['tanggal_surat'], $_POST['id']]);
            $count = $stmt->fetchColumn();

            if ($count > 0) {
                throw new Exception("Nomor surat dan tanggal surat sudah digunakan. Silakan gunakan nomor surat lain.");
            }

            $unavailable_pegawai = [];
            foreach ($_POST['pegawai'] as $nip) {
                $check = checkPegawaiAvailability(
                    $pdo,
                    $nip,
                    $_POST['tanggal_mulai'],
                    $_POST['tanggal_selesai'],
                    $_POST['id']
                );

                if (!$check['available']) {
                    $stmt = $pdo->prepare("SELECT nama FROM pegawai WHERE nip = ?");
                    $stmt->execute([$nip]);
                    $pegawai_data = $stmt->fetch();

                    $unavailable_pegawai[] = [
                        'nama' => $pegawai_data['nama'],
                        'conflicts' => $check['conflicts']
                    ];
                }
            }

            if (!empty($unavailable_pegawai)) {
                $message = "Beberapa pegawai sudah memiliki tugas di tanggal yang sama:\\n\\n";
                foreach ($unavailable_pegawai as $p) {
                    $message .= "- {$p['nama']}:\\n";
                    foreach ($p['conflicts'] as $conflict) {
                        $message .= "  • Surat {$conflict['nomor_surat']}: " .
                            date('d/m/Y', strtotime($conflict['tanggal_mulai'])) . " s/d " .
                            date('d/m/Y', strtotime($conflict['tanggal_selesai'])) . "\\n";
                    }
                }
                throw new Exception($message);
            }

            $pdo->beginTransaction();

            $id_penandatangan_edit = !empty($_POST['id_penandatangan']) ? $_POST['id_penandatangan'] : null;

            $stmt = $pdo->prepare("UPDATE surat_tugas SET 
                    nomor_surat = :nomor_surat,
                    tanggal_surat = :tanggal_surat,
                    dasar_surat = :dasar_surat,
                    untuk = :untuk,
                    tanggal_mulai = :tanggal_mulai,
                    tanggal_selesai = :tanggal_selesai,
                    id_penandatangan = :id_penandatangan,
                    tipe_surat = :tipe_surat
                    WHERE id = :id");

            $stmt->execute([
                ':id' => $_POST['id'],
                ':nomor_surat' => $_POST['nomor_surat'],
                ':tanggal_surat' => $_POST['tanggal_surat'],
                ':dasar_surat' => $_POST['dasar_surat'],
                ':untuk' => $_POST['untuk'],
                ':tanggal_mulai' => $_POST['tanggal_mulai'],
                ':tanggal_selesai' => $_POST['tanggal_selesai'],
                ':id_penandatangan' => $id_penandatangan_edit,
                ':tipe_surat' => isset($_POST['tipe_surat']) ? $_POST['tipe_surat'] : 'umum'
            ]);
            bindNomorKeSuratTugas($pdo, (int)$_POST['id'], (string)$_POST['nomor_surat'], (string)$_POST['tanggal_surat']);

            $stmt = $pdo->prepare("DELETE FROM pegawai_tugas WHERE id_surat_tugas = ?");
            $stmt->execute([$_POST['id']]);

            if (isset($_POST['pegawai']) && is_array($_POST['pegawai'])) {
                $stmt = $pdo->prepare("INSERT INTO pegawai_tugas (id_surat_tugas, nip, urutan) VALUES (:id_surat_tugas, :nip, :urutan)");
                $urutan = 1; // Start order from 1
                foreach ($_POST['pegawai'] as $nip) {
                    $stmt->execute([
                        ':id_surat_tugas' => $_POST['id'],
                        ':nip' => $nip,
                        ':urutan' => $urutan
                    ]);
                    $urutan++; // Increment order for next employee
                }
            }

            $pdo->commit();

            include '../includes/header.php';
            $redirect_url = (isset($_POST['after_save']) && $_POST['after_save'] === 'preview')
                ? 'generate-surat.php?id=' . (int)$_POST['id']
                : 'surat-tugas.php';
            echo "<script>
                    Swal.fire({
                        title: 'Berhasil!',
                        text: 'Data surat tugas berhasil diupdate',
                        icon: 'success',
                        confirmButtonText: 'OK',
                    }).then((result) => {
                        if (result.isConfirmed) {
                            window.location.href = '" . $redirect_url . "';
                        }
                    });
                </script>";
            include '../includes/footer.php';
            exit;
        }
        catch (Exception $e) {
            if ($pdo->inTransaction()) {
                $pdo->rollBack();
            }
            $error_message = $e->getMessage();
            include '../includes/header.php';
            echo "<script>
                    Swal.fire({
                        title: 'Gagal!',
                        text: " . json_encode($error_message) . ",
                        icon: 'error',
                        confirmButtonText: 'OK',
                        width: '600px'
                    }).then((result) => {
                        if (result.isConfirmed) {
                            window.location.href = 'surat-tugas.php';
                        }
                    });
                </script>";
            include '../includes/footer.php';
            exit;
        }
    }
}

// Include header setelah semua POST handling selesai
include '../includes/header.php';
?>

<style>
    /* Select2 Customization for Tailwind */
    .select2-container .select2-selection--single {
        height: 42px !important;
        border-color: #e2e8f0 !important;
        border-radius: 0.5rem !important;
    }
    .select2-container--default .select2-selection--single .select2-selection__rendered {
        line-height: 40px !important;
        padding-left: 1rem !important;
        color: #1e293b !important;
    }
    .select2-container--default .select2-selection--single .select2-selection__arrow {
        height: 40px !important;
    }
    .select2-container--default.select2-container--focus .select2-selection--single {
        border-color: #6366f1 !important;
        box-shadow: 0 0 0 2px rgba(99, 102, 241, 0.2) !important;
    }
    
    /* Drag and Drop & List Styling */
    .draggable-item {
        transition: all 0.2s ease;
        user-select: none;
    }
    .draggable-item.opacity-50 {
        opacity: 0.5;
        background-color: #f8fafc;
        border-style: dashed;
    }
    .draggable-item:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1);
    }
</style>

<div class="space-y-6">
    <div class="flex flex-col md:flex-row justify-between items-center gap-4">
        <h2 class="text-2xl font-bold text-slate-800">Surat Tugas</h2>
        <button onclick="openModal('addSuratModal')" class="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg flex items-center gap-2 transition-colors shadow-sm">
            <i class='bx bx-plus'></i> Buat Surat Tugas
        </button>
    </div>

    <!-- Form Pencarian -->
    <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-4">
        <form action="" method="GET" class="flex gap-2">
            <div class="relative flex-1">
                <i class='bx bx-search absolute left-3 top-1/2 transform -translate-y-1/2 text-slate-400'></i>
                <input type="text" name="search" class="w-full pl-10 pr-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all"
                    placeholder="Cari nomor surat, tujuan, atau pegawai"
                    value="<?php echo htmlspecialchars($search); ?>">
            </div>
            <button type="submit" class="border-2 border-indigo-600 text-indigo-600 hover:bg-indigo-50 px-6 py-2 rounded-lg transition-colors font-medium">
                Cari
            </button>
            <?php if (!empty($search)): ?>
                <a href="surat-tugas.php" class="bg-slate-200 hover:bg-slate-300 text-slate-700 px-4 py-2 rounded-lg transition-colors flex items-center gap-2">
                    <i class='bx bx-reset'></i> Reset
                </a>
            <?php endif; ?>
        </form>
        <!-- Filter Row -->
        <div class="flex flex-wrap items-center gap-3 mt-3 pt-3 border-t border-slate-100" id="filterRow">
            <div class="flex items-center gap-2">
                <label class="text-xs font-medium text-slate-500">Tipe:</label>
                <select id="filterTipe" class="text-sm border border-slate-300 rounded-lg px-3 py-1.5 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                    <option value="">Semua</option>
                    <option value="umum">Umum</option>
                    <option value="penyuluh">Penyuluh</option>
                </select>
            </div>
            <div class="flex items-center gap-2">
                <label class="text-xs font-medium text-slate-500">Dari:</label>
                <input type="date" id="filterDateFrom" class="text-sm border border-slate-300 rounded-lg px-3 py-1.5 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
            </div>
            <div class="flex items-center gap-2">
                <label class="text-xs font-medium text-slate-500">Sampai:</label>
                <input type="date" id="filterDateTo" class="text-sm border border-slate-300 rounded-lg px-3 py-1.5 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
            </div>
            <button type="button" id="btnResetFilter" class="hidden text-xs text-red-600 hover:text-red-700 font-medium px-3 py-1.5 border border-red-200 rounded-lg hover:bg-red-50 transition-colors">
                <i class='bx bx-reset mr-1'></i>Reset Filter
            </button>
        </div>
    </div>

    <!-- Tabel Surat Tugas -->
    <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
        <div class="overflow-x-auto">
            <?php if (count($surat_tugas) > 0): ?>
                <table class="w-full text-left border-collapse" id="tabelSurat">
                    <thead>
                        <tr class="bg-slate-50 text-slate-600 text-xs uppercase tracking-wider">
                            <th class="px-3 py-3 font-semibold border-b border-slate-100 w-8" title="Lihat detail">
                                <i class='bx bx-detail text-slate-400'></i>
                            </th>
                            <th class="px-3 py-3 font-semibold border-b border-slate-100">Nomor Surat</th>
                            <th class="px-3 py-3 font-semibold border-b border-slate-100 cursor-pointer select-none sortable-header" data-sort="tanggal" data-order="">
                                <span class="inline-flex items-center gap-1">Tanggal <i class='bx bx-sort-alt-2 text-slate-400 sort-icon'></i></span>
                            </th>
                            <th class="px-3 py-3 font-semibold border-b border-slate-100" style="min-width:240px;">Untuk</th>
                            <th class="px-3 py-3 font-semibold border-b border-slate-100">Pegawai</th>
                            <th class="px-3 py-3 font-semibold border-b border-slate-100 cursor-pointer select-none sortable-header" data-sort="pelaksanaan" data-order="">
                                <span class="inline-flex items-center gap-1">Pelaksanaan <i class='bx bx-sort-alt-2 text-slate-400 sort-icon'></i></span>
                            </th>
                            <th class="px-3 py-3 font-semibold border-b border-slate-100">Tipe</th>
                            <th class="px-3 py-3 font-semibold border-b border-slate-100 text-right">Aksi</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-100">
                        <?php
    $no = $start + 1;
    foreach ($surat_tugas as $st):
        $pegawai_array = explode('|||', $st['pegawai_details']);
        $jumlah_pegawai = count($pegawai_array);
        if ($pegawai_array[0] == '')
            $jumlah_pegawai = 0;
        
        // Siapkan daftar nama pegawai untuk tooltip
        $nama_pegawai_list = [];
        foreach ($pegawai_array as $pa) {
            if ($pa != '') {
                $parts = explode(':::', $pa);
                $nama_pegawai_list[] = $parts[1] ?? '';
            }
        }
        $tooltip_pegawai = '';
        if (count($nama_pegawai_list) <= 5) {
            $tooltip_pegawai = implode('&#10;', array_map('htmlspecialchars', $nama_pegawai_list));
        } else {
            $tooltip_pegawai = implode('&#10;', array_map('htmlspecialchars', array_slice($nama_pegawai_list, 0, 5)));
            $tooltip_pegawai .= '&#10;+ ' . (count($nama_pegawai_list) - 5) . ' lainnya';
        }
        
        $tipe_surat = isset($st['tipe_surat']) ? $st['tipe_surat'] : 'umum';
?>
                            <tr class="hover:bg-slate-50 transition-colors group surat-row" 
                                data-tipe="<?php echo htmlspecialchars($tipe_surat); ?>"
                                data-tanggal="<?php echo $st['tanggal_surat']; ?>"
                                data-mulai="<?php echo $st['tanggal_mulai']; ?>">
                                <td class="px-3 py-2">
                                    <button onclick="toggleDetails('details-<?php echo $st['id']; ?>')" 
                                        class="text-indigo-600 hover:text-indigo-800 transition-transform transform duration-200" 
                                        id="btn-details-<?php echo $st['id']; ?>"
                                        title="Lihat detail surat"
                                        aria-label="Lihat detail surat">
                                        <i class='bx bx-chevron-down text-lg transition-transform duration-200'></i>
                                    </button>
                                </td>
                                <td class="px-3 py-2 text-sm font-medium text-slate-900"><?php echo htmlspecialchars($st['nomor_surat']); ?></td>
                                <td class="px-3 py-2 text-sm text-slate-600"><?php echo date('d/m/Y', strtotime($st['tanggal_surat'])); ?></td>
                                <td class="px-3 py-2 text-sm text-slate-600">
                                    <span class="block truncate" style="max-width:320px;" title="<?php echo htmlspecialchars($st['untuk']); ?>">
                                        <?php echo htmlspecialchars(mb_strlen($st['untuk']) > 80 ? mb_substr($st['untuk'], 0, 80) . '...' : $st['untuk']); ?>
                                    </span>
                                </td>
                                <td class="px-3 py-2 text-sm text-slate-600">
                                    <span class="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-indigo-100 text-indigo-800 cursor-default"
                                        title="<?php echo $tooltip_pegawai; ?>">
                                        <?php echo $jumlah_pegawai; ?> Pegawai
                                    </span>
                                </td>
                                <td class="px-3 py-2 text-sm text-slate-600 whitespace-nowrap">
                                    <?php echo date('d/m/Y', strtotime($st['tanggal_mulai'])) . ' - ' . date('d/m/Y', strtotime($st['tanggal_selesai'])); ?>
                                </td>
                                <td class="px-3 py-2 text-sm">
                                    <?php if ($tipe_surat === 'penyuluh'): ?>
                                        <span class="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-emerald-100 text-emerald-800">
                                            <i class='bx bx-leaf mr-1'></i>Penyuluh
                                        </span>
                                    <?php else: ?>
                                        <span class="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-blue-100 text-blue-800">
                                            <i class='bx bx-briefcase mr-1'></i>Umum
                                        </span>
                                    <?php endif; ?>
                                </td>
                                <td class="px-3 py-2 text-sm text-right">
                                    <div class="inline-flex items-center gap-2">
                                        <button class="text-blue-500 hover:text-blue-600 transition-colors" title="Generate Surat" onclick="generateSurat(<?php echo $st['id']; ?>)">
                                            <i class='bx bx-file text-lg'></i>
                                        </button>
                                        <button class="text-amber-500 hover:text-amber-600 edit-btn transition-colors" title="Edit" 
                                            data-id="<?php echo $st['id']; ?>"
                                            data-nomor="<?php echo htmlspecialchars($st['nomor_surat']); ?>"
                                            data-tanggal="<?php echo $st['tanggal_surat']; ?>"
                                            data-dasar="<?php echo htmlspecialchars($st['dasar_surat']); ?>"
                                            data-untuk="<?php echo htmlspecialchars($st['untuk']); ?>"
                                            data-mulai="<?php echo $st['tanggal_mulai']; ?>"
                                            data-selesai="<?php echo $st['tanggal_selesai']; ?>"
                                            data-penandatangan="<?php echo isset($st['id_penandatangan']) ? $st['id_penandatangan'] : ''; ?>"
                                            data-tipe="<?php echo htmlspecialchars($tipe_surat); ?>">
                                            <i class='bx bx-edit text-lg'></i>
                                        </button>
                                        <button class="text-red-500 hover:text-red-600 delete-btn transition-colors" title="Hapus" data-id="<?php echo $st['id']; ?>">
                                            <i class='bx bx-trash text-lg'></i>
                                        </button>
                                    </div>
                                </td>
                            </tr>
                            <tr id="details-<?php echo $st['id']; ?>" class="hidden bg-slate-50 detail-row">
                                <td colspan="8" class="px-3 py-3">
                                    <div class="bg-white rounded-lg border border-slate-200 p-4">
                                        <h6 class="text-sm font-semibold text-slate-700 mb-3">Detail Pegawai yang Ditugaskan:</h6>
                                        <div class="overflow-x-auto">
                                            <table class="w-full text-sm text-left">
                                                <thead class="bg-slate-50 text-slate-500">
                                                    <tr>
                                                        <th class="px-4 py-2 border-b">No</th>
                                                        <th class="px-4 py-2 border-b">NIP</th>
                                                        <th class="px-4 py-2 border-b">Nama</th>
                                                        <th class="px-4 py-2 border-b">Jabatan</th>
                                                    </tr>
                                                </thead>
                                                <tbody class="divide-y divide-slate-100">
                                                    <?php
        $no_detail = 1;
        foreach ($pegawai_array as $p):
            if ($p != '') {
                $parts = explode(':::', $p);
                $nip = $parts[0] ?? '';
                $nama = $parts[1] ?? '';
                $jabatan = $parts[2] ?? '';
?>
                                                            <tr>
                                                                <td class="px-4 py-2"><?php echo $no_detail++; ?></td>
                                                                <td class="px-4 py-2 font-mono text-slate-600"><?php echo htmlspecialchars($nip ?? ''); ?></td>
                                                                <td class="px-4 py-2 font-medium text-slate-900"><?php echo htmlspecialchars($nama ?? ''); ?></td>
                                                                <td class="px-4 py-2 text-slate-600"><?php echo htmlspecialchars($jabatan ?? ''); ?></td>
                                                            </tr>
                                                            <?php
            }
        endforeach;
?>
                                                </tbody>
                                            </table>
                                        </div>
                                    </div>
                                </td>
                            </tr>
                        <?php
    endforeach; ?>
                    </tbody>
                </table>
            <?php
else: ?>
                <div class="p-8 text-center text-slate-500">
                    <i class='bx bx-file-blank text-4xl mb-2 text-slate-300'></i>
                    <p>
                        <?php if (!empty($search)): ?>
                            Tidak ditemukan surat tugas dengan kata kunci "<?php echo htmlspecialchars($search); ?>".
                        <?php
    else: ?>
                            Belum ada data surat tugas.
                        <?php
    endif; ?>
                    </p>
                </div>
            <?php
endif; ?>
        </div>

        <!-- Pagination -->
        <?php if ($total_pages > 0): ?>
            <div class="px-6 py-4 border-t border-slate-100 flex justify-center">
                <nav class="flex items-center gap-1">
                    <?php
                        $maxPageLinks = 7;
                        $half = (int)floor($maxPageLinks / 2);
                        $from = max(1, $page - $half);
                        $to = min($total_pages, $from + $maxPageLinks - 1);
                        $from = max(1, $to - $maxPageLinks + 1);
                    ?>
                    <?php if ($page > 1): ?>
                        <a href="?page=<?php echo($page - 1); ?><?php echo(!empty($search) ? '&search=' . urlencode($search) : ''); ?>" 
                           class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm">Previous</a>
                    <?php
    endif; ?>

                    <?php if ($from > 1): ?>
                        <a href="?page=1<?php echo(!empty($search) ? '&search=' . urlencode($search) : ''); ?>"
                           class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm">1</a>
                        <?php if ($from > 2): ?>
                            <span class="px-2 py-1 text-slate-400 text-sm">...</span>
                        <?php endif; ?>
                    <?php endif; ?>

                    <?php for ($i = $from; $i <= $to; $i++): ?>
                        <a href="?page=<?php echo $i; ?><?php echo(!empty($search) ? '&search=' . urlencode($search) : ''); ?>"
                           class="px-3 py-1 rounded border <?php echo($page == $i) ? 'bg-indigo-600 text-white border-indigo-600' : 'border-slate-200 text-slate-600 hover:bg-slate-50'; ?> text-sm">
                            <?php echo $i; ?>
                        </a>
                    <?php
    endfor; ?>

                    <?php if ($to < $total_pages): ?>
                        <?php if ($to < $total_pages - 1): ?>
                            <span class="px-2 py-1 text-slate-400 text-sm">...</span>
                        <?php endif; ?>
                        <a href="?page=<?php echo $total_pages; ?><?php echo(!empty($search) ? '&search=' . urlencode($search) : ''); ?>"
                           class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm"><?php echo $total_pages; ?></a>
                    <?php endif; ?>

                    <?php if ($page < $total_pages): ?>
                        <a href="?page=<?php echo($page + 1); ?><?php echo(!empty($search) ? '&search=' . urlencode($search) : ''); ?>"
                           class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm">Next</a>
                    <?php
    endif; ?>
                </nav>
            </div>
        <?php
endif; ?>
    </div>
</div>

<!-- Modal Tambah Surat -->
<div id="addSuratModal" class="fixed inset-0 z-50 hidden overflow-y-auto" aria-labelledby="modal-title" role="dialog" aria-modal="true">
    <div class="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
        <div class="fixed inset-0 bg-slate-900 bg-opacity-75 transition-opacity" aria-hidden="true" onclick="closeModal('addSuratModal')"></div>
        <span class="hidden sm:inline-block sm:align-middle sm:h-screen" aria-hidden="true">&#8203;</span>
        <div class="inline-block align-bottom bg-white rounded-xl text-left overflow-hidden shadow-xl transform transition-all sm:my-8 sm:align-middle sm:max-w-4xl sm:w-full">
            <div class="bg-white px-4 pt-5 pb-4 sm:p-6 sm:pb-4">
                <div class="flex justify-between items-center mb-4">
                    <h3 class="text-lg leading-6 font-medium text-slate-900">Buat Surat Tugas</h3>
                    <button onclick="closeModal('addSuratModal')" class="text-slate-400 hover:text-slate-500">
                        <i class='bx bx-x text-2xl'></i>
                    </button>
                </div>
                <form method="POST" id="formTambah">
                    <input type="hidden" name="action" value="create">
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Tanggal Surat</label>
                            <input type="date" name="tanggal_surat" id="add-tanggal-surat" required value="<?php echo date('Y-m-d'); ?>" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Nomor Surat</label>
                            <select name="nomor_surat" id="add-nomor-surat" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                                <option value="">Pilih tanggal surat terlebih dahulu</option>
                            </select>
                            <div id="nomor-manual-wrapper" class="hidden mt-2">
                                <label class="block text-xs text-slate-600 mb-1">Nomor manual (jika belum ada slot kosong)</label>
                                <input type="text" id="add-nomor-manual" class="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm" placeholder="Contoh: 762">
                            </div>
                            <p id="info-last-nomor-tanggal" class="text-xs text-indigo-600 mt-1"></p>
                            <p class="text-xs text-slate-500 mt-1">Nomor diambil dari slot kosong pada tanggal yang dipilih.</p>
                        </div>
                    </div>
                    <div class="mb-4">
                        <label class="block text-sm font-medium text-slate-700 mb-2">
                            <i class='bx bx-category text-indigo-500 mr-1'></i>Tipe Surat Tugas <span class="text-red-500">*</span>
                        </label>
                        <div class="flex gap-4">
                            <label class="flex items-center gap-2 px-4 py-2.5 border-2 border-blue-500 bg-blue-50 rounded-lg cursor-pointer transition-all tipe-surat-option" data-value="umum">
                                <input type="radio" name="tipe_surat" value="umum" checked class="h-4 w-4 text-blue-600 border-slate-300 focus:ring-blue-500">
                                <i class='bx bx-briefcase text-blue-600'></i>
                                <span class="text-sm font-medium text-slate-700">ASN Umum</span>
                            </label>
                            <label class="flex items-center gap-2 px-4 py-2.5 border-2 border-slate-200 rounded-lg cursor-pointer transition-all tipe-surat-option" data-value="penyuluh">
                                <input type="radio" name="tipe_surat" value="penyuluh" class="h-4 w-4 text-emerald-600 border-slate-300 focus:ring-emerald-500">
                                <i class='bx bx-leaf text-emerald-600'></i>
                                <span class="text-sm font-medium text-slate-700">Penyuluh</span>
                            </label>
                        </div>
                        <p class="text-xs text-slate-500 mt-1">Tipe surat menentukan template Word yang akan digunakan saat generate</p>
                    </div>
                    <div class="mb-4">
                        <label class="block text-sm font-medium text-slate-700 mb-1">Dasar Surat (Opsional)</label>
                        <textarea name="dasar_surat" rows="3" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none"></textarea>
                    </div>
                    <div class="mb-4">
                        <label class="block text-sm font-medium text-slate-700 mb-1">Untuk</label>
                        <textarea name="untuk" rows="3" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none"></textarea>
                    </div>
                    <div class="mb-4">
                        <label class="block text-sm font-medium text-slate-700 mb-1">Cari & Tambah Pegawai</label>
                        <div class="mb-2">
                            <select class="select2-add-single w-full">
                                <option value=""></option>
                                <?php foreach ($pegawai as $p): ?>
                                    <option value="<?php echo htmlspecialchars($p['nip']); ?>" data-nama="<?php echo htmlspecialchars($p['nama']); ?>" data-jabatan="<?php echo htmlspecialchars($p['jabatan']); ?>">
                                        <?php echo htmlspecialchars($p['nama']) . ' - ' . htmlspecialchars($p['jabatan']); ?>
                                    </option>
                                <?php endforeach; ?>
                            </select>
                        </div>
                        
                        <label class="block text-xs font-semibold text-slate-500 uppercase tracking-wider mb-2">Daftar Pegawai yang Ditugaskan (Seret / gunakan panah untuk mengatur urutan):</label>
                        <div id="selected-pegawai-container" class="space-y-2 border border-slate-200 rounded-lg p-3 min-h-[100px] bg-slate-50">
                            <p class="text-sm text-slate-400 text-center py-6" id="empty-state-pegawai">Belum ada pegawai yang dipilih</p>
                        </div>
                    </div>
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Tanggal Mulai</label>
                            <input type="date" name="tanggal_mulai" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Tanggal Selesai</label>
                            <input type="date" name="tanggal_selesai" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                    </div>
                    <div class="mb-4">
                        <label class="block text-sm font-medium text-slate-700 mb-1">
                            <i class='bx bx-pen text-indigo-500 mr-1'></i>Penanda Tangan
                        </label>
                        <select name="id_penandatangan" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                            <option value="">-- Pilih Penanda Tangan --</option>
                            <?php foreach ($penandatangan_list as $pt): ?>
                                <?php
    $is_kpl = isset($pt['is_kepala']) ? (int)$pt['is_kepala'] : 1;
    $tipe_label = $is_kpl ? '[Kepala]' : '[A.n.]';
    $label_detail = $is_kpl
        ? htmlspecialchars($pt['nama']) . ' - ' . htmlspecialchars($pt['jabatan'])
        : htmlspecialchars($pt['nama']) . ' - A.n. ' . htmlspecialchars($pt['jabatan_atasan'] ?? $pt['jabatan']);
    $default_mark = $pt['is_default'] ? ' ★ Default' : '';
?>
                                <option value="<?php echo $pt['id']; ?>" <?php echo $pt['is_default'] ? 'selected' : ''; ?>>
                                    <?php echo $tipe_label . ' ' . $label_detail . $default_mark; ?>
                                </option>
                            <?php
endforeach; ?>
                        </select>
                        <?php if (empty($penandatangan_list)): ?>
                            <p class="text-xs text-amber-600 mt-1">
                                <i class='bx bx-info-circle'></i> Belum ada data penanda tangan. 
                                <a href="penandatangan.php" class="text-indigo-600 hover:underline">Tambahkan di sini</a>
                            </p>
                        <?php
endif; ?>
                    </div>
                    <div class="mt-6 flex justify-end gap-3">
                        <button type="button" onclick="closeModal('addSuratModal')" class="px-4 py-2 border border-slate-300 rounded-lg text-slate-700 hover:bg-slate-50 font-medium">Batal</button>
                        <button type="submit" class="px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 font-medium">Simpan</button>
                    </div>
                </form>
            </div>
        </div>
    </div>
</div>

<!-- Modal Edit Surat -->
<div id="editSuratModal" class="fixed inset-0 z-50 hidden overflow-y-auto" aria-labelledby="modal-title" role="dialog" aria-modal="true">
    <div class="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
        <div class="fixed inset-0 bg-slate-900 bg-opacity-75 transition-opacity" aria-hidden="true" onclick="closeModal('editSuratModal')"></div>
        <span class="hidden sm:inline-block sm:align-middle sm:h-screen" aria-hidden="true">&#8203;</span>
        <div class="inline-block align-bottom bg-white rounded-xl text-left overflow-hidden shadow-xl transform transition-all sm:my-8 sm:align-middle sm:max-w-4xl sm:w-full">
            <div class="bg-white px-4 pt-5 pb-4 sm:p-6 sm:pb-4">
                <div class="flex justify-between items-center mb-4">
                    <h3 class="text-lg leading-6 font-medium text-slate-900">Edit Surat Tugas</h3>
                    <button onclick="closeModal('editSuratModal')" class="text-slate-400 hover:text-slate-500">
                        <i class='bx bx-x text-2xl'></i>
                    </button>
                </div>
                <form method="POST" id="formEdit">
                    <input type="hidden" name="action" value="update">
                    <input type="hidden" name="id" id="edit-id">
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Nomor Surat</label>
                            <input type="text" name="nomor_surat" id="edit-nomor" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Tanggal Surat</label>
                            <input type="date" name="tanggal_surat" id="edit-tanggal" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                    </div>
                    <div class="mb-4">
                        <label class="block text-sm font-medium text-slate-700 mb-2">
                            <i class='bx bx-category text-indigo-500 mr-1'></i>Tipe Surat Tugas <span class="text-red-500">*</span>
                        </label>
                        <div class="flex gap-4">
                            <label class="flex items-center gap-2 px-4 py-2.5 border-2 border-blue-500 bg-blue-50 rounded-lg cursor-pointer transition-all edit-tipe-surat-option" data-value="umum">
                                <input type="radio" name="tipe_surat" id="edit-tipe-umum" value="umum" checked class="h-4 w-4 text-blue-600 border-slate-300 focus:ring-blue-500">
                                <i class='bx bx-briefcase text-blue-600'></i>
                                <span class="text-sm font-medium text-slate-700">ASN Umum</span>
                            </label>
                            <label class="flex items-center gap-2 px-4 py-2.5 border-2 border-slate-200 rounded-lg cursor-pointer transition-all edit-tipe-surat-option" data-value="penyuluh">
                                <input type="radio" name="tipe_surat" id="edit-tipe-penyuluh" value="penyuluh" class="h-4 w-4 text-emerald-600 border-slate-300 focus:ring-emerald-500">
                                <i class='bx bx-leaf text-emerald-600'></i>
                                <span class="text-sm font-medium text-slate-700">Penyuluh</span>
                            </label>
                        </div>
                    </div>
                    <div class="mb-4">
                        <label class="block text-sm font-medium text-slate-700 mb-1">Dasar Surat (Opsional)</label>
                        <textarea name="dasar_surat" id="edit-dasar" rows="3" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none"></textarea>
                    </div>
                    <div class="mb-4">
                        <label class="block text-sm font-medium text-slate-700 mb-1">Untuk</label>
                        <textarea name="untuk" id="edit-untuk" rows="3" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none"></textarea>
                    </div>
                    <div class="mb-4">
                        <label class="block text-sm font-medium text-slate-700 mb-1">Cari & Tambah Pegawai</label>
                        <div class="mb-2">
                            <select class="select2-edit-single w-full">
                                <option value=""></option>
                                <?php foreach ($pegawai as $p): ?>
                                    <option value="<?php echo htmlspecialchars($p['nip']); ?>" data-nama="<?php echo htmlspecialchars($p['nama']); ?>" data-jabatan="<?php echo htmlspecialchars($p['jabatan']); ?>">
                                        <?php echo htmlspecialchars($p['nama']) . ' - ' . htmlspecialchars($p['jabatan']); ?>
                                    </option>
                                <?php endforeach; ?>
                            </select>
                        </div>
                        
                        <label class="block text-xs font-semibold text-slate-500 uppercase tracking-wider mb-2">Daftar Pegawai yang Ditugaskan (Seret / gunakan panah untuk mengatur urutan):</label>
                        <div id="edit-selected-pegawai-container" class="space-y-2 border border-slate-200 rounded-lg p-3 min-h-[100px] bg-slate-50">
                            <p class="text-sm text-slate-400 text-center py-6" id="edit-empty-state-pegawai">Belum ada pegawai yang dipilih</p>
                        </div>
                    </div>
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Tanggal Mulai</label>
                            <input type="date" name="tanggal_mulai" id="edit-mulai" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Tanggal Selesai</label>
                            <input type="date" name="tanggal_selesai" id="edit-selesai" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                    </div>
                    <div class="mb-4">
                        <label class="block text-sm font-medium text-slate-700 mb-1">
                            <i class='bx bx-pen text-indigo-500 mr-1'></i>Penanda Tangan
                        </label>
                        <select name="id_penandatangan" id="edit-penandatangan" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                            <option value="">-- Pilih Penanda Tangan --</option>
                            <?php foreach ($penandatangan_list as $pt): ?>
                                <?php
    $is_kpl = isset($pt['is_kepala']) ? (int)$pt['is_kepala'] : 1;
    $tipe_label = $is_kpl ? '[Kepala]' : '[A.n.]';
    $label_detail = $is_kpl
        ? htmlspecialchars($pt['nama']) . ' - ' . htmlspecialchars($pt['jabatan'])
        : htmlspecialchars($pt['nama']) . ' - A.n. ' . htmlspecialchars($pt['jabatan_atasan'] ?? $pt['jabatan']);
    $default_mark = $pt['is_default'] ? ' ★ Default' : '';
?>
                                <option value="<?php echo $pt['id']; ?>">
                                    <?php echo $tipe_label . ' ' . $label_detail . $default_mark; ?>
                                </option>
                            <?php
endforeach; ?>
                        </select>
                    </div>
                    <div class="mt-6 flex flex-wrap justify-end gap-3">
                        <input type="hidden" name="after_save" id="edit-after-save" value="">
                        <button type="button" onclick="closeModal('editSuratModal')" class="px-4 py-2 border border-slate-300 rounded-lg text-slate-700 hover:bg-slate-50 font-medium">Batal</button>
                        <button type="submit" onclick="document.getElementById('edit-after-save').value=''" class="px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 font-medium">Update</button>
                        <button type="submit" onclick="document.getElementById('edit-after-save').value='preview'" class="inline-flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 font-medium">
                            <i class="bx bx-show"></i> Simpan & Preview
                        </button>
                    </div>
                </form>
            </div>
        </div>
    </div>
</div>

<script>
    async function loadNomorKosongByTanggal() {
        const tanggalInput = document.getElementById('add-tanggal-surat');
        const nomorSelect = document.getElementById('add-nomor-surat');
        const manualWrapper = document.getElementById('nomor-manual-wrapper');
        const manualInput = document.getElementById('add-nomor-manual');
        if (!tanggalInput || !nomorSelect) return;

        const tanggal = tanggalInput.value;
        nomorSelect.innerHTML = '<option value="">Memuat nomor kosong...</option>';
        manualWrapper?.classList.add('hidden');
        if (manualInput) manualInput.value = '';
        const infoLast = document.getElementById('info-last-nomor-tanggal');
        if (infoLast) infoLast.textContent = '';

        if (!tanggal) {
            nomorSelect.innerHTML = '<option value="">Pilih tanggal surat terlebih dahulu</option>';
            return;
        }

        try {
            const resp = await fetch(`get_nomor_kosong.php?tanggal_surat=${encodeURIComponent(tanggal)}`);
            const data = await resp.json();
            nomorSelect.innerHTML = '';

            if (data.success && typeof data.last_nomor_tanggal !== 'undefined' && infoLast) {
                const lastNo = Number(data.last_nomor_tanggal) || 0;
                infoLast.textContent = lastNo > 0
                    ? `Nomor terakhir pada tanggal ini: ${lastNo}`
                    : 'Belum ada nomor terbit pada tanggal ini.';
            }

            if (data.success && Array.isArray(data.nomor) && data.nomor.length > 0) {
                nomorSelect.append(new Option('Pilih nomor kosong', ''));
                data.nomor.forEach(item => {
                    const statusLabel = item.status ? String(item.status).toLowerCase() : 'kosong';
                    let label = `${item.nomor_surat} - ${statusLabel}`;
                    if (item.tanggal_slot) {
                        const d = new Date(item.tanggal_slot);
                        const dd = String(d.getDate()).padStart(2, '0');
                        const mm = String(d.getMonth() + 1).padStart(2, '0');
                        const yyyy = d.getFullYear();
                        label += ` (slot ${dd}/${mm}/${yyyy})`;
                    }
                    nomorSelect.append(new Option(label, item.nomor_surat));
                });
                nomorSelect.required = true;
            } else {
                nomorSelect.append(new Option('Tidak ada nomor kosong untuk tanggal ini', ''));
                manualWrapper?.classList.remove('hidden');
                nomorSelect.required = false;
            }
        } catch (error) {
            nomorSelect.innerHTML = '<option value="">Gagal memuat nomor kosong</option>';
            manualWrapper?.classList.remove('hidden');
            nomorSelect.required = false;
        }
    }

    function syncNomorManualToSubmit(event) {
        const nomorSelect = document.getElementById('add-nomor-surat');
        const manualInput = document.getElementById('add-nomor-manual');
        if (!nomorSelect || !manualInput) return;

        if (!nomorSelect.value && manualInput.value.trim() !== '') {
            nomorSelect.innerHTML = '';
            nomorSelect.append(new Option(manualInput.value.trim(), manualInput.value.trim(), true, true));
            nomorSelect.required = true;
            return;
        }

        if (!nomorSelect.value && manualInput.value.trim() === '') {
            event.preventDefault();
            Swal.fire('Nomor belum dipilih', 'Pilih nomor kosong atau isi nomor manual.', 'warning');
        }
    }

    function openModal(modalId) {
        document.getElementById(modalId).classList.remove('hidden');
        if (modalId === 'addSuratModal') {
            loadNomorKosongByTanggal();
        }
    }

    function closeModal(modalId) {
        document.getElementById(modalId).classList.add('hidden');
    }

    function toggleDetails(id) {
        const element = document.getElementById(id);
        const btn = document.getElementById('btn-' + id);
        if (!element) return;
        
        const isHidden = element.classList.contains('hidden') || element.style.display === 'none';
        if (isHidden) {
            element.classList.remove('hidden');
            element.style.display = '';
        } else {
            element.classList.add('hidden');
            element.style.display = 'none';
        }
        if (btn) {
            btn.querySelector('i').classList.toggle('rotate-180');
        }
    }

    function generateSurat(id) {
        window.location.href = `generate-surat.php?id=${id}`;
    }

    // --- Logika Seleksi & Pengurutan Pegawai Custom ---
    function updateEmptyState(containerId) {
        const container = document.getElementById(containerId);
        if (!container) return;
        const items = container.querySelectorAll('.draggable-item');
        const emptyState = container.querySelector(containerId === 'selected-pegawai-container' ? '#empty-state-pegawai' : '#edit-empty-state-pegawai');
        
        if (items.length === 0) {
            if (emptyState) {
                emptyState.style.display = 'block';
            }
        } else {
            if (emptyState) {
                emptyState.style.display = 'none';
            }
        }
    }

    function addPegawaiRow(containerId, nip, nama, jabatan) {
        const container = document.getElementById(containerId);
        if (!container) return;

        // Cek jika sudah ada
        if (container.querySelector(`[data-nip="${nip}"]`)) {
            Swal.fire({
                title: 'Info',
                text: 'Pegawai ini sudah ada dalam daftar.',
                icon: 'info',
                timer: 1500,
                showConfirmButton: false
            });
            return;
        }

        const row = document.createElement('div');
        row.className = 'flex items-center justify-between p-3 bg-white border border-slate-200 rounded-lg shadow-sm group hover:border-indigo-300 transition-colors draggable-item';
        row.setAttribute('draggable', 'true');
        row.setAttribute('data-nip', nip);
        
        row.innerHTML = `
            <div class="flex items-center gap-3 flex-1 min-w-0">
                <div class="flex items-center gap-1.5 text-slate-400">
                    <i class='bx bx-menu cursor-move text-xl p-1 hover:text-slate-600' title='Seret untuk mengatur urutan'></i>
                    <div class="flex flex-col">
                        <button type="button" class="btn-move-up hover:text-indigo-600 hover:bg-slate-100 rounded p-0.5 transition-colors" title="Pindah Ke Atas">
                            <i class='bx bx-chevron-up text-base leading-none'></i>
                        </button>
                        <button type="button" class="btn-move-down hover:text-indigo-600 hover:bg-slate-100 rounded p-0.5 transition-colors" title="Pindah Ke Bawah">
                            <i class='bx bx-chevron-down text-base leading-none'></i>
                        </button>
                    </div>
                </div>
                <div class="min-w-0 flex-1">
                    <p class="font-semibold text-slate-800 text-sm truncate">${nama}</p>
                    <p class="text-xs text-slate-500 font-mono truncate">${nip} • ${jabatan}</p>
                    <input type="hidden" name="pegawai[]" value="${nip}">
                </div>
            </div>
            <button type="button" class="btn-remove-pegawai text-rose-500 hover:text-rose-700 hover:bg-rose-50 p-2 rounded-lg transition-colors flex items-center justify-center shrink-0 ml-2" title="Hapus dari daftar">
                <i class='bx bx-trash text-lg'></i>
            </button>
        `;
        
        container.appendChild(row);
        updateEmptyState(containerId);
    }

    function initDragAndDrop(containerId) {
        const container = document.getElementById(containerId);
        if (!container) return;
        let dragSource = null;

        container.addEventListener('dragstart', function(e) {
            const item = e.target.closest('.draggable-item');
            if (!item) return;
            dragSource = item;
            item.classList.add('opacity-50');
            e.dataTransfer.effectAllowed = 'move';
        });

        container.addEventListener('dragend', function(e) {
            const item = e.target.closest('.draggable-item');
            if (item) item.classList.remove('opacity-50');
            container.querySelectorAll('.draggable-item').forEach(el => {
                el.classList.remove('border-indigo-400', 'bg-indigo-50');
            });
        });

        container.addEventListener('dragover', function(e) {
            e.preventDefault();
            e.dataTransfer.dropEffect = 'move';
            return false;
        });

        container.addEventListener('dragenter', function(e) {
            const target = e.target.closest('.draggable-item');
            if (target && target !== dragSource) {
                target.classList.add('border-indigo-400', 'bg-indigo-50');
            }
        });

        container.addEventListener('dragleave', function(e) {
            const target = e.target.closest('.draggable-item');
            if (target && target !== dragSource) {
                target.classList.remove('border-indigo-400', 'bg-indigo-50');
            }
        });

        container.addEventListener('drop', function(e) {
            e.stopPropagation();
            const target = e.target.closest('.draggable-item');
            if (target && target !== dragSource) {
                target.classList.remove('border-indigo-400', 'bg-indigo-50');
                
                const rect = target.getBoundingClientRect();
                const next = (e.clientY - rect.top) / (rect.bottom - rect.top) > 0.5;
                
                container.insertBefore(dragSource, next ? target.nextSibling : target);
                updateEmptyState(containerId);
            }
            return false;
        });
    }

    $(document).ready(function () {
        // Inisialisasi Select2 Single
        $('.select2-add-single').select2({
            placeholder: "Cari pegawai untuk ditambahkan...",
            allowClear: true,
            width: '100%'
        }).on('select2:select', function(e) {
            const nip = e.params.data.id;
            if (!nip) return;
            const $option = $(e.params.data.element);
            const nama = $option.data('nama');
            const jabatan = $option.data('jabatan');
            
            addPegawaiRow('selected-pegawai-container', nip, nama, jabatan);
            $(this).val('').trigger('change'); // reset selection
        });

        $('.select2-edit-single').select2({
            placeholder: "Cari pegawai untuk ditambahkan...",
            allowClear: true,
            width: '100%'
        }).on('select2:select', function(e) {
            const nip = e.params.data.id;
            if (!nip) return;
            const $option = $(e.params.data.element);
            const nama = $option.data('nama');
            const jabatan = $option.data('jabatan');
            
            addPegawaiRow('edit-selected-pegawai-container', nip, nama, jabatan);
            $(this).val('').trigger('change'); // reset selection
        });

        // Init drag and drop untuk container
        initDragAndDrop('selected-pegawai-container');
        initDragAndDrop('edit-selected-pegawai-container');

        // Handler untuk tombol move up/down
        $(document).on('click', '.btn-move-up', function() {
            const item = $(this).closest('.draggable-item');
            const prev = item.prev('.draggable-item');
            if (prev.length > 0) {
                item.insertBefore(prev);
            }
        });

        $(document).on('click', '.btn-move-down', function() {
            const item = $(this).closest('.draggable-item');
            const next = item.next('.draggable-item');
            if (next.length > 0) {
                item.insertAfter(next);
            }
        });

        // Handler untuk tombol hapus
        $(document).on('click', '.btn-remove-pegawai', function() {
            const container = $(this).closest('.draggable-item').parent();
            const containerId = container.attr('id');
            $(this).closest('.draggable-item').remove();
            updateEmptyState(containerId);
        });
    });

    // Nomor surat by tanggal
    document.getElementById('add-tanggal-surat')?.addEventListener('change', loadNomorKosongByTanggal);
    
    // Submit handler formTambah dengan validasi
    document.getElementById('formTambah')?.addEventListener('submit', function(e) {
        const container = document.getElementById('selected-pegawai-container');
        const items = container ? container.querySelectorAll('.draggable-item') : [];
        if (items.length === 0) {
            e.preventDefault();
            Swal.fire('Validasi Gagal', 'Minimal harus memilih satu pegawai.', 'warning');
            return false;
        }
        syncNomorManualToSubmit(e);
    });

    // Submit handler formEdit dengan validasi
    document.getElementById('formEdit')?.addEventListener('submit', function(e) {
        const container = document.getElementById('edit-selected-pegawai-container');
        const items = container ? container.querySelectorAll('.draggable-item') : [];
        if (items.length === 0) {
            e.preventDefault();
            Swal.fire('Validasi Gagal', 'Minimal harus memilih satu pegawai.', 'warning');
            return false;
        }
    });

    // Script untuk edit
    document.querySelectorAll('.edit-btn').forEach(button => {
        button.addEventListener('click', async function () {
            const id = this.dataset.id;

            try {
                const response = await fetch(`get_assigned_pegawai.php?id=${id}`);
                const assignedPegawai = await response.json();

                document.getElementById('edit-id').value = id;
                document.getElementById('edit-nomor').value = this.dataset.nomor;
                document.getElementById('edit-tanggal').value = this.dataset.tanggal;
                document.getElementById('edit-dasar').value = this.dataset.dasar;
                document.getElementById('edit-untuk').value = this.dataset.untuk;
                document.getElementById('edit-mulai').value = this.dataset.mulai;
                document.getElementById('edit-selesai').value = this.dataset.selesai;

                // Set penanda tangan
                const penandatanganSelect = document.getElementById('edit-penandatangan');
                if (penandatanganSelect) {
                    penandatanganSelect.value = this.dataset.penandatangan || '';
                }

                // Set tipe surat
                const tipeSurat = this.dataset.tipe || 'umum';
                if (tipeSurat === 'penyuluh') {
                    document.getElementById('edit-tipe-penyuluh').checked = true;
                } else {
                    document.getElementById('edit-tipe-umum').checked = true;
                }
                // Update styling for edit tipe surat radio buttons
                document.querySelectorAll('.edit-tipe-surat-option').forEach(opt => {
                    if (opt.dataset.value === tipeSurat) {
                        opt.classList.remove('border-slate-200');
                        if (tipeSurat === 'penyuluh') {
                            opt.classList.add('border-emerald-500', 'bg-emerald-50');
                        } else {
                            opt.classList.add('border-blue-500', 'bg-blue-50');
                        }
                    } else {
                        opt.classList.remove('border-blue-500', 'bg-blue-50', 'border-emerald-500', 'bg-emerald-50');
                        opt.classList.add('border-slate-200');
                    }
                });

                // Bersihkan kontainer edit pegawai
                const editContainer = document.getElementById('edit-selected-pegawai-container');
                if (editContainer) {
                    editContainer.querySelectorAll('.draggable-item').forEach(el => el.remove());
                    updateEmptyState('edit-selected-pegawai-container');
                }

                // Masukkan pegawai yang ditugaskan ke kontainer
                if (Array.isArray(assignedPegawai)) {
                    assignedPegawai.forEach(nip => {
                        const $opt = $(`.select2-edit-single option[value="${nip}"]`);
                        if ($opt.length) {
                            const nama = $opt.data('nama');
                            const jabatan = $opt.data('jabatan');
                            addPegawaiRow('edit-selected-pegawai-container', nip, nama, jabatan);
                        }
                    });
                }
                openModal('editSuratModal');

            } catch (error) {
                console.error('Error:', error);
                Swal.fire('Error!', 'Gagal mengambil data pegawai', 'error');
            }
        });
    });

    // Script untuk delete
    document.querySelectorAll('.delete-btn').forEach(button => {
        button.addEventListener('click', function () {
            const id = this.dataset.id;
            Swal.fire({
                title: 'Apakah Anda yakin?',
                text: "Data surat tugas akan dihapus permanen!",
                icon: 'warning',
                showCancelButton: true,
                confirmButtonColor: '#ef4444',
                cancelButtonColor: '#3b82f6',
                confirmButtonText: 'Ya, hapus!',
                cancelButtonText: 'Batal'
            }).then((result) => {
                if (result.isConfirmed) {
                    const formData = new FormData();
                    formData.append('id', id);

                    fetch('delete_surat.php', {
                        method: 'POST',
                        body: formData
                    })
                        .then(response => response.json())
                        .then(data => {
                            if (data.success) {
                                Swal.fire({
                                    icon: 'success',
                                    title: 'Berhasil!',
                                    text: 'Data surat tugas berhasil dihapus',
                                    showConfirmButton: false,
                                    timer: 1500
                                }).then(() => {
                                    window.location.reload();
                                });
                            } else {
                                throw new Error(data.message || 'Terjadi kesalahan saat menghapus data');
                            }
                        })
                        .catch(error => {
                            console.error('Error:', error);
                            Swal.fire({
                                icon: 'error',
                                title: 'Error!',
                                text: error.message
                            });
                        });
                }
            });
        });
    });

    // Tipe Surat radio button styling - Add form
    document.querySelectorAll('.tipe-surat-option input[type="radio"]').forEach(radio => {
        radio.addEventListener('change', function() {
            document.querySelectorAll('.tipe-surat-option').forEach(opt => {
                opt.classList.remove('border-blue-500', 'bg-blue-50', 'border-emerald-500', 'bg-emerald-50');
                opt.classList.add('border-slate-200');
            });
            const parent = this.closest('.tipe-surat-option');
            parent.classList.remove('border-slate-200');
            if (this.value === 'penyuluh') {
                parent.classList.add('border-emerald-500', 'bg-emerald-50');
            } else {
                parent.classList.add('border-blue-500', 'bg-blue-50');
            }
        });
    });

    // Tipe Surat radio button styling - Edit form
    document.querySelectorAll('.edit-tipe-surat-option input[type="radio"]').forEach(radio => {
        radio.addEventListener('change', function() {
            document.querySelectorAll('.edit-tipe-surat-option').forEach(opt => {
                opt.classList.remove('border-blue-500', 'bg-blue-50', 'border-emerald-500', 'bg-emerald-50');
                opt.classList.add('border-slate-200');
            });
            const parent = this.closest('.edit-tipe-surat-option');
            parent.classList.remove('border-slate-200');
            if (this.value === 'penyuluh') {
                parent.classList.add('border-emerald-500', 'bg-emerald-50');
            } else {
                parent.classList.add('border-blue-500', 'bg-blue-50');
            }
        });
    });

    // Auto-open modal edit jika URL mengandung ?edit=ID (diarahkan dari halaman generate-surat)
    (function() {
        const urlParams = new URLSearchParams(window.location.search);
        const editId = urlParams.get('edit');
        if (editId) {
            const editBtn = document.querySelector(`.edit-btn[data-id="${editId}"]`);
            if (editBtn) {
                editBtn.click();
            }
            // Hapus parameter 'edit' dari URL agar bersih tanpa reload
            const newUrl = window.location.pathname + (urlParams.toString().replace(/edit=[^&]*&?/, '').replace(/&$/, '') ? '?' + urlParams.toString().replace(/edit=[^&]*&?/, '').replace(/&$/, '') : '');
            history.replaceState(null, '', newUrl);
        }
    })();

    // ===== CLIENT-SIDE FILTER & SORT =====
    (function() {
        const filterTipe = document.getElementById('filterTipe');
        const filterDateFrom = document.getElementById('filterDateFrom');
        const filterDateTo = document.getElementById('filterDateTo');
        const btnResetFilter = document.getElementById('btnResetFilter');

        function getRows() {
            return document.querySelectorAll('#tabelSurat tbody .surat-row');
        }

        function applyFilters() {
            const rows = getRows();
            const tipe = filterTipe ? filterTipe.value : '';
            const dateFrom = filterDateFrom ? filterDateFrom.value : '';
            const dateTo = filterDateTo ? filterDateTo.value : '';
            let hasActiveFilter = (tipe !== '' || dateFrom !== '' || dateTo !== '');

            rows.forEach(row => {
                let show = true;
                const rowTipe = row.dataset.tipe || '';
                const rowTanggal = row.dataset.tanggal || '';

                // Filter tipe
                if (tipe && rowTipe !== tipe) show = false;

                // Filter date range
                if (dateFrom && rowTanggal < dateFrom) show = false;
                if (dateTo && rowTanggal > dateTo) show = false;

                row.style.display = show ? '' : 'none';
                // Also hide/show the associated detail row
                const detailId = row.querySelector('button[id^="btn-details-"]');
                if (detailId) {
                    const id = detailId.id.replace('btn-', '');
                    const detailRow = document.getElementById(id);
                    if (detailRow) {
                        if (!show) {
                            detailRow.style.display = 'none';
                        } else {
                            // Restore: only show if it was already expanded (not hidden class)
                            detailRow.style.display = detailRow.classList.contains('hidden') ? 'none' : '';
                        }
                    }
                }
            });

            // Show/hide reset button
            if (btnResetFilter) {
                btnResetFilter.classList.toggle('hidden', !hasActiveFilter);
            }
        }

        if (filterTipe) filterTipe.addEventListener('change', applyFilters);
        if (filterDateFrom) filterDateFrom.addEventListener('change', applyFilters);
        if (filterDateTo) filterDateTo.addEventListener('change', applyFilters);

        if (btnResetFilter) {
            btnResetFilter.addEventListener('click', function() {
                if (filterTipe) filterTipe.value = '';
                if (filterDateFrom) filterDateFrom.value = '';
                if (filterDateTo) filterDateTo.value = '';
                applyFilters();
            });
        }

        // ===== SORT BY COLUMN =====
        document.querySelectorAll('.sortable-header').forEach(header => {
            header.addEventListener('click', function() {
                const sortField = this.dataset.sort;
                let order = this.dataset.order;
                
                // Toggle order
                if (order === 'asc') order = 'desc';
                else order = 'asc';
                this.dataset.order = order;

                // Reset other headers
                document.querySelectorAll('.sortable-header').forEach(h => {
                    if (h !== this) {
                        h.dataset.order = '';
                        h.querySelector('.sort-icon').className = 'bx bx-sort-alt-2 text-slate-400 sort-icon';
                    }
                });

                // Update icon
                const icon = this.querySelector('.sort-icon');
                icon.className = order === 'asc' 
                    ? 'bx bx-sort-up text-indigo-600 sort-icon' 
                    : 'bx bx-sort-down text-indigo-600 sort-icon';

                // Sort rows
                const tbody = document.querySelector('#tabelSurat tbody');
                if (!tbody) return;

                // Collect row pairs (data row + its detail row by ID)
                const rowPairs = [];
                const suratRows = tbody.querySelectorAll('.surat-row');
                suratRows.forEach(row => {
                    const btnEl = row.querySelector('button[id^="btn-details-"]');
                    let detailRow = null;
                    if (btnEl) {
                        const detailId = btnEl.id.replace('btn-', '');
                        detailRow = document.getElementById(detailId);
                    }
                    rowPairs.push({ main: row, detail: detailRow });
                });

                rowPairs.sort((a, b) => {
                    let valA, valB;
                    if (sortField === 'tanggal') {
                        valA = a.main.dataset.tanggal || '';
                        valB = b.main.dataset.tanggal || '';
                    } else if (sortField === 'pelaksanaan') {
                        valA = a.main.dataset.mulai || '';
                        valB = b.main.dataset.mulai || '';
                    }
                    if (order === 'asc') return valA.localeCompare(valB);
                    return valB.localeCompare(valA);
                });

                // Re-append sorted rows (main + detail always together)
                rowPairs.forEach(pair => {
                    tbody.appendChild(pair.main);
                    if (pair.detail) tbody.appendChild(pair.detail);
                });

                // Re-apply filters after sort
                applyFilters();
            });
        });
    })();
</script>

<?php include '../includes/footer.php'; ?>