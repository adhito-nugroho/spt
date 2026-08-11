<?php
$pdo = require_once '../config/database.php';

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

function ensurePerformanceIndexes(PDO $pdo): void
{
    $indexStatements = [
        "ALTER TABLE buku_nomor_surat_tugas ADD INDEX idx_buku_tahun_status_nomor (tahun, status, nomor_surat)",
        "ALTER TABLE buku_nomor_surat_tugas ADD INDEX idx_buku_tanggal (tanggal_surat)",
    ];
    foreach ($indexStatements as $sql) {
        try {
            $pdo->exec($sql);
        } catch (PDOException $e) {
            // ignore existing index
        }
    }
}

function normalizeBukuNomorStatus(PDO $pdo): void
{
    $pdo->exec("UPDATE buku_nomor_surat_tugas
        SET status = 'kosong'
        WHERE status IN ('cadangan','batal') AND id_surat_tugas IS NULL");
}

ensureBukuNomorSuratTugasTable($pdo);
ensurePerformanceIndexes($pdo);
normalizeBukuNomorStatus($pdo);
syncSuratTugasTerpakaiKeBukuNomor($pdo);

function getLastNomorByTahun(PDO $pdo, int $tahun): int
{
    $stmt = $pdo->prepare("SELECT COALESCE(MAX(CAST(nomor_surat AS UNSIGNED)), 0) AS last_no
        FROM buku_nomor_surat_tugas
        WHERE tahun = :tahun
          AND nomor_surat REGEXP '^[0-9]+$'");
    $stmt->execute([':tahun' => $tahun]);
    return (int)$stmt->fetchColumn();
}

function getLastTanggalByTahun(PDO $pdo, int $tahun): ?string
{
    $stmt = $pdo->prepare("SELECT MAX(tanggal_surat) AS last_tanggal
        FROM buku_nomor_surat_tugas
        WHERE tahun = :tahun
          AND tanggal_surat IS NOT NULL");
    $stmt->execute([':tahun' => $tahun]);
    $value = $stmt->fetchColumn();
    return $value ?: null;
}

/**
 * Slot manual (Tambah Slot Kosong): tanggal boleh mundur asalkan tanggal sebelum tanggal terakhir buku
 * dan nomor numerik di bawah nomor terakhir buku (tahun sama). Nomor tidak boleh sudah terpakai surat — dicek di pemanggil.
 */
function validateSlotManualTanggalNomor(PDO $pdo, int $tahun, string $tanggalYmd, string $nomor): void
{
    $lastTanggal = getLastTanggalByTahun($pdo, $tahun);
    if ($lastTanggal === null || $tanggalYmd >= $lastTanggal) {
        return;
    }
    $lastNomor = getLastNomorByTahun($pdo, $tahun);
    if ($lastNomor <= 0) {
        throw new Exception(
            'Tanggal sebelum tanggal terakhir buku (' . date('d/m/Y', strtotime($lastTanggal)) . ') memerlukan perbandingan nomor; belum ada nomor numerik terakhir di tahun ini.'
        );
    }
    if (!preg_match('/^[0-9]+$/', $nomor)) {
        throw new Exception(
            'Untuk tanggal sebelum tanggal terakhir buku (' . date('d/m/Y', strtotime($lastTanggal)) . '), nomor surat harus angka dan lebih kecil dari nomor terakhir (' . $lastNomor . ').'
        );
    }
    if ((int)$nomor >= $lastNomor) {
        throw new Exception(
            'Tanggal mundur hanya diperbolehkan jika nomor lebih kecil dari nomor terakhir buku (' . $lastNomor . ') dan tanggal lebih kecil dari tanggal terakhir (' . date('d/m/Y', strtotime($lastTanggal)) . ').'
        );
    }
}

function createCadanganSpaceByTanggal(PDO $pdo, int $tahun, string $tanggal, int $jumlah): array
{
    $lastTanggal = getLastTanggalByTahun($pdo, $tahun);
    if ($lastTanggal !== null && $tanggal < $lastTanggal) {
        throw new Exception(
            'Buat Space otomatis menambah nomor setelah nomor terakhir, jadi tidak bisa memakai tanggal sebelum tanggal terakhir buku (' . date('d/m/Y', strtotime($lastTanggal)) . '). Untuk tanggal mundur gunakan Tambah Slot Kosong: nomor numerik harus lebih kecil dari nomor terakhir dan tanggal lebih kecil dari tanggal terakhir, serta nomor belum terpakai surat.'
        );
    }
    $jumlah = max(1, min(100, $jumlah));
    $last = getLastNomorByTahun($pdo, $tahun);
    $candidate = $last + 1;
    $created = 0;
    $firstCreated = null;
    $lastCreated = null;
    $guard = 0;

    while ($created < $jumlah && $guard < ($jumlah * 200)) {
        $guard++;
        $nomor = (string)$candidate;

        $stmtExist = $pdo->prepare("SELECT id, id_surat_tugas
            FROM buku_nomor_surat_tugas
            WHERE tahun = :tahun AND nomor_surat = :nomor
            LIMIT 1");
        $stmtExist->execute([
            ':tahun' => $tahun,
            ':nomor' => $nomor,
        ]);
        $exist = $stmtExist->fetch(PDO::FETCH_ASSOC);

        if (!$exist) {
            $stmtIns = $pdo->prepare("INSERT INTO buku_nomor_surat_tugas
                (tahun, nomor_surat, tanggal_surat, status, id_surat_tugas, keterangan)
                VALUES
                (:tahun, :nomor, :tanggal, 'kosong', NULL, 'Space otomatis per tanggal')");
            $stmtIns->execute([
                ':tahun' => $tahun,
                ':nomor' => $nomor,
                ':tanggal' => $tanggal,
            ]);
            if ($firstCreated === null) {
                $firstCreated = $nomor;
            }
            $lastCreated = $nomor;
            $created++;
        } elseif (empty($exist['id_surat_tugas'])) {
            $stmtUpd = $pdo->prepare("UPDATE buku_nomor_surat_tugas
                SET tanggal_surat = :tanggal, status = 'kosong', keterangan = 'Space otomatis per tanggal'
                WHERE id = :id");
            $stmtUpd->execute([
                ':tanggal' => $tanggal,
                ':id' => $exist['id'],
            ]);
            if ($firstCreated === null) {
                $firstCreated = $nomor;
            }
            $lastCreated = $nomor;
            $created++;
        }
        $candidate++;
    }

    return [
        'created' => $created,
        'first' => $firstCreated,
        'last' => $lastCreated,
        'base_last' => $last,
    ];
}

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['action'])) {
    try {
        if ($_POST['action'] === 'reserve_slot') {
            $tahun = (int)($_POST['tahun'] ?? date('Y'));
            $nomor = trim((string)($_POST['nomor_surat'] ?? ''));
            $tanggal = !empty($_POST['tanggal_surat']) ? $_POST['tanggal_surat'] : '';
            $keterangan = trim((string)($_POST['keterangan'] ?? 'Slot kosong manual'));

            if ($nomor === '') {
                throw new Exception('Nomor surat wajib diisi.');
            }
            if ($tahun < 1900) {
                throw new Exception('Tahun tidak valid.');
            }
            if ($tanggal === '' || strtotime($tanggal) === false) {
                throw new Exception('Tanggal surat wajib diisi dan harus valid.');
            }
            $tanggal = date('Y-m-d', strtotime($tanggal));
            validateSlotManualTanggalNomor($pdo, $tahun, $tanggal, $nomor);

            $stmt = $pdo->prepare("SELECT id_surat_tugas, status FROM buku_nomor_surat_tugas WHERE tahun = ? AND nomor_surat = ? LIMIT 1");
            $stmt->execute([$tahun, $nomor]);
            $existing = $stmt->fetch();
            if ($existing && !empty($existing['id_surat_tugas'])) {
                throw new Exception('Nomor sudah terpakai surat terbit, tidak bisa dijadikan slot kosong.');
            }

            $stmt = $pdo->prepare("INSERT INTO buku_nomor_surat_tugas (tahun, nomor_surat, tanggal_surat, status, id_surat_tugas, keterangan)
                VALUES (:tahun, :nomor_surat, :tanggal_surat, 'kosong', NULL, :keterangan)
                ON DUPLICATE KEY UPDATE
                    tanggal_surat = VALUES(tanggal_surat),
                    status = 'kosong',
                    id_surat_tugas = NULL,
                    keterangan = VALUES(keterangan)");
            $stmt->execute([
                ':tahun' => $tahun,
                ':nomor_surat' => $nomor,
                ':tanggal_surat' => $tanggal,
                ':keterangan' => $keterangan,
            ]);
        } elseif ($_POST['action'] === 'set_status') {
            $id = (int)($_POST['id'] ?? 0);
            $status = (string)($_POST['status'] ?? '');
            $allowed = ['kosong'];
            if ($id <= 0 || !in_array($status, $allowed, true)) {
                throw new Exception('Permintaan tidak valid.');
            }

            $stmt = $pdo->prepare("SELECT id_surat_tugas FROM buku_nomor_surat_tugas WHERE id = ? LIMIT 1");
            $stmt->execute([$id]);
            $row = $stmt->fetch();
            if (!$row) {
                throw new Exception('Data buku nomor tidak ditemukan.');
            }
            if (!empty($row['id_surat_tugas'])) {
                throw new Exception('Nomor sudah terikat ke surat terbit, status tidak bisa diubah manual.');
            }

            $stmt = $pdo->prepare("UPDATE buku_nomor_surat_tugas SET status = ?, keterangan = ? WHERE id = ?");
            $stmt->execute([$status, 'Diubah manual dari buku nomor', $id]);
        } elseif ($_POST['action'] === 'delete_slot') {
            $id = (int)($_POST['id'] ?? 0);
            if ($id <= 0) {
                throw new Exception('ID slot tidak valid.');
            }

            $stmt = $pdo->prepare("SELECT id_surat_tugas FROM buku_nomor_surat_tugas WHERE id = ? LIMIT 1");
            $stmt->execute([$id]);
            $row = $stmt->fetch();
            if (!$row) {
                throw new Exception('Data buku nomor tidak ditemukan.');
            }
            if (!empty($row['id_surat_tugas'])) {
                throw new Exception('Nomor sudah terikat ke surat terbit, tidak bisa dihapus.');
            }

            $stmt = $pdo->prepare("DELETE FROM buku_nomor_surat_tugas WHERE id = ?");
            $stmt->execute([$id]);
        } elseif ($_POST['action'] === 'create_space_tanggal') {
            $tanggal = trim((string)($_POST['tanggal_surat'] ?? ''));
            $jumlah = (int)($_POST['jumlah_space'] ?? 10);
            if ($tanggal === '' || strtotime($tanggal) === false) {
                throw new Exception('Tanggal surat tidak valid.');
            }
            $tahun = (int)date('Y', strtotime($tanggal));
            $pdo->beginTransaction();
            try {
                $result = createCadanganSpaceByTanggal($pdo, $tahun, date('Y-m-d', strtotime($tanggal)), $jumlah);
                $pdo->commit();
            } catch (Throwable $e) {
                if ($pdo->inTransaction()) $pdo->rollBack();
                throw $e;
            }
            $msg = "Space berhasil dibuat: {$result['created']} nomor";
            if ($result['first'] !== null && $result['last'] !== null) {
                $msg .= " ({$result['first']} s/d {$result['last']})";
            }
            header('Location: buku-nomor.php?ok=1&msg=' . urlencode($msg));
            exit;
        }

        header('Location: buku-nomor.php?ok=1');
        exit;
    } catch (Exception $e) {
        header('Location: buku-nomor.php?error=' . urlencode($e->getMessage()));
        exit;
    }
}

$tahun = isset($_GET['tahun']) ? (int)$_GET['tahun'] : (int)date('Y');
$status = isset($_GET['status']) ? trim((string)$_GET['status']) : '';
$search = isset($_GET['search']) ? trim((string)$_GET['search']) : '';
$tanggal_dari_raw = isset($_GET['tanggal_dari']) ? trim((string)$_GET['tanggal_dari']) : '';
$tanggal_sampai_raw = isset($_GET['tanggal_sampai']) ? trim((string)$_GET['tanggal_sampai']) : '';
$page = isset($_GET['page']) ? max(1, (int)$_GET['page']) : 1;
$limit = 30;

$tanggalDari = null;
$tanggalSampai = null;
if ($tanggal_dari_raw !== '') {
    $ts = strtotime($tanggal_dari_raw);
    if ($ts !== false) {
        $tanggalDari = date('Y-m-d', $ts);
    }
}
if ($tanggal_sampai_raw !== '') {
    $ts = strtotime($tanggal_sampai_raw);
    if ($ts !== false) {
        $tanggalSampai = date('Y-m-d', $ts);
    }
}
if ($tanggalDari !== null && $tanggalSampai !== null && $tanggalDari > $tanggalSampai) {
    $tmp = $tanggalDari;
    $tanggalDari = $tanggalSampai;
    $tanggalSampai = $tmp;
}
$tanggalDariInput = $tanggalDari ?? '';
$tanggalSampaiInput = $tanggalSampai ?? '';

$where = ["b.tahun = :tahun"];
$params = [':tahun' => $tahun];
if ($status !== '' && in_array($status, ['kosong', 'terisi'], true)) {
    $where[] = "b.status = :status";
    $params[':status'] = $status;
}
if ($search !== '') {
    $where[] = "(b.nomor_surat LIKE :search OR COALESCE(b.keterangan,'') LIKE :search)";
    $params[':search'] = '%' . $search . '%';
}
if ($tanggalDari !== null) {
    $where[] = "b.tanggal_surat >= :tgl_dari";
    $params[':tgl_dari'] = $tanggalDari;
}
if ($tanggalSampai !== null) {
    $where[] = "b.tanggal_surat <= :tgl_sampai";
    $params[':tgl_sampai'] = $tanggalSampai;
}
$whereSql = implode(' AND ', $where);

$stmtCount = $pdo->prepare("SELECT COUNT(*) AS total FROM buku_nomor_surat_tugas b WHERE $whereSql");
$stmtCount->execute($params);
$totalRecords = (int)($stmtCount->fetch()['total'] ?? 0);
$totalPages = $totalRecords > 0 ? (int)ceil($totalRecords / $limit) : 1;
$page = min($page, $totalPages);
$start = ($page - 1) * $limit;

// LIMIT dengan placeholder sering bermasalah pada native prepared MySQL; offset/limit sudah integer aman.
$startSql = (int)$start;
$limitSql = (int)$limit;
$stmtData = $pdo->prepare("SELECT b.*, st.untuk
    FROM buku_nomor_surat_tugas b
    LEFT JOIN surat_tugas st ON st.id = b.id_surat_tugas
    WHERE $whereSql
    ORDER BY CAST(b.nomor_surat AS UNSIGNED) ASC, b.nomor_surat ASC
    LIMIT {$startSql}, {$limitSql}");
foreach ($params as $key => $value) {
    $stmtData->bindValue($key, $value);
}
$stmtData->execute();
$rows = $stmtData->fetchAll();
$lastNomorTahun = getLastNomorByTahun($pdo, $tahun);
$lastTanggalTahun = getLastTanggalByTahun($pdo, $tahun);

include '../includes/header.php';
?>

<div class="space-y-6">
    <div class="flex flex-col md:flex-row justify-between items-center gap-4">
        <div>
            <h1 class="page-header-title">Buku Nomor Surat Tugas</h1>
            <p class="page-header-sub">Penomoran dan alokasi slot surat tugas</p>
        </div>
        <div class="flex gap-2">
            <button onclick="openModal('spaceModal')" class="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg flex items-center gap-2 transition-colors shadow-sm text-sm font-medium">
                <i class='bx bx-grid-alt'></i> Buat Space 10 Nomor
            </button>
            <button onclick="openModal('reserveModal')" class="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg flex items-center gap-2 transition-colors shadow-sm text-sm font-medium">
                <i class='bx bx-plus'></i> Tambah Slot Kosong
            </button>
        </div>
    </div>

    <?php if (isset($_GET['ok'])): ?>
        <div class="bg-emerald-50 border border-emerald-200 text-emerald-700 px-4 py-3 rounded-lg text-sm">
            <?php echo isset($_GET['msg']) ? htmlspecialchars((string)$_GET['msg']) : 'Perubahan buku nomor berhasil disimpan.'; ?>
        </div>
    <?php endif; ?>
    <?php if (isset($_GET['error'])): ?>
        <div class="bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded-lg text-sm">
            <?php echo htmlspecialchars((string)$_GET['error']); ?>
        </div>
    <?php endif; ?>

    <div class="card p-4 bg-indigo-50/50 border-indigo-100 text-indigo-900 flex flex-col sm:flex-row sm:flex-wrap sm:items-baseline gap-1 sm:gap-x-6 gap-y-1 text-sm">
        <span>
            Nomor terakhir tahun <?php echo $tahun; ?>: <span class="font-semibold text-indigo-700"><?php echo $lastNomorTahun; ?></span>
            <span class="text-indigo-500 text-xs ml-2">(nomor baru akan lanjut dari <?php echo $lastNomorTahun + 1; ?>)</span>
        </span>
        <span class="text-sm">
            Tanggal terakhir pada buku (tahun <?php echo $tahun; ?>):
            <span class="font-semibold text-indigo-700"><?php echo $lastTanggalTahun !== null ? htmlspecialchars(date('d/m/Y', strtotime($lastTanggalTahun))) : '—'; ?></span>
            <span class="text-indigo-500 text-xs ml-1">(tanggal surat tertinggi yang tercatat)</span>
        </span>
    </div>

    <div class="card p-4">
        <form method="GET" class="space-y-3">
            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
                <div>
                    <label class="block text-xs font-medium text-gray-500 mb-1">Tahun</label>
                    <input type="number" name="tahun" value="<?php echo $tahun; ?>" class="input-premium py-1.5">
                </div>
                <div>
                    <label class="block text-xs font-medium text-gray-500 mb-1">Status</label>
                    <select name="status" class="input-premium py-1.5">
                        <option value="">Semua</option>
                        <option value="kosong" <?php echo $status === 'kosong' ? 'selected' : ''; ?>>Kosong</option>
                        <option value="terisi" <?php echo $status === 'terisi' ? 'selected' : ''; ?>>Terisi</option>
                    </select>
                </div>
                <div>
                    <label class="block text-xs font-medium text-gray-500 mb-1">Tanggal surat dari</label>
                    <input type="date" name="tanggal_dari" value="<?php echo htmlspecialchars($tanggalDariInput); ?>" class="input-premium py-1.5">
                </div>
                <div>
                    <label class="block text-xs font-medium text-gray-500 mb-1">Tanggal surat sampai</label>
                    <input type="date" name="tanggal_sampai" value="<?php echo htmlspecialchars($tanggalSampaiInput); ?>" class="input-premium py-1.5">
                </div>
            </div>
            <div class="grid grid-cols-1 md:grid-cols-4 gap-3">
                <div class="md:col-span-2">
                    <label class="block text-xs font-medium text-gray-500 mb-1">Cari Nomor/Keterangan</label>
                    <input type="text" name="search" value="<?php echo htmlspecialchars($search); ?>" class="input-premium py-1.5">
                </div>
                <div class="flex items-end gap-2 md:col-span-2">
                    <button type="submit" class="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg text-sm font-medium transition-colors">Tampilkan</button>
                    <a href="buku-nomor.php?tahun=<?php echo date('Y'); ?>" class="bg-gray-100 hover:bg-gray-200 text-gray-700 px-4 py-2 rounded-lg text-sm font-medium transition-colors">Reset</a>
                </div>
            </div>
            <p class="text-xs text-gray-400">Berdasarkan tanggal surat; entri tanpa tanggal tidak tampil saat filter tanggal aktif. Boleh isi hanya "dari" atau hanya "sampai".</p>
        </form>
    </div>

    <div class="card overflow-hidden">
        <div class="overflow-x-auto">
            <table class="w-full text-left table-premium">
                <thead>
                    <tr>
                        <th style="width: 50px;">No</th>
                        <th>Nomor</th>
                        <th>Tanggal Surat</th>
                        <th>Status</th>
                        <th>Tujuan/Keterangan</th>
                        <th class="text-right">Aksi</th>
                    </tr>
                </thead>
                <tbody>
                    <?php if (count($rows) > 0): ?>
                        <?php $no = $start + 1; foreach ($rows as $r): ?>
                            <tr>
                                <td class="text-gray-500"><?php echo $no++; ?></td>
                                <td class="font-semibold text-gray-900"><?php echo htmlspecialchars($r['nomor_surat']); ?></td>
                                <td class="text-gray-500">
                                    <?php echo !empty($r['tanggal_surat']) ? date('d/m/Y', strtotime($r['tanggal_surat'])) : '-'; ?>
                                </td>
                                <td>
                                    <?php
                                        $badge = 'badge-gray';
                                        if ($r['status'] === 'terisi') $badge = 'badge-green';
                                        elseif ($r['status'] === 'kosong') $badge = 'badge-blue';
                                    ?>
                                    <span class="badge <?php echo $badge; ?>">
                                        <?php echo ucfirst($r['status']); ?>
                                    </span>
                                </td>
                                <td class="text-gray-600">
                                    <?php
                                        $desc = $r['untuk'] ?: $r['keterangan'];
                                        echo htmlspecialchars($desc ?: '-');
                                    ?>
                                </td>
                                <td class="text-right">
                                    <?php if (empty($r['id_surat_tugas'])): ?>
                                        <div class="inline-flex gap-1.5">
                                            <form method="POST" class="inline">
                                                <input type="hidden" name="action" value="set_status">
                                                <input type="hidden" name="id" value="<?php echo (int)$r['id']; ?>">
                                                <input type="hidden" name="status" value="kosong">
                                                <button type="submit" class="btn-action">Kosongkan</button>
                                            </form>
                                            <form method="POST" class="inline" onsubmit="return confirm('Hapus nomor ini dari buku nomor?');">
                                                <input type="hidden" name="action" value="delete_slot">
                                                <input type="hidden" name="id" value="<?php echo (int)$r['id']; ?>">
                                                <button type="submit" class="btn-action danger">Hapus</button>
                                            </form>
                                        </div>
                                    <?php else: ?>
                                        <a href="generate-surat.php?id=<?php echo (int)$r['id_surat_tugas']; ?>" class="btn-action primary">
                                            <i class='bx bx-show'></i> Lihat Surat
                                        </a>
                                    <?php endif; ?>
                                </td>
                            </tr>
                        <?php endforeach; ?>
                    <?php else: ?>
                        <tr><td colspan="6" class="px-4 py-8 text-center text-gray-400">Belum ada data buku nomor.</td></tr>
                    <?php endif; ?>
                </tbody>
            </table>
        </div>

        <?php
        $paginationQuery = [
            'tahun' => $tahun,
            'status' => $status,
            'search' => $search,
        ];
        if ($tanggalDari !== null) {
            $paginationQuery['tanggal_dari'] = $tanggalDari;
        }
        if ($tanggalSampai !== null) {
            $paginationQuery['tanggal_sampai'] = $tanggalSampai;
        }
        $paginationBase = 'buku-nomor.php?' . http_build_query($paginationQuery, '', '&', PHP_QUERY_RFC3986);
        $fromRow = $totalRecords > 0 ? $start + 1 : 0;
        $toRow = $totalRecords > 0 ? min($start + count($rows), $totalRecords) : 0;
        $rangeLabel = $totalRecords > 0 ? "{$fromRow}–{$toRow}" : '0';
        ?>
        <div class="px-4 py-3 border-t border-slate-100 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 text-sm text-slate-600">
            <p class="text-center sm:text-left">
                Menampilkan <?php echo $rangeLabel; ?> dari <?php echo $totalRecords; ?> entri
                <span class="text-slate-400">(halaman <?php echo $page; ?> / <?php echo $totalPages; ?>)</span>
            </p>
            <?php if ($totalPages > 1): ?>
                <div class="flex justify-center flex-wrap gap-1">
                    <?php
                        $maxPageLinks = 7;
                        $half = (int)floor($maxPageLinks / 2);
                        $from = max(1, $page - $half);
                        $to = min($totalPages, $from + $maxPageLinks - 1);
                        $from = max(1, $to - $maxPageLinks + 1);
                    ?>
                    <?php if ($page > 1): ?>
                        <a href="<?php echo htmlspecialchars($paginationBase . '&page=' . ($page - 1)); ?>"
                           class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm">Sebelumnya</a>
                    <?php endif; ?>

                    <?php if ($from > 1): ?>
                        <a href="<?php echo htmlspecialchars($paginationBase . '&page=1'); ?>"
                           class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm">1</a>
                        <?php if ($from > 2): ?>
                            <span class="px-2 py-1 text-slate-400 text-sm">...</span>
                        <?php endif; ?>
                    <?php endif; ?>

                    <?php for ($i = $from; $i <= $to; $i++): ?>
                        <a href="<?php echo htmlspecialchars($paginationBase . '&page=' . $i); ?>"
                           class="px-3 py-1 rounded border text-sm <?php echo $page === $i ? 'bg-indigo-600 text-white border-indigo-600' : 'border-slate-200 text-slate-600 hover:bg-slate-50'; ?>">
                            <?php echo $i; ?>
                        </a>
                    <?php endfor; ?>

                    <?php if ($to < $totalPages): ?>
                        <?php if ($to < $totalPages - 1): ?>
                            <span class="px-2 py-1 text-slate-400 text-sm">...</span>
                        <?php endif; ?>
                        <a href="<?php echo htmlspecialchars($paginationBase . '&page=' . $totalPages); ?>"
                           class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm"><?php echo $totalPages; ?></a>
                    <?php endif; ?>

                    <?php if ($page < $totalPages): ?>
                        <a href="<?php echo htmlspecialchars($paginationBase . '&page=' . ($page + 1)); ?>"
                           class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm">Berikutnya</a>
                    <?php endif; ?>
                </div>
            <?php endif; ?>
        </div>
    </div>
</div>

<div id="spaceModal" class="fixed inset-0 z-50 hidden overflow-y-auto" aria-hidden="true">
    <div class="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
        <div class="fixed inset-0 bg-slate-900 bg-opacity-75 transition-opacity" onclick="closeModal('spaceModal')"></div>
        <span class="hidden sm:inline-block sm:align-middle sm:h-screen">&#8203;</span>
        <div class="inline-block align-bottom bg-white rounded-xl text-left overflow-hidden shadow-xl transform transition-all sm:my-8 sm:align-middle sm:max-w-lg sm:w-full">
            <div class="bg-white px-4 pt-5 pb-4 sm:p-6">
                <div class="flex justify-between items-center mb-4">
                    <h3 class="text-lg leading-6 font-medium text-slate-900">Buat Space Nomor per Tanggal</h3>
                    <button onclick="closeModal('spaceModal')" class="text-slate-400 hover:text-slate-500"><i class='bx bx-x text-2xl'></i></button>
                </div>
                <form method="POST" class="space-y-3">
                    <input type="hidden" name="action" value="create_space_tanggal">
                    <div>
                        <label class="block text-sm text-slate-600 mb-1">Tanggal Surat</label>
                        <input type="date" name="tanggal_surat" id="space-tanggal-surat" required class="w-full px-3 py-2 border border-slate-300 rounded-lg">
                    </div>
                    <div>
                        <p id="space-last-info" class="text-xs text-indigo-600">Pilih tanggal untuk melihat nomor terakhir tahunan.</p>
                        <p class="text-xs text-slate-500 mt-1">Tanggal sebelum tanggal terakhir buku tidak bisa di sini (nomor selalu ditambah setelah terakhir). Pakai <strong>Tambah Slot Kosong</strong> untuk isi mundur.</p>
                    </div>
                    <div>
                        <label class="block text-sm text-slate-600 mb-1">Jumlah Space</label>
                        <input type="number" name="jumlah_space" value="10" min="1" max="100" required class="w-full px-3 py-2 border border-slate-300 rounded-lg">
                    </div>
                    <div class="pt-2 flex justify-end gap-2">
                        <button type="button" onclick="closeModal('spaceModal')" class="px-4 py-2 bg-slate-200 text-slate-700 rounded-lg hover:bg-slate-300">Batal</button>
                        <button type="submit" class="px-4 py-2 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700">Buat Space</button>
                    </div>
                </form>
            </div>
        </div>
    </div>
</div>

<div id="reserveModal" class="fixed inset-0 z-50 hidden overflow-y-auto" aria-hidden="true">
    <div class="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
        <div class="fixed inset-0 bg-slate-900 bg-opacity-75 transition-opacity" onclick="closeModal('reserveModal')"></div>
        <span class="hidden sm:inline-block sm:align-middle sm:h-screen">&#8203;</span>
        <div class="inline-block align-bottom bg-white rounded-xl text-left overflow-hidden shadow-xl transform transition-all sm:my-8 sm:align-middle sm:max-w-lg sm:w-full">
            <div class="bg-white px-4 pt-5 pb-4 sm:p-6">
                <div class="flex justify-between items-center mb-4">
                    <h3 class="text-lg leading-6 font-medium text-slate-900">Tambah Slot Nomor Kosong</h3>
                    <button onclick="closeModal('reserveModal')" class="text-slate-400 hover:text-slate-500"><i class='bx bx-x text-2xl'></i></button>
                </div>
                <form method="POST" class="space-y-3">
                    <input type="hidden" name="action" value="reserve_slot">
                    <div>
                        <label class="block text-sm text-slate-600 mb-1">Tahun</label>
                        <input type="number" name="tahun" value="<?php echo $tahun; ?>" required class="w-full px-3 py-2 border border-slate-300 rounded-lg">
                    </div>
                    <div>
                        <label class="block text-sm text-slate-600 mb-1">Nomor Surat</label>
                        <input type="text" name="nomor_surat" id="reserve-nomor-surat" required class="w-full px-3 py-2 border border-slate-300 rounded-lg" placeholder="Contoh: 762">
                        <p id="reserve-last-info" class="text-xs text-indigo-600 mt-1">Nomor terakhir tahun ini: <?php echo $lastNomorTahun; ?>.</p>
                        <button type="button" onclick="isiNomorBerikutnya()" class="mt-2 px-2 py-1 text-xs bg-indigo-100 hover:bg-indigo-200 text-indigo-700 rounded">
                            Isi nomor berikutnya (<?php echo $lastNomorTahun + 1; ?>)
                        </button>
                    </div>
                    <div>
                        <label class="block text-sm text-slate-600 mb-1">Tanggal Surat</label>
                        <input type="date" name="tanggal_surat" required class="w-full px-3 py-2 border border-slate-300 rounded-lg">
                        <p class="text-xs text-slate-500 mt-1">Tanggal boleh mundur jika nomor (angka) lebih kecil dari nomor terakhir buku, tanggal lebih kecil dari tanggal terakhir, dan nomor belum dipakai surat.</p>
                    </div>
                    <div>
                        <label class="block text-sm text-slate-600 mb-1">Keterangan</label>
                        <input type="text" name="keterangan" class="w-full px-3 py-2 border border-slate-300 rounded-lg" value="Slot kosong manual">
                    </div>
                    <div class="pt-2 flex justify-end gap-2">
                        <button type="button" onclick="closeModal('reserveModal')" class="px-4 py-2 bg-slate-200 text-slate-700 rounded-lg hover:bg-slate-300">Batal</button>
                        <button type="submit" class="px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700">Simpan</button>
                    </div>
                </form>
            </div>
        </div>
    </div>
</div>

<script>
    async function updateSpaceLastInfo() {
        const tanggalInput = document.getElementById('space-tanggal-surat');
        const info = document.getElementById('space-last-info');
        if (!tanggalInput || !info) return;

        const tanggal = tanggalInput.value;
        if (!tanggal) {
            info.textContent = 'Pilih tanggal untuk melihat nomor terakhir.';
            return;
        }
        try {
            const resp = await fetch(`get_nomor_kosong.php?tanggal_surat=${encodeURIComponent(tanggal)}`);
            const data = await resp.json();
            if (data.success) {
                const lastNo = Number(data.last_nomor_tahun) || 0;
                info.textContent = lastNo > 0
                    ? `Nomor terakhir tahun ${data.tahun}: ${lastNo}. Space baru dimulai dari ${lastNo + 1} (tetap diset ke tanggal yang dipilih).`
                    : `Belum ada nomor tahun ${data.tahun}. Space baru dimulai dari 1.`;
            } else {
                info.textContent = 'Gagal membaca nomor terakhir.';
            }
        } catch (e) {
            info.textContent = 'Gagal membaca nomor terakhir.';
        }
    }

    function openModal(modalId) {
        document.getElementById(modalId).classList.remove('hidden');
        if (modalId === 'spaceModal') {
            const input = document.getElementById('space-tanggal-surat');
            if (input && !input.value) {
                const now = new Date();
                const m = String(now.getMonth() + 1).padStart(2, '0');
                const d = String(now.getDate()).padStart(2, '0');
                input.value = `${now.getFullYear()}-${m}-${d}`;
            }
            updateSpaceLastInfo();
        }
        if (modalId === 'reserveModal') {
            isiNomorBerikutnya();
        }
    }
    function closeModal(modalId) {
        document.getElementById(modalId).classList.add('hidden');
    }
    function isiNomorBerikutnya() {
        const input = document.getElementById('reserve-nomor-surat');
        const info = document.getElementById('reserve-last-info');
        if (!input || !info) return;
        const match = info.textContent.match(/(\d+)/g);
        const last = match && match.length ? Number(match[match.length - 1]) : null;
        if (last !== null && !Number.isNaN(last)) {
            input.value = String(last + 1);
        }
    }
    document.getElementById('space-tanggal-surat')?.addEventListener('change', updateSpaceLastInfo);
</script>

<?php include '../includes/footer.php'; ?>

