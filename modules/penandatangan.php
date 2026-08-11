<?php
error_reporting(E_ALL);
ini_set('display_errors', 1);

$pdo = require_once '../config/database.php';

// Pastikan tabel penandatangan ada
try {
    $pdo->query("SELECT 1 FROM penandatangan LIMIT 1");
} catch (PDOException $e) {
    // Buat tabel jika belum ada
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

// Pastikan kolom is_kepala dan jabatan_atasan ada
try {
    $pdo->query("SELECT is_kepala FROM penandatangan LIMIT 1");
} catch (PDOException $e) {
    try {
        $pdo->exec("ALTER TABLE penandatangan ADD COLUMN is_kepala TINYINT(1) DEFAULT 1 AFTER jabatan");
    } catch (PDOException $e2) {}
}
try {
    $pdo->query("SELECT jabatan_atasan FROM penandatangan LIMIT 1");
} catch (PDOException $e) {
    try {
        $pdo->exec("ALTER TABLE penandatangan ADD COLUMN jabatan_atasan VARCHAR(255) DEFAULT NULL AFTER is_kepala");
    } catch (PDOException $e2) {}
}

// Pastikan kolom tanda_tangan ada
try {
    $pdo->query("SELECT tanda_tangan FROM penandatangan LIMIT 1");
} catch (PDOException $e) {
    try {
        $pdo->exec("ALTER TABLE penandatangan ADD COLUMN tanda_tangan VARCHAR(255) DEFAULT NULL AFTER jabatan_atasan");
    } catch (PDOException $e2) {}
}

// Pastikan kolom id_penandatangan ada di surat_tugas
try {
    $pdo->query("SELECT id_penandatangan FROM surat_tugas LIMIT 1");
} catch (PDOException $e) {
    try {
        $pdo->exec("ALTER TABLE surat_tugas ADD COLUMN id_penandatangan INT DEFAULT NULL AFTER tanggal_selesai");
    } catch (PDOException $e2) {
        // Abaikan jika kolom sudah ada
    }
}

// Handle JSON responses (sebelum include header)
if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_POST['action'])) {
    // Set Default
    if ($_POST['action'] == 'set_default') {
        header('Content-Type: application/json');
        try {
            $pdo->beginTransaction();
            // Reset semua default
            $pdo->exec("UPDATE penandatangan SET is_default = 0");
            // Set default baru
            $stmt = $pdo->prepare("UPDATE penandatangan SET is_default = 1 WHERE id = ?");
            $stmt->execute([$_POST['id']]);
            $pdo->commit();
            echo json_encode(['success' => true, 'message' => 'Penanda tangan default berhasil diubah']);
        } catch (Exception $e) {
            if ($pdo->inTransaction()) $pdo->rollBack();
            echo json_encode(['success' => false, 'message' => $e->getMessage()]);
        }
        exit;
    }

    // Toggle Aktif
    if ($_POST['action'] == 'toggle_aktif') {
        header('Content-Type: application/json');
        try {
            $stmt = $pdo->prepare("UPDATE penandatangan SET aktif = NOT aktif WHERE id = ?");
            $stmt->execute([$_POST['id']]);
            echo json_encode(['success' => true, 'message' => 'Status penanda tangan berhasil diubah']);
        } catch (Exception $e) {
            echo json_encode(['success' => false, 'message' => $e->getMessage()]);
        }
        exit;
    }

    // Delete
    if ($_POST['action'] == 'delete') {
        header('Content-Type: application/json');
        try {
            // Cek apakah penandatangan sedang digunakan
            $stmt = $pdo->prepare("SELECT COUNT(*) FROM surat_tugas WHERE id_penandatangan = ?");
            $stmt->execute([$_POST['id']]);
            $count = $stmt->fetchColumn();
            
            if ($count > 0) {
                throw new Exception("Penanda tangan ini sedang digunakan di {$count} surat tugas dan tidak dapat dihapus. Nonaktifkan saja jika tidak ingin digunakan lagi.");
            }
            
            // Hapus file tanda tangan jika ada
            $stmt_get = $pdo->prepare("SELECT tanda_tangan FROM penandatangan WHERE id = ?");
            $stmt_get->execute([$_POST['id']]);
            $pt_data = $stmt_get->fetch();
            if ($pt_data && !empty($pt_data['tanda_tangan'])) {
                $file_path = __DIR__ . '/../assets/img/tanda_tangan/' . $pt_data['tanda_tangan'];
                if (file_exists($file_path)) {
                    unlink($file_path);
                }
            }
            
            $stmt = $pdo->prepare("DELETE FROM penandatangan WHERE id = ?");
            $stmt->execute([$_POST['id']]);
            echo json_encode(['success' => true, 'message' => 'Penanda tangan berhasil dihapus']);
        } catch (Exception $e) {
            echo json_encode(['success' => false, 'message' => $e->getMessage()]);
        }
        exit;
    }
}

// Include header
include '../includes/header.php';

// Handle Create & Update (setelah header karena butuh HTML)
if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_POST['action'])) {
    if ($_POST['action'] == 'create') {
        try {
            if (empty($_POST['nama']) || empty($_POST['nip']) || empty($_POST['jabatan'])) {
                throw new Exception("Nama, NIP, dan Jabatan wajib diisi.");
            }

            $is_default = isset($_POST['is_default']) ? 1 : 0;
            $is_kepala = isset($_POST['is_kepala']) ? (int)$_POST['is_kepala'] : 1;
            $jabatan_atasan = ($is_kepala == 0 && !empty($_POST['jabatan_atasan'])) ? trim($_POST['jabatan_atasan']) : null;
            
            // Handle upload tanda tangan
            $tanda_tangan_filename = null;
            if (isset($_FILES['tanda_tangan']) && $_FILES['tanda_tangan']['error'] === UPLOAD_ERR_OK) {
                $allowed_types = ['image/png'];
                $file_type = $_FILES['tanda_tangan']['type'];
                $file_ext = strtolower(pathinfo($_FILES['tanda_tangan']['name'], PATHINFO_EXTENSION));
                
                if (!in_array($file_type, $allowed_types) && $file_ext !== 'png') {
                    throw new Exception("Format file tanda tangan harus PNG.");
                }
                
                if ($_FILES['tanda_tangan']['size'] > 2 * 1024 * 1024) {
                    throw new Exception("Ukuran file tanda tangan maksimal 2MB.");
                }
                
                $tanda_tangan_filename = 'ttd_' . time() . '_' . uniqid() . '.png';
                $upload_path = __DIR__ . '/../assets/img/tanda_tangan/' . $tanda_tangan_filename;
                
                if (!move_uploaded_file($_FILES['tanda_tangan']['tmp_name'], $upload_path)) {
                    throw new Exception("Gagal mengupload file tanda tangan.");
                }
            }
            
            if ($is_default) {
                $pdo->exec("UPDATE penandatangan SET is_default = 0");
            }

            $stmt = $pdo->prepare("INSERT INTO penandatangan (nama, nip, pangkat, jabatan, is_kepala, jabatan_atasan, tanda_tangan, is_default) VALUES (?, ?, ?, ?, ?, ?, ?, ?)");
            $stmt->execute([
                $_POST['nama'],
                $_POST['nip'],
                $_POST['pangkat'] ?? '',
                $_POST['jabatan'],
                $is_kepala,
                $jabatan_atasan,
                $tanda_tangan_filename,
                $is_default
            ]);

            echo "<script>
                Swal.fire({
                    title: 'Berhasil!',
                    text: 'Penanda tangan berhasil ditambahkan',
                    icon: 'success',
                    confirmButtonText: 'OK'
                }).then(() => { window.location.href = 'penandatangan.php'; });
            </script>";
            include '../includes/footer.php';
            exit;
        } catch (Exception $e) {
            echo "<script>
                Swal.fire({
                    title: 'Gagal!',
                    text: " . json_encode($e->getMessage()) . ",
                    icon: 'error',
                    confirmButtonText: 'OK'
                }).then(() => { window.location.href = 'penandatangan.php'; });
            </script>";
            include '../includes/footer.php';
            exit;
        }
    }

    if ($_POST['action'] == 'update') {
        try {
            if (empty($_POST['nama']) || empty($_POST['nip']) || empty($_POST['jabatan'])) {
                throw new Exception("Nama, NIP, dan Jabatan wajib diisi.");
            }

            $is_default = isset($_POST['is_default']) ? 1 : 0;
            $is_kepala = isset($_POST['is_kepala']) ? (int)$_POST['is_kepala'] : 1;
            $jabatan_atasan = ($is_kepala == 0 && !empty($_POST['jabatan_atasan'])) ? trim($_POST['jabatan_atasan']) : null;
            
            // Handle upload tanda tangan
            $tanda_tangan_filename = null;
            $update_ttd = false;
            
            // Cek apakah user ingin menghapus tanda tangan
            if (isset($_POST['hapus_tanda_tangan']) && $_POST['hapus_tanda_tangan'] == '1') {
                $update_ttd = true;
                $tanda_tangan_filename = null;
                // Hapus file lama
                $stmt_old = $pdo->prepare("SELECT tanda_tangan FROM penandatangan WHERE id = ?");
                $stmt_old->execute([$_POST['id']]);
                $old_data = $stmt_old->fetch();
                if ($old_data && !empty($old_data['tanda_tangan'])) {
                    $old_file = __DIR__ . '/../assets/img/tanda_tangan/' . $old_data['tanda_tangan'];
                    if (file_exists($old_file)) unlink($old_file);
                }
            }
            
            if (isset($_FILES['tanda_tangan']) && $_FILES['tanda_tangan']['error'] === UPLOAD_ERR_OK) {
                $allowed_types = ['image/png'];
                $file_type = $_FILES['tanda_tangan']['type'];
                $file_ext = strtolower(pathinfo($_FILES['tanda_tangan']['name'], PATHINFO_EXTENSION));
                
                if (!in_array($file_type, $allowed_types) && $file_ext !== 'png') {
                    throw new Exception("Format file tanda tangan harus PNG.");
                }
                
                if ($_FILES['tanda_tangan']['size'] > 2 * 1024 * 1024) {
                    throw new Exception("Ukuran file tanda tangan maksimal 2MB.");
                }
                
                // Hapus file lama
                $stmt_old = $pdo->prepare("SELECT tanda_tangan FROM penandatangan WHERE id = ?");
                $stmt_old->execute([$_POST['id']]);
                $old_data = $stmt_old->fetch();
                if ($old_data && !empty($old_data['tanda_tangan'])) {
                    $old_file = __DIR__ . '/../assets/img/tanda_tangan/' . $old_data['tanda_tangan'];
                    if (file_exists($old_file)) unlink($old_file);
                }
                
                $tanda_tangan_filename = 'ttd_' . time() . '_' . uniqid() . '.png';
                $upload_path = __DIR__ . '/../assets/img/tanda_tangan/' . $tanda_tangan_filename;
                
                if (!move_uploaded_file($_FILES['tanda_tangan']['tmp_name'], $upload_path)) {
                    throw new Exception("Gagal mengupload file tanda tangan.");
                }
                $update_ttd = true;
            }
            
            if ($is_default) {
                $pdo->exec("UPDATE penandatangan SET is_default = 0");
            }

            if ($update_ttd) {
                $stmt = $pdo->prepare("UPDATE penandatangan SET nama = ?, nip = ?, pangkat = ?, jabatan = ?, is_kepala = ?, jabatan_atasan = ?, tanda_tangan = ?, is_default = ? WHERE id = ?");
                $stmt->execute([
                    $_POST['nama'],
                    $_POST['nip'],
                    $_POST['pangkat'] ?? '',
                    $_POST['jabatan'],
                    $is_kepala,
                    $jabatan_atasan,
                    $tanda_tangan_filename,
                    $is_default,
                    $_POST['id']
                ]);
            } else {
                $stmt = $pdo->prepare("UPDATE penandatangan SET nama = ?, nip = ?, pangkat = ?, jabatan = ?, is_kepala = ?, jabatan_atasan = ?, is_default = ? WHERE id = ?");
                $stmt->execute([
                    $_POST['nama'],
                    $_POST['nip'],
                    $_POST['pangkat'] ?? '',
                    $_POST['jabatan'],
                    $is_kepala,
                    $jabatan_atasan,
                    $is_default,
                    $_POST['id']
                ]);
            }

            echo "<script>
                Swal.fire({
                    title: 'Berhasil!',
                    text: 'Penanda tangan berhasil diupdate',
                    icon: 'success',
                    confirmButtonText: 'OK'
                }).then(() => { window.location.href = 'penandatangan.php'; });
            </script>";
            include '../includes/footer.php';
            exit;
        } catch (Exception $e) {
            echo "<script>
                Swal.fire({
                    title: 'Gagal!',
                    text: " . json_encode($e->getMessage()) . ",
                    icon: 'error',
                    confirmButtonText: 'OK'
                }).then(() => { window.location.href = 'penandatangan.php'; });
            </script>";
            include '../includes/footer.php';
            exit;
        }
    }
}

// Ambil semua penandatangan
$stmt = $pdo->query("SELECT * FROM penandatangan ORDER BY is_default DESC, aktif DESC, nama ASC");
$penandatangan_list = $stmt->fetchAll();

// Ambil semua pegawai untuk dropdown "Ambil dari Data Pegawai"
$stmt_pegawai = $pdo->query("SELECT nip, nama, pangkat, jabatan FROM pegawai ORDER BY nama ASC");
$all_pegawai = $stmt_pegawai->fetchAll();
?>

<div class="space-y-6">
    <div class="flex flex-col md:flex-row justify-between items-center gap-4">
        <div>
            <h2 class="text-2xl font-bold text-slate-800">Penanda Tangan</h2>
            <p class="text-sm text-slate-500 mt-1">Kelola data pejabat penanda tangan surat tugas</p>
        </div>
        <button onclick="openModal('addModal')" class="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg flex items-center gap-2 transition-colors shadow-sm">
            <i class='bx bx-plus'></i> Tambah Penanda Tangan
        </button>
    </div>

    <!-- Info Card -->
    <div class="bg-gradient-to-r from-indigo-50 to-blue-50 rounded-xl border border-indigo-100 p-4">
        <div class="flex items-start gap-3">
            <div class="flex-shrink-0 bg-indigo-100 rounded-lg p-2">
                <i class='bx bx-info-circle text-indigo-600 text-xl'></i>
            </div>
            <div>
                <h4 class="text-sm font-semibold text-indigo-800 mb-1">Tentang Penanda Tangan</h4>
                <p class="text-sm text-indigo-700 mb-2">
                    Penanda tangan yang ditandai sebagai <strong>Default</strong> akan otomatis terpilih saat membuat surat tugas baru.
                </p>
                <p class="text-sm text-indigo-700 mb-1 font-semibold">Placeholder dasar:</p>
                <p class="text-sm text-indigo-700 mb-2">
                    <code class="bg-indigo-100 px-1.5 py-0.5 rounded text-xs font-mono">${penandatangan_nama}</code>, 
                    <code class="bg-indigo-100 px-1.5 py-0.5 rounded text-xs font-mono">${penandatangan_nip}</code>, 
                    <code class="bg-indigo-100 px-1.5 py-0.5 rounded text-xs font-mono">${penandatangan_pangkat}</code>
                </p>
                <p class="text-sm text-indigo-700 mb-1 font-semibold">Placeholder otomatis (Kepala vs A.n):</p>
                <p class="text-sm text-indigo-700">
                    <code class="bg-indigo-100 px-1.5 py-0.5 rounded text-xs font-mono">${penandatangan_header_jabatan}</code> → Otomatis: jabatan (Kepala) atau "A.n [jabatan atasan]" (A.n)<br>
                    <code class="bg-indigo-100 px-1.5 py-0.5 rounded text-xs font-mono">${penandatangan_sub_jabatan}</code> → Kosong (Kepala) atau jabatan penandatangan (A.n)
                </p>
                <p class="text-sm text-indigo-700 mt-2 mb-1 font-semibold">Placeholder gambar tanda tangan:</p>
                <p class="text-sm text-indigo-700">
                    <code class="bg-indigo-100 px-1.5 py-0.5 rounded text-xs font-mono">${tanda_tangan}</code> → Gambar tanda tangan PNG (disisipkan otomatis ke template)
                </p>
            </div>
        </div>
    </div>

    <!-- Tabel Penanda Tangan -->
    <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
        <?php if (count($penandatangan_list) > 0): ?>
            <div class="overflow-x-auto">
                <table class="w-full text-left border-collapse">
                    <thead>
                        <tr class="bg-slate-50 text-slate-600 text-sm uppercase tracking-wider">
                            <th class="px-6 py-4 font-semibold border-b border-slate-100">No</th>
                            <th class="px-6 py-4 font-semibold border-b border-slate-100">Nama</th>
                            <th class="px-6 py-4 font-semibold border-b border-slate-100">NIP</th>
                            <th class="px-6 py-4 font-semibold border-b border-slate-100">Pangkat</th>
                            <th class="px-6 py-4 font-semibold border-b border-slate-100">Jabatan</th>
                            <th class="px-6 py-4 font-semibold border-b border-slate-100">Tanda Tangan</th>
                            <th class="px-6 py-4 font-semibold border-b border-slate-100">Status</th>
                            <th class="px-6 py-4 font-semibold border-b border-slate-100 text-right">Aksi</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-100">
                        <?php $no = 1; foreach ($penandatangan_list as $pt): ?>
                            <tr class="hover:bg-slate-50 transition-colors <?php echo !$pt['aktif'] ? 'opacity-50' : ''; ?>">
                                <td class="px-6 py-4 text-sm text-slate-600"><?php echo $no++; ?></td>
                                <td class="px-6 py-4">
                                    <div class="flex items-center gap-2 flex-wrap">
                                        <span class="text-sm font-medium text-slate-900"><?php echo htmlspecialchars($pt['nama']); ?></span>
                                        <?php if ($pt['is_default']): ?>
                                            <span class="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-emerald-100 text-emerald-800">
                                                <i class='bx bx-check-circle mr-1'></i> Default
                                            </span>
                                        <?php endif; ?>
                                        <?php
                                        $is_kpl = isset($pt['is_kepala']) ? $pt['is_kepala'] : 1;
                                        if ($is_kpl): ?>
                                            <span class="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-blue-100 text-blue-800">
                                                <i class='bx bx-crown mr-1'></i> Kepala
                                            </span>
                                        <?php else: ?>
                                            <span class="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-amber-100 text-amber-800">
                                                <i class='bx bx-user mr-1'></i> A.n
                                            </span>
                                        <?php endif; ?>
                                    </div>
                                </td>
                                <td class="px-6 py-4 text-sm font-mono text-slate-600"><?php echo htmlspecialchars($pt['nip']); ?></td>
                                <td class="px-6 py-4 text-sm text-slate-600"><?php echo htmlspecialchars($pt['pangkat'] ?? '-'); ?></td>
                                <td class="px-6 py-4 text-sm text-slate-600">
                                    <?php echo htmlspecialchars($pt['jabatan']); ?>
                                    <?php if (isset($pt['is_kepala']) && !$pt['is_kepala'] && !empty($pt['jabatan_atasan'])): ?>
                                        <br><span class="text-xs text-amber-600 italic">A.n <?php echo htmlspecialchars($pt['jabatan_atasan']); ?></span>
                                    <?php endif; ?>
                                </td>
                                <td class="px-6 py-4">
                                    <?php if (!empty($pt['tanda_tangan'])): ?>
                                        <div class="relative group">
                                            <img src="../assets/img/tanda_tangan/<?php echo htmlspecialchars($pt['tanda_tangan']); ?>" 
                                                 alt="Tanda Tangan" 
                                                 class="h-12 w-auto rounded border border-slate-200 bg-white p-1 cursor-pointer hover:shadow-md transition-shadow"
                                                 onclick="openPreviewTTD(this.src)">
                                            <span class="absolute -top-1 -right-1 bg-green-500 rounded-full w-4 h-4 flex items-center justify-center">
                                                <i class='bx bx-check text-white text-xs'></i>
                                            </span>
                                        </div>
                                    <?php else: ?>
                                        <span class="text-xs text-slate-400 italic">Belum ada</span>
                                    <?php endif; ?>
                                </td>
                                <td class="px-6 py-4">
                                    <?php if ($pt['aktif']): ?>
                                        <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-800">
                                            <i class='bx bx-check mr-1'></i> Aktif
                                        </span>
                                    <?php else: ?>
                                        <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-slate-100 text-slate-600">
                                            <i class='bx bx-x mr-1'></i> Nonaktif
                                        </span>
                                    <?php endif; ?>
                                </td>
                                <td class="px-6 py-4 text-sm text-right">
                                    <div class="flex items-center justify-end gap-1">
                                        <?php if (!$pt['is_default'] && $pt['aktif']): ?>
                                            <button class="text-emerald-500 hover:text-emerald-600 transition-colors p-1 set-default-btn" 
                                                    title="Set sebagai Default" data-id="<?php echo $pt['id']; ?>">
                                                <i class='bx bx-star text-xl'></i>
                                            </button>
                                        <?php elseif ($pt['is_default']): ?>
                                            <span class="text-emerald-400 p-1" title="Sudah Default">
                                                <i class='bx bxs-star text-xl'></i>
                                            </span>
                                        <?php endif; ?>
                                        <button class="text-blue-500 hover:text-blue-600 transition-colors p-1 toggle-aktif-btn" 
                                                title="<?php echo $pt['aktif'] ? 'Nonaktifkan' : 'Aktifkan'; ?>" 
                                                data-id="<?php echo $pt['id']; ?>" data-aktif="<?php echo $pt['aktif']; ?>">
                                            <i class='bx <?php echo $pt['aktif'] ? 'bx-toggle-right text-green-500' : 'bx-toggle-left text-slate-400'; ?> text-xl'></i>
                                        </button>
                                        <button class="text-amber-500 hover:text-amber-600 transition-colors p-1 edit-btn" title="Edit"
                                                data-id="<?php echo $pt['id']; ?>"
                                                data-nama="<?php echo htmlspecialchars($pt['nama']); ?>"
                                                data-nip="<?php echo htmlspecialchars($pt['nip']); ?>"
                                                data-pangkat="<?php echo htmlspecialchars($pt['pangkat'] ?? ''); ?>"
                                                data-jabatan="<?php echo htmlspecialchars($pt['jabatan']); ?>"
                                                data-is-kepala="<?php echo isset($pt['is_kepala']) ? $pt['is_kepala'] : 1; ?>"
                                                data-jabatan-atasan="<?php echo htmlspecialchars($pt['jabatan_atasan'] ?? ''); ?>"
                                                data-tanda-tangan="<?php echo htmlspecialchars($pt['tanda_tangan'] ?? ''); ?>"
                                                data-default="<?php echo $pt['is_default']; ?>">
                                            <i class='bx bx-edit text-xl'></i>
                                        </button>
                                        <button class="text-red-500 hover:text-red-600 transition-colors p-1 delete-btn" title="Hapus" 
                                                data-id="<?php echo $pt['id']; ?>">
                                            <i class='bx bx-trash text-xl'></i>
                                        </button>
                                    </div>
                                </td>
                            </tr>
                        <?php endforeach; ?>
                    </tbody>
                </table>
            </div>
        <?php else: ?>
            <div class="p-12 text-center">
                <div class="bg-slate-100 rounded-full w-20 h-20 flex items-center justify-center mx-auto mb-4">
                    <i class='bx bx-pen text-4xl text-slate-400'></i>
                </div>
                <h3 class="text-lg font-semibold text-slate-700 mb-2">Belum Ada Penanda Tangan</h3>
                <p class="text-slate-500 mb-4">Tambahkan pejabat penanda tangan untuk digunakan di surat tugas</p>
                <button onclick="openModal('addModal')" class="inline-flex items-center gap-2 px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition-colors">
                    <i class='bx bx-plus'></i> Tambah Penanda Tangan Pertama
                </button>
            </div>
        <?php endif; ?>
    </div>
</div>

<!-- Modal Tambah -->
<div id="addModal" class="fixed inset-0 z-50 hidden overflow-y-auto" role="dialog" aria-modal="true">
    <div class="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
        <div class="fixed inset-0 bg-slate-900 bg-opacity-75 transition-opacity" onclick="closeModal('addModal')"></div>
        <span class="hidden sm:inline-block sm:align-middle sm:h-screen">&#8203;</span>
        <div class="inline-block align-bottom bg-white rounded-xl text-left overflow-hidden shadow-xl transform transition-all sm:my-8 sm:align-middle sm:max-w-lg sm:w-full">
            <div class="bg-white px-6 pt-6 pb-4">
                <div class="flex justify-between items-center mb-5">
                    <h3 class="text-lg font-semibold text-slate-900">Tambah Penanda Tangan</h3>
                    <button onclick="closeModal('addModal')" class="text-slate-400 hover:text-slate-500">
                        <i class='bx bx-x text-2xl'></i>
                    </button>
                </div>
                <form method="POST" id="formTambah" enctype="multipart/form-data">
                    <input type="hidden" name="action" value="create">
                    <div class="space-y-4">
                        <!-- Ambil dari Data Pegawai -->
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">
                                <i class='bx bx-user-plus text-indigo-500 mr-1'></i>Ambil dari Data Pegawai
                            </label>
                            <select id="add-pilih-pegawai" class="select2-pegawai-add w-full">
                                <option value="">-- Pilih pegawai untuk auto-fill --</option>
                                <?php foreach ($all_pegawai as $p): ?>
                                    <option value="<?php echo htmlspecialchars($p['nip']); ?>"
                                            data-nama="<?php echo htmlspecialchars($p['nama']); ?>"
                                            data-nip="<?php echo htmlspecialchars($p['nip']); ?>"
                                            data-pangkat="<?php echo htmlspecialchars($p['pangkat'] ?? ''); ?>"
                                            data-jabatan="<?php echo htmlspecialchars($p['jabatan']); ?>">
                                        <?php echo htmlspecialchars($p['nama']) . ' - ' . htmlspecialchars($p['nip']); ?>
                                    </option>
                                <?php endforeach; ?>
                            </select>
                            <p class="text-xs text-slate-500 mt-1">Data akan otomatis terisi, Anda tetap bisa mengedit secara manual</p>
                        </div>

                        <div class="border-t border-slate-200 pt-4">
                            <p class="text-xs font-semibold text-slate-400 uppercase tracking-wider mb-3">Data Penanda Tangan</p>
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Nama Lengkap <span class="text-red-500">*</span></label>
                            <input type="text" name="nama" required placeholder="Contoh: Dr. Ahmad Fauzi, S.Hut., M.Si." 
                                   class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">NIP <span class="text-red-500">*</span></label>
                            <input type="text" name="nip" required placeholder="Contoh: 197501012000031001" 
                                   class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Pangkat/Golongan</label>
                            <input type="text" name="pangkat" placeholder="Contoh: Pembina Tk.I / IV/b" 
                                   class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Jabatan <span class="text-red-500">*</span></label>
                            <input type="text" name="jabatan" required placeholder="Contoh: Kepala Dinas Kehutanan" 
                                   class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                        <!-- Tipe Penanda Tangan -->
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-2">Tipe Penanda Tangan <span class="text-red-500">*</span></label>
                            <div class="grid grid-cols-2 gap-3">
                                <label class="flex items-center gap-3 p-3 border-2 rounded-lg cursor-pointer transition-all add-tipe-option border-indigo-500 bg-indigo-50" data-value="1">
                                    <input type="radio" name="is_kepala" value="1" checked class="h-4 w-4 text-indigo-600 border-slate-300 focus:ring-indigo-500">
                                    <div>
                                        <span class="text-sm font-medium text-slate-800"><i class='bx bx-crown text-blue-600 mr-1'></i>Kepala</span>
                                        <p class="text-xs text-slate-500">Menandatangani langsung</p>
                                    </div>
                                </label>
                                <label class="flex items-center gap-3 p-3 border-2 rounded-lg cursor-pointer transition-all add-tipe-option border-slate-200" data-value="0">
                                    <input type="radio" name="is_kepala" value="0" class="h-4 w-4 text-indigo-600 border-slate-300 focus:ring-indigo-500">
                                    <div>
                                        <span class="text-sm font-medium text-slate-800"><i class='bx bx-user text-amber-600 mr-1'></i>A.n (Atas Nama)</span>
                                        <p class="text-xs text-slate-500">Menandatangani atas nama atasan</p>
                                    </div>
                                </label>
                            </div>
                        </div>
                        <!-- Field Jabatan Atasan (muncul jika A.n) -->
                        <div id="add-jabatan-atasan-wrap" class="hidden">
                            <label class="block text-sm font-medium text-slate-700 mb-1">
                                <i class='bx bx-buildings text-amber-500 mr-1'></i>Jabatan Atasan <span class="text-red-500">*</span>
                            </label>
                            <input type="text" name="jabatan_atasan" id="add-jabatan-atasan" placeholder="Contoh: Kepala Cabang Dinas Kehutanan Wilayah Bojonegoro" 
                                   class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                            <p class="text-xs text-slate-500 mt-1">Jabatan pejabat yang diatasnamakan (akan tampil setelah "A.n")</p>
                        </div>

                        <!-- Upload Tanda Tangan -->
                        <div class="border-t border-slate-200 pt-4">
                            <p class="text-xs font-semibold text-slate-400 uppercase tracking-wider mb-3">Gambar Tanda Tangan</p>
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">
                                <i class='bx bx-pen text-indigo-500 mr-1'></i>Upload Tanda Tangan (PNG)
                            </label>
                            <div class="relative">
                                <input type="file" name="tanda_tangan" id="add-tanda-tangan" accept="image/png" 
                                       class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none text-sm"
                                       onchange="previewTTD(this, 'add-ttd-preview')">
                            </div>
                            <p class="text-xs text-slate-500 mt-1">Format: PNG, ukuran maksimal 2MB. Sebaiknya gunakan gambar dengan background transparan.</p>
                            <div id="add-ttd-preview" class="mt-2 hidden">
                                <div class="inline-block relative border border-slate-200 rounded-lg p-2 bg-white">
                                    <img id="add-ttd-preview-img" src="" alt="Preview" class="h-20 w-auto">
                                    <button type="button" onclick="clearTTDPreview('add')" class="absolute -top-2 -right-2 bg-red-500 text-white rounded-full w-5 h-5 flex items-center justify-center text-xs hover:bg-red-600">
                                        <i class='bx bx-x'></i>
                                    </button>
                                </div>
                            </div>
                        </div>

                        <div class="flex items-center gap-2">
                            <input type="checkbox" name="is_default" id="add_is_default" class="h-4 w-4 text-indigo-600 border-slate-300 rounded focus:ring-indigo-500">
                            <label for="add_is_default" class="text-sm text-slate-700">Jadikan sebagai penanda tangan default</label>
                        </div>
                    </div>
                    <div class="mt-6 flex justify-end gap-3">
                        <button type="button" onclick="closeModal('addModal')" class="px-4 py-2 border border-slate-300 rounded-lg text-slate-700 hover:bg-slate-50 font-medium">Batal</button>
                        <button type="submit" class="px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 font-medium">Simpan</button>
                    </div>
                </form>
            </div>
        </div>
    </div>
</div>

<!-- Modal Edit -->
<div id="editModal" class="fixed inset-0 z-50 hidden overflow-y-auto" role="dialog" aria-modal="true">
    <div class="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
        <div class="fixed inset-0 bg-slate-900 bg-opacity-75 transition-opacity" onclick="closeModal('editModal')"></div>
        <span class="hidden sm:inline-block sm:align-middle sm:h-screen">&#8203;</span>
        <div class="inline-block align-bottom bg-white rounded-xl text-left overflow-hidden shadow-xl transform transition-all sm:my-8 sm:align-middle sm:max-w-lg sm:w-full">
            <div class="bg-white px-6 pt-6 pb-4">
                <div class="flex justify-between items-center mb-5">
                    <h3 class="text-lg font-semibold text-slate-900">Edit Penanda Tangan</h3>
                    <button onclick="closeModal('editModal')" class="text-slate-400 hover:text-slate-500">
                        <i class='bx bx-x text-2xl'></i>
                    </button>
                </div>
                <form method="POST" id="formEdit" enctype="multipart/form-data">
                    <input type="hidden" name="action" value="update">
                    <input type="hidden" name="id" id="edit-id">
                    <input type="hidden" name="hapus_tanda_tangan" id="edit-hapus-ttd" value="0">
                    <div class="space-y-4">
                        <!-- Ambil dari Data Pegawai -->
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">
                                <i class='bx bx-user-plus text-indigo-500 mr-1'></i>Ambil dari Data Pegawai
                            </label>
                            <select id="edit-pilih-pegawai" class="select2-pegawai-edit w-full">
                                <option value="">-- Pilih pegawai untuk auto-fill --</option>
                                <?php foreach ($all_pegawai as $p): ?>
                                    <option value="<?php echo htmlspecialchars($p['nip']); ?>"
                                            data-nama="<?php echo htmlspecialchars($p['nama']); ?>"
                                            data-nip="<?php echo htmlspecialchars($p['nip']); ?>"
                                            data-pangkat="<?php echo htmlspecialchars($p['pangkat'] ?? ''); ?>"
                                            data-jabatan="<?php echo htmlspecialchars($p['jabatan']); ?>">
                                        <?php echo htmlspecialchars($p['nama']) . ' - ' . htmlspecialchars($p['nip']); ?>
                                    </option>
                                <?php endforeach; ?>
                            </select>
                            <p class="text-xs text-slate-500 mt-1">Data akan otomatis terisi, Anda tetap bisa mengedit secara manual</p>
                        </div>

                        <div class="border-t border-slate-200 pt-4">
                            <p class="text-xs font-semibold text-slate-400 uppercase tracking-wider mb-3">Data Penanda Tangan</p>
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Nama Lengkap <span class="text-red-500">*</span></label>
                            <input type="text" name="nama" id="edit-nama" required 
                                   class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">NIP <span class="text-red-500">*</span></label>
                            <input type="text" name="nip" id="edit-nip" required 
                                   class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Pangkat/Golongan</label>
                            <input type="text" name="pangkat" id="edit-pangkat" 
                                   class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Jabatan <span class="text-red-500">*</span></label>
                            <input type="text" name="jabatan" id="edit-jabatan" required 
                                   class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                        <!-- Tipe Penanda Tangan -->
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-2">Tipe Penanda Tangan <span class="text-red-500">*</span></label>
                            <div class="grid grid-cols-2 gap-3">
                                <label class="flex items-center gap-3 p-3 border-2 rounded-lg cursor-pointer transition-all edit-tipe-option border-indigo-500 bg-indigo-50" data-value="1">
                                    <input type="radio" name="is_kepala" value="1" id="edit-is-kepala-1" checked class="h-4 w-4 text-indigo-600 border-slate-300 focus:ring-indigo-500">
                                    <div>
                                        <span class="text-sm font-medium text-slate-800"><i class='bx bx-crown text-blue-600 mr-1'></i>Kepala</span>
                                        <p class="text-xs text-slate-500">Menandatangani langsung</p>
                                    </div>
                                </label>
                                <label class="flex items-center gap-3 p-3 border-2 rounded-lg cursor-pointer transition-all edit-tipe-option border-slate-200" data-value="0">
                                    <input type="radio" name="is_kepala" value="0" id="edit-is-kepala-0" class="h-4 w-4 text-indigo-600 border-slate-300 focus:ring-indigo-500">
                                    <div>
                                        <span class="text-sm font-medium text-slate-800"><i class='bx bx-user text-amber-600 mr-1'></i>A.n (Atas Nama)</span>
                                        <p class="text-xs text-slate-500">Menandatangani atas nama atasan</p>
                                    </div>
                                </label>
                            </div>
                        </div>
                        <!-- Field Jabatan Atasan (muncul jika A.n) -->
                        <div id="edit-jabatan-atasan-wrap" class="hidden">
                            <label class="block text-sm font-medium text-slate-700 mb-1">
                                <i class='bx bx-buildings text-amber-500 mr-1'></i>Jabatan Atasan <span class="text-red-500">*</span>
                            </label>
                            <input type="text" name="jabatan_atasan" id="edit-jabatan-atasan" placeholder="Contoh: Kepala Cabang Dinas Kehutanan Wilayah Bojonegoro" 
                                   class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                            <p class="text-xs text-slate-500 mt-1">Jabatan pejabat yang diatasnamakan</p>
                        </div>

                        <!-- Upload Tanda Tangan -->
                        <div class="border-t border-slate-200 pt-4">
                            <p class="text-xs font-semibold text-slate-400 uppercase tracking-wider mb-3">Gambar Tanda Tangan</p>
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">
                                <i class='bx bx-pen text-indigo-500 mr-1'></i>Upload Tanda Tangan (PNG)
                            </label>
                            <div id="edit-ttd-current" class="mb-2 hidden">
                                <p class="text-xs text-slate-500 mb-1">Tanda tangan saat ini:</p>
                                <div class="inline-block relative border border-slate-200 rounded-lg p-2 bg-white">
                                    <img id="edit-ttd-current-img" src="" alt="Tanda Tangan Saat Ini" class="h-16 w-auto">
                                    <button type="button" onclick="removeTTDEdit()" class="absolute -top-2 -right-2 bg-red-500 text-white rounded-full w-5 h-5 flex items-center justify-center text-xs hover:bg-red-600" title="Hapus Tanda Tangan">
                                        <i class='bx bx-x'></i>
                                    </button>
                                </div>
                            </div>
                            <div class="relative">
                                <input type="file" name="tanda_tangan" id="edit-tanda-tangan" accept="image/png" 
                                       class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none text-sm"
                                       onchange="previewTTD(this, 'edit-ttd-preview')">
                            </div>
                            <p class="text-xs text-slate-500 mt-1">Format: PNG, ukuran maksimal 2MB. Kosongkan jika tidak ingin mengubah.</p>
                            <div id="edit-ttd-preview" class="mt-2 hidden">
                                <div class="inline-block relative border border-slate-200 rounded-lg p-2 bg-white">
                                    <img id="edit-ttd-preview-img" src="" alt="Preview" class="h-20 w-auto">
                                    <button type="button" onclick="clearTTDPreview('edit')" class="absolute -top-2 -right-2 bg-red-500 text-white rounded-full w-5 h-5 flex items-center justify-center text-xs hover:bg-red-600">
                                        <i class='bx bx-x'></i>
                                    </button>
                                </div>
                            </div>
                        </div>

                        <div class="flex items-center gap-2">
                            <input type="checkbox" name="is_default" id="edit_is_default" class="h-4 w-4 text-indigo-600 border-slate-300 rounded focus:ring-indigo-500">
                            <label for="edit_is_default" class="text-sm text-slate-700">Jadikan sebagai penanda tangan default</label>
                        </div>
                    </div>
                    <div class="mt-6 flex justify-end gap-3">
                        <button type="button" onclick="closeModal('editModal')" class="px-4 py-2 border border-slate-300 rounded-lg text-slate-700 hover:bg-slate-50 font-medium">Batal</button>
                        <button type="submit" class="px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 font-medium">Update</button>
                    </div>
                </form>
            </div>
        </div>
    </div>
</div>

<script>
    function openModal(modalId) {
        document.getElementById(modalId).classList.remove('hidden');
    }

    function closeModal(modalId) {
        document.getElementById(modalId).classList.add('hidden');
        // Reset select2 pegawai saat modal ditutup
        if (modalId === 'addModal') {
            $('#add-pilih-pegawai').val('').trigger('change');
        }
        if (modalId === 'editModal') {
            $('#edit-pilih-pegawai').val('').trigger('change');
        }
    }

    // Inisialisasi Select2 untuk dropdown pegawai
    $(document).ready(function() {
        $('.select2-pegawai-add').select2({
            placeholder: 'Cari pegawai berdasarkan nama atau NIP...',
            allowClear: true,
            width: '100%',
            dropdownParent: $('#addModal')
        });

        $('.select2-pegawai-edit').select2({
            placeholder: 'Cari pegawai berdasarkan nama atau NIP...',
            allowClear: true,
            width: '100%',
            dropdownParent: $('#editModal')
        });

        // Auto-fill saat pegawai dipilih di modal Tambah
        $('#add-pilih-pegawai').on('select2:select', function(e) {
            const option = e.params.data.element;
            const nama = option.getAttribute('data-nama');
            const nip = option.getAttribute('data-nip');
            const pangkat = option.getAttribute('data-pangkat');
            const jabatan = option.getAttribute('data-jabatan');
            
            // Auto-fill ke form fields
            const form = document.getElementById('formTambah');
            form.querySelector('input[name="nama"]').value = nama || '';
            form.querySelector('input[name="nip"]').value = nip || '';
            form.querySelector('input[name="pangkat"]').value = pangkat || '';
            form.querySelector('input[name="jabatan"]').value = jabatan || '';

            // Highlight efek
            form.querySelectorAll('input[name="nama"], input[name="nip"], input[name="pangkat"], input[name="jabatan"]').forEach(el => {
                el.classList.add('ring-2', 'ring-green-300', 'bg-green-50');
                setTimeout(() => {
                    el.classList.remove('ring-2', 'ring-green-300', 'bg-green-50');
                }, 1500);
            });
        });

        // Auto-fill saat pegawai dipilih di modal Edit
        $('#edit-pilih-pegawai').on('select2:select', function(e) {
            const option = e.params.data.element;
            const nama = option.getAttribute('data-nama');
            const nip = option.getAttribute('data-nip');
            const pangkat = option.getAttribute('data-pangkat');
            const jabatan = option.getAttribute('data-jabatan');
            
            document.getElementById('edit-nama').value = nama || '';
            document.getElementById('edit-nip').value = nip || '';
            document.getElementById('edit-pangkat').value = pangkat || '';
            document.getElementById('edit-jabatan').value = jabatan || '';

            // Highlight efek
            ['edit-nama', 'edit-nip', 'edit-pangkat', 'edit-jabatan'].forEach(id => {
                const el = document.getElementById(id);
                el.classList.add('ring-2', 'ring-green-300', 'bg-green-50');
                setTimeout(() => {
                    el.classList.remove('ring-2', 'ring-green-300', 'bg-green-50');
                }, 1500);
            });
        });

        // Clear auto-fill saat pilihan dihapus
        $('#add-pilih-pegawai').on('select2:clear', function() {
            const form = document.getElementById('formTambah');
            form.querySelector('input[name="nama"]').value = '';
            form.querySelector('input[name="nip"]').value = '';
            form.querySelector('input[name="pangkat"]').value = '';
            form.querySelector('input[name="jabatan"]').value = '';
        });
    });

    // Toggle tipe penanda tangan (Add Modal)
    document.querySelectorAll('input[name="is_kepala"]').forEach(radio => {
        radio.addEventListener('change', function() {
            const form = this.closest('form');
            const isAdd = form.id === 'formTambah';
            const prefix = isAdd ? 'add' : 'edit';
            const wrap = document.getElementById(prefix + '-jabatan-atasan-wrap');
            const options = form.querySelectorAll(isAdd ? '.add-tipe-option' : '.edit-tipe-option');
            
            options.forEach(opt => {
                opt.classList.remove('border-indigo-500', 'bg-indigo-50');
                opt.classList.add('border-slate-200');
            });
            
            const selectedLabel = this.closest('label');
            selectedLabel.classList.remove('border-slate-200');
            selectedLabel.classList.add('border-indigo-500', 'bg-indigo-50');
            
            if (this.value === '0') {
                wrap.classList.remove('hidden');
            } else {
                wrap.classList.add('hidden');
            }
        });
    });

    // Edit Button
    document.querySelectorAll('.edit-btn').forEach(button => {
        button.addEventListener('click', function() {
            document.getElementById('edit-id').value = this.dataset.id;
            document.getElementById('edit-nama').value = this.dataset.nama;
            document.getElementById('edit-nip').value = this.dataset.nip;
            document.getElementById('edit-pangkat').value = this.dataset.pangkat;
            document.getElementById('edit-jabatan').value = this.dataset.jabatan;
            document.getElementById('edit_is_default').checked = this.dataset.default === '1';
            
            // Set tipe penanda tangan
            const isKepala = this.dataset.isKepala;
            if (isKepala === '0') {
                document.getElementById('edit-is-kepala-0').checked = true;
                document.getElementById('edit-jabatan-atasan-wrap').classList.remove('hidden');
                document.getElementById('edit-jabatan-atasan').value = this.dataset.jabatanAtasan || '';
                // Update styling
                document.querySelectorAll('.edit-tipe-option').forEach(opt => {
                    opt.classList.remove('border-indigo-500', 'bg-indigo-50');
                    opt.classList.add('border-slate-200');
                });
                document.querySelector('.edit-tipe-option[data-value="0"]').classList.remove('border-slate-200');
                document.querySelector('.edit-tipe-option[data-value="0"]').classList.add('border-indigo-500', 'bg-indigo-50');
            } else {
                document.getElementById('edit-is-kepala-1').checked = true;
                document.getElementById('edit-jabatan-atasan-wrap').classList.add('hidden');
                document.getElementById('edit-jabatan-atasan').value = '';
                // Update styling
                document.querySelectorAll('.edit-tipe-option').forEach(opt => {
                    opt.classList.remove('border-indigo-500', 'bg-indigo-50');
                    opt.classList.add('border-slate-200');
                });
                document.querySelector('.edit-tipe-option[data-value="1"]').classList.remove('border-slate-200');
                document.querySelector('.edit-tipe-option[data-value="1"]').classList.add('border-indigo-500', 'bg-indigo-50');
            }
            
            // Set tanda tangan
            const tandaTangan = this.dataset.tandaTangan;
            document.getElementById('edit-hapus-ttd').value = '0';
            document.getElementById('edit-tanda-tangan').value = '';
            document.getElementById('edit-ttd-preview').classList.add('hidden');
            
            if (tandaTangan) {
                document.getElementById('edit-ttd-current').classList.remove('hidden');
                document.getElementById('edit-ttd-current-img').src = '../assets/img/tanda_tangan/' + tandaTangan;
            } else {
                document.getElementById('edit-ttd-current').classList.add('hidden');
            }
            
            openModal('editModal');
        });
    });

    // Set Default Button
    document.querySelectorAll('.set-default-btn').forEach(button => {
        button.addEventListener('click', function() {
            const id = this.dataset.id;
            Swal.fire({
                title: 'Set sebagai Default?',
                text: 'Penanda tangan ini akan menjadi default untuk surat tugas baru.',
                icon: 'question',
                showCancelButton: true,
                confirmButtonColor: '#10b981',
                cancelButtonColor: '#64748b',
                confirmButtonText: 'Ya, Set Default',
                cancelButtonText: 'Batal'
            }).then((result) => {
                if (result.isConfirmed) {
                    const formData = new FormData();
                    formData.append('action', 'set_default');
                    formData.append('id', id);

                    fetch('penandatangan.php', { method: 'POST', body: formData })
                        .then(r => r.json())
                        .then(data => {
                            if (data.success) {
                                Swal.fire({ icon: 'success', title: 'Berhasil!', text: data.message, timer: 1500, showConfirmButton: false })
                                    .then(() => location.reload());
                            } else {
                                throw new Error(data.message);
                            }
                        })
                        .catch(error => Swal.fire({ icon: 'error', title: 'Error!', text: error.message }));
                }
            });
        });
    });

    // Toggle Aktif Button
    document.querySelectorAll('.toggle-aktif-btn').forEach(button => {
        button.addEventListener('click', function() {
            const id = this.dataset.id;
            const isAktif = this.dataset.aktif === '1';
            const actionText = isAktif ? 'menonaktifkan' : 'mengaktifkan';

            Swal.fire({
                title: `${isAktif ? 'Nonaktifkan' : 'Aktifkan'} Penanda Tangan?`,
                text: `Apakah Anda yakin ingin ${actionText} penanda tangan ini?`,
                icon: 'question',
                showCancelButton: true,
                confirmButtonColor: '#6366f1',
                cancelButtonColor: '#64748b',
                confirmButtonText: 'Ya',
                cancelButtonText: 'Batal'
            }).then((result) => {
                if (result.isConfirmed) {
                    const formData = new FormData();
                    formData.append('action', 'toggle_aktif');
                    formData.append('id', id);

                    fetch('penandatangan.php', { method: 'POST', body: formData })
                        .then(r => r.json())
                        .then(data => {
                            if (data.success) {
                                Swal.fire({ icon: 'success', title: 'Berhasil!', text: data.message, timer: 1500, showConfirmButton: false })
                                    .then(() => location.reload());
                            } else {
                                throw new Error(data.message);
                            }
                        })
                        .catch(error => Swal.fire({ icon: 'error', title: 'Error!', text: error.message }));
                }
            });
        });
    });

    // Delete Button
    document.querySelectorAll('.delete-btn').forEach(button => {
        button.addEventListener('click', function() {
            const id = this.dataset.id;
            Swal.fire({
                title: 'Hapus Penanda Tangan?',
                text: 'Data penanda tangan akan dihapus permanen!',
                icon: 'warning',
                showCancelButton: true,
                confirmButtonColor: '#ef4444',
                cancelButtonColor: '#64748b',
                confirmButtonText: 'Ya, Hapus!',
                cancelButtonText: 'Batal'
            }).then((result) => {
                if (result.isConfirmed) {
                    const formData = new FormData();
                    formData.append('action', 'delete');
                    formData.append('id', id);

                    fetch('penandatangan.php', { method: 'POST', body: formData })
                        .then(r => r.json())
                        .then(data => {
                            if (data.success) {
                                Swal.fire({ icon: 'success', title: 'Berhasil!', text: data.message, timer: 1500, showConfirmButton: false })
                                    .then(() => location.reload());
                            } else {
                                throw new Error(data.message);
                            }
                        })
                        .catch(error => Swal.fire({ icon: 'error', title: 'Error!', text: error.message }));
                }
            });
        });
    });
    // === Fungsi Tanda Tangan ===
    
    // Preview gambar tanda tangan saat dipilih
    function previewTTD(input, previewId) {
        const preview = document.getElementById(previewId);
        const img = document.getElementById(previewId + '-img');
        
        if (input.files && input.files[0]) {
            const file = input.files[0];
            
            // Validasi tipe file
            if (!file.type.includes('png')) {
                Swal.fire({
                    icon: 'warning',
                    title: 'Format Tidak Valid',
                    text: 'Hanya file PNG yang diperbolehkan untuk tanda tangan.',
                    confirmButtonColor: '#6366f1'
                });
                input.value = '';
                preview.classList.add('hidden');
                return;
            }
            
            // Validasi ukuran file
            if (file.size > 2 * 1024 * 1024) {
                Swal.fire({
                    icon: 'warning',
                    title: 'Ukuran File Terlalu Besar',
                    text: 'Ukuran file maksimal 2MB.',
                    confirmButtonColor: '#6366f1'
                });
                input.value = '';
                preview.classList.add('hidden');
                return;
            }
            
            const reader = new FileReader();
            reader.onload = function(e) {
                img.src = e.target.result;
                preview.classList.remove('hidden');
            };
            reader.readAsDataURL(file);
        }
    }
    
    // Clear preview tanda tangan
    function clearTTDPreview(prefix) {
        const input = document.getElementById(prefix + '-tanda-tangan');
        const preview = document.getElementById(prefix + '-ttd-preview');
        input.value = '';
        preview.classList.add('hidden');
    }
    
    // Hapus tanda tangan dari edit modal
    function removeTTDEdit() {
        Swal.fire({
            title: 'Hapus Tanda Tangan?',
            text: 'Gambar tanda tangan akan dihapus saat data di-update.',
            icon: 'question',
            showCancelButton: true,
            confirmButtonColor: '#ef4444',
            cancelButtonColor: '#64748b',
            confirmButtonText: 'Ya, Hapus',
            cancelButtonText: 'Batal'
        }).then((result) => {
            if (result.isConfirmed) {
                document.getElementById('edit-ttd-current').classList.add('hidden');
                document.getElementById('edit-hapus-ttd').value = '1';
            }
        });
    }
    
    // Preview tanda tangan di lightbox (klik dari tabel)
    function openPreviewTTD(src) {
        Swal.fire({
            imageUrl: src,
            imageAlt: 'Tanda Tangan',
            showConfirmButton: false,
            showCloseButton: true,
            background: '#ffffff',
            width: 'auto',
            customClass: {
                image: 'max-h-64'
            }
        });
    }

</script>

<?php include '../includes/footer.php'; ?>
