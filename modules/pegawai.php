<?php
$pdo = require_once '../config/database.php';

// Proses Delete (harus sebelum include header agar JSON response bersih)
if (isset($_GET['delete'])) {
    try {
        // Cek apakah pegawai terkait dengan surat tugas
        $check_stmt = $pdo->prepare("SELECT COUNT(*) FROM pegawai_tugas WHERE nip = ?");
        $check_stmt->execute([$_GET['delete']]);
        if ($check_stmt->fetchColumn() > 0) {
            header('Content-Type: application/json');
            echo json_encode(['success' => false, 'message' => 'Pegawai tidak dapat dihapus karena terkait dengan surat tugas']);
            exit;
        }

        $stmt = $pdo->prepare("DELETE FROM pegawai WHERE nip = ?");
        $stmt->execute([$_GET['delete']]);

        header('Content-Type: application/json');
        echo json_encode(['success' => true]);
        exit;
    } catch (PDOException $e) {
        header('Content-Type: application/json');
        echo json_encode(['success' => false, 'message' => $e->getMessage()]);
        exit;
    }
}

// Proses CRUD (POST)
if ($_SERVER['REQUEST_METHOD'] == 'POST') {
    if (isset($_POST['action'])) {
        // Tambah Pegawai
        if ($_POST['action'] == 'create') {
            try {
                // Cek apakah NIP sudah ada
                $check_stmt = $pdo->prepare("SELECT COUNT(*) FROM pegawai WHERE nip = ?");
                $check_stmt->execute([$_POST['nip']]);
                if ($check_stmt->fetchColumn() > 0) {
                    $error_create = "NIP sudah terdaftar!";
                } else {
                    // Insert data pegawai baru
                    $stmt = $pdo->prepare("INSERT INTO pegawai (nip, nama, pangkat, jabatan) VALUES (?, ?, ?, ?)");
                    $stmt->execute([$_POST['nip'], $_POST['nama'], $_POST['pangkat'], $_POST['jabatan']]);
                    $success_create = true;
                }
            } catch (PDOException $e) {
                $error_create = $e->getMessage();
            }
        }

        // Update Pegawai
        if ($_POST['action'] == 'update') {
            try {
                $old_nip = $_POST['old_nip'];
                $new_nip = $_POST['nip'];
                
                // Jika NIP berubah, cek apakah NIP baru sudah ada
                if ($new_nip !== $old_nip) {
                    $check_stmt = $pdo->prepare("SELECT COUNT(*) FROM pegawai WHERE nip = ?");
                    $check_stmt->execute([$new_nip]);
                    if ($check_stmt->fetchColumn() > 0) {
                        $error_update = "NIP baru sudah terdaftar!";
                    } else {
                        // Update termasuk NIP, disable FK check karena constraint ON UPDATE RESTRICT
                        $pdo->beginTransaction();
                        $pdo->exec("SET FOREIGN_KEY_CHECKS = 0");
                        
                        // Update data pegawai termasuk NIP
                        $stmt = $pdo->prepare("UPDATE pegawai SET nip = ?, nama = ?, pangkat = ?, jabatan = ? WHERE nip = ?");
                        $stmt->execute([$new_nip, $_POST['nama'], $_POST['pangkat'], $_POST['jabatan'], $old_nip]);
                        
                        // Update NIP di tabel pegawai_tugas
                        $stmt_tugas = $pdo->prepare("UPDATE pegawai_tugas SET nip = ? WHERE nip = ?");
                        $stmt_tugas->execute([$new_nip, $old_nip]);
                        
                        $pdo->exec("SET FOREIGN_KEY_CHECKS = 1");
                        $pdo->commit();
                        $success_update = true;
                    }
                } else {
                    // NIP tidak berubah, update field lainnya saja
                    $stmt = $pdo->prepare("UPDATE pegawai SET nama = ?, pangkat = ?, jabatan = ? WHERE nip = ?");
                    $stmt->execute([$_POST['nama'], $_POST['pangkat'], $_POST['jabatan'], $old_nip]);
                    $success_update = true;
                }
            } catch (PDOException $e) {
                if ($pdo->inTransaction()) {
                    $pdo->rollBack();
                }
                $error_update = $e->getMessage();
            }
        }
    }
}

include '../includes/header.php';

// Tampilkan notifikasi jika ada
if (isset($success_create)) {
    echo "<script>
        Swal.fire({
            title: 'Berhasil!',
            text: 'Pegawai berhasil ditambahkan',
            icon: 'success',
            confirmButtonText: 'OK'
        }).then((result) => {
            if (result.isConfirmed) {
                window.location.href = 'pegawai.php';
            }
        });
    </script>";
}

if (isset($error_create)) {
    echo "<script>
        Swal.fire({
            title: 'Error!',
            text: '" . addslashes($error_create) . "',
            icon: 'error',
            confirmButtonText: 'OK'
        });
    </script>";
}

if (isset($success_update)) {
    echo "<script>
        Swal.fire({
            title: 'Berhasil!',
            text: 'Data pegawai berhasil diperbarui',
            icon: 'success',
            confirmButtonText: 'OK'
        }).then((result) => {
            if (result.isConfirmed) {
                window.location.href = 'pegawai.php" . (!empty($search) ? "?search=" . urlencode($search) : "") . "';
            }
        });
    </script>";
}

if (isset($error_update)) {
    echo "<script>
        Swal.fire({
            title: 'Error!',
            text: '" . addslashes($error_update) . "',
            icon: 'error',
            confirmButtonText: 'OK'
        });
    </script>";
}

// Konfigurasi Pencarian
$search = isset($_GET['search']) ? $_GET['search'] : '';
$search_condition = '';

// Konfigurasi Pagination
$limit = 10; // jumlah data per halaman
$page = isset($_GET['page']) ? (int) $_GET['page'] : 1;
$start = ($page - 1) * $limit;

// Persiapkan query sesuai apakah ada pencarian atau tidak
if (!empty($search)) {
    $search_condition = "WHERE nip LIKE :search OR nama LIKE :search OR pangkat LIKE :search OR jabatan LIKE :search";
}

// Hitung total data dengan kondisi pencarian
$query_count = "SELECT COUNT(*) AS total FROM pegawai $search_condition";
$stmt_count = $pdo->prepare($query_count);
if (!empty($search)) {
    $stmt_count->bindValue(':search', "%$search%", PDO::PARAM_STR);
}
$stmt_count->execute();
$total_records = $stmt_count->fetch()['total'];
$total_pages = ceil($total_records / $limit);

// Ambil data pegawai dengan limit dan pencarian
$query = "SELECT * FROM pegawai $search_condition ORDER BY nama LIMIT :start, :limit";
$stmt = $pdo->prepare($query);
$stmt->bindValue(':start', $start, PDO::PARAM_INT);
$stmt->bindValue(':limit', $limit, PDO::PARAM_INT);
if (!empty($search)) {
    $stmt->bindValue(':search', "%$search%", PDO::PARAM_STR);
}
$stmt->execute();
$pegawai = $stmt->fetchAll();

// Proses CRUD
?>


<div class="space-y-6">
    <div class="flex flex-col md:flex-row justify-between items-center gap-4">
        <h2 class="text-2xl font-bold text-slate-800">Data Pegawai</h2>
        <button onclick="openModal('addPegawaiModal')" class="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg flex items-center gap-2 transition-colors shadow-sm">
            <i class='bx bx-plus'></i> Tambah Pegawai
        </button>
    </div>

    <!-- Form Pencarian -->
    <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-4">
        <form action="" method="GET" class="flex gap-2">
            <div class="relative flex-1">
                <i class='bx bx-search absolute left-3 top-1/2 transform -translate-y-1/2 text-slate-400'></i>
                <input type="text" name="search" class="w-full pl-10 pr-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all"
                    placeholder="Cari NIP, nama, pangkat, atau jabatan"
                    value="<?php echo htmlspecialchars($search); ?>">
            </div>
            <button type="submit" class="border-2 border-indigo-600 text-indigo-600 hover:bg-indigo-50 px-6 py-2 rounded-lg transition-colors font-medium">
                Cari
            </button>
            <?php if (!empty($search)): ?>
                <a href="pegawai.php" class="bg-slate-200 hover:bg-slate-300 text-slate-700 px-4 py-2 rounded-lg transition-colors flex items-center gap-2">
                    <i class='bx bx-reset'></i> Reset
                </a>
            <?php endif; ?>
        </form>
        <!-- Filter Row -->
        <div class="flex flex-wrap items-center gap-3 mt-3 pt-3 border-t border-slate-100">
            <div class="flex items-center gap-2">
                <label class="text-xs font-medium text-slate-500">Pangkat:</label>
                <select id="filterPangkat" class="text-sm border border-slate-300 rounded-lg px-3 py-1.5 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                    <option value="">Semua Pangkat</option>
                    <?php
                    $pangkat_list = [];
                    foreach ($pegawai as $p) {
                        $val = trim($p['pangkat']);
                        if ($val !== '' && !in_array($val, $pangkat_list)) $pangkat_list[] = $val;
                    }
                    sort($pangkat_list);
                    foreach ($pangkat_list as $pk): ?>
                        <option value="<?php echo htmlspecialchars($pk); ?>"><?php echo htmlspecialchars($pk); ?></option>
                    <?php endforeach; ?>
                </select>
            </div>
            <div class="flex items-center gap-2">
                <label class="text-xs font-medium text-slate-500">Jabatan:</label>
                <select id="filterJabatan" class="text-sm border border-slate-300 rounded-lg px-3 py-1.5 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                    <option value="">Semua Jabatan</option>
                    <?php
                    $jabatan_list = [];
                    foreach ($pegawai as $p) {
                        $val = trim($p['jabatan']);
                        if ($val !== '' && !in_array($val, $jabatan_list)) $jabatan_list[] = $val;
                    }
                    sort($jabatan_list);
                    foreach ($jabatan_list as $jb): ?>
                        <option value="<?php echo htmlspecialchars($jb); ?>"><?php echo htmlspecialchars($jb); ?></option>
                    <?php endforeach; ?>
                </select>
            </div>
            <button type="button" id="btnResetFilterPegawai" class="hidden text-xs text-red-600 hover:text-red-700 font-medium px-3 py-1.5 border border-red-200 rounded-lg hover:bg-red-50 transition-colors">
                <i class='bx bx-reset mr-1'></i>Reset Filter
            </button>
        </div>
    </div>

    <!-- Tabel Pegawai -->
    <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
        <div class="overflow-x-auto">
            <?php if (count($pegawai) > 0): ?>
                <table class="w-full text-left border-collapse" id="tabelPegawai">
                    <thead>
                        <tr class="bg-slate-50 text-slate-600 text-xs uppercase tracking-wider">
                            <th class="px-4 py-3 font-semibold border-b border-slate-100" style="width:140px;">NIP</th>
                            <th class="px-4 py-3 font-semibold border-b border-slate-100 cursor-pointer select-none sortable-col" data-sort="nama" data-order="asc">
                                <span class="inline-flex items-center gap-1">Nama <i class='bx bx-sort-up text-indigo-600 sort-icon'></i></span>
                            </th>
                            <th class="px-4 py-3 font-semibold border-b border-slate-100 cursor-pointer select-none sortable-col" data-sort="pangkat" data-order="">
                                <span class="inline-flex items-center gap-1">Pangkat <i class='bx bx-sort-alt-2 text-slate-400 sort-icon'></i></span>
                            </th>
                            <th class="px-4 py-3 font-semibold border-b border-slate-100 cursor-pointer select-none sortable-col" data-sort="jabatan" data-order="" style="min-width:200px;">
                                <span class="inline-flex items-center gap-1">Jabatan <i class='bx bx-sort-alt-2 text-slate-400 sort-icon'></i></span>
                            </th>
                            <th class="px-4 py-3 font-semibold border-b border-slate-100 text-right">Aksi</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-100">
                        <?php foreach ($pegawai as $p): 
                            $nip_full = htmlspecialchars($p['nip']);
                            $nip_short = (strlen($p['nip']) > 12) 
                                ? substr($p['nip'], 0, 8) . '...' . substr($p['nip'], -4) 
                                : $p['nip'];
                        ?>
                            <tr class="hover:bg-slate-50 transition-colors pegawai-row"
                                data-nama="<?php echo htmlspecialchars($p['nama']); ?>"
                                data-pangkat="<?php echo htmlspecialchars($p['pangkat']); ?>"
                                data-jabatan="<?php echo htmlspecialchars($p['jabatan']); ?>">
                                <td class="px-4 py-2.5 text-sm text-slate-600">
                                    <div class="group/nip inline-flex items-center gap-1">
                                        <span class="font-mono text-xs" title="<?php echo $nip_full; ?>"><?php echo htmlspecialchars($nip_short); ?></span>
                                        <button type="button" class="copy-nip opacity-0 group-hover/nip:opacity-100 text-slate-400 hover:text-indigo-600 transition-opacity" data-nip="<?php echo $nip_full; ?>" title="Salin NIP">
                                            <i class='bx bx-copy text-sm'></i>
                                        </button>
                                    </div>
                                </td>
                                <td class="px-4 py-2.5 text-sm font-medium text-slate-900"><?php echo htmlspecialchars($p['nama']); ?></td>
                                <td class="px-4 py-2.5 text-sm text-slate-600"><?php echo htmlspecialchars($p['pangkat']); ?></td>
                                <td class="px-4 py-2.5 text-sm text-slate-600"><?php echo htmlspecialchars($p['jabatan']); ?></td>
                                <td class="px-4 py-2.5 text-sm text-right">
                                    <div class="inline-flex items-center gap-2">
                                        <button class="text-blue-500 hover:text-blue-600 view-btn transition-colors" title="Lihat Detail"
                                            data-nip="<?php echo htmlspecialchars($p['nip']); ?>"
                                            data-nama="<?php echo htmlspecialchars($p['nama']); ?>"
                                            data-pangkat="<?php echo htmlspecialchars($p['pangkat']); ?>"
                                            data-jabatan="<?php echo htmlspecialchars($p['jabatan']); ?>">
                                            <i class='bx bx-show text-lg'></i>
                                        </button>
                                        <button class="text-amber-500 hover:text-amber-600 edit-btn transition-colors" title="Edit"
                                            data-nip="<?php echo htmlspecialchars($p['nip']); ?>"
                                            data-nama="<?php echo htmlspecialchars($p['nama']); ?>"
                                            data-pangkat="<?php echo htmlspecialchars($p['pangkat']); ?>"
                                            data-jabatan="<?php echo htmlspecialchars($p['jabatan']); ?>">
                                            <i class='bx bx-edit text-lg'></i>
                                        </button>
                                        <button class="text-red-500 hover:text-red-600 delete-btn transition-colors" title="Hapus"
                                            data-nip="<?php echo htmlspecialchars($p['nip']); ?>"
                                            data-nama="<?php echo htmlspecialchars($p['nama']); ?>">
                                            <i class='bx bx-trash text-lg'></i>
                                        </button>
                                    </div>
                                </td>
                            </tr>
                        <?php endforeach; ?>
                    </tbody>
                </table>
            <?php else: ?>
                <div class="p-8 text-center text-slate-500">
                    <i class='bx bx-search text-4xl mb-2 text-slate-300'></i>
                    <p>
                        <?php if (!empty($search)): ?>
                            Tidak ditemukan data pegawai dengan kata kunci "<?php echo htmlspecialchars($search); ?>".
                        <?php else: ?>
                            Belum ada data pegawai.
                        <?php endif; ?>
                    </p>
                </div>
            <?php endif; ?>
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
                        <a href="?page=<?php echo ($page - 1); ?><?php echo (!empty($search) ? '&search=' . urlencode($search) : ''); ?>" 
                           class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm">Previous</a>
                    <?php endif; ?>

                    <?php if ($from > 1): ?>
                        <a href="?page=1<?php echo (!empty($search) ? '&search=' . urlencode($search) : ''); ?>"
                           class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm">1</a>
                        <?php if ($from > 2): ?>
                            <span class="px-2 py-1 text-slate-400 text-sm">...</span>
                        <?php endif; ?>
                    <?php endif; ?>

                    <?php for ($i = $from; $i <= $to; $i++): ?>
                        <a href="?page=<?php echo $i; ?><?php echo (!empty($search) ? '&search=' . urlencode($search) : ''); ?>"
                           class="px-3 py-1 rounded border <?php echo ($page == $i) ? 'bg-indigo-600 text-white border-indigo-600' : 'border-slate-200 text-slate-600 hover:bg-slate-50'; ?> text-sm">
                            <?php echo $i; ?>
                        </a>
                    <?php endfor; ?>

                    <?php if ($to < $total_pages): ?>
                        <?php if ($to < $total_pages - 1): ?>
                            <span class="px-2 py-1 text-slate-400 text-sm">...</span>
                        <?php endif; ?>
                        <a href="?page=<?php echo $total_pages; ?><?php echo (!empty($search) ? '&search=' . urlencode($search) : ''); ?>"
                           class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm"><?php echo $total_pages; ?></a>
                    <?php endif; ?>

                    <?php if ($page < $total_pages): ?>
                        <a href="?page=<?php echo ($page + 1); ?><?php echo (!empty($search) ? '&search=' . urlencode($search) : ''); ?>"
                           class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm">Next</a>
                    <?php endif; ?>
                </nav>
            </div>
        <?php endif; ?>
    </div>
</div>

<!-- Modal Detail Pegawai (Read-Only) -->
<div id="viewPegawaiModal" class="fixed inset-0 z-50 hidden overflow-y-auto" aria-labelledby="modal-title" role="dialog" aria-modal="true">
    <div class="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
        <div class="fixed inset-0 bg-slate-900 bg-opacity-75 transition-opacity" aria-hidden="true" onclick="closeModal('viewPegawaiModal')"></div>
        <span class="hidden sm:inline-block sm:align-middle sm:h-screen" aria-hidden="true">&#8203;</span>
        <div class="inline-block align-bottom bg-white rounded-xl text-left overflow-hidden shadow-xl transform transition-all sm:my-8 sm:align-middle sm:max-w-lg sm:w-full">
            <div class="bg-white px-4 pt-5 pb-4 sm:p-6 sm:pb-4">
                <div class="flex justify-between items-center mb-4">
                    <h3 class="text-lg leading-6 font-medium text-slate-900">Detail Pegawai</h3>
                    <button onclick="closeModal('viewPegawaiModal')" class="text-slate-400 hover:text-slate-500">
                        <i class='bx bx-x text-2xl'></i>
                    </button>
                </div>
                <div class="space-y-4">
                    <div>
                        <label class="block text-xs font-medium text-slate-500 uppercase tracking-wider mb-1">NIP</label>
                        <p id="view-nip" class="text-sm font-mono text-slate-800 bg-slate-50 px-3 py-2 rounded-lg"></p>
                    </div>
                    <div>
                        <label class="block text-xs font-medium text-slate-500 uppercase tracking-wider mb-1">Nama</label>
                        <p id="view-nama" class="text-sm font-medium text-slate-900 bg-slate-50 px-3 py-2 rounded-lg"></p>
                    </div>
                    <div>
                        <label class="block text-xs font-medium text-slate-500 uppercase tracking-wider mb-1">Pangkat</label>
                        <p id="view-pangkat" class="text-sm text-slate-800 bg-slate-50 px-3 py-2 rounded-lg"></p>
                    </div>
                    <div>
                        <label class="block text-xs font-medium text-slate-500 uppercase tracking-wider mb-1">Jabatan</label>
                        <p id="view-jabatan" class="text-sm text-slate-800 bg-slate-50 px-3 py-2 rounded-lg"></p>
                    </div>
                </div>
                <div class="mt-6 flex justify-end">
                    <button type="button" onclick="closeModal('viewPegawaiModal')" class="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg font-medium transition-colors">Tutup</button>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- Modal Tambah Pegawai -->
<div id="addPegawaiModal" class="fixed inset-0 z-50 hidden overflow-y-auto" aria-labelledby="modal-title" role="dialog" aria-modal="true">
    <div class="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
        <div class="fixed inset-0 bg-slate-900 bg-opacity-75 transition-opacity" aria-hidden="true" onclick="closeModal('addPegawaiModal')"></div>
        <span class="hidden sm:inline-block sm:align-middle sm:h-screen" aria-hidden="true">&#8203;</span>
        <div class="inline-block align-bottom bg-white rounded-xl text-left overflow-hidden shadow-xl transform transition-all sm:my-8 sm:align-middle sm:max-w-lg sm:w-full">
            <div class="bg-white px-4 pt-5 pb-4 sm:p-6 sm:pb-4">
                <div class="flex justify-between items-center mb-4">
                    <h3 class="text-lg leading-6 font-medium text-slate-900" id="modal-title">Tambah Pegawai</h3>
                    <button onclick="closeModal('addPegawaiModal')" class="text-slate-400 hover:text-slate-500">
                        <i class='bx bx-x text-2xl'></i>
                    </button>
                </div>
                <form method="POST">
                    <input type="hidden" name="action" value="create">
                    <div class="space-y-4">
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">NIP</label>
                            <input type="text" name="nip" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Nama</label>
                            <input type="text" name="nama" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Pangkat</label>
                            <input type="text" name="pangkat" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Jabatan</label>
                            <input type="text" name="jabatan" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                    </div>
                    <div class="mt-6 flex justify-end gap-3">
                        <button type="button" onclick="closeModal('addPegawaiModal')" class="px-4 py-2 border border-slate-300 rounded-lg text-slate-700 hover:bg-slate-50 font-medium">Batal</button>
                        <button type="submit" class="px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 font-medium">Simpan</button>
                    </div>
                </form>
            </div>
        </div>
    </div>
</div>

<!-- Modal Edit Pegawai -->
<div id="editPegawaiModal" class="fixed inset-0 z-50 hidden overflow-y-auto" aria-labelledby="modal-title" role="dialog" aria-modal="true">
    <div class="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
        <div class="fixed inset-0 bg-slate-900 bg-opacity-75 transition-opacity" aria-hidden="true" onclick="closeModal('editPegawaiModal')"></div>
        <span class="hidden sm:inline-block sm:align-middle sm:h-screen" aria-hidden="true">&#8203;</span>
        <div class="inline-block align-bottom bg-white rounded-xl text-left overflow-hidden shadow-xl transform transition-all sm:my-8 sm:align-middle sm:max-w-lg sm:w-full">
            <div class="bg-white px-4 pt-5 pb-4 sm:p-6 sm:pb-4">
                <div class="flex justify-between items-center mb-4">
                    <h3 class="text-lg leading-6 font-medium text-slate-900">Edit Pegawai</h3>
                    <button onclick="closeModal('editPegawaiModal')" class="text-slate-400 hover:text-slate-500">
                        <i class='bx bx-x text-2xl'></i>
                    </button>
                </div>
                <form method="POST">
                    <input type="hidden" name="action" value="update">
                    <input type="hidden" name="old_nip" id="edit-old-nip">
                    <div class="space-y-4">
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">NIP</label>
                            <input type="text" name="nip" id="edit-nip" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Nama</label>
                            <input type="text" name="nama" id="edit-nama" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Pangkat</label>
                            <input type="text" name="pangkat" id="edit-pangkat" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Jabatan</label>
                            <input type="text" name="jabatan" id="edit-jabatan" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                    </div>
                    <div class="mt-6 flex justify-end gap-3">
                        <button type="button" onclick="closeModal('editPegawaiModal')" class="px-4 py-2 border border-slate-300 rounded-lg text-slate-700 hover:bg-slate-50 font-medium">Batal</button>
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
    }

    // Event listener untuk tombol lihat detail
    document.querySelectorAll('.view-btn').forEach(button => {
        button.addEventListener('click', function() {
            document.getElementById('view-nip').textContent = this.dataset.nip;
            document.getElementById('view-nama').textContent = this.dataset.nama;
            document.getElementById('view-pangkat').textContent = this.dataset.pangkat;
            document.getElementById('view-jabatan').textContent = this.dataset.jabatan;
            openModal('viewPegawaiModal');
        });
    });

    // Event listener untuk tombol edit
    document.querySelectorAll('.edit-btn').forEach(button => {
        button.addEventListener('click', function() {
            const nip = this.dataset.nip;
            const nama = this.dataset.nama;
            const pangkat = this.dataset.pangkat;
            const jabatan = this.dataset.jabatan;
            
            document.getElementById('edit-old-nip').value = nip;
            document.getElementById('edit-nip').value = nip;
            document.getElementById('edit-nama').value = nama;
            document.getElementById('edit-pangkat').value = pangkat;
            document.getElementById('edit-jabatan').value = jabatan;
            openModal('editPegawaiModal');
        });
    });

    // Script untuk delete dengan konfirmasi nama pegawai
    document.querySelectorAll('.delete-btn').forEach(button => {
        button.addEventListener('click', function () {
            const nip = this.dataset.nip;
            const nama = this.dataset.nama;
            Swal.fire({
                title: 'Hapus ' + nama + '?',
                text: "Tindakan ini tidak dapat dibatalkan.",
                icon: 'warning',
                showCancelButton: true,
                confirmButtonColor: '#ef4444',
                cancelButtonColor: '#64748b',
                confirmButtonText: 'Hapus',
                cancelButtonText: 'Batal',
                reverseButtons: true
            }).then((result) => {
                if (result.isConfirmed) {
                    fetch(`pegawai.php?delete=${encodeURIComponent(nip)}`)
                        .then(response => response.json())
                        .then(data => {
                            if (data.success) {
                                Swal.fire(
                                    'Terhapus!',
                                    'Data pegawai berhasil dihapus.',
                                    'success'
                                ).then(() => {
                                    window.location.reload();
                                });
                            } else {
                                Swal.fire(
                                    'Error!',
                                    data.message || 'Terjadi kesalahan saat menghapus data.',
                                    'error'
                                );
                            }
                        })
                        .catch(error => {
                            Swal.fire(
                                'Error!',
                                'Terjadi kesalahan saat menghapus data.',
                                'error'
                            );
                        });
                }
            });
        });
    });

    // ===== COPY NIP =====
    document.querySelectorAll('.copy-nip').forEach(btn => {
        btn.addEventListener('click', function(e) {
            e.stopPropagation();
            const nip = this.dataset.nip;
            navigator.clipboard.writeText(nip).then(() => {
                const icon = this.querySelector('i');
                icon.className = 'bx bx-check text-sm text-green-600';
                setTimeout(() => { icon.className = 'bx bx-copy text-sm'; }, 1500);
            });
        });
    });

    // ===== CLIENT-SIDE FILTER =====
    (function() {
        const filterPangkat = document.getElementById('filterPangkat');
        const filterJabatan = document.getElementById('filterJabatan');
        const btnReset = document.getElementById('btnResetFilterPegawai');

        function applyFilters() {
            const pangkat = filterPangkat ? filterPangkat.value : '';
            const jabatan = filterJabatan ? filterJabatan.value : '';
            const hasFilter = pangkat !== '' || jabatan !== '';
            const rows = document.querySelectorAll('.pegawai-row');

            rows.forEach(row => {
                let show = true;
                if (pangkat && row.dataset.pangkat !== pangkat) show = false;
                if (jabatan && row.dataset.jabatan !== jabatan) show = false;
                row.style.display = show ? '' : 'none';
            });

            if (btnReset) btnReset.classList.toggle('hidden', !hasFilter);
        }

        if (filterPangkat) filterPangkat.addEventListener('change', applyFilters);
        if (filterJabatan) filterJabatan.addEventListener('change', applyFilters);
        if (btnReset) {
            btnReset.addEventListener('click', function() {
                if (filterPangkat) filterPangkat.value = '';
                if (filterJabatan) filterJabatan.value = '';
                applyFilters();
            });
        }
    })();

    // ===== SORT BY COLUMN =====
    (function() {
        const headers = document.querySelectorAll('.sortable-col');
        headers.forEach(header => {
            header.addEventListener('click', function() {
                const sortField = this.dataset.sort;
                let order = this.dataset.order;

                // Toggle
                if (order === 'asc') order = 'desc';
                else order = 'asc';
                this.dataset.order = order;

                // Reset others
                headers.forEach(h => {
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

                // Sort
                const tbody = document.querySelector('#tabelPegawai tbody');
                if (!tbody) return;
                const rows = Array.from(tbody.querySelectorAll('.pegawai-row'));

                rows.sort((a, b) => {
                    const valA = (a.dataset[sortField] || '').toLowerCase();
                    const valB = (b.dataset[sortField] || '').toLowerCase();
                    if (order === 'asc') return valA.localeCompare(valB);
                    return valB.localeCompare(valA);
                });

                rows.forEach(row => tbody.appendChild(row));
            });
        });
    })();
</script>

<?php include '../includes/footer.php'; ?>