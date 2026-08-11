<?php
$pdo = require_once '../config/database.php';

// Pastikan direktori templates ada
$templates_dir = __DIR__ . '/../assets/templates/';
if (!is_dir($templates_dir)) {
    mkdir($templates_dir, 0755, true);
}

// Pastikan kolom tipe_surat ada di tabel templates
try {
    $pdo->query("SELECT tipe_surat FROM templates LIMIT 1");
} catch (PDOException $e) {
    try {
        $pdo->exec("ALTER TABLE templates ADD COLUMN tipe_surat VARCHAR(20) NOT NULL DEFAULT 'umum' AFTER kategori_template");
    } catch (PDOException $e2) {}
}

// Handle JSON responses (harus sebelum include header)
if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_POST['action'])) {
    // Set Default Template
    if ($_POST['action'] == 'set_default') {
        header('Content-Type: application/json');
        try {
            $pdo->beginTransaction();
            
            // Ambil kategori template yang akan di-set sebagai default
            $stmt = $pdo->prepare("SELECT kategori_template, tipe_surat FROM templates WHERE id = ?");
            $stmt->execute([$_POST['id']]);
            $template = $stmt->fetch();
            
            if (!$template) {
                throw new Exception('Template tidak ditemukan');
            }
            
            // Set semua template dalam kategori dan tipe yang sama menjadi non-default
            $stmt_update = $pdo->prepare("UPDATE templates SET is_default = 0 WHERE kategori_template = ? AND tipe_surat = ?");
            $stmt_update->execute([$template['kategori_template'], $template['tipe_surat'] ?? 'umum']);
            
            // Set template yang dipilih menjadi default
            $stmt = $pdo->prepare("UPDATE templates SET is_default = 1 WHERE id = ?");
            $stmt->execute([$_POST['id']]);
            
            $pdo->commit();
            
            echo json_encode(['success' => true, 'message' => 'Template default berhasil diubah']);
            exit;
        } catch (Exception $e) {
            if ($pdo->inTransaction()) {
                $pdo->rollBack();
            }
            echo json_encode(['success' => false, 'message' => $e->getMessage()]);
            exit;
        }
    }

    // Delete Template
    if ($_POST['action'] == 'delete') {
        header('Content-Type: application/json');
        try {
            $stmt = $pdo->prepare("SELECT nama_file, path_file FROM templates WHERE id = ?");
            $stmt->execute([$_POST['id']]);
            $template = $stmt->fetch();

            if (!$template) {
                echo json_encode(['success' => false, 'message' => 'Template tidak ditemukan']);
                exit;
            }

            // Hapus file dari server
            $file_path = __DIR__ . '/../assets/templates/' . $template['nama_file'];
            if (file_exists($file_path)) {
                if (!unlink($file_path)) {
                    throw new Exception('Gagal menghapus file dari server');
                }
            }

            // Hapus dari database
            $stmt = $pdo->prepare("DELETE FROM templates WHERE id = ?");
            $stmt->execute([$_POST['id']]);

            echo json_encode(['success' => true, 'message' => 'Template berhasil dihapus']);
            exit;
        } catch (Exception $e) {
            echo json_encode(['success' => false, 'message' => $e->getMessage()]);
            exit;
        }
    }

    // Get Template Data (for edit modal)
    if ($_POST['action'] == 'get_template') {
        header('Content-Type: application/json');
        try {
            $stmt = $pdo->prepare("SELECT * FROM templates WHERE id = ?");
            $stmt->execute([$_POST['id']]);
            $template = $stmt->fetch();
            
            if (!$template) {
                echo json_encode(['success' => false, 'message' => 'Template tidak ditemukan']);
            } else {
                echo json_encode(['success' => true, 'data' => $template]);
            }
            exit;
        } catch (Exception $e) {
            echo json_encode(['success' => false, 'message' => $e->getMessage()]);
            exit;
        }
    }
}

// Include header untuk HTML responses
include '../includes/header.php';

// Proses Edit/Update Template (setelah include header karena perlu HTML)
if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_POST['action']) && $_POST['action'] == 'update') {
    try {
        $id = $_POST['id'];
        $nama_template = isset($_POST['nama_template']) ? trim($_POST['nama_template']) : '';
        $kategori_template = isset($_POST['kategori_template']) ? $_POST['kategori_template'] : 'banyak_pegawai';
        $deskripsi = isset($_POST['deskripsi']) ? trim($_POST['deskripsi']) : '';
        $is_default = isset($_POST['is_default']) && $_POST['is_default'] == '1' ? 1 : 0;

        if (empty($nama_template)) {
            throw new Exception('Nama template harus diisi.');
        }

        $allowed_kategori = ['1_pegawai', '2_pegawai', '3_pegawai', 'banyak_pegawai'];
        if (!in_array($kategori_template, $allowed_kategori)) {
            throw new Exception('Kategori template tidak valid.');
        }

        $tipe_surat = isset($_POST['tipe_surat']) ? $_POST['tipe_surat'] : 'umum';
        $allowed_tipe = ['umum', 'penyuluh'];
        if (!in_array($tipe_surat, $allowed_tipe)) {
            throw new Exception('Tipe surat tidak valid.');
        }

        // Ambil data template lama
        $stmt = $pdo->prepare("SELECT * FROM templates WHERE id = ?");
        $stmt->execute([$id]);
        $old_template = $stmt->fetch();
        if (!$old_template) {
            throw new Exception('Template tidak ditemukan.');
        }

        // Jika ada file baru diupload
        $nama_file = $old_template['nama_file'];
        $ukuran_file = $old_template['ukuran_file'];

        if (isset($_FILES['template_file']) && $_FILES['template_file']['error'] === UPLOAD_ERR_OK) {
            $file = $_FILES['template_file'];
            $allowed_extensions = ['docx', 'doc'];
            $file_extension = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));

            if (!in_array($file_extension, $allowed_extensions)) {
                throw new Exception('Format file tidak didukung. Hanya file .docx atau .doc yang diperbolehkan.');
            }

            $max_size = 10 * 1024 * 1024;
            if ($file['size'] > $max_size) {
                throw new Exception('Ukuran file terlalu besar. Maksimal 10MB.');
            }

            // Generate nama file baru
            $nama_file_baru = time() . '_' . preg_replace('/[^a-zA-Z0-9_-]/', '_', $nama_template) . '.' . $file_extension;
            $path_file_baru = $templates_dir . $nama_file_baru;

            if (!move_uploaded_file($file['tmp_name'], $path_file_baru)) {
                throw new Exception('Gagal menyimpan file template baru.');
            }

            // Hapus file lama
            $old_file_path = $templates_dir . $old_template['nama_file'];
            if (file_exists($old_file_path) && $old_template['nama_file'] !== $nama_file_baru) {
                unlink($old_file_path);
            }

            $nama_file = $nama_file_baru;
            $ukuran_file = $file['size'];
        }

        // Jika set sebagai default, reset yang lain di kategori dan tipe yang sama
        if ($is_default) {
            $stmt_update = $pdo->prepare("UPDATE templates SET is_default = 0 WHERE kategori_template = ? AND tipe_surat = ?");
            $stmt_update->execute([$kategori_template, $tipe_surat]);
        }

        // Update database
        $stmt = $pdo->prepare("UPDATE templates SET nama = ?, kategori_template = ?, tipe_surat = ?, nama_file = ?, path_file = ?, ukuran_file = ?, is_default = ?, deskripsi = ? WHERE id = ?");
        $stmt->execute([
            $nama_template,
            $kategori_template,
            $tipe_surat,
            $nama_file,
            '../assets/templates/' . $nama_file,
            $ukuran_file,
            $is_default,
            $deskripsi,
            $id
        ]);

        echo "<script>
                Swal.fire({
                    title: 'Berhasil!',
                    text: 'Template berhasil diupdate',
                    icon: 'success',
                    confirmButtonText: 'OK'
                }).then((result) => {
                    if (result.isConfirmed) {
                        window.location.href = 'template.php';
                    }
                });
            </script>";
        exit;

    } catch (Exception $e) {
        echo "<script>
                Swal.fire({
                    title: 'Gagal!',
                    text: '" . addslashes($e->getMessage()) . "',
                    icon: 'error',
                    confirmButtonText: 'OK'
                });
            </script>";
    }
}

// Proses Upload Template (setelah include header karena perlu HTML)
if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_POST['action']) && $_POST['action'] == 'upload') {
    try {
        if (!isset($_FILES['template_file']) || $_FILES['template_file']['error'] !== UPLOAD_ERR_OK) {
            throw new Exception('File tidak berhasil diupload. Pastikan file yang diupload valid.');
        }

        $file = $_FILES['template_file'];
        $allowed_extensions = ['docx', 'doc'];
        $file_extension = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));

        if (!in_array($file_extension, $allowed_extensions)) {
            throw new Exception('Format file tidak didukung. Hanya file .docx atau .doc yang diperbolehkan.');
        }

        // Validasi ukuran file (max 10MB)
        $max_size = 10 * 1024 * 1024; // 10MB
        if ($file['size'] > $max_size) {
            throw new Exception('Ukuran file terlalu besar. Maksimal 10MB.');
        }

        // Generate nama file unik
        $nama_template = isset($_POST['nama_template']) ? trim($_POST['nama_template']) : '';
        if (empty($nama_template)) {
            $nama_template = pathinfo($file['name'], PATHINFO_FILENAME);
        }

        $nama_file = time() . '_' . preg_replace('/[^a-zA-Z0-9_-]/', '_', $nama_template) . '.' . $file_extension;
        $path_file = $templates_dir . $nama_file;

        // Upload file
        if (!move_uploaded_file($file['tmp_name'], $path_file)) {
            throw new Exception('Gagal menyimpan file template.');
        }

        // Validasi kategori template
        $kategori_template = isset($_POST['kategori_template']) ? $_POST['kategori_template'] : 'banyak_pegawai';
        $allowed_kategori = ['1_pegawai', '2_pegawai', '3_pegawai', 'banyak_pegawai'];
        if (!in_array($kategori_template, $allowed_kategori)) {
            throw new Exception('Kategori template tidak valid.');
        }

        // Validasi tipe surat
        $tipe_surat = isset($_POST['tipe_surat']) ? $_POST['tipe_surat'] : 'umum';
        $allowed_tipe = ['umum', 'penyuluh'];
        if (!in_array($tipe_surat, $allowed_tipe)) {
            throw new Exception('Tipe surat tidak valid.');
        }

        // Simpan ke database
        $is_default = isset($_POST['is_default']) && $_POST['is_default'] == '1' ? 1 : 0;
        $deskripsi = isset($_POST['deskripsi']) ? trim($_POST['deskripsi']) : '';

        // Jika ini template default, set yang lain dalam kategori dan tipe yang sama menjadi non-default
        if ($is_default) {
            $stmt_update = $pdo->prepare("UPDATE templates SET is_default = 0 WHERE kategori_template = ? AND tipe_surat = ?");
            $stmt_update->execute([$kategori_template, $tipe_surat]);
        }

        // Cek apakah kolom kategori_template sudah ada, jika belum tambahkan
        try {
            $pdo->query("SELECT kategori_template FROM templates LIMIT 1");
        } catch (PDOException $e) {
            // Kolom belum ada, tambahkan
            try {
                $pdo->exec("ALTER TABLE templates ADD COLUMN kategori_template ENUM('1_pegawai', '2_pegawai', '3_pegawai', 'banyak_pegawai') NOT NULL DEFAULT 'banyak_pegawai' AFTER nama");
            } catch (PDOException $e2) {
                // Ignore jika sudah ada
            }
        }

        $stmt = $pdo->prepare("INSERT INTO templates (nama, kategori_template, tipe_surat, nama_file, path_file, ukuran_file, is_default, deskripsi) 
                               VALUES (?, ?, ?, ?, ?, ?, ?, ?)");
        $stmt->execute([
            $nama_template,
            $kategori_template,
            $tipe_surat,
            $nama_file,
            '../assets/templates/' . $nama_file,
            $file['size'],
            $is_default,
            $deskripsi
        ]);

        echo "<script>
                Swal.fire({
                    title: 'Berhasil!',
                    text: 'Template berhasil diupload',
                    icon: 'success',
                    confirmButtonText: 'OK'
                }).then((result) => {
                    if (result.isConfirmed) {
                        window.location.href = 'template.php';
                    }
                });
            </script>";
        exit;

    } catch (Exception $e) {
        echo "<script>
                Swal.fire({
                    title: 'Gagal!',
                    text: '" . addslashes($e->getMessage()) . "',
                    icon: 'error',
                    confirmButtonText: 'OK'
                });
            </script>";
    }
}

// Ambil semua template
try {
    // Cek apakah kolom kategori_template ada, jika belum tambahkan
    try {
        $pdo->query("SELECT kategori_template FROM templates LIMIT 1");
    } catch (PDOException $e) {
        try {
            $pdo->exec("ALTER TABLE templates ADD COLUMN kategori_template ENUM('1_pegawai', '2_pegawai', '3_pegawai', 'banyak_pegawai') NOT NULL DEFAULT 'banyak_pegawai' AFTER nama");
        } catch (PDOException $e2) {
            // Ignore jika sudah ada
        }
    }
    
    $stmt = $pdo->query("SELECT * FROM templates ORDER BY kategori_template, tipe_surat, is_default DESC, created_at DESC");
    $templates = $stmt->fetchAll();
} catch (PDOException $e) {
    // Jika tabel belum ada, buat tabel
    if (strpos($e->getMessage(), "doesn't exist") !== false || strpos($e->getMessage(), "Unknown table") !== false) {
        $sql = "CREATE TABLE IF NOT EXISTS templates (
            id INT AUTO_INCREMENT PRIMARY KEY,
            nama VARCHAR(255) NOT NULL,
            kategori_template ENUM('1_pegawai', '2_pegawai', '3_pegawai', 'banyak_pegawai') NOT NULL DEFAULT 'banyak_pegawai',
            tipe_surat VARCHAR(20) NOT NULL DEFAULT 'umum',
            nama_file VARCHAR(255) NOT NULL,
            path_file VARCHAR(500) NOT NULL,
            ukuran_file BIGINT NOT NULL,
            is_default TINYINT(1) DEFAULT 0,
            deskripsi TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            INDEX idx_kategori_tipe (kategori_template, tipe_surat, is_default)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci";
        $pdo->exec($sql);
        $templates = [];
    } else {
        $templates = [];
    }
}
?>

<div class="space-y-6">
    <div class="flex flex-col md:flex-row justify-between items-center gap-4">
        <h2 class="text-2xl font-bold text-slate-800">Kelola Template Word</h2>
        <button onclick="openModal('uploadTemplateModal')" class="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg flex items-center gap-2 transition-colors shadow-sm">
            <i class='bx bx-upload'></i> Upload Template
        </button>
    </div>

    <!-- Info Box -->
    <div class="bg-blue-50 border border-blue-200 rounded-xl p-4">
        <div class="flex items-start gap-3">
            <i class='bx bx-info-circle text-blue-600 text-xl mt-0.5'></i>
            <div class="text-sm text-blue-800">
                <p class="font-semibold mb-1">Informasi Template</p>
                <p>Upload template Word (.docx atau .doc) yang akan digunakan untuk generate surat tugas. Pilih kategori template sesuai jumlah pegawai dan <strong>tipe surat</strong> (Umum untuk ASN biasa, Penyuluh untuk penyuluh kehutanan). Template harus menggunakan placeholder seperti <code class="bg-blue-100 px-1 rounded">${nomor_surat}</code>, <code class="bg-blue-100 px-1 rounded">${tanggal_surat}</code>, dll. Sistem akan otomatis memilih template yang sesuai saat generate surat.</p>
            </div>
        </div>
    </div>

    <!-- Daftar Template -->
    <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
        <?php if (count($templates) > 0): ?>
            <div class="overflow-x-auto">
                <table class="w-full text-left border-collapse">
                    <thead>
                        <tr class="bg-slate-50 text-slate-600 text-sm uppercase tracking-wider">
                            <th class="px-6 py-4 font-semibold border-b border-slate-100">No</th>
                            <th class="px-6 py-4 font-semibold border-b border-slate-100">Nama Template</th>
                            <th class="px-6 py-4 font-semibold border-b border-slate-100">Kategori</th>
                            <th class="px-6 py-4 font-semibold border-b border-slate-100">Tipe</th>
                            <th class="px-6 py-4 font-semibold border-b border-slate-100">File</th>
                            <th class="px-6 py-4 font-semibold border-b border-slate-100">Ukuran</th>
                            <th class="px-6 py-4 font-semibold border-b border-slate-100">Status</th>
                            <th class="px-6 py-4 font-semibold border-b border-slate-100">Deskripsi</th>
                            <th class="px-6 py-4 font-semibold border-b border-slate-100 text-right">Aksi</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-100">
                        <?php $no = 1; foreach ($templates as $template): ?>
                            <tr class="hover:bg-slate-50 transition-colors">
                                <td class="px-6 py-4 text-sm text-slate-600"><?php echo $no++; ?></td>
                                <td class="px-6 py-4 text-sm font-medium text-slate-900">
                                    <?php echo htmlspecialchars($template['nama']); ?>
                                </td>
                                <td class="px-6 py-4 text-sm">
                                    <?php
                                    $kategori_labels = [
                                        '1_pegawai' => '1 Pegawai',
                                        '2_pegawai' => '2 Pegawai',
                                        '3_pegawai' => '3 Pegawai',
                                        'banyak_pegawai' => 'Banyak Pegawai'
                                    ];
                                    $kategori = isset($template['kategori_template']) ? $template['kategori_template'] : 'banyak_pegawai';
                                    $kategori_label = isset($kategori_labels[$kategori]) ? $kategori_labels[$kategori] : 'Banyak Pegawai';
                                    ?>
                                    <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-indigo-100 text-indigo-800">
                                        <?php echo $kategori_label; ?>
                                    </span>
                                </td>
                                <td class="px-6 py-4 text-sm">
                                    <?php 
                                    $tipe_surat = isset($template['tipe_surat']) ? $template['tipe_surat'] : 'umum';
                                    if ($tipe_surat === 'penyuluh'): ?>
                                        <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-emerald-100 text-emerald-800">
                                            <i class='bx bx-leaf mr-1'></i>Penyuluh
                                        </span>
                                    <?php else: ?>
                                        <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-blue-100 text-blue-800">
                                            <i class='bx bx-briefcase mr-1'></i>Umum
                                        </span>
                                    <?php endif; ?>
                                </td>
                                <td class="px-6 py-4 text-sm text-slate-600">
                                    <div class="flex items-center gap-2">
                                        <i class='bx bx-file text-indigo-600'></i>
                                        <span class="font-mono text-xs"><?php echo htmlspecialchars($template['nama_file']); ?></span>
                                    </div>
                                </td>
                                <td class="px-6 py-4 text-sm text-slate-600">
                                    <?php echo number_format($template['ukuran_file'] / 1024, 2); ?> KB
                                </td>
                                <td class="px-6 py-4 text-sm">
                                    <?php if ($template['is_default']): ?>
                                        <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-800">
                                            <i class='bx bx-check-circle mr-1'></i> Default
                                        </span>
                                    <?php else: ?>
                                        <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-slate-100 text-slate-600">
                                            Standar
                                        </span>
                                    <?php endif; ?>
                                </td>
                                <td class="px-6 py-4 text-sm text-slate-600 max-w-xs truncate" title="<?php echo htmlspecialchars($template['deskripsi']); ?>">
                                    <?php echo htmlspecialchars($template['deskripsi'] ?: '-'); ?>
                                </td>
                                <td class="px-6 py-4 text-sm text-right">
                                    <div class="flex items-center justify-end gap-1">
                                        <?php if (!$template['is_default']): ?>
                                            <button class="text-green-600 hover:text-green-700 set-default-btn transition-colors p-1" 
                                                    data-id="<?php echo $template['id']; ?>"
                                                    title="Set sebagai Default">
                                                <i class='bx bx-check text-xl'></i>
                                            </button>
                                        <?php endif; ?>
                                        <button class="text-amber-500 hover:text-amber-600 edit-template-btn transition-colors p-1"
                                                data-id="<?php echo $template['id']; ?>"
                                                data-nama="<?php echo htmlspecialchars($template['nama']); ?>"
                                                data-kategori="<?php echo htmlspecialchars($template['kategori_template'] ?? 'banyak_pegawai'); ?>"
                                                data-tipe="<?php echo htmlspecialchars($template['tipe_surat'] ?? 'umum'); ?>"
                                                data-deskripsi="<?php echo htmlspecialchars($template['deskripsi'] ?? ''); ?>"
                                                data-default="<?php echo $template['is_default']; ?>"
                                                data-file="<?php echo htmlspecialchars($template['nama_file']); ?>"
                                                title="Edit">
                                            <i class='bx bx-edit text-xl'></i>
                                        </button>
                                        <a href="../assets/templates/<?php echo htmlspecialchars($template['nama_file']); ?>" 
                                           download
                                           class="text-blue-600 hover:text-blue-700 transition-colors p-1"
                                           title="Download">
                                            <i class='bx bx-download text-xl'></i>
                                        </a>
                                        <button class="text-red-600 hover:text-red-700 delete-btn transition-colors p-1"
                                                data-id="<?php echo $template['id']; ?>"
                                                data-nama="<?php echo htmlspecialchars($template['nama']); ?>"
                                                title="Hapus">
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
            <div class="p-12 text-center text-slate-500">
                <i class='bx bx-file-blank text-6xl mb-4 text-slate-300'></i>
                <p class="text-lg font-medium mb-2">Belum ada template</p>
                <p class="text-sm mb-4">Upload template Word pertama Anda untuk memulai</p>
                <button onclick="openModal('uploadTemplateModal')" class="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg inline-flex items-center gap-2 transition-colors">
                    <i class='bx bx-upload'></i> Upload Template
                </button>
            </div>
        <?php endif; ?>
    </div>
</div>

<!-- Modal Upload Template -->
<div id="uploadTemplateModal" class="fixed inset-0 z-50 hidden overflow-y-auto" aria-labelledby="modal-title" role="dialog" aria-modal="true">
    <div class="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
        <div class="fixed inset-0 bg-slate-900 bg-opacity-75 transition-opacity" aria-hidden="true" onclick="closeModal('uploadTemplateModal')"></div>
        <span class="hidden sm:inline-block sm:align-middle sm:h-screen" aria-hidden="true">&#8203;</span>
        <div class="inline-block align-bottom bg-white rounded-xl text-left overflow-hidden shadow-xl transform transition-all sm:my-8 sm:align-middle sm:max-w-2xl sm:w-full">
            <div class="bg-white px-4 pt-5 pb-4 sm:p-6 sm:pb-4">
                <div class="flex justify-between items-center mb-4">
                    <h3 class="text-lg leading-6 font-medium text-slate-900">Upload Template Word</h3>
                    <button onclick="closeModal('uploadTemplateModal')" class="text-slate-400 hover:text-slate-500">
                        <i class='bx bx-x text-2xl'></i>
                    </button>
                </div>
                <form method="POST" enctype="multipart/form-data" id="uploadForm">
                    <input type="hidden" name="action" value="upload">
                    <div class="space-y-4">
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Nama Template</label>
                            <input type="text" name="nama_template" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none" 
                                   placeholder="Contoh: Template Surat Tugas Standar">
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Kategori Template <span class="text-red-500">*</span></label>
                            <select name="kategori_template" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                                <option value="">Pilih Kategori</option>
                                <option value="1_pegawai">1 Pegawai</option>
                                <option value="2_pegawai">2 Pegawai</option>
                                <option value="3_pegawai">3 Pegawai</option>
                                <option value="banyak_pegawai">Banyak Pegawai (4+ pegawai)</option>
                            </select>
                            <p class="text-xs text-slate-500 mt-1">Pilih kategori sesuai jumlah pegawai yang akan menggunakan template ini</p>
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Tipe Surat <span class="text-red-500">*</span></label>
                            <select name="tipe_surat" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                                <option value="umum">ASN Umum</option>
                                <option value="penyuluh">Khusus Penyuluh</option>
                            </select>
                            <p class="text-xs text-slate-500 mt-1">Template akan digunakan sesuai tipe surat tugas yang dibuat</p>
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">File Template</label>
                            <div class="mt-1 flex justify-center px-6 pt-5 pb-6 border-2 border-slate-300 border-dashed rounded-lg hover:border-indigo-400 transition-colors">
                                <div class="space-y-1 text-center">
                                    <i class='bx bx-cloud-upload text-4xl text-slate-400'></i>
                                    <div class="flex text-sm text-slate-600">
                                        <label for="template_file" class="relative cursor-pointer bg-white rounded-md font-medium text-indigo-600 hover:text-indigo-500 focus-within:outline-none focus-within:ring-2 focus-within:ring-offset-2 focus-within:ring-indigo-500">
                                            <span>Pilih file</span>
                                            <input id="template_file" name="template_file" type="file" accept=".docx,.doc" class="sr-only" required onchange="updateFileName(this)">
                                        </label>
                                        <p class="pl-1">atau drag and drop</p>
                                    </div>
                                    <p class="text-xs text-slate-500">DOCX atau DOC (maks. 10MB)</p>
                                    <p id="file-name" class="text-sm text-slate-600 mt-2"></p>
                                </div>
                            </div>
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Deskripsi (Opsional)</label>
                            <textarea name="deskripsi" rows="3" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none" 
                                      placeholder="Deskripsi template..."></textarea>
                        </div>
                        <div class="flex items-center">
                            <input type="checkbox" name="is_default" value="1" id="is_default" class="h-4 w-4 text-indigo-600 focus:ring-indigo-500 border-slate-300 rounded">
                            <label for="is_default" class="ml-2 block text-sm text-slate-700">
                                Set sebagai template default untuk kombinasi kategori dan tipe ini
                            </label>
                        </div>
                    </div>
                    <div class="mt-6 flex justify-end gap-3">
                        <button type="button" onclick="closeModal('uploadTemplateModal')" class="px-4 py-2 border border-slate-300 rounded-lg text-slate-700 hover:bg-slate-50 font-medium">Batal</button>
                        <button type="submit" class="px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 font-medium">Upload</button>
                    </div>
                </form>
            </div>
        </div>
    </div>
</div>

<!-- Modal Edit Template -->
<div id="editTemplateModal" class="fixed inset-0 z-50 hidden overflow-y-auto" aria-labelledby="modal-title" role="dialog" aria-modal="true">
    <div class="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
        <div class="fixed inset-0 bg-slate-900 bg-opacity-75 transition-opacity" aria-hidden="true" onclick="closeModal('editTemplateModal')"></div>
        <span class="hidden sm:inline-block sm:align-middle sm:h-screen" aria-hidden="true">&#8203;</span>
        <div class="inline-block align-bottom bg-white rounded-xl text-left overflow-hidden shadow-xl transform transition-all sm:my-8 sm:align-middle sm:max-w-2xl sm:w-full">
            <div class="bg-white px-4 pt-5 pb-4 sm:p-6 sm:pb-4">
                <div class="flex justify-between items-center mb-4">
                    <h3 class="text-lg leading-6 font-medium text-slate-900">Edit Template Word</h3>
                    <button onclick="closeModal('editTemplateModal')" class="text-slate-400 hover:text-slate-500">
                        <i class='bx bx-x text-2xl'></i>
                    </button>
                </div>
                <form method="POST" enctype="multipart/form-data" id="editForm">
                    <input type="hidden" name="action" value="update">
                    <input type="hidden" name="id" id="edit-template-id">
                    <div class="space-y-4">
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Nama Template <span class="text-red-500">*</span></label>
                            <input type="text" name="nama_template" id="edit-template-nama" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Kategori Template <span class="text-red-500">*</span></label>
                            <select name="kategori_template" id="edit-template-kategori" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                                <option value="1_pegawai">1 Pegawai</option>
                                <option value="2_pegawai">2 Pegawai</option>
                                <option value="3_pegawai">3 Pegawai</option>
                                <option value="banyak_pegawai">Banyak Pegawai (4+ pegawai)</option>
                            </select>
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Tipe Surat <span class="text-red-500">*</span></label>
                            <select name="tipe_surat" id="edit-template-tipe" required class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                                <option value="umum">ASN Umum</option>
                                <option value="penyuluh">Khusus Penyuluh</option>
                            </select>
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Ganti File Template</label>
                            <div class="bg-slate-50 rounded-lg p-3 mb-2 border border-slate-200">
                                <div class="flex items-center gap-2 text-sm text-slate-600">
                                    <i class='bx bx-file text-indigo-600'></i>
                                    <span>File saat ini: <strong id="edit-current-file" class="font-mono text-xs"></strong></span>
                                </div>
                            </div>
                            <div class="mt-1 flex justify-center px-6 pt-4 pb-4 border-2 border-slate-300 border-dashed rounded-lg hover:border-indigo-400 transition-colors">
                                <div class="space-y-1 text-center">
                                    <i class='bx bx-cloud-upload text-3xl text-slate-400'></i>
                                    <div class="flex text-sm text-slate-600">
                                        <label for="edit_template_file" class="relative cursor-pointer bg-white rounded-md font-medium text-indigo-600 hover:text-indigo-500">
                                            <span>Pilih file baru</span>
                                            <input id="edit_template_file" name="template_file" type="file" accept=".docx,.doc" class="sr-only" onchange="updateEditFileName(this)">
                                        </label>
                                        <p class="pl-1">(opsional)</p>
                                    </div>
                                    <p class="text-xs text-slate-500">Kosongkan jika tidak ingin mengganti file</p>
                                    <p id="edit-file-name" class="text-sm text-indigo-600 font-medium mt-1"></p>
                                </div>
                            </div>
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Deskripsi (Opsional)</label>
                            <textarea name="deskripsi" id="edit-template-deskripsi" rows="3" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none" 
                                      placeholder="Deskripsi template..."></textarea>
                        </div>
                        <div class="flex items-center">
                            <input type="checkbox" name="is_default" value="1" id="edit_is_default" class="h-4 w-4 text-indigo-600 focus:ring-indigo-500 border-slate-300 rounded">
                            <label for="edit_is_default" class="ml-2 block text-sm text-slate-700">
                                Set sebagai template default untuk kombinasi kategori dan tipe ini
                            </label>
                        </div>
                    </div>
                    <div class="mt-6 flex justify-end gap-3">
                        <button type="button" onclick="closeModal('editTemplateModal')" class="px-4 py-2 border border-slate-300 rounded-lg text-slate-700 hover:bg-slate-50 font-medium">Batal</button>
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
        // Reset form
        if (modalId === 'uploadTemplateModal') {
            document.getElementById('uploadForm').reset();
            document.getElementById('file-name').textContent = '';
        }
        if (modalId === 'editTemplateModal') {
            document.getElementById('edit-file-name').textContent = '';
            const editFileInput = document.getElementById('edit_template_file');
            if (editFileInput) editFileInput.value = '';
        }
    }

    function updateFileName(input) {
        const fileName = input.files[0] ? input.files[0].name : '';
        document.getElementById('file-name').textContent = fileName;
    }

    function updateEditFileName(input) {
        const fileName = input.files[0] ? input.files[0].name : '';
        document.getElementById('edit-file-name').textContent = fileName ? '📄 ' + fileName : '';
    }

    // Set Default Template
    document.querySelectorAll('.set-default-btn').forEach(button => {
        button.addEventListener('click', function() {
            const id = this.dataset.id;
            
            Swal.fire({
                title: 'Set sebagai Default?',
                text: 'Template ini akan digunakan sebagai template default untuk generate surat',
                icon: 'question',
                showCancelButton: true,
                confirmButtonColor: '#6366f1',
                cancelButtonColor: '#6b7280',
                confirmButtonText: 'Ya, Set Default',
                cancelButtonText: 'Batal'
            }).then((result) => {
                if (result.isConfirmed) {
                    const formData = new FormData();
                    formData.append('action', 'set_default');
                    formData.append('id', id);

                    fetch('template.php', {
                        method: 'POST',
                        body: formData
                    })
                    .then(response => response.json())
                    .then(data => {
                        if (data.success) {
                            Swal.fire({
                                icon: 'success',
                                title: 'Berhasil!',
                                text: data.message,
                                showConfirmButton: false,
                                timer: 1500
                            }).then(() => {
                                window.location.reload();
                            });
                        } else {
                            Swal.fire({
                                icon: 'error',
                                title: 'Error!',
                                text: data.message
                            });
                        }
                    })
                    .catch(error => {
                        Swal.fire({
                            icon: 'error',
                            title: 'Error!',
                            text: 'Terjadi kesalahan saat mengubah template default'
                        });
                    });
                }
            });
        });
    });

    // Edit Template
    document.querySelectorAll('.edit-template-btn').forEach(button => {
        button.addEventListener('click', function() {
            document.getElementById('edit-template-id').value = this.dataset.id;
            document.getElementById('edit-template-nama').value = this.dataset.nama;
            document.getElementById('edit-template-kategori').value = this.dataset.kategori;
            document.getElementById('edit-template-tipe').value = this.dataset.tipe || 'umum';
            document.getElementById('edit-template-deskripsi').value = this.dataset.deskripsi;
            document.getElementById('edit_is_default').checked = this.dataset.default === '1';
            document.getElementById('edit-current-file').textContent = this.dataset.file;
            document.getElementById('edit-file-name').textContent = '';
            
            // Reset file input
            const editFileInput = document.getElementById('edit_template_file');
            if (editFileInput) editFileInput.value = '';
            
            openModal('editTemplateModal');
        });
    });

    // Delete Template
    document.querySelectorAll('.delete-btn').forEach(button => {
        button.addEventListener('click', function() {
            const id = this.dataset.id;
            const nama = this.dataset.nama;
            
            Swal.fire({
                title: 'Apakah Anda yakin?',
                text: `Template "${nama}" akan dihapus permanen!`,
                icon: 'warning',
                showCancelButton: true,
                confirmButtonColor: '#ef4444',
                cancelButtonColor: '#3b82f6',
                confirmButtonText: 'Ya, hapus!',
                cancelButtonText: 'Batal'
            }).then((result) => {
                if (result.isConfirmed) {
                    const formData = new FormData();
                    formData.append('action', 'delete');
                    formData.append('id', id);

                    fetch('template.php', {
                        method: 'POST',
                        body: formData
                    })
                    .then(response => response.json())
                    .then(data => {
                        if (data.success) {
                            Swal.fire({
                                icon: 'success',
                                title: 'Terhapus!',
                                text: data.message,
                                showConfirmButton: false,
                                timer: 1500
                            }).then(() => {
                                window.location.reload();
                            });
                        } else {
                            Swal.fire({
                                icon: 'error',
                                title: 'Error!',
                                text: data.message
                            });
                        }
                    })
                    .catch(error => {
                        Swal.fire({
                            icon: 'error',
                            title: 'Error!',
                            text: 'Terjadi kesalahan saat menghapus template'
                        });
                    });
                }
            });
        });
    });
</script>

<?php include '../includes/footer.php'; ?>

