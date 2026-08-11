<?php
require_once '../config/database.php';

echo "<!DOCTYPE html>";
echo "<html lang='id'>";
echo "<head>";
echo "<meta charset='UTF-8'>";
echo "<meta name='viewport' content='width=device-width, initial-scale=1.0'>";
echo "<title>Check & Fix Templates</title>";
echo "<style>";
echo "body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }";
echo ".container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }";
echo "h2 { color: #333; border-bottom: 2px solid #4CAF50; padding-bottom: 10px; }";
echo "h3 { color: #555; margin-top: 30px; }";
echo "table { width: 100%; border-collapse: collapse; margin: 15px 0; }";
echo "th, td { padding: 12px; text-align: left; border: 1px solid #ddd; }";
echo "th { background: #4CAF50; color: white; }";
echo "tr:nth-child(even) { background: #f9f9f9; }";
echo ".success { background: #d4edda; color: #155724; padding: 15px; border-radius: 5px; border-left: 4px solid #28a745; margin: 15px 0; }";
echo ".error { background: #f8d7da; color: #721c24; padding: 15px; border-radius: 5px; border-left: 4px solid #dc3545; margin: 15px 0; }";
echo ".warning { background: #fff3cd; color: #856404; padding: 15px; border-radius: 5px; border-left: 4px solid #ffc107; margin: 15px 0; }";
echo ".info { background: #d1ecf1; color: #0c5460; padding: 15px; border-radius: 5px; border-left: 4px solid #17a2b8; margin: 15px 0; }";
echo ".btn { display: inline-block; padding: 10px 20px; margin: 5px; background: #4CAF50; color: white; text-decoration: none; border-radius: 5px; border: none; cursor: pointer; }";
echo ".btn:hover { background: #45a049; }";
echo ".btn-danger { background: #dc3545; }";
echo ".btn-danger:hover { background: #c82333; }";
echo ".btn-primary { background: #007bff; }";
echo ".btn-primary:hover { background: #0056b3; }";
echo ".badge { display: inline-block; padding: 3px 8px; border-radius: 3px; font-size: 12px; font-weight: bold; }";
echo ".badge-success { background: #28a745; color: white; }";
echo ".badge-danger { background: #dc3545; color: white; }";
echo ".badge-warning { background: #ffc107; color: #333; }";
echo "</style>";
echo "</head>";
echo "<body>";
echo "<div class='container'>";

echo "<h2>🔍 Diagnosa & Perbaikan Template</h2>";
echo "<p>Script ini akan mengecek dan memperbaiki masalah template yang hilang di database.</p>";

echo "<hr>";

// 1. CEK TEMPLATE DI DATABASE
echo "<h3>1️⃣ Template di Database</h3>";
try {
    // Pastikan kolom kategori_template ada
    try {
        $pdo->query("SELECT kategori_template FROM templates LIMIT 1");
    } catch (PDOException $e) {
        echo "<div class='warning'>⚠ Kolom kategori_template belum ada, menambahkan...</div>";
        try {
            $pdo->exec("ALTER TABLE templates ADD COLUMN kategori_template ENUM('1_pegawai', '2_pegawai', '3_pegawai', 'banyak_pegawai') NOT NULL DEFAULT 'banyak_pegawai' AFTER nama");
            echo "<div class='success'>✓ Kolom kategori_template berhasil ditambahkan</div>";
        } catch (PDOException $e2) {
            echo "<div class='error'>✗ Error menambahkan kolom: " . $e2->getMessage() . "</div>";
        }
    }
    
    $stmt = $pdo->query("SELECT id, nama, kategori_template, nama_file, path_file, ukuran_file, is_default, created_at FROM templates ORDER BY kategori_template, is_default DESC");
    $templates_db = $stmt->fetchAll();
    
    if (count($templates_db) > 0) {
        echo "<table>";
        echo "<tr><th>ID</th><th>Nama</th><th>Kategori</th><th>File</th><th>Ukuran</th><th>Default</th><th>Tanggal</th></tr>";
        foreach ($templates_db as $tmpl) {
            $default_badge = $tmpl['is_default'] ? "<span class='badge badge-success'>DEFAULT</span>" : "";
            $file_size = round($tmpl['ukuran_file'] / 1024, 2) . " KB";
            echo "<tr>";
            echo "<td>{$tmpl['id']}</td>";
            echo "<td>{$tmpl['nama']}</td>";
            echo "<td><strong>{$tmpl['kategori_template']}</strong></td>";
            echo "<td>{$tmpl['nama_file']}</td>";
            echo "<td>{$file_size}</td>";
            echo "<td>{$default_badge}</td>";
            echo "<td>" . date('d/m/Y H:i', strtotime($tmpl['created_at'])) . "</td>";
            echo "</tr>";
        }
        echo "</table>";
        echo "<div class='info'>📊 Total: " . count($templates_db) . " template di database</div>";
    } else {
        echo "<div class='warning'>⚠ Tidak ada template di database!</div>";
    }
} catch (PDOException $e) {
    echo "<div class='error'>✗ Error: " . $e->getMessage() . "</div>";
}

echo "<hr>";

// 2. CEK FILE FISIK DI FOLDER
echo "<h3>2️⃣ File Template di Folder</h3>";
$template_dir = __DIR__ . '/../assets/templates/';
$physical_files = [];

if (is_dir($template_dir)) {
    $files = scandir($template_dir);
    echo "<table>";
    echo "<tr><th>No</th><th>Nama File</th><th>Ukuran</th><th>Status di DB</th></tr>";
    $no = 1;
    foreach ($files as $file) {
        if ($file != '.' && $file != '..' && pathinfo($file, PATHINFO_EXTENSION) == 'docx') {
            $file_path = $template_dir . $file;
            $file_size = round(filesize($file_path) / 1024, 2) . " KB";
            
            // Cek apakah file ini ada di database
            $in_db = false;
            foreach ($templates_db as $tmpl) {
                if ($tmpl['nama_file'] == $file) {
                    $in_db = true;
                    break;
                }
            }
            
            $status_badge = $in_db ? 
                "<span class='badge badge-success'>✓ Ada di DB</span>" : 
                "<span class='badge badge-danger'>✗ Tidak di DB</span>";
            
            echo "<tr>";
            echo "<td>{$no}</td>";
            echo "<td>{$file}</td>";
            echo "<td>{$file_size}</td>";
            echo "<td>{$status_badge}</td>";
            echo "</tr>";
            
            $physical_files[] = [
                'nama_file' => $file,
                'path' => $file_path,
                'size' => filesize($file_path),
                'in_db' => $in_db
            ];
            
            $no++;
        }
    }
    echo "</table>";
    echo "<div class='info'>📁 Total: " . count($physical_files) . " file .docx di folder</div>";
} else {
    echo "<div class='error'>✗ Folder templates tidak ditemukan: {$template_dir}</div>";
}

echo "<hr>";

// 3. CEK KATEGORI YANG HILANG
echo "<h3>3️⃣ Analisa Kategori Template</h3>";
$required_categories = ['1_pegawai', '2_pegawai', '3_pegawai', 'banyak_pegawai'];
$existing_categories = array_column($templates_db, 'kategori_template');
$missing_categories = array_diff($required_categories, $existing_categories);

$kategori_labels = [
    '1_pegawai' => '1 Pegawai',
    '2_pegawai' => '2 Pegawai',
    '3_pegawai' => '3 Pegawai',
    'banyak_pegawai' => 'Banyak Pegawai (4+ pegawai)'
];

echo "<table>";
echo "<tr><th>Kategori</th><th>Label</th><th>Status</th></tr>";
foreach ($required_categories as $kat) {
    $status = in_array($kat, $existing_categories) ? 
        "<span class='badge badge-success'>✓ Ada</span>" : 
        "<span class='badge badge-danger'>✗ Hilang</span>";
    echo "<tr>";
    echo "<td><strong>{$kat}</strong></td>";
    echo "<td>{$kategori_labels[$kat]}</td>";
    echo "<td>{$status}</td>";
    echo "</tr>";
}
echo "</table>";

if (count($missing_categories) > 0) {
    echo "<div class='error'>";
    echo "<strong>⚠ MASALAH DITEMUKAN!</strong><br>";
    echo "Kategori yang hilang: <strong>" . implode(', ', array_map(function($k) use ($kategori_labels) { return $kategori_labels[$k]; }, $missing_categories)) . "</strong>";
    echo "</div>";
} else {
    echo "<div class='success'>✓ Semua kategori sudah ada di database</div>";
}

echo "<hr>";

// 4. SOLUSI OTOMATIS
echo "<h3>4️⃣ Perbaikan Otomatis</h3>";

if (count($missing_categories) > 0 && count($physical_files) > 0) {
    // Cari file template yang bisa digunakan
    $reference_file = null;
    
    // Prioritas: gunakan file yang sudah ada di DB
    foreach ($physical_files as $file) {
        if ($file['in_db']) {
            $reference_file = $file;
            break;
        }
    }
    
    // Jika tidak ada, gunakan file pertama
    if (!$reference_file && count($physical_files) > 0) {
        $reference_file = $physical_files[0];
    }
    
    if ($reference_file) {
        echo "<div class='info'>";
        echo "<strong>📋 Rencana Perbaikan:</strong><br>";
        echo "File referensi: <strong>{$reference_file['nama_file']}</strong> (" . round($reference_file['size']/1024, 2) . " KB)<br>";
        echo "File ini akan digunakan untuk kategori yang hilang.";
        echo "</div>";
        
        if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_POST['action']) && $_POST['action'] == 'fix_now') {
            echo "<div class='warning'>⏳ Memproses perbaikan...</div>";
            
            try {
                $pdo->beginTransaction();
                $fixed_count = 0;
                
                foreach ($missing_categories as $kategori) {
                    $nama_template = "Template Surat Tugas - " . $kategori_labels[$kategori];
                    
                    // Set sebagai default untuk kategori ini
                    $stmt = $pdo->prepare("INSERT INTO templates (nama, kategori_template, nama_file, path_file, ukuran_file, is_default, deskripsi) 
                                          VALUES (?, ?, ?, ?, ?, 1, ?)");
                    $stmt->execute([
                        $nama_template,
                        $kategori,
                        $reference_file['nama_file'],
                        'assets/templates/' . $reference_file['nama_file'],
                        $reference_file['size'],
                        "Template otomatis untuk kategori {$kategori} (dibuat oleh sistem)"
                    ]);
                    
                    $fixed_count++;
                    echo "<div class='success'>✓ Template untuk kategori '<strong>{$kategori_labels[$kategori]}</strong>' berhasil ditambahkan (ID: " . $pdo->lastInsertId() . ")</div>";
                }
                
                $pdo->commit();
                
                echo "<div class='success'>";
                echo "<strong>🎉 PERBAIKAN SELESAI!</strong><br>";
                echo "Berhasil menambahkan {$fixed_count} template baru ke database.<br><br>";
                echo "<a href='check_and_fix_templates.php' class='btn btn-primary'>🔄 Refresh Halaman</a> ";
                echo "<a href='template.php' class='btn'>📄 Lihat Daftar Template</a> ";
                echo "<a href='surat-tugas.php' class='btn'>📋 Kembali ke Surat Tugas</a>";
                echo "</div>";
                
            } catch (PDOException $e) {
                if ($pdo->inTransaction()) {
                    $pdo->rollBack();
                }
                echo "<div class='error'>✗ Error saat perbaikan: " . $e->getMessage() . "</div>";
            }
        } else {
            echo "<form method='POST'>";
            echo "<input type='hidden' name='action' value='fix_now'>";
            echo "<p><strong>Kategori yang akan diperbaiki:</strong></p>";
            echo "<ul>";
            foreach ($missing_categories as $kat) {
                echo "<li>{$kategori_labels[$kat]} (<code>{$kat}</code>)</li>";
            }
            echo "</ul>";
            echo "<button type='submit' class='btn' onclick='return confirm(\"Apakah Anda yakin ingin menambahkan template untuk kategori yang hilang?\")'>🔧 Perbaiki Sekarang</button>";
            echo "</form>";
        }
    } else {
        echo "<div class='error'>";
        echo "<strong>✗ Tidak dapat memperbaiki otomatis</strong><br>";
        echo "Tidak ada file template yang tersedia. Silakan upload template terlebih dahulu melalui halaman Template Word.";
        echo "</div>";
    }
} else if (count($missing_categories) == 0) {
    echo "<div class='success'>✓ Tidak ada masalah yang perlu diperbaiki. Semua kategori sudah memiliki template.</div>";
} else {
    echo "<div class='warning'>";
    echo "⚠ Tidak ada file template di folder. Silakan upload template terlebih dahulu.<br>";
    echo "<a href='template.php' class='btn'>📤 Upload Template</a>";
    echo "</div>";
}

echo "<hr>";

// 5. INFORMASI TAMBAHAN
echo "<h3>ℹ️ Informasi</h3>";
echo "<div class='info'>";
echo "<p><strong>Cara kerja sistem template:</strong></p>";
echo "<ol>";
echo "<li>Sistem memerlukan template untuk 4 kategori: 1 pegawai, 2 pegawai, 3 pegawai, dan banyak pegawai (4+)</li>";
echo "<li>Satu file .docx yang sama bisa digunakan untuk beberapa kategori</li>";
echo "<li>Sistem akan memilih template berdasarkan jumlah pegawai dalam surat tugas</li>";
echo "<li>Jika template kategori spesifik tidak ada, sistem akan fallback ke template 'banyak_pegawai'</li>";
echo "</ol>";
echo "<p><strong>Penyebab error di server:</strong></p>";
echo "<ul>";
echo "<li>Database di server belum memiliki record template (meskipun file fisik sudah ada)</li>";
echo "<li>Saat migrasi, hanya file yang dipindah, tapi data database tidak di-export/import</li>";
echo "</ul>";
echo "<p><strong>Solusi:</strong></p>";
echo "<ul>";
echo "<li>Gunakan script ini untuk menambahkan record template ke database</li>";
echo "<li>Atau upload ulang template melalui halaman Template Word</li>";
echo "</ul>";
echo "</div>";

echo "<hr>";
echo "<div style='text-align: center; margin-top: 30px;'>";
echo "<a href='template.php' class='btn btn-primary'>📄 Kelola Template</a> ";
echo "<a href='surat-tugas.php' class='btn'>📋 Daftar Surat Tugas</a> ";
echo "<a href='check_and_fix_templates.php' class='btn'>🔄 Refresh</a>";
echo "</div>";

echo "</div>";
echo "</body>";
echo "</html>";
?>
