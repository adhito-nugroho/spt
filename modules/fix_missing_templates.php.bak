<?php
require_once '../config/database.php';

echo "<h2>Fix Missing Templates</h2>";
echo "<hr>";

// Cek template yang ada
echo "<h3>1. Cek Template yang Ada</h3>";
try {
    $stmt = $pdo->query("SELECT id, nama, kategori_template, nama_file, is_default FROM templates ORDER BY kategori_template");
    $existing_templates = $stmt->fetchAll();
    
    echo "<table border='1' cellpadding='5'>";
    echo "<tr><th>ID</th><th>Nama</th><th>Kategori</th><th>Nama File</th><th>Default</th></tr>";
    foreach ($existing_templates as $tmpl) {
        echo "<tr>";
        echo "<td>{$tmpl['id']}</td>";
        echo "<td>{$tmpl['nama']}</td>";
        echo "<td>{$tmpl['kategori_template']}</td>";
        echo "<td>{$tmpl['nama_file']}</td>";
        echo "<td>" . ($tmpl['is_default'] ? 'YES' : 'NO') . "</td>";
        echo "</tr>";
    }
    echo "</table>";
} catch (PDOException $e) {
    echo "<p style='color: red;'>Error: " . $e->getMessage() . "</p>";
}

echo "<hr>";

// Cek kategori yang hilang
echo "<h3>2. Cek Kategori yang Hilang</h3>";
$required_categories = ['1_pegawai', '2_pegawai', '3_pegawai', 'banyak_pegawai'];
$existing_categories = array_column($existing_templates, 'kategori_template');
$missing_categories = array_diff($required_categories, $existing_categories);

if (count($missing_categories) > 0) {
    echo "<p style='color: orange;'>⚠ Kategori yang hilang: <strong>" . implode(', ', $missing_categories) . "</strong></p>";
} else {
    echo "<p style='color: green;'>✓ Semua kategori sudah ada</p>";
}

echo "<hr>";

// Solusi: Gunakan template banyak_pegawai sebagai fallback untuk kategori yang hilang
echo "<h3>3. Solusi: Tambahkan Template untuk Kategori yang Hilang</h3>";

if (count($missing_categories) > 0) {
    // Cari template banyak_pegawai yang akan dijadikan referensi
    $stmt = $pdo->prepare("SELECT * FROM templates WHERE kategori_template = 'banyak_pegawai' LIMIT 1");
    $stmt->execute();
    $template_reference = $stmt->fetch();
    
    if ($template_reference) {
        echo "<p>Menggunakan template '<strong>{$template_reference['nama']}</strong>' sebagai referensi untuk kategori yang hilang.</p>";
        echo "<p style='color: blue;'>ℹ File yang sama akan digunakan untuk semua kategori yang hilang (sistem akan otomatis memilih yang sesuai).</p>";
        
        echo "<form method='POST'>";
        echo "<input type='hidden' name='action' value='add_missing_templates'>";
        echo "<input type='hidden' name='reference_file' value='" . htmlspecialchars($template_reference['nama_file']) . "'>";
        echo "<input type='hidden' name='reference_path' value='" . htmlspecialchars($template_reference['path_file']) . "'>";
        echo "<input type='hidden' name='reference_size' value='" . $template_reference['ukuran_file'] . "'>";
        
        echo "<table border='1' cellpadding='5'>";
        echo "<tr><th>Kategori</th><th>Nama Template</th><th>Aksi</th></tr>";
        
        foreach ($missing_categories as $kategori) {
            $kategori_labels = [
                '1_pegawai' => '1 Pegawai',
                '2_pegawai' => '2 Pegawai',
                '3_pegawai' => '3 Pegawai',
                'banyak_pegawai' => 'Banyak Pegawai'
            ];
            
            $label = $kategori_labels[$kategori];
            $nama_template = "Template Surat Tugas - {$label}";
            
            echo "<tr>";
            echo "<td><strong>{$kategori}</strong></td>";
            echo "<td><input type='text' name='nama_{$kategori}' value='{$nama_template}' style='width: 300px;'></td>";
            echo "<td><input type='checkbox' name='add_{$kategori}' value='1' checked> Tambahkan</td>";
            echo "</tr>";
        }
        
        echo "</table>";
        echo "<br>";
        echo "<button type='submit' style='background: #4CAF50; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer;'>Tambahkan Template yang Hilang</button>";
        echo "</form>";
    } else {
        echo "<p style='color: red;'>✗ Tidak ada template referensi yang bisa digunakan. Silakan upload template terlebih dahulu.</p>";
    }
} else {
    echo "<p style='color: green;'>✓ Tidak ada kategori yang hilang. Semua kategori sudah memiliki template.</p>";
}

echo "<hr>";

// Proses penambahan template
if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_POST['action']) && $_POST['action'] == 'add_missing_templates') {
    echo "<h3>4. Proses Penambahan Template</h3>";
    
    try {
        $pdo->beginTransaction();
        
        $reference_file = $_POST['reference_file'];
        $reference_path = $_POST['reference_path'];
        $reference_size = $_POST['reference_size'];
        
        $added_count = 0;
        
        foreach ($missing_categories as $kategori) {
            if (isset($_POST["add_{$kategori}"]) && $_POST["add_{$kategori}"] == '1') {
                $nama_template = $_POST["nama_{$kategori}"];
                
                // Set sebagai default untuk kategori ini
                $is_default = 1;
                
                $stmt = $pdo->prepare("INSERT INTO templates (nama, kategori_template, nama_file, path_file, ukuran_file, is_default, deskripsi) 
                                      VALUES (?, ?, ?, ?, ?, ?, ?)");
                $stmt->execute([
                    $nama_template,
                    $kategori,
                    $reference_file,
                    $reference_path,
                    $reference_size,
                    $is_default,
                    "Template otomatis untuk kategori {$kategori} (menggunakan file yang sama dengan template lain)"
                ]);
                
                $added_count++;
                echo "<p style='color: green;'>✓ Template untuk kategori '<strong>{$kategori}</strong>' berhasil ditambahkan</p>";
            }
        }
        
        $pdo->commit();
        
        echo "<p style='color: green; font-weight: bold;'>✓ Berhasil menambahkan {$added_count} template baru!</p>";
        echo "<p><a href='check_templates.php' style='background: #2196F3; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block;'>Cek Ulang Template</a></p>";
        echo "<p><a href='template.php' style='background: #6366f1; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block;'>Kembali ke Halaman Template</a></p>";
        
    } catch (PDOException $e) {
        if ($pdo->inTransaction()) {
            $pdo->rollBack();
        }
        echo "<p style='color: red;'>✗ Error: " . $e->getMessage() . "</p>";
    }
}

echo "<hr>";
echo "<h3>Informasi</h3>";
echo "<div style='background: #e3f2fd; padding: 15px; border-radius: 5px; border-left: 4px solid #2196F3;'>";
echo "<p><strong>Catatan:</strong></p>";
echo "<ul>";
echo "<li>Script ini akan menambahkan template untuk kategori yang hilang menggunakan file template yang sudah ada</li>";
echo "<li>Satu file template yang sama bisa digunakan untuk beberapa kategori</li>";
echo "<li>Sistem akan otomatis memilih template yang sesuai berdasarkan jumlah pegawai</li>";
echo "<li>Anda bisa upload template khusus untuk setiap kategori nanti melalui halaman Template Word</li>";
echo "</ul>";
echo "</div>";
?>
