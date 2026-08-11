<?php
require_once '../config/database.php';

echo "<h2>Debugging Template System</h2>";
echo "<hr>";

// 1. Cek apakah tabel templates ada
echo "<h3>1. Cek Tabel Templates</h3>";
try {
    $stmt = $pdo->query("SHOW TABLES LIKE 'templates'");
    $table_exists = $stmt->rowCount() > 0;
    
    if ($table_exists) {
        echo "<p style='color: green;'>✓ Tabel 'templates' ditemukan</p>";
        
        // Cek struktur tabel
        echo "<h4>Struktur Tabel:</h4>";
        $stmt = $pdo->query("DESCRIBE templates");
        $columns = $stmt->fetchAll();
        echo "<table border='1' cellpadding='5'>";
        echo "<tr><th>Field</th><th>Type</th><th>Null</th><th>Key</th><th>Default</th><th>Extra</th></tr>";
        foreach ($columns as $col) {
            echo "<tr>";
            echo "<td>{$col['Field']}</td>";
            echo "<td>{$col['Type']}</td>";
            echo "<td>{$col['Null']}</td>";
            echo "<td>{$col['Key']}</td>";
            echo "<td>{$col['Default']}</td>";
            echo "<td>{$col['Extra']}</td>";
            echo "</tr>";
        }
        echo "</table>";
    } else {
        echo "<p style='color: red;'>✗ Tabel 'templates' TIDAK ditemukan</p>";
    }
} catch (PDOException $e) {
    echo "<p style='color: red;'>Error: " . $e->getMessage() . "</p>";
}

echo "<hr>";

// 2. Cek data template yang ada
echo "<h3>2. Data Template di Database</h3>";
try {
    $stmt = $pdo->query("SELECT * FROM templates ORDER BY kategori_template, is_default DESC");
    $templates = $stmt->fetchAll();
    
    if (count($templates) > 0) {
        echo "<p>Total template: <strong>" . count($templates) . "</strong></p>";
        echo "<table border='1' cellpadding='5'>";
        echo "<tr><th>ID</th><th>Nama</th><th>Kategori</th><th>Nama File</th><th>Path File</th><th>Ukuran</th><th>Is Default</th><th>Created At</th></tr>";
        foreach ($templates as $tmpl) {
            echo "<tr>";
            echo "<td>{$tmpl['id']}</td>";
            echo "<td>{$tmpl['nama']}</td>";
            echo "<td>" . (isset($tmpl['kategori_template']) ? $tmpl['kategori_template'] : 'N/A') . "</td>";
            echo "<td>{$tmpl['nama_file']}</td>";
            echo "<td>{$tmpl['path_file']}</td>";
            echo "<td>" . number_format($tmpl['ukuran_file'] / 1024, 2) . " KB</td>";
            echo "<td>" . ($tmpl['is_default'] ? 'YES' : 'NO') . "</td>";
            echo "<td>{$tmpl['created_at']}</td>";
            echo "</tr>";
        }
        echo "</table>";
    } else {
        echo "<p style='color: orange;'>⚠ Tidak ada data template di database</p>";
    }
} catch (PDOException $e) {
    echo "<p style='color: red;'>Error: " . $e->getMessage() . "</p>";
}

echo "<hr>";

// 3. Cek file template di folder
echo "<h3>3. File Template di Folder assets/templates/</h3>";
$templates_dir = __DIR__ . '/../assets/templates/';
if (is_dir($templates_dir)) {
    echo "<p style='color: green;'>✓ Folder templates ditemukan: <code>{$templates_dir}</code></p>";
    
    $files = scandir($templates_dir);
    $template_files = array_filter($files, function($file) {
        return !in_array($file, ['.', '..']) && (pathinfo($file, PATHINFO_EXTENSION) === 'docx' || pathinfo($file, PATHINFO_EXTENSION) === 'doc');
    });
    
    if (count($template_files) > 0) {
        echo "<p>Total file template: <strong>" . count($template_files) . "</strong></p>";
        echo "<table border='1' cellpadding='5'>";
        echo "<tr><th>Nama File</th><th>Ukuran</th><th>Path Lengkap</th><th>Readable</th><th>Exists</th></tr>";
        foreach ($template_files as $file) {
            $full_path = $templates_dir . $file;
            $size = filesize($full_path);
            $readable = is_readable($full_path) ? 'YES' : 'NO';
            $exists = file_exists($full_path) ? 'YES' : 'NO';
            
            echo "<tr>";
            echo "<td>{$file}</td>";
            echo "<td>" . number_format($size / 1024, 2) . " KB</td>";
            echo "<td><code>{$full_path}</code></td>";
            echo "<td>{$readable}</td>";
            echo "<td>{$exists}</td>";
            echo "</tr>";
        }
        echo "</table>";
    } else {
        echo "<p style='color: orange;'>⚠ Tidak ada file template di folder</p>";
    }
} else {
    echo "<p style='color: red;'>✗ Folder templates TIDAK ditemukan</p>";
}

echo "<hr>";

// 4. Cek mapping antara database dan file
echo "<h3>4. Validasi Mapping Database vs File</h3>";
try {
    $stmt = $pdo->query("SELECT id, nama, kategori_template, nama_file, is_default FROM templates");
    $templates = $stmt->fetchAll();
    
    if (count($templates) > 0) {
        echo "<table border='1' cellpadding='5'>";
        echo "<tr><th>ID</th><th>Nama</th><th>Kategori</th><th>Nama File</th><th>Default</th><th>File Exists</th><th>Status</th></tr>";
        
        foreach ($templates as $tmpl) {
            $file_path = __DIR__ . '/../assets/templates/' . $tmpl['nama_file'];
            $file_exists = file_exists($file_path);
            $status = $file_exists ? "<span style='color: green;'>✓ OK</span>" : "<span style='color: red;'>✗ FILE NOT FOUND</span>";
            
            echo "<tr>";
            echo "<td>{$tmpl['id']}</td>";
            echo "<td>{$tmpl['nama']}</td>";
            echo "<td>" . (isset($tmpl['kategori_template']) ? $tmpl['kategori_template'] : 'N/A') . "</td>";
            echo "<td>{$tmpl['nama_file']}</td>";
            echo "<td>" . ($tmpl['is_default'] ? 'YES' : 'NO') . "</td>";
            echo "<td>" . ($file_exists ? 'YES' : 'NO') . "</td>";
            echo "<td>{$status}</td>";
            echo "</tr>";
        }
        echo "</table>";
    } else {
        echo "<p style='color: orange;'>⚠ Tidak ada data untuk divalidasi</p>";
    }
} catch (PDOException $e) {
    echo "<p style='color: red;'>Error: " . $e->getMessage() . "</p>";
}

echo "<hr>";

// 5. Simulasi pencarian template untuk berbagai jumlah pegawai
echo "<h3>5. Simulasi Pencarian Template</h3>";
$jumlah_pegawai_tests = [1, 2, 3, 4, 5, 10];

foreach ($jumlah_pegawai_tests as $jumlah) {
    echo "<h4>Untuk {$jumlah} Pegawai:</h4>";
    
    // Tentukan kategori
    $kategori_template = 'banyak_pegawai';
    if ($jumlah == 1) {
        $kategori_template = '1_pegawai';
    } elseif ($jumlah == 2) {
        $kategori_template = '2_pegawai';
    } elseif ($jumlah == 3) {
        $kategori_template = '3_pegawai';
    }
    
    echo "<p>Kategori dicari: <strong>{$kategori_template}</strong></p>";
    
    try {
        // Cari template default
        $stmt = $pdo->prepare("SELECT id, nama, kategori_template, nama_file, is_default FROM templates WHERE kategori_template = ? AND is_default = 1 LIMIT 1");
        $stmt->execute([$kategori_template]);
        $template_default = $stmt->fetch();
        
        if ($template_default) {
            $file_path = __DIR__ . '/../assets/templates/' . $template_default['nama_file'];
            $file_exists = file_exists($file_path);
            
            echo "<p style='color: green;'>✓ Template default ditemukan:</p>";
            echo "<ul>";
            echo "<li>ID: {$template_default['id']}</li>";
            echo "<li>Nama: {$template_default['nama']}</li>";
            echo "<li>File: {$template_default['nama_file']}</li>";
            echo "<li>File exists: " . ($file_exists ? 'YES' : 'NO') . "</li>";
            echo "</ul>";
        } else {
            echo "<p style='color: orange;'>⚠ Template default TIDAK ditemukan</p>";
            
            // Coba cari template non-default
            $stmt = $pdo->prepare("SELECT id, nama, kategori_template, nama_file, is_default FROM templates WHERE kategori_template = ? ORDER BY created_at DESC LIMIT 1");
            $stmt->execute([$kategori_template]);
            $template_fallback = $stmt->fetch();
            
            if ($template_fallback) {
                $file_path = __DIR__ . '/../assets/templates/' . $template_fallback['nama_file'];
                $file_exists = file_exists($file_path);
                
                echo "<p style='color: blue;'>ℹ Template fallback ditemukan:</p>";
                echo "<ul>";
                echo "<li>ID: {$template_fallback['id']}</li>";
                echo "<li>Nama: {$template_fallback['nama']}</li>";
                echo "<li>File: {$template_fallback['nama_file']}</li>";
                echo "<li>File exists: " . ($file_exists ? 'YES' : 'NO') . "</li>";
                echo "</ul>";
            } else {
                echo "<p style='color: red;'>✗ Template fallback juga TIDAK ditemukan</p>";
                
                // Coba fallback ke banyak_pegawai
                if ($kategori_template !== 'banyak_pegawai') {
                    echo "<p>Mencoba fallback ke 'banyak_pegawai'...</p>";
                    $stmt = $pdo->prepare("SELECT id, nama, kategori_template, nama_file, is_default FROM templates WHERE kategori_template = 'banyak_pegawai' ORDER BY is_default DESC, created_at DESC LIMIT 1");
                    $stmt->execute();
                    $template_universal = $stmt->fetch();
                    
                    if ($template_universal) {
                        $file_path = __DIR__ . '/../assets/templates/' . $template_universal['nama_file'];
                        $file_exists = file_exists($file_path);
                        
                        echo "<p style='color: blue;'>ℹ Template universal (banyak_pegawai) ditemukan:</p>";
                        echo "<ul>";
                        echo "<li>ID: {$template_universal['id']}</li>";
                        echo "<li>Nama: {$template_universal['nama']}</li>";
                        echo "<li>File: {$template_universal['nama_file']}</li>";
                        echo "<li>File exists: " . ($file_exists ? 'YES' : 'NO') . "</li>";
                        echo "</ul>";
                    } else {
                        echo "<p style='color: red;'>✗ Template universal juga TIDAK ditemukan</p>";
                    }
                }
            }
        }
    } catch (PDOException $e) {
        echo "<p style='color: red;'>Error: " . $e->getMessage() . "</p>";
    }
    
    echo "<hr>";
}

echo "<h3>Kesimpulan dan Rekomendasi</h3>";
echo "<div style='background: #f0f0f0; padding: 15px; border-radius: 5px;'>";
echo "<p><strong>Untuk mengatasi error 'template tidak ditemukan', pastikan:</strong></p>";
echo "<ol>";
echo "<li>Tabel 'templates' memiliki kolom 'kategori_template'</li>";
echo "<li>Ada minimal 1 template untuk setiap kategori (1_pegawai, 2_pegawai, 3_pegawai, banyak_pegawai)</li>";
echo "<li>File template yang tercatat di database benar-benar ada di folder assets/templates/</li>";
echo "<li>Minimal ada 1 template dengan is_default = 1 untuk setiap kategori</li>";
echo "<li>File template bisa dibaca (readable) dan tidak corrupt</li>";
echo "</ol>";
echo "</div>";
?>
