<?php
require_once '../config/database.php';

header('Content-Type: text/html; charset=utf-8');
?>
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Debug Template - Detail</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            min-height: 100vh;
        }
        .container { 
            max-width: 1400px; 
            margin: 0 auto; 
            background: white; 
            padding: 30px; 
            border-radius: 15px; 
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
        }
        h1 { 
            color: #333; 
            margin-bottom: 10px;
            font-size: 28px;
            border-bottom: 3px solid #667eea;
            padding-bottom: 15px;
        }
        h2 { 
            color: #555; 
            margin: 30px 0 15px 0;
            font-size: 20px;
            background: #f8f9fa;
            padding: 12px 15px;
            border-left: 4px solid #667eea;
            border-radius: 5px;
        }
        .info-box {
            background: #e3f2fd;
            border-left: 4px solid #2196F3;
            padding: 15px;
            margin: 15px 0;
            border-radius: 5px;
        }
        .success-box {
            background: #e8f5e9;
            border-left: 4px solid #4CAF50;
            padding: 15px;
            margin: 15px 0;
            border-radius: 5px;
        }
        .error-box {
            background: #ffebee;
            border-left: 4px solid #f44336;
            padding: 15px;
            margin: 15px 0;
            border-radius: 5px;
        }
        .warning-box {
            background: #fff3e0;
            border-left: 4px solid #ff9800;
            padding: 15px;
            margin: 15px 0;
            border-radius: 5px;
        }
        table { 
            width: 100%; 
            border-collapse: collapse; 
            margin: 15px 0;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        th, td { 
            padding: 12px 15px; 
            text-align: left; 
            border: 1px solid #ddd; 
        }
        th { 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white; 
            font-weight: 600;
            text-transform: uppercase;
            font-size: 12px;
            letter-spacing: 0.5px;
        }
        tr:nth-child(even) { background: #f8f9fa; }
        tr:hover { background: #e9ecef; }
        .badge {
            display: inline-block;
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 11px;
            font-weight: bold;
            text-transform: uppercase;
        }
        .badge-success { background: #4CAF50; color: white; }
        .badge-danger { background: #f44336; color: white; }
        .badge-warning { background: #ff9800; color: white; }
        .badge-info { background: #2196F3; color: white; }
        .btn {
            display: inline-block;
            padding: 12px 24px;
            margin: 5px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            text-decoration: none;
            border-radius: 25px;
            border: none;
            cursor: pointer;
            font-weight: 600;
            transition: transform 0.2s, box-shadow 0.2s;
        }
        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
        }
        .btn-success { background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%); }
        .btn-danger { background: linear-gradient(135deg, #f44336 0%, #da190b 100%); }
        code {
            background: #f5f5f5;
            padding: 2px 6px;
            border-radius: 3px;
            font-family: 'Courier New', monospace;
            color: #d63384;
        }
        pre {
            background: #2d2d2d;
            color: #f8f8f2;
            padding: 15px;
            border-radius: 5px;
            overflow-x: auto;
            margin: 10px 0;
        }
        .stat-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin: 20px 0;
        }
        .stat-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 10px;
            text-align: center;
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }
        .stat-number {
            font-size: 36px;
            font-weight: bold;
            margin: 10px 0;
        }
        .stat-label {
            font-size: 14px;
            opacity: 0.9;
        }
    </style>
</head>
<body>
<div class="container">
    <h1>🔍 Debug Template - Analisa Detail</h1>
    <p style="color: #666; margin-bottom: 20px;">Script ini akan mengecek secara detail kondisi template di database dan file fisik</p>

    <?php
    // STATISTIK CEPAT
    echo "<div class='stat-grid'>";
    
    // Hitung template di database
    try {
        $stmt = $pdo->query("SELECT COUNT(*) as total FROM templates");
        $total_db = $stmt->fetch()['total'];
        echo "<div class='stat-card'>";
        echo "<div class='stat-label'>Template di Database</div>";
        echo "<div class='stat-number'>{$total_db}</div>";
        echo "</div>";
    } catch (PDOException $e) {
        echo "<div class='stat-card' style='background: #f44336;'>";
        echo "<div class='stat-label'>Template di Database</div>";
        echo "<div class='stat-number'>ERROR</div>";
        echo "</div>";
    }
    
    // Hitung file fisik
    $template_dir = __DIR__ . '/../assets/templates/';
    $file_count = 0;
    if (is_dir($template_dir)) {
        $files = scandir($template_dir);
        foreach ($files as $file) {
            if ($file != '.' && $file != '..' && pathinfo($file, PATHINFO_EXTENSION) == 'docx') {
                $file_count++;
            }
        }
    }
    echo "<div class='stat-card'>";
    echo "<div class='stat-label'>File Template (.docx)</div>";
    echo "<div class='stat-number'>{$file_count}</div>";
    echo "</div>";
    
    // Kategori yang diperlukan
    echo "<div class='stat-card'>";
    echo "<div class='stat-label'>Kategori Diperlukan</div>";
    echo "<div class='stat-number'>4</div>";
    echo "</div>";
    
    echo "</div>";
    
    // CEK KOLOM KATEGORI_TEMPLATE
    echo "<h2>1️⃣ Cek Struktur Tabel</h2>";
    try {
        $stmt = $pdo->query("DESCRIBE templates");
        $columns = $stmt->fetchAll();
        
        $has_kategori = false;
        echo "<table>";
        echo "<tr><th>Kolom</th><th>Tipe</th><th>Null</th><th>Default</th></tr>";
        foreach ($columns as $col) {
            if ($col['Field'] == 'kategori_template') {
                $has_kategori = true;
            }
            echo "<tr>";
            echo "<td><strong>{$col['Field']}</strong></td>";
            echo "<td><code>{$col['Type']}</code></td>";
            echo "<td>{$col['Null']}</td>";
            echo "<td>" . ($col['Default'] ?: '-') . "</td>";
            echo "</tr>";
        }
        echo "</table>";
        
        if ($has_kategori) {
            echo "<div class='success-box'>✅ Kolom <code>kategori_template</code> sudah ada</div>";
        } else {
            echo "<div class='error-box'>❌ Kolom <code>kategori_template</code> TIDAK ADA! Ini penyebab error.</div>";
            echo "<div class='warning-box'>";
            echo "<strong>Solusi:</strong> Jalankan query berikut di phpMyAdmin:<br>";
            echo "<pre>ALTER TABLE templates ADD COLUMN kategori_template ENUM('1_pegawai', '2_pegawai', '3_pegawai', 'banyak_pegawai') NOT NULL DEFAULT 'banyak_pegawai' AFTER nama;</pre>";
            echo "</div>";
        }
    } catch (PDOException $e) {
        echo "<div class='error-box'>❌ Error: " . $e->getMessage() . "</div>";
    }
    
    // CEK TEMPLATE DI DATABASE
    echo "<h2>2️⃣ Template di Database</h2>";
    try {
        $stmt = $pdo->query("SELECT * FROM templates ORDER BY kategori_template, is_default DESC");
        $templates = $stmt->fetchAll();
        
        if (count($templates) > 0) {
            echo "<table>";
            echo "<tr><th>ID</th><th>Nama</th><th>Kategori</th><th>File</th><th>Ukuran</th><th>Default</th><th>Tanggal</th></tr>";
            foreach ($templates as $t) {
                $kategori = isset($t['kategori_template']) ? $t['kategori_template'] : '<span class="badge badge-danger">TIDAK ADA</span>';
                $default_badge = $t['is_default'] ? '<span class="badge badge-success">DEFAULT</span>' : '';
                $size = round($t['ukuran_file'] / 1024, 2) . ' KB';
                
                echo "<tr>";
                echo "<td>{$t['id']}</td>";
                echo "<td>{$t['nama']}</td>";
                echo "<td><strong>{$kategori}</strong></td>";
                echo "<td>{$t['nama_file']}</td>";
                echo "<td>{$size}</td>";
                echo "<td>{$default_badge}</td>";
                echo "<td>" . date('d/m/Y H:i', strtotime($t['created_at'])) . "</td>";
                echo "</tr>";
            }
            echo "</table>";
        } else {
            echo "<div class='error-box'>❌ <strong>TIDAK ADA TEMPLATE DI DATABASE!</strong><br>Ini penyebab utama error. Database kosong.</div>";
        }
    } catch (PDOException $e) {
        echo "<div class='error-box'>❌ Error query: " . $e->getMessage() . "</div>";
    }
    
    // CEK KATEGORI
    echo "<h2>3️⃣ Analisa Kategori</h2>";
    $required = ['1_pegawai', '2_pegawai', '3_pegawai', 'banyak_pegawai'];
    $labels = [
        '1_pegawai' => '1 Pegawai',
        '2_pegawai' => '2 Pegawai', 
        '3_pegawai' => '3 Pegawai',
        'banyak_pegawai' => 'Banyak Pegawai (4+ pegawai)'
    ];
    
    try {
        $stmt = $pdo->query("SELECT DISTINCT kategori_template FROM templates");
        $existing = $stmt->fetchAll(PDO::FETCH_COLUMN);
        
        echo "<table>";
        echo "<tr><th>Kategori</th><th>Label</th><th>Status</th><th>Jumlah</th></tr>";
        
        $missing = [];
        foreach ($required as $kat) {
            $exists = in_array($kat, $existing);
            
            // Hitung jumlah template untuk kategori ini
            $stmt = $pdo->prepare("SELECT COUNT(*) as total FROM templates WHERE kategori_template = ?");
            $stmt->execute([$kat]);
            $count = $stmt->fetch()['total'];
            
            echo "<tr>";
            echo "<td><code>{$kat}</code></td>";
            echo "<td>{$labels[$kat]}</td>";
            if ($exists && $count > 0) {
                echo "<td><span class='badge badge-success'>✅ ADA</span></td>";
                echo "<td>{$count} template</td>";
            } else {
                echo "<td><span class='badge badge-danger'>❌ TIDAK ADA</span></td>";
                echo "<td>0 template</td>";
                $missing[] = $kat;
            }
            echo "</tr>";
        }
        echo "</table>";
        
        if (count($missing) > 0) {
            echo "<div class='error-box'>";
            echo "<strong>❌ MASALAH DITEMUKAN!</strong><br>";
            echo "Kategori yang hilang: <strong>" . implode(', ', array_map(function($k) use ($labels) { return $labels[$k]; }, $missing)) . "</strong><br><br>";
            echo "Ini sebabnya error terjadi saat generate surat dengan kategori tersebut.";
            echo "</div>";
        } else {
            echo "<div class='success-box'>✅ Semua kategori sudah ada!</div>";
        }
    } catch (PDOException $e) {
        echo "<div class='error-box'>❌ Error: " . $e->getMessage() . "</div>";
    }
    
    // CEK FILE FISIK
    echo "<h2>4️⃣ File Template di Server</h2>";
    echo "<div class='info-box'>📁 Folder: <code>{$template_dir}</code></div>";
    
    if (is_dir($template_dir)) {
        $files = scandir($template_dir);
        $docx_files = [];
        
        foreach ($files as $file) {
            if ($file != '.' && $file != '..' && pathinfo($file, PATHINFO_EXTENSION) == 'docx') {
                $file_path = $template_dir . $file;
                $docx_files[] = [
                    'name' => $file,
                    'size' => filesize($file_path),
                    'path' => $file_path,
                    'readable' => is_readable($file_path)
                ];
            }
        }
        
        if (count($docx_files) > 0) {
            echo "<table>";
            echo "<tr><th>No</th><th>Nama File</th><th>Ukuran</th><th>Readable</th><th>Di Database?</th></tr>";
            $no = 1;
            foreach ($docx_files as $f) {
                // Cek apakah file ini ada di database
                $stmt = $pdo->prepare("SELECT COUNT(*) as total FROM templates WHERE nama_file = ?");
                $stmt->execute([$f['name']]);
                $in_db = $stmt->fetch()['total'] > 0;
                
                echo "<tr>";
                echo "<td>{$no}</td>";
                echo "<td><strong>{$f['name']}</strong></td>";
                echo "<td>" . round($f['size']/1024, 2) . " KB</td>";
                echo "<td>" . ($f['readable'] ? '<span class="badge badge-success">✅ YES</span>' : '<span class="badge badge-danger">❌ NO</span>') . "</td>";
                echo "<td>" . ($in_db ? '<span class="badge badge-success">✅ ADA</span>' : '<span class="badge badge-danger">❌ TIDAK</span>') . "</td>";
                echo "</tr>";
                $no++;
            }
            echo "</table>";
            
            echo "<div class='success-box'>✅ Ditemukan " . count($docx_files) . " file template .docx</div>";
        } else {
            echo "<div class='error-box'>❌ Tidak ada file .docx di folder templates!</div>";
        }
    } else {
        echo "<div class='error-box'>❌ Folder templates tidak ditemukan!</div>";
    }
    
    // KESIMPULAN & SOLUSI
    echo "<h2>5️⃣ Kesimpulan & Solusi</h2>";
    
    $problems = [];
    $solutions = [];
    
    // Cek masalah
    try {
        $stmt = $pdo->query("SELECT COUNT(*) as total FROM templates");
        $total_templates = $stmt->fetch()['total'];
        
        if ($total_templates == 0) {
            $problems[] = "Database tidak memiliki template sama sekali";
            $solutions[] = "Gunakan script <code>check_and_fix_templates.php</code> untuk auto-fix";
        }
        
        $stmt = $pdo->query("SELECT DISTINCT kategori_template FROM templates");
        $existing_kat = $stmt->fetchAll(PDO::FETCH_COLUMN);
        $missing_kat = array_diff($required, $existing_kat);
        
        if (count($missing_kat) > 0) {
            $problems[] = "Kategori template tidak lengkap: " . implode(', ', $missing_kat);
            $solutions[] = "Tambahkan template untuk kategori yang hilang";
        }
        
        if (count($problems) == 0) {
            echo "<div class='success-box'>";
            echo "<strong>✅ TIDAK ADA MASALAH DITEMUKAN!</strong><br>";
            echo "Semua template sudah lengkap. Jika masih error, kemungkinan masalah lain.";
            echo "</div>";
        } else {
            echo "<div class='error-box'>";
            echo "<strong>❌ MASALAH DITEMUKAN:</strong><br><ul>";
            foreach ($problems as $p) {
                echo "<li>{$p}</li>";
            }
            echo "</ul></div>";
            
            echo "<div class='warning-box'>";
            echo "<strong>🔧 SOLUSI:</strong><br><ul>";
            foreach ($solutions as $s) {
                echo "<li>{$s}</li>";
            }
            echo "</ul></div>";
        }
    } catch (PDOException $e) {
        echo "<div class='error-box'>❌ Error: " . $e->getMessage() . "</div>";
    }
    
    // TOMBOL AKSI
    echo "<div style='text-align: center; margin-top: 30px; padding: 20px; background: #f8f9fa; border-radius: 10px;'>";
    echo "<h3 style='margin-bottom: 15px;'>🛠️ Aksi Cepat</h3>";
    echo "<a href='check_and_fix_templates.php' class='btn btn-success'>🔧 Auto-Fix Template</a> ";
    echo "<a href='template.php' class='btn'>📄 Kelola Template</a> ";
    echo "<a href='debug_templates.php' class='btn'>🔄 Refresh</a>";
    echo "</div>";
    ?>
</div>
</body>
</html>
