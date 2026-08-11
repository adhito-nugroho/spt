<?php
require_once '../config/database.php';
require '../vendor/autoload.php';

use PhpOffice\PhpWord\TemplateProcessor;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\Settings;

/**
 * Konversi file .docx ke .pdf menggunakan Microsoft Word via COM Object.
 * Hasil PDF 100% sama persis dengan tampilan di Word.
 * Syarat: Windows Server, Microsoft Word terinstall, PHP extension COM aktif (php_com_dotnet.dll).
 * Return path file PDF jika berhasil, null jika gagal atau COM tidak tersedia.
 *
 * Aktifkan COM: di php.ini pastikan extension=php_com_dotnet.dll tidak dikomentari; restart web server.
 * User PHP harus punya akses ke Word (IIS: Application Pool identity; Laragon: run as user).
 * Diagnostik: buka modules/debug_word_com.php. Detail: lihat README_KONVERSI_PDF.md.
 */
function konversi_docx_ke_pdf_word_com($pathDocx, $maxRetries = 3) {
    if (!class_exists('COM')) {
        error_log('Word COM: class COM tidak ada (extension php_com_dotnet tidak aktif)');
        return null;
    }
    if (!file_exists($pathDocx) || !is_readable($pathDocx)) {
        error_log('Word COM: file docx tidak ditemukan atau tidak bisa dibaca: ' . $pathDocx);
        return null;
    }

    $pathDocxReal = realpath($pathDocx);
    if ($pathDocxReal === false) {
        return null;
    }

    // Cari direktori yang benar-benar writable
    $candidateDirs = [
        dirname($pathDocxReal), // folder yang sama dengan file docx (pasti writable)
        'C:\\laragon\\tmp',
        'C:\\Windows\\Temp',
        sys_get_temp_dir(),
    ];
    $dir = null;
    foreach ($candidateDirs as $candidate) {
        if (is_dir($candidate) && is_writable($candidate)) {
            $dir = $candidate;
            break;
        }
    }
    if ($dir === null) {
        error_log('Word COM: tidak ada direktori writable ditemukan');
        return null;
    }

    $pathPdf = $dir . DIRECTORY_SEPARATOR . pathinfo($pathDocxReal, PATHINFO_FILENAME) . '.pdf';
    error_log('Word COM: menggunakan direktori: ' . $dir);

    for ($attempt = 1; $attempt <= $maxRetries; $attempt++) {
        clearstatcache(true, $pathDocxReal);
        $word = null;
        try {
            error_log("Word COM (Percobaan $attempt/$maxRetries): membuat instance Word.Application");
            $word = new COM('Word.Application');
            $word->Visible = false;
            $word->DisplayAlerts = 0; // wdAlertsNone

            error_log("Word COM (Percobaan $attempt/$maxRetries): membuka dokumen " . $pathDocxReal);
            $doc = $word->Documents->Open(
                $pathDocxReal,  // FileName
                false,          // ConfirmConversions
                true,           // ReadOnly
                false,          // AddToRecentFiles
                '',             // PasswordDocument
                '',             // PasswordTemplate
                true            // Revert
            );

            error_log("Word COM (Percobaan $attempt/$maxRetries): export ke PDF " . $pathPdf);
            $doc->ExportAsFixedFormat(
                $pathPdf,   // OutputFileName
                17,         // ExportFormat: wdExportFormatPDF = 17
                false,      // OpenAfterExport
                0,          // OptimizeFor: wdExportOptimizeForPrint = 0
                0,          // Range: wdExportAllDocument = 0
                1,          // From
                1,          // To
                0,          // Item: wdExportDocumentContent
                true,       // IncludeDocProps
                true,       // KeepIRM
                0,          // CreateBookmarks: wdExportCreateNoBookmarks
                true,       // DocStructureTags
                true,       // BitmapMissingFonts
                false       // UseISO19005_1
            );

            $doc->Close(false);
            error_log("Word COM (Percobaan $attempt/$maxRetries): dokumen ditutup");
            $word->Quit();
            $word = null;
            error_log("Word COM (Percobaan $attempt/$maxRetries): Word.Quit() selesai");

            if (file_exists($pathPdf) && filesize($pathPdf) > 100) {
                error_log("Word COM: PDF berhasil pada percobaan ke-$attempt: " . $pathPdf);
                return $pathPdf;
            }
            error_log("Word COM: file PDF tidak ada atau terlalu kecil pada percobaan ke-$attempt");
        } catch (Throwable $e) {
            error_log("Word COM Error (Percobaan $attempt/$maxRetries): " . $e->getMessage());
        } finally {
            if ($word !== null) {
                try {
                    $word->Quit();
                } catch (Throwable $e) {
                    error_log('Word COM Quit finally: ' . $e->getMessage());
                }
                $word = null;
            }
        }

        if ($attempt < $maxRetries) {
            usleep(500000); // Jeda 0.5 detik sebelum mencoba lagi
        }
    }

    return null;
}

/**
 * Konversi file .docx ke .pdf menggunakan LibreOffice (soffice).
 * Hasil PDF mendekati tampilan Word. Mengembalikan path file PDF jika berhasil, null jika gagal.
 */
function konversi_docx_ke_pdf_libreoffice($pathDocx) {
    if (!file_exists($pathDocx) || !is_readable($pathDocx)) {
        return null;
    }
    $dir = dirname($pathDocx);
    $pathPdf = $dir . DIRECTORY_SEPARATOR . pathinfo($pathDocx, PATHINFO_FILENAME) . '.pdf';

    $sofficePaths = [
        'C:\Program Files\LibreOffice\program\soffice.exe',
        'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
        'D:\laragon\bin\libreoffice\program\soffice.exe',
        '/usr/bin/soffice',
        '/usr/bin/libreoffice',
    ];
    if (defined('PATH_SOFFICE') && PATH_SOFFICE !== '') {
        array_unshift($sofficePaths, PATH_SOFFICE);
    }
    // Scan C:\Program Files\LibreOffice\*\program\soffice.exe (berbagai versi)
    if (PHP_OS_FAMILY === 'Windows' && is_dir('C:\Program Files\LibreOffice')) {
        foreach (new DirectoryIterator('C:\Program Files\LibreOffice') as $e) {
            if ($e->isDir() && !$e->isDot()) {
                $exe = $e->getPathname() . '\program\soffice.exe';
                if (file_exists($exe)) {
                    $sofficePaths[] = $exe;
                }
            }
        }
    }

    // Di Windows is_executable() sering false untuk .exe, jadi cukup file_exists
    $soffice = null;
    foreach ($sofficePaths as $p) {
        $norm = str_replace(['/', '\\'], DIRECTORY_SEPARATOR, $p);
        if (file_exists($norm)) {
            if (PHP_OS_FAMILY !== 'Windows' && !is_executable($norm)) {
                continue;
            }
            $soffice = $norm;
            break;
        }
    }
    if ($soffice === null) {
        return null;
    }

    $pathDocxReal = realpath($pathDocx);
    $dirReal = realpath($dir);
    if ($pathDocxReal === false || $dirReal === false) {
        return null;
    }

    if (PHP_OS_FAMILY === 'Windows') {
        // Argumen sebagai array agar path dengan spasi aman
        $descriptorspec = [0 => ['pipe', 'r'], 1 => ['pipe', 'w'], 2 => ['pipe', 'w']];
        $cmdArray = [$soffice, '--headless', '--convert-to', 'pdf', '--outdir', $dirReal, $pathDocxReal];
        $proc = @proc_open(
            $cmdArray,
            $descriptorspec,
            $pipes,
            $dirReal,
            null,
            ['bypass_shell' => true]
        );
        if (!is_resource($proc)) {
            return null;
        }
        fclose($pipes[0]);
        stream_get_contents($pipes[1]);
        stream_get_contents($pipes[2]);
        fclose($pipes[1]);
        fclose($pipes[2]);
        $code = proc_close($proc);
    } else {
        $cmd = sprintf(
            '%s --headless --convert-to pdf --outdir %s %s 2>/dev/null',
            escapeshellarg($soffice),
            escapeshellarg($dirReal),
            escapeshellarg($pathDocxReal)
        );
        exec($cmd, $out, $code);
    }

    // LibreOffice kadang butuh waktu menulis file; tunggu sampai 10 detik
    $maxWait = 10;
    for ($waited = 0; $waited < $maxWait; $waited++) {
        if (file_exists($pathPdf) && filesize($pathPdf) >= 100) {
            return $pathPdf;
        }
        sleep(1);
    }
    return null;
}

if (!isset($_GET['id'])) {
    die("ID surat tugas tidak ditemukan");
}

try {
    // Ambil data surat tugas
    $stmt = $pdo->prepare("SELECT st.*, GROUP_CONCAT(p.nama ORDER BY pt.urutan SEPARATOR '|\n') as pegawai_names,
                          GROUP_CONCAT(p.nip ORDER BY pt.urutan SEPARATOR '|\n') as pegawai_nip,
                          GROUP_CONCAT(p.pangkat ORDER BY pt.urutan SEPARATOR '|\n') as pegawai_pangkat,
                          GROUP_CONCAT(p.jabatan ORDER BY pt.urutan SEPARATOR '|\n') as pegawai_jabatan
                          FROM surat_tugas st 
                          LEFT JOIN pegawai_tugas pt ON st.id = pt.id_surat_tugas 
                          LEFT JOIN pegawai p ON pt.nip = p.nip 
                          WHERE st.id = ?
                          GROUP BY st.id");
    $stmt->execute([$_GET['id']]);
    $surat = $stmt->fetch();

    if (!$surat) {
        die("Surat tugas tidak ditemukan");
    }

    // Ambil data penanda tangan jika ada
    $penandatangan = null;
    if (!empty($surat['id_penandatangan'])) {
        $stmt_pt = $pdo->prepare("SELECT * FROM penandatangan WHERE id = ?");
        $stmt_pt->execute([$surat['id_penandatangan']]);
        $penandatangan = $stmt_pt->fetch();
    }
    
    // Jika tidak ada penandatangan dipilih, coba ambil default
    if (!$penandatangan) {
        try {
            $stmt_pt = $pdo->prepare("SELECT * FROM penandatangan WHERE is_default = 1 AND aktif = 1 LIMIT 1");
            $stmt_pt->execute();
            $penandatangan = $stmt_pt->fetch();
        } catch (PDOException $e) {
            // Abaikan jika tabel belum ada
        }
    }

    // Pisahkan data pegawai
    $nama_pegawai = !empty($surat['pegawai_names']) ? explode("|\n", $surat['pegawai_names']) : [];
    $nip_pegawai = !empty($surat['pegawai_nip']) ? explode("|\n", $surat['pegawai_nip']) : [];
    $pangkat_pegawai = !empty($surat['pegawai_pangkat']) ? explode("|\n", $surat['pegawai_pangkat']) : [];
    $jabatan_pegawai = !empty($surat['pegawai_jabatan']) ? explode("|\n", $surat['pegawai_jabatan']) : [];
    $jumlah_pegawai = count($nama_pegawai);

    // Validasi jumlah pegawai
    if ($jumlah_pegawai == 0) {
        die("Surat tugas tidak memiliki pegawai yang ditugaskan.");
    }

    // Tentukan kategori template berdasarkan jumlah pegawai
    // Logika: 1 pegawai -> 1_pegawai, 2 pegawai -> 2_pegawai, 3 pegawai -> 3_pegawai, 4+ -> banyak_pegawai
    $kategori_template = 'banyak_pegawai'; // default untuk 4+ pegawai
    $kategori_label = 'Banyak Pegawai (4+ pegawai)';
    
    if ($jumlah_pegawai == 1) {
        $kategori_template = '1_pegawai';
        $kategori_label = '1 Pegawai';
    } elseif ($jumlah_pegawai == 2) {
        $kategori_template = '2_pegawai';
        $kategori_label = '2 Pegawai';
    } elseif ($jumlah_pegawai == 3) {
        $kategori_template = '3_pegawai';
        $kategori_label = '3 Pegawai';
    }
    // else: tetap menggunakan banyak_pegawai untuk 4+ pegawai

    // Tentukan tipe surat (umum/penyuluh)
    $tipe_surat = isset($surat['tipe_surat']) ? $surat['tipe_surat'] : 'umum';
    $tipe_label = ($tipe_surat === 'penyuluh') ? 'Penyuluh' : 'Umum';

    // Cari template default dari database berdasarkan kategori
    $template_path = null;
    $template_found = false;
    $template_from_db = false;
    $template_nama = null;
    $template_kategori_terpilih = null;
    
    try {
        // Cek apakah kolom kategori_template ada, jika belum tambahkan
        try {
            $pdo->query("SELECT kategori_template FROM templates LIMIT 1");
        } catch (PDOException $e) {
            // Kolom belum ada, tambahkan
            try {
                $pdo->exec("ALTER TABLE templates ADD COLUMN kategori_template ENUM('1_pegawai', '2_pegawai', '3_pegawai', 'banyak_pegawai') NOT NULL DEFAULT 'banyak_pegawai' AFTER nama");
            } catch (PDOException $e2) {
                // Ignore jika sudah ada atau error lain
            }
        }

        // Cek apakah kolom tipe_surat ada, jika belum tambahkan
        try {
            $pdo->query("SELECT tipe_surat FROM templates LIMIT 1");
        } catch (PDOException $e) {
            try {
                $pdo->exec("ALTER TABLE templates ADD COLUMN tipe_surat VARCHAR(20) NOT NULL DEFAULT 'umum' AFTER kategori_template");
            } catch (PDOException $e2) {}
        }
        
        // Debug: Log informasi pencarian template
        error_log("=== DEBUG TEMPLATE SELECTION ===");
        error_log("Jumlah Pegawai: " . $jumlah_pegawai);
        error_log("Kategori Template Dicari: " . $kategori_template);
        error_log("Tipe Surat: " . $tipe_surat);
        
        // Debug: Cek semua template yang ada untuk kategori dan tipe ini
        $stmt_debug = $pdo->prepare("SELECT id, nama, kategori_template, tipe_surat, nama_file, is_default FROM templates WHERE kategori_template = ? AND tipe_surat = ?");
        $stmt_debug->execute([$kategori_template, $tipe_surat]);
        $templates_debug = $stmt_debug->fetchAll();
        error_log("Template ditemukan untuk kategori '{$kategori_template}' tipe '{$tipe_surat}': " . count($templates_debug));
        foreach ($templates_debug as $tmpl) {
            error_log("  - ID: {$tmpl['id']}, Nama: {$tmpl['nama']}, File: {$tmpl['nama_file']}, Default: {$tmpl['is_default']}");
        }
        
        // Cari template default untuk kategori dan tipe yang sesuai
        $stmt_template = $pdo->prepare("SELECT id, nama_file, kategori_template, tipe_surat, nama FROM templates WHERE kategori_template = ? AND tipe_surat = ? AND is_default = 1 LIMIT 1");
        $stmt_template->execute([$kategori_template, $tipe_surat]);
        $template_default = $stmt_template->fetch();
        
        if ($template_default) {
            error_log("Template default ditemukan: ID={$template_default['id']}, Nama={$template_default['nama']}, File={$template_default['nama_file']}, Kategori={$template_default['kategori_template']}");
        } else {
            error_log("Template default TIDAK ditemukan untuk kategori '{$kategori_template}'");
        }
        
        if ($template_default && !empty($template_default['nama_file'])) {
            $template_path = __DIR__ . '/../assets/templates/' . $template_default['nama_file'];
            
            // Cek file exists
            if (!file_exists($template_path)) {
                // File tidak ada, skip template ini dan cari yang lain
                $template_path = null;
            } elseif (!is_readable($template_path)) {
                // File tidak bisa dibaca, skip
                $template_path = null;
            } elseif (filesize($template_path) < 1000) {
                // File terlalu kecil, skip
                $template_path = null;
            } else {
                // Cek format file harus .docx
                $file_info = pathinfo($template_path);
                if (strtolower($file_info['extension']) !== 'docx') {
                    error_log("ERROR: Format file tidak valid. File harus .docx, ditemukan: .{$file_info['extension']}");
                    $template_path = null;
                } else {
                    // Validasi file dengan mencoba membuka sebagai ZIP
                    $zip = new ZipArchive();
                    $zip_result = $zip->open($template_path);
                    if ($zip_result === TRUE) {
                        $zip->close();
                        // Template valid dan sesuai kategori
                        if ($template_default['kategori_template'] == $kategori_template) {
                            $template_found = true;
                            $template_from_db = true;
                            $template_nama = $template_default['nama'];
                            $template_kategori_terpilih = $template_default['kategori_template'];
                            error_log("SUKSES: Template default digunakan - {$template_nama} (kategori: {$template_kategori_terpilih})");
                        } else {
                            // Kategori tidak sesuai, skip
                            error_log("ERROR: Kategori template tidak sesuai. Ditemukan: {$template_default['kategori_template']}, Dibutuhkan: {$kategori_template}");
                            $template_path = null;
                        }
                    } else {
                        // File corrupt atau tidak valid ZIP
                        error_log("ERROR: File template corrupt atau tidak valid ZIP: {$template_path}");
                        $template_path = null;
                    }
                }
            }
        }
        
        // Jika tidak ada default yang valid, ambil template pertama dari kategori dan tipe tersebut
        if (!$template_found) {
            $stmt_template = $pdo->prepare("SELECT id, nama_file, kategori_template, tipe_surat, nama FROM templates WHERE kategori_template = ? AND tipe_surat = ? ORDER BY created_at DESC LIMIT 1");
            $stmt_template->execute([$kategori_template, $tipe_surat]);
            $template_fallback = $stmt_template->fetch();
            
            if ($template_fallback && !empty($template_fallback['nama_file'])) {
                $template_path = __DIR__ . '/../assets/templates/' . $template_fallback['nama_file'];
                
                // Cek file exists dan valid
                if (file_exists($template_path) && is_readable($template_path) && filesize($template_path) > 1000) {
                    // Cek format file harus .docx
                    $file_info = pathinfo($template_path);
                    if (strtolower($file_info['extension']) !== 'docx') {
                        error_log("ERROR: Format file fallback tidak valid. File harus .docx, ditemukan: .{$file_info['extension']}");
                    } else {
                        // Validasi file dengan mencoba membuka sebagai ZIP
                        $zip = new ZipArchive();
                        $zip_result = $zip->open($template_path);
                        if ($zip_result === TRUE) {
                            $zip->close();
                            // Validasi kategori
                            if ($template_fallback['kategori_template'] == $kategori_template) {
                                $template_found = true;
                                $template_from_db = true;
                                $template_nama = $template_fallback['nama'];
                                $template_kategori_terpilih = $template_fallback['kategori_template'];
                                error_log("SUKSES: Template fallback digunakan - {$template_nama} (kategori: {$template_kategori_terpilih})");
                            } else {
                                error_log("ERROR: Kategori template fallback tidak sesuai. Ditemukan: {$template_fallback['kategori_template']}, Dibutuhkan: {$kategori_template}");
                            }
                        } else {
                            error_log("ERROR: File template fallback corrupt atau tidak valid ZIP: {$template_path}");
                        }
                    }
                }
            }
        }
        
        // FALLBACK: Jika template kategori spesifik tidak ditemukan, coba gunakan banyak_pegawai dengan tipe yang sama
        if (!$template_found && $kategori_template !== 'banyak_pegawai') {
            error_log("INFO: Template untuk kategori '{$kategori_template}' tipe '{$tipe_surat}' tidak ditemukan, mencoba fallback ke 'banyak_pegawai' tipe '{$tipe_surat}'");
            
            $stmt_fallback = $pdo->prepare("SELECT id, nama_file, kategori_template, tipe_surat, nama FROM templates WHERE kategori_template = 'banyak_pegawai' AND tipe_surat = ? AND is_default = 1 LIMIT 1");
            $stmt_fallback->execute([$tipe_surat]);
            $template_fallback_universal = $stmt_fallback->fetch();
            
            if (!$template_fallback_universal) {
                // Jika tidak ada default, ambil yang pertama dengan tipe yang sama
                $stmt_fallback = $pdo->prepare("SELECT id, nama_file, kategori_template, tipe_surat, nama FROM templates WHERE kategori_template = 'banyak_pegawai' AND tipe_surat = ? ORDER BY created_at DESC LIMIT 1");
                $stmt_fallback->execute([$tipe_surat]);
                $template_fallback_universal = $stmt_fallback->fetch();
            }
            
            if ($template_fallback_universal && !empty($template_fallback_universal['nama_file'])) {
                $template_path = __DIR__ . '/../assets/templates/' . $template_fallback_universal['nama_file'];
                
                if (file_exists($template_path) && is_readable($template_path) && filesize($template_path) > 1000) {
                    $file_info = pathinfo($template_path);
                    if (strtolower($file_info['extension']) === 'docx') {
                        $zip = new ZipArchive();
                        if ($zip->open($template_path) === TRUE) {
                            $zip->close();
                            $template_found = true;
                            $template_from_db = true;
                            $template_nama = $template_fallback_universal['nama'] . " (Fallback)";
                            $template_kategori_terpilih = $template_fallback_universal['kategori_template'];
                            error_log("SUKSES: Menggunakan template fallback 'banyak_pegawai' tipe '{$tipe_surat}' - {$template_nama}");
                        }
                    }
                }
            }
        }

        // FALLBACK 2: Jika masih tidak ditemukan dan tipe_surat bukan umum, coba dengan tipe umum
        if (!$template_found && $tipe_surat !== 'umum') {
            error_log("INFO: Template untuk tipe '{$tipe_surat}' tidak ditemukan, mencoba fallback ke tipe 'umum' kategori '{$kategori_template}'");
            
            $stmt_fallback2 = $pdo->prepare("SELECT id, nama_file, kategori_template, tipe_surat, nama FROM templates WHERE kategori_template = ? AND tipe_surat = 'umum' AND is_default = 1 LIMIT 1");
            $stmt_fallback2->execute([$kategori_template]);
            $template_fallback2 = $stmt_fallback2->fetch();
            
            if (!$template_fallback2) {
                $stmt_fallback2 = $pdo->prepare("SELECT id, nama_file, kategori_template, tipe_surat, nama FROM templates WHERE kategori_template = ? AND tipe_surat = 'umum' ORDER BY created_at DESC LIMIT 1");
                $stmt_fallback2->execute([$kategori_template]);
                $template_fallback2 = $stmt_fallback2->fetch();
            }

            if (!$template_fallback2 && $kategori_template !== 'banyak_pegawai') {
                $stmt_fallback2 = $pdo->prepare("SELECT id, nama_file, kategori_template, tipe_surat, nama FROM templates WHERE kategori_template = 'banyak_pegawai' AND tipe_surat = 'umum' AND is_default = 1 LIMIT 1");
                $stmt_fallback2->execute();
                $template_fallback2 = $stmt_fallback2->fetch();
            }

            if (!$template_fallback2 && $kategori_template !== 'banyak_pegawai') {
                $stmt_fallback2 = $pdo->prepare("SELECT id, nama_file, kategori_template, tipe_surat, nama FROM templates WHERE kategori_template = 'banyak_pegawai' AND tipe_surat = 'umum' ORDER BY created_at DESC LIMIT 1");
                $stmt_fallback2->execute();
                $template_fallback2 = $stmt_fallback2->fetch();
            }
            
            if ($template_fallback2 && !empty($template_fallback2['nama_file'])) {
                $template_path = __DIR__ . '/../assets/templates/' . $template_fallback2['nama_file'];
                
                if (file_exists($template_path) && is_readable($template_path) && filesize($template_path) > 1000) {
                    $file_info = pathinfo($template_path);
                    if (strtolower($file_info['extension']) === 'docx') {
                        $zip = new ZipArchive();
                        if ($zip->open($template_path) === TRUE) {
                            $zip->close();
                            $template_found = true;
                            $template_from_db = true;
                            $template_nama = $template_fallback2['nama'] . " (Fallback Umum)";
                            $template_kategori_terpilih = $template_fallback2['kategori_template'];
                            error_log("SUKSES: Menggunakan template fallback tipe 'umum' - {$template_nama}");
                        }
                    }
                }
            }
        }
    } catch (PDOException $e) {
        // Jika tabel belum ada atau kolom belum ada, lanjutkan pengecekan
        // Log error untuk debugging (bisa dihapus di production)
        error_log("Error mencari template: " . $e->getMessage());
    }
    
    // Jika template tidak ditemukan, tampilkan error sederhana
    if (!$template_found) {
        die('Error: Template untuk kategori "' . $kategori_label . '" tipe "' . $tipe_label . '" tidak ditemukan. Silakan upload template yang sesuai.');
    }

    // Periksa apakah file template ada (double check)
    if (!file_exists($template_path)) {
        include '../includes/header.php';
        
        echo '<div class="space-y-6">';
        echo '<div class="bg-red-50 border-l-4 border-red-500 p-6 rounded-lg">';
        echo '<div class="flex items-start">';
        echo '<div class="flex-shrink-0">';
        echo '<i class="bx bx-error-circle text-red-500 text-3xl"></i>';
        echo '</div>';
        echo '<div class="ml-4 flex-1">';
        echo '<h3 class="text-lg font-semibold text-red-800 mb-2">File Template Tidak Ditemukan</h3>';
        echo '<p class="text-red-700 mb-4">File template untuk kategori <strong>"' . htmlspecialchars($kategori_label) . '"</strong> tidak ditemukan di server.</p>';
        echo '<div class="flex gap-3">';
        echo '<a href="template.php" class="inline-flex items-center gap-2 px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg transition-colors">';
        echo '<i class="bx bx-file-blank"></i>';
        echo '<span>Buka Halaman Template Word</span>';
        echo '</a>';
        echo '<a href="surat-tugas.php" class="inline-flex items-center gap-2 px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg transition-colors">';
        echo '<i class="bx bx-arrow-back"></i>';
        echo '<span>Kembali ke Daftar Surat Tugas</span>';
        echo '</a>';
        echo '</div>';
        echo '</div>';
        echo '</div>';
        echo '</div>';
        echo '</div>';
        
        include '../includes/footer.php';
        exit;
    }

    // Validasi file template sebelum memproses
    if (!file_exists($template_path)) {
        include '../includes/header.php';
        echo '<div class="space-y-6">';
        echo '<div class="bg-red-50 border-l-4 border-red-500 p-6 rounded-lg">';
        echo '<div class="flex items-start">';
        echo '<div class="flex-shrink-0"><i class="bx bx-error-circle text-red-500 text-3xl"></i></div>';
        echo '<div class="ml-4 flex-1">';
        echo '<h3 class="text-lg font-semibold text-red-800 mb-2">File Template Tidak Ditemukan</h3>';
        echo '<p class="text-red-700 mb-4">File template tidak ditemukan di: ' . htmlspecialchars($template_path) . '</p>';
        echo '<a href="surat-tugas.php" class="inline-flex items-center gap-2 px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg transition-colors">';
        echo '<i class="bx bx-arrow-back"></i><span>Kembali ke Daftar Surat Tugas</span></a>';
        echo '</div></div></div></div>';
        include '../includes/footer.php';
        exit;
    }

    // Cek apakah file bisa dibaca
    if (!is_readable($template_path)) {
        include '../includes/header.php';
        echo '<div class="space-y-6">';
        echo '<div class="bg-red-50 border-l-4 border-red-500 p-6 rounded-lg">';
        echo '<div class="flex items-start">';
        echo '<div class="flex-shrink-0"><i class="bx bx-error-circle text-red-500 text-3xl"></i></div>';
        echo '<div class="ml-4 flex-1">';
        echo '<h3 class="text-lg font-semibold text-red-800 mb-2">File Template Tidak Dapat Dibaca</h3>';
        echo '<p class="text-red-700 mb-4">File template tidak memiliki permission untuk dibaca. Silakan hubungi administrator.</p>';
        echo '<a href="surat-tugas.php" class="inline-flex items-center gap-2 px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg transition-colors">';
        echo '<i class="bx bx-arrow-back"></i><span>Kembali ke Daftar Surat Tugas</span></a>';
        echo '</div></div></div></div>';
        include '../includes/footer.php';
        exit;
    }

    // Cek apakah file adalah valid .docx (minimal cek extension dan ukuran)
    $file_info = pathinfo($template_path);
    if (strtolower($file_info['extension']) !== 'docx') {
        include '../includes/header.php';
        echo '<div class="space-y-6">';
        echo '<div class="bg-red-50 border-l-4 border-red-500 p-6 rounded-lg">';
        echo '<div class="flex items-start">';
        echo '<div class="flex-shrink-0"><i class="bx bx-error-circle text-red-500 text-3xl"></i></div>';
        echo '<div class="ml-4 flex-1">';
        echo '<h3 class="text-lg font-semibold text-red-800 mb-2">Format File Tidak Valid</h3>';
        echo '<p class="text-red-700 mb-4">File template harus berformat .docx. File yang ditemukan: ' . htmlspecialchars($file_info['extension']) . '</p>';
        echo '<a href="surat-tugas.php" class="inline-flex items-center gap-2 px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg transition-colors">';
        echo '<i class="bx bx-arrow-back"></i><span>Kembali ke Daftar Surat Tugas</span></a>';
        echo '</div></div></div></div>';
        include '../includes/footer.php';
        exit;
    }

    // Cek apakah file kosong atau terlalu kecil (kemungkinan corrupt)
    if (filesize($template_path) < 1000) {
        include '../includes/header.php';
        echo '<div class="space-y-6">';
        echo '<div class="bg-red-50 border-l-4 border-red-500 p-6 rounded-lg">';
        echo '<div class="flex items-start">';
        echo '<div class="flex-shrink-0"><i class="bx bx-error-circle text-red-500 text-3xl"></i></div>';
        echo '<div class="ml-4 flex-1">';
        echo '<h3 class="text-lg font-semibold text-red-800 mb-2">File Template Corrupt</h3>';
        echo '<p class="text-red-700 mb-4">File template terlalu kecil atau mungkin corrupt. Silakan upload ulang template untuk kategori "' . htmlspecialchars($kategori_label) . '".</p>';
        echo '<div class="flex gap-3">';
        echo '<a href="template.php" class="inline-flex items-center gap-2 px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg transition-colors">';
        echo '<i class="bx bx-file-blank"></i><span>Buka Halaman Template</span></a>';
        echo '<a href="surat-tugas.php" class="inline-flex items-center gap-2 px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg transition-colors">';
        echo '<i class="bx bx-arrow-back"></i><span>Kembali ke Daftar Surat Tugas</span></a>';
        echo '</div></div></div></div></div>';
        include '../includes/footer.php';
        exit;
    }

    // Validasi akhir: Pastikan template yang dipilih sesuai kategori (jika dari database)
    if ($template_from_db && $template_kategori_terpilih !== null) {
        if ($template_kategori_terpilih !== $kategori_template) {
            // Template tidak sesuai kategori, cari lagi atau gunakan fallback
            include '../includes/header.php';
            echo '<div class="space-y-6">';
            echo '<div class="bg-yellow-50 border-l-4 border-yellow-500 p-6 rounded-lg">';
            echo '<div class="flex items-start">';
            echo '<div class="flex-shrink-0"><i class="bx bx-error-circle text-yellow-500 text-3xl"></i></div>';
            echo '<div class="ml-4 flex-1">';
            echo '<h3 class="text-lg font-semibold text-yellow-800 mb-2">Template Tidak Sesuai Kategori</h3>';
            echo '<p class="text-yellow-700 mb-4">';
            echo 'Template yang dipilih memiliki kategori <strong>"' . htmlspecialchars($template_kategori_terpilih) . '"</strong>, ';
            echo 'tetapi surat tugas ini memerlukan kategori <strong>"' . htmlspecialchars($kategori_label) . '"</strong> ';
            echo 'karena memiliki <strong>' . $jumlah_pegawai . ' pegawai</strong>.';
            echo '</p>';
            echo '<p class="text-sm text-yellow-600 mb-4">Silakan upload template yang sesuai untuk kategori "' . htmlspecialchars($kategori_label) . '".</p>';
            echo '<div class="flex gap-3">';
            echo '<a href="template.php" class="inline-flex items-center gap-2 px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg transition-colors">';
            echo '<i class="bx bx-file-blank"></i><span>Buka Halaman Template</span></a>';
            echo '<a href="surat-tugas.php" class="inline-flex items-center gap-2 px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg transition-colors">';
            echo '<i class="bx bx-arrow-back"></i><span>Kembali ke Daftar Surat Tugas</span></a>';
            echo '</div></div></div></div></div>';
            include '../includes/footer.php';
            exit;
        }
    }

    // Format tanggal dan data untuk preview/download (digunakan juga di setValue nanti)
    $bulan = array(
        1 => 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
        'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
    );
    $tanggal_surat = date('d', strtotime($surat['tanggal_surat'])) . ' ' .
        $bulan[date('n', strtotime($surat['tanggal_surat']))] . ' ' .
        date('Y', strtotime($surat['tanggal_surat']));
    $tanggal_mulai = date('d', strtotime($surat['tanggal_mulai'])) . ' ' .
        $bulan[date('n', strtotime($surat['tanggal_mulai']))] . ' ' .
        date('Y', strtotime($surat['tanggal_mulai']));
    $tanggal_selesai = date('d', strtotime($surat['tanggal_selesai'])) . ' ' .
        $bulan[date('n', strtotime($surat['tanggal_selesai']))] . ' ' .
        date('Y', strtotime($surat['tanggal_selesai']));
    $datetime1 = new DateTime($surat['tanggal_mulai']);
    $datetime2 = new DateTime($surat['tanggal_selesai']);
    $interval = $datetime1->diff($datetime2);
    $durasi = $interval->days + 1;
    if (!function_exists('angkaTerbilang')) {
        function angkaTerbilang($n) {
            $terbilang = [
                1 => 'satu', 2 => 'dua', 3 => 'tiga', 4 => 'empat', 5 => 'lima',
                6 => 'enam', 7 => 'tujuh', 8 => 'delapan', 9 => 'sembilan', 10 => 'sepuluh',
                11 => 'sebelas', 12 => 'dua belas', 13 => 'tiga belas', 14 => 'empat belas',
                15 => 'lima belas', 16 => 'enam belas', 17 => 'tujuh belas', 18 => 'delapan belas',
                19 => 'sembilan belas', 20 => 'dua puluh', 30 => 'tiga puluh'
            ];
            if (isset($terbilang[$n])) return $terbilang[$n];
            if ($n < 100) {
                $puluhan = floor($n / 10) * 10;
                $sisa = $n % 10;
                return $terbilang[$puluhan] . ' ' . $terbilang[$sisa];
            }
            return $n;
        }
    }
    if ($surat['tanggal_mulai'] == $surat['tanggal_selesai']) {
        $waktu_pelaksanaan = "pada tanggal $tanggal_mulai selama 1 (satu) hari";
    } else {
        $waktu_pelaksanaan = "dari tanggal $tanggal_mulai sampai dengan $tanggal_selesai selama " .
            $durasi . " (" . angkaTerbilang($durasi) . ") hari";
    }
    $pegawaiText = "";
    for ($i = 0; $i < $jumlah_pegawai; $i++) {
        $pegawaiText .= ($i + 1) . ". Nama    : " . $nama_pegawai[$i] . "\n";
        $pegawaiText .= "   NIP     : " . $nip_pegawai[$i] . "\n";
        $pegawaiText .= "   Pangkat : " . $pangkat_pegawai[$i] . "\n";
        $pegawaiText .= "   Jabatan : " . $jabatan_pegawai[$i] . "\n\n";
    }

    // Load template dengan error handling yang lebih spesifik
    try {
        // Pastikan ZipArchive extension tersedia
        if (!class_exists('ZipArchive')) {
            throw new Exception('ZipArchive extension tidak tersedia. Silakan aktifkan extension zip di PHP.');
        }
        
        $templateProcessor = new TemplateProcessor($template_path);
        
        // Verifikasi akhir: Pastikan template yang digunakan sesuai kategori
        // Jika template dari database, pastikan kategori sesuai
        if ($template_from_db && $template_kategori_terpilih !== null) {
            if ($template_kategori_terpilih !== $kategori_template) {
                // Template tidak sesuai kategori - ini seharusnya tidak terjadi karena sudah divalidasi sebelumnya
                include '../includes/header.php';
                echo '<div class="space-y-6">';
                echo '<div class="bg-red-50 border-l-4 border-red-500 p-6 rounded-lg">';
                echo '<div class="flex items-start">';
                echo '<div class="flex-shrink-0"><i class="bx bx-error-circle text-red-500 text-3xl"></i></div>';
                echo '<div class="ml-4 flex-1">';
                echo '<h3 class="text-lg font-semibold text-red-800 mb-2">Template Tidak Sesuai Kategori</h3>';
                echo '<p class="text-red-700 mb-4">';
                echo 'Template yang dipilih memiliki kategori <strong>"' . htmlspecialchars($template_kategori_terpilih) . '"</strong>, ';
                echo 'tetapi surat tugas ini memerlukan kategori <strong>"' . htmlspecialchars($kategori_label) . '"</strong> ';
                echo 'karena memiliki <strong>' . $jumlah_pegawai . ' pegawai</strong>.';
                echo '</p>';
                echo '<p class="text-sm text-red-600 mb-4">Template: ' . htmlspecialchars($template_nama ?: basename($template_path)) . '</p>';
                echo '<div class="flex gap-3">';
                echo '<a href="template.php" class="inline-flex items-center gap-2 px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg transition-colors">';
                echo '<i class="bx bx-file-blank"></i><span>Buka Halaman Template</span></a>';
                echo '<a href="surat-tugas.php" class="inline-flex items-center gap-2 px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg transition-colors">';
                echo '<i class="bx bx-arrow-back"></i><span>Kembali ke Daftar Surat Tugas</span></a>';
                echo '</div></div></div></div></div>';
                include '../includes/footer.php';
                exit;
            }
        }
        
        // Debug info (untuk verifikasi - bisa dihapus di production)
        // Informasi ini membantu memastikan template yang digunakan sesuai
        // error_log("Generate Surat - Jumlah Pegawai: {$jumlah_pegawai}, Kategori Dicari: {$kategori_template}, Template Terpilih: " . ($template_nama ?: basename($template_path)) . ", Kategori Template: " . ($template_kategori_terpilih ?: 'fallback'));
    } catch (ValueError $e) {
        // Error khusus untuk ZipArchive invalid
        include '../includes/header.php';
        echo '<div class="space-y-6">';
        echo '<div class="bg-red-50 border-l-4 border-red-500 p-6 rounded-lg">';
        echo '<div class="flex items-start">';
        echo '<div class="flex-shrink-0"><i class="bx bx-error-circle text-red-500 text-3xl"></i></div>';
        echo '<div class="ml-4 flex-1">';
        echo '<h3 class="text-lg font-semibold text-red-800 mb-2">File Template Corrupt atau Tidak Valid</h3>';
        echo '<p class="text-red-700 mb-2">File template untuk kategori "' . htmlspecialchars($kategori_label) . '" tidak dapat dibuka karena corrupt atau format tidak valid.</p>';
        echo '<div class="bg-white rounded-lg p-3 border border-red-200 mb-4">';
        echo '<p class="text-sm text-red-800 font-mono">' . htmlspecialchars($e->getMessage()) . '</p>';
        echo '</div>';
        echo '<p class="text-sm text-red-600 mb-4 font-medium">Solusi:</p>';
        echo '<ol class="list-decimal list-inside text-sm text-red-700 mb-4 space-y-1">';
        echo '<li>Hapus template yang corrupt dari halaman Template Word</li>';
        echo '<li>Upload ulang template yang valid untuk kategori "' . htmlspecialchars($kategori_label) . '"</li>';
        echo '<li>Pastikan file template adalah file .docx yang valid dan tidak corrupt</li>';
        echo '</ol>';
        echo '<div class="flex gap-3">';
        echo '<a href="template.php" class="inline-flex items-center gap-2 px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg transition-colors">';
        echo '<i class="bx bx-file-blank"></i><span>Buka Halaman Template</span></a>';
        echo '<a href="surat-tugas.php" class="inline-flex items-center gap-2 px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg transition-colors">';
        echo '<i class="bx bx-arrow-back"></i><span>Kembali ke Daftar Surat Tugas</span></a>';
        echo '</div></div></div></div></div>';
        include '../includes/footer.php';
        exit;
    } catch (Exception $e) {
        include '../includes/header.php';
        echo '<div class="space-y-6">';
        echo '<div class="bg-red-50 border-l-4 border-red-500 p-6 rounded-lg">';
        echo '<div class="flex items-start">';
        echo '<div class="flex-shrink-0"><i class="bx bx-error-circle text-red-500 text-3xl"></i></div>';
        echo '<div class="ml-4 flex-1">';
        echo '<h3 class="text-lg font-semibold text-red-800 mb-2">Error Memproses Template</h3>';
        echo '<p class="text-red-700 mb-2">Terjadi kesalahan saat memproses template untuk kategori "' . htmlspecialchars($kategori_label) . '":</p>';
        echo '<div class="bg-white rounded-lg p-3 border border-red-200 mb-4">';
        echo '<p class="text-sm text-red-800 font-mono">' . htmlspecialchars($e->getMessage()) . '</p>';
        echo '</div>';
        echo '<p class="text-sm text-red-600 mb-4">Kemungkinan penyebab:</p>';
        echo '<ul class="list-disc list-inside text-sm text-red-700 mb-4 space-y-1">';
        echo '<li>File template corrupt atau tidak valid</li>';
        echo '<li>Format file tidak sesuai standar .docx</li>';
        echo '<li>File sedang digunakan oleh aplikasi lain</li>';
        echo '</ul>';
        echo '<div class="flex gap-3">';
        echo '<a href="template.php" class="inline-flex items-center gap-2 px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg transition-colors">';
        echo '<i class="bx bx-file-blank"></i><span>Buka Halaman Template</span></a>';
        echo '<a href="surat-tugas.php" class="inline-flex items-center gap-2 px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg transition-colors">';
        echo '<i class="bx bx-arrow-back"></i><span>Kembali ke Daftar Surat Tugas</span></a>';
        echo '</div></div></div></div></div>';
        include '../includes/footer.php';
        exit;
    }

    // Replace placeholder umum
    $templateProcessor->setValue('nomor_surat', $surat['nomor_surat']);
    $templateProcessor->setValue('tanggal_surat', $tanggal_surat);

    $dasarSurat = trim((string)($surat['dasar_surat'] ?? ''));
    if ($dasarSurat === '') {
        // Jika dasar surat kosong, kosongkan semua placeholder terkait item dasar surat.
        $templateProcessor->setValue('dasar_surat', '');
        $templateProcessor->setValue('n', '');
        $templateProcessor->setValue('dasar_surat_item', '');
    } else {
        // Jika ada dasar surat, tampilkan nomor + isi dasar surat.
        $templateProcessor->setValue('dasar_surat', $dasarSurat);
        $templateProcessor->setValue('n', '5.');
        $templateProcessor->setValue('dasar_surat_item', '5. ' . $dasarSurat);
    }

    $templateProcessor->setValue('untuk', $surat['untuk']);
    $templateProcessor->setValue('waktu_pelaksanaan', $waktu_pelaksanaan);

    // STRATEGI PENGISIAN DATA PEGAWAI (Hybrid Approach)
    
    // 1. Isi variabel individual (nama1, nip1, dst) sampai maksimal 10 pegawai
    for ($i = 0; $i < 10; $i++) {
        $index = $i + 1;
        if ($i < $jumlah_pegawai) {
            $templateProcessor->setValue('nama' . $index, $nama_pegawai[$i]);
            $templateProcessor->setValue('nip' . $index, $nip_pegawai[$i]);
            $templateProcessor->setValue('pangkat' . $index, $pangkat_pegawai[$i]);
            $templateProcessor->setValue('jabatan' . $index, $jabatan_pegawai[$i]);
        } else {
            // Kosongkan jika tidak ada data
            $templateProcessor->setValue('nama' . $index, '');
            $templateProcessor->setValue('nip' . $index, '');
            $templateProcessor->setValue('pangkat' . $index, '');
            $templateProcessor->setValue('jabatan' . $index, '');
        }
    }

    // 2. Blok teks daftar pegawai (variabel $pegawaiText sudah dihitung di atas)
    $templateProcessor->setValue('daftar_pegawai', $pegawaiText);

    // 3. Coba isi Tabel Dinamis (jika template menggunakan row cloning)
    try {
        $pegawaiData = [];
        for ($i = 0; $i < $jumlah_pegawai; $i++) {
            $pegawaiData[] = [
                'no' => ($i + 1),
                'nama' => $nama_pegawai[$i],
                'nip' => $nip_pegawai[$i],
                'pangkat' => $pangkat_pegawai[$i],
                'jabatan' => $jabatan_pegawai[$i]
            ];
        }
        // Asumsi nama row yang akan di-clone adalah 'no'
        $templateProcessor->cloneRowAndSetValues('no', $pegawaiData);
    } catch (Exception $e) {
        // Abaikan error jika row cloning gagal
    }

    // 4. Isi data penanda tangan
    if ($penandatangan) {
        $templateProcessor->setValue('penandatangan_nama', $penandatangan['nama']);
        $templateProcessor->setValue('penandatangan_nip', $penandatangan['nip']);
        $templateProcessor->setValue('penandatangan_jabatan', $penandatangan['jabatan']);
        $templateProcessor->setValue('penandatangan_pangkat', $penandatangan['pangkat'] ?? '');

        // Tentukan tipe penanda tangan (Kepala atau A.n)
        $is_kepala = isset($penandatangan['is_kepala']) ? (int)$penandatangan['is_kepala'] : 1;
        $jabatan_atasan = $penandatangan['jabatan_atasan'] ?? '';

        if ($is_kepala) {
            // Tipe Kepala: Jabatan langsung ditampilkan
            // Contoh: "Kepala Cabang Dinas Kehutanan Wilayah Bojonegoro"
            $templateProcessor->setValue('penandatangan_header_jabatan', $penandatangan['jabatan']);
            $templateProcessor->setValue('penandatangan_sub_jabatan', '');
            $templateProcessor->setValue('penandatangan_an', '');
            $templateProcessor->setValue('penandatangan_jabatan_atasan', '');
        } else {
            // Tipe A.n: "A.n [Jabatan Atasan]" lalu jabatan sendiri di bawah
            // Contoh: "A.n Kepala Cabang Dinas Kehutanan Wilayah Bojonegoro"
            //         "Kepala Sub Bagian Tata Usaha"
            $templateProcessor->setValue('penandatangan_header_jabatan', 'A.n ' . $jabatan_atasan);
            $templateProcessor->setValue('penandatangan_sub_jabatan', $penandatangan['jabatan']);
            $templateProcessor->setValue('penandatangan_an', 'A.n');
            $templateProcessor->setValue('penandatangan_jabatan_atasan', $jabatan_atasan);
        }

        // 5. Isi gambar tanda tangan
        if ($penandatangan && !empty($penandatangan['tanda_tangan'])) {
            $ttd_path = __DIR__ . '/../assets/img/tanda_tangan/' . $penandatangan['tanda_tangan'];
            if (file_exists($ttd_path) && is_readable($ttd_path)) {
                // Tentukan ukuran gambar berdasarkan jabatan penandatangan
                $jabatanPenandatangan = isset($penandatangan['jabatan']) ? trim($penandatangan['jabatan']) : '';
                if (strcasecmp($jabatanPenandatangan, 'Kepala Cabang Dinas Kehutanan Wilayah Bojonegoro') === 0) {
                    $imgWidth = 150;
                    $imgHeight = 110;
                } else {
                    $imgWidth = 120;
                    $imgHeight = 80;
                }

                try {
                    $templateProcessor->setImageValue('tanda_tangan', [
                        'path' => $ttd_path,
                        'width' => $imgWidth,
                        'height' => $imgHeight,
                        'ratio' => true,
                        'wrappingStyle' => 'inline'
                    ]);
                } catch (Exception $imgEx) {
                    // Jika gagal menyisipkan gambar, kosongkan placeholder
                    error_log("Gagal menyisipkan gambar tanda tangan: " . $imgEx->getMessage());
                    try {
                        $templateProcessor->setValue('tanda_tangan', '');
                    } catch (Exception $e2) {}
                }
            } else {
                try {
                    $templateProcessor->setValue('tanda_tangan', '');
                } catch (Exception $e2) {}
            }
        } else {
            try {
                $templateProcessor->setValue('tanda_tangan', '');
            } catch (Exception $e2) {}
        }
    } else {
        // Kosongkan placeholder jika tidak ada penandatangan
        $templateProcessor->setValue('penandatangan_nama', '.........................');
        $templateProcessor->setValue('penandatangan_nip', '.........................');
        $templateProcessor->setValue('penandatangan_jabatan', '.........................');
        $templateProcessor->setValue('penandatangan_pangkat', '');
        $templateProcessor->setValue('penandatangan_header_jabatan', '.........................');
        $templateProcessor->setValue('penandatangan_sub_jabatan', '');
        $templateProcessor->setValue('penandatangan_an', '');
        $templateProcessor->setValue('penandatangan_jabatan_atasan', '');
        try {
            $templateProcessor->setValue('tanda_tangan', '');
        } catch (Exception $e2) {}
    }

    $format = isset($_GET['format']) ? strtolower(trim($_GET['format'])) : '';
    $inlinePdf = isset($_GET['inline']) && $_GET['inline'] === '1';

    // Preview (tanpa format): tampilkan halaman dengan iframe PDF agar sama persis dengan file Word/PDF
    if ($format === '') {
        include '../includes/header.php';
        $sid = (int)$_GET['id'];
        $previewPdfUrl = 'generate-surat.php?id=' . $sid . '&format=pdf&inline=1';
        ?>
        <div class="space-y-6">
            <div class="flex flex-wrap items-center justify-between gap-4">
                <h2 class="text-xl font-semibold text-slate-800">Preview Surat Tugas</h2>
                <div class="flex flex-wrap gap-3">
                    <a href="generate-surat.php?id=<?= $sid ?>&format=docx" class="inline-flex items-center gap-2 px-4 py-2 bg-green-600 hover:bg-green-700 text-white rounded-lg transition-colors">
                        <i class="bx bx-file-blank"></i><span>Download Word (.docx)</span>
                    </a>
                    <a href="generate-surat.php?id=<?= $sid ?>&format=pdf" class="inline-flex items-center gap-2 px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded-lg transition-colors">
                        <i class="bx bxs-file-pdf"></i><span>Download PDF</span>
                    </a>
                    <a href="surat-tugas.php" class="inline-flex items-center gap-2 px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg transition-colors">
                        <i class="bx bx-arrow-back"></i><span>Kembali</span>
                    </a>
                </div>
            </div>
            <div class="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <p class="text-sm text-slate-500 px-4 py-2 border-b border-slate-200">Tampilan di bawah sama dengan file Word/PDF yang diunduh.</p>
                <div class="p-4">
                    <iframe src="<?= htmlspecialchars($previewPdfUrl) ?>" class="w-full border border-slate-200 rounded-lg bg-white" style="min-height: 80vh; height: 297mm;" title="Preview surat (PDF)"></iframe>
                </div>
            </div>
            <div class="flex justify-center py-4">
                <a href="surat-tugas.php?edit=<?= $sid ?>" class="inline-flex items-center gap-2 px-6 py-2.5 bg-amber-500 hover:bg-amber-600 text-white font-medium rounded-lg transition-colors shadow-sm">
                    <i class="bx bx-edit text-lg"></i>
                    <span>Edit Surat</span>
                </a>
            </div>
        </div>
        <?php
        include '../includes/footer.php';
        exit;
    }

    // Cari folder writable untuk file temp
// Buat folder tmp di root project jika belum ada
$rootTmp = dirname(__DIR__) . DIRECTORY_SEPARATOR . 'tmp';
if (!is_dir($rootTmp)) {
    @mkdir($rootTmp, 0755, true);
}

$tempCandidates = [
    $rootTmp,
    __DIR__,  // folder modules (sudah terbukti writable)
    'C:\\laragon\\tmp',
    sys_get_temp_dir(),
];
$tempDir = __DIR__; // default ke folder modules
foreach ($tempCandidates as $candidate) {
    if (is_dir($candidate) && is_writable($candidate)) {
        $tempDir = $candidate;
        break;
    }
}
    $tempDocx = $tempDir . DIRECTORY_SEPARATOR . 'spt_' . uniqid() . '.docx';
    $templateProcessor->saveAs($tempDocx);

    try {
        if ($format === 'pdf') {
            set_time_limit(120); // Word COM / LibreOffice bisa memakan waktu
            $tempPdf = null;
            $usedMethod = '';
            $hasComSupport = class_exists('COM');
            $comRetryCount = isset($_GET['com_retry']) ? (int)$_GET['com_retry'] : 0;

            try {
                // Prioritas 1: Microsoft Word COM (hasil 100% sama dengan Word)
                if ($hasComSupport) {
                    $tempPdf = konversi_docx_ke_pdf_word_com($tempDocx, 3);
                    if ($tempPdf !== null) {
                        $usedMethod = 'Word COM';
                    }
                }

                // Jika server mendukung COM tetapi Word COM belum berhasil dan com_retry < 3,
                // tampilkan halaman auto-reload agar browser mencoba kembali secara otomatis sampai Word COM merespons.
                if ($tempPdf === null && $hasComSupport && $comRetryCount < 3) {
                    @unlink($tempDocx);
                    $nextRetry = $comRetryCount + 1;
                    $protocol = (isset($_SERVER['HTTPS']) && $_SERVER['HTTPS'] === 'on') ? 'https' : 'http';
                    $currentUrl = $protocol . '://' . $_SERVER['HTTP_HOST'] . $_SERVER['REQUEST_URI'];
                    
                    $urlParts = parse_url($currentUrl);
                    $queryParams = [];
                    if (isset($urlParts['query'])) {
                        parse_str($urlParts['query'], $queryParams);
                    }
                    $queryParams['com_retry'] = $nextRetry;
                    $reloadUrl = $urlParts['path'] . '?' . http_build_query($queryParams);

                    header('Content-Type: text/html; charset=utf-8');
                    ?>
                    <!DOCTYPE html>
                    <html lang="id">
                    <head>
                        <meta charset="UTF-8">
                        <title>Menyiapkan PDF Surat Tugas...</title>
                        <script src="https://cdn.tailwindcss.com"></script>
                    </head>
                    <body class="bg-slate-100 flex items-center justify-center min-h-screen p-4">
                        <div class="bg-white p-8 rounded-2xl shadow-xl max-w-md w-full text-center space-y-4 border border-slate-200">
                            <div class="inline-block animate-spin rounded-full h-12 w-12 border-4 border-indigo-600 border-t-transparent"></div>
                            <h2 class="text-lg font-semibold text-slate-800">Menyiapkan Dokumen PDF</h2>
                            <p class="text-sm text-slate-600">Sedang memproses dokumen dengan Microsoft Word COM. Halaman akan dimuat ulang secara otomatis...</p>
                            <p class="text-xs text-slate-400">Percobaan <?= $nextRetry ?> dari 3</p>
                        </div>
                        <script>
                            setTimeout(function() {
                                window.location.href = <?= json_encode($reloadUrl) ?>;
                            }, 1200);
                        </script>
                    </body>
                    </html>
                    <?php
                    exit;
                }

                // Prioritas 2: LibreOffice (fallback jika Word COM tidak ada atau tidak dapat digunakan)
                if ($tempPdf === null) {
                    $tempPdf = konversi_docx_ke_pdf_libreoffice($tempDocx);
                    if ($tempPdf !== null) {
                        $usedMethod = 'LibreOffice';
                    }
                }

                // Prioritas 3: DomPDF (fallback terakhir hanya jika COM dan LibreOffice tidak ada/gagal)
                if ($tempPdf === null) {
                    $baseDir = dirname(__DIR__);
                    $dompdfPath = $baseDir . DIRECTORY_SEPARATOR . 'vendor' . DIRECTORY_SEPARATOR . 'dompdf' . DIRECTORY_SEPARATOR . 'dompdf';
                    if (is_dir($dompdfPath)) {
                        Settings::setPdfRendererName(Settings::PDF_RENDERER_DOMPDF);
                        Settings::setPdfRendererPath($dompdfPath);
                        $tempPdf = $tempDir . DIRECTORY_SEPARATOR . 'spt_' . uniqid() . '.pdf';
                        $oldErrorReporting = error_reporting(E_ALL & ~E_DEPRECATED);
                        try {
                            $phpWord = IOFactory::load($tempDocx, 'Word2007');
                            $pdfWriter = IOFactory::createWriter($phpWord, 'PDF');
                            $pdfWriter->save($tempPdf);
                            if (file_exists($tempPdf) && filesize($tempPdf) >= 100) {
                                $usedMethod = 'DomPDF';
                            } else {
                                $tempPdf = null;
                            }
                        } finally {
                            error_reporting($oldErrorReporting);
                        }
                    }
                }
                error_log("PDF dibuat menggunakan: " . $usedMethod);
                if ($tempPdf === null || !file_exists($tempPdf) || filesize($tempPdf) < 100) {
                    throw new Exception(
                        'Gagal membuat PDF. Cek: (1) Microsoft Word terinstall & COM aktif (php.ini: extension=php_com_dotnet), ' .
                        '(2) LibreOffice terinstall, atau (3) composer install untuk DomPDF. Jalankan debug_word_com.php untuk diagnosa.'
                    );
                }

                @unlink($tempDocx);
                header('Content-Type: application/pdf');
                if ($inlinePdf) {
                    header('Content-Disposition: inline; filename="Surat_Tugas_' . preg_replace('/[^a-zA-Z0-9\-_]/', '_', $surat['nomor_surat']) . '.pdf"');
                } else {
                    header('Content-Disposition: attachment; filename="Surat_Tugas_' . preg_replace('/[^a-zA-Z0-9\-_]/', '_', $surat['nomor_surat']) . '.pdf"');
                }
                header('Content-Length: ' . filesize($tempPdf));
                readfile($tempPdf);
                @unlink($tempPdf);
            } catch (Exception $pdfEx) {
                if ($tempPdf && file_exists($tempPdf)) @unlink($tempPdf);
                throw $pdfEx;
            }
            exit;
        }

        // format=docx: unduh file Word
        header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        header('Content-Disposition: attachment; filename="Surat_Tugas_' . $surat['nomor_surat'] . '.docx"');
        header('Content-Length: ' . filesize($tempDocx));
        readfile($tempDocx);
        @unlink($tempDocx);
    } catch (Exception $e) {
        if (file_exists($tempDocx)) @unlink($tempDocx);
        throw $e;
    }

} catch (Exception $e) {
    die("Error: " . $e->getMessage());
}