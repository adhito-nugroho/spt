<?php
/**
 * Renderer PDF dan pengukur halaman untuk Surat Tugas.
 *
 * Memakai satu sesi Microsoft Word (COM) per request agar pengukuran halaman
 * dan ekspor PDF tidak membuat instance Word berulang kali.
 */

function spt_render_tmp_dir()
{
    static $dir = null;
    if ($dir !== null) {
        return $dir;
    }
    $root = dirname(__DIR__);
    foreach ([$root . DIRECTORY_SEPARATOR . 'tmp', sys_get_temp_dir()] as $candidate) {
        if (!is_dir($candidate)) {
            @mkdir($candidate, 0755, true);
        }
        if (!is_dir($candidate)) {
            continue;
        }
        $probe = $candidate . DIRECTORY_SEPARATOR . 'sptw_' . uniqid() . '.tmp';
        $handle = @fopen($probe, 'wb');
        if ($handle !== false) {
            fclose($handle);
            @unlink($probe);
            return $dir = $candidate;
        }
    }
    return $dir = sys_get_temp_dir();
}

function spt_render_word(&$error = null)
{
    static $word = null;
    static $failed = false;
    if ($word !== null) {
        return $word;
    }
    if ($failed || !class_exists('COM')) {
        $error = $failed ? 'Word COM gagal diinisialisasi' : 'Ekstensi COM tidak aktif';
        return null;
    }
    try {
        $word = new COM('Word.Application');
        $word->Visible = false;
        $word->DisplayAlerts = 0;
        return $word;
    } catch (Throwable $e) {
        $failed = true;
        $error = $e->getMessage();
        $word = null;
        return null;
    }
}

function spt_render_word_quit()
{
    $error = null;
    $word = spt_render_word($error);
    if ($word === null) {
        return;
    }
    try {
        $word->Quit();
    } catch (Throwable $e) {
        // abaikan
    }
}

/**
 * Buka dokumen di Word. Return objek dokumen atau null.
 */
function spt_render_word_open($docxPath, &$error = null, $maxAttempts = 2)
{
    $word = spt_render_word($error);
    if ($word === null) {
        return null;
    }
    if (!is_file($docxPath)) {
        $error = 'File DOCX tidak ditemukan';
        return null;
    }
    for ($attempt = 1; $attempt <= $maxAttempts; $attempt++) {
        try {
            return $word->Documents->Open($docxPath, false, true, false);
        } catch (Throwable $e) {
            $error = $e->getMessage();
            if ($attempt < $maxAttempts) {
                usleep(400000);
            }
        }
    }
    return null;
}

/**
 * Jumlah halaman dokumen menurut tata letak Word. Null jika tidak tersedia.
 */
function spt_render_measure_pages($docxPath, &$error = null)
{
    $doc = spt_render_word_open($docxPath, $error);
    if ($doc === null) {
        return null;
    }
    try {
        $pages = (int)$doc->ComputeStatistics(2); // wdStatisticPages
        $doc->Close(false);
        return $pages > 0 ? $pages : null;
    } catch (Throwable $e) {
        $error = $e->getMessage();
        try {
            $doc->Close(false);
        } catch (Throwable $e2) {
            // abaikan
        }
        return null;
    }
}

function spt_render_pdf_page_count($pdfPath)
{
    if (!is_file($pdfPath) || filesize($pdfPath) < 100) {
        return 0;
    }
    $content = @file_get_contents($pdfPath);
    if ($content === false) {
        return 0;
    }
    if (preg_match('/\/Count\s+(\d+)/', $content, $match)) {
        return (int)$match[1];
    }
    return preg_match_all('/\/Type\s*\/Page[^s]/', $content);
}

/**
 * Ekspor DOCX ke PDF. Return ['path','method','pages'] atau null.
 * Urutan: Word COM -> LibreOffice.
 */
function spt_render_pdf($docxPath, $pdfPath = null, &$error = null)
{
    if (!is_file($docxPath)) {
        $error = 'File DOCX tidak ditemukan';
        return null;
    }
    if ($pdfPath === null) {
        $pdfPath = spt_render_tmp_dir() . DIRECTORY_SEPARATOR
            . pathinfo($docxPath, PATHINFO_FILENAME) . '_' . uniqid() . '.pdf';
    }

    $doc = spt_render_word_open($docxPath, $openError);
    if ($doc !== null) {
        try {
            $pages = (int)$doc->ComputeStatistics(2);
            @unlink($pdfPath);
            $doc->ExportAsFixedFormat(
                $pdfPath, 17, false, 0, 0, 1, 1, 0, true, true, 0, true, true, false
            );
            $doc->Close(false);
            if (is_file($pdfPath) && filesize($pdfPath) > 100) {
                return ['path' => $pdfPath, 'method' => 'Word COM', 'pages' => $pages > 0 ? $pages : null];
            }
            $error = 'Word COM tidak menghasilkan PDF';
        } catch (Throwable $e) {
            $error = $e->getMessage();
            try {
                $doc->Close(false);
            } catch (Throwable $e2) {
                // abaikan
            }
        }
    } else {
        $error = $openError;
    }

    $outDir = null;
    $pdfError = null;
    $pdf = spt_render_libreoffice($docxPath, $outDir, $pdfError);
    if ($pdf !== null) {
        return ['path' => $pdf, 'method' => 'LibreOffice', 'pages' => spt_render_pdf_page_count($pdf)];
    }
    $error = trim(($error ?? '') . ' ' . ($pdfError ?? ''));
    return null;
}

/**
 * Konversi DOCX ke PDF lewat LibreOffice headless.
 */
function spt_render_libreoffice($docxPath, &$outDir = null, &$error = null)
{
    $soffice = null;
    foreach (['soffice', 'soffice.exe', 'libreoffice'] as $name) {
        $found = trim((string)@shell_exec('where ' . $name . ' 2>NUL'));
        if ($found !== '') {
            $soffice = strtok($found, "\r\n");
            break;
        }
    }
    if ($soffice === null) {
        $error = 'LibreOffice tidak ditemukan';
        return null;
    }
    $outDir = spt_render_tmp_dir() . DIRECTORY_SEPARATOR . 'lo_' . uniqid();
    if (!is_dir($outDir) && !@mkdir($outDir, 0755, true)) {
        $error = 'Folder sementara LibreOffice gagal dibuat';
        return null;
    }
    $command = '"' . $soffice . '" --headless --norestore --convert-to pdf --outdir "'
        . $outDir . '" "' . $docxPath . '" 2>&1';
    @shell_exec($command);
    $expected = $outDir . DIRECTORY_SEPARATOR . pathinfo($docxPath, PATHINFO_FILENAME) . '.pdf';
    if (is_file($expected) && filesize($expected) > 100) {
        return $expected;
    }
    $error = 'LibreOffice gagal membuat PDF';
    return null;
}

/**
 * Cari pengaturan paling longgar yang tetap muat dalam $targetPages halaman.
 *
 * $templateDocx : DOCX hasil generate sebelum layout diterapkan.
 * $settings     : pengaturan layout awal.
 * Return ['settings','pages','docx','ratio','attempts','measured'].
 *   measured=false berarti Word tidak tersedia sehingga penyusutan tidak bisa
 *   diverifikasi; settings dikembalikan apa adanya.
 */
function spt_render_fit($templateDocx, array $settings, $targetPages = 1)
{
    require_once __DIR__ . DIRECTORY_SEPARATOR . 'spt-layout-engine.php';

    $result = [
        'settings' => $settings,
        'pages' => null,
        'docx' => null,
        'ratio' => 1.0,
        'attempts' => [],
        'measured' => false,
    ];

    $workDir = spt_render_tmp_dir() . DIRECTORY_SEPARATOR . 'fit_' . uniqid();
    if (!is_dir($workDir) && !@mkdir($workDir, 0755, true)) {
        return $result;
    }

    // Hapus semua percobaan kecuali kandidat yang dikembalikan, agar folder
    // kerja solver tidak menumpuk file DOCX.
    $purge = static function ($keep) use (&$files, $workDir) {
        foreach ($files as $file) {
            if ($keep === null || $file !== $keep) {
                @unlink($file);
            }
        }
        if ($keep === null) {
            @rmdir($workDir);
        }
    };

    $ratios = [1.0, 0.94, 0.88, 0.84, 0.80, 0.76, 0.72];
    $files = [];
    $best = null;      // kandidat paling rapat
    $leastCompressed = null; // kandidat paling longgar dengan halaman tersedikit
    foreach ($ratios as $ratio) {
        $candidate = $ratio >= 1.0 ? $settings : spt_le_compress($settings, $ratio);
        $file = $workDir . DIRECTORY_SEPARATOR . 'fit_' . str_replace('.', '_', (string)$ratio) . '.docx';
        if (!@copy($templateDocx, $file)) {
            continue;
        }
        if (!spt_le_apply_docx($file, $candidate)) {
            @unlink($file);
            continue;
        }
        $files[] = $file;
        $pages = spt_render_measure_pages($file);
        if ($pages === null) {
            // Word tidak tersedia: berhenti di percobaan pertama.
            $result['settings'] = $candidate;
            $result['docx'] = $file;
            $result['ratio'] = $ratio;
            $purge($file);
            return $result;
        }
        $result['measured'] = true;
        $result['attempts'][] = ['ratio' => $ratio, 'pages' => $pages];
        $candidateInfo = ['ratio' => $ratio, 'pages' => $pages, 'file' => $file, 'settings' => $candidate];
        if ($pages <= $targetPages) {
            $result['settings'] = $candidate;
            $result['pages'] = $pages;
            $result['docx'] = $file;
            $result['ratio'] = $ratio;
            $purge($file);
            return $result;
        }
        if ($leastCompressed === null || $pages < $leastCompressed['pages']) {
            $leastCompressed = $candidateInfo;
        }
        $best = $candidateInfo;
    }

    // Target tidak tercapai: pakai kompresi paling ringan yang menghasilkan
    // jumlah halaman paling sedikit (bukan yang paling rapat).
    $fallback = $leastCompressed ?? $best;
    if ($fallback !== null) {
        $result['settings'] = $fallback['settings'];
        $result['pages'] = $fallback['pages'];
        $result['docx'] = $fallback['file'];
        $result['ratio'] = $fallback['ratio'];
        $purge($fallback['file']);
    } else {
        $purge(null);
    }
    return $result;
}
