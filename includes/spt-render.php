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

function &spt_render_word_state()
{
    static $state = ['word' => null, 'failed' => false, 'cleanup' => false];
    return $state;
}

function spt_render_word(&$error = null)
{
    $state = &spt_render_word_state();
    if ($state['word'] !== null) {
        return $state['word'];
    }
    if ($state['failed'] || !class_exists('COM')) {
        $error = $state['failed'] ? 'Word COM gagal diinisialisasi' : 'Ekstensi COM tidak aktif';
        return null;
    }
    try {
        $word = new COM('Word.Application');
        $word->Visible = false;
        $word->DisplayAlerts = 0;
        $state['word'] = $word;
        if (!$state['cleanup']) {
            // Word harus selalu ditutup, termasuk saat request berakhir karena
            // error/timeout, agar proses WINWORD tidak menumpuk di server.
            $state['cleanup'] = true;
            register_shutdown_function('spt_render_word_quit');
        }
        return $word;
    } catch (Throwable $e) {
        $state['failed'] = true;
        $error = $e->getMessage();
        $state['word'] = null;
        return null;
    }
}

function spt_render_word_quit()
{
    $state = &spt_render_word_state();
    $word = $state['word'];
    $state['word'] = null;
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
function spt_render_fit($templateDocx, array $settings, $targetPages = null)
{
    require_once __DIR__ . DIRECTORY_SEPARATOR . 'spt-layout-engine.php';

    $hasLampiran = function_exists('spt_le_docx_has_lampiran')
        ? spt_le_docx_has_lampiran($templateDocx)
        : false;
    if ($targetPages === null) {
        // Surat berlampiran butuh 1 halaman naskah + 1 halaman lampiran.
        $targetPages = $hasLampiran ? 2 : 1;
    }
    $targetPages = max(1, (int)$targetPages);

    $result = [
        'settings' => $settings,
        'pages' => null,
        'docx' => null,
        'ratio' => 1.0,
        'attempts' => [],
        'measured' => false,
        'target_pages' => $targetPages,
        'has_lampiran' => $hasLampiran,
    ];

    $workDir = spt_render_tmp_dir() . DIRECTORY_SEPARATOR . 'fit_' . uniqid();
    if (!is_dir($workDir) && !@mkdir($workDir, 0755, true)) {
        return $result;
    }

    // Hapus semua percobaan kecuali kandidat yang dikembalikan, agar folder
    // kerja solver tidak menumpuk file DOCX.
    $files = [];
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

    // Bangun + ukur satu kandidat. Return null bila file gagal dibuat.
    $attempt = function ($index) use (&$files, &$result, $templateDocx, $settings, $ratios, $workDir) {
        $ratio = $ratios[$index];
        $candidate = $ratio >= 1.0 ? $settings : spt_le_compress($settings, $ratio);
        $file = $workDir . DIRECTORY_SEPARATOR . 'fit_' . str_replace('.', '_', (string)$ratio) . '.docx';
        if (!@copy($templateDocx, $file)) {
            return null;
        }
        if (!spt_le_apply_docx($file, $candidate)) {
            @unlink($file);
            return null;
        }
        $files[] = $file;
        $pages = spt_render_measure_pages($file);
        if ($pages !== null) {
            $result['measured'] = true;
        }
        $result['attempts'][] = ['ratio' => $ratio, 'pages' => $pages];
        return [
            'ratio' => $ratio,
            'pages' => $pages,
            'file' => $file,
            'settings' => $candidate,
        ];
    };

    $first = $attempt(0);
    if ($first === null) {
        $purge(null);
        return $result;
    }
    if ($first['pages'] === null) {
        // Word tidak tersedia: pakai pengaturan apa adanya.
        $result['settings'] = $first['settings'];
        $result['docx'] = $first['file'];
        $result['ratio'] = $first['ratio'];
        $purge($first['file']);
        return $result;
    }
    if ($first['pages'] <= $targetPages) {
        $result['settings'] = $first['settings'];
        $result['pages'] = $first['pages'];
        $result['docx'] = $first['file'];
        $result['ratio'] = $first['ratio'];
        $purge($first['file']);
        return $result;
    }

    // Binary search pada indeks 1..6: cari rasio paling longgar yang muat.
    // Kandidat paling ringan dengan halaman paling sedikit disimpan sebagai
    // fallback bila target tidak tercapai.
    $least = $first;
    $fit = null;
    $low = 1;
    $high = count($ratios) - 1;
    while ($low <= $high) {
        $mid = intdiv($low + $high, 2);
        $candidate = $attempt($mid);
        if ($candidate === null) {
            $low = $mid + 1;
            continue;
        }
        if ($candidate['pages'] === null) {
            $result['settings'] = $candidate['settings'];
            $result['docx'] = $candidate['file'];
            $result['ratio'] = $candidate['ratio'];
            $purge($candidate['file']);
            return $result;
        }
        if ($candidate['pages'] <= $targetPages) {
            $fit = $candidate;
            $high = $mid - 1;
        } else {
            if ($candidate['pages'] < $least['pages']) {
                $least = $candidate;
            }
            $low = $mid + 1;
        }
    }

    $chosen = $fit ?? $least;
    $result['settings'] = $chosen['settings'];
    $result['pages'] = $chosen['pages'];
    $result['docx'] = $chosen['file'];
    $result['ratio'] = $chosen['ratio'];
    $purge($chosen['file']);
    return $result;
}
