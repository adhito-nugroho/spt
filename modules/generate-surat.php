<?php
require_once __DIR__ . '/../config/database.php';
require_once __DIR__ . '/../vendor/autoload.php';

use PhpOffice\PhpWord\TemplateProcessor;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\Settings;

/**
 * Cek apakah direktori benar-benar bisa ditulisi file.
 * Di Windows, is_writable() pada direktori sering menghasilkan false meski berizin tulis.
 */
function spt_is_dir_writable($dir) {
    if (!is_string($dir) || $dir === '' || !is_dir($dir)) {
        return false;
    }
    $testFile = rtrim($dir, '\\/') . DIRECTORY_SEPARATOR . 'spt_w_' . uniqid() . '.tmp';
    $fp = @fopen($testFile, 'wb');
    if ($fp === false) {
        return false;
    }
    fclose($fp);
    @unlink($testFile);
    return true;
}

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
        dirname(__DIR__) . DIRECTORY_SEPARATOR . 'tmp',
        'C:\\laragon\\tmp',
        'C:\\Windows\\Temp',
        sys_get_temp_dir(),
    ];
    $dir = null;
    foreach ($candidateDirs as $candidate) {
        if (spt_is_dir_writable($candidate)) {
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

/**
 * =====================================================================
 * INTERACTIVE LAYOUT & MARGIN ADJUSTER
 * =====================================================================
 * Semua parameter layout dikirim lewat query string, misalnya:
 * generate-surat.php?id=1&format=pdf&layout=1&margin_top=15&font_size=11...
 *
 * Karena PDF utama dihasilkan dari DOCX (Word COM / LibreOffice), parameter
 * layout diterapkan langsung ke properti OOXML (margin halaman, ukuran font,
 * line-height, jarak paragraf, indentasi blok TTD). Untuk fallback Dompdf,
 * parameter yang sama juga disuntikkan sebagai CSS ke HTML sebelum
 * $dompdf->loadHtml() dipanggil oleh renderer PHPWord.
 */

function spt_layout_default_values()
{
    return [
        'margin_top' => 15,
        'margin_bottom' => 15,
        'margin_left' => 20,
        'margin_right' => 20,
        'font_size' => 11,
        'line_height' => 1.0,
        'kop_line_spacing' => 0,
        'kop_bottom_spacing' => 10,
        'judul_top_spacing' => 0,
        'nomor_surat_spacing' => 4,
        'judul_bottom_spacing' => 10,
        'dasar_top_spacing' => 6,
        'dasar_item_spacing' => 3,
        'dasar_bottom_spacing' => 8,
        'memerintahkan_spacing' => 8,
        'pegawai_top_spacing' => 6,
        'pegawai_gap_spacing' => 8,
        'pegawai_line_spacing' => 2,
        'untuk_top_spacing' => 6,
        'untuk_item_spacing' => 3,
        'untuk_bottom_spacing' => 10,
        'ttd_offset' => 10,
        'ttd_x' => 0,
        'ttd_date_to_jabatan' => 4,
        'ttd_space_height' => 60,
        'ttd_name_nip_spacing' => 4,
        'lock_ttd' => 0,
    ];
}

function spt_layout_ranges()
{
    return [
        'margin_top' => [5, 40, 0.5],
        'margin_bottom' => [5, 40, 0.5],
        'margin_left' => [10, 40, 0.5],
        'margin_right' => [10, 40, 0.5],
        'font_size' => [9, 13, 0.5],
        'line_height' => [1.0, 2.0, 0.05],
        'kop_line_spacing' => [0, 20, 1],
        'kop_bottom_spacing' => [0, 50, 1],
        'judul_top_spacing' => [-15, 40, 1],
        'nomor_surat_spacing' => [2, 30, 1],
        'judul_bottom_spacing' => [0, 40, 1],
        'dasar_top_spacing' => [0, 30, 1],
        'dasar_item_spacing' => [0, 20, 1],
        'dasar_bottom_spacing' => [0, 30, 1],
        'memerintahkan_spacing' => [0, 30, 1],
        'pegawai_top_spacing' => [0, 30, 1],
        'pegawai_gap_spacing' => [0, 35, 1],
        'pegawai_line_spacing' => [0, 15, 1],
        'untuk_top_spacing' => [0, 30, 1],
        'untuk_item_spacing' => [0, 20, 1],
        'untuk_bottom_spacing' => [0, 30, 1],
        'ttd_offset' => [-20, 80, 1],
        'ttd_x' => [-50, 50, 1],
        'ttd_date_to_jabatan' => [0, 25, 1],
        'ttd_space_height' => [20, 120, 2],
        'ttd_name_nip_spacing' => [0, 25, 1],
    ];
}

function spt_layout_clamp($value, $min, $max)
{
    $value = (float)$value;
    if ($value < $min) {
        return $min;
    }
    if ($value > $max) {
        return $max;
    }
    return $value;
}

function spt_layout_round($value, $step)
{
    if ($step <= 0) {
        return $value;
    }
    return round(round((float)$value / $step) * $step, 3);
}

/**
 * Baca parameter layout dari query string. Return null jika user belum
 * mengaktifkan penyesuaian (agar halaman awal tetap memakai template asli).
 */
function spt_layout_request_settings()
{
    $keys = array_keys(spt_layout_ranges());
    $active = isset($_GET['layout']) && (string)$_GET['layout'] === '1';
    if (!$active) {
        foreach ($keys as $key) {
            if (isset($_GET[$key])) {
                $active = true;
                break;
            }
        }
    }
    if (!$active) {
        return null;
    }

    $settings = [];
    foreach (spt_layout_ranges() as $key => $range) {
        if (isset($_GET[$key]) && is_scalar($_GET[$key])) {
            $settings[$key] = spt_layout_round(
                spt_layout_clamp($_GET[$key], $range[0], $range[1]),
                $range[2]
            );
        }
    }
    $settings['lock_ttd'] = (isset($_GET['lock_ttd']) && (string)$_GET['lock_ttd'] === '1') ? 1 : 0;

    return $settings;
}

function spt_layout_query_params(array $settings)
{
    $params = ['layout' => '1'];
    foreach (array_keys(spt_layout_ranges()) as $key) {
        $params[$key] = $settings[$key] ?? spt_layout_default_values()[$key];
    }
    $params['lock_ttd'] = !empty($settings['lock_ttd']) ? '1' : '0';
    return $params;
}

/**
 * Normalisasi setting layout yang tersimpan di database.
 * Hanya key yang dikenal yang dipakai, nilainya di-clamp ke range slider.
 */
function spt_layout_normalize_stored_settings($raw)
{
    if (!is_array($raw)) {
        return null;
    }

    $settings = [];
    foreach (spt_layout_ranges() as $key => $range) {
        if (isset($raw[$key]) && is_numeric($raw[$key])) {
            $settings[$key] = spt_layout_round(
                spt_layout_clamp($raw[$key], $range[0], $range[1]),
                $range[2]
            );
        }
    }
    if (array_key_exists('lock_ttd', $raw)) {
        $settings['lock_ttd'] = !empty($raw['lock_ttd']) ? 1 : 0;
    }

    return empty($settings) ? null : $settings;
}

/**
 * Ambil pengaturan layout default yang tersimpan untuk sebuah template.
 * Return null jika kolom belum ada atau template belum punya setting.
 */
function spt_layout_load_template_settings($pdo, $templateId)
{
    $templateId = (int)$templateId;
    if ($templateId <= 0) {
        return null;
    }

    try {
        $stmt = $pdo->prepare("SELECT layout_settings FROM templates WHERE id = ? LIMIT 1");
        $stmt->execute([$templateId]);
        $row = $stmt->fetch();
    } catch (PDOException $e) {
        // Kolom layout_settings belum ada (database lama).
        return null;
    }

    if (!$row || empty($row['layout_settings'])) {
        return null;
    }

    $decoded = json_decode($row['layout_settings'], true);
    return spt_layout_normalize_stored_settings($decoded);
}

function spt_layout_mm_to_twips($mm)
{
    return (int)round(((float)$mm) * 1440 / 25.4);
}

function spt_layout_px_to_twips($px)
{
    return (int)round(((float)$px) * 15);
}

function spt_layout_w_ns()
{
    return 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
}

function spt_layout_load_dom($xml)
{
    if (!is_string($xml) || $xml === '') {
        return null;
    }
    $dom = new DOMDocument('1.0', 'UTF-8');
    $dom->preserveWhiteSpace = true;
    $dom->formatOutput = false;
    $prev = libxml_use_internal_errors(true);
    $ok = $dom->loadXML($xml, LIBXML_NONET);
    libxml_clear_errors();
    libxml_use_internal_errors($prev);

    return $ok ? $dom : null;
}

function spt_layout_child(DOMElement $parent, $localName)
{
    foreach ($parent->childNodes as $child) {
        if ($child instanceof DOMElement && $child->localName === $localName) {
            return $child;
        }
    }
    return null;
}

function spt_layout_insert_ordered(DOMElement $parent, DOMElement $element, array $order)
{
    $rank = array_search($element->localName, $order, true);
    foreach ($parent->childNodes as $child) {
        if (!$child instanceof DOMElement) {
            continue;
        }
        $childRank = array_search($child->localName, $order, true);
        if ($childRank !== false && ($rank === false || $childRank > $rank)) {
            $parent->insertBefore($element, $child);
            return;
        }
    }
    $parent->appendChild($element);
}

function spt_layout_paragraph_properties(DOMElement $paragraph, DOMDocument $dom)
{
    $pPr = spt_layout_child($paragraph, 'pPr');
    if ($pPr) {
        return $pPr;
    }

    $pPr = $dom->createElementNS(spt_layout_w_ns(), 'w:pPr');
    if ($paragraph->firstChild) {
        $paragraph->insertBefore($pPr, $paragraph->firstChild);
    } else {
        $paragraph->appendChild($pPr);
    }
    return $pPr;
}

function spt_layout_spacing(DOMElement $paragraph, DOMDocument $dom)
{
    $pPr = spt_layout_paragraph_properties($paragraph, $dom);
    $spacing = spt_layout_child($pPr, 'spacing');
    if ($spacing) {
        return $spacing;
    }

    $spacing = $dom->createElementNS(spt_layout_w_ns(), 'w:spacing');
    spt_layout_insert_ordered($pPr, $spacing, [
        'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
        'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens',
        'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE',
        'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
        'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
        'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
        'sectPr', 'pPrChange',
    ]);
    return $spacing;
}

function spt_layout_indentation(DOMElement $paragraph, DOMDocument $dom)
{
    $pPr = spt_layout_paragraph_properties($paragraph, $dom);
    $ind = spt_layout_child($pPr, 'ind');
    if ($ind) {
        return $ind;
    }

    $ind = $dom->createElementNS(spt_layout_w_ns(), 'w:ind');
    spt_layout_insert_ordered($pPr, $ind, [
        'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
        'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens',
        'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE',
        'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
        'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
        'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
        'sectPr', 'pPrChange',
    ]);
    return $ind;
}

function spt_layout_paragraph_text(DOMElement $paragraph)
{
    $text = '';
    foreach ($paragraph->getElementsByTagNameNS(spt_layout_w_ns(), 't') as $node) {
        $text .= $node->textContent;
    }
    return trim(preg_replace('/\s+/u', ' ', $text));
}

function spt_layout_reduce_spacing(DOMElement $spacing, $attribute, $amount)
{
    $ns = spt_layout_w_ns();
    $qualified = 'w:' . $attribute;
    if ($amount <= 0 || !$spacing->hasAttributeNS($ns, $attribute)) {
        return $amount;
    }
    $current = (int)$spacing->getAttributeNS($ns, $attribute);
    $used = min($current, $amount);
    $spacing->setAttributeNS($ns, $qualified, (string)($current - $used));
    return $amount - $used;
}

/**
 * Geser jarak vertikal paragraf. Nilai positif menambah spacing sebelum
 * paragraf; nilai negatif mengurangi spacing paragraf sebelumnya.
 */
function spt_layout_shift_gap(DOMElement $paragraph, $deltaTwips, DOMDocument $dom)
{
    if ($deltaTwips === 0) {
        return;
    }

    $ns = spt_layout_w_ns();
    $spacing = spt_layout_spacing($paragraph, $dom);

    if ($deltaTwips > 0) {
        $current = $spacing->hasAttributeNS($ns, 'before')
            ? (int)$spacing->getAttributeNS($ns, 'before')
            : 0;
        $spacing->setAttributeNS($ns, 'w:before', (string)($current + $deltaTwips));
        return;
    }

    $remaining = abs((int)$deltaTwips);
    $remaining = spt_layout_reduce_spacing($spacing, 'before', $remaining);

    $previous = $paragraph->previousSibling;
    while ($previous && $remaining > 0) {
        if ($previous instanceof DOMElement && $previous->localName === 'p') {
            $previousSpacing = spt_layout_spacing($previous, $dom);
            $remaining = spt_layout_reduce_spacing($previousSpacing, 'after', $remaining);
            if ($remaining > 0) {
                $remaining = spt_layout_reduce_spacing($previousSpacing, 'before', $remaining);
            }
        }
        $previous = $previous->previousSibling;
    }
}

/**
 * Ciutkan paragraf kosong agar tidak memakan ruang vertikal dokumen.
 * Menjadikan line height minimal (1 pt) dan spasi sebelum/sesudah 0.
 */
function spt_layout_collapse_paragraph(DOMElement $paragraph, DOMDocument $dom)
{
    $spacing = spt_layout_spacing($paragraph, $dom);
    $ns = spt_layout_w_ns();
    $spacing->setAttributeNS($ns, 'w:before', '0');
    $spacing->setAttributeNS($ns, 'w:after', '0');
    $spacing->setAttributeNS($ns, 'w:line', '20');
    $spacing->setAttributeNS($ns, 'w:lineRule', 'exact');
}

function spt_layout_add_keep_flag(DOMElement $paragraph, $flag, DOMDocument $dom)
{
    $pPr = spt_layout_paragraph_properties($paragraph, $dom);
    if (spt_layout_child($pPr, $flag)) {
        return;
    }

    $element = $dom->createElementNS(spt_layout_w_ns(), 'w:' . $flag);
    spt_layout_insert_ordered($pPr, $element, [
        'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
        'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens',
        'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE',
        'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
        'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
        'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
        'sectPr', 'pPrChange',
    ]);
}

/**
 * Rata-atas semua sel tabel agar label dan isi (mis. "Dasar", biodata
 * pegawai) tetap sejajar di baris pertama.
 */
function spt_layout_align_tables_top(DOMDocument $document)
{
    $ns = spt_layout_w_ns();
    $xpath = new DOMXPath($document);
    $xpath->registerNamespace('w', $ns);

    $tcPrOrder = [
        'cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd',
        'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark',
        'headers', 'cellIns', 'cellDel', 'cellMerge', 'tcPrChange',
    ];

    foreach ($xpath->query('//w:tbl//w:tc') as $cell) {
        $tcPr = spt_layout_child($cell, 'tcPr');
        if (!$tcPr) {
            $tcPr = $document->createElementNS($ns, 'w:tcPr');
            if ($cell->firstChild) {
                $cell->insertBefore($tcPr, $cell->firstChild);
            } else {
                $cell->appendChild($tcPr);
            }
        }
        $vAlign = spt_layout_child($tcPr, 'vAlign');
        if (!$vAlign) {
            $vAlign = $document->createElementNS($ns, 'w:vAlign');
            spt_layout_insert_ordered($tcPr, $vAlign, $tcPrOrder);
        }
        $vAlign->setAttributeNS($ns, 'w:val', 'top');
    }

    // Tinggi baris "exact" bisa memotong teks saat font/spasi dikompres.
    // Ubah ke "atLeast" agar baris tumbuh mengikuti isi.
    foreach ($xpath->query('//w:tbl//w:tr/w:trPr/w:trHeight') as $rowHeight) {
        $rowHeight->setAttributeNS($ns, 'w:hRule', 'atLeast');
    }
}

/**
 * Hitung line-height aman (twips) untuk judul/nomor berdasarkan ukuran
 * huruf terbesar di paragraf tersebut.
 */
function spt_layout_heading_line_twips(DOMElement $paragraph, $multiplier = 1.3)
{
    $ns = spt_layout_w_ns();
    $maxHalfPoints = 0;
    foreach ($paragraph->getElementsByTagNameNS($ns, 'sz') as $sizeNode) {
        $maxHalfPoints = max($maxHalfPoints, (int)$sizeNode->getAttributeNS($ns, 'val'));
    }
    if ($maxHalfPoints <= 0) {
        $maxHalfPoints = 24; // fallback 12 pt
    }

    $multiplier = max(1.3, (float)$multiplier);
    return (int)round(($maxHalfPoints / 2) * $multiplier * 20);
}

/**
 * Normalisasi template tanpa mengubah pengaturan layout pengguna:
 * - bersihkan kata berulang pada teks template (termasuk "tentang tentang"),
 * - pastikan judul "SURAT TUGAS" dan "Nomor" tidak bertumpuk,
 * - ratakan semua sel tabel ke atas.
 * Fungsi ini selalu dipanggil agar preview default pun rapi.
 */
function spt_normalize_docx($docxPath)
{
    if (!class_exists('ZipArchive') || !is_file($docxPath)) {
        return false;
    }

    $zip = new ZipArchive();
    if ($zip->open($docxPath) !== true) {
        return false;
    }

    $documentXml = $zip->getFromName('word/document.xml');
    if ($documentXml === false) {
        $zip->close();
        return false;
    }

    $documentXml = spt_fix_common_typos($documentXml);
    $document = spt_layout_load_dom($documentXml);
    if (!$document) {
        $zip->close();
        return false;
    }

    $ns = spt_layout_w_ns();
    $xpath = new DOMXPath($document);
    $xpath->registerNamespace('w', $ns);
    $xpath->registerNamespace('v', 'urn:schemas-microsoft-com:vml');

    // 1. Bersihkan kata berulang di dalam setiap run teks.
    foreach ($xpath->query('//w:t') as $textNode) {
        $original = $textNode->textContent;
        $cleaned = spt_clean_repeated_words($original);
        if ($cleaned !== $original) {
            $textNode->textContent = $cleaned;
        }
    }

    // 1b. Bersihkan kata berulang yang terpecah di antara dua run berdekatan.
    foreach ($xpath->query('//w:p') as $paragraph) {
        $textNodes = [];
        foreach ($xpath->query('.//w:t', $paragraph) as $textNode) {
            $textNodes[] = $textNode;
        }
        for ($i = 0; $i < count($textNodes) - 1; $i++) {
            $left = $textNodes[$i]->textContent;
            $right = $textNodes[$i + 1]->textContent;
            if (!preg_match('/\b([\p{L}]+)\s*$/u', $left, $mLeft)) {
                continue;
            }
            if (!preg_match('/^\s*([\p{L}]+)\b/u', $right, $mRight)) {
                continue;
            }
            if (mb_strtolower($mLeft[1], 'UTF-8') !== mb_strtolower($mRight[1], 'UTF-8')) {
                continue;
            }
            $newLeft = preg_replace('/\s*\b' . preg_quote($mLeft[1], '/') . '\s*$/u', '', $left);
            if ($newLeft !== null && $newLeft !== $left) {
                $textNodes[$i]->textContent = $newLeft;
            }
        }
    }

    // 2. Amankan jarak judul "SURAT TUGAS" ke "Nomor".
    $paragraphs = [];
    foreach ($xpath->query('//w:body//w:p') as $paragraph) {
        $paragraphs[] = $paragraph;
    }

    $titleIndex = null;
    $numberIndex = null;
    $awaitingNumber = false;
    foreach ($paragraphs as $index => $paragraph) {
        $text = spt_layout_paragraph_text($paragraph);
        if ($text === '') {
            continue;
        }
        if ($titleIndex === null && stripos($text, 'SURAT TUGAS') !== false) {
            $titleIndex = $index;
            if (preg_match('/SURAT\s+TUGAS.*?Nomor\s*:/isu', $text)) {
                $numberIndex = $index;
            } else {
                $awaitingNumber = true;
            }
            continue;
        }
        if ($awaitingNumber) {
            $awaitingNumber = false;
            if (stripos($text, 'Nomor') !== false) {
                $numberIndex = $index;
            }
        }
    }

    // Rapatkan paragraf kosong antara judul dan nomor agar jarak benar-benar
    // dikendalikan oleh parameter nomor_surat_spacing.
    if ($titleIndex !== null && $numberIndex !== null && $numberIndex > $titleIndex) {
        for ($i = $titleIndex + 1; $i < $numberIndex; $i++) {
            if (isset($paragraphs[$i]) && spt_layout_paragraph_text($paragraphs[$i]) === '') {
                spt_layout_collapse_paragraph($paragraphs[$i], $document);
            }
        }
    }

    // 2. Baris kop dengan line-height "exact" yang terlalu kecil bisa terpotong.
    if ($titleIndex !== null) {
        for ($i = 0; $i < $titleIndex; $i++) {
            if (!isset($paragraphs[$i]) || spt_layout_paragraph_text($paragraphs[$i]) === '') {
                continue;
            }
            $spacing = spt_layout_spacing($paragraphs[$i], $document);
            if (!$spacing->hasAttributeNS($ns, 'lineRule')) {
                continue;
            }
            if (strtolower($spacing->getAttributeNS($ns, 'lineRule')) !== 'exact') {
                continue;
            }
            $line = $spacing->hasAttributeNS($ns, 'line') ? (int)$spacing->getAttributeNS($ns, 'line') : 0;
            $singleLine = spt_layout_heading_line_twips($paragraphs[$i], 1.0);
            if ($line < $singleLine) {
                $spacing->setAttributeNS($ns, 'w:line', (string)max(240, $singleLine));
                $spacing->setAttributeNS($ns, 'w:lineRule', 'auto');
            }
        }
    }

    // 2b. Garis kop pada template berupa shape absolut sehingga bisa menabrak
    // teks saat spasi berubah. Ganti dengan border paragraf yang mengikuti teks.
    $linePicts = [];
    foreach ($xpath->query('//w:pict') as $pict) {
        if ($xpath->query('.//v:imagedata', $pict)->length > 0) {
            continue;
        }
        if ($xpath->query('.//w:txbxContent', $pict)->length > 0) {
            continue;
        }
        $lineShapes = 0;
        foreach ($xpath->query('.//v:shape', $pict) as $shape) {
            if ($shape->hasAttribute('path') && $shape->hasAttribute('strokeweight')) {
                $lineShapes++;
            }
        }
        if ($lineShapes > 0) {
            $linePicts[] = $pict;
        }
    }

    if (!empty($linePicts)) {
        foreach ($linePicts as $pict) {
            if ($pict->parentNode) {
                $pict->parentNode->removeChild($pict);
            }
        }

        $kopEndIndex = null;
        if ($titleIndex !== null) {
            for ($i = $titleIndex - 1; $i >= 0; $i--) {
                if (isset($paragraphs[$i]) && spt_layout_paragraph_text($paragraphs[$i]) !== '') {
                    $kopEndIndex = $i;
                    break;
                }
            }
        }

        $alreadyHasKopBorder = false;
        if ($kopEndIndex !== null && isset($paragraphs[$kopEndIndex])) {
            $kopProperties = spt_layout_child($paragraphs[$kopEndIndex], 'pPr');
            if ($kopProperties && spt_layout_child($kopProperties, 'pBdr')) {
                $alreadyHasKopBorder = true;
            }
        }

        if (!$alreadyHasKopBorder && $kopEndIndex !== null && isset($paragraphs[$kopEndIndex])) {
            $ruleParagraph = $document->createElementNS($ns, 'w:p');
            $ruleProperties = $document->createElementNS($ns, 'w:pPr');

            $paragraphBorder = $document->createElementNS($ns, 'w:pBdr');
            $bottomBorder = $document->createElementNS($ns, 'w:bottom');
            $bottomBorder->setAttributeNS($ns, 'w:val', 'single');
            $bottomBorder->setAttributeNS($ns, 'w:sz', '12');
            $bottomBorder->setAttributeNS($ns, 'w:space', '0');
            $bottomBorder->setAttributeNS($ns, 'w:color', 'auto');
            $paragraphBorder->appendChild($bottomBorder);
            $ruleProperties->appendChild($paragraphBorder);

            $ruleSpacing = $document->createElementNS($ns, 'w:spacing');
            $ruleSpacing->setAttributeNS($ns, 'w:before', '0');
            $ruleSpacing->setAttributeNS($ns, 'w:after', '0');
            $ruleSpacing->setAttributeNS($ns, 'w:line', '20');
            $ruleSpacing->setAttributeNS($ns, 'w:lineRule', 'exact');
            $ruleProperties->appendChild($ruleSpacing);

            $ruleParagraph->appendChild($ruleProperties);
            $paragraphs[$kopEndIndex]->parentNode->insertBefore(
                $ruleParagraph,
                $paragraphs[$kopEndIndex]->nextSibling
            );
        }
    }

    $sameHeadingParagraph = ($titleIndex !== null && $numberIndex === $titleIndex);
    $headingIndices = [];
    foreach ([$titleIndex, $numberIndex] as $index) {
        if ($index === null || in_array($index, $headingIndices, true)) {
            continue;
        }
        $headingIndices[] = $index;
    }

    foreach ($headingIndices as $index) {
        if (!isset($paragraphs[$index])) {
            continue;
        }
        $paragraph = $paragraphs[$index];
        $spacing = spt_layout_spacing($paragraph, $document);
        $safeLine = spt_layout_heading_line_twips($paragraph, 1.3);
        if ($sameHeadingParagraph) {
            $safeLine += 60;
        }
        // Set tepat agar line-height template yang besar tidak mendominasi
        // jarak antara judul dan nomor.
        $spacing->setAttributeNS($ns, 'w:line', (string)$safeLine);
        $spacing->setAttributeNS($ns, 'w:lineRule', 'auto');
        foreach (['before', 'after'] as $attribute) {
            if ($spacing->hasAttributeNS($ns, $attribute) && (int)$spacing->getAttributeNS($ns, $attribute) < 0) {
                $spacing->setAttributeNS($ns, 'w:' . $attribute, '0');
            }
        }
    }

    if ($numberIndex !== null && $numberIndex !== $titleIndex && isset($paragraphs[$numberIndex])) {
        $spacing = spt_layout_spacing($paragraphs[$numberIndex], $document);
        $spacing->setAttributeNS($ns, 'w:before', '60');
        if ($titleIndex !== null && isset($paragraphs[$titleIndex])) {
            $titleSpacing = spt_layout_spacing($paragraphs[$titleIndex], $document);
            $titleSpacing->setAttributeNS($ns, 'w:after', '0');
        }
    }

    // 3. Semua sel tabel rata atas.
    spt_layout_align_tables_top($document);

    $zip->addFromString('word/document.xml', $document->saveXML());
    $zip->close();
    return true;
}

function spt_layout_detect_dominant_font_size(DOMDocument $document)
{
    $xpath = new DOMXPath($document);
    $xpath->registerNamespace('w', spt_layout_w_ns());
    $counts = [];
    foreach ($xpath->query('//w:sz') as $node) {
        $value = (int)$node->getAttributeNS(spt_layout_w_ns(), 'val');
        if ($value > 0) {
            $counts[$value] = ($counts[$value] ?? 0) + 1;
        }
    }
    if (empty($counts)) {
        return 22;
    }
    arsort($counts);
    return (int)key($counts);
}

function spt_layout_scale_font_sizes(DOMDocument $document, $factor)
{
    if (abs($factor - 1.0) < 0.0001) {
        return;
    }

    $ns = spt_layout_w_ns();
    $xpath = new DOMXPath($document);
    $xpath->registerNamespace('w', $ns);
    foreach ($xpath->query('//w:sz | //w:szCs') as $node) {
        $value = (int)$node->getAttributeNS($ns, 'val');
        if ($value <= 0) {
            continue;
        }
        $newValue = max(2, min(3276, (int)round($value * $factor)));
        $node->setAttributeNS($ns, 'w:val', (string)$newValue);
    }
}

function spt_layout_set_page_margins(DOMDocument $document, array $settings)
{
    $ns = spt_layout_w_ns();
    $xpath = new DOMXPath($document);
    $xpath->registerNamespace('w', $ns);
    $margins = [
        'top' => spt_layout_mm_to_twips($settings['margin_top']),
        'bottom' => spt_layout_mm_to_twips($settings['margin_bottom']),
        'left' => spt_layout_mm_to_twips($settings['margin_left']),
        'right' => spt_layout_mm_to_twips($settings['margin_right']),
    ];

    foreach ($xpath->query('//w:sectPr') as $section) {
        $pgMar = spt_layout_child($section, 'pgMar');
        if (!$pgMar) {
            $pgMar = $document->createElementNS($ns, 'w:pgMar');
            spt_layout_insert_ordered($section, $pgMar, [
                'headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type',
                'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType',
                'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection',
                'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange',
            ]);
        }
        foreach ($margins as $side => $twips) {
            $pgMar->setAttributeNS($ns, 'w:' . $side, (string)$twips);
        }
    }
}

/**
 * Terapkan semua parameter layout ke file DOCX (in-place).
 * Dipakai untuk pipeline Word COM & LibreOffice agar hasil PDF ikut berubah.
 */
function spt_layout_apply_to_docx($docxPath, array $settings)
{
    if (!class_exists('ZipArchive') || !is_file($docxPath)) {
        return false;
    }

    $zip = new ZipArchive();
    if ($zip->open($docxPath) !== true) {
        return false;
    }

    try {
        $documentXml = $zip->getFromName('word/document.xml');
        $stylesXml = $zip->getFromName('word/styles.xml');
        $document = spt_layout_load_dom($documentXml);
        if (!$document) {
            $zip->close();
            return false;
        }

        $ns = spt_layout_w_ns();
        $xpath = new DOMXPath($document);
        $xpath->registerNamespace('w', $ns);

        // Ukuran dasar diambil dari paragraf isi surat (sama seperti pratinjau),
        // bukan dari hitungan semua run. Run kosong/spacer sering berukuran lain
        // dan membuat seluruh dokumen ikut membesar.
        $bodyDefaults = spt_layout_docx_body_defaults($documentXml, spt_layout_docx_styles($stylesXml));
        if (!empty($bodyDefaults['font_size'])) {
            $baseHalfPoints = (float)$bodyDefaults['font_size'] * 2;
        } else {
            $baseHalfPoints = (float)spt_layout_detect_dominant_font_size($document);
        }
        $targetHalfPoints = (float)$settings['font_size'] * 2;
        $factor = $baseHalfPoints > 0 ? ($targetHalfPoints / $baseHalfPoints) : 1.0;

        spt_layout_set_page_margins($document, $settings);
        spt_layout_scale_font_sizes($document, $factor);

        // Line-height: hanya diubah jika user memang mengubahnya dari default
        // template. Spacing yang sudah ada diskalakan agar jarak khusus (mis.
        // ruang gambar TTD) tetap proporsional.
        $baseLineHeight = spt_layout_detect_default_line_height($documentXml);
        $lineChanged = abs((float)$settings['line_height'] - $baseLineHeight) > 0.0001;
        if ($lineChanged) {
            $lineRatio = (float)$settings['line_height'] / max(0.1, $baseLineHeight);
            $lineTwips = (int)round(((float)$settings['line_height']) * 240);
            foreach ($xpath->query('//w:body//w:p') as $paragraph) {
                $spacing = spt_layout_spacing($paragraph, $document);
                if ($spacing->hasAttributeNS($ns, 'line')) {
                    $oldLine = (int)$spacing->getAttributeNS($ns, 'line');
                    if ($oldLine > 0) {
                        $spacing->setAttributeNS($ns, 'w:line', (string)max(1, (int)round($oldLine * $lineRatio)));
                    } else {
                        $spacing->setAttributeNS($ns, 'w:line', (string)$lineTwips);
                    }
                    if (!$spacing->hasAttributeNS($ns, 'lineRule')) {
                        $spacing->setAttributeNS($ns, 'w:lineRule', 'auto');
                    }
                } else {
                    $spacing->setAttributeNS($ns, 'w:line', (string)$lineTwips);
                    $spacing->setAttributeNS($ns, 'w:lineRule', 'auto');
                }
            }
        }

        $paragraphs = [];
        foreach ($xpath->query('//w:body//w:p') as $paragraph) {
            $paragraphs[] = $paragraph;
        }
        $paragraphCount = count($paragraphs);

        // Cari landmark index per blok
        $idxJudul = null;
        $idxNomor = null;
        $idxDasar = null;
        $idxMemerintahkan = null;
        $idxKepada = null;
        $idxUntuk = null;
        $ttdStart = null;
        $ttdEnd = null;

        $awaitingNomor = false;
        foreach ($paragraphs as $index => $paragraph) {
            $text = spt_layout_paragraph_text($paragraph);
            if ($text === '') {
                continue;
            }

            if ($idxJudul === null && stripos($text, 'SURAT TUGAS') !== false) {
                $idxJudul = $index;
                if (preg_match('/SURAT\s+TUGAS.*?Nomor\s*:/isu', $text)) {
                    $idxNomor = $index;
                } else {
                    $awaitingNomor = true;
                }
                continue;
            }

            if ($awaitingNomor) {
                $awaitingNomor = false;
                if (stripos($text, 'Nomor') !== false) {
                    $idxNomor = $index;
                }
                continue;
            }

            if ($idxDasar === null && preg_match('/^DASAR\b/iu', $text)) {
                $idxDasar = $index;
                continue;
            }

            if ($idxMemerintahkan === null && preg_match('/^MEMERINTAHKAN\b/iu', $text)) {
                $idxMemerintahkan = $index;
                continue;
            }

            if ($idxKepada === null && preg_match('/^KEPADA\b/iu', $text)) {
                $idxKepada = $index;
                continue;
            }

            if ($idxUntuk === null && preg_match('/^UNTUK\b/iu', $text)) {
                $idxUntuk = $index;
                continue;
            }

            if ($ttdStart === null) {
                $isPlaceLine = stripos($text, 'Bojonegoro') !== false
                    && stripos($text, 'Jalan') === false
                    && stripos($text, 'DINAS') === false
                    && stripos($text, 'KABUPATEN') === false;
                if ($isPlaceLine) {
                    $ttdStart = $index;
                }
            }
        }

        if ($ttdStart === null) {
            foreach ($paragraphs as $index => $paragraph) {
                $text = spt_layout_paragraph_text($paragraph);
                if (stripos($text, 'Kepala') !== false && stripos($text, 'Cabang Dinas') !== false) {
                    $ttdStart = max(0, $index - 1);
                    break;
                }
            }
        }

        if ($ttdStart !== null) {
            $ttdEnd = $ttdStart;
            $limit = min($paragraphCount - 1, $ttdStart + 15);
            for ($i = $ttdStart; $i <= $limit; $i++) {
                $text = spt_layout_paragraph_text($paragraphs[$i]);
                if (stripos($text, 'Lampiran') !== false && $i > $ttdStart) {
                    break;
                }
                $ttdEnd = $i;
                if (stripos($text, 'NIP') !== false) {
                    break;
                }
            }
        }

        // 1. BLOK KOP SURAT
        $kopBottomTwips = spt_layout_px_to_twips($settings['kop_bottom_spacing'] ?? 10);
        $kopLineTwips = spt_layout_px_to_twips($settings['kop_line_spacing'] ?? 0);

        // Kerapatan baris di dalam kop (0 = ikuti template).
        if ($kopLineTwips > 0 && $idxJudul !== null) {
            for ($k = 0; $k < $idxJudul; $k++) {
                if (!isset($paragraphs[$k]) || spt_layout_paragraph_text($paragraphs[$k]) === '') {
                    continue;
                }
                $spacing = spt_layout_spacing($paragraphs[$k], $document);
                $currentLine = $spacing->hasAttributeNS($ns, 'line')
                    ? (int)$spacing->getAttributeNS($ns, 'line')
                    : 0;
                $baseLine = max($currentLine, spt_layout_heading_line_twips($paragraphs[$k], 1.0));
                $spacing->setAttributeNS($ns, 'w:line', (string)($baseLine + $kopLineTwips));
                $spacing->setAttributeNS($ns, 'w:lineRule', 'auto');
            }
        }

        $idxKopEnd = $idxJudul !== null ? max(0, $idxJudul - 1) : null;
        if ($idxKopEnd !== null && isset($paragraphs[$idxKopEnd])) {
            $spacing = spt_layout_spacing($paragraphs[$idxKopEnd], $document);
            $spacing->setAttributeNS($ns, 'w:after', (string)$kopBottomTwips);
        }
        if ($idxJudul !== null) {
            for ($k = ($idxKopEnd !== null ? $idxKopEnd : 0); $k < $idxJudul; $k++) {
                if (isset($paragraphs[$k]) && spt_layout_paragraph_text($paragraphs[$k]) === '') {
                    spt_layout_collapse_paragraph($paragraphs[$k], $document);
                }
            }
        }

        // 2. BLOK JUDUL & NOMOR SURAT
        if ($idxJudul !== null && isset($paragraphs[$idxJudul])) {
            $judulTopTwips = spt_layout_px_to_twips($settings['judul_top_spacing'] ?? 0);
            $nomorTwips = spt_layout_px_to_twips($settings['nomor_surat_spacing'] ?? 4);
            $spacing = spt_layout_spacing($paragraphs[$idxJudul], $document);
            $effectiveJudulTop = max(0, $kopBottomTwips + $judulTopTwips);
            $spacing->setAttributeNS($ns, 'w:before', (string)$effectiveJudulTop);
            $spacing->setAttributeNS($ns, 'w:after', '0');
        }
        if ($idxJudul !== null && $idxNomor !== null) {
            for ($k = $idxJudul + 1; $k < $idxNomor; $k++) {
                if (isset($paragraphs[$k]) && spt_layout_paragraph_text($paragraphs[$k]) === '') {
                    spt_layout_collapse_paragraph($paragraphs[$k], $document);
                }
            }
        }
        if ($idxNomor !== null && isset($paragraphs[$idxNomor])) {
            $nomorTwips = spt_layout_px_to_twips($settings['nomor_surat_spacing'] ?? 4);
            $judulBottomTwips = spt_layout_px_to_twips($settings['judul_bottom_spacing'] ?? 10);
            $spacing = spt_layout_spacing($paragraphs[$idxNomor], $document);
            $spacing->setAttributeNS($ns, 'w:before', (string)$nomorTwips);
            $spacing->setAttributeNS($ns, 'w:after', (string)$judulBottomTwips);
            if ($idxDasar !== null) {
                for ($k = $idxNomor + 1; $k < $idxDasar; $k++) {
                    if (isset($paragraphs[$k]) && spt_layout_paragraph_text($paragraphs[$k]) === '') {
                        spt_layout_collapse_paragraph($paragraphs[$k], $document);
                    }
                }
            }
        }

        // Judul "SURAT TUGAS" dan "Nomor" harus selalu punya line-height aman
        // agar tidak saling menimpa meskipun jarak antar keduanya disetel minimum.
        $headingMultiplier = max(1.3, (float)($settings['line_height'] ?? 1.0));
        $sameHeadingParagraph = ($idxJudul !== null && $idxNomor === $idxJudul);
        $headingGapTwips = spt_layout_px_to_twips($settings['nomor_surat_spacing'] ?? 4);
        $headingIndices = [];
        foreach ([$idxJudul, $idxNomor] as $idxHeading) {
            if ($idxHeading === null || in_array($idxHeading, $headingIndices, true)) {
                continue;
            }
            $headingIndices[] = $idxHeading;
        }
        foreach ($headingIndices as $idxHeading) {
            if (!isset($paragraphs[$idxHeading])) {
                continue;
            }
            $headingSpacing = spt_layout_spacing($paragraphs[$idxHeading], $document);
            $headingLineTwips = spt_layout_heading_line_twips($paragraphs[$idxHeading], $headingMultiplier);
            // Jika judul & nomor berada di paragraf yang sama (soft line break),
            // line spacing adalah satu-satunya pengendali jarak keduanya.
            if ($sameHeadingParagraph) {
                $headingLineTwips += $headingGapTwips;
            }
            // Set tepat (bukan hanya minimum) agar line-height template yang
            // terlalu besar tidak mendominasi jarak judul ke nomor.
            $headingSpacing->setAttributeNS($ns, 'w:line', (string)$headingLineTwips);
            $headingSpacing->setAttributeNS($ns, 'w:lineRule', 'auto');
        }

        // 3. BLOK DASAR
        if ($idxDasar !== null) {
            $dasarTopTwips = spt_layout_px_to_twips($settings['dasar_top_spacing'] ?? 6);
            $dasarItemTwips = spt_layout_px_to_twips($settings['dasar_item_spacing'] ?? 3);
            $dasarBottomTwips = spt_layout_px_to_twips($settings['dasar_bottom_spacing'] ?? 8);

            $limitDasar = $idxMemerintahkan !== null ? max($idxDasar, $idxMemerintahkan - 1) : $idxDasar;
            $endDasar = $idxDasar;
            for ($k = $limitDasar; $k >= $idxDasar; $k--) {
                if (isset($paragraphs[$k]) && spt_layout_paragraph_text($paragraphs[$k]) !== '') {
                    $endDasar = $k;
                    break;
                }
            }

            for ($i = $idxDasar; $i <= $endDasar; $i++) {
                if (!isset($paragraphs[$i])) continue;
                if (spt_layout_paragraph_text($paragraphs[$i]) === '') {
                    spt_layout_collapse_paragraph($paragraphs[$i], $document);
                    continue;
                }
                $spacing = spt_layout_spacing($paragraphs[$i], $document);
                if ($i === $idxDasar) {
                    $spacing->setAttributeNS($ns, 'w:before', (string)$dasarTopTwips);
                }
                if ($i === $endDasar) {
                    $spacing->setAttributeNS($ns, 'w:after', (string)$dasarBottomTwips);
                } else {
                    $spacing->setAttributeNS($ns, 'w:after', (string)$dasarItemTwips);
                }
            }
            if ($idxMemerintahkan !== null) {
                for ($k = $endDasar + 1; $k < $idxMemerintahkan; $k++) {
                    if (isset($paragraphs[$k]) && spt_layout_paragraph_text($paragraphs[$k]) === '') {
                        spt_layout_collapse_paragraph($paragraphs[$k], $document);
                    }
                }
            }
        }

        // 4. BLOK MEMERINTAHKAN
        if ($idxMemerintahkan !== null && isset($paragraphs[$idxMemerintahkan])) {
            $memSpacingTwips = spt_layout_px_to_twips($settings['memerintahkan_spacing'] ?? 8);
            $spacing = spt_layout_spacing($paragraphs[$idxMemerintahkan], $document);
            $spacing->setAttributeNS($ns, 'w:before', (string)$memSpacingTwips);
            $spacing->setAttributeNS($ns, 'w:after', (string)$memSpacingTwips);
        }
        $targetPegawaiStart = $idxKepada !== null ? $idxKepada : ($idxMemerintahkan !== null ? $idxMemerintahkan + 1 : null);
        if ($idxMemerintahkan !== null && $targetPegawaiStart !== null) {
            for ($k = $idxMemerintahkan + 1; $k < $targetPegawaiStart; $k++) {
                if (isset($paragraphs[$k]) && spt_layout_paragraph_text($paragraphs[$k]) === '') {
                    spt_layout_collapse_paragraph($paragraphs[$k], $document);
                }
            }
        }

        // 5. BLOK PEGAWAI
        $pegawaiTopTwips = spt_layout_px_to_twips($settings['pegawai_top_spacing'] ?? 6);
        $pegawaiGapTwips = spt_layout_px_to_twips($settings['pegawai_gap_spacing'] ?? 8);
        $pegawaiLineTwips = spt_layout_px_to_twips($settings['pegawai_line_spacing'] ?? 2);

        if ($idxKepada !== null && isset($paragraphs[$idxKepada])) {
            $spacing = spt_layout_spacing($paragraphs[$idxKepada], $document);
            $spacing->setAttributeNS($ns, 'w:before', (string)$pegawaiTopTwips);
        }

        $startPegawai = $idxKepada !== null ? $idxKepada + 1 : ($idxMemerintahkan !== null ? $idxMemerintahkan + 1 : null);
        $endPegawai = $idxUntuk !== null ? max($startPegawai ?? 0, $idxUntuk - 1) : ($ttdStart !== null ? $ttdStart - 1 : null);

        if ($startPegawai !== null && $endPegawai !== null && $startPegawai <= $endPegawai) {
            $pegawaiCountSeen = 0;
            for ($i = $startPegawai; $i <= $endPegawai; $i++) {
                if (!isset($paragraphs[$i])) continue;
                $t = spt_layout_paragraph_text($paragraphs[$i]);
                if ($t === '') {
                    spt_layout_collapse_paragraph($paragraphs[$i], $document);
                    continue;
                }

                $isNewPegawai = preg_match('/^(?:\d+\.|\bNama\b|\$\{nama[2-9]|\$\{nama1[0-9])/iu', $t);
                if ($isNewPegawai) {
                    $pegawaiCountSeen++;
                }

                $spacing = spt_layout_spacing($paragraphs[$i], $document);
                if ($isNewPegawai && $pegawaiCountSeen > 1) {
                    $spacing->setAttributeNS($ns, 'w:before', (string)$pegawaiGapTwips);
                }
                $spacing->setAttributeNS($ns, 'w:after', (string)$pegawaiLineTwips);
            }
        }

        if ($endPegawai !== null && $idxUntuk !== null) {
            for ($k = $endPegawai + 1; $k < $idxUntuk; $k++) {
                if (isset($paragraphs[$k]) && spt_layout_paragraph_text($paragraphs[$k]) === '') {
                    spt_layout_collapse_paragraph($paragraphs[$k], $document);
                }
            }
        }

        // Sesuaikan baris tabel pegawai jika ada
        foreach ($xpath->query('//w:tbl//w:tr') as $tr) {
            $trText = trim($tr->textContent);
            if (stripos($trText, 'NAMA') !== false || stripos($trText, 'NIP') !== false || stripos($trText, '${nama') !== false) {
                foreach ($tr->getElementsByTagNameNS($ns, 'p') as $cellP) {
                    $sp = spt_layout_spacing($cellP, $document);
                    $sp->setAttributeNS($ns, 'w:before', (string)$pegawaiLineTwips);
                    $sp->setAttributeNS($ns, 'w:after', (string)$pegawaiLineTwips);
                }
            }
        }

        // 6. BLOK UNTUK
        if ($idxUntuk !== null) {
            $untukTopTwips = spt_layout_px_to_twips($settings['untuk_top_spacing'] ?? 6);
            $untukItemTwips = spt_layout_px_to_twips($settings['untuk_item_spacing'] ?? 3);
            $untukBottomTwips = spt_layout_px_to_twips($settings['untuk_bottom_spacing'] ?? 10);

            $limitUntuk = $ttdStart !== null ? max($idxUntuk, $ttdStart - 1) : $paragraphCount - 1;

            // Butir terakhir yang benar-benar berisi teks. Paragraf kosong di
            // ekor blok tidak boleh "mencuri" jarak bawah.
            $endUntuk = $idxUntuk;
            for ($i = $limitUntuk; $i >= $idxUntuk; $i--) {
                if (isset($paragraphs[$i]) && spt_layout_paragraph_text($paragraphs[$i]) !== '') {
                    $endUntuk = $i;
                    break;
                }
            }

            // Blok UNTUK mencakup paragraf label ("UNTUK :") beserta butir
            // pertama di sebelahnya, lalu seluruh butir tugas berikutnya.
            for ($i = $idxUntuk; $i <= $endUntuk; $i++) {
                if (!isset($paragraphs[$i])) continue;
                if (spt_layout_paragraph_text($paragraphs[$i]) === '') {
                    spt_layout_collapse_paragraph($paragraphs[$i], $document);
                    continue;
                }
                $spacing = spt_layout_spacing($paragraphs[$i], $document);
                if ($i === $idxUntuk) {
                    $spacing->setAttributeNS($ns, 'w:before', (string)$untukTopTwips);
                }
                if ($i === $endUntuk) {
                    $spacing->setAttributeNS($ns, 'w:after', (string)$untukBottomTwips);
                } else {
                    $spacing->setAttributeNS($ns, 'w:after', (string)$untukItemTwips);
                }
            }
            if ($ttdStart !== null) {
                for ($k = $endUntuk + 1; $k < $ttdStart; $k++) {
                    if (isset($paragraphs[$k]) && spt_layout_paragraph_text($paragraphs[$k]) === '') {
                        spt_layout_collapse_paragraph($paragraphs[$k], $document);
                    }
                }
            }
        }

        // 7. BLOK TANDA TANGAN
        if ($ttdStart !== null) {
            $ttdOffsetTwips = spt_layout_px_to_twips($settings['ttd_offset'] ?? 10);
            $ttdDateToJabatanTwips = spt_layout_px_to_twips($settings['ttd_date_to_jabatan'] ?? 4);
            $ttdSpaceHeightTwips = spt_layout_px_to_twips($settings['ttd_space_height'] ?? 60);
            $ttdNameNipTwips = spt_layout_px_to_twips($settings['ttd_name_nip_spacing'] ?? 4);
            $ttdIndentTwips = spt_layout_px_to_twips($settings['ttd_x'] ?? 0);

            // Baris Tempat & Tanggal
            if (isset($paragraphs[$ttdStart])) {
                $spacing = spt_layout_spacing($paragraphs[$ttdStart], $document);
                $spacing->setAttributeNS($ns, 'w:before', (string)$ttdOffsetTwips);
                $spacing->setAttributeNS($ns, 'w:after', (string)$ttdDateToJabatanTwips);
            }

            // Identifikasi baris jabatan terakhir dan baris nama penandatangan
            $lastJabatanIdx = null;
            $nameIdx = null;

            for ($i = $ttdStart + 1; $i <= $ttdEnd; $i++) {
                if (!isset($paragraphs[$i])) continue;
                $t = spt_layout_paragraph_text($paragraphs[$i]);
                if (stripos($t, 'Kepala') !== false || stripos($t, 'Cabang Dinas') !== false || stripos($t, 'Wilayah') !== false) {
                    $lastJabatanIdx = $i;
                } elseif ($nameIdx === null && (stripos($t, 'NIP') !== false || stripos($t, '${penandatangan') !== false || preg_match('/[A-Z\s\.,]{4,}/u', $t))) {
                    if ($lastJabatanIdx !== null && $i > $lastJabatanIdx && $t !== '') {
                        $nameIdx = $i;
                    }
                }
            }

            for ($i = $ttdStart + 1; $i <= $ttdEnd; $i++) {
                if (!isset($paragraphs[$i])) continue;
                $t = spt_layout_paragraph_text($paragraphs[$i]);

                // Paragraf kosong antara jabatan dan nama diciutkan agar ttd_space_height menjadi pengendali utama
                if ($t === '') {
                    spt_layout_collapse_paragraph($paragraphs[$i], $document);
                    continue;
                }

                $spacing = spt_layout_spacing($paragraphs[$i], $document);

                if ($lastJabatanIdx !== null && $i === $lastJabatanIdx) {
                    $spacing->setAttributeNS($ns, 'w:after', (string)$ttdSpaceHeightTwips);
                } elseif (stripos($t, 'NIP') !== false || ($nameIdx !== null && $i >= $nameIdx)) {
                    $spacing->setAttributeNS($ns, 'w:before', (string)$ttdNameNipTwips);
                }
            }

            if ($ttdIndentTwips !== 0) {
                for ($i = $ttdStart; $i <= $ttdEnd; $i++) {
                    if (!isset($paragraphs[$i])) continue;
                    $ind = spt_layout_indentation($paragraphs[$i], $document);
                    $current = $ind->hasAttributeNS($ns, 'left')
                        ? (int)$ind->getAttributeNS($ns, 'left')
                        : 0;
                    $ind->setAttributeNS($ns, 'w:left', (string)($current + $ttdIndentTwips));
                }
            }

            if (!empty($settings['lock_ttd'])) {
                for ($i = $ttdStart; $i <= $ttdEnd; $i++) {
                    if (!isset($paragraphs[$i])) continue;
                    spt_layout_add_keep_flag($paragraphs[$i], 'keepLines', $document);
                    if ($i < $ttdEnd) {
                        spt_layout_add_keep_flag($paragraphs[$i], 'keepNext', $document);
                    }
                }
            }
        }

        // Semua sel tabel (biodata pegawai, dasar, lampiran) rata atas agar
        // titik dua dan nomor urut tetap sejajar di baris pertama.
        spt_layout_align_tables_top($document);

        $styles = $stylesXml !== false ? spt_layout_load_dom($stylesXml) : null;
        if ($styles) {
            spt_layout_scale_font_sizes($styles, $factor);
        }

        $zip->addFromString('word/document.xml', $document->saveXML());
        if ($styles) {
            $zip->addFromString('word/styles.xml', $styles->saveXML());
        }
        $zip->close();
        return true;
    } catch (Throwable $e) {
        error_log('Layout adjuster DOCX error: ' . $e->getMessage());
        try {
            $zip->close();
        } catch (Throwable $e2) {
            // ignore
        }
        return false;
    }
}

/**
 * Baca style paragraf dari word/styles.xml agar ukuran huruf yang
 * diwarisi dari style (mis. Normal) tetap terhitung.
 */
function spt_layout_docx_styles($stylesXml)
{
    $styles = [];
    if ($stylesXml === false || trim($stylesXml) === '') {
        return $styles;
    }
    $doc = new DOMDocument();
    if (!@$doc->loadXML($stylesXml)) {
        return $styles;
    }
    $wNs = spt_layout_w_ns();
    $xp = new DOMXPath($doc);
    $xp->registerNamespace('w', $wNs);
    foreach ($xp->query('//w:style[@w:type="paragraph"]') as $style) {
        $id = $style->getAttributeNS($wNs, 'styleId');
        if ($id === '') {
            continue;
        }
        $basedOn = '';
        $based = $xp->query('w:basedOn', $style)->item(0);
        if ($based) {
            $basedOn = $based->getAttributeNS($wNs, 'val');
        }
        $size = 0;
        $sz = $xp->query('w:rPr/w:sz', $style)->item(0);
        if ($sz) {
            $size = (int)$sz->getAttributeNS($wNs, 'val');
        }
        $line = 0;
        $spacing = $xp->query('w:pPr/w:spacing', $style)->item(0);
        if ($spacing) {
            $line = (int)$spacing->getAttributeNS($wNs, 'line');
        }
        $styles[$id] = ['based' => $basedOn, 'sz' => $size, 'line' => $line];
    }
    return $styles;
}

function spt_layout_style_value(array $styles, $styleId, $field, $depth = 0)
{
    if ($styleId === '' || $depth > 20 || !isset($styles[$styleId])) {
        return 0;
    }
    $style = $styles[$styleId];
    $value = (int)($style[$field] ?? 0);
    if ($value > 0) {
        return $value;
    }
    return spt_layout_style_value($styles, (string)($style['based'] ?? ''), $field, $depth + 1);
}

/**
 * Default huruf & jarak baris diambil dari paragraf isi surat (level body),
 * bukan dari sel tabel/kop yang ukurannya berbeda.
 */
function spt_layout_docx_body_defaults($documentXml, array $styles)
{
    $doc = new DOMDocument();
    if (!@$doc->loadXML($documentXml)) {
        return null;
    }
    $wNs = spt_layout_w_ns();
    $xp = new DOMXPath($doc);
    $xp->registerNamespace('w', $wNs);

    $sizeCounts = [];
    $lineCounts = [];
    foreach ($xp->query('//w:body/w:p') as $p) {
        $text = '';
        foreach ($xp->query('.//w:t', $p) as $t) {
            $text .= $t->textContent;
        }
        if (trim($text) === '') {
            continue;
        }

        $styleId = 'Normal';
        $pStyle = $xp->query('w:pPr/w:pStyle', $p)->item(0);
        if ($pStyle) {
            $styleId = $pStyle->getAttributeNS($wNs, 'val');
        }

        $size = 0;
        foreach ($xp->query('.//w:r/w:rPr/w:sz', $p) as $sz) {
            $value = (int)$sz->getAttributeNS($wNs, 'val');
            if ($value > 0) {
                $size = $value;
                break;
            }
        }
        if ($size <= 0) {
            $size = spt_layout_style_value($styles, $styleId, 'sz');
        }
        if ($size <= 0) {
            $size = 22;
        }
        $sizeCounts[$size] = ($sizeCounts[$size] ?? 0) + 1;

        $line = 0;
        $spacing = $xp->query('w:pPr/w:spacing', $p)->item(0);
        if ($spacing) {
            $line = (int)$spacing->getAttributeNS($wNs, 'line');
        }
        if ($line <= 0) {
            $line = spt_layout_style_value($styles, $styleId, 'line');
        }
        if ($line >= 240) {
            $lineCounts[$line] = ($lineCounts[$line] ?? 0) + 1;
        }
    }

    if (empty($sizeCounts) && empty($lineCounts)) {
        return null;
    }

    $result = ['font_size' => 0.0, 'line_height' => 0.0];
    if (!empty($sizeCounts)) {
        arsort($sizeCounts);
        $result['font_size'] = ((int)key($sizeCounts)) / 2;
    }
    if (!empty($lineCounts)) {
        arsort($lineCounts);
        $result['line_height'] = ((int)key($lineCounts)) / 240;
    }
    return $result;
}

function spt_layout_detect_default_line_height($documentXml)
{
    // Utamakan paragraf isi surat; paragraf kosong dan sel tabel sering
    // memakai jarak baris berbeda dan membuat deteksi meleset.
    $doc = new DOMDocument();
    if (@$doc->loadXML($documentXml)) {
        $wNs = spt_layout_w_ns();
        $xp = new DOMXPath($doc);
        $xp->registerNamespace('w', $wNs);
        $bodyCounts = [];
        foreach ($xp->query('//w:body/w:p') as $p) {
            $text = '';
            foreach ($xp->query('.//w:t', $p) as $t) {
                $text .= $t->textContent;
            }
            if (trim($text) === '') {
                continue;
            }
            $spacing = $xp->query('w:pPr/w:spacing', $p)->item(0);
            if (!$spacing) {
                continue;
            }
            $line = (int)$spacing->getAttributeNS($wNs, 'line');
            if ($line >= 240) {
                $bodyCounts[$line] = ($bodyCounts[$line] ?? 0) + 1;
            }
        }
        if (!empty($bodyCounts)) {
            arsort($bodyCounts);
            return spt_layout_round(
                spt_layout_clamp(((int)key($bodyCounts)) / 240, 1.0, 2.0),
                0.05
            );
        }
    }

    $counts = [];
    if (preg_match_all('/<w:spacing\b[^>]*w:line="(\d+)"[^>]*w:lineRule="auto"[^>]*>/i', $documentXml, $matches)) {
        foreach ($matches[1] as $line) {
            $line = (int)$line;
            if ($line > 0) {
                $counts[$line] = ($counts[$line] ?? 0) + 1;
            }
        }
    }
    if (empty($counts) && preg_match_all('/<w:spacing\b[^>]*w:line="(\d+)"/i', $documentXml, $matches)) {
        foreach ($matches[1] as $line) {
            $line = (int)$line;
            if ($line > 0) {
                $counts[$line] = ($counts[$line] ?? 0) + 1;
            }
        }
    }
    if (empty($counts)) {
        return 1.0;
    }
    arsort($counts);
    $dominant = (int)key($counts);
    return spt_layout_round(
        spt_layout_clamp($dominant / 240, 1.0, 2.0),
        0.05
    );
}

/**
 * Cari nilai default (margin & font) dari template DOCX supaya slider
 * menampilkan angka yang mendekati template asli.
 */
function spt_layout_read_docx_defaults($docxPath)
{
    $defaults = spt_layout_default_values();
    if (!class_exists('ZipArchive') || !is_file($docxPath)) {
        return $defaults;
    }

    $zip = new ZipArchive();
    if ($zip->open($docxPath) !== true) {
        return $defaults;
    }

    $documentXml = $zip->getFromName('word/document.xml');
    $stylesXml = $zip->getFromName('word/styles.xml');
    $zip->close();
    if ($documentXml === false) {
        return $defaults;
    }

    if (preg_match('/<w:pgMar\b[^>]*>/i', $documentXml, $match)) {
        foreach (['top', 'bottom', 'left', 'right'] as $side) {
            if (preg_match('/w:' . $side . '="(-?\d+)"/i', $match[0], $value)) {
                $defaults['margin_' . $side] = round(((int)$value[1]) * 25.4 / 1440, 1);
            }
        }
    }

    $bodyDefaults = spt_layout_docx_body_defaults(
        $documentXml,
        spt_layout_docx_styles($stylesXml)
    );
    if ($bodyDefaults && $bodyDefaults['font_size'] > 0) {
        $defaults['font_size'] = spt_layout_round($bodyDefaults['font_size'], 0.5);
    } else {
        $counts = [];
        if (preg_match_all('/<w:sz\s+w:val="(\d+)"/i', $documentXml, $matches)) {
            foreach ($matches[1] as $halfPoint) {
                $halfPoint = (int)$halfPoint;
                if ($halfPoint > 0) {
                    $counts[$halfPoint] = ($counts[$halfPoint] ?? 0) + 1;
                }
            }
            if (!empty($counts)) {
                arsort($counts);
                $defaults['font_size'] = spt_layout_round(key($counts) / 2, 0.5);
            }
        }
    }

    $defaults['line_height'] = ($bodyDefaults && $bodyDefaults['line_height'] > 0)
        ? spt_layout_round($bodyDefaults['line_height'], 0.05)
        : spt_layout_detect_default_line_height($documentXml);

    // Samakan dengan batas slider agar UI tidak menampilkan nilai di luar range.
    foreach (spt_layout_ranges() as $key => $range) {
        if (isset($defaults[$key])) {
            $defaults[$key] = spt_layout_round(
                spt_layout_clamp($defaults[$key], $range[0], $range[1]),
                $range[2]
            );
        }
    }

    return $defaults;
}

/**
 * CSS dinamis untuk fallback Dompdf. Blok ini disuntikkan ke HTML
 * sebelum renderer memanggil $dompdf->loadHtml().
 */
function spt_layout_build_css(array $settings)
{
    $mt = number_format((float)$settings['margin_top'], 2, '.', '');
    $mb = number_format((float)$settings['margin_bottom'], 2, '.', '');
    $ml = number_format((float)$settings['margin_left'], 2, '.', '');
    $mr = number_format((float)$settings['margin_right'], 2, '.', '');
    $font = number_format((float)$settings['font_size'], 2, '.', '');
    $line = number_format((float)$settings['line_height'], 2, '.', '');

    $kopLine = max(0, (int)($settings['kop_line_spacing'] ?? 0));
    $kopBottom = max(0, (int)($settings['kop_bottom_spacing'] ?? 10));
    $judulTop = (int)($settings['judul_top_spacing'] ?? 0);
    $nomorSpacing = max(0, (int)($settings['nomor_surat_spacing'] ?? 4));
    $judulBottom = max(0, (int)($settings['judul_bottom_spacing'] ?? 10));
    $dasarTop = max(0, (int)($settings['dasar_top_spacing'] ?? 6));
    $dasarItem = max(0, (int)($settings['dasar_item_spacing'] ?? 3));
    $dasarBottom = max(0, (int)($settings['dasar_bottom_spacing'] ?? 8));
    $memSpacing = max(0, (int)($settings['memerintahkan_spacing'] ?? 8));
    $pegawaiTop = max(0, (int)($settings['pegawai_top_spacing'] ?? 6));
    $pegawaiGap = max(0, (int)($settings['pegawai_gap_spacing'] ?? 8));
    $pegawaiLine = max(0, (int)($settings['pegawai_line_spacing'] ?? 2));
    $untukTop = max(0, (int)($settings['untuk_top_spacing'] ?? 6));
    $untukItem = max(0, (int)($settings['untuk_item_spacing'] ?? 3));
    $untukBottom = max(0, (int)($settings['untuk_bottom_spacing'] ?? 10));
    $ttdOffset = max(0, (int)($settings['ttd_offset'] ?? 10));
    $ttdX = (int)($settings['ttd_x'] ?? 0);
    $ttdDateToJabatan = max(0, (int)($settings['ttd_date_to_jabatan'] ?? 4));
    $ttdSpace = max(20, (int)($settings['ttd_space_height'] ?? 60));
    $ttdNameNip = max(0, (int)($settings['ttd_name_nip_spacing'] ?? 4));

    $css = '<style>' . PHP_EOL;
    $css .= '@page { margin: ' . $mt . 'mm ' . $mr . 'mm ' . $mb . 'mm ' . $ml . 'mm; }' . PHP_EOL;
    $css .= 'body { font-size: ' . $font . 'pt; line-height: ' . $line . '; }' . PHP_EOL;
    $css .= 'p, td, th, li, span, h1, h2, h3, h4, h5, h6 { font-size: ' . $font . 'pt !important; line-height: ' . $line . ' !important; }' . PHP_EOL;
    $css .= 'table td, table th { vertical-align: top !important; }' . PHP_EOL;
    if ($kopLine > 0) {
        $kopLineHeightPt = ((float)$settings['font_size'] * (float)$settings['line_height']) + ($kopLine * 0.75);
        $css .= '.spt-kop-line, .spt-kop p, .spt-kop td { line-height: ' . number_format($kopLineHeightPt, 2, '.', '') . 'pt !important; }' . PHP_EOL;
    }
    $css .= 'table.spt-kop, .spt-kop { margin-bottom: ' . $kopBottom . 'px !important; }' . PHP_EOL;
    $css .= 'h1.spt-title, .spt-title { display: block !important; margin-top: ' . $judulTop . 'px !important; margin-bottom: 0 !important; line-height: 1.3 !important; }' . PHP_EOL;
    $css .= 'h5.spt-nomor, .spt-nomor { display: block !important; margin-top: ' . $nomorSpacing . 'px !important; margin-bottom: ' . $judulBottom . 'px !important; line-height: 1.3 !important; }' . PHP_EOL;
    $css .= '.spt-dasar { margin-top: ' . $dasarTop . 'px !important; margin-bottom: ' . $dasarItem . 'px !important; }' . PHP_EOL;
    $css .= '.spt-dasar-last { margin-bottom: ' . $dasarBottom . 'px !important; }' . PHP_EOL;
    $css .= '.spt-memerintahkan { margin-top: ' . $memSpacing . 'px !important; margin-bottom: ' . $memSpacing . 'px !important; }' . PHP_EOL;
    $css .= '.spt-kepada { margin-top: ' . $pegawaiTop . 'px !important; }' . PHP_EOL;
    $css .= '.spt-pegawai-gap { margin-top: ' . $pegawaiGap . 'px !important; }' . PHP_EOL;
    $css .= '.spt-pegawai-line { margin-bottom: ' . $pegawaiLine . 'px !important; }' . PHP_EOL;
    $css .= '.spt-untuk { margin-top: ' . $untukTop . 'px !important; margin-bottom: ' . $untukItem . 'px !important; }' . PHP_EOL;
    $css .= '.spt-untuk-last { margin-bottom: ' . $untukBottom . 'px !important; }' . PHP_EOL;
    $css .= '.spt-ttd-date { margin-top: ' . $ttdOffset . 'px !important; margin-bottom: ' . $ttdDateToJabatan . 'px !important; margin-left: ' . $ttdX . 'px !important; }' . PHP_EOL;
    $css .= '.spt-ttd-jabatan { margin-bottom: ' . $ttdSpace . 'px !important; margin-left: ' . $ttdX . 'px !important; }' . PHP_EOL;
    $css .= '.spt-ttd-name { margin-top: ' . $ttdNameNip . 'px !important; margin-left: ' . $ttdX . 'px !important; }' . PHP_EOL;
    if (!empty($settings['lock_ttd'])) {
        $css .= '.spt-ttd, .spt-ttd-date, .spt-ttd-jabatan, .spt-ttd-name, table { page-break-inside: avoid !important; }' . PHP_EOL;
    }
    $css .= '</style>';

    return $css;
}

/**
 * Tambahkan class penanda pada HTML hasil PHPWord agar CSS layout bisa
 * menyasar judul, bagian, dan blok tanda tangan. Lalu suntikkan <style>.
 */
function spt_layout_inject_html($html, array $settings)
{
    // Rapikan typo kop surat pada HTML sebelum CSS layout disuntikkan.
    $html = spt_fix_common_typos($html);

    // Beri class spt-kop pada tabel kop surat (tabel pertama yang memuat teks kop)
    $html = preg_replace_callback(
        '/<table\b([^>]*)>(.*?)<\/table>/is',
        function ($match) {
            $attrs = $match[1];
            $inner = $match[2];
            $text = strip_tags($inner);
            if (stripos($text, 'PEMERINTAH') !== false || stripos($text, 'KEHUTANAN') !== false || stripos($text, 'CABANG DINAS') !== false) {
                if (preg_match('/class="([^"]*)"/i', $attrs, $cm)) {
                    $attrs = preg_replace('/class="[^"]*"/i', 'class="' . trim($cm[1] . ' spt-kop') . '"', $attrs, 1);
                } else {
                    $attrs .= ' class="spt-kop"';
                }
            }
            return '<table' . $attrs . '>' . $inner . '</table>';
        },
        $html,
        1
    );

    // Target paragraf dan heading (hindari tag div agar pembungkus halaman tidak terpengaruh)
    $html = preg_replace_callback(
        '/<(p|h[1-6]|li)\b([^>]*)>(.*?)<\/\1>/is',
        function ($match) {
            $tag = $match[1];
            $attributes = $match[2];
            $inner = $match[3];
            $text = trim(preg_replace('/\s+/u', ' ', strip_tags($inner)));
            if ($text === '') {
                return $match[0];
            }

            $classes = [];
            if (stripos($text, 'PEMERINTAH PROVINSI') !== false
                || stripos($text, 'DINAS KEHUTANAN') !== false
                || (stripos($text, 'CABANG DINAS') !== false && stripos($text, 'Kepala') === false)
                || stripos($text, 'Jalan Hayam Wuruk') !== false
                || stripos($text, 'Telepon (0353)') !== false) {
                $classes[] = 'spt-kop-line';
            }
            if (stripos($text, 'SURAT TUGAS') !== false) {
                $classes[] = 'spt-title';
            }
            if (stripos($text, 'Nomor :') !== false || stripos($text, 'Nomor:') !== false) {
                $classes[] = 'spt-nomor';
            }
            if (preg_match('/^DASAR\b/iu', $text)) {
                $classes[] = 'spt-dasar';
            }
            if (preg_match('/^MEMERINTAHKAN\b/iu', $text)) {
                $classes[] = 'spt-memerintahkan';
            }
            if (preg_match('/^KEPADA\b/iu', $text)) {
                $classes[] = 'spt-kepada';
            }
            if (preg_match('/^UNTUK\b/iu', $text)) {
                $classes[] = 'spt-untuk';
            }
            if (stripos($text, 'Bojonegoro') !== false && stripos($text, 'DINAS') === false && stripos($text, 'KABUPATEN') === false) {
                $classes[] = 'spt-ttd-date';
            } elseif (stripos($text, 'Kepala') !== false && stripos($text, 'Cabang Dinas') !== false) {
                $classes[] = 'spt-ttd-jabatan';
            } elseif (stripos($text, 'NIP') !== false) {
                $classes[] = 'spt-ttd-name';
            }

            if (empty($classes)) {
                return $match[0];
            }

            if (preg_match('/class="([^"]*)"/i', $attributes, $classMatch)) {
                $attributes = preg_replace(
                    '/class="[^"]*"/i',
                    'class="' . trim($classMatch[1] . ' ' . implode(' ', $classes)) . '"',
                    $attributes,
                    1
                );
            } else {
                $attributes .= ' class="' . implode(' ', $classes) . '"';
            }

            return '<' . $tag . $attributes . '>' . $inner . '</' . $tag . '>';
        },
        $html
    );

    $css = spt_layout_build_css($settings);
    if (stripos($html, '</head>') !== false) {
        return preg_replace('/<\/head>/i', $css . '</head>', $html, 1);
    }
    return $css . $html;
}

/**
 * Hapus kata berulang yang tidak disengaja (mis. "tentang tentang").
 */
function spt_clean_repeated_words($text)
{
    $text = (string)$text;
    if ($text === '') {
        return $text;
    }

    $previous = null;
    while ($previous !== $text) {
        $previous = $text;
        $result = preg_replace('/\b([\p{L}]+)(\s+\1)+\b/iu', '$1', $text);
        if ($result === null) {
            break;
        }
        $text = $result;
    }

    return $text;
}

/**
 * Perbaikan typo umum pada teks kop surat dan isi surat.
 */
function spt_fix_common_typos($text)
{
    $text = (string)$text;
    if ($text === '') {
        return $text;
    }

    return str_ireplace(
        ['Bojongeoro', 'Telepn', 'Telepo n', 'Telepun'],
        ['Bojonegoro', 'Telepon', 'Telepon', 'Telepon'],
        $text
    );
}

/**
 * Perbaiki typo kop surat di dalam DOCX hasil generate (semua XML text part).
 * Juga menangani typo yang terpecah antar-run Word.
 */
function spt_fix_docx_typos($docxPath)
{
    if (!class_exists('ZipArchive') || !is_file($docxPath)) {
        return false;
    }

    $zip = new ZipArchive();
    if ($zip->open($docxPath) !== true) {
        return false;
    }

    $changed = false;
    $parts = [];
    for ($i = 0; $i < $zip->numFiles; $i++) {
        $name = $zip->getNameIndex($i);
        if (preg_match('#^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$#i', $name)) {
            $parts[] = $name;
        }
    }

    // Typo yang mungkin terpecah ke beberapa run <w:t>.
    $splitPatterns = [
        '/B\s*(?:<[^>]+>\s*)*o\s*(?:<[^>]+>\s*)*j\s*(?:<[^>]+>\s*)*o\s*(?:<[^>]+>\s*)*n\s*(?:<[^>]+>\s*)*g\s*(?:<[^>]+>\s*)*e\s*(?:<[^>]+>\s*)*o\s*(?:<[^>]+>\s*)*r\s*(?:<[^>]+>\s*)*o/ui',
        '/T\s*(?:<[^>]+>\s*)*e\s*(?:<[^>]+>\s*)*l\s*(?:<[^>]+>\s*)*e\s*(?:<[^>]+>\s*)*p\s*(?:<[^>]+>\s*)*n/ui',
    ];

    foreach ($parts as $name) {
        $content = $zip->getFromName($name);
        if ($content === false) {
            continue;
        }

        $fixed = spt_fix_common_typos($content);
        $splitFixed = preg_replace($splitPatterns, ['Bojonegoro', 'Telepon'], $fixed);
        if ($splitFixed !== null) {
            $fixed = $splitFixed;
        }

        if ($fixed !== $content) {
            $zip->addFromString($name, $fixed);
            $changed = true;
        }
    }

    $zip->close();
    return $changed;
}

/**
 * Hitung jumlah halaman PDF secara best-effort dari isi file.
 */
function spt_pdf_page_count($path)
{
    if (!is_file($path) || !is_readable($path)) {
        return 0;
    }

    $data = @file_get_contents($path);
    if ($data === false || $data === '') {
        return 0;
    }

    if (preg_match_all('/\/Type\s*\/Page[^s]/i', $data, $matches)) {
        return count($matches[0]);
    }
    if (preg_match_all('/\/Count\s+(\d+)/i', $data, $matches)) {
        return max(array_map('intval', $matches[1]));
    }

    return 0;
}

/**
 * Render satu kontrol slider + number input + stepper untuk panel inspector.
 */
function spt_layout_render_control($key, $label, array $settings, $unit, $help = '')
{
    $ranges = spt_layout_ranges();
    if (!isset($ranges[$key])) {
        return '';
    }

    [$min, $max, $step] = $ranges[$key];
    $defaults = spt_layout_default_values();
    $default = $defaults[$key] ?? $min;
    $value = $settings[$key] ?? $default;
    $decimals = $step < 0.1 ? 2 : ($step < 1 ? 1 : 0);
    $display = number_format((float)$value, $decimals, '.', '');
    $minAttr = htmlspecialchars((string)$min, ENT_QUOTES, 'UTF-8');
    $maxAttr = htmlspecialchars((string)$max, ENT_QUOTES, 'UTF-8');
    $stepAttr = htmlspecialchars((string)$step, ENT_QUOTES, 'UTF-8');
    $valueAttr = htmlspecialchars((string)$value, ENT_QUOTES, 'UTF-8');
    $labelAttr = htmlspecialchars($label, ENT_QUOTES, 'UTF-8');
    $unitHtml = htmlspecialchars($unit, ENT_QUOTES, 'UTF-8');
    $helpHtml = $help !== '' ? '<p class="spt-help">' . htmlspecialchars($help, ENT_QUOTES, 'UTF-8') . '</p>' : '';

    return <<<HTML
<div class="spt-control" data-control="{$key}">
    <div class="spt-control-head">
        <label class="spt-control-label" for="num-{$key}">{$labelAttr}</label>
        <button type="button" class="layout-reset spt-reset-btn" data-target="{$key}" title="Reset {$labelAttr} ke default" aria-label="Reset {$labelAttr}">&#8634;</button>
    </div>
    <div class="spt-control-row">
        <button type="button" class="layout-step spt-step-btn" data-target="{$key}" data-dir="-1" aria-label="Kurangi {$labelAttr}">&minus;</button>
        <div class="spt-number-wrap">
            <input type="number" id="num-{$key}" class="layout-number spt-number" data-number-key="{$key}" min="{$minAttr}" max="{$maxAttr}" step="{$stepAttr}" value="{$valueAttr}">
            <span class="spt-unit">{$unitHtml}</span>
        </div>
        <button type="button" class="layout-step spt-step-btn" data-target="{$key}" data-dir="1" aria-label="Tambah {$labelAttr}">+</button>
    </div>
    <input type="range" class="layout-range spt-range" data-layout-key="{$key}" min="{$minAttr}" max="{$maxAttr}" step="{$stepAttr}" value="{$valueAttr}" aria-label="{$labelAttr}">
    {$helpHtml}
</div>
HTML;
}

/**
 * =====================================================================
 * PRATINJAU INTERAKTIF (HTML)
 * =====================================================================
 * Pratinjau dirender dari DOCX hasil generate, lalu diberi penanda blok
 * dan CSS variabel agar jarak antar blok bisa ditarik langsung di halaman.
 * Hasil akhir tetap PDF/Word lewat pipeline DOCX yang sudah ada.
 */

function spt_preview_norm_text($text)
{
    return trim(preg_replace('/\s+/u', ' ', (string)$text));
}

function spt_preview_key_text($text)
{
    return preg_replace('/\s+/u', '', (string)$text);
}

function spt_preview_node_text(DOMXPath $xpath, DOMNode $node)
{
    $text = '';
    foreach ($xpath->query('.//w:t', $node) as $t) {
        $text .= $t->textContent;
    }
    return $text;
}

function spt_preview_twips_pt($twips)
{
    return ((float)$twips) / 20.0;
}

function spt_preview_emu_mm($emu)
{
    return ((float)$emu) / 36000.0;
}

function spt_preview_emu_pt($emu)
{
    return ((float)$emu) / 12700.0;
}

function spt_preview_emu_px($emu)
{
    return ((float)$emu) / 9525.0;
}

/**
 * Konversi nilai CSS (mis. "1in", "35.25pt", "54px") ke piksel.
 */
function spt_preview_css_length_px($value)
{
    if (!preg_match('/^\s*(-?[0-9.]+)\s*(pt|in|cm|mm|px|pc)?\s*$/i', (string)$value, $m)) {
        return 0.0;
    }
    $number = (float)$m[1];
    $unit = strtolower($m[2] ?? 'px');
    switch ($unit) {
        case 'pt':
            return $number * 96 / 72;
        case 'in':
            return $number * 96;
        case 'cm':
            return $number * 96 / 2.54;
        case 'mm':
            return $number * 96 / 25.4;
        case 'pc':
            return $number * 16;
        default:
            return $number;
    }
}

function spt_preview_css_length_mm($value)
{
    return spt_preview_css_length_px($value) * 25.4 / 96;
}

/**
 * Kumpulkan informasi DOCX yang dibutuhkan pratinjau: ukuran halaman,
 * border tabel, text box (blok TTD), properti paragraf, dan gambar.
 */
function spt_preview_docx_bundle($docxPath)
{
    $bundle = [
        'page' => [
            'width_mm' => 210.0,
            'height_mm' => 297.0,
            'margin_left' => 20.0,
            'margin_right' => 20.0,
            'margin_top' => 15.0,
            'margin_bottom' => 15.0,
        ],
        'tables' => [],
        'textboxes' => [],
        'paragraphs' => [],
        'media' => [],
        'images' => [],
        'floats' => [],
    ];

    if (!class_exists('ZipArchive') || !is_file($docxPath)) {
        return $bundle;
    }

    $zip = new ZipArchive();
    if ($zip->open($docxPath) !== true) {
        return $bundle;
    }

    $documentXml = $zip->getFromName('word/document.xml');
    $stylesXml = $zip->getFromName('word/styles.xml');
    $relsXml = $zip->getFromName('word/_rels/document.xml.rels');
    if ($documentXml === false) {
        $zip->close();
        return $bundle;
    }
    $docStyles = spt_layout_docx_styles($stylesXml === false ? '' : $stylesXml);

    $doc = new DOMDocument();
    if (!@$doc->loadXML($documentXml)) {
        $zip->close();
        return $bundle;
    }

    $wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    $xp = new DOMXPath($doc);
    $xp->registerNamespace('w', $wNs);
    $xp->registerNamespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing');
    $xp->registerNamespace('wps', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape');
    $xp->registerNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
    $xp->registerNamespace('v', 'urn:schemas-microsoft-com:vml');
    $xp->registerNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

    $pgSz = $xp->query('//w:body/w:sectPr/w:pgSz')->item(0);
    if ($pgSz) {
        $w = (int)$pgSz->getAttributeNS($wNs, 'w');
        $h = (int)$pgSz->getAttributeNS($wNs, 'h');
        if ($w > 0) {
            $bundle['page']['width_mm'] = round($w * 25.4 / 1440, 2);
        }
        if ($h > 0) {
            $bundle['page']['height_mm'] = round($h * 25.4 / 1440, 2);
        }
    }

    $pgMar = $xp->query('//w:body/w:sectPr/w:pgMar')->item(0);
    if ($pgMar) {
        foreach (['top', 'bottom', 'left', 'right'] as $side) {
            $value = (int)$pgMar->getAttributeNS($wNs, $side);
            if ($value > 0) {
                $bundle['page']['margin_' . $side] = round($value * 25.4 / 1440, 2);
            }
        }
    }

    // Peta relationship gambar (logo, tanda tangan) ke data URI.
    $mediaRels = [];
    if ($relsXml !== false) {
        $relsDoc = new DOMDocument();
        if (@$relsDoc->loadXML($relsXml)) {
            $relsXp = new DOMXPath($relsDoc);
            $relsXp->registerNamespace('r', 'http://schemas.openxmlformats.org/package/2006/relationships');
            foreach ($relsXp->query('//r:Relationship') as $rel) {
                $target = str_replace('\\', '/', $rel->getAttribute('Target'));
                if (stripos($target, 'media/') !== false) {
                    $mediaRels[$rel->getAttribute('Id')] = 'word/' . ltrim($target, '/');
                }
            }
        }
    }
    foreach ($mediaRels as $rid => $path) {
        $data = $zip->getFromName($path);
        if ($data === false) {
            continue;
        }
        $ext = strtolower(pathinfo($path, PATHINFO_EXTENSION));
        $mime = [
            'png' => 'image/png',
            'jpg' => 'image/jpeg',
            'jpeg' => 'image/jpeg',
            'gif' => 'image/gif',
            'bmp' => 'image/bmp',
            'svg' => 'image/svg+xml',
        ][$ext] ?? 'application/octet-stream';
        $bundle['media'][$rid] = 'data:' . $mime . ';base64,' . base64_encode($data);
    }

    // Border, lebar kolom, dan posisi tabel sesuai DOCX. PHPWord selalu
    // menambahkan border dan mengabaikan lebar kolom saat export HTML.
    foreach ($xp->query('//w:body//w:tbl') as $tbl) {
        $flags = [];
        foreach (['top', 'left', 'bottom', 'right', 'insideH', 'insideV'] as $side) {
            $node = $xp->query('w:tblPr/w:tblBorders/w:' . $side, $tbl)->item(0);
            $val = $node ? strtolower($node->getAttributeNS($wNs, 'val')) : '';
            $flags[$side] = ($val !== '' && $val !== 'nil' && $val !== 'none');
        }

        $widths = [];
        foreach ($xp->query('w:tblGrid/w:gridCol', $tbl) as $col) {
            $width = (int)$col->getAttributeNS($wNs, 'w');
            if ($width > 0) {
                $widths[] = round($width / 15, 1);
            }
        }

        $align = '';
        $jc = $xp->query('w:tblPr/w:jc', $tbl)->item(0);
        if ($jc) {
            $align = $jc->getAttributeNS($wNs, 'val');
        }

        $indentPx = 0.0;
        $ind = $xp->query('w:tblPr/w:tblInd', $tbl)->item(0);
        if ($ind) {
            $indentPx = round(((int)$ind->getAttributeNS($wNs, 'w')) / 15, 1);
        }

        $bundle['tables'][] = [
            'borders' => $flags,
            'widths' => $widths,
            'align' => $align,
            'indent_px' => $indentPx,
        ];
    }

    // Ukuran gambar dari DOCX. PHPWord memakai ukuran intrinsik PNG,
    // bukan ukuran tampilan yang diatur di dokumen.
    foreach ($xp->query('//w:body//wp:inline | //w:body//wp:anchor') as $drawing) {
        if ($xp->query('ancestor::wps:txbx', $drawing)->length > 0) {
            continue;
        }
        $extent = $xp->query('wp:extent', $drawing)->item(0);
        if (!$extent) {
            continue;
        }
        $width = (int)$extent->getAttribute('cx');
        $height = (int)$extent->getAttribute('cy');
        if ($width <= 0 || $height <= 0) {
            continue;
        }
        $bundle['images'][] = [
            'width_px' => round(spt_preview_emu_px($width), 1),
            'height_px' => round(spt_preview_emu_px($height), 1),
        ];
    }

    // Logo kop template lama berupa shape VML melayang. PHPWord merendernya
    // inline dengan ukuran intrinsik, sehingga baris kop terdorong; posisi
    // dan ukuran asli diambil dari style shape agar sama dengan Word.
    $seenFloat = [];
    foreach ($xp->query('//w:body//v:shape[v:imagedata]') as $shape) {
        $shapeStyle = $shape->getAttribute('style');
        if (stripos($shapeStyle, 'position:absolute') === false) {
            continue;
        }
        $shapeProps = [];
        foreach (explode(';', $shapeStyle) as $part) {
            $pair = explode(':', $part, 2);
            if (count($pair) === 2) {
                $shapeProps[strtolower(trim($pair[0]))] = trim($pair[1]);
            }
        }
        $width = spt_preview_css_length_px($shapeProps['width'] ?? '');
        $height = spt_preview_css_length_px($shapeProps['height'] ?? '');
        if ($width <= 0 || $height <= 0) {
            continue;
        }
        $left = spt_preview_css_length_mm($shapeProps['margin-left'] ?? '0');
        $top = spt_preview_css_length_mm($shapeProps['margin-top'] ?? '0');
        $floatKey = $shapeStyle . '|' . $left . '|' . $top;
        if (isset($seenFloat[$floatKey])) {
            continue;
        }
        $seenFloat[$floatKey] = true;
        $imageData = $shape->getElementsByTagNameNS(
            'urn:schemas-microsoft-com:vml',
            'imagedata'
        )->item(0);
        $imageRid = $imageData
            ? $imageData->getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')
            : '';
        $bundle['floats'][] = [
            'width_px' => round($width, 1),
            'height_px' => round($height, 1),
            'left_mm' => round($left, 2),
            'top_mm' => round($top, 2),
            'src' => $imageRid !== '' && isset($bundle['media'][$imageRid]) ? $bundle['media'][$imageRid] : '',
        ];
    }

    // Properti paragraf level atas untuk memperbaiki alignment/indentasi
    // yang hilang saat PHPWord mengonversi heading ke HTML.
    foreach ($xp->query('//w:body/w:p') as $p) {
        $text = spt_preview_norm_text(spt_preview_node_text($xp, $p));
        $key = spt_preview_key_text($text);
        if ($key === '') {
            continue;
        }
        $align = '';
        $jc = $xp->query('w:pPr/w:jc', $p)->item(0);
        if ($jc) {
            $value = $jc->getAttributeNS($wNs, 'val');
            $align = ($value === 'both') ? 'justify' : $value;
        }
        $leftPt = 0.0;
        $firstPt = 0.0;
        $ind = $xp->query('w:pPr/w:ind', $p)->item(0);
        if ($ind) {
            $leftPt = spt_preview_twips_pt((int)$ind->getAttributeNS($wNs, 'left'));
            $firstPt = spt_preview_twips_pt((int)$ind->getAttributeNS($wNs, 'firstLine'));
            $hanging = (int)$ind->getAttributeNS($wNs, 'hanging');
            if ($firstPt == 0.0 && $hanging > 0) {
                $firstPt = -spt_preview_twips_pt($hanging);
            }
        }
        $pageBreak = $xp->query('w:pPr/w:pageBreakBefore', $p)->length > 0
            || $xp->query('./w:r/w:br[@w:type="page"]', $p)->length > 0
            || preg_match('/^Lampiran\s+Surat\b/i', $text);

        // PHPWord tidak menulis ukuran huruf untuk run yang mewarisi style
        // paragraf, jadi ukurannya diambil dari DOCX agar judul/blok lain
        // tidak ikut mengecil ke ukuran isi.
        $fontHalfPoints = 0;
        foreach ($xp->query('.//w:r/w:rPr/w:sz', $p) as $sz) {
            $value = (int)$sz->getAttributeNS($wNs, 'val');
            if ($value > 0) {
                $fontHalfPoints = $value;
                break;
            }
        }
        if ($fontHalfPoints <= 0) {
            $pStyleId = 'Normal';
            $pStyle = $xp->query('w:pPr/w:pStyle', $p)->item(0);
            if ($pStyle) {
                $pStyleId = $pStyle->getAttributeNS($wNs, 'val');
            }
            $fontHalfPoints = spt_layout_style_value($docStyles, $pStyleId, 'sz');
        }

        $bundle['paragraphs'][] = [
            'key' => $key,
            'align' => $align,
            'left_pt' => $leftPt,
            'first_pt' => $firstPt,
            'page_break' => $pageBreak,
            'font_size_pt' => $fontHalfPoints > 0 ? $fontHalfPoints / 2 : 0,
        ];
    }

    // Text box (blok TTD). Ambil dari wps:txbx agar tidak duplikat dengan fallback VML.
    $seenTextBox = [];
    foreach ($xp->query('//w:body//wps:txbx') as $txbx) {
        $anchor = $xp->query('ancestor::wp:anchor[1]', $txbx)->item(0);
        $plainParts = [];
        foreach ($xp->query('./w:txbxContent/w:p', $txbx) as $p) {
            $plainParts[] = spt_preview_node_text($xp, $p);
        }
        $plain = implode('', $plainParts);
        $key = spt_preview_key_text($plain);
        if ($key === '') {
            continue;
        }

        $posX = 0;
        $posY = 0;
        $relH = 'column';
        $relV = 'paragraph';
        $cx = 0;
        $cy = 0;
        if ($anchor) {
            $posH = $xp->query('wp:positionH', $anchor)->item(0);
            $posV = $xp->query('wp:positionV', $anchor)->item(0);
            if ($posH) {
                $relH = $posH->getAttribute('relativeFrom') ?: 'column';
                $offset = $xp->query('wp:posOffset', $posH)->item(0);
                $posX = $offset ? (int)$offset->textContent : 0;
            }
            if ($posV) {
                $relV = $posV->getAttribute('relativeFrom') ?: 'paragraph';
                $offset = $xp->query('wp:posOffset', $posV)->item(0);
                $posY = $offset ? (int)$offset->textContent : 0;
            }
            $extent = $xp->query('wp:extent', $anchor)->item(0);
            if ($extent) {
                $cx = (int)$extent->getAttribute('cx');
                $cy = (int)$extent->getAttribute('cy');
            }
        }

        $dedupe = $key . '|' . $posX . '|' . $posY;
        if (isset($seenTextBox[$dedupe])) {
            continue;
        }
        $seenTextBox[$dedupe] = true;

        $xMm = spt_preview_emu_mm($posX);
        if ($relH === 'page') {
            $xMm -= $bundle['page']['margin_left'];
        }
        $yPt = spt_preview_emu_pt($posY);
        if ($relV === 'margin' || $relV === 'page') {
            $yPt -= $bundle['page']['margin_top'] * 72 / 25.4;
        }
        $widthMm = $cx > 0 ? spt_preview_emu_mm($cx) : 90.0;
        $heightPt = $cy > 0 ? spt_preview_emu_pt($cy) : 0.0;

        $lines = [];
        foreach ($xp->query('./w:txbxContent/w:p', $txbx) as $p) {
            $lineText = spt_preview_norm_text(spt_preview_node_text($xp, $p));
            $rendered = spt_preview_render_txbx_paragraph($xp, $p, $bundle['media'], $wNs);
            if ($lineText === '' && $rendered['html'] === '') {
                continue;
            }
            $lines[] = ['text' => $lineText, 'html' => $rendered['html'], 'style' => $rendered['style']];
        }

        $isDate = function ($text) {
            return (bool)preg_match('/^(Bojonegoro\b|\d{1,2}\s+\S+\s+\d{4})/i', $text);
        };
        $isNip = function ($text) {
            return (bool)preg_match('/^(NIP|NIP\.)\b/i', trim($text));
        };

        $isMain = false;
        $nipIndex = -1;
        foreach ($lines as $index => $line) {
            if ($isDate($line['text'])) {
                $isMain = true;
            }
            if ($nipIndex < 0 && $isNip($line['text'])) {
                $nipIndex = $index;
            }
        }
        // Nama penanda tangan adalah baris tepat sebelum NIP.
        $nameIndex = $nipIndex > 0 ? $nipIndex - 1 : -1;
        if ($nameIndex < 0) {
            for ($i = count($lines) - 1; $i >= 0; $i--) {
                if ($lines[$i]['text'] !== '' && !$isDate($lines[$i]['text'])) {
                    $nameIndex = $i;
                    break;
                }
            }
        }

        foreach ($lines as $index => &$line) {
            $text = $line['text'];
            if ($isNip($text)) {
                $line['class'] = 'spt-ttd-nip';
            } elseif ($isDate($text)) {
                $line['class'] = 'spt-ttd-date';
            } elseif ($index === $nameIndex) {
                $line['class'] = 'spt-ttd-name';
            } elseif ($nameIndex >= 0 && $index < $nameIndex) {
                $line['class'] = 'spt-ttd-jabatan';
            } else {
                $line['class'] = 'spt-ttd-line-plain';
            }
        }
        unset($line);

        $lastJabatan = -1;
        foreach ($lines as $index => $line) {
            if ($line['class'] === 'spt-ttd-jabatan') {
                $lastJabatan = $index;
            }
        }
        $html = '';
        foreach ($lines as $index => $line) {
            $class = 'spt-ttd-line ' . $line['class'];
            if ($index === $lastJabatan) {
                $class .= ' spt-ttd-jabatan-last';
            }
            $styleAttr = !empty($line['style']) ? ' style="' . $line['style'] . '"' : '';
            $html .= '<p class="' . $class . '"' . $styleAttr . '>' . $line['html'] . '</p>';
        }

        $bundle['textboxes'][] = [
            'key' => $key,
            'html' => $html,
            'x_mm' => round($xMm, 2),
            'y_pt' => round($yPt, 2),
            'width_mm' => round($widthMm, 2),
            'height_pt' => round($heightPt, 2),
            'main' => $isMain,
        ];
    }

    $zip->close();
    return $bundle;
}

/**
 * Render satu paragraf di dalam text box TTD menjadi HTML sederhana
 * (teks, tebal, garis bawah, dan gambar tanda tangan bila ada).
 */
function spt_preview_render_txbx_paragraph(DOMXPath $xp, DOMElement $p, array $media, $wNs)
{
    $styles = [];
    $jc = $xp->query('w:pPr/w:jc', $p)->item(0);
    if ($jc) {
        $value = $jc->getAttributeNS($wNs, 'val');
        $styles[] = 'text-align:' . ($value === 'both' ? 'justify' : $value);
    }
    $ind = $xp->query('w:pPr/w:ind', $p)->item(0);
    if ($ind) {
        $left = (int)$ind->getAttributeNS($wNs, 'left');
        $first = (int)$ind->getAttributeNS($wNs, 'firstLine');
        if ($left > 0) {
            $styles[] = 'margin-left:' . round(spt_preview_twips_pt($left), 1) . 'pt';
        }
        if ($first > 0) {
            $styles[] = 'text-indent:' . round(spt_preview_twips_pt($first), 1) . 'pt';
        }
    }
    $spacing = $xp->query('w:pPr/w:spacing', $p)->item(0);
    if ($spacing) {
        $before = (int)$spacing->getAttributeNS($wNs, 'before');
        $after = (int)$spacing->getAttributeNS($wNs, 'after');
        if ($before > 0) {
            $styles[] = 'margin-top:' . round(spt_preview_twips_pt($before), 1) . 'pt';
        }
        if ($after > 0) {
            $styles[] = 'margin-bottom:' . round(spt_preview_twips_pt($after), 1) . 'pt';
        }
    }
    $styleAttr = implode(';', $styles);

    $html = '';
    foreach ($xp->query('./w:r | ./w:hyperlink/w:r', $p) as $run) {
        $runStyles = [];
        $sz = $xp->query('w:rPr/w:sz', $run)->item(0);
        if ($sz) {
            $runStyles[] = 'font-size:' . (((int)$sz->getAttributeNS($wNs, 'val')) / 2) . 'pt';
        }
        $fonts = $xp->query('w:rPr/w:rFonts', $run)->item(0);
        if ($fonts) {
            $family = $fonts->getAttributeNS($wNs, 'ascii');
            if ($family !== '') {
                $runStyles[] = "font-family:'" . str_replace("'", '', $family) . "'";
            }
        }
        if ($xp->query('w:rPr/w:b', $run)->length > 0) {
            $runStyles[] = 'font-weight:bold';
        }
        if ($xp->query('w:rPr/w:i', $run)->length > 0) {
            $runStyles[] = 'font-style:italic';
        }
        $underline = $xp->query('w:rPr/w:u', $run)->item(0);
        if ($underline && $underline->getAttributeNS($wNs, 'val') !== 'none') {
            $runStyles[] = 'text-decoration:underline';
        }
        $runStyle = implode(';', $runStyles);

        foreach ($run->childNodes as $child) {
            if ($child->nodeType !== XML_ELEMENT_NODE) {
                continue;
            }
            $name = $child->localName;
            if ($name === 't') {
                $html .= '<span' . ($runStyle !== '' ? ' style="' . $runStyle . '"' : '') . '>'
                    . htmlspecialchars($child->textContent, ENT_QUOTES, 'UTF-8') . '</span>';
            } elseif ($name === 'tab') {
                $html .= '<span class="spt-tab"></span>';
            } elseif ($name === 'br') {
                $html .= '<br>';
            } elseif ($name === 'drawing' || $name === 'pict') {
                $rid = '';
                $blip = $xp->query('.//a:blip', $child)->item(0);
                if ($blip) {
                    $rid = $blip->getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed');
                } else {
                    $img = $xp->query('.//v:imagedata', $child)->item(0);
                    if ($img) {
                        $rid = $img->getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
                    }
                }
                if ($rid !== '' && isset($media[$rid])) {
                    $width = 0;
                    $height = 0;
                    $extent = $xp->query('.//wp:extent', $child)->item(0);
                    if ($extent) {
                        $width = (int)spt_preview_emu_px((int)$extent->getAttribute('cx'));
                        $height = (int)spt_preview_emu_px((int)$extent->getAttribute('cy'));
                    }
                    $size = '';
                    if ($width > 0 && $height > 0) {
                        $size = ' style="width:' . $width . 'px;height:' . $height . 'px;"';
                    }
                    $html .= '<img src="' . $media[$rid] . '"' . $size . ' alt="">';
                }
            }
        }
    }

    if ($html === '') {
        return ['html' => '', 'style' => ''];
    }
    return ['html' => $html, 'style' => $styleAttr];
}

/**
 * Bangun halaman HTML pratinjau: perbaiki text box, border tabel,
 * properti paragraf, pecah per halaman, lalu beri penanda blok.
 */
function spt_preview_build_pages($phpWordHtml, array $bundle)
{
    if (preg_match('/<body\b[^>]*>(.*)<\/body>/is', $phpWordHtml, $match)) {
        $fragment = $match[1];
    } else {
        $fragment = $phpWordHtml;
    }

    // Skalakan ukuran huruf inline lewat variabel faktor agar slider font
    // hanya mengubah proporsi, bukan menyamakan semua ukuran huruf.
    $fragment = preg_replace_callback('/style\s*=\s*"([^"]*)"/i', function ($m) {
        return 'style="' . preg_replace(
            '/(?<![-\w])font-size\s*:\s*([0-9.]+)\s*pt/i',
            'font-size: calc($1pt * var(--spt-font-factor))',
            $m[1]
        ) . '"';
    }, $fragment);

    // Tab pada DOCX jangan menyatu menjadi satu spasi.
    $fragment = str_replace("\t", '<span class="spt-tab"></span>', $fragment);

    $dom = new DOMDocument();
    @$dom->loadHTML(
        '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>' . $fragment . '</body></html>',
        LIBXML_NOERROR | LIBXML_NOWARNING
    );
    $xp = new DOMXPath($dom);

    // Ganti paragraf hasil flatten text box dengan blok TTD yang diposisikan.
    $usedTextBox = [];
    foreach ($bundle['textboxes'] as $tbx) {
        $target = null;
        foreach ($xp->query('//body//p | //body//h0 | //body//h1 | //body//h2 | //body//h3 | //body//h4 | //body//h5 | //body//h6 | //body//h7 | //body//h8 | //body//h9') as $el) {
            if (isset($usedTextBox[spl_object_id($el)])) {
                continue;
            }
            if (spt_preview_key_text($el->textContent) === $tbx['key']) {
                $target = $el;
                break;
            }
        }
        if (!$target) {
            continue;
        }
        $usedTextBox[spl_object_id($target)] = true;

        $wrapper = $dom->createElement('div');
        $wrapper->setAttribute('class', 'spt-ttd-anchor');
        $wrapper->setAttribute('data-spt-block', 'ttd');
        $boxClass = 'spt-txbx ' . ($tbx['main'] ? 'spt-txbx-main' : 'spt-txbx-static');
        $boxStyle = '--txbx-left:' . $tbx['x_mm'] . 'mm;--txbx-top:' . $tbx['y_pt'] . 'pt;'
            . 'width:' . $tbx['width_mm'] . 'mm;min-height:' . $tbx['height_pt'] . 'pt;';
        $boxHtml = '<div class="' . $boxClass . '" data-spt-block="ttd"'
            . ($tbx['main'] ? ' data-spt-main="1"' : '')
            . ' style="' . $boxStyle . '">' . $tbx['html'] . '</div>';
        $frag = $dom->createDocumentFragment();
        if (@$frag->appendXML($boxHtml)) {
            $wrapper->appendChild($frag);
            $target->parentNode->replaceChild($wrapper, $target);
        }
    }

    // Samakan ukuran gambar dengan ukuran tampilan di Word.
    $imageIndex = 0;
    foreach ($xp->query('//body//img') as $img) {
        $size = $bundle['images'][$imageIndex] ?? null;
        $imageIndex++;
        if (!$size) {
            continue;
        }
        $img->setAttribute('style', 'width:' . $size['width_px'] . 'px;height:' . $size['height_px'] . 'px;');
        $img->setAttribute('width', (string)round($size['width_px']));
        $img->setAttribute('height', (string)round($size['height_px']));
    }

    // Gambar melayang (logo kop template lama) diposisikan absolut terhadap
    // halaman. Margin halaman dikurangkan lewat calc agar posisi tetap saat
    // slider margin diubah.
    if (!empty($bundle['floats'])) {
        $floatsById = [];
        foreach ($bundle['floats'] as $float) {
            if ($float['src'] !== '') {
                $floatsById[$float['src']] = $float;
            }
        }
        foreach ($xp->query('//body//img') as $img) {
            $src = $img->getAttribute('src');
            if ($src === '' || !isset($floatsById[$src])) {
                continue;
            }
            $float = $floatsById[$src];
            unset($floatsById[$src]);
            $img->setAttribute('style',
                'position:absolute;'
                . 'left:calc(' . $float['left_mm'] . 'mm - var(--spt-margin-left));'
                . 'top:calc(' . $float['top_mm'] . 'mm - var(--spt-margin-top));'
                . 'width:' . $float['width_px'] . 'px;height:' . $float['height_px'] . 'px;'
                . 'max-width:none;z-index:2;'
            );
        }
    }

    // Terapkan alignment/indentasi dari DOCX untuk heading yang kehilangan
    // styling saat dikonversi PHPWord.
    $htmlBlocks = [];
    foreach ($xp->query('//body//p | //body//h0 | //body//h1 | //body//h2 | //body//h3 | //body//h4 | //body//h5 | //body//h6 | //body//h7 | //body//h8 | //body//h9') as $el) {
        if ($xp->query('ancestor::div[contains(@class,"spt-txbx")]', $el)->length > 0) {
            continue;
        }
        $htmlBlocks[] = $el;
    }
    $cursor = 0;
    $blockCount = count($htmlBlocks);
    foreach ($bundle['paragraphs'] as $para) {
        for ($i = $cursor; $i < $blockCount; $i++) {
            $el = $htmlBlocks[$i];
            if (spt_preview_key_text($el->textContent) !== $para['key']) {
                continue;
            }
            $cursor = $i + 1;
            $style = $el->getAttribute('style');
            $add = [];
            if ($para['align'] !== '' && stripos($style, 'text-align') === false) {
                $add[] = 'text-align:' . $para['align'];
            }
            if ($para['left_pt'] > 0 && stripos($style, 'margin-left') === false) {
                $add[] = 'margin-left:' . round($para['left_pt'], 1) . 'pt';
            }
            if (abs($para['first_pt']) > 0 && stripos($style, 'text-indent') === false) {
                $add[] = 'text-indent:' . round($para['first_pt'], 1) . 'pt';
            }
            $hasFontSize = stripos($style, 'font-size') !== false
                || $xp->query('.//span[contains(@style,"font-size")]', $el)->length > 0;
            if (!$hasFontSize && !empty($para['font_size_pt'])) {
                $add[] = 'font-size:calc(' . $para['font_size_pt'] . 'pt * var(--spt-font-factor))';
            }
            if ($add) {
                $el->setAttribute('style', trim($style . ';' . implode(';', $add), ';'));
            }
            if ($para['page_break']) {
                $el->setAttribute('data-spt-page-break', '1');
            }
            break;
        }
    }

    // Border, lebar kolom, dan perataan tabel mengikuti DOCX.
    $tableIndex = 0;
    foreach ($xp->query('//body//table') as $table) {
        $info = $bundle['tables'][$tableIndex] ?? [];
        $tableIndex++;
        $flags = $info['borders'] ?? [];
        foreach (['top', 'left', 'bottom', 'right', 'insideH', 'insideV'] as $side) {
            if (!empty($flags[$side])) {
                $table->setAttribute('class', trim($table->getAttribute('class') . ' spt-b-' . strtolower($side)));
            }
        }

        $widths = $info['widths'] ?? [];
        if ($widths) {
            $total = round(array_sum($widths), 1);
            $style = 'width:' . $total . 'px;';
            if (($info['align'] ?? '') === 'center') {
                $style .= 'margin-left:auto;margin-right:auto;';
            } elseif (($info['align'] ?? '') === 'right') {
                $style .= 'margin-left:auto;';
            } elseif (!empty($info['indent_px'])) {
                $style .= 'margin-left:' . $info['indent_px'] . 'px;';
            }
            $table->setAttribute('style', trim($table->getAttribute('style') . ';' . $style, ';'));
            foreach ($xp->query('./tr | ./tbody/tr', $table) as $row) {
                $cells = $xp->query('./td | ./th', $row);
                foreach ($cells as $index => $cell) {
                    if (!isset($widths[$index])) {
                        continue;
                    }
                    $cell->setAttribute(
                        'style',
                        trim($cell->getAttribute('style') . ';width:' . $widths[$index] . 'px', ';')
                    );
                }
            }
        }
    }

    // Tentukan isi halaman mengikuti page-break dari DOCX/PHPWord.
    $pageRoot = null;
    foreach ($xp->query('//body/div') as $div) {
        if (stripos($div->getAttribute('style'), 'page:') !== false) {
            $pageRoot = $div;
            break;
        }
    }
    if (!$pageRoot) {
        $pageRoot = $xp->query('//body')->item(0);
    }

    $pages = [];
    $page = null;
    foreach (iterator_to_array($pageRoot->childNodes) as $child) {
        if ($child->nodeType !== XML_ELEMENT_NODE) {
            continue;
        }
        $isBreak = $child->getAttribute('data-spt-page-break') === '1';
        $style = $child->getAttribute('style');
        if (stripos($style, 'page-break-before') !== false && stripos($style, 'always') !== false) {
            $isBreak = true;
            $child->setAttribute('style', preg_replace('/page-break-before\s*:\s*always\s*;?/i', '', $style));
        }
        if ($isBreak && $page !== null && $page->hasChildNodes()) {
            $pages[] = $page;
            $page = null;
        }
        if ($page === null) {
            $page = $dom->createElement('div');
            $page->setAttribute('class', 'spt-page');
            $inner = $dom->createElement('div');
            $inner->setAttribute('class', 'spt-page-inner');
            $page->appendChild($inner);
        }
        $page->firstChild->appendChild($child);
    }
    if ($page !== null && $page->hasChildNodes()) {
        $pages[] = $page;
    }

    // Beri penanda blok pada setiap halaman.
    foreach ($pages as $page) {
        spt_preview_tag_page($dom, $xp, $page->firstChild);
    }

    $container = $dom->createElement('div');
    $container->setAttribute('id', 'sptPages');
    foreach ($pages as $page) {
        $container->appendChild($page);
    }

    // Paragraf kosong sesudah blok TTD mengambang hanya penyangga posisi di
    // Word. Di pratinjau blok TTD sudah diposisikan absolut, jadi penyangga
    // itu dibuang agar tidak menciptakan halaman tambahan yang hampir kosong.
    foreach ($xp->query('.//div[contains(@class,"spt-page-inner")]//div[contains(@class,"spt-ttd-anchor")]', $container) as $anchor) {
        $node = $anchor->nextSibling;
        while ($node) {
            $nextNode = $node->nextSibling;
            if ($node->nodeType === XML_ELEMENT_NODE) {
                $tag = strtolower($node->nodeName);
                $nodeText = trim(str_replace("\xC2\xA0", ' ', $node->textContent));
                if (!preg_match('/^(p|h[0-9])$/', $tag)
                    || $nodeText !== ''
                    || $node->getAttribute('data-spt-page-break') === '1') {
                    break;
                }
                $node->parentNode->removeChild($node);
            }
            $node = $nextNode;
        }
    }

    return $dom->saveHTML($container);
}

/**
 * Tandai blok semantik surat pada satu halaman pratinjau.
 */
function spt_preview_tag_page(DOMDocument $dom, DOMXPath $xp, DOMElement $inner)
{
    $children = [];
    foreach ($inner->childNodes as $child) {
        if ($child->nodeType === XML_ELEMENT_NODE) {
            $children[] = $child;
        }
    }

    $addClass = function (DOMElement $el, $class) {
        $classes = preg_split('/\s+/', trim($el->getAttribute('class')));
        if (!in_array($class, $classes, true)) {
            $classes[] = $class;
        }
        $el->setAttribute('class', trim(implode(' ', $classes)));
    };
    $setBlock = function (DOMElement $el, $name) use ($addClass) {
        $addClass($el, 'spt-block-' . $name);
        $el->setAttribute('data-spt-block', $name);
    };
    $textOf = function (DOMElement $el) {
        return spt_preview_norm_text($el->textContent);
    };

    // Kop surat: kumpulkan baris/tabel kop yang berurutan sampai judul.
    $kopGroup = [];
    foreach ($children as $el) {
        $text = $textOf($el);
        if ($text === '') {
            continue;
        }
        if (preg_match('/^SURAT\s+TUGAS\b/i', $text)) {
            break;
        }
        if (stripos($text, 'PEMERINTAH') !== false
            || stripos($text, 'DINAS KEHUTANAN') !== false
            || stripos($text, 'CABANG DINAS') !== false
            || stripos($text, 'Jalan Hayam Wuruk') !== false
            || stripos($text, 'Telepon (0353)') !== false) {
            $kopGroup[] = $el;
        }
    }
    if ($kopGroup) {
        $first = $kopGroup[0];
        if ($first->parentNode === $inner) {
            $wrap = $dom->createElement('div');
            $wrap->setAttribute('class', 'spt-kop');
            $wrap->setAttribute('data-spt-block', 'kop');
            $inner->insertBefore($wrap, $first);
            foreach ($kopGroup as $el) {
                $wrap->appendChild($el);
                if ($el->nodeName === 'table') {
                    $addClass($el, 'spt-kop-table');
                } else {
                    $addClass($el, 'spt-kop-line');
                }
            }
        }
    }

    $inDasar = false;
    $inUntuk = false;
    $inLampiran = false;
    $dasarItems = [];
    $untukItems = [];

    foreach ($children as $el) {
        if (!in_array($el->parentNode, [$inner, $inner->firstChild ?? null], true)
            && !$inner->isSameNode($el->parentNode)) {
            // Elemen yang sudah dipindah ke pembungkus kop tidak diproses lagi.
            if (!$xp->query('ancestor::div[contains(@class,"spt-kop")]', $el)->length) {
                // biarkan; elemen lain tetap diproses di bawah
            } else {
                continue;
            }
        }
        $text = $textOf($el);
        if ($text === '') {
            continue;
        }

        // Paragraf sesudah "Lampiran Surat Tugas" bukan bagian badan surat,
        // jadi jangan ditandai sebagai nomor/dasar/untuk.
        if ($inLampiran) {
            if (stripos($text, 'Bojonegoro') !== false && stripos($text, 'NIP') !== false) {
                $addClass($el, 'spt-ttd');
                $setBlock($el, 'ttd');
            }
            continue;
        }

        if ($el->nodeName === 'table') {
            $isEmployee = stripos($text, 'NAMA') !== false
                || stripos($text, 'PANGKAT/GOLONGAN') !== false
                || (stripos($text, 'KEPADA') !== false && stripos($text, 'Nama') !== false);
            if ($isEmployee) {
                $addClass($el, 'spt-pegawai');
                if (stripos($text, 'KEPADA') !== false) {
                    $addClass($el, 'spt-kepada');
                }
                $setBlock($el, 'pegawai');
                $rowIndex = 0;
                foreach ($xp->query('./tr | ./tbody/tr', $el) as $row) {
                    $rowText = $textOf($row);
                    if ($rowText === '') {
                        continue;
                    }
                    if ($rowIndex === 0 && (stripos($rowText, 'NAMA') !== false || preg_match('/^No\b/i', $rowText))) {
                        $addClass($row, 'spt-pegawai-head');
                    } else {
                        $addClass($row, 'spt-pegawai-row');
                    }
                    foreach ($xp->query('.//p', $row) as $p) {
                        $addClass($p, 'spt-pegawai-line');
                    }
                    $rowIndex++;
                }
            } elseif (stripos($text, 'PEMERINTAH') !== false || stripos($text, 'CABANG DINAS') !== false) {
                $addClass($el, 'spt-kop');
                $setBlock($el, 'kop');
                foreach ($xp->query('.//p', $el) as $p) {
                    $addClass($p, 'spt-kop-line');
                }
            }
            continue;
        }

        if (stripos($el->getAttribute('class'), 'spt-ttd-anchor') !== false) {
            continue;
        }

        if (preg_match('/^SURAT\s+TUGAS\b/i', $text)) {
            $addClass($el, 'spt-title');
            $setBlock($el, 'title');
            continue;
        }
        if (preg_match('/^Nomor\b/i', $text)) {
            $addClass($el, 'spt-nomor');
            $setBlock($el, 'nomor');
            continue;
        }
        if (preg_match('/^DASAR\b/i', $text)) {
            $addClass($el, 'spt-dasar');
            $setBlock($el, 'dasar');
            $inDasar = true;
            $inUntuk = false;
            $dasarItems = [];
            continue;
        }
        if ($inDasar && preg_match('/^\d+[\.\)]\s*/', $text)) {
            $addClass($el, 'spt-dasar-item');
            $el->setAttribute('data-spt-item', 'dasar');
            $dasarItems[] = $el;
            continue;
        }
        if (preg_match('/^MEMERINTAHKAN\b/i', $text)) {
            $addClass($el, 'spt-memerintahkan');
            $setBlock($el, 'memerintahkan');
            $inDasar = false;
            $inUntuk = false;
            continue;
        }
        if (preg_match('/^KEPADA\b/i', $text)) {
            $addClass($el, 'spt-kepada');
            $setBlock($el, 'pegawai');
            $inDasar = false;
            $inUntuk = false;
            continue;
        }
        if (preg_match('/^UNTUK\b/i', $text)) {
            $addClass($el, 'spt-untuk');
            $setBlock($el, 'untuk');
            $inDasar = false;
            $inUntuk = true;
            $untukItems = [];
            continue;
        }
        if ($inUntuk && preg_match('/^\d+[\.\)]\s*/', $text)) {
            $addClass($el, 'spt-untuk-item');
            $el->setAttribute('data-spt-item', 'untuk');
            $untukItems[] = $el;
            continue;
        }
        if (preg_match('/^Lampiran\s+Surat\b/i', $text)) {
            $addClass($el, 'spt-lampiran');
            $setBlock($el, 'lampiran');
            $inLampiran = true;
            $inDasar = false;
            $inUntuk = false;
            continue;
        }
        if (stripos($text, 'Bojonegoro') !== false && stripos($text, 'NIP') !== false) {
            $addClass($el, 'spt-ttd');
            $setBlock($el, 'ttd');
        }
    }

    if ($dasarItems) {
        $addClass(end($dasarItems), 'spt-dasar-last');
    }
    if ($untukItems) {
        $addClass(end($untukItems), 'spt-untuk-last');
    }
}

/**
 * CSS pratinjau. Nilai yang bisa diubah memakai CSS variable agar
 * perubahan slider/drag tidak perlu memuat ulang halaman.
 */
function spt_preview_css(array $settings, array $page, $dominantFont)
{
    $defaults = spt_layout_default_values();
    $mmKeys = ['margin_top', 'margin_bottom', 'margin_left', 'margin_right'];
    $vars = [];
    foreach ($defaults as $key => $default) {
        if ($key === 'lock_ttd') {
            continue;
        }
        $value = $settings[$key] ?? $default;
        if (in_array($key, $mmKeys, true)) {
            $string = rtrim(rtrim(number_format((float)$value, 2, '.', ''), '0'), '.') . 'mm';
        } elseif ($key === 'font_size') {
            $string = rtrim(rtrim(number_format((float)$value, 2, '.', ''), '0'), '.') . 'pt';
        } elseif ($key === 'line_height') {
            $string = number_format((float)$value, 2, '.', '');
        } else {
            $string = rtrim(rtrim(number_format((float)$value, 2, '.', ''), '0'), '.') . 'px';
        }
        $vars[] = '--spt-' . str_replace('_', '-', $key) . ':' . $string;
    }
    $factor = $dominantFont > 0 ? ((float)$settings['font_size'] / $dominantFont) : 1.0;
    $vars[] = '--spt-font-factor:' . number_format($factor, 4, '.', '');
    $vars[] = '--spt-page-width:' . $page['width_mm'] . 'mm';
    $vars[] = '--spt-page-height:' . $page['height_mm'] . 'mm';

    $css = '<style>' . PHP_EOL;
    $css .= ':root{' . implode(';', $vars) . '}' . PHP_EOL;
    $css .= 'html,body{margin:0;padding:0;background:#e2e8f0;}' . PHP_EOL;
    $css .= 'body{font-family:Arial,Helvetica,sans-serif;color:#000;overflow-x:hidden;}' . PHP_EOL;
    $css .= '#sptPages{padding:38px 30px 46px 34px;}' . PHP_EOL;
    $css .= '.spt-page{width:var(--spt-page-width);height:var(--spt-page-height);margin:0 auto 18px;background:#fff;'
        . 'box-shadow:0 2px 10px rgba(15,23,42,.18);padding:var(--spt-margin-top) var(--spt-margin-right) '
        . 'var(--spt-margin-bottom) var(--spt-margin-left);box-sizing:border-box;position:relative;'
        // Arial di browser sedikit lebih lebar daripada Arial versi Word;
        // pengurangan tipis ini menjaga titik pembungkusan baris tetap sama.
        . 'letter-spacing:-0.005em;}' . PHP_EOL;
    $css .= '.spt-page-inner{position:relative;}' . PHP_EOL;
    $css .= '.spt-page *{box-sizing:border-box;}' . PHP_EOL;
    // Word menghitung "single" sebagai ±1.15 em; kalikan agar pratinjau
    // setinggi hasil PDF.
    $css .= 'p,h0,h1,h2,h3,h4,h5,h6,h7,h8,h9,td,th,li{font-size:var(--spt-font-size);'
        . 'line-height:calc(var(--spt-line-height) * 1.15)!important;}' . PHP_EOL;
    $css .= 'h0,h1,h2,h3,h4,h5,h6,h7,h8,h9{display:block;font-weight:bold;margin-top:0;margin-bottom:0;}' . PHP_EOL;
    $css .= 'p{margin:0;}' . PHP_EOL;
    $css .= 'p:empty,h0:empty,h1:empty,h2:empty,h3:empty,h4:empty,h5:empty,h6:empty,h7:empty,h8:empty,h9:empty'
        . '{min-height:calc(1em * var(--spt-line-height) * 1.15);}' . PHP_EOL;
    $css .= 'table{border-collapse:collapse;border:0;width:100%;}' . PHP_EOL;
    $css .= 'td,th{vertical-align:top;padding:0 2px;}' . PHP_EOL;
    $css .= '.spt-tab{display:inline-block;width:.5in;}' . PHP_EOL;
    $css .= '.spt-kop{margin-bottom:var(--spt-kop-bottom-spacing)!important;}' . PHP_EOL;
    $css .= '.spt-kop-line,.spt-kop p,.spt-kop td{line-height:calc(1.15em + var(--spt-kop-line-spacing))!important;}' . PHP_EOL;
    // Baris kop di Word hampir selebar area cetak (mis. nama cabang 16pt bold).
    // Arial di browser sedikit lebih lebar sehingga baris itu pecah dua;
    // perapatannya tipis saja agar lebar baris tetap sama dengan Word.
    $css .= '.spt-kop-line{letter-spacing:-0.014em;}' . PHP_EOL;
    $css .= '.spt-title{text-align:center!important;margin-top:var(--spt-judul-top-spacing)!important;margin-bottom:0!important;line-height:1.3!important;}' . PHP_EOL;
    $css .= '.spt-nomor{text-align:center!important;margin-top:var(--spt-nomor-surat-spacing)!important;margin-bottom:var(--spt-judul-bottom-spacing)!important;line-height:1.3!important;}' . PHP_EOL;
    $css .= '.spt-dasar{margin-top:var(--spt-dasar-top-spacing)!important;margin-bottom:var(--spt-dasar-item-spacing)!important;}' . PHP_EOL;
    $css .= '.spt-dasar-item{margin-bottom:var(--spt-dasar-item-spacing)!important;}' . PHP_EOL;
    $css .= '.spt-dasar-last{margin-bottom:var(--spt-dasar-bottom-spacing)!important;}' . PHP_EOL;
    $css .= '.spt-memerintahkan{margin-top:var(--spt-memerintahkan-spacing)!important;margin-bottom:var(--spt-memerintahkan-spacing)!important;}' . PHP_EOL;
    $css .= '.spt-kepada{margin-top:var(--spt-pegawai-top-spacing)!important;}' . PHP_EOL;
    $css .= '.spt-pegawai-line{margin-top:0!important;margin-bottom:var(--spt-pegawai-line-spacing)!important;}' . PHP_EOL;
    $css .= '.spt-pegawai-row + .spt-pegawai-row > td{padding-top:var(--spt-pegawai-gap-spacing)!important;}' . PHP_EOL;
    $css .= '.spt-untuk{margin-top:var(--spt-untuk-top-spacing)!important;margin-bottom:var(--spt-untuk-item-spacing)!important;}' . PHP_EOL;
    $css .= '.spt-untuk-item{margin-bottom:var(--spt-untuk-item-spacing)!important;}' . PHP_EOL;
    $css .= '.spt-untuk-last{margin-bottom:var(--spt-untuk-bottom-spacing)!important;}' . PHP_EOL;
    $css .= '.spt-b-insideh tr + tr > td{border-top:1px solid #000;}' . PHP_EOL;
    $css .= '.spt-b-insidev td + td{border-left:1px solid #000;}' . PHP_EOL;
    $css .= '.spt-b-top{border-top:1px solid #000;}' . PHP_EOL;
    $css .= '.spt-b-left > tbody > tr > td:first-child,.spt-b-left > tr > td:first-child{border-left:1px solid #000;}' . PHP_EOL;
    $css .= '.spt-b-right > tbody > tr > td:last-child,.spt-b-right > tr > td:last-child{border-right:1px solid #000;}' . PHP_EOL;
    $css .= '.spt-b-bottom > tbody > tr:last-child > td,.spt-b-bottom > tr:last-child > td{border-bottom:1px solid #000;}' . PHP_EOL;
    $css .= '.spt-ttd-anchor{position:relative;height:0;}' . PHP_EOL;
    $css .= '.spt-txbx{position:absolute;z-index:3;}' . PHP_EOL;
    $css .= '.spt-txbx-main{left:calc(var(--txbx-left,0mm) + var(--spt-ttd-x));top:calc(var(--txbx-top,0pt) + var(--spt-ttd-offset));}'
        . PHP_EOL;
    $css .= '.spt-txbx-static{left:var(--txbx-left,0mm);top:var(--txbx-top,0pt);}' . PHP_EOL;
    $css .= '.spt-txbx .spt-ttd-line{white-space:pre-wrap;margin:0;}' . PHP_EOL;
    $css .= '.spt-txbx .spt-ttd-date{margin-bottom:var(--spt-ttd-date-to-jabatan)!important;}' . PHP_EOL;
    $css .= '.spt-txbx .spt-ttd-jabatan-last{margin-bottom:var(--spt-ttd-space-height)!important;}' . PHP_EOL;
    $css .= '.spt-txbx .spt-ttd-name{margin-top:var(--spt-ttd-name-nip-spacing)!important;}' . PHP_EOL;
    $css .= '[data-spt-block]{cursor:pointer;}' . PHP_EOL;
    $css .= '[data-spt-block]:hover{outline:1px dashed rgba(5,150,105,.55);outline-offset:2px;}' . PHP_EOL;
    $css .= '[data-spt-block].is-selected{outline:2px solid #059669;outline-offset:2px;background:rgba(16,185,129,.04);}' . PHP_EOL;
    $css .= '.spt-ruler{position:fixed;background:#f8fafc;z-index:70;pointer-events:none;}' . PHP_EOL;
    $css .= '#sptRulerH{left:0;top:0;width:100%;height:22px;border-bottom:1px solid #cbd5e1;}' . PHP_EOL;
    $css .= '#sptRulerV{left:0;top:0;width:22px;height:100%;border-right:1px solid #cbd5e1;}' . PHP_EOL;
    $css .= '.spt-margin-marker{position:fixed;z-index:80;background:#059669;border:1px solid #fff;box-shadow:0 1px 3px rgba(0,0,0,.3);}'
        . PHP_EOL;
    $css .= '.spt-margin-marker.spt-mm-h{width:5px;height:22px;cursor:col-resize;}' . PHP_EOL;
    $css .= '.spt-margin-marker.spt-mm-v{width:22px;height:5px;cursor:row-resize;}' . PHP_EOL;
    $css .= '.spt-gap-handle{position:absolute;left:0;right:0;height:9px;margin-top:-4px;z-index:40;cursor:row-resize;'
        . 'display:flex;align-items:center;justify-content:center;pointer-events:auto;}' . PHP_EOL;
    $css .= '.spt-gap-handle::before{content:"";position:absolute;left:0;right:0;top:4px;border-top:1px dashed transparent;transition:border-color .12s;}'
        . PHP_EOL;
    $css .= '.spt-gap-handle:hover::before,.spt-gap-handle.is-dragging::before{border-top-color:#059669;}' . PHP_EOL;
    $css .= '.spt-gap-pill{position:relative;z-index:1;padding:1px 6px;border-radius:999px;background:#059669;color:#fff;'
        . 'font:600 9px/1.4 Arial,sans-serif;white-space:nowrap;opacity:0;transition:opacity .12s;}' . PHP_EOL;
    $css .= '.spt-gap-handle:hover .spt-gap-pill,.spt-gap-handle.is-dragging .spt-gap-pill{opacity:1;}' . PHP_EOL;
    $css .= '.spt-ttd-move{position:absolute;z-index:45;width:22px;height:22px;border-radius:6px;background:#059669;color:#fff;'
        . 'display:flex;align-items:center;justify-content:center;cursor:move;font:700 12px/1 Arial,sans-serif;'
        . 'box-shadow:0 2px 6px rgba(0,0,0,.25);}' . PHP_EOL;
    $css .= '.spt-ttd-move:hover{background:#047857;}' . PHP_EOL;
    $css .= '@media print{#sptRulerH,#sptRulerV,.spt-gap-handle,.spt-ttd-move,.spt-margin-marker{display:none!important;}}' . PHP_EOL;
    $css .= '</style>';

    return $css;
}

function spt_preview_unit_for_key($key)
{
    if (in_array($key, ['margin_top', 'margin_bottom', 'margin_left', 'margin_right'], true)) {
        return 'mm';
    }
    if ($key === 'font_size') {
        return 'pt';
    }
    if ($key === 'line_height') {
        return '';
    }
    return 'px';
}

/**
 * Cetak halaman pratinjau interaktif lalu hentikan request.
 */
function spt_preview_output($docxPath, array $settings, $tempDir)
{
    $bundle = spt_preview_docx_bundle($docxPath);
    $defaults = spt_layout_read_docx_defaults($docxPath);
    $dominantFont = (float)($defaults['font_size'] ?? 12.0);

    $phpWordHtml = '';
    $htmlPath = $tempDir . DIRECTORY_SEPARATOR . 'spt_preview_' . uniqid() . '.html';
    $oldReporting = error_reporting(E_ALL & ~E_DEPRECATED & ~E_WARNING & ~E_NOTICE);
    try {
        $phpWord = IOFactory::load($docxPath, 'Word2007');
        $writer = IOFactory::createWriter($phpWord, 'HTML');
        $writer->save($htmlPath);
        if (is_file($htmlPath)) {
            $phpWordHtml = (string)file_get_contents($htmlPath);
        }
    } finally {
        error_reporting($oldReporting);
        if (is_file($htmlPath)) {
            @unlink($htmlPath);
        }
    }

    header('Content-Type: text/html; charset=utf-8');
    header('Cache-Control: no-store, max-age=0');

    if ($phpWordHtml === '') {
        echo '<!DOCTYPE html><html lang="id"><meta charset="utf-8"><body style="font-family:Arial;padding:24px;">'
            . '<p>Pratinjau interaktif tidak dapat dibuat. Gunakan mode PDF.</p></body></html>';
        exit;
    }

    $pagesHtml = spt_preview_build_pages($phpWordHtml, $bundle);
    $css = spt_preview_css($settings, $bundle['page'], $dominantFont);

    $ranges = [];
    foreach (spt_layout_ranges() as $key => $range) {
        $ranges[$key] = ['min' => $range[0], 'max' => $range[1], 'step' => $range[2]];
    }
    $units = [];
    foreach (array_keys(spt_layout_default_values()) as $key) {
        if ($key !== 'lock_ttd') {
            $units[$key] = spt_preview_unit_for_key($key);
        }
    }
    $config = [
        'settings' => $settings,
        'ranges' => $ranges,
        'units' => $units,
        'dominantFont' => $dominantFont,
        'pageWidthMm' => $bundle['page']['width_mm'],
        'pageHeightMm' => $bundle['page']['height_mm'],
    ];

    echo '<!DOCTYPE html>' . PHP_EOL;
    echo '<html lang="id"><head><meta charset="utf-8">' . PHP_EOL;
    echo '<meta name="viewport" content="width=device-width, initial-scale=1">' . PHP_EOL;
    echo '<title>Pratinjau Surat Tugas</title>' . PHP_EOL;
    echo $css . PHP_EOL;
    echo '</head><body>' . PHP_EOL;
    echo '<canvas id="sptRulerH" class="spt-ruler"></canvas>' . PHP_EOL;
    echo '<canvas id="sptRulerV" class="spt-ruler"></canvas>' . PHP_EOL;
    echo $pagesHtml . PHP_EOL;
    echo '<script>window.SPT_PREVIEW_CONFIG=' . json_encode($config, JSON_UNESCAPED_UNICODE | JSON_HEX_TAG | JSON_HEX_AMP) . ';</script>' . PHP_EOL;
    echo '<script src="../assets/js/spt-preview.js?v=2"></script>' . PHP_EOL;
    echo '</body></html>';
    exit;
}

if (defined('SPT_TEST_MODE')) {
    return;
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
    $template_id_terpilih = null;
    $template_nama = null;
    $template_kategori_terpilih = null;
    $template_berlaku_mulai = null;
    
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

        // Cek apakah kolom berlaku_mulai ada (versi template pra/pasca PAPBD), jika belum tambahkan
        try {
            $pdo->query("SELECT berlaku_mulai FROM templates LIMIT 1");
        } catch (PDOException $e) {
            try {
                $pdo->exec("ALTER TABLE templates ADD COLUMN berlaku_mulai DATE DEFAULT NULL AFTER tipe_surat");
            } catch (PDOException $e2) {}
        }

        // Tanggal surat menentukan versi template yang dipakai:
        // - berlaku_mulai IS NULL (template lama) -> berlaku untuk semua tanggal
        // - berlaku_mulai terisi (mis. tanggal berlakunya PAPBD, dasar baru) -> hanya untuk
        //   surat dengan tanggal_surat >= berlaku_mulai.
        // Surat bertanggal sebelum masa berlaku template baru otomatis tetap memakai template lama.
        $tgl_surat_key = date('Y-m-d', strtotime($surat['tanggal_surat']));

        // Debug: Log informasi pencarian template
        error_log("=== DEBUG TEMPLATE SELECTION ===");
        error_log("Jumlah Pegawai: " . $jumlah_pegawai);
        error_log("Kategori Template Dicari: " . $kategori_template);
        error_log("Tipe Surat: " . $tipe_surat);
        error_log("Tanggal Surat: " . $tgl_surat_key);
        
        // Debug: Cek semua template yang ada untuk kategori dan tipe ini
        $stmt_debug = $pdo->prepare("SELECT id, nama, kategori_template, tipe_surat, berlaku_mulai, nama_file, is_default FROM templates WHERE kategori_template = ? AND tipe_surat = ?");
        $stmt_debug->execute([$kategori_template, $tipe_surat]);
        $templates_debug = $stmt_debug->fetchAll();
        error_log("Template ditemukan untuk kategori '{$kategori_template}' tipe '{$tipe_surat}': " . count($templates_debug));
        foreach ($templates_debug as $tmpl) {
            error_log("  - ID: {$tmpl['id']}, Nama: {$tmpl['nama']}, File: {$tmpl['nama_file']}, Berlaku: " . ($tmpl['berlaku_mulai'] ?? 'lama') . ", Default: {$tmpl['is_default']}");
        }
        
        // Cari template default untuk kategori dan tipe yang sesuai, yang sudah berlaku
        // pada tanggal surat (template lama selalu berlaku; template baru hanya jika
        // berlaku_mulai <= tanggal_surat). Urutan: masa berlaku terbaru dulu.
        $stmt_template = $pdo->prepare("SELECT id, nama_file, kategori_template, tipe_surat, berlaku_mulai, nama FROM templates WHERE kategori_template = ? AND tipe_surat = ? AND is_default = 1 AND (berlaku_mulai IS NULL OR berlaku_mulai <= ?) ORDER BY berlaku_mulai IS NULL ASC, berlaku_mulai DESC LIMIT 1");
        $stmt_template->execute([$kategori_template, $tipe_surat, $tgl_surat_key]);
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
                            $template_id_terpilih = (int)$template_default['id'];
                            $template_nama = $template_default['nama'];
                            $template_kategori_terpilih = $template_default['kategori_template'];
                            $template_berlaku_mulai = $template_default['berlaku_mulai'] ?? null;
                            error_log("SUKSES: Template default digunakan - {$template_nama} (kategori: {$template_kategori_terpilih}, berlaku: " . ($template_berlaku_mulai ?? 'lama') . ", tgl_surat: {$tgl_surat_key})");
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
        
        // Jika tidak ada default yang valid, ambil template terbaru yang sudah berlaku
        // pada tanggal surat dari kategori dan tipe tersebut
        if (!$template_found) {
            $stmt_template = $pdo->prepare("SELECT id, nama_file, kategori_template, tipe_surat, berlaku_mulai, nama FROM templates WHERE kategori_template = ? AND tipe_surat = ? AND (berlaku_mulai IS NULL OR berlaku_mulai <= ?) ORDER BY berlaku_mulai IS NULL ASC, berlaku_mulai DESC, is_default DESC, created_at DESC LIMIT 1");
            $stmt_template->execute([$kategori_template, $tipe_surat, $tgl_surat_key]);
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
                                $template_id_terpilih = (int)$template_fallback['id'];
                                $template_nama = $template_fallback['nama'];
                                $template_kategori_terpilih = $template_fallback['kategori_template'];
                                $template_berlaku_mulai = $template_fallback['berlaku_mulai'] ?? null;
                                error_log("SUKSES: Template fallback digunakan - {$template_nama} (kategori: {$template_kategori_terpilih}, berlaku: " . ($template_berlaku_mulai ?? 'lama') . ", tgl_surat: {$tgl_surat_key})");
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
            
            $stmt_fallback = $pdo->prepare("SELECT id, nama_file, kategori_template, tipe_surat, berlaku_mulai, nama FROM templates WHERE kategori_template = 'banyak_pegawai' AND tipe_surat = ? AND is_default = 1 AND (berlaku_mulai IS NULL OR berlaku_mulai <= ?) ORDER BY berlaku_mulai IS NULL ASC, berlaku_mulai DESC LIMIT 1");
            $stmt_fallback->execute([$tipe_surat, $tgl_surat_key]);
            $template_fallback_universal = $stmt_fallback->fetch();
            
            if (!$template_fallback_universal) {
                // Jika tidak ada default, ambil yang pertama dengan tipe yang sama
                $stmt_fallback = $pdo->prepare("SELECT id, nama_file, kategori_template, tipe_surat, berlaku_mulai, nama FROM templates WHERE kategori_template = 'banyak_pegawai' AND tipe_surat = ? AND (berlaku_mulai IS NULL OR berlaku_mulai <= ?) ORDER BY berlaku_mulai IS NULL ASC, berlaku_mulai DESC, is_default DESC, created_at DESC LIMIT 1");
                $stmt_fallback->execute([$tipe_surat, $tgl_surat_key]);
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
                            $template_id_terpilih = (int)$template_fallback_universal['id'];
                            $template_nama = $template_fallback_universal['nama'] . " (Fallback)";
                            $template_kategori_terpilih = $template_fallback_universal['kategori_template'];
                            $template_berlaku_mulai = $template_fallback_universal['berlaku_mulai'] ?? null;
                            error_log("SUKSES: Menggunakan template fallback 'banyak_pegawai' tipe '{$tipe_surat}' - {$template_nama}");
                        }
                    }
                }
            }
        }

        // FALLBACK 2: Jika masih tidak ditemukan dan tipe_surat bukan umum, coba dengan tipe umum
        if (!$template_found && $tipe_surat !== 'umum') {
            error_log("INFO: Template untuk tipe '{$tipe_surat}' tidak ditemukan, mencoba fallback ke tipe 'umum' kategori '{$kategori_template}'");
            
            $stmt_fallback2 = $pdo->prepare("SELECT id, nama_file, kategori_template, tipe_surat, berlaku_mulai, nama FROM templates WHERE kategori_template = ? AND tipe_surat = 'umum' AND is_default = 1 AND (berlaku_mulai IS NULL OR berlaku_mulai <= ?) ORDER BY berlaku_mulai IS NULL ASC, berlaku_mulai DESC LIMIT 1");
            $stmt_fallback2->execute([$kategori_template, $tgl_surat_key]);
            $template_fallback2 = $stmt_fallback2->fetch();
            
            if (!$template_fallback2) {
                $stmt_fallback2 = $pdo->prepare("SELECT id, nama_file, kategori_template, tipe_surat, berlaku_mulai, nama FROM templates WHERE kategori_template = ? AND tipe_surat = 'umum' AND (berlaku_mulai IS NULL OR berlaku_mulai <= ?) ORDER BY berlaku_mulai IS NULL ASC, berlaku_mulai DESC, is_default DESC, created_at DESC LIMIT 1");
                $stmt_fallback2->execute([$kategori_template, $tgl_surat_key]);
                $template_fallback2 = $stmt_fallback2->fetch();
            }

            if (!$template_fallback2 && $kategori_template !== 'banyak_pegawai') {
                $stmt_fallback2 = $pdo->prepare("SELECT id, nama_file, kategori_template, tipe_surat, berlaku_mulai, nama FROM templates WHERE kategori_template = 'banyak_pegawai' AND tipe_surat = 'umum' AND is_default = 1 AND (berlaku_mulai IS NULL OR berlaku_mulai <= ?) ORDER BY berlaku_mulai IS NULL ASC, berlaku_mulai DESC LIMIT 1");
                $stmt_fallback2->execute([$tgl_surat_key]);
                $template_fallback2 = $stmt_fallback2->fetch();
            }

            if (!$template_fallback2 && $kategori_template !== 'banyak_pegawai') {
                $stmt_fallback2 = $pdo->prepare("SELECT id, nama_file, kategori_template, tipe_surat, berlaku_mulai, nama FROM templates WHERE kategori_template = 'banyak_pegawai' AND tipe_surat = 'umum' AND (berlaku_mulai IS NULL OR berlaku_mulai <= ?) ORDER BY berlaku_mulai IS NULL ASC, berlaku_mulai DESC, is_default DESC, created_at DESC LIMIT 1");
                $stmt_fallback2->execute([$tgl_surat_key]);
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
                            $template_id_terpilih = (int)$template_fallback2['id'];
                            $template_nama = $template_fallback2['nama'] . " (Fallback Umum)";
                            $template_kategori_terpilih = $template_fallback2['kategori_template'];
                            $template_berlaku_mulai = $template_fallback2['berlaku_mulai'] ?? null;
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

    // Info versi template terpilih untuk ditampilkan di halaman preview.
    // $template_berlaku_mulai == null -> template lama (pra-PAPBD);
    // terisi tanggal -> template dasar baru yang berlaku mulai tanggal tersebut.
    $info_versi_template = null;
    $peringatan_template_lama = false;
    try {
        $stmt_cut = $pdo->query("SELECT MIN(berlaku_mulai) AS cutoff FROM templates WHERE berlaku_mulai IS NOT NULL");
        $cutoff_row = $stmt_cut ? $stmt_cut->fetch() : null;
        $cutoff_global = ($cutoff_row && !empty($cutoff_row['cutoff'])) ? $cutoff_row['cutoff'] : null;
    } catch (PDOException $e) {
        $cutoff_global = null;
    }
    if (!empty($template_berlaku_mulai)) {
        $info_versi_template = 'Dasar baru (berlaku mulai ' . date('d/m/Y', strtotime($template_berlaku_mulai)) . ')';
    } else {
        $info_versi_template = 'Dasar lama (pra-PAPBD)';
        // Surat bertanggal sama/lewat cutoff tapi masih memakai template lama:
        // kemungkinan template dasar baru untuk kombinasi ini belum di-upload.
        if ($cutoff_global && $tgl_surat_key >= $cutoff_global) {
            $peringatan_template_lama = true;
        }
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
            $n = abs((int)$n);
            $dasar = ['', 'satu', 'dua', 'tiga', 'empat', 'lima', 'enam', 'tujuh', 'delapan', 'sembilan', 'sepuluh', 'sebelas'];
            
            if ($n < 12) {
                return $dasar[$n];
            } elseif ($n < 20) {
                return $dasar[$n - 10] . ' belas';
            } elseif ($n < 100) {
                $puluh = (int)floor($n / 10);
                $sisa = $n % 10;
                return trim($dasar[$puluh] . ' puluh ' . ($sisa ? $dasar[$sisa] : ''));
            } elseif ($n < 200) {
                return trim('seratus ' . ($n - 100 > 0 ? angkaTerbilang($n - 100) : ''));
            } elseif ($n < 1000) {
                $ratus = (int)floor($n / 100);
                $sisa = $n % 100;
                return trim($dasar[$ratus] . ' ratus ' . ($sisa ? angkaTerbilang($sisa) : ''));
            } elseif ($n < 2000) {
                return trim('seribu ' . ($n - 1000 > 0 ? angkaTerbilang($n - 1000) : ''));
            } elseif ($n < 1000000) {
                $ribu = (int)floor($n / 1000);
                $sisa = $n % 1000;
                return trim(angkaTerbilang($ribu) . ' ribu ' . ($sisa ? angkaTerbilang($sisa) : ''));
            }
            return (string)$n;
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
    $dasarSurat = spt_clean_repeated_words(spt_fix_common_typos($dasarSurat));
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

    $untukSurat = spt_clean_repeated_words(spt_fix_common_typos((string)($surat['untuk'] ?? '')));
    $templateProcessor->setValue('untuk', $untukSurat);
    $templateProcessor->setValue('kota_tujuan', isset($surat['kota_tujuan']) ? (string)$surat['kota_tujuan'] : '');
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

    // Parameter layout dari panel adjuster. Jika template sudah punya setting
    // default tersimpan, setting itu otomatis dipakai saat halaman dibuka.
    $layout_explicit_settings = spt_layout_request_settings();
    $layout_defaults = spt_layout_read_docx_defaults($template_path);
    $layout_saved_settings = spt_layout_load_template_settings($pdo, $template_id_terpilih);
    $layout_has_saved_default = $layout_saved_settings !== null;
    if ($layout_has_saved_default) {
        $layout_defaults = array_merge($layout_defaults, $layout_saved_settings);
    }

    $layout_active = $layout_explicit_settings !== null || $layout_has_saved_default;
    if ($layout_explicit_settings !== null) {
        // Lengkapi parameter yang tidak ada di URL dengan nilai default template.
        $layout_settings = array_merge($layout_defaults, $layout_explicit_settings);
    } elseif ($layout_has_saved_default) {
        $layout_settings = $layout_defaults;
    } else {
        $layout_settings = null;
    }
    $layout_initial = $layout_settings !== null ? $layout_settings : $layout_defaults;

    // Preview (tanpa format): tampilkan halaman dengan iframe PDF agar sama persis dengan file Word/PDF
    if ($format === '') {
        $sid = (int)$_GET['id'];
        $previewQuery = ['id' => $sid, 'format' => 'pdf', 'inline' => '1'];
        $downloadQuery = ['id' => $sid, 'format' => 'pdf'];
        if ($layout_active) {
            $previewQuery = array_merge($previewQuery, spt_layout_query_params($layout_settings));
            $downloadQuery = array_merge($downloadQuery, spt_layout_query_params($layout_settings));
        }
        $previewPdfUrl = 'generate-surat.php?' . http_build_query($previewQuery);
        $downloadPdfUrl = 'generate-surat.php?' . http_build_query($downloadQuery);
        $layoutJsonDefaults = json_encode($layout_defaults, JSON_UNESCAPED_UNICODE);
        $layoutJsonInitial = json_encode($layout_initial, JSON_UNESCAPED_UNICODE);

        include '../includes/header.php';
        ?>
<style>
main { overflow: hidden !important; padding: 0 !important; }
main > .max-w-7xl { max-width: none !important; margin: 0 !important; height: 100%; }
.spt-app-shell { height: 100%; min-height: 0; display: flex; flex-direction: column; overflow: hidden; }
.spt-split { display: flex; flex: 1 1 auto; min-height: 0; min-width: 0; height: 100%; overflow: hidden; background: #f1f5f9; }
.spt-sidebar-collapsed { display: none !important; }
.spt-hidden { display: none !important; }

.spt-inspector { width: 380px; flex: 0 0 380px; min-width: 320px; max-width: 520px; display: flex; flex-direction: column; height: 100%; min-height: 0; background: #fff; border-right: 1px solid #e2e8f0; }
.spt-inspector-head { flex: 0 0 auto; padding: 13px 16px 11px; border-bottom: 1px solid #e2e8f0; background: #fff; }
.spt-inspector-title { display: flex; align-items: center; gap: 7px; font-size: 15px; font-weight: 700; color: #0f172a; letter-spacing: -.01em; }
.spt-inspector-sub { margin-top: 2px; font-size: 10.5px; line-height: 1.4; color: #64748b; }
.spt-preset-row { display: grid; grid-template-columns: 1fr 1fr; gap: 7px; margin-top: 9px; }
.spt-preset-btn { position: relative; display: inline-flex; align-items: center; justify-content: center; gap: 5px; padding: 7px 8px; border-radius: 8px; border: 1px solid #e2e8f0; background: #fff; color: #334155; font-size: 10.5px; font-weight: 600; transition: .15s ease; cursor: pointer; }
.spt-preset-btn:hover { background: #f8fafc; border-color: #cbd5e1; }
.spt-preset-btn.is-active { border-color: #10b981; background: #ecfdf5; color: #047857; }
.spt-preset-btn.is-loading { color: transparent !important; pointer-events: none; }
.spt-preset-btn.is-loading > * { visibility: hidden; }
.spt-preset-btn.is-loading::after { content: ''; position: absolute; width: 14px; height: 14px; border: 2px solid #10b981; border-top-color: transparent; border-radius: 999px; animation: spt-spin .7s linear infinite; }
.spt-version-line { display: flex; align-items: flex-start; gap: 6px; margin-top: 8px; font-size: 10px; line-height: 1.35; color: #64748b; }
.spt-inspector-body { flex: 1 1 auto; min-height: 0; overflow-y: auto; overscroll-behavior: contain; padding: 11px 4px 16px 16px; scrollbar-gutter: stable; }
.spt-inspector-foot { flex: 0 0 auto; position: sticky; bottom: 0; z-index: 20; padding: 11px 16px 12px; background: #fff; border-top: 1px solid #e2e8f0; box-shadow: 0 -8px 18px -14px rgba(15,23,42,.35); }

.spt-acc { border: 1px solid #e2e8f0; border-radius: 10px; background: #fff; margin-bottom: 7px; overflow: hidden; }
.spt-acc-head { width: 100%; display: flex; align-items: center; justify-content: space-between; gap: 8px; padding: 8px 11px; background: #f8fafc; font-size: 11px; font-weight: 600; color: #334155; cursor: pointer; transition: background .15s ease; }
.spt-acc-head:hover { background: #f1f5f9; }
.spt-acc-body { padding: 10px 11px 11px; border-top: 1px solid #f1f5f9; }
.spt-acc-body.is-collapsed { display: none; }
.spt-acc-icon { transition: transform .15s ease; }
.spt-acc-head[aria-expanded="true"] .spt-acc-icon { transform: rotate(180deg); }
.spt-acc-dot { width: 6px; height: 6px; border-radius: 999px; background: #10b981; box-shadow: 0 0 0 3px #d1fae5; flex: 0 0 auto; }
.spt-acc-count { min-width: 17px; height: 17px; padding: 0 5px; border-radius: 999px; background: #d1fae5; color: #047857; font-size: 9.5px; font-weight: 700; display: inline-flex; align-items: center; justify-content: center; flex: 0 0 auto; }

.spt-control { display: flex; flex-direction: column; gap: 3px; }
.spt-control-head { display: flex; align-items: center; justify-content: space-between; gap: 6px; }
.spt-control-label { font-size: 10.5px; font-weight: 500; color: #334155; line-height: 1.25; }
.spt-reset-btn { min-width: 18px; height: 18px; display: inline-flex; align-items: center; justify-content: center; border-radius: 5px; color: #94a3b8; font-size: 12px; line-height: 1; background: transparent; border: 0; cursor: pointer; transition: .15s ease; }
.spt-reset-btn:hover { background: #f1f5f9; color: #0f172a; }
.spt-control-row { display: flex; align-items: center; gap: 4px; }
.spt-step-btn { width: 22px; height: 22px; flex: 0 0 22px; display: inline-flex; align-items: center; justify-content: center; border: 1px solid #e2e8f0; border-radius: 6px; background: #fff; color: #475569; font-size: 13px; line-height: 1; cursor: pointer; }
.spt-step-btn:hover { background: #f8fafc; }
.spt-step-btn:active { background: #f1f5f9; }
.spt-number-wrap { position: relative; flex: 1 1 auto; min-width: 0; }
.spt-number { width: 100%; height: 22px; padding: 0 19px 0 5px; border: 1px solid #e2e8f0; border-radius: 6px; background: #fff; text-align: center; font-size: 10.5px; font-weight: 600; color: #0f172a; font-variant-numeric: tabular-nums; -moz-appearance: textfield; }
.spt-number::-webkit-outer-spin-button, .spt-number::-webkit-inner-spin-button { -webkit-appearance: none; margin: 0; }
.spt-number:focus { outline: none; border-color: #10b981; box-shadow: 0 0 0 2px #d1fae5; }
.spt-unit { position: absolute; right: 4px; top: 50%; transform: translateY(-50%); font-size: 8.5px; color: #94a3b8; pointer-events: none; }
.spt-range { width: 100%; height: 13px; margin: 0; accent-color: #059669; cursor: pointer; }
.spt-help { font-size: 9.5px; color: #94a3b8; line-height: 1.3; }
.spt-grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 9px 8px; }

.spt-apply-btn { width: 100%; display: inline-flex; align-items: center; justify-content: center; gap: 6px; padding: 9px 10px; border-radius: 8px; border: 0; background: #1e293b; color: #fff; font-size: 11.5px; font-weight: 600; cursor: pointer; transition: .15s ease; }
.spt-apply-btn:hover { background: #0f172a; }
.spt-apply-btn:disabled { opacity: .75; cursor: wait; }
.spt-apply-btn.is-dirty { background: #059669; box-shadow: 0 0 0 3px #d1fae5; }
.spt-apply-btn.is-dirty:hover { background: #047857; }
.spt-btn-spinner { width: 12px; height: 12px; flex: 0 0 12px; border: 2px solid rgba(255,255,255,.45); border-top-color: #fff; border-radius: 999px; animation: spt-spin .7s linear infinite; }
.spt-save-btn { width: 100%; display: inline-flex; align-items: center; justify-content: center; gap: 6px; margin-top: 6px; padding: 8px 10px; border-radius: 8px; border: 1px solid #a7f3d0; background: #fff; color: #047857; font-size: 11.5px; font-weight: 600; cursor: pointer; transition: .15s ease; }
.spt-save-btn:hover { background: #ecfdf5; }
.spt-save-btn:disabled { opacity: .7; cursor: wait; }
.spt-foot-note { margin-top: 6px; text-align: center; font-size: 9.5px; color: #94a3b8; }

.spt-preview { flex: 1 1 auto; min-width: 0; min-height: 0; display: flex; flex-direction: column; height: 100%; overflow: hidden; background: #f1f5f9; }
.spt-toolbar { flex: 0 0 auto; display: flex; align-items: center; justify-content: space-between; gap: 8px; padding: 7px 12px; background: #fff; border-bottom: 1px solid #e2e8f0; }
.spt-toolbar-left, .spt-toolbar-center, .spt-toolbar-right { display: flex; align-items: center; gap: 7px; min-width: 0; }
.spt-toolbar-left { flex: 1 1 auto; }
.spt-toolbar-center { flex: 0 0 auto; }
.spt-toolbar-right { flex: 0 0 auto; }
.spt-icon-btn { width: 28px; height: 28px; flex: 0 0 28px; display: inline-flex; align-items: center; justify-content: center; border: 1px solid #e2e8f0; border-radius: 7px; background: #fff; color: #475569; cursor: pointer; }
.spt-icon-btn:hover { background: #f8fafc; }
.spt-back-link { display: inline-flex; align-items: center; gap: 4px; font-size: 11px; font-weight: 600; color: #475569; white-space: nowrap; text-decoration: none; }
.spt-back-link:hover { color: #059669; }
.spt-doc-id { font-size: 10.5px; color: #64748b; font-weight: 500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.spt-badge { display: inline-flex; align-items: center; gap: 4px; padding: 3px 8px; border-radius: 999px; font-size: 10px; font-weight: 600; white-space: nowrap; }
.spt-badge-neutral { background: #f1f5f9; color: #475569; }
.spt-badge-amber { background: #fef3c7; color: #92400e; }
.spt-badge-emerald { background: #d1fae5; color: #047857; }
.spt-mode { display: inline-flex; align-items: center; gap: 2px; padding: 2px; border: 1px solid #e2e8f0; border-radius: 8px; background: #f8fafc; }
.spt-mode-btn { display: inline-flex; align-items: center; gap: 4px; padding: 4px 9px; border: 0; border-radius: 6px; background: transparent; color: #64748b; font-size: 10.5px; font-weight: 600; white-space: nowrap; cursor: pointer; transition: .15s ease; }
.spt-mode-btn:hover { color: #0f172a; }
.spt-mode-btn.is-active { background: #fff; color: #047857; box-shadow: 0 1px 3px rgba(15,23,42,.14); }
.spt-btn { display: inline-flex; align-items: center; justify-content: center; gap: 4px; padding: 6px 9px; border-radius: 7px; font-size: 10.5px; font-weight: 600; white-space: nowrap; transition: .15s ease; text-decoration: none; cursor: pointer; }
.spt-btn-emerald { background: #059669; color: #fff; border: 1px solid #059669; }
.spt-btn-emerald:hover { background: #047857; border-color: #047857; }
.spt-btn-outline { border: 1px solid #cbd5e1; background: #fff; color: #334155; }
.spt-btn-outline:hover { background: #f8fafc; }
.spt-canvas { position: relative; flex: 1 1 auto; min-height: 0; padding: 11px; display: flex; overflow: hidden; }
.spt-frame { flex: 1 1 auto; width: 100%; height: 100%; border: 1px solid #e2e8f0; border-radius: 10px; background: #fff; }
.spt-overlay { position: absolute; inset: 11px; z-index: 10; display: flex; align-items: center; justify-content: center; border-radius: 10px; background: rgba(255,255,255,.82); backdrop-filter: blur(2px); }
.spt-spinner { width: 32px; height: 32px; border: 3px solid #059669; border-top-color: transparent; border-radius: 999px; animation: spt-spin .7s linear infinite; }
.spt-resizer { flex: 0 0 5px; width: 5px; cursor: col-resize; background: transparent; position: relative; }
.spt-resizer::after { content: ''; position: absolute; top: 50%; left: 1px; width: 2px; height: 36px; transform: translateY(-50%); border-radius: 2px; background: #cbd5e1; }
.spt-resizer:hover::after, .spt-resizer.is-dragging::after { background: #059669; }
@keyframes spt-spin { to { transform: rotate(360deg); } }

@media (max-width: 1023px) {
    .spt-split { flex-direction: column; }
    .spt-inspector { width: 100% !important; flex: 0 0 auto; max-width: none; max-height: 46vh; border-right: 0; border-bottom: 1px solid #e2e8f0; }
    .spt-resizer { display: none; }
    .spt-preview { flex: 1 1 auto; min-height: 52vh; }
    .spt-toolbar { flex-wrap: wrap; }
    .spt-toolbar-center { order: 3; width: 100%; justify-content: center; }
}
</style>

<div class="spt-app-shell" id="sptAppShell">
    <div class="spt-split">
        <aside class="spt-inspector" id="sptInspector">
            <div class="spt-inspector-head">
                <div style="display:flex;align-items:flex-start;justify-content:space-between;gap:8px;">
                    <div style="min-width:0;">
                        <h1 class="spt-inspector-title">
                            <i class="bx bx-slider-alt" style="color:#059669;font-size:16px;"></i>
                            <span>Penyesuaian Tata Letak</span>
                        </h1>
                        <p class="spt-inspector-sub">Geser slider atau tarik blok/margin langsung di pratinjau. Mode Interaktif langsung memperbarui tampilan.</p>
                    </div>
                </div>
                <div class="spt-preset-row">
                    <button type="button" id="layoutAutoFit" class="spt-preset-btn" title="Kompres tata letak agar pas 1 halaman">
                        <i class="bx bx-collapse-vertical" style="font-size:13px;"></i>
                        <span>Paksa Muat 1 Halaman</span>
                    </button>
                    <button type="button" id="layoutReset" class="spt-preset-btn" title="Kembalikan ke default template">
                        <i class="bx bx-reset" style="font-size:13px;"></i>
                        <span>Reset ke Default</span>
                    </button>
                </div>
                <div class="spt-version-line">
                    <?php if (!empty($template_berlaku_mulai)): ?>
                        <i class="bx bx-calendar-check" style="color:#059669;font-size:13px;margin-top:1px;"></i>
                        <span><strong>Template dasar baru (PAPBD)</strong> &mdash; <?= htmlspecialchars($template_nama ?? basename($template_path)) ?>, berlaku mulai <?= date('d/m/Y', strtotime($template_berlaku_mulai)) ?>.</span>
                    <?php elseif (!empty($peringatan_template_lama)): ?>
                        <i class="bx bx-error" style="color:#d97706;font-size:13px;margin-top:1px;"></i>
                        <span><strong>Template dasar lama</strong> &mdash; <?= htmlspecialchars($template_nama ?? basename($template_path)) ?>; template dasar baru untuk kategori ini belum di-upload.</span>
                    <?php else: ?>
                        <i class="bx bx-history" style="color:#64748b;font-size:13px;margin-top:1px;"></i>
                        <span><strong>Template dasar lama</strong> &mdash; <?= htmlspecialchars($template_nama ?? basename($template_path)) ?>, sesuai tanggal surat <?= date('d/m/Y', strtotime($surat['tanggal_surat'])) ?>.</span>
                    <?php endif; ?>
                </div>
            </div>

            <div class="spt-inspector-body" id="layoutPanelBody">
                <div style="display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:7px;">
                    <span style="font-size:9.5px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:#94a3b8;">Pengaturan Blok</span>
                    <button type="button" id="layoutExpandAll" class="spt-reset-btn" style="width:auto;padding:0 6px;font-size:10px;">Tutup Semua</button>
                </div>
                <div id="layoutAccordions">
                    <div class="spt-acc" data-keys="margin_top margin_bottom margin_left margin_right font_size line_height">
                        <button type="button" class="layout-accordion-btn spt-acc-head" data-target="acc-global" aria-expanded="true">
                            <span style="display:flex;align-items:center;gap:6px;min-width:0;">
                                <i class="bx bx-font" style="color:#059669;font-size:13px;"></i>
                                <span>Halaman &amp; Huruf Global</span>
                                <span class="spt-acc-dot spt-hidden" data-acc-dot></span>
                                <span class="spt-acc-count spt-hidden" data-acc-count></span>
                            </span>
                            <i class="bx bx-chevron-down spt-acc-icon" style="font-size:14px;"></i>
                        </button>
                        <div id="acc-global" class="layout-accordion-content spt-acc-body">
                            <div class="spt-grid-2">
                                <?= spt_layout_render_control('margin_top', 'Margin Atas', $layout_initial, 'mm') ?>
                                <?= spt_layout_render_control('margin_bottom', 'Margin Bawah', $layout_initial, 'mm') ?>
                                <?= spt_layout_render_control('margin_left', 'Margin Kiri', $layout_initial, 'mm') ?>
                                <?= spt_layout_render_control('margin_right', 'Margin Kanan', $layout_initial, 'mm') ?>
                            </div>
                            <div style="display:flex;flex-direction:column;gap:10px;margin-top:10px;">
                                <?= spt_layout_render_control('font_size', 'Ukuran Huruf Global', $layout_initial, 'pt') ?>
                                <?= spt_layout_render_control('line_height', 'Jarak Antar Baris', $layout_initial, '×') ?>
                            </div>
                        </div>
                    </div>

                    <div class="spt-acc" data-keys="kop_line_spacing kop_bottom_spacing">
                        <button type="button" class="layout-accordion-btn spt-acc-head" data-target="acc-kop" aria-expanded="false">
                            <span style="display:flex;align-items:center;gap:6px;min-width:0;">
                                <i class="bx bx-image" style="color:#059669;font-size:13px;"></i>
                                <span>Kop Surat</span>
                                <span class="spt-acc-dot spt-hidden" data-acc-dot></span>
                                <span class="spt-acc-count spt-hidden" data-acc-count></span>
                            </span>
                            <i class="bx bx-chevron-down spt-acc-icon" style="font-size:14px;"></i>
                        </button>
                        <div id="acc-kop" class="layout-accordion-content spt-acc-body is-collapsed">
                            <div style="display:flex;flex-direction:column;gap:10px;">
                                <?= spt_layout_render_control('kop_line_spacing', 'Kerapatan Baris Kop', $layout_initial, 'px', 'Tambahan jarak antar baris di dalam kop (0 = ikuti template).') ?>
                                <?= spt_layout_render_control('kop_bottom_spacing', 'Jarak Bawah Kop', $layout_initial, 'px', 'Jarak dari baris terakhir kop ke judul SURAT TUGAS.') ?>
                            </div>
                        </div>
                    </div>

                    <div class="spt-acc" data-keys="judul_top_spacing nomor_surat_spacing judul_bottom_spacing">
                        <button type="button" class="layout-accordion-btn spt-acc-head" data-target="acc-judul" aria-expanded="false">
                            <span style="display:flex;align-items:center;gap:6px;min-width:0;">
                                <i class="bx bx-heading" style="color:#059669;font-size:13px;"></i>
                                <span>Judul &amp; Nomor Surat</span>
                                <span class="spt-acc-dot spt-hidden" data-acc-dot></span>
                                <span class="spt-acc-count spt-hidden" data-acc-count></span>
                            </span>
                            <i class="bx bx-chevron-down spt-acc-icon" style="font-size:14px;"></i>
                        </button>
                        <div id="acc-judul" class="layout-accordion-content spt-acc-body is-collapsed">
                            <div style="display:flex;flex-direction:column;gap:10px;">
                                <?= spt_layout_render_control('judul_top_spacing', 'Spasi Atas SURAT TUGAS', $layout_initial, 'px') ?>
                                <?= spt_layout_render_control('nomor_surat_spacing', 'Jarak Judul ke Nomor', $layout_initial, 'px', 'Minimum 2 px; line-height judul dikunci 1.3 agar tidak bertumpuk.') ?>
                                <?= spt_layout_render_control('judul_bottom_spacing', 'Jarak Sesudah Nomor', $layout_initial, 'px') ?>
                            </div>
                        </div>
                    </div>

                    <div class="spt-acc" data-keys="dasar_top_spacing dasar_item_spacing dasar_bottom_spacing memerintahkan_spacing">
                        <button type="button" class="layout-accordion-btn spt-acc-head" data-target="acc-dasar" aria-expanded="false">
                            <span style="display:flex;align-items:center;gap:6px;min-width:0;">
                                <i class="bx bx-list-ol" style="color:#059669;font-size:13px;"></i>
                                <span>Blok Dasar &amp; Memerintahkan</span>
                                <span class="spt-acc-dot spt-hidden" data-acc-dot></span>
                                <span class="spt-acc-count spt-hidden" data-acc-count></span>
                            </span>
                            <i class="bx bx-chevron-down spt-acc-icon" style="font-size:14px;"></i>
                        </button>
                        <div id="acc-dasar" class="layout-accordion-content spt-acc-body is-collapsed">
                            <div style="display:flex;flex-direction:column;gap:10px;">
                                <?= spt_layout_render_control('dasar_top_spacing', 'Jarak Sebelum DASAR', $layout_initial, 'px') ?>
                                <?= spt_layout_render_control('dasar_item_spacing', 'Kerapatan Butir Dasar', $layout_initial, 'px') ?>
                                <?= spt_layout_render_control('dasar_bottom_spacing', 'Jarak Sesudah DASAR', $layout_initial, 'px') ?>
                                <?= spt_layout_render_control('memerintahkan_spacing', 'Spasi MEMERINTAHKAN', $layout_initial, 'px') ?>
                            </div>
                        </div>
                    </div>

                    <div class="spt-acc" data-keys="pegawai_top_spacing pegawai_gap_spacing pegawai_line_spacing">
                        <button type="button" class="layout-accordion-btn spt-acc-head" data-target="acc-pegawai" aria-expanded="false">
                            <span style="display:flex;align-items:center;gap:6px;min-width:0;">
                                <i class="bx bx-user" style="color:#059669;font-size:13px;"></i>
                                <span>Blok Pegawai</span>
                                <span class="spt-acc-dot spt-hidden" data-acc-dot></span>
                                <span class="spt-acc-count spt-hidden" data-acc-count></span>
                            </span>
                            <i class="bx bx-chevron-down spt-acc-icon" style="font-size:14px;"></i>
                        </button>
                        <div id="acc-pegawai" class="layout-accordion-content spt-acc-body is-collapsed">
                            <div style="display:flex;flex-direction:column;gap:10px;">
                                <?= spt_layout_render_control('pegawai_top_spacing', 'Jarak Sebelum KEPADA', $layout_initial, 'px') ?>
                                <?= spt_layout_render_control('pegawai_gap_spacing', 'Jarak Antar Pegawai', $layout_initial, 'px', 'Jarak pemisah antar orang pada daftar pegawai.') ?>
                                <?= spt_layout_render_control('pegawai_line_spacing', 'Kerapatan Baris Identitas', $layout_initial, 'px', 'Jarak antar baris Nama, NIP, Pangkat, Jabatan.') ?>
                            </div>
                        </div>
                    </div>

                    <div class="spt-acc" data-keys="untuk_top_spacing untuk_item_spacing untuk_bottom_spacing">
                        <button type="button" class="layout-accordion-btn spt-acc-head" data-target="acc-untuk" aria-expanded="false">
                            <span style="display:flex;align-items:center;gap:6px;min-width:0;">
                                <i class="bx bx-task" style="color:#059669;font-size:13px;"></i>
                                <span>Blok UNTUK &amp; Butir Tugas</span>
                                <span class="spt-acc-dot spt-hidden" data-acc-dot></span>
                                <span class="spt-acc-count spt-hidden" data-acc-count></span>
                            </span>
                            <i class="bx bx-chevron-down spt-acc-icon" style="font-size:14px;"></i>
                        </button>
                        <div id="acc-untuk" class="layout-accordion-content spt-acc-body is-collapsed">
                            <p class="spt-help" style="margin-bottom:8px;">Berlaku untuk label <strong>UNTUK</strong> dan butir <strong>&ldquo;1. Melaksanakan&hellip;&rdquo;</strong> di sebelahnya, termasuk butir-butir berikutnya.</p>
                            <div style="display:flex;flex-direction:column;gap:10px;">
                                <?= spt_layout_render_control('untuk_top_spacing', 'Jarak Sebelum Blok UNTUK', $layout_initial, 'px') ?>
                                <?= spt_layout_render_control('untuk_item_spacing', 'Kerapatan Antar Butir Tugas', $layout_initial, 'px', 'Termasuk jarak dari butir pertama di samping label UNTUK.') ?>
                                <?= spt_layout_render_control('untuk_bottom_spacing', 'Jarak Sesudah Blok Tugas', $layout_initial, 'px') ?>
                            </div>
                        </div>
                    </div>

                    <div class="spt-acc" data-keys="ttd_offset ttd_x ttd_date_to_jabatan ttd_space_height ttd_name_nip_spacing lock_ttd">
                        <button type="button" class="layout-accordion-btn spt-acc-head" data-target="acc-ttd" aria-expanded="false">
                            <span style="display:flex;align-items:center;gap:6px;min-width:0;">
                                <i class="bx bx-edit-alt" style="color:#059669;font-size:13px;"></i>
                                <span>Blok Tanda Tangan</span>
                                <span class="spt-acc-dot spt-hidden" data-acc-dot></span>
                                <span class="spt-acc-count spt-hidden" data-acc-count></span>
                            </span>
                            <i class="bx bx-chevron-down spt-acc-icon" style="font-size:14px;"></i>
                        </button>
                        <div id="acc-ttd" class="layout-accordion-content spt-acc-body is-collapsed">
                            <div style="display:flex;flex-direction:column;gap:10px;">
                                <?= spt_layout_render_control('ttd_offset', 'Jarak Atas Blok TTD', $layout_initial, 'px') ?>
                                <?= spt_layout_render_control('ttd_x', 'Posisi Horizontal (X)', $layout_initial, 'px') ?>
                                <?= spt_layout_render_control('ttd_date_to_jabatan', 'Jarak Tanggal ke Jabatan', $layout_initial, 'px') ?>
                                <?= spt_layout_render_control('ttd_space_height', 'Ruang TTD & Stempel', $layout_initial, 'px') ?>
                                <?= spt_layout_render_control('ttd_name_nip_spacing', 'Jarak ke Nama & NIP', $layout_initial, 'px') ?>
                                <label style="display:flex;align-items:flex-start;gap:8px;padding:8px;border:1px solid #e2e8f0;border-radius:8px;background:#f8fafc;cursor:pointer;">
                                    <input type="checkbox" id="layoutLockTtd" style="margin-top:2px;accent-color:#059669;" <?= !empty($layout_initial['lock_ttd']) ? 'checked' : '' ?>>
                                    <span style="font-size:10.5px;color:#334155;line-height:1.35;">
                                        <span style="font-weight:600;">Kunci Blok TTD</span>
                                        <span style="display:block;color:#94a3b8;">Cegah blok tanda tangan terpotong ke halaman berikutnya.</span>
                                    </span>
                                </label>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <div class="spt-inspector-foot">
                <div style="display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:6px;">
                    <span id="layoutStatus" style="font-size:10px;color:#64748b;">Pengaturan default template</span>
                    <span style="font-size:9.5px;color:#94a3b8;"><span style="display:inline-block;padding:1px 4px;border:1px solid #e2e8f0;border-radius:4px;background:#f8fafc;">Ctrl</span> + <span style="display:inline-block;padding:1px 4px;border:1px solid #e2e8f0;border-radius:4px;background:#f8fafc;">Enter</span></span>
                </div>
                <button type="button" id="layoutApplyBtn" class="spt-apply-btn">
                    <span data-apply-spinner class="spt-btn-spinner spt-hidden"></span>
                    <i class="bx bx-check-double" style="font-size:14px;"></i>
                    <span data-apply-label>Terapkan Perubahan</span>
                </button>
                <button type="button" id="layoutSaveDefaultBtn" class="spt-save-btn">
                    <i class="bx bx-save" style="font-size:14px;"></i>
                    <span>Simpan sebagai Default Template</span>
                </button>
                <p class="spt-foot-note">Default template berlaku untuk semua surat yang memakai template ini.</p>
            </div>
        </aside>

        <div class="spt-resizer" id="layoutResizer" title="Geser untuk mengatur lebar panel inspector"></div>

        <section class="spt-preview">
            <div class="spt-toolbar">
                <div class="spt-toolbar-left">
                    <button type="button" id="sidebarToggle" class="spt-icon-btn" title="Tampilkan/sembunyikan menu navigasi" aria-label="Toggle menu navigasi">
                        <i class="bx bx-menu" style="font-size:16px;"></i>
                    </button>
                    <a href="surat-tugas.php?edit=<?= $sid ?>" class="spt-back-link">
                        <i class="bx bx-arrow-back"></i>
                        <span>Kembali ke Edit Surat</span>
                    </a>
                    <span class="spt-doc-id" title="Surat Tugas #<?= htmlspecialchars($surat['nomor_surat']) ?>">Surat Tugas #<?= htmlspecialchars($surat['nomor_surat']) ?></span>
                </div>
                <div class="spt-toolbar-center">
                    <div class="spt-mode" id="previewMode" role="group" aria-label="Mode pratinjau">
                        <button type="button" class="spt-mode-btn is-active" data-mode="html" title="Klik blok surat lalu geser untuk menata letak">
                            <i class="bx bx-move" style="font-size:13px;"></i>
                            <span>Interaktif</span>
                        </button>
                        <button type="button" class="spt-mode-btn" data-mode="pdf" title="Pratinjau PDF, sama persis dengan hasil cetak">
                            <i class="bx bxs-file-pdf" style="font-size:13px;"></i>
                            <span>PDF</span>
                        </button>
                    </div>
                    <span id="pageBadge" class="spt-badge spt-badge-neutral">Memeriksa halaman&hellip;</span>
                    <span id="saveBadge" class="spt-badge spt-badge-emerald">Tersimpan</span>
                </div>
                <div class="spt-toolbar-right">
                    <a id="downloadDocxBtn" href="generate-surat.php?id=<?= $sid ?>&format=docx" class="spt-btn spt-btn-outline" title="Unduh Word (.docx)">
                        <i class="bx bx-file-blank" style="font-size:14px;"></i>
                        <span>Word</span>
                    </a>
                    <a id="downloadPdfBtn" href="<?= htmlspecialchars($downloadPdfUrl) ?>" class="spt-btn spt-btn-emerald">
                        <i class="bx bxs-file-pdf" style="font-size:14px;"></i>
                        <span>Unduh PDF</span>
                    </a>
                    <button type="button" id="printPdfBtn" class="spt-btn spt-btn-outline">
                        <i class="bx bx-printer" style="font-size:14px;"></i>
                        <span>Cetak</span>
                    </button>
                </div>
            </div>
            <div class="spt-canvas" id="pdfCanvas">
                <iframe id="pdfPreviewFrame" class="spt-frame" title="Preview surat (PDF)"></iframe>
                <div id="pdfPreviewLoading" class="spt-overlay spt-hidden">
                    <div style="display:flex;flex-direction:column;align-items:center;gap:8px;text-align:center;">
                        <div class="spt-spinner"></div>
                        <p style="font-size:12px;font-weight:600;color:#334155;">Memperbarui pratinjau&hellip;</p>
                        <p style="font-size:10.5px;color:#64748b;">Menyusun ulang tata letak surat</p>
                    </div>
                </div>
            </div>
        </section>
    </div>
</div>

<script>
(function () {
    const SURAT_ID = <?= json_encode((string)$sid) ?>;
    const TEMPLATE_ID = <?= json_encode((int)$template_id_terpilih) ?>;
    const DEFAULTS = <?= $layoutJsonDefaults ?>;
    const INITIAL = <?= $layoutJsonInitial ?>;
    const HAS_SAVED_DEFAULT = <?= $layout_has_saved_default ? 'true' : 'false' ?>;
    const INITIAL_PREVIEW_URL = <?= json_encode($previewPdfUrl) ?>;
    const INITIAL_DOWNLOAD_URL = <?= json_encode($downloadPdfUrl) ?>;
    const STORAGE_KEY = 'spt_layout_' + SURAT_ID;
    const WIDTH_KEY = 'spt_inspector_width';
    const SIDEBAR_KEY = 'spt_sidebar_collapsed';
    const MODE_KEY = 'spt_preview_mode';

    const STEP_CONFIG = {
        margin_top: { min: 5, max: 40, step: 0.5, decimals: 1, unit: 'mm' },
        margin_bottom: { min: 5, max: 40, step: 0.5, decimals: 1, unit: 'mm' },
        margin_left: { min: 10, max: 40, step: 0.5, decimals: 1, unit: 'mm' },
        margin_right: { min: 10, max: 40, step: 0.5, decimals: 1, unit: 'mm' },
        font_size: { min: 9, max: 13, step: 0.5, decimals: 1, unit: 'pt' },
        line_height: { min: 1.0, max: 2.0, step: 0.05, decimals: 2, unit: '×' },
        kop_line_spacing: { min: 0, max: 20, step: 1, decimals: 0, unit: 'px' },
        kop_bottom_spacing: { min: 0, max: 50, step: 1, decimals: 0, unit: 'px' },
        judul_top_spacing: { min: -15, max: 40, step: 1, decimals: 0, unit: 'px' },
        nomor_surat_spacing: { min: 2, max: 30, step: 1, decimals: 0, unit: 'px' },
        judul_bottom_spacing: { min: 0, max: 40, step: 1, decimals: 0, unit: 'px' },
        dasar_top_spacing: { min: 0, max: 30, step: 1, decimals: 0, unit: 'px' },
        dasar_item_spacing: { min: 0, max: 20, step: 1, decimals: 0, unit: 'px' },
        dasar_bottom_spacing: { min: 0, max: 30, step: 1, decimals: 0, unit: 'px' },
        memerintahkan_spacing: { min: 0, max: 30, step: 1, decimals: 0, unit: 'px' },
        pegawai_top_spacing: { min: 0, max: 30, step: 1, decimals: 0, unit: 'px' },
        pegawai_gap_spacing: { min: 0, max: 35, step: 1, decimals: 0, unit: 'px' },
        pegawai_line_spacing: { min: 0, max: 15, step: 1, decimals: 0, unit: 'px' },
        untuk_top_spacing: { min: 0, max: 30, step: 1, decimals: 0, unit: 'px' },
        untuk_item_spacing: { min: 0, max: 20, step: 1, decimals: 0, unit: 'px' },
        untuk_bottom_spacing: { min: 0, max: 30, step: 1, decimals: 0, unit: 'px' },
        ttd_offset: { min: -20, max: 80, step: 1, decimals: 0, unit: 'px' },
        ttd_x: { min: -50, max: 50, step: 1, decimals: 0, unit: 'px' },
        ttd_date_to_jabatan: { min: 0, max: 25, step: 1, decimals: 0, unit: 'px' },
        ttd_space_height: { min: 20, max: 120, step: 2, decimals: 0, unit: 'px' },
        ttd_name_nip_spacing: { min: 0, max: 25, step: 1, decimals: 0, unit: 'px' }
    };
    const KEYS = Object.keys(STEP_CONFIG);
    const AUTOFIT_VALUES = {
        margin_top: 8, margin_bottom: 8, margin_left: 15, margin_right: 15,
        font_size: 9.5, line_height: 1.0, kop_line_spacing: 0, kop_bottom_spacing: 4,
        judul_top_spacing: -5, nomor_surat_spacing: 2, judul_bottom_spacing: 4,
        dasar_top_spacing: 3, dasar_item_spacing: 1, dasar_bottom_spacing: 4,
        memerintahkan_spacing: 4, pegawai_top_spacing: 3, pegawai_gap_spacing: 4,
        pegawai_line_spacing: 1, untuk_top_spacing: 3, untuk_item_spacing: 1,
        untuk_bottom_spacing: 5, ttd_offset: 4, ttd_x: 0, ttd_date_to_jabatan: 2,
        ttd_space_height: 40, ttd_name_nip_spacing: 2, lock_ttd: 1
    };

    const frame = document.getElementById('pdfPreviewFrame');
    const loading = document.getElementById('pdfPreviewLoading');
    const downloadBtn = document.getElementById('downloadPdfBtn');
    const docxBtn = document.getElementById('downloadDocxBtn');
    const printBtn = document.getElementById('printPdfBtn');
    const modeGroup = document.getElementById('previewMode');
    const status = document.getElementById('layoutStatus');
    const lockCheckbox = document.getElementById('layoutLockTtd');
    const applyBtn = document.getElementById('layoutApplyBtn');
    const pageBadge = document.getElementById('pageBadge');
    const saveBadge = document.getElementById('saveBadge');
    const autoFitBtn = document.getElementById('layoutAutoFit');
    const resetBtn = document.getElementById('layoutReset');
    const expandAllBtn = document.getElementById('layoutExpandAll');
    const sidebarToggle = document.getElementById('sidebarToggle');
    const mainSidebar = document.getElementById('mainSidebar');
    const inspector = document.getElementById('sptInspector');
    const resizer = document.getElementById('layoutResizer');
    const accordions = Array.from(document.querySelectorAll('.spt-acc'));
    const ranges = {};
    const numbers = {};
    document.querySelectorAll('[data-layout-key]').forEach(function (el) {
        ranges[el.dataset.layoutKey] = el;
    });
    document.querySelectorAll('[data-number-key]').forEach(function (el) {
        numbers[el.dataset.numberKey] = el;
    });

    let state = Object.assign({}, INITIAL, { lock_ttd: INITIAL.lock_ttd ? 1 : 0 });
    let templateDefaults = Object.assign({}, DEFAULTS, { lock_ttd: DEFAULTS.lock_ttd ? 1 : 0 });
    let appliedState = Object.assign({}, state);
    let hasSavedDefault = HAS_SAVED_DEFAULT;
    let active = HAS_SAVED_DEFAULT;
    let pageCount = null;
    let currentBlobUrl = null;
    let autoFitTimer = null;
    let pendingCountUrl = null;
    let restoredFromStorage = false;
    let previewMode = 'html';
    let syncTimer = null;

    function clamp(key, value) {
        const cfg = STEP_CONFIG[key];
        if (!cfg) return Number(value);
        let next = Number(value);
        if (!isFinite(next)) return Number(DEFAULTS[key] || 0);
        next = Math.min(cfg.max, Math.max(cfg.min, next));
        return Number(next.toFixed(cfg.decimals));
    }

    function fmt(key, value) {
        const cfg = STEP_CONFIG[key];
        if (!cfg) return String(value);
        return Number(value).toFixed(cfg.decimals);
    }

    function matchesState(candidate) {
        return KEYS.every(function (key) {
            return Math.abs(Number(state[key]) - Number(candidate[key])) < 0.0001;
        }) && (!!state.lock_ttd === !!candidate.lock_ttd);
    }

    function isDefaultState() {
        return matchesState(templateDefaults);
    }

    function isAutoFitState() {
        return matchesState(AUTOFIT_VALUES);
    }

    function isDirty() {
        return KEYS.some(function (key) {
            return Math.abs(Number(state[key]) - Number(appliedState[key])) > 0.0001;
        }) || (!!state.lock_ttd !== !!appliedState.lock_ttd);
    }

    function updateDirtyUi() {
        const dirty = isDirty();
        if (applyBtn) {
            applyBtn.classList.toggle('is-dirty', dirty);
            const label = applyBtn.querySelector('[data-apply-label]');
            if (label) label.textContent = dirty ? 'Terapkan Perubahan • Ada Perubahan' : 'Terapkan Perubahan';
        }
        if (saveBadge) {
            saveBadge.textContent = dirty ? 'Belum Diterapkan' : 'Tersimpan';
            saveBadge.className = 'spt-badge ' + (dirty ? 'spt-badge-amber' : 'spt-badge-emerald');
        }
    }

    function updateAccordionBadges() {
        accordions.forEach(function (acc) {
            const keys = (acc.dataset.keys || '').split(/\s+/).filter(Boolean);
            let changed = 0;
            keys.forEach(function (key) {
                if (key === 'lock_ttd') {
                    if (!!state.lock_ttd !== !!templateDefaults.lock_ttd) changed++;
                    return;
                }
                if (Math.abs(Number(state[key]) - Number(templateDefaults[key])) > 0.0001) changed++;
            });
            const dot = acc.querySelector('[data-acc-dot]');
            const count = acc.querySelector('[data-acc-count]');
            if (dot) dot.classList.toggle('spt-hidden', changed === 0);
            if (count) {
                count.textContent = String(changed);
                count.classList.toggle('spt-hidden', changed === 0);
            }
        });
    }

    function updatePresetState() {
        if (!autoFitBtn) return;
        const isActive = isAutoFitState() && pageCount === 1 && !isDirty();
        autoFitBtn.classList.toggle('is-active', isActive);
    }

    function updateStatus() {
        if (!status) return;
        if (!active) {
            status.textContent = 'Pengaturan default template';
            return;
        }
        if (isDirty()) {
            status.textContent = 'Ada perubahan — belum diterapkan';
            return;
        }
        if (isDefaultState()) {
            status.textContent = hasSavedDefault ? 'Default template tersimpan' : 'Pengaturan default template';
            return;
        }
        status.textContent = 'Kustom aktif — belum disimpan sebagai default';
    }

    function updatePageBadge(count) {
        pageCount = (typeof count === 'number' && count > 0) ? count : null;
        if (!pageBadge) return;
        if (pageCount === null) {
            pageBadge.textContent = 'Jumlah halaman tidak diketahui';
            pageBadge.className = 'spt-badge spt-badge-neutral';
        } else if (pageCount === 1) {
            pageBadge.textContent = '1 Halaman — Ideal';
            pageBadge.className = 'spt-badge spt-badge-emerald';
        } else {
            pageBadge.textContent = pageCount + ' Halaman — Belum Muat 1 Halaman';
            pageBadge.className = 'spt-badge spt-badge-amber';
        }
        updatePresetState();
    }

    function syncUi() {
        KEYS.forEach(function (key) {
            if (ranges[key]) ranges[key].value = state[key];
            if (numbers[key] && document.activeElement !== numbers[key]) {
                numbers[key].value = fmt(key, state[key]);
            }
        });
        if (lockCheckbox) lockCheckbox.checked = !!state.lock_ttd;
        updateAccordionBadges();
        updateDirtyUi();
        updatePresetState();
        updateStatus();
    }

    function markDirty() {
        active = true;
        updateDirtyUi();
        updateAccordionBadges();
        updatePresetState();
        updateStatus();
        schedulePreviewSync();
    }

    function setAccordion(acc, open) {
        const body = acc.querySelector('.layout-accordion-content');
        const head = acc.querySelector('.layout-accordion-btn');
        if (!body) return;
        body.classList.toggle('is-collapsed', !open);
        if (head) head.setAttribute('aria-expanded', open ? 'true' : 'false');
    }

    function anyAccordionOpen() {
        return accordions.some(function (acc) {
            const body = acc.querySelector('.layout-accordion-content');
            return body && !body.classList.contains('is-collapsed');
        });
    }

    function updateExpandAllLabel() {
        if (expandAllBtn) expandAllBtn.textContent = anyAccordionOpen() ? 'Tutup Semua' : 'Buka Semua';
    }

    function toggleAccordion(acc) {
        const body = acc.querySelector('.layout-accordion-content');
        const willOpen = body ? body.classList.contains('is-collapsed') : false;
        accordions.forEach(function (other) {
            if (other !== acc) setAccordion(other, false);
        });
        setAccordion(acc, willOpen);
        updateExpandAllLabel();
    }

    function showLoading(show) {
        if (loading) loading.classList.toggle('spt-hidden', !show);
    }

    function setApplyLoading(isLoading) {
        if (!applyBtn) return;
        applyBtn.disabled = !!isLoading;
        const spinner = applyBtn.querySelector('[data-apply-spinner]');
        if (spinner) spinner.classList.toggle('spt-hidden', !isLoading);
    }

    function setAutoFitLoading(isLoading) {
        if (!autoFitBtn) return;
        autoFitBtn.classList.toggle('is-loading', !!isLoading);
        if (autoFitTimer) {
            clearTimeout(autoFitTimer);
            autoFitTimer = null;
        }
        if (isLoading) {
            autoFitTimer = setTimeout(function () {
                autoFitBtn.classList.remove('is-loading');
            }, 15000);
        }
    }

    function buildParams(withInline) {
        const params = new URLSearchParams();
        params.set('id', SURAT_ID);
        params.set('format', 'pdf');
        if (withInline) params.set('inline', '1');
        if (active) {
            params.set('layout', '1');
            KEYS.forEach(function (key) {
                params.set(key, state[key]);
            });
            params.set('lock_ttd', state.lock_ttd ? '1' : '0');
        }
        return params;
    }

    function currentSettings() {
        const values = {};
        KEYS.forEach(function (key) { values[key] = state[key]; });
        values.lock_ttd = state.lock_ttd ? 1 : 0;
        return values;
    }

    function buildHtmlParams() {
        const params = buildParams(false);
        params.set('format', 'preview');
        return params;
    }

    function htmlPreviewUrl() {
        return 'generate-surat.php?' + buildHtmlParams().toString();
    }

    function updateDownloadLinks() {
        if (downloadBtn) {
            downloadBtn.href = 'generate-surat.php?' + buildParams(false).toString();
        }
        if (docxBtn) {
            const params = buildParams(false);
            params.set('format', 'docx');
            docxBtn.href = 'generate-surat.php?' + params.toString();
        }
    }

    function loadHtmlPreview() {
        if (!frame) return false;
        if (currentBlobUrl) {
            try { URL.revokeObjectURL(currentBlobUrl); } catch (e) {}
            currentBlobUrl = null;
        }
        showLoading(true);
        frame.src = htmlPreviewUrl();
        return true;
    }

    // Kirim nilai terbaru ke pratinjau interaktif tanpa memuat ulang halaman.
    function schedulePreviewSync() {
        if (previewMode !== 'html' || !frame || !frame.contentWindow) return;
        if (syncTimer) clearTimeout(syncTimer);
        syncTimer = setTimeout(function () {
            syncTimer = null;
            try {
                frame.contentWindow.postMessage({ type: 'set-settings', settings: currentSettings() }, '*');
            } catch (e) {
                // iframe belum siap; pengaturan dikirim lagi saat pesan ready.
            }
        }, 60);
    }

    function persistLocal() {
        try {
            localStorage.setItem(STORAGE_KEY, JSON.stringify({
                active: true,
                values: Object.assign({}, state, { lock_ttd: state.lock_ttd ? 1 : 0 })
            }));
        } catch (e) {
            // localStorage bisa diblokir browser; pratinjau tetap jalan.
        }
    }

    function updateHistory() {
        const historyParams = buildParams(false);
        historyParams.delete('format');
        const nextUrl = new URL(window.location.href);
        nextUrl.search = historyParams.toString();
        window.history.replaceState(null, '', nextUrl.toString());
    }

    function schedulePageCountRefresh(url) {
        pendingCountUrl = url;
        setTimeout(async function () {
            if (!pendingCountUrl) return;
            const target = pendingCountUrl;
            pendingCountUrl = null;
            const sep = target.indexOf('?') === -1 ? '?' : '&';
            try {
                const response = await fetch(target + sep + '_ts=' + Date.now(), { cache: 'no-store' });
                const contentType = (response.headers.get('content-type') || '').toLowerCase();
                if (response.ok && contentType.indexOf('application/pdf') !== -1) {
                    const count = parseInt(response.headers.get('X-SPT-Page-Count') || '0', 10);
                    updatePageBadge(count > 0 ? count : null);
                }
            } catch (e) {
                // abaikan; badge tetap menampilkan status terakhir
            }
        }, 3500);
    }

    async function loadPreview(url, attempt) {
        attempt = attempt || 0;
        if (!frame) return false;
        showLoading(true);
        setApplyLoading(true);
        const sep = url.indexOf('?') === -1 ? '?' : '&';
        try {
            const response = await fetch(url + sep + '_ts=' + Date.now(), {
                cache: 'no-store',
                credentials: 'same-origin'
            });
            const contentType = (response.headers.get('content-type') || '').toLowerCase();
            if (response.ok && contentType.indexOf('application/pdf') !== -1) {
                const count = parseInt(response.headers.get('X-SPT-Page-Count') || '0', 10);
                const blob = await response.blob();
                if (blob && blob.size > 0) {
                    const nextBlobUrl = URL.createObjectURL(blob);
                    if (currentBlobUrl) {
                        try { URL.revokeObjectURL(currentBlobUrl); } catch (e) {}
                    }
                    currentBlobUrl = nextBlobUrl;
                    frame.src = nextBlobUrl;
                    updatePageBadge(count > 0 ? count : null);
                    appliedState = Object.assign({}, state, { lock_ttd: state.lock_ttd ? 1 : 0 });
                    updateDirtyUi();
                    updateAccordionBadges();
                    updatePresetState();
                    updateStatus();
                    showLoading(false);
                    setApplyLoading(false);
                    setAutoFitLoading(false);
                    return true;
                }
            }

            // Bukan PDF (mis. halaman retry Word COM). Biarkan iframe menanganinya.
            frame.src = url;
            showLoading(false);
            setApplyLoading(false);
            setAutoFitLoading(false);
            schedulePageCountRefresh(url);
            return false;
        } catch (error) {
            frame.src = url;
            showLoading(false);
            setApplyLoading(false);
            setAutoFitLoading(false);
            return false;
        }
    }

    async function applyLayout() {
        if (!frame) return false;
        active = true;
        persistLocal();
        updateHistory();
        updateDownloadLinks();
        if (previewMode === 'html') {
            // Pratinjau interaktif memperbarui dirinya lewat postMessage.
            schedulePreviewSync();
            markApplied();
            return true;
        }
        const previewUrl = 'generate-surat.php?' + buildParams(true).toString();
        return loadPreview(previewUrl);
    }

    function markApplied() {
        appliedState = Object.assign({}, state, { lock_ttd: state.lock_ttd ? 1 : 0 });
        updateDirtyUi();
        updateAccordionBadges();
        updatePresetState();
        updateStatus();
    }

    // ------------------------------------------------------------------
    // Jembatan pesan dengan pratinjau interaktif.
    // ------------------------------------------------------------------
    const BLOCK_ACCORDION = {
        kop: 'acc-kop',
        title: 'acc-judul',
        nomor: 'acc-judul',
        dasar: 'acc-dasar',
        memerintahkan: 'acc-dasar',
        pegawai: 'acc-pegawai',
        untuk: 'acc-untuk',
        ttd: 'acc-ttd'
    };

    function openAccordionById(id) {
        const body = document.getElementById(id);
        if (!body) return;
        const acc = body.closest('.spt-acc');
        if (!acc) return;
        if (body.classList.contains('is-collapsed')) {
            accordions.forEach(function (other) {
                if (other !== acc) setAccordion(other, false);
            });
            setAccordion(acc, true);
            updateExpandAllLabel();
        }
        try {
            acc.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
        } catch (e) {
            acc.scrollIntoView();
        }
    }

    function syncFrameSettings() {
        if (!frame || !frame.contentWindow) return;
        try {
            frame.contentWindow.postMessage({ type: 'set-settings', settings: currentSettings() }, '*');
        } catch (e) {
            // iframe belum siap
        }
    }

    window.addEventListener('message', function (event) {
        if (!frame || event.source !== frame.contentWindow) return;
        const data = event.data || {};
        if (data.type === 'ready') {
            showLoading(false);
            setApplyLoading(false);
            setAutoFitLoading(false);
            syncFrameSettings();
            if (typeof data.count === 'number') updatePageBadge(data.count);
        } else if (data.type === 'page-estimate') {
            if (typeof data.count === 'number') updatePageBadge(data.count);
        } else if (data.type === 'settings-change' && data.settings) {
            KEYS.forEach(function (key) {
                if (data.settings[key] !== undefined) state[key] = clamp(key, data.settings[key]);
            });
            if (data.settings.lock_ttd !== undefined) {
                state.lock_ttd = data.settings.lock_ttd ? 1 : 0;
            }
            syncUi();
            markDirty();
        } else if (data.type === 'block-select' && data.block) {
            const accId = BLOCK_ACCORDION[data.block];
            if (accId) openAccordionById(accId);
        }
    });

    function setPreviewMode(mode, persist) {
        previewMode = mode === 'pdf' ? 'pdf' : 'html';
        renderModeButtons();
        if (persist) {
            try { localStorage.setItem(MODE_KEY, previewMode); } catch (e) {}
        }
        if (previewMode === 'pdf') {
            loadPreview('generate-surat.php?' + buildParams(true).toString());
        } else {
            loadHtmlPreview();
        }
    }

    function renderModeButtons() {
        if (!modeGroup) return;
        modeGroup.querySelectorAll('.spt-mode-btn').forEach(function (btn) {
            btn.classList.toggle('is-active', btn.dataset.mode === previewMode);
        });
    }

    function resetKey(key) {
        if (!key) return;
        if (key === 'lock_ttd') {
            state.lock_ttd = templateDefaults.lock_ttd ? 1 : 0;
            if (lockCheckbox) lockCheckbox.checked = !!state.lock_ttd;
        } else if (STEP_CONFIG[key]) {
            state[key] = clamp(key, templateDefaults[key]);
            if (ranges[key]) ranges[key].value = state[key];
            if (numbers[key]) numbers[key].value = fmt(key, state[key]);
        }
        markDirty();
    }

    // Restore state: prioritas URL, lalu localStorage.
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get('layout') === '1') {
        active = true;
        KEYS.forEach(function (key) {
            if (urlParams.has(key)) state[key] = clamp(key, urlParams.get(key));
        });
        state.lock_ttd = urlParams.get('lock_ttd') === '1' ? 1 : 0;
        appliedState = Object.assign({}, state);
    } else {
        try {
            const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null');
            if (saved && saved.active) {
                KEYS.forEach(function (key) {
                    if (saved.values && saved.values[key] !== undefined) {
                        state[key] = clamp(key, saved.values[key]);
                    }
                });
                state.lock_ttd = saved.values && saved.values.lock_ttd ? 1 : 0;
                active = true;
                restoredFromStorage = true;
            }
        } catch (e) {
            // localStorage rusak; abaikan.
        }
    }

    // Accordion: single-expand + buka/tutup semua.
    accordions.forEach(function (acc) {
        const head = acc.querySelector('.layout-accordion-btn');
        if (head) head.addEventListener('click', function () { toggleAccordion(acc); });
    });
    if (expandAllBtn) {
        expandAllBtn.addEventListener('click', function () {
            const open = !anyAccordionOpen();
            accordions.forEach(function (acc) { setAccordion(acc, open); });
            updateExpandAllLabel();
        });
    }
    updateExpandAllLabel();

    // Slider + number input (two-way binding).
    KEYS.forEach(function (key) {
        const range = ranges[key];
        if (range) {
            range.addEventListener('input', function () {
                const value = clamp(key, range.value);
                state[key] = value;
                if (numbers[key]) numbers[key].value = fmt(key, value);
                markDirty();
            });
            range.addEventListener('dblclick', function () { resetKey(key); });
        }

        const number = numbers[key];
        if (number) {
            number.addEventListener('input', function () {
                const raw = number.value.trim();
                if (raw === '' || raw === '-' || raw === '.' || raw === '-.') return;
                const parsed = Number(raw);
                if (!isFinite(parsed)) return;
                state[key] = clamp(key, parsed);
                if (range) range.value = state[key];
                markDirty();
            });
            number.addEventListener('blur', function () {
                const value = clamp(key, number.value);
                state[key] = value;
                number.value = fmt(key, value);
                if (range) range.value = value;
                markDirty();
            });
            number.addEventListener('keydown', function (event) {
                if (event.key === 'Enter') {
                    event.preventDefault();
                    number.blur();
                }
            });
        }
    });

    // Stepper - / +
    document.querySelectorAll('.layout-step').forEach(function (button) {
        button.addEventListener('click', function () {
            const key = button.dataset.target;
            const cfg = STEP_CONFIG[key];
            if (!cfg) return;
            const direction = Number(button.dataset.dir || 0);
            state[key] = clamp(key, Number(state[key]) + direction * cfg.step);
            if (ranges[key]) ranges[key].value = state[key];
            if (numbers[key]) numbers[key].value = fmt(key, state[key]);
            markDirty();
        });
    });

    // Reset per-item
    document.querySelectorAll('.layout-reset').forEach(function (button) {
        button.addEventListener('click', function () { resetKey(button.dataset.target); });
    });

    if (lockCheckbox) {
        lockCheckbox.addEventListener('change', function () {
            state.lock_ttd = lockCheckbox.checked ? 1 : 0;
            markDirty();
        });
    }

    if (applyBtn) {
        applyBtn.addEventListener('click', function () {
            if (!applyBtn.disabled) applyLayout();
        });
    }

    if (modeGroup) {
        modeGroup.addEventListener('click', function (event) {
            const btn = event.target.closest('.spt-mode-btn');
            if (!btn || btn.dataset.mode === previewMode) return;
            setPreviewMode(btn.dataset.mode, true);
        });
    }

    // Keyboard shortcut Ctrl/Cmd + Enter
    document.addEventListener('keydown', function (event) {
        if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
            event.preventDefault();
            if (applyBtn && !applyBtn.disabled) applyLayout();
        }
    });

    if (downloadBtn) {
        downloadBtn.addEventListener('click', async function (event) {
            if (!isDirty()) return;
            event.preventDefault();
            await applyLayout();
            window.location.href = downloadBtn.href;
        });
    }

    if (docxBtn) {
        docxBtn.addEventListener('click', async function (event) {
            if (!isDirty()) return;
            event.preventDefault();
            await applyLayout();
            window.location.href = docxBtn.href;
        });
    }

    if (printBtn) {
        printBtn.addEventListener('click', async function () {
            const win = window.open('', '_blank');
            if (isDirty()) await applyLayout();
            const targetUrl = currentBlobUrl || ('generate-surat.php?' + buildParams(true).toString());
            if (win) {
                win.location.href = targetUrl;
                setTimeout(function () {
                    try { win.print(); } catch (e) {}
                }, 1500);
            } else {
                window.open(targetUrl, '_blank');
            }
        });
    }

    const autoFit = autoFitBtn;
    if (autoFit) {
        autoFit.addEventListener('click', function () {
            KEYS.forEach(function (key) { state[key] = clamp(key, AUTOFIT_VALUES[key]); });
            state.lock_ttd = AUTOFIT_VALUES.lock_ttd;
            syncUi();
            markDirty();
            setAutoFitLoading(true);
            applyLayout();
        });
    }

    if (resetBtn) {
        resetBtn.addEventListener('click', function () {
            state = Object.assign({}, templateDefaults, { lock_ttd: templateDefaults.lock_ttd ? 1 : 0 });
            active = hasSavedDefault;
            try { localStorage.removeItem(STORAGE_KEY); } catch (e) {}
            syncUi();
            appliedState = Object.assign({}, state);
            updateDirtyUi();
            if (active) {
                applyLayout();
            } else {
                const nextUrl = new URL(window.location.href);
                nextUrl.search = 'id=' + encodeURIComponent(SURAT_ID);
                window.history.replaceState(null, '', nextUrl.toString());
                updateDownloadLinks();
                if (previewMode === 'html') {
                    loadHtmlPreview();
                } else {
                    loadPreview('generate-surat.php?id=' + encodeURIComponent(SURAT_ID) + '&format=pdf&inline=1');
                }
            }
        });
    }

    const saveDefaultBtn = document.getElementById('layoutSaveDefaultBtn');
    if (saveDefaultBtn) {
        saveDefaultBtn.addEventListener('click', async function () {
            if (!TEMPLATE_ID) {
                if (window.Swal) {
                    Swal.fire({ icon: 'error', title: 'Template tidak diketahui', text: 'Pengaturan tidak dapat disimpan karena template tidak terhubung ke database.' });
                }
                return;
            }

            const originalHtml = saveDefaultBtn.innerHTML;
            saveDefaultBtn.disabled = true;
            saveDefaultBtn.innerHTML = '<span class="spt-btn-spinner" style="border-color:#a7f3d0;border-top-color:#047857;"></span><span>Menyimpan…</span>';

            try {
                const payload = Object.assign({}, state, { lock_ttd: state.lock_ttd ? 1 : 0 });
                const response = await fetch('save-template-layout.php', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ template_id: TEMPLATE_ID, settings: payload })
                });
                const text = await response.text();
                let result;
                try {
                    result = JSON.parse(text);
                } catch (parseError) {
                    throw new Error('Respons server tidak valid.');
                }
                if (!response.ok || !result.success) {
                    throw new Error(result.message || 'Gagal menyimpan pengaturan.');
                }

                templateDefaults = Object.assign({}, state, { lock_ttd: state.lock_ttd ? 1 : 0 });
                hasSavedDefault = true;
                active = true;
                appliedState = Object.assign({}, state);
                syncUi();
                applyLayout();

                if (window.Swal) {
                    Swal.fire({
                        icon: 'success',
                        title: 'Tersimpan',
                        text: 'Pengaturan ini sekarang menjadi default untuk template tersebut.',
                        timer: 2200,
                        showConfirmButton: false
                    });
                }
            } catch (error) {
                if (window.Swal) {
                    Swal.fire({ icon: 'error', title: 'Gagal menyimpan', text: error.message || 'Terjadi kesalahan.' });
                } else {
                    alert(error.message || 'Gagal menyimpan pengaturan.');
                }
            } finally {
                saveDefaultBtn.disabled = false;
                saveDefaultBtn.innerHTML = originalHtml;
            }
        });
    }

    // Sidebar navigasi: auto-collapse saat halaman dibuka, bisa di-toggle.
    function readSidebarCollapsed() {
        try { return localStorage.getItem(SIDEBAR_KEY) !== '0'; } catch (e) { return true; }
    }
    function writeSidebarCollapsed(collapsed) {
        try { localStorage.setItem(SIDEBAR_KEY, collapsed ? '1' : '0'); } catch (e) {}
    }
    function applySidebarState() {
        if (!mainSidebar) return;
        if (window.innerWidth < 768) {
            mainSidebar.classList.remove('spt-sidebar-collapsed');
            return;
        }
        mainSidebar.classList.toggle('spt-sidebar-collapsed', readSidebarCollapsed());
    }
    applySidebarState();
    if (sidebarToggle) {
        sidebarToggle.addEventListener('click', function () {
            if (!mainSidebar) return;
            if (window.innerWidth < 768) {
                mainSidebar.classList.toggle('hidden');
                mainSidebar.classList.toggle('fixed');
                mainSidebar.classList.toggle('inset-0');
                mainSidebar.classList.toggle('w-full');
                mainSidebar.classList.toggle('z-50');
                return;
            }
            writeSidebarCollapsed(!readSidebarCollapsed());
            applySidebarState();
        });
    }
    window.addEventListener('resize', applySidebarState);

    // Resizer panel inspector.
    (function initResizer() {
        if (!resizer || !inspector) return;
        let dragging = false;
        function setWidth(width) {
            const clamped = Math.min(520, Math.max(320, width));
            inspector.style.width = clamped + 'px';
            inspector.style.flexBasis = clamped + 'px';
            try { localStorage.setItem(WIDTH_KEY, String(clamped)); } catch (e) {}
        }
        try {
            const savedWidth = parseInt(localStorage.getItem(WIDTH_KEY) || '', 10);
            if (savedWidth) setWidth(savedWidth);
        } catch (e) {}
        resizer.addEventListener('pointerdown', function (event) {
            if (window.innerWidth < 1024) return;
            dragging = true;
            resizer.classList.add('is-dragging');
            try { resizer.setPointerCapture(event.pointerId); } catch (e) {}
            document.body.style.cursor = 'col-resize';
            document.body.style.userSelect = 'none';
        });
        resizer.addEventListener('pointermove', function (event) {
            if (!dragging) return;
            const rect = inspector.getBoundingClientRect();
            setWidth(event.clientX - rect.left);
        });
        function stopDrag() {
            dragging = false;
            resizer.classList.remove('is-dragging');
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
        }
        resizer.addEventListener('pointerup', stopDrag);
        resizer.addEventListener('pointercancel', stopDrag);
    })();

    if (frame) {
        frame.addEventListener('load', function () {
            if (frame.src && frame.src !== 'about:blank') {
                showLoading(false);
                setApplyLoading(false);
            }
        });
    }

    syncUi();
    updateDownloadLinks();

    try {
        if (localStorage.getItem(MODE_KEY) === 'pdf') previewMode = 'pdf';
    } catch (e) {}
    renderModeButtons();

    if (restoredFromStorage) {
        persistLocal();
        updateHistory();
        markApplied();
    }

    if (previewMode === 'html') {
        loadHtmlPreview();
    } else if (restoredFromStorage) {
        loadPreview('generate-surat.php?' + buildParams(true).toString());
    } else {
        loadPreview(INITIAL_PREVIEW_URL);
    }
})();
</script>
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
    if (spt_is_dir_writable($candidate)) {
        $tempDir = $candidate;
        break;
    }
}
Settings::setTempDir($tempDir);
    $tempDocx = $tempDir . DIRECTORY_SEPARATOR . 'spt_' . uniqid() . '.docx';
    $templateProcessor->saveAs($tempDocx);
    spt_fix_docx_typos($tempDocx);
    spt_normalize_docx($tempDocx);

    if ($format === 'preview') {
        // Pratinjau interaktif: HTML dari DOCX + CSS dinamis, tanpa konversi PDF.
        register_shutdown_function(function () use ($tempDocx) {
            @unlink($tempDocx);
        });
        spt_preview_output($tempDocx, $layout_initial, $tempDir);
        exit;
    }

    // Salinan DOCX yang sudah disesuaikan layout, dipakai untuk unduhan Word
    // dan untuk Word COM/LibreOffice. File asli tetap dipakai fallback Dompdf.
    $layout_adjusted_docx = null;
    $pdf_source_docx = $tempDocx;
    if (($format === 'pdf' || $format === 'docx') && $layout_active && $layout_settings !== null) {
        $layout_adjusted_docx = $tempDir . DIRECTORY_SEPARATOR . 'spt_layout_' . uniqid() . '.docx';
        if (@copy($tempDocx, $layout_adjusted_docx)) {
            if (spt_layout_apply_to_docx($layout_adjusted_docx, $layout_settings)) {
                $pdf_source_docx = $layout_adjusted_docx;
            } else {
                @unlink($layout_adjusted_docx);
                $layout_adjusted_docx = null;
            }
        } else {
            $layout_adjusted_docx = null;
        }
    }

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
                    $tempPdf = konversi_docx_ke_pdf_word_com($pdf_source_docx, 3);
                    if ($tempPdf !== null) {
                        $usedMethod = 'Word COM';
                    }
                }

                // Jika server mendukung COM tetapi Word COM belum berhasil dan com_retry < 3,
                // tampilkan halaman auto-reload agar browser mencoba kembali secara otomatis sampai Word COM merespons.
                if ($tempPdf === null && $hasComSupport && $comRetryCount < 3) {
                    @unlink($tempDocx);
                    if ($layout_adjusted_docx) {
                        @unlink($layout_adjusted_docx);
                    }
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
                    $tempPdf = konversi_docx_ke_pdf_libreoffice($pdf_source_docx);
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
                            // Fallback Dompdf memakai DOCX yang sudah disesuaikan layout + CSS dinamis
                            $phpWord = IOFactory::load($pdf_source_docx, 'Word2007');
                            $pdfWriter = IOFactory::createWriter($phpWord, 'PDF');
                            if ($layout_active && $layout_settings !== null && method_exists($pdfWriter, 'getRenderer')) {
                                $renderer = $pdfWriter->getRenderer();
                                if ($renderer && method_exists($renderer, 'setEditCallback')) {
                                    $layoutSettingsForCss = $layout_settings;
                                    $renderer->setEditCallback(function ($html) use ($layoutSettingsForCss) {
                                        return spt_layout_inject_html($html, $layoutSettingsForCss);
                                    });
                                }
                            }
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
                if ($layout_adjusted_docx) {
                    @unlink($layout_adjusted_docx);
                }
                if (ob_get_length()) {
                    ob_clean();
                }
                header('Content-Type: application/pdf');
                $sptPageCount = spt_pdf_page_count($tempPdf);
                if ($sptPageCount > 0) {
                    header('X-SPT-Page-Count: ' . $sptPageCount);
                }
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
        if (ob_get_length()) {
            ob_clean();
        }
        $docxSource = $layout_adjusted_docx ?: $tempDocx;
        header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        header('Content-Disposition: attachment; filename="Surat_Tugas_' . $surat['nomor_surat'] . '.docx"');
        header('Content-Length: ' . filesize($docxSource));
        readfile($docxSource);
        @unlink($tempDocx);
        if ($layout_adjusted_docx) @unlink($layout_adjusted_docx);
    } catch (Exception $e) {
        if (file_exists($tempDocx)) @unlink($tempDocx);
        if ($layout_adjusted_docx && file_exists($layout_adjusted_docx)) @unlink($layout_adjusted_docx);
        throw $e;
    }

} catch (Exception $e) {
    die("Error: " . $e->getMessage());
}