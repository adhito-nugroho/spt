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
        'kop_offset' => 0,
        'section_spacing' => 0,
        'ttd_offset' => 0,
        'ttd_x' => 0,
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
        'kop_offset' => [-20, 60, 1],
        'section_spacing' => [0, 30, 1],
        'ttd_offset' => [-20, 100, 1],
        'ttd_x' => [-50, 50, 1],
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

        $baseHalfPoints = spt_layout_detect_dominant_font_size($document);
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

        // Jarak kop ke judul: cari paragraf judul "SURAT TUGAS".
        $kopOffset = spt_layout_px_to_twips($settings['kop_offset']);
        if ($kopOffset !== 0) {
            foreach ($paragraphs as $paragraph) {
                $text = spt_layout_paragraph_text($paragraph);
                if ($text !== '' && stripos($text, 'SURAT TUGAS') !== false) {
                    spt_layout_shift_gap($paragraph, $kopOffset, $document);
                    break;
                }
            }
        }

        // Jarak antar bagian DASAR / KEPADA / UNTUK.
        $sectionSpacing = spt_layout_px_to_twips($settings['section_spacing']);
        if ($sectionSpacing > 0) {
            foreach ($paragraphs as $paragraph) {
                $text = spt_layout_paragraph_text($paragraph);
                if ($text !== '' && preg_match('/^(DASAR|KEPADA|UNTUK|MEMERINTAHKAN)\b/iu', $text)) {
                    spt_layout_shift_gap($paragraph, $sectionSpacing, $document);
                }
            }
        }

        // Cari awal blok tanda tangan: baris "Bojonegoro <tanggal>".
        $ttdStart = null;
        $ttdEnd = null;
        $paragraphCount = count($paragraphs);
        foreach ($paragraphs as $index => $paragraph) {
            $text = spt_layout_paragraph_text($paragraph);
            if ($text === '') {
                continue;
            }
            $isPlaceLine = stripos($text, 'Bojonegoro') !== false
                && stripos($text, 'Jalan') === false
                && stripos($text, 'DINAS') === false
                && stripos($text, 'KABUPATEN') === false;
            if ($isPlaceLine) {
                $ttdStart = $index;
                break;
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

            $ttdOffset = spt_layout_px_to_twips($settings['ttd_offset']);
            if ($ttdOffset !== 0) {
                spt_layout_shift_gap($paragraphs[$ttdStart], $ttdOffset, $document);
            }

            $ttdIndent = spt_layout_px_to_twips($settings['ttd_x']);
            if ($ttdIndent !== 0) {
                for ($i = $ttdStart; $i <= $ttdEnd; $i++) {
                    $ind = spt_layout_indentation($paragraphs[$i], $document);
                    $current = $ind->hasAttributeNS($ns, 'left')
                        ? (int)$ind->getAttributeNS($ns, 'left')
                        : 0;
                    $ind->setAttributeNS($ns, 'w:left', (string)($current + $ttdIndent));
                }
            }

            if (!empty($settings['lock_ttd'])) {
                for ($i = $ttdStart; $i <= $ttdEnd; $i++) {
                    spt_layout_add_keep_flag($paragraphs[$i], 'keepLines', $document);
                    if ($i < $ttdEnd) {
                        spt_layout_add_keep_flag($paragraphs[$i], 'keepNext', $document);
                    }
                }
            }
        }

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

function spt_layout_detect_default_line_height($documentXml)
{
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

    $defaults['line_height'] = spt_layout_detect_default_line_height($documentXml);

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
    $titleGap = max(0, 12 + (int)$settings['kop_offset']);
    $sectionGap = max(0, (int)$settings['section_spacing']);
    $ttdGap = max(0, (int)$settings['ttd_offset']);
    $ttdX = (int)$settings['ttd_x'];

    $css = '<style>' . PHP_EOL;
    $css .= '@page { margin: ' . $mt . 'mm ' . $mr . 'mm ' . $mb . 'mm ' . $ml . 'mm; }' . PHP_EOL;
    $css .= 'body { font-size: ' . $font . 'pt; line-height: ' . $line . '; }' . PHP_EOL;
    $css .= 'p, td, th, li, span { font-size: ' . $font . 'pt !important; line-height: ' . $line . ' !important; }' . PHP_EOL;
    $css .= '.spt-title { margin-top: ' . $titleGap . 'px !important; }' . PHP_EOL;
    $css .= '.spt-section { margin-top: ' . $sectionGap . 'px !important; }' . PHP_EOL;
    $css .= '.spt-ttd { margin-top: ' . $ttdGap . 'px !important; margin-left: ' . $ttdX . 'px !important; }' . PHP_EOL;
    if (!empty($settings['lock_ttd'])) {
        $css .= '.spt-ttd, table { page-break-inside: avoid !important; }' . PHP_EOL;
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
    $html = preg_replace_callback(
        '/<(p|h[1-6]|td|th|li|div)\b([^>]*)>(.*?)<\/\1>/is',
        function ($match) {
            $text = trim(preg_replace('/\s+/u', ' ', strip_tags($match[3])));
            if ($text === '') {
                return $match[0];
            }

            $classes = [];
            if (stripos($text, 'SURAT TUGAS') !== false) {
                $classes[] = 'spt-title';
            }
            if (preg_match('/^(DASAR|KEPADA|UNTUK|MEMERINTAHKAN)\b/iu', $text)) {
                $classes[] = 'spt-section';
            }
            if (
                stripos($text, 'Bojonegoro') !== false
                || (stripos($text, 'Kepala') !== false && stripos($text, 'Cabang Dinas') !== false)
                || stripos($text, 'NIP') !== false
            ) {
                $classes[] = 'spt-ttd';
            }
            if (empty($classes)) {
                return $match[0];
            }

            $attributes = $match[2];
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

            return '<' . $match[1] . $attributes . '>' . $match[3] . '</' . $match[1] . '>';
        },
        $html
    );

    $css = spt_layout_build_css($settings);
    if (stripos($html, '</head>') !== false) {
        return preg_replace('/<\/head>/i', $css . '</head>', $html, 1);
    }
    return $css . $html;
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

    // Parameter layout dari panel adjuster (null = pakai template asli).
    $layout_settings = spt_layout_request_settings();
    $layout_defaults = spt_layout_read_docx_defaults($template_path);
    $layout_active = $layout_settings !== null;
    if ($layout_active) {
        // Lengkapi parameter yang tidak ada di URL dengan nilai default template.
        $layout_settings = array_merge($layout_defaults, $layout_settings);
    }
    $layout_initial = $layout_active ? $layout_settings : $layout_defaults;

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
        <div class="space-y-6">
            <div class="flex flex-wrap items-center justify-between gap-4">
                <div>
                    <h2 class="text-xl font-semibold text-slate-800">Preview Surat Tugas</h2>
                    <p class="text-sm text-slate-500 mt-1">Atur tata letak di panel kiri, lalu unduh PDF dengan hasil yang sama persis.</p>
                </div>
                <div class="flex flex-wrap gap-3">
                    <a href="generate-surat.php?id=<?= $sid ?>&format=docx" class="inline-flex items-center gap-2 px-4 py-2 bg-green-600 hover:bg-green-700 text-white rounded-lg transition-colors">
                        <i class="bx bx-file-blank"></i><span>Download Word (.docx)</span>
                    </a>
                    <a id="downloadPdfBtn" href="<?= htmlspecialchars($downloadPdfUrl) ?>" class="inline-flex items-center gap-2 px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded-lg transition-colors">
                        <i class="bx bxs-file-pdf"></i><span>Download PDF</span>
                    </a>
                    <a href="surat-tugas.php" class="inline-flex items-center gap-2 px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg transition-colors">
                        <i class="bx bx-arrow-back"></i><span>Kembali</span>
                    </a>
                </div>
            </div>
            <?php if (!empty($template_berlaku_mulai)): ?>
            <div class="rounded-xl border border-emerald-200 bg-emerald-50 px-4 py-3 text-sm text-emerald-900 flex items-start gap-3">
                <i class='bx bx-calendar-check text-emerald-600 text-xl mt-0.5'></i>
                <div>
                    <p class="font-semibold">Memakai template dasar baru (PAPBD)</p>
                    <p>Template: <strong><?= htmlspecialchars($template_nama ?? basename($template_path)) ?></strong> — berlaku mulai <?= date('d/m/Y', strtotime($template_berlaku_mulai)) ?>, sesuai tanggal surat <?= date('d/m/Y', strtotime($surat['tanggal_surat'])) ?>.</p>
                </div>
            </div>
            <?php elseif (!empty($peringatan_template_lama)): ?>
            <div class="rounded-xl border border-amber-300 bg-amber-50 px-4 py-3 text-sm text-amber-900 flex items-start gap-3">
                <i class='bx bx-error text-amber-600 text-xl mt-0.5'></i>
                <div>
                    <p class="font-semibold">Memakai template dasar lama — perlu perhatian</p>
                    <p>Template: <strong><?= htmlspecialchars($template_nama ?? basename($template_path)) ?></strong>. Tanggal surat (<?= date('d/m/Y', strtotime($surat['tanggal_surat'])) ?>) sudah masuk masa dasar baru PAPBD, tetapi template dasar baru untuk kombinasi kategori &ldquo;<?= htmlspecialchars($kategori_label) ?>&rdquo; tipe &ldquo;<?= htmlspecialchars($tipe_label) ?>&rdquo; belum di-upload. Segera upload di halaman Template Word dan isi kolom Berlaku Mulai<?= $cutoff_global ? ' dengan ' . date('d/m/Y', strtotime($cutoff_global)) : ' sesuai tanggal berlakunya PAPBD'; ?>.</p>
                </div>
            </div>
            <?php else: ?>
            <div class="rounded-xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm text-slate-700 flex items-start gap-3">
                <i class='bx bx-history text-slate-500 text-xl mt-0.5'></i>
                <div>
                    <p class="font-semibold">Memakai template dasar lama</p>
                    <p>Template: <strong><?= htmlspecialchars($template_nama ?? basename($template_path)) ?></strong> — sesuai tanggal surat <?= date('d/m/Y', strtotime($surat['tanggal_surat'])) ?> (sebelum masa PAPBD).</p>
                </div>
            </div>
            <?php endif; ?>

            <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
                <!-- Kolom kontrol (setara col-lg-4) -->
                <div class="lg:col-span-4">
                    <div class="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden lg:sticky lg:top-4">
                        <div class="flex items-start justify-between gap-3 px-4 py-3 border-b border-slate-200 bg-slate-50/60">
                            <div>
                                <h3 class="font-semibold text-slate-800 flex items-center gap-2">
                                    <i class="bx bx-slider-alt text-indigo-600"></i>
                                    Penyesuaian Tata Letak
                                </h3>
                                <p class="text-xs text-slate-500 mt-0.5">Geser slider atau gunakan tombol + / − untuk presisi.</p>
                            </div>
                            <button type="button" id="layoutToggle" class="p-1.5 rounded-lg text-slate-500 hover:bg-slate-200 transition-colors" title="Buka/tutup panel">
                                <i class="bx bx-chevron-up text-xl" id="layoutToggleIcon"></i>
                            </button>
                        </div>

                        <div id="layoutPanelBody" class="p-4 space-y-5 max-h-[calc(100vh-8rem)] overflow-y-auto">
                            <div class="space-y-2">
                                <p class="text-xs font-semibold uppercase tracking-wider text-slate-400">Preset Cepat</p>
                                <div class="grid grid-cols-2 gap-2">
                                    <button type="button" id="layoutAutoFit" class="inline-flex items-center justify-center gap-1.5 px-3 py-2 bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-medium rounded-lg transition-colors">
                                        <i class="bx bx-collapse-vertical text-base"></i>
                                        <span>Paksa Muat 1 Halaman</span>
                                    </button>
                                    <button type="button" id="layoutReset" class="inline-flex items-center justify-center gap-1.5 px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-medium rounded-lg transition-colors">
                                        <i class="bx bx-reset text-base"></i>
                                        <span>Reset ke Default</span>
                                    </button>
                                </div>
                            </div>

                            <div class="space-y-3.5">
                                <p class="text-xs font-semibold uppercase tracking-wider text-slate-400">Pengaturan Halaman &amp; Huruf</p>

                                <div class="space-y-1.5">
                                    <div class="flex items-center justify-between gap-2">
                                        <label class="text-sm font-medium text-slate-700">Margin Atas</label>
                                        <div class="flex items-center gap-1">
                                            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="margin_top" data-dir="-1">−</button>
                                            <span class="text-xs font-semibold text-slate-700 tabular-nums w-14 text-center"><span data-value="margin_top"><?= htmlspecialchars($layout_initial['margin_top']) ?></span> mm</span>
                                            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="margin_top" data-dir="1">+</button>
                                        </div>
                                    </div>
                                    <input type="range" data-layout-key="margin_top" min="5" max="40" step="0.5" value="<?= htmlspecialchars($layout_initial['margin_top']) ?>" class="w-full accent-indigo-600">
                                    <div class="flex items-center justify-between gap-2">
                                        <span class="text-[11px] text-slate-400">Rentang 5 – 40 mm</span>
                                    </div>
                                </div>

                                <div class="space-y-1.5">
                                    <div class="flex items-center justify-between gap-2">
                                        <label class="text-sm font-medium text-slate-700">Margin Bawah</label>
                                        <div class="flex items-center gap-1">
                                            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="margin_bottom" data-dir="-1">−</button>
                                            <span class="text-xs font-semibold text-slate-700 tabular-nums w-14 text-center"><span data-value="margin_bottom"><?= htmlspecialchars($layout_initial['margin_bottom']) ?></span> mm</span>
                                            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="margin_bottom" data-dir="1">+</button>
                                        </div>
                                    </div>
                                    <input type="range" data-layout-key="margin_bottom" min="5" max="40" step="0.5" value="<?= htmlspecialchars($layout_initial['margin_bottom']) ?>" class="w-full accent-indigo-600">
                                </div>

                                <div class="space-y-1.5">
                                    <div class="flex items-center justify-between gap-2">
                                        <label class="text-sm font-medium text-slate-700">Margin Kiri</label>
                                        <div class="flex items-center gap-1">
                                            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="margin_left" data-dir="-1">−</button>
                                            <span class="text-xs font-semibold text-slate-700 tabular-nums w-14 text-center"><span data-value="margin_left"><?= htmlspecialchars($layout_initial['margin_left']) ?></span> mm</span>
                                            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="margin_left" data-dir="1">+</button>
                                        </div>
                                    </div>
                                    <input type="range" data-layout-key="margin_left" min="10" max="40" step="0.5" value="<?= htmlspecialchars($layout_initial['margin_left']) ?>" class="w-full accent-indigo-600">
                                    <div class="flex items-center justify-between gap-2">
                                        <span class="text-[11px] text-slate-400">Rentang 10 – 40 mm</span>
                                    </div>
                                </div>

                                <div class="space-y-1.5">
                                    <div class="flex items-center justify-between gap-2">
                                        <label class="text-sm font-medium text-slate-700">Margin Kanan</label>
                                        <div class="flex items-center gap-1">
                                            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="margin_right" data-dir="-1">−</button>
                                            <span class="text-xs font-semibold text-slate-700 tabular-nums w-14 text-center"><span data-value="margin_right"><?= htmlspecialchars($layout_initial['margin_right']) ?></span> mm</span>
                                            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="margin_right" data-dir="1">+</button>
                                        </div>
                                    </div>
                                    <input type="range" data-layout-key="margin_right" min="10" max="40" step="0.5" value="<?= htmlspecialchars($layout_initial['margin_right']) ?>" class="w-full accent-indigo-600">
                                </div>

                                <div class="space-y-1.5">
                                    <div class="flex items-center justify-between gap-2">
                                        <label class="text-sm font-medium text-slate-700">Ukuran Huruf Global</label>
                                        <div class="flex items-center gap-1">
                                            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="font_size" data-dir="-1">−</button>
                                            <span class="text-xs font-semibold text-slate-700 tabular-nums w-14 text-center"><span data-value="font_size"><?= htmlspecialchars($layout_initial['font_size']) ?></span> pt</span>
                                            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="font_size" data-dir="1">+</button>
                                        </div>
                                    </div>
                                    <input type="range" data-layout-key="font_size" min="9" max="13" step="0.5" value="<?= htmlspecialchars($layout_initial['font_size']) ?>" class="w-full accent-indigo-600">
                                </div>

                                <div class="space-y-1.5">
                                    <div class="flex items-center justify-between gap-2">
                                        <label class="text-sm font-medium text-slate-700">Jarak Antar Baris</label>
                                        <div class="flex items-center gap-1">
                                            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="line_height" data-dir="-1">−</button>
                                            <span class="text-xs font-semibold text-slate-700 tabular-nums w-14 text-center"><span data-value="line_height"><?= htmlspecialchars($layout_initial['line_height']) ?></span>×</span>
                                            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="line_height" data-dir="1">+</button>
                                        </div>
                                    </div>
                                    <input type="range" data-layout-key="line_height" min="1.0" max="2.0" step="0.05" value="<?= htmlspecialchars($layout_initial['line_height']) ?>" class="w-full accent-indigo-600">
                                </div>
                            </div>

                            <div class="space-y-3.5 pt-1 border-t border-slate-100">
                                <p class="text-xs font-semibold uppercase tracking-wider text-slate-400 pt-3">Pengaturan Posisi Bagian</p>

                                <div class="space-y-1.5">
                                    <div class="flex items-center justify-between gap-2">
                                        <label class="text-sm font-medium text-slate-700">Jarak Kop ke Judul</label>
                                        <div class="flex items-center gap-1">
                                            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="kop_offset" data-dir="-1">−</button>
                                            <span class="text-xs font-semibold text-slate-700 tabular-nums w-14 text-center"><span data-value="kop_offset"><?= htmlspecialchars($layout_initial['kop_offset']) ?></span> px</span>
                                            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="kop_offset" data-dir="1">+</button>
                                        </div>
                                    </div>
                                    <input type="range" data-layout-key="kop_offset" min="-20" max="60" step="1" value="<?= htmlspecialchars($layout_initial['kop_offset']) ?>" class="w-full accent-indigo-600">
                                </div>

                                <div class="space-y-1.5">
                                    <div class="flex items-center justify-between gap-2">
                                        <label class="text-sm font-medium text-slate-700">Jarak Antar Bagian</label>
                                        <div class="flex items-center gap-1">
                                            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="section_spacing" data-dir="-1">−</button>
                                            <span class="text-xs font-semibold text-slate-700 tabular-nums w-14 text-center"><span data-value="section_spacing"><?= htmlspecialchars($layout_initial['section_spacing']) ?></span> px</span>
                                            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="section_spacing" data-dir="1">+</button>
                                        </div>
                                    </div>
                                    <input type="range" data-layout-key="section_spacing" min="0" max="30" step="1" value="<?= htmlspecialchars($layout_initial['section_spacing']) ?>" class="w-full accent-indigo-600">
                                </div>

                                <div class="space-y-1.5">
                                    <div class="flex items-center justify-between gap-2">
                                        <label class="text-sm font-medium text-slate-700">Jarak Atas Blok TTD</label>
                                        <div class="flex items-center gap-1">
                                            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="ttd_offset" data-dir="-1">−</button>
                                            <span class="text-xs font-semibold text-slate-700 tabular-nums w-14 text-center"><span data-value="ttd_offset"><?= htmlspecialchars($layout_initial['ttd_offset']) ?></span> px</span>
                                            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="ttd_offset" data-dir="1">+</button>
                                        </div>
                                    </div>
                                    <input type="range" data-layout-key="ttd_offset" min="-20" max="100" step="1" value="<?= htmlspecialchars($layout_initial['ttd_offset']) ?>" class="w-full accent-indigo-600">
                                </div>

                                <div class="space-y-1.5">
                                    <div class="flex items-center justify-between gap-2">
                                        <label class="text-sm font-medium text-slate-700">Posisi Horizontal TTD</label>
                                        <div class="flex items-center gap-1">
                                            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="ttd_x" data-dir="-1">−</button>
                                            <span class="text-xs font-semibold text-slate-700 tabular-nums w-14 text-center"><span data-value="ttd_x"><?= htmlspecialchars($layout_initial['ttd_x']) ?></span> px</span>
                                            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="ttd_x" data-dir="1">+</button>
                                        </div>
                                    </div>
                                    <input type="range" data-layout-key="ttd_x" min="-50" max="50" step="1" value="<?= htmlspecialchars($layout_initial['ttd_x']) ?>" class="w-full accent-indigo-600">
                                </div>

                                <label class="flex items-start gap-2.5 p-3 rounded-lg bg-slate-50 border border-slate-200 cursor-pointer">
                                    <input type="checkbox" id="layoutLockTtd" class="mt-0.5 accent-indigo-600" <?= !empty($layout_initial['lock_ttd']) ? 'checked' : '' ?>>
                                    <span class="text-sm text-slate-700">
                                        Kunci Blok TTD
                                        <span class="block text-xs text-slate-500">Cegah blok tanda tangan terpotong ke halaman berikutnya.</span>
                                    </span>
                                </label>
                            </div>

                            <button type="button" id="layoutApplyBtn" class="w-full inline-flex items-center justify-center gap-2 px-4 py-2.5 bg-slate-800 hover:bg-slate-900 text-white text-sm font-medium rounded-lg transition-colors">
                                <i class="bx bx-check-double text-base"></i>
                                <span>Terapkan Perubahan</span>
                            </button>
                            <p class="text-[11px] text-slate-400 text-center" id="layoutStatus">Pengaturan default template</p>
                        </div>
                    </div>
                </div>

                <!-- Kolom preview (setara col-lg-8) -->
                <div class="lg:col-span-8">
                    <div class="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                        <div class="flex flex-wrap items-center justify-between gap-2 px-4 py-2 border-b border-slate-200">
                            <p class="text-sm text-slate-500">Tampilan di bawah sama dengan file PDF yang diunduh.</p>
                            <span class="inline-flex items-center gap-1.5 text-xs font-medium text-indigo-600">
                                <i class="bx bx-show"></i> Preview PDF
                            </span>
                        </div>
                        <div class="p-4 relative">
                            <div id="pdfPreviewLoading" class="hidden absolute inset-0 z-10 bg-white/80 backdrop-blur-[1px] flex items-center justify-center rounded-lg">
                                <div class="flex flex-col items-center gap-2 text-center">
                                    <div class="animate-spin rounded-full h-9 w-9 border-4 border-indigo-600 border-t-transparent"></div>
                                    <p class="text-sm font-medium text-slate-700">Memperbarui PDF…</p>
                                    <p class="text-xs text-slate-500">Menyusun ulang tata letak surat</p>
                                </div>
                            </div>
                            <iframe id="pdfPreviewFrame" src="<?= htmlspecialchars($previewPdfUrl) ?>" class="w-full border border-slate-200 rounded-lg bg-white" style="min-height: 80vh; height: 297mm;" title="Preview surat (PDF)"></iframe>
                        </div>
                    </div>
                </div>
            </div>

            <div class="flex justify-center py-4">
                <a href="surat-tugas.php?edit=<?= $sid ?>" class="inline-flex items-center gap-2 px-6 py-2.5 bg-amber-500 hover:bg-amber-600 text-white font-medium rounded-lg transition-colors shadow-sm">
                    <i class="bx bx-edit text-lg"></i>
                    <span>Edit Surat</span>
                </a>
            </div>
        </div>

        <script>
        (function () {
            const SURAT_ID = <?= json_encode((string)$sid) ?>;
            const DEFAULTS = <?= $layoutJsonDefaults ?>;
            const INITIAL = <?= $layoutJsonInitial ?>;
            const STORAGE_KEY = 'spt_layout_' + SURAT_ID;
            const STEP_CONFIG = {
                margin_top: { min: 5, max: 40, step: 0.5, decimals: 1 },
                margin_bottom: { min: 5, max: 40, step: 0.5, decimals: 1 },
                margin_left: { min: 10, max: 40, step: 0.5, decimals: 1 },
                margin_right: { min: 10, max: 40, step: 0.5, decimals: 1 },
                font_size: { min: 9, max: 13, step: 0.5, decimals: 1 },
                line_height: { min: 1.0, max: 2.0, step: 0.05, decimals: 2 },
                kop_offset: { min: -20, max: 60, step: 1, decimals: 0 },
                section_spacing: { min: 0, max: 30, step: 1, decimals: 0 },
                ttd_offset: { min: -20, max: 100, step: 1, decimals: 0 },
                ttd_x: { min: -50, max: 50, step: 1, decimals: 0 }
            };
            const KEYS = Object.keys(STEP_CONFIG);
            const frame = document.getElementById('pdfPreviewFrame');
            const loading = document.getElementById('pdfPreviewLoading');
            const downloadBtn = document.getElementById('downloadPdfBtn');
            const status = document.getElementById('layoutStatus');
            const lockCheckbox = document.getElementById('layoutLockTtd');
            const inputs = {};
            document.querySelectorAll('[data-layout-key]').forEach(function (input) {
                inputs[input.dataset.layoutKey] = input;
            });

            let state = Object.assign({}, DEFAULTS, { lock_ttd: 0 });
            let active = false;
            let restoredFromStorage = false;

            function clamp(key, value) {
                const cfg = STEP_CONFIG[key];
                value = Number(value);
                if (!isFinite(value)) return Number(DEFAULTS[key] || 0);
                value = Math.min(cfg.max, Math.max(cfg.min, value));
                return Number(value.toFixed(cfg.decimals));
            }

            function formatValue(key, value) {
                const cfg = STEP_CONFIG[key];
                return Number(value).toFixed(cfg.decimals);
            }

            // Prioritas 1: parameter di URL.
            const urlParams = new URLSearchParams(window.location.search);
            if (urlParams.get('layout') === '1') {
                active = true;
                KEYS.forEach(function (key) {
                    if (urlParams.has(key)) state[key] = clamp(key, urlParams.get(key));
                });
                state.lock_ttd = urlParams.get('lock_ttd') === '1' ? 1 : 0;
            } else {
                // Prioritas 2: pengaturan terakhir di localStorage.
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
                    // abaikan localStorage rusak
                }
            }

            function syncUi() {
                KEYS.forEach(function (key) {
                    if (inputs[key]) inputs[key].value = state[key];
                    const output = document.querySelector('[data-value="' + key + '"]');
                    if (output) output.textContent = formatValue(key, state[key]);
                });
                if (lockCheckbox) lockCheckbox.checked = !!state.lock_ttd;
                updateStatus();
            }

            function updateStatus() {
                if (!status) return;
                status.textContent = active ? 'Kustom aktif — tersimpan otomatis' : 'Pengaturan default template';
                status.className = active
                    ? 'text-[11px] text-emerald-600 text-center font-medium'
                    : 'text-[11px] text-slate-400 text-center';
            }

            function markActive() {
                active = true;
                updateStatus();
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

            function showLoading(show) {
                if (!loading) return;
                loading.classList.toggle('hidden', !show);
            }

            function applyLayout() {
                if (!active) {
                    showLoading(false);
                    return;
                }

                const cacheBuster = String(Date.now());
                const previewParams = buildParams(true);
                previewParams.set('_ts', cacheBuster);
                showLoading(true);
                frame.src = 'generate-surat.php?' + previewParams.toString();

                const downloadParams = buildParams(false);
                downloadParams.set('_ts', cacheBuster);
                downloadBtn.href = 'generate-surat.php?' + downloadParams.toString();

                try {
                    localStorage.setItem(STORAGE_KEY, JSON.stringify({
                        active: true,
                        values: Object.assign({}, state, { lock_ttd: state.lock_ttd ? 1 : 0 })
                    }));
                } catch (e) {
                    // localStorage bisa diblokir browser; preview tetap jalan.
                }

                const historyParams = buildParams(false);
                historyParams.delete('format');
                const nextUrl = new URL(window.location.href);
                nextUrl.search = historyParams.toString();
                window.history.replaceState(null, '', nextUrl.toString());
            }

            let debounceTimer = null;
            function scheduleApply(delay) {
                if (debounceTimer) clearTimeout(debounceTimer);
                debounceTimer = setTimeout(applyLayout, typeof delay === 'number' ? delay : 450);
            }

            KEYS.forEach(function (key) {
                const input = inputs[key];
                if (!input) return;
                input.addEventListener('input', function () {
                    state[key] = clamp(key, input.value);
                    const output = document.querySelector('[data-value="' + key + '"]');
                    if (output) output.textContent = formatValue(key, state[key]);
                    markActive();
                    scheduleApply();
                });
            });

            document.querySelectorAll('.layout-step').forEach(function (button) {
                button.addEventListener('click', function () {
                    const key = button.dataset.target;
                    const direction = Number(button.dataset.dir || 0);
                    const cfg = STEP_CONFIG[key];
                    if (!cfg) return;
                    state[key] = clamp(key, Number(state[key]) + (direction * cfg.step));
                    if (inputs[key]) inputs[key].value = state[key];
                    const output = document.querySelector('[data-value="' + key + '"]');
                    if (output) output.textContent = formatValue(key, state[key]);
                    markActive();
                    scheduleApply(250);
                });
            });

            if (lockCheckbox) {
                lockCheckbox.addEventListener('change', function () {
                    state.lock_ttd = lockCheckbox.checked ? 1 : 0;
                    markActive();
                    scheduleApply(250);
                });
            }

            const applyBtn = document.getElementById('layoutApplyBtn');
            if (applyBtn) {
                applyBtn.addEventListener('click', function () {
                    if (debounceTimer) clearTimeout(debounceTimer);
                    applyLayout();
                });
            }

            if (downloadBtn) {
                downloadBtn.addEventListener('click', function () {
                    // Pastikan parameter terbaru sudah masuk ke href sebelum browser
                    // memulai unduhan (jika masih ada debounce yang tertunda).
                    if (debounceTimer) {
                        clearTimeout(debounceTimer);
                        applyLayout();
                    }
                });
            }

            const autoFit = document.getElementById('layoutAutoFit');
            if (autoFit) {
                autoFit.addEventListener('click', function () {
                    state.margin_top = 8;
                    state.margin_bottom = 8;
                    state.margin_left = 15;
                    state.margin_right = 15;
                    state.font_size = 9;
                    state.line_height = 1.0;
                    state.kop_offset = 0;
                    state.section_spacing = 0;
                    state.ttd_offset = 0;
                    state.ttd_x = 0;
                    state.lock_ttd = 1;
                    markActive();
                    syncUi();
                    if (debounceTimer) clearTimeout(debounceTimer);
                    applyLayout();
                });
            }

            const resetBtn = document.getElementById('layoutReset');
            if (resetBtn) {
                resetBtn.addEventListener('click', function () {
                    active = false;
                    state = Object.assign({}, DEFAULTS, { lock_ttd: 0 });
                    try {
                        localStorage.removeItem(STORAGE_KEY);
                    } catch (e) {
                        // ignore
                    }
                    syncUi();
                    if (debounceTimer) clearTimeout(debounceTimer);
                    const params = new URLSearchParams();
                    params.set('id', SURAT_ID);
                    params.set('format', 'pdf');
                    params.set('inline', '1');
                    params.set('_ts', String(Date.now()));
                    showLoading(true);
                    frame.src = 'generate-surat.php?' + params.toString();

                    downloadBtn.href = 'generate-surat.php?id=' + encodeURIComponent(SURAT_ID) + '&format=pdf';
                    const nextUrl = new URL(window.location.href);
                    nextUrl.search = 'id=' + encodeURIComponent(SURAT_ID);
                    window.history.replaceState(null, '', nextUrl.toString());
                    updateStatus();
                });
            }

            const toggle = document.getElementById('layoutToggle');
            const panelBody = document.getElementById('layoutPanelBody');
            const toggleIcon = document.getElementById('layoutToggleIcon');
            if (toggle && panelBody) {
                toggle.addEventListener('click', function () {
                    const hidden = panelBody.classList.toggle('hidden');
                    if (toggleIcon) toggleIcon.classList.toggle('bx-rotate-180', hidden);
                });
            }

            if (frame) {
                frame.addEventListener('load', function () {
                    showLoading(false);
                });
            }

            syncUi();

            // Jika konfigurasi dipulihkan dari localStorage, langsung render ulang.
            if (restoredFromStorage) {
                applyLayout();
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
    if (is_dir($candidate) && is_writable($candidate)) {
        $tempDir = $candidate;
        break;
    }
}
    $tempDocx = $tempDir . DIRECTORY_SEPARATOR . 'spt_' . uniqid() . '.docx';
    $templateProcessor->saveAs($tempDocx);

    // Salinan DOCX yang sudah disesuaikan layout, dipakai untuk Word COM/LibreOffice.
    // File asli tetap disimpan untuk download Word dan fallback Dompdf (CSS injection).
    $layout_adjusted_docx = null;
    $pdf_source_docx = $tempDocx;
    if ($format === 'pdf' && $layout_active && $layout_settings !== null) {
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
                            // Fallback Dompdf memakai DOCX asli + CSS dinamis yang
                            // disuntikkan ke HTML tepat sebelum loadHtml().
                            $phpWord = IOFactory::load($tempDocx, 'Word2007');
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
        header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        header('Content-Disposition: attachment; filename="Surat_Tugas_' . $surat['nomor_surat'] . '.docx"');
        header('Content-Length: ' . filesize($tempDocx));
        readfile($tempDocx);
        @unlink($tempDocx);
    } catch (Exception $e) {
        if (file_exists($tempDocx)) @unlink($tempDocx);
        if ($layout_adjusted_docx && file_exists($layout_adjusted_docx)) @unlink($layout_adjusted_docx);
        throw $e;
    }

} catch (Exception $e) {
    die("Error: " . $e->getMessage());
}