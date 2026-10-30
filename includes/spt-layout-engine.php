<?php
/**
 * Mesin tata letak Surat Tugas berbasis peran blok.
 *
 * Template surat bisa berbentuk paragraf biasa (template lama) maupun satu
 * tabel besar (template PAPBD). Mesin ini membaca dokumen sebagai daftar
 * "unit" (baris tabel atau paragraf) lalu mengklasifikasikan setiap unit ke
 * peran surat: kop, judul, nomor, dasar, memerintahkan, pegawai, untuk,
 * tanda tangan, lampiran.
 *
 * Jarak vertikal diterapkan pada kontainer yang benar:
 * - unit berupa baris tabel -> margin sel (w:tcMar) baris tersebut,
 * - unit berupa paragraf    -> spacing sebelum/sesudah paragraf.
 *
 * Pendekatan ini membuat satu set pengaturan berlaku untuk kedua jenis
 * template tanpa heuristik indeks paragraf yang rapuh.
 */

function spt_le_ns()
{
    return 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
}

function spt_le_wp_ns()
{
    return 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
}

/**
 * Paragraf kosong bergambar anchor (mis. stempel) yang dipindah ke atas tabel
 * saat tabel mengambang dijadikan inline. Dikembalikan ke tinggi minimal di
 * akhir penerapan agar tidak menambah tinggi blok.
 */
function &spt_le_anchor_paragraphs()
{
    static $list = [];
    return $list;
}

/**
 * Anchor gambar stempel (bukan garis kop datar). Garis kop memakai anchor
 * dengan tinggi nol, jadi hanya anchor yang punya tinggi berarti atau blip
 * yang dianggap stempel.
 */
function spt_le_has_stamp_anchor(DOMElement $node)
{
    $wpNs = spt_le_wp_ns();
    $aNs = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    foreach ($node->getElementsByTagNameNS($wpNs, 'anchor') as $anchor) {
        $extent = null;
        foreach ($anchor->childNodes as $child) {
            if ($child instanceof DOMElement && $child->localName === 'extent') {
                $extent = $child;
                break;
            }
        }
        $cy = $extent ? (int)$extent->getAttribute('cy') : 0;
        if ($cy > 100000) {
            return true;
        }
        if ($anchor->getElementsByTagNameNS($aNs, 'blip')->length > 0) {
            return true;
        }
    }
    return false;
}

function spt_le_load_dom($xml)
{
    if (!is_string($xml) || $xml === '') {
        return null;
    }
    $dom = new DOMDocument('1.0', 'UTF-8');
    $dom->preserveWhiteSpace = true;
    $previous = libxml_use_internal_errors(true);
    $ok = $dom->loadXML($xml, LIBXML_NONET | LIBXML_PARSEHUGE);
    libxml_clear_errors();
    libxml_use_internal_errors($previous);
    return $ok ? $dom : null;
}

function spt_le_el(DOMDocument $doc, $name)
{
    return $doc->createElementNS(spt_le_ns(), 'w:' . $name);
}

function spt_le_child(DOMElement $parent, $localName)
{
    foreach ($parent->childNodes as $child) {
        if ($child instanceof DOMElement
            && $child->localName === $localName
            && $child->namespaceURI === spt_le_ns()) {
            return $child;
        }
    }
    return null;
}

function spt_le_insert_ordered(DOMElement $parent, DOMElement $node, array $order)
{
    $position = array_search($node->localName, $order, true);
    if ($position === false) {
        $parent->appendChild($node);
        return;
    }
    foreach ($parent->childNodes as $child) {
        if (!$child instanceof DOMElement) {
            continue;
        }
        $childPosition = array_search($child->localName, $order, true);
        if ($childPosition !== false && $childPosition > $position) {
            $parent->insertBefore($node, $child);
            return;
        }
    }
    $parent->appendChild($node);
}

function spt_le_pPr(DOMElement $paragraph, DOMDocument $doc)
{
    $pPr = spt_le_child($paragraph, 'pPr');
    if (!$pPr) {
        $pPr = spt_le_el($doc, 'pPr');
        if ($paragraph->firstChild) {
            $paragraph->insertBefore($pPr, $paragraph->firstChild);
        } else {
            $paragraph->appendChild($pPr);
        }
    }
    return $pPr;
}

function spt_le_spacing(DOMElement $paragraph, DOMDocument $doc)
{
    $pPr = spt_le_pPr($paragraph, $doc);
    $spacing = spt_le_child($pPr, 'spacing');
    if (!$spacing) {
        $spacing = spt_le_el($doc, 'spacing');
        spt_le_insert_ordered($pPr, $spacing, [
            'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
            'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd',
            'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
            'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN',
            'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
            'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
            'textDirection', 'textAlignment', 'textboxTightWrap',
            'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange',
        ]);
    }
    return $spacing;
}

function spt_le_indent(DOMElement $paragraph, DOMDocument $doc)
{
    $pPr = spt_le_pPr($paragraph, $doc);
    $ind = spt_le_child($pPr, 'ind');
    if (!$ind) {
        $ind = spt_le_el($doc, 'ind');
        spt_le_insert_ordered($pPr, $ind, [
            'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
            'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd',
            'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
            'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN',
            'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
            'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
            'textDirection', 'textAlignment', 'textboxTightWrap',
            'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange',
        ]);
    }
    return $ind;
}

function spt_le_add_keep_flag(DOMElement $paragraph, $flag, DOMDocument $doc)
{
    $pPr = spt_le_pPr($paragraph, $doc);
    if (spt_le_child($pPr, $flag)) {
        return;
    }
    $element = spt_le_el($doc, $flag);
    spt_le_insert_ordered($pPr, $element, [
        'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
        'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd',
        'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
        'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
        'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents',
        'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
        'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
        'sectPr', 'pPrChange',
    ]);
}

function spt_le_unit_text(DOMElement $node)
{
    $text = '';
    foreach ($node->getElementsByTagNameNS(spt_le_ns(), 't') as $textNode) {
        $text .= $textNode->textContent;
    }
    return trim(preg_replace('/\s+/u', ' ', $text));
}

/**
 * Kumpulkan unit dokumen secara berurutan.
 * Unit = satu baris tabel (w:tr) atau satu paragraf badan (w:p).
 */
function spt_le_collect_units(DOMDocument $doc)
{
    $ns = spt_le_ns();
    $xpath = new DOMXPath($doc);
    $xpath->registerNamespace('w', $ns);
    $body = $xpath->query('//w:body')->item(0);
    if (!$body) {
        return [];
    }

    $units = [];
    foreach ($body->childNodes as $child) {
        if (!$child instanceof DOMElement || $child->namespaceURI !== $ns) {
            continue;
        }
        if ($child->localName === 'tbl') {
            foreach ($child->childNodes as $row) {
                if ($row instanceof DOMElement
                    && $row->localName === 'tr'
                    && $row->namespaceURI === $ns) {
                    $units[] = [
                        'kind' => 'row',
                        'node' => $row,
                        'table' => $child,
                        'text' => spt_le_unit_text($row),
                    ];
                }
            }
        } elseif ($child->localName === 'p') {
            $units[] = [
                'kind' => 'p',
                'node' => $child,
                'table' => null,
                'text' => spt_le_unit_text($child),
            ];
        }
    }
    return $units;
}

function spt_le_first_match(array $units, $from, $pattern)
{
    $count = count($units);
    for ($i = max(0, $from); $i < $count; $i++) {
        if ($units[$i]['text'] !== '' && preg_match($pattern, $units[$i]['text'])) {
            return $i;
        }
    }
    return null;
}

function spt_le_city_date_sub($text)
{
    if (preg_match('/^[A-Z][\p{L}\.\']{2,30}\s*,?\s*\d{1,2}\s+[A-Za-z]+\s+\d{4}/u', $text)) {
        return true;
    }
    return false;
}

function spt_le_ttd_sub($text)
{
    if ($text === '') {
        return 'spacer';
    }
    if (spt_le_city_date_sub($text)) {
        return 'date';
    }
    if (preg_match('/^A\.?\s*n\b/iu', $text)
        || preg_match('/^Kepala\b/iu', $text)
        || preg_match('/^Wilayah\b/iu', $text)
        || preg_match('/Cabang\s+Dinas/iu', $text)
        || preg_match('/^Direktur\b/iu', $text)) {
        return 'jabatan';
    }
    if (preg_match('/\bNIP\b/iu', $text) || preg_match('/^NIP\b/iu', $text)) {
        return 'nip';
    }
    if (preg_match('/\((?:I{1,3}|IV|V|[1-4])(?:[\/\w\.]*)\)/u', $text)
        || preg_match('/\bPembina\b|\bPenata\b|\bPengatur\b|\bJuru\b/iu', $text)) {
        return 'pangkat';
    }
    return 'name';
}

/**
 * Tentukan peran setiap unit.
 * Return: ['roles' => [i => ['role'=>..,'sub'=>..,'employee'=>int]], 'anchors' => [...]]
 */
function spt_le_detect_roles(array $units)
{
    $count = count($units);
    $roles = [];
    for ($i = 0; $i < $count; $i++) {
        $roles[$i] = ['role' => 'other', 'sub' => '', 'employee' => null];
    }

    $judul = spt_le_first_match($units, 0, '/SURAT\s+TUGAS/iu');
    // "Lampiran Surat Tugas" bukan judul surat utama.
    if ($judul !== null && preg_match('/Lampiran/iu', $units[$judul]['text'])) {
        $judul = null;
        for ($i = 0; $i < $count; $i++) {
            if (preg_match('/SURAT\s+TUGAS/iu', $units[$i]['text'])
                && !preg_match('/Lampiran/iu', $units[$i]['text'])) {
                $judul = $i;
                break;
            }
        }
    }
    $nomor = $judul !== null ? spt_le_first_match($units, $judul, '/^\s*NOMOR\b|^\s*Nomor\s*:/iu') : null;
    if ($nomor === $judul) {
        $nomor = null;
        $nomor = spt_le_first_match($units, $judul + 1, '/^\s*NOMOR\b|^\s*Nomor\s*:/iu');
    }
    $dasar = $nomor !== null ? spt_le_first_match($units, $nomor, '/^\s*DASAR\b/iu') : null;
    $mem = $dasar !== null ? spt_le_first_match($units, $dasar + 1, '/^\s*MEMERINTAHKAN\b/iu') : null;
    $kepada = $mem !== null ? spt_le_first_match($units, $mem + 1, '/^\s*KEPADA\b/iu') : null;
    $untuk = ($kepada ?? $mem) !== null
        ? spt_le_first_match($units, ($kepada ?? $mem) + 1, '/^\s*UNTUK\b/iu')
        : null;

    $ttd = null;
    $ttdFrom = ($untuk ?? $kepada ?? $mem) !== null ? ($untuk ?? $kepada ?? $mem) + 1 : 0;
    for ($i = $ttdFrom; $i < $count; $i++) {
        $text = $units[$i]['text'];
        if ($text === '') {
            continue;
        }
        if (spt_le_city_date_sub($text) || preg_match('/^A\.?\s*n\b/iu', $text)) {
            $ttd = $i;
            break;
        }
        if (preg_match('/^Kepala\b/iu', $text) && preg_match('/Cabang\s+Dinas/iu', $text)) {
            // Mundur satu unit jika baris sebelumnya adalah tempat/tanggal.
            $ttd = $i;
            for ($k = $i - 1; $k >= $ttdFrom; $k--) {
                if ($units[$k]['text'] !== '') {
                    if (spt_le_city_date_sub($units[$k]['text'])) {
                        $ttd = $k;
                    }
                    break;
                }
            }
            break;
        }
    }

    // Fallback: baris NIP terakhir menandai blok tanda tangan.
    if ($ttd === null) {
        for ($i = $count - 1; $i >= 0; $i--) {
            if (preg_match('/\bNIP\b/iu', $units[$i]['text'])) {
                $start = $i;
                for ($k = $i; $k >= 0 && $k >= $i - 8; $k--) {
                    $start = $k;
                    if ($units[$k]['text'] !== '' && spt_le_city_date_sub($units[$k]['text'])) {
                        break;
                    }
                }
                $ttd = $start;
                break;
            }
        }
    }

    $lampiran = $ttd !== null
        ? spt_le_first_match($units, $ttd + 1, '/Lampiran\s+(?:Surat\s+)?Tugas/iu')
        : null;

    // Blok TTD berhenti di baris NIP. Tanpa batas ini, paragraf kosong
    // "page break" sebelum lampiran ikut dianggap spacer TTD dan judul
    // lampiran bisa tertarik ke halaman surat.
    $ttdEnd = $ttd !== null ? (($lampiran ?? $count) - 1) : null;
    if ($ttd !== null) {
        $nonEmpty = 0;
        for ($i = $ttd; $i < $count; $i++) {
            if ($lampiran !== null && $i >= $lampiran) {
                $ttdEnd = $i - 1;
                break;
            }
            if ($units[$i]['text'] === '') {
                continue;
            }
            $nonEmpty++;
            $ttdEnd = $i;
            if (preg_match('/\bNIP\b/iu', $units[$i]['text']) || $nonEmpty >= 7) {
                break;
            }
        }
    }

    $ranges = [
        ['kop', 0, ($judul ?? $count) - 1],
        ['judul', $judul, ($nomor ?? ($judul === null ? -1 : $judul + 1)) - 1],
        ['nomor', $nomor, ($dasar ?? ($nomor === null ? -1 : $nomor + 1)) - 1],
        ['dasar', $dasar, ($mem ?? ($dasar === null ? -1 : $dasar + 1)) - 1],
        ['memerintahkan', $mem, ($kepada ?? ($mem === null ? -1 : $mem + 1)) - 1],
        ['pegawai', $kepada, ($untuk ?? ($kepada === null ? -1 : $kepada + 1)) - 1],
        ['untuk', $untuk, ($ttd ?? ($untuk === null ? -1 : $untuk + 1)) - 1],
        ['ttd', $ttd, $ttdEnd],
        ['lampiran', $lampiran, $count - 1],
    ];

    $ordered = ['kop', 'judul', 'nomor', 'dasar', 'memerintahkan', 'pegawai', 'untuk', 'ttd', 'lampiran'];
    foreach ($ordered as $name) {
        foreach ($ranges as $range) {
            if ($range[0] !== $name || $range[1] === null || $range[2] === null) {
                continue;
            }
            $from = max(0, (int)$range[1]);
            $to = min($count - 1, (int)$range[2]);
            for ($i = $from; $i <= $to; $i++) {
                if ($roles[$i]['role'] === 'other') {
                    $roles[$i]['role'] = $name;
                }
            }
        }
    }

    // Subperan blok pegawai.
    $employee = 0;
    foreach ($roles as $i => $info) {
        if ($info['role'] !== 'pegawai') {
            continue;
        }
        $text = $units[$i]['text'];
        if ($text === '') {
            $roles[$i]['sub'] = 'spacer';
            continue;
        }
        if (preg_match('/\bNama\s*:/iu', $text)
            || preg_match('/^\d+[\.\)]?\s*Nama\s*:/iu', $text)) {
            $employee++;
            $roles[$i]['sub'] = 'start';
            $roles[$i]['employee'] = $employee;
        } elseif (preg_match('/\bNIP\b/iu', $text)
            || preg_match('/\bPangkat\b|\bGol(?:ongan)?\b/iu', $text)
            || preg_match('/\bJabatan\b/iu', $text)) {
            $roles[$i]['sub'] = 'identity';
            $roles[$i]['employee'] = $employee > 0 ? $employee : null;
        } elseif ($i === $kepada) {
            $roles[$i]['sub'] = 'label';
        } else {
            $roles[$i]['sub'] = 'item';
        }
    }

    // Subperan blok tanda tangan.
    $seenName = false;
    foreach ($roles as $i => $info) {
        if ($info['role'] !== 'ttd') {
            continue;
        }
        $text = $units[$i]['text'];
        $sub = spt_le_ttd_sub($text);
        if ($sub === 'spacer' && spt_le_unit_has_image($units[$i])) {
            // Baris gambar tanda tangan/stempel: jangan diperlakukan sebagai ruang kosong.
            $sub = 'signature';
        }
        if ($sub === 'name') {
            if ($seenName) {
                $sub = 'other';
            } else {
                $seenName = true;
            }
        }
        $roles[$i]['sub'] = $sub;
    }

    return [
        'roles' => $roles,
        'anchors' => [
            'judul' => $judul,
            'nomor' => $nomor,
            'dasar' => $dasar,
            'memerintahkan' => $mem,
            'kepada' => $kepada,
            'untuk' => $untuk,
            'ttd' => $ttd,
            'lampiran' => $lampiran,
        ],
    ];
}

function spt_le_px_to_twips($px)
{
    return (int)round(((float)$px) * 15);
}

function spt_le_mm_to_twips($mm)
{
    return (int)round(((float)$mm) * 1440 / 25.4);
}

/* ---------------------------------------------------------------------
 * Penulisan jarak pada unit
 * ------------------------------------------------------------------- */

function spt_le_row_zero_paragraph_spacing(DOMElement $row, DOMDocument $doc)
{
    $ns = spt_le_ns();
    foreach ($row->getElementsByTagNameNS($ns, 'p') as $paragraph) {
        $spacing = spt_le_spacing($paragraph, $doc);
        $spacing->setAttributeNS($ns, 'w:before', '0');
        $spacing->setAttributeNS($ns, 'w:after', '0');
    }
}

function spt_le_paragraph_has_image(DOMElement $paragraph)
{
    $ns = spt_le_ns();
    return $paragraph->getElementsByTagNameNS($ns, 'drawing')->length > 0
        || $paragraph->getElementsByTagNameNS($ns, 'pict')->length > 0;
}

function spt_le_paragraph_is_blank(DOMElement $paragraph)
{
    $ns = spt_le_ns();
    foreach ($paragraph->getElementsByTagNameNS($ns, 't') as $text) {
        if (trim((string)$text->nodeValue) !== '') {
            return false;
        }
    }
    return !spt_le_paragraph_has_image($paragraph);
}

function spt_le_unit_has_image(array $unit)
{
    $ns = spt_le_ns();
    $node = $unit['node'];
    return $node->getElementsByTagNameNS($ns, 'drawing')->length > 0
        || $node->getElementsByTagNameNS($ns, 'pict')->length > 0;
}

function spt_le_collapse_paragraph(DOMElement $paragraph, DOMDocument $doc)
{
    if (spt_le_paragraph_has_image($paragraph)) {
        return;
    }
    $ns = spt_le_ns();
    $spacing = spt_le_spacing($paragraph, $doc);
    $spacing->setAttributeNS($ns, 'w:before', '0');
    $spacing->setAttributeNS($ns, 'w:after', '0');
    $spacing->setAttributeNS($ns, 'w:line', '20');
    $spacing->setAttributeNS($ns, 'w:lineRule', 'exact');
    if (!spt_le_paragraph_is_blank($paragraph)) {
        return;
    }
    // Tanda paragraf 1pt: tanpa ini Word tetap memakai tinggi font asli untuk
    // baris kosong meski w:line sudah 20 twips.
    $pPr = spt_le_pPr($paragraph, $doc);
    $rPr = spt_le_child($pPr, 'rPr');
    if (!$rPr) {
        $rPr = spt_le_el($doc, 'rPr');
        spt_le_insert_ordered($pPr, $rPr, [
            'pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr',
            'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd',
            'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
            'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN',
            'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
            'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
            'textDirection', 'textAlignment', 'textboxTightWrap',
            'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange',
        ]);
    }
    foreach (['sz', 'szCs'] as $tag) {
        $size = spt_le_child($rPr, $tag);
        if (!$size) {
            $size = spt_le_el($doc, $tag);
            $rPr->appendChild($size);
        }
        $size->setAttributeNS($ns, 'w:val', '2');
    }
}

function spt_le_row_set_margins(DOMElement $row, DOMDocument $doc, $topTwips = null, $bottomTwips = null)
{
    $ns = spt_le_ns();
    $tcPrOrder = [
        'cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd',
        'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark',
        'headers', 'cellIns', 'cellDel', 'cellMerge', 'tcPrChange',
    ];
    $marginOrder = ['top', 'start', 'left', 'bottom', 'end', 'right'];

    foreach ($row->childNodes as $cell) {
        if (!$cell instanceof DOMElement || $cell->localName !== 'tc' || $cell->namespaceURI !== $ns) {
            continue;
        }
        $tcPr = spt_le_child($cell, 'tcPr');
        if (!$tcPr) {
            $tcPr = spt_le_el($doc, 'tcPr');
            if ($cell->firstChild) {
                $cell->insertBefore($tcPr, $cell->firstChild);
            } else {
                $cell->appendChild($tcPr);
            }
        }
        // Sel lanjutan vMerge juga perlu tcMar agar padding bawah baris
        // benar-benar diterapkan Word pada seluruh tinggi gabungan sel.
        $margins = spt_le_child($tcPr, 'tcMar');
        if (!$margins) {
            $margins = spt_le_el($doc, 'tcMar');
            spt_le_insert_ordered($tcPr, $margins, $tcPrOrder);
        }
        foreach ([['top', $topTwips], ['bottom', $bottomTwips]] as $pair) {
            list($side, $twips) = $pair;
            if ($twips === null) {
                continue;
            }
            $node = spt_le_child($margins, $side);
            if (!$node) {
                $node = spt_le_el($doc, $side);
                spt_le_insert_ordered($margins, $node, $marginOrder);
            }
            $node->setAttributeNS($ns, 'w:w', (string)max(0, (int)round($twips)));
            $node->setAttributeNS($ns, 'w:type', 'dxa');
        }
    }
}

function spt_le_row_ensure_height(DOMElement $row, DOMDocument $doc, $twips, $rule = 'atLeast')
{
    $ns = spt_le_ns();
    $trPr = spt_le_child($row, 'trPr');
    if (!$trPr) {
        $trPr = spt_le_el($doc, 'trPr');
        if ($row->firstChild) {
            $row->insertBefore($trPr, $row->firstChild);
        } else {
            $row->appendChild($trPr);
        }
    }
    $height = spt_le_child($trPr, 'trHeight');
    if (!$height) {
        $height = spt_le_el($doc, 'trHeight');
        spt_le_insert_ordered($trPr, $height, [
            'cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter',
            'cantSplit', 'trHeight', 'tblHeader', 'tblCellSpacing', 'jc',
            'hidden', 'ins', 'del', 'trPrChange',
        ]);
    }
    $height->setAttributeNS($ns, 'w:val', (string)max(0, (int)$twips));
    $height->setAttributeNS($ns, 'w:hRule', $rule);
}

function spt_le_unit_set_margins(array $unit, DOMDocument $doc, $topTwips = null, $bottomTwips = null, $preserveHeight = false)
{
    if ($unit['kind'] === 'row') {
        spt_le_row_zero_paragraph_spacing($unit['node'], $doc);
        spt_le_row_set_margins($unit['node'], $doc, $topTwips, $bottomTwips);
        if (!$preserveHeight) {
            // trHeight yang diwarisi template menahan penyusutan baris;
            // biarkan teks + tcMar yang menentukan tinggi baris.
            $trPr = spt_le_child($unit['node'], 'trPr');
            if ($trPr) {
                $height = spt_le_child($trPr, 'trHeight');
                if ($height) {
                    $trPr->removeChild($height);
                }
            }
        }
        return;
    }
    $ns = spt_le_ns();
    $spacing = spt_le_spacing($unit['node'], $doc);
    if ($topTwips !== null) {
        $spacing->setAttributeNS($ns, 'w:before', (string)max(0, (int)$topTwips));
    }
    if ($bottomTwips !== null) {
        $spacing->setAttributeNS($ns, 'w:after', (string)max(0, (int)$bottomTwips));
    }
}

function spt_le_unit_collapse(array $unit, DOMDocument $doc)
{
    if ($unit['kind'] === 'row') {
        $ns = spt_le_ns();
        spt_le_row_zero_paragraph_spacing($unit['node'], $doc);
        foreach ($unit['node']->getElementsByTagNameNS($ns, 'p') as $paragraph) {
            spt_le_collapse_paragraph($paragraph, $doc);
        }
        spt_le_row_set_margins($unit['node'], $doc, 0, 0);
        spt_le_row_ensure_height($unit['node'], $doc, 20, 'exact');
        return;
    }
    spt_le_collapse_paragraph($unit['node'], $doc);
}

function spt_le_unit_add_line(array $unit, DOMDocument $doc, $extraTwips)
{
    if ($extraTwips <= 0) {
        return;
    }
    $ns = spt_le_ns();
    $paragraphs = $unit['kind'] === 'row'
        ? iterator_to_array($unit['node']->getElementsByTagNameNS($ns, 'p'))
        : [$unit['node']];
    foreach ($paragraphs as $paragraph) {
        $spacing = spt_le_spacing($paragraph, $doc);
        $current = $spacing->hasAttributeNS($ns, 'line')
            ? (int)$spacing->getAttributeNS($ns, 'line')
            : 0;
        $base = max($current, 240);
        $spacing->setAttributeNS($ns, 'w:line', (string)($base + (int)$extraTwips));
        $spacing->setAttributeNS($ns, 'w:lineRule', 'auto');
    }
}

function spt_le_unit_reserve_space(array $unit, DOMDocument $doc, $twips)
{
    $ns = spt_le_ns();
    if ($unit['kind'] === 'row') {
        spt_le_row_zero_paragraph_spacing($unit['node'], $doc);
        spt_le_row_set_margins($unit['node'], $doc, 0, 0);
        foreach ($unit['node']->getElementsByTagNameNS($ns, 'p') as $paragraph) {
            spt_le_collapse_paragraph($paragraph, $doc);
        }
        spt_le_row_ensure_height($unit['node'], $doc, $twips, 'atLeast');
        return;
    }
    $spacing = spt_le_spacing($unit['node'], $doc);
    $spacing->setAttributeNS($ns, 'w:before', '0');
    $spacing->setAttributeNS($ns, 'w:after', '0');
    $spacing->setAttributeNS($ns, 'w:line', (string)max(20, (int)$twips));
    $spacing->setAttributeNS($ns, 'w:lineRule', 'atLeast');
}

/* ---------------------------------------------------------------------
 * Transformasi global
 * ------------------------------------------------------------------- */

function spt_le_dominant_font(array $units)
{
    $counts = [];
    foreach ($units as $unit) {
        $scope = $unit['kind'] === 'row'
            ? $unit['node']->getElementsByTagNameNS(spt_le_ns(), 'r')
            : $unit['node']->getElementsByTagNameNS(spt_le_ns(), 'r');
        foreach ($scope as $run) {
            if ($run->getElementsByTagNameNS(spt_le_ns(), 't')->length === 0) {
                continue;
            }
            $rPr = spt_le_child($run, 'rPr');
            $sz = $rPr ? spt_le_child($rPr, 'sz') : null;
            if (!$sz) {
                continue;
            }
            $value = (int)$sz->getAttributeNS(spt_le_ns(), 'val');
            if ($value > 0) {
                $counts[$value] = ($counts[$value] ?? 0) + 1;
            }
        }
    }
    if (empty($counts)) {
        return 22.0;
    }
    arsort($counts);
    return (float)key($counts);
}

function spt_le_dominant_line(array $units)
{
    $counts = [];
    foreach ($units as $unit) {
        $paragraphs = $unit['kind'] === 'row'
            ? $unit['node']->getElementsByTagNameNS(spt_le_ns(), 'p')
            : [$unit['node']];
        foreach ($paragraphs as $paragraph) {
            $pPr = spt_le_child($paragraph, 'pPr');
            $spacing = $pPr ? spt_le_child($pPr, 'spacing') : null;
            if (!$spacing || !$spacing->hasAttributeNS(spt_le_ns(), 'line')) {
                continue;
            }
            $value = (int)$spacing->getAttributeNS(spt_le_ns(), 'line');
            // Nilai exact di bawah 180 adalah paragraf yang sengaja
            // diciutkan (1pt), bukan spasi baris isi surat.
            if ($value >= 180) {
                $counts[$value] = ($counts[$value] ?? 0) + 1;
            }
        }
    }
    if (empty($counts)) {
        return 240.0;
    }
    arsort($counts);
    return (float)key($counts);
}

function spt_le_scale_font_sizes(DOMDocument $doc, $factor)
{
    if (abs($factor - 1.0) < 0.0001) {
        return;
    }
    $ns = spt_le_ns();
    $xpath = new DOMXPath($doc);
    $xpath->registerNamespace('w', $ns);
    foreach ($xpath->query('//w:sz | //w:szCs') as $node) {
        $value = (int)$node->getAttributeNS($ns, 'val');
        if ($value <= 0) {
            continue;
        }
        $node->setAttributeNS($ns, 'w:val', (string)max(2, min(3276, (int)round($value * $factor))));
    }
}

function spt_le_apply_line_height(array $units, DOMDocument $doc, $baseLine, $targetHeight, array $exclude = [])
{
    if ($targetHeight <= 0) {
        return;
    }
    $ns = spt_le_ns();
    $targetTwips = (int)round($targetHeight * 240);
    $ratio = $baseLine > 0 ? ($targetTwips / $baseLine) : 1.0;

    foreach ($units as $index => $unit) {
        if (isset($exclude[$index])) {
            continue;
        }
        $paragraphs = $unit['kind'] === 'row'
            ? $unit['node']->getElementsByTagNameNS($ns, 'p')
            : [$unit['node']];
        foreach ($paragraphs as $paragraph) {
            $spacing = spt_le_spacing($paragraph, $doc);
            $hasLine = $spacing->hasAttributeNS($ns, 'line');
            $rule = $hasLine ? $spacing->getAttributeNS($ns, 'lineRule') : '';
            $current = $hasLine ? (int)$spacing->getAttributeNS($ns, 'line') : 0;
            if (spt_le_paragraph_is_blank($paragraph)
                && $rule === 'exact'
                && $current <= 40) {
                // Sudah diciutkan menjadi 1pt.
                continue;
            }
            if ($rule === 'exact') {
                // Template PAPBD mengunci spasi baris dengan lineRule exact.
                // Ubah ke auto agar slider line_height dan kompresi fit
                // benar-benar mengubah jarak baris isi surat.
                $spacing->setAttributeNS($ns, 'w:line', (string)$targetTwips);
                $spacing->setAttributeNS($ns, 'w:lineRule', 'auto');
                continue;
            }
            if ($hasLine && $current > 0) {
                $spacing->setAttributeNS($ns, 'w:line', (string)max(20, (int)round($current * $ratio)));
                if (!$spacing->hasAttributeNS($ns, 'lineRule')) {
                    $spacing->setAttributeNS($ns, 'w:lineRule', 'auto');
                }
                continue;
            }
            $spacing->setAttributeNS($ns, 'w:line', (string)$targetTwips);
            $spacing->setAttributeNS($ns, 'w:lineRule', 'auto');
        }
    }
}

function spt_le_set_page_margins(DOMDocument $doc, array $settings)
{
    $ns = spt_le_ns();
    $xpath = new DOMXPath($doc);
    $xpath->registerNamespace('w', $ns);
    $margins = [
        'top' => spt_le_mm_to_twips($settings['margin_top']),
        'bottom' => spt_le_mm_to_twips($settings['margin_bottom']),
        'left' => spt_le_mm_to_twips($settings['margin_left']),
        'right' => spt_le_mm_to_twips($settings['margin_right']),
    ];
    foreach ($xpath->query('//w:sectPr') as $section) {
        $pgMar = spt_le_child($section, 'pgMar');
        if (!$pgMar) {
            $pgMar = spt_le_el($doc, 'pgMar');
            spt_le_insert_ordered($section, $pgMar, [
                'headerReference', 'footerReference', 'footnotePr', 'endnotePr',
                'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType',
                'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote',
                'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid',
                'printerSettings', 'sectPrChange',
            ]);
        }
        foreach ($margins as $side => $twips) {
            $pgMar->setAttributeNS($ns, 'w:' . $side, (string)$twips);
        }
    }
}

function spt_le_align_tables_top(DOMDocument $doc, array $units = [], array $roles = [])
{
    $ns = spt_le_ns();
    $xpath = new DOMXPath($doc);
    $xpath->registerNamespace('w', $ns);
    $tcPrOrder = [
        'cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd',
        'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark',
        'headers', 'cellIns', 'cellDel', 'cellMerge', 'tcPrChange',
    ];
    $rowRoles = [];
    foreach ($units as $i => $unit) {
        if ($unit['kind'] === 'row') {
            $rowRoles[spl_object_id($unit['node'])] = [
                'role' => $roles[$i]['role'] ?? '',
                'sub' => $roles[$i]['sub'] ?? '',
            ];
        }
    }
    foreach ($xpath->query('//w:tbl//w:tc') as $cell) {
        $tcPr = spt_le_child($cell, 'tcPr');
        if (!$tcPr) {
            $tcPr = spt_le_el($doc, 'tcPr');
            if ($cell->firstChild) {
                $cell->insertBefore($tcPr, $cell->firstChild);
            } else {
                $cell->appendChild($tcPr);
            }
        }
        $vAlign = spt_le_child($tcPr, 'vAlign');
        if (!$vAlign) {
            $vAlign = spt_le_el($doc, 'vAlign');
            spt_le_insert_ordered($tcPr, $vAlign, $tcPrOrder);
        }
        $vAlign->setAttributeNS($ns, 'w:val', 'top');
    }
    foreach ($xpath->query('//w:tbl//w:tr/w:trPr/w:trHeight') as $rowHeight) {
        $tr = $rowHeight->parentNode ? $rowHeight->parentNode->parentNode : null;
        $info = $tr instanceof DOMElement
            ? ($rowRoles[spl_object_id($tr)] ?? null)
            : null;
        $val = (int)$rowHeight->getAttributeNS($ns, 'val');
        $rule = $rowHeight->getAttributeNS($ns, 'hRule');
        if ($rule === 'exact' && $val > 0 && $val <= 40) {
            // Baris kosong yang sudah diciutkan ke 1pt.
            continue;
        }
        if ($val > 40) {
            if ($rule === 'exact') {
                $rowHeight->setAttributeNS($ns, 'w:hRule', 'atLeast');
            }
            $isKop = $info !== null && $info['role'] === 'kop';
            $isTtdSpacer = $info !== null && $info['role'] === 'ttd' && $info['sub'] === 'spacer';
            if (!$isKop && !$isTtdSpacer) {
                // Tinggi baris tidak boleh menahan penyusutan jarak.
                $rowHeight->setAttributeNS($ns, 'w:val', '0');
            }
            continue;
        }
        if ($rule !== 'exact') {
            $rowHeight->setAttributeNS($ns, 'w:hRule', 'atLeast');
        }
    }
}

/* ---------------------------------------------------------------------
 * Penerapan per blok
 * ------------------------------------------------------------------- */

function spt_le_block_items(array $units, array $roles, $role)
{
    $items = [];
    foreach ($roles as $i => $info) {
        if ($info['role'] === $role) {
            $items[] = $i;
        }
    }
    return $items;
}

function spt_le_apply_block(array $units, array $roles, DOMDocument $doc, $role, $topPx, $itemPx, $bottomPx)
{
    $items = spt_le_block_items($units, $roles, $role);
    if (empty($items)) {
        return;
    }
    $content = [];
    foreach ($items as $i) {
        if ($units[$i]['text'] === '') {
            spt_le_unit_collapse($units[$i], $doc);
            continue;
        }
        $content[] = $i;
    }
    $count = count($content);
    foreach ($content as $position => $i) {
        $top = $position === 0 ? spt_le_px_to_twips($topPx) : 0;
        $bottom = $position === $count - 1
            ? spt_le_px_to_twips($bottomPx)
            : spt_le_px_to_twips($itemPx);
        spt_le_unit_set_margins($units[$i], $doc, $top, $bottom);
    }
}

function spt_le_apply_pegawai(array $units, array $roles, DOMDocument $doc, array $settings)
{
    $items = spt_le_block_items($units, $roles, 'pegawai');
    if (empty($items)) {
        return;
    }
    $topTwips = spt_le_px_to_twips($settings['pegawai_top_spacing'] ?? 6);
    $gapTwips = spt_le_px_to_twips($settings['pegawai_gap_spacing'] ?? 8);
    $lineTwips = spt_le_px_to_twips($settings['pegawai_line_spacing'] ?? 2);

    $contentSeen = 0;
    foreach ($items as $i) {
        $unit = $units[$i];
        $info = $roles[$i];
        if ($unit['text'] === '') {
            spt_le_unit_collapse($unit, $doc);
            continue;
        }
        $contentSeen++;
        $top = 0;
        if ($contentSeen === 1) {
            $top = $topTwips;
        } elseif ($info['sub'] === 'start' && ($info['employee'] ?? 0) > 1) {
            $top = $gapTwips;
        }
        $bottom = ($info['sub'] === 'identity' || $info['sub'] === 'start') ? $lineTwips : 0;
        spt_le_unit_set_margins($unit, $doc, $top, $bottom);
    }
}

function spt_le_apply_ttd(array $units, array $roles, DOMDocument $doc, array $settings)
{
    $items = spt_le_block_items($units, $roles, 'ttd');
    if (empty($items)) {
        return;
    }
    $offsetTwips = spt_le_px_to_twips($settings['ttd_offset'] ?? 10);
    $dateToJabatan = spt_le_px_to_twips($settings['ttd_date_to_jabatan'] ?? 4);
    $spaceTwips = spt_le_px_to_twips($settings['ttd_space_height'] ?? 60);
    $nameTwips = spt_le_px_to_twips($settings['ttd_name_nip_spacing'] ?? 4);

    $spacerUsed = false;
    $signatureSeen = false;
    $hasSignature = false;
    foreach ($items as $i) {
        if ($roles[$i]['sub'] === 'signature') {
            $hasSignature = true;
            break;
        }
    }
    foreach ($items as $i) {
        $unit = $units[$i];
        $sub = $roles[$i]['sub'];
        if ($sub === 'signature') {
            // Tinggi baris mengikuti gambar; jangan diciutkan atau diberi line exact.
            spt_le_unit_set_margins($unit, $doc, 0, 0);
            $spacerUsed = true;
            $signatureSeen = true;
            continue;
        }
        if ($sub === 'spacer') {
            if ($hasSignature && !$signatureSeen) {
                // Area tanda tangan sudah diisi gambar, jadi tidak perlu ruang
                // kosong seukuran tanda tangan basah. Sisakan satu baris.
                spt_le_unit_set_margins($unit, $doc, 0, 0);
                continue;
            }
            if (!$spacerUsed) {
                spt_le_unit_reserve_space($unit, $doc, $spaceTwips);
                $spacerUsed = true;
            } else {
                spt_le_unit_collapse($unit, $doc);
            }
            continue;
        }
        if ($sub === 'date') {
            spt_le_unit_set_margins($unit, $doc, $offsetTwips, $dateToJabatan);
            continue;
        }
        if ($sub === 'jabatan') {
            spt_le_unit_set_margins($unit, $doc, 0, 0);
            continue;
        }
        if ($sub === 'name' || $sub === 'nip' || $sub === 'pangkat') {
            spt_le_unit_set_margins($unit, $doc, $nameTwips, 0);
            continue;
        }
        spt_le_unit_set_margins($unit, $doc, 0, 0);
    }

    // Jika tidak ada baris kosong sebagai area TTD, beri ruang pada jabatan
    // terakhir agar tanda tangan tidak menempel ke teks jabatan.
    if (!$spacerUsed) {
        $lastJabatan = null;
        foreach ($items as $i) {
            if ($roles[$i]['sub'] === 'jabatan') {
                $lastJabatan = $i;
            }
        }
        if ($lastJabatan !== null) {
            spt_le_unit_set_margins($units[$lastJabatan], $doc, null, $spaceTwips);
        }
    }

    // Posisi horizontal blok TTD.
    $offsetX = spt_le_px_to_twips($settings['ttd_x'] ?? 0);
    if ($offsetX !== 0) {
        $tables = [];
        foreach ($items as $i) {
            if ($units[$i]['kind'] === 'row' && $units[$i]['table'] instanceof DOMElement) {
                $table = $units[$i]['table'];
                $tables[spl_object_id($table)] = $table;
            } else {
                $ind = spt_le_indent($units[$i]['node'], $doc);
                $current = $ind->hasAttributeNS(spt_le_ns(), 'left')
                    ? (int)$ind->getAttributeNS(spt_le_ns(), 'left')
                    : 0;
                $ind->setAttributeNS(spt_le_ns(), 'w:left', (string)max(0, $current + $offsetX));
            }
        }
        foreach ($tables as $table) {
            $tblPr = spt_le_child($table, 'tblPr');
            if (!$tblPr) {
                continue;
            }
            $tblInd = spt_le_child($tblPr, 'tblInd');
            if (!$tblInd) {
                $tblInd = spt_le_el($doc, 'tblInd');
                spt_le_insert_ordered($tblPr, $tblInd, [
                    'tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize',
                    'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd',
                    'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook',
                    'tblCaption', 'tblDescription', 'tblPrChange',
                ]);
            }
            $current = $tblInd->hasAttributeNS(spt_le_ns(), 'w')
                ? (int)$tblInd->getAttributeNS(spt_le_ns(), 'w')
                : 0;
            $tblInd->setAttributeNS(spt_le_ns(), 'w:w', (string)max(0, $current + $offsetX));
            $tblInd->setAttributeNS(spt_le_ns(), 'w:type', 'dxa');
        }
    }

    // Kunci blok TTD.
    if (!empty($settings['lock_ttd'])) {
        $ns = spt_le_ns();
        $count = count($items);
        foreach ($items as $position => $i) {
            $unit = $units[$i];
            $paragraphs = $unit['kind'] === 'row'
                ? iterator_to_array($unit['node']->getElementsByTagNameNS($ns, 'p'))
                : [$unit['node']];
            foreach ($paragraphs as $paragraph) {
                spt_le_add_keep_flag($paragraph, 'keepLines', $doc);
                if ($position < $count - 1 && $roles[$i]['sub'] !== 'nip') {
                    spt_le_add_keep_flag($paragraph, 'keepNext', $doc);
                }
            }
            if ($unit['kind'] === 'row') {
                $trPr = spt_le_child($unit['node'], 'trPr');
                if (!$trPr) {
                    $trPr = spt_le_el($doc, 'trPr');
                    if ($unit['node']->firstChild) {
                        $unit['node']->insertBefore($trPr, $unit['node']->firstChild);
                    } else {
                        $unit['node']->appendChild($trPr);
                    }
                }
                if (!spt_le_child($trPr, 'cantSplit')) {
                    $cantSplit = spt_le_el($doc, 'cantSplit');
                    spt_le_insert_ordered($trPr, $cantSplit, [
                        'cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore',
                        'wAfter', 'cantSplit', 'trHeight', 'tblHeader',
                        'tblCellSpacing', 'jc', 'hidden', 'ins', 'del', 'trPrChange',
                    ]);
                }
            }
        }
    }
}

function spt_le_set_table_alignment(DOMElement $table, DOMDocument $doc, $value)
{
    $tblPr = spt_le_child($table, 'tblPr');
    if (!$tblPr) {
        return;
    }
    $jc = spt_le_child($tblPr, 'jc');
    if (!$jc) {
        $jc = spt_le_el($doc, 'jc');
        spt_le_insert_ordered($tblPr, $jc, [
            'tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize',
            'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd',
            'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook',
            'tblCaption', 'tblDescription', 'tblPrChange',
        ]);
    }
    $jc->setAttributeNS(spt_le_ns(), 'w:val', $value);
}

/**
 * Jadikan halaman lampiran selalu mulai di halaman baru. Template lama
 * memakai deretan paragraf kosong sebagai "page break"; paragraf itu ikut
 * terhitung sebagai spacer TTD sehingga lampiran bisa tertarik ke halaman 1.
 */
function spt_le_isolate_lampiran(DOMDocument $doc)
{
    $body = null;
    $root = $doc->documentElement;
    if ($root) {
        if ($root->localName === 'body') {
            $body = $root;
        } else {
            foreach ($root->childNodes as $child) {
                if ($child instanceof DOMElement && $child->localName === 'body') {
                    $body = $child;
                    break;
                }
            }
        }
    }
    if (!$body) {
        return;
    }
    $lampiran = null;
    foreach ($body->childNodes as $child) {
        if ($child instanceof DOMElement
            && $child->localName === 'p'
            && preg_match('/^\s*Lampiran\s+(?:Surat\s+)?Tugas/iu', spt_le_unit_text($child))) {
            $lampiran = $child;
            break;
        }
    }
    if (!$lampiran) {
        return;
    }
    spt_le_add_keep_flag($lampiran, 'pageBreakBefore', $doc);

    $node = $lampiran->previousSibling;
    while ($node !== null) {
        $previous = $node->previousSibling;
        if ($node instanceof DOMElement) {
            if ($node->localName !== 'p') {
                break;
            }
            if (trim(spt_le_unit_text($node)) !== '' || spt_le_paragraph_has_image($node)) {
                break;
            }
            $body->removeChild($node);
        } elseif ($node->nodeType === XML_TEXT_NODE && trim((string)$node->nodeValue) === '') {
            $body->removeChild($node);
        } else {
            break;
        }
        $node = $previous;
    }
}

/**
 * Rapikan struktur dokumen sebelum pengaturan jarak diterapkan:
 *  - Tabel tanda tangan yang mengambang (tblpPr) dijadikan tabel inline
 *    rata kanan. Tabel mengambang dengan offset negatif menabrak tabel isi
 *    sehingga Word mendorong seluruh blok TTD ke halaman berikutnya.
 *  - Penanda lastRenderedPageBreak lama dibuang agar Word menghitung ulang.
 *  - Paragraf kosong berderet di akhir dokumen dipangkas; paragraf ini
 *    menambah tinggi halaman tanpa isi dan sering memicu halaman kedua.
 */
function spt_le_prepare_document(DOMDocument $doc)
{
    $ns = spt_le_ns();
    $xpath = new DOMXPath($doc);
    $xpath->registerNamespace('w', $ns);

    $anchorParagraphs = &spt_le_anchor_paragraphs();
    $anchorParagraphs = [];
    foreach (iterator_to_array($xpath->query('//w:tbl[w:tblPr/w:tblpPr]')) as $table) {
        $tblPr = spt_le_child($table, 'tblPr');
        if ($tblPr) {
            $tblpPr = spt_le_child($tblPr, 'tblpPr');
            if ($tblpPr) {
                $tblPr->removeChild($tblpPr);
            }
            spt_le_set_table_alignment($table, $doc, 'right');
        }

        // Tabel mengambang tidak memakan tinggi alur teks, sehingga paragraf
        // kosong bergambar anchor sesudahnya duduk di awal blok. Begitu tabel
        // dijadikan inline paragraf itu terdorong ke bawah tabel dan gambar
        // (mis. stempel) ikut turun. Pindahkan ke sebelum tabel agar posisinya
        // tetap relatif ke awal blok tanda tangan. Beberapa template menyisipkan
        // paragraf kosong tanpa anchor sebelum paragraf stempel, jadi pindai
        // beberapa paragraf kosong ke depan.
        $node = $table->nextSibling;
        $emptySeen = 0;
        while ($node !== null && $emptySeen <= 4) {
            if ($node->nodeType === XML_TEXT_NODE) {
                if (trim((string)$node->nodeValue) === '') {
                    $node = $node->nextSibling;
                    continue;
                }
                break;
            }
            if (!($node instanceof DOMElement) || $node->nodeName !== 'w:p') {
                break;
            }
            if (trim(spt_le_unit_text($node)) !== '') {
                break;
            }
            if (spt_le_has_stamp_anchor($node)) {
                $next = $node->nextSibling;
                $table->parentNode->insertBefore($node, $table);
                $anchorParagraphs[] = $node;
                $node = $next;
                continue;
            }
            if (spt_le_paragraph_has_image($node)) {
                // Gambar lain (bukan stempel) bukan milik blok TTD.
                break;
            }
            $emptySeen++;
            $node = $node->nextSibling;
        }
    }

    spt_le_isolate_lampiran($doc);

    foreach (iterator_to_array($xpath->query('//w:lastRenderedPageBreak')) as $marker) {
        if ($marker->parentNode) {
            $marker->parentNode->removeChild($marker);
        }
    }

    $body = $xpath->query('//w:body')->item(0);
    if (!$body) {
        return;
    }
    $blocks = [];
    foreach ($body->childNodes as $child) {
        if ($child->nodeType === 1) {
            $blocks[] = $child;
        }
    }
    $lastContent = null;
    foreach ($blocks as $block) {
        if ($block->nodeName === 'w:sectPr') {
            continue;
        }
        if ($block->nodeName === 'w:tbl') {
            $lastContent = $block;
            continue;
        }
        if (trim(spt_le_unit_text($block)) !== '') {
            $lastContent = $block;
        }
    }
    if ($lastContent === null) {
        return;
    }

    $after = false;
    $spacerKept = false;
    foreach ($blocks as $block) {
        if ($block === $lastContent) {
            $after = true;
            continue;
        }
        if (!$after || $block->nodeName === 'w:sectPr') {
            continue;
        }
        if ($block->nodeName === 'w:p' && !$spacerKept) {
            $spacerKept = true;
            spt_le_spacing($block, $doc)->setAttributeNS($ns, 'w:line', '20');
            spt_le_spacing($block, $doc)->setAttributeNS($ns, 'w:lineRule', 'exact');
            continue;
        }
        $body->removeChild($block);
    }
}

/**
 * Kembalikan paragraf anchor yang dipindah ke tinggi minimal. Gambar anchor
 * mengambang tidak terpotong oleh tinggi baris, jadi baris tidak perlu tinggi.
 */
function spt_le_settle_anchor_paragraphs(DOMDocument $doc)
{
    $anchors = &spt_le_anchor_paragraphs();
    $ns = spt_le_ns();
    foreach ($anchors as $paragraph) {
        if (!$paragraph->parentNode) {
            continue;
        }
        $spacing = spt_le_spacing($paragraph, $doc);
        $spacing->setAttributeNS($ns, 'w:before', '0');
        $spacing->setAttributeNS($ns, 'w:after', '0');
        $spacing->setAttributeNS($ns, 'w:line', '20');
        $spacing->setAttributeNS($ns, 'w:lineRule', 'exact');
    }
}

/**
 * Terapkan seluruh pengaturan layout ke DOM dokumen.
 * Return ringkasan blok yang terdeteksi (untuk diagnostik).
 */
function spt_le_apply(DOMDocument $doc, array $settings)
{
    spt_le_prepare_document($doc);

    $units = spt_le_collect_units($doc);
    $detected = spt_le_detect_roles($units);
    $roles = $detected['roles'];

    $baseHalfPoints = spt_le_dominant_font($units);
    $targetHalfPoints = (float)($settings['font_size'] ?? ($baseHalfPoints / 2)) * 2;
    $factor = $baseHalfPoints > 0 ? ($targetHalfPoints / $baseHalfPoints) : 1.0;

    $baseLine = spt_le_dominant_line($units);

    // Blok yang punya pengaturan jarak sendiri tidak ikut normalisasi baris:
    // kop (kop_line_spacing), judul/nomor (exact heading), TTD (ttd_*),
    // dan lampiran (tidak dimampatkan).
    $lineExclude = [];
    foreach ($roles as $i => $info) {
        if (in_array($info['role'], ['kop', 'judul', 'nomor', 'ttd', 'lampiran'], true)) {
            $lineExclude[$i] = true;
        }
    }

    spt_le_set_page_margins($doc, $settings);
    spt_le_scale_font_sizes($doc, $factor);
    spt_le_apply_line_height($units, $doc, $baseLine, (float)($settings['line_height'] ?? 1.0), $lineExclude);

    // Kop: kerapatan baris + jarak bawah.
    $kopLineTwips = spt_le_px_to_twips($settings['kop_line_spacing'] ?? 0);
    $kopBottomTwips = spt_le_px_to_twips($settings['kop_bottom_spacing'] ?? 10);
    $kopItems = spt_le_block_items($units, $roles, 'kop');
    $kopContent = [];
    foreach ($kopItems as $i) {
        if ($units[$i]['text'] === '') {
            spt_le_unit_collapse($units[$i], $doc);
            continue;
        }
        $kopContent[] = $i;
        if ($kopLineTwips > 0) {
            spt_le_unit_add_line($units[$i], $doc, $kopLineTwips);
        }
    }
    if (!empty($kopContent)) {
        $last = $kopContent[count($kopContent) - 1];
        spt_le_unit_set_margins($units[$last], $doc, null, $kopBottomTwips, true);
    }

    // Judul & nomor.
    $judulItems = spt_le_block_items($units, $roles, 'judul');
    $nomorItems = spt_le_block_items($units, $roles, 'nomor');
    $judulTopTwips = spt_le_px_to_twips(max(0, $settings['judul_top_spacing'] ?? 0));
    $nomorGapTwips = spt_le_px_to_twips($settings['nomor_surat_spacing'] ?? 4);
    $judulBottomTwips = spt_le_px_to_twips($settings['judul_bottom_spacing'] ?? 10);

    $headingMultiplier = max(1.3, (float)($settings['line_height'] ?? 1.0));
    $sameHeading = !empty($judulItems) && !empty($nomorItems) && $judulItems === $nomorItems;
    foreach ([['judul', $judulItems, $judulTopTwips, 0], ['nomor', $nomorItems, $nomorGapTwips, $judulBottomTwips]] as $heading) {
        list($roleName, $items, $top, $bottom) = $heading;
        $content = [];
        foreach ($items as $i) {
            if ($units[$i]['text'] === '') {
                spt_le_unit_collapse($units[$i], $doc);
                continue;
            }
            $content[] = $i;
        }
        foreach ($content as $position => $i) {
            $unitTop = $position === 0 ? $top : 0;
            $unitBottom = $position === count($content) - 1 ? $bottom : 0;
            spt_le_unit_set_margins($units[$i], $doc, $unitTop, $unitBottom);
        }
        foreach ($content as $i) {
            if ($units[$i]['kind'] === 'row') {
                // Tinggi baris judul/nomor tidak boleh mengunci jarak: biarkan isi menentukan.
                spt_le_row_ensure_height($units[$i]['node'], $doc, 0, 'atLeast');
            }
            $paragraphs = $units[$i]['kind'] === 'row'
                ? $units[$i]['node']->getElementsByTagNameNS(spt_le_ns(), 'p')
                : [$units[$i]['node']];
            foreach ($paragraphs as $paragraph) {
                if (spt_le_paragraph_is_blank($paragraph)) {
                    spt_le_collapse_paragraph($paragraph, $doc);
                    continue;
                }
                $spacing = spt_le_spacing($paragraph, $doc);
                // Hanya ukuran pada run berisi teks; w:sz di w:pPr/w:rPr milik penanda paragraf.
                $maxSize = 0;
                foreach ($paragraph->getElementsByTagNameNS(spt_le_ns(), 'r') as $run) {
                    if ($run->getElementsByTagNameNS(spt_le_ns(), 't')->length === 0) {
                        continue;
                    }
                    $rPr = spt_le_child($run, 'rPr');
                    $sz = $rPr ? spt_le_child($rPr, 'sz') : null;
                    if ($sz) {
                        $maxSize = max($maxSize, (int)$sz->getAttributeNS(spt_le_ns(), 'val'));
                    }
                }
                if ($maxSize <= 0) {
                    $maxSize = max(24, (int)round($targetHalfPoints));
                }
                // lineRule=exact: w:line dibaca sebagai twips, bukan 240-an baris.
                $lineTwips = (int)round(($maxSize / 2) * $headingMultiplier * 20);
                if ($sameHeading) {
                    $lineTwips += $nomorGapTwips;
                }
                $spacing->setAttributeNS(spt_le_ns(), 'w:line', (string)$lineTwips);
                $spacing->setAttributeNS(spt_le_ns(), 'w:lineRule', 'exact');
            }
        }
    }

    spt_le_apply_block($units, $roles, $doc, 'dasar',
        $settings['dasar_top_spacing'] ?? 6,
        $settings['dasar_item_spacing'] ?? 3,
        $settings['dasar_bottom_spacing'] ?? 8);

    $memItems = spt_le_block_items($units, $roles, 'memerintahkan');
    if (!empty($memItems)) {
        $memTwips = spt_le_px_to_twips($settings['memerintahkan_spacing'] ?? 8);
        foreach ($memItems as $i) {
            if ($units[$i]['text'] === '') {
                spt_le_unit_collapse($units[$i], $doc);
                continue;
            }
            spt_le_unit_set_margins($units[$i], $doc, $memTwips, $memTwips);
        }
    }

    spt_le_apply_pegawai($units, $roles, $doc, $settings);

    spt_le_apply_block($units, $roles, $doc, 'untuk',
        $settings['untuk_top_spacing'] ?? 6,
        $settings['untuk_item_spacing'] ?? 3,
        $settings['untuk_bottom_spacing'] ?? 10);

    spt_le_apply_ttd($units, $roles, $doc, $settings);

    // Lampiran tidak dimampatkan, tetapi sel tabel tetap rata atas.
    spt_le_align_tables_top($doc, $units, $roles);
    spt_le_settle_anchor_paragraphs($doc);

    return [
        'units' => count($units),
        'roles' => array_count_values(array_map(static function ($info) {
            return $info['role'];
        }, $roles)),
        'anchors' => $detected['anchors'],
        'font_factor' => round($factor, 4),
        'base_font' => $baseHalfPoints / 2,
    ];
}

/**
 * Normalisasi spasi warisan style (docDefaults). Template sering menyimpan
 * w:after=160/200 di sini; tanpa dinetralkan, paragraf yang tidak punya
 * <w:spacing> sendiri tetap diberi jarak liar.
 */
function spt_le_neutralize_styles(DOMDocument $styles)
{
    $ns = spt_le_ns();
    $xpath = new DOMXPath($styles);
    $xpath->registerNamespace('w', $ns);
    foreach ($xpath->query('//w:docDefaults//w:pPr//w:spacing') as $spacing) {
        $spacing->setAttributeNS($ns, 'w:before', '0');
        $spacing->setAttributeNS($ns, 'w:after', '0');
    }
}

/**
 * Apakah dokumen punya blok "Lampiran Surat Tugas" (halaman lampiran).
 * Dipakai solver fit untuk menentukan target halaman: 1 tanpa lampiran,
 * 2 dengan lampiran.
 */
function spt_le_docx_has_lampiran($docxPath)
{
    if (!class_exists('ZipArchive') || !is_file($docxPath)) {
        return false;
    }
    $zip = new ZipArchive();
    if ($zip->open($docxPath) !== true) {
        return false;
    }
    $documentXml = $zip->getFromName('word/document.xml');
    $zip->close();
    $doc = spt_le_load_dom($documentXml);
    if (!$doc) {
        return false;
    }
    // Deteksi langsung dari teks: tidak bergantung pada baris tanggal TTD
    // yang baru terisi setelah placeholder diganti.
    foreach (spt_le_collect_units($doc) as $unit) {
        if ($unit['text'] !== ''
            && preg_match('/^\s*Lampiran\s+(?:Surat\s+)?Tugas/iu', $unit['text'])) {
            return true;
        }
    }
    return false;
}

/**
 * Terapkan layout langsung ke file .docx (in-place).
 * Return array ringkasan jika berhasil, false jika gagal.
 */
function spt_le_apply_docx($docxPath, array $settings)
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
        $doc = spt_le_load_dom($documentXml);
        if (!$doc) {
            $zip->close();
            return false;
        }
        $report = spt_le_apply($doc, $settings);

        $styles = $stylesXml !== false ? spt_le_load_dom($stylesXml) : null;
        if ($styles) {
            spt_le_scale_font_sizes($styles, (float)($report['font_factor'] ?? 1.0));
            spt_le_neutralize_styles($styles);
        }

        $zip->addFromString('word/document.xml', $doc->saveXML());
        if ($styles) {
            $zip->addFromString('word/styles.xml', $styles->saveXML());
        }
        $zip->close();
        return $report;
    } catch (Throwable $e) {
        error_log('spt_le_apply_docx: ' . $e->getMessage());
        try {
            $zip->close();
        } catch (Throwable $e2) {
            // abaikan
        }
        return false;
    }
}

/**
 * Kompres pengaturan secara proporsional. Dipakai solver "Muat 1 Halaman".
 * $ratio 0.9 berarti menyusutkan hampir semua jarak dan huruf sekitar 10%.
 */
function spt_le_compress(array $settings, $ratio)
{
    $r = max(0.72, min(1.0, (float)$ratio));
    // t = 1 (tanpa perubahan) .. 0 (batas paling rapat).
    $t = ($r - 0.72) / 0.28;
    $out = $settings;

    $fontMin = 9.0;
    $out['font_size'] = round(($fontMin + (((float)$settings['font_size'] - $fontMin) * $t)) * 2) / 2;
    $out['line_height'] = round((1.0 + (((float)$settings['line_height'] - 1.0) * $t)) * 20) / 20;
    foreach (['margin_top' => 8, 'margin_bottom' => 8, 'margin_left' => 10, 'margin_right' => 10] as $side => $min) {
        $value = max($min, (float)$settings[$side]);
        $out[$side] = round(($min + (($value - $min) * $t)) * 2) / 2;
    }

    $spacingKeys = [
        'kop_line_spacing', 'kop_bottom_spacing', 'judul_top_spacing',
        'nomor_surat_spacing', 'judul_bottom_spacing', 'dasar_top_spacing',
        'dasar_item_spacing', 'dasar_bottom_spacing', 'memerintahkan_spacing',
        'pegawai_top_spacing', 'pegawai_gap_spacing', 'pegawai_line_spacing',
        'untuk_top_spacing', 'untuk_item_spacing', 'untuk_bottom_spacing',
        'ttd_offset', 'ttd_date_to_jabatan', 'ttd_space_height',
        'ttd_name_nip_spacing',
    ];
    foreach ($spacingKeys as $key) {
        if (!isset($out[$key])) {
            continue;
        }
        $out[$key] = max(0, round((float)$out[$key] * $t));
    }
    // Ruang tanda tangan tidak boleh ikut menyusut berlebihan.
    $out['ttd_space_height'] = max(40, (float)($out['ttd_space_height'] ?? 60));
    $out['ttd_x'] = (float)($settings['ttd_x'] ?? 0);
    $out['lock_ttd'] = !empty($settings['lock_ttd']) ? 1 : 0;
    return $out;
}
