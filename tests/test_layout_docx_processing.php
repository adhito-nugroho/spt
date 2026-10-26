<?php
define('SPT_TEST_MODE', true);
require_once __DIR__ . '/../modules/generate-surat.php';

$sourceDocx = __DIR__ . '/../assets/templates/1766636680_template_satu_orang.docx';
if (!file_exists($sourceDocx)) {
    echo "SKIP: Template not found\n";
    exit(0);
}

$tmpDocx = sys_get_temp_dir() . '/test_layout_' . uniqid() . '.docx';
if (!copy($sourceDocx, $tmpDocx)) {
    fwrite(STDERR, "FAIL: Unable to copy source docx\n");
    exit(1);
}

$settings = spt_layout_default_values();
$settings['kop_bottom_spacing'] = 25;
$settings['judul_top_spacing'] = 15;
$settings['nomor_surat_spacing'] = 8;
$settings['judul_bottom_spacing'] = 14;
$settings['dasar_top_spacing'] = 10;
$settings['dasar_item_spacing'] = 5;
$settings['dasar_bottom_spacing'] = 12;
$settings['memerintahkan_spacing'] = 12;
$settings['pegawai_top_spacing'] = 10;
$settings['pegawai_gap_spacing'] = 18;
$settings['pegawai_line_spacing'] = 4;
$settings['untuk_top_spacing'] = 10;
$settings['untuk_item_spacing'] = 6;
$settings['untuk_bottom_spacing'] = 15;
$settings['ttd_offset'] = 20;
$settings['ttd_date_to_jabatan'] = 6;
$settings['ttd_space_height'] = 75;
$settings['ttd_name_nip_spacing'] = 8;
$settings['lock_ttd'] = 1;

$result = spt_layout_apply_to_docx($tmpDocx, $settings);
if (!$result) {
    fwrite(STDERR, "FAIL: spt_layout_apply_to_docx returned false\n");
    @unlink($tmpDocx);
    exit(1);
}

// Verify that tmpDocx can still be opened and parsed as XML
$zip = new ZipArchive();
if ($zip->open($tmpDocx) !== true) {
    fwrite(STDERR, "FAIL: Resulting DOCX is not a valid zip\n");
    @unlink($tmpDocx);
    exit(1);
}
$xml = $zip->getFromName('word/document.xml');
$zip->close();
@unlink($tmpDocx);

$dom = new DOMDocument();
if (!$dom->loadXML($xml)) {
    fwrite(STDERR, "FAIL: Resulting document.xml is invalid XML\n");
    exit(1);
}

// Verify that spacing attributes exist
$spacings = $dom->getElementsByTagName('spacing');
if ($spacings->length === 0) {
    fwrite(STDERR, "FAIL: No spacing elements found in XML\n");
    exit(1);
}

echo "PASS: DOCX per-block layout applied successfully.\n";
