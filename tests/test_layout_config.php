<?php
define('SPT_TEST_MODE', true);
require_once __DIR__ . '/../modules/generate-surat.php';

$defaults = spt_layout_default_values();
$ranges = spt_layout_ranges();

$expectedKeys = [
    'margin_top', 'margin_bottom', 'margin_left', 'margin_right',
    'font_size', 'line_height',
    'kop_line_spacing', 'kop_bottom_spacing',
    'judul_top_spacing', 'nomor_surat_spacing', 'judul_bottom_spacing',
    'dasar_top_spacing', 'dasar_item_spacing', 'dasar_bottom_spacing',
    'memerintahkan_spacing',
    'pegawai_top_spacing', 'pegawai_gap_spacing', 'pegawai_line_spacing',
    'untuk_top_spacing', 'untuk_item_spacing', 'untuk_bottom_spacing',
    'ttd_offset', 'ttd_x', 'ttd_date_to_jabatan', 'ttd_space_height', 'ttd_name_nip_spacing',
    'lock_ttd'
];

foreach ($expectedKeys as $key) {
    if (!array_key_exists($key, $defaults)) {
        fwrite(STDERR, "FAIL: Key '$key' not found in defaults\n");
        exit(1);
    }
    if ($key !== 'lock_ttd' && !array_key_exists($key, $ranges)) {
        fwrite(STDERR, "FAIL: Key '$key' not found in ranges\n");
        exit(1);
    }
}

// Test clamping and normalization
$mockRaw = [
    'kop_bottom_spacing' => 999, // Should clamp to 50
    'judul_top_spacing' => -999, // Should clamp to -15
    'pegawai_gap_spacing' => 20,
    'lock_ttd' => 1
];
$norm = spt_layout_normalize_stored_settings($mockRaw);
if ($norm['kop_bottom_spacing'] != 50 || $norm['judul_top_spacing'] != -15 || $norm['pegawai_gap_spacing'] != 20 || $norm['lock_ttd'] !== 1) {
    fwrite(STDERR, "FAIL: Normalization / clamping unexpected result\n");
    exit(1);
}

// Jarak judul ke nomor tidak boleh turun ke 0 agar teks tidak bertumpuk.
if ($ranges['nomor_surat_spacing'][0] < 2) {
    fwrite(STDERR, "FAIL: nomor_surat_spacing minimum must be at least 2\n");
    exit(1);
}

// Pembersihan kata berulang dan typo kop surat.
if (spt_clean_repeated_words('Peraturan tentang tentang Anggaran') !== 'Peraturan tentang Anggaran') {
    fwrite(STDERR, "FAIL: repeated word cleanup failed\n");
    exit(1);
}
if (spt_fix_common_typos('Bojongeoro Telepn') !== 'Bojonegoro Telepon') {
    fwrite(STDERR, "FAIL: typo cleanup failed\n");
    exit(1);
}

echo "PASS: Layout config test passed.\n";
