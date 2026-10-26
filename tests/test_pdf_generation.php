<?php
require_once __DIR__ . '/../config/database.php';

$stmt = $pdo->query("SELECT id, nomor_surat FROM surat_tugas ORDER BY id DESC LIMIT 1");
$surat = $stmt->fetch(PDO::FETCH_ASSOC);

if (!$surat) {
    echo "Tidak ada surat_tugas.\n";
    exit(1);
}

$id = $surat['id'];
$urlBase = "http://127.0.0.1:8000/modules/generate-surat.php?id=" . $id . "&format=pdf&inline=1";
$urlNoLayout = $urlBase;
$urlWithLayout = $urlBase . "&layout=1&margin_top=35&kop_bottom_spacing=50&ttd_space_height=120";

$ctx = stream_context_create([
    'http' => [
        'timeout' => 30,
        'ignore_errors' => true,
        'header' => "Host: spt-php.test\r\n"
    ]
]);

echo "1. Mengambil PDF tanpa layout kustom...\n";
$pdf1 = @file_get_contents($urlNoLayout, false, $ctx);
$h1 = $http_response_header ?? [];
echo "Status: " . ($h1[0] ?? 'NO RESPONSE') . ", Size: " . (is_string($pdf1) ? strlen($pdf1) : 0) . " bytes\n";

echo "2. Mengambil PDF DENGAN layout kustom ekstrim (margin_top=35, kop_bottom=50, ttd_space=120)...\n";
$pdf2 = @file_get_contents($urlWithLayout, false, $ctx);
$h2 = $http_response_header ?? [];
echo "Status: " . ($h2[0] ?? 'NO RESPONSE') . ", Size: " . (is_string($pdf2) ? strlen($pdf2) : 0) . " bytes\n";

if ($pdf1 && $pdf2 && strpos($h1[0] ?? '', '200') !== false) {
    if ($pdf1 === $pdf2) {
        echo "PERINGATAN / MASALAH: File PDF 1 dan PDF 2 IDENTIK PERSIS! Layout TIDAK berpengaruh pada output PDF!\n";
    } else {
        echo "INFO: File PDF 1 dan PDF 2 berbeda (Ukuran 1: " . strlen($pdf1) . ", Ukuran 2: " . strlen($pdf2) . ")\n";
    }
    file_put_contents(__DIR__ . '/pdf_no_layout.pdf', $pdf1);
    file_put_contents(__DIR__ . '/pdf_with_layout.pdf', $pdf2);
} else {
    echo "Gagal mengambil salah satu atau kedua PDF.\n";
    print_r($h1);
    print_r($h2);
    if ($pdf1) echo "Body 1: " . substr($pdf1, 0, 500) . "\n";
}
