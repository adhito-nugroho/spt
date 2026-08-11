<?php
/**
 * Diagnostik: cek apakah Microsoft Word COM bisa digunakan untuk konversi docx → PDF.
 * Buka di browser: http://localhost/spt-php/modules/debug_word_com.php
 * Syarat: Windows, PHP dengan extension COM (php_com_dotnet.dll), Microsoft Word terinstall.
 */
header('Content-Type: text/html; charset=utf-8');
echo "<!DOCTYPE html><html><head><meta charset='utf-8'><title>Debug Word COM</title>";
echo "<style>body{font-family:Consolas,monospace;padding:20px;background:#1e1e1e;color:#d4d4d4;}";
echo "h1{color:#4ec9b0;} .ok{color:#4ec9b0;} .fail{color:#f48771;} .info{color:#9cdcfe;} pre{background:#252526;padding:12px;border-radius:4px;}</style></head><body>";

echo "<h1>Diagnostik Word COM</h1>";

$results = [];
$word = null;

// 1. Cek extension COM
$comExists = class_exists('COM');
$results['PHP COM class'] = $comExists ? 'BERHASIL (class COM ada)' : 'GAGAL (class COM tidak ada)';

// 2. Cek php_com_dotnet
$dotnetLoaded = extension_loaded('com_dotnet');
$results['Extension php_com_dotnet'] = $dotnetLoaded ? 'Aktif' : 'Tidak aktif (uncomment extension=php_com_dotnet.dll di php.ini)';

// 3. Versi PHP & OS
$results['PHP Version'] = PHP_VERSION;
$results['OS'] = PHP_OS_FAMILY . ' / ' . PHP_OS;
$results['Server API'] = php_sapi_name();

if ($comExists) {
    try {
        $word = new COM('Word.Application');
        $results['Instansiasi Word.Application'] = 'BERHASIL';
        $word->Visible = false;
        $word->DisplayAlerts = 0;
        $results['Word Version'] = isset($word->Version) ? (string)$word->Version : 'tidak terbaca';
        $results['Word Build'] = isset($word->Build) ? (string)$word->Build : '-';
    } catch (Throwable $e) {
        $results['Instansiasi Word.Application'] = 'GAGAL';
        $results['Error'] = $e->getMessage();
    } finally {
        if ($word !== null) {
            try {
                $word->Quit();
            } catch (Throwable $e) {
                $results['Quit Error'] = $e->getMessage();
            }
            $word = null;
        }
    }
} else {
    $results['Instansiasi Word.Application'] = 'Dilewati (COM tidak ada)';
}

// Tampilkan hasil
echo "<h2>Hasil</h2>";
echo "<pre>";
foreach ($results as $k => $v) {
    $cls = (strpos($v, 'GAGAL') !== false || strpos($v, 'Tidak') !== false) ? 'fail' : 'ok';
    echo "<span class='$cls'>" . htmlspecialchars($k) . ": " . htmlspecialchars($v) . "</span>\n";
}
echo "</pre>";

$allOk = $comExists && isset($results['Instansiasi Word.Application']) && strpos($results['Instansiasi Word.Application'], 'BERHASIL') !== false;
echo "<p class='" . ($allOk ? 'ok' : 'fail') . "'><strong>" . ($allOk ? 'BERHASIL: Word COM siap dipakai untuk konversi PDF.' : 'GAGAL: Word COM tidak siap.') . "</strong></p>";

echo "<h2>Instruksi mengaktifkan COM</h2>";
echo "<ul class='info'>";
echo "<li>Di <strong>php.ini</strong>: pastikan baris <code>extension=php_com_dotnet</code> tidak dikomentari (tanpa ; di depan).</li>";
echo "<li>Restart web server (Apache/IIS) setelah mengubah php.ini.</li>";
echo "<li>PHP harus berjalan sebagai user yang punya akses ke Microsoft Word (biasanya user yang login atau Application Pool identity).</li>";
echo "<li>IIS: Application Pool → Identity = account yang punya Word; atau set Load User Profile = true.</li>";
echo "<li>Laragon/Apache: jalankan sebagai user yang bisa buka Word (bukan SYSTEM).</li>";
echo "<li>Jika perlu, tambah DCOM permission: DCOM Config → Microsoft Word → Security → Launch and Activation.</li>";
echo "</ul>";

echo "</body></html>";
