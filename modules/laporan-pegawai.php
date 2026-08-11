<?php
require_once '../vendor/autoload.php';
use Dompdf\Dompdf;
use Dompdf\Options;

// Set timezone ke Indonesia (WIB)
date_default_timezone_set('Asia/Jakarta');

$pdo = require_once '../config/database.php';

// Inisialisasi variabel filter
$start_date = isset($_GET['start_date']) ? $_GET['start_date'] : date('Y-m-01');
$end_date = isset($_GET['end_date']) ? $_GET['end_date'] : date('Y-m-d');
$filter_nip = isset($_GET['nip']) ? $_GET['nip'] : '';
$export = isset($_GET['export']) ? strtolower(trim($_GET['export'])) : '';
$laporan = [];

// Pagination (hanya untuk tampilan daftar per-pegawai)
$limit = 5; // jumlah pegawai per halaman (karena tiap pegawai berisi banyak baris surat)
$page = isset($_GET['page']) ? (int)$_GET['page'] : 1;
if ($page < 1) $page = 1;
$start = ($page - 1) * $limit;
$total_records = 0;
$total_pages = 0;

// Ambil daftar pegawai untuk dropdown filter
$stmt_pegawai = $pdo->query("SELECT nip, nama FROM pegawai ORDER BY nama");
$daftar_pegawai = $stmt_pegawai->fetchAll();

// Query untuk mengambil data laporan per pegawai
$query_conditions = [];
$query_params = [];

// Filter periode
if ($start_date && $end_date) {
    $query_conditions[] = "st.tanggal_mulai BETWEEN :start_date AND :end_date";
    $query_params[':start_date'] = $start_date;
    $query_params[':end_date'] = $end_date;
}

// Filter pegawai
if (!empty($filter_nip)) {
    $query_conditions[] = "p.nip = :nip";
    $query_params[':nip'] = $filter_nip;
}

$where_clause = !empty($query_conditions) ? "WHERE " . implode(" AND ", $query_conditions) : "";

if ($start_date && $end_date) {
    if ($export === 'excel' || $export === 'pdf') {
        // Untuk export atau PDF: ambil semua data sesuai filter periode (tanpa pagination)
        $query = "SELECT
            p.nip,
            p.nama,
            p.pangkat,
            p.jabatan,
            st.id as id_surat,
            st.nomor_surat,
            st.tanggal_surat,
            st.untuk,
            st.dasar_surat,
            st.tanggal_mulai,
            st.tanggal_selesai,
            DATEDIFF(st.tanggal_selesai, st.tanggal_mulai) + 1 as durasi_hari
        FROM pegawai p
        INNER JOIN pegawai_tugas pt ON p.nip = pt.nip
        INNER JOIN surat_tugas st ON pt.id_surat_tugas = st.id
        $where_clause
        ORDER BY p.nama ASC, st.tanggal_mulai DESC";

        $stmt = $pdo->prepare($query);
        foreach ($query_params as $key => $value) {
            $stmt->bindValue($key, $value);
        }
        $stmt->execute();
        $raw_data = $stmt->fetchAll();

        // Kelompokkan data per pegawai
        $laporan_per_pegawai = [];
        foreach ($raw_data as $row) {
            $nip = $row['nip'];
            if (!isset($laporan_per_pegawai[$nip])) {
                $laporan_per_pegawai[$nip] = [
                    'nip' => $row['nip'],
                    'nama' => $row['nama'],
                    'pangkat' => $row['pangkat'],
                    'jabatan' => $row['jabatan'],
                    'total_tugas' => 0,
                    'total_hari' => 0,
                    'surat_tugas' => []
                ];
            }

            $laporan_per_pegawai[$nip]['surat_tugas'][] = [
                'id' => $row['id_surat'],
                'nomor_surat' => $row['nomor_surat'],
                'tanggal_surat' => $row['tanggal_surat'],
                'untuk' => $row['untuk'],
                'dasar_surat' => $row['dasar_surat'],
                'tanggal_mulai' => $row['tanggal_mulai'],
                'tanggal_selesai' => $row['tanggal_selesai'],
                'durasi_hari' => $row['durasi_hari']
            ];

            $laporan_per_pegawai[$nip]['total_tugas']++;
            $laporan_per_pegawai[$nip]['total_hari'] += $row['durasi_hari'];
        }

        // Konversi ke array untuk memudahkan iterasi
        $laporan = array_values($laporan_per_pegawai);
    } else {
        // Untuk tampilan: paginasi per-pegawai (supaya 1 pegawai tidak terpecah antar halaman)
        $query_count = "SELECT COUNT(DISTINCT p.nip) AS total
            FROM pegawai p
            INNER JOIN pegawai_tugas pt ON p.nip = pt.nip
            INNER JOIN surat_tugas st ON pt.id_surat_tugas = st.id
            $where_clause";

        $stmt_count = $pdo->prepare($query_count);
        foreach ($query_params as $key => $value) {
            $stmt_count->bindValue($key, $value);
        }
        $stmt_count->execute();
        $total_records = (int)($stmt_count->fetch()['total'] ?? 0);
        $total_pages = (int)ceil($total_records / $limit);

        // Jika user meminta page di luar range, paksa ke halaman terakhir
        if ($total_pages > 0 && $page > $total_pages) {
            $page = $total_pages;
            $start = ($page - 1) * $limit;
        }

        $query_pegawai = "SELECT DISTINCT
                p.nip,
                p.nama,
                p.pangkat,
                p.jabatan
            FROM pegawai p
            INNER JOIN pegawai_tugas pt ON p.nip = pt.nip
            INNER JOIN surat_tugas st ON pt.id_surat_tugas = st.id
            $where_clause
            ORDER BY p.nama ASC
            LIMIT :start, :limit";

        $stmt_pegawai_page = $pdo->prepare($query_pegawai);
        foreach ($query_params as $key => $value) {
            $stmt_pegawai_page->bindValue($key, $value);
        }
        $stmt_pegawai_page->bindValue(':start', $start, PDO::PARAM_INT);
        $stmt_pegawai_page->bindValue(':limit', $limit, PDO::PARAM_INT);
        $stmt_pegawai_page->execute();
        $pegawai_page = $stmt_pegawai_page->fetchAll();

        if (!empty($pegawai_page)) {
            $nip_list = [];
            foreach ($pegawai_page as $row) {
                $nip_list[] = $row['nip'];
            }
            $nip_placeholders = [];
            $nip_params = [];
            foreach ($nip_list as $idx => $nip) {
                $key = ':nip' . $idx;
                $nip_placeholders[] = $key;
                $nip_params[$key] = $nip;
            }

            $query_detail = "SELECT
                    p.nip,
                    p.nama,
                    p.pangkat,
                    p.jabatan,
                    st.id as id_surat,
                    st.nomor_surat,
                    st.tanggal_surat,
                    st.untuk,
                    st.dasar_surat,
                    st.tanggal_mulai,
                    st.tanggal_selesai,
                    DATEDIFF(st.tanggal_selesai, st.tanggal_mulai) + 1 as durasi_hari
                FROM pegawai p
                INNER JOIN pegawai_tugas pt ON p.nip = pt.nip
                INNER JOIN surat_tugas st ON pt.id_surat_tugas = st.id
                WHERE st.tanggal_mulai BETWEEN :start_date AND :end_date
                    AND p.nip IN (" . implode(',', $nip_placeholders) . ")
                ORDER BY p.nama ASC, st.tanggal_mulai DESC";

            $stmt_detail = $pdo->prepare($query_detail);
            $stmt_detail->bindValue(':start_date', $start_date);
            $stmt_detail->bindValue(':end_date', $end_date);
            foreach ($nip_params as $key => $value) {
                $stmt_detail->bindValue($key, $value);
            }
            $stmt_detail->execute();
            $raw_data = $stmt_detail->fetchAll();

            // Kelompokkan data per pegawai
            $laporan_per_pegawai = [];
            foreach ($raw_data as $row) {
                $nip = $row['nip'];
                if (!isset($laporan_per_pegawai[$nip])) {
                    $laporan_per_pegawai[$nip] = [
                        'nip' => $row['nip'],
                        'nama' => $row['nama'],
                        'pangkat' => $row['pangkat'],
                        'jabatan' => $row['jabatan'],
                        'total_tugas' => 0,
                        'total_hari' => 0,
                        'surat_tugas' => []
                    ];
                }

                $laporan_per_pegawai[$nip]['surat_tugas'][] = [
                    'id' => $row['id_surat'],
                    'nomor_surat' => $row['nomor_surat'],
                    'tanggal_surat' => $row['tanggal_surat'],
                    'untuk' => $row['untuk'],
                    'dasar_surat' => $row['dasar_surat'],
                    'tanggal_mulai' => $row['tanggal_mulai'],
                    'tanggal_selesai' => $row['tanggal_selesai'],
                    'durasi_hari' => $row['durasi_hari']
                ];

                $laporan_per_pegawai[$nip]['total_tugas']++;
                $laporan_per_pegawai[$nip]['total_hari'] += $row['durasi_hari'];
            }

            // Konversi ke array untuk memudahkan iterasi
            $laporan = array_values($laporan_per_pegawai);
        }
    }
}

// Export Excel berdasarkan data/filter yang sedang ditampilkan
if ($export === 'excel') {
    $suffixPegawai = !empty($filter_nip) ? '_' . $filter_nip : '_semua_pegawai';
    $filename = 'laporan_per_pegawai_' . $start_date . '_sd_' . $end_date . $suffixPegawai . '.xls';
    $filename = preg_replace('/[^a-zA-Z0-9_\-\.]/', '_', $filename);

    header('Content-Type: application/vnd.ms-excel; charset=utf-8');
    header('Content-Disposition: attachment; filename="' . $filename . '"');
    header('Pragma: no-cache');
    header('Expires: 0');

    echo '<table border="1">';
    echo '<tr><th colspan="10">Laporan Surat Tugas Per Pegawai</th></tr>';
    echo '<tr><th colspan="10">Periode: ' .
        htmlspecialchars(date('d/m/Y', strtotime($start_date)), ENT_QUOTES, 'UTF-8') .
        ' s/d ' .
        htmlspecialchars(date('d/m/Y', strtotime($end_date)), ENT_QUOTES, 'UTF-8') .
        '</th></tr>';
    if (!empty($filter_nip) && isset($laporan[0]['nama'])) {
        echo '<tr><th colspan="10">Pegawai: ' . htmlspecialchars($laporan[0]['nama'], ENT_QUOTES, 'UTF-8') . '</th></tr>';
    }
    echo '<tr>';
    echo '<th>No</th>';
    echo '<th>NIP</th>';
    echo '<th>Nama</th>';
    echo '<th>Pangkat</th>';
    echo '<th>Jabatan</th>';
    echo '<th>Nomor Surat</th>';
    echo '<th>Tanggal Surat</th>';
    echo '<th>Tujuan / Maksud</th>';
    echo '<th>Pelaksanaan</th>';
    echo '<th>Durasi (hari)</th>';
    echo '</tr>';

    if (count($laporan) > 0) {
        $no = 1;
        foreach ($laporan as $pegawai) {
            foreach ($pegawai['surat_tugas'] as $surat) {
                $pelaksanaan = date('d/m/Y', strtotime($surat['tanggal_mulai']));
                if ($surat['tanggal_mulai'] != $surat['tanggal_selesai']) {
                    $pelaksanaan .= ' - ' . date('d/m/Y', strtotime($surat['tanggal_selesai']));
                }

                echo '<tr>';
                echo '<td>' . $no++ . '</td>';
                echo '<td>' . htmlspecialchars($pegawai['nip'], ENT_QUOTES, 'UTF-8') . '</td>';
                echo '<td>' . htmlspecialchars($pegawai['nama'], ENT_QUOTES, 'UTF-8') . '</td>';
                echo '<td>' . htmlspecialchars($pegawai['pangkat'], ENT_QUOTES, 'UTF-8') . '</td>';
                echo '<td>' . htmlspecialchars($pegawai['jabatan'], ENT_QUOTES, 'UTF-8') . '</td>';
                echo '<td>' . htmlspecialchars($surat['nomor_surat'], ENT_QUOTES, 'UTF-8') . '</td>';
                echo '<td>' . htmlspecialchars(date('d/m/Y', strtotime($surat['tanggal_surat'])), ENT_QUOTES, 'UTF-8') . '</td>';
                echo '<td>' . htmlspecialchars($surat['untuk'], ENT_QUOTES, 'UTF-8') . '</td>';
                echo '<td>' . htmlspecialchars($pelaksanaan, ENT_QUOTES, 'UTF-8') . '</td>';
                echo '<td>' . htmlspecialchars((string)$surat['durasi_hari'], ENT_QUOTES, 'UTF-8') . '</td>';
                echo '</tr>';
            }
        }
    } else {
        echo '<tr><td colspan="10">Tidak ada data surat tugas dengan filter yang dipilih.</td></tr>';
    }
    echo '</table>';
    exit;
}

// Export PDF berdasarkan data/filter yang sedang ditampilkan
if ($export === 'pdf') {
    ob_start();
    ?>
    <!DOCTYPE html>
    <html lang="id">
    <head>
        <meta charset="UTF-8">
        <title>Laporan Surat Tugas Per Pegawai</title>
        <style>
            body {
                font-family: 'Helvetica', Arial, sans-serif;
                font-size: 11px;
                line-height: 1.4;
                color: #333;
                margin: 0;
                padding: 0;
            }
            h1 {
                text-align: center;
                font-size: 16px;
                margin-bottom: 2px;
                text-transform: uppercase;
            }
            .subtitle {
                text-align: center;
                font-size: 11px;
                margin-bottom: 20px;
                color: #555;
            }
            .employee-card {
                border: 1px solid #94a3b8;
                margin-bottom: 20px;
                page-break-inside: avoid;
            }
            .employee-header {
                background-color: #4f46e5;
                color: #ffffff;
                padding: 10px;
            }
            .employee-name {
                font-size: 14px;
                font-weight: bold;
                margin: 0 0 5px 0;
            }
            .employee-info {
                font-size: 10px;
                margin: 0;
            }
            table {
                width: 100%;
                border-collapse: collapse;
            }
            th {
                background-color: #f3f4f6;
                border: 1px solid #000;
                padding: 6px;
                font-weight: bold;
                text-align: left;
                text-transform: uppercase;
                font-size: 9px;
            }
            td {
                border: 1px solid #000;
                padding: 6px;
                vertical-align: top;
            }
            .text-center { text-align: center; }
            .text-nowrap { white-space: nowrap; }
            .badge {
                background-color: #e0e7ff;
                color: #3730a3;
                padding: 2px 6px;
                border-radius: 4px;
                font-size: 9px;
            }
            .signature-container {
                margin-top: 30px;
                float: right;
                width: 200px;
                text-align: center;
            }
            .signature-space {
                height: 50px;
            }
        </style>
    </head>
    <body>
        <h1>Laporan Surat Tugas Per Pegawai</h1>
        <p class="subtitle">
            Periode: <?php echo date('d/m/Y', strtotime($start_date)); ?> s/d <?php echo date('d/m/Y', strtotime($end_date)); ?>
            <?php if (!empty($filter_nip) && isset($laporan[0]['nama'])): ?>
                | Pegawai: <?php echo htmlspecialchars($laporan[0]['nama'], ENT_QUOTES, 'UTF-8'); ?>
            <?php endif; ?>
        </p>

        <?php if (count($laporan) > 0): ?>
            <?php foreach ($laporan as $pegawai): ?>
                <div class="employee-card">
                    <div class="employee-header">
                        <table style="width: 100%; border: none;">
                            <tr style="border: none;">
                                <td style="width: 70%; border: none; padding: 0; color: white; vertical-align: middle;">
                                    <div class="employee-name"><?php echo htmlspecialchars($pegawai['nama'], ENT_QUOTES, 'UTF-8'); ?></div>
                                    <p class="employee-info">NIP: <?php echo htmlspecialchars($pegawai['nip'], ENT_QUOTES, 'UTF-8'); ?> | Pangkat: <?php echo htmlspecialchars($pegawai['pangkat'], ENT_QUOTES, 'UTF-8'); ?> | Jabatan: <?php echo htmlspecialchars($pegawai['jabatan'], ENT_QUOTES, 'UTF-8'); ?></p>
                                </td>
                                <td style="width: 30%; border: none; padding: 0; text-align: right; color: white; vertical-align: middle;">
                                    <strong>Total Tugas: <?php echo $pegawai['total_tugas']; ?></strong> (<?php echo $pegawai['total_hari']; ?> Hari)
                                </td>
                            </tr>
                        </table>
                    </div>
                    <table style="width: 100%;">
                        <thead>
                            <tr>
                                <th style="width: 5%;" class="text-center">No</th>
                                <th style="width: 25%;">Nomor Surat</th>
                                <th style="width: 15%;">Tanggal Surat</th>
                                <th style="width: 35%;">Tujuan / Maksud</th>
                                <th style="width: 15%;">Pelaksanaan</th>
                                <th style="width: 8%;" class="text-center">Durasi</th>
                            </tr>
                        </thead>
                        <tbody>
                            <?php $no = 1; foreach ($pegawai['surat_tugas'] as $surat): ?>
                                <tr>
                                    <td class="text-center"><?php echo $no++; ?></td>
                                    <td><?php echo htmlspecialchars($surat['nomor_surat'], ENT_QUOTES, 'UTF-8'); ?></td>
                                    <td class="text-nowrap"><?php echo date('d/m/Y', strtotime($surat['tanggal_surat'])); ?></td>
                                    <td><?php echo htmlspecialchars($surat['untuk'], ENT_QUOTES, 'UTF-8'); ?></td>
                                    <td class="text-nowrap">
                                        <?php 
                                        $pelaksanaan = date('d/m/Y', strtotime($surat['tanggal_mulai']));
                                        if ($surat['tanggal_mulai'] != $surat['tanggal_selesai']) {
                                            $pelaksanaan .= ' - ' . date('d/m/Y', strtotime($surat['tanggal_selesai']));
                                        }
                                        echo $pelaksanaan;
                                        ?>
                                    </td>
                                    <td class="text-center">
                                        <span class="badge"><?php echo $surat['durasi_hari']; ?> hari</span>
                                    </td>
                                </tr>
                            <?php endforeach; ?>
                        </tbody>
                    </table>
                </div>
            <?php endforeach; ?>
        <?php else: ?>
            <div style="text-align: center; padding: 30px; border: 1px solid #94a3b8;">
                <p>Tidak ada data surat tugas dengan filter yang dipilih.</p>
            </div>
        <?php endif; ?>

        <div class="signature-container">
            <p>Mengetahui,</p>
            <div class="signature-space"></div>
            <p style="font-weight: bold; text-decoration: underline;">Nama Pejabat</p>
            <p>NIP. ...........................</p>
        </div>
    </body>
    </html>
    <?php
    $html = ob_get_clean();

    $options = new Options();
    $options->set('isHtml5ParserEnabled', true);
    $options->set('isRemoteEnabled', true);
    
    $dompdf = new Dompdf($options);
    $dompdf->loadHtml($html);
    $dompdf->setPaper('A4', 'portrait');
    $dompdf->render();
    
    $suffixPegawai = !empty($filter_nip) ? '_' . $filter_nip : '_semua_pegawai';
    $filename = 'laporan_per_pegawai_' . $start_date . '_sd_' . $end_date . $suffixPegawai . '.pdf';
    $filename = preg_replace('/[^a-zA-Z0-9_\-\.]/', '_', $filename);
    $dompdf->stream($filename, ['Attachment' => false]);
    exit;
}

include '../includes/header.php';
?>

<div class="space-y-6">
    <div class="flex flex-col md:flex-row justify-between items-center gap-4 print:hidden">
        <h2 class="text-2xl font-bold text-slate-800">Laporan Surat Tugas Per Pegawai</h2>
        <div class="flex gap-2">
            <a href="?start_date=<?php echo urlencode($start_date); ?>&end_date=<?php echo urlencode($end_date); ?>&nip=<?php echo urlencode($filter_nip); ?>&export=excel" class="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg flex items-center gap-2 transition-colors shadow-sm">
                <i class='bx bx-spreadsheet'></i> Export Excel
            </a>
            <a href="?start_date=<?php echo urlencode($start_date); ?>&end_date=<?php echo urlencode($end_date); ?>&nip=<?php echo urlencode($filter_nip); ?>&export=pdf" target="_blank" class="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg flex items-center gap-2 transition-colors shadow-sm">
                <i class='bx bx-printer'></i> Cetak Laporan
            </a>
        </div>
    </div>

    <!-- Filter Section (Hidden on Print) -->
    <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-4 print:hidden">
        <form action="" method="GET" class="space-y-4">
            <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                    <label class="block text-sm font-medium text-slate-700 mb-1">Dari Tanggal</label>
                    <input type="date" name="start_date" value="<?php echo $start_date; ?>" 
                        class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                </div>
                <div>
                    <label class="block text-sm font-medium text-slate-700 mb-1">Sampai Tanggal</label>
                    <input type="date" name="end_date" value="<?php echo $end_date; ?>" 
                        class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                </div>
                <div>
                    <label class="block text-sm font-medium text-slate-700 mb-1">Pegawai (Opsional)</label>
                    <select name="nip" class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
                        <option value="">Semua Pegawai</option>
                        <?php foreach ($daftar_pegawai as $peg): ?>
                            <option value="<?php echo $peg['nip']; ?>" <?php echo ($filter_nip == $peg['nip']) ? 'selected' : ''; ?>>
                                <?php echo htmlspecialchars($peg['nama']); ?>
                            </option>
                        <?php endforeach; ?>
                    </select>
                </div>
            </div>
            <div class="flex gap-2">
                <button type="submit" class="bg-indigo-600 hover:bg-indigo-700 text-white px-6 py-2 rounded-lg transition-colors">
                    <i class='bx bx-search mr-2'></i> Tampilkan
                </button>
                <a href="laporan-pegawai.php" class="bg-slate-200 hover:bg-slate-300 text-slate-700 px-6 py-2 rounded-lg transition-colors flex items-center gap-2">
                    <i class='bx bx-reset'></i> Reset
                </a>
            </div>
        </form>
    </div>

    <!-- Report Header (Visible only on Print) -->
    <div class="hidden print:block text-center mb-8">
        <h1 class="text-2xl font-bold uppercase mb-2">Laporan Surat Tugas Per Pegawai</h1>
        <p class="text-slate-600">
            Periode: <?php echo date('d/m/Y', strtotime($start_date)); ?> s/d <?php echo date('d/m/Y', strtotime($end_date)); ?>
            <?php if (!empty($filter_nip)): ?>
                | Pegawai: <?php echo htmlspecialchars($laporan[0]['nama'] ?? ''); ?>
            <?php endif; ?>
        </p>
    </div>

    <!-- Laporan Content -->
    <?php if (count($laporan) > 0): ?>
        <div class="space-y-6">
            <?php foreach ($laporan as $pegawai): ?>
                <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden print:break-inside-avoid">
                    <!-- Header Pegawai -->
                    <div class="bg-gradient-to-r from-indigo-500 to-indigo-600 p-6 text-white">
                        <div class="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                            <div>
                                <h3 class="text-xl font-bold mb-1"><?php echo htmlspecialchars($pegawai['nama']); ?></h3>
                                <div class="text-sm text-indigo-100 space-y-1">
                                    <p><span class="font-medium">NIP:</span> <?php echo htmlspecialchars($pegawai['nip']); ?></p>
                                    <p><span class="font-medium">Pangkat:</span> <?php echo htmlspecialchars($pegawai['pangkat']); ?></p>
                                    <p><span class="font-medium">Jabatan:</span> <?php echo htmlspecialchars($pegawai['jabatan']); ?></p>
                                </div>
                            </div>
                            <div class="text-right">
                                <div class="bg-white/20 backdrop-blur-sm rounded-lg px-4 py-2">
                                    <p class="text-xs text-indigo-100 mb-1">Total Tugas</p>
                                    <p class="text-3xl font-bold"><?php echo $pegawai['total_tugas']; ?></p>
                                </div>
                                <div class="mt-2 text-sm text-indigo-100">
                                    <p>Total Hari: <span class="font-bold"><?php echo $pegawai['total_hari']; ?> hari</span></p>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- Daftar Surat Tugas -->
                    <div class="p-6">
                        <div class="overflow-x-auto">
                            <table class="w-full text-left border-collapse">
                                <thead>
                                    <tr class="bg-slate-50 text-slate-600 text-sm uppercase tracking-wider print:bg-white print:border-b-2 print:border-slate-800">
                                        <th class="px-4 py-3 font-semibold border-b border-slate-100 print:px-2 print:py-2">No</th>
                                        <th class="px-4 py-3 font-semibold border-b border-slate-100 print:px-2 print:py-2">Nomor Surat</th>
                                        <th class="px-4 py-3 font-semibold border-b border-slate-100 print:px-2 print:py-2">Tanggal Surat</th>
                                        <th class="px-4 py-3 font-semibold border-b border-slate-100 print:px-2 print:py-2">Tujuan / Maksud</th>
                                        <th class="px-4 py-3 font-semibold border-b border-slate-100 print:px-2 print:py-2">Pelaksanaan</th>
                                        <th class="px-4 py-3 font-semibold border-b border-slate-100 print:px-2 print:py-2">Durasi</th>
                                    </tr>
                                </thead>
                                <tbody class="divide-y divide-slate-100 print:divide-slate-200">
                                    <?php $no = 1; foreach ($pegawai['surat_tugas'] as $surat): ?>
                                        <tr class="hover:bg-slate-50 transition-colors print:hover:bg-white">
                                            <td class="px-4 py-3 text-sm text-slate-600 print:px-2 print:py-2 align-top"><?php echo $no++; ?></td>
                                            <td class="px-4 py-3 text-sm font-medium text-slate-900 print:px-2 print:py-2 align-top">
                                                <?php echo htmlspecialchars($surat['nomor_surat']); ?>
                                            </td>
                                            <td class="px-4 py-3 text-sm text-slate-600 print:px-2 print:py-2 align-top whitespace-nowrap">
                                                <?php echo date('d/m/Y', strtotime($surat['tanggal_surat'])); ?>
                                            </td>
                                            <td class="px-4 py-3 text-sm text-slate-600 print:px-2 print:py-2 align-top">
                                                <?php echo htmlspecialchars($surat['untuk']); ?>
                                            </td>
                                            <td class="px-4 py-3 text-sm text-slate-600 print:px-2 print:py-2 align-top whitespace-nowrap">
                                                <?php 
                                                echo date('d/m/Y', strtotime($surat['tanggal_mulai']));
                                                if ($surat['tanggal_mulai'] != $surat['tanggal_selesai']) {
                                                    echo ' - ' . date('d/m/Y', strtotime($surat['tanggal_selesai']));
                                                }
                                                ?>
                                            </td>
                                            <td class="px-4 py-3 text-sm text-slate-600 print:px-2 print:py-2 align-top">
                                                <span class="inline-flex items-center px-2 py-1 rounded-full text-xs font-medium bg-indigo-100 text-indigo-800">
                                                    <?php echo $surat['durasi_hari']; ?> hari
                                                </span>
                                            </td>
                                        </tr>
                                    <?php endforeach; ?>
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            <?php endforeach; ?>
        </div>
    <?php else: ?>
        <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-12 text-center">
            <i class='bx bx-search text-6xl mb-4 text-slate-300'></i>
            <p class="text-lg font-medium text-slate-600 mb-2">Tidak ada data</p>
            <p class="text-sm text-slate-500">
                <?php if (!empty($filter_nip) || $start_date || $end_date): ?>
                    Tidak ditemukan surat tugas dengan filter yang dipilih.
                <?php else: ?>
                    Belum ada surat tugas yang tercatat.
                <?php endif; ?>
            </p>
        </div>
    <?php endif; ?>

    <!-- Pagination -->
    <?php if ($total_pages > 1): ?>
        <div class="px-6 py-4 border-t border-slate-100 flex justify-center print:hidden">
            <nav class="flex items-center gap-1">
                <?php if ($page > 1): ?>
                    <a href="?start_date=<?php echo urlencode($start_date); ?>&end_date=<?php echo urlencode($end_date); ?>&nip=<?php echo urlencode($filter_nip); ?>&page=<?php echo ($page - 1); ?>"
                       class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm">Previous</a>
                <?php endif; ?>

                <?php
                    $maxPageLinks = 7; // batasi tampilan nomor halaman
                    $half = (int)floor($maxPageLinks / 2);
                    $from = max(1, $page - $half);
                    $to = min($total_pages, $from + $maxPageLinks - 1);
                    $from = max(1, $to - $maxPageLinks + 1);
                ?>
                <?php for ($i = $from; $i <= $to; $i++): ?>
                    <a href="?start_date=<?php echo urlencode($start_date); ?>&end_date=<?php echo urlencode($end_date); ?>&nip=<?php echo urlencode($filter_nip); ?>&page=<?php echo $i; ?>"
                       class="px-3 py-1 rounded border <?php echo($page == $i) ? 'bg-indigo-600 text-white border-indigo-600' : 'border-slate-200 text-slate-600 hover:bg-slate-50'; ?> text-sm">
                        <?php echo $i; ?>
                    </a>
                <?php endfor; ?>

                <?php if ($page < $total_pages): ?>
                    <a href="?start_date=<?php echo urlencode($start_date); ?>&end_date=<?php echo urlencode($end_date); ?>&nip=<?php echo urlencode($filter_nip); ?>&page=<?php echo ($page + 1); ?>"
                       class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm">Next</a>
                <?php endif; ?>
            </nav>
        </div>
    <?php endif; ?>
</div>

<style>
    @media print {
        @page {
            size: A4;
            margin: 1cm;
        }
        body {
            background: white;
        }
        /* Hide sidebar and other non-printable elements */
        aside, header, .print\:hidden {
            display: none !important;
        }
        /* Ensure main content takes full width */
        .flex-1 {
            margin: 0 !important;
            padding: 0 !important;
            width: 100% !important;
        }
        /* Reset shadows and borders for clean print */
        .shadow-sm, .shadow-lg {
            box-shadow: none !important;
        }
        .border {
            border: 1px solid #000 !important;
        }
        /* Show hidden print elements */
        .print\:block {
            display: block !important;
        }
        /* Table styling for print */
        table {
            width: 100% !important;
            border: 1px solid #000 !important;
        }
        th, td {
            border: 1px solid #000 !important;
            padding: 8px !important;
        }
        thead tr {
            background-color: #f0f0f0 !important;
            -webkit-print-color-adjust: exact;
        }
        /* Avoid breaking inside cards */
        .print\:break-inside-avoid {
            break-inside: avoid;
            page-break-inside: avoid;
        }
    }
</style>

<?php include '../includes/footer.php'; ?>

